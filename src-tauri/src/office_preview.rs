// SPDX-License-Identifier: GPL-3.0-or-later
// License: GNU GPLv3 or later. See the license file in the project root for more information.
// Copyright © 2021 - present Aleksey Hoffman. All rights reserved.

//! Office 文档预览：把 .docx / .xlsx / .pptx 转换成 PDF，复用现有的 PDF 预览通道（iframe）。
//!
//! 设计要点：
//! - **不捆绑任何转换器**（项目以体积为卖点，LibreOffice 会带来数百 MB）。
//!   改为探测用户已安装的转换器：优先 LibreOffice（跨平台），Windows 上回退 WPS。
//! - 转换产物写进系统临时目录，并按「源文件路径 + 修改时间 + 大小」缓存，
//!   避免每次选中文件都重新转换。
//! - 转换器缺失或转换失败时返回明确错误字符串，由前端降级处理。

use std::path::{Path, PathBuf};
use std::process::Command;
use std::time::UNIX_EPOCH;

/// 支持的 Office 扩展名（与前端 OFFICE_EXTENSIONS 保持一致）
const OFFICE_EXTENSIONS: &[&str] = &["docx", "doc", "xlsx", "xls", "pptx", "ppt", "rtf", "odt", "ods", "odp"];

#[cfg(windows)]
const CREATE_NO_WINDOW: u32 = 0x0800_0000;

/// 转换结果
#[derive(serde::Serialize)]
pub struct OfficePreviewResult {
    /// 转换后 PDF 的绝对路径（前端用 convertFileSrc 转成 asset URL 喂给 iframe）
    pub pdf_path: String,
    /// 是否命中了缓存
    pub cached: bool,
    /// 使用的转换器标识（便于诊断）
    pub converter: String,
}

/// 转换器描述
#[derive(Clone, Debug)]
struct Converter {
    /// 可执行文件路径
    program: PathBuf,
    /// 供前端/日志显示的标识
    kind: &'static str,
}

/// Prober 可执行文件的候选名称。
/// 注意：WPS 的转换能力不在 wps.exe 上，而在同目录的 **kwpsconvert.exe**（wpscli）。
/// 实测 `wps.exe --convert-to` 无效；`kwpsconvert.exe word2pdf ...` 才可用。
#[cfg(windows)]
const PROBER_NAMES: &[&str] = &["wpscli.exe", "kwpsconvert.exe"];

/// 在 PATH 中查找可执行文件
fn find_in_path(name: &str) -> Option<PathBuf> {
    let path_var = std::env::var_os("PATH")?;
    let exts: Vec<String> = if cfg!(windows) {
        std::env::var("PATHEXT")
            .unwrap_or_else(|_| ".EXE;.COM;.BAT;.CMD".to_string())
            .split(';')
            .map(|s| s.to_lowercase())
            .collect()
    } else {
        vec![String::new()]
    };

    for dir in std::env::split_paths(&path_var) {
        for ext in &exts {
            let candidate = dir.join(format!("{name}{ext}"));
            if candidate.is_file() {
                return Some(candidate);
            }
        }
    }
    None
}

/// 按扩展名选择 wpscli 的转换子命令
fn wps_subcommand(ext: &str) -> Option<&'static str> {
    match ext {
        "docx" | "doc" | "dot" | "rtf" | "wps" | "odt" => Some("word2pdf"),
        "xlsx" | "xls" | "ods" => Some("excel2pdf"),
        "pptx" | "ppt" | "odp" => Some("ppt2pdf"),
        _ => None,
    }
}

/// 找到 WPS 的 kwpsconvert.exe（真正的转换器）
#[cfg(windows)]
fn find_wps_cli() -> Option<PathBuf> {
    // 1) PATH 里找
    for name in PROBER_NAMES {
        if let Some(p) = find_in_path(name) {
            return Some(p);
        }
    }

    // 2) office6 目录里找（wps.exe 通常不在 PATH 上）
    let dir = find_wps_office_dir()?;
    for name in PROBER_NAMES {
        let candidate = dir.join(name);
        if candidate.is_file() {
            return Some(candidate);
        }
    }
    None
}

/// 读取注册表某个**命名值**（不是默认值）。
///
/// 实测教训：WPS 把安装路径放在 `InstallRoot` 命名值里，而**默认值是空字符串**——
/// 早期版本只读默认值，导致完全探测不到 WPS。
#[cfg(windows)]
fn reg_read_value(root: &str, subkey: &str, name: &str) -> Option<String> {
    let mut cmd = Command::new("reg.exe");
    cmd.args(["query", &format!("{root}\\{subkey}"), "/v", name]);
    {
        use std::os::windows::process::CommandExt;
        cmd.creation_flags(CREATE_NO_WINDOW);
    }
    let output = cmd.output().ok()?;

    let text = String::from_utf8_lossy(&output.stdout);
    for line in text.lines() {
        let trimmed = line.trim();
        // 形如: "    InstallRoot    REG_SZ    D:\\...\\WPS Office\\12.1.0.28505"
        if trimmed.to_lowercase().starts_with(&name.to_lowercase()) && trimmed.contains("REG_SZ") {
            if let Some((_, value)) = trimmed.split_once("REG_SZ") {
                let v = value.trim();
                if !v.is_empty() {
                    return Some(v.to_string());
                }
            }
        }
    }
    None
}

/// 从 `reg query ... /ve` 的输出里解析出 exe 所在目录。
///
/// 输入行形如：`    (默认)    REG_SZ    "D:\...\office6\wps.exe" /prometheus /wps "%1"`
/// 抽成纯函数是为了可单元测试（不需要真实注册表）。
fn parse_wps_dir_from_reg_output(text: &str) -> Option<PathBuf> {
    for line in text.lines() {
        let Some((_, value)) = line.split_once("REG_SZ") else {
            continue;
        };
        let value = value.trim();

        // 取到第一个 ".exe" 为止。
        //
        // 安全要点：这里**必须按字符**定位与切片，不能混用字节/小写字符串下标——
        // 路径可能含非 ASCII 字符（例如中文用户名），
        // 而 `to_lowercase()` 对某些字符会改变**字符数**，用它的下标去切原串会错位或越界。
        let chars: Vec<char> = value.chars().collect();
        let lower: Vec<char> = value.to_lowercase().chars().collect();

        let mut end: Option<usize> = None;
        if lower.len() >= 4 {
            for i in 0..=(lower.len() - 4) {
                if lower[i] == '.' && lower[i + 1] == 'e' && lower[i + 2] == 'x' && lower[i + 3] == 'e' {
                    end = Some(i + 4);
                    break;
                }
            }
        }

        let Some(end) = end else {
            continue;
        };
        // 小写化可能改变长度；取两者较小值保证不越界
        let end = end.min(chars.len());
        let exe_path_str: String = chars[..end].iter().collect();
        let exe_path = PathBuf::from(exe_path_str.trim().trim_matches('"'));

        if let Some(dir) = exe_path.parent() {
            return Some(dir.to_path_buf());
        }
    }
    None
}

/// 通过已注册的文件关联反查 WPS 安装目录。
///
/// `HKCR\WPS.Docx.6\shell\open\command` 的默认值形如：
/// `"D:\...\WPS Office\12.1.0.28505\office6\wps.exe" /prometheus /wps "%1"`
/// 这是**比目录扫描更可靠**的线索（能正确处理装到非系统盘的情况）。
#[cfg(windows)]
fn find_wps_dir_from_association() -> Option<PathBuf> {
    use std::os::windows::process::CommandExt;

    let keys = [
        r"HKCR\WPS.Docx.6\shell\open\command",
        r"HKCR\ET.Xlsx.6\shell\open\command",
        r"HKCR\WPP.PPTX.6\shell\open\command",
    ];

    for key in keys {
        let mut cmd = Command::new("reg.exe");
        cmd.args(["query", key, "/ve"]);
        cmd.creation_flags(CREATE_NO_WINDOW);

        let Ok(output) = cmd.output() else {
            continue;
        };
        let text = String::from_utf8_lossy(&output.stdout);

        if let Some(dir) = parse_wps_dir_from_reg_output(&text) {
            if dir.join("wps.exe").is_file() || dir.join("kwpsconvert.exe").is_file() {
                return Some(dir);
            }
        }
    }
    None
}

/// 在 Windows 上定位 WPS 的 office6 目录（含 wps.exe / kwpsconvert.exe）
#[cfg(windows)]
fn find_wps_office_dir() -> Option<PathBuf> {
    let mut roots: Vec<PathBuf> = Vec::new();

    // 1) 注册表 `InstallRoot` 命名值 —— 实测最可靠（能处理非系统盘安装）
    let mut last_parent: Option<PathBuf> = None;
    for (root, subkey) in [
        ("HKLM", r"SOFTWARE\Kingsoft\Office\6.0\common"),
        ("HKCU", r"SOFTWARE\Kingsoft\Office\6.0\common"),
        ("HKLM", r"SOFTWARE\WOW6432Node\Kingsoft\Office\6.0\common"),
        ("HKCU", r"SOFTWARE\WOW6432Node\Kingsoft\Office\6.0\common"),
    ] {
        if let Some(value) = reg_read_value(root, subkey, "InstallRoot") {
            let p = PathBuf::from(value);
            last_parent = p.parent().map(|x| x.to_path_buf());
            roots.push(p);
        }
    }

    // 2) 文件关联反查
    if let Some(dir) = find_wps_dir_from_association() {
        roots.push(dir);
    }

    // 3) 常规候选目录（同盘 + 跨盘：KINGSOFT 可能装在 D:）
    let mut candidates: Vec<PathBuf> = Vec::new();
    for var in ["LOCALAPPDATA", "APPDATA", "ProgramFiles", "ProgramFiles(x86)"] {
        if let Some(base) = std::env::var_os(var) {
            candidates.push(PathBuf::from(base).join("Kingsoft").join("WPS Office"));
        }
    }
    // 从 InstallRoot 的父目录（即 "...\Kingsoft\WPS Office"）也扫一遍，
    // 这样能覆盖"新版装在 D 盘、LOCALAPPDATA 在 C 盘"的情况
    if let Some(p) = last_parent {
        candidates.push(p);
    }

    for candidate in candidates {
        for dir in expand_version_dirs(&candidate) {
            roots.push(dir);
        }
    }

    // 逐个校验：目录里有 wps.exe 或 kwpsconvert.exe
    for r in roots {
        for probe in [r.clone(), r.join("office6")] {
            if probe.join("wps.exe").is_file() || probe.join("kwpsconvert.exe").is_file() {
                return Some(probe);
            }
        }
    }

    None
}

/// 展开 "<root>\<版本号>\office6" 这类目录（WPS 版本号目录名会变，故扫一层）
#[cfg(windows)]
fn expand_version_dirs(root: &Path) -> Vec<PathBuf> {
    let mut out = Vec::new();
    let Ok(entries) = std::fs::read_dir(root) else {
        return out;
    };
    for entry in entries.flatten() {
        let p = entry.path();
        if p.is_dir() {
            let o6 = p.join("office6");
            if o6.is_dir() {
                out.push(o6);
            } else {
                out.push(p);
            }
        }
    }
    out
}

/// 探测可用的转换器
fn detect_converter() -> Option<Converter> {
    // 1) LibreOffice / OpenOffice（跨平台）
    for name in ["soffice", "soffice.exe", "libreoffice"] {
        if let Some(program) = find_in_path(name) {
            return Some(Converter {
                program,
                kind: "libreoffice",
            });
        }
    }

    // macOS 常见位置
    #[cfg(target_os = "macos")]
    {
        let p = PathBuf::from("/Applications/LibreOffice.app/Contents/MacOS/soffice");
        if p.is_file() {
            return Some(Converter {
                program: p,
                kind: "libreoffice",
            });
        }
    }

    // 2) Windows 上回退 WPS（用 kwpsconvert.exe，不是 wps.exe）
    #[cfg(windows)]
    {
        if let Some(cli) = find_wps_cli() {
            return Some(Converter {
                program: cli,
                kind: "wps",
            });
        }
    }

    None
}

/// 简易稳定哈希（不引入依赖；用于缓存文件名）
fn fnv1a_64(input: &str) -> u64 {
    let mut hash: u64 = 0xcbf2_9ce4_8422_2325;
    for byte in input.as_bytes() {
        hash ^= *byte as u64;
        hash = hash.wrapping_mul(0x0000_0100_0000_01b3);
    }
    hash
}

/// 取源文件的「路径 + mtime + 大小」指纹，用于缓存失效判断
fn source_fingerprint(path: &Path) -> Result<(u64, String), String> {
    let meta = std::fs::metadata(path).map_err(|error| format!("Failed to stat file: {error}"))?;
    let mtime = meta
        .modified()
        .ok()
        .and_then(|t| t.duration_since(UNIX_EPOCH).ok())
        .map(|d| d.as_secs())
        .unwrap_or(0);
    let key = format!("{}|{}|{}", path.display(), mtime, meta.len());
    Ok((mtime, key))
}

/// 执行转换，成功则返回产物路径
fn run_conversion(converter: &Converter, source: &Path, out_dir: &Path) -> Result<PathBuf, String> {
    std::fs::create_dir_all(out_dir)
        .map_err(|error| format!("Failed to create temp dir: {error}"))?;

    // 产物名固定为源文件主名，便于定位
    let stem = source
        .file_stem()
        .map(|s| s.to_string_lossy().to_string())
        .unwrap_or_else(|| "output".to_string());
    let out_pdf = out_dir.join(format!("{stem}.pdf"));

    // 已经转好就直接用（同一目录下重名时）
    if out_pdf.is_file() {
        let _ = std::fs::remove_file(&out_pdf);
    }

    let mut cmd = Command::new(&converter.program);
    match converter.kind {
        "wps" => {
            // WPS: kwpsconvert.exe <subcommand> <input> --output <file.pdf>
            // 实测语法（wpscli --help）：
            //   wpscli word2pdf  <input> [--output <file.pdf>]
            //   wpscli excel2pdf <input> [--output <file.pdf>]
            //   wpscli ppt2pdf   <input> [--output <file.pdf>]
            let ext = source
                .extension()
                .map(|e| e.to_string_lossy().to_lowercase())
                .unwrap_or_default();
            let subcommand = wps_subcommand(&ext)
                .ok_or_else(|| format!("No wpscli subcommand for extension .{ext}"))?;

            cmd.arg(subcommand)
                .arg(source)
                .arg("--output")
                .arg(&out_pdf);
        }
        _ => {
            // LibreOffice: soffice --headless --convert-to pdf --outdir <dir> <src>
            cmd.arg("--headless")
                .arg("--norestore")
                .arg("--convert-to")
                .arg("pdf")
                .arg("--outdir")
                .arg(out_dir)
                .arg(source);
        }
    }

    // 关键：不要捕获 stdout/stderr 管道（沙箱/侧车环境下管道可能不可用），
    // 也让子进程不弹控制台窗口。转换成功与否用退出码 + 产物文件判断。
    cmd.stdout(std::process::Stdio::null())
        .stderr(std::process::Stdio::null())
        .stdin(std::process::Stdio::null());

    #[cfg(windows)]
    {
        use std::os::windows::process::CommandExt;
        cmd.creation_flags(CREATE_NO_WINDOW);
    }

    let status = cmd
        .status()
        .map_err(|error| format!("Failed to run {}: {error}", converter.kind))?;

    if !status.success() {
        let code = status.code();
        // wpscli 的退出码语义（见 wpscli <sub> --help）
        let hint = match code {
            Some(100) => " (WPS is not signed in)",
            Some(101) => " (WPS account lacks the required privilege)",
            Some(207) => " (input file exceeds 200 MB)",
            Some(209) => " (source document requires a password)",
            Some(210) => " (wrong document password)",
            Some(203) => " (input file not found)",
            Some(211) => " (output directory does not exist)",
            Some(212) => " (no write permission to output directory)",
            _ => "",
        };
        return Err(format!(
            "{} exited with status {code:?}{hint}",
            converter.kind
        ));
    }

    // 有些转换器会把产物命名成别的名字，兜底扫一遍目录里最新的 pdf
    if out_pdf.is_file() {
        return Ok(out_pdf);
    }

    let mut newest: Option<(std::time::SystemTime, PathBuf)> = None;
    if let Ok(entries) = std::fs::read_dir(out_dir) {
        for entry in entries.flatten() {
            let p = entry.path();
            if p.extension().map(|e| e.eq_ignore_ascii_case("pdf")).unwrap_or(false) {
                if let Ok(meta) = entry.metadata() {
                    if let Ok(modified) = meta.modified() {
                        if newest.as_ref().map(|(t, _)| modified > *t).unwrap_or(true) {
                            newest = Some((modified, p));
                        }
                    }
                }
            }
        }
    }

    newest
        .map(|(_, p)| p)
        .ok_or_else(|| format!("{} did not produce a PDF", converter.kind))
}

/// 把 Office 文档转换成 PDF 并返回产物路径。
///
/// 前端用法：
/// ```ts
/// const result = await invoke<{ pdf_path: string; cached: boolean; converter: string }>(
///   'convert_office_to_pdf', { path: entry.path }
/// );
/// const url = convertFileSrc(result.pdf_path);
/// ```
#[tauri::command]
pub async fn convert_office_to_pdf(path: String) -> Result<OfficePreviewResult, String> {
    let source = PathBuf::from(&path);

    if !source.is_file() {
        return Err(format!("File not found: {path}"));
    }

    let ext = source
        .extension()
        .map(|e| e.to_string_lossy().to_lowercase())
        .unwrap_or_default();

    if !OFFICE_EXTENSIONS.contains(&ext.as_str()) {
        return Err(format!("Unsupported office extension: .{ext}"));
    }

    let converter = detect_converter().ok_or_else(|| {
        "No office document converter found. Install LibreOffice, or WPS Office on Windows."
            .to_string()
    })?;

    let (_mtime, key) = source_fingerprint(&source)?;
    let hash = fnv1a_64(&key);

    let cache_root = std::env::temp_dir().join("sigma-file-manager").join("office-preview");
    let out_dir = cache_root.join(format!("{hash:016x}"));

    // 缓存命中：产物存在且比源文件新
    let stem = source
        .file_stem()
        .map(|s| s.to_string_lossy().to_string())
        .unwrap_or_default();
    let cached_pdf = out_dir.join(format!("{stem}.pdf"));

    if cached_pdf.is_file() {
        let fresh = match (std::fs::metadata(&cached_pdf), std::fs::metadata(&source)) {
            (Ok(out_meta), Ok(src_meta)) => match (out_meta.modified(), src_meta.modified()) {
                (Ok(out_t), Ok(src_t)) => out_t >= src_t,
                _ => true,
            },
            _ => false,
        };

        if fresh {
            return Ok(OfficePreviewResult {
                pdf_path: cached_pdf.to_string_lossy().to_string(),
                cached: true,
                converter: converter.kind.to_string(),
            });
        }
        let _ = std::fs::remove_file(&cached_pdf);
    }

    let pdf_path = run_conversion(&converter, &source, &out_dir)?;

    Ok(OfficePreviewResult {
        pdf_path: pdf_path.to_string_lossy().to_string(),
        cached: false,
        converter: converter.kind.to_string(),
    })
}

/// 诊断用：报告当前探测到的转换器（前端可在设置页展示）
#[tauri::command]
pub fn get_office_converter_info() -> Result<String, String> {
    match detect_converter() {
        Some(c) => Ok(format!(
            "{}: {}",
            c.kind,
            c.program.to_string_lossy()
        )),
        None => Ok("none".to_string()),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn fingerprint_changes_with_content() {
        // 只验证指纹函数本身稳定可复现（不依赖真实文件）
        let a = fnv1a_64("a|1|2");
        let b = fnv1a_64("a|1|2");
        let c = fnv1a_64("a|1|3");
        assert_eq!(a, b);
        assert_ne!(a, c);
    }

    #[test]
    fn extension_list_is_lowercase() {
        for ext in OFFICE_EXTENSIONS {
            assert_eq!(*ext, ext.to_lowercase());
        }
    }

    #[test]
    fn wps_subcommand_mapping_is_correct() {
        // 实测确认的 wpscli 子命令映射
        assert_eq!(wps_subcommand("docx"), Some("word2pdf"));
        assert_eq!(wps_subcommand("doc"), Some("word2pdf"));
        assert_eq!(wps_subcommand("rtf"), Some("word2pdf"));
        assert_eq!(wps_subcommand("xlsx"), Some("excel2pdf"));
        assert_eq!(wps_subcommand("xls"), Some("excel2pdf"));
        assert_eq!(wps_subcommand("pptx"), Some("ppt2pdf"));
        assert_eq!(wps_subcommand("ppt"), Some("ppt2pdf"));
        // 非 Office 类型不应有子命令
        assert_eq!(wps_subcommand("pdf"), None);
        assert_eq!(wps_subcommand("txt"), None);
    }

    #[test]
    fn every_office_extension_has_a_subcommand() {
        // 保证扩展名列表与子命令映射不脱节
        for ext in OFFICE_EXTENSIONS {
            assert!(
                wps_subcommand(ext).is_some(),
                "扩展名 .{ext} 在 OFFICE_EXTENSIONS 里，但没有对应的 wpscli 子命令"
            );
        }
    }

    #[test]
    fn parses_wps_dir_from_reg_output() {
        // 真实机器上的 `reg query HKCR\WPS.Docx.6\shell\open\command /ve` 输出形态
        let sample = concat!(
            "\r\n",
            "HKEY_CLASSES_ROOT\\WPS.Docx.6\\shell\\open\\command\r\n",
            "    (默认)    REG_SZ    \"D:\\Users\\Administrator\\AppData\\Local\\Kingsoft\\WPS Office\\12.1.0.28505\\office6\\wps.exe\" /prometheus /wps \"%1\"\r\n",
        );
        let dir = parse_wps_dir_from_reg_output(sample).expect("应能解析出目录");
        let s = dir.to_string_lossy().replace('/', "\\");
        assert!(s.ends_with(r"WPS Office\12.1.0.28505\office6"), "实际: {s}");
    }

    #[test]
    fn reg_output_without_exe_yields_none() {
        let sample = "HKEY_CLASSES_ROOT\\Foo\r\n    (默认)    REG_SZ    \r\n";
        assert!(parse_wps_dir_from_reg_output(sample).is_none());
    }

    #[test]
    fn parses_non_ascii_path_without_panic() {
        // 中文用户名/中文路径不能让解析 panic（曾用字节切片导致越界）
        let sample = "    (默认)    REG_SZ    \"D:\\用户\\张三\\Kingsoft\\WPS Office\\12.1.0.28505\\office6\\wps.exe\" /prometheus\n";
        let dir = parse_wps_dir_from_reg_output(sample).expect("非 ASCII 路径也应能解析");
        let s = dir.to_string_lossy().replace('/', "\\");
        assert!(s.ends_with(r"office6"), "实际: {s}");
    }

    // ─────────────────────────────────────────────────────────────
    // 集成测试：真正调用转换器（默认忽略，因为多数环境没装转换器）
    //
    // 运行方式：
    //   cargo test --lib office_preview -- --ignored --nocapture
    //
    // 该测试覆盖：detect_converter -> run_conversion -> PDF 产物校验，
    // 也就是把本模块的**真实执行路径**跑一遍，而不只是逻辑分支。
    // ─────────────────────────────────────────────────────────────

    /// 极简 DOCX 生成（DOCX 就是 ZIP + 若干 XML）。
    /// 手写 ZIP 是为了不引入额外依赖 —— 只用 stored(不压缩) 方式。
    fn build_minimal_docx(path: &Path) -> std::io::Result<()> {
        let xml_files: Vec<(&str, String)> = vec![
            (
                "[Content_Types].xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_string(),
            ),
            (
                "_rels/.rels",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_string(),
            ),
            (
                "word/document.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Sigma File Manager office preview integration test</w:t></w:r></w:p></w:body></w:document>"#.to_string(),
            ),
        ];

        let mut out: Vec<u8> = Vec::new();
        let mut central: Vec<u8> = Vec::new();

        for (name, content) in &xml_files {
            let data = content.as_bytes();
            let crc = crc32(data);
            let offset = out.len() as u32;

            // Local file header
            out.extend_from_slice(&0x0403_4b50u32.to_le_bytes());
            out.extend_from_slice(&20u16.to_le_bytes()); // version needed
            out.extend_from_slice(&0u16.to_le_bytes()); // flags
            out.extend_from_slice(&0u16.to_le_bytes()); // method = stored
            out.extend_from_slice(&0u16.to_le_bytes()); // mod time
            out.extend_from_slice(&0u16.to_le_bytes()); // mod date
            out.extend_from_slice(&crc.to_le_bytes());
            out.extend_from_slice(&(data.len() as u32).to_le_bytes()); // compressed
            out.extend_from_slice(&(data.len() as u32).to_le_bytes()); // uncompressed
            out.extend_from_slice(&(name.len() as u16).to_le_bytes());
            out.extend_from_slice(&0u16.to_le_bytes()); // extra len
            out.extend_from_slice(name.as_bytes());
            out.extend_from_slice(data);

            // Central directory entry
            central.extend_from_slice(&0x0201_4b50u32.to_le_bytes());
            central.extend_from_slice(&20u16.to_le_bytes()); // version made by
            central.extend_from_slice(&20u16.to_le_bytes()); // version needed
            central.extend_from_slice(&0u16.to_le_bytes());
            central.extend_from_slice(&0u16.to_le_bytes());
            central.extend_from_slice(&0u16.to_le_bytes());
            central.extend_from_slice(&0u16.to_le_bytes());
            central.extend_from_slice(&crc.to_le_bytes());
            central.extend_from_slice(&(data.len() as u32).to_le_bytes());
            central.extend_from_slice(&(data.len() as u32).to_le_bytes());
            central.extend_from_slice(&(name.len() as u16).to_le_bytes());
            central.extend_from_slice(&0u16.to_le_bytes()); // extra
            central.extend_from_slice(&0u16.to_le_bytes()); // comment
            central.extend_from_slice(&0u16.to_le_bytes()); // disk
            central.extend_from_slice(&0u16.to_le_bytes()); // internal attrs
            central.extend_from_slice(&0u32.to_le_bytes()); // external attrs
            central.extend_from_slice(&offset.to_le_bytes());
            central.extend_from_slice(name.as_bytes());
        }

        let central_offset = out.len() as u32;
        let central_size = central.len() as u32;
        let count = xml_files.len() as u16;
        out.extend_from_slice(&central);

        // End of central directory
        out.extend_from_slice(&0x0605_4b50u32.to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes()); // disk
        out.extend_from_slice(&0u16.to_le_bytes()); // disk with central
        out.extend_from_slice(&count.to_le_bytes());
        out.extend_from_slice(&count.to_le_bytes());
        out.extend_from_slice(&central_size.to_le_bytes());
        out.extend_from_slice(&central_offset.to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes()); // comment len

        std::fs::write(path, out)
    }

    /// 标准 CRC-32（ZIP 用）
    fn crc32(data: &[u8]) -> u32 {
        let mut crc: u32 = 0xffff_ffff;
        for byte in data {
            crc ^= *byte as u32;
            for _ in 0..8 {
                let mask = (crc & 1).wrapping_neg();
                crc = (crc >> 1) ^ (0xedb8_8320 & mask);
            }
        }
        !crc
    }

    #[test]
    fn crc32_matches_known_vector() {
        // 标准测试向量
        assert_eq!(crc32(b"123456789"), 0xcbf4_3926);
    }

    #[test]
    fn generates_valid_zip_signature() {
        let dir = std::env::temp_dir().join("sfm-office-preview-test-zip");
        let _ = std::fs::create_dir_all(&dir);
        let docx = dir.join("gen.docx");
        build_minimal_docx(&docx).expect("生成 docx 失败");

        let bytes = std::fs::read(&docx).expect("读回失败");
        assert!(bytes.len() > 100, "文件太小: {}", bytes.len());
        assert_eq!(&bytes[0..2], b"PK", "ZIP 头应为 PK");
        // 结尾应是 EOCD 签名
        assert!(
            bytes.windows(4).any(|w| w == [0x50, 0x4b, 0x05, 0x06]),
            "缺少 EOCD 记录"
        );
    }

    #[test]
    #[ignore = "需要本机已安装 LibreOffice 或 WPS；用 --ignored 运行"]
    fn integration_converts_real_document_to_pdf() {
        let converter = match detect_converter() {
            Some(c) => c,
            None => {
                eprintln!("跳过：本机未探测到 LibreOffice / WPS 转换器");
                return;
            }
        };
        eprintln!(
            "使用转换器: {} -> {}",
            converter.kind,
            converter.program.display()
        );

        let dir = std::env::temp_dir().join("sfm-office-preview-integration");
        let _ = std::fs::create_dir_all(&dir);
        let docx = dir.join("integration.docx");
        build_minimal_docx(&docx).expect("生成测试 docx 失败");
        eprintln!("测试文档: {} ({} 字节)", docx.display(), std::fs::metadata(&docx).unwrap().len());

        let out_dir = dir.join("out");
        let pdf = run_conversion(&converter, &docx, &out_dir).expect("转换失败");
        eprintln!("产物: {} ({} 字节)", pdf.display(), std::fs::metadata(&pdf).unwrap().len());

        let head = std::fs::read(&pdf).expect("读取产物失败");
        assert!(head.len() > 400, "PDF 过小，可能是空产物");
        assert_eq!(&head[0..5], b"%PDF-", "产物不是有效 PDF");
    }
}
