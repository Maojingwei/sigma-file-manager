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

/// Windows 注册表读取工具（避免为一个字符串引入额外依赖）
#[cfg(windows)]
fn reg_read_default(root: &str, subkey: &str) -> Option<String> {
    let output = {
        let mut cmd = Command::new("reg.exe");
        cmd.args(["query", &format!("{root}\\{subkey}"), "/ve"]);
        #[cfg(windows)]
        {
            use std::os::windows::process::CommandExt;
            cmd.creation_flags(CREATE_NO_WINDOW);
        }
        cmd.output().ok()?
    };

    let text = String::from_utf8_lossy(&output.stdout);
    // 形如: "    (默认)    REG_SZ    C:\\path\\WPS Office"
    for line in text.lines() {
        if line.contains("REG_SZ") {
            if let Some((_, value)) = line.split_once("REG_SZ") {
                let v = value.trim();
                if !v.is_empty() {
                    return Some(v.to_string());
                }
            }
        }
    }
    None
}

/// 在 Windows 上定位 WPS 的 office6 目录（含 wps.exe / et.exe / wpp.exe）
#[cfg(windows)]
fn find_wps_office_dir() -> Option<PathBuf> {
    // 1) 注册表常见位置
    let reg_candidates: [(&str, &str); 4] = [
        ("HKLM", r"SOFTWARE\Kingsoft\Office\6.0\common"),
        ("HKCU", r"SOFTWARE\Kingsoft\Office\6.0\common"),
        ("HKLM", r"SOFTWARE\WOW6432Node\Kingsoft\Office\6.0\common"),
        ("HKCU", r"SOFTWARE\WOW6432Node\Kingsoft\Office\6.0\common"),
    ];

    for (root, subkey) in reg_candidates {
        if let Some(value) = reg_read_default(root, subkey) {
            let p = PathBuf::from(&value);
            for candidate in [p.join("office6"), p.clone()] {
                if candidate.join("wps.exe").is_file() {
                    return Some(candidate);
                }
            }
        }
    }

    // 2) 在常见安装根的下一级目录里找（WPS 版本号目录名会变，所以扫一层）
    let mut roots: Vec<PathBuf> = Vec::new();
    if let Some(local) = std::env::var_os("LOCALAPPDATA") {
        roots.push(PathBuf::from(local).join("Kingsoft").join("WPS Office"));
    }
    if let Some(pf) = std::env::var_os("ProgramFiles") {
        roots.push(PathBuf::from(pf).join("Kingsoft").join("WPS Office"));
    }
    if let Some(pf) = std::env::var_os("ProgramFiles(x86)") {
        roots.push(PathBuf::from(pf).join("Kingsoft").join("WPS Office"));
    }

    for root in roots {
        let Ok(entries) = std::fs::read_dir(&root) else {
            continue;
        };
        for entry in entries.flatten() {
            let office6 = entry.path().join("office6");
            if office6.join("wps.exe").is_file() {
                return Some(office6);
            }
        }
    }

    None
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

    // 2) Windows 上回退 WPS
    #[cfg(windows)]
    {
        if let Some(dir) = find_wps_office_dir() {
            let wps = dir.join("wps.exe");
            if wps.is_file() {
                return Some(Converter {
                    program: wps,
                    kind: "wps",
                });
            }
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
            // WPS: wps.exe <src> --convert-to pdf --outdir <dir>
            cmd.arg(source)
                .arg("--convert-to")
                .arg("pdf")
                .arg("--outdir")
                .arg(out_dir);
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
    // 也让子进程不弹控制台窗口。转换成功与否用产物文件判断。
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
        return Err(format!(
            "{} exited with status {:?}",
            converter.kind,
            status.code()
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
}
