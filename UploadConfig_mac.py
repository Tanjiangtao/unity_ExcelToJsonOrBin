#!/usr/bin/env python3
"""Beginner-friendly macOS OSS upload entry point."""

from __future__ import annotations

import os
import platform
import shutil
import subprocess
import sys
from pathlib import Path

REMOTE_DIR = "oss://wminigame/VoodooDemo3/Config1/"


def base_dir() -> Path:
    if getattr(sys, "frozen", False):
        executable_dir = Path(sys.executable).resolve().parent
        if executable_dir.name == "MacOS" and executable_dir.parent.name == "Contents":
            # Prefer the directory containing Config/, supporting both
            # 配置表/dist/App.app and 配置表/App.app layouts.
            for candidate in (executable_dir.parents[3], executable_dir.parents[2]):
                if (candidate / "Config").is_dir():
                    return candidate
            return executable_dir.parents[3]
        return executable_dir
    return Path(__file__).resolve().parent


def show_dialog(title: str, message: str, success: bool = False) -> None:
    escaped = message.replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\n")
    title_escaped = title.replace("\\", "\\\\").replace('"', '\\"')
    script = f'display dialog "{escaped}" with title "{title_escaped}" buttons {{"确定"}} default button "确定"'
    if success:
        script += " giving up after 5"
    try:
        subprocess.run(["/usr/bin/osascript", "-e", script], check=False)
    except OSError:
        print(f"{title}:\n{message}", file=sys.stderr if not success else sys.stdout)


def find_ossutil() -> str | None:
    bundled_ossutil = Path(getattr(sys, "_MEIPASS", "")) / "ossutil" if getattr(sys, "frozen", False) else None
    candidates = [
        Path(os.environ["OSSUTIL_PATH"]) if os.environ.get("OSSUTIL_PATH") else None,
        bundled_ossutil,
        base_dir() / "ossutil",
        base_dir() / "ossutilmac64",
        Path("/usr/local/bin/ossutil"),
        Path("/opt/homebrew/bin/ossutil"),
    ]
    for candidate in candidates:
        if candidate and candidate.is_file() and os.access(candidate, os.X_OK):
            return str(candidate)
    return shutil.which("ossutil") or shutil.which("ossutilmac64")


def run_ossutil(ossutil: str, args: list[str], config_file: Path, timeout: int = 60) -> subprocess.CompletedProcess[str]:
    command = [ossutil, *args]
    command.extend(["-c", str(config_file)])
    return subprocess.run(command, text=True, capture_output=True, timeout=timeout)


def fail(title: str, message: str, no_dialog: bool) -> int:
    if no_dialog:
        print(f"{title}: {message}", file=sys.stderr)
    else:
        show_dialog(title, message)
    return 1


def main() -> int:
    no_dialog = "--no-dialog" in sys.argv[1:]
    if platform.system() != "Darwin":
        return fail("平台不兼容", "此 App 只能在 macOS 上运行。", no_dialog)

    ossutil = find_ossutil()
    if not ossutil:
        return fail(
            "未找到 ossutil",
            "请先安装 macOS 版 ossutil。\n\n"
            "终端执行：\n"
            "sudo -v ; curl https://gosspublic.alicdn.com/ossutil/install.sh | sudo bash",
            no_dialog,
        )

    configured_file = os.environ.get("OSSUTIL_CONFIG_FILE", "").strip()
    if configured_file:
        config_path = Path(configured_file).expanduser()
    else:
        local_config = base_dir() / ".ossutilconfig"
        config_path = local_config if local_config.is_file() else Path.home() / ".ossutilconfig"
    if not config_path.is_file():
        return fail(
            "未找到 OSS 账号配置",
            f"没有找到配置文件：\n{config_path}\n\n"
            "请先在终端执行：\nossutil config\n\n"
            "完成 Endpoint、AccessKey ID 和 AccessKey Secret 配置后再重试。",
            no_dialog,
        )

    try:
        check = run_ossutil(ossutil, ["ls", REMOTE_DIR], config_path, timeout=30)
    except (OSError, subprocess.TimeoutExpired) as exc:
        return fail("OSS 账号检查失败", f"无法运行 ossutil：{exc}", no_dialog)
    if check.returncode != 0:
        detail = (check.stderr or check.stdout).strip()
        return fail(
            "OSS 账号检查失败",
            "账号配置无法访问目标 CDN 路径。\n\n"
            f"目标：{REMOTE_DIR}\n"
            f"错误：{detail[-800:]}",
            no_dialog,
        )

    config_dir = base_dir() / "Config"
    files = [config_dir / "game_config.bin", config_dir / "config_version.json"]
    missing = [str(path) for path in files if not path.is_file()]
    if missing:
        return fail("缺少导出文件", "请先运行导表工具生成：\n" + "\n".join(missing), no_dialog)

    for path in files:
        try:
            result = run_ossutil(ossutil, ["cp", "-f", str(path), REMOTE_DIR], config_path, timeout=120)
        except (OSError, subprocess.TimeoutExpired) as exc:
            return fail("上传失败", f"文件：{path.name}\n错误：{exc}", no_dialog)
        if result.returncode != 0:
            detail = (result.stderr or result.stdout).strip()
            return fail("上传失败", f"文件：{path.name}\n错误：{detail[-800:]}", no_dialog)

    message = f"账号检查通过，已上传 2 个文件到：\n{REMOTE_DIR}"
    if no_dialog:
        print(message)
    else:
        show_dialog("上传成功", message, success=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
