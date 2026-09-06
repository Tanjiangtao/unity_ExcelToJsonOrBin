#!/usr/bin/env python3
"""macOS entry point for the Excel -> Unity game config exporter.

The export implementation lives in ExcelToGameConfig.py so Windows and macOS
always produce the same game_config.bin format.
"""

from __future__ import annotations

import argparse
import os
import platform
import sys
import traceback
from datetime import datetime

import ExcelToGameConfig as exporter


def _default_base_dir() -> str:
    if getattr(sys, "frozen", False):
        # PyInstaller .app layout: <dir>/<app>.app/Contents/MacOS/<binary>.
        executable_dir = os.path.dirname(os.path.abspath(sys.executable))
        if executable_dir.endswith(os.path.join("Contents", "MacOS")):
            return os.path.dirname(os.path.dirname(os.path.dirname(executable_dir)))
        return executable_dir
    return os.path.dirname(os.path.abspath(__file__))


def _parse_args() -> argparse.Namespace:
    base_dir = _default_base_dir()
    parser = argparse.ArgumentParser(description="将 Excel 配置表导出为 Unity game_config.bin/json")
    parser.add_argument(
        "--input-dir",
        default=base_dir,
        help="Excel 文件目录，默认是工具所在目录",
    )
    parser.add_argument(
        "--output-dir",
        default=os.path.join(base_dir, "Config"),
        help="导出目录，默认是工具目录下的 Config",
    )
    parser.add_argument(
        "--version",
        default=datetime.now().strftime("%Y.%m.%d.%H%M%S"),
        help="配置版本号，默认使用当前时间",
    )
    parser.add_argument(
        "--no-dialog",
        action="store_true",
        help="只输出终端信息，不显示 macOS 结果弹窗",
    )
    return parser.parse_args()


def main() -> int:
    args = _parse_args()
    if platform.system() != "Darwin":
        print("Warning: this entry point is intended for macOS; continuing anyway.", file=sys.stderr)

    input_dir = os.path.abspath(os.path.expanduser(args.input_dir))
    output_dir = os.path.abspath(os.path.expanduser(args.output_dir))

    try:
        tables = exporter.build_tables(input_dir)
        config_path = exporter.write_game_config_bin(output_dir, tables)
        json_path = exporter.write_game_config_json(output_dir, tables)
        version_path = exporter.write_version_file(output_dir, args.version, config_path)

        print("Build game_config.bin success")
        print(f"InputDir: {input_dir}")
        print(f"OutputDir: {output_dir}")
        print(f"Version: {args.version}")
        print(f"Config: {config_path}")
        print(f"Json: {json_path}")
        print(f"VersionFile: {version_path}")
        print(f"TableCount: {len(tables)}")

        if not args.no_dialog:
            exporter._show_result_dialog(
                "导表成功",
                f"共导出 {len(tables)} 个配置表\n输出目录：{output_dir}\n版本：{args.version}",
                success=True,
                auto_close_ms=3000,
            )
        return 0
    except Exception as exc:
        traceback.print_exc()
        error_message = exporter._format_export_error(exc)
        if not args.no_dialog:
            exporter._show_result_dialog("导表失败", error_message, success=False)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
