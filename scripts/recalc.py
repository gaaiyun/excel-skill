#!/usr/bin/env python3
"""
recalc.py -- Recalculate formulas in openpyxl-generated .xlsx files.

openpyxl 写出来的公式只是字符串，.xlsx 里没有缓存的计算结果（用
`data_only=True` 读会拿到 None）。本脚本用 LibreOffice headless 把文件
打开 → 重算 → 另存，从而把缓存值写回，使下游 `data_only=True` 读取或
非 Excel 工具能拿到数值。

依赖：需要本机安装 LibreOffice（提供 soffice / libreoffice 命令）。没装时
脚本会明确报错并提示——直接用 Excel 打开文件也会自动重算。

用法：
    python scripts/recalc.py input.xlsx                 # 原地重算
    python scripts/recalc.py input.xlsx -o out.xlsx     # 重算到新文件
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def find_libreoffice() -> str | None:
    """跨平台定位 LibreOffice 可执行文件。"""
    # 1. PATH
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path

    # 2. 常见安装路径（PATH 里没有时兜底）
    candidates = [
        # Windows
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        # macOS
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        # Linux
        "/usr/bin/soffice",
        "/usr/bin/libreoffice",
        "/snap/bin/libreoffice",
    ]
    for c in candidates:
        if Path(c).exists():
            return c
    return None


def recalc_with_libreoffice(input_path: Path, output_path: Path) -> int:
    """用 LibreOffice headless 打开 → 重算 → 另存为 xlsx。

    思路：``--convert-to xlsx:"Calc MS Excel 2007 XML"`` 在转换时会触发
    一次完整重算，把缓存值写回。LibreOffice 默认 ``AutoCalculate`` 开启，
    convert 过程即重算（等价于打开文件时 Excel 的自动重算）。
    """
    lo_path = find_libreoffice()
    if lo_path is None:
        print(
            "[error] 未找到 LibreOffice，无法重算公式。\n"
            "        安装：https://www.libreoffice.org/download/\n"
            "        替代：直接用 Excel 打开该文件（打开时会自动重算）。",
            file=sys.stderr,
        )
        return 1

    # LibreOffice 的 --convert-to 只能把结果写进 --outdir，且文件名与输入同名。
    # 为避免污染、并支持原地覆盖，先转换到临时目录再搬到目标路径。
    with tempfile.TemporaryDirectory() as tmpdir:
        cmd = [
            lo_path,
            "--headless",
            "--calc",
            "--convert-to",
            "xlsx:Calc MS Excel 2007 XML",
            "--outdir",
            tmpdir,
            str(input_path.resolve()),
        ]
        try:
            proc = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=180,
            )
        except subprocess.TimeoutExpired:
            print("[error] LibreOffice 重算超时（>180s）。", file=sys.stderr)
            return 3
        except OSError as e:
            print(f"[error] 调用 LibreOffice 失败：{e}", file=sys.stderr)
            return 3

        produced = Path(tmpdir) / (input_path.stem + ".xlsx")
        if proc.returncode != 0 or not produced.exists():
            print(
                f"[error] LibreOffice 转换失败（returncode={proc.returncode}）。\n"
                f"        stderr: {proc.stderr.strip()[:500]}",
                file=sys.stderr,
            )
            return 3

        output_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(produced), str(output_path))

    print(f"[OK] 已重算并写入：{output_path}")
    return 0


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="用 LibreOffice 重算 openpyxl 生成的 .xlsx 公式缓存值"
    )
    parser.add_argument("input", type=Path, help="输入 .xlsx 文件")
    parser.add_argument("--output", "-o", type=Path, default=None,
                        help="输出路径（默认原地覆盖输入文件）")
    args = parser.parse_args(argv)

    if not args.input.exists():
        print(f"[error] 文件不存在：{args.input}", file=sys.stderr)
        return 1

    output = args.output or args.input
    return recalc_with_libreoffice(args.input, output)


if __name__ == "__main__":
    sys.exit(main())
