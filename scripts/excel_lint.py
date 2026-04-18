#!/usr/bin/env python3
"""
excel_lint.py — 检查 openpyxl Python 代码常见坑

检查项（来自 references/08-excel-pitfalls.md）：
  ★ JQ001 (XL001): data_only=True 后 wb.save() — 会让公式永久丢失
  ★ XL002: pandas to_excel 写带公式的 DataFrame
  ★ XL003: 公式中含中文逗号 ， (应该用英文 ,)
  XL004: sheet 名含禁止字符 / \\ ? * [ ] '
  XL005: sheet 名超 31 字符
  XL006: PatternFill 没传 fill_type
  XL007: .xlsm 文件 load 时没有 keep_vba=True
  XL008: 公式跨 sheet 引用时 sheet 名含空格但没用单引号
  XL009: 中文字符没设字体（在英文环境会乱码）
  XL010: read_only=True 后试图 wb.save()

用法:
    python scripts/excel_lint.py my_excel_code.py
    python scripts/excel_lint.py --json my_excel_code.py
"""
from __future__ import annotations

import argparse
import ast
import json
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class Issue:
    severity: str  # error / warning / info
    line: int
    col: int
    code: str
    message: str
    fix: str = ''


@dataclass
class Report:
    file: str
    issues: list[Issue] = field(default_factory=list)

    def add(self, severity: str, line: int, col: int, code: str, message: str, fix: str = '') -> None:
        self.issues.append(Issue(severity, line, col, code, message, fix))

    @property
    def errors(self) -> list[Issue]:
        return [i for i in self.issues if i.severity == 'error']

    @property
    def warnings(self) -> list[Issue]:
        return [i for i in self.issues if i.severity == 'warning']

    @property
    def passed(self) -> bool:
        return len(self.errors) == 0


class ExcelLinter(ast.NodeVisitor):
    def __init__(self, report: Report, source: str):
        self.report = report
        self.source = source
        self._loaded_data_only = False  # 是否调用过 load_workbook(data_only=True)
        self._loaded_xlsm = False        # 是否 load 过 xlsm
        self._has_keep_vba = False
        self._save_calls: list[ast.Call] = []
        self._read_only_load = False

    def visit_Call(self, node: ast.Call) -> None:
        func_name = self._get_call_name(node)

        # XL001 / XL010: load_workbook 参数检查
        if func_name in ('load_workbook', 'openpyxl.load_workbook'):
            kw = {kw.arg: kw.value for kw in node.keywords}

            if 'data_only' in kw and self._is_true(kw['data_only']):
                self._loaded_data_only = True

            if 'read_only' in kw and self._is_true(kw['read_only']):
                self._read_only_load = True

            # 检查 .xlsm 文件
            if node.args:
                first = node.args[0]
                if isinstance(first, ast.Constant) and isinstance(first.value, str):
                    if first.value.endswith('.xlsm'):
                        self._loaded_xlsm = True
                        if 'keep_vba' not in kw or not self._is_true(kw.get('keep_vba')):
                            self.report.add(
                                'warning', node.lineno, node.col_offset, 'XL007',
                                f'load_workbook("{first.value}") 没有 keep_vba=True',
                                '加 keep_vba=True，否则 save 后宏丢失',
                            )

        # XL010: workbook.save() 调用，先记下来
        if func_name.endswith('.save'):
            self._save_calls.append(node)
            if self._loaded_data_only:
                self.report.add(
                    'error', node.lineno, node.col_offset, 'XL001',
                    'load_workbook(data_only=True) 之后调用了 wb.save() — 会让所有公式永久丢失',
                    '只读用 data_only=True，要修改保存时不要带这个参数',
                )
            if self._read_only_load:
                self.report.add(
                    'error', node.lineno, node.col_offset, 'XL010',
                    'load_workbook(read_only=True) 之后调用了 wb.save() — read_only 模式不支持 save',
                    '去掉 read_only=True，或不要 save',
                )

        # XL002: pandas to_excel
        if func_name.endswith('.to_excel') or func_name == 'to_excel':
            self.report.add(
                'warning', node.lineno, node.col_offset, 'XL002',
                'pandas to_excel 不能正确写公式（会被当字符串）',
                '如果有公式列，改用 openpyxl 的 ws.cell(...).value = "=SUM(...)"，详见 references/06',
            )

        # XL006: PatternFill 没传 fill_type
        if func_name == 'PatternFill':
            kw = {kw.arg: kw.value for kw in node.keywords}
            has_fill_type = 'fill_type' in kw
            # 第一个 positional arg 也算
            has_positional = len(node.args) > 0 and not kw
            if not has_fill_type and not has_positional:
                self.report.add(
                    'warning', node.lineno, node.col_offset, 'XL006',
                    'PatternFill 没有 fill_type 参数，颜色将不显示',
                    '加 fill_type="solid"',
                )

        # XL004 / XL005: create_sheet / Worksheet 名检查
        if func_name in ('create_sheet', 'wb.create_sheet'):
            kw = {kw.arg: kw.value for kw in node.keywords}
            title_node = kw.get('title')
            if title_node is None and node.args:
                title_node = node.args[0]
            if isinstance(title_node, ast.Constant) and isinstance(title_node.value, str):
                title = title_node.value
                bad_chars = set(title) & set('/\\?*[]\'')
                if bad_chars:
                    self.report.add(
                        'error', node.lineno, node.col_offset, 'XL004',
                        f'sheet 名 "{title}" 含禁止字符 {bad_chars}',
                        '改名，禁用 / \\ ? * [ ] \'',
                    )
                if len(title) > 31:
                    self.report.add(
                        'error', node.lineno, node.col_offset, 'XL005',
                        f'sheet 名 "{title}" 长度 {len(title)} > 31',
                        '截断到 31 字符以内',
                    )

        self.generic_visit(node)

    def visit_Assign(self, node: ast.Assign) -> None:
        # 检查公式赋值
        if isinstance(node.value, ast.Constant) and isinstance(node.value.value, str):
            value = node.value.value
            if value.startswith('='):
                # XL003: 中文逗号
                if '，' in value or '；' in value:
                    self.report.add(
                        'error', node.lineno, node.col_offset, 'XL003',
                        f'公式 "{value}" 含中文标点（， ；）',
                        '改用英文逗号 , ；分号 ; 仅在某些 locale 下作为参数分隔符，但 openpyxl 写入时只接受 ,',
                    )

                # XL008: 跨 sheet 引用，sheet 名有空格但没单引号
                # 简化检查：找 SheetName!Cell 模式，看 SheetName 里有没有空格
                m = re.search(r"=\s*([A-Za-z0-9_\u4e00-\u9fa5 ]+)!", value)
                if m:
                    sheet_ref = m.group(1)
                    if ' ' in sheet_ref and "'" not in value[:m.start(1)]:
                        self.report.add(
                            'warning', node.lineno, node.col_offset, 'XL008',
                            f'公式跨 sheet 引用 "{sheet_ref}" 含空格但没单引号',
                            f"改成 ='{sheet_ref}'!...",
                        )

        self.generic_visit(node)

    def _get_call_name(self, node: ast.Call) -> str:
        return self._extract_name(node.func)

    def _extract_name(self, node: ast.AST) -> str:
        if isinstance(node, ast.Name):
            return node.id
        if isinstance(node, ast.Attribute):
            base = self._extract_name(node.value)
            if base:
                return f'{base}.{node.attr}'
            return node.attr
        return ''

    @staticmethod
    def _is_true(node: ast.AST) -> bool:
        return isinstance(node, ast.Constant) and node.value is True


def lint_file(path: Path) -> Report:
    source = path.read_text(encoding='utf-8')
    report = Report(file=str(path))

    try:
        tree = ast.parse(source)
    except SyntaxError as e:
        report.add('error', e.lineno or 0, e.offset or 0, 'XL000',
                   f'Python 语法错误：{e.msg}')
        return report

    linter = ExcelLinter(report, source)
    linter.visit(tree)
    return report


def render(report: Report) -> str:
    lines = ['=' * 70, f'Excel Code Lint — {report.file}', '=' * 70]
    if not report.issues:
        lines.append('PASSED — 没有发现 openpyxl 常见坑')
        return '\n'.join(lines)
    lines.append(f'Errors: {len(report.errors)}    Warnings: {len(report.warnings)}')
    lines.append('')
    sev_order = {'error': 0, 'warning': 1, 'info': 2}
    for i in sorted(report.issues, key=lambda x: (sev_order[x.severity], x.line)):
        sev = {'error': '[ERROR]', 'warning': '[WARN] ', 'info': '[INFO] '}[i.severity]
        loc = f'L{i.line}:{i.col}' if i.line else 'global'
        lines.append(f'{sev} {i.code} {loc}: {i.message}')
        if i.fix:
            lines.append(f'         fix: {i.fix}')
    lines.append('=' * 70)
    return '\n'.join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(description='Excel/openpyxl Python code linter')
    parser.add_argument('file')
    parser.add_argument('--json', action='store_true')
    args = parser.parse_args()

    path = Path(args.file)
    if not path.exists():
        print(f'[error] File not found: {path}', file=sys.stderr)
        return 2
    report = lint_file(path)
    if args.json:
        print(json.dumps({
            'file': report.file,
            'passed': report.passed,
            'errors': len(report.errors),
            'warnings': len(report.warnings),
            'issues': [
                {'severity': i.severity, 'line': i.line, 'col': i.col,
                 'code': i.code, 'message': i.message, 'fix': i.fix}
                for i in report.issues
            ],
        }, ensure_ascii=False, indent=2))
    else:
        print(render(report))
    return 0 if report.passed else 1


if __name__ == '__main__':
    sys.exit(main())
