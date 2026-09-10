from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Iterable, List

EXCEL_SUFFIXES = {'.xls', '.xlsx', '.xlsm'}


def ensure_project_dir(value: str | Path) -> Path:
    p = Path(value).expanduser().resolve()
    if not p.is_dir():
        raise ValueError(f'项目文件夹不存在：{p}')
    return p


def project_output_dir(project_dir: Path) -> Path:
    return project_dir / '输出'


def project_log_dir(project_dir: Path) -> Path:
    return project_output_dir(project_dir) / 'logs'


def iter_project_excel_files(project_dir: Path) -> List[Path]:
    """扫描项目原始Excel，排除输出目录、物料导入表和明显的无意义导入文件。"""
    project_dir = project_dir.resolve()
    out_dir = (project_dir / '输出').resolve()
    result: List[Path] = []
    for p in project_dir.rglob('*'):
        if not p.is_file() or p.name.startswith('~$'):
            continue
        if p.suffix.lower() not in EXCEL_SUFFIXES:
            continue
        try:
            p.resolve().relative_to(out_dir)
            continue
        except ValueError:
            pass
        name = p.name
        if name.startswith('物料导入') and p.suffix.lower() == '.xls':
            continue
        if name.startswith('01-') and '嘉盛-导入' in name:
            continue
        result.append(p)
    return sorted(result)


def write_result_json(path: str | Path | None, result: dict) -> None:
    if not path:
        return
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')


def add_project_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument('--project', required=True, help='项目文件夹路径')
    parser.add_argument('--result-json', default='', help='可选：结构化结果JSON输出路径')
