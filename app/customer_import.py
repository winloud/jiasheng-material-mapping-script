import shutil
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from typing import Any, List, Dict, Tuple

# ============================================================
# 配置区 —— 日常使用主要修改这里
# 高速版：一个 Excel 只打开一次，并使用流式行读取。
# 当前规则：D列为空结束；B列为空时仍提取，并按D列与已有数据去重。
# 目录规则：仅扫描“待处理”；数据、备份和日志均按项目目录分类存放。
# 安全规则：写入已有汇总表前自动备份；可配置按D匹配无料号历史行并只回填B列。
# ============================================================

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = PROJECT_ROOT / "data"
MASTER_FILE = DATA_DIR / "嘉盛物料映射汇总表.xlsx"
CUSTOMER_DIR = DATA_DIR / "客户清单"
CUSTOMER_PENDING = CUSTOMER_DIR / "待处理"
CUSTOMER_DONE = CUSTOMER_DIR / "已处理"
CUSTOMER_ERROR = CUSTOMER_DIR / "异常"
BACKUP_DIR = PROJECT_ROOT / "backup"
LOG_DIR = PROJECT_ROOT / "logs"

CONFIG = {
    # 工作目录：当前脚本所在目录
    "work_dir": PROJECT_ROOT,

    # 输入目录：只处理“待处理”文件夹中的 Excel
    "input_dir": CUSTOMER_PENDING,

    # 输出文件名
    "output_file": MASTER_FILE,

    # 异常文件移动目录（读取失败的文件会移动到这里）
    "error_dir": CUSTOMER_ERROR,

    # 正常完成文件移动目录（处理成功的文件会移动到这里）
    "done_dir": CUSTOMER_DONE,

    # 日志文件名
    "log_file": LOG_DIR / "customer_import.log",

    # 每次修改已有汇总表前，是否自动备份
    "backup_before_write": True,

    # 备份目录（位于脚本同级目录）
    "backup_dir": BACKUP_DIR,

    # 当新数据 B 列已有物料号、但该物料号在汇总表中不存在时：
    # 是否继续按 D 列查找“B列为空”的历史记录。
    # True：若 D 唯一匹配，则只回填历史行的 B 列，不追加新行，
    #       M~X 等人工维护的映射列完全不改。
    # False：保持旧逻辑，B物料号不存在时直接按新行追加。
    "backfill_material_no_by_d": True,

    # 并行读取文件数（1 为串行，建议 4~8）
    "max_workers": 8,

    # 是否在最终结果后面增加来源信息
    "add_source_info": True,

    # 主去重依据：
    # 最终统一结构中的第几列，从 1 开始
    # 当前第 2 列为物料号；物料号存在时优先按此列去重
    "dedup_col": 2,

    # 备用去重依据：
    # 当主去重列（B列物料号）为空时，改按最终统一结构的 D 列去重
    # 同时会把已有数据中所有非空 D 值建立索引，
    # 因此“新物料无物料号”时，可判断 D 是否已在历史数据中出现
    "fallback_dedup_col": 4,

    # 数量所在的目标列（最终结构中的第几列，从 1 开始）
    # G 列 = 第 7 列
    # 当该列写入的值不是数字时，自动从原表头查找列名为"数量"或"用量"的列进行替换
    "quantity_col": 7,

    # 物料映射列配置
    # 从哪一列开始放物料映射（物料A、物料A数量、物料B、物料B数量……）
    "mapping_start_col": "M",

    # 最多放几种物料映射
    "mapping_count": 6,

    # 每个 Sheet 独立配置
    "sheets": [
        {
            "name": "电柜元器件",

            # 从 Excel 第几行开始读取
            "start_row": 4,

            # 原始读取范围
            "start_col": "A",
            "end_col": "F",

            # 以哪一列为空作为数据结束
            "stop_col": "D",

            # 字段映射：
            # 最终统一结构一共 7 列
            #
            # 原始：
            # A B C D E F
            #
            # 最终：
            # A B C D E 空 F
            #
            # 数字表示原始区域中的第几列
            # None 表示插入空值
            "column_mapping": [
                1,  # A
                2,  # B
                3,  # C
                4,  # D
                5,  # E
                None,
                6,  # 原 F -> 最终 G
            ],
        },

        {
            "name": "电柜半成品",

            "start_row": 3,
            "start_col": "A",
            "end_col": "G",
            "stop_col": "D",

            # A:G 原样映射到最终 7 列
            "column_mapping": [
                1,
                2,
                3,
                4,
                5,
                6,
                7,
            ],
        },
    ],
}


# ============================================================
# 日志
# ============================================================

_log_lines: List[str] = []


def log(msg: str = ""):
    """同时输出到控制台和日志缓冲区。"""
    print(msg)
    _log_lines.append(msg)


def save_log(log_path: Path):
    log_path.write_text(
        "\n".join(_log_lines),
        encoding="utf-8",
    )


# ============================================================
# 工具函数
# ============================================================

def excel_col_to_num(col: str) -> int:
    """
    Excel 列字母转数字。
    A -> 1
    B -> 2
    Z -> 26
    AA -> 27
    """
    col = col.upper().strip()

    result = 0

    for char in col:
        result = result * 26 + ord(char) - ord("A") + 1

    return result


def is_empty(value: Any) -> bool:
    """
    判断单元格是否为空。
    """
    if value is None:
        return True

    if isinstance(value, str) and value.strip() == "":
        return True

    return False


def is_number(value: Any) -> bool:
    """
    判断值是否为数字。
    """
    if isinstance(value, (int, float)):
        return True

    if isinstance(value, str):
        s = value.strip()
        if s:
            try:
                float(s)
                return True
            except ValueError:
                return False

    return False


def normalize_key(value: Any) -> str:
    """
    将物料编码标准化成字符串用于去重。

    避免 Excel 中：
    123
    123.0
    "123"
    因类型不同而导致误判。
    """
    if value is None:
        return ""

    # Excel 数字经常会读成浮点数
    if isinstance(value, float) and value.is_integer():
        return str(int(value))

    return str(value).strip()


def apply_mapping(
    source_row: List[Any],
    mapping: List[Any],
) -> List[Any]:
    """
    根据配置将原始数据转换成统一结构。

    mapping 示例：
    [1, 2, 3, 4, 5, None, 6]

    表示：

    输出1 <- 原始1
    输出2 <- 原始2
    ...
    输出6 <- 空
    输出7 <- 原始6
    """

    result = []

    for source_index in mapping:

        if source_index is None:
            result.append(None)
            continue

        index = source_index - 1

        if index < len(source_row):
            result.append(source_row[index])
        else:
            result.append(None)

    return result


def check_header_b(b_value: Any) -> bool:
    """
    检查 B 列表头是否为物料号/料号。
    """
    if is_empty(b_value):
        return False
    s = str(b_value).strip()
    return "料号" in s


def mapping_label(i: int) -> str:
    """
    生成物料映射标签：A, B, ..., Z, AA, AB, ...
    """
    result = ""
    n = i + 1
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(ord("A") + rem) + result
    return result


# ============================================================
# Excel 高速读取
# ============================================================

def _read_xlsx_sheet_from_workbook(
    workbook,
    sheet_config: Dict[str, Any],
) -> List[List[Any]]:
    """
    从已经打开的 xlsx/xlsm 工作簿读取一个 Sheet。

    性能优化：
    1. 同一个 Excel 文件只打开一次；
    2. 使用 iter_rows(values_only=True) 流式读取，避免逐单元格 ws.cell()；
    3. 一次读取到本 Sheet 实际需要的最右列（含数量备用列搜索范围）。
    """
    sheet_name = sheet_config["name"]

    if sheet_name not in workbook.sheetnames:
        log(f"    [跳过] Sheet 不存在：{sheet_name}")
        return []

    ws = workbook[sheet_name]

    start_row = sheet_config["start_row"]
    header_row = start_row - 1

    start_col = excel_col_to_num(sheet_config["start_col"])
    end_col = excel_col_to_num(sheet_config["end_col"])
    stop_col = excel_col_to_num(sheet_config["stop_col"])

    # 一次读取表头前 30 列，用于检查 B 列及查找“数量/用量”
    header_values = next(
        ws.iter_rows(
            min_row=header_row,
            max_row=header_row,
            min_col=1,
            max_col=30,
            values_only=True,
        ),
        (),
    )

    b_header = header_values[1] if len(header_values) >= 2 else None
    if not check_header_b(b_header):
        raise ValueError(
            f"B 列表头不是物料号/料号，实际为：{b_header}"
        )

    quantity_col = None
    for idx, value in enumerate(header_values, start=1):
        if value and any(
            k in str(value).strip()
            for k in ("数量", "用量")
        ):
            quantity_col = idx
            break

    qty_index = CONFIG["quantity_col"] - 1
    rows: List[List[Any]] = []

    # 只流式读取真正需要的列：
    # 原始范围、停止列、以及可能作为数量备用来源的列。
    max_needed_col = max(
        end_col,
        stop_col,
        quantity_col or 0,
    )

    for values in ws.iter_rows(
        min_row=start_row,
        min_col=1,
        max_col=max_needed_col,
        values_only=True,
    ):
        stop_value = (
            values[stop_col - 1]
            if stop_col - 1 < len(values)
            else None
        )

        if is_empty(stop_value):
            break

        source_row = list(values[start_col - 1:end_col])

        mapped_row = apply_mapping(
            source_row,
            sheet_config["column_mapping"],
        )

        if (
            quantity_col is not None
            and not is_number(mapped_row[qty_index])
        ):
            mapped_row[qty_index] = (
                values[quantity_col - 1]
                if quantity_col - 1 < len(values)
                else None
            )

        rows.append(mapped_row)

    return rows


def _read_xls_sheet_from_workbook(
    workbook,
    sheet_config: Dict[str, Any],
) -> List[List[Any]]:
    """
    从已经打开的 .xls 工作簿读取一个 Sheet。
    同一个 .xls 文件也只打开一次。
    """
    import xlrd

    sheet_name = sheet_config["name"]

    try:
        ws = workbook.sheet_by_name(sheet_name)
    except xlrd.biffh.XLRDError:
        log(f"    [跳过] Sheet 不存在：{sheet_name}")
        return []

    start_row = sheet_config["start_row"] - 1
    start_col = excel_col_to_num(sheet_config["start_col"]) - 1
    end_col = excel_col_to_num(sheet_config["end_col"]) - 1
    stop_col = excel_col_to_num(sheet_config["stop_col"]) - 1

    header_row_0 = start_row - 1
    b_header = ws.cell_value(header_row_0, 1)
    if not check_header_b(b_header):
        raise ValueError(
            f"B 列表头不是物料号/料号，实际为：{b_header}"
        )

    quantity_col = None
    for col in range(min(ws.ncols, 30)):
        value = ws.cell_value(header_row_0, col)
        if value and any(
            k in str(value).strip()
            for k in ("数量", "用量")
        ):
            quantity_col = col
            break

    rows: List[List[Any]] = []
    qty_index = CONFIG["quantity_col"] - 1

    for current_row in range(start_row, ws.nrows):
        stop_value = (
            ws.cell_value(current_row, stop_col)
            if stop_col < ws.ncols
            else None
        )

        if is_empty(stop_value):
            break

        source_row = [
            ws.cell_value(current_row, col)
            if col < ws.ncols
            else None
            for col in range(start_col, end_col + 1)
        ]

        mapped_row = apply_mapping(
            source_row,
            sheet_config["column_mapping"],
        )

        if (
            quantity_col is not None
            and not is_number(mapped_row[qty_index])
            and quantity_col < ws.ncols
        ):
            mapped_row[qty_index] = ws.cell_value(
                current_row,
                quantity_col,
            )

        rows.append(mapped_row)

    return rows


# ============================================================
# 输出 Excel
# ============================================================

def build_header() -> List[str]:
    """
    构建标题行。
    A~I: 序号、物料号、品名、型号、品牌、规格、数量、文件名、sheet名
    J: 处理时间
    K~(mapping_start_col-1): 空
    mapping_start_col 起: 物料A、物料A数量、物料B、物料B数量……
    """
    header = [
        "序号", "物料号", "品名", "型号", "品牌",
        "规格", "数量", "文件名", "sheet名", "处理时间",
    ]

    # 填充空白列到 mapping_start_col 前一列
    start_num = excel_col_to_num(CONFIG["mapping_start_col"])
    while len(header) < start_num - 1:
        header.append("")

    # 物料映射列
    for i in range(CONFIG["mapping_count"]):
        label = mapping_label(i)
        header.append(f"物料{label}")
        header.append(f"物料{label}数量")

    return header


def read_existing_output(
    output_path: Path,
) -> Tuple[set, set, Dict[str, List[int]], int]:
    """
    读取已有汇总表，建立去重与回填索引。

    返回：
    - existing_primary_keys：已有非空 B 列物料号集合
    - existing_fallback_keys：已有所有非空 D 列值集合
    - blank_primary_rows_by_d：B 为空的历史行，按 D 建立 {D: [Excel行号...]} 索引
    - max_seq：最大序号

    规则：
    1. 新数据 B 有物料号时，优先按 B 去重；
    2. 新数据 B 为空时，按 D 去重；
    3. 开启 backfill_material_no_by_d 后：
       B 有物料号但按 B 找不到时，可按 D 查找历史“B为空”记录，
       唯一匹配时仅补写该历史行的 B 列。
    """
    from openpyxl import load_workbook

    existing_primary_keys = set()
    existing_fallback_keys = set()
    blank_primary_rows_by_d: Dict[str, List[int]] = {}
    max_seq = 0

    if not output_path.exists():
        return (
            existing_primary_keys,
            existing_fallback_keys,
            blank_primary_rows_by_d,
            max_seq,
        )

    try:
        wb = load_workbook(
            output_path,
            read_only=True,
            data_only=True,
        )
        ws = wb.active

        primary_index = CONFIG["dedup_col"] - 1
        fallback_index = CONFIG["fallback_dedup_col"] - 1

        for excel_row_no, row in enumerate(
            ws.iter_rows(
                min_row=2,
                values_only=True,
            ),
            start=2,
        ):
            if not row:
                continue

            if row[0] is not None:
                try:
                    seq = int(row[0])
                    if seq > max_seq:
                        max_seq = seq
                except (ValueError, TypeError):
                    pass

            primary_key = ""
            if primary_index < len(row):
                primary_key = normalize_key(row[primary_index])
                if primary_key:
                    existing_primary_keys.add(primary_key)

            fallback_key = ""
            if fallback_index < len(row):
                fallback_key = normalize_key(row[fallback_index])
                if fallback_key:
                    existing_fallback_keys.add(fallback_key)

            # 只记录“B为空 + D非空”的历史行，供后续安全回填 B。
            if not primary_key and fallback_key:
                blank_primary_rows_by_d.setdefault(
                    fallback_key,
                    [],
                ).append(excel_row_no)

        wb.close()
    except Exception:
        pass

    return (
        existing_primary_keys,
        existing_fallback_keys,
        blank_primary_rows_by_d,
        max_seq,
    )


def write_output(
    output_path: Path,
    rows: List[List[Any]],
    header: List[str],
):

    from openpyxl import Workbook

    wb = Workbook()

    ws = wb.active
    ws.title = "汇总"

    ws.append(header)

    for row in rows:
        padded = list(row) + [None] * (
            len(header) - len(row)
        )
        ws.append(padded)

    wb.save(output_path)


def backup_existing_output(
    output_path: Path,
    backup_dir: Path,
) -> Path | None:
    """
    修改已有汇总表前创建完整备份。
    若汇总表尚不存在，则无需备份。
    """
    if not output_path.exists():
        return None

    backup_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = (
        backup_dir
        / f"{output_path.stem}_{timestamp}{output_path.suffix}"
    )

    # 极少数情况下同一秒运行多次，避免覆盖前一份备份。
    index = 1
    while backup_path.exists():
        backup_path = (
            backup_dir
            / f"{output_path.stem}_{timestamp}_{index}"
            f"{output_path.suffix}"
        )
        index += 1

    shutil.copy2(output_path, backup_path)
    return backup_path


def append_output(
    output_path: Path,
    rows: List[List[Any]],
    header: List[str],
    backfill_updates: Dict[int, Any] | None = None,
):
    """
    将本次变化一次性保存到汇总表。

    backfill_updates:
        {Excel行号: 新物料号}

    回填时只修改 CONFIG["dedup_col"] 对应的 B 列，
    其它所有列（尤其 M~X 人工映射关系）保持原值。
    """
    from openpyxl import load_workbook

    backfill_updates = backfill_updates or {}

    if not output_path.exists():
        # 新建文件时不存在历史行，因此正常情况下不会有回填任务。
        write_output(output_path, rows, header)
        return

    wb = load_workbook(output_path)
    ws = wb.active

    primary_col = CONFIG["dedup_col"]

    # 先只回填 B 物料号，不触碰其它任何列。
    for excel_row_no, material_no in backfill_updates.items():
        ws.cell(
            row=excel_row_no,
            column=primary_col,
        ).value = material_no

    # 再追加真正的新行。
    for row in rows:
        padded = list(row) + [None] * (
            len(header) - len(row)
        )
        ws.append(padded)

    wb.save(output_path)


# ============================================================
# 单文件处理（供并行调用）
# ============================================================

def process_one_file(
    file_path: Path,
) -> Tuple[Path, bool, List[Dict[str, Any]], str]:
    """
    处理单个 Excel 文件，返回：
    (file_path, file_ok, sheet_results, error_msg)

    高速版关键点：
    - 一个文件只打开一次；
    - 在同一个已打开工作簿中连续读取所有配置 Sheet；
    - 文件级并行仍由 ThreadPoolExecutor 负责。
    """
    file_ok = True
    sheet_results: List[Dict[str, Any]] = []
    errors: List[str] = []

    suffix = file_path.suffix.lower()

    try:
        if suffix in (".xlsx", ".xlsm"):
            from openpyxl import load_workbook

            workbook = load_workbook(
                file_path,
                read_only=True,
                data_only=True,
            )

            try:
                for sheet_config in CONFIG["sheets"]:
                    sheet_name = sheet_config["name"]

                    try:
                        rows = _read_xlsx_sheet_from_workbook(
                            workbook,
                            sheet_config,
                        )
                        if rows:
                            sheet_results.append({
                                "sheet_name": sheet_name,
                                "rows": rows,
                            })
                    except Exception as exc:
                        file_ok = False
                        errors.append(f"{sheet_name}: {exc}")
            finally:
                workbook.close()

        elif suffix == ".xls":
            try:
                import xlrd
            except ImportError:
                raise RuntimeError(
                    "检测到 .xls 文件，但没有安装 xlrd。\n"
                    "请执行：pip install xlrd"
                )

            workbook = xlrd.open_workbook(
                file_path,
                on_demand=True,
            )

            try:
                for sheet_config in CONFIG["sheets"]:
                    sheet_name = sheet_config["name"]

                    try:
                        rows = _read_xls_sheet_from_workbook(
                            workbook,
                            sheet_config,
                        )
                        if rows:
                            sheet_results.append({
                                "sheet_name": sheet_name,
                                "rows": rows,
                            })
                    except Exception as exc:
                        file_ok = False
                        errors.append(f"{sheet_name}: {exc}")
            finally:
                workbook.release_resources()

        else:
            raise ValueError(
                f"不支持的 Excel 类型：{suffix}"
            )

    except Exception as exc:
        file_ok = False
        errors.append(str(exc))

    return (
        file_path,
        file_ok,
        sheet_results,
        " | ".join(errors),
    )


# ============================================================
# 主程序
# ============================================================

def run_project(project_dir: Path) -> Dict[str, Any]:
    """按项目目录执行脚本1。原始项目文件只读，不移动。"""
    from project_utils import iter_project_excel_files, project_log_dir

    _log_lines.clear()
    project_dir = Path(project_dir).resolve()
    output_path = MASTER_FILE
    log_path = project_log_dir(project_dir) / "01_导入客户物料清单.log"
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    log_path.parent.mkdir(parents=True, exist_ok=True)

    batch_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log("=" * 60)
    log("脚本1：导入客户物料清单")
    log(f"处理时间：{batch_time}")
    log(f"项目目录：{project_dir}")
    log("原始项目文件只读，不移动。")
    log("=" * 60)

    excel_files = iter_project_excel_files(project_dir)
    log(f"\n发现候选 Excel 文件：{len(excel_files)} 个\n")
    if not excel_files:
        result = {
            "success": False, "can_continue": False, "new_materials": 0,
            "reason": "项目目录中没有找到可处理的电气清单 Excel 文件。",
            "files": 0,
        }
        log(f"[错误] {result['reason']}")
        save_log(log_path)
        return result

    try:
        seen_primary_keys, seen_fallback_keys, blank_primary_rows_by_d, next_seq = read_existing_output(output_path)
    except Exception as exc:
        result = {"success": False, "can_continue": False, "new_materials": 0, "reason": f"读取主映射表失败：{exc}"}
        log(f"[错误] {result['reason']}")
        save_log(log_path)
        return result

    file_success = file_failed = 0
    raw_count = duplicate_count = backfill_count = backfill_ambiguous_count = 0
    result_rows = []
    backfill_updates: Dict[int, Any] = {}

    log(f"并行读取：{CONFIG['max_workers']} 线程\n")
    all_results: List[Tuple[Path, bool, List[Dict[str, Any]], str]] = []
    with ThreadPoolExecutor(max_workers=CONFIG["max_workers"]) as executor:
        future_map = {executor.submit(process_one_file, fp): fp for fp in excel_files}
        done_count = 0
        for future in as_completed(future_map):
            fp = future_map[future]
            done_count += 1
            try:
                result = future.result()
            except Exception as exc:
                result = (fp, False, [], str(exc))
            all_results.append(result)
            _, ok, sheets, err = result
            log(f"[{done_count}/{len(excel_files)}] {'OK' if ok else 'FAIL'} {fp.relative_to(project_dir)}")
            if ok:
                for sr in sheets:
                    log(f"    {sr['sheet_name']}: {len(sr['rows'])} 行")
            else:
                log(f"    [错误] {err}")

    order = {fp: i for i, fp in enumerate(excel_files)}
    all_results.sort(key=lambda r: order[r[0]])

    for file_path, file_ok, sheet_results, error_msg in all_results:
        if not file_ok:
            file_failed += 1
            continue
        file_success += 1
        for sr in sheet_results:
            sheet_name = sr["sheet_name"]
            for row in sr["rows"]:
                raw_count += 1
                primary_index = CONFIG["dedup_col"] - 1
                fallback_index = CONFIG["fallback_dedup_col"] - 1
                primary_key = normalize_key(row[primary_index])
                fallback_key = normalize_key(row[fallback_index])

                if primary_key:
                    if primary_key in seen_primary_keys:
                        duplicate_count += 1
                        continue
                    did_backfill = False
                    if CONFIG["backfill_material_no_by_d"] and fallback_key:
                        matched_rows = blank_primary_rows_by_d.get(fallback_key, [])
                        if len(matched_rows) == 1:
                            excel_row_no = matched_rows[0]
                            backfill_updates[excel_row_no] = row[primary_index]
                            backfill_count += 1
                            did_backfill = True
                            seen_primary_keys.add(primary_key)
                            blank_primary_rows_by_d.pop(fallback_key, None)
                            log(f"    [回填物料号] D={fallback_key} -> 汇总表第{excel_row_no}行 B={primary_key}")
                        elif len(matched_rows) > 1:
                            backfill_ambiguous_count += 1
                            log(f"    [警告] D={fallback_key} 匹配到 {len(matched_rows)} 条B为空历史记录；本条按新行追加。")
                    if did_backfill:
                        continue
                    seen_primary_keys.add(primary_key)
                    if fallback_key:
                        seen_fallback_keys.add(fallback_key)
                else:
                    if not fallback_key:
                        continue
                    if fallback_key in seen_fallback_keys:
                        duplicate_count += 1
                        continue
                    seen_fallback_keys.add(fallback_key)

                output_row = list(row)
                if CONFIG["add_source_info"]:
                    output_row.extend([file_path.name, sheet_name])
                output_row.append(batch_time)
                result_rows.append(output_row)

    write_error = ""
    if result_rows or backfill_updates:
        try:
            if CONFIG["backup_before_write"] and output_path.exists():
                backup_path = backup_existing_output(output_path, Path(CONFIG["backup_dir"]))
                if backup_path is not None:
                    log(f"[备份] 已创建：{backup_path}")
            append_output(output_path, result_rows, build_header(), backfill_updates=backfill_updates)
        except Exception as exc:
            write_error = str(exc)
            log(f"[错误] 写入主映射表失败：{exc}")

    log("\n" + "=" * 60)
    log(f"成功文件：{file_success}，异常文件：{file_failed}")
    log(f"原始物料：{raw_count}，重复：{duplicate_count}")
    log(f"回填历史物料号：{backfill_count}")
    log(f"本次新增物料：{len(result_rows)}")
    log(f"日志：{log_path}")
    save_log(log_path)

    if write_error:
        return {"success": False, "can_continue": False, "new_materials": len(result_rows), "reason": f"写入主映射表失败：{write_error}"}
    if file_failed:
        return {"success": False, "can_continue": False, "new_materials": len(result_rows), "reason": f"有 {file_failed} 个文件读取失败，请查看项目日志。", "files": len(excel_files)}

    new_count = len(result_rows)
    return {
        "success": True,
        "can_continue": new_count == 0,
        "new_materials": new_count,
        "backfilled_material_nos": backfill_count,
        "files": len(excel_files),
        "reason": "" if new_count == 0 else f"发现 {new_count} 条新物料，请先完成映射配置。",
        "log_file": str(log_path),
    }


def main() -> None:
    import argparse
    from project_utils import add_project_args, ensure_project_dir, write_result_json

    parser = argparse.ArgumentParser(description="脚本1：导入客户物料清单")
    add_project_args(parser)
    args = parser.parse_args()
    try:
        result = run_project(ensure_project_dir(args.project))
    except Exception as exc:
        result = {"success": False, "can_continue": False, "new_materials": 0, "reason": str(exc)}
        print(f"[错误] {exc}")
    write_result_json(args.result_json, result)
    raise SystemExit(0 if result.get("success") else 1)


if __name__ == "__main__":
    main()
