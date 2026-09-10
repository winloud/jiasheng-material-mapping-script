from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any, Dict

from customer_import import run_project as run_import
from material_fillback import run_project as run_fillback
from erp_bom_export import run_project as run_erp
from project_utils import project_log_dir


def run_project(project_dir: Path) -> Dict[str, Any]:
    project_dir = Path(project_dir).resolve()
    log_dir = project_log_dir(project_dir)
    log_dir.mkdir(parents=True, exist_ok=True)
    workflow_log = log_dir / "05_项目一键处理.log"
    lines: list[str] = []

    def say(text: str = "") -> None:
        print(text)
        lines.append(text)

    say("=" * 70)
    say("脚本5：项目一键处理")
    say(f"项目：{project_dir}")
    say(f"时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    say("流程：1 -> 无新物料 -> 3 -> 全匹配 -> 4")
    say("=" * 70)

    say("\n[1/3] 执行脚本1：导入客户物料清单")
    r1 = run_import(project_dir)
    if not r1.get("success"):
        reason = r1.get("reason") or "脚本1执行失败"
        say(f"[停止] {reason}")
        workflow_log.write_text("\n".join(lines), encoding="utf-8")
        return {"success": False, "stage": 1, "reason": reason, "log_file": str(workflow_log)}
    if not r1.get("can_continue"):
        reason = r1.get("reason") or f"发现 {r1.get('new_materials', 0)} 条新物料，请先完成映射配置。"
        say(f"[停止] {reason}")
        workflow_log.write_text("\n".join(lines), encoding="utf-8")
        return {"success": False, "stage": 1, "reason": reason, "new_materials": r1.get("new_materials", 0), "log_file": str(workflow_log)}
    say("[通过] 未发现新物料，可以继续。")

    say("\n[2/3] 执行脚本3：生成京能物料对应表")
    r3 = run_fillback(project_dir)
    if not r3.get("success"):
        reason = r3.get("reason") or "脚本3执行失败"
        say(f"[停止] {reason}")
        workflow_log.write_text("\n".join(lines), encoding="utf-8")
        return {"success": False, "stage": 3, "reason": reason, "log_file": str(workflow_log)}
    if not r3.get("can_continue"):
        reason = r3.get("reason") or f"物料回填存在 {r3.get('abnormal_rows', 0)} 条异常。"
        say(f"[停止] {reason}")
        workflow_log.write_text("\n".join(lines), encoding="utf-8")
        return {"success": False, "stage": 3, "reason": reason, "abnormal_rows": r3.get("abnormal_rows", 0), "log_file": str(workflow_log)}
    say("[通过] 物料回填全部匹配。")

    say("\n[3/3] 执行脚本4：生成 ERP 物料清单批量导入表")
    r4 = run_erp(project_dir)
    if not r4.get("success"):
        reason = r4.get("reason") or "脚本4执行失败"
        say(f"[停止] {reason}")
        workflow_log.write_text("\n".join(lines), encoding="utf-8")
        return {"success": False, "stage": 4, "reason": reason, "log_file": str(workflow_log)}

    say("\n[完成] 项目一键处理成功。")
    say(f"ERP输出：{r4.get('output_file', '')}")
    workflow_log.write_text("\n".join(lines), encoding="utf-8")
    return {"success": True, "stage": 5, "reason": "", "output_file": r4.get("output_file", ""), "log_file": str(workflow_log)}


def main() -> None:
    import argparse
    from project_utils import add_project_args, ensure_project_dir, write_result_json
    parser = argparse.ArgumentParser(description="脚本5：项目一键处理")
    add_project_args(parser)
    args = parser.parse_args()
    try:
        result = run_project(ensure_project_dir(args.project))
    except Exception as exc:
        result = {"success": False, "stage": 0, "reason": str(exc)}
        print(f"[错误] {exc}")
    write_result_json(args.result_json, result)
    raise SystemExit(0 if result.get("success") else 1)


if __name__ == "__main__":
    main()
