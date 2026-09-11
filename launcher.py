from __future__ import annotations

import importlib.util
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
APP_DIR = ROOT / "app"

TASKS = {
    "1": ("导入客户物料清单", APP_DIR / "customer_import.py", True),
    "2": ("合并外部映射汇总表", APP_DIR / "mapping_merge.py", False),
    "3": ("生成客户清单京能物料对应表", APP_DIR / "material_fillback.py", True),
    "4": ("生成 ERP 物料清单批量导入表", APP_DIR / "erp_bom_export.py", True),
    "5": ("项目一键处理（1 -> 3 -> 4）", APP_DIR / "project_workflow.py", True),
}

REQUIRED_PACKAGES = ("openpyxl", "xlrd")
WINDOWS_ONLY_PACKAGES = ("pywin32",)


def clear_screen() -> None:
    subprocess.run("cls" if sys.platform == "win32" else "clear", shell=True)


def wait_for_enter() -> None:
    input("\n按回车键继续...")


def package_installed(pkg: str) -> bool:
    if pkg == "pywin32":
        return importlib.util.find_spec("win32com") is not None
    return importlib.util.find_spec(pkg) is not None


def check_dependencies() -> bool:
    packages = list(REQUIRED_PACKAGES)
    if sys.platform == "win32":
        packages.extend(WINDOWS_ONLY_PACKAGES)
    missing = [pkg for pkg in packages if not package_installed(pkg)]
    if not missing:
        return True
    print("检测到缺少 Python 依赖：")
    for pkg in missing:
        print(f"  - {pkg}")
    print("\n正在自动安装...")
    result = subprocess.run([sys.executable, "-m", "pip", "install", *missing])
    if result.returncode != 0:
        print("\n[错误] Python 依赖安装失败。")
        return False
    print("\n依赖安装完成。")
    return True


def select_project_folder() -> Path | None:
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        selected = filedialog.askdirectory(title="请选择项目文件夹")
        root.destroy()
    except Exception as exc:
        print(f"[错误] 无法打开文件夹选择窗口：{exc}")
        return None
    return Path(selected).resolve() if selected else None


def run_task(task_key: str) -> None:
    title, script, needs_project = TASKS[task_key]
    clear_screen()
    print("=" * 64)
    print(f"  {title}")
    print("=" * 64)
    print()

    if not script.exists():
        print(f"[错误] 找不到脚本：{script}")
        return
    if not check_dependencies():
        return
    if task_key in {"4", "5"} and sys.platform != "win32":
        print("[错误] ERP .xls 生成需要 Windows + Microsoft Excel。")
        return

    cmd = [sys.executable, str(script)]
    if needs_project:
        project = select_project_folder()
        if project is None:
            print("已取消选择项目文件夹。")
            return
        print(f"项目：{project}\n")
        cmd.extend(["--project", str(project)])

    result = subprocess.run(cmd, cwd=APP_DIR if needs_project else ROOT)
    print("\n" + "=" * 64)
    if result.returncode == 0:
        print("  本次任务已结束")
    else:
        print("  本次任务未完成，请查看上方原因和日志。")
    print("=" * 64)


def main() -> None:
    while True:
        clear_screen()
        print("=" * 64)
        print("                 嘉盛物料映射工具 v1.5.0")
        print("=" * 64)
        print()
        print("  1 - 导入客户物料清单（选择项目文件夹）")
        print("  2 - 合并外部映射汇总表（全局映射维护）")
        print("  3 - 生成客户清单京能物料对应表（选择项目文件夹）")
        print("  4 - 生成 ERP 物料清单批量导入表（选择项目文件夹）")
        print("  5 - 项目一键处理：1 -> 3 -> 4（推荐）")
        print()
        print("  Q - 退出")
        print()
        print("=" * 64)

        choice = input("请选择 1/2/3/4/5/Q：").strip().upper()
        if choice == "Q":
            return
        if choice not in TASKS:
            print("\n输入无效，请重新选择。")
            wait_for_enter()
            continue
        run_task(choice)
        while True:
            action = input("\nM - 返回主菜单    Q - 退出：").strip().upper()
            if action == "M":
                break
            if action == "Q":
                return
            print("输入无效，请输入 M 或 Q。")


if __name__ == "__main__":
    main()
