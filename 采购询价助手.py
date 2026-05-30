#!/usr/bin/env python3
from __future__ import annotations

import os
import platform
import subprocess
import sys
from pathlib import Path


WORKFLOW_DIR = Path(__file__).resolve().parent
PROJECT_DIR = WORKFLOW_DIR.parent
OUTPUT_DIR = WORKFLOW_DIR / "output"
SUPPLIER_QUOTES_DIR = WORKFLOW_DIR / "input" / "supplier_quotes"
FEEDBACK_FILE = WORKFLOW_DIR / "input" / "negotiation_feedback.xlsx"


def py() -> str:
    return sys.executable


def ensure_dirs() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    SUPPLIER_QUOTES_DIR.mkdir(parents=True, exist_ok=True)


def open_path(path: Path) -> None:
    path = path.resolve()
    if not path.exists():
        print(f"路径不存在：{path}")
        return
    system = platform.system()
    if system == "Windows":
        os.startfile(path)  # type: ignore[attr-defined]
    elif system == "Darwin":
        subprocess.run(["open", str(path)], check=False)
    else:
        subprocess.run(["xdg-open", str(path)], check=False)


def latest_dir(keyword: str) -> Path | None:
    matches = [path for path in OUTPUT_DIR.glob(f"*{keyword}*") if path.is_dir()]
    if not matches:
        return None
    return max(matches, key=lambda path: path.stat().st_mtime)


def run_script(script_name: str, *args: str) -> None:
    script = WORKFLOW_DIR / "src" / script_name
    subprocess.run([py(), str(script), *args], cwd=PROJECT_DIR, check=True)


def pause() -> None:
    input("\n按回车键返回菜单...")


def run_split() -> None:
    print("正在按样板标准生成采购包拆分总表...")
    run_script("build_split_from_standard.py")
    latest = latest_dir("按昨日样板标准拆分")
    print(f"\n已完成：{latest}")
    if latest:
        open_path(latest)


def run_invitations_all() -> None:
    latest_split = latest_dir("按昨日样板标准拆分")
    if not latest_split:
        print("未找到拆分总表，先自动生成一份。")
        run_script("build_split_from_standard.py")
        latest_split = latest_dir("按昨日样板标准拆分")
    if not latest_split:
        print("拆分总表生成失败。")
        return
    split_file = latest_split / "采购包拆分总表-按昨日样板标准.xlsx"
    print("正在批量生成全部采购包邀请清单...")
    run_script("generate_invitations_from_split.py", "--standard", str(split_file), "--packages", "all")
    latest = latest_dir("采购包邀请清单")
    print(f"\n已完成：{latest}")
    if latest:
        open_path(latest)


def run_quotes_analysis() -> None:
    print("请先把供应商返回的报价 Excel 放入：")
    print(SUPPLIER_QUOTES_DIR.resolve())
    print("\n正在分析供应商报价...")
    run_script("analyze_supplier_quotes.py")
    latest = latest_dir("供应商报价分析")
    print(f"\n已完成：{latest}")
    if latest:
        open_path(latest)


def run_final_recommendation() -> None:
    if not FEEDBACK_FILE.exists():
        print("未发现谈判反馈表，先自动创建模板。")
    print(f"谈判反馈表位置：{FEEDBACK_FILE.resolve()}")
    print("正在生成最终推荐结果...")
    run_script("finalize_recommendation.py")
    latest = latest_dir("最终推荐结果")
    print(f"\n已完成：{latest}")
    if latest:
        open_path(latest)


def run_full_loop() -> None:
    print("正在一键跑完整闭环...")
    run_script("build_split_from_standard.py")
    latest_split = latest_dir("按昨日样板标准拆分")
    if not latest_split:
        print("拆分总表生成失败。")
        return
    split_file = latest_split / "采购包拆分总表-按昨日样板标准.xlsx"
    run_script("generate_invitations_from_split.py", "--standard", str(split_file), "--packages", "all")
    run_script("analyze_supplier_quotes.py")
    run_script("finalize_recommendation.py")
    print("\n完整闭环已跑完。")
    open_path(OUTPUT_DIR)


def clear() -> None:
    os.system("cls" if platform.system() == "Windows" else "clear")


def main() -> None:
    ensure_dirs()
    while True:
        clear()
        print("========================================")
        print("        采购询价智能助手")
        print("========================================\n")
        print("1. 一键跑完整闭环")
        print("2. 只生成采购包拆分总表")
        print("3. 只批量生成全部采购包邀请清单")
        print("4. 分析供应商报价并生成谈判策略")
        print("5. 生成最终推荐合作结果")
        print("6. 打开供应商报价放置文件夹")
        print("7. 打开谈判反馈表")
        print("8. 打开输出结果文件夹")
        print("9. 打开闭环使用说明")
        print("0. 退出\n")
        choice = input("请输入数字并按回车：").strip()
        print()
        try:
            if choice == "1":
                run_full_loop()
                pause()
            elif choice == "2":
                run_split()
                pause()
            elif choice == "3":
                run_invitations_all()
                pause()
            elif choice == "4":
                run_quotes_analysis()
                pause()
            elif choice == "5":
                run_final_recommendation()
                pause()
            elif choice == "6":
                open_path(SUPPLIER_QUOTES_DIR)
                pause()
            elif choice == "7":
                if not FEEDBACK_FILE.exists():
                    run_script("finalize_recommendation.py")
                open_path(FEEDBACK_FILE)
                pause()
            elif choice == "8":
                open_path(OUTPUT_DIR)
                pause()
            elif choice == "9":
                open_path(WORKFLOW_DIR / "闭环使用说明.md")
                pause()
            elif choice == "0":
                print("已退出。")
                return
            else:
                print("无效选择，请重新输入。")
                pause()
        except subprocess.CalledProcessError as exc:
            print(f"运行失败：{exc}")
            pause()


if __name__ == "__main__":
    main()
