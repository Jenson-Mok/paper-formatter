#!/usr/bin/env python3
"""依赖安装检查脚本 for Paper Formatter Skill.

检查并安装 paper-formatter 所需的所有 Python 依赖包。
"""

import subprocess
import sys
import importlib

# Phase 1 必需依赖
REQUIRED = {
    "docx": "python-docx",
}

# Phase 1 可选但推荐的依赖
OPTIONAL = {
    "pypandoc": "pypandoc_binary",
}


def check_package(import_name: str) -> bool:
    """检查某个包是否可导入."""
    try:
        importlib.import_module(import_name)
        return True
    except ImportError:
        return False


def install_package(pip_name: str, label: str = "") -> bool:
    """使用 pip 安装包."""
    try:
        print(f"  📦 正在安装 {pip_name} ({label})...")
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", pip_name],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        return True
    except subprocess.CalledProcessError:
        return False


def main():
    print("=" * 60)
    print("  Paper Formatter Skill — 依赖检查")
    print("=" * 60)
    print()

    all_ok = True

    # 检查必需依赖
    print("🔍 检查必需依赖：")
    for import_name, pip_name in REQUIRED.items():
        if check_package(import_name):
            print(f"  ✅ {pip_name} 已安装")
        else:
            print(f"  ❌ {pip_name} 未安装")
            if install_package(pip_name, "必需"):
                print(f"  ✅ {pip_name} 安装成功")
            else:
                print(f"  ❌ {pip_name} 安装失败，请手动运行：pip install {pip_name}")
                all_ok = False

    print()

    # 检查可选依赖
    print("🔍 检查可选依赖（格式转换增强）：")
    for import_name, pip_name in OPTIONAL.items():
        if check_package(import_name):
            print(f"  ✅ {pip_name} 已安装")
        else:
            print(f"  ⚠️  {pip_name} 未安装（PDF/Markdown 格式转换需要）")
            if install_package(pip_name, "可选"):
                print(f"  ✅ {pip_name} 安装成功")
            else:
                print(f"  ⚠️  {pip_name} 安装失败，PDF/Markdown 转换功能将不可用")

    print()

    if all_ok:
        print("✅ 所有必需依赖已就绪，可以开始使用 Paper Formatter。")
    else:
        print("❌ 部分必需依赖未安装，请手动处理后重新运行。")

    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())
