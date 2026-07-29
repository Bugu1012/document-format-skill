#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""公文格式字体安装脚本（Windows）
将 fonts/ 目录中的字体安装到当前用户字体目录，无需管理员权限。
"""
import os
import sys
import shutil
import ctypes
import platform

FONT_DIR = os.path.dirname(os.path.abspath(__file__))

FONTS = [
    ("仿宋_GB2312.ttf", "仿宋_GB2312"),
    ("方正小标宋.TTF", "方正小标宋简体"),
    ("楷体_GB2312.ttf", "楷体_GB2312"),
]


def get_user_font_dir():
    return os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts")


def is_font_installed(font_name):
    try:
        import subprocess
        result = subprocess.run(
            ["powershell", "-Command",
             f"[System.Drawing.Text.InstalledFontCollection]::new().Families | "
             f"Where-Object {{ $_.Name -eq '{font_name}' }}"],
            capture_output=True, text=True, timeout=10
        )
        return font_name in result.stdout
    except Exception:
        return None


def install_font_user(src_path, font_name):
    """安装字体到当前用户目录（Windows 10 1809+，无需管理员）"""
    dest_dir = get_user_font_dir()
    os.makedirs(dest_dir, exist_ok=True)
    dest_path = os.path.join(dest_dir, os.path.basename(src_path))

    if os.path.exists(dest_path):
        print(f"  已存在，跳过：{dest_path}")
        return "skipped"

    shutil.copy2(src_path, dest_path)

    # 写注册表（当前用户）
    try:
        import winreg
        reg_path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_SET_VALUE)
        # 注册表值名格式："字体名 (TrueType)"
        winreg.SetValueEx(key, f"{font_name} (TrueType)", 0, winreg.REG_SZ, os.path.basename(src_path))
        winreg.CloseKey(key)
    except Exception as e:
        print(f"  注册表写入失败（字体文件已复制，重启后可能生效）：{e}")

    print(f"  已安装：{font_name} → {dest_path}")
    return "installed"


def main():
    if platform.system() != "Windows":
        print("本脚本仅支持 Windows。macOS/Linux 请双击字体文件手动安装。")
        sys.exit(1)

    print("公文格式字体安装")
    print("=" * 40)
    print(f"字体目录：{FONT_DIR}")
    print(f"安装目标：{get_user_font_dir()}")
    print()

    installed = 0
    skipped = 0
    missing = 0

    for filename, font_name in FONTS:
        src = os.path.join(FONT_DIR, filename)
        if not os.path.exists(src):
            print(f"[缺失] {filename}（{font_name}）— 文件不存在")
            missing += 1
            continue

        status = is_font_installed(font_name)
        if status is True:
            print(f"[已有] {font_name} — 系统已安装，跳过")
            skipped += 1
            continue

        print(f"[安装] {font_name}...")
        result = install_font_user(src, font_name)
        if result == "installed":
            installed += 1
        elif result == "skipped":
            skipped += 1

    print()
    print(f"完成：新装 {installed}，已有 {skipped}，缺失 {missing}")
    if installed > 0:
        print("提示：部分应用程序需要重启后才能识别新字体。")
    if missing > 0:
        print("提示：缺失的字体文件需从内部渠道获取后放入 fonts/ 目录。")


if __name__ == "__main__":
    main()
