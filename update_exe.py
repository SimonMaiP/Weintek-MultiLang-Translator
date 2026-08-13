#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""多语言翻译工具 - 增量更新脚本

只编译 translate_gui.py 并替换到已打包的 onedir 目录中，秒级完成。
仅在修改了 translate_gui.py 源代码后运行此脚本。

用法：
    python update_exe.py

首次使用前需要完成一次完整的 onedir 打包（耗时约 3 分钟）：
    python -m PyInstaller 多语言翻译工具_onedir.spec --distpath "C:/Users/sf-bc4/Desktop/多语言翻译" --noconfirm
"""

import importlib.util
import marshal
import os
import shutil
import struct
import sys
import time

# ---- 配置 ----
PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
SOURCE_FILE = os.path.join(PROJECT_DIR, "translate_gui.py")
# onedir 打包输出: 桌面/多语言翻译/多语言翻译工具.exe + _internal/
DIST_DIR = r"D:\麦伟鹏\Python\多语言翻译"
INTERNAL_DIR = os.path.join(DIST_DIR, "_internal")


def find_target_pyc():
    """在 _internal 目录中找到 translate_gui 的 .pyc 文件"""
    for root, dirs, files in os.walk(INTERNAL_DIR):
        for f in files:
            if f == "translate_gui.pyc":
                return os.path.join(root, f)
    return None


def fast_compile():
    """编译 translate_gui.py 为 .pyc 字节码"""
    magic = importlib.util.MAGIC_NUMBER

    with open(SOURCE_FILE, "rb") as f:
        source = f.read()
    code = compile(source, SOURCE_FILE, "exec")

    # 构建 pyc header + code
    import io
    buf = io.BytesIO()
    buf.write(magic)
    buf.write(b'\x00\x00\x00\x00')
    source_mtime = int(os.path.getmtime(SOURCE_FILE))
    buf.write(struct.pack('<I', source_mtime & 0xFFFFFFFF))
    source_size = len(source)
    buf.write(struct.pack('<I', source_size & 0xFFFFFFFF))
    marshal.dump(code, buf)
    return buf.getvalue()


def main():
    print("=" * 50)
    print("  多语言翻译工具 - 增量更新")
    print("=" * 50)
    print()

    if not os.path.exists(SOURCE_FILE):
        print("[ERROR] 找不到源文件: {}".format(SOURCE_FILE))
        return 1

    if not os.path.exists(INTERNAL_DIR):
        print("[ERROR] 找不到 _internal 目录: {}".format(INTERNAL_DIR))
        print("请先完成首次 onedir 完整打包:")
        print('  python -m PyInstaller 多语言翻译工具_onedir.spec --distpath "C:/Users/sf-bc4/Desktop/多语言翻译" --noconfirm')
        return 1

    target_pyc = find_target_pyc()
    if not target_pyc:
        print("[WARN] 在 _internal 中未找到 translate_gui.pyc，将创建新文件")
        target_pyc = os.path.join(INTERNAL_DIR, "translate_gui.pyc")
    else:
        print("[OK] 找到目标: {}".format(target_pyc))

    t0 = time.time()

    print("[编译] translate_gui.py ...")
    pyc_data = fast_compile()

    if os.path.exists(target_pyc):
        backup = target_pyc + ".bak"
        shutil.copy2(target_pyc, backup)
        print("[备份] {}".format(os.path.basename(backup)))

    with open(target_pyc, "wb") as f:
        f.write(pyc_data)

    elapsed = time.time() - t0
    print()
    print("[DONE] 更新完成! 耗时 {:.2f} 秒".format(elapsed))
    print("[目标] {} -> {}".format(os.path.basename(SOURCE_FILE), target_pyc))
    print()
    print("现在可以启动 多语言翻译工具.exe 了")
    sys.exit(0)


if __name__ == "__main__":
    main()
