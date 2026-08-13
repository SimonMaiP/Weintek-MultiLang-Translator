# 多语言翻译工具

基于 Python tkinter 的多语言翻译桌面应用，支持 XML / PDF 文件翻译，集成 6 种翻译引擎。

## 文件说明

| 文件 | 用途 |
|------|------|
| `translate_gui.py` | 主程序（GUI + 翻译逻辑） |
| `update_exe.py` | 辅助脚本 |
| `translate_excel.py` | Excel 翻译辅助脚本 |
| `requirements.txt` | Python 依赖列表 |
| `多语言翻译工具.spec` | PyInstaller 打包配置文件 |
| `icon.ico` / `icon.png` | 应用图标 |
| `更新日志.md` | 版本更新记录 |

## 环境准备

```bash
pip install -r requirements.txt
```

## 运行

```bash
python translate_gui.py
```

## 打包为 EXE

```bash
pyinstaller 多语言翻译工具.spec
```

输出在 `dist/多语言翻译工具/`，将 `多语言翻译工具.exe` + `_internal/` 一起复制到运行目录即可。
