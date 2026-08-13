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

# 🌐 Weintek-MultiLang-Translator
威纶通触摸屏多语言翻译工具

## ✨ Features
- 🖥️ 图形化GUI界面，操作简单直观
- 📄 支持Excel批量导入、导出翻译文本
- 🌍 面向威纶通HMI项目多语言文本管理
- 📦 支持打包独立EXE程序

- 🖥️ Graphical GUI interface, easy to operate
- 📄 Support Excel batch import & export translation content
- 🌍 Multilingual text management for Weintek HMI
- 📦 Can be packaged into standalone EXE

## 🛠 Environment
```bash
pip install -r requirements.txt

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
