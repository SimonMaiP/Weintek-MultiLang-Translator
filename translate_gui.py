#!/usr/bin/env python3
"""
多语言翻译工具 v12.5 — 桌面GUI版
======================================
功能：导入XML文档/PDF文档，以中文为源，按需选择对应翻译语言。
      支持6个翻译API（MyMemory / LibreTranslate / Google / 百度翻译 / 翻译狗 / 火山引擎大模型）
      右上角设置弹窗管理API配置和密钥，拖拽排序优先级，带质量校验。
      关闭窗口最小化到系统托盘，右键托盘可退出。
      单实例运行，配置持久化到 settings.json。
      翻译狗使用MD5签名认证（需APP ID + 密钥）。
      火山引擎使用大模型（豆包）Chat Completions API（OpenAI兼容格式）。

v12.5 更新:
  - UI 全面重新设计，现代高级感风格
  - 自定义主题配色（石板蓝白底色 + 蓝色强调色）
  - 卡片式布局，圆角边框，渐入动效
  - 自定义按钮样式、输入框焦点高亮、状态指示圆点
  - 大模型设置页新增简短翻译模式（长句保留核心意思）

v12.4 更新:
  - 退出/关闭软件时自动保存所有勾选和设置（PDF语言、翻译模式、窗口位置大小、路径等）

v12.3 更新:
  - 删除大模型对话窗口，简化界面
  - API Key 明文显示
  - fitz（PyMuPDF）懒加载，启动速度大幅提升
  - 切换到 onedir 打包，支持增量更新

v12.0 新增:
  - 火山引擎大模型翻译源（豆包/方舟），支持11种语言
  - 设置页面新增火山引擎API Key + 模型ID配置
  - 支持拖拽调整火山引擎大模型在翻译源中的优先级

v11.8 修复:
  - 非中文/LOGO/纯英文不再翻译（5层提取过滤）
  - 白底不透明度提高到0.97消除原中文残影
  - 窗口默认高度增大确保底部状态栏可见

工作流：
  XML模式：复制源XML到输出路径 → 在副本上解析并翻译 → 保存副本
  PDF模式：文本提取翻译 → 在原PDF副本上覆盖翻译文字，保持原排版
          翻译狗文档翻译 → 上传PDF至翻译狗服务端，直接返回译文文档
"""

import os
import sys
import json
import ctypes
import shutil
import time
import queue
import hashlib
import random
import tempfile
import threading
import xml.etree.ElementTree as ET
from pathlib import Path
from datetime import datetime
from collections import defaultdict

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

import requests
# fitz (PyMuPDF) — 懒加载，仅在 PDF 功能时导入

# ============================================================
#  UI 主题系统 — 现代高级感配色 & 样式
# ============================================================
CLR = {
    "bg":           "#f5f6fa",   # 窗口底色（淡蓝灰）
    "card":         "#ffffff",   # 卡片背景
    "primary":      "#2563eb",   # 主色调（蓝）
    "primary_dark": "#1d4ed8",   # hover
    "success":      "#059669",   # 成功绿
    "error":        "#dc2626",   # 错误红
    "warn":         "#ea580c",   # 警告橙
    "text":         "#1e293b",   # 主文字
    "subtext":      "#64748b",   # 次要文字
    "border":       "#e2e8f0",   # 边框
    "hover_bg":     "#eff6ff",   # hover底色
    "title_bar":    "linear-gradient(135deg, #2563eb 0%, #7c3aed 100%)",  # 概念用
    "title_left":   "#2563eb",   # 标题区左侧强调色
    "tab_active":   "#2563eb",
    "tab_text":     "#64748b",
    "icon_gray":    "#94a3b8",
}

FONT = {
    "title":   ("Microsoft YaHei UI", 18, "bold"),
    "heading": ("Microsoft YaHei UI", 11, "bold"),
    "body":    ("Microsoft YaHei UI", 9),
    "small":   ("Microsoft YaHei UI", 8),
    "mono":    ("Cascadia Code", 9),
    "mono_sm": ("Cascadia Code", 8),
}

# ============================================================
#  创建自定义 ttk 样式
# ============================================================
def _setup_ttk_style():
    """配置 ttk 全局样式 — 现代简洁风"""
    style = ttk.Style()
    # 使用 'clam' 引擎以获得更好的自定义效果
    style.theme_use('clam')

    # ---- 通用背景 ----
    style.configure(".", background=CLR["bg"], foreground=CLR["text"],
                    font=FONT["body"])

    # ---- Frame ----
    style.configure("TFrame", background=CLR["bg"])
    style.configure("Card.TFrame", background=CLR["card"],
                    relief="flat", borderwidth=0)

    # ---- LabelFrame (卡片容器) ----
    style.configure("Card.TLabelframe", background=CLR["card"],
                    relief="solid", borderwidth=1,
                    bordercolor=CLR["border"])
    style.configure("Card.TLabelframe.Label", background=CLR["card"],
                    foreground=CLR["primary"], font=FONT["heading"])

    # ---- Label ----
    style.configure("TLabel", background=CLR["bg"], foreground=CLR["text"])
    style.configure("Card.TLabel", background=CLR["card"], foreground=CLR["text"])
    style.configure("Subtext.TLabel", foreground=CLR["subtext"], font=FONT["small"])
    style.configure("Primary.TLabel", foreground=CLR["primary"], font=FONT["heading"])
    style.configure("Success.TLabel", foreground=CLR["success"])
    style.configure("Error.TLabel", foreground=CLR["error"])

    # ---- Button ----
    style.configure("TButton", background=CLR["primary"],
                    foreground="#ffffff", font=FONT["body"],
                    borderwidth=0, relief="flat", padding=(16, 6))
    style.map("TButton",
              background=[("active", CLR["primary_dark"]),
                          ("disabled", "#cbd5e1")],
              foreground=[("disabled", "#94a3b8")])
    style.configure("Accent.TButton", background=CLR["primary"],
                    foreground="#ffffff", font=FONT["heading"],
                    borderwidth=0, relief="flat", padding=(20, 8))
    style.map("Accent.TButton",
              background=[("active", CLR["primary_dark"])])
    style.configure("Outline.TButton", background=CLR["card"],
                    foreground=CLR["primary"], font=FONT["body"],
                    relief="solid", borderwidth=1,
                    bordercolor=CLR["primary"], padding=(12, 4))
    style.map("Outline.TButton",
              background=[("active", CLR["hover_bg"])])

    # ---- Entry ----
    style.configure("TEntry", fieldbackground=CLR["card"],
                    foreground=CLR["text"], font=FONT["mono"],
                    relief="solid", borderwidth=1,
                    bordercolor=CLR["border"], padding=6)
    style.map("TEntry",
              bordercolor=[("focus", CLR["primary"])],
              relief=[("focus", "solid")])

    # ---- Combobox ----
    style.configure("TCombobox", fieldbackground=CLR["card"],
                    foreground=CLR["text"], font=FONT["mono"],
                    relief="solid", borderwidth=1,
                    bordercolor=CLR["border"], padding=6,
                    arrowcolor=CLR["primary"])
    style.map("TCombobox",
              bordercolor=[("focus", CLR["primary"]),
                           ("hover", CLR["primary"])])

    # ---- Notebook (Tab 切换) ----
    style.configure("TNotebook", background=CLR["bg"],
                    borderwidth=0, tabmargins=(0, 0, 0, 0))
    style.configure("TNotebook.Tab", background=CLR["bg"],
                    foreground=CLR["tab_text"], font=FONT["heading"],
                    padding=(24, 10), borderwidth=0,
                    lightcolor=CLR["bg"], darkcolor=CLR["bg"])
    style.map("TNotebook.Tab",
              background=[("selected", CLR["card"]),
                          ("!selected", CLR["bg"]),
                          ("active", CLR["hover_bg"])],
              foreground=[("selected", CLR["tab_active"]),
                          ("!selected", CLR["tab_text"])],
              expand=[("selected", [0, 0, 0, 0]),
                      ("!selected", [0, 0, 0, 0])],
              padding=[("selected", (24, 10)),
                       ("!selected", (24, 10))],
              lightcolor=[("selected", CLR["card"]),
                          ("!selected", CLR["bg"])],
              darkcolor=[("selected", CLR["card"]),
                         ("!selected", CLR["bg"])])

    # ---- Checkbutton / Radiobutton ----
    style.configure("TCheckbutton", background=CLR["card"],
                    foreground=CLR["text"], font=FONT["body"])
    style.configure("TRadiobutton", background=CLR["card"],
                    foreground=CLR["text"], font=FONT["body"])

    # ---- Scrollbar ----
    style.configure("TScrollbar", background=CLR["bg"],
                    troughcolor=CLR["bg"],
                    bordercolor=CLR["border"], arrowcolor=CLR["primary"])

    # ---- Separator ----
    style.configure("T Separator", background=CLR["border"])

    # ---- 拖拽高亮 ----
    style.configure("Drag.TFrame", background="#dbeafe")

    return style

# ============================================================
#  应用路径 & 单实例 & 配置持久化
# ============================================================
def get_app_dir():
    """获取应用所在目录（兼容源码运行和PyInstaller打包）"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))


def resource_path(relative_path):
    """获取资源文件绝对路径（兼容PyInstaller onefile模式）"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(get_app_dir(), relative_path)


# 应用图标对应的 emoji（多语言翻译 → 地球 🌐）
APP_EMOJI = "\U0001F310"


def _find_emoji_font(size):
    """查找系统中支持 emoji 的字体，返回 ImageFont，找不到返回 None。"""
    from PIL import ImageFont

    if sys.platform == "win32":
        candidates = [
            "C:/Windows/Fonts/seguiemj.ttf",
            "C:/Windows/Fonts/seguiemj.ttc",
            "C:/Windows/Fonts/msyh.ttc",
            "C:/Windows/Fonts/msyhbd.ttf",
            "C:/Windows/Fonts/msyh.ttf",
        ]
    elif sys.platform == "darwin":
        candidates = [
            "/System/Library/Fonts/Apple Color Emoji.ttc",
        ]
    else:
        candidates = [
            "/usr/share/fonts/truetype/noto/NotoColorEmoji.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        ]

    for path in candidates:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                continue
    return None


def make_emoji_icon(size=64):
    """使用 emoji 渲染应用图标，返回 PIL RGBA 图像（带圆形品牌底色）。

    保留原有的 icon.png 作为兜底，仅在 emoji 渲染不可用（无 PIL / 无 emoji 字体）时回退。
    """
    try:
        from PIL import Image, ImageDraw
    except Exception:
        return None

    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    # 圆形品牌底色，保证图标在任意背景下清晰可见
    draw.ellipse([0, 0, size, size], fill=(37, 99, 235, 255))

    font = _find_emoji_font(int(size * 0.6))
    if font is None:
        try:
            from PIL import ImageFont
            font = ImageFont.load_default()
        except Exception:
            font = None

    if font is not None:
        try:
            draw.text((size / 2, size / 2), APP_EMOJI,
                      font=font, fill=(255, 255, 255, 255), anchor="mm")
        except Exception:
            pass

    return img


def load_app_icon_pil():
    """加载应用图标为PIL Image对象（用于系统托盘），优先使用 emoji 渲染。"""
    img = make_emoji_icon(64)
    if img is not None:
        return img
    # 兜底：保留原有 icon.png
    try:
        from PIL import Image
        icon_path = resource_path('icon.png')
        if os.path.exists(icon_path):
            return Image.open(icon_path).convert('RGBA')
    except Exception:
        pass
    return None


def load_app_icon_tk():
    """加载应用图标为tkinter PhotoImage对象（用于窗口标题栏），优先使用 emoji 渲染。"""
    img = make_emoji_icon(64)
    if img is not None:
        try:
            import io
            from PIL import ImageTk
            buf = io.BytesIO()
            img.save(buf, format="PNG")
            buf.seek(0)
            return ImageTk.PhotoImage(data=buf.getvalue())
        except Exception:
            pass
    # 兜底：保留原有 icon.png
    try:
        icon_path = resource_path('icon.png')
        if os.path.exists(icon_path):
            return tk.PhotoImage(file=icon_path)
    except Exception:
        pass
    return None


APP_DIR = get_app_dir()
LOCK_FILE = os.path.join(tempfile.gettempdir(), 'translate_tool_multilang.lock')
CONFIG_FILE = os.path.join(APP_DIR, 'settings.json')


def _is_process_running_windows(pid):
    """Windows: 检查指定PID的进程是否仍在运行"""
    try:
        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        hProcess = ctypes.windll.kernel32.OpenProcess(
            PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        if hProcess:
            ctypes.windll.kernel32.CloseHandle(hProcess)
            return True
        return False
    except Exception:
        return False


def _check_single_instance():
    """单实例检测：通过PID锁文件防止重复启动。
       返回 True 表示可以继续运行，False 表示已有实例在运行。"""
    if os.path.exists(LOCK_FILE):
        try:
            with open(LOCK_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            pid = data.get('pid', 0)
            if _is_process_running_windows(pid):
                return False
            os.remove(LOCK_FILE)
        except Exception:
            try:
                os.remove(LOCK_FILE)
            except Exception:
                pass
    try:
        with open(LOCK_FILE, 'w', encoding='utf-8') as f:
            json.dump({'pid': os.getpid(), 'time': time.time()}, f)
        return True
    except Exception:
        return True


def _release_lock():
    """退出时释放锁文件"""
    try:
        if os.path.exists(LOCK_FILE):
            os.remove(LOCK_FILE)
    except Exception:
        pass


def load_config():
    """从 settings.json 加载配置"""
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            order = data.get('order', [])
            enabled = data.get('enabled', {})
            baidu_aid = data.get('baidu_appid', '')
            baidu_sec = data.get('baidu_secret', '')
            fy_aid = data.get('fanyigou_appid', '')
            fy_key = data.get('fanyigou_privatekey', '')
            fy_text_aid = data.get('fanyigou_text_appid', data.get('fanyigou_appid', ''))
            fy_text_key = data.get('fanyigou_text_privatekey', data.get('fanyigou_privatekey', ''))
            force_ret = data.get('force_retranslate', False)
            vc_key = data.get('volcengine_api_key', '')
            vc_model = data.get('volcengine_model', '')
            vc_concise = data.get('volcengine_concise', False)
            pdf_langs = data.get('pdf_langs', {})
            pdf_mod = data.get('pdf_mode', 'fanyigou')
            win_geo = data.get('window_geometry', None)
            xml_in = data.get('xml_input', '')
            xml_out = data.get('xml_output', '')
            pdf_in = data.get('pdf_input', '')
            pdf_out = data.get('pdf_output', '')
            all_pids = set(PROVIDER_REGISTRY.keys())
            if not order and not enabled:
                return None
            for pid in all_pids:
                if pid not in order:
                    order.append(pid)
                if pid not in enabled:
                    enabled[pid] = False
            return (order, enabled, baidu_aid, baidu_sec, fy_aid, fy_key,
                    fy_text_aid, fy_text_key, force_ret, vc_key, vc_model,
                    vc_concise, pdf_langs, pdf_mod, win_geo,
                    xml_in, xml_out, pdf_in, pdf_out)
    except Exception:
        pass
    return None


def save_config():
    """保存当前全局配置到 settings.json"""
    try:
        data = {
            'order': list(provider_order),
            'enabled': dict(provider_enabled),
            'baidu_appid': baidu_appid,
            'baidu_secret': baidu_secret,
            'fanyigou_appid': fanyigou_appid,
            'fanyigou_privatekey': fanyigou_privatekey,
            'fanyigou_text_appid': fanyigou_text_appid,
            'fanyigou_text_privatekey': fanyigou_text_privatekey,
            'volcengine_api_key': volcengine_api_key,
            'volcengine_model': volcengine_model,
            'volcengine_concise': volcengine_concise,
            'force_retranslate': force_retranslate,
            'pdf_langs': dict(pdf_lang_enabled),
            'pdf_mode': pdf_mode,
            'window_geometry': window_geometry,
            'xml_input': xml_input_path,
            'xml_output': xml_output_path,
            'pdf_input': pdf_input_path,
            'pdf_output': pdf_output_path,
            'saved_at': time.strftime('%Y-%m-%d %H:%M:%S'),
        }
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ============================================================
#  PDF 文字提取 & 段落拆分
# ============================================================
# ============================================================
#  PDF 文字块聚类辅助函数
# ============================================================
def _cluster_words(words):
    """将 get_text('words') 返回的单字按 Y 坐标和间距聚类为文字块。
    输入: [(x0,y0,x1,y1,text,block_no,line_no,word_no), ...]
    返回: [((x0,y0,x1,y1), text), ...]
    """
    if not words:
        return []
    # 简化：提取 bbox 和 text
    items = [(w[0], w[1], w[2], w[3], w[4]) for w in words]  # (x0,y0,x1,y1,text)
    return _cluster_by_proximity(items)


def _merge_spans(spans):
    """合并 dict 模式提取的相邻 span。
    输入: [(bbox, text), ...]  bbox=(x0,y0,x1,y1)
    返回: [((x0,y0,x1,y1), text), ...]
    """
    if not spans:
        return []
    # 转为统一格式
    items = [(b[0], b[1], b[2], b[3], t) for b, t in spans]
    return _cluster_by_proximity(items)


def _cluster_pdfplumber_words(words):
    """聚类 pdfplumber 提取的单词。
    输入: [(x0, top, x1, bottom, text), ...]
    返回: [((x0,y0,x1,y1), text), ...]
    """
    if not words:
        return []
    return _cluster_by_proximity(words)


def _cluster_by_proximity(items):
    """通用聚类：按 Y 坐标分层后，同行内按 X 间距合并。
    items: [(x0,y0,x1,y1,text), ...]
    返回: [((x0,y0,x1,y1), text), ...]
    """
    if not items:
        return []

    # 计算平均字符宽度作为间距阈值
    all_texts = [t for _, _, _, _, t in items if t.strip()]
    avg_len = sum(len(t) for t in all_texts) / max(len(all_texts), 1) if all_texts else 3
    avg_w = sum(i[2] - i[0] for i in items) / max(len(items), 1)
    char_w = avg_w / max(avg_len, 1) if avg_w > 0 else 5
    gap_threshold = char_w * 4  # 4个字符宽度的间距视为同一块

    # 按 Y 中心排序
    items = sorted(items, key=lambda i: (i[1] + i[3]) / 2)

    # 分行
    lines = []
    current_line = [items[0]]
    for item in items[1:]:
        prev_cy = (current_line[-1][1] + current_line[-1][3]) / 2
        cur_cy = (item[1] + item[3]) / 2
        prev_h = current_line[-1][3] - current_line[-1][1]
        if abs(cur_cy - prev_cy) < prev_h * 1.2:
            current_line.append(item)
        else:
            lines.append(current_line)
            current_line = [item]
    lines.append(current_line)

    # 每行内按 X 排序并合并
    result = []
    for line in lines:
        line.sort(key=lambda i: i[0])  # 按 x0
        merged = [list(line[0])]
        for item in line[1:]:
            prev = merged[-1]
            gap = item[0] - prev[2]  # 新词x0 - 旧词x1
            if gap < gap_threshold and gap > -gap_threshold:
                # 合并：扩展bbox，拼接text
                prev[2] = max(prev[2], item[2])
                prev[3] = max(prev[3], item[3])
                prev[4] = prev[4] + " " + item[4]
            else:
                merged.append(list(item))
        for m in merged:
            bbox = (m[0], m[1], m[2], m[3])
            text = m[4].strip()
            if text:
                result.append((bbox, text))

    return result


def _ocr_extract_page(page, pi, log_func):
    """方法5: EasyOCR — 图片型/扫描件PDF 文字提取
    渲染页面为图片 → OCR识别 → 返回 [(bbox, text), ...] (PDF坐标系)
    """
    try:
        import easyocr
        import numpy as np
        from PIL import Image
        import io
    except ImportError:
        log_func("  [OCR] EasyOCR 未安装，跳过 (pip install easyocr)", "info")
        return []

    log_func("  [OCR] 第 {} 页为图片型，正在OCR识别中...".format(pi + 1))

    # 渲染页面为图片（200 DPI 兼顾速度与精度）
    pix = page.get_pixmap(dpi=200)
    img = Image.open(io.BytesIO(pix.tobytes("png")))
    img_np = np.array(img)

    img_h, img_w = img_np.shape[:2]
    page_w = page.rect.width
    page_h = page.rect.height

    # 坐标缩放因子：图片像素 → PDF 坐标 (points)
    scale_x = page_w / img_w
    scale_y = page_h / img_h

    # 初始化 EasyOCR reader（首次运行会下载模型 ~80MB）
    try:
        reader = easyocr.Reader(["ch_sim", "en"], gpu=False, verbose=False)
        results = reader.readtext(img_np)
    except Exception as e:
        log_func("  [OCR] 识别失败: {}".format(str(e)), "error")
        return []

    blocks = []
    for bbox, text, conf in results:
        text = text.strip()
        if len(text) < 2:
            continue
        # 跳过纯数字/符号/纯英文（只保留包含中文的文字块）
        has_chinese = any('\u4e00' <= c <= '\u9fff' for c in text)
        if not has_chinese:
            continue

        # 转换坐标：图片像素 → PDF points
        # bbox: [[x0,y0], [x1,y1], [x2,y2], [x3,y3]] (四边形)
        xs = [p[0] for p in bbox]
        ys = [p[1] for p in bbox]
        pdf_bbox = (min(xs) * scale_x, min(ys) * scale_y,
                    max(xs) * scale_x, max(ys) * scale_y)
        blocks.append((pdf_bbox, text))

    log_func("  => 第 {} 页: OCR 识别 {} 块".format(pi + 1, len(blocks)))
    return blocks


def extract_pdf_text(filepath):
    """使用 pdfplumber 提取 PDF 所有页面的文字"""
    try:
        import pdfplumber
    except ImportError:
        raise ImportError("缺少 pdfplumber 库，请安装: pip install pdfplumber")

    all_text = []
    with pdfplumber.open(filepath) as pdf:
        total = len(pdf.pages)
        for i, page in enumerate(pdf.pages):
            page_text = page.extract_text()
            if page_text:
                all_text.append("=== 第 {} 页 / 共 {} 页 ===\n{}".format(
                    i + 1, total, page_text.strip()))
    return "\n\n".join(all_text)


def split_paragraphs(text):
    """将提取的文字拆分为可翻译的段落"""
    raw_paragraphs = [p.strip() for p in text.split('\n\n') if p.strip()]
    result = []
    for p in raw_paragraphs:
        # 页面标记行保持原样
        if p.startswith('=== 第'):
            if result:
                result.append('')  # 页面间加空行
            result.append(p)
            continue
        # 长段落按换行进一步拆分
        if len(p) > 600:
            sub = [s.strip() for s in p.split('\n') if s.strip()]
            for s in sub:
                if len(s) > 800:
                    # 按句号再拆
                    parts = [x.strip() + '。' for x in s.replace('！', '。').replace('？', '。').split('。') if x.strip()]
                    result.extend(parts)
                else:
                    result.append(s)
        else:
            result.append(p)
    return [x for x in result if x]  # 去掉空字符串


# ============================================================
#  配置 — XML 标签 → (API语言代码, 显示名称)
# ============================================================
LANG_MAP = {
    "英语":     ("en",     "英语"),
    "繁体":     ("zh-TW",  "繁体中文"),
    "日语":     ("ja",     "日语"),
    "韩语":     ("ko",     "韩语"),
    "德语":     ("de",     "德语"),
    "西班牙语": ("es",     "西班牙语"),
    "俄语":     ("ru",     "俄语"),
    "法语":     ("fr",     "法语"),
    "葡萄牙语": ("pt",     "葡萄牙语"),
    "意大利语": ("it",     "意大利语"),
    "泰语":     ("th",     "泰语"),
}

# PDF 页面使用的语言列表（按显示顺序）
PDF_LANG_LIST = [
    ("en",     "英语"),
    ("zh-TW",  "繁体中文"),
    ("ja",     "日语"),
    ("ko",     "韩语"),
    ("de",     "德语"),
    ("es",     "西班牙语"),
    ("ru",     "俄语"),
    ("fr",     "法语"),
    ("pt",     "葡萄牙语"),
    ("it",     "意大利语"),
    ("th",     "泰语"),
]

# 百度翻译语言代码映射
BAIDU_LANG_MAP = {
    "en": "en", "zh-TW": "cht", "ja": "jp", "ko": "kor",
    "de": "de", "es": "spa", "ru": "ru", "fr": "fra",
    "pt": "pt", "it": "it", "th": "th",
}

# 翻译狗语言代码映射
FANYIGOU_LANG_MAP = {
    "en": "en", "zh-TW": "cht", "ja": "jp", "ko": "kor",
    "de": "de", "es": "spa", "ru": "ru", "fr": "fra",
    "pt": "pt", "it": "it", "th": "th",
}

# 火山引擎大模型 — 目标语言中文名称（用于prompt指令）
VOLCENGINE_LANG_NAME_MAP = {
    "en": "English", "zh-TW": "Traditional Chinese (繁體中文)",
    "ja": "Japanese", "ko": "Korean", "de": "German",
    "es": "Spanish", "ru": "Russian", "fr": "French",
    "pt": "Portuguese", "it": "Italian", "th": "Thai",
}

MYMEMORY_URL = "https://api.mymemory.translated.net/get"
LIBRETRANSLATE_URL = "https://libretranslate.com/translate"
GOOGLE_TRANSLATE_URL = "https://translate.googleapis.com/translate_a/single"
BAIDU_TRANSLATE_URL = "https://fanyi-api.baidu.com/api/trans/vip/translate"
FANYIGOU_TRANSLATE_URL = "https://www.fanyigou.com/TranslateApi/api/trans"
FANYIGOU_UPLOAD_URL = "https://www.fanyigou.com/TranslateApi/api/uploadTranslate"
FANYIGOU_PROGRESS_URL = "https://www.fanyigou.com/TranslateApi/api/queryTransProgress"
FANYIGOU_DOWNLOAD_URL = "https://www.fanyigou.com/TranslateApi/api/downloadFile"
FANYIGOU_DETECT_URL = "https://www.fanyigou.com/TranslateApi/api/detectDocPage"
FANYIGOU_SUBMIT_URL = "https://www.fanyigou.com/TranslateApi/api/submitForDetectDoc"
VOLCENGINE_URL = "https://ark.cn-beijing.volces.com/api/v3/chat/completions"
TRANSLATE_DELAY = 0.4
POLL_INTERVAL = 3  # 翻译狗进度轮询间隔(秒)
MAX_RETRIES = 2

ISO_CODES = {'ZH', 'EN', 'JA', 'KO', 'DE', 'ES', 'RU', 'FR', 'PT', 'IT', 'TH',
             'ZH-CN', 'ZH-TW', 'ZHTW', 'CHS', 'CHT'}

API_DESCRIPTIONS = {
    "mymemory":       "免费,无需Key,质量一般",
    "libretranslate": "免费开源,无需Key,速度较慢",
    "google":         "Google翻译,免费,质量最好",
    "baidu":          "百度翻译,需填APP ID+密钥(免费200万字/月)",
    "fanyigou":       "翻译狗,需填APP ID+密钥(文档翻译专业)",
    "volcengine":     "火山引擎豆包大模型,需填API Key+模型(按token计费)",
}


# ============================================================
#  质量校验
# ============================================================
def _has_chinese(text):
    """检查文本是否包含中文字符（CJK统一汉字）"""
    for ch in text:
        if '\u4e00' <= ch <= '\u9fff' or '\u3400' <= ch <= '\u4dbf':
            return True
    return False


def quality_check(source, result, target_lang):
    if not result:
        return False, "空结果"
    r = result.strip()
    s = source.strip()
    # 与源文本相同时：
    #   - 繁体中文跳过（简繁同形）
    #   - 源文本无中文字符（纯符号/数字/单位）也跳过，相同是正常的
    if r == s and target_lang != "zh-TW":
        if _has_chinese(s):
            return False, "与源文本相同"
    lang_names_cn = {
        'en': '英语', 'zh-TW': '繁体中文', 'ja': '日语', 'ko': '韩语',
        'de': '德语', 'es': '西班牙语', 'ru': '俄语', 'fr': '法语',
        'pt': '葡萄牙语', 'it': '意大利语', 'th': '泰语',
    }
    if r in lang_names_cn.values():
        return False, "返回了语言名称"
    if r.upper() in ISO_CODES:
        return False, "返回了国家代码: {}".format(r)
    if len(s) >= 4 and len(r) <= 2 and not r.isascii():
        return False, "翻译过短(比例失调)"
    bad_keywords = ['版本', 'version', 'versao', 'version']
    if any(kw in r.lower() for kw in bad_keywords):
        return False, "包含无关关键词"
    return True, "OK"


# ============================================================
#  翻译提供方
# ============================================================
def translate_mymemory(text, target_lang):
    langpair = "zh-CN|{}".format(target_lang)
    params = {'q': text, 'langpair': langpair, 'mt': '1'}
    for _ in range(MAX_RETRIES):
        try:
            resp = requests.get(MYMEMORY_URL, params=params, timeout=15)
            if resp.status_code != 200:
                time.sleep(1.5); continue
            data = resp.json()
            for m in data.get('matches', []):
                if m.get('model') == 'neural' or m.get('created-by') == 'MT!':
                    return m.get('translation', '')
            return data.get('responseData', {}).get('translatedText', '') or ''
        except Exception:
            time.sleep(1.5)
    return ''


def translate_libretranslate(text, target_lang):
    payload = {'q': text, 'source': 'zh', 'target': target_lang, 'format': 'text'}
    for _ in range(MAX_RETRIES):
        try:
            resp = requests.post(LIBRETRANSLATE_URL, json=payload, timeout=20)
            if resp.status_code == 200:
                return resp.json().get('translatedText', '')
            time.sleep(1.5)
        except Exception:
            time.sleep(1.5)
    return ''


def translate_google(text, target_lang):
    params = {'client': 'gtx', 'sl': 'zh-CN', 'tl': target_lang, 'dt': 't', 'q': text}
    for _ in range(MAX_RETRIES):
        try:
            resp = requests.get(GOOGLE_TRANSLATE_URL, params=params, timeout=15)
            if resp.status_code != 200:
                time.sleep(1.5); continue
            data = resp.json()
            parts = [seg[0] for seg in data[0] if seg and seg[0]]
            return ''.join(parts) if parts else ''
        except Exception:
            time.sleep(1.5)
    return ''


def translate_baidu(text, target_lang):
    appid = baidu_appid.strip()
    secret = baidu_secret.strip()
    if not appid or not secret:
        return ''
    baidu_tl = BAIDU_LANG_MAP.get(target_lang, target_lang)
    salt = str(random.randint(32768, 65536))
    sign_str = appid + text + salt + secret
    sign = hashlib.md5(sign_str.encode('utf-8')).hexdigest()
    params = {
        'q': text, 'from': 'zh', 'to': baidu_tl,
        'appid': appid, 'salt': salt, 'sign': sign,
    }
    for _ in range(MAX_RETRIES):
        try:
            resp = requests.get(BAIDU_TRANSLATE_URL, params=params, timeout=15)
            if resp.status_code != 200:
                time.sleep(1.5); continue
            data = resp.json()
            if 'trans_result' in data and data['trans_result']:
                return data['trans_result'][0].get('dst', '')
            return ''
        except Exception:
            time.sleep(1.5)
    return ''


def translate_fanyigou(text, target_lang):
    """翻译狗文字翻译API — MD5签名认证：使用文字翻译专用APP ID + 密钥"""
    appid = fanyigou_text_appid.strip()
    privatekey = fanyigou_text_privatekey.strip()
    if not appid or not privatekey:
        return ''
    fy_tl = FANYIGOU_LANG_MAP.get(target_lang, target_lang)
    nonce_str = str(random.randint(1000000000, 9999999999))

    sign_params = {
        'appid': appid, 'from': 'zh', 'nonce_str': nonce_str,
        'text': text, 'to': fy_tl, 'privatekey': privatekey,
    }
    sorted_keys = sorted(sign_params.keys())
    sign_str = '&'.join('{}={}'.format(k, sign_params[k]) for k in sorted_keys)
    token = hashlib.md5(sign_str.encode('utf-8')).hexdigest().upper()

    post_data = {
        'appid': appid, 'nonce_str': nonce_str,
        'from': 'zh', 'to': fy_tl, 'text': text, 'token': token,
    }
    for _ in range(MAX_RETRIES):
        try:
            resp = requests.post(FANYIGOU_TRANSLATE_URL, data=post_data, timeout=15)
            if resp.status_code != 200:
                time.sleep(1.5); continue
            data = resp.json()
            if data.get('code') == 0 and 'data' in data:
                return data['data'].get('transResult', '')
            return ''
        except Exception:
            time.sleep(1.5)
    return ''


def translate_volcengine(text, target_lang):
    """火山引擎大模型翻译 — 豆包/方舟 Chat Completions (OpenAI兼容)
       支持 concise 模式：长句保留核心意思即可"""
    api_key = volcengine_api_key.strip()
    model = volcengine_model.strip()
    if not api_key or not model:
        return ''

    lang_name = VOLCENGINE_LANG_NAME_MAP.get(target_lang, target_lang)
    if volcengine_concise:
        system_prompt = (
            "You are a professional concise translator. Translate the user's Chinese text "
            "into {}. Keep only the core meaning. For long sentences, condense to the essential "
            "message — shorter is better, but never lose the key information. "
            "Output ONLY the translation, nothing else."
        ).format(lang_name)
    else:
        system_prompt = (
            "You are a professional translator. Translate the user's Chinese text "
            "into {}. Output ONLY the translation, nothing else. "
            "Do not add explanations, notes, or any other text."
        ).format(lang_name)
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": text},
        ],
        "temperature": 0.1,
        "max_tokens": 2048,
    }
    headers = {
        "Authorization": "Bearer {}".format(api_key),
        "Content-Type": "application/json",
    }
    for _ in range(MAX_RETRIES):
        try:
            resp = requests.post(VOLCENGINE_URL, json=payload,
                                 headers=headers, timeout=30)
            if resp.status_code != 200:
                time.sleep(2); continue
            data = resp.json()
            if "choices" in data and data["choices"]:
                content = data["choices"][0]["message"]["content"]
                return content.strip() if content else ''
            return ''
        except Exception:
            time.sleep(2)
    return ''


# ============================================================
#  翻译狗文档翻译 API（上传/查询/下载）
# ============================================================
def _fanyigou_md5_sign(appid, privatekey, params):
    """翻译狗MD5签名：按ASCII排序拼接后MD5大写"""
    sign_params = dict(params)
    sign_params['privatekey'] = privatekey
    sorted_keys = sorted(sign_params.keys())
    sign_str = '&'.join('{}={}'.format(k, sign_params[k]) for k in sorted_keys)
    return hashlib.md5(sign_str.encode('utf-8')).hexdigest().upper()


def fanyigou_upload_translate(filepath, to_lang):
    """上传文档到翻译狗并启动翻译，返回 (tid, error_msg)"""
    appid = fanyigou_appid.strip()
    privatekey = fanyigou_privatekey.strip()
    if not appid or not privatekey:
        return None, "翻译狗未配置APP ID和密钥"

    fy_tl = FANYIGOU_LANG_MAP.get(to_lang, to_lang)

    # 计算文件MD5
    file_md5 = hashlib.md5()
    with open(filepath, 'rb') as f:
        for chunk in iter(lambda: f.read(8192), b''):
            file_md5.update(chunk)
    file_md5_str = file_md5.hexdigest()

    nonce_str = str(random.randint(1000000000, 9999999999))
    params = {
        'appid': appid, 'nonce_str': nonce_str,
        'from': 'zh', 'to': fy_tl, 'md5': file_md5_str,
    }
    token = _fanyigou_md5_sign(appid, privatekey, params)
    params['token'] = token

    try:
        with open(filepath, 'rb') as f:
            files = {'file': (os.path.basename(filepath), f)}
            resp = requests.post(FANYIGOU_UPLOAD_URL, data=params,
                                 files=files, timeout=120)
        data = resp.json()
        if data.get('code') == 100 and 'data' in data:
            return data['data'].get('tid'), None
        return None, data.get('msg', '上传失败(未知错误)')
    except Exception as e:
        return None, str(e)


def fanyigou_query_progress(tid):
    """查询翻译进度，返回完整API响应dict"""
    appid = fanyigou_appid.strip()
    privatekey = fanyigou_privatekey.strip()

    nonce_str = str(random.randint(1000000000, 9999999999))
    params = {'appid': appid, 'nonce_str': nonce_str, 'tid': str(tid)}
    token = _fanyigou_md5_sign(appid, privatekey, params)
    params['token'] = token

    try:
        resp = requests.post(FANYIGOU_PROGRESS_URL, data=params, timeout=30)
        return resp.json()
    except Exception as e:
        return {'code': -1, 'msg': str(e)}


def fanyigou_download_file(tid, dtype, output_path):
    """下载翻译完成的文档到本地，返回 (success, error_msg)
       dtype: 2=PDF输出, 3=Word/PPT/Excel输出"""
    appid = fanyigou_appid.strip()
    privatekey = fanyigou_privatekey.strip()

    nonce_str = str(random.randint(1000000000, 9999999999))
    params = {'appid': appid, 'nonce_str': nonce_str,
              'tid': str(tid), 'dtype': str(dtype)}
    token = _fanyigou_md5_sign(appid, privatekey, params)
    params['token'] = token

    try:
        resp = requests.post(FANYIGOU_DOWNLOAD_URL, data=params, timeout=120)
        ct = resp.headers.get('Content-Type', '')
        if 'application/octet-stream' in ct or 'application/pdf' in ct:
            with open(output_path, 'wb') as f:
                f.write(resp.content)
            return True, None
        try:
            err = resp.json()
            return False, err.get('msg', '下载失败')
        except Exception:
            return False, '下载返回非预期格式 (Content-Type: {})'.format(ct)
    except Exception as e:
        return False, str(e)


# ============================================================
#  提供方注册表 + 运行时配置
# ============================================================
PROVIDER_REGISTRY = {
    "mymemory":       {"name": "MyMemory",       "fn": translate_mymemory},
    "libretranslate": {"name": "LibreTranslate",  "fn": translate_libretranslate},
    "google":         {"name": "Google 翻译",     "fn": translate_google},
    "baidu":          {"name": "百度翻译",        "fn": translate_baidu},
    "fanyigou":       {"name": "翻译狗",          "fn": translate_fanyigou},
    "volcengine":     {"name": "火山引擎大模型",  "fn": translate_volcengine},
}

provider_order = ["mymemory", "libretranslate", "google", "baidu", "fanyigou", "volcengine"]
provider_enabled = {"mymemory": True, "libretranslate": True, "google": True, "baidu": False, "fanyigou": False, "volcengine": False}
baidu_appid = ""
baidu_secret = ""
fanyigou_appid = ""
fanyigou_privatekey = ""
fanyigou_text_appid = ""
fanyigou_text_privatekey = ""
volcengine_api_key = ""
volcengine_model = ""
volcengine_concise = False  # 大模型简短翻译模式
current_api_name = ""
force_retranslate = False
# PDF 界面设置
pdf_lang_enabled = {}       # {lang_code: bool}
pdf_mode = "fanyigou"       # "fanyigou" or "text"
window_geometry = None      # {"x": ..., "y": ..., "width": ..., "height": ...}
xml_input_path = ""
xml_output_path = ""
pdf_input_path = ""
pdf_output_path = ""

_config = load_config()
if _config is not None:
    (provider_order, provider_enabled, baidu_appid, baidu_secret,
     fanyigou_appid, fanyigou_privatekey, fanyigou_text_appid,
     fanyigou_text_privatekey, force_retranslate, volcengine_api_key,
     volcengine_model, volcengine_concise, pdf_lang_enabled, pdf_mode,
     window_geometry, xml_input_path, xml_output_path, pdf_input_path,
     pdf_output_path) = _config


def get_active_providers():
    result = []
    for pid in provider_order:
        if provider_enabled.get(pid, False):
            info = PROVIDER_REGISTRY[pid]
            result.append((pid, info["fn"], info["name"]))
    return result


def translate_text(text, target_lang):
    global current_api_name
    last_reason = ""
    for provider_id, translate_fn, display_name in get_active_providers():
        result = translate_fn(text, target_lang)
        passed, reason = quality_check(text, result, target_lang)
        if passed:
            current_api_name = display_name
            return result, display_name, ""
        last_reason = "{} ({})".format(reason, display_name)
    current_api_name = ""
    return '', None, last_reason or "所有API均返回空"


def analyze_xml(filepath):
    tree = ET.parse(filepath)
    root = tree.getroot()
    total_entries = 0
    cn_count = 0
    missing_by_lang = {}
    tasks = []
    for text_elem in root.findall('文本'):
        total_entries += 1
        cn_elem = text_elem.find('中文')
        if cn_elem is None:
            continue
        cn_text = (cn_elem.text or '').strip()
        if not cn_text:
            continue
        cn_count += 1
        for tag, (lang_code, lang_name) in LANG_MAP.items():
            lang_elem = text_elem.find(tag)
            if lang_elem is None:
                # 标签不存在则自动创建
                lang_elem = ET.SubElement(text_elem, tag)
            existing = (lang_elem.text or '').strip()
            if not existing:
                missing_by_lang[lang_name] = missing_by_lang.get(lang_name, 0) + 1
                tasks.append((text_elem, cn_text, tag, lang_elem, lang_code, lang_name))
            elif force_retranslate:
                missing_by_lang[lang_name] = missing_by_lang.get(lang_name, 0) + 1
                tasks.append((text_elem, cn_text, tag, lang_elem, lang_code, lang_name))
    return tree, total_entries, cn_count, missing_by_lang, tasks


def check_network():
    try:
        r = requests.get("https://www.baidu.com", timeout=5)
        if r.status_code == 200:
            return True, "已连接"
        return False, "异常 (HTTP {})".format(r.status_code)
    except Exception:
        return False, "未连接"


def check_api():
    results = []
    for pid, fn, name in get_active_providers():
        try:
            test_result = fn("test", "en")
            ok = bool(test_result and test_result.strip())
        except Exception:
            ok = False
        results.append((name, ok))
    available = [n for n, ok in results if ok]
    if available:
        return True, "翻译源: " + ", ".join(available) + " (共{}个)".format(len(available))
    return False, "翻译源: 全部不可用"


# ============================================================
#  GUI (tkinter) — 主应用类
# ============================================================


# ============================================================
#  设置弹窗 — API拖拽排序 + 密钥
# ============================================================
class SettingsDialog(tk.Toplevel):
    """设置弹窗：拖拽排序API优先级 + 密钥输入 + 大模型简译开关
       关闭即保存"""

    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.title("API 与翻译设置")
        self.configure(bg=CLR["bg"])
        self.geometry("640x820")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        parent.update_idletasks()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        x = px + (pw - 640) // 2
        y = py + (ph - 820) // 2
        self.geometry("+{}+{}".format(x, y))

        self._drag_pid = None
        self._drag_start_y = 0
        self._drop_target_pid = None

        self._temp_order = list(provider_order)
        self._temp_enabled = dict(provider_enabled)
        self._temp_appid = baidu_appid
        self._temp_secret = baidu_secret
        self._temp_fy_appid = fanyigou_appid
        self._temp_fy_privatekey = fanyigou_privatekey
        self._temp_fy_text_appid = fanyigou_text_appid
        self._temp_fy_text_privatekey = fanyigou_text_privatekey
        self._temp_vc_api_key = volcengine_api_key
        self._temp_vc_model = volcengine_model
        self._temp_vc_concise = volcengine_concise
        self._temp_force_retranslate = force_retranslate

        self._build_ui()
        self._refresh_ranks()

        self.protocol("WM_DELETE_WINDOW", self._save_and_close)

    def _build_ui(self):
        self.configure(bg=CLR["bg"])

        # ---- 可滚动容器 (Notebook + 内容) ----
        scroll_container = tk.Frame(self, bg=CLR["bg"])
        scroll_container.pack(fill=tk.BOTH, expand=True, padx=14, pady=(14, 4))

        self._settings_canvas = tk.Canvas(scroll_container, bg=CLR["bg"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(scroll_container, orient=tk.VERTICAL,
                                  command=self._settings_canvas.yview)
        self._settings_canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self._settings_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        content_frame = tk.Frame(self._settings_canvas, bg=CLR["bg"])
        self._settings_canvas.create_window((0, 0), window=content_frame,
                                            anchor=tk.NW, width=610)

        def _on_content_configure(event):
            self._settings_canvas.configure(scrollregion=self._settings_canvas.bbox("all"))
        content_frame.bind("<Configure>", _on_content_configure)

        # 鼠标滚轮绑定
        def _on_mousewheel(event):
            self._settings_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        self._settings_canvas.bind_all("<MouseWheel>", _on_mousewheel)

        # ---- Notebook ----
        self.notebook = ttk.Notebook(content_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # 切换标签时更新滚动区域
        def _on_tab_changed(event):
            self.after(50, lambda: self._settings_canvas.configure(
                scrollregion=self._settings_canvas.bbox("all")))
        self.notebook.bind("<<NotebookTabChanged>>", _on_tab_changed)

        # ====== XML 设置 Tab ======
        xml_tab = tk.Frame(self.notebook, bg=CLR["bg"])
        self.notebook.add(xml_tab, text="XML 设置")

        # API列表
        list_frame = tk.Frame(xml_tab, bg=CLR["card"], highlightbackground=CLR["border"],
                              highlightthickness=1, bd=0)
        list_frame.pack(fill=tk.X, pady=(2, 8), padx=2)

        tk.Label(list_frame, text="\U0001f504  翻译API (拖拽行可调整优先级 | 勾选=启用)",
                 font=FONT["heading"], bg=CLR["card"], fg=CLR["text"]).pack(
                 anchor=tk.W, padx=12, pady=(10, 6))

        hdr = tk.Frame(list_frame, bg=CLR["card"])
        hdr.pack(fill=tk.X, padx=12, pady=(0, 6))
        tk.Label(hdr, text="  ", width=3, bg=CLR["card"]).pack(side=tk.LEFT)
        tk.Label(hdr, text="启用", width=4, anchor=tk.CENTER,
                  font=FONT["small"], bg=CLR["card"], fg=CLR["subtext"]).pack(side=tk.LEFT, padx=(0, 8))
        tk.Label(hdr, text="序号", width=4, anchor=tk.CENTER,
                  font=FONT["small"], bg=CLR["card"], fg=CLR["subtext"]).pack(side=tk.LEFT, padx=(0, 8))
        tk.Label(hdr, text="名称", width=14, anchor=tk.W,
                  font=FONT["small"], bg=CLR["card"], fg=CLR["subtext"]).pack(side=tk.LEFT)
        tk.Label(hdr, text="说明", anchor=tk.W,
                  font=FONT["small"], bg=CLR["card"], fg=CLR["subtext"]).pack(side=tk.LEFT)

        self.api_container = tk.Frame(list_frame, bg=CLR["card"])
        self.api_container.pack(fill=tk.X, padx=8, pady=(0, 4))

        self.api_rows = {}
        self._temp_order_local = list(self._temp_order)
        self._temp_enabled_local = dict(self._temp_enabled)
        for pid in self._temp_order_local:
            self._create_api_row(pid)

        tk.Label(list_frame, text="提示: 按住行左侧拖拽到目标位置即可交换优先级",
                  font=FONT["small"], bg=CLR["card"], fg=CLR["icon_gray"]).pack(
                  anchor=tk.W, padx=12, pady=(0, 10))

        # 百度密钥
        baidu_frame = tk.Frame(xml_tab, bg=CLR["card"], highlightbackground=CLR["border"],
                               highlightthickness=1, bd=0)
        baidu_frame.pack(fill=tk.X, pady=(2, 8), padx=2)
        tk.Label(baidu_frame, text="\U0001f511  百度翻译密钥", font=FONT["heading"],
                 bg=CLR["card"], fg=CLR["text"]).pack(anchor=tk.W, padx=12, pady=(10, 6))
        tk.Label(baidu_frame, text="注册: fanyi-api.baidu.com", font=FONT["small"],
                 bg=CLR["card"], fg=CLR["subtext"]).pack(anchor=tk.W, padx=12, pady=(0, 8))

        row1 = tk.Frame(baidu_frame, bg=CLR["card"])
        row1.pack(fill=tk.X, padx=12, pady=(0, 6))
        ttk.Label(row1, text="APP ID:", font=FONT["body"]).pack(side=tk.LEFT, padx=(0, 6))
        self.baidu_appid_var = tk.StringVar(value=self._temp_appid)
        ttk.Entry(row1, textvariable=self.baidu_appid_var, width=30).pack(side=tk.LEFT, fill=tk.X, expand=True)

        row2 = tk.Frame(baidu_frame, bg=CLR["card"])
        row2.pack(fill=tk.X, padx=12)
        ttk.Label(row2, text="密  钥:", font=FONT["body"]).pack(side=tk.LEFT, padx=(0, 6))
        self.baidu_secret_var = tk.StringVar(value=self._temp_secret)
        ttk.Entry(row2, textvariable=self.baidu_secret_var, width=30).pack(side=tk.LEFT, fill=tk.X, expand=True)

        row3 = tk.Frame(baidu_frame, bg=CLR["card"])
        row3.pack(fill=tk.X, padx=12, pady=(6, 10))
        self.baidu_test_btn = ttk.Button(row3, text="测试连接",
                                         command=self._test_baidu, width=10)
        self.baidu_test_btn.pack(side=tk.LEFT, padx=(0, 8))
        self.baidu_test_result = tk.StringVar(value="")
        ttk.Label(row3, textvariable=self.baidu_test_result,
                  font=FONT["body"]).pack(side=tk.LEFT)

        # 翻译狗文字翻译密钥
        fy_text_frame = tk.Frame(xml_tab, bg=CLR["card"], highlightbackground=CLR["border"],
                                 highlightthickness=1, bd=0)
        fy_text_frame.pack(fill=tk.X, pady=(2, 8), padx=2)
        tk.Label(fy_text_frame, text="\U0001f511  翻译狗文字翻译密钥", font=FONT["heading"],
                 bg=CLR["card"], fg=CLR["text"]).pack(anchor=tk.W, padx=12, pady=(10, 6))
        tk.Label(fy_text_frame, text="注册: open.fanyigou.com → 创建应用", font=FONT["small"],
                 bg=CLR["card"], fg=CLR["subtext"]).pack(anchor=tk.W, padx=12, pady=(0, 8))

        fy_text_row1 = tk.Frame(fy_text_frame, bg=CLR["card"])
        fy_text_row1.pack(fill=tk.X, padx=12, pady=(0, 6))
        ttk.Label(fy_text_row1, text="APP ID:", font=FONT["body"]).pack(side=tk.LEFT, padx=(0, 6))
        self.fy_text_appid_var = tk.StringVar(value=self._temp_fy_text_appid)
        ttk.Entry(fy_text_row1, textvariable=self.fy_text_appid_var, width=30).pack(side=tk.LEFT, fill=tk.X, expand=True)

        fy_text_row2 = tk.Frame(fy_text_frame, bg=CLR["card"])
        fy_text_row2.pack(fill=tk.X, padx=12, pady=(0, 6))
        ttk.Label(fy_text_row2, text="密  钥:", font=FONT["body"]).pack(side=tk.LEFT, padx=(0, 6))
        self.fy_text_privatekey_var = tk.StringVar(value=self._temp_fy_text_privatekey)
        ttk.Entry(fy_text_row2, textvariable=self.fy_text_privatekey_var, width=30).pack(side=tk.LEFT, fill=tk.X, expand=True)

        fy_text_row3 = tk.Frame(fy_text_frame, bg=CLR["card"])
        fy_text_row3.pack(fill=tk.X, padx=12, pady=(4, 10))
        self.fy_text_test_btn = ttk.Button(fy_text_row3, text="测试连接",
                                           command=self._test_fanyigou_text, width=10)
        self.fy_text_test_btn.pack(side=tk.LEFT, padx=(0, 8))
        self.fy_text_test_result = tk.StringVar(value="")
        ttk.Label(fy_text_row3, textvariable=self.fy_text_test_result,
                  font=FONT["body"]).pack(side=tk.LEFT)


        # 重译选项
        retrans_frame = tk.Frame(xml_tab, bg=CLR["card"], highlightbackground=CLR["border"],
                                 highlightthickness=1, bd=0)
        retrans_frame.pack(fill=tk.X, pady=(2, 8), padx=2)
        tk.Label(retrans_frame, text="\u2699\ufe0f  翻译选项", font=FONT["heading"],
                 bg=CLR["card"], fg=CLR["text"]).pack(anchor=tk.W, padx=12, pady=(10, 6))
        self.force_retranslate_var = tk.BooleanVar(value=self._temp_force_retranslate)
        ttk.Checkbutton(retrans_frame,
                        text="重译已有翻译 (勾选后所有语言将重新翻译，覆盖已有内容)",
                        variable=self.force_retranslate_var).pack(anchor=tk.W, padx=12, pady=(0, 10))

        # ====== PDF 设置 Tab ======
        pdf_tab = tk.Frame(self.notebook, bg=CLR["bg"])
        self.notebook.add(pdf_tab, text="PDF 设置")

        # 翻译狗文档翻译密钥
        fy_frame = tk.Frame(pdf_tab, bg=CLR["card"], highlightbackground=CLR["border"],
                            highlightthickness=1, bd=0)
        fy_frame.pack(fill=tk.X, pady=(2, 8), padx=2)
        tk.Label(fy_frame, text="\U0001f511  翻译狗文档翻译密钥", font=FONT["heading"],
                 bg=CLR["card"], fg=CLR["text"]).pack(anchor=tk.W, padx=12, pady=(10, 6))
        tk.Label(fy_frame, text="注册: open.fanyigou.com", font=FONT["small"],
                 bg=CLR["card"], fg=CLR["subtext"]).pack(anchor=tk.W, padx=12, pady=(0, 8))

        fy_row1 = tk.Frame(fy_frame, bg=CLR["card"])
        fy_row1.pack(fill=tk.X, padx=12, pady=(0, 6))
        ttk.Label(fy_row1, text="APP ID:", font=FONT["body"]).pack(side=tk.LEFT, padx=(0, 6))
        self.fy_appid_var = tk.StringVar(value=self._temp_fy_appid)
        ttk.Entry(fy_row1, textvariable=self.fy_appid_var, width=30).pack(side=tk.LEFT, fill=tk.X, expand=True)

        fy_row2 = tk.Frame(fy_frame, bg=CLR["card"])
        fy_row2.pack(fill=tk.X, padx=12, pady=(0, 6))
        ttk.Label(fy_row2, text="密  钥:", font=FONT["body"]).pack(side=tk.LEFT, padx=(0, 6))
        self.fy_privatekey_var = tk.StringVar(value=self._temp_fy_privatekey)
        ttk.Entry(fy_row2, textvariable=self.fy_privatekey_var, width=30).pack(side=tk.LEFT, fill=tk.X, expand=True)

        fy_row3 = tk.Frame(fy_frame, bg=CLR["card"])
        fy_row3.pack(fill=tk.X, padx=12, pady=(4, 10))
        self.fy_test_btn = ttk.Button(fy_row3, text="测试连接",
                                      command=self._test_fanyigou)
        self.fy_test_btn.pack(side=tk.LEFT, padx=(0, 8))
        self.fy_test_result = tk.StringVar(value="")
        ttk.Label(fy_row3, textvariable=self.fy_test_result,
                  font=FONT["body"]).pack(side=tk.LEFT)

        # ====== 大模型设置 Tab ======
        vc_tab = tk.Frame(self.notebook, bg=CLR["bg"])
        self.notebook.add(vc_tab, text="大模型设置")

        # ---- 配置连接 ----
        vc_frame = tk.Frame(vc_tab, bg=CLR["card"], highlightbackground=CLR["border"],
                            highlightthickness=1, bd=0)
        vc_frame.pack(fill=tk.X, pady=(2, 8), padx=2)
        tk.Label(vc_frame, text="\U0001f916  火山引擎大模型配置 (豆包/方舟)", font=FONT["heading"],
                 bg=CLR["card"], fg=CLR["text"]).pack(anchor=tk.W, padx=12, pady=(10, 8))

        vc_row1 = tk.Frame(vc_frame, bg=CLR["card"])
        vc_row1.pack(fill=tk.X, padx=12, pady=(0, 6))
        ttk.Label(vc_row1, text="API Key:", width=8, font=FONT["body"]).pack(side=tk.LEFT, padx=(0, 6))
        self.vc_api_key_var = tk.StringVar(value=self._temp_vc_api_key)
        ttk.Entry(vc_row1, textvariable=self.vc_api_key_var, width=30).pack(side=tk.LEFT, fill=tk.X, expand=True)

        vc_row2 = tk.Frame(vc_frame, bg=CLR["card"])
        vc_row2.pack(fill=tk.X, padx=12, pady=(0, 6))
        ttk.Label(vc_row2, text="模  型:", width=8, font=FONT["body"]).pack(side=tk.LEFT, padx=(0, 6))
        # 2026最新火山方舟翻译适配模型（按推荐优先级排列）
        MODEL_OPTIONS = [
            # === 字节 Doubao-Seed 2.1 旗舰系列（翻译首选） ===
            "doubao-seed-2-1-pro-260628",
            "doubao-seed-2-1-turbo",          # 均衡性价比
            # === Doubao-Seed 2.0 成熟系列 ===
            "doubao-seed-2-0-pro-260215",
            "doubao-seed-2-0-lite-260428",     # 低成本轻量
            # === 旧代稳定系列 ===
            "doubao-seed-1-8-251228",
            # === 深度求索 DeepSeek 系列 ===
            "deepseek-v4-pro-260425",          # 旗舰100万上下文
            "deepseek-v4-flash-260425",        # 低成本高并发
            # === 智谱 GLM 系列 ===
            "glm-5-2-260617",                  # 100万上下文
            # === MiniMax ===
            "minimax-m3",                      # Agent/长文本
            # === 自定义 ===
            "ep-",                             # Endpoint ID 自定义
        ]
        self.vc_model_var = tk.StringVar(value=self._temp_vc_model)
        self.vc_model_combo = ttk.Combobox(vc_row2, textvariable=self.vc_model_var,
                                           values=MODEL_OPTIONS, font=FONT["mono"])
        self.vc_model_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 模型能力描述映射
        self._MODEL_HELP = {
            "doubao-seed-2-1-pro-260628":  "旗舰通用 — 深度思考·多模态·Agent·超大上下文，翻译质量最佳",
            "doubao-seed-2-1-turbo":        "均衡性价比 — 推理速度与质量兼顾，高并发翻译优选",
            "doubao-seed-2-0-pro-260215":   "2.0旗舰 — 多模态+深度思考，成熟稳定",
            "doubao-seed-2-0-lite-260428":  "低成本轻量 — 快速翻译，适合大批量简单文本",
            "doubao-seed-1-8-251228":       "旧代稳定版 — 久经考验，兼容性好",
            "deepseek-v4-pro-260425":      "DeepSeek旗舰 — 100万上下文，复杂文档/长文本翻译",
            "deepseek-v4-flash-260425":    "DeepSeek轻量 — 低成本高并发，快速翻译",
            "glm-5-2-260617":              "智谱旗舰 — 100万上下文，多语言长文本翻译",
            "minimax-m3":                   "MiniMax均衡 — Agent/长文本/代码兼顾",
            "ep-":                          "自定义 Endpoint ID — 填入你的推理端点",
        }
        vc_model_hint = tk.Label(vc_frame,
            text=self._MODEL_HELP.get(self._temp_vc_model, ""),
            font=FONT["small"], bg=CLR["card"], fg=CLR["subtext"])
        vc_model_hint.pack(anchor=tk.W, padx=12, pady=(2, 6))

        def _on_model_changed(event):
            sel = self.vc_model_var.get()
            hint = self._MODEL_HELP.get(sel, "输入自定义模型 ID")
            vc_model_hint.configure(text=hint)
        self.vc_model_combo.bind("<<ComboboxSelected>>", _on_model_changed)
        self.vc_model_combo.bind("<KeyRelease>", _on_model_changed)

        vc_row3 = tk.Frame(vc_frame, bg=CLR["card"])
        vc_row3.pack(fill=tk.X, padx=12, pady=(2, 10))
        self.vc_test_btn = ttk.Button(vc_row3, text="测试连接",
                                      command=self._test_volcengine, width=10)
        self.vc_test_btn.pack(side=tk.LEFT, padx=(0, 8))
        self.vc_test_result = tk.StringVar(value="")
        ttk.Label(vc_row3, textvariable=self.vc_test_result,
                  font=FONT["body"]).pack(side=tk.LEFT)

        # ---- 简短翻译模式 ----
        vc_concise_frame = tk.Frame(vc_tab, bg=CLR["card"], highlightbackground=CLR["border"],
                                    highlightthickness=1, bd=0)
        vc_concise_frame.pack(fill=tk.X, pady=(2, 8), padx=2)
        tk.Label(vc_concise_frame, text="\u2702\ufe0f  翻译风格", font=FONT["heading"],
                 bg=CLR["card"], fg=CLR["text"]).pack(anchor=tk.W, padx=12, pady=(10, 6))
        self.vc_concise_var = tk.BooleanVar(value=self._temp_vc_concise)
        ttk.Checkbutton(vc_concise_frame,
                        text="简短翻译模式 (长句仅保留核心意思，翻译更简洁)",
                        variable=self.vc_concise_var).pack(anchor=tk.W, padx=12)
        tk.Label(vc_concise_frame,
                  text="适用场景：UI文本、标题、按钮、提示语等需要简短表达的场景",
                  font=FONT["small"], bg=CLR["card"], fg=CLR["subtext"]).pack(
                  anchor=tk.W, padx=12, pady=(4, 10))

        # ---- 按钮 ----
        btn_frame = tk.Frame(self, bg=CLR["bg"])
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=16, pady=(0, 14))
        ttk.Button(btn_frame, text="取消", command=self._cancel,
                   style="Outline.TButton").pack(side=tk.RIGHT)
        ttk.Button(btn_frame, text="保存并关闭", command=self._save_and_close,
                   style="Accent.TButton").pack(side=tk.RIGHT, padx=(0, 8))

    def _create_api_row(self, pid):
        row = tk.Frame(self.api_container, bg=CLR["card"], highlightbackground=CLR["border"],
                       highlightthickness=0, bd=0)
        row.pack(fill=tk.X, pady=2)

        handle = tk.Label(row, text="\u2261", width=2, cursor="fleur",
                          font=("", 12, "bold"), bg=CLR["card"], fg=CLR["icon_gray"])
        handle.pack(side=tk.LEFT, padx=(2, 4))

        var = tk.BooleanVar(value=self._temp_enabled_local.get(pid, True))
        cb = ttk.Checkbutton(row, variable=var, command=lambda p=pid: self._on_toggle(p))
        cb.pack(side=tk.LEFT, padx=(4, 8))

        rank_lbl = tk.Label(row, text="1", width=3, anchor=tk.CENTER,
                            bg=CLR["card"], fg=CLR["primary"], font=FONT["mono"] + ("bold",))
        rank_lbl.pack(side=tk.LEFT, padx=(0, 8))

        name = PROVIDER_REGISTRY[pid]["name"]
        name_lbl = tk.Label(row, text=name, width=14, anchor=tk.W,
                            font=FONT["body"], bg=CLR["card"], fg=CLR["text"])
        name_lbl.pack(side=tk.LEFT)

        desc = API_DESCRIPTIONS.get(pid, "")
        desc_lbl = tk.Label(row, text=desc, font=FONT["small"],
                            bg=CLR["card"], fg=CLR["subtext"], anchor=tk.W)
        desc_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.api_rows[pid] = {
            'frame': row, 'var': var, 'rank': rank_lbl, 'handle': handle,
        }

        for widget in [row, handle, rank_lbl, name_lbl, desc_lbl]:
            widget.bind("<Button-1>", lambda e, p=pid: self._on_drag_start(e, p))
            widget.bind("<B1-Motion>", lambda e: self._on_drag_motion(e))
            widget.bind("<ButtonRelease-1>", lambda e: self._on_drag_release(e))

    def _on_toggle(self, pid):
        self._temp_enabled_local[pid] = self.api_rows[pid]['var'].get()

    def _on_drag_start(self, event, pid):
        self._drag_pid = pid
        self._drag_start_y = event.y_root

    def _on_drag_motion(self, event):
        if not self._drag_pid:
            return
        drop_y = event.y_root
        target = None
        for pid, data in self.api_rows.items():
            frame = data['frame']
            fy = frame.winfo_rooty()
            fh = frame.winfo_height()
            if fy <= drop_y <= fy + fh:
                target = pid
                break
        if self._drop_target_pid and self._drop_target_pid != target:
            old = self.api_rows.get(self._drop_target_pid)
            if old:
                old['frame'].configure(highlightbackground=CLR["border"],
                                       highlightthickness=0)
        self._drop_target_pid = target
        if target and target != self._drag_pid:
            self.api_rows[target]['frame'].configure(highlightbackground=CLR["primary"],
                                                      highlightthickness=1)

    def _on_drag_release(self, event):
        drag_pid = self._drag_pid
        target_pid = self._drop_target_pid

        if target_pid and self.api_rows.get(target_pid):
            self.api_rows[target_pid]['frame'].configure(highlightbackground=CLR["border"],
                                                          highlightthickness=0)

        if drag_pid and target_pid and drag_pid != target_pid:
            idx1 = self._temp_order_local.index(drag_pid)
            idx2 = self._temp_order_local.index(target_pid)
            self._temp_order_local[idx1], self._temp_order_local[idx2] = \
                self._temp_order_local[idx2], self._temp_order_local[idx1]

            for pid in self._temp_order_local:
                self.api_rows[pid]['frame'].pack_forget()
            for pid in self._temp_order_local:
                self.api_rows[pid]['frame'].pack(fill=tk.X, pady=2)

            self._refresh_ranks()

        self._drag_pid = None
        self._drop_target_pid = None

    def _refresh_ranks(self):
        for i, pid in enumerate(self._temp_order_local):
            if pid in self.api_rows:
                self.api_rows[pid]['rank'].config(text=str(i + 1))

    def _save_to_globals(self):
        global provider_order, provider_enabled, baidu_appid, baidu_secret, fanyigou_appid, fanyigou_privatekey, fanyigou_text_appid, fanyigou_text_privatekey, volcengine_api_key, volcengine_model, volcengine_concise, force_retranslate
        for pid in self._temp_order_local:
            self._temp_enabled_local[pid] = self.api_rows[pid]['var'].get()

        provider_order = list(self._temp_order_local)
        provider_enabled = dict(self._temp_enabled_local)
        baidu_appid = self.baidu_appid_var.get().strip()
        baidu_secret = self.baidu_secret_var.get().strip()
        fanyigou_appid = self.fy_appid_var.get().strip()
        fanyigou_privatekey = self.fy_privatekey_var.get().strip()
        fanyigou_text_appid = self.fy_text_appid_var.get().strip()
        fanyigou_text_privatekey = self.fy_text_privatekey_var.get().strip()
        volcengine_api_key = self.vc_api_key_var.get().strip()
        volcengine_model = self.vc_model_var.get().strip()
        volcengine_concise = self.vc_concise_var.get()
        force_retranslate = self.force_retranslate_var.get()

        save_config()

    def _save_and_close(self):
        self._save_to_globals()

        active = [PROVIDER_REGISTRY[pid]["name"] for pid in provider_order
                  if provider_enabled.get(pid, False)]
        self.app._log("[API配置] 优先级: " + " -> ".join(active), "header")

        self.app._check_on_start()
        self.app._refresh_api_summary()
        # 解绑鼠标滚轮
        try:
            self._settings_canvas.unbind_all("<MouseWheel>")
        except Exception:
            pass
        self.destroy()

    def _test_baidu(self):
        appid = self.baidu_appid_var.get().strip()
        secret = self.baidu_secret_var.get().strip()
        if not appid or not secret:
            self.baidu_test_result.set("[APP ID和密钥不能为空]")
            return

        self.baidu_test_btn.config(state=tk.DISABLED, text="测试中...")
        self.baidu_test_result.set("正在测试...")

        def _do():
            global baidu_appid, baidu_secret
            saved_appid, saved_secret = baidu_appid, baidu_secret
            baidu_appid, baidu_secret = appid, secret
            try:
                result = translate_baidu("你好", "en")
                if result:
                    self.after(0, lambda: self.baidu_test_result.set(
                        "[OK] 连接成功! 翻译结果: {}".format(result)))
                else:
                    self.after(0, lambda: self.baidu_test_result.set(
                        "[FAIL] API返回为空,请检查密钥是否正确"))
            except Exception as e:
                self.after(0, lambda: self.baidu_test_result.set(
                    "[FAIL] 连接失败: {}".format(str(e))))
            finally:
                baidu_appid, baidu_secret = saved_appid, saved_secret
                self.after(0, lambda: self.baidu_test_btn.config(
                    state=tk.NORMAL, text="测试连接"))

        t = threading.Thread(target=_do, daemon=True)
        t.start()

    def _test_fanyigou(self):
        """测试 PDF 文档翻译密钥 — 调用 detectDocPage API 发送页数验证"""
        appid = self.fy_appid_var.get().strip()
        privatekey = self.fy_privatekey_var.get().strip()
        if not appid or not privatekey:
            self.fy_test_result.set("[APP ID和密钥不能为空]")
            return

        self.fy_test_btn.config(state=tk.DISABLED, text="测试中...")
        self.fy_test_result.set("正在测试...")

        def _do():
            global fanyigou_appid, fanyigou_privatekey
            saved_appid, saved_key = fanyigou_appid, fanyigou_privatekey
            fanyigou_appid, fanyigou_privatekey = appid, privatekey
            try:
                # 调用 detectDocPage API，发送页数信息测试文档翻译连接
                nonce_str = str(random.randint(1000000000, 9999999999))
                params = {
                    'appid': appid,
                    'nonce_str': nonce_str,
                    'pageCount': '1',  # 发送页数测试
                }
                token = _fanyigou_md5_sign(appid, privatekey, params)
                params['token'] = token
                resp = requests.post(FANYIGOU_DETECT_URL, data=params, timeout=15)
                data = resp.json()
                if data.get('code') in (0, 100):
                    pages = data.get('data', {}).get('pageCount', data.get('data', '未知'))
                    self.after(0, lambda: self.fy_test_result.set(
                        "[OK] 文档翻译连接成功"))
                else:
                    self.after(0, lambda: self.fy_test_result.set(
                        "[FAIL] {}".format(data.get('msg', '请检查APP ID和密钥是否正确'))))
            except Exception as e:
                self.after(0, lambda: self.fy_test_result.set(
                    "[FAIL] 连接失败: {}".format(str(e))))
            finally:
                fanyigou_appid, fanyigou_privatekey = saved_appid, saved_key
                self.after(0, lambda: self.fy_test_btn.config(
                    state=tk.NORMAL, text="测试连接"))

        t = threading.Thread(target=_do, daemon=True)
        t.start()

    def _test_fanyigou_text(self):
        """测试文字翻译密钥"""
        appid = self.fy_text_appid_var.get().strip()
        privatekey = self.fy_text_privatekey_var.get().strip()
        if not appid or not privatekey:
            self.fy_text_test_result.set("[APP ID和密钥不能为空]")
            return

        self.fy_text_test_btn.config(state=tk.DISABLED, text="测试中...")
        self.fy_text_test_result.set("正在测试...")

        def _do():
            global fanyigou_text_appid, fanyigou_text_privatekey
            saved_appid, saved_key = fanyigou_text_appid, fanyigou_text_privatekey
            fanyigou_text_appid, fanyigou_text_privatekey = appid, privatekey
            try:
                result = translate_fanyigou("你好", "en")
                if result:
                    self.after(0, lambda: self.fy_text_test_result.set(
                        "[OK] 连接成功! 翻译结果: {}".format(result)))
                else:
                    self.after(0, lambda: self.fy_text_test_result.set(
                        "[FAIL] API返回为空,请检查APP ID和密钥是否正确"))
            except Exception as e:
                self.after(0, lambda: self.fy_text_test_result.set(
                    "[FAIL] 连接失败: {}".format(str(e))))
            finally:
                fanyigou_text_appid, fanyigou_text_privatekey = saved_appid, saved_key
                self.after(0, lambda: self.fy_text_test_btn.config(
                    state=tk.NORMAL, text="测试连接"))

        t = threading.Thread(target=_do, daemon=True)
        t.start()

    def _test_volcengine(self):
        """测试火山引擎大模型密钥"""
        api_key = self.vc_api_key_var.get().strip()
        model = self.vc_model_var.get().strip()
        if not api_key or not model:
            self.vc_test_result.set("[API Key 和模型不能为空]")
            return

        self.vc_test_btn.config(state=tk.DISABLED, text="测试中...")
        self.vc_test_result.set("正在测试...")

        def _do():
            global volcengine_api_key, volcengine_model
            saved_key, saved_model = volcengine_api_key, volcengine_model
            volcengine_api_key, volcengine_model = api_key, model
            try:
                result = translate_volcengine("你好", "en")
                if result:
                    self.after(0, lambda: self.vc_test_result.set(
                        "[OK] 连接成功! 翻译: {}".format(result[:30])))
                else:
                    self.after(0, lambda: self.vc_test_result.set(
                        "[FAIL] API返回为空, 请检查API Key和模型是否正确"))
            except Exception as e:
                self.after(0, lambda: self.vc_test_result.set(
                    "[FAIL] 连接失败: {}".format(str(e))))
            finally:
                volcengine_api_key, volcengine_model = saved_key, saved_model
                self.after(0, lambda: self.vc_test_btn.config(
                    state=tk.NORMAL, text="测试连接"))

        t = threading.Thread(target=_do, daemon=True)
        t.start()

    def _cancel(self):
        self._save_and_close()


# ============================================================
#  语言选择弹窗
# ============================================================
class LanguageSelectDialog(tk.Toplevel):
    """翻译前弹窗：列出每种缺失语言及数量，勾选确认后开始"""

    def __init__(self, parent, missing_by_lang, force_retranslate=False):
        super().__init__(parent)
        self.title("选择要翻译的语言")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.result = None
        self.force_retranslate = force_retranslate

        nlangs = len(missing_by_lang)
        h = 160 + nlangs * 34
        self.geometry("420x{}".format(h))

        parent.update_idletasks()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        x = px + (pw - 420) // 2
        y = py + (ph - h) // 2
        self.geometry("+{}+{}".format(x, y))

        self._build_ui(missing_by_lang)
        self.wait_window()

    def _build_ui(self, missing_by_lang):
        if self.force_retranslate:
            header_text = "重译模式: 以下语言将全部重新翻译，请选择目标语言："
        else:
            header_text = "以下语言存在缺失翻译，请选择要翻译的语言："
        ttk.Label(self, text=header_text,
                  font=("", 10), padding="15 12").pack()

        ctrl = ttk.Frame(self)
        ctrl.pack(fill=tk.X, padx=20, pady=(0, 6))
        ttk.Button(ctrl, text="全选", command=self._select_all, width=6).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(ctrl, text="取消全选", command=self._deselect_all, width=8).pack(side=tk.LEFT)

        list_frame = ttk.Frame(self)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 8))

        self._check_vars = {}
        for lang_name, cnt in sorted(missing_by_lang.items(), key=lambda x: -x[1]):
            row = ttk.Frame(list_frame)
            row.pack(fill=tk.X, pady=2)

            var = tk.BooleanVar(value=True)
            cb = ttk.Checkbutton(row, variable=var)
            cb.pack(side=tk.LEFT)

            ttk.Label(row, text="{}: ".format(lang_name),
                      font=("", 9, "bold"), width=10, anchor=tk.W).pack(side=tk.LEFT)
            label_text = "{} 条需重译".format(cnt) if self.force_retranslate else "{} 条缺失".format(cnt)
            ttk.Label(row, text=label_text,
                      foreground="#888", font=("", 9)).pack(side=tk.LEFT)

            self._check_vars[lang_name] = var

        btn_frame = ttk.Frame(self)
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=20, pady=(0, 15))
        ttk.Button(btn_frame, text="取消", command=self._on_cancel, width=10).pack(side=tk.RIGHT)
        ttk.Button(btn_frame, text="确认开始翻译", command=self._on_confirm,
                   width=14).pack(side=tk.RIGHT, padx=(0, 8))

    def _select_all(self):
        for var in self._check_vars.values():
            var.set(True)

    def _deselect_all(self):
        for var in self._check_vars.values():
            var.set(False)

    def _on_confirm(self):
        self.result = [name for name, var in self._check_vars.items() if var.get()]
        self.destroy()

    def _on_cancel(self):
        self.result = None
        self.destroy()


# ============================================================
#  主窗口
# ============================================================
class TranslateApp:
    def __init__(self, root):
        self.root = root
        self.root.title("多语言翻译工具")
        self.root.configure(bg=CLR["bg"])
        self.root.geometry("780x880")
        self.root.minsize(520, 580)
        self.root.resizable(True, True)

        self._style = _setup_ttk_style()

        self._app_icon = load_app_icon_tk()
        if self._app_icon:
            self.root.iconphoto(True, self._app_icon)

        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        # 恢复保存的窗口位置/大小，否则居中
        if window_geometry and isinstance(window_geometry, dict):
            wx = window_geometry.get('x', (sw - 780) // 2)
            wy = window_geometry.get('y', max(20, (sh - 880) // 2))
            ww = window_geometry.get('width', 780)
            wh = window_geometry.get('height', 880)
            wx = max(0, min(wx, sw - 200))
            wy = max(0, min(wy, sh - 100))
            ww = max(720, min(ww, sw))
            wh = max(600, min(wh, sh))
            self.root.geometry("{}x{}+{}+{}".format(ww, wh, wx, wy))
        else:
            x = (sw - 780) // 2
            y = max(20, (sh - 880) // 2)
            self.root.geometry("+{}+{}".format(x, y))

        self.msg_queue = queue.Queue()
        self.running = False
        self._stop_requested = False
        self._tray_icon = None
        self._active_translate_btn = None  # 当前翻译按钮引用

        self._build_ui()
        self._poll_queue()

        self.root.protocol("WM_DELETE_WINDOW", self._hide_to_tray)
        self._setup_tray()
        self.root.after(500, self._check_on_start)

    # -------- 系统托盘 --------
    def _setup_tray(self):
        icon_img = self._make_tray_icon()
        if icon_img is None:
            return

        import pystray

        def on_show(icon, item):
            self.root.after(0, self._show_window)

        def on_quit(icon, item):
            icon.stop()
            self.root.after(0, self._quit_app)

        menu = pystray.Menu(
            pystray.MenuItem("显示窗口", on_show, default=True),
            pystray.MenuItem("退出", on_quit),
        )

        self._tray_icon = pystray.Icon(
            "translate_tool", icon_img, "多语言翻译工具", menu)

        tray_thread = threading.Thread(target=self._tray_icon.run, daemon=True)
        tray_thread.start()

    def _make_tray_icon(self):
        img = load_app_icon_pil()
        if img:
            return img
        # 最终兜底（极少触发：无 PIL 且 emoji 渲染失败）
        try:
            from PIL import Image, ImageDraw, ImageFont
        except ImportError:
            return None
        img = Image.new('RGBA', (64, 64), (37, 99, 235, 255))
        draw = ImageDraw.Draw(img)
        try:
            font = ImageFont.truetype("C:/Windows/Fonts/arialbd.ttf", 30)
        except Exception:
            try:
                font = ImageFont.truetype("C:/Windows/Fonts/msyhbd.ttf", 30)
            except Exception:
                font = ImageFont.load_default()
        bbox = draw.textbbox((0, 0), "T", font=font)
        tw = bbox[2] - bbox[0]
        th = bbox[3] - bbox[1]
        tx = (64 - tw) // 2
        ty = (64 - th) // 2 - 2
        draw.text((tx, ty), "T", fill=(255, 255, 255, 255), font=font)
        return img

    def _hide_to_tray(self):
        if self.running:
            messagebox.showwarning("提示", "翻译正在进行中，请先停止翻译。")
            return
        self._sync_ui_to_globals()
        save_config()
        self.root.withdraw()

    def _show_window(self):
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

    def _quit_app(self):
        self._sync_ui_to_globals()
        save_config()
        if self._tray_icon:
            try:
                self._tray_icon.stop()
            except Exception:
                pass
        _release_lock()
        self.root.destroy()

    def _sync_ui_to_globals(self):
        """将当前界面所有勾选和设置同步到全局变量，供 save_config 使用"""
        global pdf_lang_enabled, pdf_mode, window_geometry
        global xml_input_path, xml_output_path, pdf_input_path, pdf_output_path

        # 窗口位置/大小
        try:
            self.root.update_idletasks()
            ww = self.root.winfo_width()
            wh = self.root.winfo_height()
            wx = self.root.winfo_rootx()
            wy = self.root.winfo_rooty()
            if ww > 0 and wh > 0:
                window_geometry = {'x': wx, 'y': wy, 'width': ww, 'height': wh}
        except Exception:
            pass

        # PDF 语言复选框
        pdf_lang_enabled = {}
        for lang_code, var in self.pdf_lang_vars.items():
            pdf_lang_enabled[lang_code] = var.get()

        # PDF 翻译模式
        pdf_mode = self.pdf_mode_var.get()

        # XML/PDF 路径
        xml_input_path = self.xml_input_var.get().strip()
        xml_output_path = self.xml_output_var.get().strip()
        pdf_input_path = self.pdf_input_var.get().strip()
        pdf_output_path = self.pdf_output_var.get().strip()

    # -------- UI 构建 --------
    def _build_ui(self):
        # 主容器（直接pack，去除Canvas/Scrollbar包装）
        main_frame = tk.Frame(self.root, bg=CLR["bg"])
        main_frame.pack(fill=tk.BOTH, expand=False, padx=16, pady=(14, 8))

        # ====== 标题卡片 ======
        header_card = tk.Frame(main_frame, bg=CLR["card"], highlightbackground=CLR["border"],
                               highlightthickness=1, bd=0)
        header_card.pack(fill=tk.X, pady=(0, 10))

        # 左侧蓝色强调条
        accent_bar = tk.Frame(header_card, bg=CLR["primary"], width=4)
        accent_bar.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 14))

        header_inner = tk.Frame(header_card, bg=CLR["card"])
        header_inner.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 12), pady=(12, 8))

        # 第一行：标题 + 齿轮
        title_row = tk.Frame(header_inner, bg=CLR["card"])
        title_row.pack(fill=tk.X)
        tk.Label(title_row, text="多语言翻译工具", font=FONT["title"],
                 bg=CLR["card"], fg=CLR["text"]).pack(side=tk.LEFT)
        # 设置齿轮（最右）
        settings_btn = tk.Label(title_row, text="\u2699", font=("", 20),
                                bg=CLR["card"], fg=CLR["icon_gray"], cursor="hand2")
        settings_btn.pack(side=tk.RIGHT, padx=(0, 4))
        settings_btn.bind("<Button-1>", lambda e: self._open_settings())
        settings_btn.bind("<Enter>", lambda e: settings_btn.config(fg=CLR["primary"]))
        settings_btn.bind("<Leave>", lambda e: settings_btn.config(fg=CLR["icon_gray"]))

        # 第二行：副标题 + 状态信息（紧挨齿轮下面）
        subtitle_row = tk.Frame(header_inner, bg=CLR["card"])
        subtitle_row.pack(fill=tk.X, pady=(2, 0))
        tk.Label(subtitle_row, text="以中文为源  \u2022  支持 XML 及 PDF 文档翻译  \u2022  6 大翻译引擎",
                 font=FONT["small"], bg=CLR["card"], fg=CLR["subtext"]).pack(side=tk.LEFT)

        # 状态信息（网络、翻译源、当前API、进度）放在齿轮旁边
        status_frame = tk.Frame(subtitle_row, bg=CLR["card"])
        status_frame.pack(side=tk.RIGHT)

        self.network_label = tk.Label(status_frame, text="网络: 检测中...",
                                      font=FONT["small"], bg=CLR["card"], fg=CLR["subtext"])
        self.network_label.pack(side=tk.LEFT, padx=(0, 2))

        tk.Label(status_frame, text="\u2022", font=FONT["small"],
                 bg=CLR["card"], fg=CLR["border"]).pack(side=tk.LEFT, padx=4)

        self.api_list_label = tk.Label(status_frame, text="翻译源: 检测中...",
                                       font=FONT["small"], bg=CLR["card"], fg=CLR["subtext"])
        self.api_list_label.pack(side=tk.LEFT, padx=(0, 2))

        tk.Label(status_frame, text="\u2022", font=FONT["small"],
                 bg=CLR["card"], fg=CLR["border"]).pack(side=tk.LEFT, padx=4)

        self.current_api_label = tk.Label(status_frame, text="",
                                          font=FONT["small"], bg=CLR["card"], fg=CLR["primary"])
        self.current_api_label.pack(side=tk.LEFT)

        tk.Label(status_frame, text="\u2022", font=FONT["small"],
                 bg=CLR["card"], fg=CLR["border"]).pack(side=tk.LEFT, padx=4)

        self.progress_var = tk.StringVar(value="就绪")
        self.progress_label = tk.Label(status_frame, textvariable=self.progress_var,
                                       font=FONT["small"], bg=CLR["card"], fg=CLR["subtext"])
        self.progress_label.pack(side=tk.LEFT)

        # ====== API 状态条 ======
        status_card = tk.Frame(main_frame, bg=CLR["card"], highlightbackground=CLR["border"],
                               highlightthickness=1, bd=0)
        status_card.pack(fill=tk.X, pady=(0, 10))

        status_inner = tk.Frame(status_card, bg=CLR["card"])
        status_inner.pack(fill=tk.X, padx=14, pady=(10, 10))

        self._status_dot = tk.Canvas(status_inner, width=10, height=10, bg=CLR["card"],
                                     highlightthickness=0)
        self._status_dot.pack(side=tk.LEFT, padx=(0, 8))
        self._status_oval = self._status_dot.create_oval(1, 1, 9, 9, fill="#94a3b8", outline="")

        self.api_summary_var = tk.StringVar(value=self._get_api_summary())
        tk.Label(status_inner, textvariable=self.api_summary_var,
                 font=FONT["body"], bg=CLR["card"], fg=CLR["text"]).pack(side=tk.LEFT)
        link_lbl = tk.Label(status_inner, text="\u2699 设置 API", cursor="hand2",
                            font=FONT["small"], bg=CLR["card"], fg=CLR["primary"])
        link_lbl.pack(side=tk.RIGHT)
        link_lbl.bind("<Button-1>", lambda e: self._open_settings())

        # ====== 工作区 Notebook ======
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.X, pady=(0, 8))

        self.xml_tab = tk.Frame(self.notebook, bg=CLR["bg"])
        self.notebook.add(self.xml_tab, text="  XML 翻译  ")
        self._build_xml_tab()

        self.pdf_tab = tk.Frame(self.notebook, bg=CLR["bg"])
        self.notebook.add(self.pdf_tab, text="  PDF 翻译  ")
        self._build_pdf_tab()

        # 绑定页面切换事件：自动同步日志区域和翻译模式
        def _on_tab_changed(event):
            idx = self.notebook.index("current")
            if idx == 0:  # XML
                self._current_mode = "xml"
                self._current_log_area = self.xml_log_area
                self.log_area = self.xml_log_area
            else:  # PDF
                self._current_mode = "pdf"
                self._current_log_area = self.pdf_log_area
                self.log_area = self.pdf_log_area
        self.notebook.bind("<<NotebookTabChanged>>", _on_tab_changed)

        # 当前激活的日志源（XML或PDF），用于快速写入
        self._current_log_area = self.xml_log_area
        # 当前翻译模式
        self._current_mode = "xml"
        # 向后兼容引用
        self.log_area = self.xml_log_area

    # ==================== XML 翻译页面 ====================
    def _build_xml_tab(self):
        """构建 XML 翻译标签页 — 卡片式高级风格"""
        pad = {"padx": 12, "pady": (8, 0)}

        # ---- 导入文件卡片 ----
        card1 = tk.Frame(self.xml_tab, bg=CLR["card"], highlightbackground=CLR["border"],
                         highlightthickness=1, bd=0)
        card1.pack(fill=tk.X, **pad)
        tk.Label(card1, text="\U0001f4c4  导入 XML 文档", font=FONT["heading"],
                 bg=CLR["card"], fg=CLR["text"]).pack(anchor=tk.W, padx=12, pady=(10, 6))
        input_row = tk.Frame(card1, bg=CLR["card"])
        input_row.pack(fill=tk.X, padx=12, pady=(0, 10))
        self.xml_input_var = tk.StringVar(value=xml_input_path)
        self.xml_input_entry = ttk.Entry(input_row, textvariable=self.xml_input_var, font=FONT["mono"])
        self.xml_input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(input_row, text="浏览...", command=self._xml_browse_input).pack(side=tk.RIGHT)

        # ---- 输出路径卡片 ----
        card2 = tk.Frame(self.xml_tab, bg=CLR["card"], highlightbackground=CLR["border"],
                         highlightthickness=1, bd=0)
        card2.pack(fill=tk.X, **pad)
        tk.Label(card2, text="\U0001f4be  输出 XML 路径", font=FONT["heading"],
                 bg=CLR["card"], fg=CLR["text"]).pack(anchor=tk.W, padx=12, pady=(10, 6))
        output_row = tk.Frame(card2, bg=CLR["card"])
        output_row.pack(fill=tk.X, padx=12, pady=(0, 10))
        self.xml_output_var = tk.StringVar(value=xml_output_path)
        self.xml_output_entry = ttk.Entry(output_row, textvariable=self.xml_output_var, font=FONT["mono"])
        self.xml_output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(output_row, text="浏览...", command=self._xml_browse_output).pack(side=tk.RIGHT)

        # ---- 操作按钮（在日志上方） ----
        btn_row = tk.Frame(self.xml_tab, bg=CLR["bg"])
        btn_row.pack(fill=tk.X, pady=(12, 0), padx=2)
        ttk.Button(btn_row, text="清除日志", command=self._clear_log_ui,
                   style="Outline.TButton").pack(side=tk.LEFT)
        self.xml_translate_btn = ttk.Button(btn_row, text="开始翻译",
                                            command=self._xml_start_translation, style="Accent.TButton")
        self.xml_translate_btn.pack(side=tk.RIGHT)

        # ---- XML 日志区域 ----
        log_outer = tk.Frame(self.xml_tab, bg=CLR["card"], highlightbackground=CLR["border"],
                             highlightthickness=1, bd=0)
        log_outer.pack(fill=tk.BOTH, expand=True, pady=(6, 0), padx=12)

        self.xml_log_area = scrolledtext.ScrolledText(
            log_outer, wrap=tk.WORD, state=tk.DISABLED,
            font=FONT["mono_sm"], bg="#f8fafc", fg=CLR["text"],
            relief=tk.FLAT, borderwidth=0, padx=10, pady=8,
            insertbackground=CLR["primary"])
        self.xml_log_area.pack(fill=tk.BOTH, expand=True)

        self.xml_log_area.tag_configure("info", foreground=CLR["text"])
        self.xml_log_area.tag_configure("success", foreground=CLR["success"])
        self.xml_log_area.tag_configure("error", foreground=CLR["error"])
        self.xml_log_area.tag_configure("header", foreground=CLR["primary"],
                                        font=FONT["mono_sm"] + ("bold",))

    # ==================== PDF 翻译页面 ====================
    def _build_pdf_tab(self):
        """构建 PDF 翻译标签页 — 卡片式高级风格"""
        pad = {"padx": 12, "pady": (8, 0)}

        # ---- 翻译模式卡片 ----
        card_mode = tk.Frame(self.pdf_tab, bg=CLR["card"], highlightbackground=CLR["border"],
                             highlightthickness=1, bd=0)
        card_mode.pack(fill=tk.X, **pad)
        tk.Label(card_mode, text="\u2699\ufe0f  翻译模式", font=FONT["heading"],
                 bg=CLR["card"], fg=CLR["text"]).pack(anchor=tk.W, padx=12, pady=(10, 6))

        self.pdf_mode_var = tk.StringVar(value=pdf_mode)
        radio_frame = tk.Frame(card_mode, bg=CLR["card"])
        radio_frame.pack(fill=tk.X, padx=12, pady=(0, 10))
        ttk.Radiobutton(radio_frame,
                        text="翻译狗文档翻译  (上传PDF → 直接返回翻译后PDF/Word文档，保持排版)",
                        variable=self.pdf_mode_var, value="fanyigou").pack(anchor=tk.W, pady=(2, 4))
        ttk.Radiobutton(radio_frame,
                        text="文本提取翻译  (在原PDF副本上覆写翻译文字，保持排版)",
                        variable=self.pdf_mode_var, value="text").pack(anchor=tk.W, pady=(2, 2))

        # ---- 源文件卡片 ----
        card_src = tk.Frame(self.pdf_tab, bg=CLR["card"], highlightbackground=CLR["border"],
                            highlightthickness=1, bd=0)
        card_src.pack(fill=tk.X, **pad)
        tk.Label(card_src, text="\U0001f4c4  源 PDF 文件", font=FONT["heading"],
                 bg=CLR["card"], fg=CLR["text"]).pack(anchor=tk.W, padx=12, pady=(10, 6))
        input_row = tk.Frame(card_src, bg=CLR["card"])
        input_row.pack(fill=tk.X, padx=12, pady=(0, 10))
        self.pdf_input_var = tk.StringVar(value=pdf_input_path)
        self.pdf_input_entry = ttk.Entry(input_row, textvariable=self.pdf_input_var, font=FONT["mono"])
        self.pdf_input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(input_row, text="浏览...", command=self._pdf_browse_input).pack(side=tk.RIGHT)

        # ---- 输出目录卡片 ----
        card_out = tk.Frame(self.pdf_tab, bg=CLR["card"], highlightbackground=CLR["border"],
                            highlightthickness=1, bd=0)
        card_out.pack(fill=tk.X, **pad)
        tk.Label(card_out, text="\U0001f4be  输出目录", font=FONT["heading"],
                 bg=CLR["card"], fg=CLR["text"]).pack(anchor=tk.W, padx=12, pady=(10, 6))
        output_row = tk.Frame(card_out, bg=CLR["card"])
        output_row.pack(fill=tk.X, padx=12, pady=(0, 10))
        self.pdf_output_var = tk.StringVar(value=pdf_output_path)
        self.pdf_output_entry = ttk.Entry(output_row, textvariable=self.pdf_output_var, font=FONT["mono"])
        self.pdf_output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(output_row, text="浏览...", command=self._pdf_browse_output).pack(side=tk.RIGHT)

        # ---- 目标语言卡片 ----
        card_lang = tk.Frame(self.pdf_tab, bg=CLR["card"], highlightbackground=CLR["border"],
                             highlightthickness=1, bd=0)
        card_lang.pack(fill=tk.X, **pad)
        tk.Label(card_lang, text="\U0001f310  目标语言", font=FONT["heading"],
                 bg=CLR["card"], fg=CLR["text"]).pack(anchor=tk.W, padx=12, pady=(10, 4))

        lang_ctrl = tk.Frame(card_lang, bg=CLR["card"])
        lang_ctrl.pack(fill=tk.X, padx=12, pady=(0, 4))
        ttk.Button(lang_ctrl, text="全选", command=self._pdf_select_all_langs,
                   style="Outline.TButton").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(lang_ctrl, text="取消全选", command=self._pdf_deselect_all_langs,
                   style="Outline.TButton").pack(side=tk.LEFT)

        self.pdf_lang_vars = {}
        langs_frame = tk.Frame(card_lang, bg=CLR["card"])
        langs_frame.pack(fill=tk.X, padx=12, pady=(2, 10))

        row1 = tk.Frame(langs_frame, bg=CLR["card"])
        row1.pack(fill=tk.X, pady=2)
        row2 = tk.Frame(langs_frame, bg=CLR["card"])
        row2.pack(fill=tk.X, pady=2)
        row3 = tk.Frame(langs_frame, bg=CLR["card"])
        row3.pack(fill=tk.X, pady=2)

        for i, (lang_code, lang_name) in enumerate(PDF_LANG_LIST):
            if i < 4:
                parent = row1
            elif i < 8:
                parent = row2
            else:
                parent = row3
            default_val = pdf_lang_enabled.get(lang_code, True)
            var = tk.BooleanVar(value=default_val)
            cb = ttk.Checkbutton(parent, text=lang_name, variable=var)
            cb.pack(side=tk.LEFT, padx=6)
            self.pdf_lang_vars[lang_code] = var

        # ---- 提示 ----
        self.fy_tip_var = tk.StringVar(value="\u26a0\ufe0f  翻译狗按翻译页数扣费，请确保账户余额充足")
        self.fy_tip_label = tk.Label(self.pdf_tab, textvariable=self.fy_tip_var,
                                     font=FONT["small"], bg=CLR["bg"], fg=CLR["warn"])
        self.fy_tip_label.pack(anchor=tk.W, pady=(6, 0), padx=4)

        # ---- 操作按钮（在日志上方） ----
        btn_row = tk.Frame(self.pdf_tab, bg=CLR["bg"])
        btn_row.pack(fill=tk.X, pady=(12, 0), padx=2)
        ttk.Button(btn_row, text="清除日志", command=self._clear_log_ui,
                   style="Outline.TButton").pack(side=tk.LEFT)
        self.pdf_translate_btn = ttk.Button(btn_row, text="开始翻译",
                                            command=self._pdf_start_translation, style="Accent.TButton")
        self.pdf_translate_btn.pack(side=tk.RIGHT)

        # ---- PDF 日志区域 ----
        log_outer = tk.Frame(self.pdf_tab, bg=CLR["card"], highlightbackground=CLR["border"],
                             highlightthickness=1, bd=0)
        log_outer.pack(fill=tk.BOTH, expand=True, pady=(6, 0), padx=12)

        self.pdf_log_area = scrolledtext.ScrolledText(
            log_outer, wrap=tk.WORD, state=tk.DISABLED,
            font=FONT["mono_sm"], bg="#f8fafc", fg=CLR["text"],
            relief=tk.FLAT, borderwidth=0, padx=10, pady=8,
            insertbackground=CLR["primary"])
        self.pdf_log_area.pack(fill=tk.BOTH, expand=True)

        self.pdf_log_area.tag_configure("info", foreground=CLR["text"])
        self.pdf_log_area.tag_configure("success", foreground=CLR["success"])
        self.pdf_log_area.tag_configure("error", foreground=CLR["error"])
        self.pdf_log_area.tag_configure("header", foreground=CLR["primary"],
                                        font=FONT["mono_sm"] + ("bold",))


    # -------- 设置弹窗 --------
    def _open_settings(self):
        SettingsDialog(self.root, self)

    def _get_api_summary(self):
        active = [PROVIDER_REGISTRY[pid]["name"] for pid in provider_order
                  if provider_enabled.get(pid, False)]
        if active:
            return "当前API: " + " -> ".join(active)
        return "当前API: 无 (请在设置中启用)"

    def _refresh_api_summary(self):
        self.api_summary_var.set(self._get_api_summary())

    # ==================== XML 事件处理 ====================
    def _xml_browse_input(self):
        path = filedialog.askopenfilename(
            title="选择 XML 文档",
            filetypes=[("XML 文件", "*.xml"), ("所有文件", "*.*")])
        if path:
            self.xml_input_var.set(path)
            stem = Path(path).stem
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            out = str(Path(path).parent / "{}_已翻译_{}.xml".format(stem, ts))
            self.xml_output_var.set(out)

    def _xml_browse_output(self):
        path = filedialog.asksaveasfilename(
            title="保存翻译结果",
            defaultextension=".xml",
            filetypes=[("XML 文件", "*.xml"), ("所有文件", "*.*")])
        if path:
            self.xml_output_var.set(path)

    # ==================== PDF 事件处理 ====================
    def _pdf_browse_input(self):
        path = filedialog.askopenfilename(
            title="选择 PDF 文档",
            filetypes=[("PDF 文件", "*.pdf"), ("所有文件", "*.*")])
        if path:
            self.pdf_input_var.set(path)
            stem = Path(path).stem
            out_dir = str(Path(path).parent / "{}_翻译结果".format(stem))
            self.pdf_output_var.set(out_dir)

    def _pdf_browse_output(self):
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.pdf_output_var.set(path)

    def _pdf_select_all_langs(self):
        for var in self.pdf_lang_vars.values():
            var.set(True)

    def _pdf_deselect_all_langs(self):
        for var in self.pdf_lang_vars.values():
            var.set(False)

    # -------- 启动检测 --------
    def _check_on_start(self):
        def _do():
            net_ok, net_msg = check_network()
            if net_ok:
                self.network_label.config(text="网络: " + net_msg, fg=CLR["success"])
                self._status_dot.itemconfig(self._status_oval, fill=CLR["success"])
            else:
                self.network_label.config(text="网络: " + net_msg, fg=CLR["error"])
                self._status_dot.itemconfig(self._status_oval, fill=CLR["error"])

            api_ok, api_msg = check_api()
            if api_ok:
                self.api_list_label.config(text=api_msg, fg=CLR["success"])
            else:
                self.api_list_label.config(text=api_msg, fg=CLR["error"])
        t = threading.Thread(target=_do, daemon=True)
        t.start()

    # ==================== XML 翻译逻辑 ====================
    def _xml_start_translation(self):
        input_path = self.xml_input_var.get().strip()
        output_path = self.xml_output_var.get().strip()

        if not input_path:
            messagebox.showwarning("提示", "请先选择要导入的 XML 文档。")
            return
        if not os.path.exists(input_path):
            messagebox.showerror("错误", "文件不存在:\n{}".format(input_path))
            return
        if not output_path:
            messagebox.showwarning("提示", "请指定输出路径。")
            return
        if not get_active_providers():
            messagebox.showwarning("提示", "请至少启用一个翻译API。\n点击右上角齿轮按钮进行设置。")
            return

        # 切换到XML模式
        self._current_mode = "xml"
        self._current_log_area = self.xml_log_area
        self._clear_log()
        active_names = [n for _, _, n in get_active_providers()]
        self._log("=" * 55, "header")
        self._log("  多语言翻译工具  v12.5  [XML模式]", "header")
        self._log("=" * 55, "header")
        self._log("  翻译API: " + " -> ".join(active_names), "info")
        self._log("")

        self._log("[1/3] 复制源文件到输出路径...")
        try:
            shutil.copy2(input_path, output_path)
            self._log("  => 副本: " + output_path)
        except Exception as e:
            self._log("[ERROR] 复制文件失败: " + str(e), "error")
            return

        self._update_progress("解析XML...")
        self._log("")
        self._log("[2/3] 解析 XML 并分析缺失翻译...")
        try:
            tree, total_entries, cn_count, missing_by_lang, all_tasks = \
                analyze_xml(output_path)
        except Exception as e:
            self._log("[ERROR] XML 解析失败: " + str(e), "error")
            return

        self._log("  共 {} 个 <文本> 条目，其中 {} 条含中文".format(total_entries, cn_count))

        missing_total = sum(missing_by_lang.values())
        if missing_total == 0:
            self._log("")
            self._log("  [OK] 无需翻译，所有语言已完整！", "success")
            self._update_progress("无需翻译")
            return

        for lang_name, cnt in sorted(missing_by_lang.items()):
            self._log("    - {}: {} 条缺失".format(lang_name, cnt))
        self._log("  总计缺失: {} 个翻译".format(missing_total))

        dlg = LanguageSelectDialog(self.root, missing_by_lang, force_retranslate)
        if not dlg.result:
            self._log("")
            self._log("  [取消] 用户取消了翻译操作", "error")
            self._update_progress("已取消")
            return
        if len(dlg.result) == 0:
            messagebox.showwarning("提示", "请至少选择一种语言。")
            self._update_progress("已取消")
            return

        selected_set = set(dlg.result)
        tasks = [t for t in all_tasks if t[5] in selected_set]
        selected_total = len(tasks)

        selected_names = ", ".join(sorted(dlg.result))
        self._log("")
        self._log("  已选择: {}".format(selected_names), "info")
        self._log("  待翻译条数: {}".format(selected_total), "info")

        self.running = True
        self._stop_requested = False
        self._translation_start_time = datetime.now().strftime('%Y%m%d_%H%M%S')
        self._active_translate_btn = self.xml_translate_btn
        self.xml_translate_btn.config(text="停止翻译", command=self._xml_stop_translation)

        t = threading.Thread(target=self._xml_run_translation,
                             args=(output_path, tree, tasks, selected_total),
                             daemon=True)
        t.start()

    def _xml_stop_translation(self):
        self._stop_requested = True
        if self._active_translate_btn:
            self._active_translate_btn.config(state=tk.DISABLED, text="停止中...")
        self._log("")
        self._log("  [STOP] 正在停止翻译...", "error")

    def _xml_reset_button(self):
        self.xml_translate_btn.config(
            state=tk.NORMAL, text="开始翻译", command=self._xml_start_translation)

    def _xml_run_translation(self, output_path, tree, tasks, selected_total):
        try:
            self._log("")
            self._log("[3/3] 开始翻译 (共 {} 条)...".format(selected_total))
            self._xml_do_translate(tasks, selected_total, tree, output_path)

            if self._stop_requested:
                self._log("")
                self._log("=" * 55, "header")
                self._log("  [STOPPED] 翻译已停止 (已保存已翻译部分)", "error")
                self._log("=" * 55, "header")
                self._update_progress("已停止")
            else:
                self._update_current_api("")
                self._log("")
                self._log("=" * 55, "header")
                self._log("  [DONE] 翻译完成！", "success")
                self._log("  输出: " + os.path.abspath(output_path), "success")
                self._log("=" * 55, "header")
                self._update_progress("完成")

        except Exception as e:
            self._log("")
            self._log("[ERROR] " + str(e), "error")
            import traceback
            self._log(traceback.format_exc(), "error")
            self._update_progress("出错")
            self._update_current_api("")
        finally:
            self.running = False
            self._stop_requested = False
            self._active_translate_btn = None
            self.root.after(0, self._xml_reset_button)
            self._save_log_to_file()

    def _xml_do_translate(self, tasks, total, tree, output_path):
        by_lang = defaultdict(list)
        for t in tasks:
            by_lang[t[4]].append(t)

        done = 0
        last_save = 0
        SAVE_INTERVAL = 50

        for lang_code, group in sorted(by_lang.items()):
            lang_name = group[0][5]
            cnt = len(group)
            self._log("", "info")
            self._log("--- [{}] {} 条 ---".format(lang_name, cnt), "header")

            for text_elem, cn_text, tag, lang_elem, lc, ln in group:
                if self._stop_requested:
                    tree.write(output_path, encoding='utf-8', xml_declaration=True)
                    self._log("  => 已保存 {} 条 (已停止)".format(done), "error")
                    return

                result, provider, fail_reason = translate_text(cn_text, lc)

                if result:
                    lang_elem.text = result
                    label = "[{}] - [{}]".format(cn_text, result)
                    if provider:
                        label += "  (via {})".format(provider)
                    self._log("  " + label, "info")
                    self._update_current_api(provider or "")
                else:
                    self._log("  [{}] - [FAIL] {}".format(cn_text, fail_reason), "error")
                    self._update_current_api("")

                done += 1
                self._update_progress("{}/{} ({}%)".format(
                    done, total, int(done * 100 / total)))

                if done - last_save >= SAVE_INTERVAL:
                    tree.write(output_path, encoding='utf-8', xml_declaration=True)
                    last_save = done

                time.sleep(TRANSLATE_DELAY)

        tree.write(output_path, encoding='utf-8', xml_declaration=True)
        self._log("")
        self._log("  => 已保存，共更新 {} 个翻译".format(done), "success")

    # ==================== PDF 翻译逻辑 ====================
    def _pdf_start_translation(self):
        """根据翻译模式分发到文本提取或翻译狗文档翻译"""
        input_path = self.pdf_input_var.get().strip()
        output_dir = self.pdf_output_var.get().strip()

        if not input_path:
            messagebox.showwarning("提示", "请先选择要翻译的 PDF 文档。")
            return
        if not os.path.exists(input_path):
            messagebox.showerror("错误", "文件不存在:\n{}".format(input_path))
            return
        if not output_dir:
            messagebox.showwarning("提示", "请指定输出目录。")
            return

        # 收集选中的语言
        selected_langs = []
        for lang_code, lang_name in PDF_LANG_LIST:
            var = self.pdf_lang_vars.get(lang_code)
            if var and var.get():
                selected_langs.append((lang_code, lang_name))

        if not selected_langs:
            messagebox.showwarning("提示", "请至少选择一种目标语言。")
            return

        mode = self.pdf_mode_var.get()

        # 翻译狗模式：需要翻译狗已配置
        if mode == "fanyigou":
            if not fanyigou_appid.strip() or not fanyigou_privatekey.strip():
                messagebox.showwarning("提示",
                    "翻译狗文档翻译需要配置APP ID和密钥。\n"
                    "请点击右上角齿轮按钮 → 选择「PDF 设置」页签填写。")
                return
            self._pdf_run_fanyigou_translation(input_path, output_dir, selected_langs)
            return

        # 文本提取模式
        if not get_active_providers():
            messagebox.showwarning("提示", "请至少启用一个翻译API。\n点击右上角齿轮按钮进行设置。")
            return

        # 切换到PDF模式
        self._current_mode = "pdf"
        self._current_log_area = self.pdf_log_area
        self._clear_log()
        active_names = [n for _, _, n in get_active_providers()]
        self._log("=" * 55, "header")
        self._log("  多语言翻译工具  v12.5  [PDF文本提取模式]", "header")
        self._log("=" * 55, "header")
        self._log("  翻译API: " + " -> ".join(active_names), "info")
        self._log("  源文件: " + input_path, "info")
        self._log("  输出目录: " + output_dir, "info")
        lang_names = ", ".join(name for _, name in selected_langs)
        self._log("  目标语言: " + lang_names, "info")
        self._log("")

        self.running = True
        self._stop_requested = False
        self._translation_start_time = datetime.now().strftime('%Y%m%d_%H%M%S')
        self._active_translate_btn = self.pdf_translate_btn
        self.pdf_translate_btn.config(text="停止翻译", command=self._pdf_stop_translation)

        t = threading.Thread(target=self._pdf_run_text_translation,
                             args=(input_path, output_dir, selected_langs),
                             daemon=True)
        t.start()

    def _pdf_run_text_translation(self, input_path, output_dir, selected_langs):
        """文本提取翻译模式：在原PDF副本上覆写翻译文字，保持排版"""
        try:
            self._log("[1/3] 正在打开 PDF 文档...")
            try:
                import fitz
                doc = fitz.open(input_path)
            except Exception as e:
                self._log("[ERROR] 无法打开 PDF: " + str(e), "error")
                self._update_progress("打开失败")
                return

            total_pages = len(doc)
            if total_pages == 0:
                self._log("[ERROR] PDF 文档为空", "error")
                self._update_progress("空文档")
                doc.close()
                return

            # 收集所有页面中需要翻译的文字块
            # 多种提取方式：blocks → words → dict/spans → pdfplumber后备
            self._log("  => 共 {} 页，正在扫描文字块...".format(total_pages))
            all_blocks = []  # [(page_idx, bbox, text)]
            total_chars = 0

            for pi in range(total_pages):
                page = doc[pi]
                page_blocks = []

                # 方法1: get_text("blocks") — 标准文字块
                blocks = page.get_text("blocks")
                for block in blocks:
                    x0, y0, x1, y1, text, block_no, block_type = block
                    text = text.strip()
                    if block_type == 0 and text and len(text) > 1:
                        has_chinese = any('\u4e00' <= c <= '\u9fff' for c in text)
                        if has_chinese:
                            page_blocks.append(((x0, y0, x1, y1), text))

                # 方法2: get_text("words") — 单字级提取（图文混排PDF常用）
                if not page_blocks:
                    words = page.get_text("words")
                    if words:
                        # 按Y坐标聚类（同行的word合并，间距<字符宽度3倍视为同一块）
                        clustered = _cluster_words(words)
                        for bbox, text in clustered:
                            has_chinese = any('\u4e00' <= c <= '\u9fff' for c in text)
                            if has_chinese and len(text.strip()) > 1:
                                page_blocks.append((bbox, text.strip()))

                # 方法3: get_text("dict") → spans级别
                if not page_blocks:
                    td = page.get_text("dict")
                    spans_list = []
                    for block in td.get("blocks", []):
                        if block.get("type") == 0:  # text block
                            for line in block.get("lines", []):
                                for span in line.get("spans", []):
                                    bbox = span["bbox"]
                                    text = span["text"].strip()
                                    if text:
                                        spans_list.append((bbox, text))
                    if spans_list:
                        # 合并相邻span
                        merged = _merge_spans(spans_list)
                        for bbox, text in merged:
                            has_chinese = any('\u4e00' <= c <= '\u9fff' for c in text)
                            if has_chinese and len(text.strip()) > 1:
                                page_blocks.append((bbox, text.strip()))

                # 方法4: pdfplumber 后备
                if not page_blocks:
                    try:
                        import pdfplumber
                        with pdfplumber.open(input_path) as ppdf:
                            if pi < len(ppdf.pages):
                                ppage = ppdf.pages[pi]
                                pwords = ppage.extract_words(
                                    extra_attrs=["size", "fontname"])
                                if pwords:
                                    pw_list = []
                                    for w in pwords:
                                        pw_list.append((w["x0"], w["top"],
                                                        w["x1"], w["bottom"],
                                                        w["text"]))
                                    clustered = _cluster_pdfplumber_words(pw_list)
                                    for bbox, text in clustered:
                                        has_chinese = any(
                                            '\u4e00' <= c <= '\u9fff' for c in text)
                                        if has_chinese and len(text.strip()) > 1:
                                            page_blocks.append((bbox, text.strip()))
                        self._log("  => 第 {} 页: pdfplumber 提取 {} 块".format(
                            pi + 1, len(page_blocks)), "info")
                    except Exception:
                        pass

                # 方法5: EasyOCR — 图片型/扫描件PDF
                if not page_blocks:
                    page_blocks = _ocr_extract_page(page, pi, self._log)

                for bbox, text in page_blocks:
                    all_blocks.append((pi, bbox, text))
                    total_chars += len(text)

            if not all_blocks:
                self._log("[ERROR] PDF 中未检测到可翻译文字", "error")
                self._log("  提示：如是扫描件图片型PDF，请先用OCR工具识别为可编辑PDF", "info")
                self._update_progress("无文字")
                doc.close()
                return

            self._log("  => 检测到 {} 个文字块，共 {} 字符".format(len(all_blocks), total_chars))
            self._log("")

            os.makedirs(output_dir, exist_ok=True)

            # 逐语言翻译并生成PDF副本
            total_tasks = len(all_blocks) * len(selected_langs)
            done = 0

            self._log("[2/3] 开始翻译并生成 PDF 副本 (共 {} 种语言)...".format(
                len(selected_langs)))
            self._log("")

            for lang_code, lang_name in selected_langs:
                if self._stop_requested:
                    self._log("  => 翻译已停止", "error")
                    break

                self._log("--- [{}] 翻译中... ---".format(lang_name), "header")

                # 翻译所有文字块
                translations = []
                for pi, bbox, text in all_blocks:
                    if self._stop_requested:
                        translations.append(None)
                        continue
                    result, provider, fail_reason = translate_text(text, lang_code)
                    translations.append(result)
                    done += 1
                    if result:
                        preview = text[:25].replace('\n', ' ')
                        self._log("  [{}] → [{}] (via {})".format(
                            preview, result[:25].replace('\n', ' '), provider or "-"))
                    else:
                        preview = text[:25].replace('\n', ' ')
                        self._log("  [{}] - [FAIL] {}".format(preview, fail_reason), "error")
                    self._update_progress("{}/{} ({}%)".format(
                        done, total_tasks, int(done * 100 / total_tasks)))
                    time.sleep(TRANSLATE_DELAY)

                # 生成该语言的PDF副本
                if not self._stop_requested:
                    self._log("")
                    self._log("  [3/3] 正在生成 {} 的 PDF 副本...".format(lang_name))
                    output_pdf = os.path.join(output_dir,
                        "{}_{}.pdf".format(
                            os.path.splitext(os.path.basename(input_path))[0],
                            lang_name))
                    try:
                        self._create_translated_pdf(
                            input_path, output_pdf, all_blocks, translations)
                        self._log("  => 已保存: {}".format(output_pdf), "success")
                    except Exception as e:
                        self._log("  [ERROR] PDF生成失败: {}".format(str(e)), "error")
                    self._log("")

            doc.close()
            self._update_current_api("")

            if self._stop_requested:
                self._log("=" * 55, "header")
                self._log("  [STOPPED] 翻译已停止", "error")
                self._log("=" * 55, "header")
                self._update_progress("已停止")
            else:
                self._log("=" * 55, "header")
                self._log("  [DONE] 翻译完成！", "success")
                self._log("  输出目录: " + os.path.abspath(output_dir), "success")
                self._log("=" * 55, "header")
                self._update_progress("完成")

        except Exception as e:
            self._log("")
            self._log("[ERROR] " + str(e), "error")
            import traceback
            self._log(traceback.format_exc(), "error")
            self._update_progress("出错")
            self._update_current_api("")
        finally:
            self.running = False
            self._stop_requested = False
            self._active_translate_btn = None
            self.root.after(0, self._pdf_reset_button)
            self._save_log_to_file()

    def _create_translated_pdf(self, src_pdf, dst_pdf, all_blocks, translations):
        """创建翻译后的PDF副本：在原文字位置覆盖白底并写入译文。
        优化：按行合并白底（减少遮罩数量），同行内相邻文字块统一翻译渲染。"""
        import fitz
        doc = fitz.open(src_pdf)

        # 按页分组
        page_blocks = defaultdict(list)
        for idx, (pi, bbox, text) in enumerate(all_blocks):
            page_blocks[pi].append((idx, bbox, text, translations[idx]))

        success_count = 0
        fail_count = 0
        overflow_count = 0

        def _group_into_rows(blocks, y_tol=8.0):
            """将blocks按Y坐标分组成行"""
            if not blocks:
                return []
            # 先按Y中心排序
            sorted_blocks = []
            for item in blocks:
                idx, bbox, src, tr = item
                if not tr:
                    continue
                cy = (bbox[1] + bbox[3]) / 2.0
                sorted_blocks.append((cy, item))
            sorted_blocks.sort(key=lambda x: (x[0], x[1][1][0]))

            rows = []
            for cy, item in sorted_blocks:
                placed = False
                for row in rows:
                    row_cy = row['_cy']
                    if abs(cy - row_cy) <= y_tol:
                        row['items'].append(item)
                        # 更新行的平均cy
                        total_cy = sum((it[1][1] + it[1][3]) / 2.0 for it in row['items'])
                        row['_cy'] = total_cy / len(row['items'])
                        placed = True
                        break
                if not placed:
                    rows.append({'_cy': cy, 'items': [item]})
            return [r['items'] for r in rows]

        def _merge_row_items(row_items, x_gap=15.0):
            """同一行内，X方向相邻（间距<x_gap）的blocks合并成段"""
            if not row_items:
                return []
            # 按X坐标排序
            sorted_items = sorted(row_items, key=lambda it: it[1][0])
            segments = []
            cur_seg = [sorted_items[0]]
            for item in sorted_items[1:]:
                prev_x1 = cur_seg[-1][1][2]
                cur_x0 = item[1][0]
                if cur_x0 - prev_x1 <= x_gap:
                    cur_seg.append(item)
                else:
                    segments.append(cur_seg)
                    cur_seg = [item]
            segments.append(cur_seg)
            return segments

        for pi in range(len(doc)):
            page = doc[pi]
            if pi not in page_blocks:
                continue

            # 清理页面内容流（修复图片型PDF的渲染层级问题）
            try:
                page.clean_contents()
            except Exception:
                pass

            # 按行分组
            rows = _group_into_rows(page_blocks[pi])

            for row in rows:
                # 行内再按X相邻性合并成段
                segments = _merge_row_items(row)

                for seg in segments:
                    # 合并段内所有文字
                    seg_src = ' '.join(it[2] for it in seg)
                    seg_tr = ' '.join(it[3] for it in seg)
                    if not seg_tr:
                        continue

                    # 计算段的合并bbox
                    sx0 = min(it[1][0] for it in seg)
                    sy0 = min(it[1][1] for it in seg)
                    sx1 = max(it[1][2] for it in seg)
                    sy1 = max(it[1][3] for it in seg)
                    s_bw = sx1 - sx0
                    s_bh = sy1 - sy0

                    # 判断译文是否含中文，选择合适的字体
                    has_cjk = any('\u4e00' <= c <= '\u9fff' or
                                  '\u3040' <= c <= '\u30ff' or
                                  '\uac00' <= c <= '\ud7af'
                                  for c in seg_tr)
                    font_name = "china-s" if has_cjk else "helv"

                    # 1) 画一个覆盖整段的细长白底
                    PAD_X = 2.0
                    PAD_Y = 1.0
                    cover_rect = fitz.Rect(
                        max(0, sx0 - PAD_X), max(0, sy0 - PAD_Y),
                        min(page.rect.width, sx1 + PAD_X),
                        min(page.rect.height, sy1 + PAD_Y))
                    shape = page.new_shape()
                    shape.draw_rect(cover_rect)
                    # 高不透明白底：fill_opacity=0.97 几乎完全遮盖原文
                    shape.finish(fill=(1, 1, 1), fill_opacity=0.97,
                                 color=(1, 1, 1), width=0, stroke_opacity=0)
                    shape.commit()

                    # 2) 计算字体大小（以段的高度为基准）
                    font_size = min(s_bh * 0.75, 11.0)
                    if font_size < 5.0:
                        font_size = 5.0

                    # 估算译文所需宽度
                    if has_cjk:
                        est_cn = sum(1 for c in seg_tr if '\u4e00' <= c <= '\u9fff')
                        est_en = len(seg_tr) - est_cn
                        est_width = est_cn * font_size + est_en * font_size * 0.55
                    else:
                        est_width = len(seg_tr) * font_size * 0.55
                    if s_bw > 0 and est_width > s_bw * 1.2:
                        new_fs = font_size * s_bw * 1.2 / max(est_width, 1)
                        font_size = max(4.0, min(font_size, new_fs))

                    # 3) 写入翻译文字（统一插入整段）
                    text_rect = fitz.Rect(
                        cover_rect.x0 + 1, cover_rect.y0 + 1,
                        cover_rect.x1 - 1, cover_rect.y1 - 1)
                    try:
                        rc = page.insert_textbox(text_rect, seg_tr,
                                                 fontname=font_name,
                                                 fontsize=font_size,
                                                 align=0,
                                                 color=(0, 0, 0))
                        if rc < 0 and font_size > 4.0:
                            overflow_count += 1
                            new_fs = max(4.0, font_size * 0.7)
                            try:
                                page.insert_textbox(text_rect, seg_tr,
                                                    fontname=font_name,
                                                    fontsize=new_fs,
                                                    align=0,
                                                    color=(0, 0, 0))
                                success_count += 1
                            except Exception as e2:
                                self._log("  [WARN] 文字段重试失败 ({}): {}".format(
                                    seg_src[:30], str(e2)), "error")
                                fail_count += 1
                        else:
                            success_count += 1
                    except Exception as e:
                        fail_count += 1
                        try:
                            page.insert_text(
                                fitz.Point(text_rect.x0 + 2,
                                           text_rect.y0 + font_size),
                                seg_tr,
                                fontname=font_name,
                                fontsize=font_size,
                                color=(0, 0, 0))
                            success_count += 1
                            fail_count -= 1
                        except Exception:
                            self._log("  [WARN] 文字段写入失败 ({}): {}".format(
                                seg_src[:30], str(e)), "error")

        self._log("  PDF生成统计: 成功 {} / 溢出 {} / 失败 {}".format(
            success_count, overflow_count, fail_count), "info")

        doc.save(dst_pdf, garbage=4, deflate=True)
        doc.close()

    def _pdf_stop_translation(self):
        self._stop_requested = True
        if self._active_translate_btn:
            self._active_translate_btn.config(state=tk.DISABLED, text="停止中...")
        self._log("")
        self._log("  [STOP] 正在停止翻译...", "error")

    def _pdf_reset_button(self):
        self.pdf_translate_btn.config(
            state=tk.NORMAL, text="开始翻译", command=self._pdf_start_translation)

    # ==================== 翻译狗文档翻译 ====================
    def _pdf_run_fanyigou_translation(self, input_path, output_dir, selected_langs):
        """翻译狗文档翻译: 上传PDF → 轮询进度 → 下载译文"""
        # 切换到PDF模式
        self._current_mode = "pdf"
        self._current_log_area = self.pdf_log_area
        self._clear_log()
        self._log("=" * 55, "header")
        self._log("  多语言翻译工具  v12.5  [翻译狗文档翻译模式]", "header")
        self._log("=" * 55, "header")
        self._log("  翻译API: 翻译狗 (文档翻译)", "info")
        self._log("  源文件: " + input_path, "info")
        self._log("  输出目录: " + output_dir, "info")
        lang_names = ", ".join(name for _, name in selected_langs)
        self._log("  目标语言: " + lang_names, "info")
        self._log("  提示: 翻译狗按页扣费，翻译完成后从主账户扣除", "info")
        self._log("")

        self.running = True
        self._stop_requested = False
        self._translation_start_time = datetime.now().strftime('%Y%m%d_%H%M%S')
        self._active_translate_btn = self.pdf_translate_btn
        self.pdf_translate_btn.config(text="停止翻译", command=self._pdf_stop_translation)

        t = threading.Thread(target=self._pdf_run_fanyigou_thread,
                             args=(input_path, output_dir, selected_langs),
                             daemon=True)
        t.start()

    def _pdf_run_fanyigou_thread(self, input_path, output_dir, selected_langs):
        """翻译狗文档翻译后台线程"""
        try:
            os.makedirs(output_dir, exist_ok=True)

            total = len(selected_langs)
            completed = 0

            for idx, (lang_code, lang_name) in enumerate(selected_langs):
                if self._stop_requested:
                    self._log("  => 翻译已停止", "error")
                    break

                self._log("--- [{}] ({}/{}) ---".format(lang_name, idx + 1, total), "header")

                # Step 1: 上传文档
                self._update_progress("[{}] 上传中...".format(lang_name))
                self._log("  [1/3] 正在上传文档到翻译狗...")
                tid, error = fanyigou_upload_translate(input_path, lang_code)
                if error:
                    self._log("  [ERROR] 上传失败: {}".format(error), "error")
                    continue
                self._log("  => 上传成功, 翻译ID: {}".format(tid))

                # Step 2: 轮询进度
                self._update_progress("[{}] 翻译中...".format(lang_name))
                self._log("  [2/3] 等待翻译完成 (每{}秒查询一次)...".format(POLL_INTERVAL))
                max_wait = 900  # 最多等15分钟
                waited = 0
                translated = False

                while waited < max_wait:
                    if self._stop_requested:
                        break
                    time.sleep(POLL_INTERVAL)
                    waited += POLL_INTERVAL

                    result = fanyigou_query_progress(tid)
                    if result.get('code') == 100:
                        data = result.get('data', {})
                        status = data.get('status', 0)
                        percent = data.get('percent', '0')
                        pages = data.get('pageCount', '?')
                        msg = data.get('msg', '')
                        self._log("    进度: {}% | 状态: {} | 页数: {} | {}".format(
                            percent, status, pages, msg))

                        if status == 314:  # 翻译完成
                            translated = True
                            self._log("  => 翻译完成 ({}页, 耗时约{}秒)".format(pages, waited))
                            break
                    else:
                        err_msg = result.get('msg', '未知错误')
                        self._log("    查询失败: [{}] {}".format(result.get('code'), err_msg))
                        # 查询失败不中断，继续重试

                if self._stop_requested:
                    continue

                if not translated:
                    self._log("  [ERROR] 翻译超时 (已等待{}秒)".format(max_wait), "error")
                    continue

                # Step 3: 下载译文
                self._update_progress("[{}] 下载中...".format(lang_name))
                self._log("  [3/3] 正在下载译文...")

                ext = os.path.splitext(input_path)[1].lower()
                if ext == '.pdf':
                    dtype = 2  # 输出PDF
                else:
                    dtype = 3  # 输出Word/PPT/Excel

                base_name = Path(input_path).stem
                output_file = os.path.join(output_dir, "{}_{}{}".format(base_name, lang_name, ext))
                success, error = fanyigou_download_file(tid, dtype, output_file)
                if success:
                    self._log("  => 下载完成: {}".format(output_file), "success")
                    completed += 1
                else:
                    self._log("  [ERROR] 下载失败: {}".format(error), "error")

                self._update_progress("{}/{} 完成".format(completed, total))
                self._log("")

            self._update_current_api("")
            self._log("=" * 55, "header")
            if self._stop_requested:
                self._log("  [STOPPED] 翻译已停止 (已完成 {}/{} 个语言)".format(completed, total), "error")
            else:
                self._log("  [DONE] 翻译完成! 成功: {}/{} 个语言".format(completed, total), "success")
            self._log("  输出目录: " + os.path.abspath(output_dir), "success")
            self._log("=" * 55, "header")
            self._update_progress("完成" if not self._stop_requested else "已停止")

        except Exception as e:
            self._log("")
            self._log("[ERROR] " + str(e), "error")
            import traceback
            self._log(traceback.format_exc(), "error")
            self._update_progress("出错")
            self._update_current_api("")
        finally:
            self.running = False
            self._stop_requested = False
            self._active_translate_btn = None
            self.root.after(0, self._pdf_reset_button)
            self._save_log_to_file()

    # -------- 日志保存 --------
    def _save_log_to_file(self):
        try:
            ts = getattr(self, '_translation_start_time', datetime.now().strftime('%Y%m%d_%H%M%S'))
            log_dir = os.path.join(get_app_dir(), "日志")
            os.makedirs(log_dir, exist_ok=True)
            mode = self._current_mode
            prefix = "xml" if mode == "xml" else "pdf"
            log_path = os.path.join(log_dir, "log_{}_{}.txt".format(prefix, ts))
            text = self._current_log_area.get("1.0", tk.END).strip()
            with open(log_path, 'w', encoding='utf-8') as f:
                f.write(text)
        except Exception:
            pass

    # -------- 状态栏更新（共享） --------
    def _update_current_api(self, name):
        self.root.after(0, lambda n=name: self._set_current_api(n))

    def _set_current_api(self, name):
        if name:
            self.current_api_label.config(
                text="正在使用: " + name, foreground="#E65100")
        else:
            self.current_api_label.config(text="", foreground="#888888")

    # -------- 日志 & 进度（共享） --------
    def _log(self, text, tag="info"):
        self.msg_queue.put(("log", text, tag))

    def _update_progress(self, text):
        self.msg_queue.put(("progress", text, None))

    def _clear_log(self):
        self.msg_queue.put(("clear", None, None))

    def _clear_log_ui(self):
        """清除当前激活的日志区域（UI线程直接操作 + 消息队列同步）"""
        target = self._current_log_area
        target.config(state=tk.NORMAL)
        target.delete("1.0", tk.END)
        target.config(state=tk.DISABLED)
        self.msg_queue.put(("clear", None, None))

    def _poll_queue(self):
        try:
            while True:
                msg_type, text, tag = self.msg_queue.get_nowait()
                if msg_type == "clear":
                    target = self._current_log_area
                    target.config(state=tk.NORMAL)
                    target.delete("1.0", tk.END)
                    target.config(state=tk.DISABLED)
                elif msg_type == "progress":
                    self.progress_var.set(text)
                elif msg_type == "log":
                    target = self._current_log_area
                    target.config(state=tk.NORMAL)
                    if tag:
                        target.insert(tk.END, text + "\n", tag)
                    else:
                        target.insert(tk.END, text + "\n")
                    target.see(tk.END)
                    target.config(state=tk.DISABLED)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)


# ============================================================
#  入口
# ============================================================
def main():
    # 启用 DPI 感知，避免高分屏下 tkinter 字体被系统缩放而模糊
    if sys.platform == "win32":
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass

    if not _check_single_instance():
        import tkinter.messagebox as _mb
        root_check = tk.Tk()
        root_check.withdraw()
        _mb.showwarning("提示", "软件已打开，请检查系统托盘或任务栏。")
        root_check.destroy()
        sys.exit(0)

    root = tk.Tk()
    app = TranslateApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
