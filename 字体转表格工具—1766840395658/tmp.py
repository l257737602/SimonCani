#!/usr/bin/env python3
"""
字体转表格工具
读取字体文件，提取已定义字形的Unicode编码和对应字符，导出为表格文件
支持多种字体格式和表格格式
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter import font as tkfont
import os
from pathlib import Path
import threading
import traceback
from datetime import datetime
import unicodedata

# 第三方库导入检查
try:
    from fontTools.ttLib import TTFont
    FONTTOOLS_AVAILABLE = True
except ImportError:
    FONTTOOLS_AVAILABLE = False
    print("警告: fontTools库未安装，请运行: pip install fonttools")

try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False
    print("警告: pandas库未安装，请运行: pip install pandas")

try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = True  # 非必需，仅用于Excel格式

class FontToTableApp:
    def __init__(self, root):
        self.root = root
        self.root.title("字体转表格工具")
        self.root.geometry("750x550")
        self.root.resizable(True, True)
        
        # 设置窗口图标（如果有）
        try:
            self.root.iconbitmap(default='')
        except:
            pass
        
        # 初始化变量
        self.font_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.table_format = tk.StringVar(value="csv")
        self.status_var = tk.StringVar(value="就绪")
        self.progress_var = tk.DoubleVar(value=0)
        
        # 支持的格式
        self.supported_font_formats = [
            ("TrueType 字体", "*.ttf"),
            ("OpenType 字体", "*.otf"),
            ("Web Open Font Format", "*.woff"),
            ("WOFF2 字体", "*.woff2"),
            ("所有文件", "*.*")
        ]
        
        self.supported_table_formats = [
            ("CSV 文件", "csv"),
            ("Excel 文件", "xlsx"),
            ("JSON 文件", "json"),
            ("HTML 表格", "html"),
            ("Markdown 表格", "md")
        ]
        
        self.setup_ui()
        
    def setup_ui(self):
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # 标题
        title_label = ttk.Label(
            main_frame, 
            text="字体转表格工具", 
            font=("Arial", 16, "bold")
        )
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # 字体文件选择
        ttk.Label(main_frame, text="字体文件:").grid(row=1, column=0, sticky=tk.W, pady=5)
        
        font_frame = ttk.Frame(main_frame)
        font_frame.grid(row=1, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        font_frame.columnconfigure(0, weight=1)
        
        self.font_entry = ttk.Entry(font_frame, textvariable=self.font_path)
        self.font_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        
        ttk.Button(
            font_frame, 
            text="浏览...", 
            command=self.browse_font_file,
            width=10
        ).grid(row=0, column=1)
        
        # 输出格式选择
        ttk.Label(main_frame, text="表格格式:").grid(row=2, column=0, sticky=tk.W, pady=5)
        
        format_frame = ttk.Frame(main_frame)
        format_frame.grid(row=2, column=1, sticky=tk.W, pady=5)
        
        for i, (format_name, format_ext) in enumerate(self.supported_table_formats):
            ttk.Radiobutton(
                format_frame,
                text=f"{format_name} (.{format_ext})",
                variable=self.table_format,
                value=format_ext
            ).grid(row=0, column=i, padx=(0, 15))
        
        # 输出文件选择
        ttk.Label(main_frame, text="输出文件:").grid(row=3, column=0, sticky=tk.W, pady=5)
        
        output_frame = ttk.Frame(main_frame)
        output_frame.grid(row=3, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        output_frame.columnconfigure(0, weight=1)
        
        self.output_entry = ttk.Entry(output_frame, textvariable=self.output_path)
        self.output_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        
        ttk.Button(
            output_frame, 
            text="浏览...", 
            command=self.browse_output_file,
            width=10
        ).grid(row=0, column=1)
        
        # 选项框架
        options_frame = ttk.LabelFrame(main_frame, text="选项", padding="10")
        options_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=15)
        
        # 包含控制字符选项
        self.include_control_chars = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            options_frame,
            text="包含控制字符（不可见字符）",
            variable=self.include_control_chars
        ).grid(row=0, column=0, sticky=tk.W)
        
        # 显示字符预览选项
        self.show_preview = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_frame,
            text="在表格中显示字符预览",
            variable=self.show_preview
        ).grid(row=0, column=1, sticky=tk.W, padx=(20, 0))
        
        # 进度条
        ttk.Label(main_frame, text="进度:").grid(row=5, column=0, sticky=tk.W, pady=(15, 5))
        
        self.progress_bar = ttk.Progressbar(
            main_frame, 
            variable=self.progress_var,
            maximum=100,
            mode='determinate'
        )
        self.progress_bar.grid(row=5, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=(15, 5))
        
        # 状态显示
        self.status_label = ttk.Label(
            main_frame, 
            textvariable=self.status_var,
            relief=tk.SUNKEN,
            padding=5
        )
        self.status_label.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=7, column=0, columnspan=3, pady=10)
        
        self.convert_button = ttk.Button(
            button_frame,
            text="开始转换",
            command=self.start_conversion,
            width=15
        )
        self.convert_button.grid(row=0, column=0, padx=5)
        
        ttk.Button(
            button_frame,
            text="预览字体",
            command=self.preview_font,
            width=15
        ).grid(row=0, column=1, padx=5)
        
        ttk.Button(
            button_frame,
            text="退出",
            command=self.root.quit,
            width=15
        ).grid(row=0, column=2, padx=5)
        
        # 信息文本
        info_text = (
            "功能说明:\n"
            "1. 选择字体文件（支持 .ttf, .otf, .woff, .woff2 格式）\n"
            "2. 选择输出表格格式（CSV, Excel, JSON, HTML, Markdown）\n"
            "3. 选择输出文件路径\n"
            "4. 点击'开始转换'按钮\n\n"
            "表格将包含以下列:\n"
            "  - 字符: 可复制的Unicode字符\n"
            "  - Unicode编码: 十六进制格式（如 U+0041）\n"
            "  - Unicode名称: 字符的官方名称（如 LATIN CAPITAL LETTER A）\n"
            "  - 区块: Unicode区块名称\n"
        )
        
        info_label = ttk.Label(
            main_frame, 
            text=info_text,
            justify=tk.LEFT,
            relief=tk.GROOVE,
            padding=10
        )
        info_label.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(20, 0))
        
    def browse_font_file(self):
        """浏览并选择字体文件"""
        filetypes = self.supported_font_formats
        
        filename = filedialog.askopenfilename(
            title="选择字体文件",
            filetypes=filetypes
        )
        
        if filename:
            self.font_path.set(filename)
            
            # 自动生成输出文件名
            if not self.output_path.get():
                font_name = Path(filename).stem
                output_dir = Path(filename).parent
                output_file = output_dir / f"{font_name}_glyphs.{self.table_format.get()}"
                self.output_path.set(str(output_file))
    
    def browse_output_file(self):
        """浏览并选择输出文件"""
        format_ext = self.table_format.get()
        filetypes = [(f"{format_ext.upper()} 文件", f"*.{format_ext}")]
        
        filename = filedialog.asksaveasfilename(
            title="保存表格文件",
            defaultextension=f".{format_ext}",
            filetypes=filetypes
        )
        
        if filename:
            self.output_path.set(filename)
    
    def get_unicode_name(self, char):
        """获取Unicode字符的名称"""
        try:
            name = unicodedata.name(char)
            return name
        except ValueError:
            return "未命名字符"
    
    def get_unicode_block(self, code_point):
        """获取Unicode字符所属的区块"""
        # Unicode区块范围定义（简化版，常用区块）
        blocks = [
            (0x0000, 0x007F, "基本拉丁文"),
            (0x0080, 0x00FF, "拉丁文补充-1"),
            (0x0100, 0x017F, "拉丁文扩展-A"),
            (0x0180, 0x024F, "拉丁文扩展-B"),
            (0x0370, 0x03FF, "希腊文及科普特文"),
            (0x0400, 0x04FF, "西里尔文"),
            (0x0500, 0x052F, "西里尔文补充"),
            (0x0530, 0x058F, "亚美尼亚文"),
            (0x0590, 0x05FF, "希伯来文"),
            (0x0600, 0x06FF, "阿拉伯文"),
            (0x0700, 0x074F, "叙利亚文"),
            (0x0750, 0x077F, "阿拉伯文补充"),
            (0x0780, 0x07BF, "它拿文"),
            (0x0800, 0x083F, "撒马利亚文"),
            (0x0840, 0x085F, "曼达文"),
            (0x0860, 0x086F, "叙利亚文补充"),
            (0x08A0, 0x08FF, "阿拉伯文扩展-A"),
            (0x0900, 0x097F, "天城文"),
            (0x0980, 0x09FF, "孟加拉文"),
            (0x0A00, 0x0A7F, "古木基文"),
            (0x0A80, 0x0AFF, "古吉拉特文"),
            (0x0B00, 0x0B7F, "奥里亚文"),
            (0x0B80, 0x0BFF, "泰米尔文"),
            (0x0C00, 0x0C7F, "泰卢固文"),
            (0x0C80, 0x0CFF, "卡纳达文"),
            (0x0D00, 0x0D7F, "马拉雅拉姆文"),
            (0x0D80, 0x0DFF, "僧伽罗文"),
            (0x0E00, 0x0E7F, "泰文"),
            (0x0E80, 0x0EFF, "老挝文"),
            (0x0F00, 0x0FFF, "藏文"),
            (0x1000, 0x109F, "缅甸文"),
            (0x10A0, 0x10FF, "格鲁吉亚文"),
            (0x1100, 0x11FF, "谚文兼容字母"),
            (0x1200, 0x137F, "埃塞俄比亚文"),
            (0x1380, 0x139F, "埃塞俄比亚文补充"),
            (0x13A0, 0x13FF, "切罗基文"),
            (0x1400, 0x167F, "统一加拿大原住民音节文字"),
            (0x1680, 0x169F, "欧甘文"),
            (0x16A0, 0x16FF, "如尼文"),
            (0x1700, 0x171F, "塔加拉文"),
            (0x1720, 0x173F, "哈努诺文"),
            (0x1740, 0x175F, "布希德文"),
            (0x1760, 0x177F, "塔格班瓦文"),
            (0x1780, 0x17FF, "高棉文"),
            (0x1800, 0x18AF, "蒙古文"),
            (0x18B0, 0x18FF, "统一加拿大原住民音节文字扩展"),
            (0x1900, 0x194F, "林布文"),
            (0x1950, 0x197F, "德宏傣文"),
            (0x1980, 0x19DF, "新傣仂文"),
            (0x19E0, 0x19FF, "高棉符号"),
            (0x1A00, 0x1A1F, "布吉文"),
            (0x1A20, 0x1AAF, "兰纳文"),
            (0x1AB0, 0x1AFF, "组合变音标记扩展"),
            (0x1B00, 0x1B7F, "巴厘文"),
            (0x1B80, 0x1BBF, "巽他文"),
            (0x1BC0, 0x1BFF, "巴塔克文"),
            (0x1C00, 0x1C4F, "雷布查文"),
            (0x1C50, 0x1C7F, "桑塔利文"),
            (0x1C80, 0x1C8F, "西里尔文扩展-C"),
            (0x1C90, 0x1CBF, "格鲁吉亚文扩展"),
            (0x1CC0, 0x1CCF, "巽他文补充"),
            (0x1CD0, 0x1CFF, "吠陀扩展"),
            (0x1D00, 0x1D7F, "音标扩展"),
            (0x1D80, 0x1DBF, "音标扩展补充"),
            (0x1DC0, 0x1DFF, "组合变音标记补充"),
            (0x1E00, 0x1EFF, "拉丁文扩展附加"),
            (0x1F00, 0x1FFF, "希腊文扩展"),
            (0x2000, 0x206F, "常用标点"),
            (0x2070, 0x209F, "上标和下标"),
            (0x20A0, 0x20CF, "货币符号"),
            (0x20D0, 0x20FF, "组合用符号"),
            (0x2100, 0x214F, "字母式符号"),
            (0x2150, 0x218F, "数字形式"),
            (0x2190, 0x21FF, "箭头"),
            (0x2200, 0x22FF, "数学运算符"),
            (0x2300, 0x23FF, "杂项技术符号"),
            (0x2400, 0x243F, "控制图片"),
            (0x2440, 0x245F, "光学识别符"),
            (0x2460, 0x24FF, "带圈字母数字"),
            (0x2500, 0x257F, "制表符"),
            (0x2580, 0x259F, "方块元素"),
            (0x25A0, 0x25FF, "几何图形"),
            (0x2600, 0x26FF, "杂项符号"),
            (0x2700, 0x27BF, "印刷符号"),
            (0x27C0, 0x27EF, "杂项数学符号-A"),
            (0x27F0, 0x27FF, "补充箭头-A"),
            (0x2800, 0x28FF, "盲文点字模型"),
            (0x2900, 0x297F, "补充箭头-B"),
            (0x2980, 0x29FF, "杂项数学符号-B"),
            (0x2A00, 0x2AFF, "补充数学运算符"),
            (0x2B00, 0x2BFF, "杂项符号和箭头"),
            (0x2C00, 0x2C5F, "格拉哥里文"),
            (0x2C60, 0x2C7F, "拉丁文扩展-C"),
            (0x2C80, 0x2CFF, "科普特文"),
            (0x2D00, 0x2D2F, "格鲁吉亚文补充"),
            (0x2D30, 0x2D7F, "提非纳文"),
            (0x2D80, 0x2DDF, "埃塞俄比亚文扩展"),
            (0x2DE0, 0x2DFF, "西里尔文扩展-A"),
            (0x2E00, 0x2E7F, "补充标点"),
            (0x2E80, 0x2EFF, "中日韩汉字部首补充"),
            (0x2F00, 0x2FDF, "康熙部首"),
            (0x2FF0, 0x2FFF, "表意文字描述字符"),
            (0x3000, 0x303F, "中日韩符号和标点"),
            (0x3040, 0x309F, "日文平假名"),
            (0x30A0, 0x30FF, "日文片假名"),
            (0x3100, 0x312F, "注音字母"),
            (0x3130, 0x318F, "谚文兼容字母"),
            (0x3190, 0x319F, "汉文注释标志"),
            (0x31A0, 0x31BF, "注音字母扩展"),
            (0x31C0, 0x31EF, "中日韩笔画"),
            (0x31F0, 0x31FF, "日文片假名拼音扩展"),
            (0x3200, 0x32FF, "带圈中日韩字母和月份"),
            (0x3300, 0x33FF, "中日韩兼容字符"),
            (0x3400, 0x4DBF, "中日韩统一表意文字扩展A"),
            (0x4DC0, 0x4DFF, "易经六十四卦符号"),
            (0x4E00, 0x9FFF, "中日韩统一表意文字"),
            (0xA000, 0xA48F, "彝文音节"),
            (0xA490, 0xA4CF, "彝文字根"),
            (0xA4D0, 0xA4FF, "老傈僳文"),
            (0xA500, 0xA63F, "瓦伊文"),
            (0xA640, 0xA69F, "西里尔文扩展-B"),
            (0xA6A0, 0xA6FF, "巴穆姆文"),
            (0xA700, 0xA71F, "声调修饰字母"),
            (0xA720, 0xA7FF, "拉丁文扩展-D"),
            (0xA800, 0xA82F, "锡尔赫特文"),
            (0xA830, 0xA83F, "通用印度数字格式"),
            (0xA840, 0xA87F, "八思巴文"),
            (0xA880, 0xA8DF, "索拉什特拉文"),
            (0xA8E0, 0xA8FF, "天城文扩展"),
            (0xA900, 0xA92F, "克耶文"),
            (0xA930, 0xA95F, "勒姜文"),
            (0xA960, 0xA97F, "谚文扩展-A"),
            (0xA980, 0xA9DF, "爪哇文"),
            (0xA9E0, 0xA9FF, "缅甸文扩展-B"),
            (0xAA00, 0xAA5F, "占文"),
            (0xAA60, 0xAA7F, "缅甸文扩展-A"),
            (0xAA80, 0xAADF, "越南傣文"),
            (0xAAE0, 0xAAFF, "梅泰文扩展"),
            (0xAB00, 0xAB2F, "埃塞俄比亚文扩展-A"),
            (0xAB30, 0xAB6F, "拉丁文扩展-E"),
            (0xAB70, 0xABBF, "切罗基文补充"),
            (0xABC0, 0xABFF, "曼尼普尔文"),
            (0xAC00, 0xD7AF, "谚文音节"),
            (0xD7B0, 0xD7FF, "谚文字母扩展-B"),
            (0xF900, 0xFAFF, "中日韩兼容表意文字"),
            (0xFB00, 0xFB4F, "字母表达形式"),
            (0xFB50, 0xFDFF, "阿拉伯表达形式-A"),
            (0xFE00, 0xFE0F, "变体选择符"),
            (0xFE10, 0xFE1F, "竖排形式"),
            (0xFE20, 0xFE2F, "组合用半符号"),
            (0xFE30, 0xFE4F, "中日韩兼容形式"),
            (0xFE50, 0xFE6F, "小写变体形式"),
            (0xFE70, 0xFEFF, "阿拉伯表达形式-B"),
            (0xFF00, 0xFFEF, "半形及全形形式"),
            (0xFFF0, 0xFFFF, "特殊"),
            (0x10000, 0x1007F, "线性文字B音节文字"),
            (0x10080, 0x100FF, "线性文字B表意文字"),
            (0x10100, 0x1013F, "爱琴海数字"),
            (0x10140, 0x1018F, "古希腊数字"),
            (0x10190, 0x101CF, "古代符号"),
            (0x101D0, 0x101FF, "费斯托斯圆盘"),
            (0x10280, 0x1029F, "吕基亚文"),
            (0x102A0, 0x102DF, "卡里亚文"),
            (0x102E0, 0x102FF, "科普特历法"),
            (0x10300, 0x1032F, "古意大利文"),
            (0x10330, 0x1034F, "哥特文"),
            (0x10350, 0x1037F, "古彼尔姆文"),
            (0x10380, 0x1039F, "乌加里特文"),
            (0x103A0, 0x103DF, "古波斯文"),
            (0x10400, 0x1044F, "德瑟雷特文"),
            (0x10450, 0x1047F, "肃伯纳文"),
            (0x10480, 0x104AF, "奥斯曼亚文"),
            (0x104B0, 0x104FF, "欧塞奇文"),
            (0x10500, 0x1052F, "爱尔巴桑文"),
            (0x10530, 0x1056F, "高加索阿尔巴尼亚文"),
            (0x10600, 0x1077F, "线性文字A"),
            (0x10800, 0x1083F, "塞浦路斯音节文字"),
            (0x10840, 0x1085F, "帝国亚拉姆文"),
            (0x10860, 0x1087F, "帕尔迈拉文"),
            (0x10880, 0x108AF, "纳巴泰文"),
            (0x108E0, 0x108FF, "哈特拉文"),
            (0x10900, 0x1091F, "腓尼基文"),
            (0x10920, 0x1093F, "吕底亚文"),
            (0x10980, 0x1099F, "麦罗埃文圣书体"),
            (0x109A0, 0x109FF, "麦罗埃文草书体"),
            (0x10A00, 0x10A5F, "佉卢文"),
            (0x10A60, 0x10A7F, "古南阿拉伯文"),
            (0x10A80, 0x10A9F, "古北阿拉伯文"),
            (0x10AC0, 0x10AFF, "曼尼安文"),
            (0x10B00, 0x10B3F, "阿维斯陀文"),
            (0x10B40, 0x10B5F, "碑铭帕提亚文"),
            (0x10B60, 0x10B7F, "碑铭巴列维文"),
            (0x10B80, 0x10BAF, "诗篇巴列维文"),
            (0x10C00, 0x10C4F, "古突厥文"),
            (0x10C80, 0x10CFF, "古匈牙利文"),
            (0x10E60, 0x10E7F, "卢米符号数字"),
            (0x11000, 0x1107F, "婆罗米文"),
            (0x11080, 0x110CF, "凯提文"),
            (0x110D0, 0x110FF, "索拉桑朋文"),
            (0x11100, 0x1114F, "查克马文"),
            (0x11150, 0x1117F, "马哈贾尼文"),
            (0x11180, 0x111DF, "夏拉达文"),
            (0x111E0, 0x111FF, "信德文"),
            (0x11200, 0x1124F, "格兰塔文"),
            (0x11280, 0x112AF, "古吉拉特文"),
            (0x112B0, 0x112FF, "索拉什特拉文"),
            (0x11300, 0x1137F, "泰米尔文"),
            (0x11400, 0x1147F, "泰卢固文"),
            (0x11480, 0x114DF, "埃塞俄比亚文扩展"),
            (0x11580, 0x115FF, "悉昙文"),
            (0x11600, 0x1165F, "蒙古文补充"),
            (0x11660, 0x1167F, "加拿大原住民音节文字扩展"),
            (0x11680, 0x116CF, "泰克里文"),
            (0x11800, 0x1184F, "瓦郎奇蒂文"),
            (0x118A0, 0x118FF, "万秋文"),
            (0x11AC0, 0x11AFF, "蒲甘文"),
            (0x11C00, 0x11C6F, "拜克舒基文"),
            (0x11C70, 0x11CBF, "玛钦文"),
            (0x12000, 0x123FF, "楔形文字"),
            (0x12400, 0x1247F, "楔形文字数字和标点"),
            (0x12480, 0x1254F, "早期王朝楔形文字"),
            (0x13000, 0x1342F, "埃及圣书体"),
            (0x13430, 0x1343F, "埃及圣书体格式控制"),
            (0x14400, 0x1467F, "安纳托利亚象形文字"),
            (0x16800, 0x16A3F, "巴姆穆文字补充"),
            (0x16A40, 0x16A6F, "穆塔文"),
            (0x16AD0, 0x16AFF, "巴萨瓦赫文"),
            (0x16B00, 0x16B8F, "帕西安文"),
            (0x16F00, 0x16F9F, "柏格理苗文"),
            (0x16FE0, 0x16FFF, "表意符号和标点"),
            (0x17000, 0x187FF, "中日韩统一表意文字扩展B"),
            (0x18800, 0x18AFF, "中日韩统一表意文字扩展C"),
            (0x1B000, 0x1B0FF, "日文假名补充"),
            (0x1B100, 0x1B12F, "日文假名扩展-A"),
            (0x1B170, 0x1B2FF, "女书"),
            (0x1BC00, 0x1BC9F, "杜普洛伊速记"),
            (0x1BCA0, 0x1BCAF, "格式控制符号的速记"),
            (0x1D000, 0x1D0FF, "拜占庭音乐符号"),
            (0x1D100, 0x1D1FF, "音乐符号"),
            (0x1D200, 0x1D24F, "古希腊音乐记号"),
            (0x1D300, 0x1D35F, "太玄经符号"),
            (0x1D360, 0x1D37F, "算筹"),
            (0x1D400, 0x1D7FF, "数学字母数字符号"),
            (0x1E800, 0x1E8DF, "门德文"),
            (0x1E900, 0x1E95F, "阿德拉姆文"),
            (0x1EE00, 0x1EEFF, "阿拉伯数学字母符号"),
            (0x1F000, 0x1F02F, "麻将牌"),
            (0x1F030, 0x1F09F, "多米诺骨牌"),
            (0x1F0A0, 0x1F0FF, "扑克牌"),
            (0x1F100, 0x1F1FF, "带圈字母数字补充"),
            (0x1F200, 0x1F2FF, "封闭式表意文字补充"),
            (0x1F300, 0x1F5FF, "杂项符号和象形文字"),
            (0x1F600, 0x1F64F, "表情符号"),
            (0x1F650, 0x1F67F, "装饰符号"),
            (0x1F680, 0x1F6FF, "交通和地图符号"),
            (0x1F700, 0x1F77F, "炼金术符号"),
            (0x1F780, 0x1F7FF, "几何图形扩展"),
            (0x1F800, 0x1F8FF, "补充箭头-C"),
            (0x1F900, 0x1F9FF, "补充符号和象形文字"),
            (0x20000, 0x2A6DF, "中日韩统一表意文字扩展C"),
            (0x2A700, 0x2B73F, "中日韩统一表意文字扩展D"),
            (0x2B740, 0x2B81F, "中日韩统一表意文字扩展E"),
            (0x2B820, 0x2CEAF, "中日韩统一表意文字扩展F"),
            (0x2CEB0, 0x2EBEF, "中日韩统一表意文字扩展G"),
            (0x2F800, 0x2FA1F, "中日韩兼容表意文字补充"),
            (0xE0000, 0xE007F, "标签"),
            (0xE0100, 0xE01EF, "变体选择符补充"),
            (0xF0000, 0xFFFFF, "专用区（补充专用区）"),
            (0x100000, 0x10FFFF, "辅助专用区"),
        ]
        
        for start, end, name in blocks:
            if start <= code_point <= end:
                return name
        
        return "其他区块"
    
    def extract_font_glyphs(self, font_path, include_control_chars=False, show_preview=True):
        """从字体文件中提取字形信息"""
        if not FONTTOOLS_AVAILABLE:
            raise ImportError("fontTools库未安装，请运行: pip install fonttools")
        
        if not PANDAS_AVAILABLE:
            raise ImportError("pandas库未安装，请运行: pip install pandas")
        
        self.status_var.set("正在加载字体文件...")
        self.progress_var.set(10)
        self.root.update()
        
        try:
            font = TTFont(font_path)
        except Exception as e:
            raise Exception(f"无法加载字体文件: {e}")
        
        self.status_var.set("正在提取字形信息...")
        self.progress_var.set(30)
        self.root.update()
        
        glyphs_data = []
        
        # 获取cmap表（字符映射表）
        if 'cmap' not in font:
            raise Exception("字体文件不包含cmap表（字符映射表）")
        
        cmap = font.getBestCmap()
        
        if not cmap:
            raise Exception("无法从字体文件中提取字符映射")
        
        total_glyphs = len(cmap)
        processed = 0
        
        for code_point, glyph_name in cmap.items():
            try:
                # 更新进度
                processed += 1
                progress = 30 + (processed / total_glyphs) * 60
                self.progress_var.set(progress)
                
                if processed % 100 == 0:
                    self.status_var.set(f"正在处理字符: {processed}/{total_glyphs}")
                    self.root.update()
                
                # 将Unicode码点转换为字符
                char = chr(code_point)
                
                # 如果不包含控制字符且当前字符是控制字符，则跳过
                if not include_control_chars and unicodedata.category(char).startswith('C'):
                    continue
                
                # 如果不显示预览，将控制字符替换为占位符
                display_char = char
                if not show_preview and unicodedata.category(char).startswith('C'):
                    display_char = "□"  # 用方框表示控制字符
                
                # 获取Unicode名称
                unicode_name = self.get_unicode_name(char)
                
                # 获取Unicode区块
                unicode_block = self.get_unicode_block(code_point)
                
                # 格式化Unicode码点
                unicode_hex = f"U+{code_point:04X}"
                
                # 添加到数据列表
                glyphs_data.append({
                    "字符": display_char,
                    "Unicode编码": unicode_hex,
                    "Unicode名称": unicode_name,
                    "区块": unicode_block,
                    "十进制码点": code_point,
                    "字符类别": unicodedata.category(char)
                })
                
            except Exception as e:
                # 跳过无法处理的字符
                continue
        
        font.close()
        
        self.status_var.set(f"成功提取 {len(glyphs_data)} 个字符")
        self.progress_var.set(95)
        self.root.update()
        
        return glyphs_data
    
    def save_table(self, data, output_path, table_format):
        """保存数据为表格文件"""
        self.status_var.set(f"正在保存为{table_format.upper()}文件...")
        self.root.update()
        
        # 创建DataFrame
        df = pd.DataFrame(data)
        
        # 按Unicode编码排序
        df = df.sort_values("十进制码点").reset_index(drop=True)
        
        # 删除临时列
        df = df.drop(columns=["十进制码点", "字符类别"])
        
        # 根据格式保存
        if table_format == "csv":
            df.to_csv(output_path, index=False, encoding='utf-8-sig')
        
        elif table_format == "xlsx":
            if not OPENPYXL_AVAILABLE:
                # 尝试安装openpyxl或使用其他引擎
                try:
                    df.to_excel(output_path, index=False, engine='openpyxl')
                except:
                    # 如果没有openpyxl，尝试使用xlwt（仅支持.xls）
                    output_path = output_path.replace('.xlsx', '.xls')
                    df.to_excel(output_path, index=False, engine='xlwt')
            else:
                df.to_excel(output_path, index=False, engine='openpyxl')
        
        elif table_format == "json":
            df.to_json(output_path, orient='records', force_ascii=False, indent=2)
        
        elif table_format == "html":
            html_table = df.to_html(index=False, classes='font-glyphs-table')
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>字体字形表</title>
    <style>
        table {{ border-collapse: collapse; width: 100%; }}
        th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
        th {{ background-color: #f2f2f2; }}
        tr:nth-child(even) {{ background-color: #f9f9f9; }}
        .char-cell {{ font-family: monospace; font-size: 24px; text-align: center; }}
    </style>
</head>
<body>
    <h1>字体字形表</h1>
    <p>生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
    {html_table}
</body>
</html>""")
        
        elif table_format == "md":
            df.to_markdown(output_path, index=False)
        
        else:
            raise ValueError(f"不支持的表格格式: {table_format}")
    
    def start_conversion(self):
        """开始转换过程"""
        # 检查库是否可用
        if not FONTTOOLS_AVAILABLE:
            messagebox.showerror("错误", "fontTools库未安装，请运行: pip install fonttools")
            return
        
        if not PANDAS_AVAILABLE:
            messagebox.showerror("错误", "pandas库未安装，请运行: pip install pandas")
            return
        
        # 检查输入文件
        font_path = self.font_path.get()
        if not font_path or not os.path.exists(font_path):
            messagebox.showerror("错误", "请选择有效的字体文件")
            return
        
        # 检查输出路径
        output_path = self.output_path.get()
        if not output_path:
            messagebox.showerror("错误", "请选择输出文件路径")
            return
        
        # 检查输出目录是否存在
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
            except Exception as e:
                messagebox.showerror("错误", f"无法创建输出目录: {e}")
                return
        
        # 禁用转换按钮，防止重复点击
        self.convert_button.config(state=tk.DISABLED)
        self.status_var.set("开始处理...")
        self.progress_var.set(0)
        
        # 在新线程中执行转换
        thread = threading.Thread(
            target=self.convert_thread,
            args=(font_path, output_path)
        )
        thread.daemon = True
        thread.start()
    
    def convert_thread(self, font_path, output_path):
        """转换线程"""
        try:
            # 提取字形数据
            glyphs_data = self.extract_font_glyphs(
                font_path,
                self.include_control_chars.get(),
                self.show_preview.get()
            )
            
            if not glyphs_data:
                self.root.after(0, lambda: messagebox.showwarning("警告", "未找到可提取的字符"))
                return
            
            # 保存表格
            table_format = self.table_format.get()
            self.save_table(glyphs_data, output_path, table_format)
            
            # 更新状态
            self.progress_var.set(100)
            self.status_var.set(f"转换完成！共提取 {len(glyphs_data)} 个字符")
            
            # 显示成功消息
            self.root.after(0, lambda: messagebox.showinfo(
                "成功", 
                f"转换完成！\n共提取 {len(glyphs_data)} 个字符\n文件已保存到: {output_path}"
            ))
            
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("错误", f"转换失败: {str(e)}\n\n{traceback.format_exc()}"))
        
        finally:
            # 重新启用转换按钮
            self.root.after(0, lambda: self.convert_button.config(state=tk.NORMAL))
            self.progress_var.set(0)
    
    def preview_font(self):
        """预览字体"""
        font_path = self.font_path.get()
        if not font_path or not os.path.exists(font_path):
            messagebox.showerror("错误", "请先选择字体文件")
            return
        
        try:
            # 创建预览窗口
            preview_window = tk.Toplevel(self.root)
            preview_window.title("字体预览")
            preview_window.geometry("600x400")
            
            # 尝试加载字体
            try:
                font_family = os.path.basename(font_path).split('.')[0]
                custom_font = tkfont.Font(family=font_family, size=20)
                
                # 如果字体加载失败，使用默认字体
                if custom_font.actual()["family"] == "TkDefaultFont":
                    raise Exception("无法加载字体")
            except:
                messagebox.showwarning("警告", "无法在系统中加载此字体进行预览\n将在表格中查看实际字符")
                preview_window.destroy()
                return
            
            # 创建文本框显示字体示例
            text_frame = ttk.Frame(preview_window, padding="10")
            text_frame.pack(fill=tk.BOTH, expand=True)
            
            ttk.Label(text_frame, text="字体预览:", font=("Arial", 12, "bold")).pack(anchor=tk.W)
            
            text_widget = tk.Text(text_frame, height=10, width=50, font=custom_font)
            text_widget.pack(fill=tk.BOTH, expand=True, pady=(5, 0))
            
            # 添加示例文本
            sample_text = """ABCDEFGHIJKLMNOPQRSTUVWXYZ
abcdefghijklmnopqrstuvwxyz
0123456789
!@#$%^&*()_+-=[]{}|;:'",.<>/?
中文字符示例：你好，世界！
Font Preview: 字体预览"""
            
            text_widget.insert(1.0, sample_text)
            text_widget.config(state=tk.DISABLED)
            
            # 添加滚动条
            scrollbar = ttk.Scrollbar(text_widget, command=text_widget.yview)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            text_widget.config(yscrollcommand=scrollbar.set)
            
            # 字体信息
            info_frame = ttk.LabelFrame(preview_window, text="字体信息", padding="10")
            info_frame.pack(fill=tk.X, padx=10, pady=10)
            
            font_name = os.path.basename(font_path)
            file_size = os.path.getsize(font_path) / 1024  # KB
            
            info_text = f"""字体文件: {font_name}
文件大小: {file_size:.2f} KB
文件路径: {font_path}
字体名称: {font_family}"""
            
            ttk.Label(info_frame, text=info_text, justify=tk.LEFT).pack(anchor=tk.W)
            
            # 关闭按钮
            ttk.Button(
                preview_window, 
                text="关闭", 
                command=preview_window.destroy
            ).pack(pady=10)
            
        except Exception as e:
            messagebox.showerror("错误", f"预览失败: {e}")

def main():
    """主函数"""
    # 检查必要库
    if not FONTTOOLS_AVAILABLE or not PANDAS_AVAILABLE:
        print("错误: 缺少必要的库")
        print("请运行以下命令安装所需库:")
        print("pip install fonttools pandas")
        if not OPENPYXL_AVAILABLE:
            print("pip install openpyxl  # 用于Excel文件支持")
        return
    
    # 创建主窗口
    root = tk.Tk()
    app = FontToTableApp(root)
    
    # 运行主循环
    root.mainloop()

if __name__ == "__main__":
    main()
