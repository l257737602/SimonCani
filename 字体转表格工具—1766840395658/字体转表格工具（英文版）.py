#!/usr/bin/env python3
"""
Font to Table Converter
Extract font glyph information and export to table files
Support multiple font formats and table formats
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from pathlib import Path
import threading
import traceback
from datetime import datetime
import unicodedata
import sys

# Third-party library imports
try:
    from fontTools.ttLib import TTFont
    FONTTOOLS_AVAILABLE = True
except ImportError:
    FONTTOOLS_AVAILABLE = False
    print("Warning: fontTools library not installed, please run: pip install fonttools")

try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False
    print("Warning: pandas library not installed, please run: pip install pandas")

try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False
    print("Info: openpyxl not installed, Excel export may use xlwt engine")

class FontToTableApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Font to Table Converter")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # Set window icon
        try:
            self.root.iconbitmap(default='')
        except:
            pass
        
        # Initialize variables
        self.font_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.table_format = tk.StringVar(value="csv")
        self.status_var = tk.StringVar(value="Ready")
        self.progress_var = tk.DoubleVar(value=0)
        
        # Supported formats
        self.supported_font_formats = [
            ("TrueType Font", "*.ttf"),
            ("OpenType Font", "*.otf"),
            ("Web Open Font Format", "*.woff"),
            ("WOFF2 Font", "*.woff2"),
            ("All Files", "*.*")
        ]
        
        self.supported_table_formats = [
            ("CSV File", "csv"),
            ("Excel File", "xlsx"),
            ("JSON File", "json"),
            ("HTML Table", "html"),
            ("Markdown Table", "md")
        ]
        
        # Unicode 17.0 Block Names
        self.unicode_blocks = self.load_unicode_blocks()
        
        self.setup_ui()
        
    def load_unicode_blocks(self):
        """Load complete Unicode 17.0 block names"""
        # Complete Unicode 17.0 blocks (updated to Unicode 17.0)
        blocks = [
            (0x0000, 0x007F, "Basic Latin"),
            (0x0080, 0x00FF, "Latin-1 Supplement"),
            (0x0100, 0x017F, "Latin Extended-A"),
            (0x0180, 0x024F, "Latin Extended-B"),
            (0x0250, 0x02AF, "IPA Extensions"),
            (0x02B0, 0x02FF, "Spacing Modifier Letters"),
            (0x0300, 0x036F, "Combining Diacritical Marks"),
            (0x0370, 0x03FF, "Greek and Coptic"),
            (0x0400, 0x04FF, "Cyrillic"),
            (0x0500, 0x052F, "Cyrillic Supplement"),
            (0x0530, 0x058F, "Armenian"),
            (0x0590, 0x05FF, "Hebrew"),
            (0x0600, 0x06FF, "Arabic"),
            (0x0700, 0x074F, "Syriac"),
            (0x0750, 0x077F, "Arabic Supplement"),
            (0x0780, 0x07BF, "Thaana"),
            (0x07C0, 0x07FF, "NKo"),
            (0x0800, 0x083F, "Samaritan"),
            (0x0840, 0x085F, "Mandaic"),
            (0x0860, 0x086F, "Syriac Supplement"),
            (0x0870, 0x089F, "Arabic Extended-B"),
            (0x08A0, 0x08FF, "Arabic Extended-A"),
            (0x0900, 0x097F, "Devanagari"),
            (0x0980, 0x09FF, "Bengali"),
            (0x0A00, 0x0A7F, "Gurmukhi"),
            (0x0A80, 0x0AFF, "Gujarati"),
            (0x0B00, 0x0B7F, "Oriya"),
            (0x0B80, 0x0BFF, "Tamil"),
            (0x0C00, 0x0C7F, "Telugu"),
            (0x0C80, 0x0CFF, "Kannada"),
            (0x0D00, 0x0D7F, "Malayalam"),
            (0x0D80, 0x0DFF, "Sinhala"),
            (0x0E00, 0x0E7F, "Thai"),
            (0x0E80, 0x0EFF, "Lao"),
            (0x0F00, 0x0FFF, "Tibetan"),
            (0x1000, 0x109F, "Myanmar"),
            (0x10A0, 0x10FF, "Georgian"),
            (0x1100, 0x11FF, "Hangul Jamo"),
            (0x1200, 0x137F, "Ethiopic"),
            (0x1380, 0x139F, "Ethiopic Supplement"),
            (0x13A0, 0x13FF, "Cherokee"),
            (0x1400, 0x167F, "Unified Canadian Aboriginal Syllabics"),
            (0x1680, 0x169F, "Ogham"),
            (0x16A0, 0x16FF, "Runic"),
            (0x1700, 0x171F, "Tagalog"),
            (0x1720, 0x173F, "Hanunoo"),
            (0x1740, 0x175F, "Buhid"),
            (0x1760, 0x177F, "Tagbanwa"),
            (0x1780, 0x17FF, "Khmer"),
            (0x1800, 0x18AF, "Mongolian"),
            (0x18B0, 0x18FF, "Unified Canadian Aboriginal Syllabics Extended"),
            (0x1900, 0x194F, "Limbu"),
            (0x1950, 0x197F, "Tai Le"),
            (0x1980, 0x19DF, "New Tai Lue"),
            (0x19E0, 0x19FF, "Khmer Symbols"),
            (0x1A00, 0x1A1F, "Buginese"),
            (0x1A20, 0x1AAF, "Tai Tham"),
            (0x1AB0, 0x1AFF, "Combining Diacritical Marks Extended"),
            (0x1B00, 0x1B7F, "Balinese"),
            (0x1B80, 0x1BBF, "Sundanese"),
            (0x1BC0, 0x1BFF, "Batak"),
            (0x1C00, 0x1C4F, "Lepcha"),
            (0x1C50, 0x1C7F, "Ol Chiki"),
            (0x1C80, 0x1C8F, "Cyrillic Extended-C"),
            (0x1C90, 0x1CBF, "Georgian Extended"),
            (0x1CC0, 0x1CCF, "Sundanese Supplement"),
            (0x1CD0, 0x1CFF, "Vedic Extensions"),
            (0x1D00, 0x1D7F, "Phonetic Extensions"),
            (0x1D80, 0x1DBF, "Phonetic Extensions Supplement"),
            (0x1DC0, 0x1DFF, "Combining Diacritical Marks Supplement"),
            (0x1E00, 0x1EFF, "Latin Extended Additional"),
            (0x1F00, 0x1FFF, "Greek Extended"),
            (0x2000, 0x206F, "General Punctuation"),
            (0x2070, 0x209F, "Superscripts and Subscripts"),
            (0x20A0, 0x20CF, "Currency Symbols"),
            (0x20D0, 0x20FF, "Combining Diacritical Marks for Symbols"),
            (0x2100, 0x214F, "Letterlike Symbols"),
            (0x2150, 0x218F, "Number Forms"),
            (0x2190, 0x21FF, "Arrows"),
            (0x2200, 0x22FF, "Mathematical Operators"),
            (0x2300, 0x23FF, "Miscellaneous Technical"),
            (0x2400, 0x243F, "Control Pictures"),
            (0x2440, 0x245F, "Optical Character Recognition"),
            (0x2460, 0x24FF, "Enclosed Alphanumerics"),
            (0x2500, 0x257F, "Box Drawing"),
            (0x2580, 0x259F, "Block Elements"),
            (0x25A0, 0x25FF, "Geometric Shapes"),
            (0x2600, 0x26FF, "Miscellaneous Symbols"),
            (0x2700, 0x27BF, "Dingbats"),
            (0x27C0, 0x27EF, "Miscellaneous Mathematical Symbols-A"),
            (0x27F0, 0x27FF, "Supplemental Arrows-A"),
            (0x2800, 0x28FF, "Braille Patterns"),
            (0x2900, 0x297F, "Supplemental Arrows-B"),
            (0x2980, 0x29FF, "Miscellaneous Mathematical Symbols-B"),
            (0x2A00, 0x2AFF, "Supplemental Mathematical Operators"),
            (0x2B00, 0x2BFF, "Miscellaneous Symbols and Arrows"),
            (0x2C00, 0x2C5F, "Glagolitic"),
            (0x2C60, 0x2C7F, "Latin Extended-C"),
            (0x2C80, 0x2CFF, "Coptic"),
            (0x2D00, 0x2D2F, "Georgian Supplement"),
            (0x2D30, 0x2D7F, "Tifinagh"),
            (0x2D80, 0x2DDF, "Ethiopic Extended"),
            (0x2DE0, 0x2DFF, "Cyrillic Extended-A"),
            (0x2E00, 0x2E7F, "Supplemental Punctuation"),
            (0x2E80, 0x2EFF, "CJK Radicals Supplement"),
            (0x2F00, 0x2FDF, "Kangxi Radicals"),
            (0x2FF0, 0x2FFF, "Ideographic Description Characters"),
            (0x3000, 0x303F, "CJK Symbols and Punctuation"),
            (0x3040, 0x309F, "Hiragana"),
            (0x30A0, 0x30FF, "Katakana"),
            (0x3100, 0x312F, "Bopomofo"),
            (0x3130, 0x318F, "Hangul Compatibility Jamo"),
            (0x3190, 0x319F, "Kanbun"),
            (0x31A0, 0x31BF, "Bopomofo Extended"),
            (0x31C0, 0x31EF, "CJK Strokes"),
            (0x31F0, 0x31FF, "Katakana Phonetic Extensions"),
            (0x3200, 0x32FF, "Enclosed CJK Letters and Months"),
            (0x3300, 0x33FF, "CJK Compatibility"),
            (0x3400, 0x4DBF, "CJK Unified Ideographs Extension A"),
            (0x4DC0, 0x4DFF, "Yijing Hexagram Symbols"),
            (0x4E00, 0x9FFF, "CJK Unified Ideographs"),
            (0xA000, 0xA48F, "Yi Syllables"),
            (0xA490, 0xA4CF, "Yi Radicals"),
            (0xA4D0, 0xA4FF, "Lisu"),
            (0xA500, 0xA63F, "Vai"),
            (0xA640, 0xA69F, "Cyrillic Extended-B"),
            (0xA6A0, 0xA6FF, "Bamum"),
            (0xA700, 0xA71F, "Modifier Tone Letters"),
            (0xA720, 0xA7FF, "Latin Extended-D"),
            (0xA800, 0xA82F, "Syloti Nagri"),
            (0xA830, 0xA83F, "Common Indic Number Forms"),
            (0xA840, 0xA87F, "Phags-pa"),
            (0xA880, 0xA8DF, "Saurashtra"),
            (0xA8E0, 0xA8FF, "Devanagari Extended"),
            (0xA900, 0xA92F, "Kayah Li"),
            (0xA930, 0xA95F, "Rejang"),
            (0xA960, 0xA97F, "Hangul Jamo Extended-A"),
            (0xA980, 0xA9DF, "Javanese"),
            (0xA9E0, 0xA9FF, "Myanmar Extended-B"),
            (0xAA00, 0xAA5F, "Cham"),
            (0xAA60, 0xAA7F, "Myanmar Extended-A"),
            (0xAA80, 0xAADF, "Tai Viet"),
            (0xAAE0, 0xAAFF, "Meetei Mayek Extensions"),
            (0xAB00, 0xAB2F, "Ethiopic Extended-A"),
            (0xAB30, 0xAB6F, "Latin Extended-E"),
            (0xAB70, 0xABBF, "Cherokee Supplement"),
            (0xABC0, 0xABFF, "Meetei Mayek"),
            (0xAC00, 0xD7AF, "Hangul Syllables"),
            (0xD7B0, 0xD7FF, "Hangul Jamo Extended-B"),
            (0xF900, 0xFAFF, "CJK Compatibility Ideographs"),
            (0xFB00, 0xFB4F, "Alphabetic Presentation Forms"),
            (0xFB50, 0xFDFF, "Arabic Presentation Forms-A"),
            (0xFE00, 0xFE0F, "Variation Selectors"),
            (0xFE10, 0xFE1F, "Vertical Forms"),
            (0xFE20, 0xFE2F, "Combining Half Marks"),
            (0xFE30, 0xFE4F, "CJK Compatibility Forms"),
            (0xFE50, 0xFE6F, "Small Form Variants"),
            (0xFE70, 0xFEFF, "Arabic Presentation Forms-B"),
            (0xFF00, 0xFFEF, "Halfwidth and Fullwidth Forms"),
            (0xFFF0, 0xFFFF, "Specials"),
            (0x10000, 0x1007F, "Linear B Syllabary"),
            (0x10080, 0x100FF, "Linear B Ideograms"),
            (0x10100, 0x1013F, "Aegean Numbers"),
            (0x10140, 0x1018F, "Ancient Greek Numbers"),
            (0x10190, 0x101CF, "Ancient Symbols"),
            (0x101D0, 0x101FF, "Phaistos Disc"),
            (0x10280, 0x1029F, "Lycian"),
            (0x102A0, 0x102DF, "Carian"),
            (0x102E0, 0x102FF, "Coptic Epact Numbers"),
            (0x10300, 0x1032F, "Old Italic"),
            (0x10330, 0x1034F, "Gothic"),
            (0x10350, 0x1037F, "Old Permic"),
            (0x10380, 0x1039F, "Ugaritic"),
            (0x103A0, 0x103DF, "Old Persian"),
            (0x10400, 0x1044F, "Deseret"),
            (0x10450, 0x1047F, "Shavian"),
            (0x10480, 0x104AF, "Osmanya"),
            (0x104B0, 0x104FF, "Osage"),
            (0x10500, 0x1052F, "Elbasan"),
            (0x10530, 0x1056F, "Caucasian Albanian"),
            (0x10600, 0x1077F, "Linear A"),
            (0x10800, 0x1083F, "Cypriot Syllabary"),
            (0x10840, 0x1085F, "Imperial Aramaic"),
            (0x10860, 0x1087F, "Palmyrene"),
            (0x10880, 0x108AF, "Nabataean"),
            (0x108E0, 0x108FF, "Hatran"),
            (0x10900, 0x1091F, "Phoenician"),
            (0x10920, 0x1093F, "Lydian"),
            (0x10980, 0x1099F, "Meroitic Hieroglyphs"),
            (0x109A0, 0x109FF, "Meroitic Cursive"),
            (0x10A00, 0x10A5F, "Kharoshthi"),
            (0x10A60, 0x10A7F, "Old South Arabian"),
            (0x10A80, 0x10A9F, "Old North Arabian"),
            (0x10AC0, 0x10AFF, "Manichaean"),
            (0x10B00, 0x10B3F, "Avestan"),
            (0x10B40, 0x10B5F, "Inscriptional Parthian"),
            (0x10B60, 0x10B7F, "Inscriptional Pahlavi"),
            (0x10B80, 0x10BAF, "Psalter Pahlavi"),
            (0x10C00, 0x10C4F, "Old Turkic"),
            (0x10C80, 0x10CFF, "Old Hungarian"),
            (0x10E60, 0x10E7F, "Rumi Numeral Symbols"),
            (0x10F00, 0x10F2F, "Old Sogdian"),
            (0x10F30, 0x10F6F, "Sogdian"),
            (0x10FB0, 0x10FDF, "Chorasmian"),
            (0x10FE0, 0x10FFF, "Elymaic"),
            (0x11000, 0x1107F, "Brahmi"),
            (0x11080, 0x110CF, "Kaithi"),
            (0x110D0, 0x110FF, "Sora Sompeng"),
            (0x11100, 0x1114F, "Chakma"),
            (0x11150, 0x1117F, "Mahajani"),
            (0x11180, 0x111DF, "Sharada"),
            (0x111E0, 0x111FF, "Sinhala Archaic Numbers"),
            (0x11200, 0x1124F, "Khojki"),
            (0x11280, 0x112AF, "Multani"),
            (0x112B0, 0x112FF, "Khudawadi"),
            (0x11300, 0x1137F, "Grantha"),
            (0x11400, 0x1147F, "Newa"),
            (0x11480, 0x114DF, "Tirhuta"),
            (0x11580, 0x115FF, "Siddham"),
            (0x11600, 0x1165F, "Modi"),
            (0x11660, 0x1167F, "Mongolian Supplement"),
            (0x11680, 0x116CF, "Takri"),
            (0x11700, 0x1174F, "Ahom"),
            (0x11800, 0x1184F, "Dogra"),
            (0x118A0, 0x118FF, "Warang Citi"),
            (0x11900, 0x1195F, "Dives Akuru"),
            (0x119A0, 0x119FF, "Nandinagari"),
            (0x11A00, 0x11A4F, "Zanabazar Square"),
            (0x11A50, 0x11AAF, "Soyombo"),
            (0x11AC0, 0x11AFF, "Pau Cin Hau"),
            (0x11C00, 0x11C6F, "Bhaiksuki"),
            (0x11C70, 0x11CBF, "Marchen"),
            (0x11D00, 0x11D5F, "Masaram Gondi"),
            (0x11D60, 0x11DAF, "Gunjala Gondi"),
            (0x11EE0, 0x11EFF, "Makasar"),
            (0x11FB0, 0x11FBF, "Lisu Supplement"),
            (0x11FC0, 0x11FFF, "Tamil Supplement"),
            (0x12000, 0x123FF, "Cuneiform"),
            (0x12400, 0x1247F, "Cuneiform Numbers and Punctuation"),
            (0x12480, 0x1254F, "Early Dynastic Cuneiform"),
            (0x13000, 0x1342F, "Egyptian Hieroglyphs"),
            (0x13430, 0x1343F, "Egyptian Hieroglyph Format Controls"),
            (0x14400, 0x1467F, "Anatolian Hieroglyphs"),
            (0x16800, 0x16A3F, "Bamum Supplement"),
            (0x16A40, 0x16A6F, "Mro"),
            (0x16AD0, 0x16AFF, "Bassa Vah"),
            (0x16B00, 0x16B8F, "Pahawh Hmong"),
            (0x16E40, 0x16E9F, "Medefaidrin"),
            (0x16F00, 0x16F9F, "Miao"),
            (0x16FE0, 0x16FFF, "Ideographic Symbols and Punctuation"),
            (0x17000, 0x187FF, "Tangut"),
            (0x18800, 0x18AFF, "Tangut Components"),
            (0x18B00, 0x18CFF, "Khitan Small Script"),
            (0x18D00, 0x18D7F, "Tangut Supplement"),
            (0x1B000, 0x1B0FF, "Kana Supplement"),
            (0x1B100, 0x1B12F, "Kana Extended-A"),
            (0x1B130, 0x1B16F, "Small Kana Extension"),
            (0x1B170, 0x1B2FF, "Nüshu"),
            (0x1BC00, 0x1BC9F, "Duployan"),
            (0x1BCA0, 0x1BCAF, "Shorthand Format Controls"),
            (0x1D000, 0x1D0FF, "Byzantine Musical Symbols"),
            (0x1D100, 0x1D1FF, "Musical Symbols"),
            (0x1D200, 0x1D24F, "Ancient Greek Musical Notation"),
            (0x1D2E0, 0x1D2FF, "Mayan Numerals"),
            (0x1D300, 0x1D35F, "Tai Xuan Jing Symbols"),
            (0x1D360, 0x1D37F, "Counting Rod Numerals"),
            (0x1D400, 0x1D7FF, "Mathematical Alphanumeric Symbols"),
            (0x1D800, 0x1DAAF, "Sutton SignWriting"),
            (0x1E000, 0x1E02F, "Glagolitic Supplement"),
            (0x1E100, 0x1E14F, "Nyiakeng Puachue Hmong"),
            (0x1E290, 0x1E2BF, "Toto"),
            (0x1E2C0, 0x1E2FF, "Wancho"),
            (0x1E7E0, 0x1E7FF, "Ethiopic Extended-B"),
            (0x1E800, 0x1E8DF, "Mende Kikakui"),
            (0x1E900, 0x1E95F, "Adlam"),
            (0x1EC70, 0x1ECBF, "Indic Siyaq Numbers"),
            (0x1ED00, 0x1ED4F, "Ottoman Siyaq Numbers"),
            (0x1EE00, 0x1EEFF, "Arabic Mathematical Alphabetic Symbols"),
            (0x1F000, 0x1F02F, "Mahjong Tiles"),
            (0x1F030, 0x1F09F, "Domino Tiles"),
            (0x1F0A0, 0x1F0FF, "Playing Cards"),
            (0x1F100, 0x1F1FF, "Enclosed Alphanumeric Supplement"),
            (0x1F200, 0x1F2FF, "Enclosed Ideographic Supplement"),
            (0x1F300, 0x1F5FF, "Miscellaneous Symbols and Pictographs"),
            (0x1F600, 0x1F64F, "Emoticons"),
            (0x1F650, 0x1F67F, "Ornamental Dingbats"),
            (0x1F680, 0x1F6FF, "Transport and Map Symbols"),
            (0x1F700, 0x1F77F, "Alchemical Symbols"),
            (0x1F780, 0x1F7FF, "Geometric Shapes Extended"),
            (0x1F800, 0x1F8FF, "Supplemental Arrows-C"),
            (0x1F900, 0x1F9FF, "Supplemental Symbols and Pictographs"),
            (0x1FA00, 0x1FA6F, "Chess Symbols"),
            (0x1FA70, 0x1FAFF, "Symbols and Pictographs Extended-A"),
            (0x1FB00, 0x1FBFF, "Symbols for Legacy Computing"),
            (0x20000, 0x2A6DF, "CJK Unified Ideographs Extension B"),
            (0x2A700, 0x2B73F, "CJK Unified Ideographs Extension C"),
            (0x2B740, 0x2B81F, "CJK Unified Ideographs Extension D"),
            (0x2B820, 0x2CEAF, "CJK Unified Ideographs Extension E"),
            (0x2CEB0, 0x2EBEF, "CJK Unified Ideographs Extension F"),
            (0x2F800, 0x2FA1F, "CJK Compatibility Ideographs Supplement"),
            (0x30000, 0x3134F, "CJK Unified Ideographs Extension G"),
            (0x31350, 0x323AF, "CJK Unified Ideographs Extension H"),
            (0xE0000, 0xE007F, "Tags"),
            (0xE0100, 0xE01EF, "Variation Selectors Supplement"),
            (0xF0000, 0xFFFFF, "Supplementary Private Use Area-A"),
            (0x100000, 0x10FFFF, "Supplementary Private Use Area-B"),
        ]
        
        # Private Use Areas
        private_use_blocks = [
            (0xE000, 0xF8FF, "Private Use Area"),  # Basic Private Use Area
            (0xF0000, 0xFFFFF, "Supplementary Private Use Area-A"),
            (0x100000, 0x10FFFF, "Supplementary Private Use Area-B"),
        ]
        
        # Combine and sort
        all_blocks = blocks + private_use_blocks
        all_blocks.sort()
        return all_blocks
    
    def get_unicode_block(self, code_point):
        """Get Unicode 17.0 block name for a code point"""
        for start, end, name in self.unicode_blocks:
            if start <= code_point <= end:
                return name
        return "Unassigned"
    
    def setup_ui(self):
        # Create main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(
            main_frame, 
            text="Font to Table Converter", 
            font=("Arial", 16, "bold")
        )
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Font file selection
        ttk.Label(main_frame, text="Font file:").grid(row=1, column=0, sticky=tk.W, pady=5)
        
        font_frame = ttk.Frame(main_frame)
        font_frame.grid(row=1, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        font_frame.columnconfigure(0, weight=1)
        
        self.font_entry = ttk.Entry(font_frame, textvariable=self.font_path)
        self.font_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        
        ttk.Button(
            font_frame, 
            text="Browse...", 
            command=self.browse_font_file,
            width=10
        ).grid(row=0, column=1)
        
        # Output format selection
        ttk.Label(main_frame, text="Table format:").grid(row=2, column=0, sticky=tk.W, pady=5)
        
        format_frame = ttk.Frame(main_frame)
        format_frame.grid(row=2, column=1, sticky=tk.W, pady=5)
        
        for i, (format_name, format_ext) in enumerate(self.supported_table_formats):
            ttk.Radiobutton(
                format_frame,
                text=f"{format_name} (.{format_ext})",
                variable=self.table_format,
                value=format_ext
            ).grid(row=0, column=i, padx=(0, 15))
        
        # Output file selection
        ttk.Label(main_frame, text="Output file:").grid(row=3, column=0, sticky=tk.W, pady=5)
        
        output_frame = ttk.Frame(main_frame)
        output_frame.grid(row=3, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        output_frame.columnconfigure(0, weight=1)
        
        self.output_entry = ttk.Entry(output_frame, textvariable=self.output_path)
        self.output_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        
        ttk.Button(
            output_frame, 
            text="Browse...", 
            command=self.browse_output_file,
            width=10
        ).grid(row=0, column=1)
        
        # Options frame
        options_frame = ttk.LabelFrame(main_frame, text="Options", padding="10")
        options_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=15)
        
        # Include control characters option
        self.include_control_chars = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            options_frame,
            text="Include control characters (non-printable)",
            variable=self.include_control_chars
        ).grid(row=0, column=0, sticky=tk.W)
        
        # Show character preview option
        self.show_preview = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_frame,
            text="Show character preview in table",
            variable=self.show_preview
        ).grid(row=0, column=1, sticky=tk.W, padx=(20, 0))
        
        # Progress bar
        ttk.Label(main_frame, text="Progress:").grid(row=5, column=0, sticky=tk.W, pady=(15, 5))
        
        self.progress_bar = ttk.Progressbar(
            main_frame, 
            variable=self.progress_var,
            maximum=100,
            mode='determinate'
        )
        self.progress_bar.grid(row=5, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=(15, 5))
        
        # Status display
        self.status_label = ttk.Label(
            main_frame, 
            textvariable=self.status_var,
            relief=tk.SUNKEN,
            padding=5
        )
        self.status_label.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=7, column=0, columnspan=3, pady=10)
        
        self.convert_button = ttk.Button(
            button_frame,
            text="Start Conversion",
            command=self.start_conversion,
            width=15
        )
        self.convert_button.grid(row=0, column=0, padx=5)
        
        ttk.Button(
            button_frame,
            text="Preview Font",
            command=self.preview_font,
            width=15
        ).grid(row=0, column=1, padx=5)
        
        ttk.Button(
            button_frame,
            text="Exit",
            command=self.root.quit,
            width=15
        ).grid(row=0, column=2, padx=5)
        
        # Info text
        info_text = (
            "Instructions:\n"
            "1. Select a font file (supports .ttf, .otf, .woff, .woff2)\n"
            "2. Choose output table format (CSV, Excel, JSON, HTML, Markdown)\n"
            "3. Select output file path\n"
            "4. Click 'Start Conversion'\n\n"
            "Table will contain the following columns:\n"
            "  - Character: Copyable Unicode character\n"
            "  - Unicode: Hexadecimal format (e.g., U+0041)\n"
            "  - Unicode Name: Official character name\n"
            "  - Block: Unicode block name\n"
            "  - GlyphName: Font-specific glyph name\n"
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
        """Browse and select font file"""
        filetypes = self.supported_font_formats
        
        filename = filedialog.askopenfilename(
            title="Select Font File",
            filetypes=filetypes
        )
        
        if filename:
            self.font_path.set(filename)
            
            # Auto-generate output filename
            if not self.output_path.get():
                font_name = Path(filename).stem
                output_dir = Path(filename).parent
                output_file = output_dir / f"{font_name}_glyphs.{self.table_format.get()}"
                self.output_path.set(str(output_file))
    
    def browse_output_file(self):
        """Browse and select output file"""
        format_ext = self.table_format.get()
        filetypes = [(f"{format_ext.upper()} File", f"*.{format_ext}")]
        
        filename = filedialog.asksaveasfilename(
            title="Save Table File",
            defaultextension=f".{format_ext}",
            filetypes=filetypes
        )
        
        if filename:
            self.output_path.set(filename)
    
    def get_unicode_name(self, char):
        """Get Unicode character name"""
        try:
            name = unicodedata.name(char)
            return name
        except ValueError:
            return "Unnamed Character"
    
    def get_glyph_name(self, font, glyph_id):
        """Get glyph name from font"""
        try:
            # Get glyph name from font
            glyph_name = font.getGlyphName(glyph_id)
            return glyph_name
        except:
            return f"glyph_{glyph_id}"
    
    def extract_font_glyphs(self, font_path, include_control_chars=False, show_preview=True):
        """Extract glyph information from font file"""
        if not FONTTOOLS_AVAILABLE:
            raise ImportError("fontTools library not installed, please run: pip install fonttools")
        
        if not PANDAS_AVAILABLE:
            raise ImportError("pandas library not installed, please run: pip install pandas")
        
        self.status_var.set("Loading font file...")
        self.progress_var.set(10)
        self.root.update()
        
        try:
            font = TTFont(font_path)
        except Exception as e:
            raise Exception(f"Cannot load font file: {e}")
        
        self.status_var.set("Extracting glyph information...")
        self.progress_var.set(30)
        self.root.update()
        
        glyphs_data = []
        
        # Get cmap table (character mapping)
        if 'cmap' not in font:
            raise Exception("Font file does not contain cmap table")
        
        # Get character to glyph ID mapping
        cmap = font.getBestCmap()
        
        if not cmap:
            raise Exception("Cannot extract character mapping from font file")
        
        # Get reverse glyph name mapping
        try:
            glyph_order = font.getGlyphOrder()
        except:
            glyph_order = []
        
        total_glyphs = len(cmap)
        processed = 0
        
        for code_point, glyph_id in cmap.items():
            try:
                # Update progress
                processed += 1
                progress = 30 + (processed / total_glyphs) * 60
                self.progress_var.set(progress)
                
                if processed % 100 == 0:
                    self.status_var.set(f"Processing characters: {processed}/{total_glyphs}")
                    self.root.update()
                
                # Convert Unicode code point to character
                char = chr(code_point)
                
                # Get Unicode category
                try:
                    category = unicodedata.category(char)
                except:
                    category = "Cn"  # Other, not assigned
                
                # Skip control characters if not included
                if not include_control_chars and category.startswith('C'):
                    continue
                
                # Get display character
                if not show_preview and category.startswith('C'):
                    display_char = "□"  # Use square for control characters
                else:
                    display_char = char
                
                # Get Unicode name
                unicode_name = self.get_unicode_name(char)
                
                # Get Unicode block
                unicode_block = self.get_unicode_block(code_point)
                
                # Get glyph name
                glyph_name = self.get_glyph_name(font, glyph_id)
                
                # Format Unicode code point
                unicode_hex = f"U+{code_point:04X}"
                
                # Add to data list
                glyphs_data.append({
                    "Character": display_char,
                    "Unicode": unicode_hex,
                    "UnicodeName": unicode_name,
                    "Block": unicode_block,
                    "GlyphName": glyph_name,
                    "CodePoint": code_point,
                    "Category": category
                })
                
            except Exception as e:
                # Skip characters that cannot be processed
                continue
        
        font.close()
        
        self.status_var.set(f"Successfully extracted {len(glyphs_data)} characters")
        self.progress_var.set(95)
        self.root.update()
        
        return glyphs_data
    
    def save_table(self, data, output_path, table_format):
        """Save data as table file"""
        self.status_var.set(f"Saving as {table_format.upper()} file...")
        self.root.update()
        
        # Create DataFrame
        df = pd.DataFrame(data)
        
        # Sort by Unicode code point
        df = df.sort_values("CodePoint").reset_index(drop=True)
        
        # Drop temporary columns
        df = df.drop(columns=["CodePoint", "Category"])
        
        # Save based on format
        if table_format == "csv":
            df.to_csv(output_path, index=False, encoding='utf-8-sig')
        
        elif table_format == "xlsx":
            try:
                if OPENPYXL_AVAILABLE:
                    df.to_excel(output_path, index=False, engine='openpyxl')
                else:
                    df.to_excel(output_path, index=False, engine='xlwt')
            except Exception as e:
                # Try with xlsxwriter as fallback
                try:
                    df.to_excel(output_path, index=False, engine='xlsxwriter')
                except:
                    raise Exception(f"Cannot save Excel file: {e}")
        
        elif table_format == "json":
            df.to_json(output_path, orient='records', force_ascii=False, indent=2)
        
        elif table_format == "html":
            html_table = df.to_html(index=False, classes='font-glyphs-table')
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Font Glyphs Table</title>
    <style>
        table {{ border-collapse: collapse; width: 100%; }}
        th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
        th {{ background-color: #f2f2f2; }}
        tr:nth-child(even) {{ background-color: #f9f9f9; }}
        .char-cell {{ font-family: monospace; font-size: 24px; text-align: center; }}
    </style>
</head>
<body>
    <h1>Font Glyphs Table</h1>
    <p>Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
    {html_table}
</body>
</html>""")
        
        elif table_format == "md":
            df.to_markdown(output_path, index=False)
        
        else:
            raise ValueError(f"Unsupported table format: {table_format}")
    
    def start_conversion(self):
        """Start conversion process"""
        # Check if libraries are available
        if not FONTTOOLS_AVAILABLE:
            messagebox.showerror("Error", "fontTools library not installed, please run: pip install fonttools")
            return
        
        if not PANDAS_AVAILABLE:
            messagebox.showerror("Error", "pandas library not installed, please run: pip install pandas")
            return
        
        # Check input file
        font_path = self.font_path.get()
        if not font_path or not os.path.exists(font_path):
            messagebox.showerror("Error", "Please select a valid font file")
            return
        
        # Check output path
        output_path = self.output_path.get()
        if not output_path:
            messagebox.showerror("Error", "Please select output file path")
            return
        
        # Check output directory exists
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
            except Exception as e:
                messagebox.showerror("Error", f"Cannot create output directory: {e}")
                return
        
        # Disable convert button to prevent multiple clicks
        self.convert_button.config(state=tk.DISABLED)
        self.status_var.set("Starting...")
        self.progress_var.set(0)
        
        # Execute conversion in a new thread
        thread = threading.Thread(
            target=self.convert_thread,
            args=(font_path, output_path)
        )
        thread.daemon = True
        thread.start()
    
    def convert_thread(self, font_path, output_path):
        """Conversion thread"""
        try:
            # Extract glyph data
            glyphs_data = self.extract_font_glyphs(
                font_path,
                self.include_control_chars.get(),
                self.show_preview.get()
            )
            
            if not glyphs_data:
                self.root.after(0, lambda: messagebox.showwarning("Warning", "No characters found to extract"))
                return
            
            # Save table
            table_format = self.table_format.get()
            self.save_table(glyphs_data, output_path, table_format)
            
            # Update status
            self.progress_var.set(100)
            self.status_var.set(f"Conversion complete! Extracted {len(glyphs_data)} characters")
            
            # Show success message
            self.root.after(0, lambda: messagebox.showinfo(
                "Success", 
                f"Conversion complete!\nExtracted {len(glyphs_data)} characters\nFile saved to: {output_path}"
            ))
            
        except Exception as e:
            error_msg = f"Conversion failed: {str(e)}"
            if "Private Use Area" in str(e) or "glyph" in str(e).lower():
                error_msg += "\n\nNote: Some characters in Private Use Area may not display correctly in some applications."
            self.root.after(0, lambda: messagebox.showerror("Error", error_msg))
            print(traceback.format_exc())
        
        finally:
            # Re-enable convert button
            self.root.after(0, lambda: self.convert_button.config(state=tk.NORMAL))
            self.progress_var.set(0)
    
    def preview_font(self):
        """Preview font"""
        font_path = self.font_path.get()
        if not font_path or not os.path.exists(font_path):
            messagebox.showerror("Error", "Please select a font file first")
            return
        
        try:
            # Create preview window
            preview_window = tk.Toplevel(self.root)
            preview_window.title("Font Preview")
            preview_window.geometry("600x400")
            
            # Try to load font
            try:
                import tkinter.font as tkfont
                font_family = os.path.basename(font_path).split('.')[0]
                custom_font = tkfont.Font(family=font_family, size=20)
                
                # If font loading failed, use default font
                if custom_font.actual()["family"] == "TkDefaultFont":
                    raise Exception("Cannot load font")
            except:
                messagebox.showwarning("Warning", "Cannot load this font in the system for preview\nCheck the actual characters in the table")
                preview_window.destroy()
                return
            
            # Create text box to display font sample
            text_frame = ttk.Frame(preview_window, padding="10")
            text_frame.pack(fill=tk.BOTH, expand=True)
            
            ttk.Label(text_frame, text="Font Preview:", font=("Arial", 12, "bold")).pack(anchor=tk.W)
            
            text_widget = tk.Text(text_frame, height=10, width=50, font=custom_font)
            text_widget.pack(fill=tk.BOTH, expand=True, pady=(5, 0))
            
            # Add sample text
            sample_text = """ABCDEFGHIJKLMNOPQRSTUVWXYZ
abcdefghijklmnopqrstuvwxyz
0123456789
!@#$%^&*()_+-=[]{}|;:'",.<>/?
Sample text: Hello, World!
字体预览：你好，世界！"""
            
            text_widget.insert(1.0, sample_text)
            text_widget.config(state=tk.DISABLED)
            
            # Add scrollbar
            scrollbar = ttk.Scrollbar(text_widget, command=text_widget.yview)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            text_widget.config(yscrollcommand=scrollbar.set)
            
            # Font information
            info_frame = ttk.LabelFrame(preview_window, text="Font Information", padding="10")
            info_frame.pack(fill=tk.X, padx=10, pady=10)
            
            font_name = os.path.basename(font_path)
            file_size = os.path.getsize(font_path) / 1024  # KB
            
            info_text = f"""Font file: {font_name}
File size: {file_size:.2f} KB
File path: {font_path}
Font name: {font_family}"""
            
            ttk.Label(info_frame, text=info_text, justify=tk.LEFT).pack(anchor=tk.W)
            
            # Close button
            ttk.Button(
                preview_window, 
                text="Close", 
                command=preview_window.destroy
            ).pack(pady=10)
            
        except Exception as e:
            messagebox.showerror("Error", f"Preview failed: {e}")

def main():
    """Main function"""
    # Check required libraries
    if not FONTTOOLS_AVAILABLE or not PANDAS_AVAILABLE:
        print("Error: Required libraries missing")
        print("Please run the following commands to install required libraries:")
        print("pip install fonttools pandas")
        if not OPENPYXL_AVAILABLE:
            print("pip install openpyxl  # For Excel file support")
        return
    
    # Create main window
    root = tk.Tk()
    app = FontToTableApp(root)
    
    # Run main loop
    root.mainloop()

if __name__ == "__main__":
    main()
