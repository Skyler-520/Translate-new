# -*- coding: utf-8 -*-
"""
自动化翻译处理工具 - GUI版本
功能描述:
    1. 支持文件夹和文件两种输入模式
    2. 支持超过100种语言的翻译
    3. 提供翻译确认机制和进度显示
    4. 自动生成各语言版本的XML文件
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, simpledialog
import xml.etree.ElementTree as ET
import json
import os
import requests
from urllib.parse import quote
import time
import threading
import hashlib
import uuid

# 语言映射字典 - 包含语言代码、名称和中文说明
# 格式: {'语言缩写': {'code': 'Google翻译代码', 'name': '语言名称 (中文语言:中文国家)'}}
LANG_MAP = {
    'CAT': {'code': 'ca', 'name': 'Catalan: Spain (加泰罗尼亚语:西班牙)'},
    'CHT': {'code': 'zh-TW', 'name': 'Chinese: Traditional (Taiwan) (中文繁体:中国台湾)'},
    'CHS': {'code': 'zh-CN', 'name': 'Chinese: Simplified (PRC) (中文简体:中国)'},
    'ZHH': {'code': 'zh-HK', 'name': 'Chinese: Hong Kong S.A.R. (中文:中国香港)'},
    'ZHI': {'code': 'zh-SG', 'name': 'Chinese: Singapore (中文:新加坡)'},
    'ZHM': {'code': 'zh-MO', 'name': 'Chinese: Macau SAR (中文:中国澳门)'},
    'ARA': {'code': 'ar-SA', 'name': 'Arabic: Saudi Arabia (阿拉伯语:沙特阿拉伯)'},
    'ARI': {'code': 'ar-IQ', 'name': 'Arabic: Iraq (阿拉伯语:伊拉克)'},
    'ARE': {'code': 'ar-EG', 'name': 'Arabic: Egypt (阿拉伯语:埃及)'},
    'ARL': {'code': 'ar-LY', 'name': 'Arabic: Libya (阿拉伯语:利比亚)'},
    'ARG': {'code': 'ar-DZ', 'name': 'Arabic: Algeria (阿拉伯语:阿尔及利亚)'},
    'ARM': {'code': 'ar-MA', 'name': 'Arabic: Morocco (阿拉伯语:摩洛哥)'},
    'ART': {'code': 'ar-TN', 'name': 'Arabic: Tunisia (阿拉伯语:突尼斯)'},
    'ARO': {'code': 'ar-OM', 'name': 'Arabic: Oman (阿拉伯语:阿曼)'},
    'ARY': {'code': 'ar-YE', 'name': 'Arabic: Yemen (阿拉伯语:也门)'},
    'ARS': {'code': 'ar-SY', 'name': 'Arabic: Syria (阿拉伯语:叙利亚)'},
    'ARJ': {'code': 'ar-JO', 'name': 'Arabic: Jordan (阿拉伯语:约旦)'},
    'ARB': {'code': 'ar-LB', 'name': 'Arabic: Lebanon (阿拉伯语:黎巴嫩)'},
    'ARK': {'code': 'ar-KW', 'name': 'Arabic: Kuwait (阿拉伯语:科威特)'},
    'ARU': {'code': 'ar-AE', 'name': 'Arabic: U.A.E. (阿拉伯语:阿联酋)'},
    'ARH': {'code': 'ar-BH', 'name': 'Arabic: Bahrain (阿拉伯语:巴林)'},
    'ARQ': {'code': 'ar-QA', 'name': 'Arabic: Qatar (阿拉伯语:卡塔尔)'},
    'BGR': {'code': 'bg', 'name': 'Bulgarian: Bulgaria (保加利亚语:保加利亚)'},
    'CSY': {'code': 'cs', 'name': 'Czech: Czech Republic (捷克语:捷克)'},
    'DAN': {'code': 'da', 'name': 'Danish: Denmark (丹麦语:丹麦)'},
    'GER': {'code': 'de-DE', 'name': 'German: Germany (德语:德国)'},
    'DES': {'code': 'de-CH', 'name': 'German: Switzerland (德语:瑞士)'},
    'DEA': {'code': 'de-AT', 'name': 'German: Austria (德语:奥地利)'},
    'DEL': {'code': 'de-LU', 'name': 'German: Luxembourg (德语:卢森堡)'},
    'DEC': {'code': 'de-LI', 'name': 'German: Liechtenstein (德语:列支敦士登)'},
    'ELL': {'code': 'el', 'name': 'Greek: Greece (希腊语:希腊)'},
    'USA': {'code': 'en-US', 'name': 'English: United States (英语:美国)'},
    'ENG': {'code': 'en-GB', 'name': 'English: United Kingdom (英语:英国)'},
    'ENA': {'code': 'en-AU', 'name': 'English: Australia (英语:澳大利亚)'},
    'ENC': {'code': 'en-CA', 'name': 'English: Canada (英语:加拿大)'},
    'ENZ': {'code': 'en-NZ', 'name': 'English: New Zealand (英语:新西兰)'},
    'ENI': {'code': 'en-IE', 'name': 'English: Ireland (英语:爱尔兰)'},
    'ENS': {'code': 'en-ZA', 'name': 'English: South Africa (英语:南非)'},
    'ESP': {'code': 'es-ES', 'name': 'Spanish: Spain (西班牙语:西班牙)'},
    'ESM': {'code': 'es-MX', 'name': 'Spanish: Mexico (西班牙语:墨西哥)'},
    'ESG': {'code': 'es-GT', 'name': 'Spanish: Guatemala (西班牙语:危地马拉)'},
    'ESC': {'code': 'es-CR', 'name': 'Spanish: Costa Rica (西班牙语:哥斯达黎加)'},
    'ESA': {'code': 'es-PA', 'name': 'Spanish: Panama (西班牙语:巴拿马)'},
    'ESD': {'code': 'es-DO', 'name': 'Spanish: Dominican Republic (西班牙语:多米尼加)'},
    'ESV': {'code': 'es-VE', 'name': 'Spanish: Venezuela (西班牙语:委内瑞拉)'},
    'ESO': {'code': 'es-CO', 'name': 'Spanish: Colombia (西班牙语:哥伦比亚)'},
    'ESR': {'code': 'es-PE', 'name': 'Spanish: Peru (西班牙语:秘鲁)'},
    'ESS': {'code': 'es-AR', 'name': 'Spanish: Argentina (西班牙语:阿根廷)'},
    'ESF': {'code': 'es-EC', 'name': 'Spanish: Ecuador (西班牙语:厄瓜多尔)'},
    'ESL': {'code': 'es-CL', 'name': 'Spanish: Chile (西班牙语:智利)'},
    'ESY': {'code': 'es-UY', 'name': 'Spanish: Uruguay (西班牙语:乌拉圭)'},
    'ESZ': {'code': 'es-PY', 'name': 'Spanish: Paraguay (西班牙语:巴拉圭)'},
    'ESB': {'code': 'es-BO', 'name': 'Spanish: Bolivia (西班牙语:玻利维亚)'},
    'ESE': {'code': 'es-SV', 'name': 'Spanish: El Salvador (西班牙语:萨尔瓦多)'},
    'ESH': {'code': 'es-HN', 'name': 'Spanish: Honduras (西班牙语:洪都拉斯)'},
    'ESI': {'code': 'es-NI', 'name': 'Spanish: Nicaragua (西班牙语:尼加拉瓜)'},
    'ESU': {'code': 'es-PR', 'name': 'Spanish: Puerto Rico (西班牙语:波多黎各)'},
    'FIN': {'code': 'fi', 'name': 'Finnish: Finland (芬兰语:芬兰)'},
    'FRA': {'code': 'fr-FR', 'name': 'French: France (法语:法国)'},
    'FRB': {'code': 'fr-BE', 'name': 'French: Belgium (法语:比利时)'},
    'FRC': {'code': 'fr-CA', 'name': 'French: Canada (法语:加拿大)'},
    'FRS': {'code': 'fr-CH', 'name': 'French: Switzerland (法语:瑞士)'},
    'FRL': {'code': 'fr-LU', 'name': 'French: Luxembourg (法语:卢森堡)'},
    'FRM': {'code': 'fr-MC', 'name': 'French: Monaco (法语:摩纳哥)'},
    'HEB': {'code': 'he', 'name': 'Hebrew: Israel (希伯来语:以色列)'},
    'HUN': {'code': 'hu', 'name': 'Hungarian: Hungary (匈牙利语:匈牙利)'},
    'ISL': {'code': 'is', 'name': 'Icelandic: Iceland (冰岛语:冰岛)'},
    'ITA': {'code': 'it-IT', 'name': 'Italian: Italy (意大利语:意大利)'},
    'ITS': {'code': 'it-CH', 'name': 'Italian: Switzerland (意大利语:瑞士)'},
    'JPN': {'code': 'ja', 'name': 'Japanese: Japan (日语:日本)'},
    'KOR': {'code': 'ko', 'name': 'Korean: Korea (韩语:韩国)'},
    'NLD': {'code': 'nl-NL', 'name': 'Dutch: Netherlands (荷兰语:荷兰)'},
    'NLB': {'code': 'nl-BE', 'name': 'Dutch: Belgium (荷兰语:比利时)'},
    'NOR': {'code': 'nb', 'name': 'Norwegian: Norway (Bokmål) (挪威语:挪威)'},
    'NON': {'code': 'nn', 'name': 'Norwegian: Norway (Nynorsk) (挪威语:挪威)'},
    'PLK': {'code': 'pl', 'name': 'Polish: Poland (波兰语:波兰)'},
    'PTB': {'code': 'pt-BR', 'name': 'Portuguese: Brazil (葡萄牙语:巴西)'},
    'PTG': {'code': 'pt-PT', 'name': 'Portuguese: Portugal (葡萄牙语:葡萄牙)'},
    'ROM': {'code': 'ro', 'name': 'Romanian: Romania (罗马尼亚语:罗马尼亚)'},
    'RUS': {'code': 'ru', 'name': 'Russian: Russia (俄语:俄罗斯)'},
    'HRV': {'code': 'hr', 'name': 'Croatian: Croatia (克罗地亚语:克罗地亚)'},
    'SRL': {'code': 'sr-Latn', 'name': 'Serbian: Serbia (Latin) (塞尔维亚语:塞尔维亚)'},
    'SRB': {'code': 'sr-Cyrl', 'name': 'Serbian: Serbia (Cyrillic) (塞尔维亚语:塞尔维亚)'},
    'SKY': {'code': 'sk', 'name': 'Slovak: Slovakia (斯洛伐克语:斯洛伐克)'},
    'SQI': {'code': 'sq', 'name': 'Albanian: Albania (阿尔巴尼亚语:阿尔巴尼亚)'},
    'SVE': {'code': 'sv-SE', 'name': 'Swedish: Sweden (瑞典语:瑞典)'},
    'SVF': {'code': 'sv-FI', 'name': 'Swedish: Finland (瑞典语:芬兰)'},
    'THA': {'code': 'th', 'name': 'Thai: Thailand (泰语:泰国)'},
    'TRK': {'code': 'tr', 'name': 'Turkish: Turkey (土耳其语:土耳其)'},
    'URP': {'code': 'ur-PK', 'name': 'Urdu: Pakistan (乌尔都语:巴基斯坦)'},
    'IND': {'code': 'id', 'name': 'Indonesian: Indonesia (印尼语:印尼)'},
    'UKR': {'code': 'uk', 'name': 'Ukrainian: Ukraine (乌克兰语:乌克兰)'},
    'BEL': {'code': 'be', 'name': 'Belarusian: Belarus (白俄罗斯语:白俄罗斯)'},
    'SLV': {'code': 'sl', 'name': 'Slovene: Slovenia (斯洛文尼亚语:斯洛文尼亚)'},
    'ETI': {'code': 'et', 'name': 'Estonian: Estonia (爱沙尼亚语:爱沙尼亚)'},
    'LVI': {'code': 'lv', 'name': 'Latvian: Latvia (拉脱维亚语:拉脱维亚)'},
    'LTH': {'code': 'lt', 'name': 'Lithuanian: Lithuania (立陶宛语:立陶宛)'},
    'LTC': {'code': 'lt', 'name': 'Classic Lithuanian: Lithuania (立陶宛语:立陶宛)'},
    'FAR': {'code': 'fa', 'name': 'Farsi: Iran (波斯语:伊朗)'},
    'VIT': {'code': 'vi', 'name': 'Vietnamese: Vietnam (越南语:越南)'},
    'HYE': {'code': 'hy', 'name': 'Armenian: Armenia (亚美尼亚语:亚美尼亚)'},
    'AZE': {'code': 'az-Latn', 'name': 'Azeri: Azerbaijan (Latin) (阿塞拜疆语:阿塞拜疆)'},
    'EUQ': {'code': 'eu', 'name': 'Basque: Spain (巴斯克语:西班牙)'},
    'MKI': {'code': 'mk', 'name': 'FYRO Macedonian (马其顿语:马其顿)'},
    'AFK': {'code': 'af', 'name': 'Afrikaans: South Africa (南非语:南非)'},
    'KAT': {'code': 'ka', 'name': 'Georgian: Georgia (格鲁吉亚语:格鲁吉亚)'},
    'FOS': {'code': 'fo', 'name': 'Faeroese: Faeroe Islands (法罗语:法罗群岛)'},
    'HIN': {'code': 'hi', 'name': 'Hindi: India (印地语:印度)'},
    'MSL': {'code': 'ms-MY', 'name': 'Malay: Malaysia (马来语:马来西亚)'},
    'MSB': {'code': 'ms-BN', 'name': 'Malay: Brunei Darussalam (马来语:文莱)'},
    'KAZ': {'code': 'kk', 'name': 'Kazak: Kazakhstan (哈萨克语:哈萨克斯坦)'},
    'SWK': {'code': 'sw', 'name': 'Swahili: Kenya (斯瓦希里语:肯尼亚)'},
    'UZB': {'code': 'uz-Latn', 'name': 'Uzbek: Uzbekistan (Latin) (乌兹别克语:乌兹别克斯坦)'},
    'TAT': {'code': 'tt', 'name': 'Tatar: Tatarstan (鞑靼语:鞑靼斯坦)'},
    'BEN': {'code': 'bn', 'name': 'Bengali: India (孟加拉语:印度)'},
    'PAN': {'code': 'pa', 'name': 'Punjabi: India (旁遮普语:印度)'},
    'GUJ': {'code': 'gu', 'name': 'Gujarati: India (古吉拉特语:印度)'},
    'ORI': {'code': 'or', 'name': 'Oriya: India (奥里亚语:印度)'},
    'TAM': {'code': 'ta', 'name': 'Tamil: India (泰米尔语:印度)'},
    'TEL': {'code': 'te', 'name': 'Telugu: India (泰卢固语:印度)'},
    'KAN': {'code': 'kn', 'name': 'Kannada: India (卡纳达语:印度)'},
    'MAL': {'code': 'ml', 'name': 'Malayalam: India (马拉雅拉姆语:印度)'},
    'ASM': {'code': 'as', 'name': 'Assamese: India (阿萨姆语:印度)'},
    'MAR': {'code': 'mr', 'name': 'Marathi: India (马拉地语:印度)'},
    'SAN': {'code': 'sa', 'name': 'Sanskrit: India (梵语:印度)'},
    'KOK': {'code': 'kok', 'name': 'Konkani: India (孔卡尼语:印度)'}
}

# 默认显示的语言列表（11种常用语言）
DEFAULT_LANGS = ['CAT', 'CHT', 'ESP', 'GER', 'ITA', 'KOR', 'PLK', 'PTG', 'RUS', 'TRK', 'VIT']

# 品牌名称集合 - 这些名称不翻译
BRAND_NAMES = {'禾川', '高创', '松下', '汇川', '迪维迅', '台达', '雷赛', '信捷', 'H系列'}

# 配置文件名
CONFIG_FILE = 'translation_config.json'


class TranslationApp:
    """
    自动化翻译处理工具主类
    提供可视化用户界面，支持XML文件翻译和多语言输出
    """
    
    def __init__(self, root):
        """
        初始化方法
        :param root: Tkinter主窗口对象
        """
        self.root = root
        self.root.title("自动化翻译处理工具")
        self.root.geometry("1000x800")
        
        # 初始化实例变量
        self.source_files = []           # 源文件列表
        self.source_folder = ""         # 源文件夹路径
        self.selected_langs = []        # 选中的语言列表
        self.translation_table = {}     # 翻译表字典
        self.output_dir = os.path.dirname(os.path.abspath(__file__))  # 输出目录
        self.default_langs = DEFAULT_LANGS.copy()  # 默认语言列表
        self.show_all_langs = False     # 是否显示所有语言
        self.translation_complete = False  # 翻译是否完成
        
        # API配置变量
        self.api_type = 'google'  # 当前使用的翻译API类型：google 或 baidu
        self.baidu_app_id = ''   # 百度翻译API App ID
        self.baidu_secret_key = ''  # 百度翻译API Secret Key
        
        # 加载配置、创建界面、加载翻译表
        self.load_config()
        self.create_widgets()
        self.load_translation_table()
    
    def load_config(self):
        """
        加载配置文件
        从translation_config.json读取用户自定义的默认语言设置和API配置
        """
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    if 'default_langs' in config:
                        self.default_langs = config['default_langs']
                    if 'api_type' in config:
                        self.api_type = config['api_type']
                    if 'baidu_app_id' in config:
                        self.baidu_app_id = config['baidu_app_id']
                    if 'baidu_secret_key' in config:
                        self.baidu_secret_key = config['baidu_secret_key']
            except Exception as e:
                self.log(f"加载配置失败: {e}")
    
    def save_config(self):
        """
        保存配置文件
        将当前默认语言设置和API配置保存到translation_config.json
        """
        config = {
            'default_langs': self.default_langs,
            'api_type': self.api_type,
            'baidu_app_id': self.baidu_app_id,
            'baidu_secret_key': self.baidu_secret_key
        }
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    
    def create_widgets(self):
        """
        创建用户界面组件
        包括输入源选择、语言选择、输出目录、进度条、日志等
        """
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        row = 0
        
        # 输入源选择区域
        ttk.Label(main_frame, text="输入源:", font=('Arial', 10, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=5)
        row += 1
        
        input_frame = ttk.Frame(main_frame)
        input_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E))
        
        self.input_mode_var = tk.StringVar(value='folder')
        ttk.Radiobutton(input_frame, text="选择文件夹", variable=self.input_mode_var, 
                        value='folder', command=self.switch_input_mode).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(input_frame, text="选择文件", variable=self.input_mode_var, 
                        value='files', command=self.switch_input_mode).pack(side=tk.LEFT, padx=5)
        row += 1
        
        self.folder_entry = ttk.Entry(main_frame, width=80)
        self.folder_entry.grid(row=row, column=0, sticky=(tk.W, tk.E), padx=5)
        self.folder_entry.insert(0, self.source_folder)
        
        folder_button_frame = ttk.Frame(main_frame)
        folder_button_frame.grid(row=row, column=1, sticky=tk.W)
        
        ttk.Button(folder_button_frame, text="浏览", command=self.select_source_folder).pack(side=tk.LEFT, padx=5)
        ttk.Button(folder_button_frame, text="刷新", command=self.refresh_file_list).pack(side=tk.LEFT, padx=5)
        row += 1
        
        # 文件列表
        self.file_listbox = tk.Listbox(main_frame, width=100, height=5)
        self.file_listbox.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        row += 1
        
        # 文件操作按钮
        file_button_frame = ttk.Frame(main_frame)
        file_button_frame.grid(row=row, column=0, columnspan=3, sticky=tk.W)
        
        ttk.Button(file_button_frame, text="添加文件", command=self.add_files).pack(side=tk.LEFT, padx=5)
        ttk.Button(file_button_frame, text="移除文件", command=self.remove_files).pack(side=tk.LEFT, padx=5)
        ttk.Button(file_button_frame, text="清空列表", command=self.clear_files).pack(side=tk.LEFT, padx=5)
        row += 1
        
        # 语言选择区域
        ttk.Label(main_frame, text="目标语言选择:", font=('Arial', 10, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=5)
        row += 1
        
        lang_control_frame = ttk.Frame(main_frame)
        lang_control_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E))
        
        self.lang_show_var = tk.BooleanVar(value=self.show_all_langs)
        ttk.Checkbutton(lang_control_frame, text="显示所有语言", variable=self.lang_show_var, 
                        command=self.toggle_lang_display).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(lang_control_frame, text="设置默认语言", command=self.set_default_langs).pack(side=tk.LEFT, padx=5)
        
        # API选择区域
        api_frame = ttk.Frame(lang_control_frame)
        api_frame.pack(side=tk.LEFT, padx=20)
        
        ttk.Label(api_frame, text="翻译API:").pack(side=tk.LEFT, padx=5)
        
        self.api_type_var = tk.StringVar(value=self.api_type)
        api_combo = ttk.Combobox(api_frame, textvariable=self.api_type_var, values=['google', 'baidu'], 
                                  state='readonly', width=8)
        api_combo.pack(side=tk.LEFT, padx=5)
        api_combo.bind('<<ComboboxSelected>>', lambda e: self.on_api_type_changed())
        
        ttk.Button(api_frame, text="API配置", command=self.configure_api).pack(side=tk.LEFT, padx=5)
        
        select_all_btn = ttk.Button(lang_control_frame, text="全选", command=self.select_all_langs)
        select_all_btn.pack(side=tk.LEFT, padx=5)
        
        deselect_all_btn = ttk.Button(lang_control_frame, text="取消全选", command=self.deselect_all_langs)
        deselect_all_btn.pack(side=tk.LEFT, padx=5)
        row += 1
        
        # 语言复选框框架
        self.lang_frame = ttk.Frame(main_frame)
        self.lang_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E))
        row += 1
        
        # 进度条
        ttk.Label(main_frame, text="处理进度:", font=('Arial', 10, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=5)
        row += 1
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100, length=800)
        self.progress_bar.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        self.progress_label = ttk.Label(main_frame, text="0%")
        self.progress_label.grid(row=row, column=3, sticky=tk.W, pady=5)
        row += 1
        
        # 日志区域
        ttk.Label(main_frame, text="处理日志:", font=('Arial', 10, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=5)
        row += 1
        
        self.log_text = scrolledtext.ScrolledText(main_frame, width=110, height=18, wrap=tk.WORD)
        self.log_text.grid(row=row, column=0, columnspan=4, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.log_text.insert(tk.END, "欢迎使用自动化翻译处理工具！\n")
        self.log_text.insert(tk.END, "请选择输入源（文件夹或文件）和目标语言，然后点击开始翻译按钮。\n")
        row += 1
        
        # 操作按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=row, column=0, columnspan=4, pady=10)
        
        self.start_btn = ttk.Button(button_frame, text="开始翻译", command=self.start_translation)
        self.start_btn.pack(side=tk.LEFT, padx=10)
        
        self.confirm_btn = ttk.Button(button_frame, text="确认翻译并生成XML", 
                                     command=self.confirm_and_generate, state='disabled')
        self.confirm_btn.pack(side=tk.LEFT, padx=10)
        
        ttk.Button(button_frame, text="查看翻译表", command=self.view_translation_table).pack(side=tk.LEFT, padx=10)
        
        ttk.Button(button_frame, text="导出日志", command=self.export_log).pack(side=tk.LEFT, padx=10)
        
        main_frame.columnconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # 更新语言显示
        self.update_lang_display()
    
    def switch_input_mode(self):
        """
        切换输入模式（文件夹/文件）
        清空当前文件列表和文件夹路径
        """
        self.source_files = []
        self.file_listbox.delete(0, tk.END)
        self.folder_entry.delete(0, tk.END)
        self.source_folder = ""
    
    def select_source_folder(self):
        """
        选择输入源（文件夹或文件）
        根据当前输入模式执行不同操作，自动设置输出目录为输入源的上一级目录
        """
        if self.input_mode_var.get() == 'folder':
            folder = filedialog.askdirectory(title="选择输入文件夹")
            if folder:
                self.source_folder = folder
                self.folder_entry.delete(0, tk.END)
                self.folder_entry.insert(0, folder)
                # 设置输出目录为输入文件夹的上一级目录
                self.output_dir = os.path.dirname(folder)
                self.log(f"输出目录已设置为: {self.output_dir}")
                self.refresh_file_list()
        else:
            files = filedialog.askopenfilenames(title="选择XML文件", filetypes=[("XML文件", "*.xml")])
            if files:
                # 设置输出目录为第一个文件所在目录的上一级目录
                first_file_dir = os.path.dirname(files[0])
                self.output_dir = os.path.dirname(first_file_dir)
                self.log(f"输出目录已设置为: {self.output_dir}")
            for file in files:
                if file not in self.source_files:
                    self.source_files.append(file)
                    self.file_listbox.insert(tk.END, file)
    
    def refresh_file_list(self):
        """
        刷新文件列表
        扫描源文件夹中的所有XML文件
        """
        self.source_files = []
        self.file_listbox.delete(0, tk.END)
        if self.source_folder and os.path.isdir(self.source_folder):
            for root_dir, dirs, files in os.walk(self.source_folder):
                for file in files:
                    if file.endswith('.xml'):
                        file_path = os.path.join(root_dir, file)
                        self.source_files.append(file_path)
                        self.file_listbox.insert(tk.END, file_path)
            self.log(f"已扫描到 {len(self.source_files)} 个XML文件")
    
    def add_files(self):
        """
        添加文件到列表
        通过文件对话框选择多个XML文件，自动设置输出目录为第一个文件所在目录的上一级目录
        """
        files = filedialog.askopenfilenames(title="选择XML文件", filetypes=[("XML文件", "*.xml")])
        if files and not self.source_files:
            # 如果是第一次添加文件，设置输出目录为第一个文件所在目录的上一级目录
            first_file_dir = os.path.dirname(files[0])
            self.output_dir = os.path.dirname(first_file_dir)
            self.log(f"输出目录已设置为: {self.output_dir}")
        for file in files:
            if file not in self.source_files:
                self.source_files.append(file)
                self.file_listbox.insert(tk.END, file)
    
    def remove_files(self):
        """
        从列表中移除选中的文件
        """
        selected = self.file_listbox.curselection()
        for idx in reversed(selected):
            self.file_listbox.delete(idx)
            del self.source_files[idx]
    
    def clear_files(self):
        """
        清空文件列表
        """
        self.file_listbox.delete(0, tk.END)
        self.source_files = []
    
    def toggle_lang_display(self):
        """
        切换语言显示模式
        显示所有语言或仅显示默认语言
        """
        self.show_all_langs = self.lang_show_var.get()
        self.update_lang_display()
    
    def update_lang_display(self):
        """
        更新语言复选框显示
        根据当前模式显示相应的语言选项
        """
        # 清除现有组件
        for widget in self.lang_frame.winfo_children():
            widget.destroy()
        
        self.lang_vars = {}
        # 根据模式选择显示的语言
        display_langs = LANG_MAP if self.show_all_langs else {k: v for k, v in LANG_MAP.items() if k in self.default_langs}
        
        # 创建复选框网格
        col = 0
        row_lang = 0
        for lang, info in sorted(display_langs.items()):
            var = tk.BooleanVar()
            chk = ttk.Checkbutton(self.lang_frame, text=f"{lang} - {info['name']}", variable=var)
            chk.grid(row=row_lang, column=col, sticky=tk.W, padx=5, pady=2)
            self.lang_vars[lang] = var
            col += 1
            if col >= 4:
                col = 0
                row_lang += 1
        
        self.lang_frame.update_idletasks()
    
    def select_all_langs(self):
        """
        全选所有语言
        """
        for var in self.lang_vars.values():
            var.set(True)
    
    def deselect_all_langs(self):
        """
        取消全选所有语言
        """
        for var in self.lang_vars.values():
            var.set(False)
    
    def set_default_langs(self):
        """
        设置默认语言列表
        创建独立窗口让用户选择默认语言（最多20个），支持搜索过滤和勾选框
        """
        # 创建顶层窗口
        top = tk.Toplevel(self.root)
        top.title("设置默认语言")
        top.geometry("700x600")
        top.transient(self.root)
        top.grab_set()
        
        # 主框架
        main_frame = ttk.Frame(top, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 说明标签
        ttk.Label(main_frame, text="请勾选最多20种默认语言:", 
                  font=('Arial', 10, 'bold')).pack(anchor=tk.W, pady=5)
        
        # 搜索框
        search_frame = ttk.Frame(main_frame)
        search_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(search_frame, text="搜索（输入国家中文名称）:").pack(side=tk.LEFT, padx=5)
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=search_var, width=50)
        search_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        search_entry.focus_set()
        
        # 创建带滚动条的Canvas
        canvas_frame = ttk.Frame(main_frame)
        canvas_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        canvas = tk.Canvas(canvas_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        
        # 保存所有语言数据和勾选框变量
        all_langs = [(code, info) for code, info in sorted(LANG_MAP.items())]
        lang_check_vars = {}  # 语言代码 -> BooleanVar
        lang_check_frames = {}  # 语言代码 -> Frame
        
        # 状态标签
        status_label = ttk.Label(main_frame, text=f"已选择 {len(self.default_langs)}/20 种语言")
        status_label.pack(anchor=tk.W, pady=5)
        
        # 更新状态标签
        def update_status():
            count = sum(1 for var in lang_check_vars.values() if var.get())
            status_label.config(text=f"已选择 {count}/20 种语言")
        
        # 检查并限制最多20个
        def on_check(lang_code):
            var = lang_check_vars[lang_code]
            if var.get():
                count = sum(1 for v in lang_check_vars.values() if v.get())
                if count > 20:
                    var.set(False)
                    messagebox.showwarning("警告", "最多只能选择20种语言")
            update_status()
        
        # 创建语言勾选框
        for lang_code, lang_info in all_langs:
            var = tk.BooleanVar(value=lang_code in self.default_langs)
            lang_check_vars[lang_code] = var
            
            frame = ttk.Frame(scrollable_frame)
            frame.pack(fill=tk.X, pady=1)
            lang_check_frames[lang_code] = frame
            
            chk = ttk.Checkbutton(frame, text=f"{lang_code} - {lang_info['name']}", 
                                  variable=var, command=lambda lc=lang_code: on_check(lc))
            chk.pack(anchor=tk.W, padx=5)
        
        # 搜索过滤功能
        def on_search(*args):
            filter_text = search_var.get().lower()
            for lang_code, frame in lang_check_frames.items():
                lang_info = LANG_MAP[lang_code]
                if filter_text:
                    if filter_text in lang_info['name'].lower():
                        frame.pack(fill=tk.X, pady=1)
                    else:
                        frame.pack_forget()
                else:
                    frame.pack(fill=tk.X, pady=1)
        
        search_var.trace('w', on_search)
        
        # 操作按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        # 取消全选按钮
        def deselect_all():
            for var in lang_check_vars.values():
                var.set(False)
            update_status()
        
        ttk.Button(button_frame, text="取消全选", command=deselect_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="确定", 
                   command=lambda: self._confirm_default_langs_from_vars(top, lang_check_vars)).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="取消", command=top.destroy).pack(side=tk.RIGHT, padx=5)
        
        # 鼠标滚轮滚动
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", on_mousewheel)
        
        def on_close():
            canvas.unbind_all("<MouseWheel>")
            top.destroy()
        top.protocol("WM_DELETE_WINDOW", on_close)
    
    def _confirm_default_langs_from_vars(self, top, lang_check_vars):
        """
        从勾选框变量确认默认语言设置
        """
        selected_langs = [code for code, var in lang_check_vars.items() if var.get()]
        
        if selected_langs:
            self.default_langs = selected_langs[:20]
            self.save_config()
            self.show_all_langs = False
            self.lang_show_var.set(False)
            self.update_lang_display()
            messagebox.showinfo("提示", f"已设置 {len(self.default_langs)} 种默认语言")
        else:
            messagebox.showwarning("警告", "请至少选择一种语言")
        
        top.destroy()
    
    def load_translation_table(self):
        """
        加载翻译表
        从输出目录读取translation_table.json
        """
        table_path = os.path.join(self.output_dir, 'translation_table.json')
        if os.path.exists(table_path):
            try:
                with open(table_path, 'r', encoding='utf-8') as f:
                    self.translation_table = json.load(f)
                self.log(f"已加载翻译表，共 {len(self.translation_table)} 条记录")
            except Exception as e:
                self.log(f"加载翻译表失败: {e}")
    
    def save_translation_table(self):
        """
        保存翻译表
        将翻译表写入translation_table.json
        """
        table_path = os.path.join(self.output_dir, 'translation_table.json')
        try:
            with open(table_path, 'w', encoding='utf-8') as f:
                json.dump(self.translation_table, f, ensure_ascii=False, indent=2)
            self.log("翻译表已保存")
        except Exception as e:
            self.log(f"保存翻译表失败: {e}")
    
    def log(self, message):
        """
        添加日志消息
        :param message: 日志消息内容
        """
        self.log_text.insert(tk.END, f"{time.strftime('%Y-%m-%d %H:%M:%S')} - {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def extract_chinese_texts(self):
        """
        从XML文件中提取中文文本
        :return: 中文文本集合
        """
        all_texts = set()
        
        for xml_file in self.source_files:
            try:
                with open(xml_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                parser = ET.XMLParser(encoding='utf-8')
                root = ET.fromstring(content, parser=parser)
                
                for elem in root.findall('.//Message'):
                    text = elem.get('Content', '')
                    if text and self.is_chinese(text):
                        all_texts.add(text)
                self.log(f"从 {os.path.basename(xml_file)} 提取到中文文本")
            except Exception as e:
                self.log(f"解析 {xml_file} 失败: {e}")
        
        # 将新文本添加到翻译表
        for text in all_texts:
            if text not in self.translation_table:
                self.translation_table[text] = {lang: '' for lang in LANG_MAP.keys()}
        
        self.save_translation_table()
        return all_texts
    
    def is_chinese(self, text):
        """
        判断文本是否包含中文
        :param text: 待检查文本
        :return: True表示包含中文，False表示不包含
        """
        for char in text:
            if '\u4e00' <= char <= '\u9fff':
                return True
        return False
    
    def on_api_type_changed(self):
        """
        API类型切换处理
        保存用户选择的API类型
        """
        self.api_type = self.api_type_var.get()
        self.save_config()
        if self.api_type == 'baidu':
            self.log(f"当前翻译API: 百度翻译")
            if not self.baidu_app_id or not self.baidu_secret_key:
                self.log("警告: 请先配置百度翻译API密钥！")
                messagebox.showwarning("警告", "请先配置百度翻译API密钥！\n点击'API配置'按钮进行设置。")
        else:
            self.log(f"当前翻译API: Google翻译")
    
    def configure_api(self):
        """
        配置翻译API
        弹出窗口让用户输入百度翻译API的App ID和Secret Key
        """
        top = tk.Toplevel(self.root)
        top.title("API配置")
        top.geometry("500x300")
        top.transient(self.root)
        top.grab_set()
        
        main_frame = ttk.Frame(top, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 当前API类型显示
        ttk.Label(main_frame, text=f"当前使用API: {self.api_type.upper()}", font=('Arial', 11, 'bold')).pack(anchor=tk.W, pady=10)
        
        # 百度API配置
        ttk.Label(main_frame, text="百度翻译API配置:", font=('Arial', 10, 'bold')).pack(anchor=tk.W, pady=5)
        ttk.Label(main_frame, text="App ID:").pack(anchor=tk.W)
        app_id_var = tk.StringVar(value=self.baidu_app_id)
        ttk.Entry(main_frame, textvariable=app_id_var, width=50).pack(anchor=tk.W, pady=5)
        
        ttk.Label(main_frame, text="Secret Key:").pack(anchor=tk.W)
        secret_key_var = tk.StringVar(value=self.baidu_secret_key)
        ttk.Entry(main_frame, textvariable=secret_key_var, width=50, show='*').pack(anchor=tk.W, pady=5)
        
        # 说明
        ttk.Label(main_frame, text="提示: 百度翻译API需要在百度翻译开放平台申请", 
                  font=('Arial', 9), foreground='gray').pack(anchor=tk.W, pady=10)
        
        # 按钮
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(anchor=tk.E, pady=20)
        
        def save_and_close():
            self.baidu_app_id = app_id_var.get().strip()
            self.baidu_secret_key = secret_key_var.get().strip()
            self.save_config()
            messagebox.showinfo("提示", "API配置已保存！")
            top.destroy()
        
        ttk.Button(btn_frame, text="保存", command=save_and_close).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=top.destroy).pack(side=tk.LEFT, padx=5)
    
    def baidu_translate(self, text, target_lang):
        """
        使用百度翻译API翻译文本
        :param text: 待翻译文本
        :param target_lang: 目标语言代码
        :return: 翻译后的文本
        """
        if not self.baidu_app_id or not self.baidu_secret_key:
            self.log("错误: 百度API密钥未配置！")
            return text
        
        # 百度翻译API支持的语言列表
        supported_langs = {
            'auto', 'zh', 'en', 'yue', 'wyw', 'ja', 'ko', 'fr', 'de', 'es', 'pt', 'pt-PT',
            'vi', 'id', 'th', 'ru', 'ar', 'hi', 'bn', 'pa', 'gu', 'or', 'ta', 'te', 'kn',
            'ml', 'mr', 'sa', 'yi', 'hi-IN', 'ne', 'my', 'si', 'km', 'lo', 'bo', 'ug',
            'ca', 'eu', 'gl', 'it', 'el', 'hu', 'cs', 'sk', 'sl', 'pl', 'ro', 'bg',
            'sr', 'hr', 'lv', 'lt', 'et', 'fi', 'sv', 'da', 'no', 'is', 'mt', 'tr',
            'hi-Latn', 'ur', 'fa', 'sq', 'bs', 'ms', 'ms-Arab', 'tl', 'ht', 'sw',
            'xh', 'zu', 'af', 'zu', 'xh', 'st', 'tn', 'ts', 'ss', 'nr', 've', 'wo',
            'rm', 'ga', 'gd', 'cy', 'kw', 'br', 'gv', 'gd', 'ga', 'eu', 'gl', 'ast',
            'ca', 'oc', 'fr-CA', 'it-CH', 'de-CH', 'gsw', 'als', 'lb', 'wa', 'nl',
            'fy', 'pap', 'sv-SE', 'sv-FI', 'nb', 'nn', 'da', 'is', 'fo', 'et', 'lv',
            'lt', 'pl', 'cs', 'sk', 'sl', 'hr', 'sr', 'bg', 'mk', 'sq', 'mt', 'tr',
            'az', 'uz', 'kk', 'ky', 'tg', 'tt', 'ba', 'cv', 'ce', 'yi', 'he', 'ar',
            'fa', 'ur', 'ps', 'sd', 'ku', 'ckb', 'syc', 'dv', 'hi', 'bn', 'as',
            'gu', 'or', 'ta', 'te', 'kn', 'ml', 'pa', 'mr', 'sa', 'kok', 'mai',
            'ne', 'si', 'my', 'km', 'lo', 'th', 'ja', 'ko', 'zh', 'zh-TW', 'zh-HK',
            'zh-CN', 'yue', 'wyw', 'vi', 'id', 'tl', 'ms', 'en', 'de', 'fr', 'es',
            'pt', 'ru', 'it', 'pl', 'ro', 'bg', 'cs', 'da', 'nl', 'fi', 'el', 'hu',
            'no', 'pt-PT', 'sv', 'tr', 'ar', 'hi', 'ja', 'ko', 'th', 'vi', 'zh',
            'zh-TW', 'en', 'de', 'fr', 'es', 'pt', 'ru', 'it', 'pl', 'ro', 'bg',
            'cs', 'da', 'nl', 'fi', 'el', 'hu', 'no', 'pt-PT', 'sv', 'tr', 'ar',
            'hi', 'ja', 'ko', 'th', 'vi', 'zh', 'zh-TW', 'en', 'de', 'fr', 'es',
            'pt', 'ru', 'it', 'pl', 'ro', 'bg', 'cs', 'da', 'nl', 'fi', 'el', 'hu',
            'no', 'pt-PT', 'sv', 'tr'
        }
        
        # 调试日志
        self.log(f"百度翻译: target_lang={target_lang}")
        
        # 百度翻译API语言代码映射表
        baidu_lang_map = {
            # 中文变体
            'zh-CN': 'zh', 'zh-TW': 'cht', 'zh-HK': 'zh', 'zh-SG': 'zh', 'zh-MO': 'zh',
            # 阿拉伯语变体
            'ar-SA': 'ar', 'ar-IQ': 'ar', 'ar-EG': 'ar', 'ar-LY': 'ar', 'ar-DZ': 'ar',
            'ar-MA': 'ar', 'ar-TN': 'ar', 'ar-OM': 'ar', 'ar-YE': 'ar', 'ar-SY': 'ar',
            'ar-JO': 'ar', 'ar-LB': 'ar', 'ar-KW': 'ar', 'ar-AE': 'ar', 'ar-BH': 'ar',
            'ar-QA': 'ar',
            # 德语变体
            'de-DE': 'de', 'de-CH': 'de', 'de-AT': 'de', 'de-LU': 'de', 'de-LI': 'de',
            # 英语变体
            'en-US': 'en', 'en-GB': 'en', 'en-AU': 'en', 'en-CA': 'en', 'en-NZ': 'en',
            'en-IE': 'en', 'en-ZA': 'en',
            # 西班牙语变体
            'es-ES': 'es', 'es-MX': 'es', 'es-GT': 'es', 'es-CR': 'es', 'es-PA': 'es',
            'es-DO': 'es', 'es-VE': 'es', 'es-CO': 'es', 'es-PE': 'es', 'es-AR': 'es',
            'es-EC': 'es', 'es-CL': 'es', 'es-UY': 'es', 'es-PY': 'es', 'es-BO': 'es',
            'es-SV': 'es', 'es-HN': 'es', 'es-NI': 'es', 'es-PR': 'es',
            # 法语变体
            'fr-FR': 'fr', 'fr-BE': 'fr', 'fr-CA': 'fr', 'fr-CH': 'fr', 'fr-LU': 'fr',
            'fr-MC': 'fr',
            # 意大利语变体
            'it-IT': 'it', 'it-CH': 'it',
            # 荷兰语变体
            'nl-NL': 'nl', 'nl-BE': 'nl',
            # 葡萄牙语变体
            'pt-PT': 'pt', 'pt-BR': 'pt',
            # 塞尔维亚语变体
            'sr-Latn': 'sr', 'sr-Cyrl': 'sr',
            # 瑞典语变体
            'sv-SE': 'sv', 'sv-FI': 'sv',
            # 乌兹别克语变体
            'uz-Latn': 'uz',
            # 阿塞拜疆语变体
            'az-Latn': 'az',
            # 其他语言
            'ca': 'ca', 'bg': 'bg', 'cs': 'cs', 'da': 'da', 'el': 'el', 'eo': 'eo',
            'et': 'et', 'fi': 'fi', 'he': 'he', 'hu': 'hu', 'is': 'is', 'ja': 'ja',
            'ko': 'ko', 'lv': 'lv', 'lt': 'lt', 'mk': 'mk', 'no': 'no', 'nn': 'no',
            'nb': 'no', 'pl': 'pl', 'ro': 'ro', 'ru': 'ru', 'hr': 'hr', 'sk': 'sk',
            'sl': 'sl', 'sq': 'sq', 'th': 'th', 'tr': 'tr', 'ur-PK': 'ur', 'id': 'id',
            'uk': 'uk', 'be': 'be', 'fa': 'fa', 'vi': 'vi', 'hy': 'hy', 'eu': 'eu',
            'af': 'af', 'ka': 'ka', 'fo': 'fo', 'hi': 'hi', 'ms-MY': 'ms', 'ms-BN': 'ms',
            'kk': 'kk', 'sw': 'sw', 'tt': 'tt', 'bn': 'bn', 'pa': 'pa', 'gu': 'gu',
            'or': 'or', 'ta': 'ta', 'te': 'te', 'kn': 'kn', 'ml': 'ml', 'as': 'as',
            'mr': 'mr', 'sa': 'sa', 'kok': 'kok'
        }
        
        # 获取百度语言代码
        if target_lang in baidu_lang_map:
            lang_code = baidu_lang_map[target_lang]
            self.log(f"百度翻译: 映射成功 {target_lang} -> {lang_code}")
        else:
            # 尝试仅使用语言代码部分
            base_lang = target_lang.split('-')[0] if '-' in target_lang else target_lang
            if base_lang in baidu_lang_map:
                lang_code = baidu_lang_map[base_lang]
                self.log(f"百度翻译: 使用基础语言代码 {target_lang} -> {base_lang} -> {lang_code}")
            else:
                self.log(f"警告: 百度翻译不支持语言 '{target_lang}'，使用英语")
                lang_code = 'en'
        
        max_retries = 3
        for attempt in range(max_retries):
            try:
                salt = str(uuid.uuid4())[:10]
                sign = self.baidu_app_id + text + salt + self.baidu_secret_key
                sign = hashlib.md5(sign.encode('utf-8')).hexdigest()
                
                url = 'https://fanyi-api.baidu.com/api/trans/vip/translate'
                data = {
                    'q': text[:50] + '...' if len(text) > 50 else text,
                    'from': 'auto',
                    'to': lang_code,
                    'appid': self.baidu_app_id,
                    'salt': salt,
                    'sign': sign
                }
                
                self.log(f"百度翻译请求: from=auto, to={lang_code}, text_len={len(text)}")
                
                response = requests.post(url, data=data, timeout=15)
                result = response.json()
                
                self.log(f"百度翻译响应: {result}")
                
                if 'trans_result' in result and result['trans_result']:
                    return result['trans_result'][0]['dst']
                else:
                    if 'error_code' in result:
                        error_code = result['error_code']
                        error_msg = result.get('error_msg', '未知错误')
                        self.log(f"百度翻译API错误[{error_code}]: {error_msg}")
                        # 如果是语言方向不支持，提示用户检查账户权限
                        if error_code == '58001':
                            self.log(f"提示: 请检查百度翻译API账户是否开通了 {lang_code} 语言的翻译服务")
                    return text
            except Exception as e:
                self.log(f"百度翻译异常: {str(e)}")
                if attempt < max_retries - 1:
                    time.sleep(2)
                    continue
                self.log(f"百度翻译失败: {text} -> {target_lang}")
                return text
        
        return text
    
    def translate_text(self, text, target_lang):
        """
        使用选择的翻译API翻译文本
        :param text: 待翻译文本
        :param target_lang: 目标语言代码
        :return: 翻译后的文本
        """
        text_strip = text.strip()
        if not text_strip:
            return text
        if text_strip in BRAND_NAMES:
            return text
        
        if self.api_type == 'baidu':
            return self.baidu_translate(text, target_lang)
        
        max_retries = 3
        for attempt in range(max_retries):
            try:
                url = f'https://translate.googleapis.com/translate_a/single?client=gtx&sl=zh-CN&tl={target_lang}&dt=t&q={quote(text)}'
                headers = {'User-Agent': 'Mozilla/5.0'}
                response = requests.get(url, headers=headers, timeout=15)
                response.raise_for_status()
                result = response.json()
                translated = ''.join([item[0] for item in result[0] if item[0]])
                return translated
            except Exception as e:
                if attempt < max_retries - 1:
                    time.sleep(3)
                    continue
                self.log(f"翻译失败: {text} -> {target_lang}")
                return text
    
    def start_translation(self):
        """
        开始翻译流程
        1. 检查输入
        2. 提取中文文本
        3. 执行翻译
        4. 更新进度
        """
        # 检查输入
        if not self.source_files:
            messagebox.showwarning("警告", "请先选择XML源文件或文件夹")
            return
        
        self.selected_langs = [lang for lang, var in self.lang_vars.items() if var.get()]
        if not self.selected_langs:
            messagebox.showwarning("警告", "请至少选择一种目标语言")
            return
        
        # 初始化状态
        self.start_btn.config(state='disabled')
        self.confirm_btn.config(state='disabled')
        self.translation_complete = False
        self.progress_var.set(0)
        self.progress_label.config(text="0%")
        self.log("开始提取中文文本...")
        
        # 提取中文文本
        texts = self.extract_chinese_texts()
        self.log(f"共需处理 {len(texts)} 条中文文本")
        
        # 计算需要翻译的数量
        total_needed = 0
        for text, langs in self.translation_table.items():
            for lang in self.selected_langs:
                if lang not in langs:
                    langs[lang] = ''
                if not langs[lang] or langs[lang] == text:
                    total_needed += 1
        
        # 如果没有需要翻译的内容
        if total_needed == 0:
            self.log("所有翻译已完成")
            self.translation_complete = True
            self.confirm_btn.config(state='normal')
            self.start_btn.config(state='normal')
            return
        
        self.log(f"需要翻译: {total_needed} 条")
        
        # 创建翻译线程
        def translate_thread():
            count = 0
            processed = 0
            
            # 只处理需要翻译的文本对
            texts_to_translate = []
            for text in self.translation_table.keys():
                for lang in self.selected_langs:
                    if lang not in self.translation_table[text]:
                        self.translation_table[text][lang] = ''
                    if not self.translation_table[text][lang] or self.translation_table[text][lang] == text:
                        texts_to_translate.append((text, lang))
            
            for text, lang in texts_to_translate:
                translated = self.translate_text(text, LANG_MAP[lang]['code'])
                self.translation_table[text][lang] = translated
                count += 1
                processed += 1
                
                # 每处理5条更新一次进度
                if processed % 5 == 0:
                    self.save_translation_table()
                    progress = round(count / total_needed * 100, 1)
                    # 使用root.after在主线程中更新UI
                    self.root.after(0, lambda p=progress, c=count: self._update_progress(p, c))
                    self.log(f"进度: {progress}% ({count}/{total_needed})")
                
                time.sleep(0.1)
            
            # 完成翻译
            self.save_translation_table()
            self.root.after(0, lambda: self._update_progress(100, total_needed))
            self.log("翻译完成！")
            self.translation_complete = True
            self.root.after(0, lambda: self.confirm_btn.config(state='normal'))
            self.root.after(0, lambda: self.start_btn.config(state='normal'))
        
        # 启动翻译线程
        threading.Thread(target=translate_thread, daemon=True).start()
    
    def _update_progress(self, progress, count):
        """
        在主线程中更新进度条
        """
        self.progress_var.set(progress)
        self.progress_label.config(text=f"{progress}%")
    
    def confirm_and_generate(self):
        """
        确认翻译并生成XML文件
        需要用户确认后才执行生成操作
        """
        if not self.translation_complete:
            messagebox.showwarning("警告", "请先完成翻译")
            return
        
        result = messagebox.askyesno("确认生成", "翻译已完成，是否确认生成XML文件？")
        if result:
            self.generate_xml_files()
    
    def generate_xml_files(self):
        """
        生成各语言版本的XML文件
        将翻译结果写入对应语言文件夹
        """
        if not self.source_files:
            messagebox.showwarning("警告", "请先选择XML源文件")
            return
        
        if not self.selected_langs:
            messagebox.showwarning("警告", "请至少选择一种目标语言")
            return
        
        self.log("开始生成XML文件...")
        
        for xml_file in self.source_files:
            try:
                with open(xml_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                parser = ET.XMLParser(encoding='utf-8')
                root = ET.fromstring(content, parser=parser)
                
                file_name = os.path.basename(xml_file)
                
                # 为每种选中的语言生成XML文件
                for lang in self.selected_langs:
                    output_path = os.path.join(self.output_dir, lang, 'String', file_name)
                    os.makedirs(os.path.dirname(output_path), exist_ok=True)
                    
                    lines = []
                    lines.append('<?xml version="1.0" encoding="utf-8"?>')
                    lines.append('<ResMap>')
                    
                    for elem in root.findall('.//Message'):
                        msg_id = elem.get('ID', '')
                        original_content = elem.get('Content', '')
                        
                        # 获取翻译内容
                        if original_content in self.translation_table and lang in self.translation_table[original_content]:
                            translated = self.translation_table[original_content][lang]
                            if translated and translated != original_content:
                                content = translated
                            else:
                                content = original_content
                        else:
                            content = original_content
                        
                        # 处理换行符
                        content = content.replace('\n', '&#xA;')
                        lines.append(f'\t<Message ID="{msg_id}" Content="{content}" />')
                    
                    lines.append('</ResMap>')
                    
                    # 写入文件
                    with open(output_path, 'w', encoding='utf-8') as f:
                        f.write('\n'.join(lines))
                    
                    self.log(f"生成: {output_path}")
            
            except Exception as e:
                self.log(f"生成XML失败: {e}")
        
        self.log("所有XML文件已生成完成！")
        messagebox.showinfo("完成", "XML文件已成功生成！")
    
    def view_translation_table(self):
        """
        打开翻译表文件
        """
        table_path = os.path.join(self.output_dir, 'translation_table.json')
        if os.path.exists(table_path):
            os.startfile(table_path)
        else:
            messagebox.showinfo("提示", "翻译表文件不存在")
    
    def export_log(self):
        """
        导出日志文件
        """
        log_content = self.log_text.get('1.0', tk.END)
        log_path = os.path.join(self.output_dir, 'translation_log.txt')
        with open(log_path, 'w', encoding='utf-8') as f:
            f.write(log_content)
        messagebox.showinfo("提示", f"日志已导出到: {log_path}")


if __name__ == '__main__':
    """
    程序入口
    创建主窗口并启动应用
    """
    root = tk.Tk()
    app = TranslationApp(root)
    root.mainloop()