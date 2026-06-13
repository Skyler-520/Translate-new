#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
快速设定logo和快捷方式工具
功能：
1. 将图片转换为128x128的ICO格式和75x94的GIF格式
2. 创建桌面快捷方式并设置自定义图标
"""

import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
import subprocess
import shutil
import ctypes
import platform

def is_admin():
    """检查当前是否以管理员权限运行"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    """以管理员权限重新运行程序"""
    if platform.system() != "Windows":
        return False
    
    script = os.path.abspath(sys.argv[0])
    params = ' '.join([f'"{arg}"' for arg in sys.argv[1:]])
    
    try:
        # 使用ShellExecute以管理员权限运行
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, f'"{script}" {params}', None, 1)
        return True
    except Exception as e:
        print(f"提权失败: {e}")
        return False

def check_and_request_admin():
    """检查并请求管理员权限"""
    if not is_admin():
        print("检测到需要管理员权限，正在请求提权...")
        if run_as_admin():
            print("程序将以管理员权限重新启动...")
            sys.exit(0)
        else:
            print("提权失败，程序将继续以普通权限运行")
            return False
    return True

class ImageConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("快速设定logo和快捷方式工具")
        self.root.geometry("600x450")  # 适当尺寸的窗口
        self.root.resizable(True, True)  # 窗口可以灵活拉伸
        
        # 设置苹果风格配色
        self.set_apple_style()
        
        # 设置图标（如果有的话）
        try:
            self.root.iconbitmap("icon.ico")
        except:
            pass
        
        # 初始化变量
        self.image_path = None
        self.preview_image = None
        self.ico_filename = tk.StringVar(value="LOGO")  # 默认ico文件名为LOGO
        
        # 创建界面
        self.create_widgets()
        
        # 检查依赖
        self.check_dependencies()
    
    def set_apple_style(self):
        """设置苹果风格样式"""
        # 苹果风格配色
        self.bg_color = "#f5f5f7"  # 浅灰色背景
        self.card_color = "#ffffff"  # 白色卡片
        self.accent_color = "#007AFF"  # 苹果蓝
        self.text_color = "#1d1d1f"  # 深灰色文字
        self.button_text_color = "#87CEEB"  # 天蓝色按钮文字
        
        # 配置样式
        style = ttk.Style()
        
        # 使用标准样式名称
        style.configure("TFrame", background=self.bg_color)
        style.configure("TLabel", background=self.bg_color, foreground=self.text_color)
        style.configure("TButton", background=self.accent_color, foreground=self.button_text_color)
        style.configure("TLabelframe", background=self.bg_color, foreground=self.text_color)
        style.configure("TLabelframe.Label", background=self.bg_color, foreground=self.text_color)
        style.configure("TProgressbar", background=self.accent_color)
        
        # 设置窗口背景色
        self.root.configure(bg=self.bg_color)
    
    def check_dependencies(self):
        """检查必要的依赖库"""
        try:
            from PIL import Image
        except ImportError:
            messagebox.showerror("错误", "请先安装Pillow库：pip install Pillow")
            sys.exit(1)
    
    def create_widgets(self):
        """创建界面组件"""
        # 创建主滚动框架
        main_canvas = tk.Canvas(self.root, bg=self.bg_color, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=main_canvas.yview)
        scrollable_frame = ttk.Frame(main_canvas, style="TFrame")
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )
        
        main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar.set)
        
        main_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 标题区域（紧凑设计）
        title_frame = ttk.Frame(scrollable_frame, style="TFrame")
        title_frame.pack(fill="x", padx=20, pady=(10, 8))
        
        title_label = ttk.Label(title_frame, 
                               text="快速设定logo和快捷方式工具",
                               font=("SF Pro Display", 16, "bold"),
                               foreground=self.text_color)
        title_label.pack()
        
        subtitle_label = ttk.Label(title_frame,
                                  text="轻松转换图片格式并创建桌面快捷方式",
                                  font=("SF Pro Text", 10),
                                  foreground="#6e6e73")
        subtitle_label.pack(pady=(2, 0))
        
        # 选择图片区域（苹果风格卡片）
        select_card = ttk.LabelFrame(scrollable_frame, 
                                   text=" 1. 选择图片 ",
                                   padding=12)
        select_card.pack(fill="x", padx=20, pady=(0, 8))
        
        select_content = ttk.Frame(select_card)
        select_content.pack(fill="x")
        
        # 文件选择按钮和标签
        select_row = ttk.Frame(select_content)
        select_row.pack(fill="x", pady=(0, 10))
        
        self.select_button = ttk.Button(select_row, 
                                       text="选择图片文件",
                                       command=self.select_image)
        self.select_button.pack(side="left")
        
        self.file_label = ttk.Label(select_row, 
                                   text="未选择文件",
                                   font=("SF Pro Text", 11),
                                   foreground="#6e6e73")
        self.file_label.pack(side="left", padx=(12, 0))
        
        # 转换设置区域
        settings_card = ttk.LabelFrame(scrollable_frame,
                                     text=" 2. 转换设置 ",
                                     padding=12)
        settings_card.pack(fill="x", padx=20, pady=(0, 8))
        
        settings_row = ttk.Frame(settings_card)
        settings_row.pack(fill="x")
        
        # 将公司名称输入框放在左侧，与按钮对齐
        ttk.Label(settings_row, 
                 text="公司名称:",
                 font=("SF Pro Text", 10),
                 foreground=self.text_color).pack(anchor="w", pady=(0, 2))
        
        ico_entry = ttk.Entry(settings_row, 
                            textvariable=self.ico_filename,
                            width=15,
                            font=("SF Pro Text", 10))
        ico_entry.pack(anchor="w")
        
        # 操作按钮区域
        buttons_card = ttk.LabelFrame(scrollable_frame,
                                     text=" 3. 执行操作 ",
                                     padding=12)
        buttons_card.pack(fill="x", padx=20, pady=(0, 8))
        
        buttons_row = ttk.Frame(buttons_card)
        buttons_row.pack(fill="x")
        
        # 创建合并按钮，使用更大的字体和尺寸
        self.modify_logo_button = ttk.Button(buttons_row,
                                        text="修改LOGO并创建快捷方式",
                                        command=self.modify_logo_and_create_shortcut,
                                        state=tk.DISABLED,
                                        width=25)  # 增加按钮宽度
        self.modify_logo_button.pack(side="left", padx=(0, 0))
        
        # 设置按钮字体大小
        style = ttk.Style()
        style.configure("TButton", font=("Microsoft YaHei", 11, "bold"))  # 增大字体并加粗
        
        # 状态显示区域
        status_card = ttk.LabelFrame(scrollable_frame,
                                   text=" 状态 ",
                                   padding=12)
        status_card.pack(fill="x", padx=20, pady=(0, 15))
        
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        
        status_label = ttk.Label(status_card,
                              textvariable=self.status_var,
                              font=("SF Pro Text", 11),
                              foreground=self.accent_color)
        status_label.pack(anchor="w")
        
        # 进度条
        self.progress = ttk.Progressbar(status_card, 
                                       mode='indeterminate')
        self.progress.pack(fill="x", pady=(8, 0))
    
    def select_image(self):
        """选择图片文件"""
        file_types = [
            ("图片文件", "*.png *.jpg *.jpeg *.bmp *.gif *.tiff *.ico"),
            ("所有文件", "*.*")
        ]
        
        filename = filedialog.askopenfilename(
            title="选择图片文件",
            filetypes=file_types
        )
        
        if filename:
            self.image_path = filename
            self.file_label.config(text=os.path.basename(filename))
            self.modify_logo_button.config(state=tk.NORMAL)
            self.status_var.set(f"已选择: {os.path.basename(filename)}")
    

    
    def convert_images(self):
        """转换图片格式"""
        if not self.image_path:
            messagebox.showwarning("警告", "请先选择图片文件")
            return
        
        try:
            self.progress.start()
            self.status_var.set("正在转换图片...")
            self.root.update()
            
            # 创建目标目录
            target_dir = os.path.join("DiskC", "OpenCNC", "Bin", "Logo")
            os.makedirs(target_dir, exist_ok=True)
            
            # 打开原始图片
            original_image = Image.open(self.image_path)
            
            # 转换为ICO格式 (128x128)，使用用户输入的文件名
            ico_filename = self.ico_filename.get().strip() or "LOGO"
            ico_path = os.path.join(target_dir, f"{ico_filename}.ico")
            
            # 创建128x128尺寸的图标
            ico_image = original_image.resize((128, 128), Image.Resampling.LANCZOS)
            
            # 如果图片有透明度，需要处理透明度
            if ico_image.mode in ('RGBA', 'LA'):
                # 转换为RGB模式，白色背景
                background = Image.new('RGB', ico_image.size, (255, 255, 255))
                background.paste(ico_image, mask=ico_image.split()[-1])
                ico_image = background
            elif ico_image.mode != 'RGB':
                ico_image = ico_image.convert('RGB')
            
            # 保存ICO文件，只包含128x128尺寸
            ico_image.save(ico_path, format='ICO')
            
            # 转换为GIF格式 (75x94)
            gif_image = original_image.resize((75, 94), Image.Resampling.LANCZOS)
            gif_path = os.path.join(target_dir, "LoadingImage.gif")
            
            # 如果图片有透明度，转换为RGBA模式
            if gif_image.mode in ('RGBA', 'LA'):
                background = Image.new('RGB', gif_image.size, (255, 255, 255))
                background.paste(gif_image, mask=gif_image.split()[-1])
                gif_image = background
            
            gif_image.save(gif_path, format="GIF")
            
            # 验证生成的文件
            ico_valid = self.validate_image(ico_path)
            gif_valid = self.validate_image(gif_path)
            
            self.progress.stop()
            
            if ico_valid and gif_valid:
                messagebox.showinfo("成功", f"图片转换完成！\\n文件已保存到: {target_dir}")
                self.status_var.set("转换完成")
            else:
                messagebox.showwarning("警告", "文件已生成，但验证时发现可能存在问题")
                self.status_var.set("转换完成（有警告）")
                
        except Exception as e:
            self.progress.stop()
            messagebox.showerror("错误", f"转换失败: {str(e)}")
            self.status_var.set("转换失败")
    
    def validate_image(self, image_path):
        """验证图片文件是否可正常打开"""
        try:
            test_image = Image.open(image_path)
            test_image.verify()  # 验证文件完整性
            test_image.close()
            return True
        except Exception:
            return False
    
    def clear_icon_cache(self):
        """清理Windows图标缓存并刷新图标显示"""
        try:
            # 获取用户目录
            user_profile = os.path.expandvars("%USERPROFILE%")
            
            # Windows图标缓存文件路径
            cache_paths = [
                os.path.join(user_profile, "AppData", "Local", "IconCache.db"),
                os.path.join(user_profile, "AppData", "Local", "Microsoft", "Windows", "Explorer"),
            ]
            
            cache_cleared = False
            
            # 清理图标缓存文件
            for cache_path in cache_paths:
                if os.path.exists(cache_path):
                    if os.path.isfile(cache_path):
                        # 删除单个文件
                        try:
                            os.remove(cache_path)
                            print(f"已删除图标缓存文件: {cache_path}")
                            cache_cleared = True
                        except Exception as e:
                            print(f"无法删除图标缓存文件 {cache_path}: {e}")
                    elif os.path.isdir(cache_path):
                        # 删除目录中的图标缓存文件
                        try:
                            import glob
                            icon_cache_files = glob.glob(os.path.join(cache_path, "iconcache*.db"))
                            for cache_file in icon_cache_files:
                                try:
                                    os.remove(cache_file)
                                    print(f"已删除图标缓存文件: {cache_file}")
                                    cache_cleared = True
                                except Exception as e:
                                    print(f"无法删除图标缓存文件 {cache_file}: {e}")
                        except Exception as e:
                            print(f"扫描图标缓存目录失败: {e}")
            
            # 刷新图标缓存而不重启资源管理器
            if cache_cleared:
                try:
                    # 方法1: 使用Windows API刷新图标缓存
                    import ctypes
                    from ctypes import wintypes
                    
                    # 定义Windows API函数
                    SHCNE_ASSOCCHANGED = 0x08000000
                    SHCNF_IDLIST = 0x0000
                    
                    # 调用SHChangeNotify刷新图标缓存
                    ctypes.windll.shell32.SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, None, None)
                    
                    print("已刷新图标缓存（无需重启资源管理器）")
                    return True
                    
                except Exception as e:
                    print(f"刷新图标缓存失败: {e}")
                    # 即使刷新失败，缓存清理可能已经生效
                    return True
            else:
                print("未找到需要清理的图标缓存文件")
                return True
                
        except Exception as e:
            print(f"清理图标缓存时出错: {e}")
            return False

    def modify_logo_and_create_shortcut(self):
        """修改LOGO并创建快捷方式"""
        if not self.image_path:
            messagebox.showwarning("警告", "请先选择图片文件")
            return
        
        try:
            # 第一步：转换图片格式
            self.progress.start()
            self.status_var.set("正在转换图片格式...")
            self.root.update()
            
            # 创建目标目录
            target_dir = os.path.join("DiskC", "OpenCNC", "Bin", "Logo")
            os.makedirs(target_dir, exist_ok=True)
            
            # 打开原始图片
            original_image = Image.open(self.image_path)
            
            # 转换为ICO格式 (128x128)，使用用户输入的文件名
            ico_filename = self.ico_filename.get().strip() or "LOGO"
            ico_path = os.path.join(target_dir, f"{ico_filename}.ico")
            
            # 创建128x128尺寸的图标
            ico_image = original_image.resize((128, 128), Image.Resampling.LANCZOS)
            
            # 如果图片有透明度，需要处理透明度
            if ico_image.mode in ('RGBA', 'LA'):
                # 转换为RGB模式，白色背景
                background = Image.new('RGB', ico_image.size, (255, 255, 255))
                background.paste(ico_image, mask=ico_image.split()[-1])
                ico_image = background
            elif ico_image.mode != 'RGB':
                ico_image = ico_image.convert('RGB')
            
            # 保存ICO文件，只包含128x128尺寸
            ico_image.save(ico_path, format='ICO')
            
            # 转换为GIF格式 (75x94)
            gif_image = original_image.resize((75, 94), Image.Resampling.LANCZOS)
            gif_path = os.path.join(target_dir, "LoadingImage.gif")
            
            # 如果图片有透明度，转换为RGBA模式
            if gif_image.mode in ('RGBA', 'LA'):
                background = Image.new('RGB', gif_image.size, (255, 255, 255))
                background.paste(gif_image, mask=gif_image.split()[-1])
                gif_image = background
            
            gif_image.save(gif_path, format="GIF")
            
            # 验证生成的文件
            ico_valid = self.validate_image(ico_path)
            gif_valid = self.validate_image(gif_path)
            
            if not (ico_valid and gif_valid):
                messagebox.showwarning("警告", "图片转换完成，但验证时发现可能存在问题")
            
            # 第二步：创建快捷方式
            self.status_var.set("正在创建桌面快捷方式...")
            self.root.update()
            
            # 检查必要的文件是否存在
            exe_path = os.path.join("DiskC", "OpenCNC", "Bin", "SyntecLaserMarking.exe")
            
            if not os.path.exists(exe_path):
                messagebox.showerror("错误", f"可执行文件不存在: {exe_path}")
                self.progress.stop()
                return
            
            # 清理图标缓存
            self.status_var.set("正在刷新图标缓存...")
            self.root.update()
            self.clear_icon_cache()
            
            # 执行VBScript创建快捷方式
            vbs_script = "create_shortcut.vbs"
            
            if os.path.exists(vbs_script):
                # 将用户输入的图标文件名和公司名称作为参数传递给VBScript
                company_name = self.ico_filename.get().strip() or ""
                result = subprocess.run(["cscript", "//Nologo", vbs_script, ico_filename, company_name], 
                                      capture_output=True, text=True)
                
                self.progress.stop()
                
                if result.returncode == 0:
                    messagebox.showinfo("成功", "LOGO修改完成！桌面快捷方式创建成功！图标缓存已刷新，新图标将立即生效。")
                    self.status_var.set("LOGO修改和快捷方式创建成功")
                else:
                    messagebox.showerror("错误", f"创建快捷方式失败: {result.stderr}")
                    self.status_var.set("快捷方式创建失败")
            else:
                self.progress.stop()
                messagebox.showerror("错误", "VBScript脚本文件不存在")
                
        except Exception as e:
            self.progress.stop()
            messagebox.showerror("错误", f"操作失败: {str(e)}")
            self.status_var.set("操作失败")

    def create_shortcut(self):
        """创建桌面快捷方式"""
        try:
            # 检查必要的文件是否存在
            exe_path = os.path.join("DiskC", "OpenCNC", "Bin", "SyntecLaserMarking.exe")
            # 使用用户输入的文件名
            ico_filename = self.ico_filename.get().strip() or "LOGO"
            ico_path = os.path.join("DiskC", "OpenCNC", "Bin", "Logo", f"{ico_filename}.ico")
            
            if not os.path.exists(exe_path):
                messagebox.showerror("错误", f"可执行文件不存在: {exe_path}")
                return
            
            if not os.path.exists(ico_path):
                messagebox.showwarning("警告", "请先转换图片生成logo.ico文件")
                return
            
            # 清理图标缓存
            self.status_var.set("正在清理图标缓存...")
            self.root.update()
            self.clear_icon_cache()
            
            # 执行VBScript创建快捷方式
            vbs_script = "create_shortcut.vbs"
            
            if os.path.exists(vbs_script):
                # 将用户输入的图标文件名和公司名称作为参数传递给VBScript
                company_name = self.ico_filename.get().strip() or ""
                result = subprocess.run(["cscript", "//Nologo", vbs_script, ico_filename, company_name], 
                                      capture_output=True, text=True)
                
                if result.returncode == 0:
                    messagebox.showinfo("成功", "桌面快捷方式创建成功！图标缓存已刷新，新图标将立即生效。")
                    self.status_var.set("快捷方式创建成功")
                else:
                    messagebox.showerror("错误", f"创建快捷方式失败: {result.stderr}")
                    self.status_var.set("快捷方式创建失败")
            else:
                messagebox.showerror("错误", "VBScript脚本文件不存在")
                
        except Exception as e:
            messagebox.showerror("错误", f"创建快捷方式时出错: {str(e)}")
            self.status_var.set("快捷方式创建失败")

def main():
    """主函数"""
    # 检查并请求管理员权限
    check_and_request_admin()
    
    root = tk.Tk()
    app = ImageConverterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()