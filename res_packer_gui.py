import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import gzip
import threading
import time


class ResPackerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("XML文件打包工具")
        self.root.geometry("800x600")
        
        self.selected_files = []
        self.progress_var = tk.DoubleVar()
        self.is_processing = False
        
        self.create_widgets()
    
    def create_widgets(self):
        # 输入文件选择区域
        input_frame = ttk.LabelFrame(self.root, text="输入文件")
        input_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 文件列表
        self.file_listbox = tk.Listbox(input_frame, selectmode=tk.MULTIPLE, height=10)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        scrollbar = ttk.Scrollbar(input_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox.config(yscrollcommand=scrollbar.set)
        
        # 按钮区域
        button_frame = ttk.Frame(self.root)
        button_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Button(button_frame, text="添加文件", command=self.add_files).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="移除文件", command=self.remove_files).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="清空列表", command=self.clear_list).pack(side=tk.LEFT, padx=5)
        
        # 进度条区域
        progress_frame = ttk.LabelFrame(self.root, text="压缩进度")
        progress_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, padx=5, pady=5)
        
        self.progress_label = ttk.Label(progress_frame, text="准备就绪")
        self.progress_label.pack(pady=2)
        
        # 日志区域
        log_frame = ttk.LabelFrame(self.root, text="处理日志")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=10)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 操作按钮
        action_frame = ttk.Frame(self.root)
        action_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.pack_button = ttk.Button(action_frame, text="开始打包", command=self.start_packing)
        self.pack_button.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(action_frame, text="导出日志", command=self.export_log).pack(side=tk.LEFT, padx=5)
    
    def add_files(self):
        files = filedialog.askopenfilenames(
            title="选择XML文件（不带后缀）",
            filetypes=[("所有文件", "*.*")]
        )
        if files:
            for file_path in files:
                if file_path not in self.selected_files:
                    self.selected_files.append(file_path)
                    self.file_listbox.insert(tk.END, os.path.basename(file_path))
            self.log(f"已添加 {len(files)} 个文件")
    
    def remove_files(self):
        selected_indices = self.file_listbox.curselection()
        if selected_indices:
            # 从后往前删除，避免索引错乱
            for index in reversed(selected_indices):
                del self.selected_files[index]
                self.file_listbox.delete(index)
            self.log(f"已移除 {len(selected_indices)} 个文件")
    
    def clear_list(self):
        self.selected_files.clear()
        self.file_listbox.delete(0, tk.END)
        self.log("已清空文件列表")
    
    def log(self, message):
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.insert(tk.END, f"{timestamp} - {message}\n")
        self.log_text.see(tk.END)
    
    def start_packing(self):
        if not self.selected_files:
            messagebox.showwarning("警告", "请先添加要打包的文件")
            return
        
        if self.is_processing:
            messagebox.showwarning("警告", "正在处理中，请稍候")
            return
        
        self.is_processing = True
        self.pack_button.config(state=tk.DISABLED)
        self.progress_var.set(0)
        
        # 在后台线程中执行打包
        threading.Thread(target=self.pack_files, daemon=True).start()
    
    def pack_files(self):
        total_files = len(self.selected_files)
        success_count = 0
        fail_count = 0
        
        for i, file_path in enumerate(self.selected_files):
            try:
                # 检查文件是否存在
                if not os.path.exists(file_path):
                    self.log(f"错误: 文件不存在 - {file_path}")
                    fail_count += 1
                    continue
                
                # 检查文件权限
                if not os.access(file_path, os.R_OK):
                    self.log(f"错误: 无法读取文件 - {file_path}")
                    fail_count += 1
                    continue
                
                # 获取原始文件名
                original_name = os.path.basename(file_path)
                
                # 创建输出路径（同目录下添加.res后缀）
                output_path = file_path + ".res"
                
                # 使用GZIP格式创建压缩文件（与原始.res文件格式一致）
                with open(file_path, 'rb') as input_file:
                    content = input_file.read()
                
                # 使用自定义GZIP打包，确保FLG=0x00
                import zlib
                crc = zlib.crc32(content) & 0xffffffff
                compressed = zlib.compress(content, 9)
                
                # 构造GZIP文件头（确保FLG=0x00）
                gzip_header = b'\x1f\x8b\x08\x00'  # ID1, ID2, CM, FLG
                gzip_header += b'\x00\x00\x00\x00'  # MTIME (0)
                gzip_header += b'\x00'              # XFL (0)
                gzip_header += b'\xff'              # OS (255 = unknown)
                
                # 写入文件
                with open(output_path, 'wb') as f_out:
                    f_out.write(gzip_header)
                    f_out.write(compressed[2:-4])  # 移除zlib头和adler32
                    # 添加GZIP尾部（CRC32和原始大小）
                    f_out.write(crc.to_bytes(4, 'little'))
                    f_out.write(len(content).to_bytes(4, 'little'))
                
                success_count += 1
                self.log(f"成功: {original_name} -> {os.path.basename(output_path)}")
                
            except Exception as e:
                self.log(f"错误: {os.path.basename(file_path)} - {str(e)}")
                fail_count += 1
            
            # 更新进度
            progress = ((i + 1) / total_files) * 100
            self.progress_var.set(progress)
            self.progress_label.config(text=f"正在处理: {i + 1}/{total_files}")
        
        # 完成处理
        self.progress_var.set(100)
        self.progress_label.config(text=f"处理完成: 成功 {success_count} 个, 失败 {fail_count} 个")
        self.is_processing = False
        self.pack_button.config(state=tk.NORMAL)
        
        if fail_count == 0:
            messagebox.showinfo("完成", f"所有文件打包成功！\n共生成 {success_count} 个 .res 文件")
        else:
            messagebox.showwarning("完成", f"打包完成，但有 {fail_count} 个文件失败\n请查看日志了解详情")
    
    def export_log(self):
        log_content = self.log_text.get('1.0', tk.END)
        log_path = os.path.join(os.getcwd(), 'packer_log.txt')
        with open(log_path, 'w', encoding='utf-8') as f:
            f.write(log_content)
        messagebox.showinfo("提示", f"日志已导出到: {log_path}")


if __name__ == '__main__':
    root = tk.Tk()
    app = ResPackerApp(root)
    root.mainloop()