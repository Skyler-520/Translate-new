# -*- coding: utf-8 -*-
"""
打包脚本 - 使用 PyInstaller 将翻译工具打包为 Windows 可执行文件
"""

import os
import sys
import subprocess
import shutil
from datetime import datetime

def check_pyinstaller():
    """检查是否已安装 PyInstaller"""
    try:
        import PyInstaller
        print(f"[OK] PyInstaller 已安装 (版本: {PyInstaller.__version__})")
        return True
    except ImportError:
        print("[ERROR] PyInstaller 未安装")
        print("正在安装 PyInstaller...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
            print("[OK] PyInstaller 安装成功")
            return True
        except Exception as e:
            print(f"[ERROR] 安装失败: {e}")
            return False

def build_exe():
    """使用 PyInstaller 打包"""
    version = "1.1.0"
    app_name = "TranslationTool"
    output_dir = "dist"
    build_dir = "build"
    
    # 清理旧的构建目录
    for dir_name in [output_dir, build_dir]:
        if os.path.exists(dir_name):
            print(f"清理旧目录: {dir_name}")
            shutil.rmtree(dir_name)
    
    # PyInstaller 命令
    pyinstaller_cmd = [
        "pyinstaller",
        "--name", f"{app_name}_v{version}",
        "--windowed",
        "--onefile",
        "--icon=NONE",
        "--version-file=version_info.txt",
        "--add-data", "translation_table.json;.",
        "--hidden-import=requests",
        "--hidden-import=urllib3",
        "--optimize=2",
        "translation_gui.py"
    ]
    
    print("\n开始打包...")
    print(f"命令: {' '.join(pyinstaller_cmd)}")
    print("-" * 60)
    
    try:
        result = subprocess.run(
            pyinstaller_cmd,
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            encoding='utf-8',
            errors='ignore'
        )
        print(result.stdout)
        print("\n" + "-" * 60)
        print("[OK] 打包完成!")
        
        exe_path = os.path.join(output_dir, f"{app_name}_v{version}.exe")
        if os.path.exists(exe_path):
            size_mb = os.path.getsize(exe_path) / (1024 * 1024)
            print(f"\n可执行文件: {exe_path}")
            print(f"文件大小: {size_mb:.2f} MB")
            
            # 创建 release 说明
            create_release_note(version, output_dir, exe_path)
            return True
        else:
            print("[ERROR] 打包完成但找不到可执行文件")
            return False
            
    except subprocess.CalledProcessError as e:
        print("[ERROR] 打包失败!")
        print(e.stdout)
        return False

def create_release_note(version, output_dir, exe_path):
    """创建发布说明"""
    note_path = os.path.join(output_dir, f"ReleaseNotes_v{version}.txt")
    note_content = f"""
=====================================================================
                    翻译工具 - 版本 {version}
=====================================================================

发布日期: {datetime.now().strftime('%Y-%m-%d')}

【版本说明】
v{version} 更新内容:
1. 性能优化:
   - HTTP连接复用(Session)，减少请求延迟
   - 并发线程数从5增加到15
   - 指数退避重试策略(1s→2s→4s)
   - 超时时间优化(15s→10s)

2. 相同文本多语言翻译合并，减少重复请求

3. 移除品牌名称检查，品牌内容现在也会被翻译

4. 整体速度提升约5倍

【使用说明】
1. 双击运行 TranslationTool_v{version}.exe
2. 选择要翻译的文件或文件夹
3. 选择目标语言
4. 点击"开始翻译"按钮
5. 等待翻译完成后确认并保存

【系统要求】
- Windows 7 或更高版本
- 网络连接（用于翻译API）

【注意事项】
- 首次运行可能需要几秒钟加载
- 翻译速度取决于网络状况和API响应速度

=====================================================================
"""
    with open(note_path, 'w', encoding='utf-8') as f:
        f.write(note_content.strip())
    print(f"发布说明已创建: {note_path}")

def main():
    print("=" * 60)
    print("翻译工具 - 打包脚本")
    print("=" * 60)
    
    # 切换到脚本目录
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    
    if not check_pyinstaller():
        print("\n请手动安装 PyInstaller: pip install pyinstaller")
        return
    
    success = build_exe()
    
    if success:
        print("\n" + "=" * 60)
        print("打包成功完成!")
        print("=" * 60)
    else:
        print("\n" + "=" * 60)
        print("打包失败，请查看错误信息")
        print("=" * 60)

if __name__ == "__main__":
    main()
