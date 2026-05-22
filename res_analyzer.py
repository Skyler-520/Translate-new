import os
import zipfile
import struct
import hashlib
import json
import sys


class ResFileAnalyzer:
    def __init__(self):
        self.results = {
            'original': {},
            'repacked': {},
            'comparison': {}
        }
    
    def analyze_file(self, file_path, label):
        """分析单个.res文件"""
        result = {
            'file_path': file_path,
            'file_size': os.path.getsize(file_path),
            'is_zip': False,
            'zip_info': None,
            'inner_files': [],
            'magic_bytes': None,
            'sha256': None,
            'compression_method': None,
            'metadata': {}
        }
        
        # 计算文件哈希
        with open(file_path, 'rb') as f:
            file_data = f.read()
            result['sha256'] = hashlib.sha256(file_data).hexdigest()
        
        # 检查魔数字节
        if len(file_data) >= 4:
            result['magic_bytes'] = file_data[:4].hex().upper()
        
        # 检查是否为ZIP格式
        if result['magic_bytes'] == '504B0304':
            result['is_zip'] = True
            try:
                with zipfile.ZipFile(file_path, 'r') as zf:
                    result['zip_info'] = {
                        'comment': zf.comment.decode('utf-8', errors='replace') if zf.comment else '',
                        'is_encrypted': zf.fp.read(8)[4] & 0x01,
                        'number_of_files': len(zf.infolist())
                    }
                    
                    for info in zf.infolist():
                        inner_file = {
                            'filename': info.filename,
                            'compressed_size': info.compress_size,
                            'uncompressed_size': info.file_size,
                            'compression_method': info.compression,
                            'timestamp': info.date_time,
                            'crc32': hex(info.CRC),
                            'is_directory': info.is_dir(),
                            'external_attr': info.external_attr,
                            'internal_attr': info.internal_attr
                        }
                        
                        # 获取压缩方法名称
                        if info.compression == zipfile.ZIP_STORED:
                            inner_file['compression_name'] = 'STORED (无压缩)'
                        elif info.compression == zipfile.ZIP_DEFLATED:
                            inner_file['compression_name'] = 'DEFLATED (Deflate压缩)'
                        elif info.compression == zipfile.ZIP_BZIP2:
                            inner_file['compression_name'] = 'BZIP2'
                        elif info.compression == zipfile.ZIP_LZMA:
                            inner_file['compression_name'] = 'LZMA'
                        else:
                            inner_file['compression_name'] = f'未知 ({info.compression})'
                        
                        # 读取内部文件内容
                        if not info.is_dir():
                            try:
                                with zf.open(info) as inner_f:
                                    inner_content = inner_f.read()
                                    inner_file['content_sha256'] = hashlib.sha256(inner_content).hexdigest()
                                    inner_file['content_size'] = len(inner_content)
                                    
                                    # 检测编码
                                    try:
                                        inner_file['encoding'] = 'UTF-8'
                                        inner_content.decode('utf-8')
                                    except:
                                        try:
                                            inner_file['encoding'] = 'GBK'
                                            inner_content.decode('gbk')
                                        except:
                                            inner_file['encoding'] = '未知'
                            except Exception as e:
                                inner_file['read_error'] = str(e)
                        
                        result['inner_files'].append(inner_file)
                        
                        # 记录压缩方法
                        result['compression_method'] = inner_file['compression_name']
            except Exception as e:
                result['zip_error'] = str(e)
        else:
            # 非ZIP格式，尝试其他分析
            result['non_zip_analysis'] = self.analyze_non_zip(file_data)
        
        self.results[label] = result
    
    def analyze_non_zip(self, data):
        """分析非ZIP格式的文件"""
        analysis = {}
        
        # 检查常见压缩格式
        magic_map = {
            '52617221': 'RAR',
            '425A68': 'BZIP2',
            '1F8B08': 'GZIP',
            '7801': 'DEFLATE (原始)',
            '789C': 'DEFLATE (常见)',
            '78DA': 'DEFLATE (最大压缩)',
            '7800': 'DEFLATE (无压缩)',
            '57415645': 'WAVE/RIFF',
            '4D5A': 'PE/EXE',
            '7F454C46': 'ELF'
        }
        
        magic_hex = data[:4].hex().upper()
        for magic, format_name in magic_map.items():
            if magic_hex.startswith(magic):
                analysis['detected_format'] = format_name
                break
        
        # 文件头分析
        analysis['header_analysis'] = {
            'total_bytes': len(data),
            'first_16_bytes': data[:16].hex(),
            'last_16_bytes': data[-16:].hex() if len(data) > 16 else data.hex()
        }
        
        return analysis
    
    def compare(self):
        """对比两个文件"""
        original = self.results['original']
        repacked = self.results['repacked']
        
        comparison = {
            'file_size_diff': repacked['file_size'] - original['file_size'],
            'file_size_percent_diff': ((repacked['file_size'] - original['file_size']) / original['file_size']) * 100 if original['file_size'] > 0 else 0,
            'same_sha256': original['sha256'] == repacked['sha256'],
            'both_zip': original['is_zip'] and repacked['is_zip'],
            'same_magic': original['magic_bytes'] == repacked['magic_bytes'],
            'inner_files_count_match': len(original.get('inner_files', [])) == len(repacked.get('inner_files', [])),
            'potential_issues': []
        }
        
        # 检查内部文件
        if original.get('inner_files') and repacked.get('inner_files'):
            orig_files = {f['filename']: f for f in original['inner_files']}
            rep_files = {f['filename']: f for f in repacked['inner_files']}
            
            all_names = set(orig_files.keys()).union(set(rep_files.keys()))
            
            comparison['inner_files_detail'] = []
            for name in all_names:
                orig = orig_files.get(name)
                rep = rep_files.get(name)
                
                detail = {
                    'filename': name,
                    'in_original': name in orig_files,
                    'in_repacked': name in rep_files,
                    'size_match': False,
                    'crc_match': False,
                    'content_match': False,
                    'compression_method_match': False,
                    'encoding_match': False
                }
                
                if orig and rep:
                    detail['size_match'] = orig['uncompressed_size'] == rep['uncompressed_size']
                    detail['crc_match'] = orig['crc32'] == rep['crc32']
                    detail['content_match'] = orig.get('content_sha256') == rep.get('content_sha256')
                    detail['compression_method_match'] = orig['compression_method'] == rep['compression_method']
                    detail['encoding_match'] = orig.get('encoding') == rep.get('encoding')
                    
                    if not detail['compression_method_match']:
                        comparison['potential_issues'].append(
                            f"文件 {name} 压缩方法不同: 原始={orig['compression_name']}, 重新打包={rep['compression_name']}"
                        )
                    if not detail['content_match']:
                        comparison['potential_issues'].append(
                            f"文件 {name} 内容不匹配"
                        )
                
                comparison['inner_files_detail'].append(detail)
        
        # 添加其他潜在问题
        if not comparison['same_magic']:
            comparison['potential_issues'].append(f"文件头魔数不同: 原始={original['magic_bytes']}, 重新打包={repacked['magic_bytes']}")
        
        if not comparison['both_zip']:
            if not original['is_zip']:
                comparison['potential_issues'].append(f"原始文件不是标准ZIP格式，检测到: {original.get('non_zip_analysis', {}).get('detected_format', '未知')}")
            if not repacked['is_zip']:
                comparison['potential_issues'].append(f"重新打包文件不是标准ZIP格式")
        
        if comparison['file_size_percent_diff'] > 10 or comparison['file_size_percent_diff'] < -10:
            comparison['potential_issues'].append(f"文件大小差异超过10%: 原始={original['file_size']}字节, 重新打包={repacked['file_size']}字节")
        
        # 检查ZIP特定差异
        if original.get('zip_info') and repacked.get('zip_info'):
            orig_zip = original['zip_info']
            rep_zip = repacked['zip_info']
            
            if orig_zip['comment'] != rep_zip['comment']:
                comparison['potential_issues'].append("ZIP注释不同")
                comparison['zip_comment_diff'] = {
                    'original': orig_zip['comment'],
                    'repacked': rep_zip['comment']
                }
        
        self.results['comparison'] = comparison
    
    def generate_report(self):
        """生成详细的对比报告"""
        report = []
        
        # 文件基本信息
        report.append("=" * 80)
        report.append("RES文件对比分析报告")
        report.append("=" * 80)
        report.append("")
        
        # 原始文件信息
        report.append("【原始文件】")
        report.append(f"  文件路径: {self.results['original']['file_path']}")
        report.append(f"  文件大小: {self.results['original']['file_size']} 字节")
        report.append(f"  SHA256: {self.results['original']['sha256']}")
        report.append(f"  魔数字节: 0x{self.results['original']['magic_bytes']}")
        report.append(f"  是否ZIP格式: {'是' if self.results['original']['is_zip'] else '否'}")
        
        if self.results['original']['is_zip']:
            report.append(f"  ZIP注释: '{self.results['original']['zip_info']['comment']}'")
            report.append(f"  内部文件数: {self.results['original']['zip_info']['number_of_files']}")
            report.append("  内部文件详情:")
            for inner in self.results['original']['inner_files']:
                report.append(f"    - {inner['filename']}")
                report.append(f"      压缩: {inner['compression_name']}")
                report.append(f"      原始大小: {inner['uncompressed_size']} 字节")
                report.append(f"      压缩大小: {inner['compressed_size']} 字节")
                report.append(f"      CRC32: {inner['crc32']}")
                if 'encoding' in inner:
                    report.append(f"      编码: {inner['encoding']}")
        
        report.append("")
        
        # 重新打包文件信息
        report.append("【重新打包文件】")
        report.append(f"  文件路径: {self.results['repacked']['file_path']}")
        report.append(f"  文件大小: {self.results['repacked']['file_size']} 字节")
        report.append(f"  SHA256: {self.results['repacked']['sha256']}")
        report.append(f"  魔数字节: 0x{self.results['repacked']['magic_bytes']}")
        report.append(f"  是否ZIP格式: {'是' if self.results['repacked']['is_zip'] else '否'}")
        
        if self.results['repacked']['is_zip']:
            report.append(f"  ZIP注释: '{self.results['repacked']['zip_info']['comment']}'")
            report.append(f"  内部文件数: {self.results['repacked']['zip_info']['number_of_files']}")
            report.append("  内部文件详情:")
            for inner in self.results['repacked']['inner_files']:
                report.append(f"    - {inner['filename']}")
                report.append(f"      压缩: {inner['compression_name']}")
                report.append(f"      原始大小: {inner['uncompressed_size']} 字节")
                report.append(f"      压缩大小: {inner['compressed_size']} 字节")
                report.append(f"      CRC32: {inner['crc32']}")
                if 'encoding' in inner:
                    report.append(f"      编码: {inner['encoding']}")
        
        report.append("")
        
        # 对比结果
        report.append("【对比结果】")
        report.append("-" * 40)
        
        comp = self.results['comparison']
        report.append(f" 文件大小差异: {comp['file_size_diff']:+d} 字节 ({comp['file_size_percent_diff']:.2f}%)")
        report.append(f" SHA256是否相同: {'是' if comp['same_sha256'] else '否'}")
        report.append(f" 魔数字节是否相同: {'是' if comp['same_magic'] else '否'}")
        report.append(f" 均为ZIP格式: {'是' if comp['both_zip'] else '否'}")
        report.append(f" 内部文件数量匹配: {'是' if comp['inner_files_count_match'] else '否'}")
        
        report.append("")
        
        # 详细对比内部文件
        if 'inner_files_detail' in comp:
            report.append("【内部文件详细对比】")
            report.append("-" * 40)
            for detail in comp['inner_files_detail']:
                report.append(f" 文件: {detail['filename']}")
                report.append(f"   存在于原始文件: {'是' if detail['in_original'] else '否'}")
                report.append(f"   存在于重新打包: {'是' if detail['in_repacked'] else '否'}")
                if detail['in_original'] and detail['in_repacked']:
                    report.append(f"   大小匹配: {'是' if detail['size_match'] else '否'}")
                    report.append(f"   CRC32匹配: {'是' if detail['crc_match'] else '否'}")
                    report.append(f"   内容匹配: {'是' if detail['content_match'] else '否'}")
                    report.append(f"   压缩方法匹配: {'是' if detail['compression_method_match'] else '否'}")
                    report.append(f"   编码匹配: {'是' if detail['encoding_match'] else '否'}")
                report.append("")
        
        report.append("")
        
        # 潜在问题分析
        report.append("【潜在问题分析】")
        report.append("-" * 40)
        if comp['potential_issues']:
            for i, issue in enumerate(comp['potential_issues'], 1):
                report.append(f" {i}. {issue}")
            
            # 根因分析
            report.append("")
        
        # 根因分析
        report.append("【根因分析】")
        report.append("-" * 40)
        
        # 根据问题类型给出分析
        has_compression_issue = any("压缩方法不同" in issue for issue in comp['potential_issues'])
        has_magic_issue = any("魔数不同" in issue for issue in comp['potential_issues'])
        has_zip_issue = any("不是标准ZIP格式" in issue for issue in comp['potential_issues'])
        has_content_issue = any("内容不匹配" in issue for issue in comp['potential_issues'])
        
        if has_magic_issue or has_zip_issue:
            report.append("*** 可能原因: 原始文件不是标准ZIP格式，而是自定义压缩格式")
            report.append("   建议: 需要分析原始文件的真实格式，可能需要特定的解压算法")
        
        if has_compression_issue:
            report.append("*** 可能原因: 压缩方法不匹配（如原始使用STORED无压缩，重新打包使用DEFLATED）")
            report.append("   建议: 修改打包工具使用与原始文件相同的压缩方法")
        
        if has_content_issue:
            report.append("*** 可能原因: 内部XML文件内容或编码发生变化")
            report.append("   建议: 检查XML文件的编码和内容是否保持原样")
        
        else:
            report.append(" OK 未发现明显差异，问题可能出在其他方面")
            report.append("   建议: 检查文件扩展名、文件权限、路径长度等")
        
        return '\n'.join(report)


def main():
    if len(sys.argv) != 3:
        print("用法: python res_analyzer.py <原始res文件> <重新打包res文件>")
        print("示例: python res_analyzer.py original/RUS.res repacked/RUS.res")
        sys.exit(1)
    
    original_path = sys.argv[1]
    repacked_path = sys.argv[2]
    
    if not os.path.exists(original_path):
        print(f"错误: 原始文件不存在 - {original_path}")
        sys.exit(1)
    
    if not os.path.exists(repacked_path):
        print(f"错误: 重新打包文件不存在 - {repacked_path}")
        sys.exit(1)
    
    analyzer = ResFileAnalyzer()
    analyzer.analyze_file(original_path, 'original')
    analyzer.analyze_file(repacked_path, 'repacked')
    analyzer.compare()
    
    report = analyzer.generate_report()
    print(report)
    
    # 保存报告
    report_path = 'res_comparison_report.txt'
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write(report)
    print(f"\n报告已保存到: {report_path}")


if __name__ == '__main__':
    main()