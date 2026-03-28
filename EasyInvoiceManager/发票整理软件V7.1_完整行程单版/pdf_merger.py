"""
PDF合并模块 - 用于将同类型发票合并成一个PDF方便打印
"""
import os
from pathlib import Path
from typing import List, Dict
from datetime import datetime
import fitz  # PyMuPDF


class PDFMerger:
    """PDF合并器"""
    
    def __init__(self, base_folder: Path):
        self.base_folder = base_folder
        # 记录每个类型文件夹中的PDF文件（排除重复）
        self.type_pdfs: Dict[str, List[Dict]] = {}
        # 打印日期（用于页脚）
        self.print_date = datetime.now().strftime('%Y年%m月%d日')
    
    def add_pdf(self, pdf_path: str, invoice_type: str, is_duplicate: bool = False):
        """
        添加PDF文件到合并列表
        
        Args:
            pdf_path: PDF文件路径
            invoice_type: 发票类型
            is_duplicate: 是否为重复发票
        """
        # 重复发票不加入合并列表
        if is_duplicate:
            return
        
        # 确保 invoice_type 不为 None
        if invoice_type is None:
            invoice_type = '其他发票'
        
        if invoice_type not in self.type_pdfs:
            self.type_pdfs[invoice_type] = []
        
        self.type_pdfs[invoice_type].append({
            'path': pdf_path,
            'filename': Path(pdf_path).name
        })
    
    def merge_by_type(self, type_amounts: Dict[str, float]) -> Dict[str, str]:
        """
        按类型合并PDF
        
        Args:
            type_amounts: 各类型发票的金额统计（不含重复）
        
        Returns:
            合并后的PDF路径字典 {invoice_type: merged_pdf_path}
        """
        merged_files = {}
        
        # 确保 type_amounts 不为 None
        if type_amounts is None:
            type_amounts = {}
        
        for invoice_type, pdf_list in self.type_pdfs.items():
            if len(pdf_list) < 1:
                continue
            
            # 获取该类型的输出文件夹
            type_folder = self.base_folder / f"{invoice_type}_计算中"
            if not type_folder.exists():
                # 尝试查找已重命名的文件夹
                for folder in self.base_folder.iterdir():
                    if folder.is_dir() and folder.name.startswith(invoice_type):
                        type_folder = folder
                        break
            
            if not type_folder.exists():
                continue
            
            # 生成合并后的文件名（使用精准两位小数的金额）
            total_amount = type_amounts.get(invoice_type, 0)
            output_name = f"{invoice_type}_总额{total_amount:.2f}元_合并.pdf"
            output_path = type_folder / output_name
            
            try:
                # 创建新的PDF文档
                merged_doc = fitz.open()
                
                for pdf_info in pdf_list:
                    pdf_path = pdf_info['path']
                    try:
                        # 打开PDF并追加到合并文档
                        src_doc = fitz.open(pdf_path)
                        merged_doc.insert_pdf(src_doc)
                        src_doc.close()
                    except Exception as e:
                        print(f"合并PDF时出错 {pdf_path}: {e}")
                        continue
                
                # 添加日期页脚到每一页
                if merged_doc.page_count > 0:
                    self._add_date_footer(merged_doc, invoice_type)
                    merged_doc.save(str(output_path))
                    merged_files[invoice_type] = str(output_path)
                
                merged_doc.close()
                
            except Exception as e:
                print(f"合并 {invoice_type} 的PDF失败: {e}")
                continue
        
        # 返回合并后的文件字典
        return merged_files
    
    def _add_date_footer(self, doc: fitz.Document, invoice_type: str):
        """在PDF每一页添加日期页脚"""
        total_pages = len(doc)
        
        for page_num in range(total_pages):
            page = doc[page_num]
            
            # 获取页面尺寸
            page_rect = page.rect
            page_width = page_rect.width
            page_height = page_rect.height
            
            # 页脚文本
            date_text = f"打印日期: {self.print_date}"
            page_text = f"第 {page_num + 1} 页 / 共 {total_pages} 页"
            type_text = f"{invoice_type}"
            
            # 计算位置（页面底部居中）
            footer_y = page_height - 30  # 距离底部30像素
            
            # 添加日期（左侧）
            date_point = fitz.Point(50, footer_y)
            page.insert_text(
                date_point,
                date_text,
                fontsize=10,
                color=(0.3, 0.3, 0.3),  # 深灰色
                fontname="china-ss"  # 使用中文字体
            )
            
            # 添加类型（中间）
            type_width = len(type_text) * 10  # 估算宽度
            type_point = fitz.Point((page_width - type_width) / 2, footer_y)
            page.insert_text(
                type_point,
                type_text,
                fontsize=10,
                color=(0.3, 0.3, 0.3),
                fontname="china-ss"
            )
            
            # 添加页码（右侧）
            page_point = fitz.Point(page_width - 120, footer_y)
            page.insert_text(
                page_point,
                page_text,
                fontsize=10,
                color=(0.3, 0.3, 0.3),
                fontname="china-ss"
            )
    
    def merge_all_types(self, type_amounts: Dict[str, float]) -> Dict[str, str]:
        """合并所有类型的PDF"""
        return self.merge_by_type(type_amounts)


def merge_pdfs_in_folder(folder_path: str, output_name: str = None) -> str:
    """
    合并指定文件夹中的所有PDF（工具函数）
    
    Args:
        folder_path: 文件夹路径
        output_name: 输出文件名（可选）
    
    Returns:
        合并后的PDF路径
    """
    folder = Path(folder_path)
    if not folder.exists():
        raise Exception(f"文件夹不存在: {folder_path}")
    
    # 获取所有PDF文件
    pdf_files = sorted(folder.glob("*.pdf"))
    pdf_files = [f for f in pdf_files if "_合并" not in f.name]  # 排除已合并的文件
    
    if len(pdf_files) < 1:
        raise Exception("文件夹中没有PDF文件")
    
    if output_name is None:
        output_name = f"{folder.name}_合并.pdf"
    
    output_path = folder / output_name
    
    # 合并PDF
    merged_doc = fitz.open()
    
    for pdf_file in pdf_files:
        try:
            src_doc = fitz.open(str(pdf_file))
            merged_doc.insert_pdf(src_doc)
            src_doc.close()
        except Exception as e:
            print(f"合并 {pdf_file.name} 时出错: {e}")
            continue
    
    if merged_doc.page_count > 0:
        # 添加日期页脚
        print_date = datetime.now().strftime('%Y年%m月%d日')
        total_pages = len(merged_doc)
        folder_name = folder.name
        
        for page_num in range(total_pages):
            page = merged_doc[page_num]
            page_rect = page.rect
            page_width = page_rect.width
            page_height = page_rect.height
            
            footer_y = page_height - 30
            
            # 日期（左侧）
            page.insert_text(
                fitz.Point(50, footer_y),
                f"打印日期: {print_date}",
                fontsize=10,
                color=(0.3, 0.3, 0.3),
                fontname="china-ss"
            )
            
            # 文件夹名（中间）
            type_point = fitz.Point((page_width - len(folder_name) * 5) / 2, footer_y)
            page.insert_text(
                type_point,
                folder_name,
                fontsize=10,
                color=(0.3, 0.3, 0.3),
                fontname="china-ss"
            )
            
            # 页码（右侧）
            page.insert_text(
                fitz.Point(page_width - 120, footer_y),
                f"第 {page_num + 1} 页 / 共 {total_pages} 页",
                fontsize=10,
                color=(0.3, 0.3, 0.3),
                fontname="china-ss"
            )
        
        merged_doc.save(str(output_path))
    
    merged_doc.close()
    
    return str(output_path)
