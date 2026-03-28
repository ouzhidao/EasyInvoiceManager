"""
文件整理模块
"""
import shutil
from datetime import datetime
from pathlib import Path
from utils.helpers import sanitize_filename, ensure_dir
from pdf_merger import PDFMerger


class FileOrganizer:
    def __init__(self, base_output_path: str):
        self.base_output = Path(base_output_path)
        self.current_task_folder = None
        self.type_folders = {}
        self.type_amounts = {}
        self.pdf_merger = None
        # 记录每个文件的重复状态，用于PDF合并判断
        self.file_duplicate_status = {}
    
    def create_folder_structure(self):
        timestamp = datetime.now().strftime('%Y%m%d_%H%M')
        task_name = f"发票整理{timestamp}"
        
        self.current_task_folder = self.base_output / task_name
        ensure_dir(self.current_task_folder)
        
        # 失败文件夹
        ensure_dir(self.current_task_folder / "失败待处理")
        
        # 初始化PDF合并器
        self.pdf_merger = PDFMerger(self.current_task_folder)
        self.file_duplicate_status = {}
        
        return self.current_task_folder
    
    def get_type_folder(self, invoice_type: str, amount: float, is_duplicate: bool = False):
        """获取类型文件夹，重复发票不计入金额统计"""
        # 确保 invoice_type 不为 None
        if invoice_type is None:
            invoice_type = '其他发票'
        
        if invoice_type not in self.type_folders:
            folder = self.current_task_folder / f"{invoice_type}_计算中"
            ensure_dir(folder)
            self.type_folders[invoice_type] = folder
            self.type_amounts[invoice_type] = 0.0  # 使用浮点数确保精度
        
        # 只有非重复发票才计入金额统计
        if not is_duplicate:
            self.type_amounts[invoice_type] += amount
        return self.type_folders[invoice_type]
    
    def finalize_folders(self):
        """完成文件夹整理：合并PDF并重命名文件夹"""
        # 确保 file_duplicate_status 不为 None
        if self.file_duplicate_status is None:
            self.file_duplicate_status = {}
        
        # 第一步：收集每个类型的PDF文件（排除重复）
        for final_path, is_dup in self.file_duplicate_status.items():
            if not is_dup:  # 只处理非重复的PDF
                path_obj = Path(final_path)
                # 找到对应的类型
                for inv_type, folder in self.type_folders.items():
                    if str(path_obj.parent) == str(folder):
                        self.pdf_merger.add_pdf(final_path, inv_type, is_duplicate=False)
                        break
        
        # 第二步：合并PDF（保存在原始文件夹中）
        merged_files = self.pdf_merger.merge_by_type(self.type_amounts)
        
        # 确保 merged_files 不为 None
        if merged_files is None:
            merged_files = {}
        
        # 第三步：重命名文件夹
        for inv_type, folder in self.type_folders.items():
            total = self.type_amounts.get(inv_type, 0)
            new_name = f"{inv_type}_总额{total:.2f}元"
            new_folder = self.current_task_folder / sanitize_filename(new_name)
            
            # 检查目标文件夹是否已存在
            if new_folder.exists() and new_folder != folder:
                # 如果目标已存在，添加序号
                counter = 1
                original_new_folder = new_folder
                while new_folder.exists():
                    new_name = f"{inv_type}_总额{total:.2f}元_{counter}"
                    new_folder = self.current_task_folder / sanitize_filename(new_name)
                    counter += 1
            
            try:
                folder.rename(new_folder)
            except Exception as e:
                print(f"重命名文件夹失败 {folder} -> {new_folder}: {e}")
                continue
            
            # 如果生成了合并PDF，更新其路径（因为文件夹已重命名）
            if inv_type in merged_files:
                old_merged_path = merged_files[inv_type]
                # 更新路径为新的文件夹路径
                merged_filename = Path(old_merged_path).name
                new_merged_path = new_folder / merged_filename
                merged_files[inv_type] = str(new_merged_path)
        
        return merged_files
    
    def generate_filename(self, data: dict, dup_info: dict, ext: str):
        # 确保 dup_info 有默认值
        if dup_info is None:
            dup_info = {'is_duplicate': False, 'index': 0}
        
        parts = []
        
        if dup_info.get('is_duplicate') and dup_info.get('index', 0) > 0:
            parts.append(f"重复{dup_info['index']}")
        
        # 确保 data 字段有默认值
        parts.extend([
            data.get('invoice_num_short', 'XXXXX') if data.get('invoice_num_short') else 'XXXXX',
            str(data.get('amount_int', 0)) if data.get('amount_int') is not None else '0',
            data.get('seller_short', '未知') if data.get('seller_short') else '未知',
            data.get('date', datetime.now().strftime('%Y%m%d')) if data.get('date') else datetime.now().strftime('%Y%m%d')
        ])
        
        filename = '-'.join(parts)
        filename = sanitize_filename(filename)
        
        return filename + ext
    
    def move_file(self, src: str, dest_folder: Path, new_name: str, is_duplicate: bool = False):
        """移动文件到目标文件夹，重复发票不复制到输出文件夹"""
        src_path = Path(src)
        
        # 如果是重复发票，不复制到输出文件夹，返回原始路径
        if is_duplicate:
            # 记录为重复，但不实际复制文件
            return str(src_path)
        
        dest = dest_folder / new_name
        
        # 处理冲突
        counter = 1
        original = dest
        while dest.exists():
            stem = original.stem
            suffix = original.suffix
            dest = original.parent / f"{stem}_{counter}{suffix}"
            counter += 1
        
        shutil.copy2(src, dest)
        final_path = str(dest)
        
        # 记录PDF文件的重复状态（用于后续合并）
        if src_path.suffix.lower() == '.pdf':
            if self.file_duplicate_status is None:
                self.file_duplicate_status = {}
            self.file_duplicate_status[final_path] = is_duplicate
        
        return final_path
