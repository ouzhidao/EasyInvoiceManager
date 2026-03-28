"""
文件整理模块 V2.0
支持增值税发票和铁路电子客票
"""
import shutil
from datetime import datetime
from pathlib import Path
from utils.helpers import sanitize_filename, ensure_dir
from pdf_merger import PDFMerger


class FileOrganizerV2:
    """文件整理器 V2.0"""
    
    def __init__(self, base_output_path: str):
        self.base_output = Path(base_output_path)
        self.current_task_folder = None
        self.type_folders = {}
        self.type_amounts = {}
        self.pdf_merger = None
        self.file_duplicate_status = {}
        
        # 新增：识别失败文件夹
        self.failed_folder = None
        # 新增：铁路客票专用文件夹
        self.railway_folder = None
        # 新增：行程单专用文件夹
        self.itinerary_folder = None
    
    def create_folder_structure(self):
        """创建文件夹结构"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M')
        task_name = f"发票整理{timestamp}"
        
        self.current_task_folder = self.base_output / task_name
        ensure_dir(self.current_task_folder)
        
        # 失败文件夹
        self.failed_folder = self.current_task_folder / "识别失败"
        ensure_dir(self.failed_folder)
        
        # 铁路电子客票专用文件夹（最终名称会在finalize时确定）
        self.railway_folder = self.current_task_folder / "铁路电子客票_计算中"
        ensure_dir(self.railway_folder)
        
        # 行程单专用文件夹
        self.itinerary_folder = self.current_task_folder / "行程单_计算中"
        ensure_dir(self.itinerary_folder)
        
        # 初始化PDF合并器
        self.pdf_merger = PDFMerger(self.current_task_folder)
        self.file_duplicate_status = {}
        
        return self.current_task_folder
    
    def get_type_folder(self, invoice_type: str, amount: float, is_duplicate: bool = False):
        """
        获取类型文件夹
        新增：铁路电子客票和行程单有专用文件夹
        """
        if invoice_type is None:
            invoice_type = '其他发票'
        
        # 铁路电子客票使用专用文件夹
        if invoice_type == '铁路电子客票':
            if '铁路电子客票' not in self.type_folders:
                self.type_folders['铁路电子客票'] = self.railway_folder
                self.type_amounts['铁路电子客票'] = 0.0
            
            if not is_duplicate:
                self.type_amounts['铁路电子客票'] += amount
            return self.railway_folder
        
        # 行程单使用专用文件夹
        if invoice_type == '行程单':
            if '行程单' not in self.type_folders:
                self.type_folders['行程单'] = self.itinerary_folder
                self.type_amounts['行程单'] = 0.0
            
            if not is_duplicate:
                self.type_amounts['行程单'] += amount
            return self.itinerary_folder
        
        # 其他类型发票
        if invoice_type not in self.type_folders:
            folder = self.current_task_folder / f"{invoice_type}_计算中"
            ensure_dir(folder)
            self.type_folders[invoice_type] = folder
            self.type_amounts[invoice_type] = 0.0
        
        if not is_duplicate:
            self.type_amounts[invoice_type] += amount
        return self.type_folders[invoice_type]
    
    def generate_filename(self, data: dict, dup_info: dict, ext: str, invoice_type: str = '增值税发票'):
        """
        生成文件名
        支持格式：
        - 增值税发票：{发票号后5位}-{金额}-{开票方简称}-{日期}
        - 铁路电子客票：铁路_{购买方简称}_{乘车日期}_{出发地}-{到达地}_{金额}_{发票后5位}
        - 行程单：行程_姓名_日期_起点-终点_金额
        """
        if dup_info is None:
            dup_info = {'is_duplicate': False, 'index': 0}
        
        # 铁路电子客票专用命名格式
        if invoice_type == '铁路电子客票':
            return self._generate_railway_filename(data, dup_info, ext)
        
        # 行程单专用命名格式
        if invoice_type == '行程单':
            return self._generate_itinerary_filename(data, dup_info, ext)
        
        # 增值税发票默认格式
        parts = []
        
        if dup_info.get('is_duplicate') and dup_info.get('index', 0) > 0:
            parts.append(f"重复{dup_info['index']}")
        
        # 使用含税金额(amount)，保留两位小数
        amount = data.get('amount', 0)
        if amount is None:
            amount = 0.0
        try:
            amount = float(amount)
        except (ValueError, TypeError):
            amount = 0.0
        
        parts.extend([
            data.get('invoice_num_short', 'XXXXX') if data.get('invoice_num_short') else 'XXXXX',
            f"{amount:.2f}",  # 保留两位小数，含税金额
            data.get('seller_short', '未知') if data.get('seller_short') else '未知',
            data.get('date', datetime.now().strftime('%Y%m%d')) if data.get('date') else datetime.now().strftime('%Y%m%d')
        ])
        
        filename = '-'.join(parts)
        filename = sanitize_filename(filename)
        return filename + ext
    
    def _generate_railway_filename(self, data: dict, dup_info: dict, ext: str) -> str:
        """生成铁路电子客票文件名"""
        if dup_info.get('is_duplicate') and dup_info.get('index', 0) > 0:
            prefix = f"重复{dup_info['index']}_"
        else:
            prefix = ""
        
        # 获取购买方简称（取前10个字符）
        buyer_name = data.get('buyer_name', '未知')
        if len(buyer_name) > 10:
            buyer_short = buyer_name[:10]
        else:
            buyer_short = buyer_name
        
        # 格式化日期（去掉横线）
        travel_date = data.get('travel_date', datetime.now().strftime('%Y%m%d'))
        travel_date = travel_date.replace('-', '')
        
        # 获取出发地和到达地
        departure = data.get('departure', '未知站')
        arrival = data.get('arrival', '未知站')
        
        # 获取金额
        amount = data.get('amount', 0)
        
        # 获取发票后5位
        invoice_num = data.get('invoice_num', '')
        invoice_suffix = invoice_num[-5:] if len(invoice_num) >= 5 else invoice_num
        
        # 组装文件名
        filename = f"{prefix}铁路_{buyer_short}_{travel_date}_{departure}-{arrival}_{amount:.2f}_{invoice_suffix}"
        filename = sanitize_filename(filename)
        return filename + ext
    
    def _generate_itinerary_filename(self, data: dict, dup_info: dict, ext: str) -> str:
        """生成行程单文件名：行程单_金额"""
        if dup_info.get('is_duplicate') and dup_info.get('index', 0) > 0:
            prefix = f"重复{dup_info['index']}_"
        else:
            prefix = ""
        
        # 获取金额
        amount = data.get('amount', 0)
        
        # 组装文件名：行程单_金额
        filename = f"{prefix}行程单_{amount:.2f}"
        filename = sanitize_filename(filename)
        return filename + ext
    
    def move_file(self, src: str, dest_folder: Path, new_name: str, is_duplicate: bool = False):
        """移动文件到目标文件夹"""
        src_path = Path(src)
        
        # 重复发票不复制
        if is_duplicate:
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
        
        # 记录PDF文件的重复状态
        if src_path.suffix.lower() == '.pdf':
            if self.file_duplicate_status is None:
                self.file_duplicate_status = {}
            self.file_duplicate_status[final_path] = is_duplicate
        
        return final_path
    
    def move_failed_file(self, src: str, filename: str):
        """移动识别失败的文件到失败文件夹"""
        if self.failed_folder is None:
            return None
        
        src_path = Path(src)
        dest = self.failed_folder / filename
        
        # 处理冲突
        counter = 1
        original = dest
        while dest.exists():
            stem = original.stem
            suffix = original.suffix
            dest = original.parent / f"{stem}_{counter}{suffix}"
            counter += 1
        
        try:
            shutil.copy2(src, dest)
            return str(dest)
        except Exception as e:
            print(f"复制失败文件时出错: {e}")
            return None
    
    def finalize_folders(self):
        """完成文件夹整理：合并PDF并重命名文件夹"""
        if self.file_duplicate_status is None:
            self.file_duplicate_status = {}
        
        # 第一步：收集每个类型的PDF文件（排除重复）
        for final_path, is_dup in self.file_duplicate_status.items():
            if not is_dup:
                path_obj = Path(final_path)
                for inv_type, folder in self.type_folders.items():
                    if str(path_obj.parent) == str(folder):
                        self.pdf_merger.add_pdf(final_path, inv_type, is_duplicate=False)
                        break
        
        # 第二步：合并PDF
        merged_files = self.pdf_merger.merge_by_type(self.type_amounts)
        
        if merged_files is None:
            merged_files = {}
        
        # 第三步：重命名文件夹
        for inv_type, folder in self.type_folders.items():
            total = self.type_amounts.get(inv_type, 0)
            new_name = f"{inv_type}_总额{total:.2f}元"
            new_folder = self.current_task_folder / sanitize_filename(new_name)
            
            # 检查目标文件夹是否已存在
            if new_folder.exists() and new_folder != folder:
                counter = 1
                while new_folder.exists():
                    new_name = f"{inv_type}_总额{total:.2f}元_{counter}"
                    new_folder = self.current_task_folder / sanitize_filename(new_name)
                    counter += 1
            
            try:
                folder.rename(new_folder)
            except Exception as e:
                print(f"重命名文件夹失败 {folder} -> {new_folder}: {e}")
                continue
            
            # 更新合并PDF的路径
            if inv_type in merged_files:
                old_merged_path = merged_files[inv_type]
                merged_filename = Path(old_merged_path).name
                new_merged_path = new_folder / merged_filename
                merged_files[inv_type] = str(new_merged_path)
        
        return merged_files
    
    def get_failed_folder(self) -> Path:
        """获取识别失败文件夹路径"""
        return self.failed_folder
