"""
压缩包处理模块 - 智能解压发票文件
"""
import os
import zipfile
import tarfile
import shutil
from pathlib import Path
from typing import List, Dict, Tuple, Optional
from archive_cache import ArchiveCache


class ArchiveHandler:
    """压缩包处理器"""
    
    # 支持的压缩格式
    SUPPORTED_EXTS = ['.zip', '.rar', '.7z', '.tar', '.tar.gz', '.tgz', '.tar.bz2']
    
    # 发票文件扩展名
    INVOICE_EXTS = ['.pdf', '.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif']
    
    def __init__(self, parent=None):
        self.parent = parent
        self.cache = ArchiveCache()
        self.temp_extract_dir = None
        self.extracted_count = 0
        self.skipped_count = 0
        self.password_skipped = []
        
        # 尝试导入7z支持
        self.has_7z = False
        try:
            import py7zr
            self.has_7z = True
        except ImportError:
            pass
        
        # 尝试导入rar支持
        self.has_rar = False
        try:
            import rarfile
            self.has_rar = True
        except ImportError:
            pass
    
    def is_archive(self, file_path: str) -> bool:
        """检查是否为支持的压缩包"""
        ext = Path(file_path).suffix.lower()
        # 处理 .tar.gz, .tar.bz2, .tgz 等情况
        lower_path = file_path.lower()
        if lower_path.endswith(('.tar.gz', '.tar.bz2', '.tgz')):
            return True
        return ext in self.SUPPORTED_EXTS
    
    def _get_archive_type(self, archive_path: str) -> str:
        """获取压缩包类型，正确处理复合扩展名"""
        lower_path = archive_path.lower()
        if lower_path.endswith('.tar.gz') or lower_path.endswith('.tgz'):
            return 'tar.gz'
        elif lower_path.endswith('.tar.bz2'):
            return 'tar.bz2'
        else:
            return Path(archive_path).suffix.lower()
    
    def scan_archive(self, archive_path: str) -> Dict:
        """
        扫描压缩包内容（不解压）
        返回: {
            'has_invoice': bool,
            'invoice_files': List[str],
            'total_files': int,
            'is_password_protected': bool,
            'error': str
        }
        """
        result = {
            'has_invoice': False,
            'invoice_files': [],
            'total_files': 0,
            'is_password_protected': False,
            'error': None
        }
        
        # 检查缓存
        cache = self.cache.get_cache(archive_path)
        if cache:
            return {
                'has_invoice': cache['has_invoice'],
                'invoice_files': cache['file_list'],
                'total_files': len(cache['file_list']),
                'is_password_protected': cache.get('password_required', False),
                'error': None,
                'from_cache': True
            }
        
        # 使用更准确的类型判断
        archive_type = self._get_archive_type(archive_path)
        
        try:
            if archive_type == '.zip' or archive_path.lower().endswith('.zip'):
                result = self._scan_zip(archive_path)
            elif archive_type == '.7z' and self.has_7z:
                result = self._scan_7z(archive_path)
            elif archive_type == '.rar' and self.has_rar:
                result = self._scan_rar(archive_path)
            elif archive_type in ['.tar', 'tar.gz', 'tar.bz2']:
                result = self._scan_tar(archive_path)
            else:
                result['error'] = f"不支持的压缩格式: {archive_type}"
        except Exception as e:
            result['error'] = str(e)
        
        # 保存缓存
        if not result['error']:
            self.cache.save_cache(
                archive_path,
                result['invoice_files'],
                result['has_invoice'],
                len(result['invoice_files']),
                'scanned',
                result['is_password_protected']
            )
        
        return result
    
    def _scan_zip(self, archive_path: str) -> Dict:
        """扫描ZIP文件"""
        result = {'has_invoice': False, 'invoice_files': [], 'total_files': 0, 
                  'is_password_protected': False, 'error': None}
        
        with zipfile.ZipFile(archive_path, 'r') as zf:
            # 检查是否需要密码
            try:
                zf.testzip()
            except RuntimeError as e:
                if 'password' in str(e).lower() or 'encrypted' in str(e).lower():
                    result['is_password_protected'] = True
                    # 尝试空密码
                    try:
                        zf.namelist()
                    except:
                        pass
            
            # 获取文件列表
            file_list = zf.namelist()
            result['total_files'] = len(file_list)
            
            # 筛选发票文件（遍历所有文件，但限制最大处理数量）
            max_invoice_files = 500  # 最多处理500个发票文件
            for name in file_list:
                if any(name.lower().endswith(ext) for ext in self.INVOICE_EXTS):
                    result['invoice_files'].append(name)
                    result['has_invoice'] = True
                    if len(result['invoice_files']) >= max_invoice_files:
                        break
        
        return result
    
    def _scan_7z(self, archive_path: str) -> Dict:
        """扫描7Z文件"""
        import py7zr
        result = {'has_invoice': False, 'invoice_files': [], 'total_files': 0,
                  'is_password_protected': False, 'error': None}
        
        try:
            with py7zr.SevenZipFile(archive_path, mode='r') as z:
                file_list = z.getnames()
                result['total_files'] = len(file_list)
                
                max_invoice_files = 500  # 最多处理500个发票文件
                for name in file_list:
                    if any(name.lower().endswith(ext) for ext in self.INVOICE_EXTS):
                        result['invoice_files'].append(name)
                        result['has_invoice'] = True
                        if len(result['invoice_files']) >= max_invoice_files:
                            break
        except py7zr.exceptions.PasswordRequired:
            result['is_password_protected'] = True
        
        return result
    
    def _scan_rar(self, archive_path: str) -> Dict:
        """扫描RAR文件"""
        import rarfile
        result = {'has_invoice': False, 'invoice_files': [], 'total_files': 0,
                  'is_password_protected': False, 'error': None}
        
        try:
            with rarfile.RarFile(archive_path) as rf:
                file_list = rf.namelist()
                result['total_files'] = len(file_list)
                
                max_invoice_files = 500  # 最多处理500个发票文件
                for name in file_list:
                    if any(name.lower().endswith(ext) for ext in self.INVOICE_EXTS):
                        result['invoice_files'].append(name)
                        result['has_invoice'] = True
                        if len(result['invoice_files']) >= max_invoice_files:
                            break
        except rarfile.PasswordRequired:
            result['is_password_protected'] = True
        
        return result
    
    def _scan_tar(self, archive_path: str) -> Dict:
        """扫描TAR文件"""
        result = {'has_invoice': False, 'invoice_files': [], 'total_files': 0,
                  'is_password_protected': False, 'error': None}
        
        try:
            with tarfile.open(archive_path, 'r:*') as tf:
                members = tf.getmembers()
                result['total_files'] = len(members)
                
                max_invoice_files = 500  # 最多处理500个发票文件
                for member in members:
                    if member.isfile():
                        name = member.name
                        if any(name.lower().endswith(ext) for ext in self.INVOICE_EXTS):
                            result['invoice_files'].append(name)
                            result['has_invoice'] = True
                            if len(result['invoice_files']) >= max_invoice_files:
                                break
        except Exception as e:
            result['error'] = str(e)
        
        return result
    
    def extract_archive(self, archive_path: str, target_folder: str, 
                        password: str = None, force: bool = False) -> Tuple[List[str], str]:
        """
        解压压缩包中的发票文件
        返回: (extracted_files, status)
        status: 'success', 'password_required', 'error', 'skipped', 'already_extracted'
        
        Args:
            archive_path: 压缩包路径
            target_folder: 解压目标文件夹
            password: 解压密码（如有）
            force: 是否强制重新解压，默认为False
        """
        # 检查是否已经解压过（除非强制重新解压）
        if not force:
            cache = self.cache.get_cache(archive_path)
            if cache and cache['process_status'] == 'extracted':
                # 已经解压过，返回空列表但标记为已处理
                return [], 'already_extracted'
        
        # 先扫描
        scan_result = self.scan_archive(archive_path)
        
        if scan_result['error']:
            return [], f"error: {scan_result['error']}"
        
        if not scan_result['has_invoice']:
            self.cache.update_status(archive_path, 'no_invoice')
            return [], 'no_invoice'
        
        # 检查是否需要密码
        if scan_result['is_password_protected'] and not password:
            return [], 'password_required'
        
        # 解压
        archive_type = self._get_archive_type(archive_path)
        extracted_files = []
        
        try:
            if archive_type == '.zip' or archive_path.lower().endswith('.zip'):
                extracted_files = self._extract_zip(archive_path, target_folder, 
                                                    scan_result['invoice_files'], password)
            elif archive_type == '.7z' and self.has_7z:
                extracted_files = self._extract_7z(archive_path, target_folder,
                                                   scan_result['invoice_files'], password)
            elif archive_type == '.rar' and self.has_rar:
                extracted_files = self._extract_rar(archive_path, target_folder,
                                                    scan_result['invoice_files'], password)
            elif archive_type in ['.tar', 'tar.gz', 'tar.bz2']:
                extracted_files = self._extract_tar(archive_path, target_folder,
                                                    scan_result['invoice_files'])
            
            self.cache.update_status(archive_path, 'extracted')
            self.extracted_count += len(extracted_files)
            
            return extracted_files, 'success'
            
        except Exception as e:
            return [], f"error: {str(e)}"
    
    def _extract_zip(self, archive_path: str, target_folder: str, 
                     file_list: List[str], password: str = None) -> List[str]:
        """解压ZIP文件"""
        extracted = []
        target = Path(target_folder)
        target.mkdir(parents=True, exist_ok=True)
        
        with zipfile.ZipFile(archive_path, 'r') as zf:
            for file_name in file_list:
                # 获取文件名（不含路径）
                base_name = Path(file_name).name
                dest_path = target / base_name
                
                # 检查文件是否已存在
                if dest_path.exists():
                    self.skipped_count += 1
                    continue
                
                # 解压单个文件
                try:
                    if password:
                        zf.extract(file_name, target_folder, pwd=password.encode())
                    else:
                        zf.extract(file_name, target_folder)
                    
                    # 处理解压后的文件路径
                    extracted_path = target / file_name
                    if extracted_path.exists():
                        # 如果文件在子目录中，移动到根目录
                        if '/' in file_name or '\\' in file_name:
                            shutil.move(str(extracted_path), str(dest_path))
                            extracted.append(str(dest_path))
                            # 清理空目录
                            try:
                                parent_dir = extracted_path.parent
                                # 只清理目标文件夹下的空目录
                                while parent_dir != target:
                                    if parent_dir.exists() and not any(parent_dir.iterdir()):
                                        parent_dir.rmdir()
                                    parent_dir = parent_dir.parent
                            except:
                                pass
                        else:
                            extracted.append(str(extracted_path))
                except Exception as e:
                    print(f"解压文件失败 {file_name}: {e}")
                    continue
        
        return extracted
    
    def _extract_7z(self, archive_path: str, target_folder: str,
                    file_list: List[str], password: str = None) -> List[str]:
        """解压7Z文件"""
        import py7zr
        extracted = []
        target = Path(target_folder)
        target.mkdir(parents=True, exist_ok=True)
        
        with py7zr.SevenZipFile(archive_path, mode='r', password=password) as z:
            all_files = z.read()
            for file_name, data in all_files.items():
                if file_name in file_list:
                    # 获取文件名（不含路径），解压到根目录
                    base_name = Path(file_name).name
                    dest_path = target / base_name
                    
                    if dest_path.exists():
                        self.skipped_count += 1
                        continue
                    
                    # 保存文件
                    with open(dest_path, 'wb') as f:
                        f.write(data)
                    extracted.append(str(dest_path))
        
        return extracted
    
    def _extract_rar(self, archive_path: str, target_folder: str,
                     file_list: List[str], password: str = None) -> List[str]:
        """解压RAR文件"""
        import rarfile
        extracted = []
        target = Path(target_folder)
        target.mkdir(parents=True, exist_ok=True)
        
        with rarfile.RarFile(archive_path) as rf:
            for file_name in file_list:
                # 获取文件名（不含路径）
                base_name = Path(file_name).name
                dest_path = target / base_name
                
                if dest_path.exists():
                    self.skipped_count += 1
                    continue
                
                # 解压到临时位置，然后移动
                rf.extract(file_name, target_folder, pwd=password)
                
                # 处理解压后的文件路径
                extracted_path = target / file_name
                if extracted_path.exists():
                    # 如果文件在子目录中，移动到根目录
                    if '/' in file_name or '\\' in file_name:
                        shutil.move(str(extracted_path), str(dest_path))
                        extracted.append(str(dest_path))
                        # 清理空目录
                        try:
                            parent_dir = extracted_path.parent
                            while parent_dir != target:
                                if parent_dir.exists() and not any(parent_dir.iterdir()):
                                    parent_dir.rmdir()
                                parent_dir = parent_dir.parent
                        except:
                            pass
                    else:
                        extracted.append(str(extracted_path))
        
        return extracted
    
    def _extract_tar(self, archive_path: str, target_folder: str,
                     file_list: List[str]) -> List[str]:
        """解压TAR文件"""
        extracted = []
        target = Path(target_folder)
        target.mkdir(parents=True, exist_ok=True)
        
        with tarfile.open(archive_path, 'r:*') as tf:
            for member in tf.getmembers():
                if member.name in file_list:
                    # 获取文件名（不含路径）
                    base_name = Path(member.name).name
                    dest_path = target / base_name
                    
                    if dest_path.exists():
                        self.skipped_count += 1
                        continue
                    
                    # 解压到临时位置，然后移动
                    tf.extract(member, target_folder)
                    
                    # 处理解压后的文件路径
                    extracted_path = target / member.name
                    if extracted_path.exists():
                        # 如果文件在子目录中，移动到根目录
                        if '/' in member.name or '\\' in member.name:
                            shutil.move(str(extracted_path), str(dest_path))
                            extracted.append(str(dest_path))
                            # 清理空目录
                            try:
                                parent_dir = extracted_path.parent
                                while parent_dir != target:
                                    if parent_dir.exists() and not any(parent_dir.iterdir()):
                                        parent_dir.rmdir()
                                    parent_dir = parent_dir.parent
                            except:
                                pass
                        else:
                            extracted.append(str(extracted_path))
        
        return extracted
    
    def cleanup_extracted_files(self, file_list: List[str]):
        """清理解压后的临时文件"""
        for file_path in file_list:
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
            except:
                pass
    
    def get_stats(self) -> Dict:
        """获取处理统计"""
        return {
            'extracted_count': self.extracted_count,
            'skipped_count': self.skipped_count,
            'password_skipped': self.password_skipped
        }


def scan_and_collect_files(folder_path: str, archive_handler: ArchiveHandler,
                           parent=None, recursive: bool = True) -> Tuple[List[str], List[Dict]]:
    """
    扫描文件夹，收集所有发票文件和压缩包信息
    
    Args:
        folder_path: 文件夹路径
        archive_handler: 压缩包处理器
        parent: 父窗口
        recursive: 是否递归扫描子文件夹（默认True）
    
    返回: (normal_files, archive_infos)
    normal_files: 普通发票文件路径列表
    archive_infos: 压缩包信息列表
    """
    from gui.password_dialog import ArchivePasswordDialog
    
    normal_files = []
    archive_infos = []
    
    folder = Path(folder_path)
    
    # 使用 rglob 递归搜索所有文件（如果 recursive=True）
    if recursive:
        file_iterator = folder.rglob('*')
    else:
        file_iterator = folder.iterdir()
    
    for item in file_iterator:
        if item.is_file():
            # 检查是否为压缩包
            if archive_handler.is_archive(str(item)):
                # 扫描压缩包
                scan_result = archive_handler.scan_archive(str(item))
                
                if scan_result['has_invoice']:
                    # 需要处理
                    info = {
                        'path': str(item),
                        'name': item.name,
                        'invoice_files': scan_result['invoice_files'],
                        'is_password_protected': scan_result['is_password_protected'],
                        'password': None,
                        'skip': False
                    }
                    
                    # 如果需要密码，弹窗询问
                    if scan_result['is_password_protected']:
                        password, skip = ArchivePasswordDialog.get_password(item.name, parent)
                        if skip:
                            info['skip'] = True
                            archive_handler.password_skipped.append(item.name)
                        elif password:
                            info['password'] = password
                        else:
                            # 用户取消，跳过
                            continue
                    
                    archive_infos.append(info)
                else:
                    # 不含发票，记录缓存
                    pass
            
            # 检查是否为普通发票文件
            elif item.suffix.lower() in ArchiveHandler.INVOICE_EXTS:
                normal_files.append(str(item))
    
    return normal_files, archive_infos
