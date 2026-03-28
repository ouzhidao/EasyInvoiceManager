"""
OCR引擎模块
支持百度OCR（在线）- 增强版支持PDF
"""
import base64
import io
import requests
import fitz  # PyMuPDF
from pathlib import Path
from PIL import Image


class BaiduOCRClient:
    """百度OCR客户端"""
    
    def __init__(self, app_id: str, api_key: str, secret_key: str):
        self.app_id = app_id
        self.api_key = api_key
        self.secret_key = secret_key
        self.access_token = None
        
        self._get_access_token()
    
    def _get_access_token(self):
        """获取访问令牌"""
        url = "https://aip.baidubce.com/oauth/2.0/token"
        params = {
            "grant_type": "client_credentials",
            "client_id": self.api_key,
            "client_secret": self.secret_key
        }
        
        try:
            response = requests.post(url, params=params, timeout=10)
            result = response.json()
            
            if 'access_token' in result:
                self.access_token = result['access_token']
            else:
                raise Exception(f"获取token失败: {result}")
        except Exception as e:
            raise Exception(f"获取access_token失败: {e}")
    
    def recognize_vat_invoice(self, image_data: bytes, timeout: int = 30):
        """
        增值税发票专用识别（高精度）
        """
        if not self.access_token:
            self._get_access_token()
        
        url = f"https://aip.baidubce.com/rest/2.0/ocr/v1/vat_invoice?access_token={self.access_token}"
        
        img_base64 = base64.b64encode(image_data).decode('utf-8')
        
        headers = {'Content-Type': 'application/x-www-form-urlencoded'}
        params = {'image': img_base64, 'type': 'normal'}
        
        response = requests.post(url, data=params, headers=headers, timeout=timeout)
        result = response.json()
        
        if 'error_code' in result:
            raise Exception(f"{result.get('error_msg', '未知错误')}")
        
        return result
    
    def recognize_general(self, image_data: bytes, timeout: int = 30):
        """
        通用文字识别（当发票识别失败时使用）
        """
        if not self.access_token:
            self._get_access_token()
        
        url = f"https://aip.baidubce.com/rest/2.0/ocr/v1/accurate_basic?access_token={self.access_token}"
        
        img_base64 = base64.b64encode(image_data).decode('utf-8')
        
        headers = {'Content-Type': 'application/x-www-form-urlencoded'}
        params = {
            'image': img_base64,
            'detect_direction': 'true'
        }
        
        response = requests.post(url, data=params, headers=headers, timeout=timeout)
        result = response.json()
        
        if 'error_code' in result:
            raise Exception(f"{result.get('error_msg', '未知错误')}")
        
        return result


class OCREngine:
    """OCR引擎主类"""
    
    def __init__(self):
        from config_manager import ConfigManager
        self.config = ConfigManager()
        self.baidu_client = None
        
        self._init_baidu()
    
    def _init_baidu(self):
        """初始化百度OCR"""
        creds = self.config.get_api_credentials()
        if all([creds['app_id'], creds['api_key'], creds['secret_key']]):
            try:
                self.baidu_client = BaiduOCRClient(
                    creds['app_id'],
                    creds['api_key'],
                    creds['secret_key']
                )
            except Exception as e:
                print(f"百度OCR初始化失败: {e}")
    
    def _convert_pdf_to_image(self, pdf_path: str, dpi: int = 200) -> bytes:
        """
        将PDF第一页转换为图片
        """
        try:
            doc = fitz.open(pdf_path)
            page = doc[0]  # 取第一页
            
            # 设置缩放
            mat = fitz.Matrix(dpi/72, dpi/72)
            pix = page.get_pixmap(matrix=mat)
            
            # 转换为PIL Image
            img_data = pix.tobytes("png")
            img = Image.open(io.BytesIO(img_data))
            
            # 转换为JPEG字节
            buffer = io.BytesIO()
            img.convert('RGB').save(buffer, format='JPEG', quality=95)
            
            doc.close()
            
            return buffer.getvalue()
        except Exception as e:
            raise Exception(f"PDF转换失败: {e}")
    
    def _convert_to_image_data(self, file_path: str) -> bytes:
        """
        将文件转换为图片字节数据
        """
        ext = Path(file_path).suffix.lower()
        
        if ext == '.pdf':
            # PDF需要转换
            return self._convert_pdf_to_image(file_path)
        elif ext in ['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif']:
            # 图片直接读取
            with open(file_path, 'rb') as f:
                return f.read()
        else:
            raise Exception(f"不支持的文件格式: {ext}")
    
    def recognize(self, file_path: str):
        """
        识别发票
        策略：先尝试增值税发票识别，失败则使用通用识别
        """
        if not self.baidu_client:
            raise Exception("百度OCR未配置，请先配置API密钥")
        
        # 转换为图片数据
        image_data = self._convert_to_image_data(file_path)
        
        # 先尝试增值税发票识别（高精度）
        try:
            result = self.baidu_client.recognize_vat_invoice(image_data)
            result['recognize_type'] = 'vat_invoice'
            result['source'] = 'baidu_vat'
            return result
        except Exception as e:
            error_msg = str(e)
            
            # 如果是格式错误，尝试通用识别
            if 'format' in error_msg.lower() or 'template' in error_msg.lower():
                try:
                    result = self.baidu_client.recognize_general(image_data)
                    result['recognize_type'] = 'general'
                    result['source'] = 'baidu_general'
                    return result
                except Exception as e2:
                    raise Exception(f"发票识别失败: {e2}")
            else:
                raise Exception(f"发票识别失败: {e}")
