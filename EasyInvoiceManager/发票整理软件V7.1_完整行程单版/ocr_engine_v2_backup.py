"""
OCR引擎模块 V2.0
双引擎支持：百度API（优先，如果配置了）+ PaddleOCR本地识别（备用）

注意：如果PaddleOCR在您的机器上报错（如ConvertPirAttribute2RuntimeAttribute错误），
      请尝试降级PaddlePaddle版本：
      pip install paddlepaddle==2.6.1 -i https://pypi.tuna.tsinghua.edu.cn/simple
"""
import base64
import io
import os
import re
import requests
import fitz  # PyMuPDF
from pathlib import Path
from PIL import Image


class PaddleOCREngine:
    """PaddleOCR本地识别引擎"""
    
    def __init__(self):
        self.ocr = None
        self._init_engine()
    
    def _init_engine(self):
        """初始化PaddleOCR引擎"""
        try:
            from paddleocr import PaddleOCR
            import logging
            # 设置Paddle日志级别为WARNING，减少冗余输出
            logging.getLogger('paddle').setLevel(logging.WARNING)
            logging.getLogger('ppocr').setLevel(logging.WARNING)
            
            self.ocr = PaddleOCR(lang='ch')
            print("PaddleOCR初始化成功")
        except ImportError as e:
            raise Exception(f"PaddleOCR未安装，请执行: pip install paddleocr -i https://pypi.tuna.tsinghua.edu.cn/simple")
        except Exception as e:
            raise Exception(f"PaddleOCR初始化失败: {e}")
    
    def recognize(self, image_path: str) -> dict:
        """
        识别图片/PDF中的文字
        """
        try:
            # 转换PDF为图片（如果是PDF）
            ext = Path(image_path).suffix.lower()
            if ext == '.pdf':
                image_data = self._convert_pdf_to_image(image_path)
                # 临时保存图片
                temp_img_path = image_path + '.temp.jpg'
                with open(temp_img_path, 'wb') as f:
                    f.write(image_data)
                try:
                    result = self.ocr.ocr(temp_img_path)
                finally:
                    # 清理临时文件
                    Path(temp_img_path).unlink(missing_ok=True)
            else:
                result = self.ocr.ocr(image_path)
            
            # 转换为统一格式
            words_result = []
            full_text_lines = []
            
            if result and result[0]:
                for line in result[0]:
                    if line:
                        text = line[1][0]  # 识别文本
                        confidence = line[1][1]  # 置信度
                        words_result.append({
                            'words': text,
                            'confidence': confidence
                        })
                        full_text_lines.append(text)
            
            return {
                'words_result': words_result,
                'words_result_num': len(words_result),
                'full_text': '\n'.join(full_text_lines),
                'source': 'paddleocr',
                'success': True
            }
            
        except Exception as e:
            error_msg = str(e)
            # 检查是否为已知的PaddlePaddle内部错误
            if "ConvertPirAttribute2RuntimeAttribute" in error_msg:
                error_msg += "\n\n【解决方案】请尝试降级PaddlePaddle版本：\n"
                error_msg += "pip install paddlepaddle==2.6.1 -i https://pypi.tuna.tsinghua.edu.cn/simple\n\n"
                error_msg += "或者配置百度API使用在线识别。"
            
            print(f"PaddleOCR识别异常: {error_msg}")
            return {
                'words_result': [],
                'words_result_num': 0,
                'full_text': '',
                'source': 'paddleocr',
                'success': False,
                'error': error_msg
            }
    
    def _convert_pdf_to_image(self, pdf_path: str, dpi: int = 200) -> bytes:
        """将PDF第一页转换为图片"""
        try:
            doc = fitz.open(pdf_path)
            page = doc[0]
            mat = fitz.Matrix(dpi/72, dpi/72)
            pix = page.get_pixmap(matrix=mat)
            img_data = pix.tobytes("png")
            img = Image.open(io.BytesIO(img_data))
            buffer = io.BytesIO()
            img.convert('RGB').save(buffer, format='JPEG', quality=95)
            doc.close()
            return buffer.getvalue()
        except Exception as e:
            raise Exception(f"PDF转换失败: {e}")


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
        """增值税发票专用识别"""
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
        """通用文字识别"""
        if not self.access_token:
            self._get_access_token()
        
        url = f"https://aip.baidubce.com/rest/2.0/ocr/v1/accurate_basic?access_token={self.access_token}"
        img_base64 = base64.b64encode(image_data).decode('utf-8')
        headers = {'Content-Type': 'application/x-www-form-urlencoded'}
        params = {'image': img_base64, 'detect_direction': 'true'}
        
        response = requests.post(url, data=params, headers=headers, timeout=timeout)
        result = response.json()
        
        if 'error_code' in result:
            raise Exception(f"{result.get('error_msg', '未知错误')}")
        
        return result


class OCREngineV2:
    """
    OCR引擎主类 V2.0
    
    识别策略：
    1. 如果配置了百度API，优先使用API（精度更高、更稳定）
    2. 如果API失败或额度用完，降级到PaddleOCR本地识别
    3. 如果没配置API，直接使用PaddleOCR本地识别
    
    注意：如果PaddleOCR在您的机器上报错，请配置百度API或降级PaddlePaddle版本
    
    环境变量：
    - FORCE_API_ONLY=1: 强制只使用API，不加载本地OCR（解决PaddleOCR兼容性问题）
    """
    
    def __init__(self):
        from config_manager import ConfigManager
        self.config = ConfigManager()
        self.paddle_engine = None
        self.baidu_client = None
        self.force_api_only = os.environ.get('FORCE_API_ONLY', '0') == '1'
        
        if self.force_api_only:
            print("[警告] 强制纯API模式（FORCE_API_ONLY=1），跳过本地OCR初始化")
        
        self._init_engines()
    
    def _init_engines(self):
        """初始化引擎"""
        # 1. 优先初始化百度API（如果配置了）
        try:
            creds = self.config.get_api_credentials()
            if all([creds.get('app_id'), creds.get('api_key'), creds.get('secret_key')]):
                self.baidu_client = BaiduOCRClient(
                    creds['app_id'],
                    creds['api_key'],
                    creds['secret_key']
                )
                print("[OK] 百度API引擎已配置（将优先使用）")
            else:
                print("[!] 百度API未配置")
        except Exception as e:
            print(f"[X] 百度API初始化失败: {e}")
            self.baidu_client = None
        
        # 2. 初始化PaddleOCR本地引擎（作为备用，除非强制纯API模式）
        if not self.force_api_only:
            try:
                self.paddle_engine = PaddleOCREngine()
                print("[OK] PaddleOCR本地引擎已就绪（作为备用）")
            except Exception as e:
                print(f"[X] PaddleOCR初始化失败: {e}")
                print("    如果报错包含'ConvertPirAttribute2RuntimeAttribute'，请尝试：")
                print("    1. 降级PaddlePaddle: pip install paddlepaddle==2.6.1")
                print("    2. 或设置环境变量 FORCE_API_ONLY=1 使用纯API模式")
                self.paddle_engine = None
        else:
            print("[!] 跳过PaddleOCR初始化（纯API模式）")
    
    def recognize(self, file_path: str):
        """
        智能识别发票
        """
        file_path = str(file_path)
        
        # 第一步：如果配置了百度API，优先使用
        if self.baidu_client:
            try:
                print(f"使用百度API识别: {Path(file_path).name}")
                image_data = self._convert_to_image_data(file_path)
                
                # 先尝试增值税发票识别
                try:
                    result = self.baidu_client.recognize_vat_invoice(image_data)
                    result['recognize_type'] = 'vat_invoice'
                    result['source'] = 'baidu_vat'
                    print("  [OK] 百度API识别成功（增值税发票接口）")
                    return result
                except Exception as e:
                    error_msg = str(e)
                    # 增值税识别失败，使用通用识别
                    if 'format' in error_msg.lower() or 'template' in error_msg.lower():
                        result = self.baidu_client.recognize_general(image_data)
                        result['recognize_type'] = 'general'
                        result['source'] = 'baidu_general'
                        print("  [OK] 百度API识别成功（通用接口）")
                        return result
                    else:
                        # API其他错误（如额度用完），降级到本地
                        print(f"  [!] 百度API错误: {error_msg}")
                        print("  将降级到PaddleOCR本地识别...")
                        # 继续往下执行，尝试本地识别
                        
            except Exception as e:
                print(f"  [!] 百度API识别失败: {e}")
                print("  将尝试PaddleOCR本地识别...")
                # 继续往下执行，尝试本地识别
        
        # 第二步：使用PaddleOCR本地识别（API未配置或失败时）
        if self.paddle_engine:
            try:
                print(f"使用PaddleOCR识别: {Path(file_path).name}")
                result = self.paddle_engine.recognize(file_path)
                
                if result.get('success') and result.get('words_result'):
                    # 检查识别质量
                    quality = self._check_quality(result)
                    
                    result['recognize_type'] = 'paddleocr'
                    result['source'] = 'paddleocr'
                    result['quality'] = quality
                    
                    print(f"  [OK] PaddleOCR识别成功（质量: {quality}）")
                    return result
                else:
                    error_msg = result.get('error', '未知错误')
                    print(f"  [X] PaddleOCR识别未成功")
                    print(f"      错误: {error_msg}")
                    
            except Exception as e:
                print(f"  [X] PaddleOCR识别异常: {e}")
        
        # 所有引擎都失败
        error_msg = "所有OCR引擎均识别失败"
        if not self.baidu_client and not self.paddle_engine:
            error_msg += "\n\n请至少配置一种OCR引擎：\n"
            error_msg += "1. 配置百度API（推荐）：点击'配置API'按钮\n"
            error_msg += "2. 或修复PaddleOCR：pip install paddlepaddle==2.6.1"
        elif not self.paddle_engine:
            error_msg += "\n\nPaddleOCR不可用，请配置百度API或修复PaddleOCR"
        
        raise Exception(error_msg)
    
    def _check_quality(self, result: dict) -> str:
        """
        检查OCR识别质量
        返回: 'high' | 'medium' | 'low'
        """
        full_text = result.get('full_text', '')
        
        if not full_text:
            return 'low'
        
        # 检查是否有关键发票标识
        has_invoice_keywords = any(kw in full_text for kw in [
            '发票', '发票号码', '价税合计', '电子发票',
            '铁路电子客票', '12306', '买票请到12306',
            '国家税务总局', '统一发票'
        ])
        
        # 检查是否有金额信息
        has_amount = bool(re.search(r'[¥￥]\s*\d+[.,]?\d*', full_text))
        
        # 检查是否有日期
        has_date = bool(re.search(r'\d{4}[年/-]\d{1,2}[月/-]\d{1,2}', full_text))
        
        score = sum([has_invoice_keywords, has_amount, has_date])
        
        if score >= 3:
            return 'high'
        elif score >= 2:
            return 'medium'
        else:
            return 'low'
    
    def _convert_to_image_data(self, file_path: str) -> bytes:
        """将文件转换为图片字节数据"""
        ext = Path(file_path).suffix.lower()
        
        if ext == '.pdf':
            return self._convert_pdf_to_image(file_path)
        elif ext in ['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif']:
            with open(file_path, 'rb') as f:
                return f.read()
        else:
            raise Exception(f"不支持的文件格式: {ext}")
    
    def _convert_pdf_to_image(self, pdf_path: str, dpi: int = 300) -> bytes:
        """将PDF第一页转换为图片"""
        try:
            doc = fitz.open(pdf_path)
            page = doc[0]
            mat = fitz.Matrix(dpi/72, dpi/72)
            pix = page.get_pixmap(matrix=mat)
            img_data = pix.tobytes("png")
            doc.close()
            return img_data
        except Exception as e:
            raise Exception(f"PDF转换失败: {e}")
