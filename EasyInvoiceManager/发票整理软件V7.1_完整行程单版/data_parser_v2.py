"""
数据解析模块 V2.0
支持增值税发票、铁路电子客票和行程单
"""
import re
from utils.helpers import extract_short_name, parse_amount, standardize_date
from railway_parser import RailwayInvoiceParser
from itinerary_parser import ItineraryParser


class DataParserV2:
    """数据解析器 V2.0"""
    
    INVOICE_TYPE_MAP = {
        '电子普通发票': '电子普通发票',
        '增值税电子普通发票': '电子普通发票',
        '电子专用发票': '电子专用发票',
        '数电票': '数电票',
        '全电发票': '数电票',
        '普通发票': '纸质普通发票',
        '专用发票': '纸质专用发票',
        '铁路电子客票': '铁路电子客票',
    }
    
    def __init__(self):
        self.railway_parser = RailwayInvoiceParser()
        self.itinerary_parser = ItineraryParser()
    
    def parse(self, ocr_result: dict) -> list:
        """
        解析OCR结果
        自动判断发票类型并调用相应解析器
        返回：结果列表（支持多段行程展开）
        """
        if ocr_result is None:
            ocr_result = {}
        
        # 获取完整文本
        full_text = self._get_full_text(ocr_result)
        
        # 判断是否为铁路电子客票
        if self.railway_parser.is_railway_invoice(full_text):
            return self._parse_railway(ocr_result, full_text)
        # 判断是否为行程单
        elif self.itinerary_parser.is_itinerary(full_text):
            return self._parse_itinerary(ocr_result, full_text)
        else:
            return self._parse_vat(ocr_result)
    
    def _get_full_text(self, ocr_result: dict) -> str:
        """从OCR结果中提取完整文本"""
        # 优先使用已有的full_text字段
        if 'full_text' in ocr_result:
            return ocr_result['full_text']
        
        words_list = ocr_result.get('words_result', [])
        
        # PaddleOCR格式
        if isinstance(words_list, list) and words_list and isinstance(words_list[0], dict):
            return '\n'.join([w.get('words', '') for w in words_list if isinstance(w, dict)])
        
        # 百度API增值税发票格式
        if isinstance(words_list, dict):
            # 提取words_result字典中的所有文本
            texts = []
            for key, value in words_list.items():
                if isinstance(value, dict) and 'word' in value:
                    texts.append(value['word'])
            return '\n'.join(texts)
        
        return ''
    
    def _parse_railway(self, ocr_result: dict, full_text: str) -> list:
        """解析铁路电子客票 - 返回列表以统一接口"""
        # 调用铁路客票专用解析器
        railway_data = self.railway_parser.parse(ocr_result)
        
        # 构建统一格式的结果
        result = {
            'invoice_type': '铁路电子客票',
            'invoice_num': railway_data.get('invoice_num', ''),
            'invoice_num_full': railway_data.get('invoice_num', ''),
            'invoice_num_short': railway_data.get('invoice_num', '')[-5:] if railway_data.get('invoice_num') else '',
            'buyer_name': railway_data.get('buyer_name', ''),
            'amount': railway_data.get('amount', 0),
            'amount_int': int(railway_data.get('amount', 0)),
            'date': railway_data.get('travel_date', ''),  # 使用乘车日期作为主日期
            'travel_date': railway_data.get('travel_date', ''),
            'invoice_date': railway_data.get('invoice_date', ''),
            'departure': railway_data.get('departure', ''),
            'arrival': railway_data.get('arrival', ''),
            'passenger_name': railway_data.get('passenger_name', ''),
            'departure_time': railway_data.get('departure_time', ''),
            'seat_type': railway_data.get('seat_type', ''),
            'ticket_number': railway_data.get('ticket_number', ''),
            'is_valid': railway_data.get('is_valid', False),
            'missing_fields': railway_data.get('missing_fields', []),
            'is_railway': True,
            'is_itinerary': False,
            'recognize_method': ocr_result.get('source', 'unknown'),
            'segment_count': 1,
            'segment_index': 1,
        }
        
        return [result]
    
    def _parse_itinerary(self, ocr_result: dict, full_text: str) -> list:
        """解析行程单 - 返回列表（支持多段行程展开）"""
        # 调用行程单解析器
        itinerary_list = self.itinerary_parser.parse(ocr_result)
        
        # 统一格式化为列表
        results = []
        for idx, item in enumerate(itinerary_list):
            result = {
                'invoice_type': '行程单',
                'sub_type': item.get('sub_type', '其他'),
                'invoice_num': '',
                'invoice_num_full': '',
                'invoice_num_short': '',
                'buyer_name': '',
                'passenger_name': item.get('passenger_name', ''),
                'amount': item.get('amount', 0),
                'amount_int': int(item.get('amount', 0)),
                'date': item.get('date', ''),
                'time': item.get('time', ''),
                'departure': item.get('departure', ''),
                'arrival': item.get('arrival', ''),
                'vehicle_info': item.get('vehicle_info', ''),
                'distance': item.get('distance', ''),
                'seat_type': item.get('seat_class', ''),
                'is_valid': item.get('is_valid', False),
                'missing_fields': item.get('missing_fields', []),
                'is_railway': False,
                'is_itinerary': True,
                'recognize_method': ocr_result.get('source', 'unknown'),
                'segment_count': item.get('segment_total', 1),
                'segment_index': item.get('segment_index', 1),
            }
            results.append(result)
        
        return results
    
    def _parse_vat(self, ocr_result: dict) -> dict:
        """解析增值税发票"""
        source = ocr_result.get('source', '')
        
        if source == 'baidu_vat':
            return self._parse_vat_baidu(ocr_result)
        else:
            return self._parse_vat_general(ocr_result)
    
    def _parse_vat_baidu(self, ocr_result: dict) -> dict:
        """解析百度API增值税发票结果"""
        words_result = ocr_result.get('words_result', {})
        if words_result is None:
            words_result = {}
        
        invoice_data = {
            'invoice_num': self._extract(words_result, ['InvoiceNum', 'invoice_num', '发票号码']),
            'date': self._extract(words_result, ['InvoiceDate', 'invoice_date', '开票日期']),
            'amount': self._extract(words_result, ['AmountInFiguers', 'amount_in_figuers', '价税合计(小写)', 'TotalAmount', 'total_amount', '价税合计']),
            'seller_name': self._extract(words_result, ['SellerName', 'seller_name', '销售方名称']),
            'buyer_name': self._extract(words_result, ['BuyerName', 'buyer_name', '购买方名称', 'PurchaserName']),
            'invoice_type': self._extract(words_result, ['InvoiceType', 'invoice_type', '发票种类']),
            'content': self._extract_content(words_result),
        }
        
        return self._format_vat_result(invoice_data, ocr_result.get('source', 'baidu_vat'))
    
    def _parse_vat_general(self, ocr_result: dict) -> dict:
        """解析通用OCR结果（增值税发票）"""
        words_list = ocr_result.get('words_result', [])
        if words_list is None:
            words_list = []
        
        # 提取完整文本
        if isinstance(words_list, list) and words_list and isinstance(words_list[0], dict):
            full_text = '\n'.join([w.get('words', '') for w in words_list if isinstance(w, dict)])
        else:
            full_text = ''
        
        # 用正则提取关键信息
        invoice_data = {
            'invoice_num': self._regex_extract(full_text, [
                r'发票号码[：:]?\s*(\d{8,20})',
                r'No[.:]?\s*(\d{8,20})',
                r'(\d{20})',
            ]),
            'date': self._regex_extract(full_text, [
                r'(\d{4}[年/-]\d{1,2}[月/-]\d{1,2}[日]?)',
                r'开票日期[：:]?\s*(\d{4}[年/-]\d{1,2}[月/-]\d{1,2}[日]?)',
            ]),
            'amount': self._regex_extract(full_text, [
                # 优先匹配价税合计（含税总金额）
                r'价税合计[（(]小写[)）]?[：:]?\s*[¥￥]?(\d[\d,]*\.?\d{0,2})',
                r'价税合计[：:]?\s*[¥￥]?(\d[\d,]*\.?\d{0,2})',
                r'价税合计.*?[¥￥](\d[\d,]*\.?\d{0,2})',
                # 然后是含税合计
                r'合\s*计[（(]小写[)）]?[：:]?\s*[¥￥]?(\d[\d,]*\.?\d{0,2})',
                # 最后是任意金额（作为备选）
                r'[¥￥](\d[\d,]*\.?\d{0,2})',
            ]),
            'seller_name': self._regex_extract(full_text, [
                r'销售方[：:]?\s*([^\n]{2,30}公司)',
                r'销\s*售\s*方[：:]?\s*名称\s*([^\n]{2,30})',
            ]),
            'buyer_name': self._regex_extract(full_text, [
                r'购买方[：:]?\s*名称\s*([^\n]{2,50})',
                r'购\s*买\s*方[：:]?\s*名称\s*([^\n]{2,50})',
                r'买方[：:]?\s*名称\s*([^\n]{2,50})',
                r'购方[：:]?\s*名称\s*([^\n]{2,50})',
            ]),
            'invoice_type': self._detect_type(full_text),
            'content': self._extract_content_from_text(full_text),
        }
        
        return self._format_vat_result(invoice_data, ocr_result.get('source', 'unknown'))
    
    def _format_vat_result(self, invoice_data: dict, source: str) -> dict:
        """格式化增值税发票结果"""
        if invoice_data is None:
            invoice_data = {}
        
        invoice_num = invoice_data.get('invoice_num') or ''
        date = invoice_data.get('date') or ''
        amount_raw = invoice_data.get('amount') or '0'
        seller_name = invoice_data.get('seller_name') or ''
        buyer_name = invoice_data.get('buyer_name') or ''
        invoice_type = invoice_data.get('invoice_type') or '其他发票'
        content = invoice_data.get('content') or ''
        
        # 解析金额
        parsed_amount = parse_amount(amount_raw)
        if parsed_amount is None:
            parsed_amount = 0.0
        
        try:
            amount_int = int(parsed_amount)
        except (ValueError, TypeError):
            amount_int = 0
        
        result = {
            'invoice_type': self.INVOICE_TYPE_MAP.get(invoice_type, invoice_type or '其他发票'),
            'invoice_num_full': invoice_num,
            'invoice_num_short': invoice_num[-5:] if len(invoice_num) >= 5 else invoice_num,
            'date': standardize_date(date),
            'amount': parsed_amount,
            'amount_int': amount_int,
            'seller_name': seller_name,
            'seller_short': extract_short_name(seller_name),
            'buyer_name': buyer_name,
            'content': content,
            'is_valid': bool(invoice_num and (date or parsed_amount > 0)),
            'is_railway': False,
            'recognize_method': source,
        }
        
        return result
    
    def _extract(self, data: dict, keys: list):
        """从字典中提取字段"""
        if data is None:
            data = {}
        
        for key in keys:
            if key in data:
                value = data[key]
                if isinstance(value, dict):
                    return value.get('word', '') or ''
                return str(value) if value is not None else ''
        return ''
    
    def _regex_extract(self, text: str, patterns: list):
        """用正则提取"""
        if text is None:
            text = ''
        
        for pattern in patterns:
            try:
                match = re.search(pattern, text, re.IGNORECASE)
                if match:
                    return match.group(1).strip() if match.group(1) else ''
            except re.error:
                continue
        return ''
    
    def _detect_type(self, text: str):
        """检测发票类型"""
        if text is None:
            text = ''
        
        text = text.upper()
        
        if '电子专票' in text or '电子专用发票' in text:
            return '电子专用发票'
        elif '电子普票' in text or '电子普通发票' in text:
            return '电子普通发票'
        elif '全电' in text or '数电' in text:
            return '数电票'
        elif '专票' in text or '专用发票' in text:
            return '纸质专用发票'
        elif '普票' in text or '普通发票' in text:
            return '纸质普通发票'
        
        return '其他发票'
    
    def _extract_content(self, words_result: dict) -> str:
        """从增值税发票结果中提取商品/服务内容"""
        if words_result is None:
            words_result = {}
        
        full_text = ''
        if 'words_result' in words_result:
            for item in words_result.get('words_result', []):
                if isinstance(item, dict):
                    full_text += item.get('words', '') + '\n'
        
        # 优先匹配星号之间的内容
        if full_text:
            star_pattern = r'\*\*([^\*]+)\*\*'
            try:
                matches = re.findall(star_pattern, full_text)
                if matches:
                    content = matches[0].strip()
                    if len(content) >= 1:
                        return content
            except re.error:
                pass
            
            single_star_pattern = r'\*([^\*]+)\*'
            try:
                matches = re.findall(single_star_pattern, full_text)
                if matches:
                    content = matches[0].strip()
                    if len(content) >= 1:
                        return content
            except re.error:
                pass
        
        # 尝试直接提取 CommodityName
        content = ''
        if 'CommodityName' in words_result:
            commodity = words_result['CommodityName']
            if isinstance(commodity, dict):
                content = commodity.get('word', '') or ''
            elif isinstance(commodity, list):
                items = [c.get('word', '') if isinstance(c, dict) else str(c) for c in commodity if c]
                content = items[0] if items else ''
            else:
                content = str(commodity) if commodity is not None else ''
        
        return self._summarize_content(content)
    
    def _extract_content_from_text(self, full_text: str) -> str:
        """从通用识别文本中提取商品/服务内容"""
        if full_text is None:
            full_text = ''
        
        # 优先匹配双星号之间的内容
        star_pattern = r'\*\*([^\*]+)\*\*'
        try:
            matches = re.findall(star_pattern, full_text)
            if matches:
                content = matches[0].strip()
                if len(content) >= 1:
                    return content
        except re.error:
            pass
        
        # 匹配单星号之间的内容
        single_star_pattern = r'\*([^\*]+)\*'
        try:
            matches = re.findall(single_star_pattern, full_text)
            if matches:
                content = matches[0].strip()
                if len(content) >= 1:
                    return content
        except re.error:
            pass
        
        return ''
    
    def _summarize_content(self, content: str) -> str:
        """概括内容"""
        if not content:
            return ''
        
        content = content.strip()
        
        # 清理内容
        cleaned = re.sub(r'[\d\*\s\.\-\\/]', '', content)
        
        if len(cleaned) <= 10 and len(cleaned) >= 2:
            return cleaned
        
        if len(cleaned) > 10:
            cleaned = cleaned[:10]
        
        return cleaned if cleaned else '其他'
