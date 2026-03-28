"""
铁路电子客票专用解析模块
"""
import re
from typing import Dict, List, Tuple


class RailwayInvoiceParser:
    """铁路电子客票解析器"""
    
    # 铁路客票关键词（用于识别）
    RAILWAY_KEYWORDS = [
        '铁路电子客票',
        '买票请到12306',
        '发货请到95306',
        '中国铁路',
        '电子客票号',
        '国家税务总局',
        '局务税省苏江',  # 江苏省税务局（竖排）
    ]
    
    def is_railway_invoice(self, full_text: str) -> bool:
        """判断是否为铁路电子客票"""
        if not full_text:
            return False
        
        full_text = full_text.replace(' ', '').replace('\n', '')
        
        # 检查关键词
        for keyword in self.RAILWAY_KEYWORDS:
            if keyword in full_text:
                return True
        
        # 检查是否同时包含"电子客票"和车站名称特征
        if '电子客票' in full_text and ('站' in full_text or '站开' in full_text):
            return True
        
        # 检查是否包含铁路特有的格式（如车次格式 K1234, G1234, D1234）
        if re.search(r'[GDCZTKYL][0-9]{1,4}', full_text) and '站' in full_text:
            return True
        
        return False
    
    def parse(self, ocr_result: dict) -> dict:
        """
        解析铁路电子客票
        
        返回字典包含以下字段:
        - invoice_num: 发票号码
        - buyer_name: 购买方名称
        - passenger_name: 乘车人姓名
        - departure: 出发地
        - arrival: 到达地
        - train_number: 车次
        - travel_date: 乘车日期
        - departure_time: 开车时间
        - seat_type: 席别
        - amount: 金额
        - invoice_date: 开票日期
        - ticket_number: 电子客票号
        """
        # 获取完整文本
        words_list = ocr_result.get('words_result', [])
        
        if isinstance(words_list, list) and words_list and isinstance(words_list[0], dict):
            # PaddleOCR格式
            full_text = '\n'.join([w.get('words', '') for w in words_list if isinstance(w, dict)])
        else:
            # 百度API格式
            full_text = ocr_result.get('full_text', '')
        
        result = {
            'invoice_num': self._extract_invoice_number(full_text),
            'buyer_name': self._extract_buyer_name(full_text),
            'passenger_name': self._extract_passenger_name(full_text),
            'departure': self._extract_departure(full_text),
            'arrival': self._extract_arrival(full_text),
            'train_number': self._extract_train_number(full_text),
            'travel_date': self._extract_travel_date(full_text),
            'departure_time': self._extract_departure_time(full_text),
            'seat_type': self._extract_seat_type(full_text),
            'amount': self._extract_amount(full_text),
            'invoice_date': self._extract_invoice_date(full_text),
            'ticket_number': self._extract_ticket_number(full_text),
        }
        
        # 验证关键字段
        result['is_valid'] = self._validate(result)
        result['missing_fields'] = self._get_missing_fields(result)
        
        return result
    
    def _extract_invoice_number(self, text: str) -> str:
        """提取发票号码（20位数字）"""
        patterns = [
            r'发票号码[：:]\s*(\d{20})',
            r'发票号码[:：]\s*(\d{20})',
            r'(\d{20})',  # 直接匹配20位数字
        ]
        return self._regex_extract(text, patterns)
    
    def _extract_buyer_name(self, text: str) -> str:
        """提取购买方名称"""
        patterns = [
            r'购买方名称[：:]\s*([^\n]+)',
            r'名称[：:]\s*([^\n]{5,50})',
        ]
        result = self._regex_extract(text, patterns)
        # 清理并过滤掉销售方名称
        if result and '中国铁路' not in result and '科技' not in result:
            return result.strip()
        
        # 尝试其他方式：查找"统一社会信用代码"前的公司名称
        match = re.search(r'([\u4e00-\u9fa5]{4,30}有限公司)\s*统一社会信用代码', text)
        if match:
            company = match.group(1).strip()
            if '中国铁路' not in company:
                return company
        
        return ''
    
    def _extract_passenger_name(self, text: str) -> str:
        """提取乘车人姓名"""
        # 匹配身份证后的姓名（通常在身份证号后面一行）
        patterns = [
            r'\d{6}\*{4,8}\d{4}\s*([\u4e00-\u9fa5]{2,4})',
            r'\d{17}[\dXx]\s*([\u4e00-\u9fa5]{2,4})',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                name = match.group(1).strip()
                # 过滤掉非人名的词
                if name and len(name) >= 2 and len(name) <= 4:
                    if '公司' not in name and '中国' not in name and '铁路' not in name:
                        return name
        
        # 备选：查找常见的中国人名模式
        lines = text.split('\n')
        for i, line in enumerate(lines):
            if re.search(r'\d{6}\*{4,8}\d{4}', line) or re.search(r'\d{17}[\dXx]', line):
                # 检查下一行
                if i + 1 < len(lines):
                    next_line = lines[i + 1].strip()
                    if re.match(r'^[\u4e00-\u9fa5]{2,4}$', next_line):
                        if next_line not in ['中国铁路', '国家税务总局', '江苏省税务局']:
                            return next_line
        
        return ''
    
    def _extract_departure(self, text: str) -> str:
        """提取出发地"""
        # 首先尝试匹配拼音+站名的组合
        # 模式：英文拼音 + 中文站名
        pattern1 = r'([A-Za-z]+)\s*([\u4e00-\u9fa5]{2,6}站)'
        matches = re.findall(pattern1, text)
        
        if len(matches) >= 2:
            # 通常第一个是出发地，第二个是到达地
            return matches[0][1]
        
        # 备选：查找"站"前的内容
        station_pattern = r'([\u4e00-\u9fa5]{2,6}站)'
        stations = re.findall(station_pattern, text)
        
        if len(stations) >= 2:
            return stations[0]
        
        # 如果只有一个站名，尝试通过上下文判断
        if stations:
            # 检查是否有箭头或方向指示
            if '→' in text or '开' in text:
                return stations[0]
        
        return ''
    
    def _extract_arrival(self, text: str) -> str:
        """提取到达地"""
        # 同样先尝试匹配拼音+站名
        pattern1 = r'([A-Za-z]+)\s*([\u4e00-\u9fa5]{2,6}站)'
        matches = re.findall(pattern1, text)
        
        if len(matches) >= 2:
            return matches[1][1]
        
        # 备选：查找所有站名
        station_pattern = r'([\u4e00-\u9fa5]{2,6}站)'
        stations = re.findall(station_pattern, text)
        
        if len(stations) >= 2:
            return stations[1]
        
        return ''
    
    def _extract_train_number(self, text: str) -> str:
        """提取车次号（如 K8353, G1234, D1234）"""
        pattern = r'([GDCZTKYL])[\s]*([0-9]{1,4})'
        match = re.search(pattern, text)
        if match:
            return match.group(1) + match.group(2)
        return ''
    
    def _extract_travel_date(self, text: str) -> str:
        """提取乘车日期"""
        # 铁路客票通常有多个日期，乘车日期是较早的那个
        # 或者在"开"字前面的日期
        patterns = [
            r'(\d{4})年(\d{1,2})月(\d{1,2})日\s*\d{2}:\d{2}开',
            r'(\d{4})年(\d{1,2})月(\d{1,2})日',
        ]
        
        dates = []
        for pattern in patterns:
            for match in re.finditer(pattern, text):
                year = match.group(1)
                month = match.group(2).zfill(2)
                day = match.group(3).zfill(2)
                dates.append(f"{year}-{month}-{day}")
        
        if dates:
            # 如果有多个日期，通常第一个（较早的）是乘车日期
            # 最后一个通常是开票日期
            return dates[0]
        
        return ''
    
    def _extract_departure_time(self, text: str) -> str:
        """提取开车时间"""
        pattern = r'(\d{2}):(\d{2})开'
        match = re.search(pattern, text)
        if match:
            return f"{match.group(1)}:{match.group(2)}"
        return ''
    
    def _extract_seat_type(self, text: str) -> str:
        """提取席别"""
        seat_patterns = [
            r'(新空调[\u4e00-\u9fa5]+)',
            r'(二等座)',
            r'(一等座)',
            r'(商务座)',
            r'(硬座)',
            r'(软座)',
            r'(硬卧)',
            r'(软卧)',
            r'(无座)',
            r'(高级软卧)',
            r'(动卧)',
            r'(高铁动卧)',
        ]
        
        for pattern in seat_patterns:
            match = re.search(pattern, text)
            if match:
                return match.group(1)
        
        return ''
    
    def _extract_amount(self, text: str) -> float:
        """提取票价金额"""
        # 优先匹配"票价:"后的金额
        patterns = [
            r'票价[：:]\s*[¥￥]\s*(\d+\.?\d*)',
            r'票价\s*[¥￥]\s*(\d+\.?\d*)',
            r'[¥￥]\s*(\d+\.\d{2})',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                try:
                    return float(match.group(1))
                except:
                    pass
        
        return 0.0
    
    def _extract_invoice_date(self, text: str) -> str:
        """提取开票日期"""
        # 查找所有日期，开票日期通常是最后一个
        pattern = r'(\d{4})年(\d{1,2})月(\d{1,2})日'
        dates = []
        
        for match in re.finditer(pattern, text):
            year = match.group(1)
            month = match.group(2).zfill(2)
            day = match.group(3).zfill(2)
            dates.append(f"{year}-{month}-{day}")
        
        if len(dates) >= 2:
            # 如果有多个日期，通常最后一个是开票日期
            return dates[-1]
        elif dates:
            return dates[0]
        
        return ''
    
    def _extract_ticket_number(self, text: str) -> str:
        """提取电子客票号"""
        pattern = r'电子客票号[：:]\s*(\d+)'
        match = re.search(pattern, text)
        if match:
            return match.group(1)
        return ''
    
    def _regex_extract(self, text: str, patterns: List[str]) -> str:
        """使用正则表达式提取"""
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                result = match.group(1).strip()
                if result:
                    return result
        return ''
    
    def _validate(self, result: dict) -> bool:
        """
        验证关键字段是否完整
        关键字段：出发地、到达地、金额
        任意一个缺失就算失败
        """
        required_fields = ['departure', 'arrival', 'amount']
        
        for field in required_fields:
            value = result.get(field)
            if not value or value == '' or value == 0.0:
                return False
        
        return True
    
    def _get_missing_fields(self, result: dict) -> List[str]:
        """获取缺失的关键字段列表"""
        required_fields = {
            'departure': '出发地',
            'arrival': '到达地',
            'amount': '金额'
        }
        
        missing = []
        for field, name in required_fields.items():
            value = result.get(field)
            if not value or value == '' or value == 0.0:
                missing.append(name)
        
        return missing
