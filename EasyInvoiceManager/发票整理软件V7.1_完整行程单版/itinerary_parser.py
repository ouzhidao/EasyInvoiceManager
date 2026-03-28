"""
行程单解析模块
支持：滴滴出行行程单、航空运输电子客票行程单、其他交通行程单
"""
import re
from typing import List, Dict, Optional
from datetime import datetime


class ItineraryParser:
    """行程单解析器"""
    
    # 行程类型关键字映射
    TYPE_KEYWORDS = {
        '飞机': ['航空运输电子客票行程单', '电子客票', '航班', '登机牌', '机票'],
        '汽车': ['滴滴出行', 'DIDI TRAVEL', '网约车', '快车', '专车', '出租车', '顺风车'],
        '火车': ['铁路电子客票', '火车票', '高铁', '动车', '列车'],
        '轮渡': ['船票', '轮渡', '客船', '渡轮'],
        '其他': ['行程单', '出行', '交通']
    }
    
    def __init__(self):
        pass
    
    def is_itinerary(self, full_text: str) -> bool:
        """判断是否为行程单"""
        if not full_text:
            return False
        
        full_text = full_text.replace(' ', '').replace('\n', '')
        
        # 检查是否包含任何行程单关键字
        for type_name, keywords in self.TYPE_KEYWORDS.items():
            for keyword in keywords:
                if keyword in full_text:
                    return True
        
        return False
    
    def detect_type(self, full_text: str) -> str:
        """检测行程单类型"""
        if not full_text:
            return '其他'
        
        full_text = full_text.replace(' ', '').replace('\n', '')
        
        # 按优先级检测类型
        for type_name in ['飞机', '汽车', '火车', '轮渡']:
            for keyword in self.TYPE_KEYWORDS[type_name]:
                if keyword in full_text:
                    return type_name
        
        return '其他'
    
    def parse(self, ocr_result: dict) -> List[Dict]:
        """
        解析行程单
        返回：行程明细列表（滴滴多段行程会展开为多条）
        """
        # 获取完整文本
        full_text = self._get_full_text(ocr_result)
        
        # 检测类型
        itinerary_type = self.detect_type(full_text)
        
        # 根据类型调用对应解析器
        if itinerary_type == '飞机':
            return self._parse_flight(ocr_result, full_text)
        elif itinerary_type == '汽车':
            return self._parse_didi(ocr_result, full_text)
        elif itinerary_type == '火车':
            return self._parse_train(ocr_result, full_text)
        else:
            return self._parse_generic(ocr_result, full_text)
    
    def _get_full_text(self, ocr_result: dict) -> str:
        """从OCR结果中提取完整文本"""
        if 'full_text' in ocr_result:
            return ocr_result['full_text']
        
        words_list = ocr_result.get('words_result', [])
        
        # PaddleOCR格式
        if isinstance(words_list, list) and words_list and isinstance(words_list[0], dict):
            return '\n'.join([w.get('words', '') for w in words_list if isinstance(w, dict)])
        
        # 百度API格式
        if isinstance(words_list, dict):
            texts = []
            for key, value in words_list.items():
                if isinstance(value, dict) and 'word' in value:
                    texts.append(value['word'])
            return '\n'.join(texts)
        
        return ''
    
    def _parse_flight(self, ocr_result: dict, full_text: str) -> List[Dict]:
        """解析航空运输电子客票行程单 - 简化版"""
        results = []
        
        # 提取基本信息
        passenger_name = self._extract_passenger_name(full_text)
        invoice_date = self._extract_flight_invoice_date(full_text)
        
        # 提取金额
        amount = self._extract_flight_total(full_text)
        if amount == 0:
            amount = self._extract_flight_fare(full_text)
        
        # 只返回一条汇总记录
        result = {
            'invoice_type': '行程单',
            'sub_type': '飞机',
            'passenger_name': passenger_name,
            'date': invoice_date,
            'time': '',
            'departure': '',
            'arrival': '',
            'amount': amount,
            'vehicle_info': '',
            'is_valid': amount > 0,
            'missing_fields': [],
            'is_itinerary': True,
            'segment_index': 1,
            'segment_total': 1
        }
        results.append(result)
        
        return results
    
    def _parse_didi(self, ocr_result: dict, full_text: str) -> List[Dict]:
        """解析滴滴出行行程单 - 只返回汇总记录"""
        results = []
        
        # 提取信息
        passenger_phone = self._extract_didi_phone(full_text)
        report_date = self._extract_didi_report_date(full_text)
        
        # 提取合计金额
        total_amount = self._extract_didi_total(full_text)
        if total_amount == 0:
            # 尝试从文本中提取任意金额
            amounts = re.findall(r'(\d+\.\d{2})', full_text)
            if amounts:
                total_amount = float(amounts[-1])
        
        # 只返回一条汇总记录
        result = {
            'invoice_type': '行程单',
            'sub_type': '汽车',
            'passenger_name': passenger_phone or '',
            'date': report_date or '',
            'time': '',
            'departure': '',
            'arrival': '',
            'amount': total_amount,
            'vehicle_info': '',
            'is_valid': total_amount > 0,
            'missing_fields': [],
            'is_itinerary': True,
            'segment_index': 1,
            'segment_total': 1
        }
        results.append(result)
        
        return results
    
    def _parse_train(self, ocr_result: dict, full_text: str) -> List[Dict]:
        """解析火车票/铁路电子客票"""
        results = []
        
        # 复用现有railway_parser的逻辑，但转换为行程单格式
        # 这里简化处理
        passenger_name = self._extract_train_passenger(full_text)
        train_number = self._extract_train_number(full_text)
        departure = self._extract_train_departure(full_text)
        arrival = self._extract_train_arrival(full_text)
        travel_date = self._extract_train_date(full_text)
        amount = self._extract_train_amount(full_text)
        
        result = {
            'invoice_type': '行程单',
            'sub_type': '火车',
            'passenger_name': passenger_name,
            'date': travel_date,
            'time': '',
            'departure': departure,
            'arrival': arrival,
            'amount': amount,
            'vehicle_info': train_number,
            'seat_class': '',
            'is_valid': bool(departure and arrival and train_number),
            'missing_fields': [],
            'is_itinerary': True,
            'segment_index': 1,
            'segment_total': 1
        }
        
        if not result['is_valid']:
            missing = []
            if not departure:
                missing.append('出发站')
            if not arrival:
                missing.append('到达站')
            if not train_number:
                missing.append('车次')
            result['missing_fields'] = missing
        
        results.append(result)
        return results
    
    def _parse_generic(self, ocr_result: dict, full_text: str) -> List[Dict]:
        """通用行程单解析"""
        results = []
        
        # 尝试提取日期
        date = self._extract_generic_date(full_text)
        # 尝试提取金额
        amount = self._extract_generic_amount(full_text)
        # 尝试提取地点
        departure, arrival = self._extract_generic_locations(full_text)
        
        result = {
            'invoice_type': '行程单',
            'sub_type': '其他',
            'passenger_name': '',
            'date': date,
            'time': '',
            'departure': departure,
            'arrival': arrival,
            'amount': amount,
            'vehicle_info': '',
            'is_valid': False,
            'missing_fields': ['类型识别失败'],
            'is_itinerary': True,
            'segment_index': 1,
            'segment_total': 1
        }
        
        results.append(result)
        return results
    
    # ==================== 飞机行程单提取方法 ====================
    
    def _extract_passenger_name(self, text: str) -> str:
        """提取旅客姓名"""
        patterns = [
            r'旅客姓名[\s:]?(\S{2,10})',
            r'姓名[\s:]?(\S{2,10})',
            r'乘机人[\s:]?(\S{2,10})',
        ]
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                return match.group(1).strip()
        return ''
    
    def _extract_flight_invoice_date(self, text: str) -> str:
        """提取填开日期"""
        patterns = [
            r'填开日期[\s:]?(\d{4}[年/-]\d{1,2}[月/-]\d{1,2}[日]?|[年月日\d]+)',
            r'日期[\s:]?(\d{4}[年/-]\d{1,2}[月/-]\d{1,2})',
        ]
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                return self._standardize_date(match.group(1))
        return ''
    
    def _extract_flight_segments(self, text: str) -> List[Dict]:
        """提取航段信息（支持多航段）"""
        flights = []
        
        # 查找航班号模式
        flight_pattern = r'([A-Z]{2}\d{3,4})'
        flight_matches = list(re.finditer(flight_pattern, text))
        
        for match in flight_matches:
            flight_num = match.group(1)
            # 在航班号附近查找出发地、目的地、日期
            start_pos = max(0, match.start() - 200)
            end_pos = min(len(text), match.end() + 200)
            context = text[start_pos:end_pos]
            
            flight = {
                'flight_number': flight_num,
                'departure': self._extract_airport(context, 'departure'),
                'arrival': self._extract_airport(context, 'arrival'),
                'date': self._extract_flight_date(context),
                'time': self._extract_flight_time(context),
                'fare': self._extract_flight_fare(context),
                'seat_class': self._extract_seat_class(context)
            }
            
            # 只添加有效航段
            if flight['departure'] or flight['arrival']:
                flights.append(flight)
        
        return flights
    
    def _extract_airport(self, text: str, type_: str) -> str:
        """提取机场/城市"""
        # 简化处理，实际可使用机场三字码映射
        city_pattern = r'(北京|上海|广州|深圳|成都|杭州|武汉|西安|重庆|青岛|大连|厦门|昆明|天津|南京|郑州|长沙|沈阳|乌鲁木齐|济南|宁波|南宁|合肥|海口|南昌|贵阳|福州|长春|兰州|太原|呼和浩特|银川|西宁|拉萨|三亚|珠海|烟台|泉州|无锡|佛山|东莞|石家庄|哈尔滨|[\u4e00-\u9fa5]{2,6}(?:机场|航站))'
        matches = re.findall(city_pattern, text)
        if matches:
            if type_ == 'departure' and len(matches) >= 1:
                return matches[0]
            elif type_ == 'arrival' and len(matches) >= 2:
                return matches[1]
        return ''
    
    def _extract_flight_date(self, text: str) -> str:
        """提取航班日期"""
        patterns = [
            r'(\d{4}[年/-]\d{1,2}[月/-]\d{1,2})',
            r'(\d{2}[年/-]\d{1,2}[月/-]\d{1,2})',
        ]
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                return self._standardize_date(match.group(1))
        return ''
    
    def _extract_flight_time(self, text: str) -> str:
        """提取航班时间"""
        patterns = [
            r'(\d{1,2}:\d{2})',
            r'(\d{1,2}时\d{1,2}分)',
        ]
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                return match.group(1)
        return ''
    
    def _extract_flight_fare(self, text: str) -> float:
        """提取票价"""
        patterns = [
            r'票价[\s:]*(\d+\.?\d*)',
            r'CNY[\s:]*(\d+\.?\d*)',
            r'¥[\s:]*(\d+\.?\d*)',
        ]
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                try:
                    return float(match.group(1))
                except:
                    return 0
        return 0
    
    def _extract_flight_total(self, text: str) -> float:
        """提取合计金额"""
        patterns = [
            r'合计[\s:]*(\d+\.?\d*)',
            r'总额[\s:]*(\d+\.?\d*)',
            r'总计[\s:]*(\d+\.?\d*)',
        ]
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                try:
                    return float(match.group(1))
                except:
                    return 0
        return 0
    
    def _extract_flight_number(self, text: str) -> str:
        """提取航班号"""
        match = re.search(r'([A-Z]{2}\d{3,4})', text)
        if match:
            return match.group(1)
        return ''
    
    def _extract_seat_class(self, text: str) -> str:
        """提取舱位等级"""
        patterns = [
            r'舱位[\s:]*(\S)',
            r'([YBMFHAECWDSPG])\s*舱',
        ]
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                return match.group(1)
        return ''
    
    def _extract_flight_departure(self, text: str) -> str:
        """提取出发地（整体文本）"""
        return self._extract_airport(text, 'departure')
    
    def _extract_flight_arrival(self, text: str) -> str:
        """提取目的地（整体文本）"""
        return self._extract_airport(text, 'arrival')
    
    # ==================== 滴滴行程单提取方法 ====================
    
    def _extract_didi_phone(self, text: str) -> str:
        """提取行程人手机号"""
        patterns = [
            r'行程人手机号[\s:]?(1\d{10})',
            r'手机[\s:]?(1\d{10})',
            r'(1\d{10})',
        ]
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                return match.group(1)
        return ''
    
    def _extract_didi_report_date(self, text: str) -> str:
        """提取申请日期"""
        patterns = [
            r'申请日期[\s:]?(\d{4}[年/-]\d{1,2}[月/-]\d{1,2})',
        ]
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                return self._standardize_date(match.group(1))
        return ''
    
    def _extract_didi_date_range(self, text: str) -> str:
        """提取行程日期范围"""
        patterns = [
            r'行程起止日期[\s:]?(\d{4}[年/-]\d{1,2}[月/-]\d{1,2}\s*[至到-]\s*\d{4}[年/-]\d{1,2}[月/-]\d{1,2})',
        ]
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                return match.group(1)
        return ''
    
    def _extract_didi_trips(self, text: str) -> List[Dict]:
        """提取滴滴行程明细 - 增强版"""
        trips = []
        
        # 清理文本
        text = text.replace(' ', '').replace('  ', ' ')
        lines = text.split('\n')
        
        # 方法1: 尝试匹配完整行程行
        # 模式：序号 日期 时间 类型 起点 终点 里程 金额
        for i, line in enumerate(lines):
            line = line.strip()
            if not line:
                continue
            
            # 尝试匹配日期时间格式
            # 支持格式：11-26 11:25 或 2025-11-26 或 11月26日
            date_patterns = [
                r'(\d{1,2}[月/-]\d{1,2})',
                r'(\d{4}[年/-]\d{1,2}[年/-]\d{1,2})',
            ]
            time_pattern = r'(\d{1,2}:\d{2})'
            amount_pattern = r'(\d+\.?\d*)'
            
            date_match = None
            for dp in date_patterns:
                m = re.search(dp, line)
                if m:
                    date_match = m
                    break
            
            time_match = re.search(time_pattern, line)
            
            # 查找金额（通常是行尾的数字）
            amounts = re.findall(r'(\d+\.\d{2})', line)
            amount = float(amounts[-1]) if amounts else 0
            
            # 如果找到日期和金额，尝试提取其他信息
            if date_match and amount > 0:
                trip_type = self._extract_trip_type(line)
                distance = self._extract_distance(line)
                departure, arrival = self._extract_trip_locations(line, lines, i)
                
                # 只要有起点或终点就认为是有效行程
                if departure or arrival:
                    trip = {
                        'date': date_match.group(1),
                        'time': time_match.group(1) if time_match else '',
                        'type': trip_type,
                        'departure': departure or '未知',
                        'arrival': arrival or '未知',
                        'distance': distance,
                        'amount': amount
                    }
                    trips.append(trip)
        
        # 方法2: 如果没找到明细，尝试提取合计金额作为整体
        if not trips:
            total_amount = self._extract_didi_total(text)
            if total_amount > 0:
                # 获取手机号作为乘客名
                phone = self._extract_didi_phone(text)
                report_date = self._extract_didi_report_date(text)
                
                # 创建一个整体行程记录
                trips.append({
                    'date': report_date or '',
                    'time': '',
                    'type': '滴滴行程',
                    'departure': '见明细',
                    'arrival': '见明细',
                    'distance': '',
                    'amount': total_amount
                })
        
        return trips
    
    def _extract_trip_type(self, line: str) -> str:
        """提取行程类型（快车/专车等）"""
        types = ['快车', '专车', '出租车', '顺风车', '豪华车', '拼车']
        for t in types:
            if t in line:
                return t
        return ''
    
    def _extract_distance(self, line: str) -> str:
        """提取里程"""
        match = re.search(r'(\d+\.?\d*)\s*公里|km', line)
        if match:
            return match.group(1)
        return ''
    
    def _extract_trip_locations(self, line: str, lines: list, idx: int) -> tuple:
        """提取起点和终点"""
        # 简化处理：在整行中查找地点
        # 实际可能需要更复杂的逻辑
        
        # 尝试在当前行和前后行中查找地点分隔符
        full_context = line
        if idx + 1 < len(lines):
            full_context += ' ' + lines[idx + 1]
        
        # 查找常见地点关键词
        location_pattern = r'([\u4e00-\u9fa5]{2,10}(?:站|机场|酒店|小区|大厦|广场|路|街|区))'
        locations = re.findall(location_pattern, full_context)
        
        if len(locations) >= 2:
            return locations[0], locations[1]
        elif len(locations) == 1:
            return locations[0], ''
        
        return '', ''
    
    def _extract_didi_total(self, text: str) -> float:
        """提取滴滴合计金额"""
        patterns = [
            r'合计\s*(\d+\.?\d*)\s*元',
            r'共\s*(\d+)\s*笔行程[\s:]*[¥￥]?(\d+\.?\d*)',
            r'合计行程.*?(\d+\.?\d*)\s*元',
        ]
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                try:
                    return float(match.group(1) if match.lastindex == 1 else match.group(2))
                except:
                    return 0
        return 0
    
    def _complete_date(self, date_str: str, ref_date: str) -> str:
        """补全年份"""
        if not date_str:
            return ''
        
        # 从参考日期提取年份
        year = ''
        if ref_date and len(ref_date) >= 4:
            year = ref_date[:4]
        else:
            year = str(datetime.now().year)
        
        # 格式化日期
        date_str = date_str.replace('月', '-').replace('/', '-')
        parts = date_str.split('-')
        
        if len(parts) == 2:
            month, day = parts
            return f"{year}-{month.zfill(2)}-{day.zfill(2)}"
        
        return date_str
    
    # ==================== 火车票提取方法 ====================
    
    def _extract_train_passenger(self, text: str) -> str:
        """提取乘车人"""
        patterns = [
            r'乘车人[\s:]?(\S{2,10})',
            r'旅客[\s:]?(\S{2,10})',
        ]
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                return match.group(1)
        return ''
    
    def _extract_train_number(self, text: str) -> str:
        """提取车次"""
        patterns = [
            r'([GDCZTKYL]\d{1,4})',
            r'车次[\s:]?([A-Z]?\d{1,4})',
        ]
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                return match.group(1)
        return ''
    
    def _extract_train_departure(self, text: str) -> str:
        """提取出发站"""
        patterns = [
            r'([\u4e00-\u9fa5]{2,6}站)',
            r'出发[站\s:]?([\u4e00-\u9fa5]{2,6})',
        ]
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                return match.group(1).replace('站', '')
        return ''
    
    def _extract_train_arrival(self, text: str) -> str:
        """提取到达站"""
        patterns = [
            r'到达[站\s:]?([\u4e00-\u9fa5]{2,6})',
            r'至[\s:]?([\u4e00-\u9fa5]{2,6})',
        ]
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                return match.group(1).replace('站', '')
        return ''
    
    def _extract_train_date(self, text: str) -> str:
        """提取乘车日期"""
        patterns = [
            r'(\d{4}[年/-]\d{1,2}[月/-]\d{1,2})',
            r'(\d{2}[年/-]\d{1,2}[月/-]\d{1,2})',
        ]
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                return self._standardize_date(match.group(1))
        return ''
    
    def _extract_train_amount(self, text: str) -> float:
        """提取票价"""
        patterns = [
            r'票价[\s:]*(\d+\.?\d*)',
            r'¥[\s:]*(\d+\.?\d*)',
            r'(\d+\.\d{2})',
        ]
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                try:
                    return float(match.group(1))
                except:
                    return 0
        return 0
    
    # ==================== 通用提取方法 ====================
    
    def _extract_generic_date(self, text: str) -> str:
        """通用日期提取"""
        patterns = [
            r'(\d{4}[年/-]\d{1,2}[月/-]\d{1,2})',
            r'(\d{2}[年/-]\d{1,2}[月/-]\d{1,2})',
        ]
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                return self._standardize_date(match.group(1))
        return ''
    
    def _extract_generic_amount(self, text: str) -> float:
        """通用金额提取"""
        patterns = [
            r'金额[\s:]*(\d+\.?\d*)',
            r'合计[\s:]*(\d+\.?\d*)',
            r'¥[\s:]*(\d+\.?\d*)',
            r'(\d+\.\d{2})',
        ]
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                try:
                    return float(match.group(1))
                except:
                    return 0
        return 0
    
    def _extract_generic_locations(self, text: str) -> tuple:
        """通用地点提取"""
        location_pattern = r'([\u4e00-\u9fa5]{2,6}(?:站|机场|港口|码头|城市))'
        locations = re.findall(location_pattern, text)
        
        if len(locations) >= 2:
            return locations[0], locations[1]
        elif len(locations) == 1:
            return locations[0], ''
        
        return '', ''
    
    def _standardize_date(self, date_str: str) -> str:
        """标准化日期格式"""
        if not date_str:
            return ''
        
        # 清理字符串
        date_str = date_str.replace('年', '-').replace('月', '-').replace('日', '').replace('/', '-')
        
        # 分割
        parts = date_str.split('-')
        if len(parts) == 3:
            year, month, day = parts
            # 处理2位年份
            if len(year) == 2:
                year_int = int(year)
                if year_int >= 50:
                    year = '19' + year
                else:
                    year = '20' + year
            return f"{year}-{month.zfill(2)}-{day.zfill(2)}"
        
        return date_str
