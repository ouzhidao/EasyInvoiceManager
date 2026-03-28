"""
数据解析模块
"""
import re
from utils.helpers import extract_short_name, parse_amount, standardize_date


# 商品/服务内容关键词映射（用于概括）
CONTENT_KEYWORDS = {
    # 餐饮类
    '餐费': ['餐费', '餐饮', '餐', '饭菜', '酒水', '食品'],
    '住宿费': ['住宿', '酒店', '宾馆', '旅店', '客房'],
    '交通费': ['打车', '出租车', '网约车', '快车', '专车', '滴滴'],
    '加油费': ['汽油', '柴油', '成品油', '加油'],
    '停车费': ['停车', '停车费'],
    '过路费': ['高速', '公路', '桥梁', '通行费', '过路费'],
    '机票': ['机票', '航空', '飞机票'],
    '火车票': ['火车', '铁路', '高铁', '动车'],
    # 办公类
    '办公用品': ['办公', '文具', '纸', '笔', '打印', '复印'],
    '电子设备': ['电脑', '手机', '平板', '笔记本', '显示器', '打印机'],
    '软件服务': ['软件', '系统', '服务', '订阅', 'license'],
    # 通讯类
    '通讯费': ['通讯', '话费', '宽带', '网络', '电信', '移动', '联通'],
    '快递费': ['快递', '物流', '运费', '邮寄'],
    # 日常类
    '日用品': ['日用', '清洁', '洗护', '用品'],
    '图书资料': ['图书', '资料', '书籍', '教材', '刊物'],
    '会议费': ['会议', '会场', '会务'],
    '培训费': ['培训', '教育', '课程', '学习'],
    '咨询费': ['咨询', '顾问', '服务'],
    '广告费': ['广告', '宣传', '推广', '设计'],
    '维修费': ['维修', '修理', '保养', '维护'],
    '租赁费': ['租赁', '租用', '租金', '房租'],
    '水电费': ['水费', '电费', '水电', '燃气'],
    '物业费': ['物业', '管理'],
    # 其他
    '材料费': ['材料', '原料', '原材料', '辅料'],
    '加工费': ['加工', '制作', '生产'],
    '检测费': ['检测', '检验', '测试', '认证'],
}


class DataParser:
    INVOICE_TYPE_MAP = {
        '电子普通发票': '电子普通发票',
        '增值税电子普通发票': '电子普通发票',
        '电子专用发票': '电子专用发票',
        '数电票': '数电票',
        '全电发票': '数电票',
        '普通发票': '纸质普通发票',
        '专用发票': '纸质专用发票',
    }
    
    def parse(self, ocr_result: dict):
        # 确保 ocr_result 不为 None
        if ocr_result is None:
            ocr_result = {}
        
        source = ocr_result.get('source', '')
        
        if source == 'baidu_vat':
            # 增值税发票专用识别结果
            return self._parse_vat(ocr_result)
        else:
            # 通用识别结果
            return self._parse_general(ocr_result)
    
    def _parse_vat(self, ocr_result: dict):
        """解析增值税发票识别结果"""
        # 确保 ocr_result 不为 None
        if ocr_result is None:
            ocr_result = {}
        
        words_result = ocr_result.get('words_result', {})
        if words_result is None:
            words_result = {}
        
        invoice_data = {
            'invoice_num': self._extract(words_result, ['InvoiceNum', 'invoice_num', '发票号码']),
            'date': self._extract(words_result, ['InvoiceDate', 'invoice_date', '开票日期']),
            'amount': self._extract(words_result, ['TotalAmount', 'total_amount', '价税合计']),
            'seller_name': self._extract(words_result, ['SellerName', 'seller_name', '销售方名称']),
            'buyer_name': self._extract(words_result, ['BuyerName', 'buyer_name', '购买方名称', 'PurchaserName']),
            'invoice_type': self._extract(words_result, ['InvoiceType', 'invoice_type', '发票种类']),
            'content': self._extract_content(words_result),
        }
        
        return self._format_result(invoice_data)
    
    def _parse_general(self, ocr_result: dict):
        """解析通用文字识别结果（用正则提取）"""
        # 确保 ocr_result 不为 None
        if ocr_result is None:
            ocr_result = {}
        
        words_list = ocr_result.get('words_result', [])
        if words_list is None:
            words_list = []
        
        full_text = '\n'.join([w.get('words', '') for w in words_list if w and isinstance(w, dict)])
        
        # 用正则提取关键信息
        invoice_data = {
            'invoice_num': self._regex_extract(full_text, [
                r'发票号码[：:]?\s*(\d{8,20})',
                r'No[.:]?\s*(\d{8,20})',
                r'(\d{20})',  # 20位发票号码
            ]),
            'date': self._regex_extract(full_text, [
                r'(\d{4}[年/-]\d{1,2}[月/-]\d{1,2}[日]?)',
                r'开票日期[：:]?\s*(\d{4}[年/-]\d{1,2}[月/-]\d{1,2}[日]?)',
            ]),
            'amount': self._regex_extract(full_text, [
                r'价税合计[（(]小写[)）][：:]?\s*[¥￥]?(\d[\d,]+\.?\d{0,2})',
                r'[¥￥](\d[\d,]+\.?\d{0,2})',
                r'合\s*计[：:]?\s*[¥￥]?(\d[\d,]+\.?\d{0,2})',
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
        
        result = self._format_result(invoice_data)
        result['recognize_method'] = 'general'  # 标记为通用识别
        result['raw_text'] = full_text[:200]  # 保留部分原文用于调试
        
        return result
    
    def _format_result(self, invoice_data: dict):
        """格式化最终结果"""
        # 确保 invoice_data 不为 None
        if invoice_data is None:
            invoice_data = {}
        
        # 获取各个字段，确保不为 None
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
        
        # 确保金额是数字
        try:
            amount_int = int(parsed_amount)
        except (ValueError, TypeError):
            amount_int = 0
        
        result = {
            'invoice_num_full': invoice_num,
            'invoice_num_short': invoice_num[-5:] if len(invoice_num) >= 5 else invoice_num,
            'date': standardize_date(date),
            'amount': parsed_amount,
            'amount_int': amount_int,
            'seller_name': seller_name,
            'seller_short': extract_short_name(seller_name),
            'buyer_name': buyer_name,
            'invoice_type': self.INVOICE_TYPE_MAP.get(invoice_type, invoice_type or '其他发票'),
            'content': content,
            'is_valid': all([
                invoice_num,
                date or (parsed_amount > 0)
            ])
        }
        
        return result
    
    def _extract(self, data: dict, keys: list):
        """从字典中提取字段"""
        # 确保 data 不为 None
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
        # 确保 text 不为 None
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
        # 确保 text 不为 None
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
        """从增值税发票结果中提取商品/服务内容（优先返回星号之间的原始内容）"""
        # 确保 words_result 不为 None
        if words_result is None:
            words_result = {}
        
        # 获取完整文本以查找星号之间的内容
        full_text = ''
        if 'words_result' in words_result:
            for item in words_result.get('words_result', []):
                if isinstance(item, dict):
                    full_text += item.get('words', '') + '\n'
        
        # 1. 优先匹配星号之间的内容（支持单星号*内容* 或 双星号**内容**）
        if full_text:
            # 先尝试双星号 **内容**
            star_pattern = r'\*\*([^\*]+)\*\*'
            try:
                matches = re.findall(star_pattern, full_text)
                if matches:
                    content = matches[0].strip()
                    if len(content) >= 1:
                        return content
            except re.error:
                pass
            
            # 再尝试单星号 *内容*（非贪婪匹配，取最短匹配）
            single_star_pattern = r'\*([^\*]+)\*'
            try:
                matches = re.findall(single_star_pattern, full_text)
                if matches:
                    # 取第一个匹配
                    content = matches[0].strip()
                    if len(content) >= 1:
                        return content
            except re.error:
                pass
        
        # 2. 尝试直接提取 CommodityName
        content = ''
        if 'CommodityName' in words_result:
            commodity = words_result['CommodityName']
            if isinstance(commodity, dict):
                content = commodity.get('word', '') or ''
            elif isinstance(commodity, list):
                # 多行商品，取第一个
                items = [c.get('word', '') if isinstance(c, dict) else str(c) for c in commodity if c]
                content = items[0] if items else ''
            else:
                content = str(commodity) if commodity is not None else ''
        
        # 2.1 如果从 CommodityName 获取到内容，尝试从中提取星号之间的内容
        if content:
            # 先尝试双星号 **内容**
            star_pattern = r'\*\*([^\*]+)\*\*'
            try:
                matches = re.findall(star_pattern, content)
                if matches:
                    matched_content = matches[0].strip()
                    if len(matched_content) >= 1:
                        return matched_content
            except re.error:
                pass
            
            # 再尝试单星号 *内容*
            single_star_pattern = r'\*([^\*]+)\*'
            try:
                matches = re.findall(single_star_pattern, content)
                if matches:
                    matched_content = matches[0].strip()
                    if len(matched_content) >= 1:
                        return matched_content
            except re.error:
                pass
        
        # 3. 如果没有，尝试其他字段
        if not content:
            content = self._extract(words_result, [
                'ServiceName', 'ItemName', 'ProjectName', 
                'GoodsName', 'Details', 'Content'
            ])
        
        return self._summarize_content(content)
    
    def _extract_content_from_text(self, full_text: str) -> str:
        """从通用识别文本中提取商品/服务内容 - 优先提取星号间的内容"""
        # 确保 full_text 不为 None
        if full_text is None:
            full_text = ''
        
        # 1. 优先匹配双星号之间的内容 **内容**
        star_pattern = r'\*\*([^\*]+)\*\*'
        try:
            matches = re.findall(star_pattern, full_text)
            if matches:
                content = matches[0].strip()
                if len(content) >= 1:
                    return content
        except re.error:
            pass
        
        # 2. 匹配单星号之间的内容 *内容*（如 *纺织产品*）
        single_star_pattern = r'\*([^\*]+)\*'
        try:
            matches = re.findall(single_star_pattern, full_text)
            if matches:
                # 取第一个匹配
                content = matches[0].strip()
                if len(content) >= 1:
                    return content
        except re.error:
            pass
        
        # 3. 尝试匹配商品/服务相关字段
        patterns = [
            r'货物或应税劳务名称[\s:：]*([^\n]{1,20})',
            r'项目名称[\s:：]*([^\n]{1,20})',
            r'商品名称[\s:：]*([^\n]{1,20})',
            r'服务内容[\s:：]*([^\n]{1,20})',
            r'(?:^|\n)([^\n]{2,10}(?:费|服务|商品|材料|设备))',
        ]
        
        for pattern in patterns:
            try:
                match = re.search(pattern, full_text, re.IGNORECASE)
                if match:
                    content = match.group(1).strip() if match.group(1) else ''
                    # 清理数字和符号
                    content = re.sub(r'^[\d\s\*\-\.]+', '', content)
                    if len(content) >= 2:
                        return self._summarize_content(content)
            except re.error:
                continue
        
        # 如果没有匹配到，尝试从文本中提取关键词
        return self._summarize_content(full_text)
    
    def _summarize_content(self, content: str) -> str:
        """概括内容，返回不超过10个汉字的简短描述"""
        if not content:
            return ''
        
        content = content.strip()
        
        # 1. 直接匹配关键词映射
        for category, keywords in CONTENT_KEYWORDS.items():
            for keyword in keywords:
                if keyword in content:
                    return category
        
        # 2. 清理内容（去除数字、符号等）
        cleaned = re.sub(r'[\d\*\s\.\-\/\\]', '', content)
        
        # 3. 如果内容本身就短（<=10字），直接返回
        if len(cleaned) <= 10 and len(cleaned) >= 2:
            return cleaned
        
        # 4. 提取核心词汇（去除通用后缀）
        common_suffixes = ['费', '服务', '商品', '用品', '设备', '材料', '费', '及', '等']
        for suffix in common_suffixes:
            if cleaned.endswith(suffix) and len(cleaned) > 2:
                cleaned = cleaned[:-1]
        
        # 5. 截断到10个字符
        if len(cleaned) > 10:
            cleaned = cleaned[:10]
        
        return cleaned if cleaned else '其他'
