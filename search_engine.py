# search_engine.py
# 文档索引和匹配引擎
import os
import re
from collections import Counter
from document_reader import DocumentReader
from excel_handler import parse_excel_template
from excel_handler import MongoDBHandler

# 专业词库
ECONOMIC_TERMS = [
    'GDP', '国内生产总值', '人均GDP', '国民总收入', 'GNI',
    'CPI', '居民消费价格', 'PPI', '工业生产者出厂价格',
    'PMI', '采购经理人指数', '财政收入', '财政支出', '税收',
    '固定资产投资', '房地产投资', '基础设施投资',
    '社会消费品零售总额', '网上零售额', '进出口', '出口', '进口',
    '第一产业', '第二产业', '第三产业', '增加值', '工业增加值',
    '粮食产量', '猪肉产量', '牛肉产量', '羊肉产量', '禽肉产量',
    '人口', '常住人口', '城镇化率', '就业人员', '失业率',
    '人均可支配收入', '人均消费支出', '恩格尔系数',
    'R&D', '研发经费', '发明专利', '技术合同'
]

URBAN_TERMS = [
    'GDP总量', '人均GDP', '常住人口', '户籍人口',
    '一般公共预算收入', '税收收入', '财政支出',
    '规上工业增加值', '固定资产投资', '社会消费品零售总额',
    '进出口总额', '实际使用外资', '金融机构本外币存款',
    '城镇居民人均可支配收入', '农村居民人均可支配收入',
    '商品房销售面积', '商品房销售额', '居民消费价格',
    '城市名', '城市排名', '百强城市'
]

class DocumentMatcher:
    def __init__(self):
        self.reader = DocumentReader()
        self.db = MongoDBHandler(db_name="document_system")
        print("✅ 文档匹配器初始化成功（使用MongoDB）")
    
    def extract_keywords(self, text):
        """综合提取关键词"""
        keywords = set()
        
        # 1. 专业词库匹配
        all_terms = ECONOMIC_TERMS + URBAN_TERMS
        for term in all_terms:
            if term in text:
                keywords.add(term)
        
        # 2. 数值指标提取
        patterns = [
            r'([\u4e00-\u9fa5]{2,10})[为是](\d+\.?\d*)[万亿千万亿]?[元人吨亩]',
            r'([\u4e00-\u9fa5]{2,10})增长?([\+\-]?\d+\.?\d*)%',
        ]
        
        for pattern in patterns:
            matches = re.findall(pattern, text)
            for match in matches:
                if match and match[0]:
                    keywords.add(match[0])
        
        # 3. 高频词统计
        words = re.findall(r'[\u4e00-\u9fa5]{2,}', text)
        word_freq = Counter(words)
        
        # 取高频词，但过滤常见词
        common_words = {'中国', '人民', '全国', '全年', '以上', '以下', '其中', '合计', '根据'}
        for word, freq in word_freq.most_common(20):
            if word not in common_words and len(word) >= 2:
                keywords.add(word)
        
        return list(keywords)
    
    def index_documents(self, doc_paths):
        """为文档建立索引"""
        for path in doc_paths:
            try:
                text = self.reader.read(path)
                keywords = self.extract_keywords(text)
                
                doc_info = {
                    'path': path,
                    'filename': os.path.basename(path),
                    'keywords': keywords,
                    'keyword_count': len(keywords),
                    'word_count': len(text),
                    'preview': text[:300],
                    'type': 'document'  # ✅ 必须有
                }
                
                # 插入数据库
                self.db.insert_data(doc_info)  # ✅ 用 insert_data
                print(f"✅ 已索引：{os.path.basename(path)} ({len(keywords)}个关键词)")
                print(f"   关键词示例：{keywords[:10]}")
                
            except Exception as e:
                print(f"❌ 索引失败 {path}: {e}")
    
    def match_template(self, template_path):
        """为模板找到最匹配的文档"""
        template_keywords = parse_excel_template(template_path)
        print(f"📋 模板字段：{template_keywords}")
        
        all_docs = self.db.query_data('document', {})
        
        if not all_docs:
            print("⚠️ 文档库为空")
            return None
        
        matches = []
        for doc in all_docs:
            matched = [k for k in template_keywords if k in doc.get('keywords', [])]
            score = len(matched) * 10
            
            if doc.get('word_count', 0) > 10000:
                score += 5
            
            matches.append({
                'path': doc['path'],
                'filename': doc['filename'],
                'score': score,
                'matched_keywords': matched,
                'match_count': len(matched),
                'preview': doc.get('preview', '')
            })
        
        matches.sort(key=lambda x: x['score'], reverse=True)
        
        if matches:
            best = matches[0]
            print(f"✅ 最佳匹配：{best['filename']} (匹配 {best['match_count']} 个字段)")
            print(f"   匹配字段：{best['matched_keywords']}")
            return best
        else:
            print("⚠️ 没有找到匹配的文档")
            return None
        
    def clear_index(self):
        """清空索引"""
        self.db.clear_collection('document')
        print("🧹 索引已清空")