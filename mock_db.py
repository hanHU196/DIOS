# mock_db.py
"""
模拟数据库（用于开发和测试）
等真正的数据库好了，直接替换成 db_handler.py
"""
import json
import os
from datetime import datetime

class MockDatabase:
    """模拟数据库（用JSON文件）"""
    
    def __init__(self, db_path="mock_index.json"):
        self.db_path = db_path
        self._init_db()
    
    def _init_db(self):
        """初始化数据库文件"""
        if not os.path.exists(self.db_path):
            with open(self.db_path, 'w', encoding='utf-8') as f:
                json.dump([], f, ensure_ascii=False)
    
    def save_document(self, doc_info):
        """保存文档信息"""
        # 读取现有数据
        with open(self.db_path, 'r', encoding='utf-8') as f:
            docs = json.load(f)
        
        # 去重（如果已存在就更新）
        for i, doc in enumerate(docs):
            if doc['path'] == doc_info['path']:
                docs[i] = doc_info
                break
        else:
            docs.append(doc_info)
        
        # 写回文件
        with open(self.db_path, 'w', encoding='utf-8') as f:
            json.dump(docs, f, ensure_ascii=False, indent=2)
        
        print(f"✅ [模拟DB] 已保存：{doc_info['filename']}")
    
    def find_by_fields(self, fields):
        """根据字段查找文档"""
        try:
            with open(self.db_path, 'r', encoding='utf-8') as f:
                docs = json.load(f)
        except FileNotFoundError:
            return []
        
        results = []
        for doc in docs:
            matched = [f for f in fields if f in doc.get('keywords', [])]
            score = len(matched) * 10
            
            # 文档长度加分
            if doc.get('word_count', 0) > 10000:
                score += 5
            
            results.append({
                'path': doc['path'],
                'filename': doc['filename'],
                'score': score,
                'matched_fields': matched,
                'preview': doc.get('preview', '')
            })
        
        results.sort(key=lambda x: x['score'], reverse=True)
        return results
    
    def get_all_documents(self):
        """获取所有文档"""
        try:
            with open(self.db_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            return []
    
    def clear(self):
        """清空数据库"""
        with open(self.db_path, 'w', encoding='utf-8') as f:
            json.dump([], f, ensure_ascii=False)