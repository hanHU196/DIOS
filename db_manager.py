# db_manager.py
from pymongo import MongoClient
from datetime import datetime
import hashlib
import json
import os

class DatabaseManager:
    """高级数据库管理模块"""
    
    def __init__(self, uri="mongodb://localhost:27017/", db_name="doc_mind"):
        try:
            self.client = MongoClient(uri)
            self.db = self.client[db_name]
            print(f"✅ 已连接到 MongoDB：{db_name}")
            self._ensure_indexes()
        except Exception as e:
            print(f"❌ 连接失败：{e}")
            raise
    
    def _ensure_indexes(self):
        """创建索引（提升查询性能）"""
        # 文档索引
        self.db.documents.create_index([('filename', 1)])
        self.db.documents.create_index([('timestamp', -1)])
        self.db.documents.create_index([('keywords', 1)])
        
        # 提取记录索引
        self.db.extracted_records.create_index([('source_files', 1)])
        self.db.extracted_records.create_index([('timestamp', -1)])
        
        # 填表历史索引
        self.db.fill_history.create_index([('template', 1)])
        self.db.fill_history.create_index([('timestamp', -1)])
        
        print("✅ 索引创建完成")
    
    # ========== 文档管理 ==========
    def save_document(self, filename, file_path, content_preview, keywords=None):
        """保存文档信息"""
        doc_hash = hashlib.md5(open(file_path, 'rb').read()).hexdigest()
        
        doc = {
            'filename': filename,
            'file_path': file_path,
            'hash': doc_hash,
            'size': os.path.getsize(file_path),
            'content_preview': content_preview[:500],
            'keywords': keywords or [],
            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        
        # 避免重复插入（用 hash 去重）
        existing = self.db.documents.find_one({'hash': doc_hash})
        if existing:
            print(f"📄 文档已存在：{filename}")
            return existing['_id']
        
        result = self.db.documents.insert_one(doc)
        print(f"✅ 文档已保存：{filename}")
        return result.inserted_id
    
    def get_document_by_hash(self, doc_hash):
        """根据哈希值获取文档"""
        return self.db.documents.find_one({'hash': doc_hash})
    
    # ========== 提取记录管理 ==========
    def save_extraction(self, source_files, fields, data, command):
        """保存提取记录"""
        record = {
            'source_files': source_files,
            'fields': fields,
            'data': data,
            'data_count': len(data) if isinstance(data, list) else 1,
            'command': command,
            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        
        result = self.db.extracted_records.insert_one(record)
        print(f"✅ 提取记录已保存，ID：{result.inserted_id}")
        return result.inserted_id
    
    def query_extractions(self, field_name=None, field_value=None, limit=50):
        """查询提取记录"""
        query = {}
        if field_name and field_value:
            query[f'data.{field_name}'] = {'$regex': field_value, '$options': 'i'}
        
        results = self.db.extracted_records.find(query).sort('timestamp', -1).limit(limit)
        return [{**r, '_id': str(r['_id'])} for r in results]
    
    # ========== 填表历史管理 ==========
    def save_fill_history(self, template, doc_count, data_count, fields, success=True):
        """保存填表历史"""
        record = {
            'template': template,
            'document_count': doc_count,
            'data_count': data_count,
            'fields': fields,
            'success': success,
            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        
        result = self.db.fill_history.insert_one(record)
        return result.inserted_id
    
    def get_fill_history(self, limit=20):
        """获取填表历史"""
        results = self.db.fill_history.find().sort('timestamp', -1).limit(limit)
        return [{**r, '_id': str(r['_id'])} for r in results]
    
    # ========== 缓存管理（AI 结果缓存）==========
    def get_cached_result(self, doc_path, fields):
        """获取缓存的 AI 提取结果"""
        doc_hash = hashlib.md5(open(doc_path, 'rb').read()).hexdigest()
        cache_key = f"{doc_hash}_{tuple(sorted(fields))}"
        
        cached = self.db.cache.find_one({'key': cache_key})
        if cached:
            print(f"📦 命中缓存：{doc_path}")
            return cached['result']
        return None
    
    def save_cache(self, doc_path, fields, result):
        """保存 AI 结果到缓存"""
        doc_hash = hashlib.md5(open(doc_path, 'rb').read()).hexdigest()
        cache_key = f"{doc_hash}_{tuple(sorted(fields))}"
        
        cache = {
            'key': cache_key,
            'doc_path': doc_path,
            'fields': fields,
            'result': result,
            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        
        # 更新或插入
        self.db.cache.update_one({'key': cache_key}, {'$set': cache}, upsert=True)
        print(f"💾 已缓存：{doc_path}")
    
    # ========== 统计和报告 ==========
    def get_statistics(self):
        """获取系统统计信息"""
        stats = {
            'total_documents': self.db.documents.count_documents({}),
            'total_extractions': self.db.extracted_records.count_documents({}),
            'total_fills': self.db.fill_history.count_documents({}),
            'cache_hits': self.db.cache.count_documents({})
        }
        return stats
    
    # ========== 关闭连接 ==========
    def close(self):
        self.client.close()
        print("🔌 数据库连接已关闭")