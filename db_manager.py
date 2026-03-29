# db_manager.py
from pymongo import MongoClient
from datetime import datetime
import hashlib
import os

class DatabaseManager:
    """
    数据库管理模块
    设置 ENABLE_DB = True 启用数据库，False 禁用数据库
    """
    
    # ========== 在这里切换数据库开关 ==========
    ENABLE_DB = False   # True: 启用数据库, False: 禁用数据库（测试模式）
    # ========================================
    
    def __init__(self, uri="mongodb://localhost:27017/", db_name="doc_mind"):
        self.enabled = self.ENABLE_DB
        
        if not self.enabled:
            print("⚠️ 数据库已禁用（测试模式），所有数据将不会保存")
            self.client = None
            self.db = None
            return
        
        try:
            self.client = MongoClient(uri, serverSelectionTimeoutMS=5000)
            # 测试连接
            self.client.admin.command('ping')
            self.db = self.client[db_name]
            print(f"✅ 已连接到 MongoDB：{db_name}")
            self._ensure_indexes()
        except Exception as e:
            print(f"⚠️ MongoDB 连接失败：{e}")
            print(f"   数据库功能将不可用，但不影响核心填表功能")
            self.enabled = False
            self.client = None
            self.db = None
    
    def _ensure_indexes(self):
        """创建索引（提升查询性能）"""
        if not self.enabled:
            return
        
        try:
            # 文档索引
            self.db.documents.create_index([('filename', 1)])
            self.db.documents.create_index([('timestamp', -1)])
            self.db.documents.create_index([('hash', 1)])
            
            # 提取记录索引
            self.db.extracted_records.create_index([('source_files', 1)])
            self.db.extracted_records.create_index([('timestamp', -1)])
            
            # 填表历史索引
            self.db.fill_history.create_index([('template', 1)])
            self.db.fill_history.create_index([('timestamp', -1)])
            
            # 缓存索引
            self.db.cache.create_index([('key', 1)], unique=True)
            self.db.cache.create_index([('timestamp', -1)])
            
            print("✅ 数据库索引创建完成")
        except Exception as e:
            print(f"⚠️ 索引创建失败：{e}")
    
    # ========== 文档管理 ==========
    def save_document(self, filename, file_path, content_preview, keywords=None):
        """保存文档信息"""
        if not self.enabled:
            return None
        
        try:
            # 计算文件哈希
            with open(file_path, 'rb') as f:
                doc_hash = hashlib.md5(f.read()).hexdigest()
            
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
                return existing['_id']
            
            result = self.db.documents.insert_one(doc)
            return result.inserted_id
        except Exception as e:
            print(f"⚠️ 保存文档失败：{e}")
            return None
    
    def get_document_by_hash(self, doc_hash):
        """根据哈希值获取文档"""
        if not self.enabled:
            return None
        try:
            return self.db.documents.find_one({'hash': doc_hash})
        except:
            return None
    
    # ========== 提取记录管理 ==========
    def save_extraction(self, source_files, fields, data, command):
        """保存提取记录"""
        if not self.enabled:
            return None
        
        try:
            record = {
                'source_files': source_files,
                'fields': fields,
                'data': data,
                'data_count': len(data) if isinstance(data, list) else 1,
                'command': command,
                'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            result = self.db.extracted_records.insert_one(record)
            return result.inserted_id
        except Exception as e:
            print(f"⚠️ 保存提取记录失败：{e}")
            return None
    
    def query_extractions(self, field_name=None, field_value=None, limit=50):
        """查询提取记录"""
        if not self.enabled:
            return []
        
        try:
            query = {}
            if field_name and field_value:
                query[f'data.{field_name}'] = {'$regex': field_value, '$options': 'i'}
            
            results = self.db.extracted_records.find(query).sort('timestamp', -1).limit(limit)
            return [{**r, '_id': str(r['_id'])} for r in results]
        except Exception as e:
            print(f"⚠️ 查询提取记录失败：{e}")
            return []
    
    # ========== 填表历史管理 ==========
    def save_fill_history(self, template, doc_count, data_count, fields, success=True):
        """保存填表历史"""
        if not self.enabled:
            return None
        
        try:
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
        except Exception as e:
            print(f"⚠️ 保存填表历史失败：{e}")
            return None
    
    def get_fill_history(self, limit=20):
        """获取填表历史"""
        if not self.enabled:
            return []
        
        try:
            results = self.db.fill_history.find().sort('timestamp', -1).limit(limit)
            return [{**r, '_id': str(r['_id'])} for r in results]
        except Exception as e:
            print(f"⚠️ 查询填表历史失败：{e}")
            return []
    
    # ========== 缓存管理 ==========
    def get_cached_result(self, doc_path, fields):
        """获取缓存的 AI 提取结果"""
        if not self.enabled:
            return None
        
        try:
            # 计算文档哈希
            with open(doc_path, 'rb') as f:
                doc_hash = hashlib.md5(f.read()).hexdigest()
            cache_key = f"{doc_hash}_{tuple(sorted(fields))}"
            
            cached = self.db.cache.find_one({'key': cache_key})
            if cached:
                return cached['result']
            return None
        except Exception as e:
            return None
    
    def save_cache(self, doc_path, fields, result):
        """保存 AI 结果到缓存"""
        if not self.enabled:
            return
        
        try:
            # 计算文档哈希
            with open(doc_path, 'rb') as f:
                doc_hash = hashlib.md5(f.read()).hexdigest()
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
        except Exception as e:
            pass  # 缓存失败不影响主流程
    
    # ========== 统计信息 ==========
    def get_statistics(self):
        """获取系统统计信息"""
        if not self.enabled:
            return {
                'total_documents': 0,
                'total_extractions': 0,
                'total_fills': 0,
                'cache_hits': 0
            }
        
        try:
            stats = {
                'total_documents': self.db.documents.count_documents({}),
                'total_extractions': self.db.extracted_records.count_documents({}),
                'total_fills': self.db.fill_history.count_documents({}),
                'cache_hits': self.db.cache.count_documents({})
            }
            return stats
        except Exception as e:
            return {
                'total_documents': 0,
                'total_extractions': 0,
                'total_fills': 0,
                'cache_hits': 0
            }
    
    # ========== 清理功能 ==========
    def clear_all(self):
        """清空所有数据（慎用）"""
        if not self.enabled:
            return
        
        try:
            self.db.documents.delete_many({})
            self.db.extracted_records.delete_many({})
            self.db.fill_history.delete_many({})
            self.db.cache.delete_many({})
            print("🧹 已清空所有数据库记录")
        except Exception as e:
            print(f"⚠️ 清空失败：{e}")
    
    # ========== 关闭连接 ==========
    def close(self):
        """关闭数据库连接"""
        if self.enabled and self.client:
            self.client.close()
            print("🔌 数据库连接已关闭")