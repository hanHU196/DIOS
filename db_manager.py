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
    ENABLE_DB = True  # True: 启用数据库, False: 禁用数据库（测试模式）
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
            # ========== 历史文档集合索引 ==========
            # 普通索引
            self.db.history_documents.create_index([('original_name', 1)])
            self.db.history_documents.create_index([('created_at', -1)])
            self.db.history_documents.create_index([('timestamp', -1)])
            # 全文索引（用于搜索）
            self.db.history_documents.create_index([('original_name', 'text')])
            print("✅ history_documents 索引创建完成")
            
            # ========== 文档集合索引 ==========
            self.db.documents.create_index([('filename', 1)])
            self.db.documents.create_index([('timestamp', -1)])
            self.db.documents.create_index([('hash', 1)])
            print("✅ documents 索引创建完成")
            
            # ========== 提取记录索引 ==========
            self.db.extracted_records.create_index([('source_files', 1)])
            self.db.extracted_records.create_index([('timestamp', -1)])
            print("✅ extracted_records 索引创建完成")
            
            # ========== 填表历史索引 ==========
            self.db.fill_history.create_index([('template', 1)])
            self.db.fill_history.create_index([('timestamp', -1)])
            print("✅ fill_history 索引创建完成")
            
            # ========== 缓存索引 ==========
            self.db.cache.create_index([('key', 1)], unique=True)
            self.db.cache.create_index([('timestamp', -1)])
            print("✅ cache 索引创建完成")
            
            print("🎉 所有数据库索引创建完成")
        
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
    def save_fill_history(self, template, doc_count, data_count, fields, success=True, template_type='excel'):
        """保存填表历史"""
        if not self.enabled:
            print("⚠️ 数据库未启用，跳过保存填表历史")
            return None
        
        try:
            record = {
                'template': template,
                'template_type': template_type,  # 新增：记录模板类型 (excel/word)
                'document_count': doc_count,
                'data_count': data_count,
                'fields': fields,
                'success': success,
                'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            result = self.db.fill_history.insert_one(record)
            print(f"✅ 填表历史已保存: {template}, 数据行数: {data_count}, 类型: {template_type}")
            return result.inserted_id
        except Exception as e:
            print(f"⚠️ 保存填表历史失败: {e}")
            return None

    def get_fill_history(self, limit=20):
        """获取填表历史"""
        if not self.enabled:
            return []
        
        try:
            results = self.db.fill_history.find().sort('timestamp', -1).limit(limit)
            return [{**r, '_id': str(r['_id'])} for r in results]
        except Exception as e:
            print(f"⚠️ 查询填表历史失败: {e}")
            return []
    
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
        
    def check_history_exists(self, original_name, size):
        """检查历史记录中是否已存在相同文件（通过文件名和大小判断）"""
        if not self.enabled:
            return False
        
        try:
            existing = self.db.history_documents.find_one({
                'original_name': original_name,
                'size': size
            })
            return existing is not None
        except Exception as e:
            print(f"⚠️ 检查历史重复失败：{e}")
            return False

    def get_history_by_name_and_size(self, original_name, size):
        """根据文件名和大小获取历史文档"""
        if not self.enabled:
            return None
        
        try:
            return self.db.history_documents.find_one({
                'original_name': original_name,
                'size': size
            })
        except Exception as e:
            print(f"⚠️ 查询历史文档失败：{e}")
            return None
    
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
            # ========== 历史文档管理 ==========

    def save_history_document(self, filename, original_name, size, file_type, content_preview, temp_path):
        """保存文档到历史记录（将文件内容存储为 base64）"""
        if not self.enabled:
            return None
        
        try:
            import base64
            
            # 读取文件内容并转为 base64
            with open(temp_path, 'rb') as f:
                file_content = base64.b64encode(f.read()).decode('utf-8')
            
            doc = {
                'original_name': original_name,
                'filename': filename,
                'size': size,
                'type': file_type,
                'content_preview': content_preview[:300],
                'file_content': file_content,  # base64 编码的文件内容
                'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'created_at': datetime.now()
            }
            
            result = self.db.history_documents.insert_one(doc)
            return result.inserted_id
            
        except Exception as e:
            print(f"⚠️ 保存历史文档失败：{e}")
            return None

    def get_history_documents(self, limit=100, skip=0):
        """获取历史文档列表"""
        if not self.enabled:
            return []
        
        try:
            results = self.db.history_documents.find()\
                .sort('created_at', -1)\
                .skip(skip)\
                .limit(limit)
            
            return [{
                '_id': str(r['_id']),
                'name': r['original_name'],
                'size': r['size'],
                'type': r['type'],
                'content_preview': r.get('content_preview', ''),
                'timestamp': r['timestamp']
            } for r in results]
            
        except Exception as e:
            print(f"⚠️ 获取历史文档失败：{e}")
            return []

    def get_history_document_by_id(self, doc_id):
        """根据 ID 获取历史文档"""
        if not self.enabled:
            return None
        
        try:
            from bson import ObjectId
            result = self.db.history_documents.find_one({'_id': ObjectId(doc_id)})
            if result:
                return {
                    '_id': str(result['_id']),
                    'original_name': result['original_name'],
                    'filename': result['filename'],
                    'size': result['size'],
                    'type': result['type'],
                    'file_content': result['file_content'],
                    'content_preview': result.get('content_preview', ''),
                    'timestamp': result['timestamp']
                }
            return None
            
        except Exception as e:
            print(f"⚠️ 获取历史文档失败：{e}")
            return None

    def delete_history_documents(self, doc_ids):
        """删除历史文档"""
        if not self.enabled:
            return 0
        
        try:
            from bson import ObjectId
            object_ids = [ObjectId(doc_id) for doc_id in doc_ids]
            result = self.db.history_documents.delete_many({'_id': {'$in': object_ids}})
            return result.deleted_count
            
        except Exception as e:
            print(f"⚠️ 删除历史文档失败：{e}")
            return 0

    def clear_all_history(self):
        """清空所有历史文档"""
        if not self.enabled:
            return
        
        try:
            self.db.history_documents.delete_many({})
            print("🧹 已清空所有历史文档")
        except Exception as e:
            print(f"⚠️ 清空历史失败：{e}")

    def search_history_documents(self, keyword):
        """搜索历史文档"""
        if not self.enabled:
            return []
        
        try:
            results = self.db.history_documents.find({
                'original_name': {'$regex': keyword, '$options': 'i'}
            }).sort('created_at', -1).limit(50)
            
            return [{
                '_id': str(r['_id']),
                'name': r['original_name'],
                'size': r['size'],
                'type': r['type'],
                'content_preview': r.get('content_preview', ''),
                'timestamp': r['timestamp']
            } for r in results]
            
        except Exception as e:
            print(f"⚠️ 搜索历史文档失败：{e}")
            return []