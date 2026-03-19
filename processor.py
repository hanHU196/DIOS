# processor.py
# 统一的调用脚本
import os
import logging
from document_reader import DocumentReader
from excel_handler import fill_excel_with_data, parse_excel_template
from ai_module import extract_entities

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class DocumentProcessor:
    """统一处理器：串联甲乙丁"""
    
    def __init__(self):
        self.reader = DocumentReader()
        print("✅ 文档处理器初始化成功")
    
    def process_single(self, doc_path, template_path, output_path, instruction=None):
        """处理单个文档"""
        try:
            # 1. 甲读文档
            logger.info(f"📖 读取文档：{doc_path}")
            text = self.reader.read(doc_path)
            logger.info(f"✅ 读取成功，长度：{len(text)} 字符")
            
            # 2. 如果有指令，解析意图
            # 在 process_single 方法中，修改意图判断部分
            if instruction:
                from ai_module import parse_instruction
                intent_result = parse_instruction(instruction)
                logger.info(f"📋 指令解析结果：{intent_result}")
                
                # 如果是字典，检查 intent 字段
                if isinstance(intent_result, dict):
                    intent = intent_result.get('intent', '')
                    if intent != 'fill_form':
                        return {'success': False, 'error': f'意图不是填表：{intent_result}'}
                else:
                    # 如果是字符串，直接比较
                    if intent_result != 'fill_form':
                        return {'success': False, 'error': f'意图不是填表：{intent_result}'}
            
            # 3. 乙解析模板字段
            from excel_handler import parse_excel_template
            fields = parse_excel_template(template_path)
            logger.info(f"📋 模板字段：{fields}")
            
          
            # 4. 丁提取数据
            from ai_module import extract_entities
            data = extract_entities(text, fields)
            logger.info(f"🔍 AI返回数据类型：{type(data)}")
            if isinstance(data, list):
                logger.info(f"📊 提取到 {len(data)} 条数据")
                if len(data) > 0:
                    logger.info(f"第一条数据预览：{data[0]}")
            else:
                logger.info(f"⚠️ AI返回的不是列表，而是：{type(data)}")
                logger.info(f"数据内容：{data}")
            
            # 5. 准备数据
            if isinstance(data, list):
                data_list = data
            else:
                data_list = [data]
            
            # 6. 乙生成Excel
            from excel_handler import fill_excel_with_data
            fill_excel_with_data(template_path, data_list, output_path)
            logger.info(f"✅ Excel生成成功：{output_path}")
            
            return {
                'success': True,
                'output': output_path,
                'fields': fields,
                'data_count': len(data_list)
            }
            
        except Exception as e:
            logger.error(f"处理失败：{e}")
            return {'success': False, 'error': str(e)}