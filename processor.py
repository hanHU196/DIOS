# processor.py
# 统一的调用脚本
import os
import logging
import pandas as pd
from document_reader import DocumentReader
from excel_handler import fill_excel_with_data, parse_excel_template
from ai_module import parse_instruction  # 导入指令解析函数
from instruction_parser import InstructionOperator
from ai_module import extract_entities_safe_parallel

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class DocumentProcessor:
    """统一处理器：串联甲乙丁"""

    def __init__(self):
        self.reader = DocumentReader()
        # 从环境变量或直接传入 API Key
        api_key = os.environ.get('DEEPSEEK_API_KEY', "sk-4cbb2ea6e387462383eaeefdbcaa3314")
        self.operator = InstructionOperator(api_key=api_key)
        print("✅ 文档处理器初始化成功")
    
    # ==================== 升级版：动态语言雷达与数据清洗 ====================
    def _clean_extracted_data(self, data_list):
        import re
        
        # 预置双向翻译字典（仅做常见地理名词的兜底，不写死）
        en_to_zh = {
            "asia": "亚洲", "europe": "欧洲", "africa": "非洲", 
            "north america": "北美洲", "south america": "南美洲", 
            "oceania": "大洋洲", "antarctica": "南极洲"
        }
        zh_to_en = {v: k.capitalize() for k, v in en_to_zh.items()}

        # 智能雷达：判断字符串中是否包含中文字符
        def has_chinese(text):
            return bool(re.search(r'[\u4e00-\u9fff]', str(text)))

        for row in data_list:
            for key, value in row.items():
                if isinstance(value, str):
                    clean_val = value.strip()
                    
                    # 1. 刮除冗余前缀 (兼容中英文前缀)
                    prefixes_to_strip = [f"{key}:", f"{key}：", f"{key} ", "约", "大约", "About ", "about "]
                    for prefix in prefixes_to_strip:
                        if clean_val.startswith(prefix):
                            clean_val = clean_val[len(prefix):].strip()
                    # 剔除序号
                    clean_val = re.sub(r'^([a-zA-Z]|\d+)[\.、]\s*', '', clean_val)
                    
                    # 2. 动态语言统一（表头驱动）
                    # 只有在处理带有大洲含义的字段时才介入
                    if "洲" in key or "continent" in key.lower() or "大洲" in key:
                        val_lower = clean_val.lower()
                        # 如果表头包含中文（比如要求填“大洲”），强制转中文
                        if has_chinese(key):
                            for eng, chn in en_to_zh.items():
                                if eng in val_lower:
                                    clean_val = chn
                                    break
                        # 如果表头是纯英文（比如要求填“Continent”），强制转英文
                        else:
                            for chn, eng in zh_to_en.items():
                                if chn in val_lower:
                                    clean_val = eng
                                    break
                                    
                    # 3. 数值转换
                    clean_num_val = clean_val.replace(',', '').strip()
                    try:
                        float_val = float(clean_num_val)
                        if float_val.is_integer():
                            row[key] = int(float_val)
                        else:
                            row[key] = float_val
                    except ValueError:
                        row[key] = clean_val
        return data_list
    # ====================================================================

    # 处理单个文档
    def process_single(self, doc_path, template_path, output_path, instruction=None):
        """处理单个文档"""
        try:
            # 1. 意图解析
            if instruction:
                intent_result = parse_instruction(instruction)
                logger.info(f"📋 指令解析结果：{intent_result}")
                
                if isinstance(intent_result, dict):
                    intent = intent_result.get('intent', '')
                    if intent != 'fill_form':
                        return {'success': False, 'error': f'意图不是填表：{intent_result}'}
                else:
                    if intent_result != 'fill_form':
                        return {'success': False, 'error': f'意图不是填表：{intent_result}'}
            
            # 2. 解析模板字段
            fields = parse_excel_template(template_path)
            logger.info(f"📋 模板字段：{fields}")
            
            # 3. 提取数据（加入 Excel 极速直通车）
            if doc_path.endswith('.xlsx') or doc_path.endswith('.xls'):
                logger.info(f"⚡ 触发极速模式: 发现源文件为 Excel，跳过 AI 直接读取 {os.path.basename(doc_path)}")
                df = pd.read_excel(doc_path)
                data_list = df.to_dict('records')
            else:
                logger.info(f"📖 读取文档：{doc_path}")
                text = self.reader.read(doc_path)
                logger.info(f"✅ 读取成功，长度：{len(text)} 字符")
                
                # 丁提取数据（使用多线程并发提取提速）
                data = extract_entities_safe_parallel(text, fields, max_workers=5)
                data_list = data if isinstance(data, list) else [data]
            
            # 4. 数据类型强制清洗
            data_list = self._clean_extracted_data(data_list)
            
            # 5. 生成Excel
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
   
    # 批量处理多个模板
    def process_batch(self, doc_paths, template_paths, output_dir):
        """批量处理多个模板"""
        results = []
        
        for i, template_path in enumerate(template_paths, 1):
            all_data = []
            
            # 提前在循环外解析模板字段，节省时间
            fields = parse_excel_template(template_path)
            
            for doc_path in doc_paths:
                # 判断文件类型，开启极速直通车
                if doc_path.endswith('.xlsx') or doc_path.endswith('.xls'):
                    # Excel 文件：因为表头一致，直接用 Pandas 读取，速度快百倍
                    logger.info(f"⚡ 直接读取 Excel (极速模式): {os.path.basename(doc_path)}")
                    try:
                        df = pd.read_excel(doc_path)
                        data = df.to_dict('records')
                        all_data.extend(data)
                        logger.info(f"   ✅ 读取到 {len(data)} 行数据")
                    except Exception as e:
                        logger.error(f"   ❌ Excel 读取失败: {e}")
                else:
                    # Word/TXT 等文本文件：用 AI 并发提取
                    logger.info(f"🤖 用 AI 处理文档: {os.path.basename(doc_path)}")
                    text = self.reader.read(doc_path)
                    
                    # 【核心修复】：原代码这里用的是单线程 extract_entities_safe，极慢！
                    # 已替换为 extract_entities_safe_parallel 开启多线程加速
                    data = extract_entities_safe_parallel(text, fields, max_workers=5)
                    
                    if isinstance(data, list):
                        all_data.extend(data)
                    elif isinstance(data, dict):
                        all_data.append(data)
            
            # 【核心修复】：统一进行数值类型清洗，解决部分是文本、部分是数值的问题
            data_list = self._clean_extracted_data(data_list)
            
            # 填表
            output_path = os.path.join(output_dir, f"result_{i}.xlsx")
            fill_excel_with_data(template_path, all_data, output_path)
            results.append({'success': True, 'output': output_path})
        
        return results

    # 智能指令操作文档
    def operate_document(self, file_path, instruction):
        """执行指令操作文档"""
        logger.info(f"📝 operate_document 被调用：{file_path}, {instruction}")
        
        # 1. AI理解指令
        command = parse_instruction(instruction)
        logger.info(f"🤖 AI理解结果：{command}")
        
        if command.get('intent') != 'operate':
            return {'success': False, 'error': '无法理解您的指令'}
        
        # 2. 直接调用 operator.execute，不构建 std_command
        result = self.operator.execute(instruction, file_path)
        logger.info(f"🔧 执行结果：{result}")
        
        return result