import os
import re
import logging
from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename
from processor import DocumentProcessor
import pandas as pd
import numpy as np
import ai_module
from document_reader import DocumentReader
from db_manager import DatabaseManager
from concurrent.futures import ThreadPoolExecutor, as_completed

# 初始化处理器
processor = DocumentProcessor()
reader = DocumentReader()
app = Flask(__name__)  
db = DatabaseManager()

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# 配置
UPLOAD_FOLDER = 'uploads'
ALLOWED_DOC_EXTENSIONS = {'txt', 'docx', 'md', 'xlsx'}
ALLOWED_TEMPLATE_EXTENSIONS = {'xlsx', 'xls','docx'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename, allowed_set=ALLOWED_DOC_EXTENSIONS):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_set

def parse_fields_from_instruction(instruction):
    """从指令中解析字段列表"""
    import re
    
    # 匹配 "提取甲方、乙方、金额"
    match = re.search(r'提取[：:]\s*(.+)', instruction)
    if not match:
        match = re.search(r'提取\s*(.+)', instruction)
    
    if match:
        fields_str = match.group(1)
        # 按中文顿号、逗号、空格分割
        fields = re.split(r'[、，,\s]+', fields_str)
        return [f.strip() for f in fields if f.strip()]
    
    # 如果没有"提取"关键词，尝试从指令中找中文词
    words = re.split(r'[、，,\s]+', instruction)
    return [w for w in words if len(w) >= 2 and not re.search(r'[a-zA-Z0-9]', w)]

def process_word_document(doc_path, fields):
    """处理单个 Word 文档（用于并行）"""
    try:
        text = reader.read(doc_path)
        
        # 检查缓存
        cached = db.get_cached_result(doc_path, fields)
        if cached:
            return doc_path, cached, True
        
        # AI 提取
        extracted = ai_module.extract_entities_safe(text, fields)
        db.save_cache(doc_path, fields, extracted)
        return doc_path, extracted, False
        
    except Exception as e:
        logger.error(f"处理文档 {doc_path} 失败: {e}")
        return doc_path, [], False

# ---------- 导入队友模块 ----------
try:
    from document_reader import DocumentReader
    logger.info("成功导入 document_reader.DocumentReader")
    
    reader = DocumentReader()
    
    def read_documents(file_paths):
        all_text = ""
        for path in file_paths:
            logger.info(f"📖 读取文件：{path}")
            try:
                text = reader.read(path)
                logger.info(f"  读取成功，长度：{len(text)} 字符")
                all_text += text + "\n"
            except Exception as e:
                logger.error(f"读取失败：{e}")
                all_text += f"[读取失败：{path}]\n"
        return all_text
        
except ImportError as e:
    logger.error(f"导入 document_reader 失败：{e}")
    def read_documents(file_paths):
        return "模拟文档内容"

# 表格生成模块
try:
    import excel_handler
    logger.info("✅ 成功导入 excel_handler")
except ImportError as e:
    logger.error(f"❌ 导入 excel_handler 失败：{e}")
    class MockExcelHandler:
        @staticmethod
        def fill_excel_with_data(template_path, data_list, output_path):
            df = pd.DataFrame(data_list)
            df.to_excel(output_path, index=False)
            return output_path
        @staticmethod
        def parse_excel_template(template_path):
            return ["甲方", "乙方", "金额"]
        @staticmethod
        def MongoDBHandler():
            class MockMongo:
                def insert_data(self, *args, **kwargs):
                    return None
                def query_data(self, *args, **kwargs):
                    return []
            return MockMongo()
    excel_handler = MockExcelHandler()

# AI模块
try:
    from ai_module import parse_instruction, extract_entities
    logger.info("成功导入 ai_module")
except ImportError:
    logger.warning("导入 ai_module 失败，使用模拟函数")
    def parse_instruction(instruction):
        if '填表' in instruction:
            return 'fill_form'
        return 'unknown'
    def extract_entities(text, targets):
        return {t: f'示例{t}' for t in targets}

# ---------------------------------------------------------

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    doc_file = request.files.get('document')
    template_file = request.files.get('template')
    instruction = request.form.get('command', '').strip()

    if not doc_file or not template_file or not instruction:
        return "请上传文档和模板文件，并输入指令", 400

    if not allowed_file(doc_file.filename, ALLOWED_DOC_EXTENSIONS):
        return "文档格式不支持", 400
    if not allowed_file(template_file.filename, ALLOWED_TEMPLATE_EXTENSIONS):
        return "模板格式不支持", 400

    doc_filename = secure_filename(doc_file.filename)
    doc_path = os.path.join(app.config['UPLOAD_FOLDER'], doc_filename)
    doc_file.save(doc_path)

    template_filename = secure_filename(template_file.filename)
    if not template_filename.endswith('.xlsx'):
        template_filename = template_filename + '.xlsx'
    template_path = os.path.join(app.config['UPLOAD_FOLDER'], 'template_' + template_filename)
    template_file.save(template_path)

    output_dir = "outputs"
    os.makedirs(output_dir, exist_ok=True)

    from datetime import datetime
    output_filename = f"result_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    output_path = os.path.join(output_dir, output_filename)

    try:
        result = processor.process_single(
            doc_path=doc_path,
            template_path=template_path,
            output_path=output_path,
            instruction=instruction
        )
        
        if result['success']:
            return send_file(output_path, as_attachment=True, download_name='result.xlsx')
        else:
            return f"处理失败：{result.get('error', '未知错误')}", 500
    except Exception as e:
        logger.error(f"处理异常：{e}")
        return f"处理异常：{str(e)}", 500

@app.route('/batch_process', methods=['POST'])
def batch_process():
    from datetime import datetime
    
    doc_files = request.files.getlist('documents')
    template_files = request.files.getlist('templates')
    
    if not doc_files or not template_files:
        return jsonify({'error': '请上传文档和模板文件'}), 400
    
    doc_paths = []
    for doc_file in doc_files:
        doc_filename = secure_filename(doc_file.filename)
        doc_path = os.path.join(app.config['UPLOAD_FOLDER'], doc_filename)
        doc_file.save(doc_path)
        doc_paths.append(doc_path)
    
    template_paths = []
    for template_file in template_files:
        template_filename = secure_filename(template_file.filename)
        if not template_filename.endswith('.xlsx'):
            template_filename = template_filename + '.xlsx'
        template_path = os.path.join(app.config['UPLOAD_FOLDER'], 'template_' + template_filename)
        template_file.save(template_path)
        template_paths.append(template_path)
    
    batch_dir = f"batch_output_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    output_dir = os.path.join('outputs', batch_dir)
    os.makedirs(output_dir, exist_ok=True)
    
    results = processor.process_batch(doc_paths, template_paths, output_dir)
    
    success_count = sum(1 for r in results if r.get('success', False))
    
    response = {
        'success': True,
        'batch_dir': batch_dir,
        'stats': {
            'total': len(results),
            'success': success_count,
            'failed': len(results) - success_count
        },
        'results': []
    }
    
    for r in results:
        result_item = {
            'template_name': os.path.basename(r['template']),
            'success': r['success']
        }
        if r['success']:
            result_item['data_count'] = r.get('data_count', 0)
            result_item['output_file'] = os.path.basename(r['output_path'])
        else:
            result_item['error'] = r.get('error', '未知错误')
        response['results'].append(result_item)
    
    return jsonify(response)

@app.route('/operate', methods=['POST'])
def operate_document():
    file = request.files.get('file')
    instruction = request.form.get('instruction', '').strip()
    
    if not file or not instruction:
        return jsonify({'error': '请上传文件并输入指令'}), 400
    
    filename = secure_filename(file.filename)
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(file_path)
    
    result = processor.operate_document(file_path, instruction)
    
    if result['success']:
        return send_file(file_path, as_attachment=True, download_name=f'operated_{filename}')
    else:
        return jsonify({'error': result['error']}), 400

@app.route('/api/extract', methods=['POST'])
def api_extract():
    """智能文档提取（增强版）"""
    try:
        files = request.files.getlist('documents')
        command = request.form.get('command', '').strip()
        
        if not files or not command:
            return jsonify({'error': '请上传文档并输入指令'}), 400
        
        doc_paths = []
        all_text = ""
        source_files = []
        
        for file in files:
            filename = secure_filename(file.filename)
            path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(path)
            doc_paths.append(path)
            source_files.append(filename)
            
            # 保存文档信息到数据库
            text = reader.read(path)
            db.save_document(filename, path, text[:500])
            all_text += f"\n--- {filename} ---\n{text}\n"
        
        # 解析字段
        fields = parse_fields_from_instruction(command)
        if not fields:
            return jsonify({'error': '无法从指令中识别字段'}), 400
        
        # 检查缓存
        cached_result = None
        for doc_path in doc_paths:
            cached = db.get_cached_result(doc_path, fields)
            if cached:
                cached_result = cached
                break
        
        if cached_result:
            extracted = cached_result
            logger.info("📦 使用缓存结果")
        else:
            # AI 提取
            extracted = ai_module.extract_entities(all_text, fields)
            # 保存缓存
            for doc_path in doc_paths:
                db.save_cache(doc_path, fields, extracted)
        
        if isinstance(extracted, dict):
            extracted = [extracted]
        
        # 保存提取记录
        db.save_extraction(source_files, fields, extracted, command)
        
        return jsonify({
            'success': True,
            'fields': fields,
            'data': extracted,
            'from_cache': cached_result is not None
        })
        
    except Exception as e:
        logger.error(f"提取失败: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/fill', methods=['POST'])
def api_fill():
    """表格智能填写（支持并行处理）"""
    try:
        doc_files = request.files.getlist('documents')
        template_file = request.files.get('template')
        command = request.form.get('command', '').strip()
        
        if not doc_files or not template_file or not command:
            return jsonify({'error': '请上传文档、模板并输入指令'}), 400
        
        # 保存文件
        doc_paths = []
        for file in doc_files:
            filename = secure_filename(file.filename)
            path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(path)
            doc_paths.append(path)
            logger.info(f"📄 文档已保存：{filename}")
        
        template_filename = secure_filename(template_file.filename)
        template_path = os.path.join(app.config['UPLOAD_FOLDER'], 'template_' + template_filename)
        template_file.save(template_path)
        logger.info(f"📊 模板已保存：{template_filename}")
        
        # 判断模板类型
        is_word_template = template_filename.endswith('.docx')
        is_excel_template = template_filename.endswith('.xlsx') or template_filename.endswith('.xls')
        
        if not is_word_template and not is_excel_template:
            return jsonify({'error': '模板格式不支持'}), 400
        
        # ========== 特殊场景：单个 Excel 数据文件 + Word 模板 ==========
        is_excel_to_word = (len(doc_paths) == 1 and
                            doc_paths[0].endswith(('.xlsx', '.xls')) and
                            is_word_template)   # 注意这里直接用 is_word_template
        
        if is_excel_to_word:
            output_filename = f"{os.path.splitext(os.path.basename(doc_paths[0]))[0]}_填表.docx"
            output_path = os.path.join('outputs', output_filename)
            os.makedirs('outputs', exist_ok=True)
            fill_word_from_excel(template_path, doc_paths[0], output_path)
            return send_file(output_path, as_attachment=True, download_name=output_filename)
        
        
        # 解析模板字段
        if is_excel_template:
            from excel_handler import parse_excel_template
            fields = parse_excel_template(template_path)
        else:
            fields = parse_word_template(template_path)
        logger.info(f"📋 模板字段：{fields}")
        
        # 筛选条件
        filters = []
        try:
            from ai_module import parse_filter_conditions
            filter_result = parse_filter_conditions(command)
            filters = filter_result.get('filters', [])
        except:
            pass
        
        # ========== 并行处理 Word 文档 ==========
        all_data = []
        word_docs = []
        excel_data = []
        
        for doc_path in doc_paths:
            if doc_path.endswith('.xlsx') or doc_path.endswith('.xls'):
                df = pd.read_excel(doc_path)
                if filters:
                    df = apply_filters_to_df(df, filters)
                excel_data.extend(df.to_dict('records'))
            else:
                word_docs.append(doc_path)
        
        # 并行处理 Word 文档
        if word_docs:
            logger.info(f"🚀 并行处理 {len(word_docs)} 个 Word 文档")
            
            with ThreadPoolExecutor(max_workers=4) as executor:
                futures = {executor.submit(process_word_document, doc, fields): doc 
                           for doc in word_docs}
                
                for future in as_completed(futures):
                    doc_path, extracted, from_cache = future.result()
                    logger.info(f"   {'📦 缓存' if from_cache else '🤖 AI'} 完成: {os.path.basename(doc_path)}")
                    
                    if isinstance(extracted, list):
                        all_data.extend(extracted)
                    elif isinstance(extracted, dict):
                        all_data.append(extracted)
        
        # 添加 Excel 数据
        all_data.extend(excel_data)
        
        logger.info(f"📊 合并后共 {len(all_data)} 行数据")
        
        # ========== 生成输出文件 ==========
        # 用第一个文档名命名
        first_doc_name = os.path.splitext(os.path.basename(doc_paths[0]))[0]

        if is_excel_template:
            output_filename = f"{first_doc_name}_填表.xlsx"
            output_path = os.path.join('outputs', output_filename)
            os.makedirs('outputs', exist_ok=True)
            from excel_handler import fill_excel_with_data
            fill_excel_with_data(template_path, all_data, output_path)
        else:
            output_filename = f"{first_doc_name}_填表.docx"
            output_path = os.path.join('outputs', output_filename)
            os.makedirs('outputs', exist_ok=True)
            fill_word_with_data(template_path, all_data, output_path)
        
        # 保存历史
        try:
            db.save_fill_history(
                template=template_filename,
                doc_count=len(doc_paths),
                data_count=len(all_data),
                fields=fields,
                success=True
            )
        except:
            pass
        
        # 返回文件
        return send_file(output_path, as_attachment=True, download_name=output_filename)
        
    except Exception as e:
        logger.error(f"填表失败: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500
            
@app.route('/api/stats', methods=['GET'])
def get_stats():
    """获取系统统计信息"""
    stats = db.get_statistics()
    return jsonify(stats)


@app.route('/api/history', methods=['GET'])
def get_history():
    """获取填表历史"""
    limit = request.args.get('limit', 20, type=int)
    history = db.get_fill_history(limit)
    return jsonify({'success': True, 'history': history})


@app.route('/api/search', methods=['POST'])
def search_data():
    """跨文档搜索"""
    data = request.get_json()
    field = data.get('field')
    value = data.get('value')
    
    if not field or not value:
        return jsonify({'error': '请提供字段和值'}), 400
    
    results = db.query_extractions(field, value)
    return jsonify({
        'success': True,
        'count': len(results),
        'results': results
    })

def parse_word_template(template_path):
    """解析 Word 模板中的占位符"""
    from docx import Document
    import re
    doc = Document(template_path)
    fields = set()
    
    for para in doc.paragraphs:
        matches = re.findall(r'\{\{(.*?)\}\}', para.text)
        for match in matches:
            fields.add(match.strip())
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                matches = re.findall(r'\{\{(.*?)\}\}', cell.text)
                for match in matches:
                    fields.add(match.strip())
    
    return list(fields)

def fill_word_with_data(template_path, data_records, output_path):
    """填充 Word 模板"""
    from docx import Document
    doc = Document(template_path)
    
    # 替换占位符
    for record in data_records:
        for key, value in record.items():
            placeholder = f'{{{{{key}}}}}'
            for para in doc.paragraphs:
                if placeholder in para.text:
                    para.text = para.text.replace(placeholder, str(value))
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        if placeholder in cell.text:
                            cell.text = cell.text.replace(placeholder, str(value))
    
    doc.save(output_path)
    logger.info(f"✅ Word 文件已生成：{output_path}")


def apply_filters_to_df(df, filters):
    """应用筛选条件到 DataFrame"""
    filtered_df = df.copy()
    
    for f in filters:
        filter_type = f.get('type')
        column = f.get('column')
        
        if column not in filtered_df.columns:
            logger.warning(f"列 '{column}' 不存在，跳过筛选")
            continue
        
        if filter_type == 'date_range':
            # 日期范围筛选
            filtered_df[column] = pd.to_datetime(filtered_df[column], errors='coerce')
            start = f.get('start')
            end = f.get('end')
            mask = (filtered_df[column] >= start) & (filtered_df[column] <= end)
            filtered_df = filtered_df[mask]
        
        elif filter_type == 'numeric':
            # 数值比较筛选
            operator = f.get('operator')
            value = f.get('value')
            
            if operator == '>':
                filtered_df = filtered_df[filtered_df[column] > value]
            elif operator == '<':
                filtered_df = filtered_df[filtered_df[column] < value]
            elif operator == '>=':
                filtered_df = filtered_df[filtered_df[column] >= value]
            elif operator == '<=':
                filtered_df = filtered_df[filtered_df[column] <= value]
            elif operator == '==':
                filtered_df = filtered_df[filtered_df[column] == value]
    
    return filtered_df


def apply_filters_to_data(data_list, filters):
    """应用筛选条件到字典列表（AI提取的数据）"""
    filtered = []
    for item in data_list:
        match = True
        for f in filters:
            filter_type = f.get('type')
            column = f.get('column')
            value = item.get(column)
            
            if value is None:
                match = False
                break
            
            if filter_type == 'numeric':
                operator = f.get('operator')
                target = f.get('value')
                try:
                    num_value = float(value) if isinstance(value, str) else value
                    if operator == '>' and not (num_value > target):
                        match = False
                    elif operator == '<' and not (num_value < target):
                        match = False
                    elif operator == '>=' and not (num_value >= target):
                        match = False
                    elif operator == '<=' and not (num_value <= target):
                        match = False
                except:
                    match = False
            elif filter_type == 'date_range':
                # 日期范围筛选（简化处理）
                pass
        
        if match:
            filtered.append(item)
    
    return filtered

@app.route('/api/format', methods=['POST'])
def api_format():
    try:
        file = request.files.get('document')
        command = request.form.get('command', '').strip()
        
        if not file or not command:
            return jsonify({'error': '请上传文档并输入指令'}), 400
        
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        
        result = processor.operate_document(file_path, command)
        
        if result.get('success'):
            return send_file(file_path, as_attachment=True, download_name=f'formatted_{filename}')
        else:
            return jsonify({'error': result.get('error', '格式调整失败')}), 400
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500


def fill_word_from_excel(template_path, excel_path, output_path):
    """动态版：根据 Word 表格前的段落文本提取城市名，根据表头自动匹配 Excel 列名"""
    from docx import Document
    import pandas as pd
    import re

    # 1. 读取 Excel 数据
    df = pd.read_excel(excel_path)

    # 2. 打开 Word 模板
    doc = Document(template_path)
    tables = doc.tables

    # 3. 识别每个表格对应的城市（从段落文本中提取城市名）
    city_table_map = {}
    table_index = 0
    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            # 尝试从段落中提取城市名（匹配“德州市”、“潍坊市”、“临沂市”等）
            # 可根据实际城市列表调整正则
            match = re.search(r'(德州市|潍坊市|临沂市)', text)
            if match:
                city_name = match.group(1)
                # 检查后面是否有表格
                next_elem = para._element.getnext()
                while next_elem is not None:
                    if next_elem.tag.endswith('tbl'):
                        if table_index < len(tables):
                            city_table_map[city_name] = tables[table_index]
                            table_index += 1
                        break
                    next_elem = next_elem.getnext()
        if table_index >= len(tables):
            break

    # 降级方案：如果未识别到，使用顺序匹配（城市列表可根据 Excel 中的实际城市动态获取）
    if not city_table_map:
        # 从 Excel 中获取所有不重复的城市名（按顺序）
        cities = df['城市'].dropna().unique().tolist()
        for i, city in enumerate(cities):
            if i < len(tables):
                city_table_map[city] = tables[i]

    # 4. 获取 Excel 列名
    excel_columns = df.columns.tolist()

    # 5. 填充每个表格
    for city, table in city_table_map.items():
        city_data = df[df['城市'] == city]

        # 获取 Word 表格表头（第一行）
        header_row = table.rows[0]
        header_texts = [cell.text.strip() for cell in header_row.cells]

        # 建立列映射
        col_mapping = {}
        for i, header in enumerate(header_texts):
            if header in excel_columns:
                col_mapping[i] = header
            else:
                # 模糊匹配（去除空格、括号）
                clean_header = header.replace(' ', '').replace('（', '').replace('）', '')
                for col in excel_columns:
                    if clean_header in col.replace(' ', '').replace('（', '').replace('）', ''):
                        col_mapping[i] = col
                        break
                if i not in col_mapping:
                    col_mapping[i] = None

        # 调整行数
        existing_rows = len(table.rows) - 1
        needed_rows = len(city_data) - existing_rows
        if needed_rows > 0:
            for _ in range(needed_rows):
                table.add_row()

        # 填充数据
        for row_idx, (_, row_data) in enumerate(city_data.iterrows()):
            target_row = table.rows[row_idx + 1]
            for col_idx, excel_col in col_mapping.items():
                if col_idx < len(target_row.cells):
                    value = row_data.get(excel_col, '') if excel_col else ''
                    target_row.cells[col_idx].text = str(value)

    doc.save(output_path)
    logger.info(f"✅ Word 文件已生成（动态版本）：{output_path}")
    return output_path

if __name__ == '__main__':
    app.run(debug=True)