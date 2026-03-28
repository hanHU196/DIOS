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

# 初始化处理器
processor = DocumentProcessor()
reader = DocumentReader()
app = Flask(__name__)  

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
    try:
        files = request.files.getlist('documents')
        command = request.form.get('command', '').strip()
        
        if not files or not command:
            return jsonify({'error': '请上传文档并输入指令'}), 400
        
        doc_paths = []
        for file in files:
            filename = secure_filename(file.filename)
            path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(path)
            doc_paths.append(path)
        
        all_text = ""
        for path in doc_paths:
            text = reader.read(path)
            all_text += f"\n--- {os.path.basename(path)} ---\n{text}\n"
        
        def parse_fields(instruction):
            import re
            match = re.search(r'提取[：:]\s*(.+)', instruction)
            if not match:
                match = re.search(r'提取\s*(.+)', instruction)
            if match:
                fields_str = match.group(1)
                fields = re.split(r'[、，,\s]+', fields_str)
                return [f.strip() for f in fields if f.strip()]
            words = re.split(r'[、，,\s]+', instruction)
            return [w for w in words if len(w) >= 2 and not re.search(r'[a-zA-Z0-9]', w)]
        
        fields = parse_fields(command)
        if not fields:
            return jsonify({'error': '无法从指令中识别字段'}), 400
        
        extracted = ai_module.extract_entities(all_text, fields)
        
        if isinstance(extracted, dict):
            extracted = [extracted]
        
        return jsonify({'success': True, 'fields': fields, 'data': extracted})
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/fill', methods=['POST'])
def api_fill():
    """表格智能填写（支持多文档）"""
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
        if not template_filename.endswith('.xlsx'):
            template_filename = template_filename + '.xlsx'
        template_path = os.path.join(app.config['UPLOAD_FOLDER'], 'template_' + template_filename)
        template_file.save(template_path)
        logger.info(f"📊 模板已保存：{template_filename}")
        
        # 解析模板字段
        from excel_handler import parse_excel_template
        fields = parse_excel_template(template_path)
        logger.info(f"📋 模板字段：{fields}")
        
        # 解析筛选条件（可选）
        try:
            from ai_module import parse_filter_conditions
            filter_result = parse_filter_conditions(command)
            filters = filter_result.get('filters', [])
            logger.info(f"🔍 筛选条件: {filters}")
        except:
            filters = []
            logger.info("🔍 无筛选条件")
        
        # 合并所有文档的数据
        all_data = []
        for doc_path in doc_paths:
            if doc_path.endswith('.xlsx') or doc_path.endswith('.xls'):
                # Excel 文件：直接读取
                logger.info(f"📊 处理 Excel: {os.path.basename(doc_path)}")
                df = pd.read_excel(doc_path)
                logger.info(f"   原始数据: {len(df)} 行")
                
                # 应用筛选
                if filters:
                    df = apply_filters_to_df(df, filters)
                    logger.info(f"   筛选后: {len(df)} 行")
                
                data = df.to_dict('records')
                all_data.extend(data)
                
            else:
                # Word/TXT 文件：用 AI 提取
                logger.info(f"🤖 处理文档: {os.path.basename(doc_path)}")
                text = reader.read(doc_path)
                extracted = ai_module.extract_entities_safe(text, fields)
                
                # 打印调试
                if isinstance(extracted, list):
                    logger.info(f"   AI 提取到 {len(extracted)} 条数据")
                    if len(extracted) > 0:
                        logger.info(f"   第一条数据: {extracted[0]}")
                else:
                    logger.info(f"   AI 返回类型: {type(extracted)}")
                
                # 直接添加，不筛选（Word 数据量小，先保证能填）
                if isinstance(extracted, list):
                    all_data.extend(extracted)
                elif isinstance(extracted, dict):
                    all_data.append(extracted)
        
        logger.info(f"📊 合并后共 {len(all_data)} 行数据")
        
        # 生成输出文件
        from datetime import datetime
        output_filename = f"filled_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        output_path = os.path.join('outputs', output_filename)
        os.makedirs('outputs', exist_ok=True)
        
        from excel_handler import fill_excel_with_data
        fill_excel_with_data(template_path, all_data, output_path)
        
        return send_file(output_path, as_attachment=True, download_name=output_filename)
        
    except Exception as e:
        logger.error(f"填表失败: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


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

if __name__ == '__main__':
    app.run(debug=True)