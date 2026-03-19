import os
import re
import logging
from flask import Flask, render_template, request, send_file,jsonify
from werkzeug.utils import secure_filename
from processor import DocumentProcessor

# 初始化处理器
processor = DocumentProcessor()
app = Flask(__name__)

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# 配置
UPLOAD_FOLDER = 'uploads'
ALLOWED_DOC_EXTENSIONS = {'txt', 'docx', 'md', 'xlsx'}  # 支持的文件类型
ALLOWED_TEMPLATE_EXTENSIONS = {'xlsx', 'xls'}  # 模板文件类型
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 限制16MB

# 确保上传文件夹存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename, allowed_set=ALLOWED_DOC_EXTENSIONS):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_set

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    # 获取上传的文件和指令
    doc_file = request.files.get('document')
    template_file = request.files.get('template')
    instruction = request.form.get('command', '').strip()

    if not doc_file or not template_file or not instruction:
        return "请上传文档和模板文件，并输入指令", 400

    # 检查文件格式
    if not allowed_file(doc_file.filename, ALLOWED_DOC_EXTENSIONS):
        return "文档格式不支持，请上传 txt/docx/md/xlsx 文件", 400
    if not allowed_file(template_file.filename, ALLOWED_TEMPLATE_EXTENSIONS):
        return "模板格式不支持，请上传 Excel 文件（.xlsx 或 .xls）", 400

    # 保存文档
    doc_filename = secure_filename(doc_file.filename)
    doc_path = os.path.join(app.config['UPLOAD_FOLDER'], doc_filename)
    doc_file.save(doc_path)

    # 保存模板（确保有.xlsx后缀）
    template_filename = secure_filename(template_file.filename)
    if not template_filename.endswith('.xlsx'):
        template_filename = template_filename + '.xlsx'
    template_path = os.path.join(app.config['UPLOAD_FOLDER'], 'template_' + template_filename)
    template_file.save(template_path)
    logger.info(f"✅ 模板保存为：{template_path}")

    # 创建输出目录
    output_dir = "outputs"
    os.makedirs(output_dir, exist_ok=True)

    # 生成唯一输出文件名
    from datetime import datetime
    output_filename = f"result_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    output_path = os.path.join(output_dir, output_filename)

    # 调用统一的处理器
    try:
        result = processor.process_single(
            doc_path=doc_path,
            template_path=template_path,
            output_path=output_path,
            instruction=instruction
        )
        
        if result['success']:
            logger.info(f"✅ 处理成功，生成文件：{output_path}")
            return send_file(output_path, as_attachment=True, download_name='result.xlsx')
        else:
            return f"处理失败：{result.get('error', '未知错误')}", 500
            
    except Exception as e:
        logger.error(f"处理异常：{e}")
        return f"处理异常：{str(e)}", 500

@app.route('/batch_process', methods=['POST'])
def batch_process():
    """批量处理多个模板"""
    import json
    from datetime import datetime
    
    # 获取上传的多个文档和多个模板
    doc_files = request.files.getlist('documents')
    template_files = request.files.getlist('templates')
    
    if not doc_files or not template_files:
        return jsonify({'error': '请上传文档和模板文件'}), 400
    
    # 保存所有文档
    doc_paths = []
    for doc_file in doc_files:
        doc_filename = secure_filename(doc_file.filename)
        doc_path = os.path.join(app.config['UPLOAD_FOLDER'], doc_filename)
        doc_file.save(doc_path)
        doc_paths.append(doc_path)
        logger.info(f"✅ 文档已保存：{doc_filename}")
    
    # 保存所有模板
    template_paths = []
    for template_file in template_files:
        template_filename = secure_filename(template_file.filename)
        # 确保文件名有 .xlsx 后缀（如果是Excel）
        if template_filename.endswith('.xlsx') or template_filename.endswith('.xls'):
            template_path = os.path.join(app.config['UPLOAD_FOLDER'], 'template_' + template_filename)
        else:
            # 其他格式（如docx）保持原样
            template_path = os.path.join(app.config['UPLOAD_FOLDER'], 'template_' + template_filename)
        template_file.save(template_path)
        template_paths.append(template_path)
        logger.info(f"✅ 模板已保存：{template_filename}")
    
    # 创建输出目录
    batch_dir = f"batch_output_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    output_dir = os.path.join('outputs', batch_dir)
    os.makedirs(output_dir, exist_ok=True)
    
    # 先建立文档索引
    logger.info("📚 开始建立文档索引...")
    processor.matcher.index_documents(doc_paths)
    
    # 执行批处理
    logger.info("🚀 开始批量处理...")
    results = processor.process_batch(doc_paths, template_paths, output_dir)
    
    # 统计结果
    success_count = sum(1 for r in results if r['success'])
    
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
    
    # 整理结果（避免路径过长）
    for r in results:
        result_item = {
            'template_name': os.path.basename(r['template']),
            'success': r['success']
        }
        if r['success']:
            result_item['matched_doc'] = r.get('matched_doc', '')
            result_item['output_file'] = os.path.basename(r['output_path'])
        else:
            result_item['error'] = r.get('error', '未知错误')
        response['results'].append(result_item)
    
    logger.info(f"✅ 批处理完成，成功：{success_count}/{len(results)}")
    return jsonify(response)

@app.route('/search', methods=['POST'])
def search_documents():
    """文档检索接口"""
    data = request.get_json()
    query = data.get('query', '').strip()
    
    if not query:
        return jsonify({'error': '请输入搜索关键词'}), 400
    
    # 获取要检索的文档（检索 uploads 文件夹下的所有文件）
    import os
    upload_folder = app.config['UPLOAD_FOLDER']
    doc_paths = []
    if os.path.exists(upload_folder):
        for filename in os.listdir(upload_folder):
            file_path = os.path.join(upload_folder, filename)
            if os.path.isfile(file_path):
                doc_paths.append(file_path)
    
    if not doc_paths:
        return jsonify({'error': '没有可检索的文档'}), 400
    
    # 执行检索
    results = processor.search_documents(query, doc_paths)
    
    return jsonify({
        'success': True,
        'query': query,
        'results': results,
        'count': len(results)
    })

@app.route('/operate', methods=['POST'])
def operate_document():
    """智能指令操作文档"""
    
    file = request.files.get('file')
    instruction = request.form.get('instruction', '').strip()
    
    if not file or not instruction:
        return jsonify({'error': '请上传文件并输入指令'}), 400
    
    # 保存文件
    filename = secure_filename(file.filename)
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(file_path)
    
    # 执行操作
    result = processor.operate_document(file_path, instruction)
    
    if result['success']:
        # 返回修改后的文件
        return send_file(file_path, as_attachment=True, 
                        download_name=f'operated_{filename}')
    else:
        return jsonify({'error': result['error']}), 400

@app.route('/match_and_fill', methods=['POST'])
def match_and_fill():
    """先匹配文档，再填表"""
    
    # 获取上传的多个文档和单个模板
    doc_files = request.files.getlist('documents')
    template_file = request.files.get('template')
    instruction = request.form.get('command', '').strip()
    
    if not doc_files or not template_file:
        return jsonify({'error': '请上传文档库和模板文件'}), 400
    
    # 保存所有文档
    doc_paths = []
    for doc_file in doc_files:
        doc_filename = secure_filename(doc_file.filename)
        doc_path = os.path.join(app.config['UPLOAD_FOLDER'], doc_filename)
        doc_file.save(doc_path)
        doc_paths.append(doc_path)
    
    # 保存模板
    template_filename = secure_filename(template_file.filename)
    if not template_filename.endswith('.xlsx'):
        template_filename = template_filename + '.xlsx'
    template_path = os.path.join(app.config['UPLOAD_FOLDER'], 'template_' + template_filename)
    template_file.save(template_path)
    
    # 创建输出目录
    from datetime import datetime
    output_filename = f"matched_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    output_path = os.path.join('outputs', output_filename)
    
    # 调用匹配填表
    try:
        result = processor.process_with_matching(
            template_path=template_path,
            doc_paths=doc_paths,
            output_path=output_path,
            instruction=instruction
        )
        
        if result['success']:
            return send_file(output_path, as_attachment=True, 
                           download_name='matched_result.xlsx')
        else:
            return jsonify({'error': result.get('error', '未知错误')}), 500
            
    except Exception as e:
        logger.error(f"处理异常：{e}")
        return jsonify({'error': str(e)}), 500




if __name__ == '__main__':
    app.run(debug=True)