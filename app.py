import os
import re
import logging
from flask import Flask, render_template, request, send_file
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

if __name__ == '__main__':
    app.run(debug=True)