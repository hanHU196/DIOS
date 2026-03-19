# instruction_parser.py
# 解析智能指令并执行操作
import re
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class InstructionOperator:
    """智能指令操作器"""
    
    def __init__(self):
        logger.info("✅ 指令操作器初始化成功")
    
    def execute_standard(self, command, file_path):
        """
        执行标准化指令
        command: {
            "operation": "bold",
            "target": "title",
            "parameters": {}
        }
        """
        operation = command.get('operation')
        target = command.get('target', 'all')
        params = command.get('parameters', {})
        
        logger.info(f"📝 执行操作：{operation}，目标：{target}")
        
        # 判断文件类型
        if file_path.endswith('.docx'):
            return self._execute_word(operation, target, params, file_path)
        elif file_path.endswith('.xlsx'):
            return self._execute_excel(operation, target, params, file_path)
        else:
            return {'success': False, 'error': f'不支持的文件类型：{file_path}'}
    
    def _execute_word(self, operation, target, params, file_path):
        """执行Word操作"""
        doc = Document(file_path)
        
        if operation == 'bold':
            for para in doc.paragraphs:
                if target == 'title' and para.style.name == 'Title':
                    for run in para.runs:
                        run.bold = True
                elif target == 'all':
                    for run in para.runs:
                        run.bold = True
        
        elif operation == 'center':
            for para in doc.paragraphs:
                if target == 'title' and para.style.name == 'Title':
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                elif target == 'all':
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        elif operation == 'font_size':
            size = params.get('size', 12)
            for para in doc.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(size)
        
        elif operation == 'insert_table':
            rows = params.get('rows', 3)
            cols = params.get('cols', 3)
            table = doc.add_table(rows=rows, cols=cols)
            # 尝试多种样式
            try:
                table.style = 'Table Grid'  # 优先尝试这个
            except:
                try:
                    table.style = 'Light Grid Accent 1'  # 另一个常见样式
                except:
                    # 如果都不行，就不设置样式
                    pass
        
        doc.save(file_path)
        return {'success': True, 'message': f'✅ 已执行 {operation}'}
    
    def _execute_excel(self, operation, target, params, file_path):
        """执行Excel操作"""
        wb = load_workbook(file_path)
        ws = wb.active
        
        if operation == 'bold':
            if target == 'header':
                for cell in ws[1]:
                    cell.font = Font(bold=True)
            elif target == 'all':
                for row in ws.iter_rows():
                    for cell in row:
                        cell.font = Font(bold=True)
        
        elif operation == 'center':
            for row in ws.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(horizontal='center')
        
        elif operation == 'column_width':
            width = params.get('width', 15)
            for col in range(1, ws.max_column + 1):
                col_letter = chr(64 + col) if col <= 26 else f"Column{col}"
                ws.column_dimensions[col_letter].width = width
        
        elif operation == 'sum':
            last_row = ws.max_row + 1
            ws.cell(row=last_row, column=1, value="合计")
            for col in range(2, ws.max_column + 1):
                col_letter = chr(64 + col)
                ws.cell(row=last_row, column=col, 
                       value=f"=SUM({col_letter}2:{col_letter}{last_row-1})")
        
        wb.save(file_path)
        return {'success': True, 'message': f'✅ 已执行 {operation}'}