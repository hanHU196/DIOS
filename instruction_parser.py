# instruction_parser.py
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from deepseek_parser import DeepSeekParser
import logging

logger = logging.getLogger(__name__)

class InstructionOperator:
     def __init__(self, api_key=None):
        self.api_key = api_key
        if api_key:
            self.deepseek = DeepSeekParser(api_key)
        else:
            self.deepseek = None
        logger.info("✅ 指令操作器初始化成功")
    
     def execute_standard(self, command, file_path):
        """兼容旧代码的 execute_standard 方法"""
        if isinstance(command, dict):
            if file_path.endswith('.docx'):
                return self._execute_word(command, file_path)
            elif file_path.endswith('.xlsx'):
                return self._execute_excel(command, file_path)
        return {'success': False, 'error': '无效的指令格式'}
    
     def execute(self, instruction, file_path):
        """执行自然语言指令（主入口）"""
        if self.deepseek:
            command = self.deepseek.parse_instruction(instruction)
            if command.get('action') == 'unknown':
                return {'success': False, 'error': '无法理解指令'}
        else:
            return {'success': False, 'error': '未配置 DeepSeek API Key'}
        
        if file_path.endswith('.docx'):
            return self._execute_word(command, file_path)
        elif file_path.endswith('.xlsx'):
            return self._execute_excel(command, file_path)
        else:
            return {'success': False, 'error': '不支持的文件类型'}
    
     def _execute_word(self, command, file_path):
        """执行 Word 操作（支持位置定位）"""
        doc = Document(file_path)
        action = command.get('action')
        position = command.get('position')
        target = command.get('target', 'all')
        size = command.get('size', 12)
        
        # 确定要操作的段落
        if position is not None and target == 'paragraph':
            if 1 <= position <= len(doc.paragraphs):
                paragraphs = [doc.paragraphs[position - 1]]
            else:
                return {'success': False, 'error': f'段落位置 {position} 超出范围（共 {len(doc.paragraphs)} 段）'}
        elif target == 'title':
            paragraphs = [doc.paragraphs[0]] if doc.paragraphs else []
        else:
            paragraphs = doc.paragraphs
        
        # 执行操作
        if action == 'bold':
            for para in paragraphs:
                for run in para.runs:
                    run.bold = True
            msg = f'✅ 已加粗 {len(paragraphs)} 段'
        elif action == 'italic':
            for para in paragraphs:
                for run in para.runs:
                    run.italic = True
            msg = f'✅ 已设置斜体 {len(paragraphs)} 段'
        elif action == 'underline':
            for para in paragraphs:
                for run in para.runs:
                    run.underline = True
            msg = f'✅ 已添加下划线 {len(paragraphs)} 段'
        elif action == 'center':
            for para in paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            msg = f'✅ 已居中 {len(paragraphs)} 段'
        elif action == 'left':
            for para in paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            msg = f'✅ 已左对齐 {len(paragraphs)} 段'
        elif action == 'right':
            for para in paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            msg = f'✅ 已右对齐 {len(paragraphs)} 段'
        elif action == 'font_size':
            for para in paragraphs:
                for run in para.runs:
                    run.font.size = Pt(size)
            msg = f'✅ 字体大小设为 {size}'
        elif action == 'insert_table':
            rows = command.get('rows', 3)
            cols = command.get('cols', 3)
            table = doc.add_table(rows=rows, cols=cols)
            try:
                table.style = 'Table Grid'
            except:
                pass
            doc.save(file_path)
            return {'success': True, 'message': f'✅ 已插入 {rows}x{cols} 表格'}
        else:
            return {'success': False, 'error': f'不支持的操作：{action}'}
        
        doc.save(file_path)
        return {'success': True, 'message': msg}
    
     def _execute_excel(self, command, file_path):
        """执行 Excel 操作（支持位置定位）"""
        wb = load_workbook(file_path)
        ws = wb.active
        action = command.get('action')
        row_num = command.get('row')
        col_num = command.get('col')
        width = command.get('width', 15)
        
        # 确定要操作的单元格
        if row_num is not None and col_num is not None:
            cells = [ws.cell(row=row_num, column=col_num)]
            msg_target = f"单元格({row_num},{col_num})"
        elif row_num is not None:
            cells = ws[row_num]
            msg_target = f"第{row_num}行"
        elif col_num is not None:
            cells = [ws.cell(row=i, column=col_num) for i in range(1, ws.max_row + 1)]
            msg_target = f"第{col_num}列"
        elif command.get('target') == 'header':
            cells = ws[1]
            msg_target = "表头"
        else:
            cells = []
            for row in ws.iter_rows():
                cells.extend(row)
            msg_target = "全部"
        
        if action == 'excel_bold':
            for cell in cells:
                cell.font = Font(bold=True)
            msg = f'✅ 已加粗 {msg_target}'
        elif action == 'excel_center':
            for cell in cells:
                cell.alignment = Alignment(horizontal='center')
            msg = f'✅ 已居中 {msg_target}'
        elif action == 'excel_width' and col_num is not None:
            col_letter = chr(64 + col_num)
            ws.column_dimensions[col_letter].width = width
            msg = f'✅ 第{col_num}列宽设为{width}'
        elif action == 'excel_sum':
            last_row = ws.max_row + 1
            ws.cell(row=last_row, column=1, value="合计")
            for col in range(2, ws.max_column + 1):
                col_letter = chr(64 + col)
                ws.cell(row=last_row, column=col, 
                       value=f"=SUM({col_letter}2:{col_letter}{last_row-1})")
            msg = '✅ 已添加合计行'
        else:
            return {'success': False, 'error': f'不支持的操作：{action}'}
        
        wb.save(file_path)
        return {'success': True, 'message': msg}