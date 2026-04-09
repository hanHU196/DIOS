# instruction_parser.py
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from deepseek_parser import DeepSeekParser
import logging

logger = logging.getLogger(__name__)

class InstructionOperator:
    def __init__(self, api_key=None):
        self.api_key = api_key
        if api_key:
            self.deepseek = DeepSeekParser(api_key)
            logger.info("✅ DeepSeek 指令解析器初始化成功")
        else:
            self.deepseek = None
            logger.warning("⚠️ 未配置 DeepSeek API Key")
    
    def execute(self, instruction, file_path):
        """执行自然语言指令"""
        if not self.deepseek:
            return {'success': False, 'error': 'DeepSeek 未配置'}
        
        # 根据文件类型选择解析方式
        if file_path.endswith('.docx'):
            file_type = "word"
        elif file_path.endswith('.xlsx'):
            file_type = "excel"
        else:
            return {'success': False, 'error': '不支持的文件类型'}
        
        # 解析指令
        command = self.deepseek.parse_instruction(instruction, file_type)
        logger.info(f"解析结果: {command}")
        
        if command.get('action') == 'unknown':
            return {'success': False, 'error': f'无法理解指令: {instruction}'}
        
        # 执行操作
        if file_type == "word":
            return self._execute_word(command, file_path)
        else:
            return self._execute_excel(command, file_path)
    
    def _execute_word(self, command, file_path):
        """执行 Word 操作"""
        try:
            doc = Document(file_path)
            action = command.get('action')
            position = command.get('position')
            target = command.get('target', 'all')
            size = command.get('size')
            font_name = command.get('font', '')
            font_color = command.get('font_color', '')
            
            # 颜色映射
            color_map = {
                '红色': RGBColor(255, 0, 0),
                '蓝色': RGBColor(0, 0, 255),
                '绿色': RGBColor(0, 255, 0),
                '黑色': RGBColor(0, 0, 0)
            }
            
            # 确定段落
            if position and target == 'paragraph':
                if 1 <= position <= len(doc.paragraphs):
                    paragraphs = [doc.paragraphs[position - 1]]
                else:
                    return {'success': False, 'error': f'段落 {position} 超出范围'}
            else:
                paragraphs = doc.paragraphs
            
            # 应用格式
            for para in paragraphs:
                for run in para.runs:
                    if action == 'bold':
                        run.bold = True
                    elif action == 'italic':
                        run.italic = True
                    elif action == 'underline':
                        run.underline = True
                    if font_name:
                        run.font.name = font_name
                    if size:
                        run.font.size = Pt(size)
                    if font_color and font_color in color_map:
                        run.font.color.rgb = color_map[font_color]
                
                # 对齐
                if action == 'center':
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                elif action == 'left':
                    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                elif action == 'right':
                    para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            
            doc.save(file_path)
            return {'success': True, 'message': '格式调整完成'}
            
        except Exception as e:
            logger.error(f"Word 操作失败: {e}")
            return {'success': False, 'error': str(e)}
    
    def _execute_excel(self, command, file_path):
        """执行 Excel 操作"""
        try:
            wb = load_workbook(file_path)
            ws = wb.active
            action = command.get('action')
            row = command.get('row')
            col = command.get('col')
            row_start = command.get('row_start')
            row_end = command.get('row_end')
            col_start = command.get('col_start')
            col_end = command.get('col_end')
            width = command.get('width', 15)
            height = command.get('height', 20)
            font_name = command.get('font', '')
            size = command.get('size', 11)
            
            # 确定操作范围
            if row and col:
                # 单个单元格
                cells = [(row, col)]
            elif row_start and row_end:
                # 行范围
                cells = [(r, 1) for r in range(row_start, row_end + 1)]
            elif row:
                # 单行
                cells = [(row, c) for c in range(1, ws.max_column + 1)]
            elif col:
                # 单列
                cells = [(r, col) for r in range(1, ws.max_row + 1)]
            else:
                # 全部
                cells = [(r, c) for r in range(1, ws.max_row + 1) for c in range(1, ws.max_column + 1)]
            
            # 应用格式
            for r, c in cells:
                cell = ws.cell(row=r, column=c)
                
                if action == 'excel_bold':
                    cell.font = Font(bold=True)
                elif action == 'excel_italic':
                    cell.font = Font(italic=True)
                elif action == 'excel_center':
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                elif action == 'excel_left':
                    cell.alignment = Alignment(horizontal='left')
                elif action == 'excel_right':
                    cell.alignment = Alignment(horizontal='right')
                
                if font_name:
                    cell.font = Font(name=font_name, size=size)
                elif size:
                    cell.font = Font(size=size)
            
            # 列宽
            if action == 'excel_width' and col:
                col_letter = get_column_letter(col)
                ws.column_dimensions[col_letter].width = width
            
            # 行高
            if action == 'excel_height' and row:
                ws.row_dimensions[row].height = height
            
            # 合并单元格
            if action == 'excel_merge' and row_start and col_start and row_end and col_end:
                start_cell = get_column_letter(col_start) + str(row_start)
                end_cell = get_column_letter(col_end) + str(row_end)
                ws.merge_cells(f'{start_cell}:{end_cell}')
            
            wb.save(file_path)
            return {'success': True, 'message': '格式调整完成'}
            
        except Exception as e:
            logger.error(f"Excel 操作失败: {e}")
            return {'success': False, 'error': str(e)}