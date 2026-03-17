# instruction_parser.py
import re

class InstructionParser:
    """自然语言指令解析器"""
    
    def __init__(self):
        # 定义支持的指令模式
        self.patterns = {
            # Word操作
            'bold': r'(加粗|粗体|bold)',
            'italic': r'(斜体|italic)',
            'underline': r'(下划线|underline)',
            'center': r'(居中|居中对齐|center)',
            'left': r'(左对齐|left)',
            'right': r'(右对齐|right)',
            'font_size': r'字体大小[设为为]?(\d+)',
            'insert_table': r'插入表格[设为]?(\d+)x(\d+)',
            'insert_image': r'插入图片',
            
            # Excel操作
            'excel_bold': r'加粗[表头|标题]',
            'excel_center': r'居中',
            'excel_width': r'列宽[设为]?(\d+)',
            'excel_sum': r'求和|合计',
            'excel_chart': r'插入图表',
            
            # 通用
            'extract': r'提取(.*)',
            'fill': r'填表|填写',
        }
    
    def parse(self, instruction):
        """
        解析指令，返回操作类型和参数
        输入："把标题加粗，字体大小16"
        输出：{
            'operations': [
                {'type': 'bold', 'target': 'title'},
                {'type': 'font_size', 'size': 16}
            ]
        }
        """
        result = {'operations': []}
        
        # 处理加粗
        if re.search(self.patterns['bold'], instruction, re.IGNORECASE):
            # 判断是加粗标题还是全文
            if '标题' in instruction:
                result['operations'].append({'type': 'bold', 'target': 'title'})
            else:
                result['operations'].append({'type': 'bold', 'target': 'all'})
        
        # 处理斜体
        if re.search(self.patterns['italic'], instruction, re.IGNORECASE):
            result['operations'].append({'type': 'italic'})
        
        # 处理下划线
        if re.search(self.patterns['underline'], instruction, re.IGNORECASE):
            result['operations'].append({'type': 'underline'})
        
        # 处理字体大小
        size_match = re.search(self.patterns['font_size'], instruction)
        if size_match:
            result['operations'].append({
                'type': 'font_size', 
                'size': int(size_match.group(1))
            })
        
        # 处理居中
        if re.search(self.patterns['center'], instruction, re.IGNORECASE):
            result['operations'].append({'type': 'center'})
        
        # 处理左对齐
        if re.search(self.patterns['left'], instruction, re.IGNORECASE):
            result['operations'].append({'type': 'left'})
        
        # 处理右对齐
        if re.search(self.patterns['right'], instruction, re.IGNORECASE):
            result['operations'].append({'type': 'right'})
        
        # 处理插入表格
        table_match = re.search(self.patterns['insert_table'], instruction)
        if table_match:
            result['operations'].append({
                'type': 'insert_table',
                'rows': int(table_match.group(1)),
                'cols': int(table_match.group(2))
            })
        
        # 处理Excel加粗表头
        if re.search(self.patterns['excel_bold'], instruction, re.IGNORECASE):
            result['operations'].append({'type': 'excel_bold'})
        
        # 处理Excel列宽
        width_match = re.search(self.patterns['excel_width'], instruction)
        if width_match:
            result['operations'].append({
                'type': 'excel_width',
                'size': int(width_match.group(1))
            })
        
        # 处理Excel求和
        if re.search(self.patterns['excel_sum'], instruction, re.IGNORECASE):
            result['operations'].append({'type': 'excel_sum'})
        
        return result
    
    def get_document_type(self, instruction):
        """判断要操作的是Word还是Excel"""
        if 'excel' in instruction.lower() or '表格' in instruction:
            return 'excel'
        elif 'word' in instruction.lower() or '文档' in instruction:
            return 'word'
        else:
            return 'unknown'


# 测试代码
if __name__ == "__main__":
    parser = InstructionParser()
    
    test_instructions = [
        "把标题加粗，居中",
        "字体大小16，插入3x4表格",
        "加粗表头，列宽15",
        "提取甲方、乙方、金额",
        "斜体并加下划线",
        "求和",
    ]
    
    print("="*50)
    print("指令解析器测试")
    print("="*50)
    
    for inst in test_instructions:
        result = parser.parse(inst)
        print(f"\n📝 指令：{inst}")
        print(f"🔍 解析：{result}")
    
    print("\n" + "="*50)
    print("测试完成")