from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Inches, Cm, Mm, Pt, Emu
import json
import numpy as np  # 用于统计分析
from check import check_file

def pt_to_font_size(pt_size):
    """将磅值转换为中文标准字号"""
    font_size_map = {
        42: '初号',      # 42pt
        36: '小初',      # 36pt
        26: '一号',      # 26pt
        24: '小一',      # 24pt
        22: '二号',      # 22pt
        18: '小二',      # 18pt
        16: '三号',      # 16pt
        15: '小三',      # 15pt
        14: '四号',      # 14pt
        12: '小四',      # 12pt
        10.5: '五号',    # 10.5pt
        9: '小五',       # 9pt
        7.5: '六号',     # 7.5pt
        5.5: '小六',     # 5.5pt
        5: '七号',       # 5pt
    }
    # 找到最接近的字号
    closest_size = min(font_size_map.keys(), key=lambda x: abs(x - pt_size))
    return font_size_map[closest_size]

def get_measurement_info(measurement):
    """Get measurement value and unit from Word measurement object"""
    if measurement is None:
        return None, None, None
    
    # 检查测量对象的类型和单位
    if hasattr(measurement, 'pt'):
        # 如果是磅
        return measurement.pt, '磅', 'pt'
    elif hasattr(measurement, 'inches'):
        # 如果是英寸
        return measurement.inches, '英寸', 'inch'
    elif hasattr(measurement, 'cm'):
        # 如果是厘米
        return measurement.cm, '厘米', 'cm'
    elif hasattr(measurement, 'mm'):
        # 如果是毫米
        return measurement.mm, '毫米', 'mm'
    elif hasattr(measurement, 'emu'):
        # 如果是EMU (English Metric Units)
        return measurement.emu, 'EMU', 'emu'
    else:
        # 如果是字符或其他单位
        return measurement, '字符', 'char'

class Node:
    def __init__(self, node_type, content, font=None, size=None, bold=False, spacing=None, line_spacing=None, indentation=None, outline_level=None, alignment=None, direction=None, paragraph_format=None):
        self.type = node_type  # e.g., 'heading1', 'heading2', 'paragraph'
        self.content = content
        self.font = font
        self.size = size  # in pt
        self.font_size_name = pt_to_font_size(size) if size else None  # 字号
        self.bold = bold
        self.spacing = spacing  # paragraph spacing after, in pt
        self.line_spacing = line_spacing  # line spacing in pt
        self.indentation = indentation  # indentation information
        self.outline_level = outline_level  # outline level from Word
        self.alignment = alignment  # paragraph alignment
        self.direction = direction  # text direction
        self.paragraph_format = paragraph_format  # complete paragraph format info
        self.children = []  # list of child nodes

    def add_child(self, child_node):
        self.children.append(child_node)

    def to_dict(self):
        return {
            'type': self.type,
            'content': self.content,
            'font': self.font,
            'size': self.size,
            'font_size_name': self.font_size_name,  # 字号
            'bold': self.bold,
            'spacing': self.spacing,
            'line_spacing': self.line_spacing,
            'indentation': self.indentation,
            'outline_level': self.outline_level,
            'alignment': self.alignment,
            'direction': self.direction,
            'paragraph_format': self.paragraph_format,
            'children': [child.to_dict() for child in self.children]
        }

def get_font_size_distribution(doc):
    """Collect font sizes from all paragraphs."""
    sizes = []
    for paragraph in doc.paragraphs:
        if not paragraph.text.strip():  # Skip empty paragraphs
            continue
        first_run = paragraph.runs[0] if paragraph.runs else None
        size = first_run.font.size.pt if first_run and first_run.font.size else 12.0
        sizes.append(size)
    return sizes

def get_alignment_from_xml(paragraph):
    """尝试从XML中直接获取对齐方式"""
    try:
        # 获取段落的XML元素
        p_element = paragraph._element
        
        # 查找段落属性
        pPr = p_element.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
        if pPr is not None:
            # 查找对齐方式元素
            jc = pPr.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}jc')
            if jc is not None:
                val = jc.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                if val:
                    # XML中的对齐方式值映射
                    xml_alignment_map = {
                        'left': '左对齐',
                        'center': '居中',
                        'right': '右对齐',
                        'both': '两端对齐',
                        'justify': '两端对齐',
                        'distribute': '分散对齐',
                        'start': '左对齐',
                        'end': '右对齐'
                    }
                    return xml_alignment_map.get(val, f'XML对齐方式({val})')
        
        return '左对齐'  # XML中没有找到时的默认值
        
    except Exception as e:
        print(f"从XML获取对齐方式时出错: {e}")
        return '左对齐'

def get_alignment_info(pf, paragraph):
    """获取段落对齐方式的修复版本"""
    # 完整的对齐方式映射（只包含python-docx实际支持的）
    alignment_map = {
        0: '左对齐',           # WD_ALIGN_PARAGRAPH.LEFT
        1: '居中',            # WD_ALIGN_PARAGRAPH.CENTER  
        2: '右对齐',          # WD_ALIGN_PARAGRAPH.RIGHT
        3: '两端对齐',        # WD_ALIGN_PARAGRAPH.JUSTIFY
        4: '分散对齐',        # WD_ALIGN_PARAGRAPH.DISTRIBUTE (如果存在)
        7: '泰文两端对齐'     # WD_ALIGN_PARAGRAPH.THAI_JUSTIFY (如果存在)
    }
    
    # 尝试多种方法获取对齐方式
    alignment_value = None
    alignment_name = '未知'
    
    try:
        # 方法1：直接获取alignment属性
        if hasattr(pf, 'alignment') and pf.alignment is not None:
            alignment_value = pf.alignment
            # 如果是枚举对象，获取其数值
            if hasattr(alignment_value, 'value'):
                alignment_numeric = alignment_value.value
            else:
                alignment_numeric = int(alignment_value) if alignment_value is not None else None
            
            if alignment_numeric is not None:
                alignment_name = alignment_map.get(alignment_numeric, f'未知对齐方式({alignment_numeric})')
            else:
                alignment_name = '左对齐'  # 默认值
        else:
            # 方法2：尝试通过XML直接读取
            alignment_name = get_alignment_from_xml(paragraph)
            
    except Exception as e:
        print(f"获取对齐方式时出错: {e}")
        alignment_name = '左对齐'  # 出错时的默认值
    
    return {
        'type': '对齐方式',
        'value': alignment_name,
        'raw_value': str(alignment_value) if alignment_value is not None else 'None'
    }

def extract_paragraph_formatting(paragraph):
    """Extract detailed formatting information from paragraph including all properties from the dialog."""
    format_info = {}
    
    # Get paragraph format
    pf = paragraph.paragraph_format
    
    # 1. 常规 (General) Section
    # Alignment (对齐方式) - 使用修复后的函数
    format_info['alignment'] = get_alignment_info(pf, paragraph)
    
    # Outline Level (大纲级别)
    outline_map = {
        0: '标题1', 1: '标题2', 2: '标题3', 3: '标题4', 4: '标题5',
        5: '标题6', 6: '标题7', 7: '标题8', 8: '标题9', 9: '正文文本'
    }
    
    if hasattr(pf, 'outline_level') and pf.outline_level is not None:
        format_info['outline_level'] = {
            'type': '大纲级别',
            'value': outline_map.get(pf.outline_level, str(pf.outline_level)),
            'raw_value': pf.outline_level
        }
    else:
        format_info['outline_level'] = {
            'type': '大纲级别',
            'value': '正文文本',
            'raw_value': 9
        }
    
    # Direction (方向) - 从右向左/从左向右
    format_info['direction'] = {
        'type': '方向',
        'value': '从左向右',
        'raw_value': 'LTR'
    }
    
    # 2. 缩进 (Indentation) Section
    # 文本之前 (Before text) - 左缩进
    if pf.left_indent:
        value, unit_name, unit_code = get_measurement_info(pf.left_indent)
        format_info['left_indent'] = {
            'type': '左缩进',
            'value': value,
            'unit': unit_name,
            'unit_code': unit_code
        }
    else:
        format_info['left_indent'] = {
            'type': '左缩进',
            'value': 0,
            'unit': '磅',
            'unit_code': 'pt'
        }
    
    # 文本之后 (After text) - 右缩进
    if pf.right_indent:
        value, unit_name, unit_code = get_measurement_info(pf.right_indent)
        format_info['right_indent'] = {
            'type': '右缩进',
            'value': value,
            'unit': unit_name,
            'unit_code': unit_code
        }
    else:
        format_info['right_indent'] = {
            'type': '右缩进',
            'value': 0,
            'unit': '磅',
            'unit_code': 'pt'
        }
    
    # 特殊格式 (Special) - 首行缩进
    if pf.first_line_indent:
        value, unit_name, unit_code = get_measurement_info(pf.first_line_indent)
        if value > 0:
            format_info['first_line_indent'] = {
                'type': '首行缩进',
                'value': value,
                'unit': unit_name,
                'unit_code': unit_code
            }
        elif value < 0:
            format_info['hanging_indent'] = {
                'type': '悬挂缩进',
                'value': abs(value),
                'unit': unit_name,
                'unit_code': unit_code
            }
    else:
        format_info['first_line_indent'] = {
            'type': '首行缩进',
            'value': 0,
            'unit': '磅',
            'unit_code': 'pt'
        }
    
    # 3. 间距 (Spacing) Section
    # 段前 (Before paragraph)
    if pf.space_before:
        value, unit_name, unit_code = get_measurement_info(pf.space_before)
        format_info['space_before'] = {
            'type': '段前间距',
            'value': value,
            'unit': unit_name,
            'unit_code': unit_code
        }
    else:
        format_info['space_before'] = {
            'type': '段前间距',
            'value': 0,
            'unit': '磅',
            'unit_code': 'pt'
        }
    
    # 段后 (After paragraph)
    if pf.space_after:
        value, unit_name, unit_code = get_measurement_info(pf.space_after)
        format_info['space_after'] = {
            'type': '段后间距',
            'value': value,
            'unit': unit_name,
            'unit_code': unit_code
        }
    else:
        format_info['space_after'] = {
            'type': '段后间距',
            'value': 0,
            'unit': '磅',
            'unit_code': 'pt'
        }
    
    # 行距 (Line spacing) - 处理不同类型的行距
    line_spacing_rule_map = {
        WD_LINE_SPACING.SINGLE: '单倍行距',
        WD_LINE_SPACING.ONE_POINT_FIVE: '1.5倍行距',
        WD_LINE_SPACING.DOUBLE: '2倍行距',
        WD_LINE_SPACING.AT_LEAST: '最小值',
        WD_LINE_SPACING.EXACTLY: '固定值',
        WD_LINE_SPACING.MULTIPLE: '多倍行距'
    }
    
    if pf.line_spacing:
        if hasattr(pf, 'line_spacing_rule') and pf.line_spacing_rule is not None:
            rule_name = line_spacing_rule_map.get(pf.line_spacing_rule, str(pf.line_spacing_rule))
            
            if pf.line_spacing_rule in [WD_LINE_SPACING.SINGLE, WD_LINE_SPACING.ONE_POINT_FIVE, WD_LINE_SPACING.DOUBLE, WD_LINE_SPACING.MULTIPLE]:
                # 倍数行距
                format_info['line_spacing'] = {
                    'type': '行距',
                    'rule': rule_name,
                    'value': pf.line_spacing,
                    'unit': '倍数',
                    'unit_code': 'multiple'
                }
            else:
                # 固定值或最小值
                value, unit_name, unit_code = get_measurement_info(pf.line_spacing)
                format_info['line_spacing'] = {
                    'type': '行距',
                    'rule': rule_name,
                    'value': value,
                    'unit': unit_name,
                    'unit_code': unit_code
                }
        else:
            # 如果没有行距规则，尝试获取数值和单位
            value, unit_name, unit_code = get_measurement_info(pf.line_spacing)
            format_info['line_spacing'] = {
                'type': '行距',
                'rule': '固定值',
                'value': value,
                'unit': unit_name,
                'unit_code': unit_code
            }
    else:
        format_info['line_spacing'] = {
            'type': '行距',
            'rule': '单倍行距',
            'value': 1.0,
            'unit': '倍数',
            'unit_code': 'multiple'
        }
    
    return format_info

def infer_heading_level(content, font, size, bold, outline_level):
    """Improved heading level inference based on content patterns and formatting."""
    
    # Check for Chinese numbering patterns
    content_stripped = content.strip()
    
    # Level 1: "一、", "二、", "三、" etc.
    if content_stripped and content_stripped[0] in '一二三四五六七八九十' and '、' in content_stripped[:3]:
        return 'heading1'
    
    # Level 2: "（一）", "（二）" etc.
    if content_stripped.startswith('（') and content_stripped[1] in '一二三四五六七八九十' and '）' in content_stripped:
        return 'heading2'
    
    # Level 3: "1.", "2." etc.
    if content_stripped and content_stripped[0].isdigit() and '.' in content_stripped[:3]:
        return 'heading3'
    
    # Level 4: "（1）", "（2）" etc.
    if content_stripped.startswith('（') and content_stripped[1].isdigit() and '）' in content_stripped:
        return 'heading4'
    
    # Check outline level from Word
    if outline_level and outline_level.get('value') != '正文文本':
        outline_value = outline_level.get('value', '')
        if outline_value == '标题1':
            return 'heading1'
        elif outline_value == '标题2':
            return 'heading2'
        elif outline_value == '标题3':
            return 'heading3'
        elif outline_value == '标题4':
            return 'heading4'
    
    # Check font characteristics
    if bold and size >= 16:
        return 'heading1'
    elif bold and size >= 14:
        return 'heading2'
    elif bold and size >= 12:
        return 'heading3'
    
    return 'paragraph'

def parse_docx_to_tree(file_path):
    doc = Document(file_path)
    root = Node('root', 'Document Root')
    stack = [root]
    
    # Step 1: Collect font size distribution
    size_distribution = get_font_size_distribution(doc)
    
    # Step 2: Parse paragraphs and build tree
    for paragraph in doc.paragraphs:
        if not paragraph.text.strip():  # Skip empty paragraphs
            continue
        
        # Extract properties
        first_run = paragraph.runs[0] if paragraph.runs else None
        font = first_run.font.name if first_run and first_run.font.name else 'Default'
        size = first_run.font.size.pt if first_run and first_run.font.size else 12.0
        bold = first_run.font.bold if first_run and first_run.font.bold else False
        
        # Get paragraph formatting
        format_info = extract_paragraph_formatting(paragraph)
        
        # 处理间距信息
        spacing_info = format_info.get('space_after', {})
        spacing = spacing_info.get('value', 0.0) if isinstance(spacing_info, dict) else 0.0
        
        # 处理行距信息
        line_spacing_info = format_info.get('line_spacing', {})
        line_spacing = line_spacing_info.get('value', None) if isinstance(line_spacing_info, dict) else None
        
        # Get outline level
        outline_level = format_info.get('outline_level', None)
        
        # Get alignment
        alignment = format_info.get('alignment', None)
        
        # Get direction
        direction = format_info.get('direction', None)
        
        # Infer node type based on content and formatting
        node_type = infer_heading_level(paragraph.text, font, size, bold, outline_level)
        content = paragraph.text.strip()
        
        # Create indentation info
        indentation = {}
        if 'left_indent' in format_info:
            indentation['left'] = format_info['left_indent']
        if 'right_indent' in format_info:
            indentation['right'] = format_info['right_indent']
        if 'first_line_indent' in format_info:
            indentation['first_line'] = format_info['first_line_indent']
        if 'hanging_indent' in format_info:
            indentation['hanging'] = format_info['hanging_indent']
        
        new_node = Node(
            node_type, content, font, size, bold, spacing, 
            line_spacing, indentation, outline_level, alignment, direction, format_info
        )
        
        # Build tree: Adjust stack based on level
        current_level = {'heading1': 1, 'heading2': 2, 'heading3': 3, 'heading4': 4, 'paragraph': 5}.get(node_type, 5)
        while len(stack) > 1 and current_level <= {'heading1': 1, 'heading2': 2, 'heading3': 3, 'heading4': 4, 'paragraph': 5}.get(stack[-1].type, 5):
            stack.pop()
        stack[-1].add_child(new_node)
        stack.append(new_node)
    
    return root

# Print as tree string
def print_tree(node, indent=0):
    spacing_info = f"spacing={node.spacing}" if node.spacing else ""
    line_spacing_info = f"line_spacing={node.line_spacing}" if node.line_spacing else ""
    indent_info = f"indent={node.indentation}" if node.indentation else ""
    alignment_info = f"alignment={node.alignment['value'] if node.alignment else 'None'}"
    outline_info = f"outline={node.outline_level['value'] if node.outline_level else 'None'}"
    font_size_info = f"字号={node.font_size_name}" if node.font_size_name else ""
    
    info_parts = [f"font={node.font}", f"size={node.size}pt", f"bold={node.bold}"]
    if font_size_info:
        info_parts.append(font_size_info)
    if spacing_info:
        info_parts.append(spacing_info)
    if line_spacing_info:
        info_parts.append(line_spacing_info)
    if indent_info:
        info_parts.append(indent_info)
    if alignment_info:
        info_parts.append(alignment_info)
    if outline_info:
        info_parts.append(outline_info)
    
    info_str = ", ".join(info_parts)
    print('  ' * indent + f"{node.type}: {node.content} ({info_str})")
    for child in node.children:
        print_tree(child, indent + 1)

# Example usage
if __name__ == "__main__":
    file_path = '1.docx'  # Replace with your file path
    tree_root = parse_docx_to_tree(file_path)

    # Print tree as JSON
    print("正在生成JSON输出...")
    json_output = json.dumps(tree_root.to_dict(), indent=4, ensure_ascii=False)
    
    # Save to JSON file
    with open('tree_output.json', 'w', encoding='utf-8') as f:
        f.write(json_output)
    
    print("JSON文件已保存为 'tree_output.json'")
    
    # Print tree structure
    print("\n" + "="*50)
    print("TREE STRUCTURE:")
    print("="*50)
    print_tree(tree_root)
    
    print(f"\n解析完成! JSON文件已保存。")
    print(f"共解析了 {len([n for n in tree_root.children])} 个顶级节点。")
   
    check_file('tree_output.json')