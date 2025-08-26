from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Inches, Cm, Mm, Pt, Emu
import json
import numpy as np  # 用于统计分析

def pt_to_font_size(pt_size):
    """将磅值转换为中文标准字号"""
    # 参考《GB/T 15834-2011 标点符号用法》和常用中文字号标准
    font_size_map = {
        42: '初号',      # 42pt
        36: '小初',      # 36pt
        32: '一号',      # 32pt
        28: '小一',      # 28pt
        24: '二号',      # 24pt
        22: '小二',      # 22pt
        18: '三号',      # 18pt
        16: '小三',      # 16pt
        15: '四号',      # 15pt
        14: '小四',      # 14pt
        12: '五号',      # 12pt
        10.5: '小五',    # 10.5pt
        9: '六号',       # 9pt
        7.5: '小六',     # 7.5pt
        5.5: '七号',     # 5.5pt
        5: '八号',       # 5pt
    }
    # 找到最接近的字号
    closest_size = min(font_size_map.keys(), key=lambda x: abs(x - pt_size))
    return font_size_map[closest_size]
    
    # Find the closest font size
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

def extract_paragraph_formatting(paragraph):
    """Extract detailed formatting information from paragraph including all properties from the dialog."""
    format_info = {}
    
    # Get paragraph format
    pf = paragraph.paragraph_format
    
    # 1. 常规 (General) Section
    # Alignment (对齐方式) - 使用正确的Word对齐常量
    alignment_map = {
        WD_ALIGN_PARAGRAPH.LEFT: '左对齐',
        WD_ALIGN_PARAGRAPH.CENTER: '居中',
        WD_ALIGN_PARAGRAPH.RIGHT: '右对齐',
        WD_ALIGN_PARAGRAPH.JUSTIFY: '两端对齐',
        WD_ALIGN_PARAGRAPH.DISTRIBUTE: '分散对齐',
        WD_ALIGN_PARAGRAPH.THAI_JUSTIFY: '泰文两端对齐',
        WD_ALIGN_PARAGRAPH.JUSTIFY_MED: '两端对齐(中等)',
        WD_ALIGN_PARAGRAPH.JUSTIFY_HI: '两端对齐(高)',
        WD_ALIGN_PARAGRAPH.JUSTIFY_LOW: '两端对齐(低)'
    }
    
    if pf.alignment is not None:
        format_info['alignment'] = {
            'type': '对齐方式',
            'value': alignment_map.get(pf.alignment, str(pf.alignment)),
            'raw_value': str(pf.alignment)
        }
    else:
        format_info['alignment'] = {
            'type': '对齐方式',
            'value': '左对齐',
            'raw_value': 'None'
        }
    
    # Outline Level (大纲级别)
    outline_map = {
        0: '标题1',
        1: '标题2', 
        2: '标题3',
        3: '标题4',
        4: '标题5',
        5: '标题6',
        6: '标题7',
        7: '标题8',
        8: '标题9',
        9: '正文文本'
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

# Example usage
file_path = 'example.docx'  # Replace with your file path
tree_root = parse_docx_to_tree(file_path)

# Print tree as JSON
print(json.dumps(tree_root.to_dict(), indent=4, ensure_ascii=False))

# Print as tree string
def print_tree(node, indent=0):
    spacing_info = f"spacing={node.spacing}" if node.spacing else ""
    line_spacing_info = f"line_spacing={node.line_spacing}" if node.line_spacing else ""
    indent_info = f"indent={node.indentation}" if node.indentation else ""
    alignment_info = f"alignment={node.alignment}" if node.alignment else ""
    outline_info = f"outline={node.outline_level}" if node.outline_level else ""
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

print("\n" + "="*50)
print("TREE STRUCTURE:")
print("="*50)
print_tree(tree_root)

# Save to JSON file
with open('tree_output.json', 'w', encoding='utf-8') as f:
    json.dump(tree_root.to_dict(), f, ensure_ascii=False, indent=4)