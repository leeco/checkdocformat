import json
from typing import List, Dict, Any, Optional
from dataclasses import dataclass

# 导入AI分析模块
from ai_analysis import create_analyzer

# 全局变量：控制上下文范围
CONTEXT_BEFORE_NODES = 3  # 前N个节点
CONTEXT_AFTER_NODES = 2   # 后M个节点

# 格式要求定义
PROMPTS = "拟稿部门或单位：没有附件时，则正文下空一行，右空四字；有附件时，则附件下空二行，右空四字。日期：用阿拉伯数字将年、月、日标全，年份应标全称，月、日不编虚位（即1不编为01），另起一行，位于拟稿部门或单位下方正中间。"

@dataclass
class NodeInfo:
    """节点信息数据类"""
    type: str
    content: str
    font: Optional[str]
    size: Optional[float]
    font_size_name: Optional[str]
    bold: bool
    spacing: Optional[float]
    line_spacing: Optional[float]
    indentation: Optional[Dict]
    outline_level: Optional[Dict]
    alignment: Optional[Dict]
    direction: Optional[Dict]
    paragraph_format: Optional[Dict]
    font_color: Optional[Dict] = None
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'NodeInfo':
        """从字典创建节点信息"""
        return cls(
            type=data.get('type'),
            content=data.get('content', ''),
            font=data.get('font'),
            size=data.get('size'),
            font_size_name=data.get('font_size_name'),
            bold=data.get('bold', False),
            spacing=data.get('spacing'),
            line_spacing=data.get('line_spacing'),
            indentation=data.get('indentation'),
            outline_level=data.get('outline_level'),
            alignment=data.get('alignment'),
            direction=data.get('direction'),
            paragraph_format=data.get('paragraph_format'),
            font_color=data.get('font_color')
        )

class DocumentAnalyzer:
    """文档AI分析器"""
    
    def __init__(self, json_file: str, api_key: str = None):
        self.json_file = json_file
        self.nodes: List[NodeInfo] = []
        self.ai_analyzer = create_analyzer(api_key)
        self.load_nodes()
    
    def load_nodes(self):
        """加载并解析JSON文件中的节点"""
        with open(self.json_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        self._extract_nodes(data)
    
    def _extract_nodes(self, node_data: Dict[str, Any]):
        """递归提取所有节点"""
        # 跳过根节点
        if node_data.get('type') != 'root':
            self.nodes.append(NodeInfo.from_dict(node_data))
        
        # 递归处理子节点
        for child in node_data.get('children', []):
            self._extract_nodes(child)
    
    def generate_context_string(self, current_index: int) -> str:
        """生成包含前后节点的完整上下文字符串"""
        context_str = "=== 上下文信息 ===\n"
        
        # 前N个节点
        start_before = max(0, current_index - CONTEXT_BEFORE_NODES)
        before_nodes = self.nodes[start_before:current_index]
        
        if before_nodes:
            context_str += f"\n前{CONTEXT_BEFORE_NODES}个节点:\n"
            for i, node in enumerate(before_nodes):
                context_str += f"\n前节点{i+1}:\n"
                context_str += self.format_node_details(node, "  ")
        else:
            context_str += "\n前节点: 无\n"
        
        # 当前节点
        current_node = self.nodes[current_index]
        context_str += f"\n当前节点:\n"
        context_str += self.format_node_details(current_node, "  ")
        
        # 后M个节点
        end_after = min(len(self.nodes), current_index + 1 + CONTEXT_AFTER_NODES)
        after_nodes = self.nodes[current_index + 1:end_after]
        
        if after_nodes:
            context_str += f"\n后{CONTEXT_AFTER_NODES}个节点:\n"
            for i, node in enumerate(after_nodes):
                context_str += f"\n后节点{i+1}:\n"
                context_str += self.format_node_details(node, "  ")
        else:
            context_str += "\n后节点: 无\n"
        
        return context_str
    
    def format_node_details(self, node: NodeInfo, prefix: str = "") -> str:
        """格式化节点详细信息"""
        details = f"{prefix}节点内容: {node.content}\n"
        details += f"{prefix}节点类型: {node.type}\n"
        details += f"{prefix}字体: {node.font}\n"
        details += f"{prefix}字号: {node.font_size_name} ({node.size}pt)\n"
        details += f"{prefix}加粗: {node.bold}\n"
        details += f"{prefix}行距: {node.line_spacing}\n"
        details += f"{prefix}段后间距: {node.spacing}\n"
        
        if node.alignment:
            details += f"{prefix}对齐方式: {node.alignment.get('value')}\n"
        
        if node.paragraph_format:
            pf = node.paragraph_format
            details += f"{prefix}首行缩进: {pf.get('first_line_indent', {}).get('value', 0)}\n"
            details += f"{prefix}左缩进: {pf.get('left_indent', {}).get('value', 0)}\n"
            details += f"{prefix}右缩进: {pf.get('right_indent', {}).get('value', 0)}\n"
            details += f"{prefix}段前间距: {pf.get('space_before', {}).get('value', 0)}\n"
            details += f"{prefix}段后间距: {pf.get('space_after', {}).get('value', 0)}\n"
        
        if node.font_color:
            details += f"{prefix}字体颜色: {node.font_color.get('value', 'None')}\n"
        
        return details
    
    def analyze_all_nodes(self) -> List[Dict[str, Any]]:
        """对所有节点进行AI分析"""
        if not self.ai_analyzer:
            print("DeepSeek API未配置，无法进行AI分析")
            return []
        
        results = []
        
        for i, node in enumerate(self.nodes):
            print(f"正在分析节点 {i+1}/{len(self.nodes)}: {node.content[:50]}...")
            
            # 生成上下文
            context = self.generate_context_string(i)
            
            # 进行AI分析
            analysis = self.ai_analyzer.analyze_single_node(node, context, PROMPTS)
            
            results.append({
                'node_index': i,
                'content': node.content,
                'analysis': analysis
            })
        
        return results

def main():
    """主函数"""
    # 检查API密钥
    api_key = "sk-cf7edd378f41409b8270cbca5baef81b"  # 请替换为您的API密钥
    
    # 创建分析器
    analyzer = DocumentAnalyzer('tree_output.json', api_key)
    
    print(f"总共解析到 {len(analyzer.nodes)} 个节点")
    print(f"上下文设置：前{CONTEXT_BEFORE_NODES}个节点，后{CONTEXT_AFTER_NODES}个节点")
    print(f"AI分析: {'启用' if api_key else '未启用'}\n")
    
    # 对所有节点进行AI分析
    if api_key:
        print("=== 开始AI分析所有节点 ===")
        results = analyzer.analyze_all_nodes()
        
        print("\n=== AI分析结果 ===")
        for i, result in enumerate(results):
            print(f"\n节点 {result['node_index']+1}: {result['content']}")
            print("="*80)
            print(result['analysis'])
            print("="*80)
            
            # 添加分隔符
            if i < len(results) - 1:
                print("\n" + "-"*100 + "\n")
    else:
        print("请设置API密钥以启用AI分析功能")

if __name__ == "__main__":
    main()