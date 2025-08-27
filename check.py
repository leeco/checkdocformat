import json
from typing import List, Dict, Any, Optional
from dataclasses import dataclass

# 导入AI分析模块
from ai_analysis import create_analyzer

# 全局变量：控制上下文范围
CONTEXT_BEFORE_NODES = 5  # 前A个节点
CONTEXT_AFTER_NODES = 5   # 后B个节点

# 全局变量：控制批量分析
BATCH_SIZE = 5  # 一次性分析的相邻节点数N，默认为1（单节点分析）

# 格式要求定义
PROMPTS = """
格式具体 要求如下：
一、标题
由单位名称、事由和文种组成。其中单位名称可视情况省略，但事由和文种必备；字体为方正小标宋简体，字号为二号且不加粗。(补充：标题中发文机关字体为方正小标宋简体，字号为小初且不加粗。)
二、正文
字体为仿宋_GB2312，字号为三号且不加粗，编排于主送机关名称下一行，每个自然段左空二字，回行顶格。一般每面排22行,每行排28个字，特殊情况除外。
三、公文的结构层次序数
文中结构层次序数依次可以用“一、”“（一）”“1.”“（1）”标注；一般第一层用黑体三号字不加粗、第二层用楷体三号字、第三层和第四层用仿宋_GB2312三号字标注。
实际行文中，如果不需要全部四个层次，只有两个层次时，可按顺序跳用，第一层使用“一、”，第二层可使用（一）或“1.”表示，只有一个层次时，使用“一、”
四、结尾
特此报告\请示\申请。另起一行，左空二字，字体为仿宋_GB2312三号。
五、落款
拟稿部门或单位：没有附件时，则正文下空一行，右空四字；有附件时，则附件下空二行，右空四字。日期：用阿拉伯数字将年、月、日标全，年份应标全称，月、日不编虚位（即1不编为01），另起一行，位于拟稿部门或单位下方正中间。
六、附件说明
正文下空一行左空二字编排“附件”二字，后标全角冒号和附件名称。如有多个附件，使用阿拉伯数字标注附件顺序号（如：“附件：1.XXXX”）；附件名称后不加标点符号。附件名称较长需回行时，应当与上一行附件名称的首字对齐。字体为仿宋_GB2312三号）
"""


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
        """递归提取所有节点，包括空行"""
        # 跳过根节点，但保留空行节点用于上下文分析
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
    
    def generate_batch_context_string(self, start_index: int, batch_size: int) -> str:
        """生成批量节点分析的上下文字符串"""
        context_str = "=== 批量分析上下文信息 ===\n"
        
        # 前A个节点
        start_before = max(0, start_index - CONTEXT_BEFORE_NODES)
        before_nodes = self.nodes[start_before:start_index]
        
        if before_nodes:
            context_str += f"\n前{CONTEXT_BEFORE_NODES}个节点:\n"
            for i, node in enumerate(before_nodes):
                context_str += f"\n前节点{i+1}:\n"
                context_str += self.format_node_details(node, "  ")
        else:
            context_str += "\n前节点: 无\n"
        
        # 当前批量节点
        end_index = min(len(self.nodes), start_index + batch_size)
        current_batch = self.nodes[start_index:end_index]
        
        context_str += f"\n当前批量分析节点 (共{len(current_batch)}个):\n"
        for i, node in enumerate(current_batch):
            context_str += f"\n批量节点{i+1}:\n"
            context_str += self.format_node_details(node, "  ")
        
        # 后B个节点
        end_after = min(len(self.nodes), end_index + CONTEXT_AFTER_NODES)
        after_nodes = self.nodes[end_index:end_after]
        
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
        """对所有节点进行批量AI分析，跳过空行节点"""
        if not self.ai_analyzer:
            print("DeepSeek API未配置，无法进行AI分析")
            return []
        
        results = []
        analyzed_count = 0
        skipped_count = 0
        i = 0
        
        while i < len(self.nodes):
            # 检查当前批次是否包含非空行节点
            batch_end = min(len(self.nodes), i + BATCH_SIZE)
            current_batch = self.nodes[i:batch_end]
            
            # 过滤出非空行节点
            non_empty_nodes = [node for node in current_batch if node.type != 'empty_line']
            
            if not non_empty_nodes:
                # 整个批次都是空行，跳过
                for j, node in enumerate(current_batch):
                    node_index = i + j
                    print(f"跳过空行节点 {node_index+1}/{len(self.nodes)}: [空行]")
                    skipped_count += 1
                    results.append({
                        'node_index': node_index,
                        'content': node.content,
                        'analysis': '[跳过AI分析 - 空行节点]',
                        'skipped': True,
                        'batch_info': {
                            'batch_start': i,
                            'batch_size': len(current_batch),
                            'batch_position': j
                        }
                    })
                i = batch_end
                continue
            
            # 处理包含非空行节点的批次
            if BATCH_SIZE == 1:
                # 单节点分析模式（原有逻辑）
                node = current_batch[0]
                if node.type == 'empty_line':
                    print(f"跳过空行节点 {i+1}/{len(self.nodes)}: [空行]")
                    skipped_count += 1
                    results.append({
                        'node_index': i,
                        'content': node.content,
                        'analysis': '[跳过AI分析 - 空行节点]',
                        'skipped': True,
                        'batch_info': {
                            'batch_start': i,
                            'batch_size': 1,
                            'batch_position': 0
                        }
                    })
                else:
                    analyzed_count += 1
                    print(f"正在分析节点 {i+1}/{len(self.nodes)}: {node.content[:50]}...")
                    
                    # 生成单节点上下文
                    context = self.generate_context_string(i)
                    
                    # 进行AI分析
                    analysis = self.ai_analyzer.analyze_single_node(node, context, PROMPTS)
                    
                    results.append({
                        'node_index': i,
                        'content': node.content,
                        'analysis': analysis,
                        'skipped': False,
                        'batch_info': {
                            'batch_start': i,
                            'batch_size': 1,
                            'batch_position': 0
                        }
                    })
                i += 1
            else:
                # 批量分析模式
                analyzed_count += len(non_empty_nodes)
                batch_content_preview = " | ".join([node.content[:20] + "..." for node in non_empty_nodes[:3]])
                print(f"正在批量分析节点 {i+1}-{batch_end}/{len(self.nodes)} (共{len(current_batch)}个，{len(non_empty_nodes)}个非空): {batch_content_preview}")
                
                # 生成批量上下文
                context = self.generate_batch_context_string(i, len(current_batch))
                
                # 进行批量AI分析
                batch_analysis = self.ai_analyzer.analyze_batch_nodes(current_batch, context, PROMPTS)
                
                # 处理批量分析结果
                for j, node in enumerate(current_batch):
                    node_index = i + j
                    if node.type == 'empty_line':
                        skipped_count += 1
                        results.append({
                            'node_index': node_index,
                            'content': node.content,
                            'analysis': '[跳过AI分析 - 空行节点]',
                            'skipped': True,
                            'batch_info': {
                                'batch_start': i,
                                'batch_size': len(current_batch),
                                'batch_position': j
                            }
                        })
                    else:
                        # 从批量分析结果中提取对应节点的分析
                        node_analysis = self._extract_node_analysis_from_batch(batch_analysis, j, node)
                        results.append({
                            'node_index': node_index,
                            'content': node.content,
                            'analysis': node_analysis,
                            'skipped': False,
                            'batch_info': {
                                'batch_start': i,
                                'batch_size': len(current_batch),
                                'batch_position': j,
                                'batch_analysis': batch_analysis
                            }
                        })
                
                i = batch_end
        
        print(f"\n分析完成统计: 总节点{len(self.nodes)}个, AI分析{analyzed_count}个, 跳过空行{skipped_count}个")
        if BATCH_SIZE > 1:
            print(f"批量分析模式: 每批{BATCH_SIZE}个节点")
        return results
    
    def _extract_node_analysis_from_batch(self, batch_analysis: str, node_position: int, node: NodeInfo) -> str:
        """从批量分析结果中提取特定节点的分析结果"""
        try:
            # 尝试按节点分割批量分析结果
            lines = batch_analysis.split('\n')
            
            # 查找节点标识符
            node_markers = [
                f"批量节点{node_position + 1}",
                f"节点{node_position + 1}",
                f"第{node_position + 1}个节点",
                f"{node_position + 1}.",
                f"({node_position + 1})"
            ]
            
            # 查找节点内容标识符
            content_markers = [
                node.content[:30],  # 节点内容前30字符
                node.content[:20],  # 节点内容前20字符
                node.content[:10]   # 节点内容前10字符
            ]
            
            start_line = -1
            end_line = len(lines)
            
            # 查找当前节点的开始位置
            for i, line in enumerate(lines):
                for marker in node_markers + content_markers:
                    if marker in line:
                        start_line = i
                        break
                if start_line != -1:
                    break
            
            # 查找下一个节点的开始位置（作为结束位置）
            if start_line != -1:
                next_node_markers = [
                    f"批量节点{node_position + 2}",
                    f"节点{node_position + 2}",
                    f"第{node_position + 2}个节点",
                    f"{node_position + 2}.",
                    f"({node_position + 2})"
                ]
                
                for i in range(start_line + 1, len(lines)):
                    for marker in next_node_markers:
                        if marker in lines[i]:
                            end_line = i
                            break
                    if end_line != len(lines):
                        break
            
            # 提取节点分析内容
            if start_line != -1:
                node_analysis_lines = lines[start_line:end_line]
                node_analysis = '\n'.join(node_analysis_lines).strip()
                
                # 如果提取的内容为空或太短，返回完整的批量分析
                if len(node_analysis) < 50:
                    return f"[批量分析结果 - 节点{node_position + 1}]\n{batch_analysis}"
                
                return node_analysis
            else:
                # 如果无法精确提取，返回带标记的完整批量分析
                return f"[批量分析结果 - 节点{node_position + 1}: {node.content[:30]}...]\n{batch_analysis}"
                
        except Exception as e:
            # 如果提取失败，返回完整的批量分析结果
            return f"[批量分析结果提取失败 - 节点{node_position + 1}: {str(e)}]\n{batch_analysis}"

def main():
    """主函数"""
    # 直接调用 check_file 实现主流程
    file_path = 'tree_output.json'
    check_file(file_path)

def check_file(file_path: str):
    """检查文件"""
    print(f"开始分析文件: {file_path}")
    analyzer = DocumentAnalyzer(file_path, api_key="sk-cf7edd378f41409b8270cbca5baef81b")
    print(f"已加载 {len(analyzer.nodes)} 个节点（包含空行），准备进行AI分析...")
    results = analyzer.analyze_all_nodes()
    print("\n=== AI分析结果 ===")
    last_result = None
    ai_analyzed_count = 0
    
    for i, result in enumerate(results):
        # 检查是否是跳过的空行节点
        if result.get('skipped', False):
            print(f"\n节点 {result['node_index']+1}: {result['content']} [已跳过]")
            continue
        
        ai_analyzed_count += 1
        print(f"\n节点 {result['node_index']+1}: {result['content']}")
        print("="*80)
        print(result['analysis'])
        print("="*80)
        
        # 只在AI分析的节点之间添加分隔符
        next_non_skipped = False
        for j in range(i + 1, len(results)):
            if not results[j].get('skipped', False):
                next_non_skipped = True
                break
        
        if next_non_skipped:
            print("\n" + "-"*100 + "\n")
        
        last_result = result
    
    print(f"\n所有节点分析完成。AI分析了 {ai_analyzed_count} 个有效节点。")
    return last_result

def set_batch_size(size: int):
    """设置批量分析的节点数"""
    global BATCH_SIZE
    BATCH_SIZE = max(1, size)  # 确保批量大小至少为1
    print(f"批量分析大小已设置为: {BATCH_SIZE}")

def set_context_range(before: int, after: int):
    """设置上下文范围"""
    global CONTEXT_BEFORE_NODES, CONTEXT_AFTER_NODES
    CONTEXT_BEFORE_NODES = max(0, before)
    CONTEXT_AFTER_NODES = max(0, after)
    print(f"上下文范围已设置为: 前{CONTEXT_BEFORE_NODES}个节点, 后{CONTEXT_AFTER_NODES}个节点")

if __name__ == "__main__":
    main()