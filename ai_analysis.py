import requests
import time
from typing import Dict, Any, Optional
from dataclasses import dataclass

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

class DeepSeekAnalyzer:
    """DeepSeek R1模型分析器"""
    
    def __init__(self, api_key: str, api_url: str = "https://api.deepseek.com/v1/chat/completions"):
        self.api_key = api_key
        self.api_url = api_url
        self.headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        }
    
    def analyze_node(self, node_info: str, context: str, format_requirements: str) -> str:
        """使用DeepSeek R1分析节点格式"""
        
        prompt = f"""你是一个专业的文档格式检查专家。请分析以下文档节点的格式是否符合要求。

文档格式要求：
{format_requirements}

当前节点信息：
{node_info}

上下文信息：
{context}

请从以下几个方面进行分析：
1. 节点类型识别（是否为拟稿部门、日期等）
2. 格式合规性检查（字体、字号、对齐、缩进等）
3. 位置关系检查（是否符合文档结构要求）
4. 具体问题描述和建议

请用中文回答，格式要清晰易读。"""

        payload = {
            "model": "deepseek-chat",
            "messages": [
                {
                    "role": "user",
                    "content": prompt
                }
            ],
            "temperature": 0.3,
            "max_tokens": 2000
        }
        
        try:
            response = requests.post(
                self.api_url,
                headers=self.headers,
                json=payload,
                timeout=30
            )
            
            if response.status_code == 200:
                result = response.json()
                return result['choices'][0]['message']['content']
            else:
                return f"API调用失败: {response.status_code} - {response.text}"
                
        except Exception as e:
            return f"分析过程中出现错误: {str(e)}"
    
    def analyze_batch(self, nodes_data: list, format_requirements: str, delay: float = 1.0) -> list:
        """批量分析多个节点"""
        results = []
        
        for i, node_data in enumerate(nodes_data):
            print(f"正在分析节点 {i+1}/{len(nodes_data)}...")
            
            analysis = self.analyze_node(
                node_data['node_info'],
                node_data['context'],
                format_requirements
            )
            
            results.append({
                'node_index': node_data['node_index'],
                'analysis': analysis
            })
            
            # 添加延迟避免API限制
            if i < len(nodes_data) - 1:
                time.sleep(delay)
        
        return results

class DocumentAnalyzer:
    """文档分析器 - 整合节点处理和AI分析"""
    
    def __init__(self, api_key: str):
        self.analyzer = DeepSeekAnalyzer(api_key)
    
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
    
    def generate_node_info(self, node: NodeInfo) -> str:
        """生成当前节点的详细信息字符串"""
        node_info = f"""当前节点详细信息：
内容: {node.content}
类型: {node.type}
字体: {node.font}
字号: {node.font_size_name} ({node.size}pt)
加粗: {node.bold}
对齐: {node.alignment.get('value') if node.alignment else 'None'}
行距: {node.line_spacing}
段后间距: {node.spacing}"""

        if node.paragraph_format:
            pf = node.paragraph_format
            node_info += f"""
首行缩进: {pf.get('first_line_indent', {}).get('value', 0)}
左缩进: {pf.get('left_indent', {}).get('value', 0)}
右缩进: {pf.get('right_indent', {}).get('value', 0)}"""

        return node_info
    
    def analyze_single_node(self, node: NodeInfo, context: str, format_requirements: str) -> str:
        """分析单个节点"""
        node_info = self.generate_node_info(node)
        return self.analyzer.analyze_node(node_info, context, format_requirements)
    
    def analyze_nodes_with_context(self, nodes: list, context_generator, format_requirements: str) -> list:
        """分析带上下文的节点列表"""
        nodes_data = []
        
        for node in nodes:
            context = context_generator(node)
            node_info = self.generate_node_info(node)
            
            nodes_data.append({
                'node_index': getattr(node, 'index', 0),
                'node_info': node_info,
                'context': context
            })
        
        return self.analyzer.analyze_batch(nodes_data, format_requirements)

# 配置常量
DEEPSEEK_API_KEY = "your_api_key_here"  # 请替换为您的API密钥
DEEPSEEK_API_URL = "https://api.deepseek.com/v1/chat/completions"

def create_analyzer(api_key: str = None) -> Optional[DocumentAnalyzer]:
    """创建文档分析器实例"""
    if api_key is None:
        api_key = DEEPSEEK_API_KEY
    
    if api_key == "your_api_key_here":
        print("警告: 请设置正确的DeepSeek API密钥")
        return None
    
    return DocumentAnalyzer(api_key)

# 示例使用
if __name__ == "__main__":
    # 测试AI分析功能
    analyzer = create_analyzer()
    
    if analyzer:
        # 创建测试节点
        test_node = NodeInfo(
            type="paragraph",
            content="测试内容",
            font="宋体",
            size=12.0,
            font_size_name="小四",
            bold=False,
            spacing=0,
            line_spacing=1.0,
            indentation=None,
            outline_level=None,
            alignment={"value": "左对齐"},
            direction=None,
            paragraph_format=None,
            font_color=None
        )
        
        # 测试分析
        test_context = "这是测试上下文信息"
        test_requirements = "测试格式要求"
        
        result = analyzer.analyze_single_node(test_node, test_context, test_requirements)
        print("测试分析结果:")
        print(result)
    else:
        print("无法创建分析器，请检查API密钥配置")
