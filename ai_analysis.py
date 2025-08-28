import requests
import time
import json
from typing import Dict, Any, Optional, Generator
from dataclasses import dataclass


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


# 全局变量：控制流式输出
ENABLE_STREAMING = True  # 是否启用流式输出
STREAMING_DELAY = 0.05  # 流式输出延迟（秒）

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
    
    def analyze_node_streaming(self, node_info: str, context: str) -> Generator[str, None, None]:
        """使用DeepSeek R1进行流式分析"""
        # 在函数内部获取格式要求，不再作为参数传入
        format_requirements = PROMPTS
        
        prompt = f"""你是一个专业的文档格式检查专家。请分析以下文档节点的格式是否符合要求。
当前节点信息：
{node_info}
--------------------------------
上下文信息：
{context}
--------------------------------
文档格式要求：
{format_requirements}
请从以下几个方面进行分析：
1. 节点类型识别（是否为拟稿部门、日期等）
2. 格式合规性检查（字体、字号、对齐、缩进等）
    -对于类型为标题的节点，可能输入的格式是错误的，需要根据节点内容和长度判断是否为标题，如果是标题，则需要判断是否符合标题的格式要求，如果不是标题，则需要根据其内容判断是否符合相关内容对格式要求。
3. 位置关系检查（是否符合文档结构要求）
4. 具体问题描述和建议
6. 需要忽略分析的节点：
    -分割线
    -页眉页脚
    -加载两分割线之间的，且可能为落款的内容部分

请根据上述要求，逐条进行审查，输出应简明扼要、结构清晰，仅限于要求范围内的信息，避免添加无关内容。
如发现不符合要求，请明确指出具体原因并给出修改建议。输出格式如下：
审查结果：通过 | 不通过
不通过原因：XXX
修改建议：XXX



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
            "max_tokens": 2000,
            "stream": True  # 启用流式输出
        }
        
        try:
            response = requests.post(
                self.api_url,
                headers=self.headers,
                json=payload,
                timeout=30,
                stream=True
            )
            
            if response.status_code == 200:
                for line in response.iter_lines():
                    if line:
                        line = line.decode('utf-8')
                        if line.startswith('data: '):
                            data = line[6:]  # 移除 'data: ' 前缀
                            if data == '[DONE]':
                                break
                            try:
                                json_data = json.loads(data)
                                if 'choices' in json_data and len(json_data['choices']) > 0:
                                    delta = json_data['choices'][0].get('delta', {})
                                    if 'content' in delta:
                                        content = delta['content']
                                        yield content
                                        if ENABLE_STREAMING:
                                            print(content, end='', flush=True)
                                            time.sleep(STREAMING_DELAY)
                            except json.JSONDecodeError:
                                continue
            else:
                error_msg = f"API调用失败: {response.status_code} - {response.text}"
                yield error_msg
                if ENABLE_STREAMING:
                    print(error_msg)
                
        except Exception as e:
            error_msg = f"分析过程中出现错误: {str(e)}"
            yield error_msg
            if ENABLE_STREAMING:
                print(error_msg)
    
    def analyze_node(self, node_info: str, context: str) -> str:
        """使用DeepSeek R1分析节点格式（非流式）"""
        # 在函数内部获取格式要求，不再作为参数传入
        format_requirements = PROMPTS
        
        prompt = f"""你是一个专业的文档格式检查专家。请分析以下文档节点的格式是否符合要求。

文档格式要求：
{format_requirements}

当前节点信息：
{node_info}

上下文信息：
{context}

请从以下几个方面进行分析：
0. 节点类型识别（是否为拟稿部门、日期等）
1. 格式合规性检查（字体、字号、对齐、缩进等）
2. 位置关系检查（是否符合文档结构要求）
3. 对以下内容需要忽略分析：
    -分割线
    -页眉页脚
    -加载两分割线之间的，且可能为落款的内容部分


输出格式：
    审查结果：通过|不通过
    不通过原因：XXX
    修改建议：XXX


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
            "max_tokens": 2000,
            "stream": False  # 非流式输出
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
    
    def analyze_node_with_streaming_control(self, node_info: str, context: str) -> str:
        """根据全局变量控制是否使用流式输出"""
        if ENABLE_STREAMING:
            print("\n=== 开始流式分析 ===")
            # 打印节点内容（最长30个字符），蓝色
            content_to_print = node_info
            if isinstance(node_info, dict) and 'content' in node_info:
                content_to_print = node_info['content']
            elif hasattr(node_info, 'content'):
                content_to_print = node_info.content
            # 截取前30个字符
            short_content = str(content_to_print)[:30]
            print(f"\033[34m节点内容: {short_content}\033[0m")
            full_response = ""
            for chunk in self.analyze_node_streaming(node_info, context):
                full_response += chunk
            print("\n=== 流式分析完成 ===")
            return full_response
        else:
            return self.analyze_node(node_info, context)
    
    def analyze_batch(self, nodes_data: list, format_requirements: str, delay: float = 1.0) -> list:
        """批量分析多个节点"""
        results = []
        
        for i, node_data in enumerate(nodes_data):
            print(f"正在分析节点 {i+1}/{len(nodes_data)}...")
            
            analysis = self.analyze_node_with_streaming_control(
                node_data['node_info'],
                node_data['context']
            )
            
            results.append({
                'node_index': node_data['node_index'],
                'analysis': analysis
            })
            
            # 添加延迟避免API限制
            if i < len(nodes_data) - 1:
                time.sleep(delay)
        
        return results
    
    def analyze_batch_nodes(self, nodes: list, context: str) -> str:
        """分析一批相邻节点（作为一个整体）"""
        # 在函数内部获取格式要求，不再作为参数传入
        format_requirements = PROMPTS
        
        # 构建批量节点信息
        batch_info = "=== 批量节点分析 ===\n"
        batch_info += f"节点数量: {len(nodes)}\n\n"
        
        for i, node in enumerate(nodes):
            batch_info += f"批量节点{i+1}:\n"
            batch_info += f"  内容: {node.content}\n"
            batch_info += f"  类型: {node.type}\n"
            batch_info += f"  字体: {node.font}\n"
            batch_info += f"  字号: {node.font_size_name} ({node.size}pt)\n"
            batch_info += f"  加粗: {node.bold}\n"
            if node.alignment:
                batch_info += f"  对齐: {node.alignment.get('value')}\n"
            if node.paragraph_format:
                pf = node.paragraph_format
                batch_info += f"  首行缩进: {pf.get('first_line_indent', {}).get('value', 0)}\n"
                batch_info += f"  左缩进: {pf.get('left_indent', {}).get('value', 0)}\n"
            batch_info += "\n"
        
        prompt = f"""你是一个专业的文档格式检查专家。请分析以下批量节点的格式是否符合要求。

文档格式要求：
{format_requirements}

{context}

批量节点信息：
{batch_info}

请逐个分析每个批量节点的格式问题，并提供改进建议。在分析结果中请明确标注每个节点的分析内容（如：批量节点1: xxx, 批量节点2: xxx）。
"""

        payload = {
            "model": "deepseek-chat",
            "messages": [
                {
                    "role": "user", 
                    "content": prompt
                }
            ],
            "temperature": 0.3,
            "max_tokens": 4000,
            "stream": False
        }
        
        try:
            response = requests.post(
                self.api_url,
                headers=self.headers,
                json=payload,
                timeout=60
            )
            
            if response.status_code == 200:
                result = response.json()
                return result['choices'][0]['message']['content']
            else:
                return f"批量分析失败: {response.status_code} - {response.text}"
                
        except Exception as e:
            return f"批量分析出错: {str(e)}"

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
    
    def analyze_single_node(self, node: NodeInfo, context: str) -> str:
        """分析单个节点"""
        node_info = self.generate_node_info(node)
        return self.analyzer.analyze_node_with_streaming_control(node_info, context)
    
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
    
    def analyze_batch_nodes(self, nodes: list, context: str) -> str:
        """批量分析相邻节点"""
        return self.analyzer.analyze_batch_nodes(nodes, context)

# 配置常量
DEEPSEEK_API_KEY = "sk-908c59db6a72469d9dcbd3607a9e2338"  # 请替换为您的API密钥
DEEPSEEK_API_URL = "https://api.deepseek.com/v1/chat/completions"

def create_analyzer(api_key: str = None) -> Optional[DocumentAnalyzer]:
    """创建文档分析器实例"""
    if api_key is None:
        api_key = DEEPSEEK_API_KEY
    
    if api_key == "your_api_key_here":
        print("警告: 请设置正确的DeepSeek API密钥")
        return None
    
    return DocumentAnalyzer(api_key)

def set_streaming_mode(enabled: bool, delay: float = 0.05):
    """设置流式输出模式"""
    global ENABLE_STREAMING, STREAMING_DELAY
    ENABLE_STREAMING = enabled
    STREAMING_DELAY = delay
    print(f"流式输出模式: {'启用' if enabled else '禁用'}")
    if enabled:
        print(f"流式输出延迟: {delay}秒")

# 示例使用
if __name__ == "__main__":
    # 设置流式输出模式
    set_streaming_mode(True, 0.03)  # 启用流式输出，延迟0.03秒
    
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
        result = analyzer.analyze_single_node(test_node, test_context)
        print("\n完整分析结果:")
        print(result)
    else:
        print("无法创建分析器，请检查API密钥配置")
