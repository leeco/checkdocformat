import requests
import json
from typing import Dict, Any, Optional
from dataclasses import dataclass

@dataclass
class NodeClassificationInfo:
    """节点分类信息"""
    content: str
    font: str
    size: float
    bold: bool
    outline_level: Optional[Dict]
    alignment: Optional[Dict]
    paragraph_format: Optional[Dict]

class AINodeClassifier:
    """AI节点分类器"""
    
    def __init__(self, api_key: str, api_url: str = "https://api.deepseek.com/v1/chat/completions"):
        self.api_key = api_key
        self.api_url = api_url
        self.headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        }
    
    def classify_node(self, node_info: NodeClassificationInfo, context_nodes: list = None) -> str:
        """使用AI对节点进行分类"""
        
        # 构建上下文信息
        context_str = ""
        if context_nodes:
            context_str = "\n上下文节点:\n"
            for i, node in enumerate(context_nodes[-3:]):  # 只取前3个节点作为上下文
                context_str += f"  节点{i+1}: {node.get('content', '')[:50]}...\n"
        
        prompt = f"""你是一个专业的公文文档结构分析专家。请分析以下文档节点的类型。

节点信息：
- 内容: {node_info.content}
- 字体: {node_info.font}
- 字号: {node_info.size}pt
- 加粗: {node_info.bold}
- 对齐: {node_info.alignment.get('value') if node_info.alignment else 'None'}
- 大纲级别: {node_info.outline_level.get('value') if node_info.outline_level else 'None'}

{context_str}

请从以下类型中选择最合适的一个：
1. 发文标题 - 由单位名称、事由和文种组成，通常居中，字体较大
2. 主送机关 - 如："XX市人民政府："
3. 一级标题 - 如：一、二、三、等
4. 二级标题 - 如：（一）（二）等  
5. 三级标题 - 如：1. 2. 3. 等
6. 四级标题 - 如：（1）（2）等
7. 普通段落 - 正文内容
8. 列表项 - 注意：不要将列表项误判为标题
9. 结尾 - 如："特此报告"、"特此请示"、"特此申请"等
10. 落款 - 发文单位名称和日期
11. 附件 - 附件说明，如："附件：1.XXXX"
12. 分隔符 - 如："———"、"＊＊＊"等
13. 空行 - 空行

判断标准：
- 发文标题：通常位于文档开头，居中对齐，包含事由和文种
- 主送机关：通常在标题下方，以机关名称开头，以冒号结尾
- 标题有明确的编号格式和层级关系
- 普通段落是正文内容，通常首行缩进2字
- 结尾：固定格式，如"特此报告"等
- 落款：包含发文单位和日期
- 附件：以"附件："开头
- 列表项：以项目符号或编号开头，但内容相对简短
- 分隔符：主要由符号组成的分割线

请只返回类型名称，不要包含其他内容。"""

        payload = {
            "model": "deepseek-chat",
            "messages": [
                {
                    "role": "user",
                    "content": prompt
                }
            ],
            "temperature": 0.1,  # 低温度确保一致性
            "max_tokens": 50
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
                classification = result['choices'][0]['message']['content'].strip().lower()
                
                # 验证返回的类型是否有效
                valid_types = ['发文标题', '主送机关', '一级标题', '二级标题', '三级标题', '四级标题', 
                              '普通段落', '列表项', '结尾', '落款', '附件', '分隔符', '空行']
                if classification in valid_types:
                    return classification
                else:
                    # 如果AI返回了无效类型，使用备用逻辑
                    return self._fallback_classification(node_info)
            else:
                print(f"AI分类失败: {response.status_code} - {response.text}")
                return self._fallback_classification(node_info)
                
        except Exception as e:
            print(f"AI分类出错: {str(e)}")
            return self._fallback_classification(node_info)
    
    def _fallback_classification(self, node_info: NodeClassificationInfo) -> str:
        """备用分类逻辑（当AI不可用时使用）"""
        content_stripped = node_info.content.strip()
        
        # 检查空行
        if not content_stripped:
            return '空行'
        
        # 检查分隔符（主要由符号组成）
        if self._is_separator(content_stripped):
            return '分隔符'
        
        # 检查附件
        if content_stripped.startswith('附件：') or content_stripped.startswith('附件:'):
            return '附件'
        
        # 检查结尾
        ending_patterns = ['特此报告', '特此请示', '特此申请', '特此函告', '特此通知', '特此通报']
        if any(pattern in content_stripped for pattern in ending_patterns):
            return '结尾'
        
        # 检查主送机关（以冒号结尾）
        if content_stripped.endswith('：') or content_stripped.endswith(':'):
            # 进一步判断是否包含机关名称
            addressee_keywords = ['政府', '委员会', '局', '厅', '部', '院', '处', '科', '司', '公司', '单位']
            if any(keyword in content_stripped for keyword in addressee_keywords):
                return '主送机关'
        
        # 检查落款（包含日期格式）
        import re
        date_pattern = r'\d{4}年\d{1,2}月\d{1,2}日'
        if re.search(date_pattern, content_stripped):
            return '落款'
        
        # 检查发文标题（通常居中，字号较大，在文档开头）
        alignment_value = node_info.alignment.get('value') if node_info.alignment else ''
        if ('居中' in alignment_value or 'CENTER' in str(alignment_value)) and node_info.size >= 16:
            # 检查是否包含文种关键词
            document_types = ['报告', '请示', '申请', '通知', '通报', '函', '意见', '决定', '通告', '公告', '令']
            if any(doc_type in content_stripped for doc_type in document_types):
                return '发文标题'
        
        # 检查是否为列表项
        if self._is_list_item(content_stripped):
            return '列表项'
        
        # 检查标题模式
        # Level 1: "一、", "二、", "三、" etc.
        if content_stripped and content_stripped[0] in '一二三四五六七八九十' and '、' in content_stripped[:3]:
            return '一级标题'
        
        # Level 2: "（一）", "（二）" etc.
        if content_stripped.startswith('（') and content_stripped[1] in '一二三四五六七八九十' and '）' in content_stripped:
            return '二级标题'
        
        # Level 3: "1.", "2." etc.
        if content_stripped and content_stripped[0].isdigit() and '.' in content_stripped[:3]:
            return '三级标题'
        
        # Level 4: "（1）", "（2）" etc.
        if content_stripped.startswith('（') and content_stripped[1].isdigit() and '）' in content_stripped:
            return '四级标题'
        
        # 检查Word大纲级别
        if node_info.outline_level and node_info.outline_level.get('value') != '正文文本':
            outline_value = node_info.outline_level.get('value', '')
            if outline_value == '标题1':
                return '一级标题'
            elif outline_value == '标题2':
                return '二级标题'
            elif outline_value == '标题3':
                return '三级标题'
            elif outline_value == '标题4':
                return '四级标题'
        
        # 检查字体特征
        if node_info.bold and node_info.size >= 16:
            return '一级标题'
        elif node_info.bold and node_info.size >= 14:
            return '二级标题'
        elif node_info.bold and node_info.size >= 12:
            return '三级标题'
        
        return '普通段落'
    
    def _is_list_item(self, content: str) -> bool:
        """判断是否为列表项"""
        content_stripped = content.strip()
        
        # 检查常见的列表项模式
        list_patterns = [
            # 项目符号
            lambda x: x.startswith('•'),
            lambda x: x.startswith('·'),
            lambda x: x.startswith('▪'),
            lambda x: x.startswith('▫'),
            lambda x: x.startswith('-'),
            lambda x: x.startswith('—'),
            # 数字列表（但内容较短，可能是列表项）
            lambda x: x[0].isdigit() and '.' in x[:3] and len(x) < 100,
            # 字母列表
            lambda x: x[0].isalpha() and '.' in x[:3] and len(x) < 100,
            # 中文数字列表（但内容较短）
            lambda x: x[0] in '一二三四五六七八九十' and '、' in x[:3] and len(x) < 100,
        ]
        
        for pattern in list_patterns:
            if pattern(content_stripped):
                return True
        
        return False
    
    def _is_separator(self, content: str) -> bool:
        """判断是否为分隔符"""
        content_stripped = content.strip()
        
        # 检查是否主要由分隔符字符组成
        separator_chars = set('—―-_*＊×※＝=')
        content_chars = set(content_stripped)
        
        # 如果内容中80%以上是分隔符字符，则认为是分隔符
        if len(content_chars) > 0:
            separator_ratio = len(content_chars & separator_chars) / len(content_chars)
            if separator_ratio >= 0.8 and len(content_stripped) >= 3:
                return True
        
        # 检查常见的分隔符模式
        separator_patterns = [
            lambda x: x.count('—') >= 3,
            lambda x: x.count('―') >= 3,
            lambda x: x.count('-') >= 5,
            lambda x: x.count('_') >= 5,
            lambda x: x.count('*') >= 3,
            lambda x: x.count('＊') >= 3,
            lambda x: x.count('=') >= 3,
            lambda x: x.count('＝') >= 3,
        ]
        
        for pattern in separator_patterns:
            if pattern(content_stripped):
                return True
        
        return False

class HybridNodeClassifier:
    """混合节点分类器 - 结合AI和规则"""
    
    def __init__(self, api_key: str = None):
        self.ai_classifier = AINodeClassifier(api_key) if api_key else None
        self.use_ai = api_key is not None
    
    def classify_node(self, content: str, font: str, size: float, bold: bool, 
                     outline_level: Optional[Dict] = None, alignment: Optional[Dict] = None, 
                     paragraph_format: Optional[Dict] = None, context_nodes: list = None) -> str:
        """分类节点类型"""
        
        node_info = NodeClassificationInfo(
            content=content,
            font=font,
            size=size,
            bold=bold,
            outline_level=outline_level,
            alignment=alignment,
            paragraph_format=paragraph_format
        )
        
        if self.use_ai and self.ai_classifier:
            try:
                return self.ai_classifier.classify_node(node_info, context_nodes)
            except Exception as e:
                print(f"AI分类失败，使用备用逻辑: {str(e)}")
                return self.ai_classifier._fallback_classification(node_info)
        else:
            # 直接使用备用逻辑
            return self.ai_classifier._fallback_classification(node_info) if self.ai_classifier else self._simple_classification(node_info)
    
    def _simple_classification(self, node_info: NodeClassificationInfo) -> str:
        """简单的分类逻辑（当没有AI时使用）"""
        return self.ai_classifier._fallback_classification(node_info) if self.ai_classifier else 'paragraph'

# 配置常量
DEEPSEEK_API_KEY = "sk-908c59db6a72469d9dcbd3607a9e2338"  # 请替换为您的API密钥

def create_classifier(api_key: str = None) -> HybridNodeClassifier:
    """创建节点分类器"""
    if api_key is None:
        api_key = DEEPSEEK_API_KEY
    return HybridNodeClassifier(api_key)

# 示例使用
if __name__ == "__main__":
    # 创建分类器
    classifier = create_classifier()
    
    # 测试用例
    test_cases = [
        {
            'content': '关于加强项目管理工作的报告',
            'font': '方正小标宋简体',
            'size': 22.0,
            'bold': False,
            'alignment': {'value': '居中'},
            'expected': '发文标题'
        },
        {
            'content': 'XX市人民政府：',
            'font': '仿宋_GB2312',
            'size': 14.0,
            'bold': False,
            'expected': '主送机关'
        },
        {
            'content': '一、项目概述',
            'font': '黑体',
            'size': 14.0,
            'bold': False,
            'expected': '一级标题'
        },
        {
            'content': '（一）项目背景',
            'font': '楷体',
            'size': 14.0,
            'bold': False,
            'expected': '二级标题'
        },
        {
            'content': '1. 技术方案',
            'font': '仿宋_GB2312',
            'size': 14.0,
            'bold': False,
            'expected': '三级标题'
        },
        {
            'content': '（1）系统架构',
            'font': '仿宋_GB2312',
            'size': 14.0,
            'bold': False,
            'expected': '四级标题'
        },
        {
            'content': '• 系统架构设计',
            'font': '仿宋_GB2312',
            'size': 14.0,
            'bold': False,
            'expected': '列表项'
        },
        {
            'content': '这是一个普通的段落内容，包含详细的描述信息。',
            'font': '仿宋_GB2312',
            'size': 14.0,
            'bold': False,
            'expected': '普通段落'
        },
        {
            'content': '特此报告',
            'font': '仿宋_GB2312',
            'size': 14.0,
            'bold': False,
            'expected': '结尾'
        },
        {
            'content': 'XX单位\n2024年1月15日',
            'font': '仿宋_GB2312',
            'size': 14.0,
            'bold': False,
            'expected': '落款'
        },
        {
            'content': '附件：1.项目实施方案',
            'font': '仿宋_GB2312',
            'size': 14.0,
            'bold': False,
            'expected': '附件'
        },
        {
            'content': '——————————————————',
            'font': '宋体',
            'size': 12.0,
            'bold': False,
            'expected': '分隔符'
        }
    ]
    
    print("=== 节点分类测试 ===")
    for i, test_case in enumerate(test_cases):
        result = classifier.classify_node(
            content=test_case['content'],
            font=test_case['font'],
            size=test_case['size'],
            bold=test_case['bold'],
            alignment=test_case.get('alignment')
        )
        
        status = "✓" if result == test_case['expected'] else "✗"
        print(f"{status} 测试{i+1}: {test_case['content'][:20]}... -> {result} (期望: {test_case['expected']})")
