# -*- coding: utf-8 -*-
import os
import json
from google import genai
from google.genai import types
from dotenv import load_dotenv

# 加载环境变量（内含 GEMINI_API_KEY）
load_dotenv()

class LLMClient:
    def __init__(self):
        """
        初始化 Gemini 客户端
        """
        # 从环境变量获取 API Key
        api_key = os.getenv("GEMINI_API_KEY")
        if not api_key:
            raise ValueError("未在 .env 文件或环境变量中找到 GEMINI_API_KEY，请检查配置！")
            
        # 初始化谷歌官方 Gemini 客户端（使用指定的 api_key）
        self.client = genai.Client(api_key=api_key)
        self.model = "gemini-3.5-flash"

    def generate_experiment_solution(self, requirements: str, rag_context: str = "") -> dict:
        """
        根据实验要求和 RAG 上下文，调用 Gemini 生成 Python 代码和实验结果分析。
        返回一个包含 'code' 和 'text_analysis' 的字典。
        """
        # 核心系统提示词，在这里我们明确要求大模型模仿用户的旧报告风格
        system_prompt = (
            "你是一个 Python 实验报告辅助生成助手。\n"
            "请仔细分析实验要求，并结合参考背景（注意：参考背景里提供了用户以前的旧报告内容，"
            "它代表了用户的个人说话语气和写作风格。请在生成 'text_analysis' 时，"
            "尽量模仿参考资料中的文风、措辞习惯和逻辑叙述，确保风格的延续，像是一个大一学生自己写的）。\n\n"
            "请生成一个包含以下两个字段的 JSON 对象：\n"
            "1. 'code': 解决该实验要求的完整 Python 代码。必须是一键可运行、自包含的代码。\n"
            "   - 若涉及 matplotlib 绘图，必须在开头设置非交互式后端：\n"
            "     import matplotlib; matplotlib.use('Agg')\n"
            "   - 绘图时需设置中文字体以防乱码：\n"
            "     plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial']\n"
            "     plt.rcParams['axes.unicode_minus'] = False\n"
            "   - 所有 plt.show() 替换为 plt.savefig() 并保存图片。\n"
            "   - 所有生成的图片保存到相对路径 './data/output/' 目录下（如 './data/output/result.png'）。\n"
            "   - 确保代码中包含 `os.makedirs('./data/output', exist_ok=True)`。\n"
            "   - 若需要读取外部数据，请从相对路径 './data/raw/' 中读取。\n"
            "2. 'text_analysis': 针对实验结果撰写的实验结果描述与文字分析（Markdown 格式，200~300字左右），"
            "说明实验的完成情况与图表展示的规律。用词要亲切自然，不要有过于生硬的 AI 腔调。\n\n"
            "注意：请务必返回合法的 JSON 格式，其中键名为 'code' 和 'text_analysis'，例如：\n"
            "{\n"
            "  \"code\": \"# 你的代码\\n...\",\n"
            "  \"text_analysis\": \"你的实验分析\\n...\"\n"
            "}"
        )
        
        contents = []
        if rag_context:
            contents.append(f"【参考背景材料（用于风格学习与知识库检索）】:\n{rag_context}\n")
        contents.append(f"实验要求:\n{requirements}")
        full_input = "\n".join(contents)
        
        try:
            # 调用官方 genai 接口生成内容
            response = self.client.models.generate_content(
                model=self.model,
                contents=full_input,
                config=types.GenerateContentConfig(
                    system_instruction=system_prompt,
                    response_mime_type="application/json",  # 强制要求大模型返回 JSON
                ),
            )
            
            # 💡 防崩防呆处理：清除首尾空白，并清洗大模型可能带出来的 ```json 标签
            raw_text = response.text.strip()
            if raw_text.startswith("```json"):
                raw_text = raw_text[7:]
            elif raw_text.startswith("```"):
                raw_text = raw_text[3:]
            if raw_text.endswith("```"):
                raw_text = raw_text[:-3]
            raw_text = raw_text.strip()
            
            # 解析大模型返回的 JSON 数据
            result = json.loads(raw_text)
            return {
                "code": result.get("code", ""),
                "text_analysis": result.get("text_analysis", "")
            }
        except Exception as e:
            print(f"❌ 调用大模型生成方案失败: {e}")
            return {
                "code": f"# 自动生成失败，错误信息: {e}",
                "text_analysis": f"生成实验分析时发生异常：{e}"
            }

    def ask_with_context(self, question: str, context: str) -> str:
        """
        结合检索出的文档上下文，回答用户的问题（用于一般性问答）。
        """
        system_prompt = (
            "你是一个专业的文档分析助手。请根据下方提供的参考文档内容回答用户的问题。\n"
            "如果文档中没有相关信息，请坦白告知，不要胡乱编造。\n\n"
            f"【参考文档内容】:\n{context}"
        )
        
        try:
            response = self.client.models.generate_content(
                model=self.model,
                contents=question,
                config=types.GenerateContentConfig(
                    system_instruction=system_prompt,
                ),
            )
            return response.text
        except Exception as e:
            return f"调用 Gemini API 出错: {e}"
