# ai_chat_client.py
import os
import requests
import json
from dotenv import load_dotenv

# 加载环境变量
load_dotenv()


class AIChatClient:
    def __init__(self):
        # 加载 DeepSeek 配置
        self.api_key = os.getenv("DEEPSEEK_API_KEY") or os.getenv("AI_API_KEY")
        self.api_url = os.getenv("DEEPSEEK_API_URL", "https://api.deepseek.com/v1/chat/completions")

        # 备用：通义千问
        self.qwen_api_key = os.getenv("QWEN_API_KEY")
        self.qwen_base_url = os.getenv("QWEN_BASE_URL", "https://dashscope.aliyuncs.com/compatible-mode/v1")

        # 检查 API 配置
        if self.api_key:
            print("DeepSeek API 已配置")
            self.model_type = "deepseek"
        elif self.qwen_api_key:
            print("通义千问 API 已配置")
            self.model_type = "tongyi"
        else:
            print("警告: 未配置API密钥，将使用模拟模式")
            self.model_type = "mock"

    def build_prompt(self, user_question, student_profile=None):
        """构造个性化提示词"""
        system_prompt = """你是一位专业的AI职业规划顾问，专注为大学生提供职业规划、求职面试、技能提升的个性化建议。

你的特点：
1. 回答要专业、具体、可落地
2. 语气亲切易懂，像朋友一样交流
3. 结合学生的实际情况给出建议
4. 提供具体的行动步骤和学习资源
5. 回答要简洁有力，使用适当的emoji让回答更生动

请根据学生的问题，给出专业、有用的回答。"""

        # 拼接学生画像信息
        profile_info = ""
        if student_profile:
            basic = student_profile.get("基本信息", {})
            skills = student_profile.get("专业技能", [])
            certs = student_profile.get("证书", [])
            job_intent = student_profile.get("求职意向", {})

            profile_info = f"""
【学生画像信息】
- 姓名: {basic.get('姓名', '同学')}
- 专业: {basic.get('专业', '未提供')}
- 学历: {basic.get('学历', '未提供')}
- 学校: {basic.get('学校', '未提供')}
- 专业技能: {', '.join(skills) if skills else '未提供'}
- 证书: {', '.join(certs) if certs else '未提供'}
- 项目经验: {len(student_profile.get('项目经验', []))}个
- 目标岗位: {job_intent.get('期望岗位', '未提供')}
"""

        # 最终提示词
        full_prompt = f"{system_prompt}\n{profile_info}\n\n学生问题：{user_question}\n\n请给出专业、具体、可落地的回答："
        return full_prompt

    def call_deepseek(self, prompt):
        """调用 DeepSeek API"""
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }
        data = {
            "model": "deepseek-chat",
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.7,
            "max_tokens": 2000
        }

        try:
            print(f"正在调用 DeepSeek API...")
            response = requests.post(
                self.api_url,
                headers=headers,
                json=data,
                timeout=60
            )
            response.raise_for_status()
            result = response.json()
            answer = result["choices"][0]["message"]["content"]
            print(f"DeepSeek 响应成功，长度: {len(answer)}")
            return answer
        except requests.exceptions.Timeout:
            print("DeepSeek API 超时")
            return None
        except Exception as e:
            print(f"DeepSeek API 调用失败: {e}")
            return None

    def call_tongyi(self, prompt):
        """调用通义千问 API"""
        if not self.qwen_api_key:
            return None

        headers = {
            "Authorization": f"Bearer {self.qwen_api_key}",
            "Content-Type": "application/json"
        }
        data = {
            "model": "qwen-turbo",
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.7,
            "max_tokens": 2000
        }

        try:
            print(f"正在调用通义千问 API...")
            response = requests.post(
                f"{self.qwen_base_url}/chat/completions",
                headers=headers,
                json=data,
                timeout=60
            )
            response.raise_for_status()
            result = response.json()
            return result["choices"][0]["message"]["content"]
        except Exception as e:
            print(f"通义千问 API 调用失败: {e}")
            return None

    def get_fallback_answer(self, user_question):
        """兜底回答（当 API 调用失败时）"""
        return f"""抱歉，AI服务暂时无法响应您的问题「{user_question[:50]}」。

可能的原因：
1. API 服务暂时不可用
2. 网络连接问题

请稍后再试。如果问题持续存在，请联系管理员。"""

    def get_answer(self, user_question, student_profile=None):
        """对外暴露的统一接口"""
        prompt = self.build_prompt(user_question, student_profile)

        answer = None

        # 尝试调用 DeepSeek
        if self.model_type == "deepseek" and self.api_key:
            answer = self.call_deepseek(prompt)
            if answer:
                return answer

        # 尝试调用通义千问
        if self.qwen_api_key and answer is None:
            answer = self.call_tongyi(prompt)
            if answer:
                return answer

        # 使用兜底回答
        return self.get_fallback_answer(user_question)


# 初始化客户端
ai_chat_client = AIChatClient()