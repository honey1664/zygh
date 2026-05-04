"""
学生画像模块 - 学生信息录入、简历解析、能力评估
"""

import json
import re
from typing import Dict, List, Optional
from dataclasses import dataclass, asdict


@dataclass
class StudentProfile:
    """学生画像数据结构"""
    # 基本信息
    name: str = ""
    student_id: str = ""
    major: str = ""
    school: str = ""
    education: str = "本科"
    grade: str = ""
    
    # 能力信息
    skills: List[str] = None
    certificates: List[str] = None
    projects: List[Dict] = None
    internship: str = ""
    
    # 求职意向
    target_cities: List[str] = None
    target_salary: str = ""
    target_jobs: List[str] = None
    
    # 软技能
    self_evaluation: str = ""
    communication: str = "中"  # 强/中/弱
    learning: str = "强"  # 强/中/弱
    pressure: str = "中"  # 强/中/弱
    innovation: str = "中"  # 强/中/弱
    
    def __post_init__(self):
        if self.skills is None:
            self.skills = []
        if self.certificates is None:
            self.certificates = []
        if self.projects is None:
            self.projects = []
        if self.target_cities is None:
            self.target_cities = []
        if self.target_jobs is None:
            self.target_jobs = []
    
    def to_dict(self) -> Dict:
        """转换为字典"""
        return asdict(self)
    
    @classmethod
    def from_dict(cls, data: Dict) -> 'StudentProfile':
        """从字典创建"""
        return cls(**{k: v for k, v in data.items() if k in cls.__annotations__})


class StudentProfileManager:
    """学生画像管理器"""
    
    # 技能分类
    SKILL_CATEGORIES = {
        "编程语言": ["Python", "Java", "JavaScript", "C++", "C#", "Go", "Rust", "PHP", "Ruby", "TypeScript"],
        "前端技术": ["HTML", "CSS", "Vue", "React", "Angular", "jQuery", "Bootstrap", "Tailwind"],
        "后端框架": ["SpringBoot", "Django", "Flask", "Express", "NestJS", "Laravel", "Ruby on Rails"],
        "数据库": ["MySQL", "PostgreSQL", "MongoDB", "Redis", "Elasticsearch", "Oracle", "SQLite"],
        "大数据": ["Hadoop", "Spark", "Hive", "Kafka", "Flink", "HBase"],
        "云技术": ["AWS", "阿里云", "腾讯云", "Docker", "Kubernetes", "Kubernetes", "Terraform"],
        "机器学习": ["TensorFlow", "PyTorch", "Scikit-learn", "Keras", "XGBoost", "LightGBM"],
        "数据分析": ["Pandas", "NumPy", "Excel", "Tableau", "PowerBI", "R", "SPSS"],
        "开发工具": ["Git", "SVN", "VS Code", "IntelliJ IDEA", "Eclipse", "Maven", "npm"],
        "操作系统": ["Linux", "Windows", "macOS", "Shell", "Bash"]
    }
    
    # 证书分类
    CERTIFICATE_CATEGORIES = {
        "技术认证": ["软考初级", "软考中级", "软考高级", "AWS认证", "阿里云认证", "华为认证", "Azure认证"],
        "专业资格": ["PMP", "NPDP", "ACP", "CISSP", "CISA"],
        "语言证书": ["英语四级", "英语六级", "雅思", "托福", "日语N1", "日语N2"],
        "竞赛获奖": ["ACM", "蓝帽杯", "计算机设计大赛", "数学建模", "挑战杯"]
    }
    
    # 常见项目关键词
    PROJECT_KEYWORDS = {
        "电商": ["电商", "商城", "购物", "订单", "支付"],
        "博客": ["博客", "文章", "CMS", "内容管理"],
        "管理系统": ["管理", "OA", "ERP", "CRM", "HR"],
        "社交": ["社交", "论坛", "聊天", "即时通讯"],
        "数据": ["数据", "可视化", "大屏", "BI", "分析"],
        "AI": ["AI", "人工智能", "机器学习", "深度学习", "NLP", "CV"]
    }
    
    def __init__(self):
        self.current_profile = None
    
    def create_profile(self, form_data: Dict) -> StudentProfile:
        """从表单数据创建画像"""
        profile = StudentProfile(
            name=form_data.get("name", ""),
            student_id=form_data.get("student_id", ""),
            major=form_data.get("major", ""),
            school=form_data.get("school", ""),
            education=form_data.get("education", "本科"),
            grade=form_data.get("grade", ""),
            skills=form_data.get("skills", []),
            certificates=form_data.get("certificates", []),
            projects=form_data.get("projects", []),
            internship=form_data.get("internship", ""),
            target_cities=form_data.get("target_cities", []),
            target_salary=form_data.get("target_salary", ""),
            target_jobs=form_data.get("target_jobs", []),
            self_evaluation=form_data.get("self_evaluation", ""),
            communication=form_data.get("communication", "中"),
            learning=form_data.get("learning", "强"),
            pressure=form_data.get("pressure", "中"),
            innovation=form_data.get("innovation", "中")
        )
        
        self.current_profile = profile
        return profile
    
    def calculate_completeness(self, profile: StudentProfile) -> float:
        """
        计算画像完整度评分 (0-100)
        
        评分标准：
        - 基本信息：30分
        - 专业技能：25分
        - 证书：15分
        - 项目经验：20分
        - 实习经历：10分
        """
        score = 0
        
        # 基本信息 (30分)
        if profile.name:
            score += 5
        if profile.major:
            score += 8
        if profile.school:
            score += 7
        if profile.education:
            score += 5
        if profile.grade:
            score += 5
        
        # 专业技能 (25分)
        skill_count = len(profile.skills)
        score += min(25, skill_count * 5)
        
        # 证书 (15分)
        cert_count = len(profile.certificates)
        score += min(15, cert_count * 5)
        
        # 项目经验 (20分)
        project_count = len(profile.projects)
        score += min(20, project_count * 10)
        
        # 实习经历 (10分)
        if profile.internship:
            score += 10
        
        return min(100.0, score)
    
    def calculate_competitiveness(self, profile: StudentProfile) -> float:
        """
        计算竞争力评分 (0-100)
        
        评分标准：
        - 技能含金量：40分
        - 证书含金量：30分
        - 项目质量：20分
        - 实习质量：10分
        """
        score = 0
        
        # 技能含金量 (40分)
        high_value_skills = [
            "机器学习", "深度学习", "AI", "NLP", "计算机视觉",
            "云原生", "Kubernetes", "大数据", "Hadoop", "Spark",
            "架构设计", "高并发", "分布式"
        ]
        
        medium_value_skills = [
            "Python", "Java", "Go", "Rust",
            "SpringBoot", "Django", "Flask",
            "MySQL", "Redis", "MongoDB",
            "Docker", "Linux", "Git"
        ]
        
        for skill in profile.skills:
            if skill in high_value_skills:
                score += 5
            elif skill in medium_value_skills:
                score += 3
            else:
                score += 1
        
        score = min(40, score)
        
        # 证书含金量 (30分)
        high_value_certs = ["软考高级", "PMP", "CISSP", "AWS专家级", "ACM"]
        medium_value_certs = ["软考中级", "NPDP", "CISSP", "AWS认证"]
        
        for cert in profile.certificates:
            if cert in high_value_certs:
                score += 6
            elif cert in medium_value_certs:
                score += 4
            else:
                score += 2
        
        score = min(30, score)
        
        # 项目质量 (20分)
        project_count = len(profile.projects)
        score += min(20, project_count * 7)
        
        # 实习质量 (10分)
        if profile.internship:
            # 检查实习公司知名度
            famous_companies = ["BAT", "字节", "阿里", "腾讯", "百度", "京东", "华为", "微软", "Google", "Amazon"]
            if any(company in profile.internship for company in famous_companies):
                score += 10
            else:
                score += 5
        
        return min(100.0, score)
    
    def parse_resume(self, text: str) -> Dict:
        """
        从简历文本中提取信息
        
        使用规则+正则表达式提取
        """
        result = {
            "skills": [],
            "certificates": [],
            "projects": [],
            "education": "本科"
        }
        
        # 提取技能
        for category, skills in self.SKILL_CATEGORIES.items():
            for skill in skills:
                if skill.lower() in text.lower():
                    if skill not in result["skills"]:
                        result["skills"].append(skill)
        
        # 提取证书
        for category, certs in self.CERTIFICATE_CATEGORIES.items():
            for cert in certs:
                if cert in text:
                    if cert not in result["certificates"]:
                        result["certificates"].append(cert)
        
        # 提取学历
        if "博士" in text:
            result["education"] = "博士"
        elif "硕士" in text:
            result["education"] = "硕士"
        
        return result
    
    def compare_with_job(self, profile: StudentProfile, job_requirements: Dict) -> Dict:
        """
        与岗位要求对比分析
        
        Returns:
            匹配度详情
        """
        # 提取岗位技能要求
        required_skills = set(job_requirements.get("skills", []))
        student_skills = set(profile.skills)
        
        # 匹配技能
        matched = required_skills.intersection(student_skills)
        missing = required_skills - student_skills
        
        # 计算匹配率
        if required_skills:
            match_rate = len(matched) / len(required_skills) * 100
        else:
            match_rate = 0
        
        return {
            "required_skills": list(required_skills),
            "matched_skills": list(matched),
            "missing_skills": list(missing),
            "match_rate": round(match_rate, 1)
        }
    
    def generate_suggestions(self, profile: StudentProfile) -> List[str]:
        """生成提升建议"""
        suggestions = []
        
        # 技能建议
        if len(profile.skills) < 5:
            suggestions.append("建议增加更多专业技能，特别是主流技术栈")
        
        # 证书建议
        if len(profile.certificates) < 2:
            suggestions.append("建议考取相关证书，如软考、PMP等")
        
        # 项目建议
        if len(profile.projects) < 2:
            suggestions.append("建议增加项目经验，特别是实战项目")
        
        # 实习建议
        if not profile.internship:
            suggestions.append("建议尽早寻找实习机会，积累工作经验")
        
        return suggestions
    
    def export_profile(self, profile: StudentProfile, format: str = 'json') -> str:
        """导出画像"""
        if format == 'json':
            return json.dumps(profile.to_dict(), ensure_ascii=False, indent=2)
        return str(profile.to_dict())


class ResumeParser:
    """简历解析器 - 支持PDF和图片"""
    
    @staticmethod
    def extract_text_from_pdf(pdf_file) -> str:
        """从PDF提取文本"""
        # 实际项目中需要使用 pdfplumber 或 PyPDF2
        # 这里返回模拟数据
        return """
姓名：张三
学校：某某大学
专业：计算机科学与技术
学历：本科

技能：Python, Java, MySQL, Docker, Linux, Git, HTML, CSS, JavaScript, Vue.js

项目经验：
1. 电商网站开发 - 使用Vue.js和SpringBoot开发了一个完整的电商网站
2. 数据分析平台 - 使用Python和Echarts开发了数据可视化平台

证书：英语四级、软考初级

实习经历：在某某互联网公司实习3个月，主要负责后端开发
"""
    
    @staticmethod
    def extract_text_from_image(image_file) -> str:
        """从图片提取文本 - 需要OCR"""
        # 实际项目中需要使用 pytesseract 或 百度OCR API
        return ResumeParser.extract_text_from_pdf(None)
    
    @classmethod
    def parse(cls, file) -> Dict:
        """解析简历文件"""
        if file.name.endswith('.pdf'):
            text = cls.extract_text_from_pdf(file)
        elif file.name.lower().endswith(('.jpg', '.jpeg', '.png')):
            text = cls.extract_text_from_image(file)
        else:
            raise ValueError("不支持的文件格式")
        
        # 解析文本
        manager = StudentProfileManager()
        return manager.parse_resume(text)