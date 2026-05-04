"""
岗位画像模块 - 岗位数据管理、职业发展图谱、薪资预测
"""

import pandas as pd
import json
import re
from typing import Dict, List, Optional, Tuple
from collections import defaultdict


class JobProfile:
    """岗位画像数据结构"""
    
    def __init__(self):
        self.job_id: str = ""
        self.job_name: str = ""
        self.industry: str = ""
        self.company: str = ""
        self.city: str = ""
        self.salary: str = ""
        self.experience: str = ""
        self.education: str = ""
        self.skills: List[str] = []
        self.requirements: str = ""
        self.description: str = ""
        self.tags: List[str] = []
    
    def to_dict(self) -> Dict:
        return {
            "job_id": self.job_id,
            "job_name": self.job_name,
            "industry": self.industry,
            "company": self.company,
            "city": self.city,
            "salary": self.salary,
            "experience": self.experience,
            "education": self.education,
            "skills": self.skills,
            "requirements": self.requirements,
            "description": self.description,
            "tags": self.tags
        }


class JobProfileManager:
    """岗位画像管理器"""
    
    def __init__(self):
        self.jobs: List[JobProfile] = []
        self.job_categories = {
            "技术类": ["开发", "工程师", "程序员", "架构师", "测试", "运维", "算法", "数据", "前端", "后端"],
            "产品类": ["产品经理", "产品运营", "产品设计"],
            "设计类": ["UI", "UX", "设计师", "交互", "视觉"],
            "运营类": ["运营", "市场", "销售", "商务", "客服"],
            "职能类": ["人力", "财务", "行政", "法务"]
        }
    
    def load_jobs_from_csv(self, file_path: str) -> bool:
        """从CSV加载岗位数据"""
        try:
            df = pd.read_csv(file_path)
            self.jobs = []
            
            for _, row in df.iterrows():
                job = JobProfile()
                job.job_id = str(row.get("岗位ID", row.get("ID", "")))
                job.job_name = str(row.get("职位名称", row.get("岗位名称", "")))
                job.industry = str(row.get("行业", ""))
                job.company = str(row.get("公司", ""))
                job.city = str(row.get("城市", ""))
                job.salary = str(row.get("薪资", ""))
                job.experience = str(row.get("经验", row.get("工作经验", "")))
                job.education = str(row.get("学历", ""))
                job.skills = self._parse_skills(str(row.get("技能", row.get("技能要求", ""))))
                job.requirements = str(row.get("要求", row.get("岗位职责", "")))
                job.description = str(row.get("描述", row.get("职位描述", "")))
                job.tags = self._parse_tags(job.job_name)
                
                if job.job_name:  # 确保有岗位名称
                    self.jobs.append(job)
            
            print(f"成功加载 {len(self.jobs)} 个岗位")
            return True
        except Exception as e:
            print(f"加载岗位数据失败: {e}")
            return False
    
    def load_jobs_from_json(self, file_path: str) -> bool:
        """从JSON加载岗位数据"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            self.jobs = []
            for item in data:
                job = JobProfile()
                job.job_id = str(item.get("job_id", ""))
                job.job_name = str(item.get("job_name", item.get("职位名称", "")))
                job.industry = str(item.get("industry", ""))
                job.company = str(item.get("company", ""))
                job.city = str(item.get("city", ""))
                job.salary = str(item.get("salary", ""))
                job.experience = str(item.get("experience", ""))
                job.education = str(item.get("education", ""))
                job.skills = item.get("skills", [])
                job.requirements = str(item.get("requirements", ""))
                job.description = str(item.get("description", ""))
                job.tags = item.get("tags", [])
                
                if job.job_name:
                    self.jobs.append(job)
            
            print(f"成功加载 {len(self.jobs)} 个岗位")
            return True
        except Exception as e:
            print(f"加载岗位数据失败: {e}")
            return False
    
    def _parse_skills(self, skills_str: str) -> List[str]:
        """解析技能字符串"""
        if not skills_str:
            return []
        
        # 尝试多种分隔符
        separators = ["，", ",", "、", " ", ";", "；"]
        for sep in separators:
            if sep in skills_str:
                return [s.strip() for s in skills_str.split(sep) if s.strip()]
        
        # 如果没有分隔符，尝试按关键词识别
        common_skills = ["Python", "Java", "JavaScript", "MySQL", "Redis", "Linux", "Git", 
                        "Vue", "React", "Spring", "Django", "Flask", "Docker", "Kubernetes"]
        found = []
        for skill in common_skills:
            if skill.lower() in skills_str.lower():
                found.append(skill)
        
        return found if found else [skills_str.strip()]
    
    def _parse_tags(self, job_name: str) -> List[str]:
        """从岗位名称提取标签"""
        tags = []
        
        # 技术栈标签
        tech_patterns = {
            "Python": ["python", "爬虫", "django", "flask"],
            "Java": ["java", "spring", "后端"],
            "前端": ["前端", "vue", "react", "html", "css", "javascript"],
            "测试": ["测试", "qa", "自动化"],
            "算法": ["算法", "ai", "机器学习", "深度学习"],
            "数据": ["数据", "分析", "bi", "大数据"],
            "运维": ["运维", "devops", "docker", "kubernetes"],
            "产品": ["产品", "pm"],
            "设计": ["设计", "ui", "ux"]
        }
        
        job_lower = job_name.lower()
        for tag, patterns in tech_patterns.items():
            if any(pattern in job_lower for pattern in patterns):
                tags.append(tag)
        
        # 经验标签
        if "初级" in job_name or "助理" in job_name:
            tags.append("初级")
        elif "中级" in job_name or "工程师" in job_name:
            tags.append("中级")
        elif "高级" in job_name or "资深" in job_name:
            tags.append("高级")
        elif "架构" in job_name:
            tags.append("架构师")
        
        return tags
    
    def search_jobs(self, keywords: str = "", city: str = "", 
                   education: str = "", category: str = "") -> List[JobProfile]:
        """搜索岗位"""
        results = self.jobs
        
        if keywords:
            kw_lower = keywords.lower()
            results = [job for job in results 
                      if kw_lower in job.job_name.lower() 
                      or kw_lower in job.description.lower()
                      or any(kw_lower in skill.lower() for skill in job.skills)]
        
        if city:
            results = [job for job in results if city in job.city]
        
        if education:
            results = [job for job in results if education in job.education]
        
        if category:
            category_keywords = self.job_categories.get(category, [])
            results = [job for job in results 
                      if any(keyword in job.job_name for keyword in category_keywords)]
        
        return results
    
    def get_job_by_id(self, job_id: str) -> Optional[JobProfile]:
        """根据ID获取岗位"""
        for job in self.jobs:
            if job.job_id == job_id:
                return job
        return None
    
    def get_job_categories(self) -> List[str]:
        """获取岗位分类列表"""
        return list(self.job_categories.keys())
    
    def get_cities(self) -> List[str]:
        """获取所有城市"""
        cities = set()
        for job in self.jobs:
            if job.city:
                cities.add(job.city)
        return sorted(list(cities))
    
    def get_salary_statistics(self, city: str = "") -> Dict:
        """获取薪资统计"""
        jobs = self.jobs
        if city:
            jobs = [job for job in jobs if city in job.city]
        
        salaries = []
        for job in jobs:
            salary = self._parse_salary(job.salary)
            if salary:
                salaries.extend(salary)
        
        if not salaries:
            return {"avg": 0, "min": 0, "max": 0, "count": 0}
        
        return {
            "avg": round(sum(salaries) / len(salaries), 1),
            "min": min(salaries),
            "max": max(salaries),
            "count": len(salaries)
        }
    
    def _parse_salary(self, salary_str: str) -> Optional[List[float]]:
        """解析薪资字符串为数值"""
        if not salary_str or "面议" in salary_str:
            return None
        
        # 提取数字
        numbers = re.findall(r'\d+\.?\d*', salary_str)
        if len(numbers) >= 2:
            return [float(numbers[0]), float(numbers[1])]
        elif len(numbers) == 1:
            return [float(numbers[0])]
        
        return None
    
    def get_top_skills(self, limit: int = 10) -> List[Tuple[str, int]]:
        """获取热门技能排行"""
        skill_counts = defaultdict(int)
        for job in self.jobs:
            for skill in job.skills:
                skill_counts[skill] += 1
        
        return sorted(skill_counts.items(), key=lambda x: x[1], reverse=True)[:limit]
    
    def get_job_recommendations(self, skills: List[str]) -> List[JobProfile]:
        """根据技能推荐岗位"""
        skill_set = set(skill.lower() for skill in skills)
        
        # 计算匹配度
        job_matches = []
        for job in self.jobs:
            job_skills = set(skill.lower() for skill in job.skills)
            matched = len(skill_set & job_skills)
            total = len(job_skills)
            
            if total > 0:
                match_rate = matched / total
                job_matches.append((job, match_rate))
        
        # 按匹配度排序
        job_matches.sort(key=lambda x: x[1], reverse=True)
        
        return [job for job, rate in job_matches[:10]]
    
    def predict_salary(self, skills: List[str], experience: str = "") -> float:
        """预测薪资"""
        # 基于技能和经验预测薪资
        skill_salary_map = {
            "Python": 8,
            "Java": 9,
            "JavaScript": 7,
            "前端": 7,
            "后端": 8,
            "算法": 15,
            "机器学习": 18,
            "深度学习": 20,
            "数据": 10,
            "测试": 6,
            "运维": 7,
            "Docker": 5,
            "Kubernetes": 8,
            "MySQL": 4,
            "Redis": 3
        }
        
        base_salary = 6  # 基础薪资
        
        # 根据技能累加
        for skill in skills:
            base_salary += skill_salary_map.get(skill, 0) * 0.5
        
        # 根据经验调整
        if "3-5年" in experience or "三年" in experience:
            base_salary *= 1.5
        elif "5年以上" in experience or "五年" in experience:
            base_salary *= 2.0
        
        return round(base_salary, 1)
    
    def generate_career_path(self, job_name: str) -> Dict:
        """生成职业发展路径"""
        career_paths = {
            "初级开发工程师": {
                "current": "初级开发工程师",
                "next": "中级开发工程师",
                "skills": ["扎实的编程基础", "熟悉主流框架", "代码规范"],
                "timeline": "1-2年",
                "next_skills": ["架构设计", "性能优化", "团队协作"]
            },
            "中级开发工程师": {
                "current": "中级开发工程师",
                "next": "高级开发工程师/技术主管",
                "skills": ["独立完成模块开发", "性能优化", "Code Review"],
                "timeline": "2-3年",
                "next_skills": ["系统架构", "技术选型", "团队管理"]
            },
            "高级开发工程师": {
                "current": "高级开发工程师",
                "next": "技术总监/架构师",
                "skills": ["系统架构设计", "技术攻关", "技术规划"],
                "timeline": "3-5年",
                "next_skills": ["战略规划", "业务理解", "跨团队协调"]
            },
            "产品经理": {
                "current": "产品经理",
                "next": "高级产品经理/产品总监",
                "skills": ["需求分析", "产品设计", "项目管理"],
                "timeline": "2-3年",
                "next_skills": ["商业思维", "数据分析", "团队管理"]
            },
            "算法工程师": {
                "current": "算法工程师",
                "next": "高级算法工程师/算法负责人",
                "skills": ["算法设计", "模型优化", "论文阅读"],
                "timeline": "2-3年",
                "next_skills": ["系统架构", "业务理解", "团队管理"]
            }
        }
        
        # 模糊匹配
        for key in career_paths:
            if key in job_name or job_name in key:
                return career_paths[key]
        
        # 默认路径
        return {
            "current": job_name,
            "next": f"资深{job_name}",
            "skills": ["专业技能提升", "项目经验积累"],
            "timeline": "2-3年",
            "next_skills": ["团队管理", "技术深度"]
        }


# 全局实例
job_manager = JobProfileManager()