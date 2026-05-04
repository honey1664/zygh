"""
职业测评模块 - 职业兴趣测评、能力评估、性格分析
"""

import json
import random
from typing import Dict, List, Tuple
from dataclasses import dataclass, asdict


@dataclass
class Question:
    """测评问题"""
    question_id: str
    question: str
    options: List[str]
    category: str  # 兴趣/能力/性格


@dataclass
class AssessmentResult:
    """测评结果"""
    # 基本信息
    user_id: str = ""
    timestamp: str = ""
    
    # 兴趣测评结果
    interest_scores: Dict[str, int] = None
    interest_type: str = ""
    interest_description: str = ""
    
    # 能力测评结果
    ability_scores: Dict[str, int] = None
    ability_ranking: List[str] = None
    
    # 性格测评结果
    personality_scores: Dict[str, int] = None
    personality_type: str = ""
    personality_description: str = ""
    
    # 综合建议
    suggested_jobs: List[str] = []
    suggested_skills: List[str] = []
    
    def __post_init__(self):
        if self.interest_scores is None:
            self.interest_scores = {}
        if self.ability_scores is None:
            self.ability_scores = {}
        if self.ability_ranking is None:
            self.ability_ranking = []
        if self.personality_scores is None:
            self.personality_scores = {}
        if self.suggested_jobs is None:
            self.suggested_jobs = []
        if self.suggested_skills is None:
            self.suggested_skills = []
    
    def to_dict(self) -> Dict:
        return asdict(self)


class CareerAssessment:
    """职业测评系统"""
    
    def __init__(self):
        # 兴趣测评题目
        self.interest_questions = [
            Question(
                "i1",
                "你更喜欢哪种活动？",
                ["编程写代码", "与人交流沟通", "设计创意方案", "数据分析"],
                "兴趣"
            ),
            Question(
                "i2",
                "你周末通常会做什么？",
                ["学习新技术", "参加社交活动", "画画或设计", "阅读书籍"],
                "兴趣"
            ),
            Question(
                "i3",
                "你解决问题的方式是？",
                ["逻辑分析", "听取他人意见", "创新思维", "数据分析"],
                "兴趣"
            ),
            Question(
                "i4",
                "你更喜欢什么样的工作环境？",
                ["安静独立工作", "团队协作", "自由创意", "结构化流程"],
                "兴趣"
            ),
            Question(
                "i5",
                "你对什么最感兴趣？",
                ["技术创新", "人际关系", "艺术设计", "商业数据"],
                "兴趣"
            ),
            Question(
                "i6",
                "你喜欢什么样的任务？",
                ["有挑战性的技术难题", "与人打交道的任务", "创意性的设计任务", "需要分析的任务"],
                "兴趣"
            ),
            Question(
                "i7",
                "你学习新东西的方式是？",
                ["通过实践", "与人讨论", "通过观察", "通过阅读"],
                "兴趣"
            ),
            Question(
                "i8",
                "你更看重什么？",
                ["技术深度", "人际关系", "美学感受", "数据驱动"],
                "兴趣"
            )
        ]
        
        # 能力测评题目
        self.ability_questions = [
            Question(
                "a1",
                "你解决复杂编程问题的能力如何？",
                ["很强", "较强", "一般", "较弱"],
                "能力"
            ),
            Question(
                "a2",
                "你学习新技术的速度如何？",
                ["很快", "较快", "一般", "较慢"],
                "能力"
            ),
            Question(
                "a3",
                "你的沟通表达能力如何？",
                ["很强", "较强", "一般", "较弱"],
                "能力"
            ),
            Question(
                "a4",
                "你的团队协作能力如何？",
                ["很强", "较强", "一般", "较弱"],
                "能力"
            ),
            Question(
                "a5",
                "你的问题分析能力如何？",
                ["很强", "较强", "一般", "较弱"],
                "能力"
            ),
            Question(
                "a6",
                "你的创新思维能力如何？",
                ["很强", "较强", "一般", "较弱"],
                "能力"
            ),
            Question(
                "a7",
                "你的时间管理能力如何？",
                ["很强", "较强", "一般", "较弱"],
                "能力"
            ),
            Question(
                "a8",
                "你的抗压能力如何？",
                ["很强", "较强", "一般", "较弱"],
                "能力"
            )
        ]
        
        # 性格测评题目
        self.personality_questions = [
            Question(
                "p1",
                "你更倾向于？",
                ["关注内心想法", "关注外部世界"],
                "性格"
            ),
            Question(
                "p2",
                "你获取信息的方式是？",
                ["通过直觉和灵感", "通过事实和经验"],
                "性格"
            ),
            Question(
                "p3",
                "你做决策的方式是？",
                ["基于逻辑分析", "基于情感和价值观"],
                "性格"
            ),
            Question(
                "p4",
                "你更喜欢的生活方式是？",
                ["有计划有条理", "灵活自由"],
                "性格"
            ),
            Question(
                "p5",
                "在团队中你通常是？",
                ["决策者", "协调者"],
                "性格"
            ),
            Question(
                "p6",
                "面对压力时你会？",
                ["冷静分析", "寻求支持"],
                "性格"
            )
        ]
        
        # 兴趣类型映射
        self.interest_types = {
            "技术型": {
                "description": "你对技术有浓厚兴趣，喜欢解决复杂问题，适合从事技术开发、算法研究等工作",
                "jobs": ["软件工程师", "算法工程师", "系统架构师", "数据工程师"]
            },
            "社交型": {
                "description": "你善于与人打交道，喜欢沟通交流，适合从事管理、市场、销售等工作",
                "jobs": ["产品经理", "项目经理", "市场专员", "人力资源"]
            },
            "艺术型": {
                "description": "你富有创造力，喜欢设计和美学，适合从事设计、创意等工作",
                "jobs": ["UI设计师", "UX设计师", "产品设计师", "视觉设计师"]
            },
            "分析型": {
                "description": "你擅长数据分析和逻辑推理，适合从事数据分析、商业分析等工作",
                "jobs": ["数据分析师", "商业分析师", "BI工程师", "量化研究员"]
            }
        }
        
        # 性格类型映射
        self.personality_types = {
            "INTJ": {
                "description": "战略型：独立、果断，善于分析和规划，适合领导和管理岗位",
                "traits": ["理性", "独立", "果断", "专注"]
            },
            "INFJ": {
                "description": "咨询师型：富有洞察力，善于理解他人，适合咨询和教育岗位",
                "traits": ["洞察", "关怀", "理想", "创意"]
            },
            "INT": {
                "description": "思考型：逻辑性强，善于分析，适合技术和分析岗位",
                "traits": ["逻辑", "分析", "理性", "专注"]
            },
            "INF": {
                "description": "理想型：富有想象力，善于创新，适合创意和设计岗位",
                "traits": ["创意", "理想", "敏感", "创新"]
            },
            "EST": {
                "description": "实干型：务实高效，善于执行，适合运营和管理岗位",
                "traits": ["务实", "高效", "果断", "组织"]
            },
            "ESF": {
                "description": "社交型：热情友好，善于沟通，适合销售和客服岗位",
                "traits": ["热情", "友好", "沟通", "合作"]
            }
        }
    
    def get_interest_questions(self) -> List[Question]:
        """获取兴趣测评题目"""
        return self.interest_questions
    
    def get_ability_questions(self) -> List[Question]:
        """获取能力测评题目"""
        return self.ability_questions
    
    def get_personality_questions(self) -> List[Question]:
        """获取性格测评题目"""
        return self.personality_questions
    
    def calculate_interest_scores(self, answers: List[int]) -> Dict[str, int]:
        """计算兴趣得分"""
        categories = ["技术型", "社交型", "艺术型", "分析型"]
        scores = {cat: 0 for cat in categories}
        
        for i, answer in enumerate(answers):
            if answer == 0:  # 编程/技术
                scores["技术型"] += 1
            elif answer == 1:  # 社交/沟通
                scores["社交型"] += 1
            elif answer == 2:  # 设计/创意
                scores["艺术型"] += 1
            elif answer == 3:  # 数据/分析
                scores["分析型"] += 1
        
        return scores
    
    def calculate_ability_scores(self, answers: List[int]) -> Dict[str, int]:
        """计算能力得分"""
        categories = ["编程能力", "学习能力", "沟通能力", "协作能力", 
                     "分析能力", "创新能力", "时间管理", "抗压能力"]
        scores = {}
        
        for i, answer in enumerate(answers):
            # 选项：很强(3)、较强(2)、一般(1)、较弱(0)
            scores[categories[i]] = 3 - answer
        
        return scores
    
    def calculate_personality_scores(self, answers: List[int]) -> Dict[str, int]:
        """计算性格得分"""
        scores = {
            "内向": 0, "外向": 0,
            "直觉": 0, "感觉": 0,
            "思考": 0, "情感": 0,
            "判断": 0, "感知": 0
        }
        
        # p1: 内向/外向
        scores["内向" if answers[0] == 0 else "外向"] += 1
        
        # p2: 直觉/感觉
        scores["直觉" if answers[1] == 0 else "感觉"] += 1
        
        # p3: 思考/情感
        scores["思考" if answers[2] == 0 else "情感"] += 1
        
        # p4: 判断/感知
        scores["判断" if answers[3] == 0 else "感知"] += 1
        
        # p5: 判断/感知 (补充)
        scores["判断" if answers[4] == 0 else "感知"] += 1
        
        # p6: 思考/情感 (补充)
        scores["思考" if answers[5] == 0 else "情感"] += 1
        
        return scores
    
    def determine_interest_type(self, scores: Dict[str, int]) -> Tuple[str, str]:
        """确定兴趣类型"""
        max_type = max(scores, key=scores.get)
        description = self.interest_types[max_type]["description"]
        return max_type, description
    
    def determine_personality_type(self, scores: Dict[str, int]) -> Tuple[str, str]:
        """确定性格类型"""
        # 判断维度
        i_e = "I" if scores["内向"] > scores["外向"] else "E"
        n_s = "N" if scores["直觉"] > scores["感觉"] else "S"
        t_f = "T" if scores["思考"] > scores["情感"] else "F"
        j_p = "J" if scores["判断"] > scores["感知"] else "P"
        
        # 简化类型
        personality_type = i_e + n_s + t_f
        
        # 获取描述
        if personality_type in self.personality_types:
            description = self.personality_types[personality_type]["description"]
        else:
            description = "你的性格比较均衡，适合多种类型的工作"
        
        return personality_type, description
    
    def get_ability_ranking(self, scores: Dict[str, int]) -> List[str]:
        """获取能力排名"""
        sorted_items = sorted(scores.items(), key=lambda x: x[1], reverse=True)
        return [item[0] for item in sorted_items]
    
    def generate_suggestions(self, interest_type: str, ability_ranking: List[str]) -> Tuple[List[str], List[str]]:
        """生成职业建议"""
        job_suggestions = self.interest_types.get(interest_type, {}).get("jobs", [])
        
        # 根据能力排名推荐技能提升方向
        skill_suggestions = []
        for ability in ability_ranking[:3]:
            if ability == "编程能力":
                skill_suggestions.append("建议深入学习Python/Java等主流编程语言")
            elif ability == "学习能力":
                skill_suggestions.append("建议多参与开源项目，积累实战经验")
            elif ability == "沟通能力":
                skill_suggestions.append("建议多参加演讲和表达训练")
            elif ability == "分析能力":
                skill_suggestions.append("建议学习数据分析工具和方法")
            elif ability == "创新能力":
                skill_suggestions.append("建议多关注行业前沿技术和趋势")
        
        return job_suggestions, skill_suggestions
    
    def complete_assessment(self, interest_answers: List[int], 
                           ability_answers: List[int],
                           personality_answers: List[int]) -> AssessmentResult:
        """完成完整测评"""
        # 计算各维度得分
        interest_scores = self.calculate_interest_scores(interest_answers)
        ability_scores = self.calculate_ability_scores(ability_answers)
        personality_scores = self.calculate_personality_scores(personality_answers)
        
        # 确定类型
        interest_type, interest_desc = self.determine_interest_type(interest_scores)
        personality_type, personality_desc = self.determine_personality_type(personality_scores)
        
        # 获取能力排名
        ability_ranking = self.get_ability_ranking(ability_scores)
        
        # 生成建议
        suggested_jobs, suggested_skills = self.generate_suggestions(interest_type, ability_ranking)
        
        # 构建结果
        result = AssessmentResult(
            interest_scores=interest_scores,
            interest_type=interest_type,
            interest_description=interest_desc,
            ability_scores=ability_scores,
            ability_ranking=ability_ranking,
            personality_scores=personality_scores,
            personality_type=personality_type,
            personality_description=personality_desc,
            suggested_jobs=suggested_jobs,
            suggested_skills=suggested_skills
        )
        
        return result
    
    def export_result(self, result: AssessmentResult, format: str = 'json') -> str:
        """导出测评结果"""
        if format == 'json':
            return json.dumps(result.to_dict(), ensure_ascii=False, indent=2)
        return str(result.to_dict())


# 全局实例
assessment_system = CareerAssessment()