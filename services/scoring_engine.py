"""
Scoring Engine - Implements weighted scoring, rules, and validation.
"""

from typing import Dict, Any, List
from core.utils import parse_number


class ScoringEngine:
    """Handles scoring calculations for PM01 and PM05."""
    
    # Weighted scoring percentages
    PRIMARY_WEIGHT = 0.60  # 60%
    SUB_WEIGHT = 0.20      # 20%
    PROCESS_WEIGHT = 0.20  # 20%
    
    # Question to category mapping (Q1-Q6)
    QUESTION_CATEGORY_MAP = {
        1: {
            'primary': '問題理解',
            'sub': '情報整理',
            'process': 'clarity'
        },
        2: {
            'primary': '論理構成',
            'sub': '因果推論',
            'process': 'structure'
        },
        3: {
            'primary': '仮説構築',
            'sub': '前提設定',
            'process': 'hypothesis'
        },
        4: {
            'primary': 'AI指示',
            'sub': '要件定義力',
            'process': 'prompt_clarity'
        },
        5: {
            'primary': 'AI成果検証力',
            'sub': '品質チェック力',
            'process': 'quality_check'
        },
        6: {
            'primary': '優先順位判断',
            'sub': '意思決定',
            'process': 'consistency'
        }
    }
    
    def __init__(self, config: Dict[str, Any]):
        """Initialize scoring engine with configuration."""
        self.config = config
    
    def aggregate_pm01_raw_scores(
        self,
        pm01_raw_results: Dict[str, Dict[str, Any]],
        questions: List[Dict[str, Any]]
    ) -> Dict[str, Any] | None:
        """
        Aggregate PM01 raw scores (without LLM analysis).
        
        Returns aggregated scores:
        {
            "scores_primary": {...},
            "scores_sub": {...},
            "process": {...},
            "aes": {...},
            "total_score": <float>,
            "per_question": {...}
        }
        """
        try:
            # Extract per-question scores
            per_question = {}
            primary_scores = {}
            sub_scores = {}
            process_scores = {}
            aes_scores = {}
            
            # Process each question Q1-Q6
            for question in questions:
                q_num = question['number']
                if q_num < 1 or q_num > 6:
                    continue
                
                question_id = f"Q{q_num}"
                q_data = pm01_raw_results.get(question_id, {})
                
                # Get category mapping for this question
                category_map = self.QUESTION_CATEGORY_MAP.get(q_num, {})
                primary_category = category_map.get('primary', '')
                sub_category = category_map.get('sub', '')
                process_item = category_map.get('process', '')
                
                # Extract scores
                primary = parse_number(q_data.get('primary_score', 0), 0)
                sub = parse_number(q_data.get('sub_score', 0), 0)
                process = parse_number(q_data.get('process_score', 0), 0)
                
                # Extract increase/decrease points
                increase_points = parse_number(q_data.get('increase_points', 0), 0)
                decrease_points = parse_number(q_data.get('decrease_points', 0), 0)
                
                # Apply increase/decrease adjustments
                primary_adjusted = max(0, min(5, primary + increase_points - decrease_points))
                sub_adjusted = max(0, min(5, sub + increase_points - decrease_points))
                process_adjusted = max(0, min(5, process + increase_points - decrease_points))
                
                # Extract AES components (clarity, logic, relevance)
                aes_clarity = parse_number(q_data.get('aes_clarity', 0), 0)
                aes_logic = parse_number(q_data.get('aes_logic', 0), 0)
                aes_relevance = parse_number(q_data.get('aes_relevance', 0), 0)
                
                # Calculate AES score: (clarity + logic + relevance) / 3
                aes_score = (aes_clarity + aes_logic + aes_relevance) / 3 if (aes_clarity + aes_logic + aes_relevance) > 0 else 0
                
                # Store per-question data
                per_question[question_id] = {
                    'primary_score': round(primary_adjusted, 2),
                    'sub_score': round(sub_adjusted, 2),
                    'process_score': round(process_adjusted, 2),
                    'aes_score': round(aes_score, 2),
                    'aes_clarity': aes_clarity,
                    'aes_logic': aes_logic,
                    'aes_relevance': aes_relevance,
                    'increase_points': increase_points,
                    'decrease_points': decrease_points,
                    'evidence': q_data.get('evidence', ''),
                    'judgment_reason': q_data.get('judgment_reason', '')
                }
                
                # Use adjusted scores for aggregation
                primary = primary_adjusted
                sub = sub_adjusted
                process = process_adjusted
                
                # Aggregate scores by category
                if primary_category:
                    if primary_category not in primary_scores:
                        primary_scores[primary_category] = []
                    primary_scores[primary_category].append(primary)
                
                if sub_category:
                    if sub_category not in sub_scores:
                        sub_scores[sub_category] = []
                    sub_scores[sub_category].append(sub)
                
                if process_item:
                    if process_item not in process_scores:
                        process_scores[process_item] = []
                    process_scores[process_item].append(process)
                
                # AES is per-question
                aes_scores[question_id] = round(aes_score, 2)
            
            # Calculate averages for aggregated scores
            primary_avg = {
                skill: sum(scores) / len(scores) if scores else 0
                for skill, scores in primary_scores.items()
            }
            sub_avg = {
                skill: sum(scores) / len(scores) if scores else 0
                for skill, scores in sub_scores.items()
            }
            process_avg = {
                skill: sum(scores) / len(scores) if scores else 0
                for skill, scores in process_scores.items()
            }
            
            # Calculate weighted total score
            # Average all primary scores
            avg_primary = sum(primary_avg.values()) / len(primary_avg) if primary_avg else 0
            avg_sub = sum(sub_avg.values()) / len(sub_avg) if sub_avg else 0
            avg_process = sum(process_avg.values()) / len(process_avg) if process_avg else 0
            avg_aes = sum(aes_scores.values()) / len(aes_scores) if aes_scores else 0
            
            # Weighted total: PRIMARY 60% + SUB 20% + PROCESS 20%
            # AES is not included in total score (used as supplementary indicator)
            weighted_total = (
                avg_primary * self.PRIMARY_WEIGHT +
                avg_sub * self.SUB_WEIGHT +
                avg_process * self.PROCESS_WEIGHT
            )
            
            return {
                'scores_primary': primary_avg,
                'scores_sub': sub_avg,
                'process': process_avg,
                'aes': aes_scores,
                'total_score': round(weighted_total, 2),
                'per_question': per_question
            }
            
        except Exception as e:
            print(f"Error aggregating PM01 raw scores: {e}")
            return None
    
    def combine_pm01_final(
        self,
        aggregated_scores: Dict[str, Any],
        pm01_final_analysis: Dict[str, Any],
        pm01_raw_results: Dict[str, Dict[str, Any]]
    ) -> Dict[str, Any] | None:
        """
        STEP 3: Combine aggregated scores with LLM analysis.
        
        Returns structured PM01 Final result:
        {
            "scores_primary": {...},
            "scores_sub": {...},
            "process": {...},
            "aes": {...},
            "total_score": <float>,
            "per_question": {...},
            "top_strengths": [...],
            "top_weaknesses": [...],
            "overall_summary": "<string>",
            "ai_use_level": "<string>",
            "recommendations": [...],
            "debug_raw": {...}
        }
        """
        try:
            # Use LLM analysis for insights
            top_strengths = pm01_final_analysis.get('top_strengths', [])
            top_weaknesses = pm01_final_analysis.get('top_weaknesses', [])
            overall_summary = pm01_final_analysis.get('overall_summary', '')
            ai_use_level = pm01_final_analysis.get('ai_use_level', '標準')
            recommendations = pm01_final_analysis.get('recommendations', [])
            
            # Combine aggregated scores with LLM analysis
            return {
                'scores_primary': aggregated_scores.get('scores_primary', {}),
                'scores_sub': aggregated_scores.get('scores_sub', {}),
                'process': aggregated_scores.get('process', {}),
                'aes': aggregated_scores.get('aes', {}),
                'total_score': aggregated_scores.get('total_score', 0),
                'per_question': aggregated_scores.get('per_question', {}),
                'top_strengths': top_strengths,
                'top_weaknesses': top_weaknesses,
                'overall_summary': overall_summary,
                'ai_use_level': ai_use_level,
                'recommendations': recommendations,
                'debug_raw': pm01_raw_results  # Store raw results for debugging
            }
            
        except Exception as e:
            print(f"Error combining PM01 Final: {e}")
            return None
    
    def process_pm05_final(
        self,
        pm05_llm_response: Dict[str, Any],
        pm01_final: Dict[str, Any]
    ) -> Dict[str, Any] | None:
        """
        STEP 4: Process PM05 Final consistency check result.
        
        Returns PM05 Final result:
        {
            "status": "valid|caution|re-evaluate",
            "consistency_score": <1-5>,
            "issues": [...],
            "summary": "<string>"
        }
        """
        try:
            consistency_score = pm05_llm_response.get('consistency_score', 0)
            status = pm05_llm_response.get('status', 'caution')
            issues = pm05_llm_response.get('issues', [])
            summary = pm05_llm_response.get('summary', '')
            
            return {
                'status': status,
                'consistency_score': round(consistency_score, 2),
                'issues': issues,
                'summary': summary
            }
        except Exception as e:
            print(f"Error processing PM05 Final: {e}")
            return None
    
    def _identify_top_items(
        self,
        primary_avg: Dict[str, float],
        sub_avg: Dict[str, float],
        process_avg: Dict[str, float],
        top_n: int,
        is_strength: bool
    ) -> List[Dict[str, Any]]:
        """Identify top strengths or weaknesses."""
        all_items = []
        
        for skill, score in primary_avg.items():
            all_items.append({'category': 'primary', 'skill': skill, 'score': score})
        for skill, score in sub_avg.items():
            all_items.append({'category': 'sub', 'skill': skill, 'score': score})
        for skill, score in process_avg.items():
            all_items.append({'category': 'process', 'skill': skill, 'score': score})
        
        # Sort by score (descending for strengths, ascending for weaknesses)
        sorted_items = sorted(all_items, key=lambda x: x['score'], reverse=is_strength)
        
        return sorted_items[:top_n]
    
    def _generate_summary(
        self,
        total_score: float,
        primary_avg: Dict[str, float],
        sub_avg: Dict[str, float],
        process_avg: Dict[str, float]
    ) -> str:
        """Generate overall summary based on scores."""
        if total_score >= 4.0:
            level = "強い"
        elif total_score >= 2.6:
            level = "標準"
        else:
            level = "弱い"
        
        return f"総合スコア: {total_score:.2f} ({level})"
    
    def _determine_ai_use_level(
        self,
        total_score: float,
        process_avg: Dict[str, float]
    ) -> str:
        """Determine AI use level based on scores."""
        prompt_clarity = process_avg.get('prompt_clarity', 0)
        quality_check = process_avg.get('quality_check', 0)
        
        if prompt_clarity >= 4.0 and quality_check >= 4.0:
            return "高度"
        elif prompt_clarity >= 3.0 and quality_check >= 3.0:
            return "標準"
        else:
            return "基礎"
    
    def calculate_pm05_validation(
        self,
        pm01_result: Dict[str, Any],
        pm05_llm_response: Dict[str, Any],
        questions: List[Dict[str, Any]]
    ) -> Dict[str, Any] | None:
        """
        Calculate PM05 validation results.
        
        Returns PM05 result:
        {
            "status": "valid | caution | re-evaluate",
            "consistency_score": 1-5,
            "issues": [],
            "comment": "<string>"
        }
        """
        try:
            # Extract reverse-scored results from PM05 LLM response
            pm05_scores = pm05_llm_response.get('reverse_scores', {})
            
            # Compare PM01 vs PM05 scores
            issues = []
            consistency_scores = []
            
            # Compare per-question scores
            pm01_per_q = pm01_result.get('per_question', {})
            
            for question in questions:
                q_num = question['number']
                if q_num < 1 or q_num > 6:
                    continue
                
                question_id = f"Q{q_num}"
                pm01_q = pm01_per_q.get(question_id, {})
                pm05_q = pm05_scores.get(question_id, {})
                
                # Compare scores
                pm01_total = (
                    pm01_q.get('primary_score', 0) * self.PRIMARY_WEIGHT +
                    pm01_q.get('sub_score', 0) * self.SUB_WEIGHT +
                    pm01_q.get('process_score', 0) * self.PROCESS_WEIGHT
                )
                pm05_total = parse_number(pm05_q.get('total_score', 0), 0)
                
                # Calculate difference
                diff = abs(pm01_total - pm05_total)
                
                # Consistency score: 5 = perfect match, 1 = large difference
                if diff < 0.5:
                    consistency = 5
                elif diff < 1.0:
                    consistency = 4
                elif diff < 1.5:
                    consistency = 3
                elif diff < 2.0:
                    consistency = 2
                else:
                    consistency = 1
                    issues.append(f"{question_id}: Large score difference ({diff:.2f})")
                
                consistency_scores.append(consistency)
            
            # Overall consistency score
            avg_consistency = sum(consistency_scores) / len(consistency_scores) if consistency_scores else 0
            
            # Determine status
            if avg_consistency >= 4.5:
                status = "valid"
            elif avg_consistency >= 3.5:
                status = "caution"
            else:
                status = "re-evaluate"
            
            # Get comment from LLM response
            comment = pm05_llm_response.get('validation_comment', '')
            
            return {
                'status': status,
                'consistency_score': round(avg_consistency, 2),
                'issues': issues,
                'comment': comment,
                'raw_pm05_response': pm05_llm_response
            }
            
        except Exception as e:
            print(f"Error calculating PM05 validation: {e}")
            return None
    

