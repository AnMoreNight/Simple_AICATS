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
    
    def __init__(self, config: Dict[str, Any]):
        """Initialize scoring engine with configuration."""
        self.config = config
    
    def calculate_pm01_scores(
        self,
        llm_response: Dict[str, Any],
        questions: List[Dict[str, Any]]
    ) -> Dict[str, Any] | None:
        """
        Calculate PM01 scores from LLM response.
        
        Returns structured PM01 result:
        {
            "primary_scores": {...},
            "sub_scores": {...},
            "process_scores": {...},
            "aes_scores": {...},
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
                q_data = llm_response.get('per_question', {}).get(question_id, {})
                
                # Extract scores
                primary = parse_number(q_data.get('primary_score', 0), 0)
                sub = parse_number(q_data.get('sub_score', 0), 0)
                process = parse_number(q_data.get('process_score', 0), 0)
                aes = parse_number(q_data.get('aes_score', 0), 0)
                
                # Store per-question data
                per_question[question_id] = {
                    'primary_score': primary,
                    'sub_score': sub,
                    'process_score': process,
                    'aes_score': aes,
                    'comment': q_data.get('comment', '')
                }
                
                # Aggregate scores by skill type
                primary_skill = question.get('primary_skill', '')
                sub_skill = question.get('sub_skill', '')
                process_skill = question.get('process_skill', '')
                
                if primary_skill:
                    if primary_skill not in primary_scores:
                        primary_scores[primary_skill] = []
                    primary_scores[primary_skill].append(primary)
                
                if sub_skill:
                    if sub_skill not in sub_scores:
                        sub_scores[sub_skill] = []
                    sub_scores[sub_skill].append(sub)
                
                if process_skill:
                    if process_skill not in process_scores:
                        process_scores[process_skill] = []
                    process_scores[process_skill].append(process)
                
                # AES is per-question
                aes_scores[question_id] = aes
            
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
            weighted_total = (
                avg_primary * self.PRIMARY_WEIGHT +
                avg_sub * self.SUB_WEIGHT +
                avg_process * self.PROCESS_WEIGHT
            )
            
            # Apply increase/decrease rules if configured
            weighted_total = self._apply_scoring_rules(weighted_total, llm_response)
            
            return {
                'primary_scores': primary_avg,
                'sub_scores': sub_avg,
                'process_scores': process_avg,
                'aes_scores': aes_scores,
                'total_score': round(weighted_total, 2),
                'per_question': per_question,
                'raw_llm_response': llm_response
            }
            
        except Exception as e:
            print(f"Error calculating PM01 scores: {e}")
            return None
    
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
    
    def _apply_scoring_rules(self, base_score: float, llm_response: Dict[str, Any]) -> float:
        """
        Apply increase/decrease scoring rules.
        
        Returns adjusted score.
        """
        # Check for increase/decrease flags in LLM response
        increase_factor = llm_response.get('increase_factor', 0)
        decrease_factor = llm_response.get('decrease_factor', 0)
        
        # Apply adjustments
        adjusted_score = base_score + increase_factor - decrease_factor
        
        # Clamp to valid range (0-5)
        return max(0.0, min(5.0, adjusted_score))

