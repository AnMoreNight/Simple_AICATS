"""
Scoring Engine - Implements weighted scoring, rules, and validation.
"""

from typing import Dict, Any, List, Optional
from core.utils import parse_number
from core.category_mapper import (
    map_to_official_category,
    OFFICIAL_PRIMARY_CATEGORIES,
    OFFICIAL_SUB_CATEGORIES,
    OFFICIAL_PROCESS_ITEMS
)


class ScoringEngine:
    """Handles scoring calculations for PM01 and PM05."""
    
    # Weighted scoring percentages
    PRIMARY_WEIGHT = 0.60  # 60%
    SUB_WEIGHT = 0.20      # 20%
    PROCESS_WEIGHT = 0.20  # 20%
    
    # Official categories (imported from shared module)
    OFFICIAL_PRIMARY_CATEGORIES = OFFICIAL_PRIMARY_CATEGORIES
    OFFICIAL_SUB_CATEGORIES = OFFICIAL_SUB_CATEGORIES
    OFFICIAL_PROCESS_ITEMS = OFFICIAL_PROCESS_ITEMS
    
    def __init__(self, config: Dict[str, Any]):
        """Initialize scoring engine with configuration."""
        self.config = config
    
    def _map_to_official_category(self, category: str, category_type: str) -> Optional[str]:
        """Map question sheet category to official category (delegates to shared mapper)."""
        return map_to_official_category(category, category_type)
    
    def _round_dict_values(self, data: Dict[str, Any], decimals: int = 1) -> Dict[str, Any]:
        """
        Round all float values in a dictionary to specified decimal places.
        This prevents floating point precision issues like 1.7999999999999998.
        """
        result = {}
        for key, value in data.items():
            if isinstance(value, float):
                result[key] = round(value, decimals)
            elif isinstance(value, dict):
                result[key] = self._round_dict_values(value, decimals)
            else:
                result[key] = value
        return result
    
    def aggregate_pm05_raw_scores(
        self,
        pm05_raw_results: Dict[str, Dict[str, Any]],
        pm01_raw_results: Dict[str, Dict[str, Any]],
        questions: List[Dict[str, Any]]
    ) -> Dict[str, Any] | None:
        """
        Aggregate PM05 Raw scores (validated scores from secondary diagnosis).
        Uses PM05 Raw for primary/sub/process scores, PM01 Raw for AES scores.
        
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
            # AES scores: collect by component (clarity, logic, relevance) not by question
            aes_clarity_list = []
            aes_logic_list = []
            aes_relevance_list = []
            
            # Process each question Q1-Q6
            for question in questions:
                q_num = question['number']
                if q_num < 1 or q_num > 6:
                    continue
                
                question_id = f"Q{q_num}"
                # Use PM05 Raw results for validated scores
                pm05_data = pm05_raw_results.get(question_id, {})
                # Use PM01 Raw results for AES scores (PM05 validates but doesn't rescore AES)
                pm01_data = pm01_raw_results.get(question_id, {})
                
                # Get category mapping for this question from the question data (required)
                question_primary_category = question.get('primary_category', '')
                question_sub_category = question.get('sub_category', '')
                question_process_item = question.get('process_category', '')
                
                # Map question sheet categories to official categories
                primary_category = self._map_to_official_category(question_primary_category, 'primary')
                sub_category = self._map_to_official_category(question_sub_category, 'sub')
                process_item = self._map_to_official_category(question_process_item, 'process')
                
                # Skip if categories cannot be mapped to official categories
                if not primary_category or not sub_category or not process_item:
                    print(f"Warning: Cannot map categories for Q{q_num} to official categories. "
                          f"Question categories: PRIMARY={question_primary_category}, SUB={question_sub_category}, PROCESS={question_process_item}. "
                          f"Skipping aggregation for this question.")
                    continue
                
                # Extract scores from PM05 Raw (validated scores)
                primary = parse_number(pm05_data.get('primary_score', 0), 0)
                sub = parse_number(pm05_data.get('sub_score', 0), 0)
                process = parse_number(pm05_data.get('process_score', 0), 0)
                
                # Extract AES scores from PM01 Raw (PM05 validates but doesn't rescore AES)
                aes_clarity = parse_number(pm01_data.get('aes_clarity', 0), 0)
                aes_logic = parse_number(pm01_data.get('aes_logic', 0), 0)
                aes_relevance = parse_number(pm01_data.get('aes_relevance', 0), 0)
                
                # Scores from PM05 Raw are already validated (include adjustments if mentioned in difference_note)
                # Use scores directly without additional adjustments
                primary_adjusted = max(1.0, min(5.0, primary))
                sub_adjusted = max(1.0, min(5.0, sub))
                process_adjusted = max(1.0, min(5.0, process))
                
                # Calculate AES score: (clarity + logic + relevance) / 3
                aes_score = (aes_clarity + aes_logic + aes_relevance) / 3 if (aes_clarity + aes_logic + aes_relevance) > 0 else 0
                
                # Store per-question data (1 decimal place)
                per_question[question_id] = {
                    'primary_score': round(primary_adjusted, 1),  # From PM05 Raw
                    'sub_score': round(sub_adjusted, 1),  # From PM05 Raw
                    'process_score': round(process_adjusted, 1),  # From PM05 Raw
                    'aes_score': round(aes_score, 1),  # From PM01 Raw
                    'aes_clarity': round(aes_clarity, 1),  # From PM01 Raw
                    'aes_logic': round(aes_logic, 1),  # From PM01 Raw
                    'aes_relevance': round(aes_relevance, 1),  # From PM01 Raw
                    'difference_note': pm05_data.get('difference_note', '')  # From PM05 Raw
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
                
                # Collect AES components for aggregation (not per-question)
                if aes_clarity > 0:
                    aes_clarity_list.append(aes_clarity)
                if aes_logic > 0:
                    aes_logic_list.append(aes_logic)
                if aes_relevance > 0:
                    aes_relevance_list.append(aes_relevance)
            
            # Calculate averages for aggregated scores
            # Ensure all official categories are included (even if empty)
            # All averages must be rounded to 1 decimal place to avoid floating point precision issues
            primary_avg = {}
            for category in self.OFFICIAL_PRIMARY_CATEGORIES:
                if category in primary_scores and primary_scores[category]:
                    avg = sum(primary_scores[category]) / len(primary_scores[category])
                    primary_avg[category] = round(avg, 1)
                else:
                    primary_avg[category] = 0.0
            
            sub_avg = {}
            for category in self.OFFICIAL_SUB_CATEGORIES:
                if category in sub_scores and sub_scores[category]:
                    avg = sum(sub_scores[category]) / len(sub_scores[category])
                    sub_avg[category] = round(avg, 1)
                else:
                    sub_avg[category] = 0.0
            
            process_avg = {}
            for item in self.OFFICIAL_PROCESS_ITEMS:
                if item in process_scores and process_scores[item]:
                    avg = sum(process_scores[item]) / len(process_scores[item])
                    process_avg[item] = round(avg, 1)
                else:
                    process_avg[item] = 0.0
            
            # Calculate weighted total score
            # Average all primary scores (already rounded to 1 decimal)
            avg_primary = sum(primary_avg.values()) / len(primary_avg) if primary_avg else 0.0
            avg_sub = sum(sub_avg.values()) / len(sub_avg) if sub_avg else 0.0
            avg_process = sum(process_avg.values()) / len(process_avg) if process_avg else 0.0
            # Calculate AES averages by component (not per-question)
            avg_aes_clarity = sum(aes_clarity_list) / len(aes_clarity_list) if aes_clarity_list else 0.0
            avg_aes_logic = sum(aes_logic_list) / len(aes_logic_list) if aes_logic_list else 0.0
            avg_aes_relevance = sum(aes_relevance_list) / len(aes_relevance_list) if aes_relevance_list else 0.0
            avg_aes = (avg_aes_clarity + avg_aes_logic + avg_aes_relevance) / 3 if (aes_clarity_list or aes_logic_list or aes_relevance_list) else 0.0
            
            # Round intermediate averages to 1 decimal place
            avg_primary = round(avg_primary, 1)
            avg_sub = round(avg_sub, 1)
            avg_process = round(avg_process, 1)
            avg_aes_clarity = round(avg_aes_clarity, 1)
            avg_aes_logic = round(avg_aes_logic, 1)
            avg_aes_relevance = round(avg_aes_relevance, 1)
            avg_aes = round(avg_aes, 1)
            
            # Weighted total: PRIMARY 60% + SUB 20% + PROCESS 20%
            # AES is not included in total score (used as supplementary indicator)
            weighted_total = (
                avg_primary * self.PRIMARY_WEIGHT +
                avg_sub * self.SUB_WEIGHT +
                avg_process * self.PROCESS_WEIGHT
            )
            
            # AES output: aggregated by component (not per-question)
            aes_output = {
                'aes_clarity': avg_aes_clarity,
                'aes_logic': avg_aes_logic,
                'aes_relevance': avg_aes_relevance
            }
            
            # Round all dictionary values to ensure no floating point precision issues
            result = {
                'scores_primary': self._round_dict_values(primary_avg, 1),
                'scores_sub': self._round_dict_values(sub_avg, 1),
                'process': self._round_dict_values(process_avg, 1),
                'aes': self._round_dict_values(aes_output, 1),
                'total_score': round(weighted_total, 1),
                'per_question': self._round_dict_values(per_question, 1)
            }
            
            return result
            
        except Exception as e:
            print(f"Error aggregating PM01 raw scores: {e}")
            return None
    
    def combine_pm01_final(
        self,
        aggregated_scores: Dict[str, Any],
        pm01_final_analysis: Dict[str, Any],
        pm05_raw_results: Dict[str, Dict[str, Any]]
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
            "overall_summary": "<string>",
            "ai_use_level": "<string>",
            "recommendations": [...],
            "debug_raw": {...}
        }
        """
        try:
            # Use LLM analysis for insights
            overall_summary = pm01_final_analysis.get('overall_summary', '')
            ai_use_level = pm01_final_analysis.get('ai_use_level', '標準')
            recommendations = pm01_final_analysis.get('recommendations', [])
            
            # Combine aggregated scores with LLM analysis
            # Ensure all scores are properly rounded (double-check to prevent any precision issues)
            scores_primary = aggregated_scores.get('scores_primary', {})
            scores_sub = aggregated_scores.get('scores_sub', {})
            process_scores = aggregated_scores.get('process', {})
            aes_scores = aggregated_scores.get('aes', {})
            per_question = aggregated_scores.get('per_question', {})
            
            return {
                'scores_primary': self._round_dict_values(scores_primary, 1),
                'scores_sub': self._round_dict_values(scores_sub, 1),
                'process': self._round_dict_values(process_scores, 1),
                'aes': self._round_dict_values(aes_scores, 1),
                'total_score': round(aggregated_scores.get('total_score', 0), 1),
                'per_question': self._round_dict_values(per_question, 1),
                'overall_summary': overall_summary,
                'ai_use_level': ai_use_level,
                'recommendations': recommendations,
                'debug_raw': pm05_raw_results  # Store PM05 Raw results for debugging
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
            "status": "妥当|注意|再評価",
            "consistency_score": <0.0-1.0>,
            "detected_issues": [...],
            "comment": "<string 80-120 chars in Japanese>"
        }
        """
        try:
            consistency_score = pm05_llm_response.get('consistency_score', 0)
            # Validate consistency_score range (0.0-1.0)
            if consistency_score < 0.0 or consistency_score > 1.0:
                print(f"Warning: consistency_score out of range: {consistency_score}, clamping to 0.0-1.0")
                consistency_score = max(0.0, min(1.0, consistency_score))
            
            status = pm05_llm_response.get('status', '注意')
            # Ensure status is in Japanese
            status_map = {
                'valid': '妥当',
                'caution': '注意',
                're-evaluate': '再評価',
                'reevaluate': '再評価'
            }
            status_lower = status.lower()
            if status_lower in status_map:
                status = status_map[status_lower]
            elif status not in ['妥当', '注意', '再評価']:
                status = '注意'  # Default to 注意 if invalid
            
            # Accept both 'detected_issues' and 'issues' for backward compatibility
            detected_issues = pm05_llm_response.get('detected_issues', pm05_llm_response.get('issues', []))
            if not isinstance(detected_issues, list):
                detected_issues = []
            
            # Accept both 'comment' and 'summary' for backward compatibility
            comment = pm05_llm_response.get('comment', pm05_llm_response.get('summary', ''))
            
            return {
                'status': status,
                'consistency_score': round(consistency_score, 2),  # 2 decimal places for 0-1 range
                'detected_issues': detected_issues,
                'comment': comment
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
                'consistency_score': round(avg_consistency, 1),
                'issues': issues,
                'comment': comment,
                'raw_pm05_response': pm05_llm_response
            }
            
        except Exception as e:
            print(f"Error calculating PM05 validation: {e}")
            return None
    

