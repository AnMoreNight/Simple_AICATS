"""
JSON parser service - parses PM01 and PM05 LLM responses with error recovery.
"""

import json
import re
from typing import Dict, Any, Optional
from core.utils import safe_json_parse, parse_number


class JsonParser:
    """Service for parsing JSON responses from LLM."""
    
    def parse_pm01_raw_response(self, raw: str, question_number: int) -> Optional[Dict[str, Any]]:
        """Parse PM01 Raw Scoring response for a single question."""
        parsed = self._parse_with_repair(raw)
        if not parsed or not isinstance(parsed, dict):
            return None
        
        # Validate scores are numbers
        for score_key in ['primary_score', 'sub_score', 'process_score']:
            score = parse_number(parsed.get(score_key), float('nan'))
            if score == float('nan') or score < 0 or score > 5:
                print(f"Warning: Invalid {score_key} for Q{question_number}: {parsed.get(score_key)}")
                return None
            parsed[score_key] = round(score, 2)
        
        # Validate AES components (clarity, logic, relevance)
        for aes_key in ['aes_clarity', 'aes_logic', 'aes_relevance']:
            score = parse_number(parsed.get(aes_key), float('nan'))
            if score == float('nan') or score < 0 or score > 5:
                print(f"Warning: Invalid {aes_key} for Q{question_number}: {parsed.get(aes_key)}")
                return None
            parsed[aes_key] = round(score, 2)
        
        # Validate increase/decrease points (optional, default 0)
        parsed['increase_points'] = parse_number(parsed.get('increase_points', 0), 0)
        parsed['decrease_points'] = parse_number(parsed.get('decrease_points', 0), 0)
        
        # Validate evidence and judgment_reason (required)
        if not parsed.get('evidence') or not parsed.get('judgment_reason'):
            print(f"Warning: Missing evidence or judgment_reason for Q{question_number}")
            return None
        
        # Add question number for reference
        parsed['question_number'] = question_number
        
        return parsed
    
    def parse_pm05_raw_response(self, raw: str, question_number: int) -> Optional[Dict[str, Any]]:
        """Parse PM05 Raw Scoring response for a single question."""
        parsed = self._parse_with_repair(raw)
        if not parsed or not isinstance(parsed, dict):
            return None
        
        # Validate scores are numbers
        for score_key in ['primary_score', 'sub_score', 'process_score']:
            score = parse_number(parsed.get(score_key), float('nan'))
            if score == float('nan') or score < 0 or score > 5:
                print(f"Warning: Invalid {score_key} for Q{question_number}: {parsed.get(score_key)}")
                return None
            parsed[score_key] = round(score, 2)
        
        # Add question number for reference
        parsed['question_number'] = question_number
        
        return parsed
    
    def parse_pm01_final_response(self, raw: str) -> Optional[Dict[str, Any]]:
        """Parse PM01 Final analysis response."""
        parsed = self._parse_with_repair(raw)
        if not parsed or not isinstance(parsed, dict):
            return None
        
        # Validate top_strengths (should be a list)
        if 'top_strengths' not in parsed or not isinstance(parsed['top_strengths'], list):
            parsed['top_strengths'] = []
        
        # Validate top_weaknesses (should be a list)
        if 'top_weaknesses' not in parsed or not isinstance(parsed['top_weaknesses'], list):
            parsed['top_weaknesses'] = []
        
        # Validate overall_summary (required)
        if not parsed.get('overall_summary'):
            print("Warning: Missing overall_summary in PM01 Final response")
            return None
        
        # Validate ai_use_level
        ai_use_level = parsed.get('ai_use_level', '')
        if ai_use_level not in ['基礎', '標準', '高度']:
            print(f"Warning: Invalid ai_use_level: {ai_use_level}")
            parsed['ai_use_level'] = '標準'  # Default
        
        # Validate recommendations (optional, should be a list)
        if 'recommendations' not in parsed or not isinstance(parsed['recommendations'], list):
            parsed['recommendations'] = []
        
        return parsed
    
    def parse_pm05_final_response(self, raw: str) -> Optional[Dict[str, Any]]:
        """Parse PM05 Final consistency check response."""
        parsed = self._parse_with_repair(raw)
        if not parsed or not isinstance(parsed, dict):
            return None
        
        # Validate consistency_score
        consistency_score = parse_number(parsed.get('consistency_score'), float('nan'))
        if consistency_score == float('nan') or consistency_score < 1 or consistency_score > 5:
            print(f"Warning: Invalid consistency_score: {parsed.get('consistency_score')}")
            return None
        parsed['consistency_score'] = round(consistency_score, 2)
        
        # Validate status
        status = parsed.get('status', '').lower()
        if status not in ['valid', 'caution', 're-evaluate']:
            print(f"Warning: Invalid status: {parsed.get('status')}")
            return None
        
        # Validate issues (should be a list)
        if 'issues' not in parsed or not isinstance(parsed['issues'], list):
            parsed['issues'] = []
        
        return parsed
    
    def parse_pm05_response(self, raw: str) -> Optional[Dict[str, Any]]:
        """Parse PM05 LLM response."""
        parsed = self._parse_with_repair(raw)
        if not parsed or not isinstance(parsed, dict):
            return None
        
        # Validate structure
        reverse_scores = parsed.get('reverse_scores', {})
        if not reverse_scores:
            return None
        
        # Validate all Q1-Q6 are present
        for i in range(1, 7):
            q_id = f"Q{i}"
            if q_id not in reverse_scores:
                print(f"Warning: Missing {q_id} in PM05 response")
                return None
            
            q_data = reverse_scores[q_id]
            if not isinstance(q_data, dict):
                return None
            
            # Validate total_score
            total_score = parse_number(q_data.get('total_score'), float('nan'))
            if total_score == float('nan') or total_score < 0 or total_score > 5:
                print(f"Warning: Invalid total_score for {q_id}: {q_data.get('total_score')}")
                return None
            q_data['total_score'] = round(total_score, 2)
        
        return parsed
    
    def _parse_with_repair(self, raw: str) -> Optional[Any]:
        """Applies heuristic fixes to malformed JSON strings."""
        sanitized = self._strip_json_code_fence(raw)
        if not sanitized:
            return None
        
        direct = safe_json_parse(sanitized)
        if direct is not None:
            return direct
        
        return self._attempt_repair(sanitized)
    
    def _attempt_repair(self, raw: str) -> Optional[Any]:
        """Attempts to fix common JSON formatting issues."""
        try:
            text = raw.strip()
            if not text.startswith('{'):
                first_brace = text.find('{')
                if first_brace >= 0:
                    text = text[first_brace:]
            
            if not text.endswith('}'):
                text = text + '}'
            
            # Remove trailing commas
            text = re.sub(r',(\s*[}\]])', r'\1', text)
            
            return json.loads(text)
        except (json.JSONDecodeError, Exception) as e:
            print(f"Auto repair failed: {e}")
            return None
    
    def _strip_json_code_fence(self, raw: str) -> Optional[str]:
        """Strips markdown code fences from JSON strings."""
        if not raw:
            return None
        
        text = raw.strip()
        if not text:
            return None
        
        if text.startswith('```'):
            first_line_end = text.find('\n')
            if first_line_end >= 0:
                fence_label = text[3:first_line_end].strip().lower()
                if not fence_label or fence_label in ['json', 'javascript']:
                    closing = text.rfind('```')
                    if closing > first_line_end:
                        text = text[first_line_end + 1:closing].strip()
        
        return text if text else None
