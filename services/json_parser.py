"""
JSON parser service - parses PM01 and PM05 LLM responses with error recovery.
"""

import json
import re
from typing import Dict, Any, Optional
from core.utils import safe_json_parse, parse_number


class JsonParser:
    """Service for parsing JSON responses from LLM."""
    
    def parse_pm01_response(self, raw: str) -> Optional[Dict[str, Any]]:
        """Parse PM01 LLM response."""
        parsed = self._parse_with_repair(raw)
        if not parsed or not isinstance(parsed, dict):
            return None
        
        # Validate structure
        per_question = parsed.get('per_question', {})
        if not per_question:
            return None
        
        # Validate all Q1-Q6 are present
        for i in range(1, 7):
            q_id = f"Q{i}"
            if q_id not in per_question:
                print(f"Warning: Missing {q_id} in PM01 response")
                return None
            
            q_data = per_question[q_id]
            if not isinstance(q_data, dict):
                return None
            
            # Validate scores are numbers
            for score_key in ['primary_score', 'sub_score', 'process_score', 'aes_score']:
                score = parse_number(q_data.get(score_key), float('nan'))
                if score == float('nan') or score < 0 or score > 5:
                    print(f"Warning: Invalid {score_key} for {q_id}: {q_data.get(score_key)}")
                    return None
                q_data[score_key] = round(score, 2)
        
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
