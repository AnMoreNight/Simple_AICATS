"""
Validation service - validates respondent data.
"""

from typing import Dict, List, Any
from datetime import datetime
from core.utils import is_empty


class ValidationService:
    """Service for validating respondent data."""
    
    REQUIRED_FIELDS = ['id', 'name']
    MAX_ANSWERS = 6  # Q1-Q6 only
    MAX_ANSWER_LENGTH = 400
    
    def __init__(self, sheets_service):
        """Initialize the validation service."""
        self.sheets = sheets_service
    
    def validate_respondents(self, rows: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Verifies respondent rows meet structural requirements."""
        valid = []
        errors = []
        now = datetime.now()
        
        if not rows:
            print("validateRespondents: No rows provided")
            return {'valid': valid, 'errors': errors}
        
        for row in rows:
            # Check required fields
            missing_fields = []
            for field in self.REQUIRED_FIELDS:
                if is_empty(row.get(field)):
                    missing_fields.append(field)
            
            if missing_fields:
                errors.append({
                    'rowIndex': row.get('rowIndex', 0),
                    'respondentId': row.get('id', ''),
                    'reason': f"Missing fields: {', '.join(missing_fields)}",
                    'timestamp': now
                })
                self.sheets.update_respondent_status(row.get('rowIndex', 0), "入力不足により無効")
                continue
            
            # Check that answers array exists and has the correct length
            answers = row.get('answers', [])
            if not isinstance(answers, list):
                errors.append({
                    'rowIndex': row.get('rowIndex', 0),
                    'respondentId': row.get('id', ''),
                    'reason': 'Answers array is missing or invalid',
                    'timestamp': now
                })
                self.sheets.update_respondent_status(row.get('rowIndex', 0), "回答データ不正")
                continue
            
            # Check that exactly MAX_ANSWERS are non-empty answers
            non_empty_count = sum(1 for a in answers if not is_empty(a))
            if non_empty_count != self.MAX_ANSWERS:
                errors.append({
                    'rowIndex': row.get('rowIndex', 0),
                    'respondentId': row.get('id', ''),
                    'reason': f'Expected {self.MAX_ANSWERS} non-empty answers, received {non_empty_count}',
                    'timestamp': now
                })
                self.sheets.update_respondent_status(row.get('rowIndex', 0), "回答数不足により無効")
                continue
            
            # Check answer length limits
            over_limit = [
                {'index': i, 'length': len(answer)}
                for i, answer in enumerate(answers)
                if len(answer) > self.MAX_ANSWER_LENGTH
            ]
            
            if over_limit:
                descriptors = ', '.join([f"Q{item['index'] + 1} ({item['length']})" for item in over_limit])
                errors.append({
                    'rowIndex': row.get('rowIndex', 0),
                    'respondentId': row.get('id', ''),
                    'reason': f'Answers exceed max length: {descriptors}',
                    'timestamp': now
                })
                self.sheets.update_respondent_status(row.get('rowIndex', 0), "回答文字数超過")
                continue
            
            valid.append(row)
        
        print(f"validateRespondents: {len(rows)} rows processed, {len(valid)} valid, {len(errors)} errors")
        
        return {'valid': valid, 'errors': errors}

