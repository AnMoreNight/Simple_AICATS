"""
Google Sheets service - handles all sheet operations for PM01/PM05.
"""

import gspread
import re
import json
from google.oauth2.service_account import Credentials
from typing import List, Dict, Any, Optional
from datetime import datetime
import os


class SheetsService:
    """Service for interacting with Google Sheets."""
    
    DIAGNOSIS_QUESTION_COUNT = 6  # Q1-Q6 only
    MAX_ANSWER_LENGTH = 400
    
    def __init__(self, config: Optional[Dict[str, Any]] = None):
        """Initialize the sheets service with configuration."""
        self.config = config
        self._spreadsheet: Optional[gspread.Spreadsheet] = None
        self._init_client()
    
    def _init_client(self):
        """Initialize the Google Sheets client."""
        scope = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
        
        # Require JSON content in environment variable
        creds_json = os.getenv('GOOGLE_APPLICATION_CREDENTIALS_JSON')
        if not creds_json:
            raise ValueError(
                "GOOGLE_APPLICATION_CREDENTIALS_JSON environment variable is required. "
                "Set it with the JSON content of your service account credentials."
            )
        
        try:
            creds_info = json.loads(creds_json)
            creds = Credentials.from_service_account_info(creds_info, scopes=scope)
        except json.JSONDecodeError as e:
            raise ValueError(
                f"Invalid JSON in GOOGLE_APPLICATION_CREDENTIALS_JSON: {e}. "
                "Please provide valid JSON credentials."
            )
        
        self._client = gspread.authorize(creds)
        
        spreadsheet_id = os.getenv('SPREADSHEET_ID')
        if not spreadsheet_id:
            raise ValueError("SPREADSHEET_ID environment variable is required.")
        
        self._spreadsheet = self._client.open_by_key(spreadsheet_id)
    
    def _require_config(self):
        """Ensures config is set before using sheet operations."""
        if not self.config:
            raise ValueError("Config not set. Set sheets.config before calling sheet operations.")
    
    def _get_sheet(self, sheet_name: str) -> Optional[gspread.Worksheet]:
        """Gets a sheet by name, returns None if not found."""
        if not self._spreadsheet:
            return None
        try:
            return self._spreadsheet.worksheet(sheet_name)
        except gspread.exceptions.WorksheetNotFound:
            return None
    
    def get_respondent_rows(self) -> List[Dict[str, Any]]:
        """Reads respondent data rows from the configured sheet."""
        self._require_config()
        sheet = self._get_sheet(self.config['respondentsSheet'])
        if not sheet:
            print(f"getRespondentRows: Sheet '{self.config['respondentsSheet']}' not found")
            return []
        
        values = sheet.get_all_values()
        if len(values) <= 1:
            print("getRespondentRows: No data rows found")
            return []
        
        rows = []
        for i, row in enumerate(values[1:], start=2):
            if len(row) < 18:
                continue
            
            respondent_id = str(row[2] or '').strip()
            family_name = str(row[3] or '').strip()
            given_name = str(row[4] or '').strip()
            name = f"{family_name} {given_name}".strip() if family_name and given_name else (family_name or given_name or '').strip()
            
            # Q1-Q6 are at indices 12-17
            answers = [self._sanitize_answer(str(row[i] or '')) for i in range(12, 18)]
            status = str(row[22] or '').strip() if len(row) > 22 else ''
            
            rows.append({
                'id': respondent_id,
                'name': name,
                'answers': answers,
                'rowIndex': i,
                'status': status
            })
        
        print(f"getRespondentRows: Read {len(rows)} rows")
        return rows
    
    def get_question_rows(self) -> List[Dict[str, Any]]:
        """Reads question rows from the question sheet."""
        self._require_config()
        sheet = self._get_sheet(self.config['questionSheet'])
        if not sheet:
            return []
        
        values = sheet.get_all_values()
        questions = []
        
        for row in values[1:]:
            if len(row) < 5:
                continue
            
            try:
                no = int(float(row[0]))
            except (ValueError, TypeError):
                continue
            
            # Only process questions 1-6
            if no < 1 or no > self.DIAGNOSIS_QUESTION_COUNT:
                continue
            
            question_text = str(row[1] or '').strip()
            primary_skill = str(row[2] or '').strip()
            sub_skill = str(row[3] or '').strip()
            process_skill = str(row[4] or '').strip()
            
            questions.append({
                'number': no,
                'questionText': question_text,
                'primary_skill': primary_skill,
                'sub_skill': sub_skill,
                'process_skill': process_skill
            })
        
        return sorted(questions, key=lambda q: q['number'])
    
    def _sanitize_answer(self, answer: str) -> str:
        """Sanitizes an answer string."""
        return answer.strip()[:self.MAX_ANSWER_LENGTH]
    
    def log_validation_errors(self, errors: List[Dict[str, Any]]):
        """Writes validation error entries to the validation log sheet."""
        if not errors or not self._spreadsheet:
            return
        self._require_config()
        
        sheet = self._get_sheet(self.config['validationLogSheet'])
        if not sheet:
            sheet = self._spreadsheet.add_worksheet(
                title=self.config['validationLogSheet'],
                rows=1000,
                cols=10
            )
            sheet.append_row(['Timestamp', 'RowIndex', 'RespondentId', 'Reason'])
        
        rows = []
        for error in errors:
            rows.append([
                self._format_date(error['timestamp']),
                error['rowIndex'],
                error['respondentId'],
                error['reason']
            ])
        
        if rows:
            sheet.append_rows(rows)
    
    def log_error(self, error: Dict[str, Any]):
        """Appends a single API error row to the error log sheet."""
        if not self._spreadsheet:
            return
        self._require_config()
        
        sheet = self._get_sheet(self.config['errorLogSheet'])
        if not sheet:
            sheet = self._spreadsheet.add_worksheet(
                title=self.config['errorLogSheet'],
                rows=1000,
                cols=10
            )
            sheet.append_row(['Timestamp', 'RespondentId', 'Category', 'Message', 'Details', 'Attempt'])
        
        details_json = json.dumps(error.get('details') or {})
        
        sheet.append_row([
            self._format_date(error['timestamp']),
            error.get('respondentId', ''),
            error['category'],
            error['message'],
            details_json,
            error.get('attempt', 1)
        ])
    
    def append_pm01_result(self, respondent: Dict[str, Any], pm01_result: Dict[str, Any]):
        """Writes PM01 result to PM01 sheet."""
        if not self._spreadsheet:
            return
        
        sheet = self._get_sheet('PM01')
        if not sheet:
            sheet = self._spreadsheet.add_worksheet(title='PM01', rows=1000, cols=50)
            # Create headers
            headers = ['Respondent_ID', 'Timestamp', 'Total_Score']
            for i in range(1, 7):
                headers.extend([f'Q{i}_Primary', f'Q{i}_Sub', f'Q{i}_Process', f'Q{i}_AES', f'Q{i}_Comment'])
            headers.extend(['Primary_Scores_JSON', 'Sub_Scores_JSON', 'Process_Scores_JSON', 'AES_Scores_JSON'])
            sheet.append_row(headers)
        
        # Build row
        row = [
            respondent['id'],
            self._format_date(datetime.now()),
            pm01_result.get('total_score', 0)
        ]
        
        # Add per-question scores
        per_question = pm01_result.get('per_question', {})
        for i in range(1, 7):
            q_id = f"Q{i}"
            q_data = per_question.get(q_id, {})
            row.extend([
                q_data.get('primary_score', 0),
                q_data.get('sub_score', 0),
                q_data.get('process_score', 0),
                q_data.get('aes_score', 0),
                q_data.get('comment', '')
            ])
        
        # Add aggregated scores as JSON
        row.extend([
            json.dumps(pm01_result.get('primary_scores', {}), ensure_ascii=False),
            json.dumps(pm01_result.get('sub_scores', {}), ensure_ascii=False),
            json.dumps(pm01_result.get('process_scores', {}), ensure_ascii=False),
            json.dumps(pm01_result.get('aes_scores', {}), ensure_ascii=False)
        ])
        
        sheet.append_row(row)
    
    def append_pm05_result(self, respondent: Dict[str, Any], pm05_result: Dict[str, Any]):
        """Writes PM05 result to PM05 sheet."""
        if not self._spreadsheet:
            return
        
        sheet = self._get_sheet('PM05')
        if not sheet:
            sheet = self._spreadsheet.add_worksheet(title='PM05', rows=1000, cols=20)
            sheet.append_row([
                'Respondent_ID', 'Timestamp', 'Status', 'Consistency_Score',
                'Issues_JSON', 'Comment'
            ])
        
        sheet.append_row([
            respondent['id'],
            self._format_date(datetime.now()),
            pm05_result.get('status', ''),
            pm05_result.get('consistency_score', 0),
            json.dumps(pm05_result.get('issues', []), ensure_ascii=False),
            pm05_result.get('comment', '')
        ])
    
    def read_pm01_rows(self) -> List[Dict[str, Any]]:
        """Reads PM01 results from sheet."""
        sheet = self._get_sheet('PM01')
        if not sheet:
            return []
        
        values = sheet.get_all_values()
        if len(values) <= 1:
            return []
        
        result = []
        for row in values[1:]:
            if len(row) < 1:
                continue
            result.append({
                'respondentId': str(row[0] or '').strip()
            })
        
        return result
    
    def update_respondent_status(self, row_index: int, status: str):
        """Updates the status column for a respondent row."""
        self._require_config()
        sheet = self._get_sheet(self.config['respondentsSheet'])
        if not sheet:
            return
        
        status_column = 23
        sheet.update_cell(row_index, status_column, status)
    
    def write_run_log(self, summary: Dict[str, Any]):
        """Logs a batch run summary into the persistent run log sheet."""
        if not self._spreadsheet:
            return
        
        sheet = self._get_sheet('RunLog')
        if not sheet:
            sheet = self._spreadsheet.add_worksheet(title='RunLog', rows=1000, cols=10)
            sheet.append_row(['Timestamp', 'RunId', 'Processed', 'Errors', 'Duration'])
        
        duration_ms = summary['durationMs']
        total_seconds = duration_ms // 1000
        minutes = total_seconds // 60
        seconds = total_seconds % 60
        duration_formatted = f"{duration_ms}ms({minutes}分{seconds}秒)"
        
        sheet.append_row([
            self._format_date(summary['timestamp']),
            summary['runId'],
            summary['processed'],
            summary['errors'],
            duration_formatted
        ])
    
    def _format_date(self, date: datetime) -> str:
        """Formats a date for display."""
        return date.strftime('%Y-%m-%d %H:%M:%S')
