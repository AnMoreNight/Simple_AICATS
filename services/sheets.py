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
import re
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

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
        """Reads respondent data rows from the configured sheet.
        
        Expected format:
        Column 0: No.
        Column 1: 作成日 (Creation Date)
        Column 2: お名前/姓 (Family Name)
        Column 3: お名前/名 (Given Name)
        Column 4: 所属部門（部署）名 (Department Name)
        Column 5: 会社名（法人名） (Company Name)
        Column 6: 年齢層 (Age Group)
        Column 7: Q1 Answer
        Column 8: Q1 Reason
        Column 9: Q2 Answer
        Column 10: Q2 Reason
        Column 11: Q3 Answer
        Column 12: Q3 Reason
        Column 13: Q4 Answer
        Column 14: Q4 Reason
        Column 15: Q5 Answer
        Column 16: Q5 Reason
        Column 17: Q6 Answer
        Column 18: Q6 Reason
        Column 19: Status
        """
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
            if len(row) < 20:
                continue
            
            # Column 0: No. (respondent ID)
            respondent_id = str(row[0] or '').strip()
            # Column 2: お名前/姓 (Family Name)
            family_name = str(row[2] or '').strip()
            # Column 3: お名前/名 (Given Name)
            given_name = str(row[3] or '').strip()
            name = f"{family_name} {given_name}".strip() if family_name and given_name else (family_name or given_name or '').strip()
            
            # Column 4: 所属部門（部署）名 (Department Name)
            department = str(row[4] or '').strip() if len(row) > 4 else ''
            
            # Column 5: 会社名（法人名） (Company Name)
            company_name = str(row[5] or '').strip() if len(row) > 5 else ''
            
            # Q1-Q6 answers are at columns 7, 9, 11, 13, 15, 17
            answer_columns = [7, 9, 11, 13, 15, 17]
            answers = [self._sanitize_answer(str(row[col] or '')) for col in answer_columns]
            
            # Q1-Q6 reasons are at columns 8, 10, 12, 14, 16, 18
            reason_columns = [8, 10, 12, 14, 16, 18]
            reasons = [self._sanitize_answer(str(row[col] or '')) for col in reason_columns]
            
            # Column 19: Status
            status = str(row[19] or '').strip() if len(row) > 19 else ''
            
            rows.append({
                'id': respondent_id,
                'name': name,
                'department': department,
                'company_name': company_name,
                'answers': answers,
                'reasons': reasons,
                'rowIndex': i,
                'status': status
            })
        
        print(f"getRespondentRows: Read {len(rows)} rows")
        return rows
    
    def get_question_rows(self) -> List[Dict[str, Any]]:
        """Reads question rows from the question sheet.
        
        Expected format:
        - Row 1 (Header): A1=Q1, B1=Q2, C1=Q3, D1=Q4, E1=Q5, F1=Q6
        - Row 2 (Main questions): A2=Q1_main, B2=Q2_main, C2=Q3_main, ...
        - Row 3 (Follow-up questions): A3=Q1_followup, B3=Q2_followup, C3=Q3_followup, ...
        - Row 4 (Categories): A4="PRIMARY: X, SUB: Y, PROCESS: Z" or separate columns G, H, I
        """
        self._require_config()
        sheet = self._get_sheet(self.config['questionSheet'])
        if not sheet:
            print(f"getQuestionRows: Sheet '{self.config['questionSheet']}' not found")
            return []
        
        values = sheet.get_all_values()
        if not values:
            print(f"getQuestionRows: Sheet '{self.config['questionSheet']}' is empty")
            return []
        
        print(f"getQuestionRows: Found {len(values)} rows in sheet")
        
        # Check if we have at least header row
        if len(values) < 1:
            print("getQuestionRows: No rows found")
            return []
        
        # Row 1 is header: A1=Q1, B1=Q2, C1=Q3, D1=Q4, E1=Q5, F1=Q6
        header_row = values[0]
        questions = []
        
        # Check if categories are in separate columns (G, H, I) or in row 4
        # Look for PRIMARY, SUB, PROCESS headers in row 1
        has_separate_category_columns = False
        primary_col_idx = None
        sub_col_idx = None
        process_col_idx = None
        
        if len(header_row) > 6:
            for idx, header in enumerate(header_row):
                header_upper = str(header or '').strip().upper()
                if header_upper == 'PRIMARY':
                    primary_col_idx = idx
                elif header_upper == 'SUB':
                    sub_col_idx = idx
                elif header_upper == 'PROCESS':
                    process_col_idx = idx
            
            if primary_col_idx is not None and sub_col_idx is not None and process_col_idx is not None:
                has_separate_category_columns = True
                print(f"getQuestionRows: Found separate category columns: PRIMARY={primary_col_idx+1}, SUB={sub_col_idx+1}, PROCESS={process_col_idx+1}")
        
        # Process each column (Q1-Q6)
        for col_idx in range(min(6, len(header_row))):
            # Extract question ID from header (e.g., "Q1")
            question_id = str(header_row[col_idx] or '').strip().upper()
            if not question_id.startswith('Q'):
                print(f"getQuestionRows: Column {col_idx+1} header doesn't start with Q: '{question_id}'")
                continue
            
            try:
                # Extract number from Q1, Q2, etc.
                no = int(question_id[1:])
            except (ValueError, TypeError):
                print(f"getQuestionRows: Column {col_idx+1} invalid question number: '{question_id}'")
                continue
            
            # Only process questions 1-6
            if no < 1 or no > self.DIAGNOSIS_QUESTION_COUNT:
                print(f"getQuestionRows: Column {col_idx+1} question number out of range: {no}")
                continue
            
            # Get main question from row 2 (index 1), column col_idx
            main_question = ''
            if len(values) > 1 and len(values[1]) > col_idx:
                main_question = str(values[1][col_idx] or '').strip()
            
            # Get follow-up question from row 3 (index 2), column col_idx
            follow_up_question = ''
            if len(values) > 2 and len(values[2]) > col_idx:
                follow_up_question = str(values[2][col_idx] or '').strip()
            
            # Combine main question and follow-up question
            question_text = main_question
            if follow_up_question:
                question_text += f"\n\n{follow_up_question}"
            
            # Extract categories
            primary_category = ''
            sub_category = ''
            process_category = ''
            
            if has_separate_category_columns:
                # Read from separate columns (G, H, I) - categories aligned with Q1-Q6
                # Row 2 has PRIMARY categories aligned with Q1-Q6 (col_idx maps to question)
                # Row 3 has SUB categories aligned with Q1-Q6
                # Row 4 has PROCESS categories aligned with Q1-Q6
                # But wait, if PRIMARY/SUB/PROCESS are headers in row 1, then:
                # Row 2, column G (primary_col_idx) = PRIMARY category for Q1
                # Row 2, column H (sub_col_idx) = SUB category for Q1
                # Row 2, column I (process_col_idx) = PROCESS category for Q1
                # Actually, this doesn't align well. Let's assume:
                # If separate columns exist, they're in columns G, H, I
                # And we need to read from the row that corresponds to this question
                # For now, let's read from row 2 (index 1) for all categories
                # This assumes categories are in row 2, columns G, H, I
                if len(values) > 1:
                    if primary_col_idx is not None and len(values[1]) > primary_col_idx:
                        primary_category = str(values[1][primary_col_idx] or '').strip()
                    if sub_col_idx is not None and len(values[1]) > sub_col_idx:
                        sub_category = str(values[1][sub_col_idx] or '').strip()
                    if process_col_idx is not None and len(values[1]) > process_col_idx:
                        process_category = str(values[1][process_col_idx] or '').strip()
            else:
                # Try to read from row 4 (index 3) in the same column as the question
                if len(values) > 3 and len(values[3]) > col_idx:
                    category_text = str(values[3][col_idx] or '').strip()
                    # Parse format: "PRIMARY: 問題理解, SUB: 情報整理, PROCESS: clarity"
                    if category_text:
                        # Try to parse the format
                        primary_match = re.search(r'PRIMARY:\s*([^,]+)', category_text, re.IGNORECASE)
                        sub_match = re.search(r'SUB:\s*([^,]+)', category_text, re.IGNORECASE)
                        process_match = re.search(r'PROCESS:\s*([^,]+)', category_text, re.IGNORECASE)
                        
                        if primary_match:
                            primary_category = primary_match.group(1).strip()
                        if sub_match:
                            sub_category = sub_match.group(1).strip()
                        if process_match:
                            process_category = process_match.group(1).strip()
            
            question_data = {
                'number': no,
                'questionText': question_text
            }
            
            # Add categories if found
            if primary_category:
                question_data['primary_category'] = primary_category
            if sub_category:
                question_data['sub_category'] = sub_category
            if process_category:
                question_data['process_category'] = process_category
            
            questions.append(question_data)
            
            print(f"getQuestionRows: Added Q{no} from column {col_idx+1}")
            if primary_category or sub_category or process_category:
                print(f"  Categories: PRIMARY={primary_category}, SUB={sub_category}, PROCESS={process_category}")
        
        print(f"getQuestionRows: Read {len(questions)} questions")
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
        
        # Ensure headers exist
        headers = ['Timestamp', 'RowIndex', 'RespondentId', 'Reason']
        existing_values = sheet.get_all_values()
        if len(existing_values) == 0:
            # Sheet is empty, add headers
            sheet.append_row(headers)
        elif existing_values[0] != headers:
            # First row doesn't match headers, insert headers at the top
            sheet.insert_row(headers, index=1)
        
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
        
        # Ensure headers exist
        headers = ['Timestamp', 'RespondentId', 'Category', 'Message', 'Details', 'Attempt']
        existing_values = sheet.get_all_values()
        if len(existing_values) == 0:
            # Sheet is empty, add headers
            sheet.append_row(headers)
        elif existing_values[0] != headers:
            # First row doesn't match headers, insert headers at the top
            sheet.insert_row(headers, index=1)
        
        details_json = json.dumps(error.get('details') or {})
        
        sheet.append_row([
            self._format_date(error['timestamp']),
            error.get('respondentId', ''),
            error['category'],
            error['message'],
            details_json,
            error.get('attempt', 1)
        ])
    
    def update_respondent_status(self, row_index: int, status: str):
        """Updates the status column for a respondent row.
        
        Status is at column 19 (0-indexed), which is column 20 in gspread (1-indexed).
        """
        self._require_config()
        sheet = self._get_sheet(self.config['respondentsSheet'])
        if not sheet:
            return
        
        status_column = 20  # Column 19 (0-indexed) = Column 20 (1-indexed in gspread)
        sheet.update_cell(row_index, status_column, status)
    
    def write_run_log(self, summary: Dict[str, Any]):
        """Logs a batch run summary into the persistent run log sheet."""
        if not self._spreadsheet:
            return
        
        sheet = self._get_sheet('RunLog')
        if not sheet:
            sheet = self._spreadsheet.add_worksheet(title='RunLog', rows=1000, cols=10)
        
        # Ensure headers exist
        headers = ['Timestamp', 'RunId', 'Processed', 'Errors', 'Duration']
        existing_values = sheet.get_all_values()
        if len(existing_values) == 0:
            # Sheet is empty, add headers
            sheet.append_row(headers)
        elif existing_values[0] != headers:
            # First row doesn't match headers, insert headers at the top
            sheet.insert_row(headers, index=1)
        
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
    
    def write_pm1raw_results(self, respondent: Dict[str, Any], pm01_raw_results: Dict[str, Dict[str, Any]]):
        """Writes PM01 Raw Scoring results to PM1Raw sheet (one row per question)."""
        if not self._spreadsheet or not pm01_raw_results:
            return
        
        sheet = self._get_sheet('PM1Raw')
        if not sheet:
            sheet = self._spreadsheet.add_worksheet(title='PM1Raw', rows=1000, cols=13)
        
        # Ensure headers exist
        headers = [
            'Respondent_ID', 'Timestamp', 'Question', 'Primary_Score', 'Sub_Score', 
            'Process_Score', 'AES_Clarity', 'AES_Logic', 'AES_Relevance', 
            'Evidence', 'Judgment_Reason'
        ]
        existing_values = sheet.get_all_values()
        if len(existing_values) == 0:
            # Sheet is empty, add headers
            sheet.append_row(headers)
        elif existing_values[0] != headers:
            # First row doesn't match headers, insert headers at the top
            # This will shift existing rows down
            sheet.insert_row(headers, index=1)
        
        timestamp = self._format_date(datetime.now())
        rows = []
        
        # Write one row per question (Q1-Q6)
        for question_id in sorted(pm01_raw_results.keys()):
            q_data = pm01_raw_results[question_id]
            rows.append([
                respondent['id'],
                timestamp,
                question_id,
                q_data.get('primary_score', 0),
                q_data.get('sub_score', 0),
                q_data.get('process_score', 0),
                q_data.get('aes_clarity', 0),
                q_data.get('aes_logic', 0),
                q_data.get('aes_relevance', 0),
                q_data.get('evidence', ''),
                q_data.get('judgment_reason', '')
            ])
        
        if rows:
            sheet.append_rows(rows)
    
    def write_pm5raw_results(self, respondent: Dict[str, Any], pm05_raw_results: Dict[str, Dict[str, Any]]):
        """Writes PM05 Raw Scoring results to PM5Raw sheet (one row per question)."""
        if not self._spreadsheet or not pm05_raw_results:
            return
        
        sheet = self._get_sheet('PM5Raw')
        if not sheet:
            sheet = self._spreadsheet.add_worksheet(title='PM5Raw', rows=1000, cols=10)
        
        # Ensure headers exist
        headers = [
            'Respondent_ID', 'Timestamp', 'Question', 'Primary_Score', 'Sub_Score', 
            'Process_Score', 'Difference_Note'
        ]
        existing_values = sheet.get_all_values()
        if len(existing_values) == 0:
            # Sheet is empty, add headers
            sheet.append_row(headers)
        elif existing_values[0] != headers:
            # First row doesn't match headers, insert headers at the top
            # This will shift existing rows down
            sheet.insert_row(headers, index=1)
        
        timestamp = self._format_date(datetime.now())
        rows = []
        
        # Write one row per question (Q1-Q6)
        for question_id in sorted(pm05_raw_results.keys()):
            q_data = pm05_raw_results[question_id]
            rows.append([
                respondent['id'],
                timestamp,
                question_id,
                q_data.get('primary_score', 0),
                q_data.get('sub_score', 0),
                q_data.get('process_score', 0),
                q_data.get('difference_note', '')
            ])
        
        if rows:
            sheet.append_rows(rows)
    
    def write_pm1final_results(self, respondent: Dict[str, Any], pm01_final: Dict[str, Any]):
        """Writes PM01 Final results to PM1Final sheet (one row per respondent)."""
        if not self._spreadsheet or not pm01_final:
            return
        
        sheet = self._get_sheet('PM1Final')
        if not sheet:
            sheet = self._spreadsheet.add_worksheet(title='PM1Final', rows=1000, cols=15)
        
        # Ensure headers exist (aggregated results only, no per-question details)
        headers = [
            'Respondent_ID', 'Company_Name', 'Timestamp', 'Total_Score', 
            'Scores_Primary_JSON', 'Scores_Sub_JSON', 'Process_JSON', 'AES_JSON',
            'Overall_Summary', 'AI_Use_Level', 'Recommendations_JSON'
        ]
        existing_values = sheet.get_all_values()
        if len(existing_values) == 0:
            # Sheet is empty, add headers
            sheet.append_row(headers)
        elif existing_values[0] != headers:
            # First row doesn't match headers, insert headers at the top
            # This will shift existing rows down
            sheet.insert_row(headers, index=1)
        
        timestamp = self._format_date(datetime.now())
        
        # Build row with aggregated results only
        row = [
            respondent['id'],
            respondent.get('company_name', ''),
            timestamp,
            pm01_final.get('total_score', 0),
            json.dumps(pm01_final.get('scores_primary', {}), ensure_ascii=False),
            json.dumps(pm01_final.get('scores_sub', {}), ensure_ascii=False),
            json.dumps(pm01_final.get('process', {}), ensure_ascii=False),
            json.dumps(pm01_final.get('aes', {}), ensure_ascii=False),
            pm01_final.get('overall_summary', ''),
            pm01_final.get('ai_use_level', ''),
            json.dumps(pm01_final.get('recommendations', []), ensure_ascii=False)
        ]
        
        sheet.append_row(row)
    
    def write_report_url(
        self,
        respondent_id: str,
        hash_id: str,
        filepath: str,
        report_url: str,
        timestamp: str,
        report_type: str = 'individual',
        company_name: Optional[str] = None,
        department: Optional[str] = None
    ):
        """
        Writes report URL mapping to reportIndSheet or reportOrgSheet (configured in Config sheet).
        Uses different column structures for individual vs organization reports.
        
        Args:
            respondent_id: Respondent ID (for individual) or company name (for organization)
            hash_id: Hash ID for the report
            filepath: File path of the report
            report_url: Full URL of the report
            timestamp: Timestamp string
            report_type: 'individual' or 'organization' (default: 'individual')
            company_name: Company name (for organization reports, optional)
            department: Department name (for organization reports, optional)
        """
        if not self._spreadsheet:
            return
        
        # Get sheet name from config based on report type
        if report_type == 'organization':
            sheet_name = self.config.get('reportOrgSheet', 'ReportOrganization') if self.config else 'ReportOrganization'
        else:
            sheet_name = self.config.get('reportIndSheet', 'ReportIndividual') if self.config else 'ReportIndividual'
        
        sheet = self._get_sheet(sheet_name)
        if not sheet:
            # Create sheet with appropriate number of columns
            num_cols = 7 if report_type == 'organization' else 6
            sheet = self._spreadsheet.add_worksheet(title=sheet_name, rows=1000, cols=num_cols)
        
        # Define headers based on report type
        if report_type == 'organization':
            headers = [
                'Hash_ID', 'Company_Name', 'Department', 'Report_URL', 'Filepath', 'Timestamp', 'Created_At'
            ]
        else:
            headers = [
                'Hash_ID', 'Respondent_ID', 'Report_URL', 'Filepath', 'Timestamp', 'Created_At'
            ]
        
        # Ensure headers exist
        existing_values = sheet.get_all_values()
        if len(existing_values) == 0:
            sheet.append_row(headers)
        elif existing_values[0] != headers:
            sheet.insert_row(headers, index=1)
        
        created_at = self._format_date(datetime.now())
        
        # Build row based on report type
        if report_type == 'organization':
            row = [
                hash_id,
                company_name or respondent_id,
                department or '',
                report_url,
                filepath,
                timestamp,
                created_at
            ]
        else:
            row = [
                hash_id,
                respondent_id,
                report_url,
                filepath,
                timestamp,
                created_at
            ]
        
        sheet.append_row(row)
    
    def write_pm5final_results(self, respondent: Dict[str, Any], pm05_final: Dict[str, Any]):
        """Writes PM05 Final results to PM5Final sheet (one row per respondent)."""
        if not self._spreadsheet or not pm05_final:
            return
        
        sheet = self._get_sheet('PM5Final')
        if not sheet:
            sheet = self._spreadsheet.add_worksheet(title='PM5Final', rows=1000, cols=10)
        
        # Ensure headers exist
        headers = [
            'Respondent_ID', 'Timestamp', 'Status', 'Consistency_Score', 
            'Detected_Issues_JSON', 'Comment'
        ]
        existing_values = sheet.get_all_values()
        if len(existing_values) == 0:
            # Sheet is empty, add headers
            sheet.append_row(headers)
        elif existing_values[0] != headers:
            # First row doesn't match headers, insert headers at the top
            sheet.insert_row(headers, index=1)
        
        sheet.append_row([
            respondent['id'],
            self._format_date(datetime.now()),
            pm05_final.get('status', ''),
            pm05_final.get('consistency_score', 0),
            json.dumps(pm05_final.get('detected_issues', []), ensure_ascii=False),
            pm05_final.get('comment', '')
        ])
