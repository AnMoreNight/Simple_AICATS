#!/usr/bin/env python3
"""
Standalone script to generate individual reports from existing PM1Final and PM5Final data.
Usage: python generate_report.py [respondent_id]
If no respondent_id is provided, generates reports for all completed respondents.
"""

import sys
import json
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

from services.sheets import SheetsService
from services.report import ReportService
from services.llm import LLMService
from core.config import Config


def generate_report_for_respondent(respondent_id: str = None):
    """Generate report for a specific respondent or all completed respondents."""
    # Initialize services
    sheets = SheetsService(config=None)
    config = Config.get_config(sheets_service=sheets)
    sheets.config = config
    
    llm_service = LLMService(config)
    report_service = ReportService(output_dir="report", sheets_service=sheets)
    
    # Get all respondents
    respondents = sheets.get_respondent_rows()
    
    if respondent_id:
        # Generate report for specific respondent
        respondent = next((r for r in respondents if r['id'] == respondent_id), None)
        if not respondent:
            print(f"Error: Respondent ID '{respondent_id}' not found")
            return
        
        generate_single_report(sheets, report_service, respondent, llm_service, config)
    else:
        # Generate reports for all completed respondents
        print(f"Generating reports for all completed respondents...")
        completed_count = 0
        
        for respondent in respondents:
            status = respondent.get('status', '').strip().lower()
            if 'pm5final完了' in status or 'pm5final完成' in status:
                if generate_single_report(sheets, report_service, respondent, llm_service, config):
                    completed_count += 1
        
        print(f"\n✓ Generated {completed_count} reports")


def generate_single_report(sheets: SheetsService, report_service: ReportService, respondent: dict, llm_service: LLMService, config: dict) -> bool:
    """Generate report for a single respondent."""
    try:
        respondent_id = respondent['id']
        print(f"\nProcessing respondent: {respondent_id} ({respondent.get('name', 'N/A')})")
        
        # Read PM1Final data
        pm1final_sheet = sheets._get_sheet('PM1Final')
        if not pm1final_sheet:
            print(f"  ✗ PM1Final sheet not found")
            return False
        
        pm01_final = None
        values = pm1final_sheet.get_all_values()
        for row in values[1:]:
            if len(row) > 0 and row[0] == respondent_id:
                # Column order: Respondent_ID, Company_Name, Timestamp, Total_Score, 
                # Scores_Primary_JSON, Scores_Sub_JSON, Process_JSON, AES_JSON,
                # Overall_Summary, AI_Use_Level, Recommendations_JSON
                pm01_final = {
                    'total_score': float(row[3]) if len(row) > 3 and row[3] else 0,
                    'scores_primary': json.loads(row[4]) if len(row) > 4 and row[4] else {},
                    'scores_sub': json.loads(row[5]) if len(row) > 5 and row[5] else {},
                    'process': json.loads(row[6]) if len(row) > 6 and row[6] else {},
                    'aes': json.loads(row[7]) if len(row) > 7 and row[7] else {},
                    'overall_summary': row[8] if len(row) > 8 else '',
                    'ai_use_level': row[9] if len(row) > 9 else '',
                    'recommendations': json.loads(row[10]) if len(row) > 10 and row[10] else []
                }
                break
        
        if not pm01_final:
            print(f"  ✗ PM1Final data not found for {respondent_id}")
            return False
        
        # Read PM5Final data
        pm5final_sheet = sheets._get_sheet('PM5Final')
        pm05_final = None
        if pm5final_sheet:
            values = pm5final_sheet.get_all_values()
            for row in values[1:]:
                if len(row) > 0 and row[0] == respondent_id:
                    # Column order: Respondent_ID, Timestamp, Status, Consistency_Score, 
                    # Detected_Issues_JSON, Comment
                    pm05_final = {
                        'status': row[2] if len(row) > 2 else '妥当',
                        'consistency_score': float(row[3]) if len(row) > 3 and row[3] else 0,
                        'detected_issues': json.loads(row[4]) if len(row) > 4 and row[4] else [],
                        'comment': row[5] if len(row) > 5 else ''
                    }
                    break
        
        # Generate report
        report_result = report_service.generate_individual_report(
            respondent=respondent,
            pm01_final=pm01_final,
            pm05_final=pm05_final,
            llm_service=llm_service,
            config=config
        )
        
        print(f"  ✓ Report generated: {report_result['filepath']}")
        print(f"  ✓ Report URL: {report_result['url']}")
        return True
        
    except Exception as e:
        print(f"  ✗ Error generating report: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    respondent_id = sys.argv[1] if len(sys.argv) > 1 else None
    generate_report_for_respondent(respondent_id)

