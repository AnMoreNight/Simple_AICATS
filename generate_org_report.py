#!/usr/bin/env python3
"""
Standalone script to generate organization reports from existing PM1Final data.
Usage: python generate_org_report.py [company_name] [department_name]
If no company_name is provided, lists all available companies.
"""

import sys
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

from services.sheets import SheetsService
from services.report import ReportService
from services.llm import LLMService
from core.config import Config


def list_companies():
    """List all companies with completed diagnoses."""
    sheets = SheetsService(config=None)
    config = Config.get_config(sheets_service=sheets)
    sheets.config = config
    
    pm1final_sheet = sheets._get_sheet('PM1Final')
    if not pm1final_sheet:
        print("Error: PM1Final sheet not found")
        return
    
    values = pm1final_sheet.get_all_values()
    if len(values) <= 1:
        print("No data found in PM1Final sheet")
        return
    
    companies = set()
    for row in values[1:]:
        if len(row) > 1 and row[1]:  # Company_Name column
            companies.add(row[1].strip())
    
    if companies:
        print("\nAvailable companies:")
        for company in sorted(companies):
            print(f"  - {company}")
    else:
        print("No companies found")


def generate_org_report(company_name: str, department: str = None):
    """Generate organization report for a specific company."""
    # Initialize services
    sheets = SheetsService(config=None)
    config = Config.get_config(sheets_service=sheets)
    sheets.config = config
    
    report_service = ReportService(output_dir="report", sheets_service=sheets)
    llm_service = LLMService(config)
    
    try:
        report_output = report_service.generate_organization_report(
            company_name=company_name,
            sheets_service=sheets,
            department_filter=department,
            llm_service=llm_service,
            config=config
        )
        print(f"\n✓ Organization report generated: {report_output['filepath']}")
        print(f"  ✓ Report URL: {report_output['url']}")
        print(f"  Open the file in your browser to view the report.")
    except ValueError as e:
        print(f"\n✗ Error: {e}")
    except Exception as e:
        print(f"\n✗ Error generating report: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python generate_org_report.py [company_name] [department_name]")
        print("\nTo list available companies, run without arguments:")
        list_companies()
    else:
        company_name = sys.argv[1]
        department = sys.argv[2] if len(sys.argv) > 2 else None
        
        if department:
            print(f"Generating organization report for {company_name} (Department: {department})...")
        else:
            print(f"Generating organization report for {company_name}...")
        
        generate_org_report(company_name, department)

