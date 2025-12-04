#!/usr/bin/env python3
"""
Main entry point for PM01 (Primary Diagnosis) and PM05 (Secondary Diagnosis).
Processes respondents one by one with Q1-Q6 answers.
"""

import sys
import traceback
from datetime import datetime
from typing import Dict, Any

from services.sheets import SheetsService
from services.validation import ValidationService
from services.llm import LLMService
from services.scoring_engine import ScoringEngine
from core.utils import create_run_id


def main():
    """Main execution function."""
    try:
        # Initialize services
        from core.config import Config
        from services.sheets import SheetsService
        
        # Create SheetsService (config not needed for initialization, only for sheet operations)
        sheets = SheetsService(config=None)
        
        # Read full config from Config sheet
        config = Config.get_config(sheets_service=sheets)
        sheets.config = config  # Set config for sheet operations
        
        validation = ValidationService(sheets)
        llm_service = LLMService(config)
        scoring_engine = ScoringEngine(config)
        
        # Read respondents
        print("Reading respondents from sheet...")
        all_rows = sheets.get_respondent_rows()
        print(f"Found {len(all_rows)} total respondents")
        
        # Filter pending respondents
        pending = [
            row for row in all_rows
            if row.get('status', '').lower() not in ['pm01完了', 'pm05完了', '診断完了']
        ]
        print(f"Found {len(pending)} pending respondents")
        
        if not pending:
            print("No respondents pending diagnosis.")
            return
        
        # Validate respondents
        print("Validating respondents...")
        validation_result = validation.validate_respondents(pending)
        valid = validation_result['valid']
        errors = validation_result['errors']
        
        print(f"Validation complete: {len(valid)} valid, {len(errors)} errors")
        
        if errors:
            sheets.log_validation_errors(errors)
        
        if not valid:
            print("No valid respondents to process.")
            return
        
        # Get already processed respondent IDs
        print("Checking for already processed respondents...")
        processed_ids = set()
        try:
            pm01_rows = sheets.read_pm01_rows()
            for row in pm01_rows:
                processed_ids.add(row['respondentId'])
            print(f"Found {len(processed_ids)} already processed respondents")
        except Exception as e:
            print(f"Warning: Could not read PM01 rows: {e}")
        
        # Filter out already processed
        unprocessed = [row for row in valid if row['id'] not in processed_ids]
        print(f"Found {len(unprocessed)} unprocessed respondents")
        
        if not unprocessed:
            print("All pending respondents have already been processed.")
            return
        
        # Read questions and metadata
        print("Reading questions and metadata...")
        questions = sheets.get_question_rows()
        
        # Initialize run tracking
        run_id = create_run_id()
        started_at = datetime.now()
        total_processed = 0
        total_errors = 0
        
        print(f"Starting diagnosis with run ID: {run_id}")
        print(f"Processing {len(unprocessed)} respondents one by one...")
        
        # Process each respondent
        for idx, respondent in enumerate(unprocessed, 1):
            try:
                print(f"\n[{idx}/{len(unprocessed)}] Processing: {respondent['id']} ({respondent['name']})")
                
                # Run PM01 (Primary Diagnosis)
                pm01_result = run_pm01(
                    respondent=respondent,
                    questions=questions,
                    config=config,
                    llm_service=llm_service,
                    scoring_engine=scoring_engine,
                    sheets=sheets
                )
                
                if not pm01_result:
                    total_errors += 1
                    print(f"✗ PM01 failed for {respondent['id']}")
                    continue
                
                # Run PM05 (Secondary Diagnosis)
                pm05_result = run_pm05(
                    respondent=respondent,
                    questions=questions,
                    pm01_result=pm01_result,
                    config=config,
                    llm_service=llm_service,
                    scoring_engine=scoring_engine,
                    sheets=sheets
                )
                
                if pm05_result:
                    total_processed += 1
                    print(f"✓ Successfully processed {respondent['id']} (PM01 + PM05)")
                else:
                    total_errors += 1
                    print(f"✗ PM05 failed for {respondent['id']}")
                    
            except Exception as e:
                total_errors += 1
                error_msg = f"Error processing {respondent['id']}: {str(e)}"
                print(f"✗ {error_msg}")
                traceback.print_exc()
                
                sheets.log_error({
                    'respondentId': respondent['id'],
                    'category': 'PROCESSING_ERROR',
                    'message': error_msg,
                    'attempt': 1,
                    'timestamp': datetime.now(),
                    'details': {'error': str(e)}
                })
        
        # Finalize run
        duration_ms = int((datetime.now() - started_at).total_seconds() * 1000)
        sheets.write_run_log({
            'runId': run_id,
            'processed': total_processed,
            'errors': total_errors,
            'timestamp': datetime.now(),
            'durationMs': duration_ms
        })
        
        print(f"\n{'='*60}")
        print(f"Diagnosis completed!")
        print(f"Run ID: {run_id}")
        print(f"Processed: {total_processed}")
        print(f"Errors: {total_errors}")
        print(f"Duration: {duration_ms}ms ({duration_ms//60000}分{(duration_ms%60000)//1000}秒)")
        print(f"{'='*60}")
        
    except Exception as e:
        print(f"Fatal error: {e}")
        traceback.print_exc()
        sys.exit(1)


def run_pm01(
    respondent: Dict[str, Any],
    questions: list,
    config: Dict[str, Any],
    llm_service: LLMService,
    scoring_engine: 'ScoringEngine',
    sheets: SheetsService
) -> Dict[str, Any] | None:
    """
    Run Primary Diagnosis (PM01).
    
    Returns PM01 result dict or None if failed.
    """
    max_retries = config.get('maxRetries', 3)
    attempt = 0
    
    while attempt < max_retries:
        attempt += 1
        print(f"  PM01 Attempt {attempt}/{max_retries}...")
        
        try:
            # Get LLM evaluation
            llm_response = llm_service.run_pm01_diagnosis(
                respondent=respondent,
                questions=questions,
                attempt=attempt
            )
            
            if not llm_response:
                continue
            
            # Apply scoring engine
            pm01_result = scoring_engine.calculate_pm01_scores(
                llm_response=llm_response,
                questions=questions
            )
            
            if pm01_result:
                # Write PM01 result to sheet
                sheets.append_pm01_result(respondent, pm01_result)
                print(f"  ✓ PM01 completed")
                return pm01_result
                
        except Exception as e:
            print(f"  Error in PM01 attempt {attempt}: {e}")
            if attempt >= max_retries:
                sheets.log_error({
                    'respondentId': respondent['id'],
                    'category': 'PM01_FAILED',
                    'message': f'PM01 failed after {max_retries} attempts',
                    'attempt': max_retries,
                    'timestamp': datetime.now(),
                    'details': {'error': str(e)}
                })
                return None
    
    return None


def run_pm05(
    respondent: Dict[str, Any],
    questions: list,
    pm01_result: Dict[str, Any],
    config: Dict[str, Any],
    llm_service: LLMService,
    scoring_engine: 'ScoringEngine',
    sheets: SheetsService
) -> Dict[str, Any] | None:
    """
    Run Secondary Diagnosis (PM05) - Reverse scoring validation.
    
    Returns PM05 result dict or None if failed.
    """
    max_retries = config.get('maxRetries', 3)
    attempt = 0
    
    while attempt < max_retries:
        attempt += 1
        print(f"  PM05 Attempt {attempt}/{max_retries}...")
        
        try:
            # Get reverse-scored LLM evaluation
            llm_response = llm_service.run_pm05_validation(
                respondent=respondent,
                questions=questions,
                pm01_result=pm01_result,
                attempt=attempt
            )
            
            if not llm_response:
                continue
            
            # Calculate consistency and validation
            pm05_result = scoring_engine.calculate_pm05_validation(
                pm01_result=pm01_result,
                pm05_llm_response=llm_response,
                questions=questions
            )
            
            if pm05_result:
                # Write PM05 result to sheet
                sheets.append_pm05_result(respondent, pm05_result)
                print(f"  ✓ PM05 completed")
                return pm05_result
                
        except Exception as e:
            print(f"  Error in PM05 attempt {attempt}: {e}")
            if attempt >= max_retries:
                sheets.log_error({
                    'respondentId': respondent['id'],
                    'category': 'PM05_FAILED',
                    'message': f'PM05 failed after {max_retries} attempts',
                    'attempt': max_retries,
                    'timestamp': datetime.now(),
                    'details': {'error': str(e)}
                })
                return None
    
    return None


if __name__ == '__main__':
    main()
