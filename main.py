#!/usr/bin/env python3
"""
Main entry point for PM01 (Primary Diagnosis) and PM05 (Secondary Diagnosis).
Processes respondents one by one with Q1-Q6 answers.
"""

import sys
import traceback
from datetime import datetime
from typing import Dict, Any
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

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
        print("Questions : ", questions)
        
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
                
                # STEP 1: PM01 Raw Scoring (Individual Q-A scoring)
                print("  STEP 1: PM01 Raw Scoring...")
                pm01_raw_results = run_pm01_raw(
                    respondent=respondent,
                    questions=questions,
                    config=config,
                    llm_service=llm_service,
                    sheets=sheets
                )
                
                if not pm01_raw_results:
                    total_errors += 1
                    print(f"✗ STEP 1 failed for {respondent['id']}")
                    continue
                
                # STEP 2: PM05 Raw Scoring (Reverse logic)
                print("  STEP 2: PM05 Raw Scoring...")
                pm05_raw_results = run_pm05_raw(
                    respondent=respondent,
                    questions=questions,
                    pm01_raw_results=pm01_raw_results,
                    config=config,
                    llm_service=llm_service,
                    sheets=sheets
                )
                
                if not pm05_raw_results:
                    total_errors += 1
                    print(f"✗ STEP 2 failed for {respondent['id']}")
                    continue
                
                # STEP 3: PM01 Final (Aggregate PM01 raw scores + LLM analysis)
                print("  STEP 3: PM01 Final (Aggregation + Analysis)...")
                
                # First, aggregate scores
                aggregated_scores = scoring_engine.aggregate_pm01_raw_scores(
                    pm01_raw_results=pm01_raw_results,
                    questions=questions
                )
                
                if not aggregated_scores:
                    total_errors += 1
                    print(f"✗ STEP 3 aggregation failed for {respondent['id']}")
                    continue
                
                # Then, get LLM analysis
                max_retries = config.get('maxRetries')
                attempt = 0
                pm01_final_analysis = None
                
                while attempt < max_retries:
                    attempt += 1
                    print(f"    PM01 Final Analysis Attempt {attempt}/{max_retries}...")
                    
                    try:
                        pm01_final_analysis = llm_service.run_pm01_final_analysis(
                            respondent=respondent,
                            pm01_raw_results=pm01_raw_results,
                            aggregated_scores=aggregated_scores,
                            attempt=attempt
                        )
                        
                        if pm01_final_analysis:
                            print(f"    ✓ PM01 Final analysis completed")
                            break
                            
                    except Exception as e:
                        print(f"    Error in PM01 Final analysis attempt {attempt}: {e}")
                        if attempt >= max_retries:
                            sheets.log_error({
                                'respondentId': respondent['id'],
                                'category': 'PM01_FINAL_FAILED',
                                'message': f'PM01 Final analysis failed after {max_retries} attempts',
                                'attempt': max_retries,
                                'timestamp': datetime.now(),
                                'details': {'error': str(e)}
                            })
                
                if not pm01_final_analysis:
                    total_errors += 1
                    print(f"✗ STEP 3 analysis failed for {respondent['id']}")
                    continue
                
                # Combine aggregated scores with LLM analysis
                pm01_final = scoring_engine.combine_pm01_final(
                    aggregated_scores=aggregated_scores,
                    pm01_final_analysis=pm01_final_analysis,
                    pm01_raw_results=pm01_raw_results
                )
                
                if not pm01_final:
                    total_errors += 1
                    print(f"✗ STEP 3 combination failed for {respondent['id']}")
                    continue
                
                # STEP 4: PM05 Final (Consistency check)
                print("  STEP 4: PM05 Final (Consistency Check)...")
                pm05_final = run_pm05_final(
                    respondent=respondent,
                    pm01_final=pm01_final,
                    config=config,
                    llm_service=llm_service,
                    scoring_engine=scoring_engine,
                    sheets=sheets
                )
                
                if pm05_final:
                    # Write final results
                    sheets.append_pm01_result(respondent, pm01_final)
                    sheets.append_pm05_result(respondent, pm05_final)
                    total_processed += 1
                    print(f"✓ Successfully processed {respondent['id']} (All steps completed)")
                else:
                    total_errors += 1
                    print(f"✗ STEP 4 failed for {respondent['id']}")
                    
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


def run_pm01_raw(
    respondent: Dict[str, Any],
    questions: list,
    config: Dict[str, Any],
    llm_service: LLMService,
    sheets: SheetsService
) -> Dict[str, Dict[str, Any]] | None:
    """
    STEP 1: PM01 Raw Scoring - Individual Q-A scoring.
    
    Returns dict mapping Q1-Q6 to their raw scoring results.
    """
    max_retries = config.get('maxRetries')
    all_question_results = {}
    
    # Process each question individually
    for i, question in enumerate(questions):
        if question['number'] < 1 or question['number'] > 6:
            continue
        
        question_id = f"Q{question['number']}"
        print(f"    Processing {question_id}...")
        
        attempt = 0
        question_result = None
        
        while attempt < max_retries:
            attempt += 1
            print(f"      {question_id} Attempt {attempt}/{max_retries}...")
            
            try:
                # Get LLM evaluation for this single question
                llm_response = llm_service.run_pm01_raw_scoring(
                    respondent=respondent,
                    question=question,
                    question_index=i,
                    attempt=attempt
                )
                
                if llm_response:
                    question_result = llm_response
                    print(f"      ✓ {question_id} raw scoring completed")
                    break
                    
            except Exception as e:
                print(f"      Error in {question_id} attempt {attempt}: {e}")
                if attempt >= max_retries:
                    sheets.log_error({
                        'respondentId': respondent['id'],
                        'category': 'PM01_RAW_FAILED',
                        'message': f'{question_id} raw scoring failed after {max_retries} attempts',
                        'attempt': max_retries,
                        'timestamp': datetime.now(),
                        'details': {'error': str(e), 'question': question_id}
                    })
        
        if not question_result:
            print(f"    ✗ {question_id} raw scoring failed")
            return None
        
        all_question_results[question_id] = question_result
    
    print(f"  ✓ STEP 1 completed (all Q-A raw scoring)")
    return all_question_results


def run_pm05_raw(
    respondent: Dict[str, Any],
    questions: list,
    pm01_raw_results: Dict[str, Dict[str, Any]],
    config: Dict[str, Any],
    llm_service: LLMService,
    sheets: SheetsService
) -> Dict[str, Dict[str, Any]] | None:
    """
    STEP 2: PM05 Raw Scoring - Reverse logic scoring using PM01 raw as reference.
    
    Returns dict mapping Q1-Q6 to their reverse-scored results.
    """
    max_retries = config.get('maxRetries')
    all_question_results = {}
    
    # Process each question individually
    for i, question in enumerate(questions):
        if question['number'] < 1 or question['number'] > 6:
            continue
        
        question_id = f"Q{question['number']}"
        pm01_raw = pm01_raw_results.get(question_id)
        
        if not pm01_raw:
            print(f"    ✗ {question_id} PM01 raw result not found")
            return None
        
        print(f"    Processing {question_id} reverse scoring...")
        
        attempt = 0
        question_result = None
        
        while attempt < max_retries:
            attempt += 1
            print(f"      {question_id} Attempt {attempt}/{max_retries}...")
            
            try:
                # Get reverse logic evaluation for this single question
                llm_response = llm_service.run_pm05_raw_scoring(
                    respondent=respondent,
                    question=question,
                    question_index=i,
                    pm01_raw_result=pm01_raw,
                    attempt=attempt
                )
                
                if llm_response:
                    question_result = llm_response
                    print(f"      ✓ {question_id} reverse scoring completed")
                    break
                    
            except Exception as e:
                print(f"      Error in {question_id} attempt {attempt}: {e}")
                if attempt >= max_retries:
                    sheets.log_error({
                        'respondentId': respondent['id'],
                        'category': 'PM05_RAW_FAILED',
                        'message': f'{question_id} reverse scoring failed after {max_retries} attempts',
                        'attempt': max_retries,
                        'timestamp': datetime.now(),
                        'details': {'error': str(e), 'question': question_id}
                    })
        
        if not question_result:
            print(f"    ✗ {question_id} reverse scoring failed")
            return None
        
        all_question_results[question_id] = question_result
    
    print(f"  ✓ STEP 2 completed (all Q-A reverse scoring)")
    return all_question_results


def run_pm05_final(
    respondent: Dict[str, Any],
    pm01_final: Dict[str, Any],
    config: Dict[str, Any],
    llm_service: LLMService,
    scoring_engine: 'ScoringEngine',
    sheets: SheetsService
) -> Dict[str, Any] | None:
    """
    STEP 4: PM05 Final - Overall consistency check using PM01 Final.
    
    Returns PM05 final result dict or None if failed.
    """
    max_retries = config.get('maxRetries')
    attempt = 0
    
    while attempt < max_retries:
        attempt += 1
        print(f"    PM05 Final Attempt {attempt}/{max_retries}...")
        
        try:
            # Get consistency check evaluation
            llm_response = llm_service.run_pm05_final_check(
                respondent=respondent,
                pm01_final=pm01_final,
                attempt=attempt
            )
            
            if not llm_response:
                continue
            
            # Process consistency check result
            pm05_final_result = scoring_engine.process_pm05_final(
                pm05_llm_response=llm_response,
                pm01_final=pm01_final
            )
            
            if pm05_final_result:
                print(f"  ✓ STEP 4 completed (consistency check)")
                return pm05_final_result
                
        except Exception as e:
            print(f"    Error in PM05 Final attempt {attempt}: {e}")
            if attempt >= max_retries:
                sheets.log_error({
                    'respondentId': respondent['id'],
                    'category': 'PM05_FINAL_FAILED',
                    'message': f'PM05 Final failed after {max_retries} attempts',
                    'attempt': max_retries,
                    'timestamp': datetime.now(),
                    'details': {'error': str(e)}
                })
                return None
    
    return None




if __name__ == '__main__':
    main()
