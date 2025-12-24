#!/usr/bin/env python3
"""
Flask API server for AI-CATS diagnosis and report generation.
Deployed on xserver with WSGI.
"""

import os
import sys
import json
import traceback
import threading
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, Optional
from flask import Flask, request, jsonify, Response, stream_with_context
from flask_cors import CORS
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Import services
from services.sheets import SheetsService
from services.validation import ValidationService
from services.llm import LLMService
from services.scoring_engine import ScoringEngine
from services.report import ReportService
from core.config import Config
from core.utils import create_run_id

# Import main processing functions
from main import (
    run_pm01_raw,
    run_pm05_raw,
    run_pm05_final
)

app = Flask(__name__)
CORS(app)  # Enable CORS for frontend

# Global state for diagnosis status
diagnosis_status = {
    'running': False,
    'total': 0,
    'pending': 0,
    'completed': 0
}


def initialize_services():
    """Initialize all services with config."""
    sheets = SheetsService(config=None)
    config = Config.get_config(sheets_service=sheets)
    sheets.config = config
    
    validation = ValidationService(sheets)
    llm_service = LLMService(config)
    scoring_engine = ScoringEngine(config)
    report_service = ReportService(output_dir="report", sheets_service=sheets)
    
    return {
        'sheets': sheets,
        'config': config,
        'validation': validation,
        'llm_service': llm_service,
        'scoring_engine': scoring_engine,
        'report_service': report_service
    }


@app.route('/api/diagnosis/start', methods=['POST'])
def start_diagnosis():
    """Start diagnosis process with streaming response."""
    
    def generate():
        """Generator function for streaming diagnosis progress."""
        try:
            yield f"data: {json.dumps({'type': 'log', 'message': '診断を開始します...', 'level': 'info'})}\n\n"
            
            services = initialize_services()
            sheets = services['sheets']
            config = services['config']
            validation = services['validation']
            llm_service = services['llm_service']
            scoring_engine = services['scoring_engine']
            
            # Read respondents
            yield f"data: {json.dumps({'type': 'log', 'message': '回答者を読み込んでいます...', 'level': 'info'})}\n\n"
            all_rows = sheets.get_respondent_rows()
            
            # Filter pending respondents
            completed_statuses = ['pm5final完了', 'pm5final完成', '診断完了', '診断完成']
            pending = [
                row for row in all_rows
                if row.get('status', '').strip().lower() not in completed_statuses
            ]
            
            yield f"data: {json.dumps({'type': 'status', 'status': '実行中', 'total': len(all_rows), 'pending': len(pending), 'count': 0})}\n\n"
            
            if not pending:
                yield f"data: {json.dumps({'type': 'log', 'message': '処理待ちの回答者がありません。', 'level': 'warning'})}\n\n"
                return
            
            # Validate respondents
            yield f"data: {json.dumps({'type': 'log', 'message': '回答者を検証しています...', 'level': 'info'})}\n\n"
            validation_result = validation.validate_respondents(pending)
            valid = validation_result['valid']
            errors = validation_result['errors']
            
            if errors:
                sheets.log_validation_errors(errors)
            
            if not valid:
                yield f"data: {json.dumps({'type': 'log', 'message': '有効な回答者がありません。', 'level': 'error'})}\n\n"
                return
            
            # Read questions
            questions = sheets.get_question_rows()
            
            # Process each respondent
            run_id = create_run_id()
            total_processed = 0
            total_errors = 0
            
            for idx, respondent in enumerate(valid, 1):
                try:
                    resp_id = respondent['id']
                    resp_name = respondent.get('name', 'N/A')
                    message = f"{resp_id} ({resp_name}) を処理中..."
                    yield f"data: {json.dumps({'type': 'progress', 'current': idx, 'total': len(valid), 'message': message})}\n\n"
                    
                    # Determine starting step
                    current_status = respondent.get('status', '').strip().lower()
                    start_from_step = 1
                    
                    if current_status in ['pm1raw完了', 'pm1raw完成']:
                        start_from_step = 2
                    elif current_status in ['pm5raw完了', 'pm5raw完成']:
                        start_from_step = 3
                    elif current_status in ['pm1final完了', 'pm1final完成']:
                        start_from_step = 4
                    
                    # STEP 1: PM01 Raw
                    pm01_raw_results = None
                    if start_from_step <= 1:
                        yield f"data: {json.dumps({'type': 'log', 'message': '  STEP 1: PM01 Raw処理中...', 'level': 'info'})}\n\n"
                        pm01_raw_results = run_pm01_raw(
                            respondent=respondent,
                            questions=questions,
                            config=config,
                            llm_service=llm_service,
                            sheets=sheets
                        )
                        if pm01_raw_results:
                            sheets.write_pm1raw_results(respondent, pm01_raw_results)
                            sheets.update_respondent_status(respondent['rowIndex'], 'PM1Raw完了')
                        else:
                            total_errors += 1
                            continue
                    else:
                        # Load from sheet
                        pm1raw_sheet = sheets._get_sheet('PM1Raw')
                        if pm1raw_sheet:
                            values = pm1raw_sheet.get_all_values()
                            pm01_raw_results = {}
                            for row in values[1:]:
                                if len(row) > 2 and row[0] == respondent['id']:
                                    q_id = row[2]
                                    pm01_raw_results[q_id] = {
                                        'primary_score': float(row[3]) if len(row) > 3 and row[3] else 0,
                                        'sub_score': float(row[4]) if len(row) > 4 and row[4] else 0,
                                        'process_score': float(row[5]) if len(row) > 5 and row[5] else 0,
                                        'aes_clarity': float(row[6]) if len(row) > 6 and row[6] else 0,
                                        'aes_logic': float(row[7]) if len(row) > 7 and row[7] else 0,
                                        'aes_relevance': float(row[8]) if len(row) > 8 and row[8] else 0,
                                        'evidence': row[9] if len(row) > 9 else '',
                                        'judgment_reason': row[10] if len(row) > 10 else ''
                                    }
                    
                    # STEP 2: PM05 Raw
                    pm05_raw_results = None
                    if start_from_step <= 2:
                        yield f"data: {json.dumps({'type': 'log', 'message': '  STEP 2: PM05 Raw処理中...', 'level': 'info'})}\n\n"
                        pm05_raw_results = run_pm05_raw(
                            respondent=respondent,
                            questions=questions,
                            pm01_raw_results=pm01_raw_results,
                            config=config,
                            llm_service=llm_service,
                            sheets=sheets
                        )
                        if pm05_raw_results:
                            sheets.write_pm5raw_results(respondent, pm05_raw_results)
                            sheets.update_respondent_status(respondent['rowIndex'], 'PM5Raw完了')
                        else:
                            total_errors += 1
                            continue
                    else:
                        # Load from sheet
                        pm5raw_sheet = sheets._get_sheet('PM5Raw')
                        if pm5raw_sheet:
                            values = pm5raw_sheet.get_all_values()
                            pm05_raw_results = {}
                            for row in values[1:]:
                                if len(row) > 2 and row[0] == respondent['id']:
                                    q_id = row[2]
                                    pm05_raw_results[q_id] = {
                                        'primary_score': float(row[3]) if len(row) > 3 and row[3] else 0,
                                        'sub_score': float(row[4]) if len(row) > 4 and row[4] else 0,
                                        'process_score': float(row[5]) if len(row) > 5 and row[5] else 0,
                                        'difference_note': row[6] if len(row) > 6 else ''
                                    }
                    
                    # STEP 3: PM01 Final
                    if start_from_step <= 3:
                        yield f"data: {json.dumps({'type': 'log', 'message': '  STEP 3: PM01 Final処理中...', 'level': 'info'})}\n\n"
                        aggregated_scores = scoring_engine.aggregate_pm05_raw_scores(
                            pm05_raw_results=pm05_raw_results,
                            pm01_raw_results=pm01_raw_results,
                            questions=questions
                        )
                        
                        if aggregated_scores:
                            max_retries = config.get('maxRetries')
                            attempt = 0
                            pm01_final_analysis = None
                            
                            while attempt < max_retries:
                                attempt += 1
                                try:
                                    pm01_final_analysis = llm_service.run_pm01_final_analysis(
                                        respondent=respondent,
                                        pm05_raw_results=pm05_raw_results,
                                        aggregated_scores=aggregated_scores,
                                        attempt=attempt
                                    )
                                    if pm01_final_analysis:
                                        break
                                except Exception as e:
                                    if attempt >= max_retries:
                                        sheets.log_error({
                                            'respondentId': respondent['id'],
                                            'category': 'PM01_FINAL_FAILED',
                                            'message': f'PM01 Final analysis failed after {max_retries} attempts',
                                            'attempt': max_retries,
                                            'timestamp': datetime.now(),
                                            'details': {'error': str(e)}
                                        })
                            
                            if pm01_final_analysis:
                                pm01_final = scoring_engine.process_pm01_final(
                                    pm01_llm_response=pm01_final_analysis,
                                    aggregated_scores=aggregated_scores,
                                    respondent=respondent
                                )
                                
                                if pm01_final:
                                    sheets.write_pm1final_results(respondent, pm01_final)
                                    sheets.update_respondent_status(respondent['rowIndex'], 'PM1Final完了')
                                else:
                                    total_errors += 1
                                    continue
                            else:
                                total_errors += 1
                                continue
                        else:
                            total_errors += 1
                            continue
                    
                    # STEP 4: PM05 Final
                    if start_from_step <= 4:
                        yield f"data: {json.dumps({'type': 'log', 'message': '  STEP 4: PM05 Final処理中...', 'level': 'info'})}\n\n"
                        # Load PM01 Final
                        pm1final_sheet = sheets._get_sheet('PM1Final')
                        pm01_final = None
                        if pm1final_sheet:
                            values = pm1final_sheet.get_all_values()
                            for row in values[1:]:
                                if len(row) > 0 and row[0] == respondent['id']:
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
                        
                        if pm01_final:
                            pm05_final_result = run_pm05_final(
                                respondent=respondent,
                                pm01_final=pm01_final,
                                config=config,
                                llm_service=llm_service,
                                scoring_engine=scoring_engine,
                                sheets=sheets
                            )
                            
                            if pm05_final_result:
                                sheets.write_pm5final_results(respondent, pm05_final_result)
                                sheets.update_respondent_status(respondent['rowIndex'], 'PM5Final完了')
                                total_processed += 1
                                resp_id = respondent['id']
                                yield f"data: {json.dumps({'type': 'log', 'message': f'  ✓ {resp_id} 完了', 'level': 'success'})}\n\n"
                            else:
                                total_errors += 1
                        else:
                            total_errors += 1
                    else:
                        total_processed += 1
                        resp_id = respondent['id']
                        yield f"data: {json.dumps({'type': 'log', 'message': f'  ✓ {resp_id} 完了', 'level': 'success'})}\n\n"
                    
                except Exception as e:
                    total_errors += 1
                    error_msg = f"Error processing {respondent['id']}: {str(e)}"
                    yield f"data: {json.dumps({'type': 'log', 'message': error_msg, 'level': 'error'})}\n\n"
                    sheets.log_error({
                        'respondentId': respondent['id'],
                        'category': 'PROCESSING_ERROR',
                        'message': error_msg,
                        'attempt': 1,
                        'timestamp': datetime.now(),
                        'details': {'error': str(e)}
                    })
            
            yield f"data: {json.dumps({'type': 'log', 'message': f'診断が完了しました。処理: {total_processed}, エラー: {total_errors}', 'level': 'success'})}\n\n"
            yield f"data: {json.dumps({'type': 'status', 'status': '完了', 'count': total_processed})}\n\n"
            
        except Exception as e:
            yield f"data: {json.dumps({'type': 'log', 'message': f'エラー: {str(e)}', 'level': 'error'})}\n\n"
    
    return Response(stream_with_context(generate()), mimetype='text/event-stream')


@app.route('/api/diagnosis/status', methods=['GET'])
def get_diagnosis_status():
    """Get current diagnosis status."""
    try:
        services = initialize_services()
        sheets = services['sheets']
        
        all_rows = sheets.get_respondent_rows()
        completed_statuses = ['pm5final完了', 'pm5final完成', '診断完了', '診断完成']
        pending = [
            row for row in all_rows
            if row.get('status', '').strip().lower() not in completed_statuses
        ]
        completed = len(all_rows) - len(pending)
        
        return jsonify({
            'total': len(all_rows),
            'pending': len(pending),
            'completed': completed,
            'running': diagnosis_status['running']
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/report/generate', methods=['POST'])
def generate_report():
    """Generate individual or organization report."""
    try:
        data = request.json
        report_type = data.get('report_type', 'individual')
        respondent_id = data.get('respondent_id')
        company_name = data.get('company_name')
        department = data.get('department')
        
        services = initialize_services()
        sheets = services['sheets']
        config = services['config']
        llm_service = services['llm_service']
        report_service = services['report_service']
        
        if report_type == 'individual':
            if respondent_id:
                respondents = sheets.get_respondent_rows()
                respondent = next((r for r in respondents if r['id'] == respondent_id), None)
                if not respondent:
                    return jsonify({'error': f'Respondent ID {respondent_id} not found'}), 404
                
                # Read PM1Final and PM5Final data
                pm1final_sheet = sheets._get_sheet('PM1Final')
                if not pm1final_sheet:
                    return jsonify({'error': 'PM1Final sheet not found'}), 404
                
                pm01_final = None
                values = pm1final_sheet.get_all_values()
                for row in values[1:]:
                    if len(row) > 0 and row[0] == respondent_id:
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
                    return jsonify({'error': f'PM1Final data not found for {respondent_id}'}), 404
                
                # Read PM5Final
                pm5final_sheet = sheets._get_sheet('PM5Final')
                pm05_final = None
                if pm5final_sheet:
                    values = pm5final_sheet.get_all_values()
                    for row in values[1:]:
                        if len(row) > 0 and row[0] == respondent_id:
                            pm05_final = {
                                'status': row[2] if len(row) > 2 else '妥当',
                                'consistency_score': float(row[3]) if len(row) > 3 and row[3] else 0,
                                'detected_issues': json.loads(row[4]) if len(row) > 4 and row[4] else [],
                                'comment': row[5] if len(row) > 5 else ''
                            }
                            break
                
                # Generate report
                result = report_service.generate_individual_report(
                    respondent=respondent,
                    pm01_final=pm01_final,
                    pm05_final=pm05_final,
                    llm_service=llm_service,
                    config=config
                )
                
                if result:
                    hash_id = Path(result['filepath']).stem
                    return jsonify({
                        'success': True,
                        'filepath': result['filepath'],
                        'url': f"/report/{hash_id}.html"
                    })
                else:
                    return jsonify({'error': 'Failed to generate individual report'}), 500
            else:
                # Generate for all completed
                respondents = sheets.get_respondent_rows()
                completed_count = 0
                for respondent in respondents:
                    status = respondent.get('status', '').strip().lower()
                    if 'pm5final完了' in status or 'pm5final完成' in status:
                        # Similar logic as above for each respondent
                        # (simplified - you may want to extract this to a helper function)
                        completed_count += 1
                
                return jsonify({
                    'success': True,
                    'message': f'Reports will be generated for {completed_count} completed respondents'
                })
        
        elif report_type == 'organization':
            if not company_name:
                return jsonify({'error': 'Company name is required for organization reports'}), 400
            
            # Generate organization report
            result = report_service.generate_organization_report(
                company_name=company_name,
                department=department,
                llm_service=llm_service,
                config=config
            )
            
            if result:
                hash_id = Path(result['filepath']).stem
                return jsonify({
                    'success': True,
                    'filepath': result['filepath'],
                    'url': f"/report/{hash_id}.html"
                })
            else:
                return jsonify({'error': 'Failed to generate organization report'}), 500
        
        return jsonify({'error': 'Invalid report type'}), 400
        
    except Exception as e:
        return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500


@app.route('/api/reports/list', methods=['GET'])
def list_reports():
    """List all generated reports."""
    try:
        report_dir = Path('report')
        if not report_dir.exists():
            return jsonify({'reports': []})
        
        reports = []
        for file in report_dir.glob('*.html'):
            hash_id = file.stem
            reports.append({
                'id': hash_id,
                'filepath': str(file),
                'url': f"/report/{hash_id}.html"
            })
        
        return jsonify({'reports': reports})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/health', methods=['GET'])
def health_check():
    """Health check endpoint."""
    return jsonify({'status': 'ok', 'timestamp': datetime.now().isoformat()})


# Serve static report files
@app.route('/report/<path:filename>')
def serve_report(filename):
    """Serve generated report HTML files."""
    from flask import send_from_directory
    report_dir = Path('report')
    if report_dir.exists() and (report_dir / filename).exists():
        return send_from_directory(str(report_dir), filename)
    return jsonify({'error': 'Report not found'}), 404


# Serve frontend static files
@app.route('/')
def serve_frontend():
    """Serve frontend index.html."""
    from flask import send_from_directory
    return send_from_directory('frontend', 'index.html')


@app.route('/<path:filename>')
def serve_frontend_static(filename):
    """Serve frontend static files (CSS, JS)."""
    from flask import send_from_directory
    return send_from_directory('frontend', filename)


if __name__ == '__main__':
    # For local development
    # Use 127.0.0.1 on Windows to avoid socket errors
    import sys
    if sys.platform == 'win32':
        # On Windows, disable reloader to avoid socket errors
        app.run(host='127.0.0.1', port=5000, debug=True, use_reloader=False)
    else:
        app.run(host='0.0.0.0', port=5000, debug=True)
else:
    # For WSGI (production)
    application = app

