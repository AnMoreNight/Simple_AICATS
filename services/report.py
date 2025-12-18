"""
Report generation service for individual and organization reports.
"""

import json
import statistics
import hashlib
import secrets
from datetime import datetime
from typing import Dict, Any, Optional, List
from pathlib import Path
from collections import defaultdict
from string import Template


class ReportService:
    """Service for generating HTML reports from diagnosis results."""
    
    REPORT_BASE_URL = "https://ai-cats.app/report"
    
    # Mapping from English PROCESS keys to Japanese labels
    PROCESS_LABELS_JP = {
        'clarity': '明瞭性',
        'structure': '構造性',
        'hypothesis': '仮説性',
        'prompt clarity': 'プロンプト明瞭性',
        'consistency': '一貫性'
    }
    
    def __init__(self, output_dir: str = "report", template_dir: str = "templates", sheets_service=None):
        """Initialize report service with output directory and template directory."""
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)
        self.template_dir = Path(template_dir)
        self.sheets_service = sheets_service
    
    def _generate_hash_id(self, respondent_id: str, timestamp: str = None) -> str:
        """Generate a unique hash ID for the report URL."""
        if timestamp is None:
            timestamp = datetime.now().isoformat()
        
        # Create a hash from respondent ID + timestamp + random salt
        salt = secrets.token_hex(8)
        hash_input = f"{respondent_id}_{timestamp}_{salt}".encode('utf-8')
        hash_id = hashlib.sha256(hash_input).hexdigest()[:16]  # Use first 16 chars for shorter URL
        
        return hash_id
    
    def _load_template(self, template_name: str) -> str:
        """Load HTML template from file."""
        template_path = self.template_dir / template_name
        if not template_path.exists():
            raise FileNotFoundError(f"Template not found: {template_path}")
        with open(template_path, 'r', encoding='utf-8') as f:
            return f.read()
    
    def generate_individual_report(
        self,
        respondent: Dict[str, Any],
        pm01_final: Dict[str, Any],
        pm05_final: Optional[Dict[str, Any]] = None
    ) -> Dict[str, str]:
        """
        Generate individual report HTML from PM01 Final and PM05 Final results.
        
        Args:
            respondent: Respondent data with id, name, company_name
            pm01_final: PM01 Final results with scores, summary, etc.
            pm05_final: PM05 Final results with consistency check (optional)
        
        Returns:
            Dictionary with 'filepath' and 'url' keys
        """
        # Extract data for report
        report_data = self._prepare_report_data(respondent, pm01_final, pm05_final)
        
        # Generate HTML
        html_content = self._generate_html(report_data)
        
        # Generate hash ID and URL
        timestamp = datetime.now().isoformat()
        hash_id = self._generate_hash_id(respondent['id'], timestamp)
        report_url = f"{self.REPORT_BASE_URL}/{hash_id}"
        
        # Save to file (use hash_id as filename for easier lookup)
        filename = f"{hash_id}.html"
        filepath = self.output_dir / filename
        
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        # Store URL mapping in Google Sheets if sheets_service is available
        if self.sheets_service:
            try:
                self.sheets_service.write_report_url(
                    respondent_id=respondent['id'],
                    hash_id=hash_id,
                    filepath=str(filepath),
                    report_url=report_url,
                    timestamp=timestamp,
                    report_type='individual'
                )
            except Exception as e:
                print(f"Warning: Could not store report URL mapping: {e}")
        
        return {
            'filepath': str(filepath),
            'url': report_url,
            'hash_id': hash_id
        }
    
    def _prepare_report_data(
        self,
        respondent: Dict[str, Any],
        pm01_final: Dict[str, Any],
        pm05_final: Optional[Dict[str, Any]]
    ) -> Dict[str, Any]:
        """Prepare data structure for report generation."""
        # Get total_score and level from diagnosis results (no default calculation)
        total_score = pm01_final.get('total_score', 0)
        level = pm01_final.get('level', '') or (pm05_final.get('level', '') if pm05_final else '')
        
        # Extract PRIMARY scores (using official categories)
        scores_primary = pm01_final.get('scores_primary', {})
        primary_data = {
            '問題理解': scores_primary.get('問題理解', 0),
            '論理思考': scores_primary.get('論理思考', 0) or scores_primary.get('論理構成', 0),  # Support old name
            '仮説構築': scores_primary.get('仮説構築', 0),
            'AI指示': scores_primary.get('AI指示', 0),
            'AI検証/優先順位判断': scores_primary.get('AI検証/優先順位判断', 0) or scores_primary.get('AI検証', 0)  # Support old name
        }
        
        # Extract PROCESS scores (use Japanese labels)
        process_scores = pm01_final.get('process', {})
        process_data = {
            self.PROCESS_LABELS_JP['clarity']: process_scores.get('clarity', 0),
            self.PROCESS_LABELS_JP['structure']: process_scores.get('structure', 0),
            self.PROCESS_LABELS_JP['hypothesis']: process_scores.get('hypothesis', 0),
            self.PROCESS_LABELS_JP['prompt clarity']: process_scores.get('prompt clarity', 0),
            self.PROCESS_LABELS_JP['consistency']: process_scores.get('consistency', 0)
        }
        
        # Calculate PRIMARY average: average of all PRIMARY items displayed
        primary_values = [v for v in primary_data.values() if isinstance(v, (int, float))]
        primary_avg = sum(primary_values) / len(primary_values) if primary_values else 0
        
        # Calculate PROCESS average: average of all PROCESS items displayed
        process_values = [v for v in process_data.values() if isinstance(v, (int, float))]
        process_avg = sum(process_values) / len(process_values) if process_values else 0
        
        # Calculate AES components (average from per_question if available)
        per_question = pm01_final.get('per_question', {})
        aes_clarity_list = []
        aes_logic_list = []
        aes_relevance_list = []
        
        for q_data in per_question.values():
            if isinstance(q_data, dict):
                if 'aes_clarity' in q_data:
                    aes_clarity_list.append(q_data['aes_clarity'])
                if 'aes_logic' in q_data:
                    aes_logic_list.append(q_data['aes_logic'])
                if 'aes_relevance' in q_data:
                    aes_relevance_list.append(q_data['aes_relevance'])
        
        aes_clarity = sum(aes_clarity_list) / len(aes_clarity_list) if aes_clarity_list else 0
        aes_logic = sum(aes_logic_list) / len(aes_logic_list) if aes_logic_list else 0
        aes_relevance = sum(aes_relevance_list) / len(aes_relevance_list) if aes_relevance_list else 0
        
        # Calculate AES average: average of the three components (Clarity, Logic, Relevance)
        # Formula: (Clarity + Logic + Relevance) / 3
        aes_avg = (aes_clarity + aes_logic + aes_relevance) / 3
        
        # Get AES comment from diagnosis results (no default generation)
        aes_comment = pm01_final.get('aes_comment', '') or (pm05_final.get('aes_comment', '') if pm05_final else '')
        
        # Get overall comment from PM05 Final if available, otherwise use PM01 Final summary
        overall_comment = ""
        if pm05_final and pm05_final.get('comment'):
            overall_comment = pm05_final['comment']
        elif pm01_final.get('overall_summary'):
            overall_comment = pm01_final['overall_summary']
        
        # Get status from diagnosis results (no default)
        status = pm05_final.get('status', '') if pm05_final else ''
        
        return {
            'respondent_id': respondent.get('id', ''),
            'respondent_name': respondent.get('name', ''),
            'diagnosis_date': datetime.now().strftime('%Y年%m月%d日'),
            'total_score': round(total_score, 1),
            'level': level,
            'primary_data': primary_data,
            'primary_avg': round(primary_avg, 1),
            'process_data': process_data,
            'process_avg': round(process_avg, 1),
            'aes_avg': round(aes_avg, 1),
            'aes_clarity': round(aes_clarity, 1),
            'aes_logic': round(aes_logic, 1),
            'aes_relevance': round(aes_relevance, 1),
            'aes_comment': aes_comment,
            'overall_comment': overall_comment,
            'status': status,
            'ai_use_level': pm01_final.get('ai_use_level', '')
        }
    
    def _json_escape(self, json_str: str) -> str:
        """Escape JSON string for embedding in JavaScript."""
        return json_str.replace('\\', '\\\\').replace('"', '\\"')
    
    def _generate_html(self, data: Dict[str, Any]) -> str:
        """Generate HTML content for individual report."""
        # Load template
        template = self._load_template("individual_report.html")
        
        # Prepare data for template (no HTML generation, only JSON data)
        respondent_name_html = f'<div class="name">{data["respondent_name"]}</div>' if data["respondent_name"] else ''
        process_data_json = json.dumps(data["process_data"], ensure_ascii=False)
        primary_data_json = json.dumps(data["primary_data"], ensure_ascii=False)
        
        # Prepare status HTML if available
        status = data.get('status', '')
        if status:
            status_display = {
                'valid': '妥当',
                'caution': '注意',
                're-eval': '再評価',
                '妥当': '妥当',
                '注意': '注意',
                '再評価': '再評価'
            }.get(status, status)
            status_html = f'<div class="status-badge status-{status.lower().replace("-", "_")}">{status_display}</div>'
        else:
            status_html = ''
        
        # Replace placeholders using Template.safe_substitute
        template_obj = Template(template)
        return template_obj.safe_substitute(
            respondent_name_html=respondent_name_html,
            diagnosis_date=data["diagnosis_date"],
            total_score=data["total_score"],
            level=data["level"],
            primary_avg=data["primary_avg"],
            process_avg=data["process_avg"],
            aes_avg=data["aes_avg"],
            primary_data_json=primary_data_json,
            aes_clarity=data["aes_clarity"],
            aes_logic=data["aes_logic"],
            aes_relevance=data["aes_relevance"],
            aes_comment=data["aes_comment"],
            status_html=status_html,
            overall_comment=data["overall_comment"],
            process_data_json=process_data_json
        )
    
    def generate_organization_report(
        self,
        company_name: str,
        sheets_service: Any,
        department_filter: Optional[str] = None
    ) -> Dict[str, str]:
        """
        Generate organization report HTML from PM1Final data grouped by company.
        
        Args:
            company_name: Company name to generate report for
            sheets_service: SheetsService instance to read data
            department_filter: Optional department name to filter by
        
        Returns:
            Dictionary with 'filepath', 'url', and 'hash_id' keys
        """
        # Read organization data
        org_data = self._read_organization_data(company_name, sheets_service, department_filter)
        
        if not org_data or org_data['count'] == 0:
            raise ValueError(f"No data found for company: {company_name}")
        
        # Prepare report data
        report_data = self._prepare_organization_data(org_data)
        
        # Generate HTML
        html_content = self._generate_organization_html(report_data)
        
        # Generate hash ID and URL
        timestamp = datetime.now().isoformat()
        hash_id = self._generate_hash_id(company_name, timestamp)
        report_url = f"{self.REPORT_BASE_URL}/{hash_id}"
        
        # Save to file (use hash_id as filename for easier lookup)
        filename = f"{hash_id}.html"
        filepath = self.output_dir / filename
        
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        # Store URL mapping in Google Sheets if sheets_service is available
        if self.sheets_service:
            try:
                self.sheets_service.write_report_url(
                    respondent_id=company_name,
                    hash_id=hash_id,
                    filepath=str(filepath),
                    report_url=report_url,
                    timestamp=timestamp,
                    report_type='organization',
                    company_name=company_name,
                    department=department_filter
                )
            except Exception as e:
                print(f"Warning: Could not store report URL mapping: {e}")
        
        return {
            'filepath': str(filepath),
            'url': report_url,
            'hash_id': hash_id
        }
    
    def _read_organization_data(
        self,
        company_name: str,
        sheets_service: Any,
        department_filter: Optional[str]
    ) -> Dict[str, Any]:
        """Read and aggregate organization data from PM1Final sheet."""
        pm1final_sheet = sheets_service._get_sheet('PM1Final')
        if not pm1final_sheet:
            return {'count': 0, 'data': []}
        
        values = pm1final_sheet.get_all_values()
        if len(values) <= 1:
            return {'count': 0, 'data': []}
        
        # Read respondents to get department info
        respondents = sheets_service.get_respondent_rows()
        respondent_map = {r['id']: r for r in respondents}
        
        # Read department info directly from respondents sheet (column 4)
        respondents_sheet = sheets_service._get_sheet(sheets_service.config['respondentsSheet'])
        department_map = {}
        if respondents_sheet:
            resp_values = respondents_sheet.get_all_values()
            for row in resp_values[1:]:
                if len(row) > 5:
                    resp_id = str(row[0] or '').strip()
                    dept = str(row[4] or '').strip() if len(row) > 4 else ''
                    if resp_id:
                        department_map[resp_id] = dept
        
        org_data = []
        for row in values[1:]:
            if len(row) < 11:
                continue
            
            row_company = row[1] if len(row) > 1 else ''  # Company_Name column
            if row_company.strip() != company_name.strip():
                continue
            
            respondent_id = row[0] if len(row) > 0 else ''
            respondent = respondent_map.get(respondent_id, {})
            
            # Filter by department if specified (department is in column 4 of respondents sheet)
            if department_filter:
                # Try to get department from respondent data or from respondents sheet
                dept = ''
                if respondent_id in respondent_map:
                    dept = respondent_map[respondent_id].get('department', '')
                if not dept and len(row) > 4:
                    # Try to get from PM1Final row if available (though it's not stored there)
                    pass
                if dept.strip() != department_filter.strip():
                    continue
            
            try:
                total_score = float(row[3]) if len(row) > 3 and row[3] else 0
                scores_primary = json.loads(row[4]) if len(row) > 4 and row[4] else {}
                scores_sub = json.loads(row[5]) if len(row) > 5 and row[5] else {}
                process_scores = json.loads(row[6]) if len(row) > 6 and row[6] else {}
                aes_scores = json.loads(row[7]) if len(row) > 7 and row[7] else {}
                ai_use_level = row[9] if len(row) > 9 else ''
                
                # Get department from department_map
                department = department_map.get(respondent_id, '')
                
                org_data.append({
                    'respondent_id': respondent_id,
                    'name': respondent.get('name', ''),
                    'department': department,
                    'total_score': total_score,
                    'scores_primary': scores_primary,
                    'scores_sub': scores_sub,
                    'process': process_scores,
                    'aes': aes_scores,
                    'ai_use_level': ai_use_level
                })
            except (ValueError, json.JSONDecodeError) as e:
                print(f"Warning: Error parsing row for {respondent_id}: {e}")
                continue
        
        return {
            'count': len(org_data),
            'data': org_data,
            'company_name': company_name
        }
    
    def _prepare_organization_data(self, org_data: Dict[str, Any]) -> Dict[str, Any]:
        """Prepare data structure for organization report."""
        data_list = org_data['data']
        count = org_data['count']
        
        if count == 0:
            return {}
        
        # Calculate average total score
        total_scores = [d['total_score'] for d in data_list]
        avg_total_score = statistics.mean(total_scores) if total_scores else 0
        
        # Aggregate PRIMARY scores by category (using official categories)
        primary_categories = ['問題理解', '論理思考', '仮説構築', 'AI指示', 'AI検証/優先順位判断']
        primary_distributions = {}
        primary_means = []
        
        for category in primary_categories:
            scores = []
            for d in data_list:
                # Support both old and new category names
                cat_score = d['scores_primary'].get(category, 0)
                if cat_score == 0:
                    # Try old names for backward compatibility
                    if category == '論理思考':
                        cat_score = d['scores_primary'].get('論理構成', 0)
                    elif category == 'AI検証/優先順位判断':
                        cat_score = d['scores_primary'].get('AI検証', 0)
                if cat_score > 0:
                    scores.append(cat_score)
            if scores:
                mean_score = statistics.mean(scores)
                primary_means.append(mean_score)
                primary_distributions[category] = {
                    'min': round(min(scores), 1),
                    'q1': round(statistics.quantiles(scores, n=4)[0] if len(scores) > 1 else scores[0], 1),
                    'median': round(statistics.median(scores), 1),
                    'q3': round(statistics.quantiles(scores, n=4)[2] if len(scores) > 1 else scores[0], 1),
                    'max': round(max(scores), 1),
                    'mean': round(mean_score, 1),
                    'values': [round(s, 1) for s in scores]
                }
        
        # Calculate PRIMARY average (average of all PRIMARY category means)
        primary_avg = round(statistics.mean(primary_means), 1) if primary_means else 0
        
        # Aggregate PROCESS scores (use Japanese labels)
        process_categories_en = ['clarity', 'structure', 'hypothesis', 'prompt clarity', 'consistency']
        process_averages = {}
        process_values = {cat: [] for cat in process_categories_en}
        process_means = []
        
        for d in data_list:
            for cat in process_categories_en:
                score = d['process'].get(cat, 0)
                if score > 0:
                    process_values[cat].append(score)
        
        for cat in process_categories_en:
            if process_values[cat]:
                mean_score = statistics.mean(process_values[cat])
                process_means.append(mean_score)
                jp_label = self.PROCESS_LABELS_JP[cat]
                process_averages[jp_label] = {
                    'mean': round(mean_score, 1),
                    'std': round(statistics.stdev(process_values[cat]) if len(process_values[cat]) > 1 else 0, 1),
                    'min': round(min(process_values[cat]), 1),
                    'max': round(max(process_values[cat]), 1),
                    'values': [round(s, 1) for s in process_values[cat]]
                }
        
        # Calculate PROCESS average (average of all PROCESS category means)
        process_avg = round(statistics.mean(process_means), 1) if process_means else 0
        
        # Aggregate AES scores
        aes_categories = ['clarity', 'logic', 'relevance']
        aes_averages = {}
        aes_values = {cat: [] for cat in aes_categories}
        aes_means = []
        
        for d in data_list:
            aes_scores = d.get('aes', {})
            for cat in aes_categories:
                # Support both 'aes_clarity' format and 'clarity' format
                score = aes_scores.get(f'aes_{cat}', aes_scores.get(cat, 0))
                if score == 0:
                    # Try alternative formats
                    if cat == 'clarity':
                        score = aes_scores.get('aes_clarity', 0)
                    elif cat == 'logic':
                        score = aes_scores.get('aes_logic', 0)
                    elif cat == 'relevance':
                        score = aes_scores.get('aes_relevance', 0)
                if score > 0:
                    aes_values[cat].append(score)
        
        aes_labels_jp = {
            'clarity': '明瞭さ',
            'logic': '論理性',
            'relevance': '関連性'
        }
        
        for cat in aes_categories:
            if aes_values[cat]:
                mean_score = statistics.mean(aes_values[cat])
                aes_means.append(mean_score)
                jp_label = aes_labels_jp[cat]
                aes_averages[jp_label] = {
                    'mean': round(mean_score, 1),
                    'std': round(statistics.stdev(aes_values[cat]) if len(aes_values[cat]) > 1 else 0, 1),
                    'min': round(min(aes_values[cat]), 1),
                    'max': round(max(aes_values[cat]), 1),
                    'values': [round(s, 1) for s in aes_values[cat]]
                }
        
        # Calculate AES average (average of all AES component means)
        aes_avg = round(statistics.mean(aes_means), 1) if aes_means else 0
        
        # Get AI maturity rating from diagnosis results (no default calculation)
        # Count AI use levels from actual data (no hardcoded categories)
        ai_level_counts = {}
        for d in data_list:
            level = d.get('ai_use_level', '')
            if level:
                ai_level_counts[level] = ai_level_counts.get(level, 0) + 1
        
        # Get maturity rating from diagnosis results (should be stored in PM1Final or PM5Final)
        maturity_rating = ''
        if data_list:
            maturity_rating = data_list[0].get('maturity_rating', '')
        
        # Get trend analysis from diagnosis results (no default generation)
        trend_analysis = ''
        if data_list:
            trend_analysis = data_list[0].get('trend_analysis', '')
        
        # Get CTA section from first respondent's data if available, otherwise use default
        cta_section_html = ''
        if data_list:
            cta_section_html = data_list[0].get('cta_section_html', '')
        
        # Default CTA section for organization reports if not provided
        if not cta_section_html:
            cta_section_html = '''<div class="cta-section">
<div class="cta-title">本診断（AI-CATS）への誘導</div>
<div class="cta-content">
<p><strong>組織診断の必要性</strong></p>
<p>この簡易診断では、組織全体の思考傾向を可視化しました。より詳細な分析と改善提案をご希望の場合は、AI-CATS本診断をご利用ください。</p>
<p><strong>標準導入プロセス</strong></p>
<ul>
<li>診断設計：組織の課題に合わせた診断項目のカスタマイズ</li>
<li>診断実施：全社員または対象部署への診断実施</li>
<li>結果分析：詳細なスキル評価と組織分析レポートの提供</li>
<li>改善支援：診断結果に基づく研修・育成プログラムの提案</li>
</ul>
<p><strong>参考納期・料金（任意）</strong></p>
<p>詳細な納期・料金については、お問い合わせください。</p>
</div>
<a href="https://ai-cats.gs-group.jp/@ai-cats1" target="_blank" class="cta-button">AI-CATS本診断へ</a>
</div>'''
        
        # Get department information (if filtered by department, show it; otherwise show all departments)
        departments = set()
        for d in data_list:
            dept = d.get('department', '')
            if dept:
                departments.add(dept)
        department_display = ', '.join(sorted(departments)) if departments else ''
        
        return {
            'company_name': org_data['company_name'],
            'department': department_display,
            'count': count,
            'avg_total_score': round(avg_total_score, 1),
            'primary_distributions': primary_distributions,
            'primary_avg': primary_avg,
            'process_averages': process_averages,
            'process_avg': process_avg,
            'aes_averages': aes_averages,
            'aes_avg': aes_avg,
            'maturity_rating': maturity_rating,
            'maturity_description': '',  # Should come from diagnosis results
            'ai_level_distribution': ai_level_counts,
            'trend_analysis': trend_analysis,
            'generation_date': datetime.now().strftime('%Y年%m月%d日'),
            'cta_section_html': cta_section_html
        }
    
    
    def _generate_organization_html(self, data: Dict[str, Any]) -> str:
        """Generate HTML content for organization report."""
        # Load template
        template = self._load_template("organization_report.html")
        
        # Prepare data for template (no HTML generation, only JSON data)
        primary_data_json = json.dumps(data['primary_distributions'], ensure_ascii=False)
        process_data_json = json.dumps(data['process_averages'], ensure_ascii=False)
        aes_data_json = json.dumps(data.get('aes_averages', {}), ensure_ascii=False)
        # Get maturity description from diagnosis results (no default generation)
        maturity_description_html = f'<div class="maturity-description">{data.get("maturity_description", "")}</div>' if data.get('maturity_description') else ''
        
        # Get CTA section from diagnosis results (no default)
        cta_section_html = data.get('cta_section_html', '')
        
        # Prepare department HTML (to be appended to company name)
        department = data.get('department', '')
        department_html = f' / {department}' if department else ''
        
        # Replace placeholders using Template.safe_substitute
        template_obj = Template(template)
        return template_obj.safe_substitute(
            company_name=data['company_name'],
            department_html=department_html,
            generation_date=data['generation_date'],
            count=data['count'],
            avg_total_score=data['avg_total_score'],
            primary_avg=data.get('primary_avg', 0),
            process_avg=data.get('process_avg', 0),
            aes_avg=data.get('aes_avg', 0),
            trend_analysis=data['trend_analysis'],
            maturity_rating=data['maturity_rating'],
            maturity_description_html=maturity_description_html,
            primary_data_json=primary_data_json,
            process_data_json=process_data_json,
            aes_data_json=aes_data_json,
            cta_section_html=cta_section_html
        )

