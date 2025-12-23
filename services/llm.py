"""
LLM service - handles PM01 and PM05 LLM API calls with strict JSON output.
"""

import json
import requests
from typing import Dict, List, Any, Optional
from services.json_parser import JsonParser
from core.category_mapper import map_to_official_category


class LLMService:
    """Service for interacting with LLM APIs for PM01 and PM05."""
    
    def __init__(self, config: Dict[str, Any]):
        """Initialize the LLM service with configuration."""
        self.config = config
        self.json_parser = JsonParser()
    
    def _map_to_official_category(self, category: str, category_type: str) -> Optional[str]:
        """Map question sheet category to official category (delegates to shared mapper)."""
        return map_to_official_category(category, category_type)
    
    def run_pm01_raw_scoring(
        self,
        respondent: Dict[str, Any],
        question: Dict[str, Any],
        question_index: int,
        attempt: int
    ) -> Optional[Dict[str, Any]]:
        """
        STEP 1: PM01 Raw Scoring - Individual Q-A scoring.
        
        Returns structured JSON with scores, evidence, judgment reason, and increase/decrease points.
        """
        prompt = self._build_pm01_raw_prompt(respondent, question, question_index)
        
        response = self._invoke_llm(prompt, attempt, "You are an expert evaluator. Output ONLY valid JSON.")
        if not response:
            return None
        
        parsed = self.json_parser.parse_pm01_raw_response(response, question['number'])
        return parsed
    
    def run_pm05_raw_scoring(
        self,
        respondent: Dict[str, Any],
        question: Dict[str, Any],
        question_index: int,
        pm01_raw_result: Dict[str, Any],
        attempt: int
    ) -> Optional[Dict[str, Any]]:
        """
        STEP 2: PM05 Raw Scoring - Reverse logic scoring using PM01 raw as reference.
        
        Returns structured JSON with reverse-scored results.
        """
        prompt = self._build_pm05_raw_prompt(respondent, question, question_index, pm01_raw_result)
        
        response = self._invoke_llm(prompt, attempt, "You are a validation evaluator using reverse logic. Output ONLY valid JSON.")
        if not response:
            return None
        
        parsed = self.json_parser.parse_pm05_raw_response(response, question['number'])
        return parsed
    
    def run_pm01_final_analysis(
        self,
        respondent: Dict[str, Any],
        pm05_raw_results: Dict[str, Dict[str, Any]],
        aggregated_scores: Dict[str, Any],
        attempt: int
    ) -> Optional[Dict[str, Any]]:
        """
        STEP 3: PM01 Final - Analysis and interpretation of aggregated scores from PM05 Raw.
        
        Returns structured JSON with analysis insights.
        """
        prompt = self._build_pm01_final_prompt(respondent, pm05_raw_results, aggregated_scores)
        
        response = self._invoke_llm(prompt, attempt, "You are an expert analyst. Output ONLY valid JSON with all text in Japanese.")
        if not response:
            return None
        
        parsed = self.json_parser.parse_pm01_final_response(response)
        return parsed
    
    def run_pm05_final_check(
        self,
        respondent: Dict[str, Any],
        pm01_final: Dict[str, Any],
        attempt: int
    ) -> Optional[Dict[str, Any]]:
        """
        STEP 4: PM05 Final - Overall consistency check using PM01 Final.
        
        Returns structured JSON with consistency evaluation.
        """
        prompt = self._build_pm05_final_prompt(respondent, pm01_final)
        
        response = self._invoke_llm(prompt, attempt, "You are a consistency evaluator. Output ONLY valid JSON with comment field in Japanese.")
        if not response:
            return None
        
        parsed = self.json_parser.parse_pm05_final_response(response)
        return parsed
    
    def _build_pm01_raw_prompt(
        self,
        respondent: Dict[str, Any],
        question: Dict[str, Any],
        question_index: int
    ) -> str:
        """Build PM01 Raw Scoring prompt for a single question."""
        lines = []
        
        lines.append("# Respondent Information")
        lines.append(f"ID: {respondent['id']}")
        lines.append(f"Name: {respondent['name']}")
        lines.append("")
        
        # Single question, answer, and reason
        question_id = f"Q{question['number']}"
        answer = respondent['answers'][question_index] if question_index < len(respondent['answers']) else ""
        reason = respondent.get('reasons', [])
        reason_text = reason[question_index] if question_index < len(reason) else ""
        
        lines.append(f"# Question {question_id}")
        lines.append(f"Question: {question['questionText']}")
        lines.append(f"Answer: {answer or '(無回答)'}")
        if reason_text:
            lines.append(f"選択理由（AI分析用）: {reason_text}")
        
        # Add category information if available (use mapped official categories)
        question_primary_cat = question.get('primary_category', '')
        question_sub_cat = question.get('sub_category', '')
        question_process_cat = question.get('process_category', '')
        
        # Map to official categories
        primary_cat = self._map_to_official_category(question_primary_cat, 'primary')
        sub_cat = self._map_to_official_category(question_sub_cat, 'sub')
        process_cat = self._map_to_official_category(question_process_cat, 'process')
        
        if primary_cat or sub_cat or process_cat:
            lines.append("")
            lines.append("# Evaluation Categories")
            if primary_cat:
                lines.append(f"PRIMARY Category: {primary_cat}")
            if sub_cat:
                lines.append(f"SUB Category: {sub_cat}")
            if process_cat:
                lines.append(f"PROCESS Category: {process_cat}")
        
        lines.append("")
        
        # Get prompt from config sheet (required for STEP 1)
        prompt_text = self.config.get('promptPM1Raw', '').strip()
        if not prompt_text:
            raise ValueError("promptPM1Raw not configured in Config sheet. Required for STEP 1 (PM01 Raw Scoring).")
        
        lines.append("# Evaluation Instructions")
        lines.append(prompt_text)
        lines.append("")
        
        lines.append("# Required JSON Schema")
        lines.append("""{
  "primary_score": <1.0-5.0 float with 1 decimal place, required - final score for スキル評価（Primary）>,
  "sub_score": <1.0-5.0 float with 1 decimal place, required - final score for スキル評価（Sub）>,
  "process_score": <1.0-5.0 float with 1 decimal place, required - final score for PROCESS評価>,
  "aes_clarity": <1.0-5.0 float with 1 decimal place, required>,
  "aes_logic": <1.0-5.0 float with 1 decimal place, required>,
  "aes_relevance": <1.0-5.0 float with 1 decimal place, required>,
  "evidence": "<string - specific evidence from answer, required>",
  "judgment_reason": "<string - reason for the scores referencing official criteria. MUST mention any bonus/penalty conditions (+0.1 to +0.5 or -0.1 to -0.5) applied to PRIMARY/SUB/PROCESS scores, required>"
}""")
        
        return "\n".join(lines)
    
    def _build_pm05_raw_prompt(
        self,
        respondent: Dict[str, Any],
        question: Dict[str, Any],
        question_index: int,
        pm01_raw_result: Dict[str, Any]
    ) -> str:
        """Build PM05 Raw Scoring prompt using reverse logic."""
        lines = []
        
        lines.append("# Respondent Information")
        lines.append(f"ID: {respondent['id']}")
        lines.append(f"Name: {respondent['name']}")
        lines.append("")
        
        # Single question, answer, and reason
        question_id = f"Q{question['number']}"
        answer = respondent['answers'][question_index] if question_index < len(respondent['answers']) else ""
        reason = respondent.get('reasons', [])
        reason_text = reason[question_index] if question_index < len(reason) else ""
        
        lines.append(f"# Question {question_id}")
        lines.append(f"Question: {question['questionText']}")
        lines.append(f"Answer: {answer or '(無回答)'}")
        if reason_text:
            lines.append(f"選択理由（AI分析用）: {reason_text}")
        
        # Add category information if available (use mapped official categories)
        question_primary_cat = question.get('primary_category', '')
        question_sub_cat = question.get('sub_category', '')
        question_process_cat = question.get('process_category', '')
        
        # Map to official categories
        primary_cat = self._map_to_official_category(question_primary_cat, 'primary')
        sub_cat = self._map_to_official_category(question_sub_cat, 'sub')
        process_cat = self._map_to_official_category(question_process_cat, 'process')
        
        if primary_cat or sub_cat or process_cat:
            lines.append("")
            lines.append("# Evaluation Categories")
            if primary_cat:
                lines.append(f"PRIMARY Category: {primary_cat}")
            if sub_cat:
                lines.append(f"SUB Category: {sub_cat}")
            if process_cat:
                lines.append(f"PROCESS Category: {process_cat}")
        
        lines.append("")
        
        # Include PM01 raw scoring result
        lines.append("# PM01 Raw Scoring Result (Reference)")
        lines.append(f"Primary Score: {pm01_raw_result.get('primary_score', 0)}")
        lines.append(f"Sub Score: {pm01_raw_result.get('sub_score', 0)}")
        lines.append(f"Process Score: {pm01_raw_result.get('process_score', 0)}")
        lines.append(f"AES Clarity: {pm01_raw_result.get('aes_clarity', 0)}")
        lines.append(f"AES Logic: {pm01_raw_result.get('aes_logic', 0)}")
        lines.append(f"AES Relevance: {pm01_raw_result.get('aes_relevance', 0)}")
        aes_score = (pm01_raw_result.get('aes_clarity', 0) + pm01_raw_result.get('aes_logic', 0) + pm01_raw_result.get('aes_relevance', 0)) / 3 if (pm01_raw_result.get('aes_clarity', 0) + pm01_raw_result.get('aes_logic', 0) + pm01_raw_result.get('aes_relevance', 0)) > 0 else 0
        lines.append(f"AES Score (Average): {round(aes_score, 1)}")
        lines.append(f"Evidence: {pm01_raw_result.get('evidence', '')}")
        lines.append(f"Judgment Reason: {pm01_raw_result.get('judgment_reason', '')}")
        lines.append("")
        
        # Get prompt from config sheet
        prompt_text = self.config.get('promptPM5Raw', '').strip()
        if prompt_text:
            lines.append("# Reverse Logic Evaluation Instructions")
            lines.append(prompt_text)
            lines.append("")
        else:
            raise ValueError("promptPM5Raw not configured in Config sheet. Required for STEP 2 (PM05 Raw Scoring).")
        
        lines.append("# Required Output JSON Schema")
        lines.append("""{
  "primary_score": <1.0-5.0 float with 1 decimal place, required - your reverse logic evaluation score>,
  "sub_score": <1.0-5.0 float with 1 decimal place, required - your reverse logic evaluation score>,
  "process_score": <1.0-5.0 float with 1 decimal place, required - your reverse logic evaluation score>,
  "difference_note": "<string in Japanese - detailed explanation covering: your reverse logic evaluation approach, comparison with PM01 Raw scores, any inconsistencies/contradictions/issues detected, explanation of score differences (if any), consistency assessment for this question, required>"
}""")
        
        return "\n".join(lines)
    
    def _build_pm01_final_prompt(
        self,
        respondent: Dict[str, Any],
        pm05_raw_results: Dict[str, Dict[str, Any]],
        aggregated_scores: Dict[str, Any]
    ) -> str:
        """Build PM01 Final analysis prompt using PM05 Raw validated scores."""
        lines = []
        
        lines.append("# Respondent Information")
        lines.append(f"ID: {respondent['id']}")
        lines.append(f"Name: {respondent['name']}")
        lines.append("")
        
        lines.append("# Aggregated Scores (from PM05 Raw validated scores)")
        lines.append(f"Primary Scores: {aggregated_scores.get('scores_primary', {})}")
        lines.append(f"Sub Scores: {aggregated_scores.get('scores_sub', {})}")
        lines.append(f"Process Scores: {aggregated_scores.get('process', {})}")
        lines.append(f"Total Score: {aggregated_scores.get('total_score', 0)}")
        lines.append("")
        
        # Get prompt from config sheet (required for STEP 3)
        prompt_text = self.config.get('promptPM1Final', '').strip()
        if not prompt_text:
            raise ValueError("promptPM1Final not configured in Config sheet. Required for STEP 3 (PM01 Final).")
        
        lines.append("# Analysis Instructions")
        lines.append(prompt_text)
        lines.append("")
        
        lines.append("# Required JSON Schema")
        lines.append("""{
  "overall_summary": "<comprehensive summary of the diagnosis>",
  "ai_use_level": "<基礎|標準|高度>",
  "recommendations": ["<recommendation1>", "<recommendation2>", ...]
}""")
        
        return "\n".join(lines)
    
    def _build_pm05_final_prompt(
        self,
        respondent: Dict[str, Any],
        pm01_final: Dict[str, Any]
    ) -> str:
        """Build PM05 Final consistency check prompt."""
        lines = []
        
        lines.append("# Respondent Information")
        lines.append(f"ID: {respondent['id']}")
        lines.append(f"Name: {respondent['name']}")
        lines.append("")
        
        lines.append("# PM01 Final Result")
        lines.append(f"Total Score: {pm01_final.get('total_score', 0)}")
        lines.append(f"Primary Scores (Aggregated): {pm01_final.get('scores_primary', {})}")
        lines.append(f"Sub Scores (Aggregated): {pm01_final.get('scores_sub', {})}")
        lines.append(f"Process Scores (Aggregated): {pm01_final.get('process', {})}")
        lines.append(f"AES Scores (Per Question): {pm01_final.get('aes', {})}")
        lines.append(f"Overall Summary: {pm01_final.get('overall_summary', '')}")
        lines.append(f"AI Use Level: {pm01_final.get('ai_use_level', '')}")
        lines.append(f"Recommendations: {pm01_final.get('recommendations', [])}")
        lines.append("")
        
        # Include per-question scores for validation
        per_question = pm01_final.get('per_question', {})
        if per_question:
            lines.append("# Per-Question Scores (Q1-Q6)")
            for q_id in ['Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6']:
                q_data = per_question.get(q_id, {})
                if q_data:
                    lines.append(f"{q_id}: Primary={q_data.get('primary_score', 0)}, "
                               f"Sub={q_data.get('sub_score', 0)}, "
                               f"Process={q_data.get('process_score', 0)}, "
                               f"AES={q_data.get('aes_score', 0)}, "
                               f"AES_Clarity={q_data.get('aes_clarity', 0)}, "
                               f"AES_Logic={q_data.get('aes_logic', 0)}, "
                               f"AES_Relevance={q_data.get('aes_relevance', 0)}")
            lines.append("")
        
        # Get prompt from config sheet (can use promptPM5Raw or separate promptPM5Final)
        prompt_text = self.config.get('promptPM5Final', '').strip()
        if not prompt_text:
            # Fallback to promptPM5Raw if promptPM5Final not set
            prompt_text = self.config.get('promptPM5Raw', '').strip()
        
        if prompt_text:
            lines.append("# Consistency Check Instructions")
            lines.append(prompt_text)
            lines.append("")
        else:
            raise ValueError("promptPM5Raw or promptPM5Final not configured in Config sheet. Required for STEP 4 (PM05 Final).")
        
        lines.append("# Required JSON Schema")
        lines.append("""{
  "consistency_score": <0.0-1.0 float with 2 decimal places, required - calculated as 1 - (score_std / 2.5), where score_std is standard deviation of PRIMARY/SUB/PROCESS scores across Q1-Q6>,
  "status": "<妥当|注意|再評価>",
  "detected_issues": ["<issue1 in Japanese>", "<issue2 in Japanese>", ...],
  "comment": "<string in Japanese, 80-120 characters - consistency evaluation, score trends, re-diagnosis recommendation if needed, required>"
}""")
        
        return "\n".join(lines)
    
    def _invoke_llm(self, prompt: str, attempt: int, system_message: str) -> Optional[str]:
        """Sends the prepared prompt to the configured LLM provider."""
        api_key = self.config.get('llmApiKey')
        if not api_key:
            raise ValueError("API key not configured. Add llmApiKey to the Config sheet in your Google Spreadsheet.")
        
        provider = self.config.get('llmProvider')
        if not provider:
            raise ValueError("llmProvider not configured in Config sheet.")
        
        try:
            if provider == 'chatgpt':
                return self._invoke_chatgpt(prompt, system_message)
            else:
                raise ValueError(f"Unsupported provider: {provider}")
        except Exception as e:
            print(f"  LLM API error: {e}")
            return None
    
    def _invoke_chatgpt(self, prompt: str, system_message: str) -> Optional[str]:
        """Invokes the ChatGPT API."""
        url = self.config.get('llmApiUrl')
        if not url:
            raise ValueError("llmApiUrl not configured in Config sheet.")
        api_key = self.config['llmApiKey']
        model = self.config.get('llmModel')
        if not model:
            raise ValueError("llmModel not configured in Config sheet.")
        
        headers = {
            'Authorization': f'Bearer {api_key}',
            'Content-Type': 'application/json'
        }
        
        body = {
            'model': model,
            'messages': [
                {'role': 'system', 'content': system_message},
                {'role': 'user', 'content': prompt}
            ],
            'temperature': 0,
            'response_format': {'type': 'json_object'}  # Force JSON output
        }
        
        response = requests.post(url, headers=headers, json=body, timeout=300)
        response.raise_for_status()
        
        data = response.json()
        if not data.get('choices') or len(data['choices']) == 0:
            raise ValueError("No choices in LLM API response")
        
        content = data['choices'][0].get('message', {}).get('content')
        if not content:
            raise ValueError("No content in LLM API response")
        
        return content.strip()
