"""
LLM service - handles PM01 and PM05 LLM API calls with strict JSON output.
"""

import json
import requests
from typing import Dict, List, Any, Optional
from services.json_parser import JsonParser


class LLMService:
    """Service for interacting with LLM APIs for PM01 and PM05."""
    
    def __init__(self, config: Dict[str, Any]):
        """Initialize the LLM service with configuration."""
        self.config = config
        self.json_parser = JsonParser()
    
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
        print(f"  PM01 Raw Q{question['number']} Prompt length: {len(prompt)} characters")
        
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
        print(f"  PM05 Raw Q{question['number']} Prompt length: {len(prompt)} characters")
        
        response = self._invoke_llm(prompt, attempt, "You are a validation evaluator using reverse logic. Output ONLY valid JSON.")
        if not response:
            return None
        
        parsed = self.json_parser.parse_pm05_raw_response(response, question['number'])
        return parsed
    
    def run_pm01_final_analysis(
        self,
        respondent: Dict[str, Any],
        pm01_raw_results: Dict[str, Dict[str, Any]],
        aggregated_scores: Dict[str, Any],
        attempt: int
    ) -> Optional[Dict[str, Any]]:
        """
        STEP 3: PM01 Final - Analysis and interpretation of aggregated scores.
        
        Returns structured JSON with analysis insights.
        """
        prompt = self._build_pm01_final_prompt(respondent, pm01_raw_results, aggregated_scores)
        print(f"  PM01 Final Prompt length: {len(prompt)} characters")
        
        response = self._invoke_llm(prompt, attempt, "You are an expert analyst. Output ONLY valid JSON.")
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
        print(f"  PM05 Final Prompt length: {len(prompt)} characters")
        
        response = self._invoke_llm(prompt, attempt, "You are a consistency evaluator. Output ONLY valid JSON.")
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
        
        # Single question and answer
        question_id = f"Q{question['number']}"
        answer = respondent['answers'][question_index] if question_index < len(respondent['answers']) else ""
        
        lines.append(f"# Question {question_id}")
        lines.append(f"Question: {question['questionText']}")
        lines.append(f"Answer: {answer or '(無回答)'}")
        lines.append("")
        
        # Get prompt from config sheet (required for STEP 1)
        prompt_text = self.config.get('promptPM1', '').strip()
        if not prompt_text:
            raise ValueError("promptPM1 not configured in Config sheet. Required for STEP 1 (PM01 Raw Scoring).")
        
        lines.append("# Evaluation Instructions")
        lines.append(prompt_text)
        lines.append("")
        
        lines.append("# Required JSON Schema")
        lines.append(f"""{{
  "primary_score": <0-5>,
  "sub_score": <0-5>,
  "process_score": <0-5>,
  "aes_clarity": <0-5>,
  "aes_logic": <0-5>,
  "aes_relevance": <0-5>,
  "increase_points": <float, optional, default 0>,
  "decrease_points": <float, optional, default 0>,
  "evidence": "<string - specific evidence from answer>",
  "judgment_reason": "<string - reason for the scores>"
}}""")
        
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
        
        # Single question and answer
        question_id = f"Q{question['number']}"
        answer = respondent['answers'][question_index] if question_index < len(respondent['answers']) else ""
        
        lines.append(f"# Question {question_id}")
        lines.append(f"Question: {question['questionText']}")
        lines.append(f"Answer: {answer or '(無回答)'}")
        lines.append("")
        
        # Include PM01 raw scoring result
        lines.append("# PM01 Raw Scoring Result (Reference)")
        lines.append(f"Primary Score: {pm01_raw_result.get('primary_score', 0)}")
        lines.append(f"Sub Score: {pm01_raw_result.get('sub_score', 0)}")
        lines.append(f"Process Score: {pm01_raw_result.get('process_score', 0)}")
        lines.append(f"Evidence: {pm01_raw_result.get('evidence', '')}")
        lines.append(f"Judgment Reason: {pm01_raw_result.get('judgment_reason', '')}")
        lines.append("")
        
        # Get prompt from config sheet
        prompt_text = self.config.get('promptPM5', '').strip()
        if prompt_text:
            lines.append("# Reverse Logic Evaluation Instructions")
            lines.append(prompt_text)
            lines.append("")
        else:
            raise ValueError("promptPM5 not configured in Config sheet. Required for STEP 2 (PM05 Raw Scoring).")
        
        lines.append("# Required JSON Schema")
        lines.append(f"""{{
  "primary_score": <0-5>,
  "sub_score": <0-5>,
  "process_score": <0-5>,
  "difference_note": "<string - note on difference from PM01 raw>"
}}""")
        
        return "\n".join(lines)
    
    def _build_pm01_final_prompt(
        self,
        respondent: Dict[str, Any],
        pm01_raw_results: Dict[str, Dict[str, Any]],
        aggregated_scores: Dict[str, Any]
    ) -> str:
        """Build PM01 Final analysis prompt."""
        lines = []
        
        lines.append("# Respondent Information")
        lines.append(f"ID: {respondent['id']}")
        lines.append(f"Name: {respondent['name']}")
        lines.append("")
        
        lines.append("# PM01 Raw Scoring Results (Q1-Q6)")
        for q_id in ['Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6']:
            raw = pm01_raw_results.get(q_id, {})
            lines.append(f"{q_id}: Primary={raw.get('primary_score', 0)}, "
                        f"Sub={raw.get('sub_score', 0)}, "
                        f"Process={raw.get('process_score', 0)}, "
                        f"Evidence={raw.get('evidence', '')[:100]}...")
        lines.append("")
        
        lines.append("# Aggregated Scores")
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
  "top_strengths": [
    {"category": "<primary|sub|process>", "skill": "<skill_name>", "score": <float>, "reason": "<string>"}
  ],
  "top_weaknesses": [
    {"category": "<primary|sub|process>", "skill": "<skill_name>", "score": <float>, "reason": "<string>"}
  ],
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
        lines.append(f"Primary Scores: {pm01_final.get('scores_primary', {})}")
        lines.append(f"Sub Scores: {pm01_final.get('scores_sub', {})}")
        lines.append(f"Process Scores: {pm01_final.get('process', {})}")
        lines.append(f"Top Strengths: {pm01_final.get('top_strengths', [])}")
        lines.append(f"Top Weaknesses: {pm01_final.get('top_weaknesses', [])}")
        lines.append(f"Overall Summary: {pm01_final.get('overall_summary', '')}")
        lines.append("")
        
        # Get prompt from config sheet (can use promptPM5 or separate promptPM5Final)
        prompt_text = self.config.get('promptPM5Final', '').strip()
        if not prompt_text:
            # Fallback to promptPM5 if promptPM5Final not set
            prompt_text = self.config.get('promptPM5', '').strip()
        
        if prompt_text:
            lines.append("# Consistency Check Instructions")
            lines.append(prompt_text)
            lines.append("")
        else:
            raise ValueError("promptPM5 or promptPM5Final not configured in Config sheet. Required for STEP 4 (PM05 Final).")
        
        lines.append("# Required JSON Schema")
        lines.append("""{
  "consistency_score": <1-5>,
  "status": "<valid|caution|re-evaluate>",
  "issues": ["<issue1>", "<issue2>", ...],
  "summary": "<overall consistency summary>"
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
