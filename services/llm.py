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
    
    def run_pm01_diagnosis(
        self,
        respondent: Dict[str, Any],
        questions: List[Dict[str, Any]],
        attempt: int
    ) -> Optional[Dict[str, Any]]:
        """
        Run PM01 (Primary Diagnosis) LLM evaluation.
        
        Returns structured JSON with per-question scores.
        """
        prompt = self._build_pm01_prompt(respondent, questions)
        print(f"  PM01 Prompt length: {len(prompt)} characters")
        
        response = self._invoke_llm(prompt, attempt, "You are an expert evaluator. Output ONLY valid JSON.")
        if not response:
            return None
        
        parsed = self.json_parser.parse_pm01_response(response)
        return parsed
    
    def run_pm05_validation(
        self,
        respondent: Dict[str, Any],
        questions: List[Dict[str, Any]],
        pm01_result: Dict[str, Any],
        attempt: int
    ) -> Optional[Dict[str, Any]]:
        """
        Run PM05 (Secondary Diagnosis) - Reverse scoring validation.
        
        Returns structured JSON with reverse-scored results.
        """
        prompt = self._build_pm05_prompt(respondent, questions, pm01_result)
        print(f"  PM05 Prompt length: {len(prompt)} characters")
        
        response = self._invoke_llm(prompt, attempt, "You are a validation evaluator. Output ONLY valid JSON.")
        if not response:
            return None
        
        parsed = self.json_parser.parse_pm05_response(response)
        return parsed
    
    def _build_pm01_prompt(
        self,
        respondent: Dict[str, Any],
        questions: List[Dict[str, Any]]
    ) -> str:
        """Build PM01 prompt for primary diagnosis."""
        lines = []
        
        lines.append("# Respondent Information")
        lines.append(f"ID: {respondent['id']}")
        lines.append(f"Name: {respondent['name']}")
        lines.append("")
        
        lines.append("# Questions and Answers (Q1-Q6)")
        for i, question in enumerate(questions):
            if question['number'] < 1 or question['number'] > 6:
                continue
            
            answer = respondent['answers'][i] if i < len(respondent['answers']) else ""
            question_id = f"Q{question['number']}"
            
            lines.append(f"## {question_id}")
            lines.append(f"Question: {question['questionText']}")
            lines.append(f"Answer: {answer or '(無回答)'}")
            lines.append(f"Primary Skill: {question.get('primary_skill', '')}")
            lines.append(f"Sub Skill: {question.get('sub_skill', '')}")
            lines.append(f"Process Skill: {question.get('process_skill', '')}")
            lines.append("")
        
        lines.append("# Scoring Instructions")
        lines.append("Evaluate each question (Q1-Q6) and provide:")
        lines.append("- primary_score: Score for primary skill (0-5)")
        lines.append("- sub_score: Score for sub skill (0-5)")
        lines.append("- process_score: Score for process skill (0-5)")
        lines.append("- aes_score: AI Evaluation Score (0-5)")
        lines.append("- comment: Evaluation comment")
        lines.append("")
        
        lines.append("# Required JSON Schema")
        lines.append("""{
  "per_question": {
    "Q1": {
      "primary_score": <0-5>,
      "sub_score": <0-5>,
      "process_score": <0-5>,
      "aes_score": <0-5>,
      "comment": "<string>"
    },
    "Q2": {...},
    "Q3": {...},
    "Q4": {...},
    "Q5": {...},
    "Q6": {...}
  },
  "increase_factor": <optional float>,
  "decrease_factor": <optional float>
}""")
        
        prompt_text = self.config.get('promptPM1', '').strip()
        if prompt_text:
            lines.append("")
            lines.append("# Additional Instructions")
            lines.append(prompt_text)
        
        return "\n".join(lines)
    
    def _build_pm05_prompt(
        self,
        respondent: Dict[str, Any],
        questions: List[Dict[str, Any]],
        pm01_result: Dict[str, Any]
    ) -> str:
        """Build PM05 prompt for reverse scoring validation."""
        lines = []
        
        lines.append("# PM05 Secondary Diagnosis - Reverse Scoring Validation")
        lines.append(f"Respondent ID: {respondent['id']}")
        lines.append("")
        
        lines.append("# PM01 Results (for comparison)")
        pm01_per_q = pm01_result.get('per_question', {})
        for i in range(1, 7):
            q_id = f"Q{i}"
            q_data = pm01_per_q.get(q_id, {})
            lines.append(f"{q_id}: Primary={q_data.get('primary_score', 0)}, "
                        f"Sub={q_data.get('sub_score', 0)}, "
                        f"Process={q_data.get('process_score', 0)}, "
                        f"AES={q_data.get('aes_score', 0)}")
        lines.append("")
        
        lines.append("# Questions and Answers (Q1-Q6)")
        for i, question in enumerate(questions):
            if question['number'] < 1 or question['number'] > 6:
                continue
            
            answer = respondent['answers'][i] if i < len(respondent['answers']) else ""
            question_id = f"Q{question['number']}"
            
            lines.append(f"## {question_id}")
            lines.append(f"Question: {question['questionText']}")
            lines.append(f"Answer: {answer or '(無回答)'}")
            lines.append("")
        
        lines.append("# PM05 Instructions")
        lines.append("Re-evaluate each question using REVERSE scoring logic.")
        lines.append("Compare your reverse-scored results with PM01 results.")
        lines.append("Detect any contradictions or inconsistencies.")
        lines.append("")
        
        lines.append("# Required JSON Schema")
        lines.append("""{
  "reverse_scores": {
    "Q1": {
      "total_score": <0-5>,
      "comment": "<string>"
    },
    "Q2": {...},
    "Q3": {...},
    "Q4": {...},
    "Q5": {...},
    "Q6": {...}
  },
  "validation_comment": "<string>"
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
