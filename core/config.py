"""
Configuration management - reads from Google Sheets Config sheet.
"""

import os
from typing import Dict, Any, Optional


class Config:
    """Configuration manager that reads from Google Sheets."""
    
    _cache: Optional[Dict[str, Any]] = None
    _sheets_service = None  # Will be set after SheetsService is initialized
    
    @classmethod
    def set_sheets_service(cls, sheets_service):
        """Sets the sheets service for reading config. Called from main.py after SheetsService is created."""
        cls._sheets_service = sheets_service
    
    @classmethod
    def get_config(cls, sheets_service) -> Dict[str, Any]:
        """Retrieves the configuration map from the Config sheet."""
        if cls._cache:
            return cls._cache
        
        if not sheets_service:
            raise ValueError("SheetsService is required to read configuration from Config sheet.")
        
        cls._sheets_service = sheets_service
        
        # Read from Config sheet
        config_sheet = sheets_service._get_sheet('Config')
        if not config_sheet:
            raise ValueError("Config sheet not found in spreadsheet. Please create a Config sheet.")
        
        values = config_sheet.get_all_values()
        if len(values) <= 1:
            raise ValueError("Config sheet is empty. Please add configuration key-value pairs.")
        
        # Parse config values
        config_map = {}
        for row in values[1:]:  # Skip header
            if len(row) >= 2:
                key = str(row[0]).strip()
                value = str(row[1]).strip()
                if key:
                    config_map[key] = value
        
        # Build config - require all values from sheet
        def get_number(key: str) -> int:
            raw = config_map.get(key)
            if not raw:
                raise ValueError(f"Required configuration key '{key}' not found in Config sheet.")
            try:
                return int(float(raw))
            except (ValueError, TypeError):
                raise ValueError(f"Invalid number value for '{key}' in Config sheet: {raw}")
        
        def get_string(key: str, required: bool = True) -> str:
            raw = config_map.get(key)
            if not raw and required:
                raise ValueError(f"Required configuration key '{key}' not found in Config sheet.")
            return raw if raw else ''
        
        config = {
            'respondentsSheet': get_string('respondentsSheet'),
            'validationLogSheet': get_string('validationLogSheet'),
            'errorLogSheet': get_string('errorLogSheet'),
            'llmProvider': get_string('llmProvider'),
            'llmApiUrl': get_string('llmApiUrl'),
            'llmApiKey': get_string('llmApiKey'),
            'llmModel': get_string('llmModel'),
            'maxRetries': get_number('maxRetries'),
            'promptDiagnosis': get_string('promptDiagnosis', required=False),
            'questionSheet': get_string('questionSheet'),
        }
        
        if not config['llmApiKey']:
            raise ValueError("API key not configured. Add llmApiKey to the Config sheet in your Google Spreadsheet.")
        
        cls._cache = config
        return config
    
    @classmethod
    def get_sheets_service(cls):
        """Gets the SheetsService instance used for reading config."""
        if not cls._sheets_service:
            raise ValueError("SheetsService not initialized. Call get_config() with a SheetsService first.")
        return cls._sheets_service
    
    @classmethod
    def clear_cache(cls):
        """Clears the in-memory configuration cache."""
        cls._cache = None

