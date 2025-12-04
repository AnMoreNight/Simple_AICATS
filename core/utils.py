"""
Utility functions.
"""

from datetime import datetime
from typing import List, TypeVar

T = TypeVar('T')


def now() -> datetime:
    """Returns the current datetime."""
    return datetime.now()


def format_date(date: datetime, timezone: str = 'Asia/Tokyo') -> str:
    """Formats a date using the configured timezone."""
    return date.strftime('%Y-%m-%d %H:%M:%S')


def parse_date(value: str) -> datetime:
    """Parses a date string back into a datetime object."""
    if not value or not value.strip():
        return now()
    
    try:
        # Try parsing the format: YYYY-MM-DD HH:MM:SS
        return datetime.strptime(value.strip(), '%Y-%m-%d %H:%M:%S')
    except ValueError:
        # Fallback to now if parsing fails
        return now()


def chunk_array(items: List[T], size: int) -> List[List[T]]:
    """Breaks an array into smaller equally sized chunks."""
    result = []
    for i in range(0, len(items), size):
        result.append(items[i:i + size])
    return result


def parse_number(value: any, fallback: float = 0.0) -> float:
    """Converts arbitrary input into a number, using the fallback when conversion fails."""
    try:
        num = float(value)
        return num if num != float('inf') and num != float('-inf') else fallback
    except (ValueError, TypeError):
        return fallback


def safe_json_parse(json_str: str) -> dict | None:
    """Safely parses JSON content and returns None when parsing fails."""
    import json
    try:
        return json.loads(json_str)
    except (json.JSONDecodeError, TypeError):
        return None


def sanitize_answer(answer: str) -> str:
    """Sanitizes a free-form answer to the maximum length used in downstream prompts."""
    return answer.strip()[:400]


def create_run_id() -> str:
    """Generates a unique run identifier based on the current timestamp."""
    current = now()
    return f"RUN_{current.strftime('%Y%m%d_%H%M%S')}"


def is_empty(value: any) -> bool:
    """Tests whether a value is null, undefined, or blank when converted to a string."""
    if value is None:
        return True
    return str(value).strip() == ''


def flatten(arrays: List[List[T]]) -> List[T]:
    """Flattens a two-dimensional array one level deep."""
    result = []
    for arr in arrays:
        result.extend(arr)
    return result

