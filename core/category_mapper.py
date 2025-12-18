"""
Category mapper - Maps question sheet categories to official categories.
Shared by ScoringEngine and LLMService to ensure consistency.
"""

from typing import Optional


# Official PRIMARY categories (5 categories)
OFFICIAL_PRIMARY_CATEGORIES = [
    '問題理解',
    '論理思考',
    '仮説構築',
    'AI指示',
    'AI検証/優先順位判断'
]

# Official SUB categories (2 categories)
OFFICIAL_SUB_CATEGORIES = [
    '情報整理',
    '因果推論'
]

# Official PROCESS items (5 items)
OFFICIAL_PROCESS_ITEMS = [
    'clarity',
    'structure',
    'hypothesis',
    'prompt clarity',
    'consistency'
]


def map_to_official_category(category: str, category_type: str) -> Optional[str]:
    """
    Map question sheet category to official category.
    
    Args:
        category: Category from question sheet
        category_type: 'primary', 'sub', or 'process'
    
    Returns:
        Official category name or None if not mappable
    """
    if not category:
        return None
    
    category = category.strip()
    
    if category_type == 'primary':
        # Map variations to official PRIMARY categories
        mapping = {
            '問題理解': '問題理解',
            '論理思考': '論理思考',
            '論理構成': '論理思考',  # Map old name to new
            '仮説構築': '仮説構築',
            'AI指示': 'AI指示',
            'AI成果検証力': 'AI検証/優先順位判断',  # Map from question sheet
            '優先順位判断': 'AI検証/優先順位判断',  # Map from question sheet
        }
        return mapping.get(category, category if category in OFFICIAL_PRIMARY_CATEGORIES else None)
    
    elif category_type == 'sub':
        # Map variations to official SUB categories
        # Note: Official SUB categories are only 2: 情報整理, 因果推論
        # Other categories from question sheet will be mapped to the closest match
        mapping = {
            '情報整理': '情報整理',
            '因果推論': '因果推論',
            # Map question sheet categories to official categories
            # Since only 2 official categories exist, map related concepts:
            '前提設定': '因果推論',  # Related to logical reasoning
            '要件定義力': '情報整理',  # Related to information organization
            '品質チェック力': '因果推論',  # Related to logical verification
            '意思決定': '因果推論',  # Related to logical decision-making
        }
        return mapping.get(category, category if category in OFFICIAL_SUB_CATEGORIES else None)
    
    elif category_type == 'process':
        # Map variations to official PROCESS items
        mapping = {
            'clarity': 'clarity',
            'structure': 'structure',
            'hypothesis': 'hypothesis',
            'prompt clarity': 'prompt clarity',
            'prompt_clarity': 'prompt clarity',  # Map underscore variant from question sheet (Q4)
            'consistency': 'consistency',
            'quality_check': 'consistency',  # Map quality_check to consistency (Q5)
        }
        normalized = mapping.get(category.lower(), category.lower() if category.lower() in [item.lower() for item in OFFICIAL_PROCESS_ITEMS] else None)
        # Return with correct casing
        if normalized:
            for official in OFFICIAL_PROCESS_ITEMS:
                if official.lower() == normalized:
                    return official
        return normalized
    
    return None

