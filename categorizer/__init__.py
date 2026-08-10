"""
Expense Categorizer - Modular Python package for automatic transaction categorization.

This package provides the core categorization engine for expense management,
organized into focused modules:

- text_processor: Text normalization and cleaning
- rules: Rule management and loading
- categorization: Core categorization engine
- learning: Pattern analysis and rule suggestion
- history: History tracking and management
- templates: Excel template handling and formatting
- ui: User interaction utilities
"""

# Module imports for convenience (optional - users can import directly if preferred)
from categorizer import (
    learning_patterns,
    rules,
    categorization,
    history,
    templates,
    text_normalizer,
    ui
)

__version__ = "2.0.0"
__all__ = [
    'text_normalizer',
    'rules',
    'categorization',
    'learning_patterns',
    'history',
    'templates',
    'ui'
]
