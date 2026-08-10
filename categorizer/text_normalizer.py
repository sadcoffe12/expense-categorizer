"""
Text processing utilities for expense categorization.
Handles normalization, cleaning, and text comparison.
"""

import unicodedata
import re


def normalize_text(text: str, is_transaction: bool = False) -> str:
    """
    Normalize text by removing accents, converting to lowercase, and cleaning up extra spaces.
    
    If is_transaction=True, also removes banking-specific noise (card numbers, transaction IDs).
    
    Args:
        text: Text to normalize
        is_transaction: If True, removes banking noise patterns
        
    Returns:
        Normalized text string
    """
    if not isinstance(text, str):
        return str(text) if text is not None else ""

    # 1. Convert to lowercase and remove accents
    text = text.lower().strip()
    text = ''.join(c for c in unicodedata.normalize('NFD', text)
                  if unicodedata.category(c) != 'Mn')

    # 2. Remove banking noise (only for transaction descriptions)
    if is_transaction:
        # Remove "tarj nro. 1234" or "tarjeta 1234" patterns
        text = re.sub(r'tarj\s?nro\.?\s?\d+', '', text)
        # Remove numbers with 5+ digits (transaction IDs)
        text = re.sub(r'\d{5,}', '', text)

    # 3. Collapse multiple spaces into single space
    text = re.sub(r'\s+', ' ', text).strip()
    
    return text


def normalize_col_name(col_name: str) -> str:
    """
    Normalize column name for flexible matching.
    Used for comparing column names, not modifying the DataFrame.
    
    Removes accents, converts to lowercase, and normalizes spaces.
    
    Args:
        col_name: Column name to normalize
        
    Returns:
        Normalized column name
    """
    if not isinstance(col_name, str):
        return ""
    
    normalized = col_name.lower().strip()
    normalized = ''.join(c for c in unicodedata.normalize('NFD', normalized)
                        if unicodedata.category(c) != 'Mn')
    normalized = re.sub(r'\s+', ' ', normalized)
    
    return normalized


def clean_descriptions(df) -> None:
    """
    Clean whitespace in description column (in-place modification).
    
    Removes leading/trailing whitespace and collapses internal duplicate spaces.
    Modifies the DataFrame column directly.
    
    Args:
        df: pandas DataFrame to clean
        
    Returns:
        None (modifies df in-place)
    """
    # Find description column (case-insensitive)
    desc_col = None
    for col in df.columns:
        if col.lower() in ['descripcion', 'description']:
            desc_col = col
            break
    
    if desc_col:
        # Remove leading/trailing whitespace
        df[desc_col] = df[desc_col].astype(str).str.strip()
        # Collapse internal duplicate spaces
        df[desc_col] = df[desc_col].str.replace(r'\s+', ' ', regex=True)
