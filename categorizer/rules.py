"""
Rules management for expense categorization.
Handles loading and maintaining categorization rules from CSV.
"""

import os
import pandas as pd
import config


def ensure_rules_file_valid() -> bool:
    """
    Verify and recreate categorization_rules.csv if missing or empty.
    
    Returns:
        True if file was recreated, False if it already existed and was valid
    """
    if not os.path.exists(config.RULES_FILE) or os.path.getsize(config.RULES_FILE) == 0:
        df_empty = pd.DataFrame(columns=['keyword', 'type', 'category', 'new_description'])
        df_empty.to_csv(config.RULES_FILE, index=False, encoding="utf-8")
        print(f"ℹ️  File '{config.RULES_FILE}' created with empty columns.")
        return True  # File was recreated
    return False  # File already existed


def load_rules() -> list:
    """
    Load categorization rules from CSV file.
    
    Returns a list of tuples: (keyword, type, category, new_description)
    Empty strings in new_description are converted to None.
    Empty keywords are skipped.
    
    Returns:
        List of rule tuples, or empty list if file doesn't exist or is invalid
    """
    ensure_rules_file_valid()
    
    if not os.path.exists(config.RULES_FILE):
        df_empty = pd.DataFrame(columns=['keyword', 'type', 'category', 'new_description'])
        df_empty.to_csv(config.RULES_FILE, index=False, encoding="utf-8")
        return []
    
    try:
        # Load CSV and fill NaN with empty strings
        df_rules = pd.read_csv(config.RULES_FILE, encoding="utf-8").fillna("")
        
        rules = []
        for _, row in df_rules.iterrows():
            keyword = str(row['keyword']).lower().strip()
            
            # Skip rows with empty keywords
            if not keyword:
                continue
            
            # Convert empty/NaN new_description to None
            new_desc = str(row['new_description']).strip()
            val_new_desc = new_desc if new_desc not in ["", "nan", "None", "NaN"] else None
            
            rules.append((
                keyword,
                str(row['type']),
                str(row['category']),
                val_new_desc
            ))
        return rules
        
    except Exception as e:
        print(f"Error loading rules from CSV: {e}")
        return []


def save_rule(keyword: str, rule_type: str, category: str, new_description: str = None) -> bool:
    """
    Save a new categorization rule to CSV file.
    
    Args:
        keyword: The keyword to match (will be lowercased)
        rule_type: Type of expense (e.g., "Fijo", "Variable")
        category: Category name (e.g., "Comida", "Alquiler")
        new_description: Optional replacement description for the transaction
        
    Returns:
        True if successful, False otherwise
    """
    try:
        new_rule = pd.DataFrame([[keyword, rule_type, category, new_description or ""]],
                               columns=['keyword', 'type', 'category', 'new_description'])
        new_rule.to_csv(config.RULES_FILE, mode='a', header=False, index=False, encoding="utf-8")
        return True
    except Exception as e:
        print(f"Error saving rule to CSV: {e}")
        return False
