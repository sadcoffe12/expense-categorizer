"""
History management for uncategorized transactions.
Tracks patterns over time to suggest new categorization rules.
"""

import os
import pandas as pd
from datetime import datetime, timedelta
import config
from categorizer import text_normalizer


def update_history(new_list: list) -> pd.DataFrame:
    """
    Update the history file with new uncategorized transaction entries.
    
    Prevents duplicates by comparing normalized text and date_added.
    Automatically cleans entries older than MAX_HISTORY_AGE_DAYS.
    
    Args:
        new_list: List of transaction descriptions to add
        
    Returns:
        Updated history DataFrame
    """
    # Normalize entries with banking noise removal
    current_date = datetime.now().strftime("%Y-%m-%d")
    new_data = [
        {
            "original": x,
            "cleaned": text_normalizer.normalize_text(x, is_transaction=True),
            "date_added": current_date
        } for x in new_list
    ]
    df_new = pd.DataFrame(new_data)
    
    # Check if history file exists and is not empty
    file_is_empty = os.path.exists(config.HISTORY_FILE) and os.path.getsize(config.HISTORY_FILE) == 0
    
    if os.path.exists(config.HISTORY_FILE) and not file_is_empty:
        df_old = pd.read_csv(config.HISTORY_FILE)
        
        # Ensure date_added column exists (for backward compatibility)
        if 'date_added' not in df_old.columns:
            df_old['date_added'] = datetime.now().strftime("%Y-%m-%d")
        
        # Create unique keys to avoid duplicates
        df_old['_key'] = df_old['cleaned'] + '|' + df_old['date_added'].astype(str)
        df_new['_key'] = df_new['cleaned'] + '|' + df_new['date_added'].astype(str)
        
        # Identify duplicates
        new_keys = set(df_new['_key'])
        old_keys = set(df_old['_key'])
        duplicated_keys = new_keys & old_keys
        
        # Filter out duplicates from new entries
        df_new_filtered = df_new[~df_new['_key'].isin(duplicated_keys)].copy()
        
        # Remove temporary key column
        df_old = df_old.drop(columns=['_key'])
        df_new_filtered = df_new_filtered.drop(columns=['_key'])
        
        # Combine old and new (non-duplicate) entries
        if not df_new_filtered.empty:
            df_final = pd.concat([df_old, df_new_filtered], ignore_index=True)
            duplicated_count = len(duplicated_keys)
            if duplicated_count > 0:
                print(f"ℹ️  {duplicated_count} entry(ies) already existed in history. Not duplicated.")
        else:
            df_final = df_old
            print(f"ℹ️  All {len(new_list)} entries already existed in history. No duplicates added.")
    else:
        df_final = df_new
    
    # Clean old entries and save
    df_final = clean_old_entries(df_final)
    df_final.to_csv(config.HISTORY_FILE, index=False, encoding="utf-8")
    
    return df_final


def clean_old_entries(df_history: pd.DataFrame) -> pd.DataFrame:
    """
    Remove history entries older than MAX_HISTORY_AGE_DAYS.
    
    Args:
        df_history: History DataFrame to clean
        
    Returns:
        Cleaned DataFrame with old entries removed
    """
    if df_history.empty or 'date_added' not in df_history.columns:
        return df_history
    
    # Convert to datetime and filter
    df_history['date_added'] = pd.to_datetime(df_history['date_added'], errors='coerce')
    cutoff_date = datetime.now() - timedelta(days=config.MAX_HISTORY_AGE_DAYS)
    
    df_clean = df_history[df_history['date_added'] >= cutoff_date].copy()
    
    removed_count = len(df_history) - len(df_clean)
    if removed_count > 0:
        print(f"ℹ️  Removed {removed_count} entries older than {config.MAX_HISTORY_AGE_DAYS} days.")
    
    return df_clean
