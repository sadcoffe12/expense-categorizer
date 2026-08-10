"""
Learning system for expense categorization.
Analyzes patterns across time periods and suggests new rules.
"""

import os
import pandas as pd
from datetime import datetime, timedelta
from difflib import SequenceMatcher
import config
from categorizer import text_normalizer
from categorizer import rules as rules_module
from categorizer import history as history_module


def learn_and_suggest(df: pd.DataFrame, source_col: str, category_col: str) -> None:
    """
    Analyze uncategorized transactions and suggest new categorization rules.
    
    Identifies patterns in both current session and historical data across
    multiple time periods (this session, last month, last 3 months, last 6 months).
    
    For each pattern detected meeting threshold triggers, prompts user to:
    1. Accept AI suggestion (if confident match found)
    2. Manually enter categorization
    3. Skip the pattern
    4. Optionally rename the transaction
    
    Saves accepted rules and updates history with new uncategorized entries.
    
    Args:
        df: DataFrame containing categorized transactions
        source_col: Column name with transaction descriptions
        category_col: Column name with assigned categories
    """
    # Verify columns exist
    if source_col not in df.columns or category_col not in df.columns:
        print(f"Error: Columns '{source_col}' or '{category_col}' not found.")
        print(f"Available: {list(df.columns)}")
        return
    
    # Find uncategorized transactions in this session
    mask_empty = df[category_col].isin(["", "nan", "None", None])
    unassigned = df[mask_empty][source_col].tolist()
    
    if not unassigned:
        print("✅ All entries have been categorized.")
        return

    # Analyze patterns across time periods
    patterns_analysis = analyze_patterns_by_period(unassigned)
    
    # Combine all patterns from all time periods
    all_patterns = {
        **patterns_analysis['this_session'],
        **patterns_analysis['last_month'],
        **patterns_analysis['last_3_months'],
        **patterns_analysis['last_6_months']
    }
    
    if not all_patterns:
        print(f"\nℹ️  No repeated patterns found in {len(unassigned)} uncategorized entries")
        print("   (Requires at least 3 repetitions this session, last month, or last 3 months)")
        history_module.update_history(unassigned)
        return

    # Load existing rules for AI suggestions
    loaded_rules = rules_module.load_rules()
    
    print("\n--- 🧠 Recurring Expense Analyzer (Temporal Analysis) ---")
    print(f"Found {len(all_patterns)} repeated patterns across time periods.\n")
    
    processed_patterns = set()
    total_patterns = len(all_patterns)
    current_pattern_num = 0
    
    for pattern in all_patterns.keys():
        if pattern in processed_patterns:
            continue
            
        processed_patterns.add(pattern)
        current_pattern_num += 1
        
        print(f"\n📋 Suggestion ({current_pattern_num}/{total_patterns})")
        
        # Determine which time periods this pattern appeared in
        period_labels = []
        if pattern in patterns_analysis['this_session']:
            count = patterns_analysis['this_session'][pattern]
            period_labels.append(f"This session ({count}x)")
        if pattern in patterns_analysis['last_month']:
            count = patterns_analysis['last_month'][pattern]
            period_labels.append(f"Last month ({count}x)")
        if pattern in patterns_analysis['last_3_months']:
            count = patterns_analysis['last_3_months'][pattern]
            period_labels.append(f"Last 3 months ({count}x)")
        if pattern in patterns_analysis['last_6_months']:
            count = patterns_analysis['last_6_months'][pattern]
            period_labels.append(f"Last 6 months ({count}x)")
        
        periods_str = " | ".join(period_labels)
        
        print(f"{'='*70}")
        print(f"Pattern: '{pattern}'")
        print(f"Time periods: {periods_str}")
        
        # Try to guess category from existing rules
        sug_type, sug_cat = guess_category(pattern, loaded_rules)
        
        # User decision flow
        if sug_type:
            print(f"🤖 AI Suggestion: Type={sug_type}, Category={sug_cat}")
            confirm = input("Accept this category? (y/n/skip): ").lower().strip()
        else:
            confirm = 'n'

        if confirm == 'y':
            tipo, cat = sug_type, sug_cat
        elif confirm == 'n':
            print("Enter manual categorization:")
            tipo = input("  Type (Fijo/Variable): ").strip()
            cat = input("  Category: ").strip()
            if not tipo or not cat:
                print("  ⚠️  Type/Category empty. Skipping this pattern.")
                continue
        else:
            print("  ➜ Skipped.")
            continue

        # Allow renaming the transaction description
        new_desc = input(f"  Rename entry [Enter for '{pattern}']: ").strip()
        final_desc = new_desc if new_desc else pattern
        
        # Save the new rule
        if rules_module.save_rule(pattern, tipo, cat, final_desc):
            print("✅ Rule saved.")
    
    # Update history with this session's uncategorized entries
    history_module.update_history(unassigned)


def analyze_patterns_by_period(current_unassigned: list) -> dict:
    """
    Analyze transaction patterns across multiple time periods.
    
    Checks for repeated patterns in:
    - This session (requires TRIGGER_THIS_SESSION repetitions)
    - Last month (requires TRIGGER_LAST_MONTH repetitions)
    - Last 3 months (requires TRIGGER_LAST_3_MONTHS repetitions)
    - Last 6 months (requires TRIGGER_LAST_6_MONTHS repetitions)
    
    Args:
        current_unassigned: List of uncategorized transaction descriptions from current session
        
    Returns:
        Dictionary with patterns organized by time period:
        {
            'this_session': {pattern: count, ...},
            'last_month': {pattern: count, ...},
            'last_3_months': {pattern: count, ...},
            'last_6_months': {pattern: count, ...}
        }
    """
    
    if not os.path.exists(config.HISTORY_FILE) or os.path.getsize(config.HISTORY_FILE) == 0:
        # No history file, only analyze this session
        new_unassigned = [
            {"cleaned": text_normalizer.normalize_text(x, is_transaction=True)}
            for x in current_unassigned
        ]
        df_new = pd.DataFrame(new_unassigned)
        counts = df_new['cleaned'].value_counts()
        return {
            'this_session': counts[counts >= config.TRIGGER_THIS_SESSION].to_dict(),
            'last_month': {},
            'last_3_months': {},
            'last_6_months': {}
        }
    
    try:
        df_history = pd.read_csv(config.HISTORY_FILE)
        
        # Ensure date_added column exists
        if 'date_added' not in df_history.columns:
            df_history['date_added'] = datetime.now().strftime("%Y-%m-%d")
        
        df_history['date_added'] = pd.to_datetime(df_history['date_added'], errors='coerce')
        now = datetime.now()
        
        # 1. Patterns in THIS SESSION
        new_unassigned = [
            {"cleaned": text_normalizer.normalize_text(x, is_transaction=True)}
            for x in current_unassigned
        ]
        df_new = pd.DataFrame(new_unassigned)
        counts_this_session = df_new['cleaned'].value_counts()
        patterns_this_session = counts_this_session[
            counts_this_session >= config.TRIGGER_THIS_SESSION
        ].to_dict()
        
        # 2. Patterns in LAST MONTH
        one_month_ago = now - timedelta(days=30)
        df_last_month = df_history[df_history['date_added'] >= one_month_ago]
        counts_month = df_last_month['cleaned'].value_counts()
        patterns_last_month = counts_month[
            counts_month >= config.TRIGGER_LAST_MONTH
        ].to_dict()
        
        # 3. Patterns in LAST 3 MONTHS
        three_months_ago = now - timedelta(days=90)
        df_last_3m = df_history[df_history['date_added'] >= three_months_ago]
        counts_3m = df_last_3m['cleaned'].value_counts()
        patterns_last_3m = counts_3m[
            counts_3m >= config.TRIGGER_LAST_3_MONTHS
        ].to_dict()
        
        # 4. Patterns in LAST 6 MONTHS
        six_months_ago = now - timedelta(days=180)
        df_last_6m = df_history[df_history['date_added'] >= six_months_ago]
        counts_6m = df_last_6m['cleaned'].value_counts()
        patterns_last_6m = counts_6m[
            counts_6m >= config.TRIGGER_LAST_6_MONTHS
        ].to_dict()
        
        return {
            'this_session': patterns_this_session,
            'last_month': patterns_last_month,
            'last_3_months': patterns_last_3m,
            'last_6_months': patterns_last_6m
        }
        
    except Exception as e:
        print(f"Error analyzing historical patterns: {e}")
        return {
            'this_session': {},
            'last_month': {},
            'last_3_months': {},
            'last_6_months': {}
        }


def guess_category(cleaned_desc: str, rules: list) -> tuple:
    """
    Attempt to predict category for a transaction using similarity matching.
    
    Compares the cleaned description against existing rule keywords and returns
    the best match if confidence exceeds MIN_CONFIDENCE_FOR_SUGGESTION.
    
    Args:
        cleaned_desc: Normalized transaction description
        rules: List of rule tuples (keyword, type, category, new_description)
        
    Returns:
        Tuple of (type, category) if confident match found, else (None, None)
    """
    best_score = 0
    best_match = (None, None)  # (Type, Category)
    
    for keyword, t_val, c_val, _ in rules:
        # Compare using sequence matching
        score = SequenceMatcher(None, cleaned_desc, keyword).ratio()
        if score > best_score:
            best_score = score
            best_match = (t_val, c_val)
    
    # Only return suggestion if confidence is above threshold
    return best_match if best_score > config.MIN_CONFIDENCE_FOR_SUGGESTION else (None, None)
