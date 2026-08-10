"""
Core categorization engine for expense transactions.
Applies rules to categorize transactions based on keyword matching.
"""

import pandas as pd
import config
from categorizer import text_normalizer
from categorizer import rules as rules_module


def categorize(df: pd.DataFrame, selected_template=None, col_identity=(None, None, None)) -> tuple:
    """
    Categorize transactions in DataFrame using loaded rules.
    
    Applies keyword matching to assign transaction type and category.
    Optionally renames transactions based on rules.
    
    Args:
        df: DataFrame containing transaction data
        selected_template: Optional template dict with column configurations and ordering
        col_identity: Tuple of (source_col, type_col, category_col) names
        
    Returns:
        Tuple of (categorized_df, source_col_name, category_col_name)
    """
    print("\n--- Categorize Records ---")
    
    src, typ, cat = col_identity

    # If column names not provided or missing, select them interactively
    if not src or src not in df.columns:
        print(f"Available columns: {list(df.columns)}")
        from categorizer import ui
        src = ui.select_column(df, "Description", "Descripcion", src)
        
    if not typ or typ not in df.columns:
        from categorizer import ui
        typ = ui.select_column(df, "Expense Type", "Tipo", typ)
        
    if not cat or cat not in df.columns:
        from categorizer import ui
        cat = ui.select_column(df, "Category", "Categoria", cat)

    # Ensure destination columns exist
    if typ not in df.columns:
        df[typ] = ""
    if cat not in df.columns:
        df[cat] = ""
    
    # Store original description before any replacements
    src_original = f"{src} Original"
    if src_original not in df.columns:
        df[src_original] = df[src].copy()
        
    # Ensure proper data types
    df[src] = df[src].astype(str)
    df[cat] = df[cat].astype(str)
    df[typ] = df[typ].astype(str)

    # Load rules and apply categorization
    loaded_rules = rules_module.load_rules()
    if not loaded_rules:
        print("Warning: No rules defined. Skipping automatic categorization.")
    else:
        categorized_count = 0
        
        def assign_category(description):
            nonlocal categorized_count
            desc_value = str(description) if pd.notna(description) else ""
            desc_lower = desc_value.lower()
            cat_v, typ_v = "", ""

            for keyword, t_val, c_val, new_desc in loaded_rules:
                if keyword in desc_lower:
                    categorized_count += 1
                    cat_v, typ_v = c_val, t_val
                    if new_desc:
                        desc_value = new_desc
                        desc_lower = desc_value.lower()
                        
            return typ_v, cat_v, desc_value

        results = df[src].apply(assign_category)
        results_df = pd.DataFrame(results.tolist(), columns=[typ, cat, src])
        df[typ], df[cat], df[src] = results_df[typ].values, results_df[cat].values, results_df[src].values
        print(f"Categorization complete. {categorized_count} matches found.")

    # Normalize column names (capitalize first letter, replace underscores)
    df.columns = [col.replace('_', ' ').capitalize() for col in df.columns]
    
    # Apply template-based column reordering if available
    if selected_template and 'ORDERED_COLS' in selected_template:
        raw_ordered_cols = selected_template['ORDERED_COLS'].replace("'", "").replace('"', "")
        ordered_cols = [col.strip() for col in raw_ordered_cols.split(',')]
        
        # Create flexible mapping: normalized_name -> original_column_in_df
        col_map = {}
        for df_col in df.columns:
            normalized_df_col = text_normalizer.normalize_col_name(df_col)
            col_map[normalized_df_col] = df_col
        
        # Find which columns from template exist in the DataFrame
        final_cols = []
        for template_col in ordered_cols:
            normalized_template_col = text_normalizer.normalize_col_name(template_col)
            if normalized_template_col in col_map:
                final_cols.append(col_map[normalized_template_col])
        
        # Reorder DataFrame with found columns
        if final_cols:
            df = df[final_cols]
    
    # Return the categorized DataFrame and the finalized column names
    src_final = src.replace('_', ' ').capitalize()
    cat_final = cat.replace('_', ' ').capitalize()
    
    return df, src_final, cat_final
