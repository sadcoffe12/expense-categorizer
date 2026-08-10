"""
User interface utilities for expense categorization.
Handles interactive column selection and user prompts.
"""


def select_column(df, prompt: str, default_name: str, template_suggestion: str = None) -> str:
    """
    Interactively select a column from DataFrame.
    
    If a template suggestion is provided and valid, uses it without prompting.
    Otherwise, displays available columns and prompts user to select.
    
    Args:
        df: pandas DataFrame with available columns
        prompt: Description of what column is being selected
        default_name: Default column name if user presses Enter
        template_suggestion: Optional pre-suggested column name from template
        
    Returns:
        Selected column name
    """
    # If template provides a valid suggestion, use it
    if template_suggestion and template_suggestion in df.columns:
        print(f"[Template] Using column '{template_suggestion}' for {prompt}.")
        return template_suggestion

    # Otherwise, prompt user to select
    print(f"\n--- Column Selection: {prompt} ---")
    for i, col_name in enumerate(df.columns):
        print(f"  {i+1}. {col_name}")
    
    user_input = input(f"Select number or name (Enter for '{default_name}'): ").strip()

    if not user_input:
        return default_name
    
    # If user entered a number, use it as index
    if user_input.isdigit():
        idx = int(user_input) - 1
        if 0 <= idx < len(df.columns):
            return df.columns[idx]
        return default_name
    
    # Otherwise, treat as column name
    return user_input
