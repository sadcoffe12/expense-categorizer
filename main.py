"""
Expense Categorizer - Main CLI Entry Point

A command-line tool for categorizing expense transactions using customizable rules
and intelligent pattern learning across time periods.

This is a thin UI/CLI wrapper that orchestrates the core categorization modules.
All business logic is in the categorizer package.
"""

import os
import pandas as pd
import config
from categorizer import templates
from categorizer import categorization
from categorizer import learning_patterns
from categorizer import text_normalizer


def load_excel_file() -> tuple:
    """
    Interactively prompt user to select and load an Excel file.
    
    Returns:
        Tuple of (DataFrame, file_path)
    """
    while True:
        file_path = input("Enter the full path to your Excel file: ").strip()
        
        # Remove surrounding quotes if present
        if (file_path.startswith(("'", '"')) and file_path.endswith(("'", '"')) 
                and file_path[0] == file_path[-1]):
            file_path = file_path[1:-1]
        
        try:
            df = pd.read_excel(file_path)
            print("\n✅ Excel file loaded successfully!")
            return df, file_path
        except Exception as e:
            print(f"Error reading file: {e}")


def save_excel_file(df: pd.DataFrame, original_file_path: str) -> None:
    """
    Save categorized DataFrame to a new Excel file.
    
    Creates a new file with "_modified" suffix in the same directory as the original.
    Cleans description whitespace before saving.
    
    Args:
        df: DataFrame to save
        original_file_path: Path of the original file (used to determine output location)
    """
    try:
        # Clean descriptions before saving
        text_normalizer.clean_descriptions(df)
        
        # Generate output file path
        new_file_path = os.path.splitext(original_file_path)[0] + "_modified.xlsx"
        
        # Save to Excel
        df.to_excel(new_file_path, index=False)
        print(f"\n✅ File saved successfully as '{new_file_path}'!")
        
    except Exception as e:
        print(f"Error saving file: {e}")


def show_main_menu() -> None:
    """Display the main menu options."""
    print("\n--- Main Menu ---")
    print("1. Use a template (Categorize + Clean)")
    print("2. Categorize records (Manual)")
    print("S. Save and exit")
    print("F. Exit without saving")


def main() -> None:
    """
    Main CLI loop for the Expense Categorizer application.
    
    Allows user to:
    1. Load an Excel file
    2. Apply templates and categorize transactions
    3. Run the learning system to suggest new rules
    4. Save categorized results
    """
    # Ensure configuration files exist
    templates.ensure_templates_file_valid()
    
    # Load Excel file
    print("\n=== Expense Categorizer ===")
    df, file_path = load_excel_file()
    
    # Main loop
    while True:
        show_main_menu()
        choice = input("\nChoose an option: ").upper()
        
        if choice == '1':
            # Apply template and categorize
            df, template, cols_info = templates.apply_format(df, file_path)
            
            if df is not None and not df.empty:
                # Run categorization with template info
                df, final_src, final_cat = categorization.categorize(df, template, cols_info)
                
                # Run learning system to suggest new rules
                learning_patterns.learn_and_suggest(df, final_src, final_cat)
                
        elif choice == '2':
            # Categorize without template (manual column selection)
            df, final_src, final_cat = categorization.categorize(df, None)
            
            # Run learning system
            learning_patterns.learn_and_suggest(df, final_src, final_cat)

        elif choice == 'S':
            # Save and exit
            save_excel_file(df, file_path)
            print("\nGoodbye!")
            break
            
        elif choice == 'F':
            # Exit without saving
            print("Exiting without saving changes.")
            break
        
        else:
            print("Invalid option. Please try again.")


if __name__ == "__main__":
    main()
