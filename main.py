# ==============================================================================
# SCRIPT: ADVANCED MATERIAL DESCRIPTION VALIDATOR
# Author: [Max Keller/DC/BDC]
# Date: [29.07.2025]
# Description: This script validates material descriptions ('Typkurzbezeichnung')
#              based on a multi-stage pipeline including character checks,
#              stop word analysis, dictionary lookups, and
#              de-compounding for compound words.
#              It updates two columns: a flag (1/"") and a reason text.
# ==============================================================================

import pandas as pd
from utils.io_utils import *
from utils.validator import TextValidator

# --- CONFIGURATION ---
# File and Column Names
SOURCE_FILENAME = 'data\input_data\TEST_Data_for_REGEX_V2.xlsx'
OUTPUT_FILENAME= 'data\output_data\output.xlsx'
SOURCE_COLUMN_NAME = 'Typkurzbezeichnung'
CONDITION_COLUMN = 'MArt'
TARGET_FLAG_COLUMN = 'B' # Will contain 1 or ""
TARGET_REASON_COLUMN = 'lsg'    # Will contain the detailed reason

# Conditional Values
CONDITION_VALUES = ['BREX', 'KAUF'] # Check if MArt is one of these

# Resource File Names
STOP_WORDS_FILENAME = 'data\dictionaries\german_stopwords.txt'
GERMAN_DICTIONARY_FILENAME = 'data\dictionaries\german_words.txt'
FRENCH_DICTIONARY_FILENAME = 'data\dictionaries\\french_words.txt'
ENGLISH_DICTIONARY_FILENAME = 'data\dictionaries\english_words.txt'
POSITIVE_WORDS_FILENAME = 'data\dictionaries\positive_words.txt'

def main():
    """Main function to orchestrate the entire data processing workflow."""
    process_name = "Advanced Material Master Data Validation"
    print(f"--- Starting: {process_name} ---")
    try:
        # Get absolute paths for resource files
        
        # Step 1: Load data and initialize validator
        df = load_data(SOURCE_FILENAME)

        validator = TextValidator(stopwords_path=STOP_WORDS_FILENAME,
                                  english_words_path= ENGLISH_DICTIONARY_FILENAME,
                                    german_words_path=GERMAN_DICTIONARY_FILENAME,
                                      french_words_path=FRENCH_DICTIONARY_FILENAME,
                                        positive_words_path=POSITIVE_WORDS_FILENAME  )
        
        # Check that all necessary columns exist
        required_cols = [SOURCE_COLUMN_NAME, CONDITION_COLUMN]
        if not all(col in df.columns for col in required_cols):
            raise ValueError(f"Required columns ({required_cols}) not found.")

        # Step 2: Apply the validation logic row-by-row
        print("Running conditional validation pipeline...")
        
        def apply_full_validation(row: pd.Series) -> pd.Series:
            """Applies validation logic to a single row and returns results for both target columns."""
            flag, reason = "", "" # Default values are "no error"
            
            # Conditional check: only run validation if MArt is one of the specified values
            if row[CONDITION_COLUMN] in CONDITION_VALUES:
                source_text = row[SOURCE_COLUMN_NAME]
                flag, reason = validator.validate(source_text)
            
            return pd.Series([flag, reason], index=[TARGET_FLAG_COLUMN, TARGET_REASON_COLUMN])

        # df.apply returns a new DataFrame with our two new columns
        results_df = df.apply(apply_full_validation, axis=1)

        # Step 3: Save the results back to the original Excel file
        save_data_to_new_excel(SOURCE_FILENAME,OUTPUT_FILENAME, results_df)


        print("\n" + "="*50)
        print("✅ SUCCESS: Workflow completed without errors.")
        print("="*50)

    except (FileNotFoundError, ValueError, PermissionError) as e:
        print(f"\n❌ ERROR: The script failed. Reason: {e}")
    except Exception as e:
    
        print(f"\n❌ An unexpected error occurred: {e}")

# ==============================================================================
# 4. SCRIPT ENTRY POINT
# ==============================================================================

if __name__ == "__main__":
    main()