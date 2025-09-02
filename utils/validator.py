
import pandas as pd
import os
import re
import openpyxl
import shutil
from openpyxl.utils.dataframe import dataframe_to_rows
from typing import List, Any, Callable, Set, Tuple
from utils.io_utils import *

class TextValidator:
    """A class to encapsulate the entire validation logic and its resources."""
    
    def __init__(self, stopwords_path: str, german_words_path: str, french_words_path: str, english_words_path: str, positive_words_path: str):
        print("Initializing validator and loading resources...")
        # Load resources into memory once
        self.stop_words = load_text_file_to_set(stopwords_path)
        self.german_dictionary = load_text_file_to_set(german_words_path)
        
        self.french_dictionary = load_text_file_to_set(french_words_path)

        self.english_dictionary = load_text_file_to_set(english_words_path)
        
        self.positive_words = load_text_file_to_set(positive_words_path)


    def _stage1_invalid_char_check(self, text: str) -> Tuple[bool, str]:
        """Stage 1: Checks for forbidden characters."""
        pattern = r"[^A-Z0-9*.,\"%/ :;=<>()&Ä´`_Ø+ \r\n-]"
        if re.search(pattern, text, re.IGNORECASE):
            return True, "Contains Invalid Characters"
        return False, ""

        
    def _stage2_word_analysis(self, text: str) -> Tuple[bool, str]:
            """Stage 2: Tokenizes text and runs checks on each token."""

            # Check 0: Extract tokens of 3 or more letters
            tokens = re.findall(r'[A-Za-z]{' + str(3) + r',}', text)
            
            for token in tokens:
                token_lower = token.lower()
                
                # Check 1: Stop words
                if token_lower in self.stop_words:
                    return True, f"Contains Stop Word: {token}"
                
                #Check 2: English Words
                if token_lower in self.english_dictionary:
                    continue
                
                #Check 3: Positive list
                if token_lower in self.positive_words:
                    continue
                
                # Check 4: Direct dictionary match
                # Check 4.1: Direct German dictionary match
                if token_lower in self.german_dictionary:
                    return True, f"Contains German Dictionary Word: {token}"
                
                # Check 4.2: Direct French dictionary match
                if token_lower in self.french_dictionary:
                    return True, f"Contains French Dictionary Word: {token}"
                
                #check 5: Compound word check DE and FR 
                for word in self.german_dictionary:
                    if len(word)>=5 and word in token_lower:
                        return True, f"Contains compound partial german word: {word} for Token: {token}"
                for word in self.french_dictionary:
                    if len(word)>=5 and word in token_lower:
                        return True, f"Contains compound partial french word: {word} for Token: {token}"

            
            return False, ""


    def validate(self, text: Any) -> Tuple[int, str]:
        """
        Runs the full validation pipeline on a single piece of text.
        Returns a tuple: (flag, reason). e.g., (1, "Contains Stop Word: FUER")
        """
        if not isinstance(text, str) or not text.strip():
            return "", "" # Not an error, just empty

        # Run Stage 1
        is_error, reason = self._stage1_invalid_char_check(text)
        if is_error:
            return 1, reason

        # Run Stage 2
        is_error, reason = self._stage2_word_analysis(text)
        if is_error:
            return 1, reason

        # If all checks pass
        return "", ""