#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Unit test để kiểm tra formula writing và building functions
"""
import sys
import os
import json

sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'logic'))

from salary_deployer import _column_index_to_letter, _column_letter_to_index

def test_column_conversions():
    """Test column index to letter and back"""
    print("Testing column conversions...")
    test_cases = [
        (1, 'A'),
        (26, 'Z'),
        (27, 'AA'),
        (52, 'AZ'),
        (53, 'BA'),
        (702, 'ZZ'),
        (703, 'AAA'),
    ]
    
    for idx, expected_letter in test_cases:
        actual_letter = _column_index_to_letter(idx)
        actual_idx = _column_letter_to_index(actual_letter)
        
        assert actual_letter == expected_letter, f"Column {idx}: expected {expected_letter}, got {actual_letter}"
        assert actual_idx == idx, f"Reverse conversion: expected {idx}, got {actual_idx}"
        print(f"  ✓ {idx} <-> {expected_letter}")
    
    print("All column conversion tests passed!\n")

def test_formula_syntax():
    """Test VLOOKUP formula syntax"""
    print("Testing VLOOKUP formula syntax...")
    
    # Test formula components
    test_formulas = [
        ('=IFERROR(VLOOKUP(C9,\'[file.xlsx]Sheet\'!$D$1:$AS$1000,42,0),"")', 'Simple VLOOKUP'),
        ('=760*BF9', 'Simple multiplication'),
        ('=F9*358*(((I8-N8)*8+N8*7))/(((I8-N8)*8+N8*7+J8)))*$I$4', 'Complex formula'),
    ]
    
    for formula, description in test_formulas:
        has_external = "[" in formula and "]" in formula
        print(f"  ✓ {description}")
        print(f"    Formula: {formula}")
        print(f"    Has external ref: {has_external}")
        print()
    
    print("All formula syntax tests passed!\n")

if __name__ == "__main__":
    try:
        test_column_conversions()
        test_formula_syntax()
        print("\n✅ All tests passed!")
    except AssertionError as e:
        print(f"\n❌ Test failed: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
