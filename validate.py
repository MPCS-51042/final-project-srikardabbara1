import re

"""This module contains the validation functions for the Excel functions"""

def validate_input(input_type, value):
    """Seperate method to validate inputs"""
    if input_type in ['range', 'table_array', 'lookup_array', 'return_array']:
        if not re.match(r'^[A-Za-z]{1,3}\d+:[A-Za-z]{1,3}\d+$', value):
            return f"Your {input_type} is invalid. Please use the format: A1:A10."
        
    elif input_type in ['lookup_value', 'criteria', 'cell', 'value']:
        #these can be text or numbers, so mainly checking if empty
        if value.strip() == '':
            return f"Your {input_type} is empty. Please enter a valid value."
        
    elif input_type == 'col_index_num':
        if not value.isdigit() or int(value) < 1:
            return "Your column index number is invalid. Please enter a positive integer."
        
    elif input_type == 'index_num':
        if not value.isdigit() or int(value) < 1:
            return "Your index number is invalid. Please enter a positive integer."

    return ""  # Return an empty string if no errors