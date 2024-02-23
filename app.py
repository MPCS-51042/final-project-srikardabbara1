from validate import validate_input
from flask import Flask, request, render_template_string, jsonify


app = Flask(__name__)

#HTML template for the home page
HOME_PAGE_TEMPLATE = """
<!doctype html>
<html>
<head><title>Excel Formula Generator</title></head>
<body>
    <h2>Select a Formula</h2>
    <form method="post" action="/generate">
        <select name="formula">
            <option value="SUM">SUM</option>
            <option value="IFERROR">IFERROR</option>
            <option value="COUNTIF">COUNTIF</option> 
            <option value="VLOOKUP">VLOOKUP</option> 
            <option value="INDEXMATCH">INDEXMATCH</option> 
            <option value="XLOOKUP">XLOOKUP</option> 
            <option value="CONCATENATE">CONCATENATE</option> 
            <option value="CHOOSE">CHOOSE</option> 
            <option value="SUBSITUTE">SUBSITUTE</option> 
            <option value="MINIF">MINIF</option> 
            <option value="MAXIF">MAXIF</option> 
        </select>
        <input type="submit" value="Choose">
    </form>
</body>
</html>
"""


#template to prompt for SUM formula inputs
SUM_TEMPLATE = """
<!doctype html>
<html>
<head><title>Sum Formula Inputs</title></head>
<body>
    <a href="/">Back to Home</a>
    <form method="post" action="/result">
        <input type="hidden" name="formula" value="SUM">
        Please enter the cell range for your data that you would like to sum: <br>
        <input type="text" name="range"><br>
        <input type="submit" value="Generate">
    </form>
</body>
</html>
"""

#template to prompt for IFERROR formula inputs
IFERROR_TEMPLATE = """
<!doctype html>
<html>
<head><title>IFERROR Formula Inputs</title></head>
<body>
    <a href="/">Back to Home</a>
    <form method="post" action="/result">
        <input type="hidden" name="formula" value="IFERROR">
        Please enter the value you'd like to change: <br>
        <input type="text" name="value"><br>
        Please enter the value you'd like to appear instead: <br>
        <input type="text" name="replacement"><br>
        <input type="submit" value="Generate">
    </form>
</body>
</html>
"""

#template to prompt for COUNTIF formula inputs
COUNTIF_TEMPLATE = """
<!doctype html>
<html>
<head><title>COUNTIF Formula Inputs</title></head>
<body>
    <a href="/">Back to Home</a>
    <form method="post" action="/result">
        <input type="hidden" name="formula" value="COUNTIF">
        <label for="range">Enter the range of cells to count from (e.g., A1:A10):</label><br>
        <input type="text" id="range" name="range"><br>
        <label for="criteria">Enter the criteria for counting (e.g., ">20", "=Done"):</label><br>
        <input type="text" id="criteria" name="criteria"><br>
        <input type="submit" value="Generate">
    </form>
</body>
</html>
"""

#VLOOKUP Template
VLOOKUP_TEMPLATE = """
<!doctype html>
<html>
<head><title>VLOOKUP Formula Inputs</title></head>
<body>
    <a href="/">Back to Home</a>
    <form method="post" action="/result">
        <input type="hidden" name="formula" value="VLOOKUP">
        Please enter the lookup value (the value you want to search for): <br>
        <input type="text" name="lookup_value"><br>
        Please enter the table range (where to look for the value): <br>
        <input type="text" name="table_array"><br>
        Please enter the column number in the range containing the return value: <br>
        <input type="text" name="col_index_num"><br>
        Should the match be exact or approximate? (Enter TRUE for approximate, FALSE for exact): <br>
        <input type="text" name="range_lookup"><br>
        <input type="submit" value="Generate">
    </form>
</body>
</html>
"""

#INDEX/MATCH Template
INDEX_MATCH_TEMPLATE = """
<!doctype html>
<html>
<head><title>INDEX/MATCH Formula Inputs</title></head>
<body>
    <a href="/">Back to Home</a>
    <form method="post" action="/result">
        <input type="hidden" name="formula" value="INDEXMATCH">
        Please enter the range that contains the return values (INDEX Range): <br>
        <input type="text" name="index_range"><br>
        Please enter the lookup value (the value you want to match): <br>
        <input type="text" name="lookup_value"><br>
        Please enter the range to search for the lookup value (MATCH Range): <br>
        <input type="text" name="match_range"><br>
        <input type="submit" value="Generate">
    </form>
</body>
</html>
"""

#CONCATENATE Template
CONCATENATE_TEMPLATE = """
<!doctype html>
<html>
<head><title>CONCATENATE Formula Inputs</title></head>
<body>
    <a href="/">Back to Home</a>
    <form method="post" action="/result">
        <input type="hidden" name="formula" value="CONCATENATE">
        Please enter the first value or string you want to concatenate: <br>
        <input type="text" name="first_value"><br>
        Please enter the second value or string you want to concatenate: <br>
        <input type="text" name="second_value"><br>
        <small>You can add more values by separating them with commas in the second input box.</small><br>
        <input type="submit" value="Generate">
    </form>
</body>
</html>
"""

#CHOOSE Template
CHOOSE_TEMPLATE = """
<!doctype html>
<html>
<head><title>CHOOSE Formula Inputs</title></head>
<body>
    <a href="/">Back to Home</a>
    <form method="post" action="/result">
        <input type="hidden" name="formula" value="CHOOSE">
        Please enter the index number (the position of the value to return): <br>
        <input type="text" name="index_num"><br>
        Please enter the list of values (separate each value with a comma): <br>
        <input type="text" name="values"><br>
        <input type="submit" value="Generate">
    </form>
</body>
</html>
"""

#SUBSTITUTE Template
SUBSTITUTE_TEMPLATE = """
<!doctype html>
<html>
<head><title>SUBSTITUTE Formula Inputs</title></head>
<body>
    <a href="/">Back to Home</a>
    <form method="post" action="/result">
        <input type="hidden" name="formula" value="SUBSTITUTE">
        Please enter the text string or cell reference: <br>
        <input type="text" name="text"><br>
        Please enter the old text you want to replace: <br>
        <input type="text" name="old_text"><br>
        Please enter the new text you want to replace it with: <br>
        <input type="text" name="new_text"><br>
        Optionally, enter which occurrence of the old text you want to replace (leave blank to replace all occurrences): <br>
        <input type="text" name="instance_num"><br>
        <input type="submit" value="Generate">
    </form>
</body>
</html>
"""

#MINIF Template
MINIF_TEMPLATE = """
<!doctype html>
<html>
<head><title>MINIF Formula Inputs</title></head>
<body>
    <a href="/">Back to Home</a>
    <form method="post" action="/result">
        <input type="hidden" name="formula" value="MINIF">
        Please enter the range to evaluate the condition (e.g., A1:A10): <br>
        <input type="text" name="condition_range"><br>
        Please enter the condition (e.g., ">20", "=MyValue"): <br>
        <input type="text" name="condition"><br>
        Please enter the range from which to find the minimum value (e.g., B1:B10): <br>
        <input type="text" name="min_range"><br>
        <input type="submit" value="Generate">
    </form>
</body>
</html>
"""

#MAXIF Template
MAXIF_TEMPLATE = """
<!doctype html>
<html>
<head><title>MAXIF Formula Inputs</title></head>
<body>
    <a href="/">Back to Home</a>
    <form method="post" action="/result">
        <input type="hidden" name="formula" value="MAXIF">
        Please enter the range to evaluate the condition (e.g., A1:A10): <br>
        <input type="text" name="condition_range"><br>
        Please enter the condition (e.g., ">20", "=MyValue"): <br>
        <input type="text" name="condition"><br>
        Please enter the range from which to find the maximum value (e.g., B1:B10): <br>
        <input type="text" name="max_range"><br>
        <input type="submit" value="Generate">
    </form>
</body>
</html>
"""

@app.route('/', methods=['GET'])
def home():
    return render_template_string(HOME_PAGE_TEMPLATE)

@app.route('/generate', methods=['POST'])
def generate():
    """
    Generate the form to prompt the user for the inputs required for the selected formula.
    - Input the user's selected formula
    - Output the form to prompt the user for the inputs required for the selected formula
    """
    formula = request.form.get('formula')
    if formula == 'SUM': # Handle SUM option
        return render_template_string(SUM_TEMPLATE)
    
    elif formula == 'IFERROR': # Handle IFERROR option
        return render_template_string(IFERROR_TEMPLATE)
    
    elif formula == 'COUNTIF':  # Handle COUNTIF option
        return render_template_string(COUNTIF_TEMPLATE)
    
    elif formula == 'VLOOKUP':  # Handle VLOOKUP option
        return render_template_string(VLOOKUP_TEMPLATE)
    
    elif formula == 'INDEXMATCH':  # Handle INDEX/MATCH option
        return render_template_string(INDEX_MATCH_TEMPLATE)
    
    elif formula == 'CONCATENATE':  # Handle CONCATENATE option
        return render_template_string(CONCATENATE_TEMPLATE)
    
    elif formula == 'CHOOSE':  # Handle CHOOSE option
        return render_template_string(CHOOSE_TEMPLATE)
    
    elif formula == 'SUBSTITUTE':  # Handle SUBSTITUTE option
        return render_template_string(SUBSTITUTE_TEMPLATE)
    
    elif formula == 'MINIF':  # Handle MINIF option
        return render_template_string(MINIF_TEMPLATE)

    elif formula == 'MAXIF':  # Handle MAXIF option
        return render_template_string(MAXIF_TEMPLATE)
    

@app.route('/result', methods=['POST'])
def result():
    """
    Generate the Excel formula based on the user's input.
        - takes in teh user's input 
        - generates the ouput formula based on the user's input
    """
    formula = request.form.get('formula')
    errors = []

    def check_input(input_type, value):
        error = validate_input(input_type, value)
        if error:
            errors.append(error)

    if formula == 'SUM':   # Generate SUM formula
        cell_range = request.form.get('range')
        if errors:  # Check for errors
            return jsonify({'errors': errors}), 400  # Return errors as JSON
        formula_result = f"=SUM({cell_range})"

    elif formula == 'IFERROR': # Generate IFERROR formula
        value = request.form.get('value')
        replacement = request.form.get('replacement')
        if errors:  # Check for errors
            return jsonify({'errors': errors}), 400  # Return errors as JSON
        formula_result = f"=IFERROR({value}, {replacement})"

    elif formula == 'COUNTIF':  # Generate COUNTIF formula
        cell_range = request.form.get('range')
        criteria = request.form.get('criteria')
        if errors:  # Check for errors
            return jsonify({'errors': errors}), 400  # Return errors as JSON
        formula_result = f"=COUNTIF({cell_range}, \"{criteria}\")"
    
    elif formula == 'VLOOKUP':  # Generate VLOOKUP formula
        lookup_value = request.form.get('lookup_value')
        table_array = request.form.get('table_array')
        col_index_num = request.form.get('col_index_num')
        range_lookup = request.form.get('range_lookup')
        if errors:  # Check for errors
            return jsonify({'errors': errors}), 400  # Return errors as JSON
        formula_result = f'=VLOOKUP({lookup_value}, {table_array}, {col_index_num}, {range_lookup})'

    elif formula == 'INDEXMATCH':  # Generate INDEX/MATCH formula
        index_range = request.form.get('index_range')
        lookup_value = request.form.get('lookup_value')
        match_range = request.form.get('match_range')
        if errors:  # Check for errors
            return jsonify({'errors': errors}), 400  # Return errors as JSON
        formula_result = f'=INDEX({index_range}, MATCH({lookup_value}, {match_range}, 0))'

    elif formula == 'CONCATENATE':  # Generate CONCATENATE formula
        first_value = request.form.get('first_value')
        second_value = request.form.get('second_value').split(',')
        values = [first_value] + second_value
        concatenated_values = ', '.join(values)
        if errors:  # Check for errors
            return jsonify({'errors': errors}), 400  # Return errors as JSON
        formula_result = f'=CONCATENATE({concatenated_values})'

    elif formula == 'CHOOSE':  # Generate CHOOSE formula
        index_num = request.form.get('index_num')
        values = request.form.get('values')
        if errors:  # Check for errors
            return jsonify({'errors': errors}), 400  # Return errors as JSON
        formula_result = f'=CHOOSE({index_num}, {values})'

    elif formula == 'SUBSTITUTE':  # Generate SUBSTITUTE formula
        text = request.form.get('text')
        old_text = request.form.get('old_text')
        new_text = request.form.get('new_text')
        instance_num = request.form.get('instance_num')
        if errors:  # Check for errors
            return jsonify({'errors': errors}), 400  # Return errors as JSON
        if instance_num:  # If an instance number is provided
            formula_result = f'=SUBSTITUTE({text}, {old_text}, {new_text}, {instance_num})'
        else:  # If no instance number is provided, replace all occurrences
            formula_result = f'=SUBSTITUTE({text}, {old_text}, {new_text})'

    elif formula == 'MINIF':  # Generate MINIF formula
        condition_range = request.form.get('condition_range')
        condition = request.form.get('condition')
        min_range = request.form.get('min_range')
        # Simulating MINIF as Excel does not have a direct function
        if errors:  # Check for errors
            return jsonify({'errors': errors}), 400  # Return errors as JSON
        formula_result = f'=MIN(IF({condition_range}{condition}, {min_range}))'

    else:  
        condition_range = request.form.get('condition_range')
        condition = request.form.get('condition')
        max_range = request.form.get('max_range')
        # Simulating MAXIF
        if errors:  # Check for errors
            return jsonify({'errors': errors}), 400  # Return errors as JSON
        formula_result = f'=MAX(IF({condition_range}{condition}, {max_range}))'
    
    return f"<h2>Your Excel formula: {formula_result}</h2><br><a href='/'>Back to Home</a>"

if __name__ == '__main__':
    app.run(debug=True)
