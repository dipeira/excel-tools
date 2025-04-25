import os
from flask import Flask, render_template, request, send_file, jsonify, url_for
import pandas as pd
from werkzeug.utils import secure_filename
import zipfile

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Ensure upload folder exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def get_column_names(file_path):
    try:
        df = pd.read_excel(file_path)
        return df.columns.tolist()
    except Exception as e:
        return []

def get_column_values(file_path, column):
    try:
        df = pd.read_excel(file_path)
        if column not in df.columns:
            return []
        # Get unique values and convert to list, handling NaN values
        values = df[column].dropna().unique().tolist()
        # Convert all values to strings for consistency
        return [str(value) for value in values]
    except Exception as e:
        return []

def compare_files(file1_path, file2_path, col1, col2):
    df1 = pd.read_excel(file1_path)
    df2 = pd.read_excel(file2_path)
    
    # Find records in file1 not in file2
    not_in_file2 = df1[~df1[col1].isin(df2[col2])]
    
    # Find records in file2 not in file1
    not_in_file1 = df2[~df2[col2].isin(df1[col1])]
    
    # Create output Excel file
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'comparison_result.xlsx')
    with pd.ExcelWriter(output_path) as writer:
        not_in_file2.to_excel(writer, sheet_name='Not in File 2', index=False)
        not_in_file1.to_excel(writer, sheet_name='Not in File 1', index=False)
    
    return output_path

def join_files(file1_path, file2_path, col1, col2):
    df1 = pd.read_excel(file1_path)
    df2 = pd.read_excel(file2_path)
    
    # Rename the join column in df2 to match df1 to avoid duplicates
    df2 = df2.rename(columns={col2: col1})
    
    # Perform the join operation (similar to SQL INNER JOIN)
    joined_df = pd.merge(df1, df2, on=col1, how='inner')
    
    # Create output Excel file
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'joined_result.xlsx')
    joined_df.to_excel(output_path, index=False)
    
    return output_path

def split_file(file_path, column, value):
    df = pd.read_excel(file_path)
    
    # Filter rows where the column value matches the selected value
    # Convert both to strings for comparison to handle different data types
    filtered_df = df[df[column].astype(str) == str(value)]
    
    # Create output Excel file
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], f'split_{value}.xlsx')
    filtered_df.to_excel(output_path, index=False)
    
    return output_path

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/compare')
def compare_page():
    return render_template('compare.html')

@app.route('/join')
def join_page():
    return render_template('join.html')

@app.route('/filter')
def filter_page():
    return render_template('filter.html')

@app.route('/split')
def split_by_column_page():
    return render_template('split.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'Δεν υπάρχει τμήμα αρχείου'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'Δεν έχει επιλεγεί αρχείο'}), 400
    
    if file and file.filename.endswith('.xlsx'):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        
        columns = get_column_names(file_path)
        return jsonify({'columns': columns, 'filename': filename})
    
    return jsonify({'error': 'Μη έγκυρος τύπος αρχείου'}), 400

@app.route('/column-values', methods=['POST'])
def column_values():
    if 'file' not in request.files:
        return jsonify({'error': 'Δεν έχει μεταφορτωθεί αρχείο'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'Δεν έχει επιλεγεί αρχείο'}), 400
    
    if not file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'Παρακαλώ μεταφορτώστε αρχείο Excel'}), 400
    
    column = request.form.get('column')
    if not column:
        return jsonify({'error': 'Δεν έχει επιλεγεί στήλη'}), 400
    
    try:
        # Save the file temporarily
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        
        # Get unique values for the column
        values = get_column_values(file_path, column)
        
        # Clean up the temporary file
        os.remove(file_path)
        
        return jsonify({'values': values})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/compare', methods=['POST'])
def compare():
    data = request.json
    file1_path = os.path.join(app.config['UPLOAD_FOLDER'], data['file1'])
    file2_path = os.path.join(app.config['UPLOAD_FOLDER'], data['file2'])
    
    try:
        output_path = compare_files(file1_path, file2_path, data['col1'], data['col2'])
        return send_file(output_path, as_attachment=True)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/join', methods=['POST'])
def join():
    data = request.json
    file1_path = os.path.join(app.config['UPLOAD_FOLDER'], data['file1'])
    file2_path = os.path.join(app.config['UPLOAD_FOLDER'], data['file2'])
    
    try:
        output_path = join_files(file1_path, file2_path, data['col1'], data['col2'])
        return send_file(output_path, as_attachment=True)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/filter', methods=['POST'])
def filter_file():
    if 'file' not in request.files:
        return render_template('filter.html', error='Δεν έχει μεταφορτωθεί αρχείο')
    
    file = request.files['file']
    if file.filename == '':
        return render_template('filter.html', error='Δεν έχει επιλεγεί αρχείο')
    
    if not file.filename.endswith(('.xlsx', '.xls')):
        return render_template('filter.html', error='Παρακαλώ μεταφορτώστε αρχείο Excel')
    
    column = request.form.get('column')
    value = request.form.get('value')
    
    if not column or not value:
        return render_template('filter.html', error='Παρακαλώ επιλέξτε στήλη και εισάγετε τιμή')
    
    try:
        # Read the Excel file
        df = pd.read_excel(file)
        
        if column not in df.columns:
            return render_template('filter.html', error=f'Η στήλη "{column}" δεν βρέθηκε στο αρχείο')
        
        # Filter the dataframe
        filtered_df = df[df[column].astype(str) == str(value)]
        
        if filtered_df.empty:
            return render_template('filter.html', error=f'Δεν βρέθηκαν γραμμές όπου {column} ισούται με "{value}"')
        
        # Create output filename
        output_filename = f'filtered_{secure_filename(file.filename)}'
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
        
        # Save the filtered dataframe
        filtered_df.to_excel(output_path, index=False)
        
        return render_template('filter.html', 
                             success=f'Βρέθηκαν {len(filtered_df)} γραμμές στις οποίες {column} ισούται με "{value}"',
                             download_link={'url': url_for('download_file', filename=output_filename),
                                          'filename': output_filename})
        
    except Exception as e:
        return render_template('filter.html', error=f'Σφάλμα κατά την επεξεργασία του αρχείου: {str(e)}')


@app.route('/get-columns', methods=['POST'])
def get_columns():
    if 'file' not in request.files:
        return jsonify({'error': 'Δεν έχει μεταφορτωθεί αρχείο'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'Δεν έχει επιλεγεί αρχείο'}), 400
    
    if not file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'Παρακαλώ μεταφορτώστε αρχείο Excel'}), 400
    
    try:
        # Read the Excel file
        df = pd.read_excel(file)
        
        # Get column names
        columns = df.columns.tolist()
        
        return jsonify({'columns': columns})
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download/<path:filename>')
def download_file(filename):
    try:
        return send_file(
            os.path.join(app.config['UPLOAD_FOLDER'], filename),
            as_attachment=True
        )
    except Exception as e:
        return str(e), 404

@app.route('/split', methods=['POST'])
def split():
    if 'file' not in request.files:
        return render_template('split.html', error='Δεν έχει μεταφορτωθεί αρχείο')
    
    file = request.files['file']
    if file.filename == '':
        return render_template('split.html', error='Δεν έχει επιλεγεί αρχείο')
    
    if not file.filename.endswith(('.xlsx', '.xls')):
        return render_template('split.html', error='Παρακαλώ μεταφορτώστε αρχείο Excel')
    
    column = request.form.get('column')
    if not column:
        return render_template('split.html', error='Παρακαλώ επιλέξτε στήλη')
    
    try:
        # Read the Excel file
        df = pd.read_excel(file)
        
        if column not in df.columns:
            return render_template('split.html', error=f'Η στήλη "{column}" δεν βρέθηκε στο αρχείο')
        
        # Create a directory for split files
        output_dir = os.path.join(app.config['UPLOAD_FOLDER'], 'split_files')
        os.makedirs(output_dir, exist_ok=True)
        
        # Split the dataframe by unique values in the selected column
        unique_values = df[column].unique()
        split_files = []
        
        for value in unique_values:
            # Create a clean filename from the value
            clean_value = str(value).replace('/', '_').replace('\\', '_')
            filename = f'split_{clean_value}.xlsx'
            filepath = os.path.join(output_dir, filename)
            
            # Save the filtered dataframe
            df[df[column] == value].to_excel(filepath, index=False)
            split_files.append(filepath)
        
        # Create a zip file containing all split files
        zip_filename = f'split_files_{secure_filename(file.filename)}.zip'
        zip_path = os.path.join(app.config['UPLOAD_FOLDER'], zip_filename)
        
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for file_path in split_files:
                arcname = os.path.basename(file_path)
                zipf.write(file_path, arcname)
        
        # Clean up individual split files
        for file_path in split_files:
            os.remove(file_path)
        
        return render_template('split.html', 
                             success=f'Το αρχείο διαχωρίστηκε σε {len(split_files)} αρχεία με βάση τη στήλη "{column}"',
                             download_link={'url': url_for('download_file', filename=zip_filename),
                                          'filename': zip_filename})
        
    except Exception as e:
        return render_template('split.html', error=f'Σφάλμα κατά την επεξεργασία του αρχείου: {str(e)}')

if __name__ == '__main__':
    # Get port from environment variable or default to 8000
    port = int(os.environ.get('PORT', 8000))
    # Run the app on all interfaces
    app.run(host='0.0.0.0', port=port) 