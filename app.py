from flask import Flask, request, send_file, render_template, jsonify
import pandas as pd
import os
import traceback
import openpyxl
import xlrd
import xlwt
from xlutils.copy import copy as xl_copy

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    try:
        large_excel = request.files['large_excel']
        same_excel = request.files['same_excel']
        if not large_excel or not same_excel:
            return jsonify({'error': 'Both files are required.'}), 400
        if not allowed_file(large_excel.filename) or not allowed_file(same_excel.filename):
            return jsonify({'error': 'Both files must be in Excel format (.xls or .xlsx).'}), 400

        large_excel_path = os.path.join(app.config['UPLOAD_FOLDER'], large_excel.filename)
        same_excel_path = os.path.join(app.config['UPLOAD_FOLDER'], same_excel.filename)
        large_excel.save(large_excel_path)
        same_excel.save(same_excel_path)

        # Read the files and convert them to dataframes
        if large_excel.filename.endswith('.xlsx'):
            wb1 = openpyxl.load_workbook(large_excel_path)
            sheet1 = wb1.active
            df1 = pd.DataFrame(sheet1.values)
        else:
            book1 = xlrd.open_workbook(large_excel_path, formatting_info=True)
            sheet1 = book1.sheet_by_index(0)
            df1 = pd.DataFrame([sheet1.row_values(rowx) for rowx in range(sheet1.nrows)])

        if same_excel.filename.endswith('.xlsx'):
            wb2 = openpyxl.load_workbook(same_excel_path)
            sheet2 = wb2.active
            df2 = pd.DataFrame(sheet2.values)
        else:
            book2 = xlrd.open_workbook(same_excel_path, formatting_info=True)
            sheet2 = book2.sheet_by_index(0)
            df2 = pd.DataFrame([sheet2.row_values(rowx) for rowx in range(sheet2.nrows)])

        # Example processing: merge the two dataframes
        # Assuming the first row is the header
        df1.columns = df1.iloc[0]
        df1 = df1[1:]
        df2.columns = df2.iloc[0]
        df2 = df2[1:]
        result_df = pd.merge(df1, df2, on='common_column')

        # Write the result dataframe to a new Excel file
        if large_excel.filename.endswith('.xlsx'):
            output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'result.xlsx')
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                result_df.to_excel(writer, index=False, sheet_name='Sheet1')

                # Get the workbook and the active sheet
                workbook = writer.book
                worksheet = writer.sheets['Sheet1']

                # Copy the styles from the original sheet to the new sheet
                for row in sheet1.iter_rows():
                    for cell in row:
                        new_cell = worksheet[cell.coordinate]
                        new_cell.font = cell.font
                        new_cell.border = cell.border
                        new_cell.fill = cell.fill
                        new_cell.number_format = cell.number_format
                        new_cell.protection = cell.protection
                        new_cell.alignment = cell.alignment
        else:
            output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'result.xls')
            book = xl_copy(book1)
            sheet = book.get_sheet(0)

            for i, row in result_df.iterrows():
                for j, value in enumerate(row):
                    sheet.write(i + 1, j, value)  # start from the second row

            book.save(output_path)

        return send_file(output_path, as_attachment=True, download_name=os.path.basename(output_path))
    except Exception as e:
        error_message = str(e)
        stack_trace = traceback.format_exc()
        print(f"Error: {error_message}")
        print(f"Stack trace:\n{stack_trace}")
        return jsonify({'error': error_message, 'stack_trace': stack_trace}), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8000)