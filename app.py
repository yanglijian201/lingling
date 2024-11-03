from flask import Flask, request, send_file, render_template, jsonify, Response, session, make_response
import os
import traceback
import openpyxl
import xlrd
import xlwt
from xlutils.copy import copy as xl_copy
import logging
import io
import time
from threading import Event
import uuid

app = Flask(__name__)
app.secret_key = 'supersecretkey'  # Replace with a real secret key
UPLOAD_FOLDER = 'uploads'

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

# Set up custom logger for processing
log_streams = {}  # Dictionary to hold log streams for each session
processing_loggers = {}  # Dictionary to hold loggers for each session

# Event to stop processing
stop_events = {}  # Dictionary to hold stop events for each session

def get_session_id():
    if 'session_id' not in session:
        session['session_id'] = str(uuid.uuid4())
    return session['session_id']

def get_logger(session_id):
    if session_id not in log_streams:
        log_streams[session_id] = io.StringIO()
        processing_logger = logging.getLogger(session_id)
        processing_logger.setLevel(logging.INFO)
        log_handler = logging.StreamHandler(log_streams[session_id])
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        log_handler.setFormatter(formatter)
        processing_logger.addHandler(log_handler)
        processing_loggers[session_id] = processing_logger
    return processing_loggers[session_id]

def get_stop_event(session_id):
    if session_id not in stop_events:
        stop_events[session_id] = Event()
    return stop_events[session_id]

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def read_excel(file_path, processing_logger, stop_event):
    processing_logger.info(f"Reading file: {file_path}")
    time.sleep(1)  # Simulate processing delay
    if stop_event.is_set():
        processing_logger.warning("Processing stopped by user during file read.")
        return None, None, None
    try:
        if file_path.endswith('.xlsx'):
            wb = openpyxl.load_workbook(file_path)
            sheet = wb.active
            data = [[cell.value for cell in row] for row in sheet.iter_rows()]
            return data, wb, sheet
        else:
            book = xlrd.open_workbook(file_path, formatting_info=True)
            sheet = book.sheet_by_index(0)
            data = [sheet.row_values(rowx) for rowx in range(sheet.nrows)]
            return data, book, sheet
    except Exception as e:
        processing_logger.error(f"Error reading file: {e}")
        return None, None, None

def update_large_with_small(large_data, small_data, processing_logger, stop_event):
    import pdb;pdb.set_trace()
    processing_logger.info("Updating large Excel file with small Excel file content")
    time.sleep(1)  # Simulate processing delay
    if stop_event.is_set():
        processing_logger.warning("Processing stopped by user during update.")
        return None

    headers1 = large_data[0]
    data1 = large_data[1:]
    headers2 = small_data[0]
    data2 = small_data[1:]

    common_column = "common_column"  # Replace this with the actual common column name
    if common_column not in headers1 or common_column not in headers2:
        processing_logger.error(f"Common column '{common_column}' not found in one of the files.")
        return None

    index1 = headers1.index(common_column)
    index2 = headers2.index(common_column)

    data2_dict = {row[index2]: row for row in data2}

    for row1 in data1:
        if stop_event.is_set():
            return None
        key = row1[index1]
        if key in data2_dict:
            for j, value in enumerate(data2_dict[key]):
                if j != index2:  # Avoid overwriting the common column
                    row1[j] = value

    return [headers1] + data1

@app.route('/')
def index():
    response = make_response(render_template('index.html'))
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response

def stream_logs(session_id):
    while True:
        if session_id not in log_streams:
            break
        log_stream = log_streams[session_id]
        log_stream.seek(0)
        logs = log_stream.read()
        if logs:
            yield f"data: {logs}\n\n"
            # Clearing the log stream for the current session while retaining the log data
            log_stream.truncate(0)
            log_stream.seek(0)
        time.sleep(1)

@app.route('/logs')
def logs():
    session_id = get_session_id()
    response = Response(stream_logs(session_id), mimetype='text/event-stream')
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response

@app.route('/upload', methods=['POST'])
def upload_files():
    try:
        session_id = get_session_id()
        stop_event = get_stop_event(session_id)
        processing_logger = get_logger(session_id)
        stop_event.clear()
        log_streams[session_id].truncate(0)
        log_streams[session_id].seek(0)

        large_excel = request.files.get('large_excel')
        small_excel = request.files.get('small_excel')

        if not large_excel or not small_excel:
            return jsonify({'error': 'Both files are required.'}), 400

        if not allowed_file(large_excel.filename) or not allowed_file(small_excel.filename):
            return jsonify({'error': 'Both files must be in Excel format (.xls or .xlsx).'}), 400

        large_excel_path = os.path.join(UPLOAD_FOLDER, 'large_' + large_excel.filename)
        small_excel_path = os.path.join(UPLOAD_FOLDER, 'small_' + small_excel.filename)
        large_excel.save(large_excel_path)
        small_excel.save(small_excel_path)

        # Process each file individually
        large_data, large_book, large_sheet = read_excel(large_excel_path, processing_logger, stop_event)
        if large_data is None:
            return jsonify({'error': 'Error reading large Excel file or processing stopped.'}), 400

        small_data, small_book, small_sheet = read_excel(small_excel_path, processing_logger, stop_event)
        if small_data is None:
            return jsonify({'error': 'Error reading small Excel file or processing stopped.'}), 400

        # Update large Excel file with small Excel file content
        updated_data = update_large_with_small(large_data, small_data, processing_logger, stop_event)
        if updated_data is None:
            return jsonify({'error': 'Error updating large Excel file or processing stopped.'}), 400

        # Write the updated data to a new Excel file, preserving the format
        output_path = large_excel_path  # Preserve the same file name and path

        if large_excel.filename.endswith('.xlsx'):
            for i, row in enumerate(updated_data):
                for j, value in enumerate(row):
                    large_sheet.cell(row=i+1, column=j+1, value=value)
            large_book.save(output_path)
        else:
            large_book = xl_copy(large_book)
            large_sheet = large_book.get_sheet(0)
            for i, row in enumerate(updated_data):
                for j, value in enumerate(row):
                    large_sheet.write(i, j, value)
            large_book.save(output_path)

        # Clean up uploaded files
        os.remove(small_excel_path)

        response = send_file(output_path, as_attachment=True, attachment_filename=large_excel.filename)
        response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
        response.headers['Pragma'] = 'no-cache'
        response.headers['Expires'] = '0'
        return response
    except Exception as e:
        session_id = get_session_id()
        processing_logger = get_logger(session_id)
        error_message = str(e)
        stack_trace = traceback.format_exc()
        processing_logger.error(f"Error: {error_message}")
        processing_logger.error(f"Stack trace:\n{stack_trace}")
        return jsonify({'error': error_message, 'stack_trace': stack_trace}), 500

@app.route('/stop', methods=['POST'])
def stop_processing():
    session_id = get_session_id()
    stop_event = get_stop_event(session_id)
    processing_logger = get_logger(session_id)
    stop_event.set()
    processing_logger.info("Processing stopped by user.")
    response = jsonify({'message': 'Processing stopped'})
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response

@app.route('/clear_logs', methods=['POST'])
def clear_logs():
    session_id = get_session_id()
    if session_id in log_streams:
        log_streams[session_id].truncate(0)
        log_streams[session_id].seek(0)
    return jsonify({'message': 'Logs cleared'})

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8000)