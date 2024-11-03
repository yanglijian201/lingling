from flask import Flask, request, send_file, render_template, jsonify, Response, session, make_response
import os
import traceback
import openpyxl
import xlrd
import xlwt
from xlutils.copy import copy as xl_copy
import logging
import io
import threading
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
log_seek_location = {}  # Dictionary to hold seek locations for each session

# Event to stop processing
stop_events = {}  # Dictionary to hold stop events for each session

def get_session_id():
    if 'session_id' not in session:
        session['session_id'] = str(uuid.uuid4())
    return session['session_id']

class DiscardAfterReadIO(io.StringIO):
    lock = {}

    def __init__(self, initial_value='', newline='\n', session_id=None):
        super().__init__(initial_value, newline)
        session_id = get_session_id()
        if session_id not in self.lock:
            self.lock[session_id] = threading.Lock()
        self._lock = self.lock[session_id]

    def read(self, size=-1):
        with self._lock:
            self.seek(0)
            result = super().read(size)
            self.seek(0)
            self.truncate(0)
            return result

    def readline(self, size=-1):
        with self._lock:
            self.seek(0)
            result = super().readline(size)
            self.seek(0)
            self.truncate(0)
            return result

    def readlines(self, hint=-1):
        with self._lock:
            self.seek(0)
            result = super().readlines(hint)
            self.seek(0)
            self.truncate(0)
            return result

    def write(self, s):
        with self._lock:
            return super().write(s)

    def getvalue(self):
        with self._lock:
            result = super().getvalue()
            self.seek(0)
            self.truncate(0)
            return result

def get_logger(session_id):
    if session_id not in log_streams:
        log_streams[session_id] = DiscardAfterReadIO()
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

def summarize_large_data(large_data, processing_logger, stop_event):
    processing_logger.info("Summarizing large Excel data")
    if stop_event.is_set():
        processing_logger.warning("Processing stopped by user during summarization.")
        return None

    headers = large_data[0]
    data = large_data[1:]

    # Example summary: count the number of rows and calculate the sum of a specific column
    row_count = len(data)
    sum_column = "ColumnToSum"  # Replace with the actual column name to summarize
    if sum_column not in headers:
        processing_logger.error(f"Column '{sum_column}' not found in large Excel file.")
        return None

    sum_index = headers.index(sum_column)
    column_sum = sum(row[sum_index] for row in data if isinstance(row[sum_index], (int, float)))

    summary_data = [
        ["Summary", ""],
        ["Total Rows", row_count],
        [f"Sum of {sum_column}", column_sum]
    ]

    return summary_data

def field_verify(fields_dict, data_line):
    for location, field_name in fields_dict.items():
        if field_name not in data_line[location]:
            raise ValueError(f"Field '{field_name}' not found in {data_line}.")

def validate_summary_data(file_name, summary_data, processing_logger):
    processing_logger.info(f"Validating data for: {file_name}")
    fields_dict = {
        0: "医疗机构分类",
        1: "医院总数量",
        2: "参加国家卫生健康委员会临床检验中心室间质量评价合格医院数量",
        3: "通过国家室间质评平均合格项目数量",
        4: "参加辽宁省临床检验中心室间质量评价合格医院数量",
        5: "通过辽宁省室间质评平均合格项目数量",
        6: "参加辽宁省医学影像质控中心影像质控认证评价合格医院数量",
        7: "要求对其他医疗机构标有互认标识的医学影像检查资料和医学检验结果认可的医院数量",
        8: "认可其他医疗机构标有互认标识的医学影像检查资料和医学检验结果互",
        9: "盟医院间通过信息系统调阅成员医院间医学影像检查资料和医学检验结果频次合计",
        10: "实施检查检验结果互认为患者节约医疗费",
    }
    data_line = summary_data[3]
    field_verify(fields_dict, data_line)
    processing_logger.info(f"Data validation successful for {file_name}.")

def validate_large_data(file_name, large_data, processing_logger):
    processing_logger.info(f"Validating data for: {file_name}")
    fields_dict = {
        0: "医疗机构名称",
        1: "医疗机构分类",
        2: "是否参加国家卫生健康委员会临床检验中心室间质量评价并合格",
        3: "通过国家室间质评医学检验结果合格项目数量",
        4: "是否参加辽宁省临床检验中心室间质量评价并合格",
        5: "通过辽宁省室间质评医学检验结果合格项目数量",
        6: "是否参加辽宁省医学影像质控中心影像质控认证评价并合格",
        7: "要求对其他医疗机构标有互认标识的医学影像检查资料和医学检验结果认可",
        8: "DR互认项目数",
        9: "DR节约检查费用",
        10: "MR互认项目数",
        11: "MR节约检查费用",
        12: "CT互认项目数",
        13: "CT节约检查费用",
        14: "临床检验互认项目数",
        15: "临床检验节约检查费用",
        16: "通过信息系统调阅成员医院间医学影像检查资料和医学检验结果频",
        17: "总项目",
        18: "总费用",
    }
    data_line = large_data[4]
    field_verify(fields_dict, data_line)
    processing_logger.info(f"Data validation successful for {file_name}.")


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
        logs = log_stream.read()
        if logs:
            for line in logs.splitlines():
                yield f"data: {line}\n\n"
            # Clearing the log stream for the current session while retaining the log data
        time.sleep(0.1)

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
        summary_excel = request.files.get('summary_excel')

        if not large_excel or not summary_excel:
            processing_logger.error("Both files are required.")
            return jsonify({'error': 'Both files are required.'}), 400

        if not allowed_file(large_excel.filename) or not allowed_file(summary_excel.filename):
            processing_logger.error("Both files must be in Excel format (.xls or .xlsx).")
            return jsonify({'error': 'Both files must be in Excel format (.xls or .xlsx).'}), 400

        large_excel_path = os.path.join(UPLOAD_FOLDER, 'large_' + large_excel.filename)
        summary_excel_path = os.path.join(UPLOAD_FOLDER, 'summary_' + summary_excel.filename)
        large_excel.save(large_excel_path)
        summary_excel.save(summary_excel_path)

        # Process each file individually
        large_data, large_book, large_sheet = read_excel(large_excel_path, processing_logger, stop_event)
        if large_data is None:
            return jsonify({'error': 'Error reading large Excel file or processing stopped.'}), 400
        validate_large_data(large_excel.filename, large_data, processing_logger)

        summary_data, summary_book, summary_sheet = read_excel(summary_excel_path, processing_logger, stop_event)
        if summary_data is None:
            return jsonify({'error': 'Error reading summary Excel file or processing stopped.'}), 400
        validate_summary_data(summary_excel.filename, summary_data, processing_logger)

        # Summarize large Excel data
        summarized_data = summarize_large_data(large_data, processing_logger, stop_event)
        if summarized_data is None:
            return jsonify({'error': 'Error summarizing large Excel data or processing stopped.'}), 400

        # Write the summarized data to the summary Excel file, preserving the format
        output_path = summary_excel_path  # Preserve the same file name and path

        if summary_excel.filename.endswith('.xlsx'):
            for i, row in enumerate(summarized_data):
                for j, value in enumerate(row):
                    summary_sheet.cell(row=i+1, column=j+1, value=value)
            summary_book.save(output_path)
        else:
            summary_book = xl_copy(summary_book)
            summary_sheet = summary_book.get_sheet(0)
            for i, row in enumerate(summarized_data):
                for j, value in enumerate(row):
                    summary_sheet.write(i, j, value)
            summary_book.save(output_path)

        # Clean up uploaded files
        os.remove(large_excel_path)

        response = send_file(output_path, as_attachment=True, attachment_filename=summary_excel.filename)
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
    # Disable Flask's default logging
    # log = logging.getLogger('werkzeug')
    # log.setLevel(logging.ERROR)
    app.run(debug=True, host='0.0.0.0', port=8000)
