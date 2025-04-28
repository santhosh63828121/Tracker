from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import sqlite3
from datetime import datetime
import openpyxl
from io import BytesIO

app = Flask(__name__, static_folder='photos')
CORS(app)

# Connect to SQLite
def get_db():
    conn = sqlite3.connect("school.db")
    conn.row_factory = sqlite3.Row
    return conn

# Create table - Directly call this function instead of using @app.before_first_request
def create_table():
    conn = get_db()
    conn.execute('''
        CREATE TABLE IF NOT EXISTS students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            qr_id TEXT UNIQUE,
            name TEXT,
            photo TEXT,
            status TEXT DEFAULT 'checked_out',
            timestamp TEXT
        )
    ''')
    conn.commit()

# Call create_table() directly
create_table()

# Insert new student route
@app.route('/add_student', methods=['POST'])
def add_student():
    # Get data from the request
    qr_id = request.json.get('qr_id')
    name = request.json.get('name')
    photo = request.json.get('photo')
    status = request.json.get('status', 'checked_out')  # Default to 'checked_out' if not provided

    if not qr_id or not name or not photo:
        return jsonify({'error': 'Missing required fields: qr_id, name, or photo'}), 400

    # Connect to database and insert new student
    conn = get_db()
    try:
        conn.execute('INSERT INTO students (qr_id, name, photo, status) VALUES (?, ?, ?, ?)', 
                     (qr_id, name, photo, status))
        conn.commit()
        return jsonify({'message': f'Student {name} added successfully'}), 201
    except sqlite3.IntegrityError:
        return jsonify({'error': 'Student with this QR ID already exists'}), 400

# Check-in/out logic (same as before)
@app.route('/scan', methods=['POST'])
def scan_qr():
    qr_id = request.json.get('qr_id')
    conn = get_db()
    student = conn.execute('SELECT * FROM students WHERE qr_id = ?', (qr_id,)).fetchone()
    
    if student:
        new_status = 'checked_out' if student['status'] == 'checked_in' else 'checked_in'
        conn.execute('UPDATE students SET status = ?, timestamp = ? WHERE qr_id = ?',
                     (new_status, datetime.now().strftime('%Y-%m-%d %H:%M:%S'), qr_id))
        conn.commit()
        return jsonify({'message': f'{student["name"]} {new_status.replace("_", " ").title()}'}), 200
    else:
        return jsonify({'error': 'Student not found'}), 404

# List students
# List checked-in students
@app.route('/students', methods=['GET'])
def get_checked_in_students():
    conn = get_db()
    students = conn.execute('SELECT * FROM students WHERE status = "checked_in"').fetchall()
    return jsonify([dict(row) for row in students])

# Export students to Excel file
@app.route('/export_students', methods=['GET'])
def export_students_to_excel():
    conn = get_db()
    students = conn.execute('SELECT * FROM students').fetchall()

    # Create a workbook and a sheet
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = "Students"

    # Add header row
    sheet.append(["ID", "QR ID", "Name", "Photo", "Status", "Timestamp"])

    # Add data rows
    for student in students:
        sheet.append([student["id"], student["qr_id"], student["name"], student["photo"], student["status"], student["timestamp"]])

    # Save workbook to a BytesIO object
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    # Send the file as a response
    return send_file(output, as_attachment=True, download_name="students.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Load sample data
@app.route('/load_sample_data', methods=['POST'])
def load_sample():
    sample_data = [
        (f"QR{i:03}", f"Student {i}", f"/photos/photo{i}.jpg") for i in range(1, 16)
    ]
    conn = get_db()
    for qr_id, name, photo in sample_data:
        try:
            conn.execute('INSERT INTO students (qr_id, name, photo) VALUES (?, ?, ?)', (qr_id, name, photo))
        except sqlite3.IntegrityError:
            continue
    conn.commit()
    return jsonify({'message': 'Sample data loaded'}), 200

if __name__ == '__main__':
    app.run(debug=True)
