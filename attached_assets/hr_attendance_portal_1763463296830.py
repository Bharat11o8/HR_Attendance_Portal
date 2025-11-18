from flask import Flask, render_template, request, send_file
from datetime import datetime, timedelta
import pandas as pd
from io import BytesIO
import re

app = Flask(__name__)

# Hardcoded employee master data
EMPLOYEE_MASTER = {
    2: "Rishi",
    3: "Ankur Jain",
    4: "Sahtosh Sharma",
    5: "Gaurav",
    7: "Gunjan",
    9: "Aarti",
    10: "Akansha Bajpai",
    11: "Chirag Channa",
    14: "Sumit",
    15: "Sanjay Dwivedi",
    17: "Himanshu Gandhi",
    19: "Sandhya Jha",
    20: "Vijaya",
    26: "Saurabh",
    31: "Anshika Singh",
    34: "Pankaj Vij",
    35: "Kiran",
    36: "Hardevi",
    37: "Sadhana",
    38: "Kanchani",
    32: "Prabhat",
    40: "Naman",
    41: "Ashish Rai",
    42: "Bharat"
}

def parse_biometric_file(file_content):
    """Parse the biometric .txt file"""
    lines = file_content.decode('utf-8').strip().split('\n')
    
    # Skip header line
    data_lines = lines[1:]
    
    records = []
    for line in data_lines:
        if not line.strip():
            continue
        
        # Split by tabs or multiple spaces
        parts = re.split(r'\t+|\s{2,}', line.strip())
        
        if len(parts) >= 10:
            try:
                en_no = int(parts[2])  # EnNo column
                date_time_str = parts[9]  # DateTime column
                
                # Parse datetime
                dt = datetime.strptime(date_time_str, '%Y-%m-%d %H:%M:%S')
                
                records.append({
                    'EnNo': en_no,
                    'DateTime': dt,
                    'Date': dt.date(),
                    'Time': dt.time()
                })
            except (ValueError, IndexError):
                continue
    
    return records

def process_attendance(records, start_date, end_date):
    """Process attendance records and generate output"""
    
    # Filter records by date range and employee list
    filtered_records = [
        r for r in records 
        if r['EnNo'] in EMPLOYEE_MASTER 
        and start_date <= r['Date'] <= end_date
    ]
    
    # Group by employee and date
    attendance_dict = {}
    for record in filtered_records:
        key = (record['EnNo'], record['Date'])
        if key not in attendance_dict:
            attendance_dict[key] = {'in_times': [], 'out_times': []}
        
        attendance_dict[key]['in_times'].append(record['DateTime'])
        attendance_dict[key]['out_times'].append(record['DateTime'])
    
    # Calculate earliest IN and latest OUT for each day
    daily_attendance = {}
    for (en_no, date), times in attendance_dict.items():
        in_time = min(times['in_times']).time()
        out_time = max(times['out_times']).time()
        daily_attendance[(en_no, date)] = {'in': in_time, 'out': out_time}
    
    # Generate output for all employees and all dates
    output_data = []
    current_date = start_date
    date_range = []
    
    while current_date <= end_date:
        date_range.append(current_date)
        current_date += timedelta(days=1)
    
    for emp_id in sorted(EMPLOYEE_MASTER.keys()):
        emp_name = EMPLOYEE_MASTER[emp_id]
        
        for date in date_range:
            key = (emp_id, date)
            
            if key in daily_attendance:
                in_time = daily_attendance[key]['in']
                out_time = daily_attendance[key]['out']
                
                # Check if late (after 9:30 AM)
                late_threshold = datetime.strptime('09:30:00', '%H:%M:%S').time()
                remark = 'Late' if in_time > late_threshold else ''
                
                in_str = in_time.strftime('%H:%M:%S')
                out_str = out_time.strftime('%H:%M:%S')
            else:
                in_str = ''
                out_str = ''
                remark = ''
            
            output_data.append({
                'ID': emp_id,
                'Name': emp_name,
                'Date': date.strftime('%Y-%m-%d'),
                'IN': in_str,
                'OUT': out_str,
                'Remark': remark
            })
    
    return output_data

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    try:
        # Get uploaded file
        biometric_file = request.files['biometric_file']
        
        # Get date range
        start_date_str = request.form['start_date']
        end_date_str = request.form['end_date']
        
        start_date = datetime.strptime(start_date_str, '%Y-%m-%d').date()
        end_date = datetime.strptime(end_date_str, '%Y-%m-%d').date()
        
        # Parse biometric data
        file_content = biometric_file.read()
        records = parse_biometric_file(file_content)
        
        # Process attendance
        output_data = process_attendance(records, start_date, end_date)
        
        # Create Excel file
        df = pd.DataFrame(output_data)
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Attendance')
        
        output.seek(0)
        
        filename = f'Attendance_{start_date_str}_to_{end_date_str}.xlsx'
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
    
    except Exception as e:
        return f"Error processing file: {str(e)}", 400

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)