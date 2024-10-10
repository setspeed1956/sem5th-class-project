from flask import Flask, request, send_file
import pandas as pd
import numpy as np
import os
from datetime import datetime

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        
        mst_file = request.files['mst_file']
        practical_file = request.files['practical_file']
        attendance_file = request.files['attendance']

        
        if not (mst_file.filename.endswith('.xlsx') and practical_file.filename.endswith('.xlsx') and attendance_file.filename.endswith('.xlsx')):
            return "Error: All uploaded files must be Excel files (.xlsx)"

        
        try:
            mst_df = pd.read_excel(mst_file)
            practical_df = pd.read_excel(practical_file)
            attendance_df = pd.read_excel(attendance_file)
        except Exception as e:
            return f"Error reading the files: {e}"

        
        choice = request.form['choice']

        
        if choice == 'custom':
            try:
                mst1_weight = float(request.form['mst1_weight'])
                mst2_weight = float(request.form['mst2_weight'])
            except ValueError:
                return "Error: Please enter valid numbers for MST1 and MST2 weightages."

            
            if mst1_weight + mst2_weight != 100:
                return "Error: MST1 and MST2 weightage must add up to 100."
            
            
            mst_marks = (mst_df['MST1'] * (mst1_weight / 100)) + (mst_df['MST2'] * (mst2_weight / 100))
        else:
            
            try:
                if choice == 'average':
                    mst_marks = (mst_df['MST1'] + mst_df['MST2']) / 2
                else:
                    mst_marks = mst_df[['MST1', 'MST2']].max(axis=1)
            except KeyError as e:
                return f"Error: Missing MST columns in the file: {e}"

        
        try:
            mst_weightage = float(request.form['mst_weightage'])
            practical_weightage = float(request.form['practical_weightage'])
            attendance_weightage = float(request.form['attendance_weightage'])
        except ValueError:
            return "Error: Please enter valid numbers for weightages."

        
        if mst_weightage + practical_weightage + attendance_weightage != 30:
            return "Error: Total weightage must add up to 30."

        
        try:
            demo1_weight = float(request.form['demo1_weight'])
            demo2_weight = float(request.form['demo2_weight'])
            quiz1_weight = float(request.form['quiz1_weight'])
            quiz2_weight = float(request.form['quiz2_weight'])
        except ValueError:
            return "Error: Please enter valid numbers for practical weightages."

        
        practical_marks = (
            (practical_df['demo1'] * (demo1_weight / 100)) +
            (practical_df['demo2'] * (demo2_weight / 100)) +
            (practical_df['quiz1'] * (quiz1_weight / 100)) +
            (practical_df['quiz2'] * (quiz2_weight / 100))
        )

        
        result_df = pd.DataFrame({'Name': mst_df['Name'], 'Roll_No': mst_df['Roll_No'], 'MST': mst_marks})
        result_df['Practical'] = practical_marks

        
        if 'Attendance' in attendance_df.columns:
            result_df['Attendance'] = attendance_df['Attendance']
        else:
            result_df['Attendance'] = 0  
        
        max_mst = result_df['MST'].max()
        max_practical = result_df['Practical'].max()
        max_attendance = result_df['Attendance'].max()

        if max_mst == 0:
            return "Error: Maximum value in MST is zero, unable to normalize MST marks."
        if max_practical == 0:
            return "Error: Maximum value in Practical is zero, unable to normalize Practical marks."
        if max_attendance == 0:
            return "Error: Maximum value in Attendance is zero, unable to normalize Attendance marks."

       
        result_df['MST'] = (result_df['MST'] / max_mst) * mst_weightage
        result_df['Practical'] = (result_df['Practical'] / max_practical) * practical_weightage
        result_df['Attendance'] = (result_df['Attendance'] / max_attendance) * attendance_weightage

        
        result_df['MST'] = result_df['MST'].round(2)
        result_df['Practical'] = result_df['Practical'].round(2)
        result_df['Attendance'] = result_df['Attendance'].round(2)

       
        result_df['Total'] = (result_df['MST'] + result_df['Practical'] + result_df['Attendance']).round(2)

        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        result_file = f'result_{timestamp}.xlsx'
        result_df.to_excel(result_file, index=False)

        
        return send_file(result_file, as_attachment=True)

    return '''
        <html>
            <head>
                <title>Upload Excel Files</title>
                <style>
                    body {
                        font-family: Arial, sans-serif;
                        background-color: #f4f4f4;
                        color: #333;
                        padding: 20px;
                    }
                    h1 {
                        color: #4CAF50;
                        text-align: center;
                    }
                    form {
                        background-color: #fff;
                        padding: 20px;
                        border-radius: 5px;
                        box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
                        max-width: 600px;
                        margin: auto;
                    }
                    input[type="file"], input[type="number"], select {
                        width: calc(100% - 20px);
                        padding: 10px;
                        margin: 10px 0;
                        border: 1px solid #ccc;
                        border-radius: 4px;
                    }
                    input[type="submit"] {
                        background-color: #4CAF50;
                        color: white;
                        padding: 10px 15px;
                        border: none;
                        border-radius: 4px;
                        cursor: pointer;
                        width: 100%;
                    }
                    input[type="submit"]:hover {
                        background-color: #45a049;
                    }
                    label {
                        margin-top: 10px;
                        font-weight: bold;
                        display: block;
                    }
                    #custom_weights {
                        display: none;
                        margin-top: 10px;
                        padding: 10px;
                        border: 1px solid #ddd;
                        border-radius: 5px;
                        background-color: #f9f9f9;
                    }
                </style>
            </head>
            <body>
                <h1>Upload Excel Files</h1>
                <form action="" method="post" enctype="multipart/form-data">
                    <h2>MST Marks</h2>
                    <input type="file" name="mst_file" required>
                    <h2>Practical Marks</h2>
                    <input type="file" name="practical_file" required>
                    <h2>Attendance Report</h2>
                    <input type="file" name="attendance" required>
                    <label for="choice">Choose an option:</label>
                    <select id="choice" name="choice">
                        <option value="average">Average MST1 and MST2</option>
                        <option value="max">Maximum of MST1 and MST2</option>
                        <option value="custom">Custom Weightage for MST1 and MST2</option>
                    </select>
                    
                    <div id="custom_weights">
                        <label for="mst1_weight">MST1 Weightage (%):</label>
                        <input type="number" id="mst1_weight" name="mst1_weight" min="0" max="100">
                        <label for="mst2_weight">MST2 Weightage (%):</label>
                        <input type="number" id="mst2_weight" name="mst2_weight" min="0" max="100">
                    </div>

                    <h2>Practical Weightages</h2>
                    <label for="demo1_weight">Demo1 Weightage (%):</label>
                    <input type="number" id="demo1_weight" name="demo1_weight" required min="0" max="100">
                    <label for="demo2_weight">Demo2 Weightage (%):</label>
                    <input type="number" id="demo2_weight" name="demo2_weight" required min="0" max="100">
                    <label for="quiz1_weight">Quiz1 Weightage (%):</label>
                    <input type="number" id="quiz1_weight" name="quiz1_weight" required min="0" max="100">
                    <label for="quiz2_weight">Quiz2 Weightage (%):</label>
                    <input type="number" id="quiz2_weight" name="quiz2_weight" required min="0" max="100">

                    <h2>Weightages</h2>
                    <label for="mst_weightage">MST Weightage:</label>
                    <input type="number" id="mst_weightage" name="mst_weightage" required>
                    <label for="practical_weightage">Practical Weightage:</label>
                    <input type="number" id="practical_weightage" name="practical_weightage" required>
                    <label for="attendance_weightage">Attendance Weightage:</label>
                    <input type="number" id="attendance_weightage" name="attendance_weightage" required>
                    
                    <input type="submit" value="Generate Result">
                </form>

                <script>
                    document.getElementById('choice').addEventListener('change', function() {
                        if (this.value === 'custom') {
                            document.getElementById('custom_weights').style.display = 'block';
                        } else {
                            document.getElementById('custom_weights').style.display = 'none';
                        }
                    });
                </script>
            </body>
        </html>
    '''

if __name__ == '__main__':
    app.run(debug=True)
