from flask import Flask, request, send_file
import pandas as pd

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Get the uploaded Excel files
        mst_file = request.files['mst_file']
        practical_file = request.files['practical_file']
        classwork_file = request.files['classwork']

        # Read the Excel files into Pandas DataFrames
        mst_df = pd.read_excel(mst_file)
        practical_df = pd.read_excel(practical_file)
        cw_df = pd.read_excel(classwork_file)

        # Get the user's choice (average or max)
        choice = request.form['choice']

        # Calculate MST marks based on user's choice
        if choice == 'average':
            mst_marks = (mst_df['MST1'] + mst_df['MST2']) / 2
        else:
            mst_marks = mst_df[['MST1', 'MST2']].max(axis=1)

        # Get the weightage for each part
        mst_weightage = float(request.form['mst_weightage'])
        practical_weightage = float(request.form['practical_weightage'])
        cw_weightage = float(request.form['cw_weightage'])

        # Check if the weightage adds up to 30
        if mst_weightage + practical_weightage + cw_weightage != 30:
            return "Error: Weightage does not add up to 30"

        # Create a new DataFrame with the processed data
        result_df = pd.DataFrame({'Name': mst_df['Name'], 'Roll_No': mst_df['Roll_No'], 'MST': mst_marks})
        result_df['Practical'] = practical_df['Practical']
        
        # Check if 'classwork' column exists in the cw_df DataFrame
        if 'classwork' in cw_df.columns:
            result_df['CW'] = cw_df['classwork']
        else:
            result_df['CW'] = np.nan  # or some default value

        # Calculate marks for each part based on weightage
        result_df['MST'] = (result_df['MST'] / result_df['MST'].max()) * mst_weightage
        result_df['Practical'] = (result_df['Practical'] / result_df['Practical'].max()) * practical_weightage
        result_df['CW'] = (result_df['CW'] / result_df['CW'].max()) * cw_weightage

        # Create an Excel file from the resulting DataFrame
        result_file = 'result.xlsx'
        result_df.to_excel(result_file, index=False)

        # Send the resulting Excel file as a response
        return send_file(result_file, as_attachment=True)

    return '''
        <html>
            <body>
                <h1>Upload Excel Files</h1>
                <form action="" method="post" enctype="multipart/form-data">
                <h1> mst marks </h1>
                    <input type="file" name="mst_file" required>
                    <h1> Practical marks </h1>
                    <input type="file" name="practical_file" required>
                    <h1> classWork marks </h1>
                    <input type="file" name="classwork" required>
                    <br><br><br>
                    <label for="choice">Choose an option:</label>
                    <select id="choice" name="choice">
                        <option value="average">Average MST1 and MST2</option>
                        <option value="max">Maximum of MST1 and MST2</option>
                    </select>
                    <br>
                    <br><br><br>
                    <label for="mst_weightage">MST Weightage:</label>
                    <input type="number" id="mst_weightage" name="mst_weightage" required>
                    <br>
                    <label for="practical_weightage">Practical Weightage:</label>
                    <input type="number" id="practical_weightage" name="practical_weightage" required>
                    <br>
                    <label for="cw_weightage">CW Weightage:</label>
                    <input type="number" id="cw_weightage" name="cw_weightage" required>
                    <br><br><br>
                    
                    <input type="submit" value="Generate Result">
                </form>
            </body>
        </html>
    '''

if __name__ == '__main__':
    app.run(debug=True)