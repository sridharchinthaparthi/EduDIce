from flask import Flask, render_template, request, redirect, send_file
import os
import pandas as pd
import random
from datetime import datetime

app = Flask(__name__)

# Set the upload folder path
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def assign_topics_randomly(roll_numbers_df, topics_df, students_per_topic):
    students = roll_numbers_df['rollnumber'].tolist()
    topics = topics_df['topics'].tolist()

    # Create a new DataFrame to store the assignments
    assignments_df = pd.DataFrame(columns=['Roll Number', 'Assigned Topic', 'Assigned Time'])

    # Shuffle the list of students
    random.shuffle(students)

    # Assign topics to students
    for i in range(0, len(students), students_per_topic):
        selected_students = students[i:i + students_per_topic]
        assigned_topic = random.choice(topics)
        assigned_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        for student in selected_students:
            assignments_df = pd.concat([assignments_df, pd.DataFrame({
                'Roll Number': [student],
                'Assigned Topic': [assigned_topic],
                'Assigned Time': [assigned_time]
            })], ignore_index=True)

    return assignments_df

def filter_latest_assignments(assignments_df):
    latest_assignments = assignments_df.sort_values(['Roll Number', 'Assigned Time'], ascending=[True, False]).drop_duplicates('Roll Number')
    return latest_assignments

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Check if the POST request has the file parts
        if 'roll_file' not in request.files or 'topic_file' not in request.files:
            return redirect(request.url)

        roll_file = request.files['roll_file']
        topic_file = request.files['topic_file']

        # If the user does not select a file, the browser submits an empty file without a filename
        if roll_file.filename == '' or topic_file.filename == '':
            return redirect(request.url)

        # If the files exist and are allowed, save them to the uploads folder
        if roll_file and topic_file:
            roll_filename = os.path.join(app.config['UPLOAD_FOLDER'], roll_file.filename)
            topic_filename = os.path.join(app.config['UPLOAD_FOLDER'], topic_file.filename)

            roll_file.save(roll_filename)
            topic_file.save(topic_filename)

            # Read Excel files
            roll_numbers_df = pd.read_excel(roll_filename)
            topics_df = pd.read_excel(topic_filename)

            # Delete the uploaded files (optional)
            os.remove(roll_filename)
            os.remove(topic_filename)

            students_per_topic = 5
            assignments = assign_topics_randomly(roll_numbers_df, topics_df, students_per_topic)
            latest_assignments = filter_latest_assignments(assignments)

            # Save the assignments to a new Excel file
            latest_assignments.to_excel('assignments.xlsx', index=False)

            return render_template('result.html', assignments=latest_assignments)

    return render_template('index.html')

@app.route('/download_excel')
def download_excel():
    # Specify the path to the generated Excel file
    excel_path = 'assignments.xlsx'
    
    # Send the file as an attachment
    return send_file(excel_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
