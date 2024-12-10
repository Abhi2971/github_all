import os
import pandas as pd

def process_student_files(model, repo_path):
    """
    Process all student folders and .java files in the given TeamTaskRepo path.
    Predict the question for each .java file and generate a consolidated report.
    """
    report_data = []

    # Traverse each student's folder
    for student_folder in os.listdir(repo_path):
        student_path = os.path.join(repo_path, student_folder)

        # Ensure it's a directory
        if os.path.isdir(student_path):
            print(f"Processing student folder: {student_folder}")

            # Find all .java files in the student folder
            for root, _, files in os.walk(student_path):
                for file in files:
                    if file.endswith('.java'):
                        file_path = os.path.join(root, file)
                        print(f"Analyzing file: {file_path}")

                        # Predict the question
                        try:
                            with open(file_path, 'r', encoding='utf-8') as f:
                                code = f.read()
                            predicted_question = model.predict([code])[0]
                            report_data.append({
                                'Student': student_folder,
                                'File': file,
                                'Predicted Question': predicted_question
                            })
                        except Exception as e:
                            print(f"Error analyzing file {file_path}: {e}")
    
    # Generate the Excel report
    report_df = pd.DataFrame(report_data)
    report_file = "TeamTaskRepo_Report.xlsx"
    report_df.to_excel(report_file, index=False)
    print(f"Report generated: {report_file}")

# Paths
repo_path = "./TeamTaskRepo"  # Replace with the actual path to TeamTaskRepo

# Process all students
process_student_files(model, repo_path)

