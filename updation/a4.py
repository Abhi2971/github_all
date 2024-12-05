import os
import csv
from pathlib import Path

def analyze_java_code(filepath):
    """Generate a question based on the Java code."""
    with open(filepath, 'r') as file:
        code = file.read()

    if 'class ' in code:
        try:
            class_name = code.split('class ')[1].split()[0]
            question = f"What is the purpose of the class {class_name}?"
        except IndexError:
            question = "What is the purpose of this class?"
    elif 'public static void main' in code:
        question = "What is the purpose of the main method?"
    else:
        question = "Explain the functionality of this code snippet."
    
    return question

def process_student_folders(base_folder):
    """Process all student folders and generate a list of results."""
    results = []
    java_file_count = 0
    base_path = Path(base_folder)
    
    for student_folder in base_path.iterdir():
        if student_folder.is_dir():
            for java_file in student_folder.rglob('*.java'):  # Recursive search
                java_file_count += 1
                question = analyze_java_code(java_file)
                results.append({
                    'student': student_folder.name,
                    'file': java_file.name,
                    'question': question,
                })
    return results, java_file_count

def save_results_to_csv(results, output_file):
    """Save results to a CSV file."""
    with open(output_file, 'w', newline='') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=['Student Name', 'File Name', 'Question'])
        writer.writeheader()
        for row in results:
            writer.writerow({
                'Student Name': row['student'],
                'File Name': row['file'],
                'Question': row['question'],
            })

# Main function
if __name__ == "__main__":
    base_folder = "./TeamTaskRepo"  # Replace with the actual path
    output_file = "questions_report.csv"

    try:
        results, java_file_count = process_student_folders(base_folder)
        save_results_to_csv(results, output_file)
        print(f"Processed {java_file_count} .java files successfully.")
        print(f"Questions report saved to {output_file}")
    except Exception as e:
        print(f"An error occurred: {e}")

