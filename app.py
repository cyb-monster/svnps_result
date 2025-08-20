import os
import pandas as pd
from flask import Flask, render_template, request

app = Flask(__name__)
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)



BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')

def list_excel_files():
    """List all Excel files in uploads folder."""
    files = [f for f in os.listdir(UPLOAD_FOLDER) if f.lower().endswith(('.xls', '.xlsx'))]
    return files

def get_classes(excel_file):
    """Return sheet names from the selected Excel file."""
    path = os.path.join(UPLOAD_FOLDER, excel_file)
    xls = pd.ExcelFile(path)
    return xls.sheet_names

def find_student_in_class(excel_file, class_name, roll_number):
    """Search roll number in selected Excel file and class sheet."""
    path = os.path.join(UPLOAD_FOLDER, excel_file)
    xls = pd.ExcelFile(path)

    if class_name not in xls.sheet_names:
        return None

    df = pd.read_excel(xls, sheet_name=class_name)
    df.columns = [col.strip() for col in df.columns]

    roll_col_candidates = [col for col in df.columns if 'Roll' in col]
    if not roll_col_candidates:
        return None

    roll_col = roll_col_candidates[0]
    student_row = df[df[roll_col] == roll_number]
    if not student_row.empty:
        return student_row.iloc[0].to_dict()

    return None

@app.route('/', methods=['GET', 'POST'])
def index():
    excel_files = list_excel_files()
    selected_excel = None
    classes = []
    selected_class = None
    roll_query = ''
    student_info = None
    message = None

    if request.method == 'POST':
        selected_excel = request.form.get('exam_select')
        selected_class = request.form.get('class_select')
        roll_query = request.form.get('roll_search', '').strip()

        if not selected_excel:
            message = "Please select the exam (Excel file)."
        else:
            classes = get_classes(selected_excel)

            if not selected_class:
                message = "Please select a class."
            elif not roll_query.isdigit():
                message = "Please enter a valid numeric roll number."
            else:
                student_info = find_student_in_class(selected_excel, selected_class, int(roll_query))
                if student_info is None:
                    message = f"Roll number {roll_query} not found in class {selected_class}."

    else:
        # GET request, no selections yet
        excel_files = list_excel_files()

    return render_template('index.html',
                           excel_files=excel_files,
                           selected_excel=selected_excel,
                           classes=classes,
                           selected_class=selected_class,
                           roll_query=roll_query,
                           student_info=student_info,
                           message=message)

if __name__ == '__main__':
    app.run(debug=True)
