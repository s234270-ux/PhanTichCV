import os
import io
import pyodbc
import pandas as pd
import pdfplumber
import spacy
# Đã thêm redirect và url_for vào đây:
from flask import Flask, render_template, request, send_file, redirect, url_for

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# --- CẤU HÌNH KẾT NỐI SQL SERVER ---
SERVER_NAME = r'LAPTOP-6C0OGPN9'
DATABASE_NAME = 'CV_Analyzer'
CONN_STR = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={SERVER_NAME};DATABASE={DATABASE_NAME};Trusted_Connection=yes;'


def get_db_connection():
    return pyodbc.connect(CONN_STR)


# --- TÍCH HỢP AI VÀ TỪ ĐIỂN ---
try:
    nlp = spacy.load("en_core_web_sm")
except:
    nlp = None

IT_SYNONYMS = {
    "js": "javascript", "react": "reactjs", "node": "nodejs",
    "vue": "vuejs", "postgres": "postgresql", "html5": "html", "css3": "css"
}


def extract_text_from_pdf(pdf_path):
    text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                extracted = page.extract_text()
                if extracted:
                    text += extracted + "\n"
    except Exception as e:
        print(f"Lỗi đọc PDF: {e}")
    return text


def analyze_cv(cv_text, jd_skills):
    if nlp:
        doc = nlp(cv_text.lower())
        cv_cleaned_text = " ".join([token.lemma_ for token in doc if not token.is_punct and not token.is_space])
    else:
        cv_cleaned_text = cv_text.lower()

    skills_list = [skill.strip().lower() for skill in jd_skills.split(',') if skill.strip()]
    found_skills = []
    missing_skills = []

    for skill in skills_list:
        standardized_skill = IT_SYNONYMS.get(skill, skill)
        if standardized_skill in cv_cleaned_text or skill in cv_text.lower():
            found_skills.append(skill.upper())
        else:
            missing_skills.append(skill.upper())

    total_skills = len(skills_list)
    match_score = int((len(found_skills) / total_skills) * 100) if total_skills > 0 else 0
    return match_score, found_skills, missing_skills


@app.route('/', methods=['GET', 'POST'])
def index():
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM candidates ORDER BY id DESC")
    history = cursor.fetchall()

    if request.method == 'POST':
        jd_skills = request.form.get('jd_skills')
        cv_files = request.files.getlist('cv_files')
        batch_results = []

        for cv_file in cv_files:
            if cv_file and cv_file.filename != '':
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], cv_file.filename)
                cv_file.save(filepath)

                cv_text = extract_text_from_pdf(filepath)
                match_score, found_skills, missing_skills = analyze_cv(cv_text, jd_skills)

                cursor.execute('''
                               INSERT INTO candidates (name, jd_skills, match_score, found_skills, missing_skills)
                               VALUES (?, ?, ?, ?, ?)
                               ''', (
                    cv_file.filename, jd_skills, match_score,
                    ", ".join(found_skills), ", ".join(missing_skills)
                ))

                batch_results.append({
                    "name": cv_file.filename,
                    "match_score": match_score,
                    "found_skills": found_skills,
                    "missing_skills": missing_skills
                })
                os.remove(filepath)

        conn.commit()
        batch_results.sort(key=lambda x: x['match_score'], reverse=True)

        cursor.execute("SELECT * FROM candidates ORDER BY id DESC")
        history = cursor.fetchall()
        conn.close()
        return render_template('index.html', batch_results=batch_results, jd_skills=jd_skills, history=history)

    conn.close()
    return render_template('index.html', batch_results=None, history=history)


# --- TÍNH NĂNG XUẤT FILE EXCEL ---
@app.route('/export')
def export_excel():
    try:
        candidate_id = request.args.get('id')
        conn = get_db_connection()

        if candidate_id:
            query = "SELECT id, name, jd_skills, match_score, found_skills, missing_skills FROM candidates WHERE id = ?"
            df = pd.read_sql(query, conn, params=(candidate_id,))
            file_name_export = f"Bao_Cao_{df.iloc[0]['name']}.xlsx"
        else:
            query = "SELECT id, name, jd_skills, match_score, found_skills, missing_skills FROM candidates ORDER BY id DESC"
            df = pd.read_sql(query, conn)
            file_name_export = "Bao_Cao_Tat_Ca_CV.xlsx"

        conn.close()

        if df.empty:
            return "Không có dữ liệu để xuất!"

        df.columns = ['ID', 'Tên File CV', 'Yêu cầu JD', 'Độ Phù Hợp (%)', 'Kỹ Năng Đạt', 'Kỹ Năng Thiếu']
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='LichSuDanhGiaCV')
        output.seek(0)

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=file_name_export
        )
    except Exception as e:
        return f"Lỗi khi xuất file: {e}"


# --- TÍNH NĂNG XÓA CV KHỎI DATABASE ---
@app.route('/delete/<int:id>')
def delete_cv(id):
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        # Chạy lệnh xóa row có ID tương ứng trong SQL Server
        cursor.execute("DELETE FROM candidates WHERE id = ?", (id,))
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"Lỗi khi xóa CSDL: {e}")
    # Quay trở lại trang chủ sau khi xóa xong
    return redirect(url_for('index'))


if __name__ == '__main__':
    app.run(debug=True)