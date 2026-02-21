from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import os
from main import StudentReportGenerator
from werkzeug.utils import secure_filename
import uuid

app = Flask(__name__)
app.secret_key = 'simple-key-123'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB

# Create upload folder
os.makedirs('uploads', exist_ok=True)

# Simple in-memory storage
processed_data = {}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'xlsx', 'xls'}

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('No file selected')
        return redirect(url_for('index'))
    
    file = request.files['file']
    
    if file.filename == '':
        flash('No file selected')
        return redirect(url_for('index'))
    
    if file and allowed_file(file.filename):
        # Save file
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        try:
            # Process using your code
            generator = StudentReportGenerator(filepath)
            generator.load_and_process_data()
            generator.clean_data()
            
            # Generate unique ID
            session_id = str(uuid.uuid4())
            
            # Store data
            processed_data[session_id] = {
                'generator': generator,
                'filename': filename,
                'summary': generator.get_class_summary(),
                'students': generator.get_students_list(),
                'chart': generator.generate_performance_chart()
            }
            
            return redirect(url_for('dashboard', session_id=session_id))
            
        except Exception as e:
            flash(f'Error: {str(e)}')
            return redirect(url_for('index'))
    
    flash('Invalid file type')
    return redirect(url_for('index'))

@app.route('/dashboard/<session_id>')
def dashboard(session_id):
    if session_id not in processed_data:
        flash('Session expired')
        return redirect(url_for('index'))
    
    data = processed_data[session_id]
    return render_template('dashboard.html',
                         summary=data['summary'],
                         students=data['students'],
                         chart=data['chart'],
                         filename=data['filename'],
                         session_id=session_id)

@app.route('/student/<session_id>/<int:student_id>')
def student_detail(session_id, student_id):
    if session_id not in processed_data:
        flash('Session expired')
        return redirect(url_for('index'))
    
    generator = processed_data[session_id]['generator']
    student = generator.get_student_details(student_id)
    
    # Generate all three charts
    subject_chart = generator.generate_student_chart(student['name'], student['subject_marks'])
    exam_trend_chart = generator.generate_student_exam_trend_chart(student['name'], student['exam_marks'])
    heatmap_chart = generator.generate_subject_exam_heatmap(student['name'], student['subject_exam_data'])
    
    return render_template('student.html',
                         student=student,
                         subject_chart=subject_chart,      # This is for the new chart
                         exam_trend_chart=exam_trend_chart,
                         heatmap_chart=heatmap_chart,
                         session_id=session_id)

@app.route('/download/pdf/<session_id>')
def download_pdf(session_id):
    if session_id not in processed_data:
        flash('Session expired')
        return redirect(url_for('index'))
    
    generator = processed_data[session_id]['generator']
    pdf_path = os.path.join('uploads', f'report_{session_id}.pdf')
    generator.create_pdf_report(pdf_path)
    
    return send_file(pdf_path, as_attachment=True, download_name='student_report.pdf')

@app.route('/download/excel/<session_id>')
def download_excel(session_id):
    if session_id not in processed_data:
        flash('Session expired')
        return redirect(url_for('index'))
    
    generator = processed_data[session_id]['generator']
    excel_path = os.path.join('uploads', f'data_{session_id}.xlsx')
    generator.export_to_excel(excel_path)
    
    return send_file(excel_path, as_attachment=True, download_name='processed_data.xlsx')

@app.route('/clear/<session_id>')
def clear_session(session_id):
    if session_id in processed_data:
        # Clean up files
        try:
            os.remove(os.path.join('uploads', f'report_{session_id}.pdf'))
            os.remove(os.path.join('uploads', f'data_{session_id}.xlsx'))
        except:
            pass
        del processed_data[session_id]
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)