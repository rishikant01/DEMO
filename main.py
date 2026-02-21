"""
Automated Student Performance Report Generation System
Specifically designed for Compilesheet IX 'D' (2025-26) format
with Subject names in Row 4 and Assessment types in Row 5
"""

import pandas as pd
import matplotlib.pyplot as plt
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER
import numpy as np
from io import BytesIO
import base64  # Make sure this is imported
import warnings
warnings.filterwarnings('ignore')

class StudentReportGenerator:
    def __init__(self, excel_file_path):
        """
        Initialize the report generator with Excel file path
        """
        self.excel_file = excel_file_path
        self.df = None
        self.processed_data = None
        self.subject_columns = {}  # Store column ranges for each subject
        self.performance_ranges = {
            '91-100%': 0,
            '80-90%': 0,
            '70-79%': 0,
            '60-69%': 0,
            '50-59%': 0,
            'Below 50%': 0
        }
        
    def load_and_process_data(self):
        """
        Load Excel file and process the specific structure:
        Row 4: Subject names
        Row 5: Assessment types (PA1, Half Yearly, PA2, Annual Exam, Total 200, Year's Avg 100)
        """
        print("📂 Loading Excel file with multi-level headers...")
        
        # Read the Excel file without headers first
        df_raw = pd.read_excel(self.excel_file, sheet_name=0, header=None)
        
        # Extract the header rows
        # Row 4 (index 3) contains subject names
        # Row 5 (index 4) contains assessment types
        subject_row = df_raw.iloc[3].fillna('')  # Subjects
        assessment_row = df_raw.iloc[4].fillna('')  # Assessment types
        
        print("\n📋 Detected Subjects and Assessments:")
        print("-" * 60)
        
        # Create multi-level columns
        multi_columns = []
        current_subject = ''
        subject_start_idx = {}
        
        for idx, (subject, assessment) in enumerate(zip(subject_row, assessment_row)):
            # If this is a new subject
            if pd.notna(subject) and subject != '' and 'Unnamed' not in str(subject):
                current_subject = subject
                subject_start_idx[current_subject] = idx
                print(f"\n📚 Subject: {current_subject}")
            
            # Create column name combining subject and assessment
            if current_subject and pd.notna(assessment) and assessment != '':
                col_name = f"{current_subject}_{assessment}"
                print(f"   - {assessment}")
            else:
                # For columns like S. No, Sch. No., Name, etc.
                col_name = subject if pd.notna(subject) else f"Column_{idx}"
            
            multi_columns.append(col_name)
        
        print("-" * 60)
        
        # Set the columns and remove header rows (first 5 rows)
        df = df_raw.iloc[5:].copy()
        df.columns = multi_columns
        
        # Reset index
        df = df.reset_index(drop=True)
        
        # Find the data rows (where S. No exists)
        # Look for column that might contain 'S. No'
        s_no_col = None
        for col in df.columns:
            if 'S. No' in str(col) or 'S.No' in str(col):
                s_no_col = col
                break
        
        if s_no_col:
            # Convert S. No to numeric, dropping rows where it's NaN
            df[s_no_col] = pd.to_numeric(df[s_no_col], errors='coerce')
            df = df[df[s_no_col].notna()].copy()
            df[s_no_col] = df[s_no_col].astype(int)
        
        # Rename important columns for easy access
        column_mapping = {}
        for col in df.columns:
            if 'Sch. No.' in str(col):
                column_mapping[col] = 'Sch_No'
            elif 'Name' in str(col):
                column_mapping[col] = 'Name'
            elif 'S. No' in str(col):
                column_mapping[col] = 'S_No'
        
        df.rename(columns=column_mapping, inplace=True)
        
        self.df = df
        print(f"\n✅ Loaded {len(df)} student records")
        
        # Display first few rows to verify
        print("\n📊 First few records:")
        print(df[['S_No', 'Sch_No', 'Name']].head() if 'Name' in df.columns else df.head())
        
        return df
    
    def clean_data(self):
        """
        Clean the data, handle 'AB' (Absent) values and calculate totals
        """
        print("\n🧹 Cleaning and processing data...")
        
        # Replace 'AB' with 0 and convert to numeric for all subject columns
        for col in self.df.columns:
            if col not in ['S_No', 'Sch_No', 'Name'] and '_' in str(col):
                self.df[col] = pd.to_numeric(
                    self.df[col].astype(str).str.replace('AB', '0').str.strip(), 
                    errors='coerce'
                ).fillna(0)
        
        # Group columns by subject
        subjects = ['English', 'Hindi', 'Maths', 'Science', 'Social Science', 'Computer']
        
        for subject in subjects:
            # Find all columns for this subject
            subject_cols = [col for col in self.df.columns if subject in str(col)]
            
            if subject_cols:
                # Find the 'Year's Avg' column for this subject (marks out of 100)
                year_avg_cols = [col for col in subject_cols if 'Year' in str(col) or 'Avg' in str(col)]
                
                if year_avg_cols:
                    self.subject_columns[f"{subject}_Year_Avg"] = year_avg_cols[0]
                
                # Find the 'Total 200' column
                total_200_cols = [col for col in subject_cols if 'Total 200' in str(col)]
                if total_200_cols:
                    self.subject_columns[f"{subject}_Total_200"] = total_200_cols[0]
        
        # Calculate final total (sum of all subject Year's Avg)
        year_avg_cols = [col for subject, col in self.subject_columns.items() if 'Year_Avg' in subject]
        if year_avg_cols:
            self.df['Final_Total'] = self.df[year_avg_cols].sum(axis=1)
        else:
            # Fallback: Try to find the Final Assessment Total 600 column
            final_total_cols = [col for col in self.df.columns if 'Total 600' in str(col)]
            if final_total_cols:
                self.df['Final_Total'] = self.df[final_total_cols[0]]
        
        # Calculate percentage
        self.df['Percentage'] = (self.df['Final_Total'] / 600) * 100
        
        # Categorize performance
        self.df['Performance_Category'] = pd.cut(
            self.df['Percentage'],
            bins=[0, 50, 60, 70, 80, 91, 100],
            labels=['Below 50%', '50-59%', '60-69%', '70-79%', '80-90%', '91-100%'],
            right=False
        )
        
        # Update performance ranges
        for category in self.df['Performance_Category'].dropna():
            self.performance_ranges[str(category)] += 1
        
        print("✅ Data cleaning complete")
        print(f"📈 Average Percentage: {self.df['Percentage'].mean():.2f}%")
        
        return self.df
    
    def extract_subject_marks(self, student_row):
        """
        Extract subject-wise marks (Year's Avg out of 100) for a student
        """
        subjects = ['English', 'Hindi', 'Maths', 'Science', 'Social Science', 'Computer']
        subject_marks = {}
        
        for subject in subjects:
            # Look for Year's Avg column for this subject
            year_avg_key = f"{subject}_Year_Avg"
            
            if year_avg_key in self.subject_columns:
                col_name = self.subject_columns[year_avg_key]
                marks = student_row[col_name]
            else:
                # Fallback: search for any column with subject and Year/Avg
                subject_cols = [col for col in self.df.columns 
                              if subject in str(col) and ('Year' in str(col) or 'Avg' in str(col))]
                if subject_cols:
                    marks = student_row[subject_cols[0]]
                else:
                    marks = 0
            
            subject_marks[subject] = float(marks) if pd.notna(marks) else 0
        
        return subject_marks
    
    def generate_performance_pie_chart(self):
        """
        Generate pie chart showing performance distribution
        """
        print("📊 Generating performance pie chart...")
        
        # Filter out empty categories
        labels = []
        sizes = []
        colors_list = ['#ff9999', '#66b3ff', '#99ff99', '#ffcc99', '#c2c2f0', '#ffb3e6']
        
        for range_name, count in self.performance_ranges.items():
            if count > 0:
                percentage = (count / len(self.df)) * 100
                labels.append(f"{range_name}\n({count} students, {percentage:.1f}%)")
                sizes.append(count)
        
        plt.figure(figsize=(12, 8))
        wedges, texts, autotexts = plt.pie(
            sizes, 
            labels=labels, 
            colors=colors_list[:len(sizes)], 
            autopct='',  # We'll add custom labels
            startangle=90, 
            textprops={'fontsize': 10},
            pctdistance=0.85
        )
        
        # Add percentage inside pie
        for i, (wedge, size) in enumerate(zip(wedges, sizes)):
            percentage = (size / len(self.df)) * 100
            ang = (wedge.theta2 + wedge.theta1) / 2
            x = 0.7 * np.cos(np.radians(ang))
            y = 0.7 * np.sin(np.radians(ang))
            plt.text(x, y, f'{percentage:.1f}%', 
                    ha='center', va='center', fontsize=10, fontweight='bold')
        
        plt.title('Student Performance Distribution', fontsize=16, fontweight='bold', pad=20)
        plt.axis('equal')
        
        # Add legend
        plt.legend(wedges, [f"{label}" for label in labels], 
                  title="Performance Ranges", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
        
        plt.tight_layout()
        
        # Save to BytesIO
        img_data = BytesIO()
        plt.savefig(img_data, format='png', dpi=150, bbox_inches='tight')
        plt.close()
        img_data.seek(0)
        
        return img_data
    
    def generate_student_progress_graph(self, student_name, subject_marks):
        """
        Generate bar chart for individual student subject-wise performance
        """
        subjects = list(subject_marks.keys())
        marks = list(subject_marks.values())
        
        plt.figure(figsize=(12, 6))
        
        # Create bar chart
        x_pos = np.arange(len(subjects))
        bars = plt.bar(x_pos, marks, color='skyblue', edgecolor='navy', alpha=0.7, width=0.6)
        
        # Color bars based on performance
        for bar, mark in zip(bars, marks):
            if mark >= 75:
                bar.set_color('lightgreen')
            elif mark >= 60:
                bar.set_color('skyblue')
            elif mark >= 45:
                bar.set_color('gold')
            else:
                bar.set_color('salmon')
        
        # Add value labels on bars
        for bar, mark in zip(bars, marks):
            height = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2., height + 1,
                    f'{mark:.0f}', ha='center', va='bottom', fontsize=11, fontweight='bold')
        
        # Add trend line
        plt.plot(x_pos, marks, 'ro-', linewidth=2, markersize=8, label='Performance Trend')
        
        plt.xlabel('Subjects', fontsize=13, fontweight='bold')
        plt.ylabel('Marks (out of 100)', fontsize=13, fontweight='bold')
        plt.title(f'Subject-wise Performance: {student_name}', fontsize=14, fontweight='bold')
        plt.xticks(x_pos, subjects, rotation=45, ha='right', fontsize=11)
        plt.yticks(fontsize=11)
        plt.ylim(0, 105)
        plt.grid(True, alpha=0.3, linestyle='--', axis='y')
        plt.axhline(y=33, color='red', linestyle='--', linewidth=2, alpha=0.7, label='Pass Mark (33%)')
        plt.legend(loc='upper right')
        
        plt.tight_layout()
        
        # Save to BytesIO
        img_data = BytesIO()
        plt.savefig(img_data, format='png', dpi=150, bbox_inches='tight')
        plt.close()
        img_data.seek(0)
        
        return img_data
    
    def create_pdf_report(self, output_filename="Student_Performance_Report.pdf"):
        """
        Create comprehensive PDF report with half-year tracking
        """
        print(f"\n📄 Generating PDF report: {output_filename}")
        
        doc = SimpleDocTemplate(output_filename, pagesize=A4, 
                            rightMargin=72, leftMargin=72,
                            topMargin=72, bottomMargin=72)
        
        styles = getSampleStyleSheet()
        styles.add(ParagraphStyle(
            name='CenterTitle',
            parent=styles['Heading1'],
            alignment=TA_CENTER,
            spaceAfter=30,
            fontSize=18,
            textColor=colors.HexColor('#2E4057')
        ))
        styles.add(ParagraphStyle(
            name='CenterSubTitle',
            parent=styles['Heading2'],
            alignment=TA_CENTER,
            spaceAfter=20,
            fontSize=14,
            textColor=colors.HexColor('#4A6FA5')
        ))
        
        story = []
        
        # Title Page
        story.append(Paragraph("CENTRAL ACADEMY ENGLISH MEDIUM SCHOOL", styles['CenterTitle']))
        story.append(Paragraph("Vijay Nagar, Jabalpur", styles['CenterSubTitle']))
        story.append(Spacer(1, 0.3*inch))
        story.append(Paragraph("Academic Year: 2025-26", styles['CenterSubTitle']))
        story.append(Spacer(1, 0.5*inch))
        story.append(Paragraph("STUDENT PERFORMANCE ANALYSIS REPORT", styles['CenterTitle']))
        story.append(Spacer(1, 0.3*inch))
        story.append(Paragraph(f"Generated on: {pd.Timestamp.now().strftime('%d-%m-%Y')}", 
                            styles['Normal']))
        story.append(PageBreak())
        
        # Class Overview Section
        story.append(Paragraph("CLASS PERFORMANCE OVERVIEW", styles['Heading1']))
        story.append(Spacer(1, 0.2*inch))
        
        # Summary Statistics
        total_students = len(self.df)
        pass_count = len(self.df[self.df['Percentage'] >= 33])
        fail_count = total_students - pass_count
        avg_percentage = self.df['Percentage'].mean()
        max_percentage = self.df['Percentage'].max()
        min_percentage = self.df['Percentage'].min()
        
        # Calculate class half-year averages
        first_half_totals = []
        second_half_totals = []
        for _, student in self.df.iterrows():
            exam_marks = self.get_student_exam_wise_marks(student)
            first_half_totals.append(exam_marks['first_half']['percentage'])
            second_half_totals.append(exam_marks['second_half']['percentage'])
        
        avg_first_half = sum(first_half_totals) / len(first_half_totals) if first_half_totals else 0
        avg_second_half = sum(second_half_totals) / len(second_half_totals) if second_half_totals else 0
        
        # Create summary table
        summary_data = [
            ['Metric', 'Value'],
            ['Total Students', str(total_students)],
            ['Students Passed', f"{pass_count} ({pass_count/total_students*100:.1f}%)"],
            ['Students Failed', f"{fail_count} ({fail_count/total_students*100:.1f}%)"],
            ['Average Percentage', f"{avg_percentage:.2f}%"],
            ['Highest Percentage', f"{max_percentage:.2f}%"],
            ['Lowest Percentage', f"{min_percentage:.2f}%"],
            ['', ''],
            ['HALF-YEAR ANALYSIS', ''],
            ['Average First Half %', f"{avg_first_half:.2f}%"],
            ['Average Second Half %', f"{avg_second_half:.2f}%"],
            ['Improvement', f"{avg_second_half - avg_first_half:+.2f}%"]
        ]
        
        summary_table = Table(summary_data, colWidths=[3*inch, 3*inch])
        summary_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 12),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4A6FA5')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('BACKGROUND', (0, 1), (-1, 7), colors.HexColor('#F5F5F5')),
            ('BACKGROUND', (0, 8), (-1, -1), colors.HexColor('#E8F0FE')),
            ('SPAN', (0, 8), (1, 8)),  # Merge cells for "HALF-YEAR ANALYSIS" header
            ('ALIGN', (0, 8), (-1, 8), 'CENTER'),
            ('FONTNAME', (0, 8), (-1, 8), 'Helvetica-Bold'),
            ('TEXTCOLOR', (0, 8), (-1, 8), colors.HexColor('#4A6FA5')),
            ('PADDING', (0, 0), (-1, -1), 8),
        ]))
        
        story.append(summary_table)
        story.append(Spacer(1, 0.3*inch))
        
        # Performance Distribution Pie Chart
        story.append(Paragraph("Performance Distribution", styles['Heading2']))
        story.append(Spacer(1, 0.1*inch))
        
        pie_chart = self.generate_performance_pie_chart()
        img = Image(pie_chart, width=6*inch, height=4.5*inch)
        story.append(img)
        story.append(PageBreak())
        
        # Top Performers Section
        story.append(Paragraph("TOP PERFORMERS", styles['Heading1']))
        story.append(Spacer(1, 0.2*inch))
        
        top_students = self.df.nlargest(5, 'Percentage')[['Name', 'Sch_No', 'Percentage', 'Final_Total']]
        
        top_data = [['Rank', 'Name', 'Sch. No.', 'Percentage', 'Total Marks']]
        for idx, (_, student) in enumerate(top_students.iterrows(), 1):
            top_data.append([
                str(idx),
                student['Name'],
                str(student['Sch_No']),
                f"{student['Percentage']:.2f}%",
                f"{student['Final_Total']:.0f}/600"
            ])
        
        top_table = Table(top_data, colWidths=[0.8*inch, 2.2*inch, 1.2*inch, 1.5*inch, 1.5*inch])
        top_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 11),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4A6FA5')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#F5F5F5')),
            ('PADDING', (0, 0), (-1, -1), 6),
        ]))
        
        story.append(top_table)
        story.append(PageBreak())
        
        # Individual Student Reports
        story.append(Paragraph("INDIVIDUAL STUDENT ANALYSIS", styles['Heading1']))
        story.append(Spacer(1, 0.2*inch))
        
        for idx, student in self.df.iterrows():
            student_name = student['Name']
            sch_no = student['Sch_No']
            
            # Get exam-wise data for this student
            exam_marks = self.get_student_exam_wise_marks(student)
            subject_marks = self.extract_subject_marks(student)
            
            story.append(Paragraph(f"{student_name} (Sch. No: {sch_no})", styles['Heading2']))
            
            # Student summary with half-year data
            student_data = [
                ['Detail', 'Value'],
                ['Total Marks (out of 600)', f"{student['Final_Total']:.0f}"],
                ['Percentage', f"{student['Percentage']:.2f}%"],
                ['Performance Category', student['Performance_Category']],
                ['Result', '✅ PASS' if student['Percentage'] >= 33 else '❌ FAIL'],
                ['', ''],
                ['HALF-YEAR BREAKDOWN', ''],
                ['First Half Total (PA1 + Half Yearly)', f"{exam_marks['first_half']['total']:.0f}/600"],
                ['First Half Percentage', f"{exam_marks['first_half']['percentage']}%"],
                ['Second Half Total (PA2 + Annual Exam)', f"{exam_marks['second_half']['total']:.0f}/600"],
                ['Second Half Percentage', f"{exam_marks['second_half']['percentage']}%"],
                ['Year Average', f"{exam_marks['year_average']}%"],
                ['Improvement', f"{exam_marks['second_half']['percentage'] - exam_marks['first_half']['percentage']:+.2f}%"]
            ]
            
            student_table = Table(student_data, colWidths=[2.5*inch, 2.5*inch])
            student_table.setStyle(TableStyle([
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4A6FA5')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('BACKGROUND', (0, 1), (-1, 4), colors.HexColor('#F5F5F5')),
                ('BACKGROUND', (0, 5), (-1, 5), colors.HexColor('#E8F0FE')),  # Separator
                ('BACKGROUND', (0, 6), (-1, 6), colors.HexColor('#4A6FA5')),  # Half-year header
                ('TEXTCOLOR', (0, 6), (-1, 6), colors.white),
                ('BACKGROUND', (0, 7), (-1, -1), colors.HexColor('#F0F7FF')),
                ('TEXTCOLOR', (-1, -1), (-1, -1), 
                colors.green if student['Percentage'] >= 33 else colors.red),
                ('PADDING', (0, 0), (-1, -1), 6),
            ]))
            
            story.append(student_table)
            story.append(Spacer(1, 0.1*inch))
            
            # Exam-wise marks table
            exam_data = [['Examination', 'Total Marks (out of 600)', 'Percentage']]
            for exam, marks in exam_marks['individual'].items():
                percentage = (marks / 600) * 100
                exam_data.append([
                    exam,
                    f"{marks:.0f}",
                    f"{percentage:.1f}%"
                ])
            
            # Add half-year summary rows
            exam_data.append(['', '', ''])
            exam_data.append(['FIRST HALF TOTAL', f"{exam_marks['first_half']['total']:.0f}", 
                            f"{exam_marks['first_half']['percentage']}%"])
            exam_data.append(['SECOND HALF TOTAL', f"{exam_marks['second_half']['total']:.0f}", 
                            f"{exam_marks['second_half']['percentage']}%"])
            exam_data.append(['YEAR AVERAGE', '', f"{exam_marks['year_average']}%"])
            
            exam_table = Table(exam_data, colWidths=[2*inch, 1.5*inch, 1.5*inch])
            exam_table.setStyle(TableStyle([
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
                ('GRID', (0, 0), (-1, -3), 1, colors.black),  # Grid for exam rows
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4A6FA5')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('BACKGROUND', (0, 1), (-1, 4), colors.HexColor('#F5F5F5')),  # Exam rows
                ('BACKGROUND', (0, 5), (-1, 5), colors.HexColor('#E8F0FE')),  # Separator
                ('BACKGROUND', (0, 6), (-1, 6), colors.HexColor('#4CAF50')),  # First half
                ('BACKGROUND', (0, 7), (-1, 7), colors.HexColor('#FF9800')),  # Second half
                ('BACKGROUND', (0, 8), (-1, 8), colors.HexColor('#667EEA')),  # Year average
                ('TEXTCOLOR', (0, 6), (-1, 8), colors.white),
                ('ALIGN', (1, 0), (2, -1), 'CENTER'),
                ('PADDING', (0, 0), (-1, -1), 6),
            ]))
            
            story.append(exam_table)
            story.append(Spacer(1, 0.2*inch))
            
            # Subject-wise progress graph
            progress_graph = self.generate_student_progress_graph(student_name, subject_marks)
            img = Image(progress_graph, width=6*inch, height=3.5*inch)
            story.append(img)
            
            # Subject marks table with exam breakdown
            subject_data = [['Subject', 'Year Avg', 'PA 1', 'Half Yearly', 'PA 2', 'Annual Exam']]
            subject_exam_data = self.get_student_subject_exam_wise(student)
            
            for subject in ['English', 'Hindi', 'Maths', 'Science', 'Social Science', 'Computer']:
                subject_data.append([
                    subject,
                    f"{subject_marks[subject]:.0f}",
                    f"{subject_exam_data[subject]['PA 1']:.0f}",
                    f"{subject_exam_data[subject]['Half Yearly']:.0f}",
                    f"{subject_exam_data[subject]['PA 2']:.0f}",
                    f"{subject_exam_data[subject]['Annual Exam']:.0f}"
                ])
            
            subject_table = Table(subject_data, colWidths=[1.5*inch, 0.8*inch, 0.8*inch, 1*inch, 0.8*inch, 1*inch])
            subject_table.setStyle(TableStyle([
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 9),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4A6FA5')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#F5F5F5')),
                ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
                ('PADDING', (0, 0), (-1, -1), 4),
            ]))
            
            story.append(subject_table)
            story.append(Spacer(1, 0.2*inch))
            
            # Add page break between students (except last)
            if idx < len(self.df) - 1:
                story.append(PageBreak())
        
        # Build PDF
        doc.build(story)
        print(f"✅ PDF report generated successfully: {output_filename}")
        return output_filename
    def generate_summary_report(self):
        """
        Generate a text summary report
        """
        print("\n" + "="*70)
        print("📊 STUDENT PERFORMANCE SUMMARY REPORT")
        print("="*70)
        print(f"📚 Total Students: {len(self.df)}")
        print(f"📈 Average Percentage: {self.df['Percentage'].mean():.2f}%")
        print(f"🏆 Highest Percentage: {self.df['Percentage'].max():.2f}%")
        print(f"📉 Lowest Percentage: {self.df['Percentage'].min():.2f}%")
        
        pass_count = len(self.df[self.df['Percentage'] >= 33])
        fail_count = len(self.df) - pass_count
        print(f"✅ Pass: {pass_count} ({pass_count/len(self.df)*100:.1f}%)")
        print(f"❌ Fail: {fail_count} ({fail_count/len(self.df)*100:.1f}%)")
        
        print("\n📊 Performance Distribution:")
        print("-"*50)
        
        for range_name, count in self.performance_ranges.items():
            if count > 0:
                percentage = (count / len(self.df)) * 100
                bar = '█' * int(percentage/2) + '░' * (50 - int(percentage/2))
                print(f"{range_name:12}: {count:2} students ({percentage:5.1f}%) {bar}")
        
        print("\n🏅 Top 5 Students:")
        print("-"*50)
        top_students = self.df.nlargest(5, 'Percentage')[['Name', 'Sch_No', 'Percentage', 'Final_Total']]
        for idx, student in top_students.iterrows():
            print(f"   {student['Name']:25} : {student['Percentage']:.2f}% (Total: {student['Final_Total']:.0f}/600)")
        
        print("="*70)
    
    def export_to_excel(self, output_filename="processed_student_data.xlsx"):
        """
        Export processed data to Excel
        """
        export_df = self.df[['S_No', 'Sch_No', 'Name', 'Final_Total', 'Percentage', 'Performance_Category']].copy()
        export_df.to_excel(output_filename, index=False)
        print(f"✅ Processed data exported to: {output_filename}")
        return output_filename

    # ==================== NEW METHODS FOR WEB INTERFACE ====================
    
    def get_class_summary(self):
        """Get class summary statistics for web display"""
        total_students = len(self.df)
        pass_count = len(self.df[self.df['Percentage'] >= 33])
        
        return {
            'total_students': total_students,
            'pass_count': pass_count,
            'fail_count': total_students - pass_count,
            'pass_percentage': round((pass_count/total_students*100), 2),
            'avg_percentage': round(self.df['Percentage'].mean(), 2),
            'max_percentage': round(self.df['Percentage'].max(), 2),
            'min_percentage': round(self.df['Percentage'].min(), 2)
        }

    def get_students_list(self):
        """Get list of all students for web display"""
        students = []
        for _, row in self.df.iterrows():
            students.append({
                's_no': int(row['S_No']),
                'sch_no': str(row['Sch_No']),
                'name': row['Name'],
                'percentage': round(row['Percentage'], 2),
                'total': int(row['Final_Total']),
                'category': row['Performance_Category'],
                'result': 'PASS' if row['Percentage'] >= 33 else 'FAIL'
            })
        return students

    def get_student_details(self, student_id):
        """Get details for a specific student including exam-wise and half-year data"""
        student = self.df[self.df['S_No'] == student_id].iloc[0]
        
        subjects = ['English', 'Hindi', 'Maths', 'Science', 'Social Science', 'Computer']
        subject_marks = {}
        
        for subject in subjects:
            # Year's Avg is out of 100
            year_avg_key = f"{subject}_Year_Avg"
            if year_avg_key in self.subject_columns:
                col_name = self.subject_columns[year_avg_key]
                marks = student[col_name]
            else:
                subject_cols = [col for col in self.df.columns 
                            if subject in str(col) and ('Year' in str(col) or 'Avg' in str(col))]
                marks = student[subject_cols[0]] if subject_cols else 0
            
            subject_marks[subject] = round(float(marks) if pd.notna(marks) else 0, 2)
        
        # Get exam-wise data with half-year calculations
        exam_marks = self.get_student_exam_wise_marks(student)
        subject_exam_data = self.get_student_subject_exam_wise(student)
        
        # Calculate total of subject year averages (should match Final_Total)
        total_year_avg = sum(subject_marks.values())
        
        return {
            's_no': int(student['S_No']),
            'sch_no': str(student['Sch_No']),
            'name': student['Name'],
            'total': int(student['Final_Total']),  # This is sum of Year's Avg (out of 600)
            'total_subject_marks': total_year_avg,  # Should match total
            'percentage': round(student['Percentage'], 2),
            'category': student['Performance_Category'],
            'result': 'PASS' if student['Percentage'] >= 33 else 'FAIL',
            'subject_marks': subject_marks,  # Each out of 100
            'subject_totals': exam_marks['subject_totals'],  # Each out of 200
            'exam_marks': exam_marks,
            'subject_exam_data': subject_exam_data
        }
    def generate_performance_chart(self):
        """Generate pie chart as base64 for web display"""
        plt.figure(figsize=(8, 6))
        
        labels = []
        sizes = []
        colors = ['#ff9999', '#66b3ff', '#99ff99', '#ffcc99', '#c2c2f0', '#ffb3e6']
        
        for range_name, count in self.performance_ranges.items():
            if count > 0:
                labels.append(f"{range_name}")
                sizes.append(count)
        
        plt.pie(sizes, labels=labels, colors=colors[:len(sizes)], 
                autopct='%1.1f%%', startangle=90)
        plt.title('Performance Distribution', fontsize=14, fontweight='bold')
        plt.axis('equal')
        
        img_data = BytesIO()
        plt.savefig(img_data, format='png', dpi=100, bbox_inches='tight')
        plt.close()
        img_data.seek(0)
        
        return base64.b64encode(img_data.getvalue()).decode()

    def generate_student_chart(self, student_name, subject_marks):
        """Generate bar chart as base64 for web display"""
        plt.figure(figsize=(10, 5))
        
        subjects = list(subject_marks.keys())
        marks = list(subject_marks.values())
        
        bars = plt.bar(subjects, marks, color='skyblue', edgecolor='navy', alpha=0.7)
        
        # Color bars based on performance
        for bar, mark in zip(bars, marks):
            if mark >= 75:
                bar.set_color('lightgreen')
            elif mark >= 60:
                bar.set_color('skyblue')
            elif mark >= 45:
                bar.set_color('gold')
            else:
                bar.set_color('salmon')
        
        # Add value labels
        for bar, mark in zip(bars, marks):
            height = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2., height + 1,
                    f'{mark:.0f}', ha='center', va='bottom', fontsize=10)
        
        plt.xlabel('Subjects')
        plt.ylabel('Marks (out of 100)')
        plt.title(f'Performance: {student_name}')
        plt.xticks(rotation=45, ha='right')
        plt.ylim(0, 105)
        plt.grid(True, alpha=0.3, axis='y')
        plt.axhline(y=33, color='red', linestyle='--', alpha=0.7)
        
        plt.tight_layout()
        
        img_data = BytesIO()
        plt.savefig(img_data, format='png', dpi=100, bbox_inches='tight')
        plt.close()
        img_data.seek(0)
        
        return base64.b64encode(img_data.getvalue()).decode()
    def get_student_exam_wise_marks(self, student_row):
        """
        Extract exam-wise marks and calculate half-yearly totals
        Each subject total is out of 200 (across all exams)
        Each exam has different weightage:
        - PA1: 20 marks per subject
        - Half Yearly: 80 marks per subject  
        - PA2: 20 marks per subject
        - Annual Exam: 80 marks per subject
        """
        # Individual exam marks (total across all 6 subjects)
        exams = {
            'PA 1': 0,
            'Half Yearly': 0,
            'PA 2': 0,
            'Annual Exam': 0
        }
        
        subjects = ['English', 'Hindi', 'Maths', 'Science', 'Social Science', 'Computer']
        n_subjects = len(subjects)
        
        # Get individual exam marks (sum across all subjects)
        for exam in exams.keys():
            total = 0
            for subject in subjects:
                exam_cols = [col for col in self.df.columns 
                            if subject in str(col) and exam in str(col)]
                if exam_cols:
                    marks = student_row[exam_cols[0]]
                    total += float(marks) if pd.notna(marks) else 0
            exams[exam] = round(total, 2)
        
        # Max possible marks for each exam type (across all subjects)
        pa_max = 20 * n_subjects  # 120 total for PA1 and PA2
        exam_max = 80 * n_subjects  # 480 total for Half Yearly and Annual Exam
        
        # Calculate half-year totals
        first_half_total = exams['PA 1'] + exams['Half Yearly']
        second_half_total = exams['PA 2'] + exams['Annual Exam']
        
        # Max possible for each half (across all subjects)
        half_max = pa_max + exam_max  # 120 + 480 = 600
        
        # Calculate percentages
        first_half_percentage = (first_half_total / half_max) * 100
        second_half_percentage = (second_half_total / half_max) * 100
        
        # Year average (average of both halves)
        year_average = (first_half_percentage + second_half_percentage) / 2
        
        # Calculate subject-wise totals (each subject out of 200)
        subject_totals = {}
        for subject in subjects:
            subject_total = 0
            for exam in exams.keys():
                exam_cols = [col for col in self.df.columns 
                            if subject in str(col) and exam in str(col)]
                if exam_cols:
                    marks = student_row[exam_cols[0]]
                    subject_total += float(marks) if pd.notna(marks) else 0
            subject_totals[subject] = round(subject_total, 2)
        
        return {
            'individual': exams,
            'max_marks': {
                'PA 1': pa_max,
                'PA 2': pa_max,
                'Half Yearly': exam_max,
                'Annual Exam': exam_max
            },
            'first_half': {
                'total': first_half_total,
                'max': half_max,
                'percentage': round(first_half_percentage, 2)
            },
            'second_half': {
                'total': second_half_total,
                'max': half_max,
                'percentage': round(second_half_percentage, 2)
            },
            'year_average': round(year_average, 2),
            'subject_totals': subject_totals  # Each subject out of 200
        }
    def get_student_subject_exam_wise(self, student_row):
        """
        Get exam-wise marks for each subject
        Returns: Dictionary with subjects and their exam-wise marks
        """
        exams = ['PA 1', 'Half Yearly', 'PA 2', 'Annual Exam']
        subjects = ['English', 'Hindi', 'Maths', 'Science', 'Social Science', 'Computer']
        
        result = {}
        for subject in subjects:
            subject_data = {}
            for exam in exams:
                exam_cols = [col for col in self.df.columns 
                            if subject in str(col) and exam in str(col)]
                if exam_cols:
                    marks = student_row[exam_cols[0]]
                    subject_data[exam] = round(float(marks) if pd.notna(marks) else 0, 2)
                else:
                    subject_data[exam] = 0
            result[subject] = subject_data
        
        return result

    def generate_student_exam_trend_chart(self, student_name, exam_marks):
        """
        Generate line chart showing student's performance across exams and halves
        """
        plt.figure(figsize=(12, 6))
        
        # Individual exams data
        individual = exam_marks['individual']
        exams = list(individual.keys())
        marks = list(individual.values())
        
        # Create line chart for individual exams
        plt.plot(exams, marks, 'bo-', linewidth=2, markersize=8, 
                markerfacecolor='blue', label='Individual Exams', alpha=0.7)
        
        # Add value labels for individual exams
        for i, (exam, mark) in enumerate(zip(exams, marks)):
            plt.annotate(f'{mark}', (exam, mark), textcoords="offset points", 
                        xytext=(0,10), ha='center', fontsize=9)
        
        # Add half-year averages as horizontal lines
        first_half_avg = exam_marks['first_half']['total']
        second_half_avg = exam_marks['second_half']['total']
        
        plt.axhline(y=first_half_avg, color='green', linestyle='--', linewidth=2, 
                    alpha=0.7, label=f"First Half Avg: {first_half_avg:.0f}")
        plt.axhline(y=second_half_avg, color='orange', linestyle='--', linewidth=2, 
                    alpha=0.7, label=f"Second Half Avg: {second_half_avg:.0f}")
        
        # Add year average line
        plt.axhline(y=exam_marks['year_average'] * 6, color='purple', linestyle='-', 
                    linewidth=2, alpha=0.7, label=f"Year Avg: {exam_marks['year_average']:.1f}%")
        
        # Add threshold line for pass
        pass_mark = (33/100) * 600
        plt.axhline(y=pass_mark, color='red', linestyle=':', linewidth=2, 
                    alpha=0.7, label=f'Pass Mark ({pass_mark:.0f})')
        
        plt.xlabel('Examinations', fontsize=12, fontweight='bold')
        plt.ylabel('Total Marks (out of 600)', fontsize=12, fontweight='bold')
        plt.title(f'Exam-wise Performance Trend: {student_name}', fontsize=14, fontweight='bold')
        plt.grid(True, alpha=0.3, linestyle='--')
        plt.legend(loc='best')
        
        # Set y-axis limit
        max_mark = max(marks) if marks else 0
        plt.ylim(0, min(650, max_mark + 100))
        
        plt.tight_layout()
        
        img_data = BytesIO()
        plt.savefig(img_data, format='png', dpi=100, bbox_inches='tight')
        plt.close()
        img_data.seek(0)
        
        return base64.b64encode(img_data.getvalue()).decode()

    def generate_subject_exam_heatmap(self, student_name, subject_exam_data):
        """
        Generate a heatmap-style visualization of subject-wise exam performance
        """
        subjects = list(subject_exam_data.keys())
        exams = ['PA 1', 'Half Yearly', 'PA 2', 'Annual Exam']
        
        # Create data matrix
        data = []
        for subject in subjects:
            row = [subject_exam_data[subject][exam] for exam in exams]
            data.append(row)
        
        plt.figure(figsize=(12, 6))
        
        # Create heatmap using imshow
        im = plt.imshow(data, cmap='YlOrRd', aspect='auto', vmin=0, vmax=100)
        
        # Add text annotations
        for i in range(len(subjects)):
            for j in range(len(exams)):
                text = plt.text(j, i, f'{data[i][j]:.0f}',
                            ha="center", va="center", color="black", fontweight='bold')
        
        # Customize axes
        plt.xticks(range(len(exams)), exams, rotation=45, ha='right')
        plt.yticks(range(len(subjects)), subjects)
        plt.title(f'Subject-wise Exam Performance: {student_name}', fontsize=14, fontweight='bold')
        
        # Add colorbar
        plt.colorbar(im, label='Marks')
        
        plt.tight_layout()
        
        img_data = BytesIO()
        plt.savefig(img_data, format='png', dpi=100, bbox_inches='tight')
        plt.close()
        img_data.seek(0)
        
        return base64.b64encode(img_data.getvalue()).decode()

    def get_class_exam_averages(self):
        """
        Calculate class average for each exam
        """
        exams = ['PA 1', 'Half Yearly', 'PA 2', 'Annual Exam']
        subjects = ['English', 'Hindi', 'Maths', 'Science', 'Social Science', 'Computer']
        
        exam_averages = {exam: {'total': 0, 'subjects': {}} for exam in exams}
        
        for exam in exams:
            subject_totals = {subject: [] for subject in subjects}
            
            for _, student in self.df.iterrows():
                for subject in subjects:
                    exam_cols = [col for col in self.df.columns 
                                if subject in str(col) and exam in str(col)]
                    if exam_cols:
                        marks = student[exam_cols[0]]
                        if pd.notna(marks):
                            subject_totals[subject].append(float(marks))
            
            # Calculate averages
            total_sum = 0
            for subject in subjects:
                if subject_totals[subject]:
                    avg = sum(subject_totals[subject]) / len(subject_totals[subject])
                    exam_averages[exam]['subjects'][subject] = round(avg, 2)
                    total_sum += avg * len(subject_totals[subject])
            
            # Overall exam average
            total_students = len(self.df)
            if total_students > 0:
                exam_averages[exam]['total'] = round(total_sum / total_students, 2)
        
        return exam_averages

def main():
    """
    Main function to run the complete pipeline
    """
    print("="*70)
    print("🚀 AUTOMATED STUDENT PERFORMANCE REPORT GENERATION SYSTEM")
    print("="*70)
    print("📁 Reading: Compilesheet IX 'D' (2025-26) (Recovered).xlsx")
    print("📊 Format: Subject names in Row 4, Assessment types in Row 5")
    print("="*70)
    
    # Initialize the generator with your Excel file
    excel_file = r"C:\Users\HP\OneDrive\Desktop\PROGRESS REPORT\students.xlsx"
    generator = StudentReportGenerator(excel_file)
    
    try:
        # Step 1: Load and process data with multi-level headers
        generator.load_and_process_data()
        
        # Step 2: Clean data and handle 'AB' values
        generator.clean_data()
        
        # Step 3: Generate summary report (console)
        generator.generate_summary_report()
        
        # Step 4: Generate PDF report
        timestamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
        pdf_file = generator.create_pdf_report(f"Class_IX_Performance_Report_{timestamp}.pdf")
        
        # Step 5: Export processed data
        excel_file_out = generator.export_to_excel(f"processed_student_data_{timestamp}.xlsx")
        
        print("\n" + "="*70)
        print("✨ ALL TASKS COMPLETED SUCCESSFULLY! ✨")
        print("="*70)
        print(f"📊 Generated Reports:")
        print(f"   • PDF Report: {pdf_file}")
        print(f"   • Excel Data: {excel_file_out}")
        print("="*70)
        
    except Exception as e:
        print(f"\n❌ Error occurred: {str(e)}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()