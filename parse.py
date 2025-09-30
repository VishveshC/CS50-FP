import os
import fitz
import sqlite3
import re

SUBJECT_CODES = []
SUBJECT_NAMES = []
student_data = []
all_students_data = []
db_path, dbname, year, result_type, department, semester, result_name, course_code = None, None, None, None, None, None, None, None

def parse_pdf(file_path, year, result_type, department, semester, result_name, course_code):

    # --- Part 1: Extracting Data ---
    try:
        doc = fitz.open(file_path)
        page = doc[0]
    except:
        print("Error: Could not open the PDF file.")
        exit(1)

    # Extract subject codes from the header line
    search_results = page.search_for("EXAMINATION")
    rect = search_results[0]
    data_block_rect = fitz.Rect(0, rect.y0 + 20, page.rect.width, rect.y1 + 90)
    text_block = page.get_text("text", clip=data_block_rect)
    lines = text_block.split('\n')

    # Pattern detection for subject codes
    subject_pattern = re.compile(r"^(?!(?:EXAMINATION|AND))\S+\s*:.*", re.IGNORECASE)

    subject_lines = [line for line in lines if subject_pattern.match(line)]

    if not subject_lines:
        print("Error: Could not find any valid subject headers in the defined area.")
        exit()

    SUBJECT_CODES = [s.split(':', 1)[0].strip() for s in subject_lines]
    SUBJECT_NAMES = [s.split(':', 1)[1].split(' (')[0].strip() for s in subject_lines]

    # --- Part 2: Parsing All Student Data ---
    try:
        # Pattern detection for seat numbers
        seat_no_pattern = re.compile(r"\b\d{10,}\b")

        # Loop through each page of the document
        for page in doc:
            words = page.get_text("words")  # Extracts a list of [x0, y0, x1, y1, "word"]

            # Find all words that match our seat number pattern
            found_seat_numbers = [w for w in words if seat_no_pattern.fullmatch(w[4])]

            for word_info in found_seat_numbers:
                num_subjects = len(SUBJECT_CODES)
                total_marks = []

                rect = fitz.Rect(word_info[:4]) 
                data_block_rect = fitz.Rect(0, rect.y0, page.rect.width, rect.y1 + 80)
                text_block = page.get_text("text", clip=data_block_rect)
                
                lines = text_block.split('\n')

                

                # Find seat numbers in the block
                current_seat_no = None
                for line in lines:
                    match = seat_no_pattern.search(line)
                    if match:
                        current_seat_no = match.group(0) 
                        line_with_name = line
                        break
                
                # If seat number wasn't found in this block, skip it
                if not current_seat_no:
                    print(f"Warning: Could not find a seat number in a data block. Skipping.")
                    continue
                
                # Find student name
                student_name = None
                temp_name_str = line_with_name.replace(current_seat_no, '').strip()
                
                college_match = re.search(r'\d+\s-', temp_name_str)
                if college_match:
                    student_name = (temp_name_str[:college_match.start()].strip().title())
                    try:
                        student_name = student_name.split(')')[1]
                    except:
                        student_name = student_name

                # Find subject total marks
                try:
                    result_index = lines.index('FAIL') if 'FAIL' in lines else lines.index('PASS') if 'PASS' in lines else lines.index('WITHHELD')
                    start_index = result_index - num_subjects
                    total_marks = [int(mark) for mark in lines[start_index:result_index]]
                except:
                    print(f"Warning: Could not parse marks for Seat No: {current_seat_no}! Skipping.")
                    continue

                # Find subject status (PASS/FAIL) for each subject
                subject_status = [
                    'FAIL' if item.replace(' ', '').split('/')[1] == 'FF' else 'PASS'
                    for item in lines if item.count('/') == 2
                ]
                fail_count = subject_status.count('FAIL')
                remarks = 0
                
                sgpa = None
                # Handle special case for "AB" and SGPA
                for i, line in enumerate(lines):
                    pattern = re.compile(r".+/.+/.+")
                    if line == "AB":
                        subject_status[i] = "AB"
                        remarks += 1
                    elif pattern.search(line):
                        # 2. Add an index safety check
                        if i >= 2:
                            sgpa1 = lines[i - 2].strip()[:6]
                            sgpa2 = lines[i - 1].strip()[:6]
                            sgpa = f'{sgpa1}, {sgpa2}'
                            break 
                        else:
                            print(f"Warning: Found SGPA pattern on line {i}, but not enough preceding data.")
                            break
                else:
                    print("Warning: Could not find SGPA for this student.")



                # Find final student status (PASS/FAIL)
                final_result = "PASS" if "PASS" in lines else "FAIL" if "FAIL" in lines else "Unknown"
                
                remark_string = ""
                if fail_count > 0:
                    remark_string = f"Failed in {fail_count} subject(s)"
                elif remarks > 0:
                    remark_string = f"Absent for {remarks} subject(s)"
                else:
                    remark_string = "All Clear"

                # Neatly store all student data
                for i, subject_code in enumerate(SUBJECT_CODES):
                    all_students_data.append({
                        "seat_no": current_seat_no,
                        "name": student_name if student_name else "",
                        "subject_code": subject_code,
                        "total_marks": total_marks[i] if i < len(total_marks) else 0,
                        "final_result": final_result,
                        "subject_status": subject_status[i] if i < len(subject_status) else "PASS",
                        "subject_name": SUBJECT_NAMES[i] if i < len(SUBJECT_NAMES) else "",
                        "remarks": remark_string,
                        "sgpa": sgpa if sgpa else ""
                    })

        student_count = len({d['seat_no'] for d in all_students_data})
        print(f"Successfully extracted data for {student_count} students.")

    except Exception as e:
        print(f"An error occurred during data extraction: {e}")


    # --- Part 3: Inserting Data into SQLite Database ---
    if all_students_data:
        try:
            dbname = (file_path.split('\\')[1]).split('.')[0].strip()
            db_path = os.path.join("databases", dbname)
            conn = sqlite3.connect(f'{db_path}.db')
            cursor = conn.cursor()
            
            cursor.execute('''
            DROP TABLE IF EXISTS results
            ''')

            cursor.execute('''
            DROP TABLE IF EXISTS subject
            ''')

            cursor.execute('''
            DROP TABLE IF EXISTS student
            ''')

            
            cursor.execute('''
            CREATE TABLE IF NOT EXISTS results (
                id INTEGER PRIMARY KEY,
                seat_no TEXT NOT NULL,
                subject_code TEXT NOT NULL,
                total_marks INTEGER,
                subject_status TEXT,
                remarks TEXT,
                sgpa TEXT
            )
            ''')

            cursor.execute('''
            CREATE TABLE IF NOT EXISTS subject (
                subject_code TEXT PRIMARY KEY,
                subject_name TEXT NOT NULL,
                professor TEXT
            )
            ''')

            cursor.execute('''
            CREATE TABLE IF NOT EXISTS student (
                seat_no TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                final_result TEXT
            )
            ''')

            
            # Insert Data into Tables

            # SUBJECT table
            unique_subjects = { (rec['subject_code'], rec['subject_name']) for rec in all_students_data }
            
            cursor.executemany('''
            INSERT OR IGNORE INTO subject (subject_code, subject_name, professor)
            VALUES (?, ?, 'Unknown')
            ''', list(unique_subjects))

            # RESULTS table
            results_to_insert_in_results = [
                (
                    record['seat_no'], record['subject_code'], record['total_marks'],
                    record['subject_status'], record['remarks'], record['sgpa']
                ) for record in all_students_data
            ]
            
            cursor.executemany('''
            INSERT INTO results (seat_no, subject_code, total_marks, subject_status, remarks, sgpa)
            VALUES (?, ?, ?, ?, ?, ?)
            ''', results_to_insert_in_results)

            # STUDENT table
            results_to_insert_in_students = [
                (
                    record['seat_no'], record['name'], record['final_result']
                ) for record in all_students_data
            ]

            cursor.executemany('''
            INSERT OR IGNORE INTO student (seat_no, name, final_result)
            VALUES (?, ?, ?)
            ''', results_to_insert_in_students)

            conn.commit()
            print("\nDatabase update complete. All student data has been saved successfully.")
            
        except sqlite3.Error as e:
            print(f"Database error: {e}")
        finally:
            if conn:
                conn.close()
                print("Database connection closed.")
    else:
        print("\nNo student data was extracted, skipping database insertion.")
    
    # Update global variables
    db_path = db_path
    year = year
    result_type = result_type
    department = department
    semester = semester
    result_name = result_name
    course_code = course_code

    return unique_subjects

