import os
import sqlite3
from flask import Flask, flash, redirect, render_template, request, session, send_from_directory, url_for
from flask_session import Session
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename
from openpyxl import load_workbook
from datetime import datetime


from formatter import format_results
from login_helper import db, login_required
from parse import parse_pdf


# Configure application
app = Flask(__name__)
app.secret_key = 'some_random_secret_string'


# Configure session to use filesystem (instead of signed cookies)
app.config["SESSION_PERMANENT"] = False
app.config["SESSION_TYPE"] = "filesystem"
app.config["UPLOAD_FOLDER"] = "uploads"
app.config['DOWNLOAD_FOLDER'] = "downloads"
DOWNLOAD_FOLDER = 'downloads'
ALLOWED_EXTENSIONS = {"pdf"}
Session(app)

conn = None

@app.after_request
def after_request(response):
    """Ensure responses aren't cached"""
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    response.headers["Expires"] = 0
    response.headers["Pragma"] = "no-cache"
    return response

@app.route('/upload', methods=['GET', 'POST'])
@login_required
def upload_file():
    # POST
    if request.method == 'POST':
        if 'teachers_are_being_submitted' in request.form:
            filename = session.get('filename')
            year = session.get('year')
            semester = session.get('semester')
            result_name = session.get('result_name')
            result_type = session.get('result_type')
            department = session.get('department')
            course_code = session.get('course_code')
            
            
            if not filename or not year or not semester or not result_type or not department or not course_code:
                flash('Session expired. Please re-upload the PDF file.', 'danger')
                return redirect(url_for('upload_file'))
            
            subject_teacher_map = {}

            for key, value in request.form.items():
                if key.startswith('teacher_'):
                    subject_code = key.replace('teacher_', '', 1)
                    subject_teacher_map[subject_code] = value
            
            print("Received teacher names:", subject_teacher_map)

            try:
                db_path = os.path.join("databases", filename.split('.')[0])
                conn = sqlite3.connect(f'{db_path}.db')
                cursor = conn.cursor()

                for subject_code, professor in subject_teacher_map.items():
                    cursor.execute('''
                    UPDATE subject
                    SET professor = ?
                    WHERE subject_code = ?
                    ''', (professor, subject_code))
                
                conn.commit()
                print("Successfully updated subject teachers in the database.")

            except sqlite3.Error as e:
                print(f"Database error while updating subject teachers: {e}")
            finally:
                if conn:
                    conn.close()
                    print("Database connection closed after updating subject teachers.")

            format_results(f'{db_path}.db', year, result_type, department, semester, result_name, course_code, session['college_id']) 

            flash('Output file generated successfully!')

            # Clear the session variables
            session.pop('filename', None)
            session.pop('year', None)
            session.pop('semester', None)
            session.pop('result_name', None)
            session.pop('result_type', None)
            session.pop('department', None)
            session.pop('course_code', None)

            return redirect(f'/download/{filename}')
        else:
            year = request.form.get("academic_year")
            semester = int(request.form.get("semester"))
            result_name = None
            result_type = request.form.get("result_type")
            department = request.form.get("department")
            course_code = request.form.get("course_code")
            session['college_id'] = int(request.form.get("college_id"))

            if 'file' not in request.files:
                flash('No file part in the request.')
                return redirect(request.url)
            
            file = request.files['file']

            # If the user does not select a file, the browser submits an empty file without a filename.
            if file.filename == '':
                flash('No file selected.')
                return redirect(request.url)
            
            # If the file exists and has an allowed extension
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                save_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                
                file.save(save_path)
                
                flash(f'File "{filename}" uploaded successfully!')
                
                subjects = parse_pdf(save_path, year, result_type, department, semester, result_name, course_code)

                session['year'] = year
                session['semester'] = semester
                session['result_name'] = result_name
                session['result_type'] = result_type
                session['department'] = department
                session['course_code'] = course_code
                session['filename'] = filename
                
                return render_template('upload.html', subjects=subjects)

    # GET
    colleges = []
    with db("users.db") as conn:
        rows = conn.execute("SELECT * FROM college ORDER BY name").fetchall()
        colleges = [dict(row) for row in rows]

    return render_template('upload.html', colleges=colleges)

@app.route('/download/<path:filename>')
@login_required
def download_file(filename):
    download_url = f'{f'{filename.split('.')[0]}.xlsx'}'
    return send_from_directory(
        app.config['DOWNLOAD_FOLDER'], 
        download_url, 
        as_attachment=True
    )

@app.route("/viewer/<path:filename>")
@login_required
def view(filename):
    book = load_workbook(os.path.join(app.config['DOWNLOAD_FOLDER'], filename))
    sheet = book.active
    flash("File preview (not all features supported); please use a proper XLSX viewer for best results.", 'info')
    return render_template("viewer.html", sheet=sheet, title=filename)

@app.route("/view", methods=["GET"])
@login_required
def view_all():
    files = os.listdir(app.config['DOWNLOAD_FOLDER'])
    flash("File preview (not all features supported); please use a proper XLSX viewer for best results.", 'info')
    return render_template("view.html", files=files)

@app.route("/college", methods=["GET", "POST"])
@login_required
def college():
    # Handle the form submission
    if request.method == "POST":
        college_name = request.form.get("college_name", "").strip()
        college_id = request.form.get("college_id")
        logo_file = request.files.get("logo")
        
        db_logo_path = None
        conn = None # Define here to ensure it's accessible in finally block

        try:
            # 1. Handle the logo file if one was uploaded
            if logo_file and logo_file.filename != '':
                # Sanitize the college name to create a safe filename
                base_filename = college_name.lower().replace(' ', '_')
                extension = secure_filename(logo_file.filename).rsplit('.', 1)[1].lower()
                filename = f"{base_filename}.{extension}"
                
                # Full path to save the file on the server
                save_path = os.path.join(os.path.join('static', 'uploads'), filename)
                logo_file.save(save_path)
                
                # Path to store in DB (relative path for use in HTML src attribute)
                db_logo_path = os.path.join('uploads', filename).replace("\\", "/")

            # 2. Interact with the database
            conn = sqlite3.connect("users.db")
            cursor = conn.cursor()

            if college_id == "None":
                # This is a new college entry
                cursor.execute("INSERT INTO college (name, logo_path) VALUES (?, ?)", 
                               (college_name, db_logo_path))
                flash("New college added successfully!", 'success')
            else:
                # This is an update to an existing college
                if db_logo_path:
                    # If a new logo was uploaded, update both name and logo path
                    cursor.execute("UPDATE college SET name = ?, logo_path = ? WHERE id = ?",
                                   (college_name, db_logo_path, int(college_id)))
                else:
                    # If no new logo was uploaded, only update the name
                    cursor.execute("UPDATE college SET name = ? WHERE id = ?",
                                   (college_name, int(college_id)))
                flash("College information updated successfully!", 'success')

            conn.commit()
            return redirect(url_for('college'))

        except sqlite3.Error as e:
            # Print the actual error to your terminal for easy debugging
            print(f"A database error occurred: {e}") 
            
            # Provide a specific user-friendly message for common errors
            if "UNIQUE constraint failed" in str(e):
                flash(f"A college with the name '{college_name}' already exists.", 'danger')
            else:
                flash("A database error occurred. Please try again.", 'danger')
            return redirect(url_for('college'))
        finally:
            if conn:
                conn.close()
    
    # Handle the page load (display the form)
    else:
        conn = sqlite3.connect("users.db")
        conn.row_factory = sqlite3.Row
        rows = conn.execute("SELECT * FROM college ORDER BY name").fetchall()
        conn.close()
        
        # Convert the list of Row objects into a list of dictionaries for JSON serialization
        colleges = [dict(row) for row in rows]
        
        return render_template("college.html", colleges=colleges)

@app.route("/")
@login_required
def index():
    try:
        conn = sqlite3.connect("users.db")
        conn.row_factory = sqlite3.Row

        cursor = conn.cursor()
        
        sql_query = "SELECT username FROM users WHERE id = ?"
        params = (session["user_id"],)
        
        cursor.execute(sql_query, params)
        
        row = cursor.fetchone()
        
        username = row["username"] if row else None

    except sqlite3.Error as e:
        print(f"Database error: {e}")
    finally:
        if conn:
            conn.close()

    # List files in the download folder to create links dynamically
    files = os.listdir(app.config['DOWNLOAD_FOLDER'])

    return render_template("index.html", files=files, username=username)

@app.route('/delete/<path:filename>', methods=["POST"])
@login_required
def delete_file(filename):
    file_path = os.path.join(app.config['DOWNLOAD_FOLDER'], filename)
    try:
        if os.path.exists(file_path):
            os.remove(file_path)
            os.remove(os.path.join(app.config['UPLOAD_FOLDER'], filename.split('.')[0] + '.pdf'))
            os.remove(os.path.join('databases', filename.split('.')[0] + '.db'))
            flash(f'File "{filename}" deleted successfully!', 'success')
        else:
            flash(f'File "{filename}" not found.', 'danger')
    except Exception as e:
        flash(f'An error occurred while deleting the file: {e}', 'danger')
    
    return redirect("/")

@app.route("/register", methods=["GET", "POST"])
def register():
    # POST
    if request.method == "POST":
        name = request.form.get("username")
        psswd_hash = generate_password_hash(request.form.get("password"))

        # Ensure username was submitted
        if not name:
            flash("Must provide username", 'danger')
            return redirect("/register")

        # If passwords do not match
        elif request.form.get("password") != request.form.get("confirmation"):
            flash("Given passwords must match", 'danger')
            return redirect("/register")

        # Ensure password was submitted
        elif not request.form.get("password") or not request.form.get("confirmation"):
            flash("Must provide password", 'danger')
            return redirect("/register")

        # If everything's good
        else:
            # Check for username already exists
            try:
                conn = sqlite3.connect("users.db")
                with conn:
                    cursor = conn.cursor()
                    cursor.execute('''
                    CREATE TABLE IF NOT EXISTS users (
                        username TEXT PRIMARY KEY,
                        hash TEXT NOT NULL
                    )
                    ''')
                    cursor.execute('''
                    CREATE TABLE IF NOT EXISTS college (
                        id NUMERIC PRIMARY KEY,
                        name TEXT NOT NULL,
                        logo_path TEXT
                    )
                    ''')
                    cursor.execute("INSERT INTO users (username, hash) VALUES(?, ?)", (name, psswd_hash))
                    return redirect("/login")
            except:
                flash("Username already taken", 'danger')
                return redirect("/register")
            finally:
                if conn:
                    conn.close()
    # GET
    else:
        return render_template("register.html")

@app.route("/login", methods=["GET", "POST"])
def login():
    session.clear()

    # POST
    if request.method == "POST":
        # Ensure username was submitted
        if not request.form.get("username"):
            flash("Must provide username!", 'danger')
            return redirect("/")

        # Ensure password was submitted
        elif not request.form.get("password"):
            flash("Must provide password!", 'danger')
            return redirect("/")

        # Query database for username
        try:
            conn = sqlite3.connect("users.db")
            conn.row_factory = sqlite3.Row # To access columns by name and not tuple index
            with conn:
                cursor = conn.cursor()
                rows = cursor.execute("SELECT * FROM users WHERE username = ?", (request.form.get("username"),)).fetchall()

                # Ensure username exists and password is correct
                if len(rows) != 1 or not check_password_hash(rows[0]["hash"], request.form.get("password")):
                    flash("Invalid username and/or password!", 'danger')
                    return redirect("/")

                # Remember which user has logged in
                session["user_id"] = rows[0]["id"]
        finally:
            if conn:
                conn.close()

        return redirect("/")
    # GET
    else:
        return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")

# Function to check if the file extension is allowed
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS