# RESULTANT
#### Video Demo:  <URL HERE>
#### Description:

So, what exactly is Resultant?

At its heart, it's a tool born out of a simple idea: manually creating detailed, well-formatted academic result analysis reports is a massive time sink. Anyone who's had to wrangle with PDF mark sheets, copy-paste data into Excel, and then spend hours formatting it to look professional knows the pain. Resultant is the cure for that headache.

It’s a web application that takes a standard university result PDF, and like magic, transforms it into a beautifully formatted, insightful Excel report, complete with summary statistics, subject-wise analysis, and even the college's logo for that professional touch. It's designed to save educators and administrative staff hours of tedious work, freeing them up to do what they actually do best: teach and support students.

This isn't just about converting a PDF to Excel; it's about adding intelligence to the process. It understands the structure of the results, calculates pass/fail counts, handles various edge cases, and presents the final data in a way that's immediately useful.

#### Inspiration

The inspiration for Resultant came from a real-world problem I observed during my internship at my father's institution. The administrative staff were spending an inordinate amount of time manually processing student results, which was not only inefficient but also prone to errors. I realized that there had to be a better way to handle this repetitive task, and thus, Resultant was born.

Excel formats and input result PDFs were sourced from my father's college, ensuring that the tool is tailored to real-world needs. The goal was to create a solution that was not only functional but also user-friendly and adaptable to various academic settings.

Hopefully, Resultant, after some polishing and additional features, will be adopted and hosted by my father's institution and potentially other colleges facing similar challenges.

#### Key Features

- Login and session gating for all core routes.
- PDF parsing for subject codes, subject names, seat numbers, student names, marks, per-subject status, overall status, and SGPA where available.
- A normalized SQLite schema with `student`, `subject`, and `results` tables created per uploaded dataset.
- A teacher assignment step that updates professors for each parsed subject and is then baked into the Excel header.
- Excel report writer with merged headers, per-subject Total Marks and Grade columns, a Remarks column, optional SGPA, thin-grid borders, conditional fills for FAIL or AB, and a summary section with COUNTIF based metrics.
- College branding support with configurable college name and logo embedded into the XLSX top banner.
- In-browser lightweight viewer for quick checks and a downloads list to retrieve generated files.

#### Architecture and Data Flow

1. User authentication and sessions are handled by Flask. Protected routes check a session key before allowing access.
2. On upload, the app saves the PDF to `uploads/`, invokes the PDF parser to extract subjects and student records, and writes a fresh SQLite database under `databases/` named after the file.
3. The UI asks you to supply teacher names for each detected subject. These are saved to the `subject` table.
4. The formatter reads the database, pivots rows to a student-by-subject table, adds remarks and SGPA columns if present, and writes a styled Excel workbook under `downloads/` with per-subject summaries and overall counts.
5. You can download the final `.xlsx` or open a limited viewer for a quick visual check.

#### Design choices

Mostly my design choices were driven by familiarity and practicality. I wanted to build something that worked well and could be rapidly developed and if need be, later on refactored; so I leaned on tools and libraries I already knew or had learned during CS50 itself.

Why Flask?
It's powerful and flexible without being bloated and most importantly I had just worked with it in PSET 9. For a focused application like this, Flask gave me everything I needed to build the web interface and manage the backend logic without getting in my way.

Why SQLite for the Database?
Honestly, this was an easy one for me. I already had experience with SQL, so why learn a whole new system when this fit perfectly? This means each report I process gets its own little .db file, keeping everything neat and self-contained. There was no need for a heavy, complex database server; though I do plan to migrate to PostgreSQL if I decide to scale this up or deploy it more widely.

The pandas and openpyxl libraries:
I quickly realized these two libraries were the perfect team for creating the final report. I used pandas for the heavy data lifting—merging the different tables and pivoting the data into the final structure. But for making minor adjustments, openpyxl was the way to go. It let me go in after the fact to control everything from merged headers and column widths to adding the college logo and coloring the "FAIL" cells. One handles the structure, the other handles the style.
Both libraries were a bit of a learning curve, but with some trial and error (and a lot of help from ChatGPT), I got the hang of it.

#### Filewise Description

I’ve structured the project into a few key Python files, each with its own special job. Completion of PSET 9 hugely helped me with the directory structures and sped up the process. This "separation of concerns" was a big deal for me. It means if we need to change how PDFs are read, we only have to mess with parse.py. If we want to redesign the Excel sheet, we only touch formatter.py. It keeps things clean and much easier to debug.

##### Files:
1. app.py
    This is the main web application file. It uses Flask to create a simple web server that you can interact with through your browser.
    Here’s what it does:

    Routing: It listens for when you visit a page, the request method, like /upload or /login (GET or POST), and decides what to do.

    The Workflow Conductor: When you upload a PDF, this file invokes the PDF parser to extract the necessary data. Then, after you've entered the teacher's names, it invokes the formatter to create the final Excel report.

    Session Management: It keeps track of who is logged in (via Flask sessions) and remembers important details between steps, like the name of the file you're working on.

    User & College Management: It handles all the logic for registering new users, logging them in, and the feature where you can add and edit college details (like names and logos).

2. parse.py
    This was the toughest, most detail-oriented script in this project. I had to experiment a lot to get this right; tried different libraries (like pdfplumber and pypdf) and approaches before settling on this one. Had to trial-error the regex patterns a lot to make sure they worked across different PDFs.

    The Job: This file is responsible for reading the uploaded PDF and extracting (structured database tables) all the relevant data..

    The Heavy Lifting: I'm using the fitz (PyMuPDF) library here, which is brilliant at tearing apart PDFs using screen-snip like functionality (fitz.Rect).

    Regex: I've used regex extensively throughout this script. It took a bit of trial and error to get these just right, but they are the key to reliably extracting the data.

    Database Packer: Once it has all the student data neatly organized, it connects to a temporary SQLite database and packs everything in. I decided on SQLite because it's a simple, file-based database and we had already worked on SQLite in PSET 7. It's fast and efficient enough for this use case, I hope.

3. formatter.py
    This file takes the raw, structured data from the database and turns it into the polished Excel report. Again, this was a bit of trial and error to get the formatting just right. Had to learn a lot about openpyxl to make the sheet look professional.

    Pandas: The first step is to pull the data from the database into a pandas DataFrame. I use it to merge the student, subject, and result info and then pivot the table. This pivot is what transforms the data from a long list of individual subject results into that wide format where each row is a single student.
    I had to take the help of ChatGPT to figure out how to do this pivoting correctly. Trying to learn pandas on my own was a bit overwhelming, but with some guidance, I got the hang of it.

    Openpyxl: While pandas can write to Excel, openpyxl has total control over the formatting. This was a deliberate choice, not a redundancy. I didn't just want the data in a grid; I wanted it to look like a professional report. 

    I've used openpyxl to do things like:
    Merge cells to create those nice multi-line headers for subject codes, names, and teachers.
    Add the college logo right onto the sheet.
    Set column widths and row heights to make everything readable.
    Apply borders, fonts, and colors. The little touch of coloring the "FAIL" and "AB" cells makes a huge difference for at-a-glance analysis.

    EXCEL Formulae: Decided on Excel formulae to account for manual future mark updation if necessary. It actually writes Excel formulas into the summary section at the bottom. So, the pass/fail counts are dynamic and calculated by Excel itself.

4. login_helper.py
    This is a small utility file (copied over from our PSET 9 Flask app) that handles user session management. Keeps the main app.py file cleaner.

5. templates (HTML Files)
  The HTML files serve as the user interface templates for different functionalities:
  - index.html: The landing page of the application.
  - login.html & register.html: Pages for user login and registration.
  - upload.html: Interface for users to upload files.
  - view.html & viewer.html: Pages to view uploaded or processed data.
  - college.html: Displays college-related information dynamically.
  - layout.html: A base template to provide consistent layout and styling across pages.

6. requirements.txt 
    Lists the project's dependencies and their specific versions necessary to run the application. This facilitates easy setup of the environment using package managers like pip.


#### Conclusion:

And that's pretty much the grand tour of Resultant! It was a really fun project to build; Completing CS50 has left me with the same feeling as finishing a great show: definite sense of 'what now?' But I learned a ton about web development, PDF parsing, data manipulation, and Excel formatting. More importantly, I built something that solves a real problem and saves people time.

If I were to continue working on this, I'd probably add features like:
    - Making it production-ready with Docker and deploying it on a cloud platform. (Currently trying to deploy on Railway but hitting some roadblocks with file paths and migrating to PostgreSQL).
    - Support for more PDF formats and layouts.
    - AI-powered error detection and correction during parsing (Probably a simple API call to OpenAI/Gemini).
    - More advanced analytics and visualizations in the Excel report.
    - Migrating from SQLite to a more robust database system for larger datasets (Probably PostgreSQL).
    - User roles and permissions for different levels of access.
    - An admin dashboard to manage users and view upload history.
    - Email notifications when reports are ready.
    - Personal 'professor dashboards' with their subject's performance metrics.
    - Bulk upload support for multiple PDFs at once.
    - And of course, lots of testing and polishing to make it rock-solid.


#### How to Run Locally
1. Python 3.11 or newer is recommended.
2. Create a virtual environment and install dependencies:
   ```bash
   python -m venv .venv
   . .venv/bin/activate  # Windows: .venv\Scripts\activate
   pip install -r requirements.txt
   flask run
   ```