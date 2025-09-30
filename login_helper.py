from flask import redirect, render_template, session
from functools import wraps
import sqlite3

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if session.get("user_id") is None:
            return redirect("/login")
        return f(*args, **kwargs)

    return decorated_function

def db(db_name):
    """
    Establishes a connection to the SQLite database.
    Configures the connection to return rows as dictionary-like objects.
    """
    conn = sqlite3.connect(db_name)
    conn.row_factory = sqlite3.Row
    return conn