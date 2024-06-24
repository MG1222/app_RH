import sqlite3
import logging

DB_PATH = 'relance_rh.db'

def init_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('''CREATE TABLE IF NOT EXISTS relance_rh
                      ( id INTEGER PRIMARY KEY AUTOINCREMENT, last_name TEXT, first_name TEXT, email TEXT UNIQUE, 
                      interview_1 TEXT, interview_2 TEXT, interview_3 TEXT, email_3 BOOLEAN DEFAULT FALSE,
                      email_6 BOOLEAN DEFAULT FALSE)''')
    conn.commit()
    conn.close()


if __name__ == '__main__':
    init_db()
    print("Database initialized successfully.")
    logging.info("SQL- Database initialized successfully.")