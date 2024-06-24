import sqlite3
import logging

class Database:
	def __init__(self, db_path='relance_rh.db'):
		self.conn = sqlite3.connect(db_path)
		self.cursor = self.conn.cursor()

	def insert_data(self, data, interview_date):
		try:
			self.cursor.execute(
				"INSERT INTO relance_rh (last_name, first_name, email, interview_1, interview_2, interview_3) VALUES (?, ?, ?, ?, ?, ?)",
				(data['last_name'], data['first_name'], data['email'], data['dates'][0], data['dates'][1],
				 data['dates'][2]))
			self.conn.commit()
			return True
		except sqlite3.Error as e:
			print(e)
			logging.error(f"User with email {data['email']} already exists.")
			return False

	def update_data(self, interviews, email):
	    self.cursor.execute("""
	        UPDATE relance_rh 
	        SET interview_1 = ?, 
	            interview_2 = ?, 
	            interview_3 = ?, 
	            email_3 = ?
	        WHERE email = ?
	    """,  (interviews[0], interviews[1], interviews[2], True,  email))
	    self.conn.commit()
	    return True

	def get_data(self):
		self.cursor.execute("SELECT * FROM relance_rh")
		return self.cursor.fetchall()

	def close(self):
		self.conn.close()
		logging.info("Database closed")