import os
import sys
import json
import logging
from relance_rh.ui import Visuel
from relance_rh.create_database import init_db
from relance_rh.database import Database

log_file = os.path.join(os.path.dirname(__file__), "logs", "log.log")
logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

logging.info("Starting the application")


def main():
	init_db()
	new_db = Database()

	ui = Visuel()
	ui.find_folder_widget()



if __name__ == "__main__":
	main()




