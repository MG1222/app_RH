import os
import logging
from relance_rh.ui import Visuel



# create a log file
log_file = os.path.join(os.path.dirname(__file__), "logs", "log.log")
logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

logging.info("Starting the application")


def main():
	try:
		ui = Visuel()
		ui.find_folder_widget()
	except Exception as e:
		logging.error(f"An error occured: {e}")
		raise e



if __name__ == "__main__":
	main()




