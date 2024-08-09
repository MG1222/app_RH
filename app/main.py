import os
import logging
from app.frame_controller import FrameController


log_dir = "logs"
log_file = os.path.join(log_dir, "log.log")


if not os.path.exists(log_dir):
    os.makedirs(log_dir)

logging.basicConfig(filename=log_file, level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

logging.warning("=== Starting the application")


def main():
    try:
        mainPage = FrameController()
        mainPage.mainloop()
    except Exception as e:
        logging.error(f"An error occured: {e}")
        raise e



if __name__ == "__main__":
    main()




