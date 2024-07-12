from tkinter import Frame, ttk, filedialog as fd, messagebox
from PIL import Image, ImageTk
import logging
import threading
import time
from datetime import date
import subprocess
import sys
import os

from relance_rh.parameters_page import ParametersPage
from relance_rh.excel_operations import ExcelOperations


def resource_path(relative_path):

	try:
		# PyInstaller creates a temp folder and stores path in _MEIPASS
		base_path = sys._MEIPASS2
	except Exception:
		base_path = os.path.abspath(".")
	
	return os.path.join(base_path, relative_path)



class Visuel(Frame):
	def __init__(self, parent, controller, *args, **kwargs):
		super().__init__(parent, *args, **kwargs)
		self.controller = controller
		
		# Set up the main grid layout
		self.grid(row=0, column=0, sticky="nsew")
		self.grid_rowconfigure(1, weight=1)
		self.grid_columnconfigure(0, weight=1)
		
		self.create_widgets()
		self.controller.title("Relance RH")
	
	def create_widgets(self):
		# Create a frame for the tree view and other elements
		main_frame = Frame(self)
		main_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
		main_frame.grid_rowconfigure(3, weight=1)
		main_frame.grid_columnconfigure(0, weight=1)
		
		# Add logos frame
		logos_frame = Frame(main_frame)
		logos_frame.grid(row=0, column=0, pady=(0, 20), sticky="n")
		
		# Load logos
		logo1 = Image.open(resource_path("./asset/logo.png"))
		logo1.thumbnail((100, 100))
		logo1 = ImageTk.PhotoImage(logo1)
		
		logo2 = Image.open(resource_path("./asset/logoGT.jpg"))
		logo2.thumbnail((150, 150))
		logo2 = ImageTk.PhotoImage(logo2)
		
		# Add logos to the frame
		logo1_label = ttk.Label(logos_frame, image=logo1)
		logo1_label.image = logo1
		logo1_label.grid(row=0, column=0, padx=20)
		
		logo2_label = ttk.Label(logos_frame, image=logo2)
		logo2_label.image = logo2
		logo2_label.grid(row=0, column=1, padx=20)
		
		# Add description and version
		description = ttk.Label(main_frame,
		                        text="Mini app d'enregistrement et creation de fichier Excel avec les "
		                             "informations des candidats.",
		                        wraplength=400, justify="center")
		description.grid(row=1, column=0, pady=(0, 10), sticky="n")
		version = ttk.Label(main_frame, text="Version 1.0")
		version.grid(row=2, column=0, pady=(0, 20), sticky="n")
		
		# Create a frame for buttons
		buttons_frame = Frame(main_frame)
		buttons_frame.grid(row=4, column=0, pady=(20, 10), sticky="n")
		
		# Configure styles
		style = ttk.Style()
		style.configure("Treeview", rowheight=30)
		style.configure('W.TButton', font=('calibri', 10, 'bold', 'underline'), foreground='red')
		style.configure('TButton', font=('calibri', 10, 'bold'), foreground='blue')
		
		# Add buttons
		search_button = ttk.Button(buttons_frame, text="Chercher", command=self.select_folder)
		search_button.grid(row=0, column=0, padx=10)
		

		
		# Create a frame for the progress bar
		self.progress_frame = Frame(self)
		self.progress_frame.grid(row=1, column=0, sticky="ew", padx=20, pady=10)
		self.progress = ttk.Progressbar(self.progress_frame, orient="horizontal", length=500, mode="determinate")
		self.progress.grid(row=0, column=0, sticky="ew")
		
		# Initially hide the progress bar
		self.progress.grid_remove()
	
	def select_folder(self):
		self.folder = fd.askdirectory()
		if self.folder:
			logging.info("Folder selected")
			self.find_files_widget()
		else:
			messagebox.showwarning("Aucun dossier sélectionné", "Vous n'avez sélectionné aucun dossier.")
			logging.error("No folder selected")
	
	def find_files_widget(self):
		self.progress.grid()
		self.progress['value'] = 0
		self.progress.update()
		threading.Thread(target=self.process_files).start()
	
	def process_files(self):
		instance_exelOpr = ExcelOperations()
		data = instance_exelOpr.process_excel_files(self.folder, self.progress)
		if data is None:
			messagebox.showerror("Aucun fichier trouvé", "Aucun fichier n'a été trouvé dans le dossier sélectionné")
			logging.error("No files found in the selected folder")
		else:
			time.sleep(1)
			messagebox.showinfo("Traitement terminé",
			                    "Le traitement des fichiers est terminé, vous pouvez maintenant sauvegarder les données")
			logging.info("Files processed successfully")
			self.prompt_save(data)
		
		self.progress.grid_remove()
	
	def prompt_save(self, data):
		file_path = self.save_widget(data)
		if file_path:
			open_file_response = messagebox.askquestion("Sauvegarde terminée",
			                                            "Les données ont été sauvegardées avec succès! Voulez-vous ouvrir le fichier?")
			if open_file_response == "yes":
				self.open_file(file_path)
	
	def save_widget(self, data):
		today = date.today().strftime("%m-%y")
		file_path = fd.asksaveasfilename(defaultextension=".xlsx", initialfile=f'relance-{today}')
		
		if file_path:
			try:
				instance_exelOpr = ExcelOperations()
				instance_exelOpr.create_new_excel_file(data, file_path)
				logging.info("Data saved successfully")
				return file_path
			except ValueError as e:
				messagebox.showerror("Erreur lors de la sauvegarde",
				                     f"Une erreur est survenue lors de la sauvegarde des données: {e}")
				logging.error(f"Error while saving data: {e}")
				return None
		else:
			messagebox.showwarning("Sauvegarde annulée",
			                       "Aucun fichier n'a été sélectionné pour la sauvegarde.")
			return None
	
	def open_file(self, file_path):
		logging.debug(f"open_file called with file_path: {file_path}")
		try:
			if sys.platform == "win32":
				os.startfile(file_path)
			elif sys.platform == "darwin":
				subprocess.call(("open", file_path))
			else:
				subprocess.call(("xdg-open", file_path))
			logging.info(f"Opened file {file_path} successfully")
		except Exception as e:
			messagebox.showerror("Erreur d'ouverture du fichier",
			                     f"Une erreur est survenue lors de l'ouverture du fichier: {e}")
			logging.error(f"Error opening file {file_path}: {e}")

