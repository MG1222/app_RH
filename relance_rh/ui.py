import sys
from tkinter import *
from tkinter import ttk, filedialog as fd, messagebox
from PIL import Image, ImageTk
from relance_rh.excel_operations import ExcelOperations
import os
import logging
import time
import subprocess
import threading


def resource_path(relative_path):
	""" Get absolute path to resource, works for dev and for PyInstaller """
	try:
		base_path = sys._MEIPASS
	except Exception:
		base_path = os.path.abspath(".")
	return os.path.join(base_path, relative_path)


class Visuel:
	def __init__(self):
		self.logo = None
		self.label = None
		self.btn = None
		self.progress = None
		self.root = None
	
	def find_folder_widget(self):
		self.root = Tk()
		self.root.title("Relance RH")
		self.root.geometry("600x300")
		self.logo = Image.open(resource_path("./asset/logo.png"))
		self.logo.thumbnail((100, 100))
		self.logo = ImageTk.PhotoImage(self.logo)
		label = ttk.Label(self.root, image=self.logo)
		label.pack(side="top", padx=10, pady=10)
		self.label = ttk.Label(self.root, text="Veuillez choisir le dossier contenant les fichiers à traiter")
		self.label.pack()
		self.btn = ttk.Button(self.root, text="Chercher", command=self.select_folder)
		self.btn.pack()

		self.progress = ttk.Progressbar(self.root, orient="horizontal", length=300, mode="determinate")
		
		self.root.mainloop()
	
	def select_folder(self):
		self.folder = fd.askdirectory()
		if self.folder:
			logging.info("Folder selected")
			self.find_files_widget()
			return True
		else:
			messagebox.showwarning("Aucun dossier sélectionné",
			                       "Vous n'avez sélectionné aucun dossier.")
			logging.error("No folder selected")
			return False
	
	def find_files_widget(self):
		self.progress.pack(pady=10)
		self.progress['value'] = 0
		self.progress.update()
		
		threading.Thread(target=self.process_files).start()
		self.btn.config(state="disabled")
	
	def process_files(self):
		instance_exelOpr = ExcelOperations()
		data = instance_exelOpr.process_excel_files(self.folder, self.progress)
		
		if data is None:
			messagebox.showerror("Aucun fichier trouvé",
			                     "Aucun fichier n'a été trouvé dans le dossier sélectionné")
			logging.error("No files found in the selected folder")
			self.label.config(text="Checher un autre dossier")
			self.btn.config(state="normal")
		else:
			time.sleep(1)
			self.btn.config(state="disabled")
			messagebox.showinfo("Traitement terminé",
			                    "Le traitement des fichiers est terminé, vous pouvez maintenant sauvegarder les données")
			logging.info("Files processed successfully")
			self.prompt_save(data)
		
		self.progress.pack_forget()
	
	def save_widget(self, data):
		file = fd.asksaveasfile(mode='w', defaultextension=".xlsx")
		if file is None:
			messagebox.showwarning("Sauvegarde non sélectionnée",
			                       "Aucun fichier n'a été sélectionné pour la sauvegarde.")
			self.label.config(text="Sauvegarder les données")
			self.btn.config(text="Souvegarder", command=lambda: self.prompt_save(data), state="normal")
			return None
		
		instance_exelOpr = ExcelOperations()
		try:
			new_excel = instance_exelOpr.create_new_excel_file(data, file.name)
			logging.info("Data saved successfully")
			return file.name
		
		except ValueError as e:
			messagebox.showerror("Erreur lors de la sauvegarde",
			                     f"Une erreur est survenue lors de la sauvegarde des données: {e}")
			logging.error(f"Error while saving data: {e}")
			return None
	
	def prompt_save(self, data):
		self.btn.config(state="disabled")
		file_path = self.save_widget(data)
		if file_path:
			self.label.config(text="Checher un autre dossier")
			self.btn.config(text="Chercher", command=self.select_folder, state="normal")
			open_file_response = messagebox.askquestion("Sauvegarde terminée",
			                                            "Les données ont été sauvegardées avec succès! Voulez-vous ouvrir le fichier?")
			
			if open_file_response == "yes":
				self.open_file(file_path)
		else:
			self.label.config(text="Sauvegarder les données")
			self.btn.config(text="Souvegarder", command=lambda: self.prompt_save(data), state="normal")
	
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