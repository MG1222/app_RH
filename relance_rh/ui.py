import sys
from tkinter import *
from tkinter import ttk, filedialog as fd, messagebox
from PIL import Image, ImageTk
from relance_rh.excel_operations import ExcelOperations
import os
import logging
import time

# https://stackoverflow.com/questions/31836104/pyinstaller-and-onefile-how-to-include-an-image-in-the-exe-file

def resource_path(relative_path):
	""" Get absolute path to resource, works for dev and for PyInstaller """
	try:
		# PyInstaller creates a temp folder and stores path in _MEIPASS
		base_path = sys._MEIPASS
	except Exception:
		base_path = os.path.abspath(".")

	return os.path.join(base_path, relative_path)


class Visuel:


	def __init__(self):
		self.logo = None
		self.label = None
		self.btn = None

	def find_folder_widget(self):

		root = Tk()
		root.title("Relance RH")
		root.geometry("600x300")
		self.logo = Image.open(resource_path("./asset/logo.png"))
		self.logo.thumbnail((100, 100))
		self.logo = ImageTk.PhotoImage(self.logo)
		label = ttk.Label(root, image=self.logo)
		label.pack(side="top", padx=10, pady=10)
		self.label = ttk.Label(root, text="Veuillez choisir le dossier contenant les fichiers à traiter")
		self.label.pack()

		self.btn = ttk.Button(root, text="Chercher", command=self.select_folder)
		self.btn.pack()

		root.mainloop()

	def select_folder(self):
		self.folder = fd.askdirectory()
		if self.folder:
			logging.info("Folder selected")
			self.find_files_widget()
			return True
		else:
			messagebox.showinfo("Aucun dossier sélectionné","Vous n'avez sélectionné aucun dossier. Voulez-vous réessayer?")
			logging.error("No folder selected")
			return False




	def find_files_widget(self):
		instance_exelOpr = ExcelOperations()

		data = instance_exelOpr.process_excel_files(self.folder)

		if data is None:
			messagebox.showinfo("Aucun fichier trouvé", "Aucun fichier n'a été trouvé dans le dossier sélectionné")
			logging.error("No files found in the selected folder")
			self.label.config(text="Checher un autre dossier")
			self.btn.config(state="normal")
		else:
			self.btn.config(state="disabled")
			time.sleep(1)
			messagebox.showinfo("Traitement terminé", "Le traitement des fichiers est terminé, vous pouvez maintenant sauvegarder les données")
			logging.info("Files processed successfully")
			print(data)
			self.save_widget(data)
			return data

	def save_widget(self, data):
		self.btn.config(state="disabled")
		file = fd.asksaveasfile(mode='w', defaultextension=".xlsx")
		instance_exelOpr = ExcelOperations()
		new_excel = instance_exelOpr.create_new_excel_file(data, file.name)
		if new_excel is True:
			messagebox.showinfo("Sauvegarde terminée",
			                    "Les données ont été sauvegardées avec succès")
			logging.info("Data saved successfully")
			self.label.config(text="Checher un autre dossier")
			self.btn.config(text="Chercher", command=self.select_folder, state="normal")
			time.sleep(2)
		else:
			messagebox.showinfo("Aucun fichier sélectionné",
			                    "Vous n'avez pas sélectionné de fichier. Voulez-vous réessayer?")
			logging.error("No file selected")
			self.save_widget(data)












