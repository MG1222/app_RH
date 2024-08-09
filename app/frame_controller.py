import logging
from tkinter import Tk

from app.view.home_page import HomePage
from app.view.setting_page import SettingPage


class FrameController(Tk):
	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)
		self.frames = {}
		
		self.grid_rowconfigure(1, weight=1)
		self.grid_columnconfigure(0, weight=1)
		
		for F in (HomePage, SettingPage):
			page_name = F.__name__
			frame = F(parent=self, controller=self)
			self.frames[page_name] = frame
			frame.grid(row=1, column=0, sticky="nsew")
		
		self.show_frame("HomePage")
	
	def show_frame(self, page_name):
		
		if page_name in self.frames:
			frame = self.frames[page_name]
			frame.tkraise()
			if page_name == "HomePage":
				self.geometry("550x500")
			else:
				self.geometry("700x400")
		else:
			logging.error(f"appController = Page not found {page_name}.")
			print(f"Page {page_name} non trouvée.")

		