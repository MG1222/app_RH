import tkinter as tk
from tkinter import ttk, Frame, messagebox


import json

class ParametersPage(Frame):
    def __init__(self, parent, controller, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.controller = controller

        self.params = self.load_params_from_json()
        self.default_params = {
            "last_name": "B6",
            "first_name": "B9",
            "tel_num": "C5",
            "email": "E5",
            "status1": "G22",
            "status2": "G23",
            "interview_dates_start": 7,
            "interview_dates_end": 9,
            "interview_managers_start": 7,
            "interview_managers_end": 9
        }
        if not self.params:
            self.params = self.default_params.copy()

        self.create_widgets()
        

    def load_params_from_json(self):
        param_json = {
            "last_name": "B6",
            "first_name": "B9",
            "tel_num": "C5",
            "tel_num_sec": "D5",
            "email": "E5",
            "status1": "G22",
            "status2": "G23",
            "interview_dates_start": 7,
            "interview_dates_end": 9,
            "interview_managers_start": 7,
            "interview_managers_end": 9
        }
        return param_json

    def save_params_to_json(self):
        param_json = {
            "last_name": self.params["last_name"],
            "first_name": self.params["first_name"],
            "tel_num": self.params["tel_num"],
	        "tel_num_sec": self.params["tel_num_sec"],
			"email": self.params["email"],
	        "status1": self.params["status1"],
	        "status2": self.params["status2"],
	        "interview_dates_start": self.params["interview_dates_start"],
	        "interview_dates_end": self.params["interview_dates_end"],
	        "interview_managers_start": self.params["interview_managers_start"],
	        "interview_managers_end": self.params["interview_managers_end"]
	        
        }
        

    def create_widgets(self):
        frame = ttk.Frame(self)
        frame.grid(padx=20, pady=20)

        ttk.Label(frame, text="Paramètres des cellules").grid(row=0, column=0, columnspan=2, pady=10)

        row = 1
        for param, value in self.params.items():
            label_text = self.get_french_label(param)
            ttk.Label(frame, text=label_text).grid(row=row, column=0, sticky="e", pady=5)
            entry = ttk.Entry(frame)
            entry.insert(0, value)
            entry.grid(row=row, column=1, sticky="w", pady=5)
            entry.param_key = param
            entry.bind("<FocusOut>", self.update_param)
            row += 1

        save_button = ttk.Button(frame, text="Enregistrer", command=self.save_params)
        save_button.grid(row=row, column=0, columnspan=2, pady=10)

    
        back_button = ttk.Button(frame, text="Retour", command=lambda: self.controller.show_frame("Visuel"))
        back_button.grid(row=row, column=2, columnspan=2, pady=10)

    def get_french_label(self, param):
        labels = {
            "last_name": "Nom de famille",
            "first_name": "Prénom",
            "tel_num": "Numéro de téléphone",
            "email": "Email",
            "status1": "Statut 1",
            "status2": "Statut 2",
            "interview_dates_start": "Début des dates d'entretien",
            "interview_dates_end": "Fin des dates d'entretien",
            "interview_managers_start": "Début des managers d'entretien",
            "interview_managers_end": "Fin des managers d'entretien"
        }
        return labels.get(param, param)

    def update_param(self, event):
        widget = event.widget
        self.params[widget.param_key] = widget.get()

    def save_params(self):
        self.save_params_to_json()
        self.controller.params = self.params
        messagebox.showinfo('Sauvegarde terminée', "Les données ont été sauvegardées avec succès!")


   
        
