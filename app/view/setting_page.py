import logging
import tkinter as tk
from tkinter import ttk, Frame, messagebox
import json
import os


class SettingPage(Frame):
    def __init__(self, parent, controller, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.controller = controller

        self.params = self.load_params_from_json()

        self.create_widgets()

    def load_params_from_json(self):
        try:
            with open('./config/setting_excel.json', 'r') as file:
                param_json = json.load(file)
        except FileNotFoundError:
            messagebox.showerror("Erreur", "Le fichier de configuration n'a pas été trouvé.")
            param_json = {}
        return param_json
    def create_widgets(self):
        frame = ttk.Frame(self)
        frame.grid(padx=20, pady=20)
        ttk.Label(frame, text="Paramètres des cellules Excel").grid(row=0, column=0, columnspan=4, pady=10)

        row = 1
        col = 0
        pairs = [
            ("email", None),
            ("last_name", "first_name"),
            ("tel_num", "tel_num_sec"),
            ("status1", "status2"),
            ("interview_dates_start", "interview_dates_end"),
            ("interview_managers_start", "interview_managers_end")
        ]

        for pair in pairs:
            for param in pair:
                if param:
                    label_text = self.get_french_label(param)
                    ttk.Label(frame, text=label_text).grid(row=row, column=col, sticky="e", pady=5)
                    entry = ttk.Entry(frame)
                    entry.insert(0, self.params[param])
                    entry.grid(row=row, column=col + 1, sticky="w", pady=5)
                    entry.param_key = param
                    entry.bind("<FocusOut>", self.update_param)
                    col += 2
            col = 0
            row += 1

        save_button = ttk.Button(frame, text="Enregistrer", command=self.save_params)
        save_button.grid(row=row, column=1, columnspan=2, pady=10)
        back_button = ttk.Button(frame, text="Retour", command=lambda: self.controller.show_frame("HomePage"))
        back_button.grid(row=row, column=3, columnspan=2, pady=10)
        style = ttk.Style()
        style.configure("TButton", foreground="black")
        back_button.config(style="TButton")
        style.configure("W.TButton", foreground="green")
        save_button.config(style="W.TButton")
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
        try:
            with open('./config/setting_excel.json', 'w') as file:
                json.dump(param_json, file, indent=4)
                return True
        except Exception as e:
            logging.error(f"Error while saving params to json: {e}")
            return False
    def get_french_label(self, param):
        labels = {
            "last_name": "Nom de famille",
            "first_name": "Prénom",
            "tel_num": "Téléphone",
            "tel_num_sec": "Téléphone 2",
            "email": "Email",
            "status1": "Disponibilité 1 ",
            "status2": "Disponibilete 2",
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
        saved = self.save_params_to_json()
        if saved:
            self.controller.params = self.params
            messagebox.showinfo('Sauvegarde terminée', "Les données ont été sauvegardées avec succès!")
        else:
            messagebox.showerror('Erreur', "Une erreur s'est produite lors de la sauvegarde des données.")