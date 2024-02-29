import tkinter as tk
from tkinter import filedialog, Label, Entry, Button, StringVar, Frame, Checkbutton
import pandas as pd
import os


class ExcelSplitterApp(tk.Tk):
    def update_filter_info(self):
        applied_filters = []
        if self.ccb_var.get():
            applied_filters.append("Verteiler für CCB")
        if self.cedus_var.get():
            applied_filters.append("Verteiler für Cedus")

        if not applied_filters:
            self.filter_info_var.set("Keine Filter angewendet")
        else:
            self.filter_info_var.set(
                f"Gefiltert nach: {', '.join(applied_filters)}")

    def __init__(self):
        super().__init__()
        self.title("Excel Aufteiler")
        self.geometry("400x300")  # Größe des Fensters angepasst
        self.configure(bg='light blue')  # Hintergrundfarbe des Fensters

        # Erstelle einen Frame als Container, der zentriert wird
        self.center_frame = Frame(self, bg='light blue')
        self.center_frame.place(relx=0.5, rely=0.5, anchor='center')

        # Stiloptionen
        label_font = ('Arial', 10, 'bold')
        button_font = ('Arial', 10, 'bold')
        feedback_font = ('Arial', 9, 'italic')

        # Eingabefeld für den Namen des Blattes
        self.label_sheet_name = Label(
            self.center_frame, text="Name des Blattes:", font=label_font, bg='light blue')
        self.label_sheet_name.grid(
            row=0, column=0, padx=10, pady=5, sticky="e")

        self.sheet_name_var = tk.StringVar(
            value="Trennen in 50")  # Standardwert
        self.entry_sheet_name = Entry(
            self.center_frame, textvariable=self.sheet_name_var, width=20)
        self.entry_sheet_name.grid(
            row=0, column=1, padx=10, pady=5, sticky="w")

        # Eingabefeld für Einträge pro Blatt
        self.label_entries_per_sheet = Label(
            self.center_frame, text="Einträge pro Blatt:", font=label_font, bg='light blue')
        self.label_entries_per_sheet.grid(
            row=1, column=0, padx=10, pady=5, sticky="e")

        self.entries_per_sheet_var = tk.IntVar(value=50)  # Standardwert ist 50
        self.entry_entries_per_sheet = Entry(
            self.center_frame, textvariable=self.entries_per_sheet_var, width=10)
        self.entry_entries_per_sheet.grid(
            row=1, column=1, padx=10, pady=5, sticky="w")

        # Checkbox für "Verteiler für CCB"
        self.ccb_var = tk.BooleanVar(value=False)
        self.check_ccb = Checkbutton(self.center_frame, text="Verteiler für CCB", variable=self.ccb_var,
                                     font=label_font, bg='light blue', command=self.update_filter_info)
        self.check_ccb.grid(row=2, column=0, padx=10, pady=5, sticky="w")

        # Checkbox für "Verteiler für Cedus"
        self.cedus_var = tk.BooleanVar(value=False)
        self.check_cedus = Checkbutton(self.center_frame, text="Verteiler für Workshops",
                                       variable=self.cedus_var, font=label_font, bg='light blue', command=self.update_filter_info)
        self.check_cedus.grid(row=2, column=1, padx=10, pady=5, sticky="w")

        # Button zum Auswählen der Datei
        self.button_select_file = Button(
            self.center_frame, text="Excel-Datei auswählen", command=self.split_excel, font=button_font, bg='light grey')
        self.button_select_file.grid(
            row=3, column=0, columnspan=2, padx=10, pady=10)

        # Label für die Rückmeldung
        self.feedback_var = tk.StringVar(
            value="Noch keine Excel aufgeteilt")  # Standardwert
        self.label_feedback = Label(self.center_frame, textvariable=self.feedback_var,
                                    font=feedback_font, bg='light blue', fg='dark green', wraplength=380)
        self.label_feedback.grid(
            row=4, column=0, columnspan=2, padx=10, pady=5)

        # Label für Filterinformationen
        self.filter_info_var = tk.StringVar(
            value="Keine Filter angewendet")  # Standardwert
        self.label_filter_info = Label(self.center_frame, textvariable=self.filter_info_var, font=(
            'Arial', 8, 'italic'), bg='light blue', fg='blue', wraplength=380)
        self.label_filter_info.grid(
            row=5, column=0, columnspan=2, padx=10, pady=5)

    def split_excel(self):
        entries_per_sheet = self.entries_per_sheet_var.get()
        filepath = filedialog.askopenfilename(
            title="Wähle eine Excel-Datei aus", filetypes=[("Excel files", "*.xlsx;*.xls")])
        if not filepath:
            return

        directory, original_filename = os.path.split(filepath)
        sheet_name = self.sheet_name_var.get()

        try:
            df = pd.read_excel(
                filepath, sheet_name=sheet_name if sheet_name else None)
        except ValueError as e:
            self.feedback_var.set(f"Fehler: {e}")
            return

        # Anwenden der spezifischen Filter
        if self.ccb_var.get():
            df = df[df["Verteiler für CCB"].str.lower() == "ja"]
        if self.cedus_var.get():
            df = df[df["Verteiler für Workshops"].str.lower() == "ja"]

        num_entries = len(df)
        sheets_needed = num_entries // entries_per_sheet + \
            (1 if num_entries % entries_per_sheet else 0)
        base_filename = "aufgeteilte_datei"
        filename = os.path.join(directory, f"{base_filename}.xlsx")
        file_counter = 1
        while os.path.exists(filename):
            filename = os.path.join(
                directory, f"{base_filename}_{file_counter}.xlsx")
            file_counter += 1

        with pd.ExcelWriter(filename, engine="openpyxl") as writer:
            for i in range(sheets_needed):
                start_row = i * entries_per_sheet
                end_row = start_row + entries_per_sheet
                df_slice = df.iloc[start_row:end_row]
                df_slice.to_excel(
                    writer, sheet_name=f'Blatt{i+1}', index=False)

        self.feedback_var.set(
            f"Datei erfolgreich aufgeteilt und gespeichert als '{os.path.basename(filename)}' im Ordner '{directory}'.")


# Erstelle und starte die App
app = ExcelSplitterApp()
app.mainloop()
