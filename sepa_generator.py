import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from docxtpl import DocxTemplate
from docx import Document
import os

class SepaGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("SEPA Mandat Generator v2.1")
        self.root.geometry("600x400")

        # --- Variablen für die Dateipfade ---
        self.excel_path = tk.StringVar()
        self.template_path = tk.StringVar()
        self.output_path = tk.StringVar()

        # --- GUI Elemente ---
        # Titel
        title_label = tk.Label(root, text="SEPA Mandat Generator", font=("Helvetica", 16, "bold"))
        title_label.pack(pady=10)

        # Frame für die Dateiauswahl
        frame = tk.Frame(root, padx=10, pady=10)
        frame.pack(fill="x", padx=10)

        # Excel-Datei
        btn_excel = tk.Button(frame, text="1. Excel-Datei auswählen...", command=self.select_excel)
        btn_excel.grid(row=0, column=0, sticky="ew", pady=5)
        lbl_excel = tk.Label(frame, textvariable=self.excel_path, fg="blue", anchor="w")
        lbl_excel.grid(row=0, column=1, sticky="ew", padx=10)

        # Word-Template
        btn_template = tk.Button(frame, text="2. Word-Vorlage auswählen...", command=self.select_template)
        btn_template.grid(row=1, column=0, sticky="ew", pady=5)
        lbl_template = tk.Label(frame, textvariable=self.template_path, fg="blue", anchor="w")
        lbl_template.grid(row=1, column=1, sticky="ew", padx=10)

        # Output-Ordner
        btn_output = tk.Button(frame, text="3. Speicherordner auswählen...", command=self.select_output)
        btn_output.grid(row=2, column=0, sticky="ew", pady=5)
        lbl_output = tk.Label(frame, textvariable=self.output_path, fg="blue", anchor="w")
        lbl_output.grid(row=2, column=1, sticky="ew", padx=10)
        
        frame.columnconfigure(1, weight=1)

        # Generate Button
        btn_generate = tk.Button(root, text="Mandate generieren", font=("Helvetica", 12, "bold"), bg="green", fg="white", command=self.generate_documents)
        btn_generate.pack(pady=20, ipadx=10, ipady=5)

        # Status Label
        self.status_var = tk.StringVar()
        self.status_var.set("Bereit. Bitte alle drei Pfade auswählen.")
        lbl_status = tk.Label(root, textvariable=self.status_var, font=("Helvetica", 10), fg="gray")
        lbl_status.pack(side="bottom", fill="x", pady=5, padx=10)

    def select_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if path:
            self.excel_path.set(os.path.basename(path))
            self._full_excel_path = path

    def select_template(self):
        path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if path:
            self.template_path.set(os.path.basename(path))
            self._full_template_path = path

    def select_output(self):
        path = filedialog.askdirectory()
        if path:
            self.output_path.set(path)
            self._full_output_path = path

    def format_iban(self, iban):
        clean_iban = str(iban).replace(' ', '')
        return ' '.join(clean_iban[i:i+4] for i in range(0, len(clean_iban), 4))

    def generate_documents(self):
        if not all([self.excel_path.get(), self.template_path.get(), self.output_path.get()]):
            messagebox.showerror("Fehler", "Bitte wählen Sie eine Excel-Datei, eine Word-Vorlage und einen Speicherordner aus.")
            return

        self.status_var.set("Verarbeitung wird gestartet...")
        self.root.update_idletasks()

        try:
            df = pd.read_excel(self._full_excel_path)
            
            mandates_data = {}
            for _, row in df.iterrows():
                kontoinhaber = row['Kontoinhaber']
                kinder = set()
                if pd.notna(row['Name Kind']):
                    kinder.add(str(row['Name Kind']).strip())
                if pd.notna(row['Geschwister']):
                    geschwister_liste = [g.strip() for g in str(row['Geschwister']).split(',') if g.strip()]
                    kinder.update(geschwister_liste)
                
                if kontoinhaber not in mandates_data:
                    mandates_data[kontoinhaber] = {
                        'IBAN': self.format_iban(row['IBAN']),
                        'BIC': row['BIC'],
                        'KINDER': kinder
                    }
                else:
                    mandates_data[kontoinhaber]['KINDER'].update(kinder)

            contexts = []
            for kontoinhaber, data in mandates_data.items():
                sorted_kinder = sorted(list(data['KINDER']))
                contexts.append({
                    'KONTOINHABER': kontoinhaber,
                    'IBAN': data['IBAN'],
                    'BIC': data['BIC'],
                    'KINDERLISTE': ', '.join(sorted_kinder),
                    'sort_key': sorted_kinder[0] if sorted_kinder else ''
                })
            
            sorted_contexts = sorted(contexts, key=lambda x: x['sort_key'])

            # --- NEUE LOGIK ZUR DOKUMENTENERSTELLUNG (Version 3) ---
            
            # Fall behandeln, falls keine Daten vorhanden sind
            if not sorted_contexts:
                self.status_var.set("Keine Daten in der Excel-Datei gefunden.")
                messagebox.showinfo("Information", "Keine Daten zum Generieren in der Excel-Datei gefunden.")
                return

            # 1. Das erste Mandat wird als Basis für das gesamte Dokument generiert
            first_context = sorted_contexts[0]
            tpl = DocxTemplate(self._full_template_path)
            tpl.render(first_context)
            final_doc = tpl.docx  # Dies ist jetzt unser Zieldokument, beginnend mit dem ersten Mandat.

            # 2. Iteriere durch die restlichen Kontexte (beginnend mit dem ZWEITEN)
            for context in sorted_contexts[1:]:
                # Füge ZUERST einen Seitenumbruch hinzu
                final_doc.add_page_break()

                # Erstelle eine neue Vorlagen-Instanz für das nächste Mandat
                tpl_next = DocxTemplate(self._full_template_path)
                tpl_next.render(context)
                
                # Füge den Inhalt des neu gefüllten Dokuments zum Zieldokument hinzu
                for element in tpl_next.docx.element.body:
                    final_doc.element.body.append(element)
            
            # 3. Speichere das finale Gesamtdokument
            output_filename = os.path.join(self._full_output_path, "SEPA-Mandate_gesamt.docx")
            final_doc.save(output_filename)
            # --- ENDE DER NEUEN LOGIK ---

            self.status_var.set(f"Erfolg! Datei wurde gespeichert unter:\n{output_filename}")
            messagebox.showinfo("Erfolg", f"Die SEPA-Mandate wurden erfolgreich generiert und in der Datei 'SEPA-Mandate_gesamt.docx' gespeichert.")

        except Exception as e:
            self.status_var.set("Ein Fehler ist aufgetreten!")
            messagebox.showerror("Fehler bei der Verarbeitung", f"Ein unerwarteter Fehler ist aufgetreten:\n\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = SepaGeneratorApp(root)
    root.mainloop()