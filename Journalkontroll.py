
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import re
import os
import sys
import subprocess

# Funksjon for å håndtere ressursbaner, spesielt for PyInstaller
def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# Filbaner for navn og ekskluderingsord
NAMES_FILE = resource_path("navn.txt")
EXCLUSION_WORDS_FILE = resource_path("ekskluderte_ord.txt")
# Hovedfunksjon for å finne journaler med nøkkelord og navn
def find_journals_with_keywords_and_names(journals_file, output_file, progress_callback=None):
    EXCLUDE_WHOLE_WORD_ONLY = False
    FIXED_NUM_SLASH_NUM_PATTERN = re.compile(r'\d{3}/\d+|plan \d+', re.IGNORECASE) # RegEx for å matche mønsteret for plannummer eller gårds- og bruksnummer
    # Laster navn og ekskluderingsord fra filer
    name_patterns_with_names = []
    try:
        with open(NAMES_FILE, 'r', encoding='utf-8') as n_file:
            for line in n_file:
                original_name = line.strip()
                if not original_name:
                    continue
                escaped_name = re.escape(original_name)
                compiled_pattern = re.compile(r'\b' + escaped_name + r'\b', re.IGNORECASE)
                name_patterns_with_names.append((original_name, compiled_pattern))
    except FileNotFoundError:
        messagebox.showerror("Feil", f"Navnefilen '{NAMES_FILE}' ble ikke funnet.")
        return
    # Laster ekskluderingsord og utelukker saker som dreier seg om gårds- og bruksnummer ved hjelp av RegEx
    exclusion_patterns = []
    try:
        with open(EXCLUSION_WORDS_FILE, 'r', encoding='utf-8') as ex_file:
            for line in ex_file:
                exclusion_word = line.strip()
                if not exclusion_word:
                    continue
                escaped_word = re.escape(exclusion_word)
                if EXCLUDE_WHOLE_WORD_ONLY:
                    compiled_pattern = re.compile(r'\b' + escaped_word + r'\b', re.IGNORECASE)
                else:
                    compiled_pattern = re.compile(escaped_word, re.IGNORECASE)
                exclusion_patterns.append((exclusion_word, compiled_pattern))
    except FileNotFoundError:
        messagebox.showwarning("Advarsel", f"Utelukkelsesfilen '{EXCLUSION_WORDS_FILE}' ble ikke funnet. Ingen utelukkelsesord fra filen vil bli brukt.")
    # Legger til RegEx-mønsteret for å utelukke saker med gårds- og bruksnummer
    exclusion_patterns.append(("fast_mønster: " + FIXED_NUM_SLASH_NUM_PATTERN.pattern, FIXED_NUM_SLASH_NUM_PATTERN))

    try:
        with open(journals_file, 'r', encoding='utf-8') as j_file:
            lines = j_file.readlines()

        with open(output_file, 'w', encoding='utf-8') as out_file:
            found_report_entries_count = 0
            total_journal_entries = len(lines)

            out_file.write("### Rapport for matchende journalposter ###\n\n")

            for i, journal_entry in enumerate(lines, 1):
                if progress_callback:
                    progress_callback(i, total_journal_entries)

                is_excluded = False
                for _, ex_pattern in exclusion_patterns:
                    if ex_pattern.search(journal_entry):
                        is_excluded = True
                        break

                if is_excluded:
                    continue

                matched_names = []
                for original_name, name_pattern in name_patterns_with_names:
                    if name_pattern.search(journal_entry):
                        matched_names.append(original_name)

                if matched_names:
                    found_report_entries_count += 1
                    out_file.write(f"Matchet av: {', '.join(matched_names)}\n")
                    out_file.write(f"Journalpost (Linje {i}): {journal_entry.strip()}\n")
                    out_file.write("---\n")

        #messagebox.showinfo("Ferdig", f"Behandling fullført.\nRapport: {output_file}")
    except Exception as e:
        messagebox.showerror("Feil", f"En feil oppsto: {e}")
#Grafisk brukergrensesnitt
def run_gui():
    root = tk.Tk()
    root.title("Journalsøk 🗳️")
    root.configure(bg="#f0f0f0")

    def browse_journal_file():
        filename = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        entry_journal.delete(0, tk.END)
        entry_journal.insert(0, filename)

    def update_progress(current, total):
        progress_var.set(int((current / total) * 100))
        root.update_idletasks()

    def process_files():
        journals_file = entry_journal.get()
        output_file = "rapport_resultat.txt"

        if not journals_file:
            messagebox.showwarning("Manglende informasjon", "Vennligst velg journalfil.")
            return

        progress_bar.grid(row=3, column=1, pady=5)
        progress_var.set(0)
        find_journals_with_keywords_and_names(journals_file, output_file, update_progress)
        progress_bar.grid_remove()
        open_report_file()  # Åpner rapporten automatisk etter prosessering

    def open_report_file():
        output_file = "rapport_resultat.txt"
        if os.path.exists(output_file):
            try:
                subprocess.Popen(["notepad", output_file])
            except Exception as e:
                messagebox.showerror("Feil", f"Kunne ikke åpne rapportfilen: {e}")
        else:
            messagebox.showwarning("Filen finnes ikke", "Rapportfilen er ikke generert enda.")

    tk.Label(root, text="Journalfil:", bg="#f0f0f0").grid(row=0, column=0, sticky="e", padx=5, pady=5)
    entry_journal = tk.Entry(root, width=60)
    entry_journal.grid(row=0, column=1, padx=5, pady=5)
    tk.Button(root, text="Velg...", command=browse_journal_file, bg="#d9d9d9").grid(row=0, column=2, padx=5, pady=5)

    tk.Button(root, text="Kjør søk 📔", command=process_files, bg="lightblue").grid(row=1, column=1, pady=10)
    #tk.Button(root, text="Åpne rapport", command=open_report_file, bg="lightgreen").grid(row=2, column=1, pady=5)

    progress_var = tk.IntVar()
    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100)
    progress_bar.grid(row=3, column=1, pady=5)
    progress_bar.grid_remove()

    root.mainloop()

if __name__ == "__main__":
    run_gui()
