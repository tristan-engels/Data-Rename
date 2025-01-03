import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
from logic import docconverter, renamedata

def funktionsauswahl_gui():
    root = ttk.Window(themename="cosmo")
    root.title("Datei Bro")
    root.geometry("800x400")

    label = ttk.Label(root, text="Dein Datei Bro", font=("Arial", 16), bootstyle="primary")
    label.pack(pady=20)

    button_convert = ttk.Button(
        root, text="Docx Converter öffnen", 
        command=lambda: converter_gui(root),  
        bootstyle="success"
    )
    button_convert.pack(pady=10)

    button_rename = ttk.Button(
        root, text="Datei umbenennen", 
        command=lambda: rename_gui(root),
        bootstyle="info"
    )

    button_rename.pack(pady=10)

    root.mainloop()



def converter_gui(parent):  
    converter_window = ttk.Toplevel(parent)
    converter_window.title("DOC zu DOCX Converter")
    converter_window.geometry("600x300")

    label = ttk.Label(converter_window, text="DOC zu DOCX Converter", font=("Arial", 14), bootstyle="info")
    label.pack(pady=20)

    button_select_file = ttk.Button(
        converter_window, text="Datei auswählen und konvertieren",

        command=dateiauswahl_convert_gui,
        bootstyle="primary"
    )
    button_select_file.pack(pady=10)

    

def dateiauswahl_convert_gui():
    pfade = filedialog.askopenfilenames(
        title = "Datei auswählen",
        filetypes=[("Word-Dateien", "*.doc")],
    )
    
    success_count = 0
    error_count = 0
    error_files = []

    for pfad in pfade:
        result_message = docconverter(pfad)
        if "Erfolgreich" in result_message:
            success_count += 1
        else:
            error_count += 1
            error_files.append(pfad)

    converter_window = ttk.Toplevel()
    converter_window.title("Status")
    converter_window.geometry("300x100")
    status_label = ttk.Label(converter_window, text="", font=("Arial", 14), bootstyle="secondary")
    status_label.pack(pady=20)

    summary = f"{success_count} Datei(en) erfolgreich konvertiert.\n"
    if error_count > 0:
        summary += f"{error_count} Fehler:\n" + "\n".join(error_files)
    status_label.config(text=summary, bootstyle="success" if error_count == 0 else "danger")

def rename_gui(parent):
    converter_window = ttk.Toplevel(parent)
    converter_window.title("Word Dokumente automatisch umbenennen")
    converter_window.geometry("600x300")

    label = ttk.Label(converter_window, text="Word/DocX Dokumente Automatisch umbenennen", font=("Arial", 14), bootstyle="info")
    label.pack(pady=20)

    button_select_file = ttk.Button(
        converter_window, text="Datei/Dateien auswählen und Umbenennen",

        command=dateiauswahl_rename_gui,
        bootstyle="primary"
    )
    button_select_file.pack(pady=10)

def dateiauswahl_rename_gui():
    pfade= filedialog.askopenfilenames(
        title = "Datei auswählen",
        filetypes=[("Word-Dateien", "*.docx")],
    )

    succes_count = 0
    error_count = 0
    error_files = []

    for pfad in pfade:
        result_message = renamedata(pfad)
        if "Erfolgreich" in result_message:
            succes_count += 1
        else:
            error_count += 1
            error_files.append(pfad)
    
    rename_window = ttk.Toplevel()
    rename_window.title("Status")
    rename_window.geometry("400x200")
    status_label = ttk.Label(rename_window, text="", font=("Arial", 14), bootstyle="secondary")
    status_label.pack(pady=20)

    summary = f"{succes_count} Datei(en) erfolgreich konvertiert.\n"
    if error_count > 0:
        summary += f"{error_count} Fehler:\n" + "\n".join(error_files)
    status_label.config(text=summary, bootstyle="success" if error_count == 0 else "danger")

if __name__ == "__main__":
    funktionsauswahl_gui()
