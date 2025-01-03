import os
import win32com.client
from docx import Document


def docconverter(pfad):
    try:
        if not pfad.endswith(".doc"):
            raise ValueError("Die Datei muss eine .doc Datei sein.")
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False

        doc = word.Documents.Open(pfad)
        docxpfad = os.path.splitext(pfad)[0] + ".docx"
        doc.SaveAs(docxpfad, FileFormat=16)
        doc.Close()
        word.Quit()
        return "Die Datei wurde Erfolgreich konvertiert, sie heißt jetzt: " + docxpfad
    except Exception as e:
        return f"Fehler beim Konvertieren: {e}"


def syntax_file_name(newname):
    for char in " :/\\*?\"<>|":
        newname = newname.replace(char, "_")
    return newname


def renamedata(pfad):
    try:
        doc = Document(pfad)
        first = doc.paragraphs[0].text.strip()
        first_extract = first.split()[:3]
        newname = syntax_file_name("_".join(first_extract)) + ".docx"
        directory = os.path.dirname(pfad)
        newname = generate_unique_name(directory, newname)
        newpath = os.path.join(directory, newname)
        doc.save(newpath)
        pfad = newpath
        return f"Die Datei wurde Erfolgreich in {pfad} umbenannt."
    except Exception as e:
        return f"Fehler beim Umbenennen: {e}"



def generate_unique_name(pfad, newname):
    base, extension = os.path.splitext(newname)
    counter = 1
    while os.path.exists(os.path.join(pfad, newname)):
        newname = f"{base}_{counter}{extension}"
        counter += 1
    return newname


