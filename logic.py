import os
import win32com.client
from docx import Document


def main():
    while True:
        funktionauswahl = input(
            'Wählen Sie eine Funktion: "1" für ein einzelnes Dokument, "2" für mehrere Dokumente oder "3" zum Konvertieren von .doc zu .docx: '
        )
        if funktionauswahl == "1":
            getsingledocument()
            break
        elif funktionauswahl == "2":
            getmultpledocument()
            break
        elif funktionauswahl == "3":
            docconverter()
            break
        else:
            print("Ungültige Eingabe. Bitte geben Sie nur 1, 2 oder 3 ein.")


def docconverter():
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False

        einodermehr = input("Möchten Sie ein einzelnes Dokument oder mehrere Dokumente konvertieren? (ein/mehr): ").lower()

        if einodermehr == "ein":
            while True:
                pfad = input("Geben Sie den Ordnerpfad an, in dem sich die Datei befindet: ").strip()
                dokument = input("Geben Sie den Dateinamen mit Endung .doc ein: ").strip()

                if not pfad.endswith("\\"):
                    pfad += "\\"
                if not dokument.endswith(".doc"):
                    dokument += ".doc"
                docpfad = os.path.join(pfad, dokument)

                if not os.path.exists(docpfad):
                    print(f"Die Datei {docpfad} existiert nicht. Bitte erneut versuchen.")
                    continue

                try:
                    print(f"Verarbeite: {docpfad}")
                    doc = word.Documents.Open(docpfad)
                    docxpfad = os.path.splitext(docpfad)[0] + ".docx"
                    doc.SaveAs(docxpfad, FileFormat=16)
                    doc.Close()
                    print(f"Die Datei wurde erfolgreich konvertiert: {docxpfad}")
                    break
                except Exception as e:
                    print(f"Fehler beim Konvertieren der Datei: {e}")
                    continue

        elif einodermehr == "mehr":
            while True:
                pfad = input("Geben Sie den Ordnerpfad an, in dem sich die zu konvertierenden Dateien befinden: ").strip()
                if not pfad.endswith("\\"):
                    pfad += "\\"
                if os.path.exists(pfad):
                    dateien = finddoc(pfad)
                    for file_path in dateien:
                        convert_doc_to_docx(file_path, pfad, word)
                    break
                else:
                    print(f"Der angegebene Pfad \"{pfad}\" existiert nicht. Bitte erneut versuchen.")
    finally:
        word.Quit()


def convert_doc_to_docx(file_path, pfad, word):
    try:
        print(f"Verarbeite: {file_path}")
        doc = word.Documents.Open(file_path)
        docxpfad = os.path.splitext(file_path)[0] + ".docx"
        doc.SaveAs(docxpfad, FileFormat=16)
        doc.Close()
        print(f"Die Datei wurde erfolgreich konvertiert: {docxpfad}")
    except Exception as e:
        print(f"Fehler beim Konvertieren der Datei {file_path}: {e}")


def finddoc(pfad):
    return [os.path.join(root, file) for root, _, files in os.walk(pfad) for file in files if file.endswith(".doc")]


def getsingledocument():
    while True:
        pfad = input("Geben Sie den Dateipfad an, in dem sich die Datei befindet: ").strip()
        dokument = input("Geben Sie den Dateinamen mit Endung .docx ein: ").strip()
        if not pfad.endswith("\\"):
            pfad += "\\"
        if not dokument.endswith(".docx"):
            dokument += ".docx"
        docpfad = os.path.join(pfad, dokument)

        if os.path.exists(docpfad):
            print(f"Die Datei {docpfad} existiert.")
            renamedata(docpfad, pfad)
            break
        else:
            print(f"Die Datei oder der Pfad existiert nicht. Bitte erneut versuchen.")


def getmultpledocument():
    while True:
        pfad = input("Geben Sie den Ordnerpfad an, in dem sich die Dateien befinden: ").strip()
        if not pfad.endswith("\\"):
            pfad += "\\"
        if os.path.exists(pfad):
            dateien = finddocx(pfad)
            renamemultipledata(dateien, pfad)
            break
        else:
            print(f"Der Pfad {pfad} existiert nicht. Bitte erneut versuchen.")


def finddocx(pfad):
    return [os.path.join(root, file) for root, _, files in os.walk(pfad) for file in files if file.endswith(".docx")]


def syntax_file_name(newname):
    for char in " :/\\*?\"<>|":
        newname = newname.replace(char, "_")
    return newname


def renamemultipledata(dateien, pfad):
    for file_path in dateien:
        try:
            doc = Document(file_path)
        except Exception as e:
            print(f"Die Datei {file_path} konnte nicht geöffnet werden. Fehler: {e}")
            continue

        first = doc.paragraphs[0].text
        first_extract = first.split()[:3]
        newname = syntax_file_name("_".join(first_extract)) + ".docx"
        newname = generate_unique_name(pfad, newname)
        newpath = os.path.join(pfad, newname)

        print(f"Datei: {file_path} -> Neuer Name: {newname}")

        while True:
            weitermachen = input("Möchten Sie die Datei umbenennen? (ja/nein): ").lower()
            if weitermachen == "ja":
                try:
                    os.rename(file_path, newpath)
                    print(f"Die Datei wurde erfolgreich umbenannt: {newname}")
                except Exception as e:
                    print(f"Fehler beim Umbenennen: {e}")
                break
            elif weitermachen == "nein":
                print("Die Datei bleibt unverändert.")
                break
            else:
                print("Ungültige Eingabe. Bitte ja oder nein eingeben.")


def renamedata(docpfad, pfad):
    doc = Document(docpfad)
    first = doc.paragraphs[0].text
    first_extract = first.split()[:3]
    newname = syntax_file_name("_".join(first_extract)) + ".docx"
    newpath = os.path.join(pfad, newname)

    while True:
        weitermachen = input("Möchten Sie die Datei umbenennen? (ja/nein): ").lower()
        if weitermachen == "ja":
            try:
                os.rename(docpfad, newpath)
                print(f"Die Datei wurde erfolgreich umbenannt: {newname}")
            except Exception as e:
                print(f"Fehler beim Umbenennen: {e}")
            break
        elif weitermachen == "nein":
            print("Die Umbenennung wurde abgebrochen.")
            break
        else:
            print("Ungültige Eingabe. Bitte ja oder nein eingeben.")


def generate_unique_name(pfad, newname):
    base, extension = os.path.splitext(newname)
    counter = 1
    while os.path.exists(os.path.join(pfad, newname)):
        newname = f"{base}_{counter}{extension}"
        counter += 1
    return newname


if __name__ == "__main__":
    main()
