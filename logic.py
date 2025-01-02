import os
import win32com.client
from docx import Document


def main():
    while True:
        funktionauswahl = input(
            'Gebe ein, welche Funktion des Programms du benutzen möchtest: "1" für ein einzelnes Dokument oder "2" für mehrere oder "3" um Dokumente von .doc in docx zu ändern: '
        )
        if funktionauswahl == "1":
            getsingledocument()
            break
        elif funktionauswahl == "2":
            getmultpledocument()
            break
        elif funktionauswahl == "3":
            docconverter()
        else:
            print("Ungültige Eingabe. Gebe nur die Zahl 1 oder 2 ein.")

def docconverter():
    word = win32com.client.Dispatch("Word.3Application")
    word.Visible = False

    einodermehr = input("Möchtest du ein einzelnes Dokument oder mehrere Dokumente konvertieren? (ein/mehr): ").lower()
    
    if einodermehr == "ein":
        while True:
            pfad = input("Geben Sie den Dateipfad an, wo sich die Datei befindet: ").strip()
            dokument = input("Geben Sie den vollen Dateinamen des Dokuments ein: ").strip()
            
            if not pfad.endswith("\\"):
                pfad += "\\"
            if not dokument.endswith(".doc"):
                dokument += ".doc"
            docpfad = os.path.join(pfad, dokument)

            if not os.path.exists(docpfad):
                print(f"Die Datei {docpfad} existiert nicht. Bitte geben Sie den Pfad erneut ein.")
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
            pfad = input("Geben Sie den Ordnerpfad an, wo alle Dateien konvertiert werden sollen: ")
            if not pfad.endswith("\\"):
                pfad += "\\"
            if os.path.exists(pfad):
                dateien = finddoc(pfad)
                for file_path in dateien:
                    convert_doc_to_docx(file_path, pfad, word)
                break
            if not os.path.exists(pfad):
                print(f"Der Pfad {pfad} existiert nicht. Bitte erneut eingeben.")
                continue
            else:
                print(f"Der angegebene Pfad \"{pfad}\" existiert nicht. Geben Sie ihn neu ein.")

def generate_unique_name(pfad, newname):
    base, extension = os.path.splitext(newname)  
    counter = 1
    unique_name = newname
    while os.path.exists(os.path.join(pfad, unique_name)):
        unique_name = f"{base}_{counter}{extension}"  
        counter += 1
    return unique_name

def getsingledocument():
    while True:
        pfad = input("Geben Sie den Dateipfad an, wo sich die Datei befindet: ")
        dokument = input("Geben Sie den vollen Dateinamen des Dokuments ein: ")
        if not pfad.endswith("\\"):
            pfad += "\\"
        if not dokument.endswith(".docx"):
            dokument += ".docx"
        docpfad = pfad + dokument

        if os.path.exists(docpfad):
            print(f"Die Datei {docpfad} existiert.")
            renamedata(docpfad, pfad)
            break
        else:
            print(
                f"Der angegebene Pfad: {pfad} oder die Datei {dokument} existiert nicht. Geben Sie beides neu ein!"
            )


def getmultpledocument():
    while True:
        pfad = input("Geben Sie den Ordnerpfad an, wo alle Dateien umbenannt werden sollen: ")
        if not pfad.endswith("\\"):
            pfad += "\\"
        if os.path.exists(pfad):
            dateien = finddocx(pfad)
            renamemultipledata(dateien, pfad)
            break
        if not os.path.exists(pfad):
            print(f"Der Pfad {pfad} existiert nicht. Bitte erneut eingeben.")
            continue
        else:
            print(f"Der angegebene Pfad \"{pfad}\" existiert nicht. Geben Sie ihn neu ein.")


def finddocx(pfad):
    dateien = []
    for root, _, dateinamen in os.walk(pfad):
        for datei in dateinamen:
            if datei.endswith(".docx"):
                dateien.append(os.path.join(root, datei))
    return dateien

def syntax_file_name(newname):
    newname = newname.replace(" ", "_")
    newname = newname.replace(":", "_")
    newname = newname.replace("/", "_")
    newname = newname.replace("\\", "_")
    newname = newname.replace("*", "_")
    newname = newname.replace("?", "_")
    newname = newname.replace("\"", "_")
    newname = newname.replace("<", "_")
    newname = newname.replace(">", "_")
    newname = newname.replace("|", "_")
    return newname


def renamemultipledata(dateien, pfad):
    for file_path in dateien:
        try:
            doc = Document(file_path)
        except Exception as e:
            print(f"Die Datei {file_path} konnte nicht geöffnet werden. Fehler: {e}")
            continue
        doc = Document(file_path)
        first = doc.paragraphs[0].text
        first_extract = first.split()[:3]
        newname = "_".join(first_extract) + ".docx"
        newname = syntax_file_name(newname)
        newname = generate_unique_name(pfad, newname)
        newpath = os.path.join(pfad, newname)

        print(f"\nAktuelle Datei: {file_path}")
        print(f"Vorgeschlagener Name: {newname}")

        while True:
            weitermachen = input(
                'Möchten Sie diese Datei umbenennen? Geben Sie "ja" oder "nein" ein: '
            ).lower()
            if weitermachen == "ja":
                try:
                    os.rename(file_path, newpath)
                    print(f"Umbenennung war erfolgreich. Die Datei heißt nun: {newname}")
                except Exception as e:
                    print(f"Umbenennung war nicht erfolgreich. Fehler: {e}")
                    continue
                break
            elif weitermachen == "nein":
                print("Die Datei wurde übersprungen und bleibt unverändert.")
                break
            else:
                print('Ungültige Eingabe. Geben Sie nur "ja" oder "nein" ein.')


def renamedata(docpfad, pfad):
    doc = Document(docpfad)
    first = doc.paragraphs[0].text
    first_extract = first.split()[:3]
    print(f"Die ersten 3 Wörter/Überschrift des Dokuments sind: {first_extract}")

    while True:
        weitermachen = input(
            'Wenn Sie fortfahren möchten, geben Sie "ja" ein, wenn nicht, "nein": '
        ).lower()
        if weitermachen == "ja":
            newname = "_".join(first_extract) + ".docx"
            newpath = os.path.join(pfad, newname)
            os.rename(docpfad, newpath)
            print(f"Umbenennung war erfolgreich. Die Datei heißt nun: {newname}")
            break
        elif weitermachen == "nein":
            print("Umbenennung wurde abgebrochen.")
            break
        else:
            print('Ungültige Eingabe. Geben Sie nur "ja" oder "nein" ein.')


if __name__ == "__main__":
    main()
