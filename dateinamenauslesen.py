import os
from docx import Document

while True:
    while True:
        pfad = input("geben sie den Dateipfad an wo sich die Datei befindet: ")
        dokument = input("Gebe den vollen Dateinamen ein des Dokumentes: ")
        if not pfad.endswith("\\"):
            pfad += "\\"
        if not dokument.endswith(".docx"):
            dokument += ".docx"
        docpfad = pfad + dokument

        if os.path.exists(docpfad):
            print(f"Die Datei {docpfad} existiert")
            break
        else:
            print(f"Der Angegebene Pfad: {pfad} oder die Datei {dokument} existiert nicht, gebe beides neu ein!")

    doc = Document(docpfad)
    print("Die Ersten 3 Wörter/Überschrifft sind:")
    first = doc.paragraphs[0].text
    first_extract = first.split()[:3]
    print(first_extract)

    while True:
        weitermachen = input("wenn du fortfahren möchtest gebe \"ja\" ein, wenn nicht dann ein \"nein\": ").lower()
        if weitermachen == "ja":
            print("Unbennennung wird gestartet")
            newname = "_".join(first_extract) + ".docx"
            newpath = pfad + newname
            os.rename(docpfad, newpath)
            print("Unbennenung war erfolgreich")
            break
        elif weitermachen == "nein":
            print("Programm wird abgebrochen")
            break
        else:
            print("Ungültige Eingabe, gebe nur \"ja\" oder \"nein\"")
    
    while True:
        weitermachen2 = input('Möchtest du das noch eine andere Datei umbennant wird, gebe "ja" oder "nein" ein.').lower()
        if weitermachen2 == "ja":
            print("Das Programm wird fortgeführt!")
            break
        elif weitermachen2 == "nein":
            print("Danke fürs benutzen.")
            exit()  
        else:
            print("Ungültige Eingabe, gebe nur \"ja\" oder \"nein\"")
