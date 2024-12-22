import os
from docx import Document

def main():
    funktionauswahl = input("Gebe ein Welche Funktion des Programm du benutzen möchtest, \"1\" für ein Einzelnes Dokument: ")
    if funktionauswahl == "1":
        getsingledocument()
    elif funktionauswahl == "2":
        getmultpledocument()
    else:
        print("Gebe nur die Zahl 1 oder 2 ein")
        main()

def getsingledocument():
    while True:
        global pfad
        pfad = input("geben sie den Dateipfad an wo sich die Datei befindet: ")
        global dokument
        dokument = input("Gebe den vollen Dateinamen ein des Dokumentes: ")
        if not pfad.endswith("\\"):
            pfad += "\\"
        if not dokument.endswith(".docx"):
            dokument += ".docx"
        global docpfad
        docpfad = pfad + dokument

        if os.path.exists(docpfad):
            print(f"Die Datei {docpfad} existiert")
            break
        else:
            print(f"Der Angegebene Pfad: {pfad} oder die Datei {dokument} existiert nicht, gebe beides neu ein!")
    renamedata()

def getmultpledocument():
    while True:
        global pfad
        pfad = input("Gebe den Ordnerpfad an wo alle Dateien umbenannt werden sollen: ")
        if not pfad.endswith("\\"):
            pfad += "\\"
            break
        if os.path.exists(pfad):
            break
        else:
            print(f"der angegebene Pfad \"{pfad}\" existiert nicht, gebe ihn neu ein")
    finddocx(pfad)

def finddocx(pfad):
    print("test")
    global dateien
    dateien = []
    for root, _, dateinamen in os.walk(pfad):
        for datei in dateinamen:
            if datei.endswith('.docx'):
                dateien.append(os.path.join(root, datei))
    return dateien


def renamedata():
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
    getsingledocument2()

def getsingledocument2():    
    while True:
        weitermachen2 = input('Möchtest du das noch eine andere Datei umbennant wird, gebe "ja" oder "nein" ein.').lower()
        if weitermachen2 == "ja":
            print("Das Programm wird fortgeführt!")
            getsingledocument()
        elif weitermachen2 == "nein":
            print("Danke fürs benutzen.")
            exit()  
        else:
            print("Ungültige Eingabe, gebe nur \"ja\" oder \"nein\"")

if __name__ == "__main__":
    main()