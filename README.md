# FastText

Un Blocco note per Windows con **tabelle**, **immagini** e **schede**, in stile Notepad++.
Un solo font, nessuna formattazione: solo testo, intestazioni, tabelle e immagini.

## Funzioni
- Schede multiple, pannello laterale con l'elenco dei file aperti e la struttura del documento
- Intestazioni (Ctrl+1/2/3) e linee di separazione
- Tabelle (Ctrl+T) e immagini incollate (Ctrl+V) o trascinate
- Ricerca avanzata: normale, estesa, espressioni regolari; trova tutti, nei tab aperti e nei file di una cartella
- Maiuscole/minuscole e operazioni sulle righe (ordina, duplica, sposta, rimuovi duplicati...)
- Apre i file Word (.docx) in vista semplificata; esporta in Word e PDF

## Formato `.ftx`
Un file ZIP con `content.md` (Markdown: `#` titoli, `|` tabelle, `![](images/...)` immagini, `---` linee)
e la cartella `images/`.

## Avvio
1. `installa.bat` (una volta: installa PySide6, python-docx, pillow)
2. `FastText.bat`
3. facoltativo: `crea_exe.bat` (crea `dist\FastText\FastText.exe`) e `associa_ftx.bat`

Dettagli e scorciatoie in [LEGGIMI.txt](LEGGIMI.txt).

## File
| File | Contenuto |
| --- | --- |
| `fasttext.py` | finestra, schede, editor |
| `ftcore.py` | modello del documento, `.ftx`, Word, conversioni |
| `ftsearch.py` | ricerca e sostituzione |
| `fttextops.py` | maiuscole/minuscole e operazioni sulle righe |
| `ftpanels.py` | pannello laterale (file aperti, struttura) |
