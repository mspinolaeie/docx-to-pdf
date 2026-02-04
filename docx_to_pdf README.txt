Ecco uno script pronto che converte in PDF tutti i `.docx` nella cartella corrente (o in una cartella a scelta), creando i **segnalibri** dai titoli e riportando correttamente i **metadati** del documento.

* Su **Windows** con **Microsoft Word** installato (via COM), ottieni la massima fedeltà: segnalibri da intestazioni o da bookmark Word e metadati inclusi nel PDF.
* In assenza di Word, usa il **fallback LibreOffice** (`soffice`), e lo script inserisce i metadati nel PDF leggendo quelli del DOCX.

### Uso rapido

```bash
# Converti tutti i .docx nella cartella corrente
python docx_to_pdf.py --dir . 

# Ricorsivo su sottocartelle, sovrascrive PDF già esistenti
python docx_to_pdf.py --dir . --recursive --overwrite

# Forza backend (auto | word | libreoffice)
python docx_to_pdf.py --use word

# Segnalibri: dai Titoli Word (default), dai bookmark Word o nessuno
python docx_to_pdf.py --bookmarks headings   # headings | word | none

# Esporta in PDF/A-1 (solo backend Word/COM)
python docx_to_pdf.py --pdfa
```

### Modalità “doppio click” (senza argomenti)

Se avvii `docx_to_pdf.py` **senza argomenti** (es. doppio click da Esplora Risorse), lo script apre una finestra per scegliere la cartella e chiede solo 2 opzioni (ricorsivo / sovrascrivi). Il log viene scritto in `conversion.log` dentro la cartella scelta.

### Requisiti

* **Backend Word (consigliato)**: Windows + Microsoft Word + `pip install pywin32`
* **Fallback LibreOffice**: avere `soffice` nel PATH
* **Inserimento metadati lato fallback**: `pip install pypdf` (lo script legge i core properties da `docProps/core.xml` del DOCX)

### Note tecniche

* Con **Word/COM** uso `ExportAsFixedFormat` con `IncludeDocProps=True` e `CreateBookmarks=Heading` (configurabile), così:

  * i **segnalibri** derivano dagli *stili Titolo* (o da bookmark Word, se richiesto);
  * i **metadati** (Titolo, Autore, Soggetto, Parole chiave) vengono riportati nel PDF.
* Con **LibreOffice**, i segnalibri sono generalmente creati dagli stili di intestazione; dopo la conversione lo script **inietta i metadati** nel PDF leggendo i core properties del DOCX (Title, Author, Subject, Keywords).
