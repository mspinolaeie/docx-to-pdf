Strumento per convertire file `.docx` in PDF con due modalità:

* **CLI batch** su cartella (`--dir`), come prima.
* **GUI PySide6** senza argomenti, con drag and drop di file e cartelle.

Su **Windows** con **Microsoft Word** installato (via COM) ottieni la massima fedeltà: segnalibri da intestazioni o bookmark Word e metadati inclusi nel PDF. In assenza di Word, viene usato il fallback **LibreOffice** (`soffice`), con iniezione dei metadati leggendo `docProps/core.xml` dal DOCX.

### Uso rapido CLI

```bash
# Converti tutti i .docx nella cartella corrente
python docx_to_pdf.py --dir .

# Ricorsivo su sottocartelle, sovrascrive PDF già esistenti
python docx_to_pdf.py --dir . --recursive --overwrite

# Forza backend (auto | word | libreoffice)
python docx_to_pdf.py --use word

# Segnalibri: headings | word | none
python docx_to_pdf.py --bookmarks headings

# Esporta in PDF/A-1 (solo backend Word/COM)
python docx_to_pdf.py --pdfa
```

### GUI Drag and Drop

Se avvii `docx_to_pdf.py` **senza argomenti**, si apre una GUI `PySide6` con:

* area drag and drop per `.docx` e cartelle;
* espansione cartelle solo al primo livello;
* coda esplicita dei file da convertire;
* opzioni base: `overwrite`, `backend`, `workers`;
* opzioni avanzate: `bookmarks`, `pdfa`, `validate_pdf`, `log_level`;
* log e progresso in tempo reale;
* output PDF sempre accanto al file sorgente.

La conversione parte solo con il pulsante `Converti`.

### Requisiti

* **GUI**: `pip install .[gui]` oppure `pip install PySide6`
* **Backend Word (consigliato)**: Windows + Microsoft Word + `pip install pywin32`
* **Fallback LibreOffice**: avere `soffice` nel PATH
* **Inserimento metadati lato fallback**: `pip install pypdf`
* **Build completa**: `pip install .[build]`

### Build PyInstaller

Lo `spec` genera due eseguibili:

* `docx-to-pdf.exe` per uso console / scripting;
* `docx-to-pdf-gui.exe` per uso finestrato senza terminale.

Entrambi includono la GUI `PySide6` necessaria per la modalità senza argomenti.

### Note tecniche

* Con **Word/COM** viene usato `ExportAsFixedFormat` con `IncludeDocProps=True` e `CreateBookmarks` configurabile, quindi segnalibri e metadati vengono preservati nel PDF.
* Con **LibreOffice**, i segnalibri derivano in genere dagli stili di intestazione; dopo la conversione lo script inietta i metadati nel PDF leggendo le core properties del DOCX.
