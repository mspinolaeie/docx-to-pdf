Strumento per convertire file `.docx` in PDF con tre modalità:

* **CLI batch** su cartella (`--dir`).
* **CLI drag-and-drop** — passa file o cartelle come argomenti posizionali (o trascina sull'eseguibile in Explorer).
* **GUI PySide6** senza argomenti, con drag and drop visivo di file e cartelle.

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

### CLI Drag-and-Drop (file espliciti)

Passa uno o più file `.docx` e/o cartelle direttamente come argomenti posizionali. Utile anche per trascinare file sull'eseguibile in Windows Explorer.

```bash
# Converti file singoli
docx-to-pdf-drop relazione.docx offerta.docx

# Mix di file e cartelle
docx-to-pdf-drop C:\docs relazione.docx

# Con opzioni (tutte le flag CLI sono supportate)
docx-to-pdf-drop *.docx --overwrite --use word
```

Tutti i flag CLI (`--overwrite`, `--use`, `--bookmarks`, `--pdfa`, `--recursive`, ecc.) funzionano anche in questa modalità.

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

### Entry point installati

Dopo `pip install -e .` sono disponibili due comandi equivalenti:

* `docx-to-pdf` — uso CLI completo (`--dir`, opzioni avanzate, ecc.);
* `docx-to-pdf-drop` — alias dedicato per DnD/file espliciti; identico a `docx-to-pdf` ma il nome chiarisce l'uso da Explorer.

### Launcher Windows Explorer (DnD senza installazione)

`docx-to-pdf-drop.cmd` è un launcher autonomo per Windows Explorer: trascina uno o più `.docx` (o cartelle) sull'icona del file e la conversione parte immediatamente. Richiede solo Python nel PATH.

Per distribuire lo strumento senza `pip install` bastano due file nella stessa cartella:

```
docx-to-pdf-drop.cmd
docx_to_pdf.py
```

### Build PyInstaller

Lo `spec` genera due eseguibili:

* `docx-to-pdf.exe` per uso console / scripting;
* `docx-to-pdf-gui.exe` per uso finestrato senza terminale.

Entrambi includono la GUI `PySide6` necessaria per la modalità senza argomenti.

### Note tecniche

* Con **Word/COM** viene usato `ExportAsFixedFormat` con `IncludeDocProps=True` e `CreateBookmarks` configurabile, quindi segnalibri e metadati vengono preservati nel PDF.
* Con **LibreOffice**, i segnalibri derivano in genere dagli stili di intestazione; dopo la conversione lo script inietta i metadati nel PDF leggendo le core properties del DOCX.
