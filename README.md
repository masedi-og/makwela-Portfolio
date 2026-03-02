# Desktop Agent MVP

Windows desktop app for processing missing supplier images from an Excel input and copying TIFF outputs into supplier upload folders.

## Features

- Loads `.xlsx` input with required columns.
- Supplier multi-select.
- Processes CHILD / PARENT / RANGE based on `Missing ... Image == Yes` flags.
- RANGE processing runs only when `Range 1 Name` or `Range 1 Number` is populated.
- Searches in supplier source folders (`PreEdit`, `Master`) for `.tif/.tiff` only.
- Matching rules:
  - filename starts with `Catalogue Code` (case-insensitive), or
  - filename contains catalogue code as token with boundaries space/dash/underscore.
- Multi-image handling:
  - primary: `CatalogueCode.tif/.tiff`
  - secondaries: `CatalogueCode_2`, `CatalogueCode_3`, ...
  - `_1` is ignored.
- Ambiguous detection when multiple candidate sets are found.
- Dry run mode (no file copy).
- Overwrite option for existing outputs.
- Writes into `TIFF_WORKING` under destination upload folders.
- Live progress, live logs, results table.
- Export results to CSV or XLSX.

## Config

Edit `config.json` to configure supplier folders.

Included City Office defaults:

- Source PreEdit: `W:\SUPPLIERS\SUPPLIER BRANDS\City Office\HIGH RES\PRE EDIT IMAGES\PRE EDIT - KEEP`
- Source Master: `W:\SUPPLIERS\SUPPLIER BRANDS\City Office\HIGH RES\MASTERED IMAGES`
- Upload Child: `W:\SUPPLIERS\SUPPLIER BRANDS\City Office\LOW RES\UPLOAD\CHILD`
- Upload Parent: `W:\SUPPLIERS\SUPPLIER BRANDS\City Office\LOW RES\UPLOAD\PARENT`
- Upload Range: `W:\SUPPLIERS\SUPPLIER BRANDS\City Office\LOW RES\UPLOAD\UPLOAD RANGE`

## Run

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

## Result schema

Each processed work item includes:

- row number
- supplier
- supplier code
- catalogue code
- process type (CHILD/PARENT/RANGE)
- status (Success/Missing/Ambiguous)
- found source (PreEdit/Master)
- match type (StartsWith/TokenContains/Mixed)
- found files
- output files
- notes
