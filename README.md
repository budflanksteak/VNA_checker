# VNA Accession Checker

Read an Excel file of accession numbers, query a DICOM **VNA** (Vendor Neutral
Archive) by accession number using a **DIMSE C-FIND**, determine whether images
are stored for each accession and which modalities are present, and write the
results back into the **same** Excel workbook.

## What it does

For each accession number in the workbook the script performs a STUDY-level
C-FIND against the Study Root Query/Retrieve Information Model and writes back
these columns:

| Column           | Meaning                                                        |
| ---------------- | ------------------------------------------------------------- |
| `Images Present` | `Yes` if one or more studies matched, `No` if none, `Error` on failure |
| `Modality`       | Comma-separated modalities in the study/studies (e.g. `CT, MR`) |
| `Study Count`    | Number of studies matching the accession number               |
| `Instance Count` | Total DICOM instances across the matching studies             |
| `Query Notes`    | `OK`, `No matching studies in VNA`, or an error detail         |

If the VNA does not return `ModalitiesInStudy`, the script falls back to a
SERIES-level query to collect the modalities (disable with `--no-series-fallback`).

Result columns are matched by header name, so re-running updates the same
columns in place instead of appending duplicates.

## Install

```bash
pip install -r requirements.txt
```

## Usage

```bash
python vna_accession_checker.py accessions.xlsx \
    --host 10.0.0.5 --port 104 \
    --called-aet VNA_SCP --calling-aet VNA_CHECK
```

Connection settings can also come from environment variables:

| Argument         | Environment variable |
| ---------------- | -------------------- |
| `--host`         | `VNA_HOST`           |
| `--port`         | `VNA_PORT`           |
| `--called-aet`   | `VNA_CALLED_AET`     |
| `--calling-aet`  | `VNA_CALLING_AET`    |

### Options

- `--accession-column` — header name, column letter (e.g. `B`), or column number
  for the accession column. Auto-detected from the header row if omitted (looks
  for headers like *Accession Number*, *Accession*, *AccNum*, …).
- `--sheet` — worksheet name to process (defaults to the active sheet).
- `--timeout` — DICOM network/association/DIMSE timeout in seconds (default 30).
- `--no-series-fallback` — do not do a SERIES-level query when the VNA omits
  `ModalitiesInStudy`.
- `-v` / `--verbose` — debug logging, including the pynetdicom protocol logs.

## Notes

- The input workbook is expected to have a header row in row 1 with the
  accession numbers in one column; every non-empty data row below is queried.
- The VNA must permit C-FIND associations from the calling AE title you use;
  the AE title / IP usually has to be whitelisted on the VNA side first.
- Verify the calling AE title, called AE title, host, and port with your PACS/VNA
  administrator before running against production.
