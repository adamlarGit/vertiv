# VERTIV Thermal Report Automation

Generate one Word report page per row in `EXCEL/data.xlsx` using `TEMPLATE/TEMPLATE.docx`.

## Usage

```powershell
python generate_report.py
```

Default output:

```text
OUTPUT/thermal_report.docx
```

The default merge engine uses Microsoft Word COM automation because FLIR Tools+ objects do not survive the pure Python DOCX merge as analyzable objects.

## Inputs

- `EXCEL/data.xlsx`
- `TEMPLATE/TEMPLATE.docx`

Temporary constants live near the top of `generate_report.py`:

- `ptu.no`
- `model`
- `rating`

If those columns are later added to Excel, non-blank Excel values override the constants row by row.

## Dependencies

```powershell
pip install -r requirements.txt
```

Microsoft Word must be installed for the default `--merge-engine word` mode.
