# VERTIV Thermal Report Automation

Generate one Word report page per row in `EXCEL/data.xlsx` using `TEMPLATE/TEMPLATE.docx`.

## Usage

```powershell
python generate_report.py
```

Default output:

```text
OUTPUT/thermal_report_part_01.docx
OUTPUT/thermal_report_part_02.docx
...
```

The script splits merged reports into 10-page chunks by default. With the current 43-row workbook, it generates five files: 10 + 10 + 10 + 10 + 3 pages.

Microsoft Word COM automation is used for merging because FLIR Tools+ objects do not survive pure Python DOCX merging as analyzable objects.

To change the chunk size:

```powershell
python generate_report.py --chunk-size 5
```

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

Microsoft Word must be installed because page merging is performed through Word automation.
