# Wafer_Summary_Report

KGD test data processing tool with **folder-prefix scan** and **XY deduplication**.

## Key Differences from Wafer_Yield_Summary_Report

| Feature | Wafer_Yield_Summary_Report | Wafer_Summary_Report (this repo) |
|---------|---------------------------|----------------------------------|
| Folder input | Drag & drop any folder | Scan all subfolders matching Wafer_ID prefix |
| Duplicate dies | All records kept | Same (X,Y) -> latest XML only |

## XY Deduplication Logic

When multiple XML files contain data for the same die coordinate (X_COORD, Y_COORD):
- Compare `os.path.getmtime()` of each XML file
- Keep only the record from the **most recently modified** XML
- Older records for the same coordinate are discarded

This handles retest scenarios where a die has been measured multiple times.

## XML Format Support

Both formats are auto-detected:

| Format | Encoding |
|--------|----------|
| A - Proprietary | hex string -> XOR 0xFF -> string reversal |
| B - Plain | Standard XML read directly |

## Output Reports

| File | Contents |
|------|----------|
| `<WaferID>_<PartNum>_Wafer_Summary.xlsx` | Wafer Info, Test Summary, Color-coded Wafer Map |
| `<WaferID>_<PartNum>_Data_Summary.xlsx` | Per-die metadata + test parameters |

Red cell = VALUE > HIGH_LIMIT or VALUE < LOW_LIMIT

## Save Paths

- Primary: `Z:\ToFTP` (both) + `Z:\KYEC` (Wafer Summary copy)
- Fallback: `C:\KGD_data\Molex_KGD_Data` (when Z: unavailable)

## Build

```
Double-click build_exe.bat
```
