# Automation Assignment Submission Notes

This project follows the `Python + Playwright` format described in the assignment instructions.

## Project files

- `run_assignment.py`: main automation script
- `image_preview_test.py`: small wrapper that runs the same main script
- `sample.png`: upload file used by the test
- `execution_results.csv`: CSV result file
- `execution_results.xlsx`: Excel result file created with `openpyxl`
- `results/`: screenshots saved during execution

## What the script does

1. Opens `https://www.pixelssuite.com/convert-to-png`
2. Uploads a PNG file
3. Checks whether preview evidence appears
4. Saves screenshots before and after upload
5. Writes the result to both CSV and Excel files

## One-time installation

Run these commands in the project folder:

```powershell
pip install -U pip
pip install playwright openpyxl
playwright install
```

If your machine uses `py`, you can also run:

```powershell
py -m pip install -U pip
py -m pip install playwright openpyxl
py -m playwright install
```

## Run the assignment

```powershell
python run_assignment.py
```

Or:

```powershell
py run_assignment.py
```

## Output files

- `results/01_before_upload.png`
- `results/02_after_upload_pass.png` or `results/02_after_upload_fail.png`
- `execution_results.csv`
- `execution_results.xlsx`

## Note

- If `sample.png` does not exist, the script creates a tiny valid PNG automatically.
- The CSV and Excel files are refreshed on each run so the latest result is easy to submit.
