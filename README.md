# PDF Drawing Generator (Excel → Fillable PDF → Flattened Output)

Generates flattened PDF drawing forms from standardized Excel input and saves outputs with safe, consistent filenames.

## What it does
- Reads a standardized Excel schema (consistent columns)
- Cleans/validates fields (trims whitespace, handles blank optional fields)
- Fills a blank fillable PDF template (fields filled by name)
- Flattens the PDF and saves as:
  `Customer - Job Name - Drawing Name.pdf`
- (Planned) Merge multiple generated PDFs into a single “packet” per job

## Why it matters
This project mirrors real operations automation work: standardize inputs → automate output generation → reduce manual effort and errors.

## Results (real-world)
Reduced a high-volume workflow from days to minutes and enabled same-day turnaround for quote + drawing deliverables.

## Project structure
- `src/` - Python scripts
- `sample_data/` - sanitized sample Excel inputs
- `templates/` - sample fillable PDF template
- `output_samples/` - optional screenshots/outputs

## Requirements
- Windows
- Adobe Acrobat (for PDF form automation)

## Data
This repo uses sample/sanitized data only. Do not upload proprietary templates or customer information.


## Run locally
```bash
pip install -r requirements.txt
python src/main.py


