PLM Loss Run PDF Conversion
Overview
This Python module extracts claim-level information from PLM loss run PDF reports, transforms the extracted data into a standardized insurance loss run schema, and returns a DataFrame that can be used by a UI application or downstream processing workflow.

The script uses pdfplumber to extract PDF text and pandas to structure and format claim data. It also applies PLM-specific coverage mappings and conforms the output to a shared template schema.
Key Features
- Extracts text from PDF pages.
- Identifies claim-detail pages and excludes summary-only pages from claim parsing.
- Extracts report metadata, including account/policy number and report valuation date.
- Detects and merges claim numbers that are split across PDF lines.
- Splits report text into individual claim blocks.
- Parses claim details, including claim number, claimant ID, loss date, status, claimant, driver, coverage code, loss description, location, financials, and litigation status.
- Normalizes money values and date formats.
- Maps PLM coverage codes to broad coverage types.
- Converts PLM-specific field names into the shared template schema.
Requirements
Python 3.9 or later is a reasonable starting point, but the script's exact supported version has not been specified. Install the required libraries in your environment:
pip install pdfplumber pandas
The script also depends on the project module that provides conform_to_template. Depending on the project layout, it imports this function from either:
from template_schema import conform_to_template
or:
from scripts.template_schema import conform_to_template
Make sure the appropriate module is available in the project before running the converter.
Project Structure
A suggested layout is:
project/
├── plm_converter.py
├── template_schema.py
├── scripts/
│   └── template_schema.py
├── input/
│   └── sample_loss_run.pdf
└── output/
Use the layout that matches your existing application. The input/ and output/ folders are illustrative; the provided module does not itself define a batch-folder or Excel-export workflow.
How It Works
1. Read the PDF: Open the PDF and extract text from each page.
2. Identify claim pages: Include pages containing the Claim # header or claim-detail continuation information such as Claimant:.
3. Read report metadata: Search all extracted pages for the account/policy number and report date.
4. Prepare claim blocks: Merge supported split claim numbers, skip repeated headers and page labels, and split text into claim-level blocks.
5. Parse claims: Extract the fields supported by the PLM report layout.
6. Normalize values: Clean text, convert financial values to numbers, and parse dates.
7. Apply coverage mapping: Convert recognized PLM coverage codes into broad coverage types.
8. Conform the schema: Rename and map fields into the shared template using conform_to_template.
9. Return results: The public extract(file_path) function returns a template-conformed pandas DataFrame.
Extracted Fields
The internal extraction DataFrame includes the following fields:
Field	Description
Valuation Date	Report valuation date, formatted as MM/DD/YYYY
Lookup #	Composite lookup value built from claim number, claimant ID, and coverage code
Account	Account or policy number found in the report
ClaimNumber	Claim number
ClaimantID	Claimant identifier parsed from the claim number
CoverageCode	Three-letter code extracted from the loss type
PaidLoss	First parsed monetary value for the claim
Outstanding	Second parsed monetary value for the claim
Expense	Third parsed monetary value; not mapped to the shared template
Lossdate	Date of loss, formatted as MM/DD/YYYY
Status	Claim status, such as Open or Closed
ClaimantName	Claimant name
DriverName	Driver name, when present
LossType	Full loss type text
LossDesc	Loss description/event description
LossAddress	Parsed loss location/address
LossCity	Parsed city
LossState	Parsed state
LossZip	Parsed ZIP code
Litigated	Litigation indicator, when present


The final output columns and layout are controlled by the shared conform_to_template function and its template schema.
Template Mapping Notes
The module applies these PLM-specific mappings before template conformance:
- Account → Client
- ClaimNumber → Claim Number
- Lossdate → Loss Date
- ClaimantName → Claimant Name
- DriverName → Driver Name
- LossType → Claim Type
- LossDesc → Loss Description/Event Description
- LossAddress → Loss Location
- LossCity → Loss City
- LossState → State
- PaidLoss → Paid/Loss NOC
- Outstanding → Reserve NOC
The coverage lookup maps recognized PLM codes to broader categories such as AL, APD, PROP, EPL, FID, and GL.
Fields not supplied by the PLM PDF
According to the script's mapping notes, the PLM report does not provide the following fields per claim, so the module does not infer them:
- Policy Number
- Policy Term
- Occupation
- Class Code
- VIN
The script also does not map Expense into the shared template because its mapping notes treat it as redundant with paid and reserve values and the template has no dedicated slot for it.
Usage
The module exposes this entry point:
import plm_converter

df = plm_converter.extract(r"path\to\loss_run.pdf")
print(df.head())
Replace plm_converter with the actual Python filename (without .py) and provide the full path to the input PDF.

The function returns a pandas DataFrame. If no claim records are found, it returns an empty DataFrame with the module's expected internal columns, passed through the template-conformance function.
Optional: Save the returned DataFrame to Excel
If you want to export the result from a separate runner script:
df.to_excel(r"path\to\output.xlsx", index=False)
This example shows how to export the returned DataFrame; Excel export is not performed directly by the provided module.
Error Handling and Validation
- The PDF is opened using pdfplumber.
- Individual claim parsing errors are printed to the console, and processing continues with subsequent claim blocks.
- A parsed record is included only when a claim number is present.
- Date conversion uses coercion for invalid or missing values before formatting.
- Unrecognized coverage codes map to an empty coverage type.
- Monetary strings that cannot be converted are returned as 0 by the current cleaning helper.
Recommended checks after conversion
1. Compare the number of claims in the PDF with the number of output rows.
2. Verify claim numbers and claimant IDs, especially where the PDF splits a claim number across lines.
3. Spot-check paid and outstanding amounts against the source report.
4. Verify loss dates, statuses, loss descriptions, and locations.
5. Review records with missing coverage types or incomplete address details.
6. Confirm that the final DataFrame matches the expected shared template schema.
Limitations
- Parsing depends on the PLM PDF's text layout and the patterns implemented in the script.
- Scanned/image-only PDFs may require OCR; OCR is not implemented in this module.
- Address parsing expects a recognizable city/state/ZIP pattern and may not split every address correctly.
- Claim financials are assigned based on the order of monetary values found in each claim block. Verify this ordering against representative source PDFs.
- Fields absent from the source report are left blank rather than inferred.
- The module relies on conform_to_template; the final schema depends on that function's implementation.
Troubleshooting
Import error for template_schema
- Confirm that template_schema.py exists in the project or that the scripts package is available.
- Run the script from a working directory where the project modules can be imported.
No claims found
- Confirm the PDF contains selectable text.
- Check whether claim-detail pages contain the expected claim headers or claimant labels.
- Test the parser against a representative PLM loss run PDF.
Unexpected financial values
- Check whether the source PDF's monetary values appear in the expected order.
- Review claims with fewer than three monetary values or unusual negative/parenthesized amounts.
Missing location details
- Review the source PDF's location formatting and compare it with the address parsing pattern.
Security and Data Handling
Loss run reports can contain confidential claim and claimant information. Store source PDFs and converted outputs only in approved locations, and avoid sharing real claim data in logs, examples, or public repositories.
License
No license information was provided with the source script. Add the appropriate license and ownership details if this project will be distributed.
