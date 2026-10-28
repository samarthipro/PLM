import pdfplumber
import pandas as pd
import re
import os

try:
    from template_schema import conform_to_template
except ImportError:
    from scripts.template_schema import conform_to_template

# =========================
# CARRIER CONFIGURATION
# =========================

CARRIER_NAME = "PLM"

# Coverage-code -> broad Coverage Type, per Mapping file.xlsx's "PLM" sheet lookup table
# (columns D/E). CoverageCode is the 3-letter prefix PLM's own "Loss Type: XXX - ..." line
# already gets split out into (see parse_claim_block's LossType/CoverageCode extraction below).
_COVERAGE_TYPE_LOOKUP = {
    "ABI": "AL", "AMP": "AL", "APD": "APD", "AUD": "AL", "AUM": "AL",
    "BIU": "PROP", "BLD": "PROP", "COL": "APD", "COM": "APD", "CRM": "PROP",
    "CTN": "PROP", "EPL": "EPL", "FID": "FID", "INL": "PROP", "OBI": "GL",
    "OMP": "GL", "OPD": "GL", "PBI": "GL", "PCD": "GL", "PIP": "APD",
    "PK2": "PROP", "PK3": "PROP", "PPD": "GL", "STK": "PROP",
}

# =========================
# HELPERS
# =========================

def clean_text(text):
    if not text:
        return ""

    text = text.replace("\x0c", " ")
    text = text.replace("", "")
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\n+', '\n', text)

    return text.strip()


def clean_money(val):
    if not val:
        return 0

    val = str(val).strip()
    
    negative = False

    if val.startswith("-"):
        negative = True
    val = (
        val.replace("$", "")
        .replace(",", "")
        .replace("(", "")
        .replace(")", "")
        .strip()
    )
    try:
        amt = float(val)

        if negative:
            amt = -abs(amt)

        return amt
    except:
        return 0


def split_claim(val):
    val = str(val).strip().replace(" ", "")

    if "-" in val:
        return val.split("-", 1)

    return val, None


def parse_address(addr):
    if not addr:
        return None, None, None, None

    addr = clean_text(addr)

    addr = addr.replace("\n", " ")

    # Match city/state/zip from end
    m = re.search(
        r'(.+?)\s+([A-Za-z .\-]+),\s*([A-Z]{2})\s*(\d{5})?',
        addr
    )

    if m:
        full_addr = m.group(1).strip()
        city = m.group(2).strip()
        state = m.group(3).strip()
        zipc = m.group(4)

        return addr, city, state, zipc

    return addr, None, None, None


# =========================
# PDF TEXT EXTRACTION
# =========================

def is_claim_detail_page(text):
    """Check if a page contains claim detail data (not summary)."""
    if not text:
        return False

    # Claim detail pages have the 'Claim #' column header
    if 'Claim #' in text:
        return True

    # Also detect continuation pages that have claim data
    # but may not repeat the header
    if re.search(r'Claimant:', text):
        return True

    return False


def extract_pdf_text(path):

    all_text = []
    metadata_text = []  # collect ALL pages for metadata

    with pdfplumber.open(path) as pdf:

        for page in pdf.pages:

            text = page.extract_text(
                x_tolerance=2,
                y_tolerance=2
            )

            if text:
                text = clean_text(text)
                metadata_text.append(text)

                # Only include claim detail pages
                if is_claim_detail_page(text):
                    all_text.append(text)

    claim_text = "\n".join(all_text)
    meta_text = "\n".join(metadata_text)

    return claim_text, meta_text


# =========================
# GLOBAL METADATA
# =========================

def extract_account(text):

    patterns = [
        r'Account #:\s*([A-Z0-9]+)',
        r'Policy #:\s*([A-Z0-9]+)'
    ]

    for pat in patterns:
        m = re.search(pat, text)

        if m:
            return m.group(1)

    return None


def extract_valuation_date(text):

    m = re.search(r'Report Date\s*(\d{1,2}/\d{1,2}/\d{4})', text)

    if m:
        try:
            return pd.to_datetime(m.group(1))
        except:
            pass

    return None


# =========================
# CLAIM BLOCK SPLITTER
# =========================

def has_claim_number(text):
    # Standard: claim number with dash and digits on same line
    if re.search(r'[A-Z0-9]{4,}-\d+', text):
        return True

    # Split: claim number ending with dash (ID is on next line)
    if re.search(r'\d{4}[A-Z]\d+-', text):
        return True

    return False


def merge_split_claim_numbers(text):
    lines = text.splitlines()
    merged = []
    i = 0

    while i < len(lines):
        line = lines[i]

        # Check if this line has a claim number ending with dash
        # but no digits after it (split format)
        m = re.search(r'(\d{4}[A-Z]\d+-)\s', line)

        if m and not re.search(r'\d{4}[A-Z]\d+-\d+', line):
            claim_prefix = m.group(1)

            # Look at the next line for the claimant ID
            if i + 1 < len(lines):
                next_line = lines[i + 1].strip()

                # Pattern: "12:00:00 645496" or just "645496"
                id_match = re.search(r'(\d{5,})', next_line)

                if id_match:
                    claimant_id = id_match.group(1)

                    # Merge: replace "XXXX-" with "XXXX-ID" in current line
                    line = line.replace(
                        claim_prefix,
                        claim_prefix + claimant_id
                    )

                    # Remove the ID from the next line
                    next_cleaned = next_line.replace(
                        claimant_id, ''
                    ).strip()

                    # Keep remaining content of next line if any
                    if next_cleaned and next_cleaned != 'AM':
                        merged.append(line)
                        merged.append(next_cleaned)
                    else:
                        merged.append(line)

                    i += 2

                    # Skip "AM" line if present
                    if i < len(lines) and lines[i].strip() == 'AM':
                        i += 1

                    continue

        merged.append(line)
        i += 1

    return '\n'.join(merged)


def split_claim_blocks(text):

    # First, merge any split claim numbers
    text = merge_split_claim_numbers(text)

    lines = text.splitlines()

    blocks = []

    current_block = []

    claim_start_pattern = re.compile(
        r'^\d{1,2}/\d{1,2}/\d{4}'
    )

    for line in lines:

        line = line.strip()

        if not line:
            continue

        # Skip page headers and repeated column headers
        if re.match(r'^Page \d+ of \d+', line):
            continue
        if 'Date Of Loss' in line or 'Claim Status' in line:
            continue
        if line.startswith('Report Date'):
            continue
        if line.startswith('Account #:') or line.startswith('Policy #:'):
            continue
        if line.startswith('Insured Name:') or line.startswith('Account Name:'):
            continue
        if line.startswith('Broker:'):
            continue
        if 'Cat#' in line and 'Paid' in line:
            continue

        # New claim starts
        if claim_start_pattern.match(line):

            # Save previous block
            if current_block:

                joined = "\n".join(current_block)

                if has_claim_number(joined):
                    blocks.append(joined)

            current_block = [line]

        else:
            current_block.append(line)

    # Last block
    if current_block:

        joined = "\n".join(current_block)

        if has_claim_number(joined):
            blocks.append(joined)

    return blocks


# =========================
# FIELD EXTRACTION
# =========================

def extract_field(block, label):

    pattern = rf'{re.escape(label)}\s*:?\s*(.*?)(?=\n[A-Z][A-Za-z ]+\s*:|\n\d{{1,2}}/\d{{1,2}}/\d{{4}}|$)'

    m = re.search(pattern, block, re.DOTALL)

    if m:
        return clean_text(m.group(1))

    return None


def extract_loss_desc(block):

    patterns = [
        r'Loss Description\s*:\s*(.*?)(?=\n\d{1,2}/\d{1,2}/\d{4}|$)',
        r'Loss Desc\s*:\s*(.*?)(?=\n\d{1,2}/\d{1,2}/\d{4}|$)'
    ]

    for pat in patterns:

        m = re.search(pat, block, re.DOTALL)

        if m:

            desc = clean_text(m.group(1))

            # Remove footer garbage
            desc = re.split(
                r'Note: Some Claim Numbers',
                desc
            )[0]
            
            # Ensure it is a single line, formatted sentence
            desc = re.sub(r'\s+', ' ', desc).strip()
            
            return desc

    return None


def extract_money_values(block):

    money_matches = re.findall(
        r'-?\$[\d,]+\.\d{2}',
        block
    )

    cleaned = [clean_money(x) for x in money_matches]

    if len(cleaned) >= 3:
        return cleaned[0], cleaned[1], cleaned[2]

    if len(cleaned) == 2:
        return cleaned[0], cleaned[1], 0

    if len(cleaned) == 1:
        return cleaned[0], 0, 0

    return 0, 0, 0


# =========================
# CLAIM PARSER
# =========================

def parse_claim_block(block, account, valuation_date):

    data = {
        "Valuation Date": valuation_date,
        "Lookup #": None,
        "Account": account,
        "ClaimNumber": None,
        "ClaimantID": None,
        "CoverageCode": None,
        "PaidLoss": 0,
        "Outstanding": 0,
        "Expense": 0,
        "Lossdate": None,
        "Status": None,
        "ClaimantName": None,
        "DriverName": None,
        "LossType": None,
        "LossDesc": None,
        "LossAddress": None,
        "LossCity": None,
        "LossState": None,
        "LossZip": None,
        "Litigated": None
    }

    block = clean_text(block)

    lines = [x.strip() for x in block.splitlines() if x.strip()]

    full_text = "\n".join(lines)

    # =====================
    # LOSS DATE
    # =====================

    for line in lines:

        m = re.match(r'^\d{1,2}/\d{1,2}/\d{4}', line)

        if m:
            try:
                data["Lossdate"] = pd.to_datetime(m.group())
                break
            except:
                pass

    # =====================
    # STATUS
    # =====================

    status_match = re.search(r'\b(Open|Closed)\b', full_text)

    if status_match:
        data["Status"] = status_match.group(1)

    # =====================
    # CLAIM NUMBER
    # =====================

    claim_match = re.search(r'([A-Z0-9]+-\d+)', full_text)

    if claim_match:

        claim_no, claimant_id = split_claim(
            claim_match.group(1)
        )

        data["ClaimNumber"] = claim_no
        data["ClaimantID"] = claimant_id

    # =====================
    # FINANCIALS
    # =====================

    money_vals = re.findall(
        r'-?\$[\d,]+\.\d{2}',
        full_text
    )

    cleaned_money = [clean_money(x) for x in money_vals]

    if len(cleaned_money) >= 1:
        data["PaidLoss"] = cleaned_money[0]

    if len(cleaned_money) >= 2:
        data["Outstanding"] = cleaned_money[1]

    # =====================
    # LOSS TYPE
    # =====================

    m = re.search(
        r'Loss Type:\s*(.+)',
        full_text
    )

    if m:

        lt = m.group(1).strip()

        data["LossType"] = lt

        cov = lt.split("-")[0].strip()

        data["CoverageCode"] = cov

    # =====================
    # CLAIMANT
    # =====================

    m = re.search(
        r'Claimant:\s*(.*?)(?=\s*(?:Examiner|Note)\s*:|\n\s*(?:Injury|Cause|Claim|Location|Driver|Loss Desc)\b|\Z)',
        full_text,
        re.DOTALL
    )

    if m:
        raw_name = m.group(1).replace('\n', ' ').strip()
        data["ClaimantName"] = re.sub(r'\s+', ' ', raw_name)

    # =====================
    # DRIVER
    # =====================

    driver_name = None

    for i, line in enumerate(lines):

        if re.match(r'^Driver\s*:', line):

            # take only text on SAME line
            val = re.sub(
                r'^Driver\s*:\s*',
                '',
                line
            ).strip()

            # if nothing after Driver:, keep blank
            if val and not val.startswith("Loss Desc"):
                driver_name = val

            break

    data["DriverName"] = driver_name

    # =====================
    # LOCATION
    # =====================

    location_match = re.search(
        r'^Claim\b(.*?)\nLocation:(.*?)(?=^Driver|^Loss Desc|^Loss Description|^Note:|\Z)',
        full_text,
        re.MULTILINE | re.DOTALL
    )

    if not location_match:
        location_match = re.search(
            r'Claim\s*Location:\s*(.*?)(?=Driver|Loss Desc|Loss Description|Note:|$)',
            full_text,
            re.DOTALL
        )

    if location_match:

        if len(location_match.groups()) == 2:
            raw_loc = location_match.group(1).strip() + " " + location_match.group(2).strip()
        else:
            raw_loc = location_match.group(1).strip()

        location = clean_text(raw_loc)
        
        # Remove footer garbage
        location = re.split(r'Note:\s*Some Claim Numbers', location)[0].strip()

        (
            data["LossAddress"],
            data["LossCity"],
            data["LossState"],
            data["LossZip"]
        ) = parse_address(location)

        if data["LossAddress"]:
            data["LossAddress"] = data["LossAddress"].replace(",", "").strip()
            if not data["LossAddress"]:
                data["LossAddress"] = None

    # =====================
    # DESCRIPTION
    # =====================

    loss_desc = None

    for i, line in enumerate(lines):

        if re.match(r'^(Loss Desc|Loss Description)\s*:', line):

            first_line = re.sub(
                r'^(Loss Desc|Loss Description)\s*:\s*',
                '',
                line
            ).strip()

            desc_lines = [first_line] if first_line else []

            for next_line in lines[i + 1:]:

                clean_next = next_line.strip()

                if re.match(
                    r'^(Examiner\s*:|Injury\s*:|Cause\s*:|Claim\s+Location|Location\s*:|Driver\s*:|Loss Type\s*:|Date Of Loss|Note:|Total\s*:?)',
                    clean_next
                ):
                    break

                if clean_next:
                    desc_lines.append(clean_next)

            # --- CHANGES MADE HERE ---
            # 1. Join with a space (" ") rather than newline ("\n")
            raw_desc = " ".join(desc_lines)
            
            # 2. Use regular expression to squish any multiple spaces/tabs/newlines into a single space,
            #    and finally .strip() to eliminate trailing and leading spaces.
            loss_desc = re.sub(r'\s+', ' ', raw_desc).strip()
            # -------------------------

            break

    data["LossDesc"] = loss_desc

    # =====================
    # LITIGATED
    # =====================

    lit_match = re.search(
        r'Litigated:\s*(\w+)',
        full_text,
        re.IGNORECASE
    )

    if lit_match:

        data["Litigated"] = lit_match.group(1).strip().title()


    # =====================
    # LOOKUP
    # =====================

    data["Lookup #"] = (
        f"{data['ClaimNumber']} - "
        f"{data['ClaimantID']} - "
        f"{data['CoverageCode'] or ''}"
    )

    return data


# =========================
# FILE PROCESSOR
# =========================

def process_file(path):

    print(f"\nProcessing: {os.path.basename(path)}")

    claim_text, meta_text = extract_pdf_text(path)

    account = extract_account(meta_text)

    valuation_date = extract_valuation_date(meta_text)

    claim_blocks = split_claim_blocks(claim_text)

    print(f"Claim Blocks Found: {len(claim_blocks)}")

    results = []

    for idx, block in enumerate(claim_blocks, start=1):

        try:

            parsed = parse_claim_block(
                block,
                account,
                valuation_date
            )

            if parsed["ClaimNumber"]:
                results.append(parsed)

        except Exception as e:

            print(f"❌ Error parsing claim #{idx}")
            print(str(e))

    return pd.DataFrame(results)


# =========================
# MAIN EXTRACTION
# =========================

def extract(file_path: str) -> pd.DataFrame:
    """Entry point for the UI application."""
    
    df = process_file(file_path)

    columns = [
        "Valuation Date",
        "Lookup #",
        "Account",
        "ClaimNumber",
        "ClaimantID",
        "CoverageCode",
        "PaidLoss",
        "Outstanding",
        "Expense",
        "Lossdate",
        "Status",
        "ClaimantName",
        "DriverName",
        "LossType",
        "LossDesc",
        "LossAddress",
        "LossCity",
        "LossState",
        "LossZip",
        "Litigated"
    ]

    if df.empty:
        return pd.DataFrame(columns=columns)

    # Ensure columns exist
    for col in columns:
        if col not in df.columns:
            df[col] = None

    df = df[columns]

    # Format dates
    for col in ["Valuation Date", "Lossdate"]:
        df[col] = pd.to_datetime(
            df[col],
            errors="coerce"
        ).dt.strftime("%m/%d/%Y")

    return _to_template(df)


# ================= TEMPLATE CONFORMANCE =================
# Reshapes PLM's own short/internal column names onto the shared 38-column Template schema (see
# Mapping file.xlsx, "PLM" sheet). PLM's PDFs report a single undifferentiated Paid/Outstanding
# figure per claim rather than splitting Medical/Legal/Indemnity, so those map to the Template's
# "...NOC" (not-otherwise-classified) columns rather than the split ones, which stay blank — same
# as the mapping file itself shows for PLM. "Expense" is dropped, not renamed: financial_mapper.py
# already treats it as a redundant Paid+Reserve total, and the Template has no slot for it. Policy
# Number, Policy Term, Occupation, Class Code and VIN stay blank because PLM's PDFs don't carry
# them per claim today (see extract()'s own field list above) — not guessed at.
def _to_template(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return conform_to_template(df)

    out = df.copy()
    out["Coverage Type"] = out["CoverageCode"].astype(str).map(_COVERAGE_TYPE_LOOKUP).fillna("")

    return conform_to_template(out, rename_map={
        "Account": "Client",
        "ClaimNumber": "Claim Number",
        "Lossdate": "Loss Date",
        "ClaimantName": "Claimant Name",
        "DriverName": "Driver Name",
        "LossType": "Claim Type",
        "LossDesc": "Loss Description/Event Description",
        "LossAddress": "Loss Location",
        "LossCity": "Loss City",
        "LossState": "State",
        "PaidLoss": "Paid/Loss NOC",
        "Outstanding": "Reserve NOC",
    })


# ================= LOCAL FILE RUNNER =================
LOCAL_PLM_DIRECTORY = r"C:\Users\Samarth\Documents\LLama\PLM"


def process_local_pdfs(directory: str = LOCAL_PLM_DIRECTORY) -> None:
    """Extract every PDF in directory and save its Excel output beside it."""
    if not os.path.isdir(directory):
        raise FileNotFoundError(f"PLM directory not found: {directory}")

    pdf_files = sorted(
        filename
        for filename in os.listdir(directory)
        if filename.lower().endswith(".pdf")
    )

    if not pdf_files:
        raise FileNotFoundError(f"No PDF files found in: {directory}")

    for filename in pdf_files:
        input_path = os.path.join(directory, filename)
        output_name = f"{os.path.splitext(filename)[0]}_output.xlsx"
        output_path = os.path.join(directory, output_name)

        extracted_data = extract(input_path)
        extracted_data.to_excel(output_path, index=False)
        print(f"✅ Output saved: {output_path}")


if __name__ == "__main__":
    process_local_pdfs()
