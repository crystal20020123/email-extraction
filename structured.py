import re

patterns = {
        "Cheque Number": r'CHEQUE #:\s*([\d]+)',
        "Cheque Date": r'CHEQUE DATE:\s*(\d{2}/\d{2}/\d{2})',
        "Provider Tax ID": r'PROVIDER TAX ID:\s*(\d+)',
        "Bulk Claim Identifier": r'BULK CLAIM IDENTIFIER:\s*(\d+)',
        # Match multiline until the next uppercase word starts (i.e., a new field)
        "Claimant Name": r'CLAIMANT NAME:\s*([A-Za-z\s\.]+)(?=\n[A-Z])',
        "Policy Number": r'POLICY NUMBER:\s*([\d-]+)',
        "Claim Number": r'CLAIM NUMBER:\s*(\w+)',
        "Dates of Service": r'DATES OF SERVICE\s*(\d{2}/\d{2}/\d{2}-\d{2}/\d{2}/\d{2})',
        "Service": r'SERVICE\s*([A-Z]+)',
        "Reference Number": r'REF. NO\s*(\d+)',
        "Charged Amount": r'CHARGE\s*([\d,.]+)',
        "Excluded Amount": r'EXCLUDED\s*([\d,.]*)',
        "Deductible": r'DEDUCTIBLE\s*([\d,.]*)',
        "Percentage": r'%\s*(\d+)',
        "Payable Amount": r'PAYABLE\s*([\d,.]+)',
        "Remark Code": r'REMARK CODE\s*(\d+)',
        "Provider Payment": r'PROVIDER PAYMENT:\s*([\d,.]+)',
        "Insured Payment": r'INSURED PAYMENT:\s*([\d,.]*)',
        "Transaction Fee": r'TRANSACTION FEE:\s*([\d,.]+)',
        "GCT": r'G\.C\.T\.:\s*([\d,.]+)'
        }

# This function applies all the patterns to the input text and returns a list of dictionaries
def extract_structured_data(text):
    # Initialize an empty list to hold all extracted claim records
    input_text = text.replace('\n\n','\n')
  
    structured_data = []

    # Split the text into sections for each claim
    claims_sections = re.split(r"={10,}", input_text)

    # Loop through each claim section and apply regex patterns
    for section in claims_sections:
        # Skip sections that don't contain relevant data
        if "CLAIMANT NAME:" not in section:
            continue
        # Create a dictionary to store extracted data for a single claim
        claim_data = {}
        for key, pattern in patterns.items():
            match = re.search(pattern, section, re.MULTILINE)
            claim_data[key] = match.group(1).strip() if match else None
        # Append the extracted data to the list
        structured_data.append(claim_data)

    return structured_data



pattern = re.compile(
    r'\b(?P<Batch_ID>\d{4})?\b\s*'    # Batch ID (optional)
    r'(?P<Date>\d{8})?\s*'            # Date (optional)
    r'(?P<Claim_No>0*[0-9]+)\s*'  # Claim No (optional)
    r'(?P<Member_ID>\d+-\d+-\d+(?:-\d+)?)?\s*'  # Member ID (optional)
    r'(?P<Name>[A-Z ]+[A-Z])\s+'      # Name
    r'(?P<Benefit>[A-Z]+)?\s*'        # Benefit (optional)
    r'(?P<Charged_Amount>-?[\d,]+\.\d{2})\s*'   # Charged Amount
    r'(?P<Copay>-?[\d,]+\.\d{2})\s*'  # Copay
    r'(?P<Payable_Amount>-?[\d,]+\.\d{2})\s*'  # Payable Amount
    r'(?P<Fee>-?[\d,]+\.\d{2})\s*'    # Fee
    r'(?P<GCT>-?[\d,]+\.\d{2})\s*'    # GCT
    r'(?P<Total>-?[\d,]+\.\d{2})\s*'  # Total
    r'(?P<Rejection_Reason>REVERSED|CLAIM REVERSED)?', # Rejection Reason (if exists)
    re.MULTILINE
)


def extract_structured_data_guardian(text):
    data = text.replace('\n\n','\n')
    matches = pattern.findall(data)
    # print(matches)

    # Create a list of dictionaries for the structured data
    list = []
  
    matches = pattern.finditer(text)

    # Iterate through the matches and print the captured groups
    for match in matches:
        match_dict = match.groupdict()
        # Cleaning up 'null' values and adapting the format if necessary
        structured_data = {k: (v.strip() if v else None) for k, v in match_dict.items()}
        # print(structured_data)
        list.append(structured_data)

    return list
