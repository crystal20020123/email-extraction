import pdfplumber

pdf_path = "test.pdf"

with pdfplumber.open(pdf_path) as pdf:
    print(pdf)
    for page in pdf.pages:
        # Extract the table from the PDF
        table = page.extract_table()
        # Do something with the extracted table...
        for row in table:
            print(row)