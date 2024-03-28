import requests
from urllib.parse import urlparse, parse_qs
import os
import PyPDF2
from io import BytesIO
from urllib.parse import unquote
from urllib.parse import unquote
from io import StringIO, BytesIO
from pdfminer.high_level import extract_text_to_fp
from pdfminer.layout import LAParams
# Assuming 'input_data' is provided by Zapier
# encoded_url = input_data['attachment']
encoded_url = "https://zapier.com/engine/hydrate/17180415/.eJwtzctqwzAQheF3mbXV6urbLptCnyEEI49GRSSRjCW3BON3r2q6_c_wzQ6YYqFYpvJaCEa4QAMh5mIj0hQcjFoZMbSyawC3XNJzy7Seg-hEz7UwDVjEtFXirNpIbjRvwAd6uCna5x_rw4Nype8_dv3KMO5nmZYU6ve1husOd3rVy6ze217P6K1jsybHNHHJZoMD67zqyDg3IFfV-qc_43cKSEy2Un1wKRnnSl-V6IW6vS3Ow3E7jl9YEUdB:1r800D:PGztx83V3p4Q2INpcpTcnKbiX7E/"

try:
    # Make a HEAD request to the encoded URL
    response = requests.head(encoded_url)
    decoded_url = unquote(encoded_url)
    response1 = requests.get(decoded_url, stream=True)
    content = None

    contains_pdf = response.headers.get('Content-Type', '').lower() == 'application/pdf'
    file_name = None

    if contains_pdf:

        content = response
        pdf_full_text= None

        file_like_object = BytesIO(response1.content)
        output_string = StringIO()
        laparams = LAParams()


        extract_text_to_fp(file_like_object, output_string, laparams=laparams)

        text_content = output_string.getvalue()

        content_disposition = response.headers.get('Content-Disposition')
        
        if content_disposition:
            # Parse the filename from the Content-Disposition header
            disposition_parts = content_disposition.split(';')
            for part in disposition_parts:
                if 'filename=' in part:
                    file_name = part.replace('filename=', '').strip().strip('"')
                    break
        else:
            # Fallback to parsing the filename from the URL if no Content-Disposition header
            url_path = urlparse(encoded_url).path
            file_name = os.path.basename(url_path)

except Exception as e:
    # Return an error message in case of an exception
    result = {'error': str(e)}
else:
    # Return a dictionary with the result
    result = {'contains_pdf': contains_pdf, 'filename' : file_name, 'content' : text_content}

# You would typically send 'result' back to Zapier or use it in further processing
print(result)