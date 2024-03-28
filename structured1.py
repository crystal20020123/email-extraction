
from openai import OpenAI
# Generating response back from gpt-3.5-turbo
import json
import re
stock_custom_functions = [
    {
        'name': 'extract_stock_info',
        'description': 'Get the structured list data from the body of the input text',
        'parameters': {
            'type': 'object',
            'properties': {
                'cheque_number':{
                    'type':'string',
                    'description':'cheque #'
                },
                'cheque_date': {
                    'type': 'string',
                    'description': 'cheque date'
                },

                'provide_tax_id': {
                    'type': 'string',
                    'description': 'provide tax id'
                },
                'bulk_claim_identifier':{
                    'type':'string',
                    'description' : 'Bulk Claim Identifier'
                },
                'claimant_name': {
                    'type': 'array',
                    "items": {"type": "string"},
                    'description': 'claimant name'
                },     
                'policy_number': {
                    'type': 'array',
                    "items": {"type": "string"},
                    'description': 'policy number'
                },  
                'claim_number': {
                'type': 'array',
                "items": {"type": "string"},
                'description': 'claim number'
                },  
                'dates_of_service': {
                    'type': 'array',
                    "items": {"type": "string"},
                    'description': 'dates of service name'
                }, 
                'service': {
                    'type': 'array',
                    "items": {"type": "string"},
                    'description': 'service'
                },
                'reference_number': {
                    'type': 'array',
                    "items": {"type": "string"},
                    'description': 'reference number'
                },
                'charged_amount': {
                    'type': 'array',
                    "items": {"type": "string"},
                    'description': 'charged amount'
                }, 
                'excluded_amount': {
                    'type': 'array',
                    "items": {"type": "string"},
                    'description': 'excluded amount'
                },
                'deductible_percentage': {
                    'type': 'array',
                    "items": {"type": "string"},
                    'description': 'deductible percentage'
                }, 
                'payable_amount': {
                    'type': 'array',
                    "items": {"type": "string"},
                    'description': 'payable amount'
                }, 
                'remark_code': {
                    'type': 'array',
                    "items": {"type": "string"},
                    'description': 'remark code'
                }, 
                'provider_payment': {
                    'type': 'array',
                    "items": {"type": "string"},
                    'description': 'provide payment'
                }, 
                'insured_payment': {
                    'type': 'array',
                    "items": {"type": "string"},
                    'description': 'insured payment'
                }, 
                'transaction_fee': {
                    'type': 'array',
                    "items": {"type": "string"},
                    'description': 'transaction fee'
                }, 
                'gct': {
                    'type': 'array',
                    "items": {"type": "string"},
                    'description': 'GCT'
                }, 

            }
        }
    }
]



client = OpenAI(api_key="")

def fuction_calling(text):

    response = client.chat.completions.create(
        model = 'gpt-4',
        messages = [{'role': 'user', 'content': text}],
        functions = stock_custom_functions,
        function_call = 'auto'
    )


    json_response = json.loads(response.choices[0].message.function_call.arguments)
    return json_response



test_text = '''PREMIER RADIOLOGY INSTITUTE AND MEDICA
8 CONSTANT SPING ROAD
KINGSTON 10

Sagicor Life Jamaica Limited
28 - 48 Barbados Avenue., Kingston 5, Jamaica W.I.
Telephone: (876) 929 - 8920-9

PAGE

1

CLAIMS PAYMENT STATEMENT

CHEQUE #:

000060090

CHEQUE DATE:

12/22/23

PROVIDER TAX ID:

0000240259

BULK CLAIM IDENTIFIER:

235179060

============================================================================================================================
CLAIMANT NAME:
NICOLE T. LORD

POLICY NUMBER:
0000910000-02000-0008972762

CHEQUE NUMBER:
000060090

CLAIM NUMBER:
233466520I

DATES OF SERVICE
12/12/23-12/12/23

SERVICE
DIAGNOSTIC

REF. NO
40000117832879

CHARGE
15,400.00

EXCLUDED
9,479.80

DEDUCTIBLE
1,204.25

%
80

PAYABLE
3,772.76

REMARK CODE
19

TOTALS:

15,400.00

9,479.80

1,204.25

3,772.76

PROVIDER PAYMENT:

3,696.84

INSURED PAYMENT:

.00

TRANSACTION FEE:

66.02

G.C.T.:

9.90

19-Charge exceed Reasonable & Customary rate for this service.
============================================================================================================================
CLAIMANT NAME:
KIERSTEN R. BOWEN

POLICY NUMBER:
0000910000-04025-0000506502

CHEQUE NUMBER:
000060090

CLAIM NUMBER:
233475805I

DATES OF SERVICE
12/13/23-12/13/23

SERVICE
DIAGNOSTIC

REF. NO
40000117859153

TOTALS:

CHARGE
8,900.00

8,900.00

EXCLUDED
.00

.00

DEDUCTIBLE

.00

.00

%
80

PAYABLE
7,120.00

7,120.00

REMARK CODE

PROVIDER PAYMENT:

6,976.71

INSURED PAYMENT:

.00

TRANSACTION FEE:

124.60

G.C.T.:

18.69

============================================================================================================================
CLAIMANT NAME:
DAHLIA M. CLARKE

POLICY NUMBER:
0000910000-04971-0001028357

CHEQUE NUMBER:
000060090

CLAIM NUMBER:
233475806I

DATES OF SERVICE
12/13/23-12/13/23

SERVICE
DIAGNOSTIC

REF. NO
40000117873136

TOTALS:

CHARGE
9,400.00

9,400.00

EXCLUDED
.00

.00

DEDUCTIBLE

.00

.00

%
80

PAYABLE
7,520.00

7,520.00

REMARK CODE

PROVIDER PAYMENT:

7,368.66

INSURED PAYMENT:

.00

TRANSACTION FEE:

131.60

G.C.T.:

19.74

============================================================================================================================
CLAIMANT NAME:
LATOYO A. DAVY

POLICY NUMBER:
0000910000-05001-0001010971

CHEQUE NUMBER:
000060090

CLAIM NUMBER:
233495942I

DATES OF SERVICE

SERVICE

REF. NO

CHARGE

EXCLUDED

DEDUCTIBLE

%

PAYABLE

REMARK CODE

THIS EXPLANATION OF BENEFITS IS CONTINUED

♀PREMIER RADIOLOGY INSTITUTE AND MEDICA
8 CONSTANT SPING ROAD
KINGSTON 10

Sagicor Life Jamaica Limited
28 - 48 Barbados Avenue., Kingston 5, Jamaica W.I.
Telephone: (876) 929 - 8920-9

PAGE

2

CLAIMS PAYMENT STATEMENT

CHEQUE #:

000060090

CHEQUE DATE:

12/22/23

PROVIDER TAX ID:

0000240259

BULK CLAIM IDENTIFIER:

235179060

============================================================================================================================

12/15/23-12/15/23

DIAGNOSTIC

40000117933583

8,900.00

3,890.60

368.80

80

3,712.48

19

TOTALS:

8,900.00

3,890.60

368.80

3,712.48

PROVIDER PAYMENT:

3,637.78

INSURED PAYMENT:

.00

TRANSACTION FEE:

64.96

G.C.T.:

9.74

19-Charge exceed Reasonable & Customary rate for this service.
============================================================================================================================
CLAIMANT NAME:
RACQUEL M. LAWSON-BURKE

POLICY NUMBER:
0000910000-05462-0008983552

CHEQUE NUMBER:
000060090

CLAIM NUMBER:
233504813I

DATES OF SERVICE
12/16/23-12/16/23

SERVICE
MATERNITY

REF. NO
40000117943294

CHARGE
13,400.00

EXCLUDED
10,000.00

TOTALS:

13,400.00

10,000.00

DEDUCTIBLE

.00

.00

%
100

PAYABLE
3,400.00

REMARK CODE
16

3,400.00

PROVIDER PAYMENT:

3,331.58

INSURED PAYMENT:

.00

TRANSACTION FEE:

59.50

G.C.T.:

8.92

16-Claim paid according to Pre-Authorization given.
============================================================================================================================

TOTALS

PAYMENT(S):

25,011.57

TRANSACTION FEE:

446.68

G.C.T.:

66.99

============================================================================================================================
YOU MAY E-MAIL ANY QUERIES TO claim-inquiries@sagicor.com OR CALL OUR CALL CENTRE MONDAYS TO SATURDAYS 8:00a.m. TO 9:00p.m.
- 929-8920 (Select option 3).'''

# fuction_calling(test_text)
# print(fuction_calling(test_text))