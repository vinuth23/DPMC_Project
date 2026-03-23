import os
import json
import requests
from flask import Flask, render_template, request, jsonify, send_file
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from io import BytesIO
from dotenv import load_dotenv
from PIL import Image
import pytesseract

load_dotenv()

app = Flask(__name__)

# Free LLM options:
# 1. Ollama (local, completely free) - http://localhost:11434
# 2. Hugging Face (free tier available)
# Using Ollama for completely free, offline capability
#comment ersr

OLLAMA_API = "http://localhost:11434/api/generate"
OLLAMA_MODEL = "mistral"  # Free model, you can also use "llama2", "neural-chats"

# Tesseract OCR path (Windows default install location)
TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
if os.path.exists(TESSERACT_PATH):
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

def build_fields_list(fields_list):
    """Helper to format field list as string"""
    return "\n".join([f"- {f['name']}" for f in fields_list]) if fields_list else "None"

def generate_save_test_cases(starting_tc_number):
    """
    Generate 3 standard save/form test cases that should appear at the end of every test case set
    
    starting_tc_number: The TC number to start from (e.g., if 10, will generate TC10, TC11, TC12)
    """
    return [
        {
            'id': f"TC{str(starting_tc_number).zfill(2)}",
            'scenario': "Verify whether save is not allowed if required fields are empty",
            'expected_result': "The form should not be saved and an error message should be displayed indicating which required fields are empty"
        },
        {
            'id': f"TC{str(starting_tc_number + 1).zfill(2)}",
            'scenario': "Verify whether the form highlights the empty required fields when save is pressed",
            'expected_result': "All empty required fields should be highlighted/marked in red or with an error indicator"
        },
        {
            'id': f"TC{str(starting_tc_number + 2).zfill(2)}",
            'scenario': "Verify whether save works as expected when every required field is complete",
            'expected_result': "The form should be saved successfully and a success message should be displayed"
        }
    ]

def generate_test_cases_from_requirement(requirement_text, custom_fields=None):
    """
    Use free Ollama LLM to generate test cases from requirement text and/or custom fields
    Ollama runs locally on your computer - completely FREE and OFFLINE
    
    custom_fields: List of dicts with 'name' and 'type' keys (type can be 'text', 'amount', or 'button')
    If no requirement_text is provided, generates test cases based ONLY on custom fields
    
    Test cases are generated differently based on field type:
    - TEXT fields: Data loading + error handling (2 test cases per field)
    - AMOUNT fields: Valid input acceptance + invalid input rejection (2 test cases per field)
    - BUTTON fields: Button functionality (1 test case per button)
    """
    
    # Build the prompt based on what we have
    if custom_fields and any(field.get('name') for field in custom_fields):
        # Separate fields by type and data field flag
        text_fields = [f for f in custom_fields if f.get('name') and f.get('type') == 'text']
        amount_fields = [f for f in custom_fields if f.get('name') and f.get('type') == 'amount']
        button_fields = [f for f in custom_fields if f.get('name') and f.get('type') == 'button']
        
        # Separate data fields from regular fields (only for TEXT fields)
        data_text_fields = [f for f in text_fields if f.get('is_data_field', False)]
        regular_text_fields = [f for f in text_fields if not f.get('is_data_field', False)]
        # Amount fields never have data field flag - they always get 2 test cases
        regular_amount_fields = amount_fields
        
        # Format field lists
        regular_text_fields_list = build_fields_list(regular_text_fields)
        data_text_fields_list = build_fields_list(data_text_fields)
        regular_amount_fields_list = build_fields_list(regular_amount_fields)
        
        # Build context prefix
        context_prefix = f"REQUIREMENT: {requirement_text}\n\n" if requirement_text.strip() else ""
        context_text = "for this requirement" if requirement_text.strip() else "for these fields in a form/module"
        
        # Calculate exact numbers
        regular_text_count = len(regular_text_fields)
        data_text_count = len(data_text_fields)
        amount_count = len(regular_amount_fields)
        
        total_expected = regular_text_count + (data_text_count * 3) + (amount_count * 2)
        
        prompt = f"""You are a QA engineer creating test cases {context_text}.

{context_prefix}*** CRITICAL: ONLY use the fields provided below. Do NOT make up, assume, or hallucinate any fields. ***

FIELD INVENTORY (COMPLETE LIST - ONLY GENERATE FOR THESE):
- REGULAR TEXT FIELDS: {regular_text_count} field(s) → Generate {regular_text_count} test case(s) total
  {regular_text_fields_list if regular_text_fields_list else '(None)'}
- DATA TEXT FIELDS: {data_text_count} field(s) → Generate {data_text_count * 3} test case(s) total
  {data_text_fields_list if data_text_fields_list else '(None)'}
- AMOUNT FIELDS: {amount_count} field(s) → Generate {amount_count * 2} test case(s) total
  {regular_amount_fields_list if regular_amount_fields_list else '(None)'}

TOTAL EXPECTED TEST CASES: {total_expected} (MUST be exactly this number)

=== GENERATION RULES ===

FOR EACH REGULAR TEXT FIELD (if any exist):
- Generate EXACTLY 1 test case
- Use template: "Verify whether the [field name] has an input limit and whether the entered amount can be successfully saved"
- NEVER generate more than 1 test case per regular text field
- NEVER add "loads up values" test case

FOR EACH DATA TEXT FIELD (if any exist):
- Generate EXACTLY 3 test cases per field
- Test 1: "Verify whether the field only accepts the value when a proper [field name] is entered"
- Test 2: "Verify whether proper errors are shown if an invalid [field name] is entered"  
- Test 3: "Verify whether the [field name] properly loads up the values in the fields below"
- ALWAYS include all 3 test cases
- NEVER skip the "loads up values" test case

FOR EACH AMOUNT FIELD (if any exist):
- Generate EXACTLY 2 test cases per field
- Test 1: "Verify whether the field only accepts the value when a proper [field name] is entered"
- Test 2: "Verify whether the form shows proper error messages to inputs like negative [field name] and zero"
- NEVER generate more than 2 test cases per amount field
- NEVER add "loads up values" test case for amount fields

IF NO FIELDS ARE PROVIDED:
- Return empty test_cases array: {{"test_cases": []}}

=== OUTPUT FORMAT ===
Return ONLY this JSON format - no markdown, no extra text, ONLY JSON:

{{"test_cases": [
  {{"id": "TC01", "scenario": "...", "expected_result": "..."}},
  {{"id": "TC02", "scenario": "...", "expected_result": "..."}}
]}}

=== VERIFICATION CHECKLIST ===
Before returning, verify:
✓ Total test cases = {total_expected}
✓ All test cases have exactly 3 fields: id, scenario, expected_result
✓ No "steps" field anywhere
✓ Test case IDs are sequential (TC01, TC02, etc.)
✓ Scenarios are SPECIFIC, DETAILED, and COMPLETE
✓ Expected results are CLEAR and MEASURABLE
✓ Only used the {regular_text_count + data_text_count + amount_count} fields provided
✓ Did NOT invent, assume, or hallucinate additional fields"""
    else:
        # Only requirement provided, no fields
        if not requirement_text.strip():
            # No requirement, but might have buttons
            if button_fields and any(button.get('name') for button in button_fields):
                button_test_cases = []
                
                # Add spelling/grammar test case FIRST
                button_test_cases.append({
                    'id': 'TC01',
                    'scenario': 'Verify whether spellings and grammar of the form are correct',
                    'expected_result': 'All text on the form should be correctly spelled with proper grammar'
                })
                
                for i, button in enumerate(button_fields):
                    if button.get('name') and button.get('name').strip():
                        button_name = button['name'].strip()
                        test_case_id = f"TC{str(len(button_test_cases) + 1).zfill(2)}"
                        button_test_cases.append({
                            'id': test_case_id,
                            'scenario': f"Verify whether the {button_name} button works as expected",
                            'expected_result': f"The {button_name} button should perform its intended action without errors"
                        })
                
                # Add standard save/form test cases at the end
                next_tc_number = len(button_test_cases) + 1
                save_test_cases = generate_save_test_cases(next_tc_number)
                button_test_cases.extend(save_test_cases)
                
                return button_test_cases
            else:
                # Nothing provided at all
                return []
        
        prompt = f"""You are a QA engineer. Generate 8-10 comprehensive test cases from this requirement:

{requirement_text}

Cover:
- Positive scenarios (things that should work)
- Negative scenarios (error conditions)
- Edge cases (boundary conditions)

RESPONSE FORMAT (CRITICAL):
Return ONLY valid JSON in this exact format. Do NOT include markdown, code blocks, or any text:

{{"test_cases": [
  {{"id": "TC01", "scenario": "test case scenario here", "expected_result": "expected result here"}},
  {{"id": "TC02", "scenario": "test case scenario here", "expected_result": "expected result here"}}
]}}

MANDATORY REQUIREMENTS:
1. EVERY test case must have EXACTLY 3 fields: id, scenario, and expected_result
2. NO "steps" field - steps field is completely removed from test cases
3. Scenarios must be SPECIFIC and DETAILED
4. Expected results must be CLEAR and MEASURABLE
5. Test case IDs must be sequential (TC01, TC02, TC03, etc.)
6. Generate 8-10 test cases total
7. No markdown, no code blocks, only JSON"""

    try:
        # Call local Ollama API (free, runs on your computer)
        response = requests.post(
            OLLAMA_API,
            json={
                "model": OLLAMA_MODEL,
                "prompt": prompt,
                "stream": False,
                "temperature": 0.7
            },
            timeout=180
        )
        
        if response.status_code != 200:
            print(f"Ollama error: Status {response.status_code}")
            print(f"Response: {response.text[:500]}")
            return []
        
        response_data = response.json()
        response_text = response_data.get("response", "").strip()
        
        print(f"Ollama response length: {len(response_text)}")
        
        # Extract JSON from response
        if '```json' in response_text:
            response_text = response_text.split('```json')[1].split('```')[0].strip()
        elif '```' in response_text:
            response_text = response_text.split('```')[1].split('```')[0].strip()
        
        # Find JSON object in response
        start = response_text.find('{')
        end = response_text.rfind('}') + 1
        if start != -1 and end > start:
            response_text = response_text[start:end]
        
        print(f"Extracted JSON length: {len(response_text)}")
        print(f"JSON content preview: {response_text[:100]}")
        
        try:
            test_cases_data = json.loads(response_text)
            test_cases = test_cases_data.get('test_cases', [])
            print(f"Successfully parsed {len(test_cases)} test cases from Ollama")
            
            # FILTER: Remove any test cases that reference fields NOT in our provided list
            # This prevents Ollama from hallucinating fields that don't exist
            all_provided_fields = [f.get('name', '').lower() for f in custom_fields if f.get('name')]
            filtered_test_cases = []
            
            for tc in test_cases:
                scenario = tc.get('scenario', '').lower()
                # Check if this test case references any of our provided fields
                references_valid_field = any(field_name in scenario for field_name in all_provided_fields)
                
                if references_valid_field:
                    filtered_test_cases.append(tc)
                else:
                    print(f"FILTERED OUT (hallucinated field): {tc.get('scenario', '')[:60]}")
            
            test_cases = filtered_test_cases
            print(f"After filtering hallucinated fields: {len(test_cases)} test cases")
            
            # INSERT SPELLING/GRAMMAR TEST CASE AS FIRST TEST CASE
            spelling_test_case = {
                'id': 'TC01',
                'scenario': 'Verify whether spellings and grammar of the form are correct',
                'expected_result': 'All text on the form should be correctly spelled with proper grammar'
            }
            test_cases.insert(0, spelling_test_case)
            
            # RE-NUMBER ALL TEST CASES after insertion (shift all IDs by 1)
            for i, tc in enumerate(test_cases):
                tc['id'] = f"TC{str(i + 1).zfill(2)}"
            
            print(f"Added spelling/grammar test case at beginning")
            
            # Generate button test cases
            if button_fields and any(button.get('name') for button in button_fields):
                button_test_cases = []
                for i, button in enumerate(button_fields):
                    if button.get('name') and button.get('name').strip():
                        button_name = button['name'].strip()
                        test_case_id = f"TC{str(len(test_cases) + i + 1).zfill(2)}"
                        button_test_cases.append({
                            'id': test_case_id,
                            'scenario': f"Verify whether the {button_name} button works as expected",
                            'expected_result': f"The {button_name} button should perform its intended action without errors"
                        })
                test_cases.extend(button_test_cases)
                print(f"Added {len(button_test_cases)} button test cases")
            
            # Add standard save/form test cases at the end
            next_tc_number = len(test_cases) + 1
            save_test_cases = generate_save_test_cases(next_tc_number)
            test_cases.extend(save_test_cases)
            print(f"Added {len(save_test_cases)} standard save test cases")
            
            return test_cases
        except json.JSONDecodeError as json_err:
            print(f"JSON parsing failed: {json_err}")
            print(f"Full response text: {response_text}")
            return []
    
    except requests.exceptions.ConnectionError as e:
        print(f"Error: Cannot connect to Ollama on {OLLAMA_API}")
        print(f"Details: {e}")
        print("Make sure Ollama is running locally!")
        return []
    except Exception as e:
        print(f"Error: {type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()
        return []

def create_excel_file(test_cases):
    """
    Create an Excel file from test cases with new format:
    - Scenario ID (empty - manual)
    - Scenario (empty - manual)
    - Test Case ID
    - Test Case (renamed from Scenario)
    - Expected Result
    - Actual Result (empty - manual)
    - Comments (empty - manual)
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Test Cases"
    
    # Define styles
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # New column headers as per your requirement
    headers = ["Scenario ID", "Scenario", "Test Case ID", "Test Case", "Expected Result", "Actual Result", "Comments"]
    ws.append(headers)
    
    # Format header row
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
        cell.border = border
    
    # Add test case data
    for test_case in test_cases:
        ws.append([
            "",  # Scenario ID (empty - to be filled manually)
            "",  # Scenario (empty - to be filled manually)
            test_case.get('id', ''),  # Test Case ID
            test_case.get('scenario', ''),  # Test Case (renamed from Scenario)
            test_case.get('expected_result', ''),  # Expected Result
            "",  # Actual Result (empty - to be filled manually)
            ""   # Comments (empty - to be filled manually)
        ])
    
    # Format data rows
    for row in ws.iter_rows(min_row=2, max_row=len(test_cases) + 1):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    
    # Set column widths
    ws.column_dimensions['A'].width = 15  # Scenario ID
    ws.column_dimensions['B'].width = 25  # Scenario
    ws.column_dimensions['C'].width = 15  # Test Case ID
    ws.column_dimensions['D'].width = 30  # Test Case
    ws.column_dimensions['E'].width = 35  # Expected Result
    ws.column_dimensions['F'].width = 30  # Actual Result
    ws.column_dimensions['G'].width = 25  # Comments
    
    # Save to bytes
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output

@app.route('/')
def index():
    """Render the main page"""
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    """
    API endpoint to generate test cases
    Uses free local Ollama LLM
    
    Expected JSON:
    {
        'requirement': 'requirement text (optional)',
        'custom_fields': [
            {'name': 'Customer Name', 'type': 'text'},
            {'name': 'Loan Amount', 'type': 'amount'},
            {'name': 'Submit', 'type': 'button'}
        ]
    }
    
    Works in multiple modes:
    1. With custom fields only - generates field-specific test cases
    2. With requirement only - generates requirement-based test cases
    3. With buttons (type='button') - generates one test case per button
    4. With any combination - generates all relevant test cases
    """
    data = request.json
    requirement = data.get('requirement', '').strip()
    custom_fields = data.get('custom_fields', [])
    
    # Check if we have at least something to work with
    has_fields = any(field.get('name') for field in custom_fields)
    
    if not has_fields and not requirement:
        return jsonify({'error': 'Please either enter a requirement or add at least one custom field/button'}), 400
    
    # Generate test cases using Ollama (free, local)
    test_cases = generate_test_cases_from_requirement(requirement, custom_fields)
    
    if not test_cases:
        return jsonify({'error': 'Failed to generate test cases. Make sure Ollama is running!'}), 500
    
    return jsonify({
        'success': True,
        'test_cases': test_cases
    })

@app.route('/export', methods=['POST'])
def export():
    """
    API endpoint to export test cases to Excel
    """
    data = request.json
    test_cases = data.get('test_cases', [])
    
    if not test_cases:
        return jsonify({'error': 'No test cases to export'}), 400
    
    # Create Excel file
    excel_file = create_excel_file(test_cases)
    
    return send_file(
        excel_file,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='generated_test_cases.xlsx'
    )

@app.route('/extract-fields', methods=['POST'])
def extract_fields():
    """Extract form fields from screenshot(s) using Tesseract OCR + Mistral"""
    if 'screenshots' not in request.files:
        return jsonify({'error': 'No screenshots uploaded'}), 400

    files = request.files.getlist('screenshots')
    if not files or all(f.filename == '' for f in files):
        return jsonify({'error': 'No screenshots selected'}), 400

    # Step 1: OCR all screenshots and combine text
    all_ocr_text = []
    for file in files:
        if not file or not file.filename:
            continue
        try:
            img = Image.open(file)
            # Convert to grayscale for better OCR
            img = img.convert('L')
            ocr_text = pytesseract.image_to_string(img)
            print(f"OCR from {file.filename}: {ocr_text[:300]}")
            all_ocr_text.append(ocr_text)
        except Exception as e:
            print(f"OCR error on {file.filename}: {e}")
            continue

    combined_text = "\n".join(all_ocr_text).strip()
    if not combined_text:
        return jsonify({'error': 'Could not extract any text from screenshots. Make sure Tesseract is installed.'}), 500

    # Step 2: Send OCR text to Mistral to identify form fields
    prompt = f"""Below is text extracted via OCR from a screenshot of a form/screen in a software application.
Identify all form fields, input boxes, dropdowns, and buttons from this text.

For each element, output one line in this exact format:
- FieldLabel: type

Where type is one of: text, amount, button

Use "text" for text inputs, dropdowns, date fields, search fields, code fields, name fields, IDs, descriptions.
Use "amount" for numeric/currency/money/quantity fields.
Use "button" for buttons (like Add, Delete, Submit, Save, Clear, Search).

OCR TEXT:
{combined_text}

List ONLY the field labels and types, one per line. Do not explain anything."""

    try:
        response = requests.post(
            OLLAMA_API,
            json={
                "model": OLLAMA_MODEL,
                "prompt": prompt,
                "stream": False,
                "temperature": 0.3
            },
            timeout=180
        )

        if response.status_code != 200:
            print(f"Mistral error: Status {response.status_code}")
            return jsonify({'error': 'Failed to process OCR text with Mistral'}), 500

        response_data = response.json()
        response_text = response_data.get("response", "").strip()
        print(f"Mistral field extraction response: {response_text[:500]}")

        # Parse lines like "- Label: type" or "Label: type"
        all_fields = []
        for line in response_text.split('\n'):
            line = line.strip().lstrip('-•*0123456789.').strip()
            if ':' not in line:
                continue
            parts = line.rsplit(':', 1)
            if len(parts) != 2:
                continue
            name = parts[0].strip().strip('"').strip("'")
            raw_type = parts[1].strip().lower().strip('"').strip("'").rstrip('.')
            if not name or len(name) > 80:
                continue

            if any(kw in raw_type for kw in ['button', 'btn']):
                field_type = 'button'
            elif any(kw in raw_type for kw in ['amount', 'number', 'numeric', 'currency', 'quantity']):
                field_type = 'amount'
            else:
                field_type = 'text'

            all_fields.append({'name': name, 'type': field_type})

        print(f"Extracted {len(all_fields)} fields total")

        if not all_fields:
            return jsonify({'error': 'Could not identify form fields from the screenshot text.'}), 500

        # Deduplicate fields by name (case-insensitive)
        seen = set()
        unique_fields = []
        for f in all_fields:
            key = f.get('name', '').strip().lower()
            if key and key not in seen:
                seen.add(key)
                unique_fields.append(f)

        return jsonify({
            'success': True,
            'fields': unique_fields
        })

    except requests.exceptions.ConnectionError:
        return jsonify({'error': 'Cannot connect to Ollama. Make sure it is running!'}), 500
    except Exception as e:
        print(f"Error extracting fields: {e}")
        return jsonify({'error': f'Error: {str(e)}'}), 500


@app.route('/health', methods=['GET'])
def health():
    """Health check endpoint"""
    try:
        # Check if Ollama is running
        response = requests.get('http://localhost:11434/api/tags', timeout=2)
        ollama_running = response.status_code == 200
    except:
        ollama_running = False
    
    return jsonify({
        'status': 'ok',
        'ollama_running': ollama_running,
        'message': 'Using FREE local Ollama LLM (no API costs!)'
    })

if __name__ == '__main__':
    app.run(debug=False, port=5000)
