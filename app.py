# app.py - Flask Web Application for Invoice Extraction
from flask import Flask, render_template, request, jsonify, send_file
from google import genai
import os
import json
import pandas as pd
from dotenv import load_dotenv
from werkzeug.utils import secure_filename
import uuid
from datetime import datetime

load_dotenv()

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['RESULTS_FOLDER'] = 'results'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['ALLOWED_EXTENSIONS'] = {'png', 'jpg', 'jpeg', 'pdf'}

# Create folders if they don't exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['RESULTS_FOLDER'], exist_ok=True)

# In-memory storage for extracted data (use database in production)
extracted_data = []

class InvoiceExtractor:
    def __init__(self):
        self.client = genai.Client(api_key=os.getenv('GOOGLE_API_KEY'))
        self.model = "gemini-2.5-flash"
        
    def extract_invoice_data(self, image_path):
        """Extract structured data from invoice image"""
        try:
            my_file = self.client.files.upload(file=image_path)
            
            prompt = """
            Extract data from this invoice/waybill and return ONLY valid JSON (no markdown, no backticks).
            
            Structure:
            {
                "invoice_number": "",
                "waybill_number": "",
                "customer_name": "",
                "order_number": "",
                "invoice_date": "",
                "line_items": [
                    {
                        "line_no": 1,
                        "item_code": "",
                        "item_description": "",
                        "quantity": 0,
                        "uom": "",
                        "unit_price": 0,
                        "total_amount": 0,
                        "discount_amount": 0,
                        "vat": 0,
                        "amount_incl_vat": 0,
                        "batch_no": "",
                        "expiry_date": ""
                    }
                ]
            }
            
            Rules:
            - Extract ALL line items from the table
            - For missing fields, use empty string "" or 0
            - Return ONLY the JSON, no other text
            """
            
            response = self.client.models.generate_content(
                model=self.model,
                contents=[my_file, prompt],
            )
            
            response_text = response.text.strip().replace('```json', '').replace('```', '').strip()
            data = json.loads(response_text)
            return data
            
        except Exception as e:
            print(f"Error: {e}")
            return None

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def flatten_to_rows(invoice_data):
    """Convert invoice data to flat rows"""
    if not invoice_data:
        return []
    
    rows = []
    invoice_number = invoice_data.get('invoice_number', '')
    waybill_number = invoice_data.get('waybill_number', '')
    customer_name = invoice_data.get('customer_name', '')
    order_number = invoice_data.get('order_number', '')
    invoice_date = invoice_data.get('invoice_date', '')
    
    for item in invoice_data.get('line_items', []):
        row = {
            'Invoice_Number': invoice_number,
            'Waybill_Number': waybill_number,
            'Customer_Name': customer_name,
            'Order_Number': order_number,
            'Invoice_Date': invoice_date,
            'Line_No': item.get('line_no', ''),
            'Item_Code': item.get('item_code', ''),
            'Item_Description': item.get('item_description', ''),
            'Quantity': item.get('quantity', 0),
            'UOM': item.get('uom', ''),
            'Unit_Price': item.get('unit_price', 0),
            'Total_Amount': item.get('total_amount', 0),
            'Discount_Amount': item.get('discount_amount', 0),
            'VAT': item.get('vat', 0),
            'Amount_Incl_VAT': item.get('amount_incl_vat', 0),
            'Batch_No': item.get('batch_no', ''),
            'Expiry_Date': item.get('expiry_date', ''),
        }
        rows.append(row)
    
    return rows

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    if 'files[]' not in request.files:
        return jsonify({'error': 'No files provided'}), 400
    
    files = request.files.getlist('files[]')
    uploaded_files = []
    
    for file in files:
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            unique_filename = f"{uuid.uuid4()}_{filename}"
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
            file.save(filepath)
            
            uploaded_files.append({
                'original_name': filename,
                'stored_name': unique_filename,
                'path': filepath,
                'size': os.path.getsize(filepath)
            })
    
    return jsonify({
        'success': True,
        'files': uploaded_files,
        'count': len(uploaded_files)
    })

@app.route('/process', methods=['POST'])
def process_files():
    global extracted_data
    
    data = request.json
    file_paths = data.get('file_paths', [])
    
    if not file_paths:
        return jsonify({'error': 'No files to process'}), 400
    
    extractor = InvoiceExtractor()
    all_rows = []
    processed = 0
    
    for filepath in file_paths:
        try:
            invoice_data = extractor.extract_invoice_data(filepath)
            if invoice_data:
                rows = flatten_to_rows(invoice_data)
                all_rows.extend(rows)
                processed += 1
        except Exception as e:
            print(f"Error processing {filepath}: {e}")
    
    # Store in memory
    extracted_data = all_rows
    
    # Calculate statistics
    df = pd.DataFrame(all_rows) if all_rows else pd.DataFrame()
    stats = {
        'total_documents': df['Invoice_Number'].nunique() if not df.empty else 0,
        'total_line_items': len(all_rows),
        'total_quantity': int(df['Quantity'].sum()) if not df.empty else 0,
        'total_value': float(df['Amount_Incl_VAT'].sum()) if not df.empty else 0
    }
    
    return jsonify({
        'success': True,
        'processed': processed,
        'total_rows': len(all_rows),
        'stats': stats
    })

@app.route('/data')
def get_data():
    return jsonify({
        'data': extracted_data,
        'count': len(extracted_data)
    })

@app.route('/export')
def export_excel():
    if not extracted_data:
        return jsonify({'error': 'No data to export'}), 400
    
    df = pd.DataFrame(extracted_data)
    
    # Generate unique filename
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f'invoice_extract_{timestamp}.xlsx'
    filepath = os.path.join(app.config['RESULTS_FOLDER'], filename)
    
    # Save to Excel
    df.to_excel(filepath, index=False, sheet_name='Invoice_Data')
    
    return send_file(filepath, as_attachment=True, download_name=filename)

@app.route('/clear')
def clear_data():
    global extracted_data
    extracted_data = []
    
    # Optionally clear upload folder
    for file in os.listdir(app.config['UPLOAD_FOLDER']):
        try:
            os.remove(os.path.join(app.config['UPLOAD_FOLDER'], file))
        except:
            pass
    
    return jsonify({'success': True})

if __name__ == '__main__':
    app.run(debug=True, port=5000)