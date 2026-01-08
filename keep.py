from google import genai
from google.genai import types
import pathlib
import os
import json
import pandas as pd
from dotenv import load_dotenv
from datetime import datetime

load_dotenv()

class InvoiceExtractor:
    def __init__(self):
        self.client = genai.Client(api_key=os.getenv('GOOGLE_API_KEY'))
        self.model = "gemini-2.5-flash"
        
    def extract_invoice_data(self, image_path):
        """Extract structured data from invoice image using Gemini API"""
        try:
            # Upload the file
            my_file = self.client.files.upload(file=image_path)
            
            # Simple extraction prompt
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
            
            # Generate content
            response = self.client.models.generate_content(
                model=self.model,
                contents=[my_file, prompt],
            )
            
            # Clean and parse response
            response_text = response.text.strip()
            response_text = response_text.replace('```json', '').replace('```', '').strip()
            
            # Parse JSON
            data = json.loads(response_text)
            return data
            
        except json.JSONDecodeError as e:
            print(f"  ✗ JSON parsing error: {e}")
            print(f"  Response: {response.text[:200]}")
            return None
        except Exception as e:
            print(f"  ✗ Error: {e}")
            return None
    
    def flatten_to_rows(self, invoice_data):
        """Convert invoice data to simple flat rows"""
        if not invoice_data:
            return []
        
        rows = []
        
        # Header info (will be repeated for each line item)
        invoice_number = invoice_data.get('invoice_number', '')
        waybill_number = invoice_data.get('waybill_number', '')
        customer_name = invoice_data.get('customer_name', '')
        order_number = invoice_data.get('order_number', '')
        invoice_date = invoice_data.get('invoice_date', '')
        
        # Process each line item
        for item in invoice_data.get('line_items', []):
            row = {
                # Header fields (repeated for each row)
                'Invoice_Number': invoice_number,
                'Waybill_Number': waybill_number,
                'Customer_Name': customer_name,
                'Order_Number': order_number,
                'Invoice_Date': invoice_date,
                
                # Line item details
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
    
    def process_multiple_invoices(self, image_paths, output_excel='extracted_invoices.xlsx'):
        """Process multiple invoices and save to Excel"""
        all_rows = []
        
        for idx, image_path in enumerate(image_paths, 1):
            print(f"\nProcessing {idx}/{len(image_paths)}: {image_path}")
            
            # Extract data
            invoice_data = self.extract_invoice_data(image_path)
            
            if invoice_data:
                # Flatten to rows
                rows = self.flatten_to_rows(invoice_data)
                all_rows.extend(rows)
                print(f"  ✓ Extracted {len(rows)} line items")
            else:
                print(f"  ✗ Failed to extract data")
        
        # Create DataFrame
        if all_rows:
            df = pd.DataFrame(all_rows)
            
            # Save to Excel with formatting
            with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Invoice_Data', index=False)
                
                # Auto-adjust column widths
                from openpyxl.utils import get_column_letter
                worksheet = writer.sheets['Invoice_Data']
                
                for idx, col in enumerate(df.columns, 1):
                    try:
                        column_letter = get_column_letter(idx)
                        max_length = max(
                            df[col].astype(str).apply(len).max(),
                            len(col)
                        )
                        adjusted_width = min(max_length + 2, 50)
                        worksheet.column_dimensions[column_letter].width = adjusted_width
                    except:
                        continue
            
            print(f"\n{'='*60}")
            print(f"✓ SUCCESS: Saved {len(all_rows)} rows to {output_excel}")
            print(f"{'='*60}")
            print(f"Documents processed: {len(image_paths)}")
            print(f"Total line items: {len(all_rows)}")
            print(f"Columns: {len(df.columns)}")
            print(f"\nColumn names:")
            for col in df.columns:
                print(f"  - {col}")
            
            return df
        else:
            print("\n✗ No data extracted")
            return None


# Main execution
if __name__ == "__main__":
    # Initialize extractor
    extractor = InvoiceExtractor()
    
    # List of invoice images to process
    invoice_images = [
        "Multiple Product Waybill-images-1.jpg",  # Commercial Invoice
        "Multiple Product Waybill-images-2.jpg",  # Sales Waybill
        # Add more image paths here
    ]
    
    # Process all invoices
    df = extractor.process_multiple_invoices(
        invoice_images,
        output_excel='aava_invoices_extracted.xlsx'
    )
    