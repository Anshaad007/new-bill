from flask import Flask, jsonify, send_file, send_from_directory, request
import csv
import os
import shutil
import json
import re
from io import StringIO, BytesIO
from reportlab.lib.pagesizes import letter, legal, A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Register fonts that support ₹ rupee symbol (Windows: Arial, Linux: DejaVu Sans)
FONT = 'Helvetica'
FONT_BOLD = 'Helvetica-Bold'
FONT_ITALIC = 'Helvetica-Oblique'

font_paths = [
    # Windows
    {'regular': 'C:/Windows/Fonts/arial.ttf', 'bold': 'C:/Windows/Fonts/arialbd.ttf', 'italic': 'C:/Windows/Fonts/ariali.ttf'},
    # Linux (Ubuntu/Render)
    {'regular': '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf', 'bold': '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf', 'italic': '/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf'},
]
for fp in font_paths:
    try:
        if os.path.exists(fp['regular']):
            pdfmetrics.registerFont(TTFont('CustomFont', fp['regular']))
            pdfmetrics.registerFont(TTFont('CustomFont-Bold', fp['bold']))
            pdfmetrics.registerFont(TTFont('CustomFont-Italic', fp['italic']))
            FONT = 'CustomFont'
            FONT_BOLD = 'CustomFont-Bold'
            FONT_ITALIC = 'CustomFont-Italic'
            break
    except Exception as e:
        print(f"Font registration failed: {e}")
from datetime import datetime
import zipfile

app = Flask(__name__, static_folder=os.path.dirname(os.path.abspath(__file__)))

@app.after_request
def add_cors_headers(response):
    response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type,Authorization'
    response.headers['Access-Control-Allow-Methods'] = 'GET,PUT,POST,DELETE,OPTIONS'
    return response

@app.errorhandler(Exception)
def handle_exception(e):
    return jsonify({
        'success': False,
        'error': f"Internal Server Error: {str(e)}"
    }), 500

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
RECEIPTS_FOLDER = os.path.join(BASE_DIR, 'receipts')
INVOICE_COUNTER_FILE = os.path.join(BASE_DIR, 'invoice_counter.txt')

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RECEIPTS_FOLDER, exist_ok=True)

def get_next_invoice_number():
    """Get the next invoice number in ascending order"""
    if os.path.exists(INVOICE_COUNTER_FILE):
        try:
            # Try UTF-8 first, then fall back to UTF-16 (handles Notepad-saved files)
            try:
                with open(INVOICE_COUNTER_FILE, 'r', encoding='utf-8') as f:
                    content = f.read().strip()
            except (UnicodeDecodeError, ValueError):
                with open(INVOICE_COUNTER_FILE, 'r', encoding='utf-16') as f:
                    content = f.read().strip()
            # Remove BOM and null bytes if any
            content = content.replace('\ufeff', '').replace('\x00', '').strip()
            current_num = int(content)
        except (ValueError, UnicodeDecodeError) as e:
            print(f"WARNING: Could not read invoice counter file, resetting to 191. Error: {e}")
            current_num = 191
    else:
        # Start from 192 (B2C 0192)
        current_num = 191
    
    next_num = current_num + 1
    
    # Save the new number (always as UTF-8)
    with open(INVOICE_COUNTER_FILE, 'w', encoding='utf-8') as f:
        f.write(str(next_num))
    
    # Format as B2C XXXX (4 digits with leading zeros)
    return f"B2C {next_num:04d}"

ALLOWED_EXTENSIONS = {'csv', 'xls', 'xlsx'}

# Store current session data in memory
session_data = {
    'headers': [],
    'rows': [],
    'filename': '',
    'column_map': {}  # Maps generic keys to actual column names
}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def read_csv_file(filepath):
    rows = []
    headers = []
    with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
        reader = csv.DictReader(f)
        headers = reader.fieldnames or []
        for row in reader:
            rows.append(row)
    return headers, rows

def read_excel_file(filepath):
    try:
        import openpyxl
    except ImportError:
        raise ImportError("openpyxl is required for Excel files. Run: pip install openpyxl")
    
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    data = list(ws.iter_rows(values_only=True))
    if not data:
        return [], []
    
    headers = [str(h) if h else f'Column_{i}' for i, h in enumerate(data[0])]
    rows = []
    for row in data[1:]:
        row_dict = {}
        for i, val in enumerate(row):
            if i < len(headers):
                row_dict[headers[i]] = str(val) if val is not None else ''
        rows.append(row_dict)
    return headers, rows

def auto_detect_columns(headers):
    """Auto-detect which columns map to name, amount, date, etc."""
    mapping = {}
    name_keywords = ['name', 'student', 'customer', 'client', 'party', 'person', 'recipient']
    amount_keywords = ['amount', 'total', 'fee', 'price', 'cost', 'charge', 'payment']
    date_keywords = ['date', 'payment date', 'paid', 'timestamp', 'time']
    id_keywords = ['transaction', 'receipt', 'invoice', 'id', 'ref', 'reference', 'txn']
    class_keywords = ['class', 'category', 'group', 'section', 'department', 'type', 'batch']
    desc_keywords = ['description', 'detail', 'particular', 'item', 'fee type', 'specify', 'term', 'month']

    lower_headers = [h.lower().strip() for h in headers]

    for key, keywords in [
        ('name', name_keywords), ('amount', amount_keywords),
        ('date', date_keywords), ('transaction_id', id_keywords),
        ('class', class_keywords), ('description', desc_keywords)
    ]:
        for kw in keywords:
            for i, lh in enumerate(lower_headers):
                if kw in lh and key not in mapping:
                    mapping[key] = headers[i]
                    break
            if key in mapping:
                break

    return mapping

def fix_scientific_notation(value):
    """Convert scientific notation strings like '5.24591E+11' to full number strings"""
    if not value or not isinstance(value, str):
        return str(value) if value else 'N/A'
    value = value.strip()
    try:
        if 'E+' in value or 'e+' in value or 'E-' in value or 'e-' in value:
            num = float(value)
            if num == int(num):
                return str(int(num))
            return str(num)
    except (ValueError, OverflowError):
        pass
    return value

def generate_receipt_pdf(row_data, filename, col_map, org_settings, invoice_number=None, fee_items=None, transaction_id=None):
    """Generate a PDF receipt for a single row using dynamic column mapping."""
    pdf_path = os.path.join(RECEIPTS_FOLDER, filename)
    
    # Use A4 page size for better proportions
    page_size = A4
    doc = SimpleDocTemplate(pdf_path, pagesize=page_size,
                            rightMargin=0.5*inch, leftMargin=0.5*inch,
                            topMargin=0.4*inch, bottomMargin=0.4*inch,
                            allowSplitting=1)
    
    page_width, page_height = page_size
    # Available width = page - margins - frame padding (6pt each side)
    avail_width = page_width - 1.0*inch - 12
    bg_color = colors.HexColor('#FFF2D8')
    accent_color = colors.HexColor('#BCA37F')
    primary_color = colors.HexColor('#113946')
    text_color = colors.HexColor('#113946')
    secondary_color = colors.HexColor('#666666')
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontSize=18, leading=22,
        textColor=primary_color, alignment=TA_LEFT, fontName=FONT_BOLD, spaceAfter=4)
    normal_style = ParagraphStyle('Normal', parent=styles['Normal'], fontSize=8, leading=12,
        textColor=text_color, alignment=TA_LEFT, fontName=FONT)
    label_style = ParagraphStyle('Label', parent=styles['Normal'], fontSize=9,
        textColor=primary_color, alignment=TA_LEFT, fontName=FONT_BOLD, leading=12)
    accent_label_style = ParagraphStyle('AccentLabel', parent=styles['Normal'], fontSize=10,
        textColor=text_color, alignment=TA_LEFT, fontName=FONT_BOLD, leading=13)
    total_style = ParagraphStyle('Total', parent=styles['Normal'], fontSize=11,
        textColor=colors.HexColor('#FFF2D8'), alignment=TA_CENTER, fontName=FONT_BOLD, leading=14)
    payment_title_style = ParagraphStyle('PaymentTitle', parent=styles['Heading3'], fontSize=10,
        textColor=primary_color, alignment=TA_LEFT, fontName=FONT_BOLD, leading=13, spaceAfter=4)
    footer_style = ParagraphStyle('Footer', parent=styles['Normal'], fontSize=7,
        textColor=secondary_color, alignment=TA_CENTER, fontName=FONT_ITALIC, leading=10)

    # Extract values using column mapping
    name_col = col_map.get('name', '')
    student_name = (row_data.get(name_col, 'N/A') if name_col else 'N/A').upper()
    
    amount_col = col_map.get('amount', '')
    amount_raw = row_data.get(amount_col, '0') if amount_col else '0'
    try:
        # Handle UTF-16 encoded strings
        if isinstance(amount_raw, str):
            amount_cleaned = str(amount_raw).replace(',', '').strip()
            # Remove UTF-16 BOM if present
            if amount_cleaned.startswith('\ufeff') or amount_cleaned.startswith('\ufffe'):
                amount_cleaned = amount_cleaned[1:]
            amount_value = float(amount_cleaned)
        else:
            amount_value = float(amount_raw)
    except:
        amount_value = 0.0

    date_col = col_map.get('date', '')
    date = row_data.get(date_col, 'N/A') if date_col else 'N/A'
    
    txn_col = col_map.get('transaction_id', '')
    original_transaction_id = row_data.get(txn_col, 'N/A') if txn_col else 'N/A'
    original_transaction_id = fix_scientific_notation(original_transaction_id)
    
    # Use the provided transaction ID if available, otherwise use original
    transaction_id = fix_scientific_notation(transaction_id) if transaction_id else original_transaction_id
    
    # Use the provided invoice number or generate one
    if invoice_number:
        receipt_number = invoice_number
    else:
        receipt_number = transaction_id
    
    class_col = col_map.get('class', '')
    student_class = (row_data.get(class_col, 'N/A') if class_col else 'N/A').upper()
    
    desc_col = col_map.get('description', '')
    description = (row_data.get(desc_col, '') if desc_col else '').strip() or 'Fee Payment'

    org_name = org_settings.get('org_name', 'ALIF ONLINE MORAL SCHOOL')
    org_address = org_settings.get('org_address', 'Othukkungal (PO), Malappuram, Kerala 676531')
    org_phone = org_settings.get('org_phone', '9061711444')
    org_email = org_settings.get('org_email', 'info1alifonlinemoralschool@gmail.com')
    org_gstin = org_settings.get('org_gstin', '32ACAFA0267H1ZY')
    org_website = org_settings.get('org_website', 'www.alifonlinemoralschool.com')

    # Use provided fee items or default to single item
    if fee_items and len(fee_items) > 0:
        items_to_use = fee_items
    else:
        try:
            tax_percent_val = float(str(org_settings.get('tax_percent', '18')).strip())
        except:
            tax_percent_val = 18.0
        base_price = amount_value / (1 + tax_percent_val / 100)
        items_to_use = [{'description': description, 'qty': 1, 'price': base_price, 'tax': tax_percent_val}]

    elements = []

    # Header with optional Logo on the right side
    logo_path = os.path.join(BASE_DIR, 'logo.png')
    logo_img = None
    if os.path.exists(logo_path):
        try:
            # ReportLab image helper
            logo_img = Image(logo_path, width=1.1*inch, height=1.1*inch)
        except Exception as img_err:
            print(f"Error loading logo in ReportLab: {img_err}")

    details_paragraphs = [
        Paragraph(org_name, title_style),
        Paragraph(org_address, normal_style),
        Paragraph(f"Phone: {org_phone}", normal_style),
        Paragraph(f"Email: {org_email}", normal_style),
        Paragraph(f"Website: {org_website}", normal_style)
    ]
    
    # Left column: details. Right column: logo
    left_cell = []
    for p in details_paragraphs:
        left_cell.append(p)
        left_cell.append(Spacer(1, 0.02*inch))
        
    header_table_data = [[left_cell, logo_img if logo_img else '']]
    header_table = Table(header_table_data, colWidths=[avail_width - 1.7*inch, 1.7*inch], hAlign='LEFT')
    header_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    elements.append(header_table)
    elements.append(Spacer(1, 0.1*inch))

    # GSTIN + Student Name Row
    gstin_cols = [avail_width * 0.55, avail_width * 0.45]
    gstin_row_data = [
        [Paragraph(f'GSTIN: {org_gstin}', label_style),
         Paragraph(f'<para alignment="right"><b>{student_name}</b></para>',
            ParagraphStyle('SN', parent=styles['Normal'], fontSize=13, textColor=primary_color,
                fontName=FONT_BOLD, leading=15))]
    ]
    gstin_table = Table(gstin_row_data, colWidths=gstin_cols, hAlign='LEFT')
    gstin_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), bg_color),
        ('LEFTPADDING', (0, 0), (-1, -1), 6), ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6), ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LINEBELOW', (0, 0), (-1, -1), 1, accent_color),
    ]))
    elements.append(gstin_table)
    elements.append(Spacer(1, 0.1*inch))

    # Bill to & Metadata info
    bill_to = Table([
        [Paragraph('<b>Bill To</b>', label_style)],
        [Paragraph(student_name, accent_label_style)],
        [Paragraph(student_class, normal_style)]
    ], colWidths=[avail_width * 0.5])
    bill_to.setStyle(TableStyle([('LEFTPADDING', (0, 0), (-1, -1), 0), ('RIGHTPADDING', (0, 0), (-1, -1), 0)]))

    invoice_meta = Table([
        [Paragraph('<b>Receipt No:</b>', label_style), Paragraph(str(receipt_number), normal_style)],
        [Paragraph('<b>Date:</b>', label_style), Paragraph(str(date), normal_style)],
    ], colWidths=[1.1*inch, avail_width * 0.45 - 1.1*inch])
    invoice_meta.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0), ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))

    details_table = Table([[bill_to, invoice_meta]], colWidths=[avail_width * 0.55, avail_width * 0.45], hAlign='LEFT')
    details_table.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP')]))
    elements.append(details_table)
    elements.append(Spacer(1, 0.1*inch))

    try:
        tax_percent_val = float(str(org_settings.get('tax_percent', '18')).strip())
    except:
        tax_percent_val = 18.0

    # Calculate total amount from fee items
    total_amount = 0
    for item in items_to_use:
        item_total = item.get('qty', 1) * item.get('price', 0) * (1 + item.get('tax', 0) / 100)
        total_amount += item_total

    # Invoice items table matching custom colors
    desc_style = ParagraphStyle('Desc', parent=styles['Normal'], fontSize=9, leading=11,
        textColor=text_color, fontName=FONT)
    desc_col_w = avail_width - 0.6*inch - 1.0*inch - 0.7*inch - 1.1*inch
    invoice_data = [
        [Paragraph('<b>Description</b>', desc_style), 'QTY', 'Price', 'Tax', 'Amount']
    ]
    for item in items_to_use:
        item_total = item.get('qty', 1) * item.get('price', 0) * (1 + item.get('tax', 0) / 100)
        invoice_data.append([
            Paragraph(str(item.get('description', 'Fee Payment')), desc_style),
            str(item.get('qty', 1)),
            f'₹{item.get("price", 0):.2f}',
            f'{item.get("tax", 0):.0f}%',
            f'₹{item_total:.2f}'
        ])
    inv_table = Table(invoice_data, colWidths=[desc_col_w, 0.6*inch, 1.0*inch, 0.7*inch, 1.1*inch], hAlign='LEFT')
    inv_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), accent_color),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor('#FFF2D8')),
        ('FONTNAME', (0, 0), (-1, 0), FONT_BOLD), ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),
        ('FONTNAME', (0, 1), (-1, -1), FONT), ('FONTSIZE', (0, 1), (-1, -1), 9),
        ('TEXTCOLOR', (0, 1), (-1, -1), text_color),
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#EAD7BB')),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8), ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 0.5, accent_color),
    ]))
    elements.append(inv_table)
    elements.append(Spacer(1, 0.1*inch))

    # Total box matching custom styling
    total_box = Table([
        [Paragraph('Total', total_style), Paragraph(f'₹{total_amount:.2f}', total_style)]
    ], colWidths=[1.8*inch, 1.2*inch])
    total_box.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), primary_color),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 8), ('RIGHTPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 8), ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
    ]))
    total_container = Table([[total_box]], colWidths=[avail_width], hAlign='LEFT')
    total_container.setStyle(TableStyle([('ALIGN', (0, 0), (-1, -1), 'RIGHT'), ('LEFTPADDING', (0, 0), (-1, -1), 0)]))
    elements.append(total_container)
    elements.append(Spacer(1, 0.1*inch))

    # Payment info
    payment_box = Table([
        [Paragraph('Payment Information', payment_title_style)],
        [Paragraph(f'<b>UPI Transaction ID:</b> {transaction_id}', normal_style)],
        [Paragraph(f'<b>Payment Date:</b> {date}', normal_style)],
        [Paragraph(f'<b>Total Amount:</b> ₹{total_amount:.2f}', normal_style)]
    ], colWidths=[avail_width], hAlign='LEFT')
    payment_box.setStyle(TableStyle([
        ('LEFTPADDING', (0, 0), (-1, -1), 6), ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 8), ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('LINEABOVE', (0, 0), (-1, 0), 1, accent_color),
        ('LINEBELOW', (0, -1), (-1, -1), 1, accent_color),
    ]))
    elements.append(payment_box)
    elements.append(Spacer(1, 0.1*inch))
    elements.append(Paragraph('This is a computer-generated receipt. No signature is required.', footer_style))

    def draw_background(canvas, document):
        canvas.saveState()
        canvas.setFillColor(bg_color)
        canvas.rect(0, 0, page_width, page_height, fill=1, stroke=0)
        canvas.restoreState()

    doc.build(elements, onFirstPage=draw_background, onLaterPages=draw_background)
    return pdf_path


@app.route('/')
def index():
    return send_from_directory(BASE_DIR, 'index.html')

@app.route('/logo.png')
def logo():
    """Serve school logo.png"""
    return send_from_directory(BASE_DIR, 'logo.png')

@app.route('/html2pdf.bundle.min.js')
def html2pdf_js():
    """Serve html2pdf.js locally to avoid CDN tracking prevention"""
    return send_from_directory(BASE_DIR, 'html2pdf.bundle.min.js')

@app.route('/favicon.ico')
def favicon():
    """Return empty response for favicon to avoid 500 error"""
    return '', 204

@app.route('/api/upload', methods=['POST'])
def upload_file():
    """Handle file upload - CSV or Excel"""
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': 'No file provided'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'error': 'No file selected'}), 400
    
    if not allowed_file(file.filename):
        return jsonify({'success': False, 'error': 'Invalid file type. Use CSV, XLS, or XLSX.'}), 400

    # Save file
    filename = file.filename
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    # Parse file
    try:
        ext = filename.rsplit('.', 1)[1].lower()
        if ext == 'csv':
            headers, rows = read_csv_file(filepath)
        else:
            headers, rows = read_excel_file(filepath)
    except Exception as e:
        return jsonify({'success': False, 'error': f'Failed to parse file: {str(e)}'}), 500

    if not headers or not rows:
        return jsonify({'success': False, 'error': 'File is empty or has no data rows.'}), 400

    # Auto-detect columns
    col_map = auto_detect_columns(headers)

    # Store in session
    session_data['headers'] = headers
    session_data['rows'] = rows
    session_data['filename'] = filename
    session_data['column_map'] = col_map

    return jsonify({
        'success': True,
        'headers': headers,
        'rows': rows,
        'total': len(rows),
        'filename': filename,
        'column_map': col_map
    })

@app.route('/api/data', methods=['GET'])
def get_data():
    """Return current session data"""
    if not session_data['rows']:
        return jsonify({'success': False, 'error': 'No file uploaded yet. Please upload a CSV or Excel file.'}), 404
    
    return jsonify({
        'success': True,
        'headers': session_data['headers'],
        'rows': session_data['rows'],
        'total': len(session_data['rows']),
        'filename': session_data['filename'],
        'column_map': session_data['column_map']
    })

@app.route('/api/update-column-map', methods=['POST'])
def update_column_map():
    """Update column mapping"""
    data = request.get_json()
    if not data:
        return jsonify({'success': False, 'error': 'No data provided'}), 400
    session_data['column_map'] = data.get('column_map', {})
    return jsonify({'success': True, 'column_map': session_data['column_map']})

@app.route('/api/download/<int:row_id>', methods=['GET'])
def download_row(row_id):
    """Download a specific row as CSV"""
    try:
        if row_id >= len(session_data['rows']):
            return jsonify({'success': False, 'error': 'Row not found'}), 404
        
        row = session_data['rows'][row_id]
        output = StringIO()
        writer = csv.DictWriter(output, fieldnames=session_data['headers'])
        writer.writeheader()
        writer.writerow(row)
        
        csv_bytes = output.getvalue().encode('utf-8')
        bytes_io = BytesIO(csv_bytes)
        
        name_col = session_data['column_map'].get('name', '')
        name = row.get(name_col, 'export') if name_col else 'export'
        
        return send_file(bytes_io, mimetype='text/csv', as_attachment=True,
                         download_name=f'row_{row_id}_{name}.csv')
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/download-all', methods=['GET'])
def download_all():
    """Download all data as CSV"""
    try:
        output = StringIO()
        writer = csv.DictWriter(output, fieldnames=session_data['headers'])
        writer.writeheader()
        for row in session_data['rows']:
            writer.writerow(row)
        
        csv_bytes = output.getvalue().encode('utf-8')
        bytes_io = BytesIO(csv_bytes)
        return send_file(bytes_io, mimetype='text/csv', as_attachment=True,
                         download_name=f'{session_data["filename"] or "data"}_all.csv')
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/download-single-pdf', methods=['POST'])
def download_single_pdf():
    """Generate and download a single PDF receipt"""
    try:
        req_data = request.get_json() or {}
        row_data = req_data.get('row_data', {})
        org_settings = req_data.get('org_settings', {})
        col_map = req_data.get('column_map', {})
        invoice_number = req_data.get('invoice_number')
        fee_items = req_data.get('fee_items')
        transaction_id = req_data.get('transaction_id')

        print(f"DEBUG: row_data keys: {list(row_data.keys()) if row_data else 'None'}")
        print(f"DEBUG: col_map: {col_map}")
        print(f"DEBUG: org_settings keys: {list(org_settings.keys()) if org_settings else 'None'}")
        print(f"DEBUG: fee_items: {fee_items}")
        print(f"DEBUG: invoice_number: {invoice_number}")
        print(f"DEBUG: transaction_id: {transaction_id}")

        if not row_data:
            return jsonify({'success': False, 'error': 'No row data provided'}), 400

        # Generate filename
        name_col = col_map.get('name', '')
        student_name = row_data.get(name_col, 'receipt') if name_col else 'receipt'
        clean_name = re.sub(r'[^a-zA-Z0-9\s]', '', str(student_name)).replace(' ', '_')
        filename = f"{clean_name}_Fee_Receipt.pdf"

        print(f"DEBUG: Generating PDF: {filename}")

        # Generate PDF
        pdf_path = generate_receipt_pdf(row_data, filename, col_map, org_settings, invoice_number, fee_items, transaction_id)

        print(f"DEBUG: PDF generated at: {pdf_path}")

        # Check if file exists
        if not os.path.exists(pdf_path):
            return jsonify({'success': False, 'error': 'PDF generation failed - file not created'}), 500

        # Send file
        return send_file(pdf_path, as_attachment=True, download_name=filename, mimetype='application/pdf')
    except Exception as e:
        import traceback
        print(f"ERROR in download_single_pdf: {str(e)}")
        print(f"TRACEBACK: {traceback.format_exc()}")
        return jsonify({'success': False, 'error': str(e), 'traceback': traceback.format_exc()}), 500

@app.route('/api/generate-receipts', methods=['POST'])
def generate_receipts():
    """Generate PDF receipts for all rows"""
    try:
        if not session_data['rows']:
            return jsonify({'success': False, 'error': 'No data loaded. Upload a file first.'}), 400

        req_data = request.get_json() or {}
        org_settings = req_data.get('org_settings', {})
        col_map = req_data.get('column_map', session_data['column_map'])
        edited_records = req_data.get('edited_records', {})

        print(f"DEBUG: Total rows: {len(session_data['rows'])}")
        print(f"DEBUG: edited_records keys: {list(edited_records.keys())}")

        if os.path.exists(RECEIPTS_FOLDER):
            shutil.rmtree(RECEIPTS_FOLDER)
        os.makedirs(RECEIPTS_FOLDER)

        generated_count = 0
        results = []

        generated_names = {}
        for idx, row in enumerate(session_data['rows']):
            name_col = col_map.get('name', '')
            student_name = row.get(name_col, f'Entry_{idx}').strip() if name_col else f'Entry_{idx}'
            clean_name = "".join(c for c in student_name if c.isalnum() or c.isspace()).strip().upper().replace(' ', '_')
            if not clean_name:
                clean_name = f"STUDENT_{idx+1}"

            if clean_name in generated_names:
                generated_names[clean_name] += 1
                filename = f"{clean_name}_{generated_names[clean_name]}.pdf"
            else:
                generated_names[clean_name] = 0
                filename = f"{clean_name}.pdf"

            try:
                # Check if this record has been edited
                edited_data = edited_records.get(str(idx), {})
                invoice_number = edited_data.get('invoice_number') or get_next_invoice_number()
                fee_items = edited_data.get('fee_items')
                transaction_id = edited_data.get('transaction_id')
                
                print(f"DEBUG: Generating PDF for {student_name}, idx={idx}")
                generate_receipt_pdf(row, filename, col_map, org_settings, invoice_number, fee_items, transaction_id)
                generated_count += 1
                results.append({'index': idx+1, 'name': student_name, 'filename': filename, 'invoice_number': invoice_number, 'status': 'success'})
            except Exception as e:
                print(f"ERROR generating PDF for {student_name}: {str(e)}")
                results.append({'index': idx+1, 'name': student_name, 'filename': filename, 'status': 'error', 'error': str(e)})

        return jsonify({
            'success': True, 'total': generated_count,
            'folder': 'receipts', 'data': results
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/download-receipts-zip', methods=['GET'])
def download_receipts_zip():
    """Download all receipts as ZIP"""
    try:
        if not os.path.exists(RECEIPTS_FOLDER):
            return jsonify({'success': False, 'error': 'No receipts generated yet.'}), 404
        
        pdf_files = [f for f in os.listdir(RECEIPTS_FOLDER) if f.endswith('.pdf')]
        if not pdf_files:
            return jsonify({'success': False, 'error': 'No receipts found.'}), 404
        
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
            for pdf_file in pdf_files:
                zf.write(os.path.join(RECEIPTS_FOLDER, pdf_file), arcname=pdf_file)
        
        zip_buffer.seek(0)
        return send_file(zip_buffer, mimetype='application/zip', as_attachment=True,
                         download_name='receipts.zip')
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/receipts-status', methods=['GET'])
def receipts_status():
    """Check receipt generation status"""
    if os.path.exists(RECEIPTS_FOLDER):
        pdf_files = [f for f in os.listdir(RECEIPTS_FOLDER) if f.endswith('.pdf')]
        return jsonify({'success': True, 'exists': True, 'count': len(pdf_files)})
    return jsonify({'success': True, 'exists': False, 'count': 0})

if __name__ == '__main__':
    print(f"Base Directory: {BASE_DIR}")
    print(f"Upload Folder: {UPLOAD_FOLDER}")
    print(f"Receipts Folder: {RECEIPTS_FOLDER}")
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=True, host='0.0.0.0', port=port)
