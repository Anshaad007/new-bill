# CSV Data Viewer & Invoice Generator with PDF Download

A powerful web-based application to view, search, and download CSV data with professional invoice generation and PDF export capabilities.

## 🎯 Features

### Core Functionality
- ✅ **View CSV Data**: Display all rows and columns in an organized table
- ✅ **Search & Filter**: Real-time search across all columns
- ✅ **Invoice Preview**: Beautiful invoice modal with professional design
- ✅ **PDF Download**: Generate professional PDF receipts from invoice data
- ✅ **CSV Download**: Download individual rows or entire dataset as CSV
- ✅ **Responsive Design**: Works seamlessly on desktop and mobile devices
- ✅ **Real-time Statistics**: Shows total rows, displayed rows, and column count

### Invoice Features
- Professional invoice layout with school branding
- GSTIN and student information display
- Itemized fee breakdown with amounts
- Payment transaction tracking
- Clean, printable invoice design
- Customizable invoice content from CSV data

## 📁 Project Structure

```
auto/
├── extract.py              # Flask backend server with API endpoints
├── index.html              # Web interface with invoice modal
├── planets.csv             # Sample CSV data file
├── .venv/                  # Python virtual environment
├── README.md              # This file
└── bill-alif-main/        # Original bill generator design files (reference)
```

## 🛠️ Setup & Installation

### Prerequisites
- Python 3.7+
- Modern web browser (Chrome, Firefox, Safari, Edge)
- Flask (will be installed)

### Installation Steps

1. **Navigate to project folder**:
   ```bash
   cd c:\Users\Muhammed Anshad K T\Desktop\auto
   ```

2. **Create virtual environment** (if not already done):
   ```bash
   python -m venv .venv
   ```

3. **Activate virtual environment**:
   ```bash
   .\.venv\Scripts\activate
   ```

4. **Install dependencies**:
   ```bash
   pip install flask
   ```

## 🚀 Running the Application

1. **Start the Flask server**:
   ```bash
   python extract.py
   ```

2. **Open browser** and navigate to:
   ```
   http://localhost:5000
   ```

3. **You will see**:
   - Statistics dashboard (Total Rows, Displayed Rows, Columns)
   - Search box for filtering data
   - Complete data table with all records
   - Action buttons for each row (Preview & CSV download)
   - Download All button for entire dataset

## 📊 How to Use

### Viewing Data

1. **Page loads** automatically with all data from `planets.csv`
2. **Statistics panel** shows total records and column count
3. **Search box** filters data in real-time across all columns

### Working with Invoices

1. **Click "🔍 Preview"** button on any row
2. **Invoice modal** opens showing:
   - School information (ALIF ONLINE MORAL SCHOOL)
   - Student name and class
   - Receipt date and transaction ID
   - Fee breakdown with amounts
   - Payment information
   - Total amount

3. **Choose download option**:
   - **📥 Download CSV**: Downloads row as CSV file
   - **📄 Download PDF**: Generates and downloads professional PDF invoice
   - **✕ Close**: Closes the preview modal

### Downloading Data

#### Download Individual Row as CSV
- Click "⬇️ CSV" button next to any row
- Downloads as `row_[index]_[student_name].csv`

#### Download Individual Row as PDF
- Click "🔍 Preview" to open invoice modal
- Click "📄 Download PDF" button
- Downloads as `[student_name]_receipt.pdf`

#### Download All Data
- Click **"⬇️ Download All"** button in search section
- Downloads entire dataset as `planets_all.csv`

## 🔧 API Endpoints

The Flask backend provides REST API endpoints:

### `GET /api/data`
Returns all CSV data as JSON
```json
{
  "success": true,
  "headers": ["Timestamp", "Fee type", "Name of student", ...],
  "rows": [{...}, {...}],
  "total": 28
}
```

### `GET /api/download/<row_id>`
Downloads specific row as CSV file
- **Parameters**: `row_id` (integer) - zero-indexed row number
- **Response**: CSV file download

### `GET /api/download-all`
Downloads entire CSV file
- **Response**: Complete CSV file as `planets_all.csv`

## 📝 CSV Format Requirements

Your CSV file should have:
- **First row**: Column headers/names
- **Data rows**: Consistent columns with actual data
- **Supported columns** (auto-detected):
  - Name of student
  - Fee type
  - Specify month or term
  - Date of payment
  - Amount
  - Transaction ID
  - Class
  - Any other columns

### Example CSV Structure
```
Timestamp,Fee type,Specify month or term,Name of student,Date of payment,Amount,Transaction ID,Class
2025/09/03 10:51:02 am GMT+5:30,,,HAYA,2/9/2025,1770,5.24591E+11,PLANETS
2025/09/03 10:52:45 am GMT+5:30,,,RAZAN,3/9/2025,1770,5.24652E+11,PLANETS
```

## 🎨 Customization

### Change CSV File
1. Replace `planets.csv` with your own CSV file
2. Keep the same filename or update reference in `extract.py`
3. Refresh browser - system auto-detects all columns

### Modify School Information
Edit the invoice generation in `index.html` (in `generateInvoicePreview()` function):
```javascript
const schoolName = 'ALIF ONLINE MORAL SCHOOL';
const schoolGST = '32ACAFA0267H1ZY';
// Update these values with your school details
```

### Update Styling
Edit the CSS in `<style>` section of `index.html`:
- Colors: Update `--brand-dark`, `--brand-gold`, `--brand-cream` variables
- Fonts: Modify `font-family` declarations
- Layout: Adjust padding, margins, and sizing

## 📋 Browser Download Handling

**Note**: PDF downloads save to your browser's default download folder:
- **Chrome/Edge**: Usually `C:\Users\[YourName]\Downloads\`
- **Firefox**: Same as above
- **Safari**: Downloads folder or specified location

## 🔐 Security Considerations

- **Local Operation**: All processing happens on your machine
- **No External Data**: Data is not sent to external servers
- **Browser-based PDF**: Uses html2pdf library for client-side PDF generation
- **No Authentication**: Add authentication if deploying to web

## ⚠️ Notes

- This is a **development server**
- For production deployment, use a WSGI server (Gunicorn, uWSGI, etc.)
- Large CSV files may take a moment to render
- PDF generation depends on browser's JavaScript engine
- Column auto-detection works with any CSV structure

## 🐛 Troubleshooting

### Server Won't Start
```bash
# Check if port 5000 is already in use
netstat -ano | findstr :5000

# If in use, kill the process or change Flask port in extract.py
```

### Data Not Loading
- Verify `planets.csv` exists in the same directory as `extract.py`
- Check browser console (F12) for JavaScript errors
- Ensure CSV file has proper headers in first row

### PDF Download Not Working
- Check if browser allows downloads from localhost
- Verify JavaScript console has no errors (F12)
- Try different browser if issue persists
- Ensure sufficient disk space for PDF generation

### Search Not Working
- Check if data loaded successfully
- Try refreshing page (F5)
- Clear browser cache and try again

## 📞 Support

For issues or questions:
1. Check the troubleshooting section above
2. Review browser console errors (press F12)
3. Verify CSV file format and data
4. Check that Flask server is running and accessible

## 📄 File Descriptions

| File | Purpose |
|------|---------|
| `extract.py` | Flask backend with API endpoints |
| `index.html` | Web UI with invoice modal and download features |
| `planets.csv` | Sample CSV data |
| `.venv/` | Python virtual environment |
| `README.md` | This documentation |

## 🎓 Design Credits

Invoice design inspired by [bill-alif](https://github.com/bill-alif) project featuring professional billing templates for educational institutions.

## ⚡ Quick Start Checklist

- [ ] Python 3.7+ installed
- [ ] Virtual environment created
- [ ] Flask installed (`pip install flask`)
- [ ] CSV file in project folder
- [ ] Server running on http://localhost:5000
- [ ] Page loads with data table
- [ ] Search functionality works
- [ ] Preview button opens invoice modal
- [ ] PDF download works
- [ ] CSV download works

## 📄 License

This project is provided as-is for personal and educational use.

---

**Last Updated**: May 20, 2026

**System Features**: CSV Viewer • Invoice Generator • PDF Export • Data Search • Multi-Format Download

