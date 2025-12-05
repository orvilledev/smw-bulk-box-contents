# 📦 Shipment Grouping Tool

<div align="center">

![Streamlit](https://img.shields.io/badge/Streamlit-FF4B4B?style=for-the-badge&logo=streamlit&logoColor=white)
![Python](https://img.shields.io/badge/Python-3.8+-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-blue?style=for-the-badge)

**An intelligent Excel processing tool that automatically groups shipments, assigns team members, and generates comprehensive reports**

[🚀 Live App](#-live-demo) • [📖 Documentation](#-features) • [🛠️ Installation](#-installation) • [💻 Usage](#-usage)

</div>

---

## ✨ Features

### 🎯 Core Functionality
- **📊 Smart Grouping**: Automatically groups rows based on the first 15 characters of Column C
- **🔤 Shipment Sorting**: Separates shipments (A, B, C...) into alphabetical order
- **📑 Multi-Sheet Export**: Creates one organized sheet per group with professional formatting
- **👥 Team Assignment**: Automatically assigns POs to team members (Paulo, JB, Sunshine, Stephanie, Orville)
- **📈 Pivot Tables**: Generates comprehensive pivot tables for UPC and quantity analysis
- **🔢 Data Preservation**: Maintains leading zeros and prevents scientific notation

### 🎨 Advanced Features
- **Color-Coded Tabs**: Each PO sheet is color-coded based on assigned team member
- **PO Summary Sheet**: Centralized dashboard with team assignments and workflow tracking
- **Box Numbering**: Automatically creates Box# column based on unique carton numbers
- **Summary Statistics**: Calculates total boxes and quantities per PO
- **Professional Formatting**: Excel files with custom headers, borders, and cell formatting

---

## 🚀 Live Demo

**Access the live application:** [View on Streamlit Cloud](https://smw-bulk-box-contents.streamlit.app)

> 💡 Simply upload your Excel file and download the processed, organized spreadsheet in seconds!

---

## 📋 Table of Contents

- [Features](#-features)
- [Installation](#-installation)
- [Usage](#-usage)
- [How It Works](#-how-it-works)
- [Project Structure](#-project-structure)
- [Requirements](#-requirements)
- [Deployment](#-deployment)
- [Contributing](#-contributing)

---

## 🛠️ Installation

### Prerequisites

- Python 3.8 or higher
- pip (Python package installer)

### Step-by-Step Setup

1. **Clone the repository**
   ```bash
   git clone https://github.com/orvilledev/smw-bulk-box-contents.git
   cd smw-bulk-box-contents
   ```

2. **Create a virtual environment** (recommended)
   ```bash
   python -m venv venv
   ```
   
   **Activate the virtual environment:**
   - Windows: `venv\Scripts\activate`
   - macOS/Linux: `source venv/bin/activate`

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

---

## 💻 Usage

### Running Locally

1. **Start the Streamlit app**
   ```bash
   streamlit run smw-bulk.py
   ```

2. **Open your browser**
   - The app will automatically open at `http://localhost:8501`
   - Or manually navigate to the URL shown in the terminal

3. **Upload and Process**
   - Click "Upload Excel File" button
   - Select your `.xlsx` file
   - Wait for processing (usually takes a few seconds)
   - Click "Download Organized Excel File" to save the result

### Input File Requirements

- **Format**: Excel file (`.xlsx`)
- **Minimum Columns**: At least 3 columns (Column C is used for grouping)
- **Column C**: Contains the shipment identifier (first 15 characters used for grouping)

### Output File Structure

The generated Excel file contains:

1. **Original Data Sheet** - Unmodified input data
2. **PO Summary Sheet** - Overview with team assignments and tracking columns
3. **Individual PO Sheets** - One sheet per unique PO number with:
   - Original data grouped by shipment
   - Box numbering
   - Summary statistics (Total Boxes, Total Quantity)
   - Pivot table analysis (UPC × Box# with quantities)

---

## 🔧 How It Works

### Processing Pipeline

```
1. Upload Excel File
   ↓
2. Extract Grouping Keys (First 15 chars of Column C)
   ↓
3. Sort by Group and Shipment Letter
   ↓
4. Assign Team Members to POs
   ↓
5. Generate Individual PO Sheets
   ↓
6. Create Pivot Tables & Summaries
   ↓
7. Apply Formatting & Color Coding
   ↓
8. Export Multi-Sheet Excel File
```

### Team Assignment Logic

- POs are evenly distributed among team members
- Orville receives lower priority for remainder assignments
- Random shuffling ensures fair distribution
- Each PO is color-coded for easy visual identification

---

## 📁 Project Structure

```
smw-bulk-box-contents/
│
├── smw-bulk.py          # Main Streamlit application
├── requirements.txt     # Python dependencies
├── README.md           # Project documentation
├── .gitignore          # Git ignore rules
└── venv/               # Virtual environment (not in repo)
```

---

## 📦 Requirements

### Python Packages

- `streamlit` - Web framework for the app
- `pandas` - Data manipulation and Excel processing
- `openpyxl` - Excel file reading
- `xlsxwriter` - Excel file writing with formatting
- `pytz` - Timezone handling for timestamps

See `requirements.txt` for specific versions.

---

## 🌐 Deployment

### Deploy to Streamlit Cloud

1. **Push to GitHub**
   ```bash
   git add .
   git commit -m "Ready for deployment"
   git push origin main
   ```

2. **Deploy on Streamlit Cloud**
   - Visit [share.streamlit.io](https://share.streamlit.io)
   - Sign in with GitHub
   - Click "New app"
   - Select repository: `orvilledev/smw-bulk-box-contents`
   - Main file: `smw-bulk.py`
   - Click "Deploy"

3. **Your app is live!**
   - Streamlit Cloud automatically redeploys on every push to main branch
   - Access your app at the provided URL

---

## 🤝 Contributing

Contributions are welcome! If you'd like to improve this tool:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

---

## 📝 License

This project is open source and available under the [MIT License](LICENSE).

---

## 👤 Author

**Orville Dev**

- GitHub: [@orvilledev](https://github.com/orvilledev)
- Repository: [smw-bulk-box-contents](https://github.com/orvilledev/smw-bulk-box-contents)

---

## 🙏 Acknowledgments

- Built with [Streamlit](https://streamlit.io/)
- Powered by [Pandas](https://pandas.pydata.org/) and [XlsxWriter](https://xlsxwriter.readthedocs.io/)

---

<div align="center">

**⭐ If you find this project helpful, please consider giving it a star! ⭐**

Made with ❤️ using Streamlit

</div>
