# Shipment Grouping Tool

A Streamlit web application that processes Excel files to group shipments and create organized spreadsheets with multiple sheets.

## Features

- Groups rows based on the first 15 characters of Column C
- Separates shipments (A, B, C...) into alphabetical order
- Creates one sheet per group
- Generates PO Summary with team member assignments
- Creates pivot tables for UPC and quantity analysis
- Preserves leading zeros and prevents scientific notation

## Deployment to Streamlit Cloud

### Prerequisites

1. A GitHub account
2. A Streamlit Cloud account (free at [share.streamlit.io](https://share.streamlit.io))

### Steps to Deploy

1. **Push your code to GitHub:**
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin <your-github-repo-url>
   git push -u origin main
   ```

2. **Deploy to Streamlit Cloud:**
   - Go to [share.streamlit.io](https://share.streamlit.io)
   - Sign in with your GitHub account
   - Click "New app"
   - Select your repository
   - Set the main file path to: `smw-bulk.py`
   - Click "Deploy"

3. **Your app will be live at:**
   `https://<your-app-name>.streamlit.app`

## Local Development

### Installation

1. Create a virtual environment (recommended):
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

### Running Locally

```bash
streamlit run smw-bulk.py
```

The app will open in your browser at `http://localhost:8501`

## Requirements

- Python 3.8 or higher
- See `requirements.txt` for package dependencies

