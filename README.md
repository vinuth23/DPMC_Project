# DPMC Test Case Generator

An AI-assisted test case generator that automatically creates manual test cases from requirement text and form field definitions.

## Features

- **Field-based generation** – define Text, Amount, and Button fields to generate targeted test cases
- **Requirement-based generation** – paste a requirement and let the AI generate 8-10 comprehensive test cases
- **Excel export** – download generated test cases as a formatted `.xlsx` file
- **Modification upload** – upload an existing Excel file to update test cases
- **Screenshot & document upload** – attach screenshots and requirement documents for context

## Prerequisites

- Python 3.8+
- [Ollama](https://ollama.ai) running locally with the `mistral` model

### Install and start Ollama

```bash
# Install Ollama (macOS / Linux)
curl -fsSL https://ollama.com/install.sh | sh

# Pull the Mistral model
ollama pull mistral

# Start the Ollama server (runs on http://localhost:11434)
ollama serve
```

## Setup

```bash
# Clone the repository
git clone https://github.com/vinuth23/DPMC_Project.git
cd DPMC_Project

# Create and activate a virtual environment
python -m venv venv
source venv/bin/activate   # Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

## Running the Application

```bash
python app.py
```

Open your browser at [http://localhost:5000](http://localhost:5000).

## How to Use

1. **Step 1 – Define Fields/Buttons** – click **Add** to add form fields and set their type (Text, Amount, or Button).
2. **Step 2 – Upload Screenshots** *(optional)* – attach UI screenshots.
3. **Step 3 – Enter Requirement** *(optional)* – paste requirement text or upload a document.
4. Click **Generate Test Cases**.
5. Review the generated test cases in **Step 4**.
6. Click **Download as Excel** to export.

## Project Structure

```
DPMC_Project/
├── app.py              # Flask backend + test-case generation logic
├── requirements.txt    # Python dependencies
├── templates/
│   └── index.html      # Single-page UI
└── static/
    ├── script.js       # Frontend logic
    └── style.css       # Styles
```

## API Endpoints

| Method | Path | Description |
|--------|------|-------------|
| GET | `/` | Render the main page |
| POST | `/generate` | Generate test cases from requirement / fields |
| POST | `/export` | Export test cases to Excel |
| GET | `/health` | Health check (checks Ollama availability) |
