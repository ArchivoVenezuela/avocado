✅ README.md for Avocado Desktop (v2.7)
Save the following content as README.md in your repository folder:

markdown
Copy
Edit
# 🥑 Avocado Desktop v2.7

Avocado is a desktop tool for retrieving and enriching bibliographic metadata from OCLC WorldCat using book titles and author names provided in a CSV file.

## 💡 Features

- Search WorldCat for OCLC numbers using book metadata
- Retrieve complete bibliographic metadata for matching records
- Export enriched records to a CSV file
- Built with Python and Tkinter for a simple desktop interface

## 🖥️ Requirements

- Python 3.9+
- OCLC WSKey and Secret (you must have a valid OCLC API account)

## 📦 Installation

1. Clone the repository or download the source code:
   ```bash
   git clone https://github.com/yourusername/avocado-desktop.git
   cd avocado-desktop
Create a virtual environment (recommended):

bash
Copy
Edit
python -m venv venv
venv\Scripts\activate  # Windows
Install required dependencies:

bash
Copy
Edit
pip install -r requirements.txt
Set your OCLC credentials as environment variables:

bash
Copy
Edit
set OCLC_WSKEY=your_key
set OCLC_WSSECRET=your_secret
🚀 Usage
Run the desktop app:

bash
Copy
Edit
python avocado_v2_7.py
Choose your CSV file (must include Title and Author columns).

Select an output folder.

Click Start Process.

A CSV with enriched metadata will be saved in the selected folder.

📁 Sample CSV Format
Title	Author
Transilvania unplugged	John Doe
Dear parent or guardian	Jane Smith

📄 Output Fields
OCLC number

Title

Publisher

ISBN

...plus additional metadata if available

🛠️ Development Notes
Uses OCLC’s WorldCat Search API v2.

Network calls include exponential backoff for rate limiting.

📜 License
MIT License. See LICENSE for details.