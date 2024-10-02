
# 🧾 Invoice Data Extractor

This Python tool extracts key information from invoice documents uploaded to a Google Drive folder, such as:

- 💰 Amount
- 📅 Date
- 📦 Item name
- 💱 Currency

## 🔑 Setting up Google Auth
- Follow the guide here [Google Developers](https://developers.google.com/workspace/guides/create-credentials) to set up your GCP.
- Download the `credentials.json` and `token.json`, and save these files at the project root.

## ✨ Features

- Extracts structured data from invoices.
- Supports multiple fields including 💰 amount, 📅 date, and 💱 currency.

## 🛠 Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/Telwha/invoice-data-extractor.git
   cd invoice-data-extractor
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## 🚀 Usage

Run the script:
```bash
python main.py
```

## 📋 Requirements

- Python 3.x
- Libraries listed in `requirements.txt`

## 🤝 Contributing

Feel free to fork this repository, make improvements, and submit pull requests.
