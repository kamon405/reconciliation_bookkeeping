import pandas as pd
from fuzzywuzzy import fuzz
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Button, Label, Frame
import os
import pdfplumber
import re
from pdf2image import convert_from_path
import pytesseract

class ReconciliationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("QuickBooks vs. Bank Statement Reconciliation")

        # UI Layout
        self.frame = Frame(root, padding=10)
        self.frame.grid(row=0, column=0, padx=20, pady=20)

        Label(self.frame, text="Select QuickBooks Files:").grid(row=0, column=0, sticky="w")
        self.qb_files = []
        Button(self.frame, text="Browse", command=self.load_qb_files).grid(row=0, column=1)

        Label(self.frame, text="Select Bank Statement Files:").grid(row=1, column=0, sticky="w")
        self.bank_files = []
        Button(self.frame, text="Browse", command=self.load_bank_files).grid(row=1, column=1)

        self.reconcile_button = Button(self.frame, text="Reconcile Transactions", command=self.run_reconciliation)
        self.reconcile_button.grid(row=2, column=0, columnspan=2, pady=10)

    def load_qb_files(self):
        files = filedialog.askopenfilenames(filetypes=[("CSV & Excel Files", "*.csv;*.xlsx")])
        if files:
            self.qb_files = files
            messagebox.showinfo("Files Selected", f"Selected {len(files)} QuickBooks files.")

    def load_bank_files(self):
        files = filedialog.askopenfilenames(filetypes=[("CSV, Excel, & PDF Files", "*.csv;*.xlsx;*.pdf")])
        if files:
            self.bank_files = files
            messagebox.showinfo("Files Selected", f"Selected {len(files)} Bank Statement files.")
    
    def load_data(self, files):
        dataframes = []
        for file in files:
            ext = file.split('.')[-1].lower()
            if ext in ['csv', 'xlsx']:
                df = pd.read_csv(file) if ext == 'csv' else pd.read_excel(file)
            elif ext == 'pdf':
                df = self.extract_data_from_pdf(file)
            else:
                raise ValueError(f"Unsupported file format: {file}")
            
            if df.empty:
                continue
            
            for col in ['amount', 'date', 'description']:
                if col not in df.columns:
                    df[col] = pd.Series(dtype='object')
            
            if 'amount' in df.columns and isinstance(df['amount'], pd.Series):
                df['amount'] = pd.to_numeric(df['amount'], errors='coerce').fillna(0.0)
            
            df['date'] = pd.to_datetime(df['date'], errors='coerce')
            df['description'] = df['description'].astype(str).fillna('')
            dataframes.append(df)
        
        if dataframes:
            combined_data = pd.concat(dataframes, ignore_index=True)
        else:
            combined_data = pd.DataFrame(columns=['date', 'description', 'amount'])
        
        return combined_data
    
    def extract_data_from_pdf(self, pdf_file):
        extracted_data = []
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    extracted_data.extend(self.parse_pdf_text(text))
        
        if not extracted_data:
            text = self.extract_text_with_ocr(pdf_file)
            extracted_data.extend(self.parse_pdf_text(text))
        
        if not extracted_data:
            return pd.DataFrame(columns=['date', 'description', 'amount'])
        
        df = pd.DataFrame(extracted_data, columns=['date', 'description', 'amount'])
        if not df.empty and 'amount' in df.columns:
            df['amount'] = pd.to_numeric(df['amount'], errors='coerce').fillna(0.0)
        df['date'] = pd.to_datetime(df['date'], errors='coerce')
        return df
    
    def extract_text_with_ocr(self, pdf_file):
        pages = convert_from_path(pdf_file)
        extracted_text = []
        for page in pages:
            text = pytesseract.image_to_string(page)
            extracted_text.append(text)
        return "\n".join(extracted_text)

    def parse_pdf_text(self, text):
        lines = text.split('\n')
        transactions = []
        for line in lines:
            match = re.search(r"(\d{2}/\d{2})\s+(.+?)\s+(-?\$?\d{1,3}(,\d{3})*(\.\d{2})?)", line)
            if match:
                date, description, amount = match.groups()[0], match.groups()[1], match.groups()[2]
                amount = amount.replace('$', '').replace(',', '')
                try:
                    amount = float(amount)
                    transactions.append([date, description, amount])
                except ValueError:
                    continue
        
        return transactions if transactions else [[]]
    
    def reconcile_transactions(self, qb_data, bank_data):
        qb_data['amount'] = pd.to_numeric(qb_data['amount'], errors='coerce').fillna(0.0)
        bank_data['amount'] = pd.to_numeric(bank_data['amount'], errors='coerce').fillna(0.0)
        qb_data['date'] = pd.to_datetime(qb_data['date'], errors='coerce')
        bank_data['date'] = pd.to_datetime(bank_data['date'], errors='coerce')

        merged = pd.merge(qb_data, bank_data, on=['date', 'amount'], suffixes=('_qb', '_bank'), how='outer', indicator=True)
        merged['match_status'] = 'Unmatched'
        merged.loc[merged['_merge'] == 'both', 'match_status'] = 'Exact Match'
        merged.loc[merged['_merge'] == 'left_only', 'match_status'] = 'Only in QuickBooks'
        merged.loc[merged['_merge'] == 'right_only', 'match_status'] = 'Only in Bank Statement'
        merged.drop(columns=['_merge'], inplace=True)

        merged['is_duplicate'] = merged.duplicated(subset=['date', 'amount'], keep=False)
        return merged
    
    def run_reconciliation(self):
        if not self.qb_files or not self.bank_files:
            messagebox.showerror("Error", "Please select both QuickBooks and Bank Statement files.")
            return

        try:
            qb_data = self.load_data(self.qb_files)
            bank_data = self.load_data(self.bank_files)
            reconciliation_report = self.reconcile_transactions(qb_data, bank_data)

            output_file = os.path.join(os.path.dirname(self.qb_files[0]), f"reconciliation_report_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
            reconciliation_report.to_excel(output_file, index=False)

            messagebox.showinfo("Success", f"Reconciliation completed!\nReport saved as: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

# Run the GUI Application
if __name__ == "__main__":
    root = tk.Tk()
    app = ReconciliationApp(root)
    root.mainloop()
