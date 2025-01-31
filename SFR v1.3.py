import os
import pandas as pd
import numpy as np
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
import win32com.client as win32
import pythoncom

class RowSpecificAnalyzer:
    def __init__(self, excel_file, row_number, company_name):
        self.excel_file = excel_file
        self.row_number = row_number - 1  # Convert to 0-indexed
        self.company_name = company_name
        self.df = pd.read_excel(excel_file, sheet_name=0)  # Read first sheet
    
    def analyze_row(self):
        """Analyze specific row data"""
        if self.row_number < 0 or self.row_number >= len(self.df):
            raise ValueError("Invalid row number")
        
        row_data = self.df.iloc[self.row_number]
        
        # Identify numeric columns
        numeric_columns = row_data[row_data.apply(np.isreal)].index.tolist()
        
        return {
            'row_data': row_data,
            'numeric_columns': numeric_columns
        }
    
    def generate_report(self):
        """Create Word report for row analysis"""
        try:
            analysis = self.analyze_row()
            
            word = win32.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Add()
            
            # Title and Company Information
            doc.Content.Text = f"{self.company_name} - Row {self.row_number + 1} Analysis\n"
            doc.Content.Font.Bold = True
            doc.Content.Paragraphs(1).Alignment = 1
            
            # Report Details
            doc.Content.Text += f"\nReport Generated: {datetime.now().strftime('%Y-%m-%d')}\n\n"
            
            # Row Data
            doc.Content.Text += "Row Details:\n"
            for column, value in analysis['row_data'].items():
                doc.Content.Text += f"{column}: {value}\n"
            
            # Numeric Column Analysis
            doc.Content.Text += "\nNumeric Column Analysis:\n"
            for col in analysis['numeric_columns']:
                doc.Content.Text += f"{col}: {analysis['row_data'][col]}\n"
            
            # Save Report
            save_path = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word Document", "*.docx")],
                title="Save Row Analysis Report"
            )
            
            if save_path:
                doc.SaveAs(save_path)
                doc.Close()
                word.Quit()
                messagebox.showinfo("Success", f"Report saved to {save_path}")
            
        except Exception as e:
            messagebox.showerror("Report Generation Error", str(e))

def main():
    pythoncom.CoInitialize()
    root = tk.Tk()
    root.withdraw()
    
    # Select Excel File
    excel_file = filedialog.askopenfilename(
        title="Select Excel File for Row Analysis",
        filetypes=[("Excel Files", "*.xlsx *.xls *.csv")]
    )
    
    if excel_file:
        # Prompt for Company Name
        company_name = simpledialog.askstring(
            "Company Name", 
            "Enter the Company Name:"
        )
        
        # Prompt for Row Number
        row_number = simpledialog.askinteger(
            "Row Number", 
            "Enter the Row Number to Analyze:"
        )
        
        if company_name and row_number:
            try:
                analyzer = RowSpecificAnalyzer(excel_file, row_number, company_name)
                analyzer.generate_report()
            except Exception as e:
                messagebox.showerror("Error", str(e))
        else:
            messagebox.showinfo("Notice", "Missing company name or row number")
    else:
        messagebox.showinfo("Notice", "No file selected")

if __name__ == "__main__":
    main()