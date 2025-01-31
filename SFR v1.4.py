import os
import pandas as pd
import numpy as np
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
import win32com.client as win32
import pythoncom

def get_row_input(df):
    """
    Get and validate row input with a fixed range of 1-78 and actual file size
    """
    max_row = min(78, len(df))  # Use either 78 or actual file size, whichever is smaller
    
    while True:
        row_number = simpledialog.askinteger(
            "Row Number",
            f"Enter the Row Number to Analyze (1-{max_row}):",
            minvalue=1,
            maxvalue=max_row
        )
        
        if row_number is None:  # User cancelled
            return None
            
        if 1 <= row_number <= max_row:
            return row_number
        else:
            messagebox.showerror(
                "Invalid Input",
                f"Please enter a row number between 1 and {max_row}"
            )

class RowColumnAnalyzer:
    def __init__(self, excel_file, row_number, start_column, end_column, company_name):
        self.excel_file = excel_file
        self.df = pd.read_excel(excel_file, sheet_name=0)  # Read first sheet
        
        # Validate row number against both fixed bounds and file size
        max_row = min(78, len(self.df))
        if not 1 <= row_number <= max_row:
            raise ValueError(f"Row number must be between 1 and {max_row}")
        
        self.row_number = row_number - 1  # Convert to 0-indexed
        self.start_column = start_column
        self.end_column = end_column
        self.company_name = company_name
    
    def _get_column_letter(self, column_number):
        """Convert column number to Excel letter"""
        result = ""
        while column_number > 0:
            column_number -= 1
            result = chr(column_number % 26 + 65) + result
            column_number //= 26
        return result
    
    def _get_column_number(self, column_letter):
        """Convert Excel letter to column number"""
        result = 0
        for char in column_letter.upper():
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result
    
    def analyze_range(self):
        """Analyze specific row and column range data"""
        # Removed the row number validation here since it's already validated in __init__
        
        # Convert column letters to numbers if needed
        start_col = self.start_column if isinstance(self.start_column, int) else self._get_column_number(self.start_column)
        end_col = self.end_column if isinstance(self.end_column, int) else self._get_column_number(self.end_column)
        
        # Get column labels
        column_labels = self.df.columns[start_col-1:end_col]
        row_data = self.df.iloc[self.row_number, start_col-1:end_col]
        
        # Identify numeric columns
        numeric_data = row_data[row_data.apply(np.isreal)]
        
        return {
            'row_data': row_data,
            'column_labels': column_labels,
            'numeric_data': numeric_data
        }
    
    def generate_report(self):
        """Create Word report for range analysis"""
        try:
            analysis = self.analyze_range()
            
            word = win32.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Add()
            
            # Title and Company Information
            doc.Content.Font.Size = 16
            doc.Content.Text = f"{self.company_name} - Detailed Range Analysis\n"
            doc.Content.Font.Bold = True
            doc.Content.Paragraphs(1).Alignment = 1
            
            # Report Details
            doc.Content.Text += f"\nReport Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            doc.Content.Text += f"Analysis for Row: {self.row_number + 1}\n"
            doc.Content.Text += f"Column Range: {self.start_column} to {self.end_column}\n\n"
            
            # Data Analysis
            doc.Content.Text += "Selected Range Data:\n"
            for col, value in analysis['row_data'].items():
                formatted_value = f"{value:,.2f}" if isinstance(value, (int, float)) else str(value)
                doc.Content.Text += f"{col}: {formatted_value}\n"
            
            # Numeric Analysis
            if not analysis['numeric_data'].empty:
                doc.Content.Text += "\nNumeric Data Analysis:\n"
                doc.Content.Text += f"Total Sum: {analysis['numeric_data'].sum():,.2f}\n"
                doc.Content.Text += f"Average Value: {analysis['numeric_data'].mean():,.2f}\n"
                doc.Content.Text += f"Maximum Value: {analysis['numeric_data'].max():,.2f}\n"
                doc.Content.Text += f"Minimum Value: {analysis['numeric_data'].min():,.2f}\n"
            
            # Save Report
            save_path = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word Document", "*.docx")],
                title="Save Range Analysis Report"
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
        title="Select Excel File for Analysis",
        filetypes=[("Excel Files", "*.xlsx *.xls *.csv")]
    )
    
    if excel_file:
        try:
            # Load DataFrame
            df = pd.read_excel(excel_file, sheet_name=0)
            
            # Prompt for Company Name
            company_name = simpledialog.askstring(
                "Company Name", 
                "Enter the Company Name:"
            )
            
            if not company_name:  # User cancelled
                messagebox.showinfo("Notice", "Operation cancelled")
                return
                
            # Get validated row number
            row_number = get_row_input(df)
            if row_number is None:  # User cancelled
                messagebox.showinfo("Notice", "Operation cancelled")
                return
            
            # Prompt for Column Range
            start_column = get_column_input("Enter Start Column (number or letter):")
            if start_column is None:  # User cancelled
                messagebox.showinfo("Notice", "Operation cancelled")
                return
                
            end_column = get_column_input("Enter End Column (number or letter):")
            if end_column is None:  # User cancelled
                messagebox.showinfo("Notice", "Operation cancelled")
                return
            
            analyzer = RowColumnAnalyzer(excel_file, row_number, start_column, end_column, company_name)
            analyzer.generate_report()
            
        except Exception as e:
            messagebox.showerror("Error", str(e))
    else:
        messagebox.showinfo("Notice", "No file selected")

if __name__ == "__main__":
    main()