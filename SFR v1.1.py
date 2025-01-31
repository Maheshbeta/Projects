import os
import sys
import pandas as pd
import numpy as np
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from pathlib import Path
import win32com.client as win32
import pythoncom

class SalesForecastDialog:
    def __init__(self, master=None):
        """Initialize the sales forecast dialog system"""
        self.master = master or tk.Tk()
        self.master.withdraw()  # Hide the main window
        self.products = []
    
    def select_excel_file(self):
        """Open file dialog to select Excel file"""
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls *.csv")]
        )
        return file_path
    
    def read_excel_data(self, file_path):
        """Read data from Excel file"""
        try:
            df = pd.read_excel(file_path)
            return df
        except Exception as e:
            messagebox.showerror("Error", f"Could not read file: {str(e)}")
            return None
    
    def input_product_details(self):
        """Input product details via dialog boxes"""
        while True:
            product_type = simpledialog.askstring(
                "Product Input", 
                "Enter Product Type (or Cancel to finish):"
            )
            
            if product_type is None:
                break
            
            try:
                current_units = simpledialog.askfloat(
                    "Product Units", 
                    f"Enter current monthly unit sales for {product_type}:"
                )
                
                growth_rate = simpledialog.askfloat(
                    "Growth Rate", 
                    f"Enter expected growth rate (%) for {product_type}:"
                )
                
                if current_units is not None and growth_rate is not None:
                    self.products.append({
                        'name': product_type,
                        'current_units': current_units,
                        'growth_rate': growth_rate
                    })
            except Exception as e:
                messagebox.showerror("Input Error", str(e))
    
    def generate_forecast(self, months=12):
        """Generate sales forecast"""
        forecast_data = []
        
        for product in self.products:
            monthly_forecast = [
                product['current_units'] * (1 + product['growth_rate']/100) ** i 
                for i in range(months)
            ]
            
            forecast_data.append({
                'product': product['name'],
                'monthly_forecast': monthly_forecast,
                'total_forecast': sum(monthly_forecast),
                'average_forecast': np.mean(monthly_forecast),
                'growth_rate': product['growth_rate']
            })
        
        return forecast_data
    
    def create_word_report(self, forecast_data):
        """Create Word report using win32com"""
        try:
            word = win32.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Add()
            
            # Title
            doc.Content.Font.Size = 16
            doc.Content.Text = "Sales Forecast Report\n"
            doc.Content.Font.Bold = True
            doc.Content.Paragraphs(1).Alignment = 1  # Center align
            
            # Date
            doc.Content.Text += f"\nGenerated: {datetime.now().strftime('%Y-%m-%d')}\n\n"
            
            # Product Forecasts
            for forecast in forecast_data:
                doc.Content.Text += f"Product: {forecast['product']}\n"
                doc.Content.Text += f"Growth Rate: {forecast['growth_rate']:.1f}%\n"
                doc.Content.Text += f"Total Forecast: {forecast['total_forecast']:,.0f} units\n"
                doc.Content.Text += f"Average Monthly Forecast: {forecast['average_forecast']:,.0f} units\n\n"
                
                # Monthly Breakdown
                doc.Content.Text += "Monthly Breakdown:\n"
                monthly_text = " | ".join([f"{val:,.0f}" for val in forecast['monthly_forecast']])
                doc.Content.Text += monthly_text + "\n\n"
            
            # Save File Dialog
            save_path = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word Document", "*.docx")],
                title="Save Sales Forecast Report"
            )
            
            if save_path:
                doc.SaveAs(save_path)
                doc.Close()
                word.Quit()
                messagebox.showinfo("Success", f"Report saved to {save_path}")
            
        except Exception as e:
            messagebox.showerror("Report Generation Error", str(e))

def main():
    pythoncom.CoInitialize()  # Initialize COM for threading
    forecast_dialog = SalesForecastDialog()
    
    # Optional: Excel file selection
    excel_file = forecast_dialog.select_excel_file()
    if excel_file:
        excel_data = forecast_dialog.read_excel_data(excel_file)
        # You can add logic to pre-populate or process Excel data if needed
    
    # Input product details
    forecast_dialog.input_product_details()
    
    # Generate forecast
    if forecast_dialog.products:
        forecast_data = forecast_dialog.generate_forecast()
        forecast_dialog.create_word_report(forecast_data)
    else:
        messagebox.showinfo("Notice", "No products entered for forecast.")

if __name__ == "__main__":
    main()