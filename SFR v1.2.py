import os
import sys
import pandas as pd
import numpy as np
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client as win32
import pythoncom
from scipy import stats

class SalesForecastAnalyzer:
    def __init__(self, excel_file):
        self.df = pd.read_excel(excel_file)
        self.forecast_data = []
    
    def _calculate_growth_rate(self, sales_column):
        """Calculate compound annual growth rate"""
        try:
            # Linear regression to estimate growth rate
            x = np.arange(len(sales_column))
            log_sales = np.log(sales_column)
            slope, _, _, _, _ = stats.linregress(x, log_sales)
            growth_rate = (np.exp(slope) - 1) * 100
            return max(min(growth_rate, 50), -20)  # Cap growth between -20% and 50%
        except Exception:
            return 5  # Default growth rate if calculation fails
    
    def generate_forecast(self, months=12):
        """Automatically generate forecast for all numeric columns"""
        for column in self.df.select_dtypes(include=[np.number]).columns:
            if 'sales' in column.lower() or 'revenue' in column.lower():
                current_value = self.df[column].iloc[-1]
                growth_rate = self._calculate_growth_rate(self.df[column])
                
                monthly_forecast = [
                    current_value * (1 + growth_rate/100) ** i 
                    for i in range(months)
                ]
                
                self.forecast_data.append({
                    'product': column,
                    'monthly_forecast': monthly_forecast,
                    'total_forecast': sum(monthly_forecast),
                    'average_forecast': np.mean(monthly_forecast),
                    'growth_rate': growth_rate,
                    'current_value': current_value
                })
        
        return self.forecast_data
    
    def create_word_report(self):
        """Create comprehensive Word report"""
        try:
            word = win32.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Add()
            
            # Title and Metadata
            doc.Content.Font.Size = 16
            doc.Content.Text = "Automated Sales Forecast Report\n"
            doc.Content.Font.Bold = True
            doc.Content.Paragraphs(1).Alignment = 1  # Center align
            doc.Content.Text += f"\nGenerated: {datetime.now().strftime('%Y-%m-%d')}\n\n"
            
            # Product Forecasts
            for forecast in self.forecast_data:
                doc.Content.Text += f"Product: {forecast['product']}\n"
                doc.Content.Text += f"Current Value: {forecast['current_value']:,.2f}\n"
                doc.Content.Text += f"Projected Growth Rate: {forecast['growth_rate']:.1f}%\n"
                doc.Content.Text += f"Total Forecast: {forecast['total_forecast']:,.0f}\n"
                doc.Content.Text += f"Average Monthly Forecast: {forecast['average_forecast']:,.0f}\n\n"
                
                # Monthly Breakdown
                doc.Content.Text += "Monthly Forecast Breakdown:\n"
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
    pythoncom.CoInitialize()
    root = tk.Tk()
    root.withdraw()
    
    # Select Excel File
    excel_file = filedialog.askopenfilename(
        title="Select Excel File for Sales Forecast",
        filetypes=[("Excel Files", "*.xlsx *.xls *.csv")]
    )
    
    if excel_file:
        try:
            analyzer = SalesForecastAnalyzer(excel_file)
            analyzer.generate_forecast()
            analyzer.create_word_report()
        except Exception as e:
            messagebox.showerror("Error", str(e))
    else:
        messagebox.showinfo("Notice", "No file selected")

if __name__ == "__main__":
    main()