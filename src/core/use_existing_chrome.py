import time
import pandas as pd
import os
import sys
import re
import json
import urllib.parse
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from threading import Thread
import webbrowser
from datetime import datetime

class ManualGoogleMapsScraper:
    def __init__(self):
        self.results = {}
        self.current_company = None

    def search_company(self, company_name, address="", district=""):
        """Open Google Maps in default browser to search for company"""
        self.current_company = company_name
        search_term = urllib.parse.quote_plus(f"{company_name} {address} {district}")
        maps_url = f"https://www.google.com/maps/search/{search_term}"
        
        print(f"Opening browser for: {company_name}")
        webbrowser.open(maps_url)
        
        # Create empty data structure for results
        self.results[company_name] = {
            "name": company_name,
            "address": address if address else None,
            "phone": None,
            "website": None,
            "email": None,
            "category": None,
            "search_timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        
        return self.results[company_name]
    
    def save_company_info(self, company_name, phone=None, website=None, email=None, address=None, category=None):
        """Save manually entered company information"""
        if company_name not in self.results:
            self.results[company_name] = {
                "name": company_name,
                "address": None,
                "phone": None,
                "website": None,
                "email": None,
                "category": None,
                "search_timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
        
        if phone:
            self.results[company_name]["phone"] = phone
        if website:
            self.results[company_name]["website"] = website
        if email:
            self.results[company_name]["email"] = email
        if address:
            self.results[company_name]["address"] = address
        if category:
            self.results[company_name]["category"] = category
        
        return self.results[company_name]
    
    def process_companies_from_excel(self, excel_file):
        """Process companies from Excel file (manually)"""
        try:
            df = pd.read_excel(excel_file)
            
            # Define possible column name formats
            column_mappings = {
                'no': ['No.', 'No', 'NO', 'no.', 'no', 'Nomor', 'nomor', '#'],
                'company': ['Nama Perusahaan', 'Nama_Perusahaan', 'NamaPerusahaan', 'Nama', 'Company', 'company', 'NAMA'],
                'address': ['Alamat', 'ALAMAT', 'alamat', 'Address', 'address'],
                'district': ['Kecamatan', 'KECAMATAN', 'kecamatan', 'District', 'district']
            }
            
            # Map actual column names to standard names
            column_map = {}
            for standard_key, possible_names in column_mappings.items():
                found = False
                for name in possible_names:
                    if name in df.columns:
                        column_map[standard_key] = name
                        found = True
                        break
            
            # Check if required columns exist
            if 'company' not in column_map:
                print("Error: Company name column not found in Excel file")
                return None
            
            # Extract companies
            companies = []
            for i, row in df.iterrows():
                company_data = {
                    "name": row[column_map['company']],
                    "address": row[column_map['address']] if 'address' in column_map else "",
                    "district": row[column_map['district']] if 'district' in column_map else ""
                }
                companies.append(company_data)
            
            print(f"Loaded {len(companies)} companies from {excel_file}")
            return companies
        
        except Exception as e:
            print(f"Error loading Excel file: {str(e)}")
            return None
    
    def save_to_excel(self, filename="manual_company_data.xlsx"):
        """Save results to Excel file"""
        if not self.results:
            print("No data to save")
            return False
        
        try:
            # Convert results dictionary to DataFrame
            data = []
            for company, info in self.results.items():
                data.append(info)
            
            df = pd.DataFrame(data)
            df.to_excel(filename, index=False)
            print(f"Data saved to {filename}")
            return True
        
        except Exception as e:
            print(f"Error saving data to Excel: {str(e)}")
            return False


class ManualScraperGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Manual Google Maps Scraper")
        self.root.geometry("700x600")
        
        self.scraper = ManualGoogleMapsScraper()
        self.companies = []
        self.current_index = -1
        
        self.setup_ui()
    
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # File selection frame
        file_frame = ttk.LabelFrame(main_frame, text="Excel File", padding=10)
        file_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Input file
        ttk.Label(file_frame, text="Input Excel File:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.input_file_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.input_file_var, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_input_file).grid(row=0, column=2, padx=5, pady=5)
        
        # Output file
        ttk.Label(file_frame, text="Output Excel File:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.output_file_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.output_file_var, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_output_file).grid(row=1, column=2, padx=5, pady=5)
        
        # Company information frame
        info_frame = ttk.LabelFrame(main_frame, text="Company Information", padding=10)
        info_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Company name
        ttk.Label(info_frame, text="Company:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.company_var = tk.StringVar()
        ttk.Entry(info_frame, textvariable=self.company_var, width=50, state="readonly").grid(row=0, column=1, columnspan=2, padx=5, pady=5)
        
        # Address
        ttk.Label(info_frame, text="Address:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.address_var = tk.StringVar()
        ttk.Entry(info_frame, textvariable=self.address_var, width=50).grid(row=1, column=1, columnspan=2, padx=5, pady=5)
        
        # Phone
        ttk.Label(info_frame, text="Phone:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.phone_var = tk.StringVar()
        ttk.Entry(info_frame, textvariable=self.phone_var, width=50).grid(row=2, column=1, columnspan=2, padx=5, pady=5)
        
        # Website
        ttk.Label(info_frame, text="Website:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        self.website_var = tk.StringVar()
        ttk.Entry(info_frame, textvariable=self.website_var, width=50).grid(row=3, column=1, columnspan=2, padx=5, pady=5)
        
        # Email
        ttk.Label(info_frame, text="Email:").grid(row=4, column=0, sticky=tk.W, padx=5, pady=5)
        self.email_var = tk.StringVar()
        ttk.Entry(info_frame, textvariable=self.email_var, width=50).grid(row=4, column=1, columnspan=2, padx=5, pady=5)
        
        # Category
        ttk.Label(info_frame, text="Category:").grid(row=5, column=0, sticky=tk.W, padx=5, pady=5)
        self.category_var = tk.StringVar()
        ttk.Entry(info_frame, textvariable=self.category_var, width=50).grid(row=5, column=1, columnspan=2, padx=5, pady=5)
        
        # Navigation frame
        nav_frame = ttk.Frame(info_frame)
        nav_frame.grid(row=6, column=0, columnspan=3, pady=10)
        
        # Navigation buttons
        ttk.Button(nav_frame, text="Previous", command=self.previous_company).pack(side=tk.LEFT, padx=5)
        ttk.Button(nav_frame, text="Save Current", command=self.save_current).pack(side=tk.LEFT, padx=5)
        ttk.Button(nav_frame, text="Next", command=self.next_company).pack(side=tk.LEFT, padx=5)
        ttk.Button(nav_frame, text="Open in Browser", command=self.open_in_browser).pack(side=tk.LEFT, padx=5)
        
        # Progress frame
        progress_frame = ttk.LabelFrame(main_frame, text="Progress", padding=10)
        progress_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Progress info
        self.progress_var = tk.StringVar(value="No companies loaded")
        ttk.Label(progress_frame, textvariable=self.progress_var).pack(padx=5, pady=5)
        
        # Action buttons frame
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Action buttons
        ttk.Button(action_frame, text="Load Excel", command=self.load_excel).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(action_frame, text="Save All Results", command=self.save_results).pack(side=tk.LEFT, padx=5, pady=5)
    
    def browse_input_file(self):
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if filename:
            self.input_file_var.set(filename)
            # Set default output filename
            if not self.output_file_var.get():
                base, ext = os.path.splitext(filename)
                self.output_file_var.set(f"{base}_manual_results{ext}")
    
    def browse_output_file(self):
        filename = filedialog.asksaveasfilename(
            title="Save As",
            defaultextension=".xlsx",
            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
        )
        if filename:
            self.output_file_var.set(filename)
    
    def load_excel(self):
        input_file = self.input_file_var.get()
        if not input_file:
            messagebox.showerror("Error", "Please select an input Excel file")
            return
        
        # Load companies from Excel
        self.companies = self.scraper.process_companies_from_excel(input_file)
        
        if not self.companies:
            messagebox.showerror("Error", "Failed to load companies from Excel file")
            return
        
        # Reset current index
        self.current_index = -1
        
        # Update progress
        self.progress_var.set(f"Loaded {len(self.companies)} companies. Ready to start.")
        
        # Move to first company
        self.next_company()
    
    def save_current(self):
        if not self.companies or self.current_index < 0 or self.current_index >= len(self.companies):
            messagebox.showinfo("Info", "No company selected")
            return
        
        # Get current company
        company = self.companies[self.current_index]
        
        # Save company info
        self.scraper.save_company_info(
            company_name=company["name"],
            address=self.address_var.get(),
            phone=self.phone_var.get(),
            website=self.website_var.get(),
            email=self.email_var.get(),
            category=self.category_var.get()
        )
        
        messagebox.showinfo("Success", f"Saved data for: {company['name']}")
    
    def next_company(self):
        if not self.companies:
            messagebox.showinfo("Info", "No companies loaded. Please load an Excel file first.")
            return
        
        # Save current company data if we have a valid index
        if self.current_index >= 0 and self.current_index < len(self.companies):
            self.save_current()
        
        # Move to next company
        self.current_index += 1
        
        # Check if we're done
        if self.current_index >= len(self.companies):
            messagebox.showinfo("Done", "You've processed all companies. You can now save the results.")
            self.current_index = len(self.companies) - 1
            return
        
        # Load next company
        self.load_company(self.current_index)
    
    def previous_company(self):
        if not self.companies:
            messagebox.showinfo("Info", "No companies loaded. Please load an Excel file first.")
            return
        
        # Save current company data
        if self.current_index >= 0 and self.current_index < len(self.companies):
            self.save_current()
        
        # Move to previous company
        self.current_index -= 1
        
        # Check if we're at the beginning
        if self.current_index < 0:
            messagebox.showinfo("Info", "Already at the first company.")
            self.current_index = 0
            return
        
        # Load previous company
        self.load_company(self.current_index)
    
    def load_company(self, index):
        if index < 0 or index >= len(self.companies):
            return
        
        # Get company data
        company = self.companies[index]
        
        # Update UI
        self.company_var.set(company["name"])
        
        # Set address if available
        self.address_var.set(company.get("address", ""))
        
        # Clear other fields
        self.phone_var.set("")
        self.website_var.set("")
        self.email_var.set("")
        self.category_var.set("")
        
        # Check if we already have data for this company
        if company["name"] in self.scraper.results:
            data = self.scraper.results[company["name"]]
            if data["phone"]:
                self.phone_var.set(data["phone"])
            if data["website"]:
                self.website_var.set(data["website"])
            if data["email"]:
                self.email_var.set(data["email"])
            if data["category"]:
                self.category_var.set(data["category"])
        
        # Update progress
        self.progress_var.set(f"Processing company {index+1} of {len(self.companies)}: {company['name']}")
    
    def open_in_browser(self):
        if not self.companies or self.current_index < 0 or self.current_index >= len(self.companies):
            messagebox.showinfo("Info", "No company selected")
            return
        
        # Get current company
        company = self.companies[self.current_index]
        
        # Open in browser
        self.scraper.search_company(
            company_name=company["name"],
            address=company.get("address", ""),
            district=company.get("district", "")
        )
    
    def save_results(self):
        if not self.scraper.results:
            messagebox.showinfo("Info", "No data to save. Please process some companies first.")
            return
        
        output_file = self.output_file_var.get()
        if not output_file:
            messagebox.showerror("Error", "Please specify an output Excel file")
            return
        
        # Save to Excel
        if self.scraper.save_to_excel(output_file):
            messagebox.showinfo("Success", f"Data saved to {output_file}")
        else:
            messagebox.showerror("Error", "Failed to save data")


if __name__ == "__main__":
    root = tk.Tk()
    app = ManualScraperGUI(root)
    root.mainloop() 