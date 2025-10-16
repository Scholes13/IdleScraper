import os
import sys
import time
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from threading import Thread
import traceback

# Update imports to use the new file structure
try:
    from src.core.maps_scraper import GoogleMapsScraper
    from src.core.batch_scraper import batch_scrape, load_companies
except ImportError:
    # Fallback for direct execution
    from maps_scraper import GoogleMapsScraper
    from batch_scraper import batch_scrape, load_companies

class GoogleMapsScraperGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Google Maps Scraper")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        self.scraper = None
        self.scraping_thread = None
        self.is_scraping = False
        
        # Default settings
        self.enable_similar_search = tk.BooleanVar(value=False)
        self.similarity_threshold = tk.DoubleVar(value=0.7)
        self.enable_website_scraping = tk.BooleanVar(value=False)
        
        self.setup_ui()
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Input frame
        input_frame = ttk.LabelFrame(main_frame, text="Input", padding=10)
        input_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Company name
        ttk.Label(input_frame, text="Company Name:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.company_entry = ttk.Entry(input_frame, width=60)
        self.company_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        ttk.Button(input_frame, text="Search", command=self.search_company).grid(row=0, column=2, padx=5, pady=5)
        
        # Settings frame
        settings_frame = ttk.LabelFrame(main_frame, text="Settings", padding=10)
        settings_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Smart search checkbox
        smart_search_check = ttk.Checkbutton(
            settings_frame, 
            text="Enable Smart Search (find similar companies)",
            variable=self.enable_similar_search
        )
        smart_search_check.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5, columnspan=3)
        
        # Website scraping checkbox
        website_scraping_check = ttk.Checkbutton(
            settings_frame, 
            text="Enable Website Scraping (extract contact info from company websites)",
            variable=self.enable_website_scraping
        )
        website_scraping_check.grid(row=1, column=0, sticky=tk.W, padx=5, pady=5, columnspan=3)
        
        # Similarity threshold
        ttk.Label(settings_frame, text="Similarity Threshold:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        threshold_scale = ttk.Scale(
            settings_frame,
            from_=0.0,
            to=1.0,
            orient=tk.HORIZONTAL,
            variable=self.similarity_threshold,
            length=300
        )
        threshold_scale.grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)
        threshold_label = ttk.Label(settings_frame, textvariable=tk.StringVar(
            value=lambda: f"{self.similarity_threshold.get():.2f}"
        ))
        threshold_label.grid(row=2, column=2, padx=5, pady=5)
        
        # Update the threshold label when the scale changes
        def update_threshold_label(*args):
            threshold_label.config(text=f"{self.similarity_threshold.get():.2f}")
        
        self.similarity_threshold.trace_add("write", update_threshold_label)
        
        # Result frame
        result_frame = ttk.LabelFrame(main_frame, text="Results", padding=10)
        result_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create treeview for results
        columns = ("field", "value")
        self.tree = ttk.Treeview(result_frame, columns=columns, show="headings")
        self.tree.heading("field", text="Field")
        self.tree.heading("value", text="Value")
        self.tree.column("field", width=150)
        self.tree.column("value", width=500)
        
        # Add a vertical scrollbar
        scrollbar = ttk.Scrollbar(result_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # Pack the treeview and scrollbar
        scrollbar.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)
        
        # Log frame
        log_frame = ttk.LabelFrame(main_frame, text="Log", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create log text widget
        self.log_text = scrolledtext.ScrolledText(log_frame, height=10)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Status bar
        status_frame = ttk.Frame(self.root)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM)
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(status_frame, textvariable=self.status_var).pack(side=tk.LEFT, padx=5, pady=2)
        
        # Redirect stdout to our log text widget
        sys.stdout = self
        
    def write(self, text):
        """Handle stdout redirection"""
        if self.log_text:
            self.log_text.insert(tk.END, text)
            self.log_text.see(tk.END)
    
    def flush(self):
        """Required for stdout redirection"""
        pass
    
    def search_company(self):
        """Search for a company and display the results"""
        company_name = self.company_entry.get().strip()
        if not company_name:
            messagebox.showerror("Error", "Please enter a company name")
            return
        
        # Clear previous results
        self.clear_results()
        
        # Update status
        self.status_var.set("Searching...")
        self.is_scraping = True
        
        # Initialize scraper with current settings
        try:
            if self.scraper:
                self.scraper.close()
                
            self.scraper = GoogleMapsScraper(
                enable_similar_search=self.enable_similar_search.get(),
                similarity_threshold=self.similarity_threshold.get(),
                enable_website_scraping=self.enable_website_scraping.get()
            )
            
            print(f"Smart search enabled: {self.enable_similar_search.get()}, " +
                  f"Similarity threshold: {self.similarity_threshold.get():.2f}, " +
                  f"Website scraping: {self.enable_website_scraping.get()}")
                  
        except Exception as e:
            self.log_error(f"Error initializing scraper: {str(e)}")
            self.status_var.set("Ready")
            self.is_scraping = False
            return
        
        # Run the search in a separate thread
        self.scraping_thread = Thread(target=self._search_thread, args=(company_name,), daemon=True)
        self.scraping_thread.start()
    
    def _search_thread(self, company_name):
        """Run the search in a separate thread"""
        try:
            # Search for the company
            company_data = self.scraper.search_company(company_name)
            
            # Display the results
            self.root.after(0, lambda: self.display_results(company_data, company_name))
            
        except Exception as e:
            self.log_error(f"Error searching for company: {str(e)}")
            traceback.print_exc(file=sys.stdout)
        finally:
            self.root.after(0, lambda: self.status_var.set("Ready"))
            self.is_scraping = False
    
    def display_results(self, data, original_name):
        """Display the search results in the treeview"""
        if not data:
            self.tree.insert("", "end", values=("No data found", ""))
            return
        
        # Special handling for matches vs. original names
        if data.get("original_name") and data.get("mapped_name"):
            self.tree.insert("", "end", values=("Original Company", data.get("original_name")))
            self.tree.insert("", "end", values=("Matched Company", data.get("mapped_name")))
            
            # Show similarity score if available
            if 'similarity_score' in data:
                score = f"{data.get('similarity_score', 0):.2f}"
                self.tree.insert("", "end", values=("Similarity Score", score))
                
            self.tree.insert("", "end", values=("Search Variation", data.get("search_variation", "Unknown")))
            # Add a separator
            self.tree.insert("", "end", values=("-"*20, "-"*20))
        
        # Handle similar company found but not used case
        elif data.get('similar_company_found') and not data.get('similar_company_used', True):
            self.tree.insert("", "end", values=("Similar Company Found", data.get("similar_company_found")))
            self.tree.insert("", "end", values=("Similar Company Phone", data.get("similar_company_phone")))
            self.tree.insert("", "end", values=("Similarity Score", f"{data.get('similarity_score', 0):.2f}"))
            self.tree.insert("", "end", values=("Note", "Similarity score too low to use this match"))
            # Add a separator
            self.tree.insert("", "end", values=("-"*20, "-"*20))
        
        # Highlight if this is updated data from Google Maps
        if data.get("is_updated"):
            self.tree.insert("", "end", values=("STATUS", "✓ UPDATED FROM GOOGLE MAPS"))
            
        # Display original query
        if data.get("original_query") and data.get("original_query") != data.get("name"):
            self.tree.insert("", "end", values=("Original Query", data.get("original_query")))
            
        # Display the standard fields
        for field in ["name", "address", "website", "email", "rating", "reviews_count", "category"]:
            value = data.get(field, "")
            if value:
                field_name = field.capitalize()
                self.tree.insert("", "end", values=(field_name, value))
                
                # Show email source if available
                if field == "email" and data.get('email_source'):
                    self.tree.insert("", "end", values=("Email Source", data.get('email_source')))
        
        # Display phone numbers - improved handling
        if 'phones' in data and data['phones'] and len(data['phones']) > 0:
            # For multiple phone numbers, display each one on a separate row
            self.tree.insert("", "end", values=("Phone Numbers", f"{len(data['phones'])} numbers found"))
            
            for i, phone in enumerate(data['phones']):
                self.tree.insert("", "end", values=(f"   Phone {i+1}", phone))
                
            # Add source info if available
            if data.get('phone_source'):
                self.tree.insert("", "end", values=("Phone Source", data.get('phone_source')))
        elif data.get('phone'):
            # Display single phone number
            self.tree.insert("", "end", values=("Phone", data.get('phone')))
            # Add source info if available
            if data.get('phone_source'):
                self.tree.insert("", "end", values=("Phone Source", data.get('phone_source')))
        
        # Display coordinates if available
        if data.get("latitude") and data.get("longitude"):
            coords = f"{data.get('latitude')}, {data.get('longitude')}"
            self.tree.insert("", "end", values=("Coordinates", coords))
            
            # Add Google Maps link
            maps_url = f"https://www.google.com/maps/place/{coords}"
            self.tree.insert("", "end", values=("Google Maps URL", maps_url))
    
    def clear_results(self):
        """Clear the results treeview"""
        for item in self.tree.get_children():
            self.tree.delete(item)
    
    def log_error(self, message):
        """Log an error message"""
        self.log_text.insert(tk.END, f"ERROR: {message}\n")
        self.log_text.see(tk.END)
    
    def on_closing(self):
        """Handle window closing"""
        if self.is_scraping:
            if messagebox.askyesno("Quit", "A search is in progress. Are you sure you want to quit?"):
                self.close_scraper()
                self.root.destroy()
        else:
            self.close_scraper()
            self.root.destroy()
    
    def close_scraper(self):
        """Close the scraper when the application is closed"""
        if self.scraper:
            try:
                self.scraper.close()
            except:
                pass

if __name__ == "__main__":
    root = tk.Tk()
    app = GoogleMapsScraperGUI(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop() 