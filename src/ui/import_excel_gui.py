import os
import sys
import time
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from threading import Thread
import io
import queue
import datetime
from collections import deque

# Add parent directory to path for imports
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
root_dir = os.path.dirname(parent_dir)
if root_dir not in sys.path:
    sys.path.append(root_dir)

# Try different import approaches
try:
    from src.core.maps_scraper import GoogleMapsScraper
except ImportError:
    try:
        import sys
        sys.path.append(os.path.join(os.path.dirname(__file__), '..', '..'))
        from src.core.maps_scraper import GoogleMapsScraper
    except ImportError:
        try:
            sys.path.append(os.path.join(os.path.dirname(__file__), '..'))
            from core.maps_scraper import GoogleMapsScraper
        except ImportError:
            # Last resort - try direct import (old path)
            from maps_scraper import GoogleMapsScraper

# Custom text redirector for logging in GUI
class TextRedirector(io.TextIOBase):
    def __init__(self, text_widget, max_lines=1000):
        self.text_widget = text_widget
        self.buffer = ""
        self.max_lines = max_lines
        self.line_count = 0
        self.line_buffer = deque(maxlen=max_lines)
        self.lock = True  # Lock for thread safety
        
    def write(self, string):
        try:
            if not string:  # Ignore empty strings
                return 0
                
            self.buffer += string
            if '\n' in self.buffer:
                lines = self.buffer.split('\n')
                self.buffer = lines[-1]  # Keep incomplete line
                
                # Create a temporary list to hold new lines
                new_lines = []
                for line in lines[:-1]:
                    if line:  # Skip empty lines
                        new_lines.append(line)
                
                # Add all lines at once to avoid mutation during iteration
                if new_lines:
                    # Lock updates to avoid concurrent modification
                    self.lock = True
                    for line in new_lines:
                        self.line_buffer.append(line)
                    
                    # Update text widget safely
                    try:
                        self.text_widget.delete(1.0, tk.END)
                        for line in list(self.line_buffer):  # Create a copy of the deque
                            self.text_widget.insert(tk.END, line + '\n')
                        self.text_widget.see(tk.END)
                    except Exception as e:
                        print(f"Error updating text widget: {str(e)}", file=sys.__stdout__)
                    finally:
                        self.lock = False
            
            return len(string)
        except Exception as e:
            print(f"Error in TextRedirector.write: {str(e)}", file=sys.__stdout__)
            return 0
    
    def flush(self):
        pass

class ExcelProcessorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Idle Scrape - Google Maps Scraper")
        self.root.geometry("800x830")  # Increased window height for footer
        
        self.scraper = None
        self.data = None
        self.scraping_thread = None
        self.is_scraping = False
        self.stop_requested = False
        self.queue = queue.Queue()
        
        # Smart search settings
        self.enable_similar_search = tk.BooleanVar(value=False)
        self.similarity_threshold = tk.DoubleVar(value=0.7)
        self.enable_website_scraping = tk.BooleanVar(value=False)
        
        self.setup_ui()
        
        # Redirect stdout to our log widget
        self.old_stdout = sys.stdout
        sys.stdout = TextRedirector(self.log_text)
        
        # Set up periodic queue check for log updates
        self.check_queue()
    
    def setup_ui(self):
        # Main frame - Using grid instead of pack
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # File selection frame
        file_frame = ttk.LabelFrame(main_frame, text="Excel File Selection", padding=10)
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
        
        # Search Settings frame
        settings_frame = ttk.LabelFrame(main_frame, text="Search Settings", padding=10)
        settings_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Smart search checkbox
        self.smart_search_check = ttk.Checkbutton(
            settings_frame, 
            text="Enable Smart Search (find similar companies)",
            variable=self.enable_similar_search
        )
        self.smart_search_check.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5, columnspan=3)
        
        # Website scraping checkbox
        self.website_scraping_check = ttk.Checkbutton(
            settings_frame, 
            text="Enable Website Scraping (extract contact info from company websites)",
            variable=self.enable_website_scraping
        )
        self.website_scraping_check.grid(row=1, column=0, sticky=tk.W, padx=5, pady=5, columnspan=3)
        
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
        self.threshold_label = ttk.Label(settings_frame, text=f"{self.similarity_threshold.get():.2f}")
        self.threshold_label.grid(row=2, column=2, padx=5, pady=5)
        
        # Update the threshold label when the scale changes
        self.similarity_threshold.trace_add("write", self.update_threshold_label)
        
        # Control buttons frame - Now after search settings
        control_frame = ttk.LabelFrame(main_frame, text="Control Buttons", padding=10)
        control_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Create button layout using grid with fixed widths
        self.load_btn = ttk.Button(control_frame, text="Load Excel", width=15, command=self.load_excel)
        self.load_btn.grid(row=0, column=0, padx=10, pady=10)
        
        self.process_btn = ttk.Button(control_frame, text="Process Data", width=15, 
                                     command=self.process_data, state=tk.DISABLED)
        self.process_btn.grid(row=0, column=1, padx=10, pady=10)
        
        self.stop_btn = ttk.Button(control_frame, text="Stop & Save", width=15,
                                  command=self.stop_processing, state=tk.DISABLED)
        self.stop_btn.grid(row=0, column=2, padx=10, pady=10)
        
        self.save_btn = ttk.Button(control_frame, text="Save Results", width=15,
                                  command=self.save_results, state=tk.DISABLED)
        self.save_btn.grid(row=0, column=3, padx=10, pady=10)
        
        # Configure columns to distribute space evenly
        for i in range(4):
            control_frame.columnconfigure(i, weight=1)
        
        # Log frame
        log_frame = ttk.LabelFrame(main_frame, text="Log Terminal", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create log text widget with scrollbar
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=10, bg="black", fg="lime")
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Log controls
        log_controls = ttk.Frame(log_frame)
        log_controls.pack(fill=tk.X, pady=5)
        self.clear_log_btn = ttk.Button(log_controls, text="Clear Log", width=15, command=self.clear_log)
        self.clear_log_btn.pack(side=tk.RIGHT, padx=5)
        
        # Preview frame
        preview_frame = ttk.LabelFrame(main_frame, text="Data Preview", padding=10)
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create Treeview for data preview
        self.tree = ttk.Treeview(preview_frame)
        
        # Add a scrollbar
        scrollbar_y = ttk.Scrollbar(preview_frame, orient="vertical", command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(preview_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        # Pack the treeview and scrollbars
        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")
        self.tree.pack(side="left", fill="both", expand=True)
        
        # Progress frame
        progress_frame = ttk.LabelFrame(main_frame, text="Processing", padding=10)
        progress_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", length=100, mode="determinate", variable=self.progress_var)
        self.progress_bar.pack(padx=5, pady=5, fill=tk.X)
        
        # Status label
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(progress_frame, textvariable=self.status_var).pack(padx=5, pady=5, anchor=tk.W)
        
        # Version footer
        footer_frame = ttk.Frame(main_frame)
        footer_frame.pack(fill=tk.X, padx=5, pady=(10, 5))
        version_label = ttk.Label(footer_frame, text="Version 1.0", font=("Arial", 8))
        version_label.pack(side=tk.RIGHT, padx=5)
    
    def log(self, message):
        """Log a message to the queue for thread-safe logging"""
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        self.queue.put(f"[{timestamp}] {message}")
    
    def check_queue(self):
        """Check the queue for new log messages"""
        try:
            while True:
                message = self.queue.get_nowait()
                print(message)
                self.queue.task_done()
        except queue.Empty:
            pass
        finally:
            # Schedule the next queue check
            self.root.after(100, self.check_queue)
    
    def update_threshold_label(self, *args):
        """Update the threshold label when the slider changes"""
        self.threshold_label.config(text=f"{self.similarity_threshold.get():.2f}")
    
    def clear_log(self):
        """Clear the log text widget"""
        self.log_text.delete(1.0, tk.END)
        self.log("Log cleared")
    
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
                self.output_file_var.set(f"{base}_processed{ext}")
            self.log(f"Selected input file: {filename}")
    
    def browse_output_file(self):
        filename = filedialog.asksaveasfilename(
            title="Save As",
            defaultextension=".xlsx",
            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
        )
        if filename:
            self.output_file_var.set(filename)
            self.log(f"Selected output file: {filename}")
    
    def load_excel(self):
        input_file = self.input_file_var.get().strip()
        
        if not input_file:
            messagebox.showerror("Error", "Please select an input Excel file")
            return
        
        try:
            # Read Excel file
            self.log(f"Loading Excel file: {input_file}")
            df = pd.read_excel(input_file)
            
            # Define possible column name formats
            column_mappings = {
                'no': ['No.', 'No', 'NO', 'no.', 'no', 'Nomor', 'nomor', '#'],
                'company': ['Nama Perusahaan', 'Nama_Perusahaan', 'NamaPerusahaan', 'Nama', 'Company', 'company', 'NAMA'],
                'address': ['Alamat', 'ALAMAT', 'alamat', 'Address', 'address'],
                'district': ['Kecamatan', 'KECAMATAN', 'kecamatan', 'District', 'district']
            }
            
            # Map actual column names to standard names
            column_map = {}
            missing_columns = []
            
            for standard_key, possible_names in column_mappings.items():
                found = False
                for name in possible_names:
                    if name in df.columns:
                        column_map[standard_key] = name
                        found = True
                        break
                
                if not found:
                    missing_columns.append(standard_key)
            
            self.log(f"Found column mapping: {column_map}")
            
            if missing_columns:
                messagebox.showerror("Error", 
                                   f"Could not find columns for: {', '.join(missing_columns)}\n\n"
                                   "Make sure your Excel file has columns for:\n"
                                   "Number, Company Name, Address, District\n\n"
                                   f"Available columns: {', '.join(df.columns)}")
                return
            
            # Rename columns to standard format if needed
            self.data = df.rename(columns={
                column_map['no']: 'No.',
                column_map['company']: 'Nama Perusahaan',
                column_map['address']: 'Alamat',
                column_map['district']: 'Kecamatan'
            })
            
            # Update the treeview
            self.update_tree(self.data)
            
            # Enable process button
            self.process_btn.config(state=tk.NORMAL)
            
            # Update status
            self.status_var.set(f"Loaded {len(df)} companies from {input_file}")
            self.log(f"Successfully loaded {len(df)} companies")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error loading Excel file: {str(e)}")
            self.log(f"Error loading Excel file: {str(e)}")
    
    def update_tree(self, df):
        # Clear existing tree
        for i in self.tree.get_children():
            self.tree.delete(i)
        
        # Configure columns
        self.tree["columns"] = list(df.columns)
        self.tree["show"] = "headings"
        
        # Set column headings
        for col in df.columns:
            self.tree.heading(col, text=col)
            # Set column width based on content
            max_width = max(len(str(df[col].iloc[i])) for i in range(min(10, len(df)))) * 10
            max_width = max(max_width, len(col) * 10)
            max_width = min(max_width, 300)  # Cap width
            self.tree.column(col, width=max_width)
        
        # Add data to the treeview
        for i, row in df.head(50).iterrows():  # Display only first 50 rows
            values = [row[col] for col in df.columns]
            self.tree.insert("", "end", values=values)
        
        if len(df) > 50:
            self.tree.insert("", "end", values=["..."] * len(df.columns))
    
    def process_data(self):
        if self.data is None or len(self.data) == 0:
            messagebox.showerror("Error", "No data to process. Please load an Excel file first.")
            return
        
        output_file = self.output_file_var.get().strip()
        if not output_file:
            messagebox.showerror("Error", "Please specify an output Excel file")
            return
        
        # Reset stop flag
        self.stop_requested = False
        
        # Disable buttons during processing
        self.load_btn.config(state=tk.DISABLED)
        self.process_btn.config(state=tk.DISABLED)
        self.save_btn.config(state=tk.DISABLED)
        
        # Enable stop button
        self.stop_btn.config(state=tk.NORMAL)
        
        # Reset progress
        self.progress_var.set(0)
        
        # Start processing in a separate thread
        self.scraping_thread = Thread(target=self._process_data_thread)
        self.scraping_thread.daemon = True
        self.scraping_thread.start()
    
    def stop_processing(self):
        """Stop the processing and save current results"""
        if not self.is_scraping:
            return
            
        self.log("Stop requested. Waiting for current company to finish...")
        self.status_var.set("Stopping - please wait for current task to finish")
        self.stop_requested = True
        self.stop_btn.config(state=tk.DISABLED)
    
    def _process_data_thread(self):
        try:
            self.is_scraping = True
            
            # Initialize scraper
            if not self.scraper:
                self.log("Initializing Google Maps scraper...")
                try:
                    self.scraper = GoogleMapsScraper(
                        enable_similar_search=self.enable_similar_search.get(),
                        similarity_threshold=self.similarity_threshold.get(),
                        enable_website_scraping=self.enable_website_scraping.get()
                    )
                    self.log(f"Smart search enabled: {self.enable_similar_search.get()}, " +
                           f"Similarity threshold: {self.similarity_threshold.get():.2f}, " +
                           f"Website scraping: {self.enable_website_scraping.get()}")
                except Exception as e:
                    self.log(f"Error initializing scraper: {str(e)}")
                    error_msg = f"Failed to initialize Google Maps scraper: {str(e)}"
                    self.root.after(0, lambda e=error_msg: messagebox.showerror("Scraper Error", e))
                    self.root.after(0, lambda: self.status_var.set(f"Error: Failed to initialize scraper"))
                    self.root.after(0, lambda: self.load_btn.config(state=tk.NORMAL))
                    self.root.after(0, lambda: self.process_btn.config(state=tk.NORMAL))
                    self.root.after(0, lambda: self.stop_btn.config(state=tk.DISABLED))
                    self.is_scraping = False
                    return
            
            # Add columns for scraped data if they don't exist
            for col in ['Phone', 'Website', 'Email', 'Rating', 'Reviews Count', 'Category', 
                       'Updated Address', 'Updated Company Name', 'Similar Company Found',
                       'Similarity Score', 'Mapped Company', 'Phone Source', 'Email Source']:
                if col not in self.data.columns:
                    self.data[col] = None
            
            total = len(self.data)
            processed_count = 0
            
            # Process each company
            for i, row in self.data.iterrows():
                # Check if stop was requested
                if self.stop_requested:
                    self.log("Processing stopped by user")
                    break
                
                # Update progress and status
                progress_pct = (i / total) * 100
                self.root.after(0, lambda p=progress_pct: self.progress_var.set(p))
                
                company_name = row['Nama Perusahaan']
                address = row['Alamat']
                district = row['Kecamatan']
                
                self.root.after(0, lambda cn=company_name, i=i, t=total: 
                              self.status_var.set(f"Processing {i+1}/{t}: {cn}"))
                
                self.log(f"Processing {i+1}/{total}: {company_name}")
                
                # Create search query (company name + address for better results)
                search_query = f"{company_name} {address} {district}"
                
                # Search for company data with error handling
                try:
                    company_data = self.scraper.search_company(search_query)
                    
                    if company_data:
                        # Preprocessing nomor telepon agar lengkap dan benar formatnya
                        phone_number = company_data.get('phone')
                        if phone_number:
                            # Debug informasi nomor telepon yang diperoleh
                            self.log(f"Raw phone number: {phone_number}")
                            
                            # Pastikan format nomor telepon lengkap
                            # Bersihkan karakter yang tidak perlu
                            clean_phone = ''.join(c for c in phone_number if c.isdigit() or c in '+-() ')
                            self.log(f"Cleaned phone: {clean_phone}")
                            
                            # Pastikan nomor telepon tidak hanya bagian kode area
                            # Misalnya, jika hanya "06", jangan gunakan
                            if len(clean_phone.strip()) < 6:
                                self.log(f"Phone number too short, ignoring: {clean_phone}")
                                phone_number = None
                            else:
                                # Gunakan nomor telepon yang sudah dibersihkan
                                phone_number = clean_phone
                                self.log(f"Final phone: {phone_number}")
                        
                        # Update data with scraped information
                        self.data.at[i, 'Phone'] = phone_number
                        self.data.at[i, 'Website'] = company_data.get('website')
                        self.data.at[i, 'Rating'] = company_data.get('rating')
                        self.data.at[i, 'Reviews Count'] = company_data.get('reviews_count')
                        self.data.at[i, 'Category'] = company_data.get('category')
                        
                        # Record phone and email sources if available
                        if company_data.get('phone_source'):
                            self.data.at[i, 'Phone Source'] = company_data.get('phone_source')
                            self.log(f"Phone source: {company_data.get('phone_source')}")
                            
                        if company_data.get('email_source'):
                            self.data.at[i, 'Email Source'] = company_data.get('email_source')
                            self.log(f"Email source: {company_data.get('email_source')}")
                        
                        # Store updated addresses and company names if available
                        if company_data.get('address') and company_data.get('address') != address:
                            self.data.at[i, 'Updated Address'] = company_data.get('address')
                            self.log(f"✓ Updated address found: {company_data.get('address')}")
                        
                        if company_data.get('name') and company_data.get('name') != company_name:
                            self.data.at[i, 'Updated Company Name'] = company_data.get('name')
                            self.log(f"✓ Updated company name found: {company_data.get('name')}")
                        
                        # Store similar company information if available
                        if self.enable_similar_search.get():
                            if company_data.get('mapped_name'):
                                if company_data.get('similarity_score', 0) >= self.similarity_threshold.get():
                                    self.data.at[i, 'Mapped Company'] = company_data.get('mapped_name')
                                    self.data.at[i, 'Similarity Score'] = company_data.get('similarity_score')
                                    self.log(f"✓ Name mapped: {company_name} → {company_data.get('mapped_name')} " +
                                           f"(Score: {company_data.get('similarity_score', 0):.2f})")
                            elif company_data.get('similar_company_found'):
                                self.data.at[i, 'Similar Company Found'] = company_data.get('similar_company_found')
                                self.data.at[i, 'Similarity Score'] = company_data.get('similarity_score')
                                self.log(f"ℹ Similar company found but not used due to low similarity: " +
                                       f"{company_data.get('similar_company_found')} " +
                                       f"(Score: {company_data.get('similarity_score', 0):.2f})")
                        
                        # Try to extract email from website field if available
                        website = company_data.get('website')
                        if website and '@' in website:
                            self.data.at[i, 'Email'] = website
                        elif company_data.get('email'):
                            self.data.at[i, 'Email'] = company_data.get('email')
                        
                        self.log(f"✓ Data found for {company_name}")
                        processed_count += 1
                    else:
                        self.log(f"✗ No data found for {company_name}")
                except Exception as search_error:
                    self.log(f"Error processing {company_name}: {str(search_error)}")
                    # Continue dengan perusahaan berikutnya meskipun ada error
                
                # Autosave every 5 companies processed
                if processed_count > 0 and processed_count % 5 == 0:
                    temp_output = self.output_file_var.get().replace(".xlsx", "_autosave.xlsx")
                    try:
                        self.data.to_excel(temp_output, index=False)
                        self.log(f"Autosaved progress to {temp_output}")
                    except Exception as e:
                        self.log(f"Error autosaving: {str(e)}")
                
                # Add delay between requests
                if i < total - 1 and not self.stop_requested:
                    time.sleep(2)
            
            completed_status = "completed" if not self.stop_requested else "stopped"
            
            # Update progress to 100% if completed, or actual percentage if stopped
            if not self.stop_requested:
                self.root.after(0, lambda: self.progress_var.set(100))
            
            self.root.after(0, lambda: self.status_var.set(f"Processing {completed_status}"))
            self.log(f"Processing {completed_status}. Processed {processed_count} companies.")
            
            # Update the treeview
            self.root.after(0, lambda: self.update_tree(self.data))
            
            # Enable save button
            self.root.after(0, lambda: self.save_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.load_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.process_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.stop_btn.config(state=tk.DISABLED))
            
            # Show completion message
            if not self.stop_requested:
                self.root.after(0, lambda: messagebox.showinfo("Complete", f"Data processing complete. Processed {processed_count} companies."))
            else:
                self.root.after(0, lambda: messagebox.showinfo("Stopped", f"Processing stopped. Processed {processed_count} companies.\nYou can save the partial results."))
                
            # If stopped, offer to save immediately
            if self.stop_requested:
                self.root.after(0, lambda: self.offer_save_partial_results())
            
        except Exception as e:
            error_msg = f"Error processing data: {str(e)}"
            self.log(error_msg)
            self.root.after(0, lambda e=error_msg: messagebox.showerror("Error", e))
            self.root.after(0, lambda e=error_msg: self.status_var.set(f"Error: {str(e)[:50]}..."))
            self.root.after(0, lambda: self.load_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.process_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.stop_btn.config(state=tk.DISABLED))
        finally:
            self.is_scraping = False
            self.stop_requested = False
    
    def offer_save_partial_results(self):
        """Offer to save partial results after stopping"""
        if messagebox.askyesno("Save Results", "Do you want to save the partial results now?"):
            self.save_results()
    
    def save_results(self):
        if self.data is None:
            messagebox.showerror("Error", "No data to save. Please process data first.")
            return
        
        output_file = self.output_file_var.get().strip()
        if not output_file:
            # Ask for output file if not specified
            output_file = filedialog.asksaveasfilename(
                title="Save As",
                defaultextension=".xlsx",
                filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
            )
            if not output_file:
                return
            self.output_file_var.set(output_file)
        
        try:
            self.log(f"Saving data to {output_file}...")
            self.data.to_excel(output_file, index=False)
            self.log(f"Data saved successfully to {output_file}")
            messagebox.showinfo("Success", f"Data saved successfully to {output_file}")
        except Exception as e:
            error_msg = f"Error saving file: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("Error", error_msg)
    
    def on_closing(self):
        if self.is_scraping:
            if messagebox.askyesno("Confirm Exit", "A scraping task is in progress. Are you sure you want to exit?"):
                self.close_scraper()
                # Restore stdout
                sys.stdout = self.old_stdout
                self.root.destroy()
        else:
            self.close_scraper()
            # Restore stdout
            sys.stdout = self.old_stdout
            self.root.destroy()
    
    def close_scraper(self):
        if self.scraper:
            try:
                self.scraper.close()
                self.log("WebDriver closed successfully")
            except Exception as e:
                self.log(f"Error closing WebDriver: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelProcessorGUI(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop() 