import os
import sys
import time
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from threading import Thread, Timer
import io
import queue
import datetime
from collections import deque
import traceback

# Import modul notifikasi desktop
try:
    from notifikasi_helper import show_desktop_notification, NOTIFICATIONS_AVAILABLE
except ImportError:
    NOTIFICATIONS_AVAILABLE = False
    print("Modul notifikasi_helper.py tidak ditemukan. Notifikasi desktop tidak akan tersedia.")

# Add src directory to path if needed
if os.path.exists('src'):
    sys.path.append(os.path.abspath('src'))
    sys.path.append(os.path.abspath('src/core'))

try:
    from src.core.maps_scraper import GoogleMapsScraper
except ImportError:
    try:
        from maps_scraper import GoogleMapsScraper
    except ImportError:
        print("ERROR: Could not import GoogleMapsScraper module.")
        print("Please make sure the src/core/maps_scraper.py file exists.")
        sys.exit(1)

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
        """Initialize the application UI."""
        self.root = root
        self.root.title("Idle Scrape - Google Maps Scraper (Beta)")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 700)
        
        # Initialize important state variables before anything else
        self.scraping_active = False
        self.scraper = None
        self.autosave_timer = None
        self.old_stdout = sys.stdout  # Save original stdout
        
    # Metode _set_initial_pane_position dihapus untuk menghindari masalah layout
        
        # State variables
        self.file_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.column_name = tk.StringVar()
        self.input_file_var = tk.StringVar()
        self.output_file_var = tk.StringVar()
        self.data = None
        self.scraped_data = []
        self.queue = queue.Queue()
        self.last_scraped_company = tk.StringVar(value="None")
        self.scraping_active = False
        self.scraper = None
        self.update_id = None
        self.df_complete = None
        self.status_text = []
        self.max_retries = tk.IntVar(value=3)
        self.retry_delay = tk.IntVar(value=5)
        self.enable_similar_search = tk.BooleanVar(value=False)
        self.similarity_threshold = tk.DoubleVar(value=0.6)
        self.enable_website_scraping = tk.BooleanVar(value=False)
        self.enable_cache = tk.BooleanVar(value=True)
        self.cache_days = tk.IntVar(value=30)
        self.preserve_phone_format = tk.BooleanVar(value=False)
        
        # Autosave variables
        self.enable_autosave = tk.BooleanVar(value=True)
        self.autosave_interval = tk.IntVar(value=10)  # Minutes
        self.autosave_timer = None
        self.last_autosave_time = None
        
        # Language selection
        self.current_language = tk.StringVar(value="English")
        
        # City prioritization
        self.priority_city = tk.StringVar(value="All")
        
        # Anti-detection variables
        self.use_rotating_user_agents = tk.BooleanVar(value=True)
        self.use_proxies = tk.BooleanVar(value=False)
        
        # Initialize translations
        self.init_translations()
        
        # Counter variables for statistics
        self.total_companies_var = tk.StringVar(value="0")
        self.processed_var = tk.StringVar(value="0 (0%)")
        self.phone_found_var = tk.StringVar(value="0 (0%)")
        self.address_found_var = tk.StringVar(value="0 (0%)")
        self.website_found_var = tk.StringVar(value="0 (0%)")
        self.email_found_var = tk.StringVar(value="0")  # Changed from "0 (0%)" to just "0"
        self.cache_hits_var = tk.StringVar(value="0 (0%)")
        
        # Additional statistics variables
        self.gmaps_count_var = tk.StringVar(value="0")
        self.website_count_var = tk.StringVar(value="0")
        self.mobile_phone_count_var = tk.StringVar(value="0")
        self.office_phone_count_var = tk.StringVar(value="0")
        
        # Apply custom styling
        self.style = self.apply_style()
        
        # Add application icon if available
        try:
            if os.path.exists('app_icon.ico'):
                self.root.iconbitmap('app_icon.ico')
            elif os.path.exists('important_files/app_icon.ico'):
                self.root.iconbitmap('important_files/app_icon.ico')
        except Exception:
            pass  # Silently fail if icon not found or not supported
        
        # Create the menu bar before setting up the main UI
        self.create_menu_bar()
        
        self.setup_ui()
        
        # Redirect stdout to our log widget (jika belum dilakukan)
        if sys.stdout == self.old_stdout:  # Hanya redirect jika belum dilakukan
            sys.stdout = TextRedirector(self.log_text)
        
        # Set up periodic queue check for log updates
        self.check_queue()
    
    def init_translations(self):
        """Initialize translation dictionaries for supported languages"""
        # English translations (default)
        self.translations = {
            "English": {
                # Menu items
                "file_menu": "File",
                "import_excel": "Import Excel...",
                "set_output_file": "Set Output File...",
                "load_data": "Load Data",
                "save_results": "Save Results",
                "exit": "Exit",
                "settings_menu": "Settings",
                "settings_item": "Settings...",
                "process_menu": "Process",
                "start_scraping": "Start Scraping",
                "stop_save": "Stop & Save",
                "clear_log": "Clear Log",
                "copy_log": "Copy Log",
                "help_menu": "Help",
                "documentation": "Documentation",
                "feature_explanations": "Feature Explanations",
                "about": "About",
                
                # Main UI elements
                "monitoring_dashboard": "Monitoring Dashboard",
                "total_companies": "Total Companies:",
                "processed": "Processed:",
                "google_maps_data": "Google Maps Data:",
                "website_data": "Website Data:",
                "mobile_phones": "Mobile Phones:",
                "office_phones": "Office Phones:",
                "current_processing": "Current Processing",
                "company": "Company:",
                "data_source": "Data Source:",
                "progress": "Progress",
                "quick_settings": "Quick Settings",
                "enable_smart_search": "Enable Smart Search",
                "enable_website_scraping": "Enable Website Scraping",
                "more_options": "More options available in the Settings menu",
                "file_selection": "File Selection",
                "excel_file": "Excel File:",
                "browse": "Browse",
                "output_file_config": "Output file can be configured in the File menu",
                "email_found": "Email Found:",
                
                # Settings dialog
                "settings_title": "Settings - Idle Scrape Google Maps Scraper",
                "scraping_tab": "Scraping",
                "anti_detection_tab": "Anti-Detection",
                "cache_tab": "Cache",
                "advanced_tab": "Advanced",
                "interface_tab": "Interface",
                "smart_search_settings": "Smart Search Settings",
                "enable_smart_search_long": "Enable Smart Search (find similar companies)",
                "similarity_threshold": "Similarity Threshold:",
                "website_scraping": "Website Scraping",
                "enable_website_scraping_long": "Enable Website Scraping (extract contact info from company websites)",
                "preserve_phone_format": "Preserve original phone number format",
                "location_priority": "Location Priority",
                "prioritize_contact": "Prioritize Contact Info from:",
                "phone_priority_help": "Phone numbers from selected city will be displayed first",
                "user_agent_settings": "User-Agent Settings",
                "use_rotating_agents": "Use rotating User-Agents (helps avoid detection)",
                "rotating_agents_help": "Automatically rotates between different browser identities",
                "proxy_settings": "Proxy Settings (Experimental)",
                "use_proxy_rotation": "Use free proxy rotation (may affect stability)",
                "proxy_warning": "Warning: Using proxies may significantly slow down scraping and reduce reliability",
                "timing_settings": "Timing Settings",
                "retry_delay": "Retry Delay (seconds):",
                "cache_settings": "Cache Settings",
                "enable_caching": "Enable result caching (speeds up repeat searches)",
                "cache_retention": "Cache retention period:",
                "days": "days",
                "clear_cache": "Clear Cache",
                "retry_settings": "Retry Settings",
                "max_retries": "Maximum Retries:",
                "retry_help": "Higher values improve reliability but may increase processing time",
                "advanced_warning": "Note: The settings on this tab are for advanced users.\n\nThe default values should work well for most users.\nChange these settings only if you understand their effects.",
                "language_settings": "Language Settings",
                "interface_language": "Interface Language:",
                "language_note": "Changes will be applied after restarting the application",
                "ok": "OK",
                "cancel": "Cancel",
                "apply": "Apply",
                "settings_applied": "Settings Applied",
                "settings_applied_msg": "Your settings have been applied.",
                "cache_cleared": "Cache Cleared",
                "cache_cleared_msg": "The cache has been successfully cleared.",
                
                # Documentation window
                "doc_title": "Documentation - Idle Scrape Google Maps Scraper",
                "doc_content": """# Idle Scrape Google Maps Scraper v2 Documentation

## Overview
Idle Scrape is a powerful tool for extracting business contact information from Google Maps. 
It supports smart searching, website scraping, and anti-detection measures.

## Key Features
- Google Maps business data extraction
- Smart searching with similarity matching
- Website scraping for additional contact details
- Anti-detection features including user agent rotation
- Results caching for improved performance
- City-based phone number prioritization

## Getting Started
1. Import your Excel file containing company names
2. Set your desired scraping options in the Settings menu
3. Start the scraping process
4. Save results when complete

## Tips for Best Results
- Include city or region information with company names for better matching
- Use the smart search feature for companies with common spelling variations
- Enable website scraping to extract additional contact information
- Adjust similarity threshold based on your needs (higher = more exact matches)

## Troubleshooting
If you encounter issues:
- Try clearing the cache
- Ensure you have a stable internet connection
- Check your Excel file format and company name spelling
- Increase retry attempts for more reliable results

For more help, contact support at support@idle-scrape.com
""",
                "close": "Close",
                
                # Feature explanations window
                "feature_explanation_title": "Feature Explanations - Idle Scrape Google Maps Scraper",
                
                # Search Features
                "search_features_tab": "Search Features",
                "search_features_content": """# Search Features

## Smart Search
Enables intelligent matching of company names using similarity algorithms. This helps find 
companies even when the exact name in your data doesn't match Google's records.

- When enabled: The system will find companies with similar names (e.g., "ABC Company" might 
  match "ABC Co Ltd" or "A.B.C. Company")
- When disabled: Only exact matches will be returned

## Similarity Threshold
Controls how strict the name matching is for Smart Search:

- Higher values (closer to 1.0): Require closer matches, reducing false positives but potentially 
  missing valid matches
- Lower values (closer to 0.1): Allow more flexible matching, finding more potential matches but 
  with higher chance of incorrect matches
- Recommended range: 0.6-0.8 for most business names

## Location Priority
When multiple phone numbers are found, this setting determines which geographic area's phone 
numbers appear first in results:

- "All": No specific priority, returns numbers in order found
- Specific City (Jakarta, Bandung, etc.): Prioritizes phone numbers from the selected city
- Especially useful for businesses with multiple locations
""",
                
                # Contact Features
                "contact_features_tab": "Contact Features",
                "contact_features_content": """# Contact Information Features

## Website Scraping
Determines whether the system extracts additional contact information from company websites:

- When enabled: The system visits each company's website and scans for phone numbers, email 
  addresses and additional contact information not found on Google Maps
- When disabled: Only information from Google Maps is collected
- Note: Enabling this significantly increases processing time but often yields more contact details

## Preserve Original Phone Number Format
Controls how phone numbers are stored in results:

- When enabled: Phone numbers are kept exactly as they appear on source websites (e.g., 
  "(021) 555-1234" or "+62 812 3456 7890")
- When disabled: Phone numbers are standardized to a consistent format
- Enable this when the exact appearance of phone numbers is important (e.g., for marketing materials)

## Phone Type Detection
The system automatically categorizes found phone numbers into:

- Mobile Phones: Detected by prefixes (08, +628, etc.) and other patterns
- Office Phones: Detected by area codes (021, 022, etc.) and other landline patterns
- These are stored in separate columns for easier targeting in your outreach campaigns
""",
                
                # Anti-Detection Features
                "anti_detection_tab": "Anti-Detection",
                "anti_detection_content": """# Anti-Detection Features

## Rotating User-Agents
Helps avoid being blocked by Google when making many searches:

- When enabled: Each search appears to come from a different browser/device
- When disabled: All searches use the same browser identity
- Recommended: Keep enabled for datasets with more than 20 companies

## Proxy Rotation (Experimental)
Uses different internet connections to avoid IP-based rate limiting:

- When enabled: Searches are distributed across different IP addresses
- When disabled: All searches use your direct internet connection
- Note: This feature may reduce reliability and slow down processing

## Retry Settings
Controls how the system handles failed searches:

- Maximum Retries: Number of additional attempts if a search fails
- Retry Delay: Waiting time between retries (in seconds)
- Higher values improve success rates but increase total processing time
""",
                
                # Cache & Performance Features
                "cache_performance_tab": "Cache & Performance",
                "cache_performance_content": """# Cache & Performance Features

## Result Caching
Stores search results locally to improve performance on repeated searches:

- When enabled: Previously searched companies are retrieved instantly from local storage
- When disabled: Every search queries Google Maps directly
- Benefits: Faster processing, reduced Google Maps API usage, fewer detection issues

## Cache Retention Period
Controls how long cached results are considered valid:

- Range: 1-365 days
- Shorter periods ensure more up-to-date information but require more fresh searches
- Longer periods maximize performance but may return outdated information
- Recommended: 30 days for most business uses

## Clear Cache
Removes all stored search results:

- Use when you want to force fresh data for all companies
- Helpful after major Google Maps updates or when you suspect cached data is outdated
- Note: This will increase processing time for your next scrape operation

## Autosave Feature
Automatically saves your results during the scraping process:

- Default interval: 10 minutes (configurable in settings)
- Creates backup files with timestamps to prevent data loss
- Displays desktop notifications when autosaves complete
- Helpful to recover data if the application or system crashes unexpectedly
""",
                
                # About dialog
                "about_title": "About - Idle Scrape Google Maps Scraper",
                "werkudara_scraper": "Idle Scrape Google Maps Scraper",
                "version": "Version 2.5",
                "copyright": "© 2025 Idle Scrape Team",
                "department": "Development Team",
                "all_rights_reserved": "All rights reserved.",
                
                # Autosave related
                "autosave_settings": "Autosave Settings",
                "enable_autosave": "Enable Autosave",
                "autosave_interval": "Autosave interval:",
                "minutes": "minutes",
                "autosave_interval_help": "Automatically save results during scraping to prevent data loss",
                "autosave_success": "Autosave Successful",
                "autosave_success_msg": "Results have been automatically saved to",
                "autosave_feature": "Autosave Results",
                "autosave_feature_desc": "Automatically save results periodically during the process",
                "autosave_feature_help": "Prevents data loss if the application crashes",
            },
            "Indonesia": {
                # Menu items
                "file_menu": "Berkas",
                "import_excel": "Impor Excel...",
                "set_output_file": "Atur File Keluaran...",
                "load_data": "Muat Data",
                "save_results": "Simpan Hasil",
                "exit": "Keluar",
                "settings_menu": "Pengaturan",
                "settings_item": "Pengaturan...",
                "process_menu": "Proses",
                "start_scraping": "Mulai Scraping",
                "stop_save": "Berhenti & Simpan",
                "clear_log": "Bersihkan Log",
                "copy_log": "Salin Log",
                "help_menu": "Bantuan",
                "documentation": "Dokumentasi",
                "feature_explanations": "Penjelasan Fitur",
                "about": "Tentang",
                
                # Main UI elements
                "monitoring_dashboard": "Dasbor Pemantauan",
                "total_companies": "Total Perusahaan:",
                "processed": "Diproses:",
                "google_maps_data": "Data Google Maps:",
                "website_data": "Data Website:",
                "mobile_phones": "Telepon Seluler:",
                "office_phones": "Telepon Kantor:",
                "current_processing": "Sedang Diproses",
                "company": "Perusahaan:",
                "data_source": "Sumber Data:",
                "progress": "Progres",
                "quick_settings": "Pengaturan Cepat",
                "enable_smart_search": "Aktifkan Pencarian Cerdas",
                "enable_website_scraping": "Aktifkan Pencarian Website",
                "more_options": "Opsi lainnya tersedia di menu Pengaturan",
                "file_selection": "Pilihan Berkas",
                "excel_file": "Berkas Excel:",
                "browse": "Telusuri",
                "output_file_config": "Berkas keluaran dapat dikonfigurasi di menu Berkas",
                "email_found": "Email Ditemukan:",
                
                # Settings dialog
                "settings_title": "Pengaturan - Idle Scrape Google Maps Scraper",
                "scraping_tab": "Pencarian",
                "anti_detection_tab": "Anti-Deteksi",
                "cache_tab": "Cache",
                "advanced_tab": "Lanjutan",
                "interface_tab": "Antarmuka",
                "smart_search_settings": "Pengaturan Pencarian Cerdas",
                "enable_smart_search_long": "Aktifkan Pencarian Cerdas (temukan perusahaan serupa)",
                "similarity_threshold": "Ambang Kemiripan:",
                "website_scraping": "Pencarian Website",
                "enable_website_scraping_long": "Aktifkan Pencarian Website (ekstrak info kontak dari website perusahaan)",
                "preserve_phone_format": "Pertahankan format asli nomor telepon",
                "location_priority": "Prioritas Lokasi",
                "prioritize_contact": "Prioritaskan Info Kontak dari:",
                "phone_priority_help": "Nomor telepon dari kota yang dipilih akan ditampilkan lebih dulu",
                "user_agent_settings": "Pengaturan User-Agent",
                "use_rotating_agents": "Gunakan User-Agents bergantian (membantu menghindari deteksi)",
                "rotating_agents_help": "Secara otomatis berganti antara identitas browser yang berbeda",
                "proxy_settings": "Pengaturan Proxy (Eksperimental)",
                "use_proxy_rotation": "Gunakan rotasi proxy gratis (dapat mempengaruhi stabilitas)",
                "proxy_warning": "Peringatan: Menggunakan proxy dapat memperlambat pencarian dan mengurangi keandalan",
                "timing_settings": "Pengaturan Waktu",
                "retry_delay": "Jeda Coba Ulang (detik):",
                "cache_settings": "Pengaturan Cache",
                "enable_caching": "Aktifkan penyimpanan hasil (mempercepat pencarian berulang)",
                "cache_retention": "Periode penyimpanan cache:",
                "days": "hari",
                "clear_cache": "Bersihkan Cache",
                "retry_settings": "Pengaturan Coba Ulang",
                "max_retries": "Maksimum Percobaan Ulang:",
                "retry_help": "Nilai lebih tinggi meningkatkan keberhasilan tetapi dapat meningkatkan waktu pemrosesan",
                "advanced_warning": "Catatan: Pengaturan pada tab ini untuk pengguna tingkat lanjut.\n\nNilai default biasanya sudah cukup baik untuk kebanyakan pengguna.\nUbah pengaturan ini hanya jika Anda memahami efeknya.",
                "language_settings": "Pengaturan Bahasa",
                "interface_language": "Bahasa Antarmuka:",
                "language_note": "Perubahan akan diterapkan setelah memulai ulang aplikasi",
                "ok": "OK",
                "cancel": "Batal",
                "apply": "Terapkan",
                "settings_applied": "Pengaturan Diterapkan",
                "settings_applied_msg": "Pengaturan Anda telah diterapkan.",
                "cache_cleared": "Cache Dibersihkan",
                "cache_cleared_msg": "Cache telah berhasil dibersihkan.",
                
                # Documentation window
                "doc_title": "Dokumentasi - Idle Scrape Google Maps Scraper",
                "doc_content": """# Dokumentasi Idle Scrape Google Maps Scraper v2

## Ikhtisar
Idle Scrape adalah alat yang ampuh untuk mengekstrak informasi kontak bisnis dari Google Maps. 
Mendukung pencarian cerdas, pengambilan data dari website, dan fitur anti-deteksi.

## Fitur Utama
- Ekstraksi data bisnis dari Google Maps
- Pencarian cerdas dengan pencocokan kesamaan
- Pengambilan data dari website untuk detail kontak tambahan
- Fitur anti-deteksi termasuk rotasi user agent
- Penyimpanan hasil dalam cache untuk performa yang lebih baik
- Prioritas nomor telepon berdasarkan kota

## Memulai
1. Impor file Excel yang berisi nama-nama perusahaan
2. Atur opsi pencarian yang diinginkan di menu Pengaturan
3. Mulai proses pencarian
4. Simpan hasil saat selesai

## Tips untuk Hasil Terbaik
- Sertakan informasi kota atau wilayah dengan nama perusahaan untuk pencocokan yang lebih baik
- Gunakan fitur pencarian cerdas untuk perusahaan dengan variasi ejaan umum
- Aktifkan pengambilan data website untuk mendapatkan informasi kontak tambahan
- Sesuaikan ambang kesamaan berdasarkan kebutuhan Anda (lebih tinggi = kecocokan lebih tepat)

## Pemecahan Masalah
Jika Anda mengalami masalah:
- Coba bersihkan cache
- Pastikan Anda memiliki koneksi internet yang stabil
- Periksa format file Excel dan ejaan nama perusahaan
- Tingkatkan upaya percobaan ulang untuk hasil yang lebih andal

Untuk bantuan lebih lanjut, hubungi dukungan di support@idle-scrape.com
""",
                "close": "Tutup",
                
                # Feature explanations window
                "feature_explanation_title": "Penjelasan Fitur - Idle Scrape Google Maps Scraper",
                
                # Search Features
                "search_features_tab": "Fitur Pencarian",
                "search_features_content": """# Fitur Pencarian

## Pencarian Cerdas
Mengaktifkan pencocokan cerdas nama perusahaan menggunakan algoritma kesamaan. Ini membantu menemukan 
perusahaan bahkan ketika nama eksak dalam data Anda tidak cocok dengan catatan Google.

- Saat diaktifkan: Sistem akan menemukan perusahaan dengan nama serupa (misal, "ABC Company" 
  mungkin cocok dengan "ABC Co Ltd" atau "A.B.C. Company")
- Saat dinonaktifkan: Hanya kecocokan persis yang akan dikembalikan

## Ambang Kesamaan
Mengontrol seberapa ketat pencocokan nama untuk Pencarian Cerdas:

- Nilai lebih tinggi (mendekati 1.0): Memerlukan kecocokan yang lebih dekat, mengurangi positif palsu tetapi 
  berpotensi kehilangan kecocokan yang valid
- Nilai lebih rendah (mendekati 0.1): Memungkinkan pencocokan yang lebih fleksibel, menemukan lebih banyak 
  kecocokan potensial tetapi dengan peluang lebih tinggi kecocokan yang tidak tepat
- Rentang yang direkomendasikan: 0.6-0.8 untuk sebagian besar nama bisnis

## Prioritas Lokasi
Saat beberapa nomor telepon ditemukan, pengaturan ini menentukan nomor telepon dari area geografis 
mana yang muncul pertama dalam hasil:

- "Semua": Tidak ada prioritas khusus, mengembalikan nomor sesuai urutan ditemukan
- Kota Tertentu (Jakarta, Bandung, dll.): Memprioritaskan nomor telepon dari kota yang dipilih
- Sangat berguna untuk bisnis dengan beberapa lokasi
""",
                
                # Contact Features
                "contact_features_tab": "Fitur Kontak",
                "contact_features_content": """# Fitur Informasi Kontak

## Pencarian Website
Menentukan apakah sistem mengekstrak informasi kontak tambahan dari website perusahaan:

- Saat diaktifkan: Sistem mengunjungi website setiap perusahaan dan memindai nomor telepon, alamat 
  email, dan informasi kontak tambahan yang tidak ditemukan di Google Maps
- Saat dinonaktifkan: Hanya informasi dari Google Maps yang dikumpulkan
- Catatan: Mengaktifkan ini secara signifikan meningkatkan waktu pemrosesan tetapi sering menghasilkan 
  lebih banyak detail kontak

## Pertahankan Format Asli Nomor Telepon
Mengontrol bagaimana nomor telepon disimpan dalam hasil:

- Saat diaktifkan: Nomor telepon disimpan persis seperti yang muncul di website sumber (misalnya, 
  "(021) 555-1234" atau "+62 812 3456 7890")
- Saat dinonaktifkan: Nomor telepon distandarisasi ke format yang konsisten
- Aktifkan ini ketika penampilan persis nomor telepon penting (misalnya, untuk materi pemasaran)

## Deteksi Jenis Telepon
Sistem secara otomatis mengkategorikan nomor telepon yang ditemukan menjadi:

- Telepon Seluler: Terdeteksi dari prefiks (08, +628, dll.) dan pola lainnya
- Telepon Kantor: Terdeteksi dari kode area (021, 022, dll.) dan pola telepon tetap lainnya
- Ini disimpan dalam kolom terpisah untuk memudahkan penargetan dalam kampanye pemasaran Anda
""",
                
                # Anti-Detection Features
                "anti_detection_features_tab": "Fitur Anti-Deteksi",
                "anti_detection_content": """# Fitur Anti-Deteksi

## Rotasi User-Agents
Membantu menghindari pemblokiran oleh Google saat melakukan banyak pencarian:

- Saat diaktifkan: Setiap pencarian tampak berasal dari browser/perangkat yang berbeda
- Saat dinonaktifkan: Semua pencarian menggunakan identitas browser yang sama
- Direkomendasikan: Tetap aktif untuk dataset dengan lebih dari 20 perusahaan

## Rotasi Proxy (Eksperimental)
Menggunakan koneksi internet berbeda untuk menghindari pembatasan berbasis IP:

- Saat diaktifkan: Pencarian didistribusikan melalui alamat IP yang berbeda
- Saat dinonaktifkan: Semua pencarian menggunakan koneksi internet langsung Anda
- Catatan: Fitur ini dapat mengurangi keandalan dan memperlambat pemrosesan

## Pengaturan Coba Ulang
Mengontrol bagaimana sistem menangani pencarian yang gagal:

- Maksimum Percobaan Ulang: Jumlah upaya tambahan jika pencarian gagal
- Jeda Coba Ulang: Waktu tunggu antara percobaan ulang (dalam detik)
- Nilai lebih tinggi meningkatkan tingkat keberhasilan tetapi meningkatkan total waktu pemrosesan
""",
                
                # Cache & Performance Features
                "cache_performance_tab": "Cache & Kinerja",
                "cache_performance_content": """# Fitur Cache & Kinerja

## Penyimpanan Hasil
Menyimpan hasil pencarian secara lokal untuk meningkatkan kinerja pada pencarian berulang:

- Saat diaktifkan: Perusahaan yang sebelumnya dicari diambil secara instan dari penyimpanan lokal
- Saat dinonaktifkan: Setiap pencarian mengkueri Google Maps secara langsung
- Manfaat: Pemrosesan lebih cepat, penggunaan API Google Maps berkurang, masalah deteksi lebih sedikit

## Periode Penyimpanan Cache
Mengontrol berapa lama hasil cache dianggap valid:

- Rentang: 1-365 hari
- Periode lebih pendek memastikan informasi lebih terkini tetapi memerlukan lebih banyak pencarian baru
- Periode lebih panjang memaksimalkan kinerja tetapi mungkin mengembalikan informasi yang sudah usang
- Direkomendasikan: 30 hari untuk sebagian besar penggunaan bisnis

## Bersihkan Cache
Menghapus semua hasil pencarian yang tersimpan:

- Gunakan ketika Anda ingin memaksa data baru untuk semua perusahaan
- Berguna setelah pembaruan besar Google Maps atau ketika Anda mencurigai data cache sudah usang
- Catatan: Ini akan meningkatkan waktu pemrosesan untuk operasi pencarian berikutnya

## Fitur Penyimpanan Otomatis
Menyimpan hasil Anda secara otomatis selama proses pencarian:

- Interval default: 10 menit (dapat dikonfigurasi di pengaturan)
- Membuat file cadangan dengan stempel waktu untuk mencegah kehilangan data
- Menampilkan notifikasi desktop ketika penyimpanan otomatis selesai
- Berguna untuk memulihkan data jika aplikasi atau sistem crash secara tak terduga
""",
                
                # About dialog
                "about_title": "Tentang - Idle Scrape Google Maps Scraper",
                "werkudara_scraper": "Idle Scrape Google Maps Scraper",
                "version": "Versi 2.0",
                "copyright": "© 2025 Tim Idle Scrape",
                "department": "Tim Pengembangan",
                "all_rights_reserved": "Seluruh hak dilindungi undang-undang.",
                
                # Autosave related
                "autosave_settings": "Pengaturan Penyimpanan Otomatis",
                "enable_autosave": "Aktifkan Penyimpanan Otomatis",
                "autosave_interval": "Interval penyimpanan otomatis:",
                "minutes": "menit",
                "autosave_interval_help": "Menyimpan hasil secara otomatis selama pencarian untuk mencegah kehilangan data",
                "autosave_success": "Penyimpanan Otomatis Berhasil",
                "autosave_success_msg": "Hasil telah disimpan secara otomatis ke",
                "autosave_feature": "Simpan Hasil Otomatis",
                "autosave_feature_desc": "Menyimpan hasil secara otomatis secara berkala selama proses",
                "autosave_feature_help": "Mencegah hilangnya data jika aplikasi crash",
            }
        }
    
    def apply_style(self):
        """Apply custom styling to the application widgets"""
        style = ttk.Style()
        
        # Configure progress bar style
        style.configure("TProgressbar", thickness=10)
        
        # Configure button styles
        style.configure("TButton", padding=5)
        
        # Configure frame styles
        style.configure("TFrame", background="#f5f5f5")
        style.configure("TLabelframe", background="#f5f5f5")
        style.configure("TLabelframe.Label", font=("Arial", 10, "bold"))
        
        # Configure other widget styles
        style.configure("TCheckbutton", background="#f5f5f5")
        style.configure("TRadiobutton", background="#f5f5f5")
        
        return style
    
    def get_text(self, key):
        """Get translated text for the current language"""
        lang = self.current_language.get()
        if lang in self.translations and key in self.translations[lang]:
            return self.translations[lang][key]
        # Fallback to English if translation not found
        if key in self.translations["English"]:
            return self.translations["English"][key]
        # Return key as is if no translation found
        return key
        
    def create_menu_bar(self):
        """Create application menu bar"""
        self.menu_bar = tk.Menu(self.root)
        self.root.config(menu=self.menu_bar)
        
        # File menu
        file_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label=self.get_text("file_menu"), menu=file_menu)
        file_menu.add_command(label=self.get_text("import_excel"), command=self.browse_input_file)
        file_menu.add_command(label=self.get_text("set_output_file"), command=self.browse_output_file)
        file_menu.add_separator()
        file_menu.add_command(label=self.get_text("load_data"), command=self.load_excel)
        file_menu.add_command(label=self.get_text("save_results"), command=self.save_results)
        file_menu.add_separator()
        file_menu.add_command(label=self.get_text("exit"), command=self.on_closing)
        
        # Settings menu - now a direct command instead of cascade
        self.menu_bar.add_command(label=self.get_text("settings_menu"), command=self.show_settings_dialog)
        
        # View menu - Menu untuk opsi tampilan
        view_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="View", menu=view_menu)
        view_menu.add_command(label="Data Preview (Full Screen)", command=self.show_fullscreen_preview)
        view_menu.add_separator()
        view_menu.add_command(label="Close Full Screen Window", command=self.restore_normal_view)
        
        # Process menu
        process_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label=self.get_text("process_menu"), menu=process_menu)
        process_menu.add_command(label=self.get_text("start_scraping"), command=self.process_data)
        process_menu.add_command(label=self.get_text("stop_save"), command=self.stop_processing)
        process_menu.add_separator()
        process_menu.add_command(label=self.get_text("clear_log"), command=self.clear_log)
        process_menu.add_command(label=self.get_text("copy_log"), command=self.copy_log)
        
        # Help menu
        help_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label=self.get_text("help_menu"), menu=help_menu)
        help_menu.add_command(label=self.get_text("documentation"), command=self.show_documentation)
        help_menu.add_command(label=self.get_text("feature_explanations"), command=self.show_feature_explanations)
        help_menu.add_command(label=self.get_text("about"), command=self.show_about)
    
    def show_settings_dialog(self):
        """Show the settings dialog"""
        # Create settings dialog window
        settings_window = tk.Toplevel(self.root)
        settings_window.title(self.get_text("settings_title"))
        settings_window.geometry("650x550")
        settings_window.minsize(600, 500)
        
        # Make it modal
        settings_window.transient(self.root)
        settings_window.grab_set()
        
        # Create notebook (tabbed interface)
        settings_notebook = ttk.Notebook(settings_window)
        settings_notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create tabs
        scraping_tab = ttk.Frame(settings_notebook, padding=10)
        anti_detection_tab = ttk.Frame(settings_notebook, padding=10)
        cache_tab = ttk.Frame(settings_notebook, padding=10)
        advanced_tab = ttk.Frame(settings_notebook, padding=10)
        interface_tab = ttk.Frame(settings_notebook, padding=10)  # New tab for interface settings
        
        # Add tabs to notebook
        settings_notebook.add(scraping_tab, text=self.get_text("scraping_tab"))
        settings_notebook.add(anti_detection_tab, text=self.get_text("anti_detection_tab"))
        settings_notebook.add(cache_tab, text=self.get_text("cache_tab"))
        settings_notebook.add(advanced_tab, text=self.get_text("advanced_tab"))
        settings_notebook.add(interface_tab, text=self.get_text("interface_tab"))
        
        # ---- Scraping Tab ----
        # Enable Smart Search
        smart_search_frame = ttk.LabelFrame(scraping_tab, text=self.get_text("smart_search_settings"), padding=10)
        smart_search_frame.pack(fill=tk.X, padx=5, pady=5)
        
        smart_search_check = ttk.Checkbutton(
            smart_search_frame,
            text=self.get_text("enable_smart_search_long"),
            variable=self.enable_similar_search
        )
        smart_search_check.grid(row=0, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)
        
        # Similarity threshold
        ttk.Label(smart_search_frame, text=self.get_text("similarity_threshold")).grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        threshold_frame = ttk.Frame(smart_search_frame)
        threshold_frame.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
        threshold_scale = ttk.Scale(
            threshold_frame,
            from_=0.1,
            to=1.0,
            orient=tk.HORIZONTAL,
            variable=self.similarity_threshold,
            length=300
        )
        threshold_scale.pack(side=tk.LEFT, padx=5)
        threshold_label = ttk.Label(threshold_frame, text=f"{self.similarity_threshold.get():.2f}")
        threshold_label.pack(side=tk.LEFT, padx=5)
        
        # Update the threshold label when the scale changes
        def update_threshold_label(*args):
            threshold_label.config(text=f"{self.similarity_threshold.get():.2f}")
        
        self.similarity_threshold.trace_add("write", update_threshold_label)
        
        # Website scraping
        website_frame = ttk.LabelFrame(scraping_tab, text=self.get_text("website_scraping"), padding=10)
        website_frame.pack(fill=tk.X, padx=5, pady=5)
        
        website_check = ttk.Checkbutton(
            website_frame,
            text=self.get_text("enable_website_scraping_long"),
            variable=self.enable_website_scraping
        )
        website_check.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        
        preserve_format_check = ttk.Checkbutton(
            website_frame,
            text=self.get_text("preserve_phone_format"),
            variable=self.preserve_phone_format
        )
        preserve_format_check.grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        
        # City prioritization
        city_frame = ttk.LabelFrame(scraping_tab, text=self.get_text("location_priority"), padding=10)
        city_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(city_frame, text=self.get_text("prioritize_contact")).grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        
        # List of major Indonesian cities
        cities = [
            "All",  # Default option to show all contacts
            "Jakarta", 
            "Bandung", 
            "Yogyakarta", 
            "Surabaya", 
            "Medan", 
            "Makassar", 
            "Semarang", 
            "Palembang", 
            "Denpasar",
            "Balikpapan"
        ]
        
        city_dropdown = ttk.Combobox(city_frame, textvariable=self.priority_city, values=cities, state="readonly", width=20)
        city_dropdown.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        if not self.priority_city.get() in cities:
            city_dropdown.current(0)  # Set default to "All"
            
        ttk.Label(city_frame, 
                text=self.get_text("phone_priority_help"), 
                font=("Arial", 8, "italic")).grid(row=1, column=0, columnspan=2, sticky=tk.W, padx=5, pady=2)
        
        # ---- Anti-Detection Tab ----
        user_agent_frame = ttk.LabelFrame(anti_detection_tab, text=self.get_text("user_agent_settings"), padding=10)
        user_agent_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ua_check = ttk.Checkbutton(
            user_agent_frame,
            text=self.get_text("use_rotating_agents"),
            variable=self.use_rotating_user_agents
        )
        ua_check.pack(anchor=tk.W, padx=5, pady=5)
        
        ttk.Label(user_agent_frame, 
                text=self.get_text("rotating_agents_help"), 
                font=("Arial", 8, "italic")).pack(anchor=tk.W, padx=5, pady=2)
        
        proxy_frame = ttk.LabelFrame(anti_detection_tab, text=self.get_text("proxy_settings"), padding=10)
        proxy_frame.pack(fill=tk.X, padx=5, pady=5)
        
        proxy_check = ttk.Checkbutton(
            proxy_frame,
            text=self.get_text("use_proxy_rotation"),
            variable=self.use_proxies
        )
        proxy_check.pack(anchor=tk.W, padx=5, pady=5)
        
        ttk.Label(proxy_frame, 
                text=self.get_text("proxy_warning"), 
                foreground="red",
                font=("Arial", 8, "italic")).pack(anchor=tk.W, padx=5, pady=2)
        
        delay_frame = ttk.LabelFrame(anti_detection_tab, text=self.get_text("timing_settings"), padding=10)
        delay_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(delay_frame, text=self.get_text("retry_delay")).grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        retry_delay_spinbox = ttk.Spinbox(
            delay_frame,
            from_=1,
            to=30,
            width=5,
            textvariable=self.retry_delay
        )
        retry_delay_spinbox.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        # ---- Cache Tab ----
        cache_enable_frame = ttk.LabelFrame(cache_tab, text=self.get_text("cache_settings"), padding=10)
        cache_enable_frame.pack(fill=tk.X, padx=5, pady=5)
        
        cache_check = ttk.Checkbutton(
            cache_enable_frame,
            text=self.get_text("enable_caching"),
            variable=self.enable_cache
        )
        cache_check.pack(anchor=tk.W, padx=5, pady=5)
        
        cache_days_frame = ttk.Frame(cache_enable_frame)
        cache_days_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(cache_days_frame, text=self.get_text("cache_retention")).pack(side=tk.LEFT, padx=5)
        cache_days_spinbox = ttk.Spinbox(
            cache_days_frame,
            from_=1,
            to=365,
            width=5,
            textvariable=self.cache_days
        )
        cache_days_spinbox.pack(side=tk.LEFT, padx=5)
        ttk.Label(cache_days_frame, text=self.get_text("days")).pack(side=tk.LEFT, padx=5)
        
        # Add autosave settings frame
        autosave_frame = ttk.LabelFrame(cache_tab, text=self.get_text("autosave_settings"), padding=10)
        autosave_frame.pack(fill=tk.X, padx=5, pady=5)
        
        autosave_check = ttk.Checkbutton(
            autosave_frame,
            text=self.get_text("enable_autosave"),
            variable=self.enable_autosave
        )
        autosave_check.pack(anchor=tk.W, padx=5, pady=5)
        
        autosave_interval_frame = ttk.Frame(autosave_frame)
        autosave_interval_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(autosave_interval_frame, text=self.get_text("autosave_interval")).pack(side=tk.LEFT, padx=5)
        autosave_interval_spinbox = ttk.Spinbox(
            autosave_interval_frame,
            from_=1,
            to=60,
            width=5,
            textvariable=self.autosave_interval
        )
        autosave_interval_spinbox.pack(side=tk.LEFT, padx=5)
        ttk.Label(autosave_interval_frame, text=self.get_text("minutes")).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(autosave_frame, 
                text=self.get_text("autosave_interval_help"), 
                font=("Arial", 8, "italic")).pack(anchor=tk.W, padx=5, pady=2)
        
        cache_actions_frame = ttk.Frame(cache_tab)
        cache_actions_frame.pack(fill=tk.X, padx=5, pady=10)
        
        def clear_cache():
            try:
                # Implement cache clearing logic here
                if hasattr(self, 'scraper') and self.scraper:
                    self.scraper.clear_cache()
                    messagebox.showinfo(self.get_text("cache_cleared"), self.get_text("cache_cleared_msg"))
                else:
                    # Create temporary scraper to clear cache
                    temp_scraper = GoogleMapsScraper(use_cache=True)
                    temp_scraper.clear_cache()
                    temp_scraper.close()
                    messagebox.showinfo(self.get_text("cache_cleared"), self.get_text("cache_cleared_msg"))
            except Exception as e:
                messagebox.showerror("Error", f"Failed to clear cache: {str(e)}")
        
        ttk.Button(
            cache_actions_frame,
            text=self.get_text("clear_cache"),
            command=clear_cache,
            width=15
        ).pack(side=tk.LEFT, padx=5)
        
        # ---- Advanced Tab ----
        retry_frame = ttk.LabelFrame(advanced_tab, text=self.get_text("retry_settings"), padding=10)
        retry_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(retry_frame, text=self.get_text("max_retries")).grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        max_retries_spinbox = ttk.Spinbox(
            retry_frame,
            from_=0,
            to=10,
            width=5,
            textvariable=self.max_retries
        )
        max_retries_spinbox.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(retry_frame, 
                text=self.get_text("retry_help"), 
                font=("Arial", 8, "italic")).grid(row=1, column=0, columnspan=2, sticky=tk.W, padx=5, pady=2)
        
        # Advanced warning text
        warning_text = self.get_text("advanced_warning")
        
        warning_label = ttk.Label(
            advanced_tab, 
            text=warning_text,
            wraplength=500,
            justify=tk.LEFT,
            foreground="#555555",
            font=("Arial", 9, "italic")
        )
        warning_label.pack(fill=tk.X, padx=10, pady=20)
        
        # ---- Interface Tab (New) ----
        language_frame = ttk.LabelFrame(interface_tab, text=self.get_text("language_settings"), padding=10)
        language_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(language_frame, text=self.get_text("interface_language")).grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        
        # Create language options
        languages = ["English", "Indonesia"]
        
        # Language dropdown
        language_dropdown = ttk.Combobox(
            language_frame, 
            textvariable=self.current_language,
            values=languages,
            state="readonly",
            width=20
        )
        language_dropdown.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        # Make sure a valid language is selected
        if not self.current_language.get() in languages:
            language_dropdown.current(0)  # Set default to English
        
        # Add note about language changes requiring restart
        ttk.Label(
            language_frame,
            text=self.get_text("language_note"),
            font=("Arial", 8, "italic"),
            foreground="#555555"
        ).grid(row=1, column=0, columnspan=2, sticky=tk.W, padx=5, pady=2)
        
        # Add a placeholder for future interface settings
        # This can be expanded later with theme selection, font size, etc.
        
        # Buttons at the bottom
        button_frame = ttk.Frame(settings_window)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Define button actions
        def on_cancel():
            settings_window.destroy()
            
        def on_apply():
            # Apply settings without closing
            # No need to do anything as we're using variable references
            messagebox.showinfo(self.get_text("settings_applied"), self.get_text("settings_applied_msg"))
            
        def on_ok():
            # Apply and close
            settings_window.destroy()
            
        ttk.Button(button_frame, text=self.get_text("ok"), command=on_ok, width=10).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text=self.get_text("cancel"), command=on_cancel, width=10).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text=self.get_text("apply"), command=on_apply, width=10).pack(side=tk.RIGHT, padx=5)
        
        # Center the window on the screen
        settings_window.update_idletasks()
        width = settings_window.winfo_width()
        height = settings_window.winfo_height()
        x = (settings_window.winfo_screenwidth() // 2) - (width // 2)
        y = (settings_window.winfo_screenheight() // 2) - (height // 2)
        settings_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))
        
        # Make sure the dialog is closed properly
        settings_window.protocol("WM_DELETE_WINDOW", on_cancel)
    
    def show_feature_explanations(self):
        """Show detailed feature explanations window"""
        explanations_window = tk.Toplevel(self.root)
        explanations_window.title(self.get_text("feature_explanation_title"))
        explanations_window.geometry("700x600")
        explanations_window.minsize(650, 500)
        
        # Make it modal
        explanations_window.transient(self.root)
        explanations_window.grab_set()
        
        # Create notebook (tabbed interface) for better organization
        explanations_notebook = ttk.Notebook(explanations_window)
        explanations_notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create tabs for different feature categories
        search_tab = ttk.Frame(explanations_notebook, padding=10)
        contact_tab = ttk.Frame(explanations_notebook, padding=10)
        anti_detection_tab = ttk.Frame(explanations_notebook, padding=10)
        cache_tab = ttk.Frame(explanations_notebook, padding=10)
        
        # Add tabs to notebook with translated labels
        explanations_notebook.add(search_tab, text=self.get_text("search_features_tab"))
        explanations_notebook.add(contact_tab, text=self.get_text("contact_features_tab"))
        explanations_notebook.add(anti_detection_tab, text=self.get_text("anti_detection_tab"))
        explanations_notebook.add(cache_tab, text=self.get_text("cache_performance_tab"))
        
        # ----- Search Features Tab -----
        search_text = scrolledtext.ScrolledText(search_tab, wrap=tk.WORD, padx=10, pady=10)
        search_text.pack(fill=tk.BOTH, expand=True)
        
        search_content = self.get_text("search_features_content")
        search_text.insert(tk.END, search_content)
        search_text.config(state=tk.DISABLED)
        
        # ----- Contact Features Tab -----
        contact_text = scrolledtext.ScrolledText(contact_tab, wrap=tk.WORD, padx=10, pady=10)
        contact_text.pack(fill=tk.BOTH, expand=True)
        
        contact_content = self.get_text("contact_features_content")
        contact_text.insert(tk.END, contact_content)
        contact_text.config(state=tk.DISABLED)
        
        # ----- Anti-Detection Tab -----
        anti_detection_text = scrolledtext.ScrolledText(anti_detection_tab, wrap=tk.WORD, padx=10, pady=10)
        anti_detection_text.pack(fill=tk.BOTH, expand=True)
        
        anti_detection_content = self.get_text("anti_detection_content")
        anti_detection_text.insert(tk.END, anti_detection_content)
        anti_detection_text.config(state=tk.DISABLED)
        
        # ----- Cache Tab -----
        cache_text = scrolledtext.ScrolledText(cache_tab, wrap=tk.WORD, padx=10, pady=10)
        cache_text.pack(fill=tk.BOTH, expand=True)
        
        cache_content = self.get_text("cache_performance_content")
        cache_text.insert(tk.END, cache_content)
        cache_text.config(state=tk.DISABLED)
        
        # Close button
        close_btn = ttk.Button(explanations_window, text=self.get_text("close"), command=explanations_window.destroy)
        close_btn.pack(pady=10)
        
        # Center the window on the screen
        explanations_window.update_idletasks()
        width = explanations_window.winfo_width()
        height = explanations_window.winfo_height()
        x = (explanations_window.winfo_screenwidth() // 2) - (width // 2)
        y = (explanations_window.winfo_screenheight() // 2) - (height // 2)
        explanations_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    
    def show_documentation(self):
        """Show documentation window"""
        help_window = tk.Toplevel(self.root)
        help_window.title(self.get_text("doc_title"))
        help_window.geometry("600x500")
        help_window.minsize(500, 400)
        
        # Make it modal
        help_window.transient(self.root)
        help_window.grab_set()
        
        # Add scrolled text for documentation
        doc_text = scrolledtext.ScrolledText(help_window, wrap=tk.WORD, padx=10, pady=10)
        doc_text.pack(fill=tk.BOTH, expand=True)
        
        # Add documentation content using translations
        doc_content = self.get_text("doc_content")
        doc_text.insert(tk.END, doc_content)
        doc_text.config(state=tk.DISABLED)
        
        # Close button
        close_btn = ttk.Button(help_window, text=self.get_text("close"), command=help_window.destroy)
        close_btn.pack(pady=10)
        
        # Center the window on the screen
        help_window.update_idletasks()
        width = help_window.winfo_width()
        height = help_window.winfo_height()
        x = (help_window.winfo_screenwidth() // 2) - (width // 2)
        y = (help_window.winfo_screenheight() // 2) - (height // 2)
        help_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    
    def show_about(self):
        """Show about dialog"""
        about_window = tk.Toplevel(self.root)
        about_window.title(self.get_text("about_title"))
        about_window.geometry("400x300")
        about_window.resizable(False, False)
        
        # Make it modal
        about_window.transient(self.root)
        about_window.grab_set()
        
        # Logo or icon display
        try:
            if os.path.exists('app_icon.ico'):
                logo_img = tk.PhotoImage(file='app_icon.ico')
                logo_label = ttk.Label(about_window, image=logo_img)
                logo_label.image = logo_img  # Keep a reference
                logo_label.pack(pady=10)
        except Exception:
            pass  # Silently fail if icon loading fails
        
        # Version and copyright info with translations
        ttk.Label(about_window, text=self.get_text("werkudara_scraper"), font=("Arial", 14, "bold")).pack(pady=5)
        ttk.Label(about_window, text=self.get_text("version"), font=("Arial", 10)).pack()
        ttk.Label(about_window, text="").pack()  # Spacer
        ttk.Label(about_window, text=self.get_text("copyright"), font=("Arial", 9)).pack()
        ttk.Label(about_window, text=self.get_text("department"), font=("Arial", 9)).pack()
        ttk.Label(about_window, text="").pack()  # Spacer
        ttk.Label(about_window, text=self.get_text("all_rights_reserved"), font=("Arial", 8)).pack()
        
        # Close button
        close_btn = ttk.Button(about_window, text=self.get_text("close"), command=about_window.destroy)
        close_btn.pack(pady=20)
        
        # Center the window on the screen
        about_window.update_idletasks()
        width = about_window.winfo_width()
        height = about_window.winfo_height()
        x = (about_window.winfo_screenwidth() // 2) - (width // 2)
        y = (about_window.winfo_screenheight() // 2) - (height // 2)
        about_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    
    def setup_ui(self):
        # Main frame - Using grid instead of pack
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create paned window dengan proporsi 50:50
        main_paned = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        main_paned.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Get the initial window width
        initial_width = self.root.winfo_width()
        if initial_width < 100:
            initial_width = 1200  # Default fallback width
        
        # Calculate panel widths (equal 50:50 split)
        panel_width = initial_width // 2
        
        # Left panel (configurations) dengan proporsi seimbang
        left_panel = ttk.Frame(main_paned, width=panel_width)  # Panel kiri 50%
        left_panel.pack_propagate(False)  # Don't shrink below requested width
        
        # Right panel (data preview) dengan proporsi seimbang
        right_panel = ttk.Frame(main_paned, width=panel_width)  # Panel kanan 50%
        right_panel.pack_propagate(False)  # Juga terapkan pack_propagate=False untuk panel kanan
        
        # Add panels to paned window dengan proporsi 50:50
        main_paned.add(left_panel, weight=1)  # Proporsi seimbang
        main_paned.add(right_panel, weight=1)  # Proporsi seimbang
        
        # Set posisi awal sash setelah window di-render
        self.root.update_idletasks()  # Force update layout
        
        # Jadwalkan pengaturan posisi sash dengan multiple attempts untuk keandalan
        self.root.after(100, lambda: self._set_sash_position(main_paned))
        self.root.after(500, lambda: self._set_sash_position(main_paned))
        self.root.after(1000, lambda: self._set_sash_position(main_paned))
        
        # Create a canvas with scrollbar for the left panel
        left_canvas = tk.Canvas(left_panel, borderwidth=0, highlightthickness=0)
        left_scrollbar = ttk.Scrollbar(left_panel, orient="vertical", command=left_canvas.yview)
        left_scrollable_frame = ttk.Frame(left_canvas)
        
        # Make scrollable frame expand to width of canvas but not beyond
        left_scrollable_frame.bind(
            "<Configure>",
            lambda e: left_canvas.configure(scrollregion=left_canvas.bbox("all"))
        )
        
        # Platform-specific mouse wheel scrolling
        def _on_mousewheel(event):
            if event.num == 4 or event.delta > 0:  # Scroll up
                left_canvas.yview_scroll(-1, "units")
            elif event.num == 5 or event.delta < 0:  # Scroll down
                left_canvas.yview_scroll(1, "units")
                
        # Bind mousewheel events for different platforms
        def _bind_mousewheel(event=None):
            # Windows scrolling
            self.root.bind_all("<MouseWheel>", lambda e: left_canvas.yview_scroll(int(-1*(e.delta/120)), "units"))
            # Linux scrolling
            self.root.bind_all("<Button-4>", lambda e: left_canvas.yview_scroll(-1, "units"))
            self.root.bind_all("<Button-5>", lambda e: left_canvas.yview_scroll(1, "units"))
            
        def _unbind_mousewheel(event=None):
            self.root.unbind_all("<MouseWheel>")
            self.root.unbind_all("<Button-4>")
            self.root.unbind_all("<Button-5>")
        
        # Initial binding
        _bind_mousewheel()
        
        # Bind focus events to the canvas to keep scrolling working properly
        left_canvas.bind("<Enter>", _bind_mousewheel)
        left_canvas.bind("<Leave>", _unbind_mousewheel)
        
        # Create a window inside the canvas and place it at top-left corner
        window_id = left_canvas.create_window((0, 0), window=left_scrollable_frame, anchor="nw")
        left_canvas.configure(yscrollcommand=left_scrollbar.set)
        
        # Make canvas resize with the frame
        def _configure_canvas(event):
            canvas_width = event.width
            left_canvas.itemconfig(window_id, width=canvas_width)
        left_canvas.bind("<Configure>", _configure_canvas)
        
        # Pack the canvas and scrollbar
        left_scrollbar.pack(side="right", fill="y")
        left_canvas.pack(side="left", fill="both", expand=True)
        
        # Use the scrollable frame as our left container
        left_top = left_scrollable_frame
        
        # Add vertical layout for monitoring dashboard
        left_bottom = ttk.Frame(left_scrollable_frame)
        left_bottom.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # ---- MONITORING DASHBOARD ----
        monitoring_frame = ttk.LabelFrame(left_top, text=self.get_text("monitoring_dashboard"), padding=10)
        monitoring_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Statistics Frame
        stats_frame = ttk.Frame(monitoring_frame)
        stats_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Create a grid of statistics indicators
        # Row 1: Processing statistics
        ttk.Label(stats_frame, text=self.get_text("total_companies")).grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(stats_frame, textvariable=self.total_companies_var, font=("Arial", 9, "bold")).grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(stats_frame, text=self.get_text("processed")).grid(row=0, column=2, sticky=tk.W, padx=5, pady=2)
        ttk.Label(stats_frame, textvariable=self.processed_var, font=("Arial", 9, "bold")).grid(row=0, column=3, sticky=tk.W, padx=5, pady=2)
        
        # Row 2: Data source statistics
        ttk.Label(stats_frame, text=self.get_text("google_maps_data")).grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(stats_frame, textvariable=self.gmaps_count_var, font=("Arial", 9, "bold")).grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(stats_frame, text=self.get_text("website_data")).grid(row=1, column=2, sticky=tk.W, padx=5, pady=2)
        ttk.Label(stats_frame, textvariable=self.website_count_var, font=("Arial", 9, "bold")).grid(row=1, column=3, sticky=tk.W, padx=5, pady=2)
        
        # Row 3: Phone types
        ttk.Label(stats_frame, text=self.get_text("mobile_phones")).grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(stats_frame, textvariable=self.mobile_phone_count_var, font=("Arial", 9, "bold")).grid(row=2, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(stats_frame, text=self.get_text("office_phones")).grid(row=2, column=2, sticky=tk.W, padx=5, pady=2)
        ttk.Label(stats_frame, textvariable=self.office_phone_count_var, font=("Arial", 9, "bold")).grid(row=2, column=3, sticky=tk.W, padx=5, pady=2)
        
        # Row 4: Email statistics (added to replace duplicate progress)
        ttk.Label(stats_frame, text=self.get_text("email_found")).grid(row=3, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(stats_frame, textvariable=self.email_found_var, font=("Arial", 9, "bold")).grid(row=3, column=1, columnspan=3, sticky=tk.W, padx=5, pady=2)
        
        # Current company frame
        current_frame = ttk.LabelFrame(monitoring_frame, text=self.get_text("current_processing"), padding=5)
        current_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(current_frame, text=self.get_text("company")).grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.current_company_var = tk.StringVar(value="Not Started")
        ttk.Label(current_frame, textvariable=self.current_company_var, font=("Arial", 9, "bold")).grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(current_frame, text=self.get_text("data_source")).grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.current_source_var = tk.StringVar(value="N/A")
        ttk.Label(current_frame, textvariable=self.current_source_var, font=("Arial", 9, "bold")).grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        
        # Progress bar - now only shown once in the monitoring dashboard
        progress_frame = ttk.LabelFrame(monitoring_frame, text=self.get_text("progress"), padding=5)
        progress_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", length=100, mode="determinate", variable=self.progress_var, style="TProgressbar")
        self.progress_bar.pack(padx=5, pady=5, fill=tk.X)
        
        # Percentage and status
        progress_detail_frame = ttk.Frame(progress_frame)
        progress_detail_frame.pack(fill=tk.X, expand=True)
        
        # Progress percentage
        self.progress_percent_var = tk.StringVar(value="0%")
        ttk.Label(progress_detail_frame, textvariable=self.progress_percent_var, font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=5)
        
        # Status label
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(progress_detail_frame, textvariable=self.status_var, font=("Arial", 9)).pack(side=tk.RIGHT, padx=5)
        
        # Quick Settings frame - Essential settings only
        quick_settings_frame = ttk.LabelFrame(left_top, text=self.get_text("quick_settings"), padding=10)
        quick_settings_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Two most important options as quick toggles
        self.quick_smart_search = ttk.Checkbutton(
            quick_settings_frame, 
            text="Enable Smart Search",
            variable=self.enable_similar_search
        )
        self.quick_smart_search.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        
        self.quick_website_scraping = ttk.Checkbutton(
            quick_settings_frame,
            text="Enable Website Scraping",
            variable=self.enable_website_scraping
        )
        self.quick_website_scraping.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        # Info label
        settings_info = ttk.Label(quick_settings_frame, 
                                text="More options available in the Settings menu", 
                                font=("Arial", 8, "italic"))
        settings_info.grid(row=1, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)
        
        # File selection frame - Simplified with info message
        file_frame = ttk.LabelFrame(left_top, text="File Selection", padding=10)
        file_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Input file only shown - cleaner interface
        input_file_frame = ttk.Frame(file_frame)
        input_file_frame.pack(fill=tk.X, expand=True, padx=5, pady=5)
        
        ttk.Label(input_file_frame, text="Excel File:").pack(side=tk.LEFT, padx=5)
        ttk.Entry(input_file_frame, textvariable=self.input_file_var, width=50).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(input_file_frame, text="Browse", command=self.browse_input_file).pack(side=tk.LEFT, padx=5)
        
        # Add info label
        file_info = ttk.Label(file_frame, text="Output file can be configured in the File menu", font=("Arial", 8, "italic"))
        file_info.pack(anchor=tk.W, padx=5, pady=2)
        
        # Quick action buttons without frame
        action_frame = ttk.Frame(left_top)
        action_frame.pack(fill=tk.X, padx=5, pady=10)
        
        # Create button layout in a row
        self.load_btn = ttk.Button(action_frame, text="Load Data", width=15, command=self.load_excel)
        self.load_btn.pack(side=tk.LEFT, padx=5)
        
        self.process_btn = ttk.Button(action_frame, text="Start Scraping", width=15, 
                                     command=self.process_data, state=tk.DISABLED)
        self.process_btn.pack(side=tk.LEFT, padx=5)
        
        self.stop_btn = ttk.Button(action_frame, text="Stop", width=10,
                                  command=self.stop_processing, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)
        
        self.save_btn = ttk.Button(action_frame, text="Save Results", width=15,
                                  command=self.save_results, state=tk.DISABLED)
        self.save_btn.pack(side=tk.LEFT, padx=5)
        
        # Configure action frame to distribute space evenly
        for i in range(4):
            action_frame.columnconfigure(i, weight=1)
        
        # Log frame with a notebook for different log views
        self.log_notebook = ttk.Notebook(left_bottom)
        self.log_notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Main log tab
        self.log_frame = ttk.Frame(self.log_notebook, padding=5)
        self.log_notebook.add(self.log_frame, text="Log Terminal")
        
        # Error log tab
        self.error_log_frame = ttk.Frame(self.log_notebook, padding=5)
        self.log_notebook.add(self.error_log_frame, text="Errors")
        
        # Success log tab
        self.success_log_frame = ttk.Frame(self.log_notebook, padding=5)
        self.log_notebook.add(self.success_log_frame, text="Found Data")
        
        # Create log text widget with scrollbar
        self.log_text = scrolledtext.ScrolledText(self.log_frame, wrap=tk.WORD, height=10, bg="black", fg="lime")
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)
        
        # Error log text widget
        self.error_log_text = scrolledtext.ScrolledText(self.error_log_frame, wrap=tk.WORD, height=10, bg="black", fg="#FF6666")
        self.error_log_text.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)
        
        # Success log text widget
        self.success_log_text = scrolledtext.ScrolledText(self.success_log_frame, wrap=tk.WORD, height=10, bg="black", fg="#66FF66")
        self.success_log_text.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)
        
        # Log controls
        log_controls = ttk.Frame(left_bottom)
        log_controls.pack(fill=tk.X, pady=5)
        
        # Add Copy Log button
        self.copy_log_btn = ttk.Button(log_controls, text="Copy Log", width=15, command=self.copy_log)
        self.copy_log_btn.pack(side=tk.RIGHT, padx=5)
        
        # Clear Log button
        self.clear_log_btn = ttk.Button(log_controls, text="Clear Log", width=15, command=self.clear_log)
        self.clear_log_btn.pack(side=tk.RIGHT, padx=5)
        
        # ------------- RIGHT PANEL - DATA PREVIEW -------------
        
        # Data Preview occupies entire right panel (tanpa tombol fullscreen)
        preview_frame = ttk.LabelFrame(right_panel, text="Data Preview", padding=10)
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create frame to contain the treeview and scrollbars
        tree_frame = ttk.Frame(preview_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)
        
        # Create Treeview for data preview dengan horizontal scrollbar di atas
        print("Initializing Treeview and scrollbars")
        
        # Frame untuk horizontal scrollbar di atas (untuk kemudahan navigasi)
        hsb_top_frame = ttk.Frame(tree_frame)
        hsb_top_frame.grid(row=0, column=0, sticky='ew')
        
        # Horizontal scrollbar di atas
        hsb_top = ttk.Scrollbar(hsb_top_frame, orient="horizontal")
        hsb_top.pack(fill=tk.X, side=tk.TOP)
        
        # Treeview di tengah
        self.tree = ttk.Treeview(tree_frame)
        
        # Scrollbars - vertical di kanan, horizontal di atas dan bawah
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        
        # Konfigurasikan scrollbar atas untuk treeview juga
        hsb_top.config(command=self.tree.xview)
        
        # Configure treeview to use scrollbars
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=lambda first, last: (hsb.set(first, last), hsb_top.set(first, last)))
        
        # Grid layout for treeview and scrollbars
        self.tree.grid(row=1, column=0, sticky='nsew')
        vsb.grid(row=1, column=1, sticky='ns')
        hsb.grid(row=2, column=0, sticky='ew')
        
        # Configure the tree frame grid to expand properly
        tree_frame.rowconfigure(1, weight=1)  # Row 1 is now for treeview
        tree_frame.rowconfigure(0, weight=0)  # Row 0 for top scrollbar - fixed height
        tree_frame.rowconfigure(2, weight=0)  # Row 2 for bottom scrollbar - fixed height
        tree_frame.columnconfigure(0, weight=5)  # Give more weight to column with treeview
        
        # Initialize empty treeview with default columns
        default_columns = [
            "No.", 
            "Nama Perusahaan", 
            "Alamat", 
            "Kecamatan", 
            "Mobile Phone", 
            "Office Phone", 
            "Website", 
            "Email", 
            "Rating"
        ]
        self.tree["columns"] = default_columns
        self.tree["show"] = "headings"
        
        # Set up default column headings
        column_widths = {
            "No.": 50,
            "Nama Perusahaan": 200,
            "Alamat": 200,
            "Kecamatan": 120,
            "Mobile Phone": 150,
            "Office Phone": 150,
            "Website": 200,
            "Email": 180,
            "Rating": 60
        }
        
        for col in default_columns:
            self.tree.heading(col, text=col)
            width = column_widths.get(col, 120)  # Default width if not specified
            self.tree.column(col, width=width, minwidth=50)
        
        # Add a top-level separator before the footer frame
        top_separator = ttk.Separator(main_frame, orient='horizontal')
        top_separator.pack(fill=tk.X, pady=5)
        
        # Version footer in main frame
        footer_frame = ttk.Frame(main_frame)
        footer_frame.pack(fill=tk.X, padx=5, pady=(0, 5))
        
        # Add separator line above footer
        separator = ttk.Separator(footer_frame, orient='horizontal')
        separator.pack(fill=tk.X, pady=5)
        
        # Copyright text on left
        copyright_label = ttk.Label(footer_frame, text="Copyright © Werkudara Group 2025 - BAS Department - Pramuji", font=("Arial", 8))
        copyright_label.pack(side=tk.LEFT, padx=5)
        
        # Version on right
        version_label = ttk.Label(footer_frame, text="Version 2.5", font=("Arial", 8, "bold"))
        version_label.pack(side=tk.RIGHT, padx=5)
    
    def log(self, message):
        """Log a message to the queue for thread-safe logging"""
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}"
        self.queue.put({"message": formatted_message, "type": "normal"})
        
        # Log to specific tabs based on message content
        if any(error_keyword in message.lower() for error_keyword in 
              ["error", "failed", "could not", "exception", "no data found"]):
            self.queue.put({"message": formatted_message, "type": "error"})
        elif any(success_keyword in message for success_keyword in 
                ["✓", "found", "successfully", "extracted", "updated"]):
            self.queue.put({"message": formatted_message, "type": "success"})
    
    def check_queue(self):
        """Check for and process any messages in the queue"""
        try:
            # Process at most 20 items per call to avoid freezing
            items_processed = 0
            max_items_per_call = 20
            
            while not self.queue.empty() and items_processed < max_items_per_call:
                item = self.queue.get_nowait()
                items_processed += 1
                
                try:
                    # Check if item is a tuple (which is expected format)
                    if isinstance(item, tuple) and len(item) == 2:
                        item_type, item_value = item
                        
                        if item_type == 'log':
                            # Directly update log widgets instead of calling self.log
                            timestamp = datetime.datetime.now().strftime("%H:%M:%S")
                            formatted_message = f"[{timestamp}] {item_value}"
                            self._append_to_log(self.log_text, formatted_message)
                            
                            # Add to specific logs if applicable
                            if any(error_keyword in str(item_value).lower() for error_keyword in 
                                ["error", "failed", "could not", "exception", "no data found"]):
                                self._append_to_log(self.error_log_text, formatted_message)
                            elif any(success_keyword in str(item_value) for success_keyword in 
                                    ["✓", "found", "successfully", "extracted", "updated"]):
                                self._append_to_log(self.success_log_text, formatted_message)
                        elif item_type == 'current':
                            self.current_company_var.set(item_value)
                        elif item_type == 'progress':
                            self.progress_var.set(item_value)
                            self.progress_percent_var.set(f"{item_value:.1f}%")
                        elif item_type == 'total':
                            self.total_companies_var.set(str(item_value))
                        elif item_type == 'source':
                            self.current_source_var.set(item_value)
                    # Backward compatibility for dictionary format
                    elif isinstance(item, dict):
                        if "message" in item:
                            # Directly process message
                            timestamp = datetime.datetime.now().strftime("%H:%M:%S")
                            formatted_message = f"[{timestamp}] {item['message']}"
                            self._append_to_log(self.log_text, formatted_message)
                            
                            # Add to specific logs if applicable
                            if any(error_keyword in str(item['message']).lower() for error_keyword in 
                                ["error", "failed", "could not", "exception", "no data found"]):
                                self._append_to_log(self.error_log_text, formatted_message)
                            elif any(success_keyword in str(item['message']) for success_keyword in 
                                    ["✓", "found", "successfully", "extracted", "updated"]):
                                self._append_to_log(self.success_log_text, formatted_message)
                        
                        if "current" in item:
                            self.current_company_var.set(item["current"])
                        if "progress" in item:
                            self.progress_var.set(item["progress"])
                            self.progress_percent_var.set(f"{item['progress']:.1f}%")
                    else:
                        # If it's just a string, treat it as a log message
                        # But don't call self.log() to avoid recursion
                        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
                        formatted_message = f"[{timestamp}] {str(item)}"
                        self._append_to_log(self.log_text, formatted_message)
                
                    self.queue.task_done()
                    
                except Exception as item_error:
                    # Log the error but continue processing other items
                    error_msg = f"Error processing queue item: {str(item_error)}"
                    print(error_msg, file=sys.__stdout__)
                    self._append_to_log(self.error_log_text, f"[ERROR] {error_msg}")
                    self.queue.task_done()
        
        except Exception as e:
            print(f"Error in check_queue: {str(e)}", file=sys.__stdout__)
            if hasattr(self, 'error_log_text'):
                try:
                    self._append_to_log(self.error_log_text, f"[SYSTEM] Queue processing error: {str(e)}")
                except:
                    pass
        finally:
            # Schedule next check - use a longer interval if the app is struggling
            self.root.after(200, self.check_queue)
    
    def _append_to_log(self, log_widget, message):
        """Append a message to the specified log widget with color formatting"""
        try:
            log_widget.config(state=tk.NORMAL)
            log_widget.insert(tk.END, message + "\n")
        
            # Apply tag-based coloring for keywords
            if "error" in message.lower():
                start_idx = message.lower().find("error")
                end_idx = start_idx + 5
                log_widget.tag_add("error", f"end-{len(message)+1-start_idx}c", f"end-{len(message)+1-end_idx}c")
                log_widget.tag_config("error", foreground="red")
        
            if "✓" in message:
                start_idx = message.find("✓")
                end_idx = start_idx + 1
                log_widget.tag_add("check", f"end-{len(message)+1-start_idx}c", f"end-{len(message)+1-end_idx}c")
                log_widget.tag_config("check", foreground="#00FF00")
            
            # Keep only the last 1000 lines
            line_count = int(log_widget.index('end-1c').split('.')[0])
            if line_count > 1000:
                log_widget.delete('1.0', f"{line_count-1000}.0")
            
            log_widget.see(tk.END)
            log_widget.config(state=tk.DISABLED)
        except Exception as e:
            # Last resort fallback to prevent crashing
            print(f"Error in _append_to_log: {str(e)}", file=sys.__stdout__)
    
    def clear_log(self):
        """Clear all log text widgets"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        
        self.error_log_text.config(state=tk.NORMAL)
        self.error_log_text.delete(1.0, tk.END)
        self.error_log_text.config(state=tk.DISABLED)
        
        self.success_log_text.config(state=tk.NORMAL)
        self.success_log_text.delete(1.0, tk.END)
        self.success_log_text.config(state=tk.DISABLED)
        
        self.log("Log cleared")
    
    def copy_log(self):
        """Copy the current log content to clipboard"""
        # Get the currently visible tab
        current_tab = self.log_notebook.index(self.log_notebook.select())
        
        # Get content from the appropriate log widget
        if current_tab == 0:  # Main log
            content = self.log_text.get(1.0, tk.END)
            log_type = "main"
        elif current_tab == 1:  # Error log
            content = self.error_log_text.get(1.0, tk.END)
            log_type = "error"
        elif current_tab == 2:  # Success log
            content = self.success_log_text.get(1.0, tk.END)
            log_type = "success"
        else:
            content = ""
            log_type = "unknown"
        
        # Copy to clipboard
        self.root.clipboard_clear()
        self.root.clipboard_append(content)
        
        # Log success message
        self.log(f"Copied {log_type} log content to clipboard")
        
        # Flash the copy button to give visual feedback
        original_bg = self.copy_log_btn.cget("background")
        self.copy_log_btn.configure(background="lightgreen")
        self.root.after(200, lambda: self.copy_log_btn.configure(background=original_bg))
    
    def update_threshold_label(self, *args):
        """Update the threshold label when the slider changes"""
        self.threshold_label.config(text=f"{self.similarity_threshold.get():.2f}")
    
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
        
        # Disable interface during loading
        self.load_btn.config(state=tk.DISABLED)
        self.process_btn.config(state=tk.DISABLED)
        self.save_btn.config(state=tk.DISABLED)
        self.status_var.set("Loading file, please wait...")
        self.root.update_idletasks()  # Force UI update
        
        # Start file loading in a separate thread
        Thread(target=self._load_excel_thread, args=(input_file,), daemon=True).start()
    
    def _load_excel_thread(self, input_file):
        """Thread function to handle Excel file loading without blocking the UI"""
        try:
            # Read Excel file
            self.log(f"Loading Excel file: {input_file}")
            
            # Load file with a timeout to prevent hanging on corrupted files
            try:
                df = pd.read_excel(input_file)
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Error", f"Failed to read Excel file: {str(e)}"))
                self.root.after(0, lambda: self.status_var.set("Error loading file"))
                self.root.after(0, lambda: self.load_btn.config(state=tk.NORMAL))
                self.log(f"Error reading Excel file: {str(e)}")
                return
            
            # Check if DataFrame is empty
            if df.empty:
                self.root.after(0, lambda: messagebox.showerror("Error", "Excel file is empty"))
                self.root.after(0, lambda: self.status_var.set("Empty file"))
                self.root.after(0, lambda: self.load_btn.config(state=tk.NORMAL))
                self.log("Excel file is empty")
                return
            
            # Define possible column name formats - now with more variations
            column_mappings = {
                'no': ['No.', 'No', 'NO', 'no.', 'no', 'Nomor', 'nomor', '#', 'Id', 'ID', 'id'],
                'company': ['Nama Perusahaan', 'Nama_Perusahaan', 'NamaPerusahaan', 'Nama', 'Company', 'company', 'NAMA', 
                           'CompanyName', 'Company Name', 'BusinessName', 'Business Name', 'Perusahaan', 'Nama Bisnis', 'nama'],
                'address': ['Alamat', 'ALAMAT', 'alamat', 'Address', 'address', 'Location', 'location', 'Lokasi', 'lokasi'],
                'district': ['Kecamatan', 'KECAMATAN', 'kecamatan', 'District', 'district', 'City', 'city', 'Kota', 'kota', 
                            'Area', 'area', 'Region', 'region', 'Wilayah', 'wilayah', 'Daerah', 'daerah', 'Provinsi', 'provinsi', 
                            'Province', 'province', 'State', 'state']
            }
            
            # First, print all available columns for debugging
            self.log(f"Available columns in file: {', '.join(df.columns)}")
            
            # Map actual column names to standard names
            column_map = {}
            required_columns = ['company']  # Only company name is absolutely required
            optional_columns = ['no', 'address', 'district']
            missing_required_columns = []
            
            # Find required columns
            for standard_key in required_columns:
                possible_names = column_mappings[standard_key]
                found = False
                for name in possible_names:
                    if name in df.columns:
                        column_map[standard_key] = name
                        found = True
                        break
                
                if not found:
                    missing_required_columns.append(standard_key)
            
            # Find optional columns if available
            for standard_key in optional_columns:
                possible_names = column_mappings[standard_key]
                for name in possible_names:
                    if name in df.columns:
                        column_map[standard_key] = name
                        break
            
            self.log(f"Found column mapping: {column_map}")
            
            if missing_required_columns:
                error_msg = (f"Could not find required column: Company Name\n\n"
                                  "Your Excel file must have at least one column for company names.\n"
                                  "Possible column names: Nama Perusahaan, Company, Business Name, etc.\n\n"
                                   f"Available columns: {', '.join(df.columns)}")
                self.root.after(0, lambda: messagebox.showerror("Error", error_msg))
                self.root.after(0, lambda: self.status_var.set("Error: Missing required columns"))
                self.root.after(0, lambda: self.load_btn.config(state=tk.NORMAL))
                return
            
            # Create a new DataFrame with standardized column names
            self.data = pd.DataFrame()
            
            # Daftar lengkap kolom yang dibutuhkan dalam urutan yang diinginkan
            desired_columns = [
                'No.', 
                'Nama Perusahaan', 
                'Alamat', 
                'Kecamatan', 
                'Mobile Phone', 
                'Office Phone', 
                'Website', 
                'Email', 
                'Rating', 
                'Reviews Count', 
                'Category', 
                'Updated Address', 
                'Updated Company Name', 
                'Similar Company Found', 
                'Similarity Score', 
                'Mapped Company', 
                'Phone Source', 
                'Email Source', 
                'Data Source', 
                'Source URL'
            ]
            
            # Tambahkan semua kolom yang dibutuhkan dengan nilai None (kosong)
            for col in desired_columns:
                self.data[col] = None
                
            # Add company name (required)
            try:
                self.data['Nama Perusahaan'] = df[column_map['company']]
                
                # Add number column if available
                if 'no' in column_map:
                    self.data['No.'] = df[column_map['no']]
                else:
                    # Create sequence numbers if no number column exists
                    self.data['No.'] = range(1, len(df) + 1)
                
                # Add address if available
                if 'address' in column_map:
                    self.data['Alamat'] = df[column_map['address']]
                
                # Add district/location if available
                if 'district' in column_map:
                    self.data['Kecamatan'] = df[column_map['district']]
                
                # Jika ada kolom lain yang sudah ada di file Excel, tambahkan juga
                for col in df.columns:
                    if col in desired_columns and col not in ['Nama Perusahaan', 'No.', 'Alamat', 'Kecamatan']:
                        self.data[col] = df[col]
                    
            except Exception as e:
                error_msg = f"Error processing column mapping: {str(e)}"
                self.root.after(0, lambda: messagebox.showerror("Error", error_msg))
                self.root.after(0, lambda: self.status_var.set("Error: Failed to process columns"))
                self.root.after(0, lambda: self.load_btn.config(state=tk.NORMAL))
                self.log(error_msg)
                return
            
            # Update UI safely from main thread
            def update_ui():
                try:
                    print(f"Running update_ui with dataframe of {len(self.data)} rows")
                    
                    # Update status first to show progress
                    self.status_var.set(f"Updating preview with {len(df)} companies...")
                    self.root.update_idletasks()
                    
                    # Update the treeview safely
                    print("Calling update_tree...")
                    self.update_tree(self.data)
                    
                    # Update statistics
                    print("Updating statistics...")
                    self._update_statistics(self.data)
                    
                    # Enable process button
                    self.process_btn.config(state=tk.NORMAL)
                    self.load_btn.config(state=tk.NORMAL)
                    
                    # Update status
                    self.status_var.set(f"Loaded {len(df)} companies from {input_file}")
                    self.progress_var.set(0)  # Reset progress bar
                    self.progress_percent_var.set("0%")  # Reset progress percentage
                    self.total_companies_var.set(str(len(df)))  # Update total companies count
                    self.processed_var.set("0 (0%)")  # Reset processed count
                    self.log(f"Successfully loaded {len(df)} companies")
                except Exception as ui_error:
                    self.log(f"Error updating UI after loading: {str(ui_error)}")
                    print(f"ERROR in update_ui: {str(ui_error)}")
                    import traceback
                    traceback.print_exc()
                    self.status_var.set("Error updating UI")
                    self.load_btn.config(state=tk.NORMAL)
            
            print("Scheduling UI update...")
            # Schedule UI update to run on main thread
            self.root.after(100, update_ui)  # Slight delay to ensure thread sync
            
        except Exception as e:
            error_msg = f"Error loading Excel file: {str(e)}"
            self.root.after(0, lambda: messagebox.showerror("Error", error_msg))
            self.root.after(0, lambda: self.status_var.set("Error loading file"))
            self.root.after(0, lambda: self.load_btn.config(state=tk.NORMAL))
            self.log(error_msg)
            traceback.print_exc(file=sys.__stdout__)
    
    def update_tree(self, df):
        try:
            # Clear existing tree - make sure tree exists
            if not hasattr(self, "tree") or not self.tree:
                print("ERROR: Treeview is not initialized")
                return
                
            # Remove existing items
            for i in self.tree.get_children():
                self.tree.delete(i)
            
            print(f"Updating treeview with dataframe containing {len(df)} rows")
            print(f"Columns: {df.columns.tolist()}")
            
            # Definisikan urutan kolom yang diinginkan - prioritaskan kolom penting di awal
            desired_order = [
                "No.", 
                "Nama Perusahaan", 
                "Alamat", 
                "Kecamatan", 
                "Mobile Phone",  # Prioritas utama sesuai permintaan
                "Office Phone",  # Prioritas utama sesuai permintaan
                "Website", 
                "Email", 
                "Rating", 
                "Category",
                "Reviews Count", 
                "Updated Address", 
                "Updated Company Name", 
                "Similar Company Found", 
                "Similarity Score", 
                "Mapped Company", 
                "Phone Source", 
                "Email Source", 
                "Data Source", 
                "Source URL"
            ]
            
            # Buat daftar kolom yang ada dalam urutan yang diinginkan
            ordered_cols = [col for col in desired_order if col in df.columns]
            
            # Tambahkan kolom yang mungkin ada di dataframe tapi tidak ada di daftar yang diinginkan
            remaining_cols = [col for col in df.columns if col not in ordered_cols]
            ordered_cols.extend(remaining_cols)
            
            # Gunakan kolom dalam urutan yang sudah diurutkan
            self.tree["columns"] = ordered_cols
            self.tree["show"] = "headings"
                
            # Define tags for different row types
            self.tree.tag_configure('oddrow', background='#f0f0f0')
            self.tree.tag_configure('evenrow', background='white')
            self.tree.tag_configure('gmaps_data', background='#e6f2ff')  # Light blue for Google Maps data
            self.tree.tag_configure('website_data', background='#e6ffe6')  # Light green for Website data
            self.tree.tag_configure('mobile_phone', foreground='#006600')  # Dark green for mobile phones
            self.tree.tag_configure('office_phone', foreground='#000066')  # Dark blue for office phones
            self.tree.tag_configure('no_data', background='#fff0f0')  # Light red for rows with no data
            self.tree.tag_configure('moredata', background='#e0e0e0')  # For "more data" indicator
            
            # Set column headings with improved styling
            for col in ordered_cols:
                self.tree.heading(col, text=col, anchor=tk.CENTER)
                
                # Determine appropriate column width based on content
                # Calculate max width from column name and first 20 values
                values = [str(x) for x in df[col].dropna().head(20).tolist()]
                values.append(col)  # Include column name in width calculation
                
                # Get max length
                if values:
                    max_len = max(len(str(v)) for v in values)
                    # Width calculation: characters * avg char width + padding
                    width = min(max(max_len * 8, 100), 300)  # Minimum 100px, Maximum 300px
                else:
                    width = 100  # Default width
                    
                # Kolom-kolom dikecilkan untuk mode non-fullscreen
                # Buat kolom lebih sempit agar lebih banyak kolom terlihat
                if col == 'No.':
                    width = 40  # Kolom No. sangat sempit
                elif col == 'Nama Perusahaan':
                    width = 120  # Nama perusahaan lebih sempit
                elif col == 'Alamat':
                    width = 120  # Alamat lebih sempit
                elif col == 'Kecamatan':
                    width = 100  # Kecamatan lebih sempit
                elif col in ['Mobile Phone', 'Office Phone']:
                    width = 120  # Nomor telepon prioritas terlihat
                elif col in ['Website', 'Email']:
                    width = 120  # Email dan website lebih sempit
                elif col in ['Rating', 'Reviews Count', 'Similarity Score']:
                    width = 70  # Kolom angka lebih sempit
                else:
                    # Kolom lainnya standard
                    width = 100
                    
                # Apply the calculated width
                self.tree.column(col, width=width, minwidth=50, anchor=tk.W)
                
            print("Column headers configured")
            
            # OPTIMASI: Gunakan batch processing untuk input data ke treeview
            # - Tampilkan maksimal 100 baris awal untuk preview
            # - Proses dalam batch 20 baris agar UI tetap responsif
            max_preview_rows = 100
            display_rows = min(len(df), max_preview_rows)
            batch_size = 20
            
            # Label untuk menunjukkan sedang loading data - gunakan grid karena parent menggunakan grid
            loading_label = ttk.Label(self.tree.master, text="Loading data preview...", font=("Arial", 10, "italic"))
            loading_label.grid(row=1, column=0, sticky='n')
            self.root.update_idletasks()  # Refresh UI untuk menampilkan label loading
            
            row_count = 0
            for start_idx in range(0, display_rows, batch_size):
                end_idx = min(start_idx + batch_size, display_rows)
                batch = df.iloc[start_idx:end_idx]
                
                # Proses setiap batch
                for i, row in batch.iterrows():
                    values = [row[col] if not pd.isna(row[col]) else "" for col in ordered_cols]
                    
                    # Determine the appropriate tags for this row
                    tags = []
                    
                    # Add basic zebra striping
                    tags.append('evenrow' if i % 2 == 0 else 'oddrow')
                    
                    # Check data source for color highlighting
                    if 'Data Source' in df.columns and row['Data Source'] is not None:
                        if 'google maps' in str(row['Data Source']).lower():
                            tags.append('gmaps_data')
                        elif 'website' in str(row['Data Source']).lower():
                            tags.append('website_data')
                    
                    # Check phone columns for text coloring
                    if 'Mobile Phone' in df.columns and row['Mobile Phone'] is not None and str(row['Mobile Phone']) != "":
                        tags.append('mobile_phone')
                    
                    if 'Office Phone' in df.columns and row['Office Phone'] is not None and str(row['Office Phone']) != "":
                        tags.append('office_phone')
                            
                    # Check if no data found
                    has_data = False
                    if ('Mobile Phone' in df.columns and row['Mobile Phone'] is not None and str(row['Mobile Phone']) != "") or \
                       ('Office Phone' in df.columns and row['Office Phone'] is not None and str(row['Office Phone']) != ""):
                        has_data = True
                        
                    if not has_data:
                        tags.append('no_data')
                    
                    # Insert row with appropriate tags
                    self.tree.insert("", "end", values=values, tags=tuple(tags))
                    row_count += 1
                
                # Update UI setiap batch untuk tetap responsif
                self.root.update_idletasks()
            
            # Hapus label loading agar tidak mengganggu tampilan
            try:
                loading_label.destroy()
            except:
                pass  # Abaikan error jika ada masalah
                
            # Indicate if there are more rows
            if len(df) > max_preview_rows:
                remaining = len(df) - max_preview_rows
                self.tree.insert("", "end", values=[f"... and {remaining} more rows (use fullscreen view to see all) ...", *["" for _ in range(len(ordered_cols)-1)]], tags=('moredata',))
                
            print(f"Inserted {row_count} rows into treeview")
            
            # Update statistics dashboard
            self._update_statistics(df)
        except Exception as e:
            print(f"ERROR in update_tree: {str(e)}")
            import traceback
            traceback.print_exc()
    
    def process_data(self):
        if self.data is None or len(self.data) == 0:
            messagebox.showerror("Error", "No data to process. Please load an Excel file first.")
            return
        
        output_file = self.output_file_var.get().strip()
        if not output_file:
            messagebox.showerror("Error", "Please specify an output Excel file")
            return
        
        # Reset stop flag
        self.scraping_active = False
        
        # Disable buttons during processing
        self.load_btn.config(state=tk.DISABLED)
        self.process_btn.config(state=tk.DISABLED)
        self.save_btn.config(state=tk.DISABLED)
        
        # Enable stop button
        self.stop_btn.config(state=tk.NORMAL)
        
        # Reset progress
        self.progress_var.set(0)
        
        # Start processing in a separate thread
        self.thread = Thread(target=self._scrape_data_thread)
        self.thread.daemon = True
        self.thread.start()
    
    def stop_processing(self):
        """Stop the processing and save current results"""
        if not self.scraping_active:
            return
            
        self.log("Stop requested. Waiting for current company to finish...")
        self.status_var.set("Stopping - please wait for current task to finish")
        self.scraping_active = False
        self.stop_btn.config(state=tk.DISABLED)
    
        # Cancel autosave timer
        if self.autosave_timer is not None:
            self.autosave_timer.cancel()
            self.autosave_timer = None
            
        # Offer to save the results now
        self.offer_save_partial_results()
    
    def offer_save_partial_results(self):
        """Offer to save partial results after stopping"""
        if messagebox.askyesno("Save Results", "Do you want to save the partial results now?"):
            self.save_results()
    
    def _get_area_codes_for_city(self, city):
        """Get area codes for the specified city"""
        area_codes = {
            "Jakarta": ["021", "21"],
            "Bandung": ["022", "22"],
            "Yogyakarta": ["0274", "274"],
            "Surabaya": ["031", "31"],
            "Medan": ["061", "61"],
            "Makassar": ["0411", "411"],
            "Semarang": ["024", "24"],
            "Palembang": ["0711", "711"],
            "Denpasar": ["0361", "361"],
            "Balikpapan": ["0542", "542"]
        }
        
        return area_codes.get(city, [])
    
    def _scrape_data_thread(self):
        """Thread function to handle the actual scraping work."""
        try:
            self.scraping_active = True
            
            # Initialize scraper
            self.log("Initializing Google Maps scraper...")
            self.scraped_data = []
            
            # Create scraper with current settings
            self.scraper = GoogleMapsScraper(
                max_retries=self.max_retries.get(),
                retry_delay=self.retry_delay.get(),
                enable_similar_search=self.enable_similar_search.get(),
                similarity_threshold=self.similarity_threshold.get(),
                enable_website_scraping=self.enable_website_scraping.get(),
                preserve_phone_format=self.preserve_phone_format.get(),
                use_cache=self.enable_cache.get(),
                cache_days=self.cache_days.get(),
                use_rotating_user_agents=self.use_rotating_user_agents.get(),
                use_proxies=self.use_proxies.get()
            )
            
            # Schedule first autosave if enabled
            if self.enable_autosave.get():
                self._schedule_autosave()
            
            # Get the data column
            company_names = self.data['Nama Perusahaan'].dropna().tolist()
            total_companies = len(company_names)
            self.log(f"Found {total_companies} companies to scrape")
            self.queue.put(('total', total_companies))
            
            # Process each company
            for i, company_name in enumerate(company_names):
                # Check if stop was requested
                if not self.scraping_active:
                    self.log("Processing stopped by user")
                    break
                
                # Update progress and status
                progress_pct = (i / total_companies) * 100
                self.progress_var.set(progress_pct)
                self.progress_percent_var.set(f"{progress_pct:.1f}%")
                
                # Update current company in GUI thread-safely
                self.root.after(0, lambda name=company_name: self.current_company_var.set(name))
                
                self.log(f"Processing {i+1}/{total_companies}: {company_name}")
                
                # Handle potential missing location data
                address = self.data.at[i, 'Alamat']
                district = self.data.at[i, 'Kecamatan']
                
                # Skip empty company names
                if not company_name or pd.isna(company_name) or str(company_name).strip() == '':
                    self.log(f"Skipping row {i+1}: Empty company name")
                    continue
                    
                self.log(f"Processing {i+1}/{total_companies}: {company_name}")
                
                # Create search query - company name is required, location is optional
                search_query = company_name
                
                # Add location details if available to improve search accuracy
                location_parts = []
                if address and not pd.isna(address) and str(address).strip() != '':
                    location_parts.append(str(address).strip())
                if district and not pd.isna(district) and str(district).strip() != '':
                    location_parts.append(str(district).strip())
                    
                # Append non-empty location parts to search query
                if location_parts:
                    search_query = f"{company_name} {' '.join(location_parts)}"
                
                self.log(f"Search query: {search_query}")
                
                try:
                    # Search for the company
                    company_data = self.scraper.search_company(search_query)
                    
                    # Apply city prioritization if applicable
                    priority_city = self.priority_city.get()
                    if company_data and priority_city != "All" and 'phones' in company_data and company_data['phones']:
                        # Get all phone numbers
                        all_phones = company_data['phones']
                        mobile_phones = []
                        office_phones = []
                        
                        # Sort phone numbers by type without altering their format
                        for phone in all_phones:
                            # Preserve exact phone format - DO NOT MODIFY THE NUMBER
                            original_phone = phone
                            
                            # Check if it's a mobile number (starting with 08 or +628 or 628)
                            if (original_phone.startswith('08') or 
                                original_phone.startswith('+628') or 
                                original_phone.startswith('628')):
                                mobile_phones.append(original_phone)
                            # Check if it's an office/landline number (starting with 021, 022, etc.)
                            elif (original_phone.startswith('021') or 
                                  original_phone.startswith('(021)') or
                                  original_phone.startswith('022') or 
                                  original_phone.startswith('(022)') or
                                  original_phone.startswith('031') or
                                  original_phone.startswith('(031)') or
                                  (original_phone.startswith('0') and not original_phone.startswith('08'))):
                                office_phones.append(original_phone)
                            else:
                                # Check if it has + prefix but isn't a mobile number
                                if original_phone.startswith('+') and not original_phone.startswith('+628'):
                                    office_phones.append(original_phone)
                                else:
                                    # Default to mobile if we can't determine
                                    mobile_phones.append(original_phone)
                        
                        # Store mobile and office phones with their original formats
                        if mobile_phones:
                            self.data.at[i, 'Mobile Phone'] = '; '.join(mobile_phones[:5])
                            self.log(f"Found {len(mobile_phones)} mobile phone numbers")
                        
                        if office_phones:
                            self.data.at[i, 'Office Phone'] = '; '.join(office_phones[:5])
                            self.log(f"Found {len(office_phones)} office phone numbers")
                    
                    # Handle single phone if available
                    elif company_data and company_data.get('phone'):
                        # Preserve exact phone format - DO NOT MODIFY THE NUMBER
                        original_phone = company_data.get('phone')
                        
                        # Check if it's a mobile number
                        if (original_phone.startswith('08') or 
                            original_phone.startswith('+628') or 
                            original_phone.startswith('628')):
                            self.data.at[i, 'Mobile Phone'] = original_phone
                            self.log(f"Found mobile phone: {original_phone}")
                        # Check if it's an office/landline number
                        elif (original_phone.startswith('021') or 
                              original_phone.startswith('(021)') or
                              original_phone.startswith('022') or 
                              original_phone.startswith('(022)') or
                              original_phone.startswith('031') or
                              original_phone.startswith('(031)') or
                              (original_phone.startswith('0') and not original_phone.startswith('08'))):
                            self.data.at[i, 'Office Phone'] = original_phone
                            self.log(f"Found office phone: {original_phone}")
                        else:
                            # Check if it has + prefix but isn't a mobile number
                            if original_phone.startswith('+') and not original_phone.startswith('+628'):
                                self.data.at[i, 'Office Phone'] = original_phone
                            else:
                                # Default to mobile if we can't determine
                                self.data.at[i, 'Mobile Phone'] = original_phone
                            self.log(f"Found phone: {original_phone}")
                    
                    if company_data:
                        # Initialize default values
                        data_source = company_data.get('data_source', 'Google Maps')
                        source_url = "N/A"  # Initialize source_url with a default value
                        
                        # For Google Maps data, use the current_url which contains the Google Maps URL
                        # Preserve exact Google Maps URL without modification
                        if data_source == 'Google Maps' and company_data.get('current_url'):
                            source_url = company_data.get('current_url')
                            self.log(f"Found Google Maps URL: {source_url}")
                        
                        # Update data source in the monitoring dashboard
                        self.current_source_var.set(data_source)
                        
                        # Website data
                        if company_data.get('website'):
                            self.data.at[i, 'Website'] = company_data.get('website')
                            self.log(f"Found website: {company_data.get('website')}")
                            # For Website data sources, use the website URL as source
                            if data_source == "Website" and source_url == "N/A":
                                source_url = company_data.get('website')
                        
                        if company_data.get('rating'):
                            self.data.at[i, 'Rating'] = company_data.get('rating')
                            self.log(f"Found rating: {company_data.get('rating')}")
                        
                        if company_data.get('reviews_count'):
                            self.data.at[i, 'Reviews Count'] = company_data.get('reviews_count')
                            self.log(f"Found reviews: {company_data.get('reviews_count')}")
                        
                        self.data.at[i, 'Category'] = company_data.get('category')
                        
                        # Record phone source information
                        if company_data.get('phone_source'):
                            phone_source = company_data.get('phone_source')
                            self.data.at[i, 'Phone Source'] = phone_source
                            self.log(f"Phone source: {phone_source}")
                            
                            # Update data source if phone came from website
                            if isinstance(phone_source, str) and 'website' in phone_source.lower():
                                data_source = "Website"
                                if company_data.get('website') and source_url == "N/A":
                                    source_url = company_data.get('website')
                                
                        # Record email source information
                        if company_data.get('email_source'):
                            email_source = company_data.get('email_source')
                            self.data.at[i, 'Email Source'] = email_source
                            self.log(f"Email source: {email_source}")
                            
                            # Update data source if email came from website
                            if isinstance(email_source, dict) and email_source.get('page') == 'website':
                                data_source = "Website"
                                if email_source.get('url') and source_url == "N/A":
                                    source_url = email_source.get('url')
                                elif company_data.get('website') and source_url == "N/A":
                                    source_url = company_data.get('website')
                            elif isinstance(email_source, str) and 'website' in email_source.lower():
                                data_source = "Website"
                                if company_data.get('website') and source_url == "N/A":
                                    source_url = company_data.get('website')
                        
                        # Extract email from website field if not found elsewhere
                        if not self.data.at[i, 'Email']:
                            website = company_data.get('website')
                            if website and '@' in website:
                                self.data.at[i, 'Email'] = website
                            elif company_data.get('email'):
                                self.data.at[i, 'Email'] = company_data.get('email')
                        
                        # Store the data source information
                        self.data.at[i, 'Data Source'] = data_source
                        self.data.at[i, 'Source URL'] = source_url  # Now this will always have a value
                        self.log(f"Data source: {data_source}, URL: {source_url}")
                        
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
                        
                        self.log(f"✓ Data found for {company_name}")
                        self.scraped_data.append(company_data)
                        
                        # Update statistics on the monitoring dashboard
                        self._update_statistics(self.data)
                    else:
                        self.log(f"No data found for {company_name}")
                    
                except Exception as e:
                    self.log(f"Error processing {company_name}: {str(e)}")
                
                # Small delay to avoid overloading the server
                time.sleep(0.5)
                
                # Update tree view every 5 companies for better monitoring
                if i % 5 == 0 or i == total_companies - 1:
                    self.root.after(0, lambda: self.update_tree(self.data))
                
            # Store the processed data for saving
            self.results = self.data
            
            # Update final statistics
            self._update_statistics(self.data)
            
            # Update the tree view with final results
            self.root.after(0, lambda: self.update_tree(self.data))
            
            # Show desktop notification when scraping is complete
            if NOTIFICATIONS_AVAILABLE:
                # Collect statistics for notification
                processed = len(self.scraped_data)
                total = len(self.data) if hasattr(self, 'data') and self.data is not None else 0
                
                # Count companies with phone numbers
                found_count = 0
                if total > 0:
                    has_data = False
                    if 'Mobile Phone' in self.data.columns:
                        has_data = has_data | self.data['Mobile Phone'].notna()
                    if 'Office Phone' in self.data.columns:
                        has_data = has_data | self.data['Office Phone'].notna()
                    found_count = has_data.sum()
                
                # Show notification in a separate thread
                notification_thread = Thread(
                    target=show_desktop_notification,
                    args=(
                        "Werkudara - Scraping Selesai",
                        f"Berhasil memproses {total} perusahaan. Data kontak ditemukan untuk {found_count} perusahaan."
                    ),
                    kwargs={'log_func': self.log},
                    daemon=True
                )
                notification_thread.start()
            
            # Enable save button and update status
            self.root.after(0, lambda: self.save_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.status_var.set(f"Processed {len(self.scraped_data)} companies"))
        
        except Exception as e:
            self.log(f"Error in processing thread: {str(e)}")
            traceback.print_exc(file=sys.stdout)
            
            # Show error notification
            if NOTIFICATIONS_AVAILABLE:
                Thread(
                    target=show_desktop_notification,
                    args=("Werkudara - Error Saat Scraping", f"Terjadi kesalahan: {str(e)}"),
                    kwargs={'log_func': self.log},
                    daemon=True
                ).start()
        finally:
            # Cancel autosave timer
            if self.autosave_timer is not None:
                self.autosave_timer.cancel()
                self.autosave_timer = None
            
            # Close scraper
            if self.scraper:
                self.scraper.close()
                
            # Reset flags and buttons
            self.scraping_active = False
            self.stop_btn.config(state=tk.DISABLED)
            self.root.after(0, lambda: self.process_btn.config(state=tk.NORMAL))
    
    def _schedule_autosave(self):
        """Schedule the next autosave"""
        if self.enable_autosave.get() and self.scraping_active:
            # Convert minutes to seconds for the timer
            interval_seconds = self.autosave_interval.get() * 60
            
            # Cancel any existing timer
            if self.autosave_timer is not None:
                self.autosave_timer.cancel()
            
            # Schedule new timer
            self.autosave_timer = Timer(interval_seconds, self._perform_autosave)
            self.autosave_timer.daemon = True
            self.autosave_timer.start()
            self.log(f"Next autosave scheduled in {self.autosave_interval.get()} minutes")
    
    def _perform_autosave(self):
        """Perform the actual autosave operation"""
        if not self.scraping_active or not hasattr(self, 'data') or self.data is None or self.data.empty:
            # Reschedule if still scraping
            if self.scraping_active:
                self._schedule_autosave()
            return
        
        try:
            # Create autosave filename with timestamp
            output_file = self.output_file_var.get().strip()
            
            # Use a safe default location if no output file specified
            if not output_file:
                try:
                    # Try to use Documents folder first (most reliable with permissions)
                    documents_dir = os.path.join(os.path.expanduser("~"), "Documents")
                    if not os.path.exists(documents_dir):
                        # Fall back to application directory
                        documents_dir = os.path.dirname(os.path.abspath(__file__))
                    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    output_file = os.path.join(documents_dir, f"werkudara_autosave_{timestamp}.xlsx")
                except Exception:
                    # Absolute fallback
                    base_dir = os.path.dirname(os.path.abspath(__file__))
                    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    output_file = os.path.join(base_dir, f"autosave_results_{timestamp}.xlsx")
            else:
                # Add timestamp to filename to avoid overwriting
                base, ext = os.path.splitext(output_file)
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                output_file = f"{base}_autosave_{timestamp}{ext}"
            
            # Save the data
            self.log(f"Preparing autosave to {output_file}...")
            
            # Verify directory exists and is writable
            output_dir = os.path.dirname(output_file)
            if output_dir == '':
                output_dir = '.'
            
            # Make sure the directory exists
            try:
                if not os.path.exists(output_dir):
                    self.log(f"Creating directory for autosave: {output_dir}")
                    os.makedirs(output_dir, exist_ok=True)
            except Exception as dir_err:
                # If can't create directory, try Documents folder instead
                self.log(f"Failed to create directory {output_dir}: {str(dir_err)}")
                documents_dir = os.path.join(os.path.expanduser("~"), "Documents")
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                output_file = os.path.join(documents_dir, f"werkudara_autosave_{timestamp}.xlsx")
                self.log(f"Using alternate location: {output_file}")
                
                # Make sure Documents exists
                if not os.path.exists(documents_dir):
                    os.makedirs(documents_dir, exist_ok=True)
            
            # Create a copy of the data for exporting
            export_data = self.data.copy()
            
            # Final cleanup of any dictionary-like strings in the dataframe
            for column in export_data.columns:
                if export_data[column].dtype == 'object':  # Only process string columns
                    export_data[column] = export_data[column].apply(lambda x: self._clean_dict_string(x) if isinstance(x, str) else x)
            
            # Save to Excel with extra error handling
            try:
                # Apply some Excel styling for better readability
                with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                    export_data.to_excel(writer, index=False, sheet_name='Companies')
                    
                    # Try to access the workbook for formatting (requires openpyxl)
                    try:
                        workbook = writer.book
                        worksheet = writer.sheets['Companies']
                        
                        # Auto-adjust column widths
                        for column in worksheet.columns:
                            max_length = 0
                            column_letter = column[0].column_letter
                            for cell in column:
                                if cell.value:
                                    max_length = max(max_length, len(str(cell.value)))
                            adjusted_width = (max_length + 2) * 1.2
                            worksheet.column_dimensions[column_letter].width = min(adjusted_width, 50)
                    except Exception as format_error:
                        self.log(f"Warning: Could not apply Excel formatting: {str(format_error)}")
                
                self.log(f"Autosave completed successfully to {output_file}")
                self.last_autosave_time = datetime.datetime.now()
                
                # Try to create a "Latest" version too
                try:
                    base, ext = os.path.splitext(output_file)
                    base_dir = os.path.dirname(base)
                    base_name = os.path.basename(base).split("_autosave_")[0]
                    latest_file = os.path.join(base_dir, f"{base_name}_latest{ext}")
                    shutil.copy2(output_file, latest_file)
                    self.log(f"Updated latest file: {latest_file}")
                except Exception as latest_err:
                    self.log(f"Warning: Could not create latest version: {str(latest_err)}")
                
                # Show desktop notification
                if NOTIFICATIONS_AVAILABLE:
                    show_desktop_notification(
                        self.get_text("autosave_success"),
                        f"{self.get_text('autosave_success_msg')} {output_file}",
                        log_func=self.log
                    )
                
            except PermissionError as perm_e:
                # Handle permission errors specifically
                self.log(f"Permission error during autosave: {str(perm_e)}")
                # Try one more time with Documents folder
                try:
                    docs_dir = os.path.join(os.path.expanduser("~"), "Documents")
                    if not os.path.exists(docs_dir):
                        os.makedirs(docs_dir, exist_ok=True)
                    fallback_file = os.path.join(docs_dir, f"werkudara_autosave_emergency_{timestamp}.xlsx")
                    self.log(f"Trying emergency autosave to: {fallback_file}")
                    
                    # Simple save without formatting to maximize chance of success
                    export_data.to_excel(fallback_file, index=False)
                    self.log(f"Emergency autosave successful to: {fallback_file}")
                    
                    if NOTIFICATIONS_AVAILABLE:
                        show_desktop_notification(
                            "Emergency Autosave",
                            f"Data saved to: {fallback_file}",
                            log_func=self.log
                        )
                    self.last_autosave_time = datetime.datetime.now()
                except Exception as emerg_e:
                    self.log(f"Emergency autosave also failed: {str(emerg_e)}")
            except Exception as e:
                self.log(f"Error during autosave: {str(e)}")
                
        except Exception as e:
            self.log(f"Error in autosave: {str(e)}")
        finally:
            # Schedule next autosave if still scraping
            if self.scraping_active:
                self._schedule_autosave()
    
    def save_results(self):
        if not hasattr(self, 'data') or self.data is None or self.data.empty:
            messagebox.showerror("Error", "No data to save")
            return
        
        output_file = self.output_file_var.get().strip()
        if not output_file:
            messagebox.showerror("Error", "Please specify an output Excel file")
            return
        
        try:
            # Check if file exists and handle accordingly
            if os.path.exists(output_file):
                response = messagebox.askyesnocancel(
                    "File Exists",
                    f"The file '{output_file}' already exists.\n\nDo you want to replace it?\n\n"
                    "Yes: Replace existing file\n"
                    "No: Choose a new filename\n"
                    "Cancel: Abort saving"
                )
                
                if response is None:  # Cancel
                    return
                elif response is False:  # No, choose new filename
                    new_file = filedialog.asksaveasfilename(
                        title="Save As",
                        defaultextension=".xlsx",
                        initialfile=os.path.basename(output_file),
                        filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
                    )
                    if not new_file:
                        return  # User canceled file selection
                    output_file = new_file
            
            self.log(f"Saving results to {output_file}...")
            
            # Create a copy of the data for exporting
            export_data = self.data.copy()
            
            # Final cleanup of any dictionary-like strings in the dataframe
            for column in export_data.columns:
                if export_data[column].dtype == 'object':  # Only process string columns
                    export_data[column] = export_data[column].apply(lambda x: self._clean_dict_string(x) if isinstance(x, str) else x)
            
            # Check for write permissions first
            try:
                # First check if we can write to the directory
                output_dir = os.path.dirname(output_file)
                if output_dir == '':
                    output_dir = '.'
                
                # Check if directory exists
                if not os.path.exists(output_dir):
                    self.log(f"Creating directory: {output_dir}")
                    try:
                        os.makedirs(output_dir, exist_ok=True)
                    except Exception as dir_err:
                        raise PermissionError(f"Cannot create directory {output_dir}: {str(dir_err)}")
                        
                # Test if we can write to the directory by creating a test file
                test_file = os.path.join(output_dir, '.test_write_permission')
                try:
                    with open(test_file, 'w') as f:
                        f.write('test')
                    os.remove(test_file)
                except Exception as perm_err:
                    raise PermissionError(f"Cannot write to directory {output_dir}. Try saving to a different location or run the application as administrator.")
                
                # Check if the output file is locked (if it exists)
                if os.path.exists(output_file):
                    try:
                        with open(output_file, 'a') as _:
                            pass  # Just testing if we can open it for writing
                    except Exception as lock_err:
                        raise PermissionError(f"Cannot access file {output_file}. The file might be in use by another program or requires admin privileges.")
                
                # Now try to save the Excel file
                self.log("Verifying permissions successful, saving Excel file...")
                with pd.ExcelWriter(output_file, engine='openpyxl', mode='w') as writer:
                    export_data.to_excel(writer, index=False, sheet_name='Companies')
                    
                    # Try to access the workbook for formatting (requires openpyxl)
                    try:
                        workbook = writer.book
                        worksheet = writer.sheets['Companies']
                        
                        # Auto-adjust column widths
                        for column in worksheet.columns:
                            max_length = 0
                            column_letter = column[0].column_letter
                            for cell in column:
                                if cell.value:
                                    max_length = max(max_length, len(str(cell.value)))
                            adjusted_width = (max_length + 2) * 1.2
                            worksheet.column_dimensions[column_letter].width = min(adjusted_width, 50)
                    except Exception as format_error:
                        self.log(f"Warning: Could not apply Excel formatting: {str(format_error)}")
                
                self.log(f"Data saved successfully to {output_file}")
                messagebox.showinfo("Success", f"Data saved successfully to {output_file}")
                
            except PermissionError as perm_e:
                # Special handling for permission errors with more helpful message
                error_msg = str(perm_e)
                self.log(f"Permission error: {error_msg}")
                messagebox.showerror("Permission Error", 
                                   f"{error_msg}\n\nTry saving to a different location such as Documents folder, or run the application as administrator.")
            
            except Exception as e:
                messagebox.showerror("Error", f"Error saving Excel file: {str(e)}")
                self.log(f"Error saving Excel file: {str(e)}")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error preparing data for export: {str(e)}")
            self.log(f"Error preparing data for export: {str(e)}")
            traceback.print_exc(file=sys.stdout)
            
    def _clean_dict_string(self, value):
        """Clean dictionary-like strings to extract useful values"""
        if not isinstance(value, str):
            return value
            
        # If the value looks like a dictionary string
        if value.startswith('{') and value.endswith('}'):
            try:
                # Try to parse it as a dictionary
                import ast
                parsed_dict = ast.literal_eval(value)
                
                # Extract the most meaningful values
                if 'type' in parsed_dict:
                    return parsed_dict['type']
                elif 'page' in parsed_dict:
                    # Special case for known_data
                    if parsed_dict['page'] == 'known_data':
                        # For phone source that's known_data, return a cleaner description
                        return 'known_data'
                    else:
                        return parsed_dict['page']
                elif 'url' in parsed_dict:
                    return parsed_dict['url']
                else:
                    # Return first value if nothing else matches
                    return next(iter(parsed_dict.values()), value)
            except:
                # If parsing fails, clean up the string manually
                value = value.replace("{", "").replace("}", "")
                value = value.replace("'page':", "").replace("'url':", "").replace("'type':", "")
                value = value.replace("'", "").replace(",", " ").strip()
                return value
                
        return value
    
    def on_closing(self):
        if self.scraping_active:
            if messagebox.askyesno("Confirm Exit", "A scraping task is in progress. Are you sure you want to exit?"):
                # Cancel autosave timer
                if self.autosave_timer is not None:
                    self.autosave_timer.cancel()
                    self.autosave_timer = None
                
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

    def show_notification(self, title, message, timeout=10):
        """Menampilkan notifikasi desktop jika tersedia"""
        if NOTIFICATIONS_AVAILABLE:
            try:
                # Get application icon path
                icon_path = None
                if os.path.exists('app_icon.ico'):
                    icon_path = os.path.abspath('app_icon.ico')
                elif os.path.exists('important_files/app_icon.ico'):
                    icon_path = os.path.abspath('important_files/app_icon.ico')
                
                notification.notify(
                    title=title,
                    message=message,
                    app_name="Werkudara Maps Scraper",
                    timeout=timeout,
                    app_icon=icon_path
                )
                self.log(f"Notifikasi desktop ditampilkan: {title}")
            except Exception as e:
                self.log(f"Gagal menampilkan notifikasi: {str(e)}")
        else:
            self.log("Notifikasi desktop tidak tersedia. Pasang library 'plyer' untuk mengaktifkan fitur ini.")

    def _update_statistics(self, data=None):
        """Update the monitoring dashboard statistics"""
        if not hasattr(self, 'data') or self.data is None:
            return
            
        # Update total companies count
        total = len(self.data)
        self.total_companies_var.set(str(total))
        
        # If no data provided, only update the total
        if data is None:
            return
            
        # Update processed count (any company with either phone or address)
        # Safely check if columns exist first
        has_mobile_phone = 'Mobile Phone' in data.columns
        has_office_phone = 'Office Phone' in data.columns
        has_email = 'Email' in data.columns
        has_website = 'Website' in data.columns
        has_source = 'Data Source' in data.columns
        
        # Count companies with any data
        processed = 0
        if data is not None and len(data) > 0:
            processed_mask = False
            if has_mobile_phone:
                processed_mask = processed_mask | data['Mobile Phone'].notna()
            if has_office_phone:
                processed_mask = processed_mask | data['Office Phone'].notna()
            if has_email:
                processed_mask = processed_mask | data['Email'].notna()
            if has_website:
                processed_mask = processed_mask | data['Website'].notna()
            
            processed = processed_mask.sum()
        
        percent_processed = (processed / total * 100) if total > 0 else 0
        self.processed_var.set(f"{processed} ({percent_processed:.1f}%)")
        
        # Update Google Maps and Website data counts
        gmaps_count = 0
        website_count = 0
        if has_source:
            gmaps_count = data[data['Data Source'] == 'Google Maps'].shape[0]
            website_count = data[data['Data Source'] == 'Website'].shape[0]
        
            self.gmaps_count_var.set(str(gmaps_count))
            self.website_count_var.set(str(website_count))
        
        # Update phone counts
        mobile_count = 0
        office_count = 0
        if has_mobile_phone:
            mobile_count = data['Mobile Phone'].notna().sum()
        if has_office_phone:
            office_count = data['Office Phone'].notna().sum()
        
        self.mobile_phone_count_var.set(str(mobile_count))
        self.office_phone_count_var.set(str(office_count))
        
        # Update email count - now without percentage
        email_count = 0
        if has_email:
            email_count = data['Email'].notna().sum()
        
        # Set email count without percentage
        self.email_found_var.set(str(email_count))
        
        # Force GUI update
        self.root.update_idletasks()

    def show_fullscreen_preview(self):
        """Show data preview in fullscreen mode (buka GUI baru tanpa menutup monitoring)"""
        # Selalu buat window baru untuk menghindari konflik dengan window sebelumnya
        # dan memastikan monitoring tetap aktif
        if hasattr(self, 'fullscreen_window') and self.fullscreen_window:
            try:
                self.fullscreen_window.destroy()  # Hapus window lama jika ada
            except:
                pass  # Abaikan error jika window sudah ditutup
        
        # Create new fullscreen window
        self.fullscreen_window = tk.Toplevel(self.root)
        self.fullscreen_window.title("Data Preview - Werkudara Maps Scraper")
        
        # Set ukuran window yang lebih besar
        window_width = 1280
        window_height = 900
        self.fullscreen_window.geometry(f"{window_width}x{window_height}")
        
        # Posisikan window di tengah layar
        screen_width = self.fullscreen_window.winfo_screenwidth()
        screen_height = self.fullscreen_window.winfo_screenheight()
        x_position = int((screen_width - window_width) / 2)
        y_position = int((screen_height - window_height) / 2)
        self.fullscreen_window.geometry(f"+{x_position}+{y_position}")
        
        # Pastikan window ini tidak memblokir aplikasi utama
        self.fullscreen_window.transient(self.root)  # Set sebagai anak dari root window
        self.fullscreen_window.grab_set()  # Terima input tapi tidak block monitoring
        
        # Create frame for the treeview
        frame = ttk.Frame(self.fullscreen_window, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Create toolbar with filter options
        toolbar = ttk.Frame(frame)
        toolbar.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(toolbar, text="Quick Filter:").pack(side=tk.LEFT, padx=5)
        filter_entry = ttk.Entry(toolbar, width=30)
        filter_entry.pack(side=tk.LEFT, padx=5)
        
        # Add search button
        search_btn = ttk.Button(toolbar, text="Search", 
                              command=lambda: self._filter_fullscreen_data(filter_entry.get()))
        search_btn.pack(side=tk.LEFT, padx=5)
        
        # Add clear filter button
        clear_btn = ttk.Button(toolbar, text="Clear Filter", 
                             command=lambda: self._reset_fullscreen_filter(filter_entry))
        clear_btn.pack(side=tk.LEFT, padx=5)
        
        # Create a frame to contain the treeview and scrollbars
        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create performant treeview that only renders visible rows
        self.fullscreen_tree = ttk.Treeview(tree_frame)
        
        # Add scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.fullscreen_tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.fullscreen_tree.xview)
        
        # Configure treeview to use scrollbars
        self.fullscreen_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # Grid layout for treeview and scrollbars
        self.fullscreen_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        # Configure the tree frame grid to expand properly
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)
        
        # Set column headings from the main treeview
        if hasattr(self, 'data') and self.data is not None and not self.data.empty:
            ordered_cols = list(self.data.columns)
            self.fullscreen_tree["columns"] = ordered_cols
            self.fullscreen_tree["show"] = "headings"
            
            # Set headings with appropriate widths
            for col in ordered_cols:
                self.fullscreen_tree.heading(col, text=col, anchor=tk.CENTER)
                width = 150  # Default width
                
                # Adjust width based on column type
                if col in ['No.']:
                    width = 50
                elif col in ['Nama Perusahaan', 'Alamat', 'Website', 'Source URL']:
                    width = 250
                elif col in ['Mobile Phone', 'Office Phone', 'Email']:
                    width = 180
                
                self.fullscreen_tree.column(col, width=width, minwidth=50, anchor=tk.W)
            
            # Add data to the treeview
            self._populate_fullscreen_tree()
            
        # Add close button at bottom
        close_frame = ttk.Frame(frame)
        close_frame.pack(fill=tk.X, pady=10)
        close_btn = ttk.Button(close_frame, text="Close Fullscreen View", 
                              command=self.restore_normal_view)
        close_btn.pack(side=tk.RIGHT, padx=10)
        
        # Set protocol for window close
        self.fullscreen_window.protocol("WM_DELETE_WINDOW", self.restore_normal_view)
    
    def restore_normal_view(self):
        """Close fullscreen view window"""
        if hasattr(self, 'fullscreen_window') and self.fullscreen_window:
            try:
                self.fullscreen_window.destroy()  # Benar-benar tutup window, bukan sekedar sembunyi
                self.fullscreen_window = None  # Reset referensi
            except:
                pass  # Abaikan error jika window sudah ditutup
    
    def _populate_fullscreen_tree(self):
        """Efficiently populate the fullscreen treeview"""
        if not hasattr(self, 'fullscreen_tree') or not hasattr(self, 'data') or self.data is None:
            return
            
        # Clear existing items
        for item in self.fullscreen_tree.get_children():
            self.fullscreen_tree.delete(item)
            
        # Use batch processing for better performance
        batch_size = 100
        total_rows = len(self.data)
        
        for start_idx in range(0, total_rows, batch_size):
            end_idx = min(start_idx + batch_size, total_rows)
            batch = self.data.iloc[start_idx:end_idx]
            
            # Add batch to treeview
            for i, row in batch.iterrows():
                values = [row[col] if not pd.isna(row[col]) else "" for col in self.fullscreen_tree["columns"]]
                
                # Determine row tags for styling
                tags = []
                tags.append('evenrow' if i % 2 == 0 else 'oddrow')
                
                # Apply conditional formatting
                if 'Data Source' in self.data.columns and row['Data Source'] is not None:
                    if 'google maps' in str(row['Data Source']).lower():
                        tags.append('gmaps_data')
                    elif 'website' in str(row['Data Source']).lower():
                        tags.append('website_data')
                
                # Check phone columns for text coloring
                if 'Mobile Phone' in self.data.columns and row['Mobile Phone'] is not None and str(row['Mobile Phone']) != "":
                    tags.append('mobile_phone')
                
                if 'Office Phone' in self.data.columns and row['Office Phone'] is not None and str(row['Office Phone']) != "":
                    tags.append('office_phone')
                
                self.fullscreen_tree.insert("", "end", values=values, tags=tuple(tags))
            
            # Allow GUI to update after each batch
            self.fullscreen_window.update_idletasks()
        
        # Configure row styling
        self.fullscreen_tree.tag_configure('oddrow', background='#f0f0f0')
        self.fullscreen_tree.tag_configure('evenrow', background='white')
        self.fullscreen_tree.tag_configure('gmaps_data', background='#e6f2ff')
        self.fullscreen_tree.tag_configure('website_data', background='#e6ffe6')
        self.fullscreen_tree.tag_configure('mobile_phone', foreground='#006600')
        self.fullscreen_tree.tag_configure('office_phone', foreground='#000066')
    
    def _filter_fullscreen_data(self, filter_text):
        """Filter the fullscreen treeview data based on search text"""
        if not hasattr(self, 'fullscreen_tree') or not self.fullscreen_tree:
            return
            
        if not filter_text.strip():
            self._populate_fullscreen_tree()
            return
            
        # Clear existing items
        for item in self.fullscreen_tree.get_children():
            self.fullscreen_tree.delete(item)
            
        # Filter data
        filter_text = filter_text.lower()
        
        for i, row in self.data.iterrows():
            # Check if any column contains the filter text
            row_matches = False
            row_values = []
            
            for col in self.fullscreen_tree["columns"]:
                cell_value = str(row[col]) if not pd.isna(row[col]) else ""
                row_values.append(cell_value)
                
                if filter_text in cell_value.lower():
                    row_matches = True
            
            # Add matching row
            if row_matches:
                tags = []
                tags.append('evenrow' if i % 2 == 0 else 'oddrow')
                
                # Apply conditional formatting (same as in _populate_fullscreen_tree)
                if 'Data Source' in self.data.columns and row['Data Source'] is not None:
                    if 'google maps' in str(row['Data Source']).lower():
                        tags.append('gmaps_data')
                    elif 'website' in str(row['Data Source']).lower():
                        tags.append('website_data')
                
                # Check phone columns for text coloring
                if 'Mobile Phone' in self.data.columns and row['Mobile Phone'] is not None and str(row['Mobile Phone']) != "":
                    tags.append('mobile_phone')
                
                if 'Office Phone' in self.data.columns and row['Office Phone'] is not None and str(row['Office Phone']) != "":
                    tags.append('office_phone')
                    
                self.fullscreen_tree.insert("", "end", values=row_values, tags=tuple(tags))
    
    def _reset_fullscreen_filter(self, filter_entry):
        """Clear filter and restore all data"""
        if hasattr(self, 'fullscreen_tree') and self.fullscreen_tree:
            filter_entry.delete(0, tk.END)
            self._populate_fullscreen_tree()

    def _set_sash_position(self, paned_window):
        """Set sash position to create 50:50 split"""
        try:
            # Dapatkan lebar total window
            total_width = self.root.winfo_width()
            # Posisikan sash tepat di tengah
            center_position = total_width // 2
            
            # Jika window terlalu kecil, coba lagi nanti
            if total_width < 100:
                # Jadwalkan ulang setelah 200ms
                self.root.after(200, lambda: self._set_sash_position(paned_window))
                return
                
            # Set sash position
            paned_window.sashpos(0, center_position)
            print(f"Set sash position to center: {center_position}px")
            
            # Schedule another check to ensure it sticks
            self.root.after(500, lambda: self._ensure_sash_position(paned_window, center_position))
        except Exception as e:
            print(f"Error setting sash position: {str(e)}")
            # Try again after a delay
            self.root.after(300, lambda: self._set_sash_position(paned_window))
    
    def _ensure_sash_position(self, paned_window, desired_position):
        """Make sure the sash stays at the desired position"""
        try:
            current_pos = paned_window.sashpos(0)
            # If position is way off (more than 50px difference), fix it
            if abs(current_pos - desired_position) > 50:
                paned_window.sashpos(0, desired_position)
                print(f"Re-adjusted sash position to: {desired_position}px")
        except Exception as e:
            print(f"Error in _ensure_sash_position: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelProcessorGUI(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop() 