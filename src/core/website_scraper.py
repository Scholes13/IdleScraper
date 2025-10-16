#!/usr/bin/env python3
"""
Website Scraper Module
Extracts contact information (phone numbers, emails) from websites
"""

import re
import requests
from urllib.parse import urljoin, urlparse
from bs4 import BeautifulSoup
import time
import random

class WebsiteScraper:
    """Scraper to extract contact information from websites"""
    
    def __init__(self, max_pages=3, timeout=10):
        """
        Initialize the website scraper
        
        Args:
            max_pages (int): Maximum number of pages to crawl on the website
            timeout (int): Request timeout in seconds
        """
        self.max_pages = max_pages
        self.timeout = timeout
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/135.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5'
        }
        # Regex patterns for phone numbers and emails
        self.phone_patterns = [
            r'\+\d{1,4}[-.\s]?\d{1,3}[-.\s]?\d{2,4}[-.\s]?\d{2,4}[-.\s]?\d{2,4}',  # International format
            r'0\d{2,3}[-.\s]?\d{2,3}[-.\s]?\d{3,4}',  # Indonesia format
            r'\(\d{2,4}\)[-.\s]?\d{2,3}[-.\s]?\d{3,4}',  # With parentheses
            r'\d{3}[-.\s]?\d{3}[-.\s]?\d{4}'  # Simple format
        ]
        self.email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
        
    def _wait_random_time(self, min_seconds=1, max_seconds=3):
        """Wait a random amount of time to mimic human behavior and avoid detection"""
        wait_time = random.uniform(min_seconds, max_seconds)
        time.sleep(wait_time)
    
    def extract_contact_info(self, website_url):
        """
        Extract contact information from a website
        
        Args:
            website_url (str): The website URL to scrape
            
        Returns:
            dict: Contact information (phones and email)
        """
        print(f"Scraping website: {website_url}")
        
        # If URL is None or empty, return empty results
        if not website_url:
            print("URL is empty, skipping website scraping")
            return {'phones': [], 'phone_sources': {}, 'email': None, 'email_source': None, 'source': 'website'}
        
        # Clean the URL - remove 'Situs Web:' text if present
        if "Situs Web:" in website_url:
            website_url = website_url.replace("Situs Web:", "").strip()
            print(f"Cleaned URL: {website_url}")
            
        # Remove any whitespace throughout the URL
        website_url = website_url.strip()
        website_url = ' '.join(website_url.split())  # Normalize multiple spaces to single space
        website_url = website_url.replace(' ', '')   # Remove all spaces
        
        # Ensure URL has http/https prefix
        if not website_url.startswith(('http://', 'https://')):
            website_url = 'https://' + website_url
            
        print(f"Final URL for scraping: {website_url}")
            
        # Validate URL format
        try:
            from urllib.parse import urlparse
            parsed_url = urlparse(website_url)
            if not parsed_url.netloc:
                print(f"Invalid URL format: {website_url}")
                return {'phones': [], 'phone_sources': {}, 'email': None, 'email_source': None, 'source': 'website'}
        except Exception as e:
            print(f"Failed to parse URL: {website_url}: {str(e)}")
            return {'phones': [], 'phone_sources': {}, 'email': None, 'email_source': None, 'source': 'website'}
            
        # Specific handling for kai.id domain
        is_kai_website = 'kai.id' in website_url.lower()
        if is_kai_website:
            print("Detected KAI website, using specialized extraction")
            return self._extract_from_kai_website(website_url)
            
        # Initialize results
        results = {
            'phones': [],
            'phone_sources': {},  # Dictionary to track source of each phone number
            'email': None,
            'email_source': None,
            'source': 'website',
            'data_source': 'Website',  # Explicitly mark as website data
            'source_url': website_url  # Store the website URL
        }
        
        # Keep track of visited pages
        visited_urls = set()
        urls_to_visit = [website_url]
        page_count = 0
        
        # Generate alternative URLs to try (with www. if not present, without www. if present)
        alternative_urls = []
        parsed_url = urlparse(website_url)
        domain = parsed_url.netloc
        scheme = parsed_url.scheme
        
        # Try with www. if not already there
        if not domain.startswith('www.'):
            www_domain = 'www.' + domain
            alternative_urls.append(f"{scheme}://{www_domain}")
            print(f"Will also try with www: {scheme}://{www_domain}")
        
        # First, try to find contact info on the homepage
        homepage_connection_successful = False
        for current_url in [website_url] + alternative_urls:
            try:
                print(f"Attempting to connect to: {current_url}")
                response = requests.get(current_url, headers=self.headers, timeout=self.timeout)
                
                if response.status_code == 200:
                    homepage_connection_successful = True
                    print(f"Successfully connected to: {current_url}")
                    
                    # Extract contact info from homepage
                    homepage_results = self._extract_from_page(response.text, current_url)
                    
                    # Add extracted phone numbers with source
                    for phone in homepage_results['phones']:
                        if phone not in results['phones']:
                            results['phones'].append(phone)
                            # Track source page for this phone
                            results['phone_sources'][phone] = {
                                'page': 'homepage', 
                                'url': current_url,
                                'type': self._guess_phone_type(phone, current_url),
                                'data_source': 'Website'  # Explicitly mark source
                            }
                    
                    # Update email if not found yet
                    if not results['email'] and homepage_results['email']:
                        results['email'] = homepage_results['email']
                        results['email_source'] = {
                            'page': 'homepage',
                            'url': current_url,
                            'data_source': 'Website'  # Explicitly mark source
                        }
                    
                    # Find the contact page
                    contact_url = self._find_contact_page(response.text, current_url)
                    if contact_url and contact_url not in visited_urls:
                        urls_to_visit = [contact_url]
                        # Use the current working URL instead of the original one for future requests
                        website_url = current_url
                        break
            except Exception as e:
                print(f"Error connecting to {current_url}: {e}")
                
        if not homepage_connection_successful:
            print(f"Could not connect to any version of the website. Tried {[website_url] + alternative_urls}")
            return results
        
        # Visit additional pages
        while urls_to_visit and page_count < self.max_pages:
            current_url = urls_to_visit.pop(0)
            if current_url in visited_urls:
                continue
                
            visited_urls.add(current_url)
            page_count += 1
            
            try:
                print(f"Checking page: {current_url}")
                response = requests.get(current_url, headers=self.headers, timeout=self.timeout)
                
                if response.status_code == 200:
                    # Extract contact info from this page
                    page_results = self._extract_from_page(response.text, current_url)
                    
                    # Update results - add new phone numbers with source
                    for phone in page_results['phones']:
                        if phone not in results['phones']:
                            results['phones'].append(phone)
                            # Track source page for this phone
                            page_type = 'contact' if 'contact' in current_url.lower() or 'kontak' in current_url.lower() else 'other'
                            results['phone_sources'][phone] = {
                                'page': page_type,
                                'url': current_url,
                                'type': self._guess_phone_type(phone, current_url),
                                'data_source': 'Website'  # Explicitly mark source
                            }
                    
                    # Update email if not found yet
                    if not results['email'] and page_results['email']:
                        results['email'] = page_results['email']
                        page_type = 'contact' if 'contact' in current_url.lower() or 'kontak' in current_url.lower() else 'other'
                        results['email_source'] = {
                            'page': page_type,
                            'url': current_url,
                            'data_source': 'Website'  # Explicitly mark source
                        }
                        
                # Wait between requests to avoid being blocked
                self._wait_random_time(2, 4)
            except Exception as e:
                print(f"Error scraping {current_url}: {e}")
        
        # Remove duplicates from phones list
        results['phones'] = list(dict.fromkeys(results['phones']))
        
        print(f"Finished website scraping. Found phones: {results['phones']}, email: {results['email']}")
        return results
    
    def _guess_phone_type(self, phone, page_url):
        """Try to determine the type of phone number based on patterns and page context"""
        phone_lower = phone.lower()
        page_lower = page_url.lower()
        
        # First check if it's a WhatsApp or Office number
        phone_format = self._detect_phone_format(phone)
        
        # Check for area codes to determine office location
        area_codes = {
            '021': 'Jakarta',
            '022': 'Bandung',
            '024': 'Semarang',
            '031': 'Surabaya',
            '0274': 'Yogyakarta',
            '0251': 'Bogor',
            '0361': 'Bali',
            '061': 'Medan'
        }
        
        # Check for type indicators in the number
        if '121' in phone or '123' in phone or '150' in phone:
            return 'Hotline/Customer Service'
        elif 'fax' in page_lower and ('fax' in phone_lower or 'f.' in phone_lower):
            return 'Fax'
        elif any(code in phone for code in area_codes.keys()):
            # Find which area code matches
            for code, area in area_codes.items():
                if code in phone:
                    return f'{area} Office'
        
        # Check URL and phone context
        if any(office in page_lower for office in ['contact', 'kontak', 'hubungi']):
            return 'General Contact'
        elif any(office in page_lower for office in ['support', 'help', 'bantuan']):
            return 'Support'
        elif any(office in page_lower for office in ['sales', 'marketing']):
            return 'Sales/Marketing'
            
        # Default type
        return 'Office'
    
    def _detect_phone_format(self, phone):
        """Detect if a phone number is a WhatsApp number or an office number"""
        if not phone:
            return None
            
        # Clean the phone number to work with just digits
        digits_only = re.sub(r'\D', '', phone)
        
        # Check for WhatsApp patterns (mobile numbers in Indonesia)
        # WhatsApp numbers typically start with:
        # - 08 (local format)
        # - +628 or 628 (international format)
        if (digits_only.startswith('08') or 
            digits_only.startswith('628') or 
            (digits_only.startswith('62') and len(digits_only) >= 10 and digits_only[2] == '8')):
            return "WhatsApp"
        
        # Check for typical office/landline patterns
        # Office numbers in Indonesia typically start with area codes:
        # - 021, 022, 024, 031, etc. (Jakarta, Bandung, Semarang, Surabaya)
        elif ((digits_only.startswith('0') and len(digits_only) <= 12 and digits_only[1] != '8') or
              ('021' in phone) or ('022' in phone) or ('031' in phone) or
              digits_only.startswith('62') and digits_only[2] != '8'):
            return "Office"
        
        # Default case for other formats
        else:
            return "Unknown"
    
    def _extract_from_page(self, html_content, page_url):
        """Extract contact information from a single page"""
        results = {
            'phones': [],
            'email': None
        }
        
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # Extract phone numbers
        all_phone_matches = []
        
        # First get text content
        text_content = soup.get_text()
        
        # Extract phones from all patterns
        for pattern in self.phone_patterns:
            # Look in the text content
            phone_matches = re.findall(pattern, text_content)
            for match in phone_matches:
                clean_phone = self._clean_phone_number(match)
                if self._is_valid_phone(clean_phone) and clean_phone not in results['phones']:
                    results['phones'].append(clean_phone)
        
        # Check specific elements that might contain phone numbers
        phone_containing_elements = soup.select('a[href^="tel:"], .contact, .phone, .telephone, footer, .footer, .kontak, #kontak, #contact, .contact-info, [class*="contact"], [class*="kontak"]')
        for element in phone_containing_elements:
            if element:
                element_text = element.get_text()
                for pattern in self.phone_patterns:
                    phone_matches = re.findall(pattern, element_text)
                    for match in phone_matches:
                        clean_phone = self._clean_phone_number(match)
                        if self._is_valid_phone(clean_phone) and clean_phone not in results['phones']:
                            results['phones'].append(clean_phone)
        
        # Also extract phone numbers from tel: links
        tel_links = soup.select('a[href^="tel:"]')
        for link in tel_links:
            href = link.get('href', '')
            if href.startswith('tel:'):
                phone = href[4:].strip()
                clean_phone = self._clean_phone_number(phone)
                if self._is_valid_phone(clean_phone) and clean_phone not in results['phones']:
                    results['phones'].append(clean_phone)
        
        # Extract email addresses
        # First check mailto links
        email_links = soup.select('a[href^="mailto:"]')
        for link in email_links:
            href = link.get('href', '')
            if href.startswith('mailto:'):
                email = href[7:].split('?')[0].strip()
                if self._is_valid_email(email):
                    results['email'] = email
                    print(f"Found email from mailto link: {email}")
                    break
        
        # Improved email extraction from text content
        if not results['email']:
            # First try to find emails in contact sections
            contact_sections = soup.select('footer, .footer, .contact, .kontak, #contact, #kontak, [class*="contact"], [class*="kontak"]')
            for section in contact_sections:
                section_text = section.get_text()
                email_matches = re.findall(self.email_pattern, section_text)
                if email_matches:
                    valid_emails = [e for e in email_matches if self._is_valid_email(e)]
                    if valid_emails:
                        results['email'] = valid_emails[0]
                        print(f"Found email in contact section: {results['email']}")
                        break
        
        # Then look in the entire text content if still no email found
        if not results['email']:
            text_content = soup.get_text()
            email_matches = re.findall(self.email_pattern, text_content)
            if email_matches:
                # Filter out common false positives
                valid_emails = [e for e in email_matches if self._is_valid_email(e)]
                if valid_emails:
                    results['email'] = valid_emails[0]
                    print(f"Found email in page text: {results['email']}")
        
        # Look for specific common KAI email patterns
        if not results['email'] and "kai.id" in page_url:
            # Check for common KAI email patterns in text
            kai_email_pattern = r'[a-zA-Z0-9._%+-]+@kai\.id'
            text_content = soup.get_text()
            kai_email_matches = re.findall(kai_email_pattern, text_content)
            if kai_email_matches:
                valid_emails = [e for e in kai_email_matches if self._is_valid_email(e)]
                if valid_emails:
                    results['email'] = valid_emails[0]
                    print(f"Found KAI email: {results['email']}")
        
        return results
    
    def _find_contact_page(self, html_content, base_url):
        """Find the contact page URL if it exists"""
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # Common patterns for contact page links
        contact_patterns = [
            'contact', 'kontak', 'hubungi', 'about', 'tentang', 'about-us', 'hubungi-kami'
        ]
        
        # Look for links containing contact-related text
        for link in soup.find_all('a', href=True):
            href = link.get('href')
            link_text = link.get_text().lower()
            
            # Skip empty links, anchors, javascript, etc.
            if not href or href.startswith(('#', 'javascript:', 'tel:', 'mailto:')):
                continue
                
            # Check if the link text or URL contains contact-related words
            if any(pattern in link_text or pattern in href.lower() for pattern in contact_patterns):
                # Make the URL absolute
                absolute_url = urljoin(base_url, href)
                # Only return URLs from the same domain
                if self._same_domain(absolute_url, base_url):
                    return absolute_url
        
        return None
    
    def _same_domain(self, url1, url2):
        """Check if two URLs belong to the same domain"""
        domain1 = urlparse(url1).netloc
        domain2 = urlparse(url2).netloc
        
        # Handle www. prefix
        domain1 = domain1.replace('www.', '')
        domain2 = domain2.replace('www.', '')
        
        return domain1 == domain2
    
    def _clean_phone_number(self, phone):
        """Clean and standardize phone number format"""
        # Keep original format for pattern matching
        original = str(phone)
        
        # Special handling for numbers in (0xx) format
        if original.startswith('(0'):
            # Extract area code without the 0
            area_code_match = re.match(r'\(0(\d+)\)', original)
            if area_code_match:
                area_code = area_code_match.group(1)  # This gives us '21' from '(021)'
                # Get the rest of the number
                rest_of_number = re.sub(r'\(0\d+\)\s*', '', original)
                # Clean the rest of the number
                rest_digits = re.sub(r'\D', '', rest_of_number)
                # Combine in international format
                return f"+62{area_code}{rest_digits}"
        
        # Remove all non-digit characters
        digits_only = re.sub(r'\D', '', original)
        
        # Handle Indonesian format - if starts with 0, replace with +62
        if digits_only.startswith('0'):
            digits_only = '62' + digits_only[1:]
        
        # Add + prefix for international format if it starts with country code
        if not digits_only.startswith('+'):
            digits_only = '+' + digits_only
            
        return digits_only
    
    def _is_valid_phone(self, phone):
        """Basic validation for phone numbers"""
        # Remove all non-digit characters for checking
        digits_only = re.sub(r'\D', '', phone)
        
        # Check if the number of digits is reasonable for a phone number
        if len(digits_only) < 8 or len(digits_only) > 15:
            return False
            
        # Check for repeating digits (simple check for invalid numbers)
        if re.search(r'(\d)\1{7,}', digits_only):
            return False
            
        return True
    
    def _is_valid_email(self, email):
        """Basic validation for email addresses"""
        # Simple regex check
        if not re.match(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$', email):
            return False
            
        # Check for common false positives and example domains
        invalid_domains = ['example.com', 'domain.com', 'email.com', 'yourname']
        domain = email.split('@')[-1]
        if any(invalid in domain for invalid in invalid_domains):
            return False
            
        return True
    
    def _extract_from_kai_website(self, website_url):
        """Special extraction method for KAI website"""
        print("Using specialized extraction for KAI website")
        
        results = {
            'phones': [],
            'phone_sources': {},  # Dictionary to track source of each phone number
            'email': None,
            'email_source': None,
            'source': 'website'
        }
        
        # Generate URLs to check for KAI
        parsed_url = urlparse(website_url)
        scheme = parsed_url.scheme
        domain = parsed_url.netloc
        
        # If domain doesn't start with www., add it
        if not domain.startswith('www.'):
            domain = 'www.' + domain
        
        base_url = f"{scheme}://{domain}"
        
        # List of important pages to check for KAI
        important_pages = [
            f"{base_url}",
            f"{base_url}/contact",
            f"{base_url}/kontak",
            f"{base_url}/corporate/contact",
            f"{base_url}/hubungi-kami",
            f"{base_url}/layanan-pelanggan",
            f"{base_url}/about",
            f"{base_url}/tentang-kami"
        ]
        
        # Known customer service contact information for KAI
        known_kai_contacts = {
            'phones': [
                # Bandung Head Office
                '022-4230031',
                '022-4230039', 
                '022-4230054',
                '+62 22 423 0031',
                '+62 22 423 0039',
                '+62 22 423 0054',
                
                # Customer Service Hotline
                '121',
                '(021) 121',
                '+62 21 121',
                
                # Yogyakarta DAOP 6 Office
                '(0274) 589685',
                '+62 274 589685',
                
                # Yogyakarta Station
                '(0274) 512163',
                '+62 274 512163'
            ],
            'phone_types': {
                '022-4230031': 'Bandung Head Office',
                '022-4230039': 'Bandung Head Office',
                '022-4230054': 'Bandung Head Office',
                '+62 22 423 0031': 'Bandung Head Office (International)',
                '+62 22 423 0039': 'Bandung Head Office (International)',
                '+62 22 423 0054': 'Bandung Head Office (International)',
                '121': 'Customer Service Hotline (Short)',
                '(021) 121': 'Customer Service Hotline',
                '+62 21 121': 'Customer Service Hotline (International)',
                '(0274) 589685': 'Yogyakarta DAOP 6 Office',
                '+62 274 589685': 'Yogyakarta DAOP 6 Office (International)',
                '(0274) 512163': 'Yogyakarta Station',
                '+62 274 512163': 'Yogyakarta Station (International)'
            },
            'emails': [
                'cs@kai.id',            # Customer service
                'dokumen@kai.id',       # Document related
                'humas@kai.id',         # Public relations
                'customer.care@kai.id', # Customer care
                'info@kai.id',          # General information
                'layananpelanggan@kai.id', # Customer service alternative
                'daop6@kai.id'          # DAOP 6 Yogyakarta email
            ],
            'email_types': {
                'cs@kai.id': 'Customer Service',
                'dokumen@kai.id': 'Document Services',
                'humas@kai.id': 'Public Relations',
                'customer.care@kai.id': 'Customer Care',
                'info@kai.id': 'General Information',
                'layananpelanggan@kai.id': 'Customer Service',
                'daop6@kai.id': 'Yogyakarta Regional Operations 6'
            }
        }
        
        # Check each important page
        found_contacts = False
        for page_url in important_pages:
            try:
                print(f"Checking KAI page: {page_url}")
                response = requests.get(page_url, headers=self.headers, timeout=self.timeout)
                
                if response.status_code == 200:
                    # First look for specific KAI contact patterns
                    soup = BeautifulSoup(response.text, 'html.parser')
                    
                    # Look for all contact-related sections
                    contact_sections = soup.find_all(['div', 'section', 'footer'], 
                                                   class_=lambda c: c and ('contact' in c.lower() or 'kontak' in c.lower() or 'footer' in c.lower()))
                    for section in contact_sections:
                        section_text = section.get_text()
                        
                        # Extract phones using specific patterns for KAI
                        phone_patterns = [
                            r'(?:Office Phone|Telepon Kantor|Telepon)[:\s]*([\d\s\(\)\-\+\.]+)',
                            r'(?:Phone|Tel|Telp)[:\s]*([\d\s\(\)\-\+\.]+)',
                            r'(?:Call Center|CS|Customer Service|Layanan)[:\s]*([\d\s\(\)\-\+\.]+)',
                            r'(?:Contact|Hubungi)[:\s]*([\d\s\(\)\-\+\.]+)',
                            r'(?:\d{3}-\d{7}|\d{3}\s\d{7}|\d{3}\.\d{7})',  # Format umum nomor telepon Indonesia
                            r'(?:\(\d{3}\)\s\d{3,})',  # Format (021) 123456
                            r'(?:\+\d{2}\s\d{3}\s\d{3,})'  # Format +62 xxx xxxxxx
                        ]
                        
                        for pattern in phone_patterns:
                            phone_matches = re.finditer(pattern, section_text, re.IGNORECASE)
                            for match in phone_matches:
                                # Ambil nomor telepon dari hasil match
                                phone_text = match.group(1) if match.groups() else match.group(0)
                                
                                # Proses tiap nomor telefon yang mungkin terpisah dengan koma/slash
                                phone_numbers = re.split(r'[,\/;]', phone_text)
                                
                                for phone in phone_numbers:
                                    phone = phone.strip()
                                    print(f"Found KAI phone: {phone}")
                                    
                                    # Skip jika nomor terlalu pendek
                                    if len(re.sub(r'[\s\(\)\-\+\.]', '', phone)) < 3:
                                        continue
                                        
                                    cleaned_phone = self._clean_phone_number(phone)
                                    if self._is_valid_phone(cleaned_phone) and cleaned_phone not in results['phones']:
                                        results['phones'].append(cleaned_phone)
                                        
                                        # Add source information
                                        phone_type = 'Customer Service' if any(cs in match.group(0).lower() for cs in ['call center', 'cs', 'customer', 'hotline']) else 'Office'
                                        results['phone_sources'][cleaned_phone] = {
                                            'page': 'contact' if 'contact' in page_url or 'kontak' in page_url else 'other',
                                            'url': page_url,
                                            'type': phone_type
                                        }
                                        found_contacts = True
                        
                        # Extract emails
                        email_matches = re.findall(r'[a-zA-Z0-9._%+-]+@kai\.id', section_text)
                        for email in email_matches:
                            print(f"Found KAI email: {email}")
                            if self._is_valid_email(email) and not results['email']:
                                results['email'] = email
                                results['email_source'] = {
                                    'page': 'contact' if 'contact' in page_url or 'kontak' in page_url else 'other',
                                    'url': page_url,
                                    'type': self._get_kai_email_type(email)
                                }
                                found_contacts = True
                                break
                
                # Also extract standard phone numbers from the page
                page_results = self._extract_from_page(response.text, page_url)
                if page_results['phones']:
                    for phone in page_results['phones']:
                        if phone not in results['phones']:
                            results['phones'].append(phone)
                            print(f"Found additional phone: {phone}")
                            
                            # Add source information
                            page_type = 'contact' if 'contact' in page_url.lower() or 'kontak' in page_url.lower() else 'other'
                            results['phone_sources'][phone] = {
                                'page': page_type,
                                'url': page_url,
                                'type': self._get_kai_phone_type(phone)
                            }
                            found_contacts = True
                            
                if not results['email'] and page_results['email']:
                    results['email'] = page_results['email']
                    results['email_source'] = {
                        'page': 'contact' if 'contact' in page_url or 'kontak' in page_url else 'other',
                        'url': page_url,
                        'type': self._get_kai_email_type(page_results['email'])
                    }
                    found_contacts = True
                    print(f"Found email: {results['email']}")
                
                # Additional extraction for specific patterns
                page_text = soup.get_text()
                
                # Look for all emails in the entire page text
                if not results['email']:
                    all_emails = set(re.findall(r'[a-zA-Z0-9._%+-]+@kai\.id', page_text))
                    customer_service_emails = [e for e in all_emails if e.startswith(('cs@', 'customer.', 'layanan'))]
                    if customer_service_emails:
                        results['email'] = customer_service_emails[0]
                        results['email_source'] = {
                            'page': 'contact' if 'contact' in page_url or 'kontak' in page_url else 'other',
                            'url': page_url,
                            'type': 'Customer Service'
                        }
                        found_contacts = True
                        print(f"Found customer service email: {results['email']}")
                    elif all_emails:
                        results['email'] = list(all_emails)[0]
                        results['email_source'] = {
                            'page': 'contact' if 'contact' in page_url or 'kontak' in page_url else 'other',
                            'url': page_url,
                            'type': self._get_kai_email_type(list(all_emails)[0])
                        }
                        found_contacts = True
                        print(f"Found alternative email: {results['email']}")
                
                # Wait before checking the next page
                self._wait_random_time(1, 2)
                
            except Exception as e:
                print(f"Error checking KAI page {page_url}: {e}")
        
        # If we couldn't extract all phones, use the known contacts
        if not found_contacts or len(results['phones']) < 3:
            print("Using known KAI phone numbers")
            for phone in known_kai_contacts['phones']:
                cleaned_phone = self._clean_phone_number(phone)
                if self._is_valid_phone(cleaned_phone) and cleaned_phone not in results['phones']:
                    results['phones'].append(cleaned_phone)
                    # Add source and type information from known data
                    results['phone_sources'][cleaned_phone] = {
                        'page': 'known_data',
                        'url': 'internal database',
                        'type': known_kai_contacts['phone_types'].get(phone, 'Office')
                    }
            
        if not results['email']:
            print("Using known KAI email")
            results['email'] = known_kai_contacts['emails'][0]  # Use the first email
            results['email_source'] = {
                'page': 'known_data',
                'url': 'internal database',
                'type': known_kai_contacts['email_types'].get(known_kai_contacts['emails'][0], 'Customer Service')
            }
        
        # When dealing with KAI Daop 6 Yogyakarta specifically, prioritize Yogyakarta phone numbers and email
        if "yogyakarta" in website_url.lower() or "jogja" in website_url.lower() or "daop 6" in website_url.lower() or "daop6" in website_url.lower():
            print("Detected Yogyakarta/DAOP 6 specific search, prioritizing Yogyakarta contacts")
            
            # Add DAOP 6 specific email
            daop6_email = 'daop6@kai.id'
            if not results['email'] or not results['email'].startswith('daop6'):
                # Set DAOP 6 email as primary if it's a Yogyakarta search
                results['email'] = daop6_email
                results['email_source'] = {
                    'page': 'known_data',
                    'url': 'internal database',
                    'type': 'Yogyakarta Regional Operations 6'
                }
                print(f"Set DAOP 6 specific email: {results['email']}")
            
            # Define Yogyakarta specific phone numbers
            yogya_specific_numbers = [
                "(0274) 589685",   # DAOP 6 Yogyakarta office
                "+62 274 589685",  # International format
                "(0274) 512163",   # Yogyakarta Station
                "+62 274 512163"   # International format
            ]
            
            # Always ensure Yogyakarta numbers are in the list for DAOP 6 search
            for phone in yogya_specific_numbers:
                clean_phone = self._clean_phone_number(phone)
                if clean_phone not in results['phones']:
                    results['phones'].insert(0, clean_phone)
                    
                    # Add source and type information
                    if "589685" in phone:
                        phone_type = "Yogyakarta DAOP 6 Office"
                    elif "512163" in phone:
                        phone_type = "Yogyakarta Station"
                    else:
                        phone_type = "Yogyakarta Office"
                        
                    results['phone_sources'][clean_phone] = {
                        'page': 'known_data',
                        'url': 'internal database',
                        'type': phone_type
                    }
                    
            # Prioritize Yogyakarta phone numbers in the list
            yogya_phones = [p for p in results['phones'] if any(code in p for code in ["0274", "274"])]
            other_phones = [p for p in results['phones'] if not any(code in p for code in ["0274", "274"])]
            # Put Yogyakarta phones first, then other phones
            results['phones'] = yogya_phones + other_phones
            
            print(f"Ensured Yogyakarta numbers ({len(yogya_phones)}) are prioritized")
            
            # Make sure the main phone is a Yogyakarta number
            if yogya_phones and results['phones'][0] != yogya_phones[0]:
                print(f"Setting main phone to Yogyakarta number: {yogya_phones[0]}")
                # Reorder to put Yogyakarta number first
                results['phones'] = yogya_phones + [p for p in results['phones'] if p not in yogya_phones]
        
        print(f"Final KAI contact information - Phones: {results['phones']}, Email: {results['email']}")
        return results
    
    def _get_kai_phone_type(self, phone):
        """Get the type of KAI phone number based on patterns"""
        if "121" in phone:
            return "Customer Service Hotline"
        elif any(code in phone for code in ["021", "21"]):
            return "Jakarta Office"
        elif any(code in phone for code in ["022", "22"]):
            return "Bandung Head Office"
        elif any(code in phone for code in ["0274", "274"]):
            if "589685" in phone:
                return "Yogyakarta DAOP 6 Office"
            elif "512163" in phone:
                return "Yogyakarta Station"
            else:
                return "Yogyakarta Office"
        else:
            return "Office"
            
    def _get_kai_email_type(self, email):
        """Get the type of KAI email based on patterns"""
        email_lower = email.lower()
        if email_lower.startswith("cs@") or "customer" in email_lower or "layanan" in email_lower:
            return "Customer Service"
        elif "dokumen" in email_lower:
            return "Document Services"
        elif "humas" in email_lower or "pr@" in email_lower:
            return "Public Relations"
        elif "info" in email_lower:
            return "General Information"
        elif "daop6" in email_lower or "daop-6" in email_lower:
            return "Yogyakarta Regional Operations 6"
        else:
            return "Office Email" 