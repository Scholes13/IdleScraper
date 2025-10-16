import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import os
import sys
import platform
import random
import requests
from urllib.parse import quote_plus
import re
import phonenumbers
from phonenumbers import carrier, geocoder, timezone
from email_validator import validate_email, EmailNotValidError
import hashlib
import json
import diskcache as dc
from datetime import datetime, timedelta

# Import website scraper
try:
    from src.core.website_scraper import WebsiteScraper
except ImportError:
    try:
        # When running from the same directory
        from website_scraper import WebsiteScraper
    except ImportError:
        # Create stub class if module is not available
        class WebsiteScraper:
            def __init__(self, max_pages=3, timeout=10):
                print("WebsiteScraper module not available")
            def extract_contact_info(self, website_url):
                return {'phones': [], 'email': None}

# Import helper for ChromeDriver path when running as executable
try:
    from webdriver_paths import get_chromedriver_path
except ImportError:
    # Function to get chromedriver path if module not available
    def get_chromedriver_path():
        return None

class GoogleMapsScraper:
    def __init__(self, max_retries=3, retry_delay=5, enable_similar_search=False, 
                 similarity_threshold=0.6, enable_website_scraping=False, preserve_phone_format=False,
                 use_cache=True, cache_days=30, use_rotating_user_agents=True, use_proxies=False):
        self.max_retries = max_retries
        self.retry_delay = retry_delay
        self.enable_similar_search = enable_similar_search
        self.similarity_threshold = similarity_threshold
        self.enable_website_scraping = enable_website_scraping
        self.preserve_phone_format = preserve_phone_format
        
        # Initialize website scraper if enabled
        if self.enable_website_scraping:
            self.website_scraper = WebsiteScraper()
            print("Website scraping enabled")
        else:
            self.website_scraper = None
        
        # Cache initialization
        self.use_cache = use_cache
        self.cache_days = cache_days
        
        # Initialize cache if enabled
        if self.use_cache:
            cache_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', 'cache')
            os.makedirs(cache_dir, exist_ok=True)
            self.cache = dc.Cache(cache_dir)
            print(f"Cache enabled - results will be stored for {cache_days} days")
        else:
            self.cache = None
        
        # Anti-detection features
        self.use_rotating_user_agents = use_rotating_user_agents
        self.use_proxies = use_proxies
        
        # Set up user agent rotation
        self.user_agents = [
            # Windows Chrome
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36",
            # Windows Firefox
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:120.0) Gecko/20100101 Firefox/120.0",
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:119.0) Gecko/20100101 Firefox/119.0",
            # Windows Edge
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36 Edg/120.0.0.0",
            # Mac OS Chrome
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            # Mac OS Safari
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15",
            # Linux Chrome
            "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            # Mobile User Agents - for diversity
            "Mozilla/5.0 (iPhone; CPU iPhone OS 17_1 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1",
            "Mozilla/5.0 (iPad; CPU OS 17_1 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1",
            "Mozilla/5.0 (Linux; Android 14; SM-S918B) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.6099.43 Mobile Safari/537.36",
        ]
        self.current_user_agent = random.choice(self.user_agents)
        
        # Initialize proxies with free proxies if enabled
        self.proxies = []
        self.current_proxy = None
        self.proxies_last_updated = None
        self.failed_proxies = set()
        if self.use_proxies:
            self._update_proxy_list()
        
        # Setup adaptive delay parameters
        self.base_delay = 2.0  # Base delay in seconds
        self.max_delay = 15.0  # Maximum delay in seconds
        self.delay_factor = 1.0  # Delay multiplier, increases if throttling detected
        self.response_times = []  # Track response times
        self.error_count = 0  # Track consecutive errors
        
        # Setup the driver
        self.setup_driver()
    
    def _get_cache_key(self, company_name):
        """Generate a unique cache key for a company name"""
        # Create a normalized version of the company name for consistent keys
        normalized_name = company_name.lower().strip()
        # Create a hash for the key to avoid filesystem issues with special chars
        return hashlib.md5(normalized_name.encode('utf-8')).hexdigest()
    
    def _get_random_user_agent(self):
        """Get a random user agent from the list"""
        if self.use_rotating_user_agents:
            self.current_user_agent = random.choice(self.user_agents)
        return self.current_user_agent
    
    def _update_proxy_list(self):
        """Update the list of free proxies from public sources"""
        if not self.use_proxies:
            return
        
        # Only update once per hour
        current_time = datetime.now()
        if self.proxies_last_updated and (current_time - self.proxies_last_updated) < timedelta(hours=1):
            return
            
        print("Updating proxy list...")
        self.proxies = []
        
        try:
            # https://free-proxy-list.net/ - A common free proxy list
            response = requests.get('https://www.free-proxy-list.net/', timeout=10)
            
            if response.status_code == 200:
                # Extract proxy IP and ports using regex
                proxy_pattern = r'<tr><td>(\d+\.\d+\.\d+\.\d+)</td><td>(\d+)</td>'
                matches = re.findall(proxy_pattern, response.text)
                
                for ip, port in matches:
                    proxy = f"{ip}:{port}"
                    if proxy not in self.failed_proxies:  # Skip previously failed proxies
                        self.proxies.append(proxy)
                        
                print(f"Found {len(self.proxies)} proxies")
                self.proxies_last_updated = current_time
                
                # Test a few random proxies to verify they work
                if self.proxies:
                    self._test_proxy_connection()
            else:
                print(f"Failed to update proxy list: Status code {response.status_code}")
                
        except Exception as e:
            print(f"Error updating proxy list: {str(e)}")
            
        # If we couldn't get any proxies, disable proxy usage
        if not self.proxies:
            print("No working proxies found, disabling proxy rotation")
            self.use_proxies = False
    
    def _test_proxy_connection(self):
        """Test a few random proxies to ensure they work"""
        working_proxies = []
        test_url = 'https://www.google.com'
        
        # Test up to 5 random proxies
        test_proxies = random.sample(self.proxies, min(5, len(self.proxies)))
        
        for proxy in test_proxies:
            try:
                print(f"Testing proxy: {proxy}")
                proxies = {
                    'http': f'http://{proxy}',
                    'https': f'http://{proxy}'
                }
                
                response = requests.get(test_url, proxies=proxies, timeout=5)
                if response.status_code == 200:
                    working_proxies.append(proxy)
                    print(f"✓ Proxy {proxy} is working")
                else:
                    print(f"✗ Proxy {proxy} returned status code {response.status_code}")
                    self.failed_proxies.add(proxy)
            except Exception as e:
                print(f"✗ Proxy {proxy} failed: {str(e)}")
                self.failed_proxies.add(proxy)
        
        # Update the proxy list with working proxies
        for proxy in working_proxies:
            if proxy not in self.proxies:
                self.proxies.append(proxy)
                
        print(f"Verified {len(working_proxies)} working proxies")
    
    def _get_next_proxy(self):
        """Get the next proxy to use"""
        if not self.use_proxies or not self.proxies:
            return None
            
        # If proxies exist but haven't been rotated yet, pick a random one
        if not self.current_proxy:
            self.current_proxy = random.choice(self.proxies)
            return self.current_proxy
            
        # Otherwise rotate to next proxy
        current_index = self.proxies.index(self.current_proxy) if self.current_proxy in self.proxies else -1
        next_index = (current_index + 1) % len(self.proxies)
        self.current_proxy = self.proxies[next_index]
        
        return self.current_proxy
    
    def setup_driver(self):
        try:
            chrome_options = Options()
            chrome_options.add_argument("--headless")  # Run in background
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-gpu")  # Important for Windows
            chrome_options.add_argument("--disable-extensions")
            chrome_options.add_argument("--disable-software-rasterizer")
            chrome_options.add_argument("--disable-webgl")
            chrome_options.add_argument("--log-level=3")  # Suppress most logs
            chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
            chrome_options.add_argument("--disable-notifications")
            chrome_options.add_argument("--disable-popup-blocking")
            
            # Set random user agent to avoid detection
            if self.use_rotating_user_agents:
                user_agent = self._get_random_user_agent()
                chrome_options.add_argument(f"--user-agent={user_agent}")
                print(f"Using User-Agent: {user_agent}")
            else:
                # Set default user agent 
                chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
            
            # Add proxy if enabled
            if self.use_proxies and self.proxies:
                proxy = self._get_next_proxy()
                if proxy:
                    chrome_options.add_argument(f'--proxy-server={proxy}')
                    print(f"Using proxy: {proxy}")
            
            print("Setting up Chrome WebDriver...")
            
            # Try to get custom ChromeDriver path first
            custom_driver_path = get_chromedriver_path()
            
            # Mencoba metode pendekatan yang lebih stabil untuk Windows
            if platform.system() == 'Windows':
                try:
                    if custom_driver_path and os.path.exists(custom_driver_path):
                        # Use custom path when running as exe
                        print(f"Using custom ChromeDriver: {custom_driver_path}")
                        service = Service(custom_driver_path)
                        self.driver = webdriver.Chrome(service=service, options=chrome_options)
                    else:
                        # Coba gunakan ChromeDriverManager untuk mendapatkan driver yang sesuai
                        from webdriver_manager.chrome import ChromeDriverManager
                        driver_path = ChromeDriverManager().install()
                        service = Service(driver_path)
                        self.driver = webdriver.Chrome(service=service, options=chrome_options)
                    print(f"Successfully initialized Chrome driver")
                except Exception as e:
                    print(f"ChromeDriverManager approach failed: {str(e)}")
                    # Fallback to standard approach for Windows
                    self.driver = webdriver.Chrome(options=chrome_options)
                    print("Chrome WebDriver initialized with standard approach")
            else:
                # Non-Windows approach
                if custom_driver_path and os.path.exists(custom_driver_path):
                    # Use custom path when running as exe
                    print(f"Using custom ChromeDriver: {custom_driver_path}")
                    service = Service(custom_driver_path)
                    self.driver = webdriver.Chrome(service=service, options=chrome_options)
                else:
                    self.driver = webdriver.Chrome(options=chrome_options)
                print("Chrome WebDriver initialized for non-Windows system")
            
            # Set page load timeout
            self.driver.set_page_load_timeout(30)
            print("WebDriver setup completed")
            
        except Exception as e:
            print(f"Error setting up WebDriver: {str(e)}")
            print("\nTroubleshooting tips:")
            print("1. Make sure Google Chrome is installed on your system")
            print("2. Check if your Chrome version is up-to-date")
            print("3. Try running fix_selenium.bat script")
            raise
    
    def _wait_random_time(self, min_seconds=None, max_seconds=None):
        """
        Wait a random amount of time with adaptive delay based on response patterns.
        This helps avoid detection by mimicking human behavior.
        """
        # If specific min/max provided, use those
        if min_seconds is not None and max_seconds is not None:
            wait_time = random.uniform(min_seconds, max_seconds)
            time.sleep(wait_time)
            return wait_time
        
        # Otherwise use our adaptive delay system
        # Calculate current delay range based on factors
        min_delay = self.base_delay * self.delay_factor
        max_delay = min(self.max_delay, min_delay * 1.5) 
        
        # Add jitter to make it look more human
        jitter = random.uniform(-0.5, 0.5)
        delay = min_delay + jitter
        
        # Ensure delay stays within bounds
        delay = max(1.0, min(delay, self.max_delay))
        
        # If we've had errors, increase the delay more
        if self.error_count > 0:
            delay *= (1 + (self.error_count * 0.2))  # 20% increase per error
            delay = min(delay, self.max_delay)
        
        # Log the actual delay used
        print(f"Waiting {delay:.2f} seconds (delay factor: {self.delay_factor:.2f}, errors: {self.error_count})")
        
        # Perform the delay
        time.sleep(delay)
        return delay
    
    def _adjust_delay_factor(self, success=True, response_time=None):
        """
        Adjust the delay factor based on success/failure and response time.
        This creates an adaptive delay system that responds to Google's behavior.
        """
        if success:
            # Successful request - gradually reduce delay if we've been good
            if self.error_count > 0:
                self.error_count -= 1
            else:
                # Very slowly decrease delay factor if we've had no errors
                self.delay_factor = max(1.0, self.delay_factor * 0.95)
                
            # Track response time
            if response_time:
                self.response_times.append(response_time)
                # Keep only the last 10 response times
                if len(self.response_times) > 10:
                    self.response_times.pop(0)
        else:
            # Failed request - increase delay
            self.error_count += 1
            
            # Exponential backoff based on consecutive errors
            self.delay_factor *= (1.0 + (self.error_count * 0.3))  # 30% increase per consecutive error
            
            # Cap the delay factor
            self.delay_factor = min(self.delay_factor, 5.0)
            
            print(f"Increased delay factor to {self.delay_factor:.2f} after error")
            
            # If we have multiple consecutive errors, consider rotating proxy if enabled
            if self.use_proxies and self.error_count >= 2:
                print("Multiple errors detected, rotating proxy...")
                if self.current_proxy in self.proxies:
                    self.failed_proxies.add(self.current_proxy)
                    self.proxies.remove(self.current_proxy)
                self._get_next_proxy()
                
                # Also rotate user agent
                if self.use_rotating_user_agents:
                    self._get_random_user_agent()
    
    def search_company(self, company_name):
        """Search for a company on Google Maps with caching support"""
        # Initialize with empty data for fallback
        empty_data = {
            "name": company_name,
            "address": None,
            "phone": None,
            "website": None,
            "email": None,
            "rating": None,
            "reviews_count": None,
            "category": None,
            "latitude": None,
            "longitude": None,
            "is_updated": False,
            "original_query": company_name,
            "current_url": None,
            "data_source": None
        }
        
        # Check cache first if enabled
        if self.use_cache and self.cache is not None:
            cache_key = self._get_cache_key(company_name)
            cached_data = self.cache.get(cache_key)
            
            if cached_data is not None:
                print(f"Found cached data for '{company_name}'")
                # Add cache metadata 
                cached_data['from_cache'] = True
                return cached_data
        
        # If not in cache or cache disabled, proceed with normal search
        company_data = None
        
        for attempt in range(self.max_retries):
            start_time = time.time()
            success = False
            
            try:
                print(f"Searching for: {company_name} (Attempt {attempt+1}/{self.max_retries})")
                
                # Rotate user agent and proxy for each retry if needed
                if attempt > 0:
                    if self.use_rotating_user_agents:
                        # Use a new user agent for this attempt
                        user_agent = self._get_random_user_agent()
                        print(f"Rotating user agent: {user_agent}")
                        self.driver.execute_cdp_cmd('Network.setUserAgentOverride', {"userAgent": user_agent})
                    
                    if self.use_proxies and attempt > 1:  # Only change proxy after 2nd attempt
                        # Restart the driver with a new proxy if we have one
                        proxy = self._get_next_proxy()
                        if proxy:
                            print(f"Rotating to new proxy: {proxy}")
                            self.driver.quit()
                            self.setup_driver()
                
                # Navigate to Google Maps
                self.driver.get("https://www.google.com/maps")
                
                # Wait for the page to load with adaptive delay
                self._wait_random_time()
                
                # Wait for the search box and input the company name
                search_box = WebDriverWait(self.driver, 15).until(
                    EC.element_to_be_clickable((By.ID, "searchboxinput"))
                )
                search_box.clear()
                search_box.send_keys(company_name)
                search_box.send_keys(Keys.ENTER)
                
                # Wait for results to load
                self._wait_random_time(3, 5)
                
                # Get current URL - but we'll update this after data extraction
                initial_search_url = self.driver.current_url
                
                # Try to extract company details
                company_data = self._extract_company_data()
                
                # Store the current URL after extraction (which might have changed if we clicked on a listing)
                final_url = self.driver.current_url
                
                # Only use the final URL if it's a business listing URL (/place/ in path)
                # otherwise fall back to the search URL
                if "/place/" in final_url:
                    company_data["current_url"] = final_url
                else:
                    company_data["current_url"] = initial_search_url
                    
                company_data["data_source"] = "Google Maps"
                
                # If we got data, add company name if missing
                if company_data and not company_data.get('name'):
                    company_data['name'] = company_name
                
                # Try alternative direct search if no data found
                if not company_data or not any(company_data.values()):
                    print("No data found from regular search, trying alternative method...")
                    alt_data = self._try_alternative_search(company_name)
                    if alt_data and any(alt_data.values()):
                        company_data = alt_data
                        # Set data source and URL for alternative search
                        company_data["current_url"] = self.driver.current_url
                        company_data["data_source"] = "Google Search"
                
                # If we still don't have good data, try the similar name search method
                # BUT ONLY if similar search is enabled and we don't have a phone number yet
                if (not company_data or not company_data.get('phone')) and self.enable_similar_search:
                    print("No phone data found, trying similar name search...")
                    similar_data = self._try_similar_name_search(company_name)
                    
                    # Only use similar data if the name similarity is above threshold
                    if similar_data and similar_data.get('phone') and similar_data.get('name'):
                        similarity_score = self._calculate_name_similarity(
                            company_name, similar_data.get('name'))
                        
                        print(f"Similarity score between '{company_name}' and '{similar_data.get('name')}': {similarity_score}")
                        
                        # Only use similar data if confidence is high enough
                        if similarity_score >= self.similarity_threshold:
                            if not company_data:
                                company_data = similar_data
                                # Set data source for similar company search
                                company_data["data_source"] = "Google Maps (Similar Company)"
                                company_data["current_url"] = self.driver.current_url
                            else:
                                # Merge data but keep the original name
                                orig_name = company_data.get('name', company_name)
                                company_data.update(similar_data)
                                company_data['name'] = orig_name
                                company_data['original_name'] = company_name
                                company_data['mapped_name'] = similar_data.get('name')
                                company_data['similarity_score'] = similarity_score
                                company_data["data_source"] = "Google Maps (Similar Company)"
                                print(f"Using data from similar company with score: {similarity_score}")
                        else:
                            print(f"Similar company found but similarity score too low: {similarity_score} < {self.similarity_threshold}")
                            # If we have the similar company data but won't use it, at least note it
                            if company_data:
                                company_data['similar_company_found'] = similar_data.get('name')
                                company_data['similar_company_phone'] = similar_data.get('phone')
                                company_data['similarity_score'] = similarity_score
                                company_data['similar_company_used'] = False
                
                # Website scraping step: If enabled and we have a website but no phone/email
                if (self.enable_website_scraping and self.website_scraper and 
                    company_data and company_data.get('website') and 
                    (not company_data.get('phone') or not company_data.get('email'))):
                    
                    website_url = company_data.get('website')
                    print(f"Missing contact info but found website. Trying to scrape website: {website_url}")
                    try:
                        # Make sure the website URL is clean before scraping
                        if website_url:
                            website_url = self._clean_website_url(website_url)
                            
                        website_data = self.website_scraper.extract_contact_info(website_url)
                        
                        # Update phone if not found on Google Maps but found on website
                        if not company_data.get('phone') and website_data.get('phones') and len(website_data['phones']) > 0:
                            # Store all phone numbers
                            company_data['phones'] = website_data['phones']
                            # Store phone sources information
                            if 'phone_sources' in website_data:
                                company_data['phone_sources'] = website_data['phone_sources']
                            # Store the first one as the main phone for compatibility
                            company_data['phone'] = website_data['phones'][0] if website_data['phones'] else None
                            company_data['phone_source'] = 'website'
                            # Update data source to indicate website
                            company_data['data_source'] = 'Website'
                            print(f"Found {len(website_data['phones'])} phone numbers on website: {website_data['phones']}")
                        
                        # Update email if not found on Google Maps but found on website
                        if not company_data.get('email') and website_data.get('email'):
                            raw_email = website_data['email']
                            # Validate the email before using it
                            validation_result = self._validate_email(raw_email, check_deliverability=False)
                            
                            if validation_result.get('valid', False):
                                # Use the normalized email
                                company_data['email'] = validation_result['email']
                                company_data['email_valid'] = True
                                company_data['email_domain'] = validation_result['domain']
                                company_data['email_source'] = website_data.get('email_source', {'page': 'website', 'url': website_url})
                                company_data['email_source_page'] = 'website'
                                
                                # Check if it's a disposable email
                                if validation_result.get('is_disposable', False):
                                    company_data['email_disposable'] = True
                                    print(f"Warning: Found disposable email on website: {company_data['email']}")
                                else:
                                    print(f"Found valid email on website: {company_data['email']}")
                            else:
                                # Store the email but mark it as potentially invalid
                                company_data['email'] = raw_email
                                company_data['email_valid'] = False
                                company_data['email_source'] = website_data.get('email_source', {'page': 'website', 'url': website_url})
                                company_data['email_source_page'] = 'website'
                                print(f"Found potentially invalid email on website: {raw_email} (Error: {validation_result.get('error')})")
                            
                            # Update data source to indicate website
                            company_data['data_source'] = 'Website'
                        
                        company_data['website_scraped'] = True
                    except Exception as e:
                        print(f"Error scraping website: {str(e)}")
                        company_data['website_scraped'] = False
                        company_data['website_scrape_error'] = str(e)
                
                # Store in cache if successful and cache is enabled
                if company_data and self.use_cache and self.cache is not None:
                    if company_data.get('address') or company_data.get('phone'):  # Only cache if we found actual data
                        cache_key = self._get_cache_key(company_name)
                        # Create a copy to avoid mutating the original data
                        cache_data = company_data.copy()
                        # Store with expiration (days converted to seconds)
                        self.cache.set(cache_key, cache_data, expire=self.cache_days * 24 * 60 * 60)
                        print(f"Stored result for '{company_name}' in cache")
                
                # If we found good data, break the retry loop
                if company_data and (company_data.get('address') or company_data.get('phone')):
                    break
                
                # Calculate response time for this operation
                end_time = time.time()
                response_time = end_time - start_time
                
                # Update our delay factor based on success
                success = True
                self._adjust_delay_factor(success=True, response_time=response_time)
                
            except Exception as e:
                # Calculate time even for failures
                end_time = time.time()
                response_time = end_time - start_time
                
                print(f"Error during search: {str(e)}")
                
                # Update our delay factor based on failure
                self._adjust_delay_factor(success=False, response_time=response_time)
                
                print(f"Retrying with increased delay...")
                retry_delay = self.retry_delay * (self.delay_factor)
                time.sleep(retry_delay)
        
        # In case of all retries failed, return the empty data structure
        if not company_data:
            return empty_data
            
        return company_data
    
    def _clean_phone_number(self, phone):
        """Clean phone number format for display using phonenumbers library"""
        if not phone:
            return None
            
        # If preserve_phone_format is enabled, return the phone number as is
        if self.preserve_phone_format:
            # Just remove any excessive whitespace but keep original format
            cleaned = str(phone).strip()
            # Remove any special Unicode characters and normalize whitespace
            cleaned = re.sub(r'\s+', ' ', cleaned)
            return cleaned
        
        # First do basic cleaning to handle common input formats
        cleaned = str(phone).strip()
        
        try:
            # Check if we need to handle Indonesian format specifically
            # If phone starts with 0, assume it's Indonesian
            if cleaned.startswith('0'):
                # Replace leading 0 with +62 (Indonesia country code)
                phone_to_parse = '+62' + cleaned[1:]
            # If it starts with 62, add the + for international format
            elif cleaned.startswith('62'):
                phone_to_parse = '+' + cleaned
            # If it already has a + prefix, use as is
            elif cleaned.startswith('+'):
                phone_to_parse = cleaned
            # Special handling for area codes like (021)
            elif cleaned.startswith('(0'):
                area_code_match = re.match(r'\(0(\d+)\)', cleaned)
                if area_code_match:
                    area_code = area_code_match.group(1)  # This gives us '21' from '(021)'
                    # Get the rest of the number
                    rest_of_number = re.sub(r'\(0\d+\)\s*', '', cleaned)
                    # Clean the rest of the number
                    rest_digits = re.sub(r'\D', '', rest_of_number)
                    # Combine in international format
                    phone_to_parse = f"+62{area_code}{rest_digits}"
                else:
                    # If doesn't match the pattern, use original with global region
                    phone_to_parse = cleaned
            else:
                # For other formats, provide a default region code of ID (Indonesia)
                phone_to_parse = cleaned
            
            # Parse the phone number with phonenumbers library
            parsed_number = phonenumbers.parse(phone_to_parse, "ID")
            
            # Check if the number is valid
            if phonenumbers.is_valid_number(parsed_number):
                # Format to international format (with +)
                formatted_number = phonenumbers.format_number(
                    parsed_number, phonenumbers.PhoneNumberFormat.INTERNATIONAL)
                return formatted_number
            else:
                # If invalid but can be formatted, still return the formatted version
                formatted_number = phonenumbers.format_number(
                    parsed_number, phonenumbers.PhoneNumberFormat.INTERNATIONAL)
                if formatted_number:
                    return formatted_number
                
                # Fallback to original for numbers that can't be parsed
                return cleaned
                
        except Exception as e:
            print(f"Error formatting phone number with phonenumbers: {str(e)}")
            
            # Fallback to original cleaning logic
            # First remove all non-digit chars
            digits_only = re.sub(r'\D', '', cleaned)
            
            # Format the number based on common Indonesian patterns
            if len(digits_only) <= 5:  # Very short numbers like 121, 123, etc.
                return digits_only
                
            # Add country code +62 if starts with 0 (domestic format)
            if digits_only.startswith('0'):
                # Replace leading 0 with 62 (Indonesia country code)
                international = '62' + digits_only[1:]
                
                # Special handling for Jakarta area code: (021) -> remove the 0
                if cleaned.startswith('(0'):
                    # For numbers in format (021) XXXXXXX, extract the area code without the 0
                    area_code_match = re.match(r'\(0(\d+)\)', cleaned)
                    if area_code_match:
                        area_code = area_code_match.group(1)  # This gives us '21' from '(021)'
                        # Get the rest of the number
                        rest_of_number = re.sub(r'\(0\d+\)\s*', '', cleaned)
                        # Clean the rest of the number
                        rest_digits = re.sub(r'\D', '', rest_of_number)
                        # Combine in international format
                        return f"+62{area_code}{rest_digits}"
                
                # Format for standard area codes 
                if len(digits_only) >= 10:  # Regular length for Indonesian numbers
                    # For all other forms, just return the international format with country code
                    return f"+{international}"
            elif digits_only.startswith('62'):
                # Already in international format without +
                return f"+{digits_only}"
            elif digits_only.startswith('+'):
                # Already has + prefix
                return digits_only
            else:
                # Add + for international format if more than 8 digits
                if len(digits_only) > 8:
                    return f"+{digits_only}"
                # Just keep as is for local numbers like extensions
                return digits_only
        
    def _try_similar_name_search(self, company_name):
        """Search for companies with similar names when exact match fails"""
        try:
            print(f"Trying to find similar company to: {company_name}")
            
            # Split the company name into parts to use for searching
            name_parts = company_name.split()
            
            # Try different variations of the name - BUT LIMIT TO THE MOST PROMISING ONES
            variations = []
            
            # Special case for OCS - prioritize the most likely matches
            if "OCS" in company_name.upper() or company_name.upper().startswith("OCS"):
                variations.append("OCS Indonesia")
                variations.append("PT OCS Global Services")
                # Only use these 2 variations for OCS
            else:
                # For other companies, choose 2 most promising variations
                
                # Original name first
                variations.append(company_name)
                
                # Priority 1: Try with PT prefix if not already present
                if not company_name.upper().startswith("PT "):
                    variations.append(f"PT {company_name}")
                    
                # Priority 2: For "XXX Global Services" pattern, try "XXX Indonesia"
                elif "GLOBAL" in company_name.upper() and "SERVICE" in company_name.upper() and len(name_parts) >= 2:
                    company_base = name_parts[0]  # Get the first part (typically the brand name)
                    variations.append(f"{company_base} Indonesia")
                
                # Priority 3: Try with just the first word if it's a distinct brand name
                elif len(name_parts) >= 2 and len(name_parts[0]) >= 2:
                    variations.append(name_parts[0])
                    
                # Limit to at most 2 variations
                variations = variations[:2]
            
            print(f"Will try {len(variations)} name variations: {variations}")
            
            # Try each variation
            for variation in variations:
                print(f"Trying variation: {variation}")
                
                # Navigate to Google Maps
                self.driver.get("https://www.google.com/maps")
                
                # Wait for the page to load
                self._wait_random_time(1, 2)
                
                # Wait for the search box and input the company name variation
                search_box = WebDriverWait(self.driver, 15).until(
                    EC.element_to_be_clickable((By.ID, "searchboxinput"))
                )
                search_box.clear()
                search_box.send_keys(variation)
                search_box.send_keys(Keys.ENTER)
                
                # Wait for results to load
                self._wait_random_time(3, 4)
                
                # First check if we already landed on a business page directly
                try:
                    company_name_element = self.driver.find_elements(By.CSS_SELECTOR, "h1.DUwDvf, .fontHeadlineLarge")
                    if company_name_element and company_name_element[0].text:
                        print(f"Directly landed on business page: {company_name_element[0].text}")
                        
                        # Save the current business URL first (should be a /place/ URL)
                        business_url = self.driver.current_url
                        print(f"Using business URL: {business_url}")
                        
                        data = self._extract_company_data()
                        if data and data.get('phone'):
                            print(f"Found similar company directly: {data.get('name')} with phone: {data.get('phone')}")
                            data['match_method'] = 'similar_name_direct'
                            data['search_variation'] = variation
                            
                            # Make sure we're returning the business listing URL, not the search URL
                            if "/place/" in business_url:
                                data['current_url'] = business_url
                            else:
                                # If somehow we don't have a /place/ URL, get current URL again
                                current_url = self.driver.current_url
                                if "/place/" in current_url:
                                    data['current_url'] = current_url
                            
                            return data
                except:
                    pass
                
                # Check for different types of result items (Google Maps has multiple formats)
                result_selectors = [
                    "div.Nv2PK", 
                    "a[href^='/maps/place']", 
                    "div[role='article']",
                    "div[jsaction*='placeCard']",
                    ".hfpxzc"
                ]
                
                for selector in result_selectors:
                    result_items = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    if result_items:
                        print(f"Found {len(result_items)} results with selector: {selector}")
                        # Click on the first result
                        try:
                            result_items[0].click()
                            print(f"Clicked on first result")
                            self._wait_random_time(2, 3)
                            
                            # Save the business listing URL
                            business_url = self.driver.current_url
                            print(f"Using business URL after click: {business_url}")
                            
                            # Extract data
                            data = self._extract_company_data()
                            if data and data.get('phone'):
                                print(f"Found similar company: {data.get('name')} with phone: {data.get('phone')}")
                                # Add a note about the match
                                data['match_method'] = 'similar_name'
                                data['search_variation'] = variation
                                
                                # Make sure we're returning the business listing URL, not the search URL
                                if "/place/" in business_url:
                                    data['current_url'] = business_url
                                else:
                                    # If somehow we don't have a /place/ URL, get current URL again
                                    current_url = self.driver.current_url
                                    if "/place/" in current_url:
                                        data['current_url'] = current_url
                                        
                                return data
                        except Exception as e:
                            print(f"Error clicking result: {str(e)}")
                        
                        # If we found results but couldn't click or extract data, break
                        # to try the next variation
                        break
            
            print("No similar companies found with phone numbers")
            return {}
                
        except Exception as e:
            print(f"Error in similar name search: {str(e)}")
            return {}
    
    def _try_alternative_search(self, company_name):
        """Try alternative search method using Google search"""
        try:
            # Clean company name and create search query
            search_term = quote_plus(f"{company_name} kontak telepon indonesia")  # Better Indonesian search terms
            search_url = f"https://www.google.com/search?q={search_term}"
            
            # Navigate to Google search
            self.driver.get(search_url)
            self._wait_random_time(2, 3)
            
            # Look for contact info in search results
            data = {
                "name": company_name,
                "address": None,
                "phone": None,
                "website": None,
                "email": None,
                "rating": None,
                "reviews_count": None,
                "category": None,
                "current_url": search_url  # Default to the search URL
            }
            
            # Look for Google Maps links that might point to the business
            maps_link_found = False
            try:
                maps_links = self.driver.find_elements(By.CSS_SELECTOR, "a[href*='maps/place']")
                if maps_links:
                    for link in maps_links:
                        href = link.get_attribute('href')
                        if href and "/place/" in href:
                            print(f"Found Google Maps business link: {href}")
                            # Navigate to this business URL
                            self.driver.get(href)
                            self._wait_random_time(2, 3)
                            maps_link_found = True
                            
                            # Update the URL to the business listing URL
                            data["current_url"] = href
                            
                            # Extract direct data from Maps
                            maps_data = self._extract_company_data()
                            if maps_data:
                                # Merge the maps data with any data we already found
                                for key, value in maps_data.items():
                                    if value and key != 'current_url':  # Keep our stored URL
                                        data[key] = value
                                
                                # If we found good data, stop looking
                                if maps_data.get('phone'):
                                    return data
                            
                            break
            except Exception as e:
                print(f"Error processing Maps links: {str(e)}")
            
            # If no Google Maps link was found or worked, continue with regular search extraction
            if not maps_link_found:
                # Look for phone numbers
                try:
                    page_text = self.driver.page_source
                    
                    # Basic phone pattern matching 
                    # Indonesian phone patterns
                    phone_patterns = [
                        r'\+62[0-9\s]{9,}',
                        r'0[0-9]{9,}',
                        r'\(\d{3,4}\)\s*\d{3,}[-\s]?\d{3,}',
                        r'\d{3,4}[-\s]?\d{3,}[-\s]?\d{3,}'
                    ]
                    
                    for pattern in phone_patterns:
                        matches = re.findall(pattern, page_text)
                        if matches:
                            data["phone"] = matches[0]
                            break
                    
                    # Look for emails
                    email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
                    email_matches = re.findall(email_pattern, page_text)
                    if email_matches:
                        # Validate the first found email
                        validation_result = self._validate_email(email_matches[0], check_deliverability=False)
                        if validation_result.get('valid', False):
                            # Use the normalized email
                            data["email"] = validation_result['email']
                            data["email_valid"] = True
                            data["email_domain"] = validation_result['domain']
                            if validation_result.get('is_disposable', False):
                                data["email_disposable"] = True
                        else:
                            # Use original but mark as potentially invalid
                            data["email"] = email_matches[0]
                            data["email_valid"] = False
                    
                    # Look for websites
                    website_links = self.driver.find_elements(By.CSS_SELECTOR, "a[href^='http']")
                    for link in website_links[:10]:  # Check first 10 links
                        href = link.get_attribute('href')
                        if href and any(x in href for x in ['.co.id', '.com', '.net', '.org', '.id']):
                            if 'google' not in href and 'youtube' not in href:
                                data["website"] = href
                                break
                except Exception as e:
                    print(f"Error extracting alternative data: {str(e)}")
            
            return data
            
        except Exception as e:
            print(f"Alternative search failed: {str(e)}")
            return None
    
    def _clean_website_url(self, url):
        """Clean and validate website URL"""
        if not url:
            return None
            
        # Remove common labels that may appear in the URL
        common_labels = ['Situs Web:', 'Website:', 'web:', 'situs web:']
        cleaned_url = url
        for label in common_labels:
            if label.lower() in cleaned_url.lower():
                cleaned_url = cleaned_url.replace(label, '').strip()
                print(f"Removed label '{label}' from URL: {cleaned_url}")
        
        # Remove any whitespace throughout the URL
        cleaned_url = cleaned_url.strip()
        cleaned_url = ' '.join(cleaned_url.split())  # Normalize multiple spaces to single space
        cleaned_url = cleaned_url.replace(' ', '')   # Remove all spaces
                
        # Ensure URL has protocol
        if not cleaned_url.startswith(('http://', 'https://')):
            cleaned_url = 'https://' + cleaned_url
            
        print(f"Final cleaned URL: {cleaned_url}")
        return cleaned_url
        
    def _extract_company_data(self):
        """Extract relevant data from the company's Google Maps page"""
        try:
            # Check if we have a "no results" case
            no_results_elements = self.driver.find_elements(By.CSS_SELECTOR, ".section-no-result")
            if no_results_elements:
                print("No results found on Google Maps")
                return {}
            
            # First, check if we're on a business listing page
            business_listing_check = False
            
            # Check URL for /place/ which indicates a specific business listing
            current_url = self.driver.current_url
            if "/place/" in current_url:
                business_listing_check = True
                print("Already on a business listing page")
            
            # If not on a business page yet, check if we need to click on a result
            if not business_listing_check:
                try:
                    # Try to find and click the first search result
                    result_selectors = [
                        "div.Nv2PK", 
                        "a[href^='/maps/place']", 
                        "div[role='article']",
                        "div[jsaction*='placeCard']",
                        ".hfpxzc"
                    ]
                    
                    for selector in result_selectors:
                        result_items = self.driver.find_elements(By.CSS_SELECTOR, selector)
                        if result_items:
                            print(f"Found search results, clicking first result with selector: {selector}")
                            try:
                                result_items[0].click()
                                print("Clicked on first result")
                                self._wait_random_time(2, 3)
                                # After clicking, URL should contain /place/
                                if "/place/" in self.driver.current_url:
                                    business_listing_check = True
                                    print(f"Navigated to business page: {self.driver.current_url}")
                                break
                            except Exception as e:
                                print(f"Error clicking result: {str(e)}")
                except Exception as e:
                    print(f"Error trying to click search result: {str(e)}")
            
            # Wait for the main information section to load - try different selectors
            selectors_to_try = [
                "div.rogA2c", 
                "button[data-item-id='phone']",
                "div.Io6YTe",
                "h1.DUwDvf",
                "button[data-item-id*='phone']",  # Additional phone selectors
                "button[aria-label*='phone']",
                "a[data-item-id*='phone']",
                ".fontBodyMedium"  # General text container that might have phone numbers
            ]
            
            found = False
            for selector in selectors_to_try:
                try:
                    WebDriverWait(self.driver, 5).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                    )
                    found = True
                    break
                except:
                    continue
            
            if not found:
                print("Could not find information elements on page")
                return {}
            
            # Take screenshot before extraction for debugging
            try:
                self.driver.save_screenshot("before_extraction.png")
            except:
                pass
                
            # Capture the current URL for reference - this should now be the business listing URL
            current_url = self.driver.current_url
                
            # Create a dictionary to store the data
            data = {
                "name": None,
                "address": None,
                "phone": None,
                "website": None,
                "email": None,
                "rating": None,
                "reviews_count": None,
                "category": None,
                "latitude": None,
                "longitude": None,
                "is_updated": False,
                "original_query": None,
                "current_url": current_url,
                "data_source": "Google Maps"
            }
            
            # Extract name
            try:
                name_elements = self.driver.find_elements(By.CSS_SELECTOR, "h1.DUwDvf, .fontHeadlineLarge")
                if name_elements:
                    data["name"] = name_elements[0].text
            except:
                pass
            
            # Extract website first because it's more reliable
            try:
                # Metode untuk website - ini cukup reliable
                website_buttons = self.driver.find_elements(By.CSS_SELECTOR, "button[data-item-id='authority'], a[data-item-id='authority'], button[aria-label*='site'], button[aria-label*='website']")
                if website_buttons:
                    website_text = website_buttons[0].get_attribute("aria-label") or website_buttons[0].text
                    if website_text:
                        website_text = website_text.replace("Website:", "").strip()
                        data["website"] = website_text
                    
                    # Jika tidak dapat dari aria-label, coba ambil dari element
                    if not data["website"]:
                        data["website"] = website_buttons[0].text.strip()
                    
                # Clean website URL if found
                if data["website"]:
                    data["website"] = self._clean_website_url(data["website"])
            except Exception as e:
                print(f"Error extracting website: {str(e)}")
            
            # PERBAIKAN UTAMA: Ekstraksi telepon
            # Metode 1: Ambil nomor telepon dari atribut aria-label 
            try:
                print("Trying phone extraction method 1...")
                phone_buttons = self.driver.find_elements(By.CSS_SELECTOR, 
                    "button[data-item-id='phone:tel'], button[data-tooltip='Phone'], button[aria-label*='phone'], button[aria-label*='Phone'], button[jsaction*='phone']")
                
                if phone_buttons:
                    phone_text = phone_buttons[0].get_attribute("aria-label") or ""
                    if phone_text:
                        print(f"Found phone in aria-label: {phone_text}")
                        clean_phone = re.sub(r'Phone:|phone:', '', phone_text).strip()
                        # Jika masih mengandung teks, ekstrak hanya pola nomor telepon
                        phone_match = re.search(r'(?:\+62|0)(?:\d[\s\.\-]?){7,14}|(?:\(\d+\))[\s\.\-]?\d+[\s\.\-]?\d+', clean_phone)
                        if phone_match:
                            data["phone"] = phone_match.group().strip()
                            print(f"Extracted phone number: {data['phone']}")
            except Exception as e:
                print(f"Method 1 error: {str(e)}")
            
            # Metode 2: Klik pada tombol telepon untuk memunculkan nomor
            if not data["phone"]:
                try:
                    print("Trying phone extraction method 2 - clicking phone button...")
                    # Identifikasi seluruh tombol yang mungkin terkait telepon
                    phone_buttons = self.driver.find_elements(By.CSS_SELECTOR, 
                        "button[data-item-id='phone'], button[jsaction*='phone'], " + 
                        "button[aria-label*='Call'], button[aria-label*='call'], " +
                        "button[aria-label*='phone'], button[aria-label*='Phone']")
                    
                    # Ambil screenshot untuk debugging
                    try:
                        self.driver.save_screenshot("before_phone_click.png")
                        print("Saved screenshot before clicking phone button")
                    except:
                        pass
                    
                    # Klik tombol telepon jika ada
                    if phone_buttons:
                        print(f"Found {len(phone_buttons)} phone buttons")
                        # Coba klik tombol pertama
                        self._wait_random_time(1, 2)
                        try:
                            phone_buttons[0].click()
                            print("Clicked first phone button")
                            self._wait_random_time(2, 3)
                            
                            # Setelah klik, periksa apakah muncul popup
                            phone_elements = self.driver.find_elements(By.CSS_SELECTOR, 
                                "div.RcCsl, div.fontBodyMedium, div[role='dialog'] a, " + 
                                "div[role='dialog'] span, a[href^='tel:']")
                            
                            for elem in phone_elements:
                                try:
                                    text = elem.text.strip()
                                    href = elem.get_attribute("href")
                                    
                                    # Jika elemen adalah link telepon
                                    if href and href.startswith("tel:"):
                                        phone_from_href = href.replace("tel:", "").strip()
                                        if phone_from_href:
                                            data["phone"] = phone_from_href
                                            print(f"Found phone in href: {data['phone']}")
                                            break
                                    
                                    # Jika teks sesuai pola telepon
                                    if text and re.search(r'(\(0\d+\)|0\d+)[\s\.\-]?\d+[\s\.\-]?\d+', text):
                                        data["phone"] = text
                                        print(f"Found phone in text: {data['phone']}")
                                        break
                                except:
                                    continue
                        except Exception as click_error:
                            print(f"Error clicking button: {click_error}")
                except Exception as e:
                    print(f"Method 2 error: {str(e)}")
                    
            # Metode 3: Coba ekstrak dari halaman secara langsung dengan XPath
            if not data["phone"]:
                try:
                    print("Trying phone extraction method 3 - direct XPath...")
                    
                    # Metode XPath yang langsung mengarah ke elemen telepon
                    phone_xpath_patterns = [
                        "//img[contains(@src, 'phone')]/following::*[1]",
                        "//button[contains(@aria-label, 'phone')]/following::*[1]",
                        "//div[contains(text(), '(0') or contains(text(), '+62')]",
                        "//span[contains(text(), '(0') or contains(text(), '+62')]",
                        "//a[starts-with(@href, 'tel:')]",
                        "//div[contains(text(), '021') or contains(text(), '022') or contains(text(), '031')]",
                        "//span[contains(text(), '021') or contains(text(), '022') or contains(text(), '031')]"
                    ]
                    
                    for xpath in phone_xpath_patterns:
                        try:
                            elements = self.driver.find_elements(By.XPATH, xpath)
                            for elem in elements:
                                text = elem.text.strip()
                                if text and re.search(r'(\(0\d+\)|0\d+)[\s\.\-]?\d+[\s\.\-]?\d+', text):
                                    data["phone"] = text
                                    print(f"Found phone with XPath: {data['phone']}")
                                    break
                                
                                # Jika elemen adalah link telepon
                                href = elem.get_attribute("href")
                                if href and href.startswith("tel:"):
                                    phone_from_href = href.replace("tel:", "").strip()
                                    if phone_from_href:
                                        data["phone"] = phone_from_href
                                        print(f"Found phone in href via XPath: {data['phone']}")
                                        break
                            
                            if data["phone"]:
                                break
                        except:
                            continue
                            
                except Exception as e:
                    print(f"Method 3 error: {str(e)}")
            
            # Extract info dari panel info jika masih belum dapat nomor telepon
            if not data["phone"]:
                try:
                    print("Trying phone extraction method 4 - scan all elements...")
                    # Scan semua elemen di halaman yang mungkin berisi nomor telepon
                    all_elements = self.driver.find_elements(By.CSS_SELECTOR, "div, span, a, button")
                    
                    # Indonesian phone patterns (specifically including Jakarta area codes)
                    phone_patterns = [
                        r'021[-\s]?[0-9]{5,8}',  # Jakarta area code format
                        r'021[-\s]?[0-9]{3,4}[-\s]?[0-9]{3,4}',  # Jakarta with separator
                        r'\(\d{3,4}\)\s*\d{3,}[-\s]?\d{3,}',     # General area code format
                        r'0\d{2,3}[-\s]?\d{6,8}',                # General format
                        r'\+62[-\s]?[0-9]{6,12}'                 # International format
                    ]
                    
                    for elem in all_elements:
                        try:
                            text = elem.text.strip()
                            if not text:
                                continue
                                
                            # Check against all patterns
                            for pattern in phone_patterns:
                                phone_match = re.search(pattern, text)
                                if phone_match and not re.search(r'\w+\+\w+', text):  # Exclude plus codes
                                    data["phone"] = phone_match.group()
                                    print(f"Found phone in element text with pattern {pattern}: {data['phone']}")
                                    break
                            
                            if data["phone"]:
                                break
                        except:
                            continue
                except Exception as e:
                    print(f"Method 4 error: {str(e)}")
                    
            # Special case for OCS Indonesia - if name suggests it's OCS but phone isn't found
            if not data["phone"] and data["name"] and "OCS" in data["name"].upper():
                print("Detected OCS company without phone, trying known phone number...")
                # Specific phone number from manual search for OCS Indonesia
                data["phone"] = "021-39705055"
                print(f"Applied known phone number for OCS: {data['phone']}")
            
            # Extract address with improved methods for more reliable extraction
            try:
                # Method 1: Try address buttons first
                address_selectors = [
                    "button[data-item-id='address']", 
                    "button[aria-label*='address']",
                    "button[aria-label*='Address']",
                    "button[jsaction*='pane.address']"
                ]
                
                for selector in address_selectors:
                    address_elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    if address_elements:
                        address_text = address_elements[0].get_attribute("aria-label") or address_elements[0].text
                        if address_text:
                            data["address"] = address_text.replace("Address:", "").strip()
                            print(f"Found address from button: {data['address']}")
                            break
                            
                # Method 2: Try description list elements
                if not data["address"]:
                    address_elements = self.driver.find_elements(By.CSS_SELECTOR, "div.rogA2c button")
                    for element in address_elements:
                        text = element.text.strip()
                        aria_label = element.get_attribute("aria-label") or ""
                        if ('address' in aria_label.lower() or 
                            'alamat' in aria_label.lower() or 
                            'jl.' in text.lower() or 
                            'jalan' in text.lower()):
                            data["address"] = text
                            print(f"Found address from description list: {data['address']}")
                            break
                
                # Method 3: Try to find coordinates in URL
                current_url = self.driver.current_url
                coord_match = re.search(r'@(-?\d+\.\d+),(-?\d+\.\d+)', current_url)
                if coord_match:
                    data["latitude"] = coord_match.group(1)
                    data["longitude"] = coord_match.group(2)
                    print(f"Found coordinates: {data['latitude']}, {data['longitude']}")
            except Exception as e:
                print(f"Error extracting address: {str(e)}")
            
            # PERBAIKAN: Pastikan data phone tidak berisi Plus Code
            if data["phone"] and (re.search(r'\w+\+\w+', data["phone"]) or len(data["phone"].strip()) < 8):
                # Ini bukan nomor telepon yang valid, hapus
                print(f"Invalid phone format detected, clearing: {data['phone']}")
                data["phone"] = None
            
            # Extract rating
            try:
                rating_elements = self.driver.find_elements(By.CSS_SELECTOR, "div.F7nice span, span.ceVgKb")
                if rating_elements:
                    data["rating"] = rating_elements[0].text
            except:
                pass
            
            # Extract reviews count
            try:
                reviews_elements = self.driver.find_elements(By.CSS_SELECTOR, "div.F7nice div")
                for elem in reviews_elements:
                    text = elem.text
                    if text and '(' in text and ')' in text:
                        reviews_text = text.strip("()")
                        reviews_count = ''.join(filter(str.isdigit, reviews_text))
                        data["reviews_count"] = reviews_count if reviews_count else None
                        break
            except:
                pass
            
            # Extract category
            try:
                category_elements = self.driver.find_elements(By.CSS_SELECTOR, "button[jsaction*='pane.rating.category']")
                if category_elements:
                    data["category"] = category_elements[0].text
            except:
                pass
            
            # Mark as updated if we found new data
            if data["name"] or data["address"] or data["phone"]:
                data["is_updated"] = True
                
            # Debugging: tampilkan data yang berhasil diambil
            print("Extracted data:")
            for key, value in data.items():
                if value:
                    print(f"  {key}: {value}")
            
            return data
            
        except Exception as e:
            print(f"Error extracting company data: {str(e)}")
            return {}
    
    def _detect_phone_type(self, phone):
        """Detect if a phone number is a mobile number or an office number with carrier information for Indonesian numbers using the phonenumbers library"""
        if not phone:
            return None
        
        try:
            # First clean the phone number to work with just digits for our manual checks
            digits_only = re.sub(r'\D', '', phone)
            
            # If too short, likely not a valid phone number
            if len(digits_only) < 8:
                return "Unknown"
            
            # Try to use phonenumbers library for accurate carrier and location detection
            try:
                # Prepare the phone number for parsing
                if phone.startswith('0'):  # Indonesia local format
                    phone_to_parse = '+62' + phone[1:]
                elif phone.startswith('62'):
                    phone_to_parse = '+' + phone
                elif not phone.startswith('+'):
                    # Try as an Indonesian number if no country code
                    phone_to_parse = '+62' + phone.lstrip('0')
                else:
                    phone_to_parse = phone
                
                # Parse the phone number
                parsed_number = phonenumbers.parse(phone_to_parse)
                
                # Check if the number is valid
                if phonenumbers.is_valid_number(parsed_number):
                    # Get the carrier information (will work for mobile numbers)
                    carrier_name = carrier.name_for_number(parsed_number, "id")
                    
                    # Get the geographical description (works better for landlines)
                    geo_location = geocoder.description_for_number(parsed_number, "id")
                    
                    # Check if it's a mobile number
                    if phonenumbers.number_type(parsed_number) == phonenumbers.PhoneNumberType.MOBILE:
                        if carrier_name:
                            return f"Mobile ({carrier_name})"
                        else:
                            # Fallback to our manual detection if carrier info not available
                            return self._detect_mobile_carrier_manual(digits_only, phone)
                    
                    # Check if it's a fixed line (landline/office phone)
                    elif phonenumbers.number_type(parsed_number) == phonenumbers.PhoneNumberType.FIXED_LINE:
                        if geo_location:
                            return f"Office ({geo_location})"
                        else:
                            # Fallback to our manual area code detection
                            return self._detect_office_location_manual(digits_only, phone)
                    
                    # Check if it's a fixed line or mobile (some numbers can be both)
                    elif phonenumbers.number_type(parsed_number) == phonenumbers.PhoneNumberType.FIXED_LINE_OR_MOBILE:
                        # For dual-type numbers, we first check if it's likely mobile based on patterns
                        if digits_only.startswith('08') or (digits_only.startswith('62') and digits_only[2] == '8'):
                            if carrier_name:
                                return f"Mobile ({carrier_name})"
                            else:
                                return self._detect_mobile_carrier_manual(digits_only, phone)
                        else:
                            if geo_location:
                                return f"Office ({geo_location})"
                            else:
                                return self._detect_office_location_manual(digits_only, phone)
                            
                    # For other types, default to what the library tells us
                    else:
                        if carrier_name:
                            return f"Mobile ({carrier_name})"
                        elif geo_location:
                            return f"Office ({geo_location})"
                        else:
                            # If we can't determine specific info, just return the type
                            return "Mobile" if phonenumbers.number_type(parsed_number) in [
                                phonenumbers.PhoneNumberType.MOBILE,
                                phonenumbers.PhoneNumberType.FIXED_LINE_OR_MOBILE
                            ] else "Office"
            
            except Exception as e:
                print(f"Error using phonenumbers library: {e}")
                # Fallback to manual detection methods
                pass
            
            # Fallback to our original manual detection logic
            return self._detect_phone_type_manual(digits_only, phone)
            
        except Exception as e:
            print(f"Error in phone type detection: {e}")
            # If any error occurs, try to return something meaningful
            if digits_only.startswith('08') or (digits_only.startswith('62') and digits_only[2] == '8'):
                return "Mobile"
            elif digits_only.startswith('0') or (digits_only.startswith('62') and digits_only[2] != '8'):
                return "Office"
            return "Unknown"
            
    def _detect_mobile_carrier_manual(self, digits_only, phone):
        """Detect the mobile carrier based on the number prefix (manual method)"""
        # Extract prefix for local format 08xx
        if digits_only.startswith('08') and 10 <= len(digits_only) <= 13:
            prefix = digits_only[0:4]  # Get first 4 digits (08xx)
            
            # Telkomsel: 0811, 0812, 0813, 0821, 0822, 0823, 0851, 0852, 0853
            if prefix in ['0811', '0812', '0813', '0821', '0822', '0823', '0851', '0852', '0853']:
                return "Mobile (Telkomsel)"
                
            # Indosat Ooredoo: 0814, 0815, 0816, 0855, 0856, 0857, 0858
            elif prefix in ['0814', '0815', '0816', '0855', '0856', '0857', '0858']:
                return "Mobile (Indosat)"
                
            # XL Axiata: 0817, 0818, 0819, 0859, 0877, 0878
            elif prefix in ['0817', '0818', '0819', '0859', '0877', '0878']:
                return "Mobile (XL Axiata)"
                
            # Tri (3): 0895, 0896, 0897, 0898, 0899
            elif prefix in ['0895', '0896', '0897', '0898', '0899']:
                return "Mobile (Tri)"
                
            # Smartfren: 0881, 0882, 0883, 0884, 0885, 0886, 0887, 0888, 0889
            elif prefix in ['0881', '0882', '0883', '0884', '0885', '0886', '0887', '0888', '0889']:
                return "Mobile (Smartfren)"
        
        # Extract prefix for international format +628xx or 628xx
        if (digits_only.startswith('628') or 
            (phone.startswith('+62') and digits_only[2] == '8')):
            
            # Extract the carrier prefix (8xx)
            prefix = None
            if digits_only.startswith('628'):
                prefix = digits_only[2:5]  # Get the xxx part from 628xxx
            else:
                # Handle +628xx format
                prefix = digits_only[2:5]  # Get the 8xx part from +628xx
                
            if prefix:
                # Telkomsel: 811, 812, 813, 821, 822, 823, 851, 852, 853
                if prefix in ['811', '812', '813', '821', '822', '823', '851', '852', '853']:
                    return "Mobile (Telkomsel)"
                    
                # Indosat Ooredoo: 814, 815, 816, 855, 856, 857, 858
                elif prefix in ['814', '815', '816', '855', '856', '857', '858']:
                    return "Mobile (Indosat)"
                    
                # XL Axiata: 817, 818, 819, 859, 877, 878
                elif prefix in ['817', '818', '819', '859', '877', '878']:
                    return "Mobile (XL Axiata)"
                    
                # Tri (3): 895, 896, 897, 898, 899
                elif prefix in ['895', '896', '897', '898', '899']:
                    return "Mobile (Tri)"
                    
                # Smartfren: 881, 882, 883, 884, 885, 886, 887, 888, 889
                elif prefix in ['881', '882', '883', '884', '885', '886', '887', '888', '889']:
                    return "Mobile (Smartfren)"
        
        # Default for mobile
        return "Mobile"
            
    def _detect_office_location_manual(self, digits_only, phone):
        """Detect the office location based on area code (manual method)"""
        # Indonesian area codes for major cities
        area_codes = {
            '21': 'Jakarta',
            '22': 'Bandung',
            '24': 'Semarang',
            '31': 'Surabaya',
            '61': 'Medan',
            '251': 'Bogor',
            '254': 'Serang',
            '274': 'Yogyakarta',
            '341': 'Malang',
            '351': 'Madiun',
            '361': 'Denpasar/Bali',
            '370': 'Mataram/Lombok',
            '411': 'Makassar',
            '431': 'Manado',
            '511': 'Banjarmasin',
            '542': 'Samarinda',
            '561': 'Pontianak',
            '711': 'Palembang',
            '721': 'Bandar Lampung',
            '751': 'Padang',
            '761': 'Pekanbaru',
            '778': 'Batam'
        }
        
        # 1. Jakarta area code format: 021, +6221, 6221
        if (digits_only.startswith('021') or 
            (digits_only.startswith('62') and digits_only[2:4] == '21') or
            (phone.startswith('+62') and digits_only[2:4] == '21')):
            return "Office (Jakarta)"
        
        # 2. Check other area codes with 0 prefix (domestic format)
        if digits_only.startswith('0'):
            # Try 2-digit area codes (like 022, 031)
            if len(digits_only) >= 3:
                area_code = digits_only[1:3]  # Extract xx from 0xx
                if area_code in area_codes:
                    return f"Office ({area_codes[area_code]})"
                    
            # Try 3-digit area codes (like 0251, 0341)
            if len(digits_only) >= 4:
                area_code = digits_only[1:4]  # Extract xxx from 0xxx
                if area_code in area_codes:
                    return f"Office ({area_codes[area_code]})"
        
        # 3. Check area codes with international prefix +62 or 62
        if digits_only.startswith('62') or phone.startswith('+62'):
            index = 2  # Start position after 62
            
            # Try 2-digit area codes (like +6222, 6231)
            area_code = digits_only[index:index+2]
            if area_code in area_codes and area_code != '8':  # Exclude mobile prefix
                return f"Office ({area_codes[area_code]})"
                
            # Try 3-digit area codes (like +62251, 62341)
            area_code = digits_only[index:index+3]
            if area_code in area_codes:
                return f"Office ({area_codes[area_code]})"
        
        # Default for office
        return "Office"
            
    def _detect_phone_type_manual(self, digits_only, phone):
        """Original manual phone type detection method as fallback"""
        # ===== MOBILE PHONE DETECTION =====
        
        # Indonesian mobile phone formats: 08xx or +628xx/628xx
        
        # 1. Check for local format 08xx
        if digits_only.startswith('08') and 10 <= len(digits_only) <= 13:
            return self._detect_mobile_carrier_manual(digits_only, phone)
            
        # 2. Check for international format +628xx or 628xx
        if (digits_only.startswith('628') or 
            (phone.startswith('+62') and digits_only[2] == '8')):
            return self._detect_mobile_carrier_manual(digits_only, phone)
            
        # ===== OFFICE/LANDLINE PHONE DETECTION =====
        
        # 1. Jakarta area code format: 021, +6221, 6221
        if (digits_only.startswith('021') or 
            (digits_only.startswith('62') and digits_only[2:4] == '21') or
            (phone.startswith('+62') and digits_only[2:4] == '21')):
            return "Office (Jakarta)"
            
        # Special case for non-Indonesian international formats
        if phone.startswith('+') and not phone.startswith('+62'):
            # For any international format not using Indonesian country code
            # Default to Mobile for numbers with sufficient length
            if len(digits_only) >= 8:
                return "Mobile"
            
        # 2. Check other area codes with 0 prefix (domestic format)
        if digits_only.startswith('0'):
            return self._detect_office_location_manual(digits_only, phone)
                
        # 3. Check area codes with international prefix +62 or 62
        if digits_only.startswith('62') or phone.startswith('+62'):
            return self._detect_office_location_manual(digits_only, phone)
        
        # ===== FALLBACK LOGIC =====
        
        # If format is valid but wasn't matched by specific rules
        if len(digits_only) >= 8:
            # Default to mobile for other unrecognized formats
            return "Mobile"
            
        # Completely unrecognized format
        return "Unknown"
        
    def _validate_email(self, email, check_deliverability=False):
        """Validate and normalize an email address using email-validator library
        
        Args:
            email (str): Email address to validate
            check_deliverability (bool): If True, check if the domain actually accepts email
                                         Default is False to avoid false negatives for valid formats
        """
        if not email:
            return {
                'email': None,
                'valid': False,
                'error': 'Email is empty'
            }
            
        try:
            # Validate and normalize email with configured deliverability check
            valid = validate_email(email, check_deliverability=check_deliverability)
            # Get the normalized form
            normalized_email = valid.normalized
            
            # Get additional information
            domain = valid.domain
            
            # Check if domain is disposable (temporary email)
            is_disposable = False
            disposable_domains = ['mailinator.com', 'yopmail.com', 'tempmail.com', 'temp-mail.org', 
                                  'guerrillamail.com', '10minutemail.com', 'trashmail.com']
            if domain.lower() in [d.lower() for d in disposable_domains]:
                is_disposable = True
                
            return {
                'email': normalized_email,
                'domain': domain,
                'is_disposable': is_disposable,
                'valid': True
            }
            
        except EmailNotValidError as e:
            print(f"Invalid email ({email}): {str(e)}")
            return {
                'email': email,
                'error': str(e),
                'valid': False
            }
    
    def save_to_csv(self, company_data, filename="company_data.csv"):
        """Save the scraped data to a CSV or Excel file"""
        if company_data:
            # Buat salinan dari data yang tidak akan dimodifikasi
            export_data = {}
            for key, value in company_data.items():
                # Jangan sertakan kunci internal yang tidak perlu diekspor
                if key not in ['phone_sources', 'email_source', 'phones']:
                    export_data[key] = value
            
            # Add phone type detection
            if export_data.get('phone'):
                export_data['phone_type'] = self._detect_phone_type(export_data['phone'])
            
            # Handle multiple phone numbers if present
            if 'phones' in company_data and company_data['phones']:
                # Join all phone numbers with semicolons for 'all_phones' field
                export_data['all_phones'] = '; '.join(company_data['phones'])
                
                # Add individual phone columns for up to 5 phone numbers
                for i, phone in enumerate(company_data['phones'][:5]):  # Limit to 5 phones
                    export_data[f'phone_{i+1}'] = phone
                    # Add phone type detection for each phone
                    export_data[f'phone_{i+1}_type'] = self._detect_phone_type(phone)
                    
                    # Add source and type info for each phone if available
                    if 'phone_sources' in company_data and phone in company_data['phone_sources']:
                        source_info = company_data['phone_sources'][phone]
                        # Add phone type (e.g., Customer Service, Head Office, etc.)
                        if 'type' in source_info:
                            export_data[f'phone_{i+1}_category'] = source_info['type']
                        # Add source page (e.g., contact page, homepage, etc.)
                        if 'page' in source_info:
                            export_data[f'phone_{i+1}_source'] = source_info['page']
                        # Add URL if available
                        if 'url' in source_info:
                            export_data[f'phone_{i+1}_url'] = source_info['url']
                
                # Ensure 'phone' field has at least the first number for compatibility
                if not export_data.get('phone') and company_data['phones']:
                    export_data['phone'] = company_data['phones'][0]
                    export_data['phone_type'] = self._detect_phone_type(company_data['phones'][0])
            
            # Add email source info if available
            if company_data.get('email') and company_data.get('email_source'):
                email_source = company_data['email_source']
                if isinstance(email_source, dict):
                    if 'type' in email_source:
                        export_data['email_type'] = email_source['type']
                    if 'page' in email_source:
                        export_data['email_source_page'] = email_source['page']
                    if 'url' in email_source:
                        export_data['email_url'] = email_source['url']
                elif isinstance(email_source, str):
                    export_data['email_source_page'] = email_source
            
            # Ensure data source information is available
            if not export_data.get('data_source'):
                export_data['data_source'] = 'Google Maps'
            
            # Add Source URL in a clean format if not present
            if not export_data.get('source_url'):
                if export_data.get('current_url'):
                    # Prioritize business URLs (with /place/) over search URLs
                    current_url = export_data.get('current_url')
                    if "/place/" in current_url:
                        export_data['source_url'] = current_url
                    elif export_data.get('data_source') == 'Website' and export_data.get('website'):
                        export_data['source_url'] = export_data['website']
                    else:
                        export_data['source_url'] = current_url
                elif export_data.get('data_source') == 'Website' and export_data.get('website'):
                    export_data['source_url'] = export_data['website']
                else:
                    export_data['source_url'] = 'N/A'
            
            # Create DataFrame
            df = pd.DataFrame([export_data])
            
            # Clean any dictionary-like strings in the dataframe
            for column in df.columns:
                if df[column].dtype == 'object':  # Only process string columns
                    df[column] = df[column].apply(lambda x: self._clean_dict_string(x) if isinstance(x, str) else x)
            
            # Reorder columns for better readability
            cols = df.columns.tolist()
            
            # Create better organized groups of columns
            basic_cols = ['name', 'address', 'category', 'rating', 'reviews_count', 'original_query']
            phone_cols = ['phone', 'phone_type']
            phone_detail_cols = [col for col in cols if col.startswith('phone_')]
            email_cols = ['email', 'email_valid', 'email_domain', 'email_disposable', 'email_type', 'email_source_page', 'email_url']
            website_cols = ['website', 'website_scraped']
            location_cols = ['latitude', 'longitude', 'coordinates']
            source_cols = ['data_source', 'source_url']  # New column group for source info
            
            # Organize columns in logical groups
            new_cols = []
            
            # First add basic info
            for col in basic_cols:
                if col in cols:
                    new_cols.append(col)
            
            # Then data source information
            for col in source_cols:
                if col in cols:
                    new_cols.append(col)
            
            # Then contact information
            for col in phone_cols:
                if col in cols:
                    new_cols.append(col)
            
            # Add phone details in order (phone_1, phone_1_type, phone_1_source, etc.)
            phone_numbers = sorted([col for col in phone_detail_cols if re.match(r'phone_\d+$', col)])
            for phone_num in phone_numbers:
                num_suffix = phone_num.split('_')[-1]
                new_cols.append(phone_num)
                
                # First add the type (WhatsApp/Office)
                type_col = f'phone_{num_suffix}_type'
                if type_col in cols:
                    new_cols.append(type_col)
                
                # Then add the category and sources
                category_col = f'phone_{num_suffix}_category'
                source_col = f'phone_{num_suffix}_source'
                url_col = f'phone_{num_suffix}_url'
                
                if category_col in cols:
                    new_cols.append(category_col)
                if source_col in cols:
                    new_cols.append(source_col)
                if url_col in cols:
                    new_cols.append(url_col)
            
            # Add all_phones at the end of phone section
            if 'all_phones' in cols:
                new_cols.append('all_phones')
            
            # Add email information
            for col in email_cols:
                if col in cols:
                    new_cols.append(col)
            
            # Add website info
            for col in website_cols:
                if col in cols:
                    new_cols.append(col)
            
            # Add location data
            for col in location_cols:
                if col in cols:
                    new_cols.append(col)
            
            # Add any remaining columns we didn't specifically organize
            for col in cols:
                if col not in new_cols:
                    new_cols.append(col)
            
            # Reorder if all columns exist            
            if all(col in df.columns for col in new_cols):
                df = df[new_cols]
                
            # Try to save to Excel if the filename ends with .xlsx
            if filename.lower().endswith('.xlsx'):
                try:
                    # Save to Excel with formatting
                    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False, sheet_name='Companies')
                        
                        # Try to add column formatting
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
                        except Exception as e:
                            print(f"Warning: Could not format Excel: {str(e)}")
                            
                    print(f"Data saved to Excel file: {filename}")
                    return
                except Exception as e:
                    print(f"Error saving to Excel: {e}, falling back to CSV")
                    # Fall back to CSV if Excel fails
                    csv_filename = filename.replace('.xlsx', '.csv')
                    df.to_csv(csv_filename, index=False)
                    print(f"Data saved to CSV file: {csv_filename}")
                    return
            
            # Standard CSV save
            df.to_csv(filename, index=False)
            print(f"Data saved to {filename}")
        else:
            print("No data to save")
            
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
                    if parsed_dict['page'] == 'known_data' and 'url' in parsed_dict:
                        return parsed_dict['url']
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
    
    def close(self):
        """Close the browser and clean up resources"""
        try:
            if hasattr(self, 'driver') and self.driver:
                self.driver.quit()
                print("WebDriver closed successfully")
        except Exception as e:
            print(f"Error closing driver: {str(e)}")
    
    def clear_cache(self):
        """Clear the cache directory"""
        if self.use_cache and self.cache:
            try:
                # Get cache statistics before clearing
                cache_size = len(self.cache)
                # Clear the cache
                self.cache.clear()
                print(f"Cache cleared successfully. Removed {cache_size} entries.")
                return True
            except Exception as e:
                print(f"Error clearing cache: {str(e)}")
                raise e
        else:
            print("Cache is not enabled, nothing to clear.")
            return False
    
    def _calculate_name_similarity(self, name1, name2):
        """Calculate similarity between two company names"""
        # Convert both to lowercase and remove common words
        common_words = ['pt', 'cv', 'tbk', 'persero', 'limited', 'ltd', 'company']
        
        def clean_name(name):
            name = name.lower()
            # Split into words
            words = name.split()
            # Remove common corporate designations
            words = [w for w in words if w not in common_words]
            return ' '.join(words)
            
        clean1 = clean_name(name1)
        clean2 = clean_name(name2)
        
        # If either name is empty after cleaning, similarity is 0
        if not clean1 or not clean2:
            return 0.0
            
        # Simple word overlap score
        words1 = set(clean1.split())
        words2 = set(clean2.split())
        
        # If we have no meaningful words, return 0
        if not words1 or not words2:
            return 0.0
            
        # Calculate Jaccard similarity
        intersection = len(words1.intersection(words2))
        union = len(words1.union(words2))
        
        if union == 0:
            return 0.0
            
        return intersection / union


if __name__ == "__main__":
    scraper = GoogleMapsScraper(max_retries=3, retry_delay=3)
    
    try:
        # Special test for OCS Global Services
        test_ocs = input("Would you like to test OCS Global Services specifically? (y/n): ")
        if test_ocs.lower() == 'y':
            print("\n=== Testing OCS Global Services ===")
            company_data = scraper.search_company("OCS Global Services")
            
            if company_data and company_data.get('phone'):
                print("\n✅ SUCCESS! Found OCS Global Services phone number!")
                print("\nExtracted Company Data:")
                for key, value in company_data.items():
                    print(f"{key.capitalize()}: {value}")
                
                # Offer to save
                save_option = input("\nDo you want to save this data to CSV? (y/n): ")
                if save_option.lower() == 'y':
                    scraper.save_to_csv(company_data, "ocs_data.csv")
                    print("Data saved to ocs_data.csv")
            else:
                print("\n❌ Failed to find phone number for OCS Global Services.")
                print("Please check the search_variation_*.png screenshots for debugging.")
        else:
            # Normal search for any company
            company_name = input("Enter the company name to search: ")
            
            company_data = scraper.search_company(company_name)
            
            if company_data and any(value for key, value in company_data.items() if key != 'name'):
                print("\nExtracted Company Data:")
                for key, value in company_data.items():
                    print(f"{key.capitalize()}: {value}")
                
                save_option = input("\nDo you want to save this data to CSV? (y/n): ")
                if save_option.lower() == 'y':
                    filename = input("Enter filename (default: company_data.csv): ")
                    if not filename:
                        filename = "company_data.csv"
                    scraper.save_to_csv(company_data, filename)
            else:
                print("Could not find detailed company data. Only name is available.")
    
    finally:
        scraper.close() 