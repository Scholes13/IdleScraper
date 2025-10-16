import time
import csv
import pandas as pd
import os

# Import the GoogleMapsScraper with proper path handling
try:
    from src.core.maps_scraper import GoogleMapsScraper
except ImportError:
    try:
        # When running from the same directory
        from maps_scraper import GoogleMapsScraper
    except ImportError:
        # Fallback for direct execution
        from maps_scraper import GoogleMapsScraper

def load_companies(file_path):
    """Load company names from a CSV or text file"""
    companies = []
    
    # Determine file type based on extension
    if file_path.endswith('.csv'):
        with open(file_path, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            # Skip header if exists
            header = next(reader, None)
            for row in reader:
                if row and len(row) > 0:
                    companies.append(row[0])
    else:
        # Assume text file with one company per line
        with open(file_path, 'r', encoding='utf-8') as f:
            companies = [line.strip() for line in f if line.strip()]
            
    return companies

def batch_scrape(input_file, output_file='companies_data.csv', enable_similar_search=False, 
                similarity_threshold=0.7, enable_website_scraping=False):
    """Scrape data for multiple companies from a file"""
    companies = load_companies(input_file)
    
    if not companies:
        print("No companies found in the input file.")
        return
    
    print(f"Loaded {len(companies)} companies from {input_file}")
    print(f"Smart search enabled: {enable_similar_search}, Similarity threshold: {similarity_threshold}")
    print(f"Website scraping enabled: {enable_website_scraping}")
    
    scraper = GoogleMapsScraper(
        max_retries=3, 
        retry_delay=2,
        enable_similar_search=enable_similar_search,
        similarity_threshold=similarity_threshold,
        enable_website_scraping=enable_website_scraping
    )
    
    results = []
    mapped_companies = []
    address_updates = []
    
    try:
        for i, company in enumerate(companies):
            print(f"\nProcessing {i+1}/{len(companies)}: {company}")
            
            data = scraper.search_company(company)
            
            if data:
                # Handle multiple phone numbers if present from website scraping
                if 'phones' in data and data['phones']:
                    # Join all phone numbers with semicolons for Excel display
                    data['all_phones'] = '; '.join(data['phones'])
                    # Ensure 'phone' field has at least the first number for compatibility
                    if not data.get('phone') and data['phones']:
                        data['phone'] = data['phones'][0]
                    print(f"Found {len(data['phones'])} phone numbers: {data['phones']}")
                
                # Track address updates
                if data.get('address'):
                    address_update = {
                        'original_name': company,
                        'updated_name': data.get('name', company),
                        'address': data.get('address'),
                        'coordinates': f"{data.get('latitude', '')},{data.get('longitude', '')}"
                    }
                    address_updates.append(address_update)
                    print(f"✓ Updated address found: {data.get('address')}")
                
                # Check if a name mapping occurred with sufficient similarity
                if (data.get('original_name') and data.get('mapped_name') and 
                    data.get('similarity_score', 0) >= similarity_threshold):
                    mapped_info = {
                        'original': company,
                        'mapped': data.get('mapped_name'),
                        'phone': data.get('phone'),
                        'all_phones': data.get('all_phones', data.get('phone', '')),
                        'variation': data.get('search_variation', 'Unknown'),
                        'address': data.get('address'),
                        'similarity_score': data.get('similarity_score', 0),
                    }
                    mapped_companies.append(mapped_info)
                    
                    print(f"✓ Name mapped: {company} → {data.get('mapped_name')} (Score: {data.get('similarity_score', 0):.2f})")
                
                results.append(data)
                print("✓ Data extracted successfully")
            else:
                print("✗ Failed to extract data")
            
            # Add delay between requests to avoid being blocked
            if i < len(companies) - 1:
                time.sleep(2)
    
    finally:
        scraper.close()
    
    # Save results to CSV
    if results:
        # Process for better excel output
        for data in results:
            # Handle multiple phone numbers
            if 'phones' in data and data['phones']:
                # Keep the all_phones field with all numbers joined
                data['all_phones'] = '; '.join(data['phones'])
                
                # Add individual phone columns for up to 5 phone numbers
                for i, phone in enumerate(data['phones'][:5]):  # Limit to 5 phones
                    data[f'phone_{i+1}'] = phone
                    
                    # Add phone type and source info if available
                    if 'phone_sources' in data and phone in data['phone_sources']:
                        source_info = data['phone_sources'][phone]
                        # Add phone type (e.g., Customer Service, Head Office, etc.)
                        if 'type' in source_info:
                            data[f'phone_{i+1}_type'] = source_info['type']
                        # Add source page (e.g., contact page, homepage, etc.)
                        if 'page' in source_info:
                            data[f'phone_{i+1}_source'] = source_info['page']
                
                # Ensure 'phone' field has at least the first number for compatibility
                if not data.get('phone') and data['phones']:
                    data['phone'] = data['phones'][0]
            
            # Add email source info if available
            if data.get('email') and data.get('email_source'):
                email_source = data['email_source']
                if 'type' in email_source:
                    data['email_type'] = email_source['type']
                if 'page' in email_source:
                    data['email_source_page'] = email_source['page']
        
        df = pd.DataFrame(results)
        
        # Reorder columns to put phones in a better order
        if not df.empty:
            cols = df.columns.tolist()
            
            # Create better organized groups of columns
            basic_cols = ['name', 'address', 'category', 'rating', 'reviews_count', 'original_query']
            phone_cols = ['phone']
            phone_detail_cols = [col for col in cols if col.startswith('phone_')]
            email_cols = ['email', 'email_type', 'email_source_page']
            website_cols = ['website', 'website_scraped']
            location_cols = ['latitude', 'longitude']
            
            # Organize columns in logical groups
            new_cols = []
            
            # First add basic info
            for col in basic_cols:
                if col in cols:
                    new_cols.append(col)
            
            # Then contact information
            for col in phone_cols:
                if col in cols:
                    new_cols.append(col)
            
            # Add phone details in order (phone_1, phone_1_type, phone_1_source, etc.)
            phone_numbers = [col for col in phone_detail_cols if col.endswith(tuple(['1','2','3','4','5']))]
            for phone_num in sorted(phone_numbers):
                num_suffix = phone_num.split('_')[-1]
                new_cols.append(phone_num)
                type_col = f'phone_{num_suffix}_type'
                source_col = f'phone_{num_suffix}_source'
                if type_col in cols:
                    new_cols.append(type_col)
                if source_col in cols:
                    new_cols.append(source_col)
            
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
                if col not in new_cols and col != 'phones' and col != 'phone_sources' and col != 'email_source':
                    new_cols.append(col)
            
            # Remove internal tracking columns to avoid confusion
            internal_cols = ['phones', 'phone_sources', 'email_source']
            new_cols = [col for col in new_cols if col not in internal_cols]
            
            # Reorder the dataframe            
            if all(col in df.columns for col in new_cols):
                df = df[new_cols]
        
        # Save to Excel instead of CSV for better formatting
        excel_file = output_file.replace('.csv', '.xlsx')
        if not excel_file.endswith('.xlsx'):
            excel_file += '.xlsx'
            
        try:
            # Apply some Excel styling for better readability
            with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Companies')
                
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
                except:
                    # Skip formatting if openpyxl advanced features aren't available
                    pass
                
            print(f"\nData for {len(results)} companies saved to {excel_file}")
        except Exception as e:
            print(f"Error saving to Excel: {e}")
            # Fallback to CSV
            df.to_csv(output_file, index=False)
            print(f"\nFell back to CSV: Data for {len(results)} companies saved to {output_file}")
        
        # Save address updates to a separate file
        if address_updates:
            address_file = os.path.splitext(output_file)[0] + "_address_updates.xlsx"
            address_df = pd.DataFrame(address_updates)
            try:
                address_df.to_excel(address_file, index=False)
                print(f"\n{len(address_updates)} address updates saved to {address_file}")
            except:
                # Fallback to CSV
                address_csv = os.path.splitext(output_file)[0] + "_address_updates.csv"
                address_df.to_csv(address_csv, index=False)
                print(f"\n{len(address_updates)} address updates saved to {address_csv}")
        
        # Display information about name mappings if any occurred
        if mapped_companies:
            print(f"\n{len(mapped_companies)} companies were mapped to similar names:")
            for item in mapped_companies:
                phones_display = item.get('all_phones', item.get('phone', 'None'))
                print(f"- {item['original']} → {item['mapped']} (Phones: {phones_display}, Score: {item['similarity_score']:.2f})")
            
            # Save mappings to a separate file
            mappings_file = os.path.splitext(output_file)[0] + "_mappings.xlsx"
            mappings_df = pd.DataFrame(mapped_companies)
            try:
                mappings_df.to_excel(mappings_file, index=False)
                print(f"\nName mappings saved to {mappings_file}")
            except:
                # Fallback to CSV
                mappings_csv = os.path.splitext(output_file)[0] + "_mappings.csv"
                mappings_df.to_csv(mappings_csv, index=False)
                print(f"\nName mappings saved to {mappings_csv}")
    else:
        print("\nNo data was extracted.")

if __name__ == "__main__":
    print("Google Maps Batch Scraper")
    print("-------------------------")
    
    input_file = input("Enter path to input file (CSV or TXT): ")
    output_file = input("Enter output CSV file path (default: companies_data.csv): ")
    
    if not output_file:
        output_file = "companies_data.csv"
    
    # Ask about using smart search
    use_smart_search = input("Enable smart search for similar companies? (y/n, default: n): ").lower() == 'y'
    similarity_threshold = 0.7
    
    if use_smart_search:
        threshold_input = input(f"Enter similarity threshold (0.0-1.0, default: {similarity_threshold}): ")
        if threshold_input and threshold_input.replace('.', '', 1).isdigit():
            similarity_threshold = float(threshold_input)
            similarity_threshold = max(0.0, min(1.0, similarity_threshold))
        
        print(f"Smart search enabled with similarity threshold: {similarity_threshold}")
    else:
        print("Smart search disabled - will only use exact matches")
    
    # Ask about using website scraping
    use_website_scraping = input("Enable website scraping for contact info? (y/n, default: n): ").lower() == 'y'
    if use_website_scraping:
        print("Website scraping enabled - will extract contact information from company websites")
    else:
        print("Website scraping disabled - will only use Google Maps data")
    
    batch_scrape(input_file, output_file, 
                 enable_similar_search=use_smart_search, 
                 similarity_threshold=similarity_threshold,
                 enable_website_scraping=use_website_scraping) 