import requests
from bs4 import BeautifulSoup
from bs4.element import NavigableString
import json
import csv
import xml.etree.ElementTree as ET
from typing import List, Dict, Any, Optional

def get_names_page(num: int) -> Optional[str]:
    # Construct the URL for the specific page number
    url = f"https://www.behindthename.com/names/{num}"

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36'
    }

    try:
        # Make the HTTP GET request
        response = requests.get(url, headers=headers)
        # Raise an exception for bad responses (like 404s or 500s)
        response.raise_for_status() 
        # Return the HTML content as a string
        return response.text
    except requests.RequestException as e:
        print(f"Error fetching page {num}: {e}")
        return None
    
def extract_page_count(page_text: str) -> int:
    soup = BeautifulSoup(page_text, 'html.parser')
    
    max_page = 98
    pagination = soup.find('div', class_='pagination')
    
    if pagination:
        # Find all links within the pagination div
        links = pagination.find_all('a')
        
        for link in links:
            # Check if the link's text is a number
            if link.text.isdigit():
                page_num = int(link.text)
                if page_num > max_page:
                    max_page = page_num

    return max_page

def extract_names_from_page(page_text: str, gender: bool = True, usage: bool = False, desc: bool = False) -> Dict[str, Dict[str, Any]]:
    """
    Parses the HTML text of a page and extracts name data.
    """
    soup = BeautifulSoup(page_text, 'html.parser')
    
    # This dictionary will hold all the names from this page
    names_dict = {}
    
    # Find all the main "name blocks". 
    name_blocks = soup.find_all('div', class_='browsename')

    # Loop through every name block found on the page
    for block in name_blocks:
        
        # Find the name itself, which is inside an <a> (link) tag
        name_tag = block.find('a', class_='nll')
        
        # If there's no <a> tag, it's not a valid entry, so we skip it.
        if not name_tag:
            continue
            
        name = name_tag.text
        data_payload: Dict[str, Any] = {} # new dictionary

        #  Find the gender IF the user wants it (gender=True)
        if gender:
            gender_tag = block.find('span', class_='gengender')
            if gender_tag:
                # The gender is in the title attribute of the inner span (e.g., <span class="masc" title="masculine">)
                data_payload['gender'] = gender_tag.text.strip()
                # gender_inner_span = gender_tag.find('span')
                # if gender_inner_span and gender_inner_span.has_attr('title'):
                #     data_payload['gender'] = gender_inner_span['title']
            
        # Find the usage IF the user wants it (usage=True)
        if usage:
            usage_list = []
            # Find all usage links within the listusage span
            usage_links = block.select('span.listusage a.usg')
            for link in usage_links:
                usage_list.append(link.text.strip())
            if usage_list:
                data_payload['usage'] = usage_list
        
        # Find the description IF the user wants it (desc=True)
        if desc:
            desc_to_add = "" # A variable to hold our description
            
            # First, try to find the description in a 'div.namedesc' tag
            desc_tag = block.find('div', class_='namedesc')
            if desc_tag:
                desc_to_add = desc_tag.text.strip()
            else:
                # If that fails, try the <br> tag method as the comment suggests
                br_tag = block.find('br')
                if br_tag:
                    # Look at the *next* sibling that is a non-empty string
                    for sibling in br_tag.next_siblings:
                        if isinstance(sibling, NavigableString):
                            sibling_text = sibling.strip()
                            if sibling_text:
                                desc_to_add = sibling_text
                                break # Found the description, stop looping
                        else:
                            # Hit another tag, stop
                            break
            
            # If we found a description by *either* method, add it
            if desc_to_add:
                data_payload['desc'] = desc_to_add

        # data to our main dictionary
        if name:
            names_dict[name] = data_payload

    return names_dict
def scrape_names(pages: List[int], output_file_path: str, gender: bool = True, usage: bool = False, desc: bool = False, output_format: str = "csv"):
    
    # This master dictionary will hold all data from all pages
    all_names_data: Dict[str, Dict[str, Any]] = {}

    print(f"Starting to scrape {len(pages)} page(s)...")

    # ---  GATHER  DATA ---
    for page_num in pages:
        print(f"Fetching page {page_num}...")
        html_text = get_names_page(page_num)
        
        if html_text:
            # Step 2: Extract data from that page's HTML
            names_from_page = extract_names_from_page(
                html_text, 
                gender=gender, 
                usage=usage, 
                desc=desc
            )
            # Add new data to master dictionary
            all_names_data.update(names_from_page)
        else:
            print(f"Warning: Could not fetch page {page_num}. Skipping.")

    print(f"Scraping complete. Total names found: {len(all_names_data)}")
    print(f"Now writing to file: {output_file_path} as {output_format}")
    
    # --- WRITE DATA TO FILE ---

    if output_format == "csv":
        # Define the header based on the parameters
        #    important for matching the project spec
        header = ['name']
        if gender:
            header.append('gender')
        if usage:
            header.append('usage')
        if desc:
            header.append('desc')

        try:
            with open(output_file_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                
                # 2. Write the header row
                writer.writerow(header)
                
                # 3. Loop through your data and write each row
                for name, data in all_names_data.items():
                    row = [name]
                    
                    # .get() is safer than data['key']
                    # It returns None if the key doesn't exist
                    
                    if gender:
                        row.append(data.get('gender', ''))
                    if usage:
                        # Join the list into a string, e.g., "Finnish, Somali"
                        # [cite: 256]
                        usage_str = ", ".join(data.get('usage', []))
                        row.append(usage_str)
                    if desc:
                        row.append(data.get('desc', ''))
                    
                    writer.writerow(row)
        except IOError as e:
            print(f"Error writing CSV file: {e}")

    elif output_format == "json":
        try:
            with open(output_file_path, 'w', encoding='utf-8') as f:
                # indent=4 makes the file human-readable (pretty-printed)
                json.dump(all_names_data, f, indent=4, ensure_ascii=False)
        except IOError as e:
            print(f"Error writing JSON file: {e}")

    elif output_format == "xml":
        try:
            # 1. Create the root <names> element [cite: 293]
            root = ET.Element("names")
            
            # 2. Loop through each name in your data
            for name, data in all_names_data.items():
                # 3. Create the <name> element with its 'value' attribute [cite: 294]
                name_elem = ET.SubElement(root, "name", value=name)
                
                # 4. Add sub-elements *if* they exist in the data
                if gender and 'gender' in data:
                    gender_elem = ET.SubElement(name_elem, "gender")
                    gender_elem.text = data['gender']
                
                if usage and 'usage' in data:
                    # Create a new <usage> tag for EACH item in the list [cite: 306, 307]
                    for usage_item in data['usage']:
                        usage_elem = ET.SubElement(name_elem, "usage")
                        usage_elem.text = usage_item
                
                if desc and 'desc' in data:
                    desc_elem = ET.SubElement(name_elem, "desc")
                    desc_elem.text = data['desc']
            
            # 5. Create the full tree and write it to the file
            tree = ET.ElementTree(root)
            # This makes the XML pretty-printed
            ET.indent(tree, space="  ") 
            tree.write(output_file_path, encoding='utf-8', xml_declaration=True)
            
        except IOError as e:
            print(f"Error writing XML file: {e}")       

if __name__ == "__main__":
    
    print("--- Testing get_names_page and extract_page_count ---")
    # First, let's test the helper functions
    test_html = get_names_page(1)
    
    if test_html:
        # Test extract_page_count
        max_page = extract_page_count(test_html)
        print(f"Successfully found max page: {max_page}")
        
        # Test extract_names_from_page
        print("\n--- Testing extract_names_from_page (first 5 names) ---")
        names_data = extract_names_from_page(test_html, gender=True, usage=True, desc=True)
        
        count = 0
        for name, data in names_data.items():
            print(f"{name}: {data}")
            count += 1
            if count >= 5:
                break
    else:
        print("Failed to fetch test HTML. Halting tests.")

    # --- Testing scrape_names for all 3 formats ---
    
    # We'll scrape just two pages for this test
    pages_to_scrape = [1, 2]
    
    print(f"\n--- Testing scrape_names for {pages_to_scrape} ---")

    # Test 1: CSV
    print("Testing CSV output...")
    scrape_names(
        pages=pages_to_scrape,
        output_file_path="names.csv",
        gender=True,
        usage=True,
        desc=True,
        output_format="csv"
    )

    # Test 2: JSON
    print("Testing JSON output...")
    scrape_names(
        pages=pages_to_scrape,
        output_file_path="names.json",
        gender=True,
        usage=True,
        desc=False, # Test with different flags
        output_format="json"
    )

    # Test 3: XML
    print("Testing XML output...")
    scrape_names(
        pages=pages_to_scrape,
        output_file_path="names.xml",
        gender=True,
        usage=True,
        desc=True,
        output_format="xml"
    )
    
    print("\nAll tests complete. Check your project folder for names.csv, names.json, and names.xml.")
