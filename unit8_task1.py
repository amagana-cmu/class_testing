import requests
from requests.exceptions import RequestException
from bs4 import BeautifulSoup
from bs4.element import NavigableString
import json
import csv
import xml.etree.ElementTree as ET
from typing import List, Dict, Any, Optional

def get_names_page(num: int) -> Optional[str]:
    """
    Fetches the HTML content for a given page number from behindthename.com.
    
    Args:
        num: The page number to fetch.
        
    Returns:
        The HTML content as a string, or None if an error occurs.
    """
    # Construct the URL for the specific page number
    url = f"https://www.behindthename.com/names/{num}"
    
    # Set a User-Agent to be a good web citizen
    headers = {
        'User-Agent': 'YourAppName/1.0 (YourSchoolOrPersonalProject; your-email@example.com)'
    }
    
    try:
        response = requests.get(url, headers=headers)
        # Raise an exception for bad status codes (4xx or 5xx)
        response.raise_for_status()
        return response.text
    except RequestException as e:
        print(f"Error fetching page {num}: {e}")
        return None

def extract_page_count(page_text: str) -> int:
    soup = BeautifulSoup(page_text, 'html.parser')
    
    max_page = 0
    pagination = soup.find('nav', class_='pagination')
    
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
    Parses the HTML text of a page and extracts name data based on parameters.
    
    Args:
        page_text: The HTML content of a page.
        gender: Include gender information.
        usage: Include usage (origin/language) information.
        desc: Include description information.
        
    Returns:
        A dictionary where keys are names and values are dicts of their data.
    """
    soup = BeautifulSoup(page_text, 'html.parser')
    results: Dict[str, Dict[str, Any]] = {}
    
    # Find all name entries, which are marked by a span with class "listname"
    name_spans = soup.find_all('span', class_='listname')
    
    for name_span in name_spans:
        # Get the name (text inside the <span>)
        name = name_span.text.strip()
        
        # The parent <div> contains the siblings
        parent_div = name_span.parent
        
        # The description is in the next sibling <div> with class "listdesc"
        desc_div = parent_div.find_next_sibling('div', class_='listdesc')
        
        name_data: Dict[str, Any] = {}
        
        # Process Gender
        if gender:
            gender_span = name_span.find_next_sibling('span', class_='listgender')
            if gender_span:
                gender_text = gender_span.text.strip().lower()
                # Apply the specific logic from your prompt
                if 'masculine' in gender_text and 'feminine' in gender_text:
                    name_data['gender'] = 'm & f'
                elif 'masculine' in gender_text:
                    name_data['gender'] = 'm'
                elif 'feminine' in gender_text:
                    name_data['gender'] = 'f'
                else:
                    name_data['gender'] = gender_text # Fallback
        
        # Process Usage
        if usage:
            usage_span = name_span.find_next_sibling('span', class_='listusage')
            if usage_span:
                # Find all <a> tags within the usage span to get each usage part
                usage_links = usage_span.find_all('a')
                name_data['usage'] = [u.text.strip() for u in usage_links]
        
        # Process Description
        if desc:
            br_tag = parent_div.find('br')
            if br_tag:
                desc_parts = []
                current_node = br_tag.next_sibling
                while current_node:
                    if getattr(current_node, 'name', None) == 'span' and 'namedesc' in getattr(current_node, 'attrs', {}).get('id', ''):
                        break

                    if isinstance(current_node, NavigableString):
                        desc_parts.append(str(current_node))
                    elif hasattr(current_node, 'get_text'):
                        desc_parts.append(current_node.get_text())
                    
                    current_node = current_node.next_sibling
                full_desc = ''.join(desc_parts).strip()
                if full_desc:
                    name_data['desc'] = full_desc
                
            # if desc_div:
            #     name_data['desc'] = desc_div.text.strip()
        
        # Add the collected data to the main dictionary
        # This will overwrite duplicates on the same page, as per the dict-key requirement
        results[name] = name_data
            
    return results


def scrape_names(pages: List[int], output_file_path: str, gender: bool = True, usage: bool = False, desc: bool = False, output_format: str = "csv"):
    """
    Scrapes name data from a list of pages and saves to a file in the specified format.
    
    Args:
        pages: A list of page numbers to scrape.
        output_file_path: The path to the output file (e.g., "names.csv").
        gender: Include gender information.
        usage: Include usage information.
        desc: Include description information.
        output_format: The format of the output file ("csv", "json", or "xml").
    """
    
    # This master dictionary will hold all data from all pages
    all_names_data: Dict[str, Dict[str, Any]] = {}

    print(f"Starting scrape for {len(pages)} pages...")
    
    for page_num in pages:
        print(f"Scraping page {page_num}...")
        html = get_names_page(page_num)
        if html:
            page_data = extract_names_from_page(html, gender, usage, desc)
            # Update the master dictionary
            # Note: names on later pages will overwrite names from earlier pages
            all_names_data.update(page_data)
        else:
            print(f"Skipping page {page_num} (failed to fetch).")
    
    print(f"Scrape complete. Total unique names found: {len(all_names_data)}")
    print(f"Writing data to {output_file_path} as {output_format}...")

    # Write the aggregated data to the specified file format
    
    if output_format == "csv":
        # Define header based on boolean flags
        fieldnames = ['name']
        if gender: fieldnames.append('gender')
        if usage: fieldnames.append('usage')
        if desc: fieldnames.append('desc')
        
        with open(output_file_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            for name, data in all_names_data.items():
                row = {'name': name}
                if gender:
                    row['gender'] = data.get('gender')
                if usage:
                    # Join the list of usages into a single string for CSV
                    row['usage'] = ', '.join(data.get('usage', []))
                if desc:
                    row['desc'] = data.get('desc')
                writer.writerow(row)
                
    elif output_format == "json":
        with open(output_file_path, 'w', encoding='utf-8') as f:
            json.dump(all_names_data, f, indent=4, ensure_ascii=False)
            
    elif output_format == "xml":
        root = ET.Element('names')
        for name, data in all_names_data.items():
            name_elem = ET.SubElement(root, 'name', value=name)
            
            if gender and data.get('gender'):
                ET.SubElement(name_elem, 'gender').text = data.get('gender')
                
            if usage and data.get('usage'):
                # Create a separate <usage> tag for each item in the list
                for u in data.get('usage', []):
                    ET.SubElement(name_elem, 'usage').text = u
            
            if desc and data.get('desc'):
                ET.SubElement(name_elem, 'desc').text = data.get('desc')
                
        tree = ET.ElementTree(root)
        # Pretty-print the XML
        ET.indent(tree, space="  ")
        tree.write(output_file_path, encoding='utf-8', xml_declaration=True)
        
    else:
        print(f"Error: Unsupported output format '{output_format}'. Please use 'csv', 'json', or 'xml'.")

    print("File writing complete.")


if __name__ == "__main__":


    # --- Test scrape_names ( formats) ---
    # test
    test_pages = [1, 2]
    
    print(f"\n--- Testing scrape_names for pages {test_pages} ---")
    
    # Test CSV
    print("\nTesting CSV output...")
    scrape_names(test_pages, "names_output.csv", gender=True, usage=True, desc=True, output_format="csv")
    
    # Test JSON
    print("\nTesting JSON output...")
    scrape_names(test_pages, "names_output.json", gender=True, usage=True, desc=True, output_format="json")

    # Test XML
    print("\nTesting XML output...")
    scrape_names(test_pages, "names_output.xml", gender=True, usage=True, desc=True, output_format="xml")
    
    print("\n--- All tests complete. Check output files. ---")
