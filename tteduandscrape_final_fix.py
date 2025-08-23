#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Final comprehensive script to parse DOCX timetable and extract all course-faculty assignments
including lab courses from paragraphs
"""

import sys
import re
import csv
import requests
from bs4 import BeautifulSoup
from docx import Document
import difflib

# Base URL template for faculty pages
BASE_URL = (
    "https://nitc.ac.in/department/computer-science-amp-engineering/"
    "faculty-and-staff/faculty/page/{page}"
)

def fetch_faculty_usernames(max_page=6):
    """Scrape faculty pages to build a mapping: faculty_name -> username"""
    name_to_username = {}
    for page in range(1, max_page + 1):
        url = BASE_URL.format(page=page)
        try:
            resp = requests.get(url)
            resp.raise_for_status()
            soup = BeautifulSoup(resp.text, 'html.parser')
            for box in soup.select('div.xc-faculty-box-content'):
                name_tag = box.select_one('h6.faculty-name')
                if not name_tag:
                    continue
                name = name_tag.get_text(strip=True)
                email_link = box.select_one("a[href^='mailto:']")
                if not email_link:
                    continue
                email = email_link['href'].replace('mailto:', '').strip()
                username = email.split('@')[0]
                name_to_username[name] = username
        except requests.RequestException as e:
            print(f"Warning: Failed to fetch page {page}: {e}", file=sys.stderr)
    return name_to_username

def parse_docx_courses(input_path):
    """Parse the .docx file to extract course-faculty assignments"""
    # Load the document
    doc = Document(input_path)

    # Build faculty code -> name mapping from the last table
    code_to_name = {}
    last_table = doc.tables[-1]  # Last table contains faculty mapping
    
    print("Building faculty code mapping...")
    # Process each row in the faculty mapping table
    for row in last_table.rows:
        cells = [c.text.strip() for c in row.cells]
        
        # Skip header rows or rows without proper data
        if len(cells) < 3:
            continue
            
        # Process left side of the table (columns 1 and 2)
        if len(cells) >= 3 and cells[1] and cells[2] and re.match(r'^[A-Za-z0-9]', cells[1]):
            faculty_code = cells[1].strip()
            faculty_name = cells[2].strip()
            code_to_name[faculty_code] = faculty_name
        
        # Process right side of the table (columns 5 and 6)
        if len(cells) >= 7 and cells[5] and cells[6] and re.match(r'^[A-Za-z0-9]', cells[5]):
            faculty_code = cells[5].strip()
            faculty_name = cells[6].strip()
            code_to_name[faculty_code] = faculty_name

    print(f"Found {len(code_to_name)} faculty code mappings")

    # List to store all course-faculty pairs
    pairs = []

    def extract_faculty_codes(faculty_str):
        """Extract individual faculty codes from a string like 'AB, AJ' or 'SK, NKB'"""
        # Remove any brackets that might be nested
        faculty_str = re.sub(r'\[[^\]]*\]', '', faculty_str).strip()
        
        # Split by comma and clean each code
        codes = []
        for code in faculty_str.split(','):
            code = code.strip()
            
            # Remove asterisks (e.g., JJ*)
            code = re.sub(r'\*', '', code).strip()
            
            # Remove any remaining special characters but preserve spaces for names
            code = re.sub(r'[^\w\s]', '', code).strip()
            
            if code and len(code) > 0:
                codes.append(code)
        
        return codes

    def process_course_text(text, source_info=""):
        """Process a text block that may contain course information"""
        # Remove [OE] tags completely
        text = re.sub(r'\[OE\]\s*', '', text)
        
        # Multiple regex patterns to handle different course formats
        patterns = [
            # Pattern 1: CS1001E: Computer Programming [TAS]
            re.compile(r'(?P<code>[A-Z]{2}\d{4}[A-Z]?)\s*:\s*(?P<name>[^[]+?)\s*\[(?P<faculty>[^\]]+)\]'),
            
            # Pattern 2: CS2002E Computer Organization [TMS] (no colon)
            re.compile(r'(?P<code>[A-Z]{2}\d{4}[A-Z]?)\s+(?P<name>[^[]+?)\s*\[(?P<faculty>[^\]]+)\]'),
            
            # Pattern 3: More flexible pattern for complex cases
            re.compile(r'(?P<code>[A-Z]{2}\d{4}[A-Z]?)[:\s]+(?P<name>.*?)\s*\[(?P<faculty>[^\]]+)\](?:\s*\[[^\]]*\])?')
        ]
        
        # Split by OR to handle multiple courses
        course_blocks = re.split(r'\s+OR\s+(?=[A-Z]{2}\d{4})', text)
        
        for block in course_blocks:
            block = block.strip()
            if not block:
                continue
                
            # Try each pattern
            for pattern in patterns:
                matches = pattern.finditer(block)
                
                for match in matches:
                    course_code = match.group('code').strip()
                    course_name = match.group('name').strip()
                    faculty_str = match.group('faculty').strip()
                    
                    # Clean up course name
                    course_name = ' '.join(course_name.split())
                    course_name = course_name.replace('\n', ' ').strip()
                    
                    # Extract faculty codes
                    faculty_codes = extract_faculty_codes(faculty_str)
                    
                    # Process each faculty code
                    for faculty_code in faculty_codes:
                        # Try to find faculty name from code mapping
                        faculty_name = code_to_name.get(faculty_code)
                        
                        if faculty_name:
                            pairs.append((faculty_name, course_code, course_name))
                        else:
                            # Check if this might be a full name instead of a code
                            if (len(faculty_code) > 3 and 
                                ' ' in faculty_code and
                                not re.match(r'^[A-Z]{2,4}\d*$', faculty_code)):
                                # Treat as full name
                                pairs.append((faculty_code, course_code, course_name))
                            else:
                                print(f"Warning: Faculty code '{faculty_code}' not found in mapping ({source_info})", file=sys.stderr)

    def process_lab_paragraph(text, source_info=""):
        """
        Process paragraph text containing lab course information in special format:
        (S1 BTech) CS1091E: Programming Lab : JJ*, AMP, KMJ, TAS, VP, JP, JCR, TV
        CS7198E/CS7298E/CS7398E: Project Phase-3 : GG, PARK, JPB
        """
        print(f"Processing lab paragraph ({source_info}): {text[:100]}...")
        
        # Pattern to match lab course entries
        # Matches: (optional semester) COURSE_CODE(s): Course Name : Faculty codes
        lab_pattern = re.compile(r'(?:\([^)]*\)\s+)?([A-Z]{2}\d{4}[A-Z]?(?:/[A-Z]{2}\d{4}[A-Z]?)*)\s*:\s*([^:]+?)\s*:\s*([A-Z][A-Z0-9*,\s]+?)(?=\s*\([^)]*\)|$)')
        
        matches = list(lab_pattern.finditer(text))
        print(f"  Found {len(matches)} lab course matches")
        
        for match in matches:
            course_codes_str = match.group(1).strip()
            course_name = match.group(2).strip()
            faculty_codes_str = match.group(3).strip()
            
            print(f"    Course codes: '{course_codes_str}'")
            print(f"    Course name: '{course_name}'")
            print(f"    Faculty codes: '{faculty_codes_str}'")
            
            # Split multiple course codes (e.g., CS7198E/CS7298E/CS7398E)
            course_codes = [code.strip() for code in course_codes_str.split('/')]
            
            # Extract faculty codes and remove asterisks
            faculty_codes = extract_faculty_codes(faculty_codes_str)
            
            print(f"    Parsed course codes: {course_codes}")
            print(f"    Parsed faculty codes: {faculty_codes}")
            
            # Special case: If we have multiple course codes and the same number of faculty codes,
            # map them one-to-one (as specified in the requirements)
            if len(course_codes) > 1 and len(faculty_codes) >= len(course_codes):
                print(f"    Using 1:1 mapping for {len(course_codes)} courses to {len(faculty_codes)} faculty")
                # Map each course to its corresponding faculty
                for i, course_code in enumerate(course_codes):
                    if i < len(faculty_codes):
                        faculty_code = faculty_codes[i]
                        faculty_name = code_to_name.get(faculty_code)
                        
                        if faculty_name:
                            pairs.append((faculty_name, course_code, course_name))
                            print(f"      Added: {faculty_name} -> {course_code} {course_name}")
                        else:
                            print(f"Warning: Faculty code '{faculty_code}' not found in mapping ({source_info})", file=sys.stderr)
            else:
                print(f"    Using regular mapping: each of {len(faculty_codes)} faculty teaches all {len(course_codes)} courses")
                # Regular case: Each faculty teaches the course(s)
                for course_code in course_codes:
                    for faculty_code in faculty_codes:
                        faculty_name = code_to_name.get(faculty_code)
                        
                        if faculty_name:
                            pairs.append((faculty_name, course_code, course_name))
                            print(f"      Added: {faculty_name} -> {course_code} {course_name}")
                        else:
                            print(f"Warning: Faculty code '{faculty_code}' not found in mapping ({source_info})", file=sys.stderr)

    # Process all tables except the last one (which contains faculty mapping)
    print("Processing course tables...")
    for table_idx, table in enumerate(doc.tables[:-1]):
        print(f"Processing table {table_idx + 1}...")
        
        # Process each cell in the table
        for row_idx, row in enumerate(table.rows):
            for col_idx, cell in enumerate(row.cells):
                # Get cell text
                cell_text = cell.text.strip()
                
                # Skip empty cells or cells that are too short
                if not cell_text or len(cell_text) < 8:
                    continue
                
                # Skip header-like cells
                if any(header in cell_text.lower() for header in ['forenoon', 'afternoon', 'slots', 'time']):
                    continue
                
                # Process the cell text
                source_info = f"Table {table_idx + 1}, Row {row_idx + 1}, Col {col_idx + 1}"
                process_course_text(cell_text, source_info)

    # Process paragraphs for lab courses and projects
    print("Processing paragraphs for lab courses and projects...")
    
    # Process all paragraphs
    for para_idx, para in enumerate(doc.paragraphs):
        para_text = para.text.strip()
        if not para_text:
            continue
            
        # Check if this paragraph contains lab course information
        if ('Lab' in para_text or 'Project' in para_text) and ':' in para_text and re.search(r'[A-Z]{2}\d{4}[A-Z]?', para_text):
            source_info = f"Paragraph {para_idx + 1}"
            process_lab_paragraph(para_text, source_info)
        # Also check for regular course format in paragraphs
        elif re.search(r'[A-Z]{2}\d{4}[A-Z]?', para_text) and '[' in para_text:
            source_info = f"Paragraph {para_idx + 1}"
            process_course_text(para_text, source_info)

    # Remove duplicates while preserving order
    unique_pairs = []
    seen = set()
    for pair in pairs:
        if pair not in seen:
            unique_pairs.append(pair)
            seen.add(pair)
    
    print(f"Extracted {len(unique_pairs)} unique course-faculty assignments")
    return unique_pairs


def normalize_name(name):
    """Normalize a name for fuzzy matching"""
    return re.sub(r'[^A-Za-z]', '', name).lower()


def main():
    """Main function"""
    # Check command line arguments
    if len(sys.argv) != 3:
        print(f"Usage: {sys.argv[0]} <timetable.docx> <output.csv>")
        print("Example: python tteduandscrape_final_fix.py TT_Mons2025.docx output.csv")
        sys.exit(1)
    
    timetable_path, output_csv = sys.argv[1], sys.argv[2]
    
    print("Step 1: Fetching faculty usernames from NITC website...")
    # Fetch faculty name to username mapping from website
    name_to_username = fetch_faculty_usernames()
    print(f"Found {len(name_to_username)} faculty members from website")
    
    print("Step 2: Parsing course assignments from DOCX file...")
    # Parse course data from DOCX file
    course_data = parse_docx_courses(timetable_path)
    
    print("Step 3: Matching faculty names and generating CSV...")
    # Create normalized name mapping for fuzzy matching
    all_faculty_names = list(name_to_username.keys())
    normalized_name_map = {normalize_name(name): name for name in all_faculty_names}
    
    # Write results to CSV file
    with open(output_csv, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        
        # Write header row
        writer.writerow(['faculty_username', 'course_code', 'full_course_name'])
        
        matched_count = 0
        unmatched_count = 0
        
        # Process each course-faculty assignment
        for faculty_name, course_code, course_name in course_data:
            # Try direct lookup first
            username = name_to_username.get(faculty_name)
            
            if not username:
                # Try fuzzy matching using normalized names
                normalized_faculty = normalize_name(faculty_name)
                
                # First try exact normalized match
                matched_name = normalized_name_map.get(normalized_faculty)
                
                if not matched_name:
                    # Try fuzzy matching with similarity threshold
                    close_matches = difflib.get_close_matches(
                        normalized_faculty, 
                        list(normalized_name_map.keys()), 
                        n=1, 
                        cutoff=0.6
                    )
                    
                    if close_matches:
                        matched_name = normalized_name_map[close_matches[0]]
                
                if matched_name:
                    username = name_to_username[matched_name]
                    print(f"Fuzzy matched '{faculty_name}' to '{matched_name}' -> '{username}'")
                    matched_count += 1
                else:
                    print(f"Warning: No username found for faculty '{faculty_name}'")
                    unmatched_count += 1
                    continue
            else:
                matched_count += 1
            
            # Generate course key in required format
            course_key = f"m202526{course_code.lower()}"
            
            # Generate full course name
            full_course_name = f"{course_code} {course_name}"
            
            # Write row to CSV
            writer.writerow([username, course_key, full_course_name])
    
    print(f"\nStep 4: Complete!")
    print(f"Successfully matched {matched_count} assignments")
    print(f"Failed to match {unmatched_count} assignments")
    print(f"Results written to {output_csv}")


if __name__ == '__main__':
    main()