import json
from docx import Document
from typing import Dict, List, Any
from collections import defaultdict
import re
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

class AbbreviationFinder:
    def __init__(self):
        """Initialize the abbreviation finder."""
        # Pattern for finding serial commas (Oxford commas)
        self.serial_comma_pattern = r'\b[^,.]+,\s+[^,.]+,\s+(?:and|or)\s+[^,.]+\b'
        
        # Regular expression patterns for finding abbreviations
        self.abbr_patterns = [
            # Pattern for abbreviations with parentheses: Word (ABBR) or (ABBR) Word
            r'\b([A-Za-z][A-Za-z\s-]+(?:[A-Za-z]|\b))\s*\(([A-Z][A-Z0-9]{1,})\)(?!\w)',  # Word (ABBR)
            r'\b\(([A-Z][A-Z0-9]{1,})\)\s*([A-Za-z][A-Za-z\s-]+(?:[A-Za-z]|\b))(?!\w)',  # (ABBR) Word
            # Pattern for standalone uppercase abbreviations with context
            r'\b([A-Za-z][A-Za-z\s-]+(?:[A-Za-z]|\b))\s+\(([A-Z][A-Z0-9]{1,})\)(?:(?:\s+|,\s+|;\s+)\2\b)+',  # Full form (ABBR) followed by ABBR
            # Pattern for abbreviations with periods
            r'\b(?:[A-Z]\.){2,}(?:[A-Z])?\b',  # Matches A.B. or A.B.C.
            # Pattern for multi-word abbreviations
            r'\b([A-Za-z][A-Za-z\s-]+(?:[A-Za-z]|\b))\s*\(([A-Z][.-]?(?:[A-Z][.-]?)+)\)(?!\w)',  # Matches multi-word abbr like C.E.O.
            # Pattern for abbreviations followed by definition
            r'\b([A-Z][A-Z0-9]{1,})\s*[-:]\s*([A-Za-z][A-Za-z\s-]+(?:[A-Za-z]|\b))(?!\w)'  # Matches ABC: Definition
        ]
        # Pattern for author initials
        self.author_initial_patterns = [
            r'\b[A-Z]\.[A-Z]\.\s*[A-Z]?\.',  # A.B. or A.B.C.
            r'\b[A-Z]\.[s-][A-Z]\.',  # A. B. or A.-B.
            r'\b[A-Za-z]+,\s*[A-Z]\.',  # Smith, J.
            r'\b[A-Z]\.[s-][A-Z]\.[s-][A-Z]\.'  # A. B. C.
        ]
        # Common document section headings to exclude
        self.excluded_headings = {
            'abstract', 'introduction', 'background', 'methodology', 'methods',
            'results', 'discussion', 'conclusion', 'conclusions', 'references',
            'acknowledgments', 'acknowledgements', 'appendix', 'appendices',
            'materials', 'objectives', 'hypothesis', 'literature review',
            'experimental', 'data analysis', 'findings', 'recommendations',
            'bibliography', 'keywords', 'key words'
        }

    def _is_author_initial(self, text: str) -> bool:
        """Check if the text matches author initial patterns.

        Args:
            text (str): Text to check

        Returns:
            bool: True if text matches author initial patterns
        """
        return any(re.search(pattern, text) for pattern in self.author_initial_patterns)

    def find_abbreviations(self, input_path: str) -> Dict[str, Any]:
        """Find abbreviations in a Word document.

        Args:
            input_path (str): Path to the input Word document

        Returns:
            Dict[str, Any]: Dictionary containing found abbreviations
        """
        try:
            # Load the Word document
            doc = Document(input_path)

            # Collect all abbreviation-full form pairs in the document
            abbr_to_full = {}
            full_to_abbr = {}
            # Scan all paragraphs and table cells for explicit pairs
            all_texts = [para.text for para in doc.paragraphs]
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        all_texts.append(cell.text)
            for text in all_texts:
                # Skip headings and author initials
                if self._is_author_initial(text) or text.strip().lower() in self.excluded_headings:
                    continue

                # Find words inside angular brackets to exclude them completely
                excluded_bracket_words = set()
                for bracket_content in re.findall(r'<([a-zA-Z0-9_\s.-]{1,100})>', text):
                    # 1. Add raw content
                    excluded_bracket_words.add(bracket_content.strip().upper())
                    excluded_bracket_words.add(bracket_content.strip().lower())
                    excluded_bracket_words.add(bracket_content.strip())
                    
                    # 2. Add cleaned content (alphanumeric only)
                    cleaned = re.sub(r"[^-a-zA-Z0-9 ]", "", bracket_content).strip()
                    excluded_bracket_words.add(cleaned.upper())
                    excluded_bracket_words.add(cleaned.lower())
                    excluded_bracket_words.add(cleaned)
                    
                    # 3. Add individual words
                    for w in re.split(r'[^a-zA-Z0-9-]+', bracket_content):
                        if w:
                            excluded_bracket_words.add(w.upper())
                            excluded_bracket_words.add(w.lower())
                            excluded_bracket_words.add(w)

                # Clean text of angular brackets
                clean_text = re.sub(r'<[a-zA-Z0-9_\s.-]{1,100}>', ' ', text)

                # Look for explicit pairs
                for pattern in self.abbr_patterns[:2]:
                    matches = re.finditer(pattern, clean_text)
                    for match in matches:
                        if len(match.groups()) == 2:
                            if pattern.startswith('\\b([A-Za-z'):
                                full_form, abbr = match.groups()
                            else:
                                abbr, full_form = match.groups()

                            if (abbr in excluded_bracket_words or
                                abbr.upper() in excluded_bracket_words or
                                abbr.lower() in excluded_bracket_words):
                                continue

                            abbr_key = abbr.upper()
                            if abbr_key not in abbr_to_full:
                                abbr_to_full[abbr_key] = full_form.strip()
                            if full_form.strip() not in full_to_abbr:
                                full_to_abbr[full_form.strip()] = abbr_key

            # Initialize results dictionary with temporary storage
            temp_results = defaultdict(lambda: {
                'abbreviation': '',
                'full_form': None,
                'count': 0
            })

            # Process document content
            for para_index, paragraph in enumerate(doc.paragraphs, 1):
                self._process_paragraph(paragraph, para_index, temp_results)

            # Process tables
            for table_index, table in enumerate(doc.tables, 1):
                for row_index, row in enumerate(table.rows, 1):
                    for cell_index, cell in enumerate(row.cells, 1):
                        self._process_text(cell.text, f'Table {table_index}, Row {row_index}, Cell {cell_index}', temp_results)

            # After collecting, fill in missing full forms from abbr_to_full
            for abbr_key, data in temp_results.items():
                if not data['full_form'] and abbr_key in abbr_to_full:
                    data['full_form'] = abbr_to_full[abbr_key]

            # Convert temporary results to final format
            results = {
                'document_path': input_path,
                'abbreviations_found': [
                    {
                        'abbreviation': data['abbreviation'],
                        'full_form': data['full_form'],
                        'occurrences': data['count']
                    }
                    for data in temp_results.values()
                ]
            }

            logging.info(f'Successfully found abbreviations in {input_path}')
            # Find serial commas
            serial_comma_results = self.find_serial_commas(doc)
            results['serial_commas_found'] = serial_comma_results
            
            return results

        except Exception as e:
            logging.error(f'Error analyzing document: {str(e)}')
            raise
            
    def find_serial_commas(self, doc: Document) -> List[Dict[str, Any]]:
        """Find serial commas (Oxford commas) in the document.
        
        Args:
            doc (Document): The Word document object
            
        Returns:
            List[Dict[str, Any]]: List of dictionaries containing serial comma information
        """
        serial_commas = []
        
        # Process paragraphs
        for para_index, paragraph in enumerate(doc.paragraphs, 1):
            matches = re.finditer(self.serial_comma_pattern, paragraph.text)
            for match in matches:
                serial_commas.append({
                    'text': match.group(),
                    'location': f'Paragraph {para_index}'
                })
        
        # Process tables
        for table_index, table in enumerate(doc.tables, 1):
            for row_index, row in enumerate(table.rows, 1):
                for cell_index, cell in enumerate(row.cells, 1):
                    matches = re.finditer(self.serial_comma_pattern, cell.text)
                    for match in matches:
                        serial_commas.append({
                            'text': match.group(),
                            'location': f'Table {table_index}, Row {row_index}, Cell {cell_index}'
                        })
        
        return serial_commas

    def _process_paragraph(self, paragraph: Any, para_index: int, results: Dict[str, Any]) -> None:
        """Process a paragraph to find abbreviations.

        Args:
            paragraph (Any): The paragraph object
            para_index (int): Index of the paragraph
            results (Dict[str, Any]): Results dictionary to update
        """
        text = paragraph.text
        location = f'Paragraph {para_index}'
        self._process_text(text, location, results)

    def _process_table(self, table: Any, table_index: int, results: Dict[str, Any]) -> None:
        """Process a table to find abbreviations.

        Args:
            table (Any): The table object
            table_index (int): Index of the table
            results (Dict[str, Any]): Results dictionary to update
        """
        for row_index, row in enumerate(table.rows, 1):
            for cell_index, cell in enumerate(row.cells, 1):
                text = cell.text
                location = f'Table {table_index}, Row {row_index}, Cell {cell_index}'
                self._process_text(text, location, results)

    def _process_text(self, text: str, location: str, results: Dict[str, Any]) -> None:
        """Process text to find and consolidate abbreviations.

        Args:
            text (str): Text to process
            location (str): Location in the document
            results (Dict[str, Any]): Results dictionary to update
        """
        # Skip if text appears to be author initials
        if self._is_author_initial(text):
            return

        # Skip if text is a common document heading
        if text.strip().lower() in self.excluded_headings:
            return

        # Find words inside angular brackets to exclude them completely
        excluded_bracket_words = set()
        for bracket_content in re.findall(r'<([a-zA-Z0-9_\s.-]{1,100})>', text):
            # 1. Add raw content
            excluded_bracket_words.add(bracket_content.strip().upper())
            excluded_bracket_words.add(bracket_content.strip().lower())
            excluded_bracket_words.add(bracket_content.strip())
            
            # 2. Add cleaned content (alphanumeric only)
            cleaned = re.sub(r"[^-a-zA-Z0-9 ]", "", bracket_content).strip()
            excluded_bracket_words.add(cleaned.upper())
            excluded_bracket_words.add(cleaned.lower())
            excluded_bracket_words.add(cleaned)
            
            # 3. Add individual words
            for w in re.split(r'[^a-zA-Z0-9-]+', bracket_content):
                if w:
                    excluded_bracket_words.add(w.upper())
                    excluded_bracket_words.add(w.lower())
                    excluded_bracket_words.add(w)

        # Clean text of angular brackets
        clean_text = re.sub(r'<[a-zA-Z0-9_\s.-]{1,100}>', ' ', text)

        # Process each abbreviation pattern
        for pattern_index, pattern in enumerate(self.abbr_patterns):
            matches = re.finditer(pattern, clean_text)
            for match in matches:
                if len(match.groups()) == 2:
                    # Handle different pattern formats
                    if pattern_index <= 2:  # Standard patterns (Word ABBR) or (ABBR Word)
                        if pattern.startswith('\\b([A-Za-z'):
                            full_form, abbr = match.groups()
                        else:
                            abbr, full_form = match.groups()
                    elif pattern_index == 3:  # Period-separated abbreviations
                        abbr = match.group()
                        full_form = None
                    elif pattern_index == 4:  # Multi-word abbreviations
                        full_form, abbr = match.groups()
                    else:  # ABBR: Definition format
                        abbr, full_form = match.groups()

                    # Skip common headings, author initials, and bracketed words
                    if (abbr.lower() in self.excluded_headings or 
                        (full_form and full_form.lower() in self.excluded_headings) or
                        self._is_author_initial(abbr) or
                        abbr in excluded_bracket_words or
                        abbr.upper() in excluded_bracket_words or
                        abbr.lower() in excluded_bracket_words):
                        continue

                    # Normalize abbreviation key
                    abbr_key = abbr.upper().replace('.', '')
                    
                    # Update results
                    if not results[abbr_key]['abbreviation']:
                        results[abbr_key]['abbreviation'] = abbr
                    
                    # Update full form if we have a better match
                    if full_form:
                        full_form = full_form.strip()
                        if not results[abbr_key]['full_form'] or \
                           (len(full_form) > len(results[abbr_key]['full_form']) and \
                            full_form.lower() != results[abbr_key]['full_form'].lower()):
                            results[abbr_key]['full_form'] = full_form
                    
                    results[abbr_key]['count'] += 1

def main():
    """Main function to run the abbreviation finder."""
    import argparse
    
    parser = argparse.ArgumentParser(
        description='Find abbreviations in Word documents'
    )
    parser.add_argument('input_file', help='Path to input Word document')
    parser.add_argument('output_file', help='Path to output JSON file')
    parser.add_argument('--debug', action='store_true', 
                      help='Enable debug mode to show processing details')
    #args = parser.parse_args()
    args = ['Sample_for_Abbreviation.docx', 'output.json']
    finder = AbbreviationFinder()
    try:
        results = finder.find_abbreviations(args[0])
        
        # Save results to JSON file
        with open(args[1], 'w', encoding='utf-8') as f:
            json.dump(results, f, indent=2)
            
        print(f'Successfully analyzed {args[0]} and saved results to {args[1]}')
    except Exception as e:
        print(f'Error: {str(e)}')
        exit(1)

if __name__ == '__main__':
    main()