import docx2txt
from docx import Document
from collections import defaultdict
import os
import json
import re
import argparse
import sys
import zipfile
from xml.etree import ElementTree
from typing import Dict, Any, List

def load_dictionary(dict_path: str) -> set:
    """Load dictionary words from a file"""
    with open(dict_path, 'r', encoding='utf-8') as f:
        return set(word.strip().lower() for word in f.readlines() if word.strip())

def extract_text_from_docx(file_path: str) -> str:
    """Extract text content from DOCX file"""
    try:
        return docx2txt.process(file_path)
    except Exception as e:
        raise Exception(f"DOCX processing failed: {str(e)}")

def analyze_docx(file_path: str, us_words: set, uk_words: set) -> Dict[str, Any]:
    """Analyze DOCX content for UK/US word matches"""
    try:
        text = extract_text_from_docx(file_path)
        words = text.lower().split()
        
        results = {
            'us': {'total': 0, 'words': defaultdict(int)},
            'uk': {'total': 0, 'words': defaultdict(int)},
            'total_words_in_document': len(words)
        }
        
        for word in words:
            clean_word = word.strip(".,!?()\"'")
            if clean_word:
                if clean_word in us_words:
                    results['us']['words'][clean_word] += 1
                    results['us']['total'] += 1
                if clean_word in uk_words:
                    results['uk']['words'][clean_word] += 1
                    results['uk']['total'] += 1
        
        results['us']['words'] = dict(results['us']['words'])
        results['uk']['words'] = dict(results['uk']['words'])
        
        return results
    except Exception as e:
        raise Exception(f"Analysis failed: {str(e)}")

def generate_report_content(results: dict) -> str:
    """Generate report content as string"""
    report_content = "=== DOCUMENT ANALYSIS REPORT ===\n"
    report_content += f"Total_Words: {results['total_words_in_document']}\n\n"
    
    report_content += "US Words \n"
    report_content += f"Total Words : {results['us']['total']}\n"
    for word, count in results['us']['words'].items():
        report_content += f"{word}: {count}\n"
    
    report_content += "\nUK Words \n"
    report_content += f"Total Words : {results['uk']['total']}\n"
    for word, count in results['uk']['words'].items():
        report_content += f"{word}: {count}\n"
    
    return report_content

def extract_abbreviation_list(document_content: str) -> Dict[str, Any]:
    """Extract abbreviations from document content"""
    try:
        abbreviation_list = []
        abbreviation_count = []
        all_found_abbreviations = set()

        # Find words inside angular brackets to exclude them completely
        excluded_bracket_words = set()
        for bracket_content in re.findall(r'<([a-zA-Z0-9_\s.-]{1,100})>', document_content):
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

        # Remove angular brackets content for searching/counting
        clean_content = re.sub(r'<[a-zA-Z0-9_\s.-]{1,100}>', ' ', document_content)

        pattern = r"(([A-Z]+ [0-9]+)|\b([A-Z]+(-?[A-Z0-9]+)*(s)?)\b)|([a-z]*[A-Z][a-z]*)+"
        matches = re.findall(pattern, clean_content)

        for match_group in matches:
            found_abbreviation = next((m for m in match_group if m), "").strip()
            found_abbreviation = re.sub(r"[^-a-zA-Z0-9 ]", "", found_abbreviation)
            found_abbreviation = re.sub(r"(^[- ]|[- ]$)", "", found_abbreviation)

            if found_abbreviation.endswith("s"):
                found_abbreviation = found_abbreviation[:-1]

            if (2 <= len(found_abbreviation) <= 6 and
                    len(re.findall(r"[A-Z]", found_abbreviation)) > 1 and
                    found_abbreviation not in all_found_abbreviations and
                    found_abbreviation not in excluded_bracket_words and
                    found_abbreviation.upper() not in excluded_bracket_words and
                    found_abbreviation.lower() not in excluded_bracket_words):

                count_pattern = rf"\b{re.escape(found_abbreviation)}s?\b"
                count = len(re.findall(count_pattern, clean_content))

                abbreviation_list.append(found_abbreviation)
                abbreviation_count.append(count)
                all_found_abbreviations.add(found_abbreviation)

        expansion_array = resolve_search_priority(abbreviation_list, clean_content)

        result = []
        for idx, abbr in enumerate(abbreviation_list):
            expansion = expansion_array[idx] if idx < len(expansion_array) else ""
            expansion = expansion.lstrip('|').split('|')[0] if expansion else ""
            result.append({
                "abbreviation": abbr,
                "full_form": expansion,
                "occurrences": abbreviation_count[idx]
            })

        return {"abbreviations_found": result}
    except Exception as e:
        raise Exception(f"Error extracting abbreviations: {str(e)}")

# [Include all the other helper functions from your original code here:
# resolve_search_priority, get_expansion_from_document, new_approach_retrieve_expansion,
# filter_searched_expansion, new_approach_expansion_as_is_form, check_abbreviation_match,
# build_single_word_regex_pattern, clean_surrounding_nonword, search_pattern_1, search_pattern]
def resolve_search_priority(abbreviation_list, document_full_content):
    expansion_array = [''] * (len(abbreviation_list) + 1)
    #db_expansion_array = [''] * (len(abbreviation_list) + 1)

    try:
        expansion_array = get_expansion_from_document(abbreviation_list, document_full_content)

    except Exception as e:
        print(f"Error resolving abbreviation expansions in resolve_search_priority def: {e}")

    return expansion_array

def get_expansion_from_document(abbreviation_list, document_full_content):
    """
    Retrieves possible expansions for abbreviations from the document content.

    :param abbreviation_list: List of abbreviations (e.g. ['NASA', 'WHO']).
    :param document_full_content: The full document text.
    :return: List of expansions (same length as abbreviation_list + 1).
    """
    try:
        # Initialize expansion list (+1 for compatibility with original C# structure)
        expansion_array = [''] * (len(abbreviation_list) + 1)

        # Retrieve expansions using a custom heuristic
        expansion_array = new_approach_retrieve_expansion(abbreviation_list, document_full_content, expansion_array)

        # Add XML tags or formatting
        #expansion_array = add_xml_tags(expansion_array, document_full_content)

        return expansion_array

    except Exception as e:
        print(f"Error in get_expansion_from_document in: {e}")
        return [''] * (len(abbreviation_list) + 1)

def new_approach_retrieve_expansion(abbreviation_list, document_text, expansion_array=None):
    try:
        if expansion_array is None:
            expansion_array = [''] * (len(abbreviation_list) + 1)

        preposition_list = 'at|by|for|of|in|on|to|and|the|at'
        list_of_all_expansions = set()

        for idx in range(0, len(abbreviation_list)):
            abbr = abbreviation_list[idx]
            filtered_text = ''
            is_expansion = False

            if len(abbr) > 2:
                # Try abbreviation as-is (e.g., "World Health Organization (WHO)")
                filtered_text = new_approach_expansion_as_is_form(document_text, abbr, preposition_list)
                filtered_text = clean_surrounding_nonword(filtered_text)

                if filtered_text and filtered_text not in list_of_all_expansions:
                    list_of_all_expansions.add(filtered_text)
                    expansion_array[idx] += f"|{filtered_text.strip()}"
                else:
                    pattern = search_pattern_1(abbr, preposition_list) if re.search(r'[a-z]', abbr) else search_pattern(abbr, preposition_list)
                    matches = re.findall(pattern, document_text, flags=re.IGNORECASE)
                    for match_tuple in matches:
                        for match in match_tuple:
                            if not isinstance(match, str):
                                continue
                            filtered_text = clean_surrounding_nonword(match)
                            if not filtered_text or filtered_text in list_of_all_expansions:
                                continue

                            s_text = filtered_text
                            if s_text.lower().startswith(abbr.lower()):
                                s_text = s_text[len(abbr):].strip()

                            if len(s_text) > len(abbr):
                                i_pos = 0
                                is_expansion = True
                                for c in abbr:
                                    found = re.search(re.escape(c), s_text[i_pos:], flags=re.IGNORECASE)
                                    if found:
                                        i_pos += found.start() + 1
                                    else:
                                        is_expansion = False
                                        break

                                if is_expansion:
                                    list_of_all_expansions.add(filtered_text)
                                    expansion_array[idx] += f"|{filtered_text.strip()}"
                                    break
            elif len(abbr) == 2:
                filtered_text = new_approach_expansion_as_is_form(document_text, abbr, preposition_list)
                filtered_text = clean_surrounding_nonword(filtered_text)

                if filtered_text and filtered_text not in list_of_all_expansions:
                    s_text = filtered_text
                    if s_text.lower().startswith(abbr.lower()):
                        s_text = s_text[len(abbr):].strip()

                    if len(s_text) > len(abbr):
                        i_pos = 0
                        is_expansion = True
                        for c in abbr:
                            found = re.search(re.escape(c), s_text[i_pos:], flags=re.IGNORECASE)
                            if found:
                                i_pos += found.start() + 1
                            else:
                                is_expansion = False
                                break

                        if is_expansion:
                            list_of_all_expansions.add(filtered_text)
                            expansion_array[idx] += f"|{filtered_text.strip()}"
                        else:
                            expansion_array[idx] = ""
                    else:
                        expansion_array[idx] = ""
                else:
                    expansion_array[idx] = ""

        filter_searched_expansion(expansion_array, abbreviation_list, document_text)
        return expansion_array
    except Exception as e:
        print(f"Error in new_approach_retrieve_expansion def: {e}")

def search_pattern(abbreviation: str, preposition_list: str) -> str:
    try:
        abbr_chars = list(abbreviation)
        
        # Start the pattern with the first character's word
        pattern = (
            f"[{abbr_chars[0].upper()}{abbr_chars[0].lower()}][a-zA-Z\\'\\’\\-]+"
            f"({preposition_list})?"
        )

        for i in range(1, len(abbr_chars)):
            ch = abbr_chars[i]

            if i < len(abbr_chars) - 1:
                if re.match(r"[0-9]", ch):
                    pattern += (
                        f"(\\s)?([{ch.upper()}][a-z\\'\\’\\-]*)"
                        f"({preposition_list})?(\\s)?"
                    )
                elif re.match(r"[- ]", ch):
                    pattern += (
                        f"(\\s)?([{ch.upper()}][a-z\\'\\’\\-]*)?"
                        f"({preposition_list})?(\\s)?"
                    )
                else:
                    pattern += (
                        f"(\\s)+[{ch.upper()}{ch.lower()}][a-z\\'\\’\\-]+"
                        f"({preposition_list})?(\\s)?"
                    )
            else:
                if re.match(r"[0-9]", ch):
                    pattern += (
                        f"(\\s)?([{ch.lower()}][a-zA-Z\\-\\'\\’]*)(\\W)"
                    )
                elif re.match(r"[- ]", ch):
                    pattern += (
                        f"(\\s)?([{ch.lower()}][a-zA-Z\\-\\'\\’]*)?(\\W)"
                    )
                else:
                    pattern += (
                        f"(\\s)+[{ch.upper()}{ch.lower()}][a-zA-Z\\-\\'\\’]+(\\W)"
                    )

        pattern = r"(\W)" + pattern
        return pattern
    except Exception as e:
        print(f"Error resolving abbreviation expansions in search_pattern def: {e}")
                       
def search_pattern_1(abbreviation: str, preposition_list: str) -> str:
    try:
        abbr_chars = list(abbreviation)
        pattern = (
            f"[{abbr_chars[0].upper()}{abbr_chars[0].lower()}][a-z\\'\\’\\-]*"
            f"({preposition_list})?"
        )

        for i in range(1, len(abbr_chars)):
            ch = abbr_chars[i]

            if i < len(abbr_chars) - 1:
                if re.match(r"[0-9]", ch):
                    pattern += (
                        f"(\\s)?([{ch.upper()}][a-z\\'\\’\\-]*)"
                        f"({preposition_list})?(\\s)?"
                    )
                elif re.match(r"[- ]", ch):
                    pattern += (
                        f"(\\s)?([{ch.upper()}][a-z\\'\\’\\-]*)?"
                        f"({preposition_list})?(\\s)?"
                    )
                else:
                    pattern += (
                        f"(\\s)?[{ch.upper()}{ch.lower()}][a-z\\'\\’\\-]*"
                        f"({preposition_list})?(\\s)?"
                    )
            else:
                if re.match(r"[0-9]", ch):
                    pattern += (
                        f"(\\s)?([{ch.lower()}][a-zA-Z\\-\\'\\’]*)(\\W)"
                    )
                elif re.match(r"[- ]", ch):
                    pattern += (
                        f"(\\s)?([{ch.lower()}][a-zA-Z\\-\\'\\’]*)?(\\W)"
                    )
                else:
                    pattern += (
                        f"(\\s)?[{ch.upper()}{ch.lower()}][a-z\\-\\'\\’]*(\\W)"
                    )

        pattern = r"(\W)" + pattern
        return pattern
    except Exception as e:
        print(f"Error resolving abbreviation expansions in search_pattern_1 def: {e}")

def clean_surrounding_nonword(text):
    try:
        if not text:
            return ""
        return re.sub(r"^\W+|\W+$", "", text)
    except Exception as e:
        print(f"Error in clean_surrounding_nonword in clean_surrounding_nonword def: {e}")

def reconstruct_paragraph_text(para) -> str:
    """Reconstruct paragraph text replacing endnote/footnote elements with markers."""
    parts = []
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    
    def traverse(node):
        # Skip deleted text
        if node.tag.endswith('del'):
            return
            
        if node.tag.endswith('r'):
            # Check for footnote/endnote references
            if node.find('.//w:footnoteReference', ns) is not None:
                parts.append("[[^f]]")
            elif node.find('.//w:endnoteReference', ns) is not None:
                parts.append("[[^e]]")
            else:
                # Extract text from w:t elements
                for t in node.findall('.//w:t', ns):
                    if t.text:
                        parts.append(t.text)
        else:
            for child in node:
                traverse(child)
                
    traverse(para._p)
    return "".join(parts)


def extract_footnotes_and_endnotes(docx_path: str) -> List[Dict[str, str]]:
    """Extract footnote and endnote text content from raw XML parts."""
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    results = []
    
    try:
        with zipfile.ZipFile(docx_path) as z:
            # Footnotes
            if 'word/footnotes.xml' in z.namelist():
                xml_content = z.read('word/footnotes.xml')
                root = ElementTree.fromstring(xml_content)
                for fn in root.findall('.//w:footnote', ns):
                    fn_id = fn.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
                    if fn_id and int(fn_id) > 0:
                        texts = [t.text for t in fn.findall('.//w:t', ns) if t.text]
                        results.append({
                            'id': fn_id,
                            'type': 'Footnote',
                            'text': "".join(texts)
                        })
            # Endnotes
            if 'word/endnotes.xml' in z.namelist():
                xml_content = z.read('word/endnotes.xml')
                root = ElementTree.fromstring(xml_content)
                for en in root.findall('.//w:endnote', ns):
                    en_id = en.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
                    if en_id and int(en_id) > 0:
                        texts = [t.text for t in en.findall('.//w:t', ns) if t.text]
                        results.append({
                            'id': en_id,
                            'type': 'Endnote',
                            'text': "".join(texts)
                        })
    except Exception as e:
        print(f"Warning: Could not read footnotes/endnotes: {e}")
    return results

def load_formatting_rules() -> Dict[str, Dict[str, str]]:
    """Load formatting rules from config file if present, otherwise fall back to defaults."""
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        
    config_path = os.path.join(base_dir, 'formatting_rules.json')
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"Warning: Could not load formatting_rules.json: {e}")
    return {
        'double_spaces': {
            'pattern': r'  +',
            'message': 'Double Spaces',
            'target': 'both'
        },
        'serial_commas': {
            'pattern': r'\b[^,.]+,\s+[^,.]+,\s+(?:and|or)\s+[^,.]+\b',
            'message': 'Serial Commas',
            'target': 'body'
        },
        'section_words': {
            'pattern': r'\b(Section|Part|Chapter|Epilogue|Preface|Acknowledgement)s?\b',
            'message': 'Prohibited section terms',
            'target': 'both'
        },
        'space_period_para': {
            'pattern': r'\. +$',
            'message': 'Space between period and paragraph mark',
            'target': 'body'
        },
        'double_punctuation': {
            'pattern': r',{2,}|::+|;;+|(?<!\.)\.\.(?!\.)',
            'message': 'Double punctuation detected',
            'target': 'both'
        },
        'space_before_endnote': {
            'pattern': r' +\[\[\^e\]\]',
            'message': 'Space before endnote citation',
            'target': 'body'
        },
        'no_space_endnote_text': {
            'pattern': r'\[\[\^e\]\][a-zA-Z]',
            'message': 'No space between endnote and text',
            'target': 'body'
        },
        'endnote_colon': {
            'pattern': r'\[\[\^e\]\]:',
            'message': 'Endnote citation followed by colon',
            'target': 'body'
        },
        'endnote_semicolon': {
            'pattern': r'\[\[\^e\]\];',
            'message': 'Endnote citation followed by semicolon',
            'target': 'body'
        },
        'endnote_comma': {
            'pattern': r'\[\[\^e\]\],',
            'message': 'Endnote citation followed by comma',
            'target': 'body'
        },
        'comma_after_quotes_in_endnotes': {
            'pattern': r'["”],',
            'message': 'Comma placed outside of double quotes in endnotes',
            'target': 'endnotes'
        },
        'hyphen_spaced': {
            'pattern': r' - ',
            'message': 'Hyphen surrounded by spaces',
            'target': 'both'
        },
        'spaced_endash': {
            'pattern': r'\s+\u2013\s+|\s+\u2014\s+|--',
            'message': 'Double hyphen or spaced en-dash/em-dash used',
            'target': 'both'
        }
    }

FORMATTING_RULES = load_formatting_rules()

def check_body_rule(doc, rule_id: str, pattern: str, message: str) -> List[Dict[str, Any]]:
    findings = []
    regex = re.compile(pattern)
    
    # Check paragraphs
    for para_index, paragraph in enumerate(doc.paragraphs, 1):
        virtual_text = reconstruct_paragraph_text(paragraph)
        for match in regex.finditer(virtual_text):
            start = max(0, match.start() - 15)
            end = min(len(virtual_text), match.end() + 15)
            snippet = virtual_text[start:end]
            snippet_highlighted = snippet[:match.start() - start] + f"[{match.group(0)}]" + snippet[match.end() - start:]
            findings.append({
                'rule_id': rule_id,
                'message': message,
                'text': snippet_highlighted.strip(),
                'location': f'Paragraph {para_index}'
            })
            
    # Check tables
    for table_index, table in enumerate(doc.tables, 1):
        for row_index, row in enumerate(table.rows, 1):
            for cell_index, cell in enumerate(row.cells, 1):
                for para_index, paragraph in enumerate(cell.paragraphs, 1):
                    virtual_text = reconstruct_paragraph_text(paragraph)
                    for match in regex.finditer(virtual_text):
                        start = max(0, match.start() - 15)
                        end = min(len(virtual_text), match.end() + 15)
                        snippet = virtual_text[start:end]
                        snippet_highlighted = snippet[:match.start() - start] + f"[{match.group(0)}]" + snippet[match.end() - start:]
                        findings.append({
                            'rule_id': rule_id,
                            'message': message,
                            'text': snippet_highlighted.strip(),
                            'location': f'Table {table_index}, Row {row_index}, Cell {cell_index}, Para {para_index}'
                        })
    return findings

def check_endnotes_rule(doc_path: str, rule_id: str, pattern: str, message: str) -> List[Dict[str, Any]]:
    findings = []
    regex = re.compile(pattern)
    notes = extract_footnotes_and_endnotes(doc_path)
    
    for note in notes:
        text = note['text']
        for match in regex.finditer(text):
            start = max(0, match.start() - 15)
            end = min(len(text), match.end() + 15)
            snippet = text[start:end]
            snippet_highlighted = snippet[:match.start() - start] + f"[{match.group(0)}]" + snippet[match.end() - start:]
            findings.append({
                'rule_id': rule_id,
                'message': message,
                'text': snippet_highlighted.strip(),
                'location': f'{note["type"]} ID {note["id"]}'
            })
    return findings

def find_space_period_para(doc) -> List[Dict[str, Any]]:
    """Flags paragraphs ending with a period followed by space(s)."""
    rule = FORMATTING_RULES['space_period_para']
    return check_body_rule(doc, 'space_period_para', rule['pattern'], rule['message'])

def find_double_punctuations(doc, doc_path: str = None) -> List[Dict[str, Any]]:
    """Flags adjacent duplicate punctuation marks (e.g. ,,, ::, ;;, ..)."""
    rule = FORMATTING_RULES['double_punctuation']
    findings = check_body_rule(doc, 'double_punctuation', rule['pattern'], rule['message'])
    if doc_path:
        findings.extend(check_endnotes_rule(doc_path, 'double_punctuation', rule['pattern'], rule['message']))
    return findings

def find_space_before_endnote(doc) -> List[Dict[str, Any]]:
    """Flags spaces immediately preceding an endnote reference."""
    rule = FORMATTING_RULES['space_before_endnote']
    return check_body_rule(doc, 'space_before_endnote', rule['pattern'], rule['message'])

def find_no_space_endnote_text(doc) -> List[Dict[str, Any]]:
    """Flags cases where an endnote reference is immediately followed by a letter without spacing."""
    rule = FORMATTING_RULES['no_space_endnote_text']
    return check_body_rule(doc, 'no_space_endnote_text', rule['pattern'], rule['message'])

def find_endnote_colon(doc) -> List[Dict[str, Any]]:
    """Flags cases where an endnote reference is followed by a colon."""
    rule = FORMATTING_RULES['endnote_colon']
    return check_body_rule(doc, 'endnote_colon', rule['pattern'], rule['message'])

def find_endnote_semicolon(doc) -> List[Dict[str, Any]]:
    """Flags cases where an endnote reference is followed by a semicolon."""
    rule = FORMATTING_RULES['endnote_semicolon']
    return check_body_rule(doc, 'endnote_semicolon', rule['pattern'], rule['message'])

def find_endnote_comma(doc) -> List[Dict[str, Any]]:
    """Flags cases where an endnote reference is followed by a comma."""
    rule = FORMATTING_RULES['endnote_comma']
    return check_body_rule(doc, 'endnote_comma', rule['pattern'], rule['message'])

def find_comma_after_quotes_in_endnotes(doc_path: str) -> List[Dict[str, Any]]:
    """Checks for comma placed outside of double quotes in endnotes."""
    rule = FORMATTING_RULES['comma_after_quotes_in_endnotes']
    return check_endnotes_rule(doc_path, 'comma_after_quotes_in_endnotes', rule['pattern'], rule['message'])

def find_hyphen_spaced(doc, doc_path: str = None) -> List[Dict[str, Any]]:
    """Flags hyphens surrounded by spaces."""
    rule = FORMATTING_RULES['hyphen_spaced']
    findings = check_body_rule(doc, 'hyphen_spaced', rule['pattern'], rule['message'])
    if doc_path:
        findings.extend(check_endnotes_rule(doc_path, 'hyphen_spaced', rule['pattern'], rule['message']))
    return findings

def find_spaced_endash(doc, doc_path: str = None) -> List[Dict[str, Any]]:
    """Flags double hyphens or spaced en-dashes/em-dashes used for parenthetical dashes."""
    rule = FORMATTING_RULES['spaced_endash']
    findings = check_body_rule(doc, 'spaced_endash', rule['pattern'], rule['message'])
    if doc_path:
        findings.extend(check_endnotes_rule(doc_path, 'spaced_endash', rule['pattern'], rule['message']))
    return findings

def process_formatting(input_path: str) -> Dict[str, Any]:
    """Process all formatting checks on the docx file."""
    doc = Document(input_path)
    
    findings_by_rule = {}
    for rule_id, rule in FORMATTING_RULES.items():
        rule_findings = []
        target = rule.get('target', 'both')
        pattern = rule['pattern']
        message = rule['message']
        
        if target in ('body', 'both'):
            rule_findings.extend(check_body_rule(doc, rule_id, pattern, message))
        if target in ('endnotes', 'both'):
            rule_findings.extend(check_endnotes_rule(input_path, rule_id, pattern, message))
            
        findings_by_rule[rule_id] = rule_findings
    
    # Flatten findings and count rule totals
    all_findings = []
    rule_totals = {}
    for rule_id, rule_findings in findings_by_rule.items():
        all_findings.extend(rule_findings)
        rule_totals[rule_id] = len(rule_findings)
        
    return {
        'total': len(all_findings),
        'rule_totals': rule_totals,
        'findings_by_rule': findings_by_rule,
        'findings': all_findings
    }

def find_serial_commas(doc) -> Dict[str, Any]:
    """Find serial commas (Oxford commas) in the document"""
    try:
        serial_commas = []
        serial_comma_pattern = FORMATTING_RULES.get('serial_commas', {}).get('pattern', r'\b[^,.]+,\s+[^,.]+,\s+(?:and|or)\s+[^,.]+\b')
        
        for para_index, paragraph in enumerate(doc.paragraphs, 1):
            matches = re.finditer(serial_comma_pattern, paragraph.text)
            for match in matches:
                serial_commas.append({
                    'text': match.group(),
                    'location': f'Paragraph {para_index}'
                })
        
        for table_index, table in enumerate(doc.tables, 1):
            for row_index, row in enumerate(table.rows, 1):
                for cell_index, cell in enumerate(row.cells, 1):
                    matches = re.finditer(serial_comma_pattern, cell.text)
                    for match in matches:
                        serial_commas.append({
                            'text': match.group(),
                            'location': f'Table {table_index}, Row {row_index}, Cell {cell_index}'
                        })
        return {"serial_commas": serial_commas}
    except Exception as e:
        raise Exception(f"Error finding serial commas: {str(e)}")

def find_double_spaces(doc) -> Dict[str, Any]:
    """Find and count double spaces (two consecutive spaces) in the document"""
    try:
        double_spaces = []
        double_space_pattern = FORMATTING_RULES.get('double_spaces', {}).get('pattern', r'  +')
        
        for para_index, paragraph in enumerate(doc.paragraphs, 1):
            matches = list(re.finditer(double_space_pattern, paragraph.text))
            for match in matches:
                start = max(0, match.start() - 15)
                end = min(len(paragraph.text), match.end() + 15)
                snippet = paragraph.text[start:end]
                snippet_highlighted = snippet[:match.start() - start] + f"[{match.group(0)}]" + snippet[match.end() - start:]
                double_spaces.append({
                    'text': snippet_highlighted.strip(),
                    'location': f'Paragraph {para_index}'
                })
        
        for table_index, table in enumerate(doc.tables, 1):
            for row_index, row in enumerate(table.rows, 1):
                for cell_index, cell in enumerate(row.cells, 1):
                    matches = list(re.finditer(double_space_pattern, cell.text))
                    for match in matches:
                        start = max(0, match.start() - 15)
                        end = min(len(cell.text), match.end() + 15)
                        snippet = cell.text[start:end]
                        snippet_highlighted = snippet[:match.start() - start] + f"[{match.group(0)}]" + snippet[match.end() - start:]
                        double_spaces.append({
                            'text': snippet_highlighted.strip(),
                            'location': f'Table {table_index}, Row {row_index}, Cell {cell_index}'
                        })
        return {
            'total': len(double_spaces),
            'double_spaces': double_spaces
        }
    except Exception as e:
        raise Exception(f"Error finding double spaces: {str(e)}")

def filter_searched_expansion(expansion_array: list[str], abbreviation_list: list[str], document_full_content: str) -> None:
    try:
        # 1. Filter based on casing
        for i in range(1, len(expansion_array)):
            if expansion_array[i]:
                split_expansion = expansion_array[i].split('|')
                filtered_text = ""
                for item in split_expansion[1:]:
                    temp_string = re.sub(r"[a-z\W]", "", item)
                    if temp_string == abbreviation_list[i]:
                        filtered_text += "|" + item
                if filtered_text:
                    expansion_array[i] = filtered_text

        # 2. Filter based on same paragraph presence
        split_para = document_full_content.split('\n')
        for i in range(1, len(expansion_array)):
            if expansion_array[i]:
                split_expansion = expansion_array[i].split('|')
                if len(split_expansion) > 2:
                    filtered_text = ""
                    for item in split_expansion[1:]:
                        for para in split_para:
                            if (re.search(r"\W" + re.escape(item) + r"\W", para, re.IGNORECASE) and
                                re.search(r"\W" + re.escape(abbreviation_list[i]) + r"\W", para)):
                                filtered_text += "|" + item
                                break  # Stop searching after the first matching paragraph
                    expansion_array[i] = filtered_text if filtered_text else ""
    except Exception as e:
        print(f"Error resolving abbreviation expansions in filter_searched_expansion def: {e}")

def new_approach_expansion_as_is_form(previous_text, abbreviation, preposition_list):
    try:
        filtered_text = ''
        
        # Build the pattern for expansion
        pattern = build_single_word_regex_pattern(abbreviation, preposition_list)

        # Try pattern like: expansion text + abbreviation (e.g., World Health Organization (WHO))
        pattern_temp = rf"{pattern}(\W)+{re.escape(abbreviation)}(s)?(\W)+"
        match = re.search(pattern_temp, previous_text, re.IGNORECASE)
        
        if match:
            filtered_text = match.group(0)
            filtered_text = re.sub(rf"\({re.escape(abbreviation)}\)", "", filtered_text, flags=re.IGNORECASE)
        else:
            # Try pattern like: abbreviation + expansion (e.g., WHO (World Health Organization))
            pattern_temp = rf"{re.escape(abbreviation)}(\W)*{pattern}(\W)+"
            match = re.search(pattern_temp, previous_text, re.IGNORECASE)
            if match:
                filtered_text = match.group(0)
                filtered_text = re.sub(re.escape(abbreviation), "", filtered_text, flags=re.IGNORECASE)
        
        if not filtered_text:
            # Try pattern like: expansion , abbreviation
            pattern_temp = rf"{pattern}(\W)*,{re.escape(abbreviation)}"
            match = re.search(pattern_temp, previous_text, re.IGNORECASE)
            if match:
                filtered_text = match.group(0)
        if not filtered_text:
            pattern_temp = rf"{pattern}(\W)+({preposition_list})?(\s)?[A-Z][a-z\-\'\’]+(\s)?(\W){re.escape(abbreviation)}(\s)?(\W)+"
            match = re.search(pattern_temp, previous_text, re.IGNORECASE)
            if match:
                filtered_text = match.group(0)
                #filtered_text = re.sub(rf"(\W)?{re.escape(abbreviation)}(s)?(\W)?", "", filtered_text, flags=re.IGNORECASE)
                filtered_text = check_abbreviation_match(abbreviation, filtered_text)

        # Clean abbreviation from result
        if filtered_text:
            filtered_text = re.sub(rf"(\W)?{re.escape(abbreviation)}(s)?(\W)?", "", filtered_text, flags=re.IGNORECASE)

        return filtered_text.strip()
    except Exception as e:
        print(f"Error resolving abbreviation expansions in new_approach_expansion_as_is_form def: {e}")

def check_abbreviation_match(abbreviation, filtered_text):
    try:
        # List of words to ignore when checking first letters
        ignore_words = {'at', 'by', 'for', 'of', 'in', 'on', 'to', 'and', 'the'}
        
        # Remove the abbreviation in parentheses if present
        filtered_text = re.sub(rf"\({re.escape(abbreviation)}\)", "", filtered_text, flags=re.IGNORECASE)
        filtered_text = re.sub(r'[^a-zA-Z\s]', ' ', filtered_text)
        filtered_text = re.sub(rf"\b{re.escape(abbreviation)}\b", "", filtered_text, flags=re.IGNORECASE).strip()
        if re.search(rf'\b{re.escape(abbreviation)}\b', filtered_text, flags=re.IGNORECASE):
            return ""
        # Split the filtered text into words
        words = filtered_text.split()
        
        # Filter out ignored words and get first letters
        first_letters = []
        for word in words:
            lower_word = word.lower()
            if lower_word not in ignore_words:
                if word:  # check if word is not empty
                    first_letters.append(word[0].lower())
        
        # Check if each character in abbreviation matches the first letters
        if len(abbreviation) > len(first_letters):
            return ""
        
        for i in range(len(abbreviation)):
            if i >= len(first_letters):
                return ""
            if abbreviation[i].lower() != first_letters[i]:
                return ""
        
        return filtered_text
    except Exception as e:
        print(f"Error in check_abbreviation_match def: {e}")

def build_single_word_regex_pattern(abbreviation: str, preposition_list: str) -> str:
    try:
        abbr_chars = list(abbreviation)
        pattern = (
            f"[{abbr_chars[0].upper()}{abbr_chars[0].lower()}][a-zA-Z\\-]+(\\s)?"
            f"({preposition_list})?(\\s)?"
        )

        for i in range(1, len(abbr_chars)):
            ch = abbr_chars[i]

            if i < len(abbr_chars) - 1:
                if re.match(r"[0-9]", ch):
                    pattern += (
                        f"([{ch.upper()}][a-z\\-\\'\\’]*)(\\s)?"
                        f"({preposition_list})?(\\s)?"
                    )
                elif re.match(r"[- ]", ch):
                    pattern += (
                        f"([{ch.upper()}][a-z\\-\\'\\’]*)?(\\s)?"
                        f"({preposition_list})?(\\s)?"
                    )
                else:
                    pattern += (
                        f"[{ch.upper()}{ch.lower()}][a-z\\-\\'\\’]+(\\s)?"
                        f"({preposition_list})?(\\s)?"
                    )
            else:
                if re.match(r"[- 0-9]", ch):
                    pattern += f"([{ch.upper()}][a-z\\-\\'\\’]*)"
                else:
                    pattern += f"[{ch.upper()}{ch.lower()}][a-z\\-\\'\\’]+"

        pattern = "(\\W)" + pattern
        return pattern
    except Exception as e:
        print(f"Error resolving abbreviation expansions in build_single_word_regex_pattern def: {e}")

def process_analysis(input_path: str, us_dict: str, uk_dict: str):
    try:
        """Process document analysis for UK/US words"""
        us_words = load_dictionary(us_dict)
        uk_words = load_dictionary(uk_dict)
        results = analyze_docx(input_path, us_words, uk_words)
        report_content = generate_report_content(results)
        return report_content    
    except Exception as e:
            print(f"Error in process_analysis def: {e}")

def process_abbreviations(input_path: str):
    try:
        """Process document for abbreviations and serial commas"""
        text = extract_text_from_docx(input_path)
        #doc = Document(input_path)
        
        abbreviations = extract_abbreviation_list(text)
        
        
        result = abbreviations
        #result.update(serial_commas)
        return result
    except Exception as e:
        print(f"Error in process_abbreviations def: {e}")

def generate_combined_report(results: dict) -> str:
    try:
        """Generate a comprehensive text report from combined results"""
        report = []
        
        # Word analysis section
        word_data = results['word_analysis']
        report.append("=== COMBINED DOCUMENT ANALYSIS REPORT ===")
        report.append("\n=== WORD ANALYSIS ===")
        report.append(f"Total Words: {word_data['total_words_in_document']}")
        
        # US Words
        report.append("\nUS Words:")
        report.append(f"Total: {word_data['us']['total']}")
        for word, count in word_data['us']['words'].items():
            report.append(f"{word}: {count}")
        
        # UK Words
        report.append("\nUK Words:")
        report.append(f"Total: {word_data['uk']['total']}")
        for word, count in word_data['uk']['words'].items():
            report.append(f"{word}: {count}")
        
        # Abbreviations section
        abbrev_data = results['abbreviations'].get('abbreviations_found', [])
        report.append("\n=== ABBREVIATIONS ===")
        report.append(f"Total Abbreviations Found: {len(abbrev_data)}")
        for item in abbrev_data:
            report.append(f"\nAbbreviation: {item.get('abbreviation', '')}")
            report.append(f"Full Form: {item.get('full_form', '')}")
            report.append(f"Occurrences: {item.get('occurrences', 0)}")
        
        # Serial commas section
        serial_data = results['serial_commas'].get('serial_commas', [])
        report.append("\n=== SERIAL COMMAS ===")
        report.append(f"Total Serial Commas Found: {len(serial_data)}")
        for item in serial_data:
            report.append(f"\nText: {item.get('text', '')}")
            report.append(f"Location: {item.get('location', '')}")
        
        # Double spaces section
        double_space_data = results.get('double_spaces', {})
        double_spaces = double_space_data.get('double_spaces', [])
        report.append("\n=== DOUBLE SPACES ===")
        report.append(f"Total Double Spaces Found: {double_space_data.get('total', 0)}")
        for item in double_spaces:
            report.append(f"\nText: {item.get('text', '')}")
            report.append(f"Location: {item.get('location', '')}")
        
        # Formatting section
        formatting_data = results.get('formatting', {})
        report.append("\n=== FORMATTING ISSUES ===")
        for rule_id, rule_info in FORMATTING_RULES.items():
            rule_findings = formatting_data.get('findings_by_rule', {}).get(rule_id, [])
            report.append(f"\nTotal {rule_info['message']} Found: {len(rule_findings)}")
            for item in rule_findings:
                report.append(f"  Text: {item.get('text', '')}")
                report.append(f"  Location: {item.get('location', '')}")
            
        return '\n'.join(report)
    except Exception as e:
        print(f"Error in generate_combined_report def: {e}")

def count_hyphenated_word_variants(document_content: str) -> Dict[str, Any]:
    """
    Counts variants of hyphenated words in a DOCX file, including words with multiple hyphens.
    
    Args:
        document_content: The text content to analyze
        
    Returns:
        Dictionary with structure: {
            "base_word": {
                "hyphenated": count, 
                "space": count, 
                "combined": count,
                "total": count
            }
        }
    """
     # First find all unique hyphenated words (including multi-hyphen)
    hyphen_words = set(re.findall(r'\b([a-zA-Z]+(?:-[a-zA-Z]+)+)\b', document_content, re.IGNORECASE))
    
    word_counts = defaultdict(lambda: {"hyphenated": 0, "space": 0, "combined": 0, "total": 0})
    
    for word in hyphen_words:
        base_word = word.lower()
        
        # Count exact hyphenated matches (case insensitive)
        hyphenated_count = len(re.findall(
            rf'(?<!\w){re.escape(word)}(?!\w)',  # Negative lookarounds for strict boundaries
            document_content,
            re.IGNORECASE
        ))
        
        # Generate and count space version
        space_version = word.replace('-', ' ')
        space_count = len(re.findall(
            rf'(?<!\w){re.escape(space_version)}(?!\w)',
            document_content,
            re.IGNORECASE
        )) if ' ' in space_version else 0
        
        # Generate and count combined version
        combined_version = word.replace('-', '')
        combined_count = len(re.findall(
            rf'(?<!\w){re.escape(combined_version)}(?!\w)',
            document_content,
            re.IGNORECASE
        ))
        
        word_counts[base_word] = {
            "hyphenated": hyphenated_count,
            "space": space_count,
            "closed-up": combined_count,
            "total": hyphenated_count + space_count + combined_count
        }
    
    return dict(word_counts)

def highlight_findings(text: str) -> str:
    """Helper to convert bracketed matches (e.g. [,,] or [   ]) to styled HTML marks."""
    def replacer(m):
        content = m.group(1)
        if content.isspace():
            html_content = "&nbsp;" * len(content)
        else:
            html_content = content.replace('[[^e]]', '<sup style="color: #c0392b; font-weight: bold;">[Endnote]</sup>')
            html_content = html_content.replace('[[^f]]', '<sup style="color: #c0392b; font-weight: bold;">[Footnote]</sup>')
        return f'<mark style="background-color: #f39c12; color: white; padding: 0 2px;">{html_content}</mark>'
        
    return re.sub(
        r'\[((?:[^\[\]]|\[\[\^e\]\]|\[\[\^f\]\])+)\]',
        replacer,
        text
    )

def format_html_text(text: str) -> str:
    """Highlight matches and convert remaining markers to clean HTML superscript tags."""
    text = highlight_findings(text)
    text = text.replace('[[^e]]', '<sup style="color: #7f8c8d;">[Endnote]</sup>')
    text = text.replace('[[^f]]', '<sup style="color: #7f8c8d;">[Footnote]</sup>')
    return text

# Example usage:
# save_as_html(combined_results, "analysis_report.html")
def save_as_html(combined_results, output_file):
    output_file = os.path.splitext(output_file)[0] + '.html'
    """
    Save the analysis results as an HTML file.
    
    Args:
        combined_results (dict): Dictionary containing word analysis, abbreviations, and serial commas data
        output_file (str): Path to the output HTML file
    """
    formatting_data = combined_results.get('formatting', {})
    rule_totals = formatting_data.get('rule_totals', {})
    findings_by_rule = formatting_data.get('findings_by_rule', {})
    
    # Build summary box items for formatting rules
    summary_items = []
    for rule_id, rule_info in FORMATTING_RULES.items():
        count = rule_totals.get(rule_id, 0)
        summary_items.append(f'<div class="summary-item">Total {rule_info["message"]}: {count}</div>')
    formatting_summary_html = "\n".join(summary_items)
    
    # Build detailed sections for formatting rules
    detailed_sections = []
    for rule_id, rule_info in FORMATTING_RULES.items():
        rule_findings = findings_by_rule.get(rule_id, [])
        if rule_findings:
            cards = []
            for item in rule_findings:
                cards.append(f"""
                <div class="serial-comma" style="background-color: #fcf3cf; border-left-color: #f1c40f;">
                    <p>{format_html_text(item['text'])}</p>
                    <p class="location">{item['location']}</p>
                </div>
                """)
            cards_html = "\n".join(cards)
            detailed_sections.append(f"""
            <div class="section">
                <h2>{rule_info['message']} Detected</h2>
                {cards_html}
            </div>
            """)
    formatting_details_html = "\n".join(detailed_sections)

    html_content = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Document Analysis Results</title>
        <style>
            body {{ font-family: Arial, sans-serif; line-height: 1.6; margin: 20px; }}
            h1 {{ color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 5px; }}
            h2 {{ color: #2980b9; margin-top: 25px; }}
            table {{ border-collapse: collapse; width: 100%; margin-bottom: 20px; }}
            th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
            th {{ background-color: #f2f2f2; }}
            tr:nth-child(even) {{ background-color: #f9f9f9; }}
            .section {{ margin-bottom: 30px; }}
            .serial-comma {{ background-color: #e8f4f8; padding: 10px; margin: 5px 0; border-left: 4px solid #3498db; }}
            .location {{ font-style: italic; color: #7f8c8d; }}
            .summary-box {{
                background-color: #e6f2ff;
                border: 1px solid #99c2ff;
                border-radius: 5px;
                padding: 15px;
                margin-bottom: 20px;
            }}
            .summary-item {{
                margin-bottom: 5px;
                font-weight: bold;
            }}
        </style>
    </head>
    <body>
        <h1>Document Analysis Results</h1>
        
        <div class="section">
            <h2>Word Analysis</h2>
            
            <div class="summary-box">
                <div class="summary-item">Total words in document: {combined_results['word_analysis']['total_words_in_document']:,}</div>
                <div class="summary-item">Total US spelling variants: {combined_results['word_analysis']['us']['total']}</div>
                <div class="summary-item">Total UK spelling variants: {combined_results['word_analysis']['uk']['total']}</div>
                {formatting_summary_html}
            </div>
            
            <h3>US Spelling Variants</h3>
            <table>
                <tr>
                    <th>Word</th>
                    <th>Count</th>
                </tr>
                {"".join(
                    f"<tr><td>{word}</td><td>{count}</td></tr>"
                    for word, count in combined_results['word_analysis']['us']['words'].items()
                )}
            </table>
            
            <h3>UK Spelling Variants</h3>
            <table>
                <tr>
                    <th>Word</th>
                    <th>Count</th>
                </tr>
                {"".join(
                    f"<tr><td>{word}</td><td>{count}</td></tr>"
                    for word, count in combined_results['word_analysis']['uk']['words'].items()
                )}
            </table>
        </div>
        <div class="section">
            <h2>Hyphenated Words Analysis</h2>
            <div class="summary-box">
                <div class="summary-item">Total hyphenated word forms found: {sum(v['total'] for v in combined_results['hyphenated_words'].values()):,}</div>
                <div class="summary-item">Unique hyphenated word bases: {len(combined_results['hyphenated_words'])}</div>
            </div>
            
            <table>
                <tr>
                    <th>Word-Searched</th>
                    <th>Hyphenated</th>
                    <th>Spaced</th>
                    <th>Closed-up</th>
                    <th>Total</th>
                </tr>
                {"".join(
                    f"""
                    <tr>
                        <td class="hyphenated-word">{base_word}</td>
                        <td><span class="variant-count">{stats['hyphenated']}</span></td>
                        <td><span class="variant-count">{stats['space']}</span></td>
                        <td><span class="variant-count">{stats['closed-up']}</span></td>
                        <td>{stats['total']}</td>
                    </tr>
                    """
                    for base_word, stats in combined_results['hyphenated_words'].items()
                )}
            </table>
        </div>
        <div class="section">
            <h2>Abbreviations Found</h2>
            <table>
                <tr>
                    <th>Abbreviation</th>
                    <th>Full Form</th>
                    <th>Occurrences</th>
                </tr>
                {"".join(
                    f"<tr><td>{abbr['abbreviation']}</td><td>{abbr['full_form'] if abbr['full_form'] else 'Not specified'}</td><td>{abbr['occurrences']}</td></tr>"
                    for abbr in combined_results['abbreviations']['abbreviations_found']
                )}
            </table>
        </div>
        
        {formatting_details_html}
    </body>
    </html>
    """

    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(html_content)
        
    print(f"Results saved successfully to {output_file}")

# Example usage:
# save_as_html(combined_results, "analysis_report.html")        
def save_analysis_as_html(combined_results, output_file):
    # Replace the .json extension with .html in the output_file path
    output_file = os.path.splitext(output_file)[0] + '.html'

    word_results = combined_results.get('word_analysis', {})
    abbreviation_results = combined_results.get('abbreviations', {})
    serial_commas = combined_results.get('serial_commas', {})
    double_spaces = combined_results.get('double_spaces', {})

    # Extract US, UK, and total word counts
    us_total = word_results.get('us', {}).get('total', 0)
    uk_total = word_results.get('uk', {}).get('total', 0)
    total_words_in_document = word_results.get('total_words_in_document', 0)

    html = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <title>Document Analysis Report</title>
        <style>
            body {{ font-family: Arial, sans-serif; padding: 20px; }}
            h2 {{ color: #2a4d69; }}
            table {{ border-collapse: collapse; width: 100%; margin-bottom: 20px; }}
            th, td {{ border: 1px solid #ccc; padding: 8px; text-align: left; }}
            th {{ background-color: #f2f2f2; }}
        </style>
    </head>
    <body>
        <h1>Document Analysis Report</h1>

        <h2>Word Analysis</h2>
        <p><strong>Total Words in Document:</strong> {total_words_in_document}</p>

        <h3>US Words</h3>
        <p><strong>Total US Words:</strong> {us_total}</p>
        <table>
            <tr><th>Word</th><th>Occurrences</th></tr>
    """
    for word, count in word_results.get('us', {}).get('words', {}).items():
        html += f"<tr><td>{word}</td><td>{count}</td></tr>"

    html += f"""
        </table>

        <h3>UK Words</h3>
        <p><strong>Total UK Words:</strong> {uk_total}</p>
        <table>
            <tr><th>Word</th><th>Occurrences</th></tr>
    """
    for word, count in word_results.get('uk', {}).get('words', {}).items():
        html += f"<tr><td>{word}</td><td>{count}</td></tr>"

    html += """
        </table>

        <h2>Abbreviations Found</h2>
        <table>
            <tr><th>Abbreviation</th><th>Full Form</th><th>Occurrences</th></tr>
    """
    for abbr in abbreviation_results.get('abbreviations_found', []):
        html += f"<tr><td>{abbr['abbreviation']}</td><td>{abbr['full_form'] or '-'}</td><td>{abbr['occurrences']}</td></tr>"

    html += """
        </table>
    """

    formatting_data = combined_results.get('formatting', {})
    findings_by_rule = formatting_data.get('findings_by_rule', {})
    
    for rule_id, rule_info in FORMATTING_RULES.items():
        rule_findings = findings_by_rule.get(rule_id, [])
        if rule_findings:
            html += f"\n        <h2>{rule_info['message']} Detected</h2>\n"
            html += "        <table>\n"
            html += "            <tr><th>Text</th><th>Location</th></tr>\n"
            for item in rule_findings:
                html += f"            <tr><td>{format_html_text(item['text'])}</td><td>{item['location']}</td></tr>\n"
            html += "        </table>\n"

    html += """
    </body>
    </html>
    """

    # Save the generated HTML to the file
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(html)
    
    print(f"HTML report saved as '{output_file}'")

def main():
    parser = argparse.ArgumentParser(description='Document Analyzer CLI Tool')
    subparsers = parser.add_subparsers(dest='command', required=True)
    
    # Analysis command
    analysis_parser = subparsers.add_parser('analyze', help='Analyze document for UK/US words')
    analysis_parser.add_argument('-i', '--input', required=True, help='Input DOCX file path')
    analysis_parser.add_argument('-o', '--output', required=True, help='Output file path')
    analysis_parser.add_argument('--us_dict', default='us_dict.txt', help='Path to US dictionary')
    analysis_parser.add_argument('--uk_dict', default='uk_dict.txt', help='Path to UK dictionary')
    
    # Abbreviations command
    abbrev_parser = subparsers.add_parser('abbreviations', help='Find abbreviations and serial commas')
    abbrev_parser.add_argument('-i', '--input', required=True, help='Input DOCX file path')
    abbrev_parser.add_argument('-o', '--output', required=True, help='Output JSON file path')
    
    # Formatting command
    format_parser = subparsers.add_parser('formatting', help='Run formatting/styling checks')
    format_parser.add_argument('-i', '--input', required=True, help='Input DOCX file path')
    format_parser.add_argument('-o', '--output', required=True, help='Output JSON file path')
    
    # Combined analysis command
    combined_parser = subparsers.add_parser('analyze-all', help='Run all analyses (words and abbreviations)')
    combined_parser.add_argument('-i', '--input', required=True, help='Input DOCX file path')
    combined_parser.add_argument('-o', '--output', required=True, help='Output file path')
    combined_parser.add_argument('--us_dict', default='us_dict.txt', help='Path to US dictionary')
    combined_parser.add_argument('--uk_dict', default='uk_dict.txt', help='Path to UK dictionary')

    # Open Section Words Form command
    subparsers.add_parser('open-form', help='Open Section Words Management form in default web browser')

    args = parser.parse_args()

    if args.command == 'open-form':
        if getattr(sys, 'frozen', False):
            # Check MEIPASS (bundled inside exe) or next to executable
            meipass_dir = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
            exe_dir = os.path.dirname(sys.executable)
            
            form_path = os.path.join(exe_dir, 'section_words_form.html')
            meipass_form = os.path.join(meipass_dir, 'section_words_form.html')
            
            if os.path.exists(meipass_form):
                try:
                    import shutil
                    shutil.copy2(meipass_form, form_path)
                    print(f"Extracted section_words_form.html to: {form_path}")
                except Exception:
                    form_path = meipass_form
            elif not os.path.exists(form_path):
                form_path = meipass_form
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            form_path = os.path.join(base_dir, 'section_words_form.html')
            
        if os.path.exists(form_path):
            import webbrowser
            webbrowser.open(f"file:///{os.path.abspath(form_path)}")
            print(f"Opened Section Words Form in default web browser: {form_path}")
        else:
            print(f"Error: section_words_form.html not found.")
        return

    try:
        # Common validations
        if not os.path.exists(args.input):
            print(f"Error: Input file not found: {args.input}")
            sys.exit(1)
            
        if not args.input.endswith('.docx'):
            print("Error: Only DOCX files are supported")
            sys.exit(1)
        
        if args.command == 'analyze':
            # Validate dictionaries
            if not os.path.exists(args.us_dict):
                print(f"Error: US dictionary file not found: {args.us_dict}")
                sys.exit(1)
            if not os.path.exists(args.uk_dict):
                print(f"Error: UK dictionary file not found: {args.uk_dict}")
                sys.exit(1)
            
            # Process word analysis
            us_words = load_dictionary(args.us_dict)
            uk_words = load_dictionary(args.uk_dict)
            word_results = analyze_docx(args.input, us_words, uk_words)
            report_content = generate_report_content(word_results)
            
            # Save results
            with open(args.output, 'w', encoding='utf-8') as f:
                f.write(report_content)
            print(f"Analysis complete. Report saved to: {args.output}")
            
        elif args.command == 'abbreviations':
            # Process abbreviations and serial commas
            text = extract_text_from_docx(args.input)
            doc = Document(args.input)
            
            abbreviation_results = extract_abbreviation_list(text)
            serial_commas = find_serial_commas(doc)
            
            # Combine results
            combined_results = {
                'abbreviations': abbreviation_results,
                'serial_commas': serial_commas
            }
            
            # Save as JSON
            with open(args.output, 'w', encoding='utf-8') as f:
                json.dump(combined_results, f, indent=4)
            print(f"Abbreviation analysis complete. Results saved to: {args.output}")
            
        elif args.command == 'formatting':
            formatting_results = process_formatting(args.input)
            with open(args.output, 'w', encoding='utf-8') as f:
                json.dump(formatting_results, f, indent=4)
            print(f"Formatting analysis complete. Results saved to: {args.output}")
            
        elif args.command == 'analyze-all':
            # Validate dictionaries
            if not os.path.exists(args.us_dict):
                print(f"Error: US dictionary file not found: {args.us_dict}")
                sys.exit(1)
            if not os.path.exists(args.uk_dict):
                print(f"Error: UK dictionary file not found: {args.uk_dict}")
                sys.exit(1)
            
            # Process all analyses
            us_words = load_dictionary(args.us_dict)
            uk_words = load_dictionary(args.uk_dict)
            text = extract_text_from_docx(args.input)
            doc = Document(args.input)
            
            word_results = analyze_docx(args.input, us_words, uk_words)
            abbreviation_results = extract_abbreviation_list(text)
            serial_commas = find_serial_commas(doc)
            hyphenated_words = count_hyphenated_word_variants(text)
            double_spaces = find_double_spaces(doc)
            formatting = process_formatting(args.input)
            
            # Combine all results
            combined_results = {
                'word_analysis': word_results,
                'abbreviations': abbreviation_results,
                'serial_commas': serial_commas,
                'hyphenated_words': hyphenated_words,
                'double_spaces': double_spaces,
                'formatting': formatting
            }
            
            # Determine output format
            if args.output.lower().endswith('.json'):
                with open(args.output, 'w', encoding='utf-8') as f:
                    json.dump(combined_results, f, indent=4)
            else:
                # Generate text report
                report = generate_combined_report(combined_results)
                with open(args.output, 'w', encoding='utf-8') as f:
                    f.write(report)
            #save_analysis_as_html(combined_results, args.output)
            save_as_html(combined_results, args.output)
            # Save as HTML
            print(f"Combined analysis complete. Results saved to: {args.output}")
            
    except Exception as e:
        print(f"Error: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    # If executed without CLI arguments (e.g., via the editor's Run/Debug button), inject default arguments
    if len(sys.argv) == 1:
        sys.argv.extend(["analyze-all", "-i", "input.docx", "-o", "input_report.json"])
    main()