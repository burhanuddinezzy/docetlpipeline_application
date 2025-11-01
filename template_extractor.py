import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox
import json
import os
import re
from typing import List, Dict, Tuple, Optional
import difflib
from dataclasses import dataclass
import io

@dataclass
class TemplateMatch:
    template_name: str
    confidence: float
    template_data: dict

@dataclass
class TextElement:
    """Represents a text element with its coordinates and metadata."""
    text: str
    x0: float
    y0: float
    x1: float
    y1: float
    center_x: float
    center_y: float
    box_assignment: Optional[str] = None  # Which box this text belongs to, if any

class TextNormalizer:
    """Handles text normalization and fuzzy matching for BOL template matching."""
    
    @staticmethod
    def normalize_text(text: str) -> str:
        """Normalize text for consistent matching."""
        if not text:
            return ""
        
        # Convert to uppercase for case-insensitive matching
        normalized = text.upper()
        
        # Normalize whitespace (spaces, newlines, tabs -> single space)
        normalized = re.sub(r'\s+', ' ', normalized)
        
        # Remove common OCR artifacts and special characters
        normalized = re.sub(r'[^\w\s\-\.\,\(\)\/]', '', normalized)
        
        # Strip leading/trailing whitespace
        return normalized.strip()
        
    @staticmethod
    def calculate_fuzzy_similarity(template_phrase: str, pdf_text: str) -> float:
        """Calculate similarity between template phrase and PDF text using multiple strategies."""
        if not template_phrase or not pdf_text:
            return 0.0
        
        # Normalize both texts
        norm_template = TextNormalizer.normalize_text(template_phrase)
        norm_pdf = TextNormalizer.normalize_text(pdf_text)
                
        # Strategy 4: Fuzzy string matching (difflib)
        fuzzy_score = difflib.SequenceMatcher(None, norm_template, norm_pdf).ratio()
                
        return fuzzy_score

class BOLTemplateExtractor:
    def __init__(self, templates_dir: str = "templates", confidence_threshold: float = 0.5):
        """
        Initialize BOL template-based extractor with fuzzy matching.
        
        Args:
            templates_dir: Directory containing template JSON files
            confidence_threshold: Minimum confidence to use a template (0.0-1.0)
                                 Lowered default to 0.5 for fuzzy matching
        """
        self.templates_dir = templates_dir
        self.confidence_threshold = confidence_threshold
        self.templates = {}
        
        # Create templates directory if it doesn't exist
        os.makedirs(templates_dir, exist_ok=True)
        
        # Load all templates
        self.load_templates()
    
    def load_templates(self):
        """Load all template files from the templates directory."""
        self.templates = {}

        template_files = [f for f in os.listdir(self.templates_dir) if f.endswith('.json')]

        for template_file in template_files:
            file_path = os.path.join(self.templates_dir, template_file)
            with open(file_path, 'r', encoding='utf-8') as f:
                template_data = json.load(f)

            template_name = template_data.get('template_name', template_file.replace('.json', ''))

            self.templates[template_name] = template_data
        
    def _extract_with_template(self, pdf_path: str, template_data: dict) -> str:
        """
        Extract text from PDF using the provided template data.
        """
        try:
            doc = fitz.open(pdf_path)
            template_name = template_data.get('template_name', 'Unknown')
            total_doc_pages = len(doc)
            
            # --- UNIFIED TEMPLATE HANDLING ---
            # Create a standardized 'pages' structure regardless of input template type
            unified_pages = {}
            
            if 'pages' in template_data and isinstance(template_data['pages'], dict):
                # It's a true multi-page template
                unified_pages = template_data['pages']
                total_template_pages = len(unified_pages)
            else:
                # It's a legacy single-page template
                # Wrap its data into a pages structure for page 1
                unified_pages = {
                    "1": {
                        "page_raw_text": template_data.get('template_raw_text', ''),
                        "boxes": template_data.get('boxes', [])
                    }
                }
                total_template_pages = 1
            
            # Create header for the entire document
            output_parts = [
                f"## BOL Extraction Results", 
                f"**Template Used:** {template_name}",
                f"**Document Pages:** {total_doc_pages}",
                f"**Template Pages:** {total_template_pages}\n"
            ]
            
            # Process each page in the DOCUMENT
            for page_num in range(1, total_doc_pages + 1):
                page = doc[page_num - 1]  # 0-indexed for fitz
                
                # Add page separator
                output_parts.append(f"\n---\n**Page {page_num}**\n---\n")
                
                # Get template data for this page.
                # If template has fewer pages than document, use the last available template page.
                # If template has more pages, ignore extra template pages.
                template_page_key = str(min(page_num, total_template_pages))
                template_page_data = unified_pages.get(template_page_key, {})
                boxes = sorted(template_page_data.get('boxes', []), key=lambda b: b.get('extraction_order', 999))
                
                # Step 1: Extract ALL text elements with coordinates
                all_text_elements = self._extract_all_text_elements(page)
                
                # Step 2: Assign text elements to boxes based on center point intersection
                box_assignments = self._assign_text_to_boxes(all_text_elements, boxes)
                
                # Step 3: Create unified content blocks for this page
                content_blocks = []
                
                # Add box content blocks
                for box in boxes:
                    box_label = box.get('label', 'Unknown')
                    if box_label in box_assignments and box_assignments[box_label]:
                        elements = box_assignments[box_label]
                        box_text = self._process_box_elements(elements, box, page)
                        box_text = self._post_process_text(box_text)

                        # Get box coordinates for sorting
                        coords = box.get('coordinates', [0, 0, 0, 0])
                        y_pos = coords[1] if len(coords) >= 2 else 0  # Use top of box
                        x_pos = coords[0] if len(coords) >= 1 else 0  # Use left of box
                        
                        content_blocks.append({
                            'content': box_text,
                            'y_pos': y_pos,
                            'x_pos': x_pos,
                            'type': 'box',
                            'label': box_label
                        })
                
                # Add unboxed content blocks (only if template allows it)
                include_unboxed = template_data.get('include_unboxed_content', True)  # Default to True for backward compatibility
                if include_unboxed and '_UNBOXED_' in box_assignments:
                    unboxed_elements = box_assignments['_UNBOXED_']
                    if unboxed_elements:
                        unboxed_blocks = self._group_unboxed_into_blocks(unboxed_elements)
                        for block in unboxed_blocks:
                            content_blocks.append({
                                'content': block['content'],
                                'y_pos': block['y_pos'],
                                'x_pos': block['x_pos'],
                                'type': 'unboxed'
                            })
                
                # Step 4: Sort all content blocks by reading order (top-to-bottom, then left-to-right)
                content_blocks.sort(key=lambda block: (block['y_pos'], block['x_pos']))
                
                # Step 5: Add all content in reading order to output
                for block in content_blocks:
                    output_parts.append(block['content'])
                    output_parts.append("")  # Add blank line after each block for readability
            
            doc.close()
            return "\n".join(output_parts)
            
        except Exception as e:
            return f"ERROR: Template extraction failed: {e}"
    
    def extract_bol_text(self, pdf_path: str) -> str:
        """Extract text from BOL using template matching."""
        try:            
            # Find best matching template
            best_match = self.find_best_template(pdf_path)
            
            if not best_match or best_match.confidence < self.confidence_threshold:
                return f"ERROR: No suitable template found (best match: {best_match.template_name if best_match else 'None'} with confidence {best_match.confidence if best_match else 0.0:.2f})"
            
            # Use a unified extraction function for both single and multi-page
            return self._extract_with_template(pdf_path, best_match.template_data)
            
        except Exception as e:
            return f"ERROR: Failed to process BOL: {e}"

    def find_best_template(self, pdf_path: str) -> Optional[TemplateMatch]:
        """Find the best matching template for the given PDF by examining all pages."""
        
        # Extract text from ALL pages of the PDF for comprehensive fingerprinting
        doc = fitz.open(pdf_path)
        total_pages = len(doc)
        all_pages_text = []
        
        # Extract text from each page
        for page_num in range(total_pages):
            page = doc[page_num]
            page_text = page.get_text()
            all_pages_text.append(page_text)
        
        doc.close()
        
        # Combine all pages text for global matching
        combined_text = " ".join(all_pages_text).upper()  # Case-insensitive matching
        
        best_match = None
        best_confidence = 0.0
                
        for template_name, template_data in self.templates.items():
            confidence = self._calculate_template_confidence(combined_text, template_data, pdf_path)
            
            if confidence > best_confidence:
                best_confidence = confidence
                best_match = TemplateMatch(template_name, confidence, template_data)
        
        return best_match
    
    def _calculate_template_confidence(self, sample_text: str, template_data: dict, pdf_path: str) -> float:
        """Calculate confidence by comparing full raw text of template vs document."""
        template_fingerprint = template_data.get('template_raw_text', '')
        
        if not template_fingerprint:
            return 0.0
                
        # Use fuzzy similarity matching between full texts
        similarity = TextNormalizer.calculate_fuzzy_similarity(template_fingerprint, sample_text)
            
        return similarity
    
    def _extract_all_text_elements(self, page) -> List[TextElement]:
        """Extract all text elements from the page with their coordinates."""
        words = page.get_text("words")
        text_elements = []
        
        for word_data in words:
            x0, y0, x1, y1, text = word_data[0], word_data[1], word_data[2], word_data[3], word_data[4]
            center_x = (x0 + x1) / 2
            center_y = (y0 + y1) / 2
            
            text_elements.append(TextElement(
                text=text,
                x0=x0, y0=y0, x1=x1, y1=y1,
                center_x=center_x, center_y=center_y
            ))
        
        return text_elements
    
    def _assign_text_to_boxes(self, text_elements: List[TextElement], boxes: List[dict]) -> Dict[str, List[TextElement]]:
        """Assign text elements to boxes based on center point intersection."""
        # Create box assignments dictionary
        box_assignments = {}
        unboxed_elements = []
        
        # Initialize box assignments
        for box in boxes:
            box_label = box.get('label', 'Unknown')
            box_assignments[box_label] = []
        
        # Check each text element against all boxes
        for element in text_elements:
            assigned = False
            
            for box in boxes:
                coords = box.get('coordinates')
                if not coords or len(coords) != 4:
                    continue
                
                x0_box, y0_box, x1_box, y1_box = coords
                box_label = box.get('label', 'Unknown')
                
                # Check if center point is within box
                if (x0_box <= element.center_x <= x1_box and 
                    y0_box <= element.center_y <= y1_box):
                    element.box_assignment = box_label
                    box_assignments[box_label].append(element)
                    assigned = True
                    break
            
            if not assigned:
                unboxed_elements.append(element)
        
        # Add unboxed elements to the assignments
        if unboxed_elements:
            box_assignments['_UNBOXED_'] = unboxed_elements
        
        return box_assignments
     
    def _process_box_elements(self, elements: List[TextElement], box_data: dict, page) -> str:
        """Process text elements for a specific box using the box's extraction logic."""
        if not elements:
            return ""
        
        box_type = box_data.get('box_type', 'general')
        
        # Convert TextElements back to word format for existing logic
        words = []
        for element in elements:
            words.append([element.x0, element.y0, element.x1, element.y1, element.text])
        
        if box_type == 'table':
            return self._extract_table_text_from_elements(elements, box_data)
        elif box_type == 'paragraph':
            return self._extract_paragraph_from_words(words)
        else:  # general box
            return self._extract_with_layout_detection(words)
    
    def _extract_table_text_from_elements(self, elements: List[TextElement], box_data: dict) -> str:
        """Extract table text using assigned text elements and cell structure."""
        cells = box_data.get('table_cells', [])
        if not cells:
            # Fallback to regular extraction if no cell data
            words = [[e.x0, e.y0, e.x1, e.y1, e.text] for e in elements]
            return self._extract_with_layout_detection(words)
        
        # Sort cells by their cell_id to maintain original order
        ordered_cells = sorted(cells, key=lambda c: c['cell_id'])
        
        # Create cell_id to text mapping
        cell_text_map = {}
        max_row = max(cell['row'] for cell in cells)
        max_col = max(cell['col'] for cell in cells)
        
        # Initialize all cells
        for cell in ordered_cells:
            cell_text_map[cell['cell_id']] = {
                'text': "",
                'row': cell['row'],
                'col': cell['col'],
                'elements': []
            }
        
        # Assign elements to cells based on center point intersection
        tolerance = 2
        unassigned_elements = []
        
        for element in elements:
            assigned = False
            for cell in ordered_cells:
                cx0, cy0, cx1, cy1 = cell['coordinates']
                if (cx0 - tolerance <= element.center_x <= cx1 + tolerance and 
                    cy0 - tolerance <= element.center_y <= cy1 + tolerance):
                    cell_text_map[cell['cell_id']]['elements'].append(element)
                    assigned = True
                    break
            
            if not assigned:
                unassigned_elements.append(element)
        
        # Assign unassigned elements to closest existing cells
        for element in unassigned_elements:
            min_distance = float('inf')
            closest_cell_id = None
            
            for cell in ordered_cells:
                cx0, cy0, cx1, cy1 = cell['coordinates']
                # Calculate distance to cell center
                cell_center_x = (cx0 + cx1) / 2
                cell_center_y = (cy0 + cy1) / 2
                distance = ((element.center_x - cell_center_x) ** 2 + 
                        (element.center_y - cell_center_y) ** 2) ** 0.5
                
                if distance < min_distance:
                    min_distance = distance
                    closest_cell_id = cell['cell_id']
            
            # Assign to closest cell (no distance threshold)
            if closest_cell_id is not None:
                cell_text_map[closest_cell_id]['elements'].append(element)
        
        # Process each cell's elements into text
        for cell_id, cell_info in cell_text_map.items():
            if cell_info['elements']:
                # Convert elements to words format and group into lines
                words = [[e.x0, e.y0, e.x1, e.y1, e.text] for e in cell_info['elements']]
                lines = self._group_words_into_lines(words)
                cell_info['text'] = "\n".join(line['text'] for line in lines)
        
        # Convert to markdown table format
        if not cell_text_map:
            return ""
                
        # Build markdown table by following numeric cell IDs in order
        markdown_lines = []
        row_cells = []
        current_row = 0
        
        # Process cells in numeric order of cell IDs
        for cell_id in sorted(cell_text_map.keys(), key=int):
            cell_info = cell_text_map[cell_id]
            cell_text = cell_info['text'].replace("\n", " ").replace("|", "\\|").strip()
            
            # If we're on a new row
            if cell_info['row'] > current_row:
                # Add the completed row
                if row_cells:
                    markdown_lines.append("| " + " | ".join(row_cells) + " |")
                    # Add separator after header row
                    if current_row == 0 and max_row > 0:
                        markdown_lines.append("| " + " | ".join(["---"] * (max_col + 1)) + " |")
                # Start new row
                row_cells = []
                current_row = cell_info['row']
                
            row_cells.append(cell_text)
        
        # Add the last row if there are remaining cells
        if row_cells:
            markdown_lines.append("| " + " | ".join(row_cells) + " |")
        
        result = "\n".join(markdown_lines)
        return result

    def _extract_paragraph_from_words(self, words: List) -> str:
        """Extract paragraph text from words."""
        if not words:
            return ""
        
        lines = self._group_words_into_lines(words)
        
        text_lines = []
        for i, line in enumerate(lines):
            text = line['text']
            
            # Add line breaks for large vertical gaps (paragraph breaks)
            if i > 0:
                prev_line = lines[i-1]
                vertical_gap = line['center_y'] - prev_line['center_y']
                
                if vertical_gap > 20:
                    text_lines.append("")  # Add blank line
            
            text_lines.append(text)
        
        return "\n".join(text_lines)
    
    def _process_unboxed_elements(self, elements: List[TextElement]) -> str:
        """Process unboxed text elements using general layout detection."""
        if not elements:
            return ""
        
        # Convert TextElements back to word format
        words = []
        for element in elements:
            words.append([element.x0, element.y0, element.x1, element.y1, element.text])
        
        return self._extract_with_layout_detection(words)
    
    def _merge_content_by_reading_order(self, box_assignments: Dict[str, List[TextElement]], 
                                       processed_content: Dict[str, str], boxes: List[dict]) -> str:
        """Merge all content based on reading order (top-to-bottom, left-to-right)."""
        
        # Create content blocks with their positions
        content_blocks = []
        
        # Add box content blocks
        for box in boxes:
            box_label = box.get('label', 'Unknown')
            if box_label in processed_content and processed_content[box_label]:
                coords = box.get('coordinates', [0, 0, 0, 0])
                # Use top-left corner for sorting
                y_pos = coords[1] if len(coords) >= 2 else 0
                x_pos = coords[0] if len(coords) >= 1 else 0
                
                content_blocks.append({
                    'content': f"**=== {box_label.upper()} ===**\n{processed_content[box_label]}",
                    'y_pos': y_pos,
                    'x_pos': x_pos,
                    'type': 'box'
                })
        
        # Add unboxed content blocks (group by approximate lines)
        if '_UNBOXED_' in box_assignments:
            unboxed_elements = box_assignments['_UNBOXED_']
            if unboxed_elements:
                # Group unboxed elements into logical blocks
                unboxed_blocks = self._group_unboxed_into_blocks(unboxed_elements)
                
                for block in unboxed_blocks:
                    content_blocks.append({
                        'content': f"**=== OTHER TEXT ===**\n{block['content']}",
                        'y_pos': block['y_pos'],
                        'x_pos': block['x_pos'],
                        'type': 'unboxed'
                    })
        
        # Sort by reading order: top-to-bottom, then left-to-right
        content_blocks.sort(key=lambda block: (block['y_pos'], block['x_pos']))
        
        # Combine all content
        all_content = [block['content'] for block in content_blocks]
        
        return "\n\n".join(all_content)
    
    def _group_unboxed_into_blocks(self, elements: List[TextElement]) -> List[dict]:
        """Group unboxed elements into logical content blocks based on vertical proximity."""
        if not elements:
            return []
        
        # Sort elements by position (top-to-bottom, left-to-right)
        sorted_elements = sorted(elements, key=lambda e: (e.y0, e.x0))
        
        # Group into blocks based on vertical proximity
        blocks = []
        current_block_elements = []
        vertical_threshold = 20  # pixels
        
        for element in sorted_elements:
            if (current_block_elements and 
                element.y0 - current_block_elements[-1].y1 > vertical_threshold):
                
                # Process current block
                if current_block_elements:
                    block_content = self._process_unboxed_block(current_block_elements)
                    if block_content.strip():
                        blocks.append({
                            'content': block_content,
                            'y_pos': min(e.y0 for e in current_block_elements),
                            'x_pos': min(e.x0 for e in current_block_elements)
                        })
                
                # Start new block
                current_block_elements = [element]
            else:
                current_block_elements.append(element)
        
        # Process final block
        if current_block_elements:
            block_content = self._process_unboxed_block(current_block_elements)
            if block_content.strip():
                blocks.append({
                    'content': block_content,
                    'y_pos': min(e.y0 for e in current_block_elements),
                    'x_pos': min(e.x0 for e in current_block_elements)
                })
        
        return blocks

    
    def _process_unboxed_block(self, elements: List[TextElement]) -> str:
        """Process a block of unboxed elements using general layout detection."""
        if not elements:
            return ""
        
        # Convert TextElements back to word format
        words = []
        for element in elements:
            words.append([element.x0, element.y0, element.x1, element.y1, element.text])
        
        return self._extract_with_layout_detection(words)    
    
    def _extract_with_layout_detection(self, words: List) -> str:
        """Extract text with simple layout detection for general boxes."""
        if not words:
            return ""
        
        # Simple line grouping and text extraction
        lines = self._group_words_into_lines(words)
        
        # Simple paragraph formatting
        text_lines = []
        for i, line in enumerate(lines):
            text = line['text']
            
            # Add line breaks for large vertical gaps
            if i > 0:
                prev_line = lines[i-1]
                vertical_gap = line['center_y'] - prev_line['center_y']
                
                if vertical_gap > 25:
                    text_lines.append("")  # Add blank line
            
            text_lines.append(text)
        
        return "\n".join(text_lines)
    
    def _group_words_into_lines(self, words: List) -> List[dict]:
        """Group words into lines based on vertical position."""
        lines = []
        tolerance = 8  # pixels tolerance for line grouping
        
        for word in words:
            x0, y0, x1, y1, text = word[0], word[1], word[2], word[3], word[4]
            word_center_y = (y0 + y1) / 2
            
            # Find existing line within tolerance
            found_line = False
            for line in lines:
                if abs(word_center_y - line['center_y']) <= tolerance:
                    line['words'].append((x0, y0, x1, y1, text))
                    # Update line center to average
                    all_centers = [(w[1] + w[3]) / 2 for w in line['words']]
                    line['center_y'] = sum(all_centers) / len(all_centers)
                    found_line = True
                    break
            
            if not found_line:
                lines.append({
                    'center_y': word_center_y,
                    'words': [(x0, y0, x1, y1, text)]
                })
        
        # Sort lines top to bottom and words left to right within each line
        lines.sort(key=lambda line: line['center_y'])
        for line in lines:
            line['words'].sort(key=lambda w: w[0])
            line['text'] = " ".join(w[4] for w in line['words']).strip()
            
        return [line for line in lines if line['text']]  # Remove empty lines
    
    def _post_process_text(self, text: str) -> str:
        """Apply final formatting fixes after extraction."""
        if not text:
            return text
        
        # 1. Fix spaced out words (e.g., "H E L L O" → "HELLO")
        text = self._fix_spaced_words(text)
                
        # 3. Normalize whitespace
        text = self._normalize_whitespace(text)
        
        # 4. Fix punctuation spacing
        text = self._fix_punctuation(text)
        
        # 5. Common OCR error corrections
        text = self._fix_ocr_errors(text)
        
        # 6. Case normalization for specific patterns
        #text = self._normalize_case(text)
        
        return text
    
    def _fix_spaced_words(self, text: str) -> str:
        """Fix words with excessive spacing between letters."""
        # Pattern: single letters with spaces (min 3 letters)
        # "H E L L O" or "A B C D E F"
        pattern = r'\b([A-Za-z])\s+([A-Za-z])\s+([A-Za-z]+(?:\s+[A-Za-z])*)\b'
        
        def collapse_spaced(match):
            # Remove all spaces from the matched group
            return re.sub(r'\s+', '', match.group(0))
        
        return re.sub(pattern, collapse_spaced, text)

    def _normalize_whitespace(self, text: str) -> str:
        """Clean up excessive whitespace."""
        # Multiple spaces → single space
        text = re.sub(r' {2,}', ' ', text)
        
        # Multiple newlines → double newline max
        text = re.sub(r'\n{3,}', '\n\n', text)
        
        # Remove trailing spaces on lines
        text = re.sub(r' +\n', '\n', text)
        
        return text.strip()
    
    def _fix_punctuation(self, text: str) -> str:
        """Fix spacing around punctuation."""
        # Add space after punctuation if missing
        text = re.sub(r'([,;:])([A-Za-z])', r'\1 \2', text)
        
        # Remove space before punctuation
        text = re.sub(r'\s+([,;:.])', r'\1', text)
        
        return text
    
    def _fix_ocr_errors(self, text: str) -> str:
        """Fix common OCR recognition errors."""
        ocr_fixes = {
            # Number/letter confusion
            r'\b0(?=[A-Za-z])': 'O',  # 0 at start of word → O
            r'(?<=[A-Za-z])0\b': 'O',  # 0 at end of word → O
            r'\bl(?=[0-9])': '1',      # lowercase l before numbers → 1
            
            # Common character substitutions
            r'\brn\b': 'm',            # rn → m
            r'\bvv\b': 'w',            # vv → w  
            r'(?<=[a-z])I(?=[a-z])': 'l',  # I between lowercase → l
        }
        
        for pattern, replacement in ocr_fixes.items():
            text = re.sub(pattern, replacement, text)
        
        return text
    
    def _normalize_case(self, text: str) -> str:
        """Normalize case for specific patterns."""
        # Fix ALL CAPS words that should be title case
        def fix_caps(match):
            word = match.group(0)
            # Keep short words (like IDs) in caps
            if len(word) <= 3:
                return word
            # Make title case for longer words
            return word.title()
        
        # Apply to words that are all caps (but preserve intentional caps)
        text = re.sub(r'\b[A-Z]{4,}\b', fix_caps, text)
        
        return text
        
    def _extract_paragraph_text(self, page, rect) -> str:
        """Extract text from a paragraph box using simple line-by-line extraction."""
        words = page.get_text("words", clip=rect)
        if not words:
            return ""
        
        # Group words into lines with minimal processing
        lines = self._group_words_into_lines(words)
        
        # Format as flowing paragraph text
        text_lines = []
        for i, line in enumerate(lines):
            text = line['text']
            
            # Add line breaks for large vertical gaps (paragraph breaks)
            if i > 0:
                prev_line = lines[i-1]
                vertical_gap = line['center_y'] - prev_line['center_y']
                
                # Large vertical gap suggests paragraph break
                if vertical_gap > 20:
                    text_lines.append("")  # Add blank line
            
            text_lines.append(text)
        
        return "\n".join(text_lines)
    
def select_pdf_file():
    """Select a PDF file using file dialog."""
    try:
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        
        file_path = filedialog.askopenfilename(
            title="Select BOL PDF File (NOT template file)",
            filetypes=[("PDF files", "*.pdf")],  # Only show PDFs
            initialdir=os.getcwd()  # Start in current directory, not templates
        )
        
        root.destroy()
        return file_path if file_path else None
        
    except ImportError:
        print("tkinter not available.")
        return None


def main():
    # Initialize extractor with fuzzy matching
    extractor = BOLTemplateExtractor(
        templates_dir="templates",
        confidence_threshold=0.5  # Lowered for fuzzy matching
    )
    
    if not extractor.templates:
        print("\nNo templates found!")
        return
        
    pdf_path = select_pdf_file()
    if not pdf_path:
        print("No file selected.")
        return
    
    print(f"\nProcessing: {os.path.basename(pdf_path)}")
    
    # Extract text
    result = extractor.extract_bol_text(pdf_path)

    # Save result
    output_filename = f"extracted_{os.path.splitext(os.path.basename(pdf_path))[0]}_template.md"
    try:
        with open(output_filename, 'w', encoding='utf-8') as f:
            f.write(result)
        print(f"\nResults saved to: {output_filename}")
    except Exception as e:
        print(f"Error saving results: {e}")


if __name__ == "__main__":
    main()