import subprocess
import sys
from pathlib import Path
import shutil
import os
import re
import zipfile
import logging
import chardet
import glob

# Configuration
DOCX_FILE = "Paper.docx"
#DOCX_FILE = "C:\\Users\\earth\\OneDrive\\Documents\\Ashish Kumari PhD\\Paper3\\researchpaper\\Research_paper_ashish3_v2_title+abstract.docx"
#DOCX_FILE = "C:\\Users\\earth\\OneDrive\\Documents\\Annu bhure Phd MSIT\\Papers as in paper\\Annu_research_paper_160625.docx"
#LATEX_TEMPLATE = "C:\\Users\\earth\\Downloads\\ET\\MSP-Latex-Template\\template.tex"
#LATEX_TEMPLATE = "C:\\Users\\earth\\Downloads\\springer\\sn-article-template\\sn-article.tex"
LATEX_TEMPLATE  = "output/corrected.tex"
#OUTPUT_DIR = "C:\\Users\\earth\\Downloads\\ET\\MSP-Latex-Template"
#OUTPUT_DIR = "C:\\Users\\earth\\Downloads\\springer\\sn-article-template"
OUTPUT_DIR = ""
TEMP_DIR = "temp_conversion"
COMPILE_PDF = True

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)

class SimplifiedDOCXConverter:
    def __init__(self, docx_path, template_path, output_dir):
        self.docx_path = Path(docx_path)
        self.template_path = Path(template_path)
        self.output_dir = Path(output_dir)
        self.temp_dir = Path(TEMP_DIR)
    
    @staticmethod
    def find_cls_file(template_directory):
        """
        Find the .cls file in the LaTeX template directory
        
        Args:
            template_directory (str): Path to the LaTeX template directory
        
        Returns:
            str: Path to the .cls file or None if not found
        """
        # Search for .cls files in the template directory
        cls_files = glob.glob(os.path.join(template_directory, "*.cls"))
        
        # If not found in root, search one level deep
        if not cls_files:
            cls_files = glob.glob(os.path.join(template_directory, "**", "*.cls"), recursive=True)
        
        # Return the first .cls file found
        return cls_files[0] if cls_files else None

    def check_dependencies(self):
        """Check if required tools are installed"""
        required_tools = {
            'pandoc': 'pandoc --version',
            'pdflatex': 'pdflatex --version'
        }
        
        missing = []
        for tool, cmd in required_tools.items():
            try:
                subprocess.run(cmd.split(), capture_output=True, check=True)
                print(f"✓ {tool} found")
            except (subprocess.CalledProcessError, FileNotFoundError):
                missing.append(tool)
                print(f"❌ {tool} not found")
        
        if missing:
            print(f"\nPlease install missing tools: {', '.join(missing)}")
            return False
        return True
    
    def extract_images_from_docx(self):
        """Extract images from DOCX file"""
        images_dir = self.temp_dir / "images"
        images_dir.mkdir(parents=True, exist_ok=True)
        
        try:
            with zipfile.ZipFile(self.docx_path, 'r') as docx_zip:
                media_files = [f for f in docx_zip.namelist() if f.startswith('word/media/')]
                
                for media_file in media_files:
                    filename = Path(media_file).name
                    with docx_zip.open(media_file) as source, \
                         open(images_dir / filename, 'wb') as target:
                        shutil.copyfileobj(source, target)
                
                print(f"✓ Extracted {len(media_files)} media files")
                return list(images_dir.glob("*"))
                
        except Exception as e:
            print(f"⚠ Warning: Could not extract images: {e}")
            return []
    
    def convert_docx_to_latex(self):
        """Step 1: Convert DOCX to LaTeX using pandoc"""
        print("Step 1: Converting DOCX to LaTeX...")
        
        if not self.docx_path.exists():
            print(f"❌ Input file not found: {self.docx_path}")
            return None
        
        latex_file = self.temp_dir / "converted.tex"
        
        pandoc_cmd = [
            'pandoc',
            str(self.docx_path),
            '-o', str(latex_file),
            '--to=latex',
            '--standalone',  # Include full document structure
            '--extract-media', str(self.temp_dir),
        ]
        
        try:
            result = subprocess.run(pandoc_cmd, capture_output=True, text=True, check=True)
            print("✓ DOCX converted to LaTeX")
            return latex_file
        except subprocess.CalledProcessError as e:
            print(f"❌ Pandoc conversion failed: {e.stderr}")
            return None
        except FileNotFoundError:
            print("❌ Pandoc not found. Please install pandoc first.")
            return None

    def extract_preamble_from_template(self):
        """Step 2: Extract everything before \begin{document} from template"""
        print("Step 2: Extracting preamble from template...")
        
        if not self.template_path.exists():
            print(f"❌ Template file {self.template_path} not found!")
            return ""
        
        # Try multiple encodings to read the template file
        encodings_to_try = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
        
        for encoding in encodings_to_try:
            try:
                with open(self.template_path, 'r', encoding=encoding) as f:
                    template_content = f.read()
                print(f"✓ Successfully read template with {encoding} encoding")
                break
            except UnicodeDecodeError:
                continue
        else:
            # If all encodings fail, try with error handling
            try:
                with open(self.template_path, 'r', encoding='utf-8', errors='ignore') as f:
                    template_content = f.read()
                print("✓ Read template with utf-8 encoding (ignoring errors)")
            except Exception as e:
                print(f"❌ Failed to read template file: {e}")
                return ""
        
        begin_doc_match = re.search(r'\\begin\{document\}', template_content)
        if begin_doc_match:
            # Keep only content from \begin{document} onwards
            preamble = template_content[:begin_doc_match.start()].strip()
            print(f"✓ Extracted preamble ({len(preamble)} characters)")
            return preamble
        else:
            print("⚠ Warning: \\begin{document} not found in template")
            return template_content.strip()
    
    @staticmethod
    def detect_encoding(file_path):
        """Detect file encoding with better error handling"""
        try:
            with open(file_path, 'rb') as f:
                raw_data = f.read()
                result = chardet.detect(raw_data)
                encoding = result['encoding']
                confidence = result['confidence']
                
                if encoding and confidence > 0.7:
                    print(f"Detected encoding: {encoding} (confidence: {confidence:.2f})")
                    return encoding
                else:
                    print(f"Low confidence encoding detection: {encoding} (confidence: {confidence:.2f})")
                    return 'utf-8'  # Fallback to utf-8
        except Exception as e:
            print(f"Error detecting encoding: {e}")
            return 'utf-8'  # Default fallback

    def detect_column_layout(self):
        """
        Comprehensive detection with auto-encoding detection
        """
        # Get the template directory
        template_directory = str(self.template_path.parent)
        
        # Detect the cls file
        cls_file_path = self.find_cls_file(template_directory)
        if cls_file_path:
            print(f'Found the cls file: {cls_file_path}')
        else:
            print('No .cls file found')

        # Detect and use correct encoding for .tex file
        tex_encoding = self.detect_encoding(self.template_path) or 'utf-8'
        
        # Read template with detected encoding and error handling
        try:
            with open(self.template_path, 'r', encoding=tex_encoding) as f:
                tex_content = f.read()
        except UnicodeDecodeError:
            # Fallback to utf-8 with error ignoring
            with open(self.template_path, 'r', encoding='utf-8', errors='ignore') as f:
                tex_content = f.read()
            print("Used utf-8 fallback with error ignoring for template")
        
        # Check documentclass options
        docclass_pattern = r'\\documentclass\[(.*?)\]\{.*?\}'
        match = re.search(docclass_pattern, tex_content)
        
        if match:
            options = match.group(1)
            options_list = [opt.strip() for opt in options.split(',')]
            
            if 'onecolumn' in options_list:
                return 'onecolumn'
            elif 'twocolumn' in options_list:
                return 'twocolumn'
        
        # If no explicit option, check .cls defaults
        if cls_file_path:
            cls_encoding = self.detect_encoding(cls_file_path) or 'utf-8'
            try:
                with open(cls_file_path, 'r', encoding=cls_encoding) as f:
                    cls_content = f.read()
            except UnicodeDecodeError:
                with open(cls_file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    cls_content = f.read()
                print("Used utf-8 fallback with error ignoring for cls file")
            
            # Look for ExecuteOptions pattern
            execute_pattern = r'\\ExecuteOptions\{([^}]*)\}'
            matches = re.findall(execute_pattern, cls_content)
            
            for options_str in matches:
                options = [opt.strip() for opt in options_str.split(',')]
                if 'twocolumn' in options:
                    return 'twocolumn'
                elif 'onecolumn' in options:
                    return 'onecolumn'
        
        return 'onecolumn'  # Default for IEEE

    def get_optimal_image_settings(self, column_layout):
        """Get optimal image settings based on column layout"""
        if column_layout == "twocolumn":
            return {
                'small': '[width=0.48\\columnwidth]',      # Small images fit in one column
                'medium': '[width=0.48\\columnwidth]',     # Medium images fit in one column  
                'large': '[width=0.48\\textwidth]',        # Large images span both columns
                'figure_env': 'figure'                   # Use figure* for spanning images
            }
        else:  # onecolumn
            return {
                'small': '[width=0.6\\textwidth]',        # Small images
                'medium': '[width=0.8\\textwidth]',       # Medium images
                'large': '[width=0.9\\textwidth]',        # Large images
                'figure_env': 'figure'                    # Regular figure environment
            }

    def process_images_optimally(self, content, column_layout):
        """Process images with optimal sizing based on layout"""
        image_settings = self.get_optimal_image_settings(column_layout)
        
        def replace_image(match):
            original_options = match.group(1) or ''
            image_path = match.group(2)
            image_name = Path(image_path).name
            
            if 'width=' in original_options:
                # Extract width value to determine size category
                width_match = re.search(r'width=([0-9.]+)', original_options)
                if width_match:
                    width_val = float(width_match.group(1))
                    if width_val <= 0.4:
                        size_category = 'small'
                    elif width_val <= 0.7:
                        size_category = 'medium'
                    else:
                        size_category = 'large'
                else:
                    size_category = 'medium'
            else:
                # Default to medium if no width specified
                size_category = 'medium'
            
            # Get optimal settings for this size category
            optimal_width = image_settings[size_category]
            
            if column_layout == "twocolumn" and size_category == 'large':
                return (f'\\begin{{{image_settings["figure_env"]}}}[htbp]\n'
                       f'\\centering\n'
                       f'\\includegraphics{optimal_width}{{images/{image_name}}}\n'
                       f'\\end{{{image_settings["figure_env"]}}}')
            else:
                return f'\\includegraphics{optimal_width}{{images/{image_name}}}'
        
        # Replace all includegraphics commands
        processed_content = re.sub(
            r'\\includegraphics(\[[^\]]*\])?\{([^}]+)\}',
            replace_image,
            content
        )
        
        return processed_content

    def is_algorithm_table(self, table_content):
        """Check if this is an algorithm table that should be skipped"""
        algorithm_patterns = [
            r'\\begin\{algorithm\}',
            r'\\begin\{alg\w*\}',
            r'\\Procedure',
            r'\\Function',
            r'\\For\{',
            r'\\While\{',
            r'\\State',
            r'\\End',
            r'\\algorithmic'
        ]
        return any(re.search(pattern, table_content) for pattern in algorithm_patterns)

    # def process_single_table(self, match):
    #     full_table = match.group(0)
    #     print(f"  Processing table with {len(full_table)} characters")

    #     # Skip algorithm tables completely
    #     if self.is_algorithm_table(full_table):
    #         print("  [Skipping algorithm table]")
    #         return full_table  # Return the original untouched

    #     # Original table processing logic from working version
    #     col_count = 4  # Default fallback
        
    #     # Method 1: Look for column specification in table definition
    #     col_spec_patterns = [
    #         r'\\begin\{longtable\}(?:\[\])?\{([^}]+)\}',
    #         r'\\begin\{tabular\}(?:\[[^\]]*\])?\{([^}]+)\}',
    #         r'\\begin\{table\}.*?\\begin\{tabular\}(?:\[[^\]]*\])?\{([^}]+)\}'
    #     ]
        
    #     for pattern in col_spec_patterns:
    #         col_match = re.search(pattern, full_table, re.DOTALL)
    #         if col_match:
    #             col_spec = col_match.group(1)
    #             print(f"  Found column spec: {col_spec}")
    #             # Count actual column types, ignoring formatting
    #             col_spec_clean = re.sub(r'[^lcrp]', '', col_spec)  # Keep only column types
    #             if col_spec_clean:
    #                 col_count = len(col_spec_clean)
    #                 print(f"  Detected {col_count} columns from spec")
    #                 break
        
    #     # Method 2: If still no columns found, count & in content
    #     if col_count == 0 or col_count > 10:  # Sanity check
    #         amp_lines = re.findall(r'[^\\\\]*&[^\\\\]*(?:\\\\|$)', full_table)
    #         if amp_lines:
    #             # Take the most common & count + 1
    #             counts = [line.count('&') + 1 for line in amp_lines[:5]]
    #             col_count = max(set(counts), key=counts.count) if counts else 4
    #             print(f"  Detected {col_count} columns from & count")
        
    #     # Ensure reasonable column count
    #     if col_count <= 0:
    #         col_count = 4
    #         print(f"  Using default {col_count} columns")
    #     elif col_count > 10:
    #         col_count = 6  # Reasonable maximum
    #         print(f"  Limiting to {col_count} columns")
        
    #     # Find header texts - look for \textbf patterns
    #     header_texts = []
    #     textbf_matches = re.findall(r'\\textbf\{([^}]+)\}', full_table)
        
    #     print(f"  Found textbf headers: {textbf_matches}")
        
    #     # Take first col_count headers
    #     for i in range(col_count):
    #         if i < len(textbf_matches):
    #             # Clean up the header text
    #             header = textbf_matches[i].strip()
    #             # Remove any LaTeX commands from header
    #             header = re.sub(r'\\[a-zA-Z]+\{[^}]*\}', '', header)
    #             header = re.sub(r'\\[a-zA-Z]+', '', header).strip()
    #             header_texts.append(header if header else f'Column {i+1}')
    #         else:
    #             header_texts.append(f'Column {i+1}')
        
    #     print(f"  Using headers: {header_texts}")
        
    #     # Create the new table structure
    #     col_spec = 'l' * col_count
    #     new_table_lines = [
    #         f'\\resizebox{{\\linewidth}}{{!}}{{%',
    #         f'  \\begin{{tabular}}{{{col_spec}}}'
    #     ]
        
    #     # Add headers
    #     header_row = '    ' + ' & '.join([f'\\textbf{{{text}}}' for text in header_texts]) + ' \\\\'
    #     new_table_lines.extend([
    #         '    \\toprule',
    #         header_row,
    #         '    \\midrule'
    #     ])
        
    #     # Extract all content after midrule to end of table
    #     midrule_match = re.search(r'\\midrule(.*?)\\end\{longtable\}', full_table, re.DOTALL)
    #     if midrule_match:
    #         # Keep all content after midrule
    #         table_data = midrule_match.group(1).strip()
    #         # Add the data lines directly
    #         new_table_lines.append(table_data)
    #     else:
    #         # Fallback: try to find any data content
    #         lines = full_table.split('\n')
    #         data_lines = []
            
    #         for line in lines:
    #             # Skip unwanted lines
    #             if any(skip in line for skip in [
    #                 '\\begin{', '\\end{', '\\toprule', '\\midrule', '\\bottomrule',
    #                 '\\textbf{', '\\raggedright', '\\arraybackslash', '\\endhead',
    #                 '\\endfoot', '\\minipage', '\\noalign'
    #             ]):
    #                 continue
                
    #             # Look for data lines (contain & and end with \\)
    #             if '&' in line and '\\\\' in line:
    #                 # Clean the line
    #                 clean_line = line.strip()
    #                 clean_line = re.sub(r'\\raggedright\s*', '', clean_line)
    #                 clean_line = re.sub(r'\\arraybackslash\s*', '', clean_line)
    #                 if clean_line:
    #                     data_lines.append(f'    {clean_line}')
            
    #         # Add data lines
    #         new_table_lines.extend(data_lines)
        
    #     # Close table
    #     new_table_lines.extend([
    #         '    \\bottomrule',
    #         '  \\end{tabular}%',
    #         '}'
    #     ])
        
    #     result = '\n'.join(new_table_lines)
    #     print(f"  Table processed successfully")
    #     return result

     
    # def process_tables_custom(self, content):
    #     """Process longtables and regular tables with custom formatting"""
    #     print("Processing tables with custom formatting...")
        
    #     # Pattern to match longtable environments
    #     processed_content = re.sub(
    #         r'\\begin\{longtable\}.*?\\end\{longtable\}',
    #         self.process_single_table,
    #         content,
    #         flags=re.DOTALL
    #     )
        
    #     # Pattern to match table environments (which may contain tabular)
    #     # This is more general and should catch most cases
    #     processed_content = re.sub(
    #         r'\\begin\{table\}(?:\*)?.*?\\end\{table\*?\}',
    #         self.process_single_table,
    #         processed_content,
    #         flags=re.DOTALL
    #     )
        
    #     # Pattern to match standalone tabular environments not wrapped in table/longtable
    #     # This is a bit trickier as pandoc might output raw tabulars
    #     # We'll look for tabulars that are not immediately preceded by \begin{table} or \begin{longtable}
        
    #     # First, find all tabular environments
    #     tabular_matches = list(re.finditer(
    #         r'\\begin\{tabular\}.*?\\end\{tabular\}',
    #         processed_content,
    #         flags=re.DOTALL
    #     ))
        
    #     # Iterate through them and process if they are standalone
    #     for match in reversed(tabular_matches): # Process in reverse to avoid index issues
    #         start_idx = match.start()
            
    #         # Check if this tabular is already inside a \begin{table} or \begin{longtable}
    #         # Look for the closest preceding \begin{table} or \begin{longtable}
    #         # and the closest preceding \end{table} or \end{longtable}
            
    #         text_before = processed_content[:start_idx]
            
    #         last_table_begin = max(text_before.rfind('\\begin{table}'), text_before.rfind('\\begin{longtable}'))
    #         last_table_end = max(text_before.rfind('\\end{table}'), text_before.rfind('\\end{longtable}'))
            
    #         if last_table_begin == -1 or last_table_end > last_table_begin:
    #             # This tabular is likely standalone or not properly wrapped
    #             # We'll treat it as a full table for processing
    #             processed_content = processed_content[:start_idx] + \
    #                                 self.process_single_table(match) + \
    #                                 processed_content[match.end():]
        
    #     print("✓ Custom table processing complete")
    #     return processed_content


    

    def process_tables_custom(self, content):
        """
        New robust method: cleans tables in content, replacing in-place
        """
        print("Processing tables (new logic) ...")
        def detect_num_columns(table_str):
            # Look for begin{longtable} or begin{tabular}, extract column spec
            colspec_match = re.search(r'\\begin\{(?:longtable|tabular)[^}]*\}\s*\{([^}]*)\}', table_str)
            if not colspec_match:
                # fallback: count max "&" in lines + 1
                amp_counts = [line.count('&') for line in table_str.splitlines() if '&' in line]
                return max(amp_counts) + 1 if amp_counts else 1
            colspec = colspec_match.group(1)

            # Count 'p{' occurrences or [lcr]
            n_p = len(re.findall(r'p\{[^\}]+\}', colspec))
            n_basic = len(re.findall(r'[lcr]', re.sub(r'p\{[^\}]+\}', '', colspec)))
            return n_p + n_basic

        def clean_latex_table(table_str):
            ncol = detect_num_columns(table_str)
            print(f'Found columns are {ncol}')
            # Remove minipage blocks
            table_str = re.sub(r'\\begin\{minipage\}.*?\\end\{minipage\}', '', table_str, flags=re.DOTALL)
            # Remove formatting lines
            lines = table_str.split('\n')
            cleaned_lines = []
            header_skipped = False

            for line in lines:
                lstripped = line.strip()
                if lstripped == '' or lstripped.startswith('\\toprule') or lstripped.startswith('\\midrule') \
                        or lstripped.startswith('\\bottomrule') or lstripped.startswith('\\noalign') \
                        or lstripped.startswith('\\endhead') or lstripped.startswith('\\endlastfoot'):
                    continue

                # Skip header line with column count - 1 &
                if not header_skipped and line.count('&') == (ncol-1):
                    header_skipped = True
                    continue

                cleaned_lines.append(line)

            # Remove any lines containing only whitespace after cleaning
            cleaned_lines = [l for l in cleaned_lines if l.strip() != '']
            return '\n'.join(cleaned_lines)

        def rep_table_table(match):
            return clean_latex_table(match.group(0))

        # Process all longtables
        content = re.sub(r'\\begin\{longtable\}.*?\\end\{longtable\}', rep_table_table, content, flags=re.DOTALL)
        # Process all tabulars not inside table/longtable already (standalone)
        # This is a simplified but robust approach. Adjust if you see odd behaviors.
        content = re.sub(r'\\begin\{tabular\}.*?\\end\{tabular\}', rep_table_table, content, flags=re.DOTALL)
        return content
    

    def merge_latex_with_template_preamble(self, converted_latex_file, template_preamble):
        """Step 2 continued: Merge converted LaTeX with template preamble"""
        print("Step 2: Merging LaTeX with template preamble...")
        
        # Try multiple encodings to read the converted LaTeX file
        encodings_to_try = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
        latex_content = ""
        
        for encoding in encodings_to_try:
            try:
                with open(converted_latex_file, 'r', encoding=encoding) as f:
                    latex_content = f.read()
                print(f"✓ Successfully read converted LaTeX with {encoding} encoding")
                break
            except UnicodeDecodeError:
                continue
        else:
            # If all encodings fail, try with error handling
            try:
                with open(converted_latex_file, 'r', encoding='utf-8', errors='ignore') as f:
                    latex_content = f.read()
                print("✓ Read converted LaTeX with utf-8 encoding (ignoring errors)")
            except Exception as e:
                print(f"❌ Failed to read converted LaTeX file: {e}")
                return ""
        
        begin_doc_match = re.search(r'\\begin\{document\}', latex_content)
        if begin_doc_match:
            document_content = latex_content[begin_doc_match.start():]
        else:
            document_content = f"\\begin{{document}}\n{latex_content}\n\\end{{document}}"
        
        required_packages = [
            "\\usepackage{graphicx}",
            "\\usepackage{booktabs}",
        ]
        
        for package in required_packages:
            if package not in template_preamble:
                template_preamble += f"\n{package}"
        
        final_content = template_preamble + "\n\n" + document_content
        
        column_layout = self.detect_column_layout()
        print(f"✓ Detected {column_layout} layout")
        
        final_content = self.process_tables_custom(final_content)
        print("✓ Tables processed with custom formatting")
        
        final_content = self.process_images_optimally(final_content, column_layout)
        print("✓ Images processed with optimal sizing")
        
        return final_content
    
    def compile_pdf(self, tex_file):
        """Step 3: Generate PDF"""
        if not COMPILE_PDF:
            print("Step 3: PDF compilation skipped (COMPILE_PDF = False)")
            return True
            
        print("Step 3: Compiling LaTeX to PDF...")
        
        original_dir = os.getcwd()
        
        try:
            os.chdir(self.output_dir)
            
            for i in range(2):
                print(f"  LaTeX pass {i+1}/2...")
                result = subprocess.run([
                    'pdflatex', 
                    '-interaction=nonstopmode',
                    '-file-line-error',
                    tex_file.name
                ], capture_output=True, text=True, timeout=120)  # 2 minute timeout
                
                if result.returncode != 0:
                    print(f"⚠ LaTeX pass {i+1} had issues (return code: {result.returncode})")
                    
                    # Show relevant error messages
                    if result.stdout:
                        error_lines = [line for line in result.stdout.split('\n') 
                                     if any(keyword in line.lower() for keyword in 
                                           ['error', 'undefined', 'missing', 'emergency stop'])]
                        if error_lines:
                            print("  Key errors found:")
                            for error_line in error_lines[:5]:  # Show first 5 errors
                                print(f"    {error_line.strip()}")
                    
                    if i == 1:  # Only show detailed error on final pass
                        print("\n  Full LaTeX output (last 800 chars):")
                        print(result.stdout[-800:] if result.stdout else "No stdout")
                        break
                else:
                    print(f"  ✓ LaTeX pass {i+1} completed successfully")
            
            pdf_file = tex_file.with_suffix('.pdf')
            if pdf_file.exists():
                print(f"✅ PDF created successfully: {pdf_file}")
                return True
            else:
                print("❌ PDF not created, but .tex file is available")
                print("Check the LaTeX file manually for compilation issues")
                return False
                
        except subprocess.TimeoutExpired:
            print("❌ PDF compilation timed out after 2 minutes")
            print("The LaTeX compilation may be stuck - check your .tex file")
            return False
        except FileNotFoundError:
            print("❌ pdflatex not found. Install LaTeX distribution.")
            return False
        except Exception as e:
            print(f"❌ PDF compilation failed: {e}")
            return False
        finally:
            os.chdir(original_dir)
    
    def convert(self):
        """Main simplified conversion process"""
        print("Starting simplified DOCX to LaTeX conversion...")
        print(f"Input file: {self.docx_path}")
        print(f"Template: {self.template_path}")
        print(f"Output directory: {self.output_dir}")
        
        if not self.check_dependencies():
            return False
        
        try:
            self.temp_dir.mkdir(exist_ok=True)
            self.output_dir.mkdir(exist_ok=True)
        except Exception as e:
            print(f"❌ Failed to create directories: {e}")
            return False
        
        try:
            self.extract_images_from_docx()
            
            converted_latex = self.convert_docx_to_latex()
            if not converted_latex:
                return False
            
            template_preamble = self.extract_preamble_from_template()
            final_latex = self.merge_latex_with_template_preamble(converted_latex, template_preamble)
            
            output_file = self.output_dir / "paper.tex"
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(final_latex)
            
            print(f"✅ LaTeX file created: {output_file}")
            
            images_src = self.temp_dir / "images"
            images_dst = self.output_dir / "images"
            if images_src.exists() and any(images_src.iterdir()):
                if images_dst.exists():
                    shutil.rmtree(images_dst)
                shutil.copytree(images_src, images_dst)
                print(f"✓ Images copied to: {images_dst}")
            
            pdf_success = self.compile_pdf(output_file)
            
            print(f"\n{'='*50}")
            print("CONVERSION COMPLETE!")
            print(f"{'='*50}")
            print(f"📄 LaTeX file: {output_file}")
            if pdf_success:
                print(f"📕 PDF file: {output_file.with_suffix('.pdf')}")
            else:
                print("📄 LaTeX file is available for manual compilation")
            if images_dst.exists():
                print(f"🖼️ Images: {images_dst}")
            
            return True
            
        except Exception as e:
            print(f"❌ Conversion failed: {e}")
            logger.exception("Detailed error information:")
            return False
        
        finally:
            if self.temp_dir.exists():
                try:
                    shutil.rmtree(self.temp_dir, ignore_errors=True)
                    print("✓ Cleaned up temporary files")
                except Exception as e:
                    print(f"⚠ Warning: Could not clean up temp directory: {e}")

def main():
    """Run the simplified converter"""
    print("Simplified DOCX to LaTeX Converter (Custom Table Processing)")
    print("=" * 60)
    
    if not Path(DOCX_FILE).exists():
        print(f"❌ Input file not found: {DOCX_FILE}")
        print("Please update the DOCX_FILE variable with the correct path.")
        return False
    
    converter = SimplifiedDOCXConverter(DOCX_FILE, LATEX_TEMPLATE, OUTPUT_DIR)
    success = converter.convert()
    
    if success:
        print("\n" + "="*50)
        print("CONVERSION SUCCESSFUL!")
        print("="*50)
        print("Tables processed with custom formatting requirements.")
        print("You can manually edit the LaTeX file if further adjustments are needed.")
    else:
        print("\n" + "="*50)
        print("CONVERSION FAILED")
        print("="*50)

if __name__ == "__main__":
    main()