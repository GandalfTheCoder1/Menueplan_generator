import pandas as pd
import os
import subprocess
import shutil
from tqdm import tqdm
import image_gen as ig
from PyPDF2 import PdfMerger, PdfReader
import tempfile


class LaTeXMenuGenerator:
    """Generate LaTeX menu tables from CSV files with customizable columns."""
    
    # Class constants
    DAY_NAMES = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag"]
    HEADER_COLORS = ['headerYellow', 'headerGreen', 'headerBlue', 'headerRed', 
                     'headerOrange', 'headerWhite', 'headerPink']
    PIKTOS = {"A": "A.jpg", "B": "B.jpg", "C": "C.jpg", "D": "D.jpg"}
    
    # Default custom column values
    DEFAULT_VALUES_WEEKDAY = ["", "", "T", "E", "S", ""]
    DEFAULT_VALUES_TUESDAY_THURSDAY = ["", "", "E", "S", "A"]
    
    # LaTeX special characters mapping
    LATEX_SPECIAL_CHARS = {
        '&': r'\&', '%': r'\%', '$': r'\$', '#': r'\#', '_': r'\_',
        '{': r'\{', '}': r'\}', '~': r'\textasciitilde{}',
        '^': r'\textasciicircum{}', '\\': r'\textbackslash{}'
    }
    
    def __init__(self, csv_folder, tex_folder="output_tex", pdf_folder="Menues", 
                 img_folder="output_img", aux_folder="log"):
        """Initialize the LaTeX menu generator."""
        self.csv_folder = csv_folder
        self.tex_folder = tex_folder
        self.pdf_folder = pdf_folder
        self.img_folder = img_folder
        self.aux_folder = aux_folder
        
        self._create_directories()
        self._copy_pikto_images()
    
    def _create_directories(self):
        """Create necessary directories."""
        for folder in [self.tex_folder, self.pdf_folder, self.img_folder, self.aux_folder]:
            os.makedirs(folder, exist_ok=True)
    
    def _copy_pikto_images(self):
        """Copy pikto images to tex folder."""
        for pikto_file in self.PIKTOS.values():
            src = os.path.join("Piktos", pikto_file)
            if os.path.exists(src):
                dst = os.path.join(self.tex_folder, pikto_file)
                shutil.copy2(src, dst)
    
    @staticmethod
    def has_content(items_list):
        """Check if the items list has any meaningful content."""
        return any(item.strip() for item in items_list 
                  if item and str(item).strip() not in ['-', '*', '', 'nan'])
    
    @staticmethod
    def is_pdf_blank(pdf_path):
        """Check if a PDF is effectively blank by examining its content."""
        try:
            reader = PdfReader(pdf_path)
            if len(reader.pages) == 0:
                return True
            
            for page in reader.pages:
                text = page.extract_text().strip()
                if len(text) > 50:
                    return False
            return True
        except Exception:
            return True
    
    @staticmethod
    def normalize_path_for_latex(path):
        """Convert path to forward slashes for LaTeX compatibility."""
        return path.replace(os.sep, '/')
    
    @staticmethod
    def escape_latex_text(text):
        """Escape LaTeX special characters in text."""
        escaped_text = str(text)
        for char, escape in LaTeXMenuGenerator.LATEX_SPECIAL_CHARS.items():
            escaped_text = escaped_text.replace(char, escape)
        return escaped_text
    
    @staticmethod
    def unescape_latex_text(text):
        """Unescape LaTeX special characters from text."""
        unescaped_text = text
        reverse_mapping = {
            '\\&': '&', '\\%': '%', '\\$': '$', '\\#': '#', '\\_': '_',
            '\\{': '{', '\\}': '}', '\\textasciitilde{}': '~',
            '\\textasciicircum{}': '^', '\\textbackslash{}': '\\'
        }
        for escape, char in reverse_mapping.items():
            unescaped_text = unescaped_text.replace(escape, char)
        return unescaped_text
    
    def get_latex_preamble(self):
        """Generate a robust LaTeX preamble."""
        return r"""\documentclass[a4paper,20pt]{article}

% Font setup - robust across platforms
\usepackage{lmodern}
\usepackage[T1]{fontenc}
\usepackage[utf8]{inputenc}

% Use a font that's available on most systems
\renewcommand{\rmdefault}{lmss}
\renewcommand{\sfdefault}{lmss}
\renewcommand{\ttdefault}{lmtt}

% Core packages
\usepackage[table]{xcolor}
\usepackage{booktabs}
\usepackage{array}
\usepackage{graphicx}
\usepackage{tabularx}
\usepackage{multirow}
\usepackage{colortbl}

% Geometry with consistent margins
\usepackage[a4paper, margin=2.54cm]{geometry}

% Keep normal spacing - we'll control row heights individually
\renewcommand{\arraystretch}{1.0}
\setlength{\tabcolsep}{6pt}

% Remove page numbers and headers
\pagestyle{empty}

% Ensure consistent image handling
\graphicspath{{./}}
\DeclareGraphicsExtensions{.png,.jpg,.jpeg,.pdf}

% Color definitions for consistency
\definecolor{headerYellow}{RGB}{255,255,204}
\definecolor{headerGreen}{RGB}{204,255,204}
\definecolor{headerBlue}{RGB}{204,204,255}
\definecolor{headerRed}{RGB}{255,204,204}
\definecolor{headerOrange}{RGB}{255,229,204}
\definecolor{headerWhite}{RGB}{248,248,248}
\definecolor{headerPink}{RGB}{255,204,255}

\definecolor{rowCyan}{RGB}{230,248,255}
\definecolor{rowYellow}{RGB}{255,255,230}
\definecolor{rowGreen}{RGB}{240,255,240}
\definecolor{rowRed}{RGB}{255,240,240}

\begin{document}
\sffamily"""
    
    @staticmethod
    def get_latex_postamble():
        """Generate LaTeX document ending."""
        return r"\end{document}"
    
    def compile_latex_robust(self, tex_file):
        """Compile LaTeX with multiple attempts and error handling."""
        tex_basename = os.path.basename(tex_file)
        tex_name = os.path.splitext(tex_basename)[0]
        tex_dir = os.path.dirname(tex_file)
        
        # Common pdflatex arguments
        latex_args = [
            'pdflatex',
            '-interaction=nonstopmode',
            '-halt-on-error',
            f'-output-directory={os.path.abspath(self.aux_folder)}',
            '-synctex=1',
            tex_basename
        ]
        
        # Set environment variables
        env = os.environ.copy()
        env['TEXMFOUTPUT'] = self.aux_folder
        env['TEXMFCACHE'] = self.aux_folder
        
        result = None
        try:
            # Run compilation twice for references
            for _ in range(2):
                result = subprocess.run(
                    latex_args,
                    cwd=tex_dir,
                    capture_output=True,
                    text=False,
                    env=env,
                    timeout=60
                )
            
            # Check if PDF was generated
            generated_pdf = os.path.join(self.aux_folder, f"{tex_name}.pdf")
            if os.path.exists(generated_pdf):
                return generated_pdf
            else:
                self._handle_compilation_error(result, tex_basename)
                return None
                
        except subprocess.TimeoutExpired:
            print(f"LaTeX compilation timed out for {tex_basename}")
            return None
        except Exception as e:
            print(f"LaTeX compilation error for {tex_basename}: {e}")
            return None
    
    def _handle_compilation_error(self, result, tex_basename):
        """Handle LaTeX compilation errors with proper encoding."""
        print(f"LaTeX compilation failed for {tex_basename}")
        
        for output_type, output_data in [("STDOUT", result.stdout), ("STDERR", result.stderr)]:
            if output_data:
                try:
                    text = output_data.decode('utf-8')
                except UnicodeDecodeError:
                    try:
                        text = output_data.decode('latin-1')
                    except UnicodeDecodeError:
                        text = output_data.decode('cp1252', errors='replace')
                
                if text:
                    print(f"{output_type}:", text[-500:])
    
    def read_csv_with_encoding(self, csv_path):
        """Read CSV file with multiple encoding attempts."""
        encodings = ['utf-8', 'latin-1', 'cp1252']
        
        for encoding in encodings:
            try:
                return pd.read_csv(csv_path, encoding=encoding)
            except UnicodeDecodeError:
                continue
        
        raise ValueError(f"Could not read CSV file {csv_path} with any encoding")
    
    def get_row_configurations(self, is_tuesday_thursday):
        """Get row configurations based on day type."""
        if is_tuesday_thursday:
            return [
                (0, "A", "rowCyan"), (1, None, "rowCyan"),
                (2, "C", "rowGreen"), (3, None, "rowGreen"),
                (4, "D", "rowRed")
            ]
        else:
            return [
                (0, "A", "rowCyan"), (1, None, "rowCyan"),
                (2, "B", "rowYellow"),
                (3, "C", "rowGreen"), (4, None, "rowGreen"),
                (5, "D", "rowRed")
            ]
    
    def generate_image_for_item(self, item, sheet_name, row, col_idx):
        """Generate image for a menu item."""
        img_filename = f"{sheet_name}_row{row}_c{col_idx}.png"
        img_path = os.path.join(self.img_folder, img_filename)
        
        # Use unescaped text for image generation
        unescaped_item = self.unescape_latex_text(item)
        ig.generate_image_best(unescaped_item, img_path)
        
        # Return LaTeX-compatible path
        img_rel_path = self.normalize_path_for_latex(
            os.path.relpath(img_path, start=self.tex_folder)
        )
        return f"\\centering\\raisebox{{-0.5\\height}}{{\\includegraphics[width=3.5cm,height=3.5cm,keepaspectratio]{{{img_rel_path}}}}}"
    
    def create_table_row(self, custom_value, pikto_key, img_cell, text_cell, row_color):
        """Create a LaTeX table row."""
        escaped_custom_value = self.escape_latex_text(custom_value)
        
        # Handle pikto image
        pikto_img = ""
        if pikto_key:
            pikto_filename = self.PIKTOS[pikto_key]
            pikto_img = f"\\centering\\raisebox{{-0.5\\height}}{{\\includegraphics[width=3.5cm,height=3.5cm,keepaspectratio]{{{pikto_filename}}}}}"
        
        # Create table row
        color_prefix = f"\\rowcolor{{{row_color}}}\n" if row_color else ""
        return f"{color_prefix}\\rule{{0pt}}{{70pt}}\\centering\\textbf{{{escaped_custom_value}}} & {pikto_img} & {img_cell} & {text_cell} \\\\ \\hline"
    
    def create_latex_table(self, col_name, table_rows, col_idx):
        """Create complete LaTeX table."""
        table_content = "\n".join(table_rows)
        header_color = self.HEADER_COLORS[(col_idx-1) % len(self.HEADER_COLORS)]
        colored_header = f"\\multicolumn{{4}}{{|c|}}{{\\cellcolor{{{header_color}}} \\rule{{0pt}}{{20pt}}\\textbf{{{col_name}}}}} \\\\"
        
        return f"""\\begin{{center}}
\\Large
\\begin{{tabular}}{{|m{{2.5cm}}|m{{3cm}}|m{{3.5cm}}|m{{5.5cm}}|}}
\\hline
{colored_header}
\\hline
{table_content}
\\end{{tabular}}
\\end{{center}}"""
    
    def process_menu_items(self, col_df):
        """Extract and clean menu items from DataFrame column."""
        excluded_items = {
            '-', '*', "Tagesmenü:", "Vegetarisch:", "Salatteller:", 
            "Ausweichmenü:", "Vegi di", "Vegi do"
        }
        
        all_items = []
        for row_idx in range(len(col_df)):
            cell = col_df.iat[row_idx, 0]
            if pd.notna(cell) and cell not in excluded_items:
                clean_cell = str(cell).strip()
                escaped_cell = self.escape_latex_text(clean_cell)
                all_items.append(escaped_cell)
        
        return all_items
    
    def generate_tables(self, custom_column_values=None, custom_column_header="Zeit"):
        """
        Generate LaTeX tables with custom column on the left.
        
        Parameters:
        custom_column_values (dict): Dictionary mapping day names to lists of custom values
        custom_column_header (str): Header text for the custom column
        """
        csv_files = [f for f in os.listdir(self.csv_folder) 
                    if f.endswith(".csv") and f.startswith("K")]
        generated_pdfs = []
        
        for filename in tqdm(csv_files, desc="Menüpläne Konvertieren"):
            csv_path = os.path.join(self.csv_folder, filename)
            sheet_name = os.path.splitext(filename)[0]
            df = self.read_csv_with_encoding(csv_path)
            
            for col_idx in tqdm(range(1, len(df.columns)), 
                              desc=f"Menüpläne kreieren für {sheet_name}", leave=False):
                col_name = df.columns[col_idx]
                col_df = df[[col_name]].copy()
                
                day_name = (self.DAY_NAMES[col_idx - 1] 
                           if col_idx - 1 < len(self.DAY_NAMES) 
                           else f"Tag{col_idx}")
                is_tuesday_thursday = day_name in ["Dienstag", "Donnerstag"]
                
                # Process menu items
                all_items = self.process_menu_items(col_df)
                
                # Skip if no meaningful content
                if not self.has_content([self.unescape_latex_text(item) for item in all_items]):
                    print(f"Skipping {sheet_name}_{day_name} - no content")
                    continue
                
                # Generate table for this day
                pdf_path = self._generate_day_table(
                    sheet_name, day_name, col_name, col_idx, all_items,
                    is_tuesday_thursday, custom_column_values
                )
                
                if pdf_path:
                    generated_pdfs.append(pdf_path)
        
        # Merge PDFs and cleanup
        self._merge_and_cleanup_pdfs(generated_pdfs)
    
    def _generate_day_table(self, sheet_name, day_name, col_name, col_idx, 
                           all_items, is_tuesday_thursday, custom_column_values):
        """Generate LaTeX table for a single day."""
        items_to_show = all_items[:6] + [""] * (6 - len(all_items))
        
        # Get custom values for this day
        if custom_column_values and day_name in custom_column_values:
            custom_values = custom_column_values[day_name]
        else:
            custom_values = (self.DEFAULT_VALUES_TUESDAY_THURSDAY 
                           if is_tuesday_thursday 
                           else self.DEFAULT_VALUES_WEEKDAY)
        
        # Generate table rows
        row_configs = self.get_row_configurations(is_tuesday_thursday)
        table_rows = []
        
        for row, (item_idx, pikto_key, row_color) in enumerate(row_configs):
            if item_idx >= len(items_to_show):
                continue
            
            item = items_to_show[item_idx]
            custom_value = custom_values[row] if row < len(custom_values) else ""
            
            # Generate image if item has content
            img_cell = ""
            if item and item.strip():
                img_cell = self.generate_image_for_item(item, sheet_name, row, col_idx)
            
            # Create table row
            table_row = self.create_table_row(
                custom_value, pikto_key, img_cell, item, row_color
            )
            table_rows.append(table_row)
        
        # Create and compile LaTeX document
        table_latex = self.create_latex_table(col_name, table_rows, col_idx)
        document = self.get_latex_preamble() + "\n\n" + table_latex + "\n\n" + self.get_latex_postamble()
        
        # Write and compile
        tex_file = os.path.join(self.tex_folder, f"{sheet_name}_{day_name}.tex")
        with open(tex_file, 'w', encoding='utf-8', newline='\n') as f:
            f.write(document)
        
        generated_pdf_path = self.compile_latex_robust(tex_file)
        
        if generated_pdf_path and not self.is_pdf_blank(generated_pdf_path):
            final_pdf_path = os.path.join(self.pdf_folder, f"{sheet_name}_{day_name}.pdf")
            shutil.move(generated_pdf_path, final_pdf_path)
            return final_pdf_path
        elif generated_pdf_path:
            print(f"Skipping blank PDF: {sheet_name}_{day_name}")
            os.remove(generated_pdf_path)
        
        return None
    
    def _merge_and_cleanup_pdfs(self, generated_pdfs):
        """Merge generated PDFs and cleanup temporary files."""
        # Filter valid PDFs
        valid_pdfs = []
        for pdf_path in generated_pdfs:
            if os.path.exists(pdf_path) and not self.is_pdf_blank(pdf_path):
                valid_pdfs.append(pdf_path)
            elif os.path.exists(pdf_path):
                os.remove(pdf_path)
        
        if not valid_pdfs:
            return
        
        # Split into two parts and merge
        first_part = valid_pdfs[:5]
        second_part = valid_pdfs[5:]
        
        for part, filename in [(first_part, "Kantine 1.pdf"), (second_part, "Kantine 2.pdf")]:
            if part:
                merger = PdfMerger()
                try:
                    for pdf in part:
                        merger.append(pdf)
                    merger.write(os.path.join(self.pdf_folder, filename))
                finally:
                    merger.close()
        
        # Clean up individual PDFs
        for pdf in valid_pdfs:
            if os.path.exists(pdf):
                os.remove(pdf)
        
        # Clean up temporary directories
        for folder in [self.tex_folder, self.img_folder, "csv_files"]:
            shutil.rmtree(folder, ignore_errors=True)


# Convenience functions for common use cases
def create_tables_with_time_column(csv_folder):
    """Create tables with time values in the left column."""
    time_values = {
        "Montag": ["08:00", "12:00", "14:00", "16:00", "18:00", "20:00"],
        "Dienstag": ["08:00", "12:00", "14:00", "16:00", "18:00"],
        "Mittwoch": ["08:00", "12:00", "14:00", "16:00", "18:00", "20:00"],
        "Donnerstag": ["08:00", "12:00", "14:00", "16:00", "18:00"],
        "Freitag": ["08:00", "12:00", "14:00", "16:00", "18:00", "20:00"],
        "Samstag": ["10:00", "13:00", "15:00", "17:00", "19:00", "21:00"],
        "Sonntag": ["10:00", "13:00", "15:00", "17:00", "19:00", "21:00"]
    }
    
    generator = LaTeXMenuGenerator(csv_folder)
    generator.generate_tables(
        custom_column_values=time_values,
        custom_column_header="Uhrzeit"
    )


def create_tables_with_meal_type_column(csv_folder):
    """Create tables with meal type values in the left column."""
    meal_values = {
        "Montag": ["Frühstück", "Mittagessen", "Zwischenmahlzeit", "Abendessen", "Spätmahlzeit", "Snack"],
        "Dienstag": ["Frühstück", "Mittagessen", "Zwischenmahlzeit", "Abendessen", "Spätmahlzeit"],
        "Mittwoch": ["Frühstück", "Mittagessen", "Zwischenmahlzeit", "Abendessen", "Spätmahlzeit", "Snack"],
        "Donnerstag": ["Frühstück", "Mittagessen", "Zwischenmahlzeit", "Abendessen", "Spätmahlzeit"],
        "Freitag": ["Frühstück", "Mittagessen", "Zwischenmahlzeit", "Abendessen", "Spätmahlzeit", "Snack"],
        "Samstag": ["Brunch", "Mittagessen", "Kaffee", "Abendessen", "Spätmahlzeit", "Snack"],
        "Sonntag": ["Brunch", "Mittagessen", "Kaffee", "Abendessen", "Spätmahlzeit", "Snack"]
    }
    
    generator = LaTeXMenuGenerator(csv_folder)
    generator.generate_tables(
        custom_column_values=meal_values,
        custom_column_header="Mahlzeit"
    )


def create_tables_with_custom_values(csv_folder, custom_values, header_name):
    """Create tables with fully custom values."""
    generator = LaTeXMenuGenerator(csv_folder)
    generator.generate_tables(
        custom_column_values=custom_values,
        custom_column_header=header_name
    )


# Legacy function for backward compatibility
def create_latex_tables_from_folder(csv_folder, tex_folder="output_tex", pdf_folder="Menues", 
                                   img_folder="output_img", aux_folder="log", wrap_document=True, 
                                   custom_column_values=None, custom_column_header="Zeit", **latex_kwargs):
    """Legacy function - use LaTeXMenuGenerator class instead."""
    generator = LaTeXMenuGenerator(csv_folder, tex_folder, pdf_folder, img_folder, aux_folder)
    generator.generate_tables(custom_column_values, custom_column_header)