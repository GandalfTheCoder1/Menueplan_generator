import pandas as pd
import os
import subprocess
import shutil
from tqdm import tqdm
import image_gen as ig
from PyPDF2 import PdfMerger, PdfReader
import platform
import tempfile

def has_content(items_list):
    """Check if the items list has any meaningful content"""
    return any(item.strip() for item in items_list if item and str(item).strip() not in ['-', '*', '', 'nan'])

def is_pdf_blank(pdf_path):
    """Check if a PDF is effectively blank by examining its content"""
    try:
        reader = PdfReader(pdf_path)
        if len(reader.pages) == 0:
            return True
        
        # Check if all pages are essentially empty
        for page in reader.pages:
            text = page.extract_text().strip()
            # If page has significant text content, it's not blank
            if len(text) > 50:  # Adjust threshold as needed
                return False
        return True
    except:
        return True

def normalize_path_for_latex(path):
    """Convert path to forward slashes for LaTeX compatibility across platforms"""
    return path.replace(os.sep, '/')

def get_latex_preamble():
    """Generate a robust LaTeX preamble that works across platforms"""
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

def get_latex_postamble():
    """Generate LaTeX document ending"""
    return r"\end{document}"

def compile_latex_robust(tex_file, output_dir, aux_dir):
    """Compile LaTeX with multiple attempts and error handling"""
    tex_basename = os.path.basename(tex_file)
    tex_name = os.path.splitext(tex_basename)[0]
    
    # Ensure directories exist
    os.makedirs(aux_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)
    
    # Get the directory containing the tex file
    tex_dir = os.path.dirname(tex_file)
    
    # Common pdflatex arguments for consistency
    latex_args = [
        'pdflatex',
        '-interaction=nonstopmode',
        '-halt-on-error',
        f'-output-directory={os.path.abspath(aux_dir)}',
        '-synctex=1',
        tex_basename
    ]
    
    # Set environment variables for consistent behavior
    env = os.environ.copy()
    env['TEXMFOUTPUT'] = aux_dir
    env['TEXMFCACHE'] = aux_dir
    
    try:
        # First compilation - handle encoding issues
        result1 = subprocess.run(
            latex_args,
            cwd=tex_dir,
            capture_output=True,
            text=False,  # Get bytes instead of text
            env=env,
            timeout=60  # Prevent hanging
        )
        
        # Second compilation for references (if needed)
        result2 = subprocess.run(
            latex_args,
            cwd=tex_dir,
            capture_output=True,
            text=False,  # Get bytes instead of text
            env=env,
            timeout=60
        )
        
        # Check if PDF was generated
        generated_pdf = os.path.join(aux_dir, f"{tex_name}.pdf")
        if os.path.exists(generated_pdf):
            return generated_pdf
        else:
            print(f"LaTeX compilation failed for {tex_basename}")
            
            # Safely decode output for error reporting
            try:
                stdout_text = result2.stdout.decode('utf-8')
            except UnicodeDecodeError:
                try:
                    stdout_text = result2.stdout.decode('latin-1')
                except UnicodeDecodeError:
                    stdout_text = result2.stdout.decode('cp1252', errors='replace')
            
            try:
                stderr_text = result2.stderr.decode('utf-8')
            except UnicodeDecodeError:
                try:
                    stderr_text = result2.stderr.decode('latin-1')
                except UnicodeDecodeError:
                    stderr_text = result2.stderr.decode('cp1252', errors='replace')
            
            if stdout_text:
                print("STDOUT:", stdout_text[-500:])  # Last 500 chars
            if stderr_text:
                print("STDERR:", stderr_text[-500:])
            return None
            
    except subprocess.TimeoutExpired:
        print(f"LaTeX compilation timed out for {tex_basename}")
        return None
    except Exception as e:
        print(f"LaTeX compilation error for {tex_basename}: {e}")
        return None

def create_latex_tables_from_folder(csv_folder, tex_folder="output_tex", pdf_folder="Menues", 
                                   img_folder="output_img", aux_folder="log", wrap_document=True, 
                                   custom_column_values=None, custom_column_header="Zeit", **latex_kwargs):
    """
    Generate LaTeX tables with custom column on the left.
    
    Parameters:
    custom_column_values (dict): Dictionary mapping day names to lists of custom values
                                Example: {"Montag": ["08:00", "12:00", "14:00", "16:00", "18:00", "20:00"],
                                         "Dienstag": ["08:00", "12:00", "14:00", "16:00", "18:00"]}
    custom_column_header (str): Header text for the custom column (default: "Zeit")
    """
    
    # Default custom values if none provided
    default_values_weekday = ["", "", "T", "E", "S", ""]  # 6 rows
    default_values_tuesday_thursday = ["", "", "E", "S", "A"]  # 5 rows
    
    # Create directories
    os.makedirs(tex_folder, exist_ok=True)
    os.makedirs(pdf_folder, exist_ok=True)
    os.makedirs(img_folder, exist_ok=True)
    os.makedirs(aux_folder, exist_ok=True)

    # Copy pikto images with normalized paths
    all_piktos = {"A": "A.jpg", "B": "B.jpg", "C": "C.jpg", "D": "D.jpg"}
    for pikto_file in all_piktos.values():
        src = os.path.join("Piktos", pikto_file)
        if os.path.exists(src):
            dst = os.path.join(tex_folder, pikto_file)
            shutil.copy2(src, dst)  # copy2 preserves metadata

    csv_files = [f for f in os.listdir(csv_folder) if f.endswith(".csv") and f.startswith("K")]
    day_names = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag"]
    
    # Use consistent color names
    header_colors = [
        'headerYellow', 'headerGreen', 'headerBlue', 'headerRed', 
        'headerOrange', 'headerWhite', 'headerPink'
    ]

    generated_pdfs = []

    for filename in tqdm(csv_files, desc="Menüpläne Konvertieren"):
        csv_path = os.path.join(csv_folder, filename)
        sheet_name = os.path.splitext(filename)[0]

        # Read CSV with explicit encoding
        try:
            df = pd.read_csv(csv_path, encoding='utf-8')
        except UnicodeDecodeError:
            try:
                df = pd.read_csv(csv_path, encoding='latin-1')
            except UnicodeDecodeError:
                df = pd.read_csv(csv_path, encoding='cp1252')

        for col_idx in tqdm(range(1, len(df.columns)), desc=f"Menüpläne kreieren für {sheet_name}", leave=False):
            col_name = df.columns[col_idx]
            col_df = df[[col_name]].copy()

            day_name = day_names[col_idx - 1] if col_idx - 1 < len(day_names) else f"Tag{col_idx}"
            is_tuesday_thursday = day_name in ["Dienstag", "Donnerstag"]

            all_items = []
            for row_idx in range(len(col_df)):
                cell = col_df.iat[row_idx, 0]
                if pd.notna(cell) and cell not in ['-', '*', "Tagesmenü:", "Vegetarisch:", "Salatteller:", "Ausweichmenü:", "Vegi di", "Vegi do"]:
                    # Clean and escape LaTeX special characters
                    clean_cell = str(cell).strip()
                    # Escape LaTeX special characters
                    latex_special_chars = {
                        '&': r'\&',
                        '%': r'\%',
                        '$': r'\$',
                        '#': r'\#',
                        '_': r'\_',
                        '{': r'\{',
                        '}': r'\}',
                        '~': r'\textasciitilde{}',
                        '^': r'\textasciicircum{}',
                        '\\': r'\textbackslash{}'
                    }
                    for char, escape in latex_special_chars.items():
                        clean_cell = clean_cell.replace(char, escape)
                    all_items.append(clean_cell)

            # Skip this day if there's no meaningful content
            if not has_content([item.replace('\\', '') for item in all_items]):  # Check unescaped version
                print(f"Skipping {sheet_name}_{day_name} - no content")
                continue

            items_to_show = all_items[:6] + [""] * (6 - len(all_items))
            table_rows = []

            # Get custom values for this day
            if custom_column_values and day_name in custom_column_values:
                custom_values = custom_column_values[day_name]
            else:
                # Use default values based on day type
                if is_tuesday_thursday:
                    custom_values = default_values_tuesday_thursday
                else:
                    custom_values = default_values_weekday

            if is_tuesday_thursday:
                row_configs = [
                    (0, "A", "rowCyan"), (1, None, "rowCyan"),
                    (2, "C", "rowGreen"), (3, None, "rowGreen"),
                    (4, "D", "rowRed")
                ]
                max_rows = 5
            else:
                row_configs = [
                    (0, "A", "rowCyan"), (1, None, "rowCyan"),
                    (2, "B", "rowYellow"),
                    (3, "C", "rowGreen"), (4, None, "rowGreen"),
                    (5, "D", "rowRed")
                ]
                max_rows = 6

            for row, (item_idx, pikto_key, row_color) in enumerate(row_configs):
                if item_idx >= len(items_to_show):
                    continue
                    
                item = items_to_show[item_idx]
                img_cell = ""
                text_cell = ""

                # Get custom value for this row
                custom_value = custom_values[row] if row < len(custom_values) else ""
                
                # Escape custom value for LaTeX
                escaped_custom_value = custom_value
                latex_special_chars = {
                    '&': r'\&', '%': r'\%', '$': r'\$', '#': r'\#', '_': r'\_',
                    '{': r'\{', '}': r'\}', '~': r'\textasciitilde{}',
                    '^': r'\textasciicircum{}', '\\': r'\textbackslash{}'
                }
                for char, escape in latex_special_chars.items():
                    escaped_custom_value = escaped_custom_value.replace(char, escape)

                if item and item.strip():
                    # Generate image
                    img_filename = f"{sheet_name}_row{row}_c{col_idx}.png"
                    img_path = os.path.join(img_folder, img_filename)
                    
                    # Use unescaped text for image generation
                    unescaped_item = item
                    for escape, char in [('\\&', '&'), ('\\%', '%'), ('\\$', '$'), 
                                       ('\\#', '#'), ('\\_', '_'), ('\\{', '{'), 
                                       ('\\}', '}'), ('\\textasciitilde{}', '~'), 
                                       ('\\textasciicircum{}', '^'), ('\\textbackslash{}', '\\')]:
                        unescaped_item = unescaped_item.replace(escape, char)
                    
                    ig.generate_image_best(unescaped_item, img_path)
                    
                    # Use forward slashes for LaTeX path
                    img_rel_path = normalize_path_for_latex(os.path.relpath(img_path, start=tex_folder))
                    img_cell = f"\\centering\\raisebox{{-0.5\\height}}{{\\includegraphics[width=3.5cm,height=3.5cm,keepaspectratio]{{{img_rel_path}}}}}"
                    text_cell = item

                # Handle pikto image
                pikto_img = ""
                if pikto_key:
                    pikto_filename = all_piktos[pikto_key]
                    pikto_img = f"\\centering\\raisebox{{-0.5\\height}}{{\\includegraphics[width=3.5cm,height=3.5cm,keepaspectratio]{{{pikto_filename}}}}}"

                # Create table row with custom column (now 4 columns total) - with explicit row height
                color_prefix = f"\\rowcolor{{{row_color}}}\n" if row_color else ""
                row_latex = f"{color_prefix}\\rule{{0pt}}{{70pt}}\\centering\\textbf{{{escaped_custom_value}}} & {pikto_img} & {img_cell} & {text_cell} \\\\ \\hline"
                table_rows.append(row_latex)

            table_content = "\n".join(table_rows)
            
            # Use consistent header coloring (now spans 4 columns) - with fixed height
            header_color = header_colors[(col_idx-1) % len(header_colors)]
            colored_header = f"\\multicolumn{{4}}{{|c|}}{{\\cellcolor{{{header_color}}} \\rule{{0pt}}{{20pt}}\\textbf{{{col_name}}}}} \\\\"

            # Create table with consistent formatting (now 4 columns)
            table_latex = f"""\\begin{{center}}
\\Large
\\begin{{tabular}}{{|m{{2.5cm}}|m{{3cm}}|m{{3.5cm}}|m{{5.5cm}}|}}
\\hline
{colored_header}
\\hline
{table_content}
\\end{{tabular}}
\\end{{center}}"""

            # Write LaTeX file
            col_tex_file = os.path.join(tex_folder, f"{sheet_name}_{day_name}.tex")

            if wrap_document:
                document = get_latex_preamble() + "\n\n" + table_latex + "\n\n" + get_latex_postamble()
            else:
                document = table_latex

            # Write with explicit UTF-8 encoding and Unix line endings
            with open(col_tex_file, 'w', encoding='utf-8', newline='\n') as f:
                f.write(document)

            # Compile LaTeX
            generated_pdf_path = compile_latex_robust(col_tex_file, pdf_folder, aux_folder)
            
            if generated_pdf_path and not is_pdf_blank(generated_pdf_path):
                final_pdf_path = os.path.join(pdf_folder, f"{sheet_name}_{day_name}.pdf")
                shutil.move(generated_pdf_path, final_pdf_path)
                generated_pdfs.append(final_pdf_path)
            elif generated_pdf_path:
                print(f"Skipping blank PDF: {sheet_name}_{day_name}")
                os.remove(generated_pdf_path)

    # Filter valid PDFs and merge
    valid_pdfs = []
    for pdf_path in generated_pdfs:
        if os.path.exists(pdf_path) and not is_pdf_blank(pdf_path):
            valid_pdfs.append(pdf_path)
        elif os.path.exists(pdf_path):
            os.remove(pdf_path)

    if valid_pdfs:
        
        first_part = valid_pdfs[:5]
        second_part = valid_pdfs[5:]

        if first_part:
            merger1 = PdfMerger()
            try:
                for pdf in first_part:
                    merger1.append(pdf)
                merger1.write(os.path.join(pdf_folder, "Kantine 1.pdf"))
            finally:
                merger1.close()

        if second_part:
            merger2 = PdfMerger()
            try:
                for pdf in second_part:
                    merger2.append(pdf)
                merger2.write(os.path.join(pdf_folder, "Kantine 2.pdf"))
            finally:
                merger2.close()

        # Clean up individual PDFs
        for pdf in valid_pdfs:
            if os.path.exists(pdf):
                os.remove(pdf)

    # Clean up temporary directories
    shutil.rmtree(tex_folder, ignore_errors=True)
    shutil.rmtree(img_folder, ignore_errors=True)
    shutil.rmtree("csv_files", ignore_errors=True)

# Example usage functions:

def create_tables_with_time_column(csv_folder):
    """Create tables with time values in the left column"""
    time_values = {
        "Montag": ["08:00", "12:00", "14:00", "16:00", "18:00", "20:00"],
        "Dienstag": ["08:00", "12:00", "14:00", "16:00", "18:00"],
        "Mittwoch": ["08:00", "12:00", "14:00", "16:00", "18:00", "20:00"],
        "Donnerstag": ["08:00", "12:00", "14:00", "16:00", "18:00"],
        "Freitag": ["08:00", "12:00", "14:00", "16:00", "18:00", "20:00"],
        "Samstag": ["10:00", "13:00", "15:00", "17:00", "19:00", "21:00"],
        "Sonntag": ["10:00", "13:00", "15:00", "17:00", "19:00", "21:00"]
    }
    
    create_latex_tables_from_folder(
        csv_folder, 
        custom_column_values=time_values, 
        custom_column_header="Uhrzeit"
    )

def create_tables_with_meal_type_column(csv_folder):
    """Create tables with meal type values in the left column"""
    meal_values = {
        "Montag": ["Frühstück", "Mittagessen", "Zwischenmahlzeit", "Abendessen", "Spätmahlzeit", "Snack"],
        "Dienstag": ["Frühstück", "Mittagessen", "Zwischenmahlzeit", "Abendessen", "Spätmahlzeit"],
        "Mittwoch": ["Frühstück", "Mittagessen", "Zwischenmahlzeit", "Abendessen", "Spätmahlzeit", "Snack"],
        "Donnerstag": ["Frühstück", "Mittagessen", "Zwischenmahlzeit", "Abendessen", "Spätmahlzeit"],
        "Freitag": ["Frühstück", "Mittagessen", "Zwischenmahlzeit", "Abendessen", "Spätmahlzeit", "Snack"],
        "Samstag": ["Brunch", "Mittagessen", "Kaffee", "Abendessen", "Spätmahlzeit", "Snack"],
        "Sonntag": ["Brunch", "Mittagessen", "Kaffee", "Abendessen", "Spätmahlzeit", "Snack"]
    }
    
    create_latex_tables_from_folder(
        csv_folder, 
        custom_column_values=meal_values, 
        custom_column_header="Mahlzeit"
    )

def create_tables_with_custom_values(csv_folder, custom_values, header_name):
    """Create tables with fully custom values"""
    create_latex_tables_from_folder(
        csv_folder, 
        custom_column_values=custom_values, 
        custom_column_header=header_name
    )