import create_table as ct
import excel_to_csv as ec

if __name__ == "__main__":
    ec.convert_excel_to_csv()

    csv_folder = "csv_files"
    tex_folder = "output_tex"
    pdf_folder = "output_pdf"

    ct.create_latex_tables_from_folder(
        csv_folder,
        tex_folder=tex_folder,
        pdf_folder=pdf_folder,
        wrap_document=True,  # Wrap in full LaTeX document
        caption="Generated Table",
        label="tab:generated"
    )

