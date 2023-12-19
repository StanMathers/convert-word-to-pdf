import os
import comtypes.client

def convert_to_pdf(input_path, output_path):
    if not os.path.exists(input_path):
        print("Input file not found.")
        return

    # Create a Word application object
    word = comtypes.client.CreateObject("Word.Application")
    word.Visible = False  # Set to True if you want to see the Word application

    try:
        doc = word.Documents.Open(input_path)

        # Refresh the table of contents if it exists
        if doc.TablesOfContents.Count > 0:
            for toc in doc.TablesOfContents:
                toc.Update()

        doc.SaveAs(output_path, FileFormat=17)  # 17 corresponds to PDF format
        doc.Close()

        print(f"Conversion successful. PDF saved at {output_path}")
    except Exception as e:
        print(f"Conversion failed: {str(e)}")
    finally:
        word.Quit()

# For windows, especially when using pycharm, it looks for files in sys32.
# For this reason, os library is used to get the current working directory.

# Example usage:
input_docx_file = os.getcwd() + "/index.docx"
output_pdf_file = os.getcwd() + "/index.pdf"

convert_to_pdf(input_docx_file, output_pdf_file)
