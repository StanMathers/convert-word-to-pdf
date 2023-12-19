### Convert docx to pdf

This is a simple python script from upwork project. It converts docx to pdf

It converts Word docs to PDFs

Following things that are important to note:

1. If word docs have table of contents, those will be be refreshed prior to export

2. Formatting of different components and images remains intact

3. Page orientations are maintaned

4. All hyperlinks ramain functional

### Run

```bash
git clone <repository_url>
cd <repository_name>
pip install -r requirements.txt
python convert.py
```

### Selecting a word document

Initial project structure didn't require input fields that's why word document must be editen in **convert.py**

Change the following lines in convert.py

```python
input_docx_file = os.getcwd() + "/index.docx"
output_pdf_file = os.getcwd() + "/index.pdf"
```

Replace **index.docx** with desired file name and replace **index.pdf** with desired output file name and run the script again.
