# fillablePDFextract
Extracts fillable PDF data form a table into an Excel file
Python-
To use you will need a PDF path and an .xlsx output path.  
This will pull the fillable data from a table within a PDF. I then iterates through each cell.  If the cell contents meet the correct conditions it will add the contents a worksheet. 
Libraries needed:
  pip install pdfminer.six
  pip install openpyxl
  
