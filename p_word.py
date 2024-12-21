from docxtpl import DocxTemplate
import openpyxl as op

def create_doc(akt_list):
    """
    Create a Word document based on the provided data.
    
    :param akt_list: List containing document_number and document_date
    :return: None
    """
    document_number, document_date = akt_list
    
    # Load the template
    doc = DocxTemplate("шаблон.docx")
    
    # Prepare context dictionary
    context = {
        'document_number': document_number,
        'document_date': document_date,
    }
    
    # Render the template
    doc.render(context)
    
    # Save the generated document
    doc.save(f"шаблон-{document_number}.docx")

def excel_read(path_file):
    """
    Read data from an Excel file and process it.
    
    :param path_file: Path to the Excel file
    :return: None
    """
    try:
        # Open the workbook
        wb_read = op.load_workbook(filename=path_file, data_only=True)
        
        # Get the first sheet
        sheet_read = wb_read.worksheets[0]
        
        # Initialize variables
        i = 1
        
        while True:
            # Check if there are more rows
            if sheet_read[f"A{i}"].value is None:
                break
            
            # Extract data
            document_number = sheet_read[f"A{i}"].value
            document_date = sheet_read[f"B{i}"].value
            
            # Create a list with extracted data
            akt_list = [document_number, document_date]
            
            # Call create_doc function
            create_doc(akt_list)
            
            # Move to the next row
            i += 1
    
    except Exception as e:
        print(f"An error occurred while processing the Excel file: {str(e)}")

if __name__ == '__main__':
    path_file = 'data.xlsx'
    excel_read(path_file)

