import pandas as pd 
import PIL 
import xlwings as xw
import os 

def split_excel(multisheet_excel_file, output_path):
    """
        Description: A simple function to wrap pandas methods and split all the sheets present in an 
                     excel file into multiple single sheet files

    	Args:
            - multisheet_excel_filepath: an excel file containing multiple sheets (tabs)
            - output_path: path to output the split excel files
            
        Outputs: 
            - originalfilename_sheetname.xlsx
    """

    file_name = os.path.basename(multisheet_excel_file)
    base_name, _ = os.path.splitext(file_name)

    excel_in = pd.ExcelFile(multisheet_excel_file)
    sheets = excel_in.sheet_names

    for sheet in sheets:
        df = pd.read_excel(excel_in, sheet_name=sheet)  
        df.to_excel(f'{output_path}/{base_name}_{sheet}.xlsx', index=False)


def split_excel_xlwings(multisheet_excel_file, output_path):
    """
        Description: A simple function to split all the sheets present in an 
                     excel file into multiple single sheet files, using the xlwings backend 
                     for more complex file (images, color data, etc)

    	Args:
            - multisheet_excel_filepath: an excel file containing multiple sheets (tabs)
            - output_path: path to output the split excel files
            
        Outputs: 
            - originalfilename_sheetname.xlsx
    """

    file_name = os.path.basename(multisheet_excel_file)
    base_name, _ = os.path.splitext(file_name)

    try: 
        excel_app = xw.App(visible=False)
        wb = excel_app.books.open(multisheet_excel_file)

        for sheet in wb.sheets:
            wb_new = xw.Book()
            sheet.copy(after=wb_new.sheets[0])
            wb_new.sheets[0].delete()
            wb_new.save(f'{output_path}/{base_name}_{sheet.name}.xlsx')
            wb_new.close()
    finally:
        excel_app.quit()        