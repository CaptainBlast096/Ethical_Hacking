import openpyxl # For Excel
import hashlib # Hashign algorithms
class Excel:
    def write_data(file_name, data):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        
        # Merging cells
        sheet.merge_cells('A1:C1')
        
        # Add Header
        sheet['A1'] = "Binary Value"
        
        # Add Data
        for row_index, value in enumerate(data, start = 2):
            sheet.cell(row = row_index, column = 1, value = value)
            
        #Save to file    
        workbook.save(file_name)
        
def main():
    
    # Converts the integers to binary values and slices the last two value to only show the binary values
    binary_strings = [bin(x)[2:] for x in range(10000000, 10000300)]
    
    # File location
    excel_file = "D:\\binary_values.xlsx"
    
    excel = Excel()
    Excel.write_data(excel_file, binary_strings)
    
if __name__ == "__main__":
    main()
