import openpyxl # For Excel
import hashlib # Hashign algorithms
class Excel:
    def write_data(file_name, data):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        
        # Merging cells
        sheet.merge_cells('A1:C1')
        sheet.merge_cells('D1:G1')
        sheet.merge_cells('H1:O1')
        # Add Header
        sheet['A1'] = "Binary Value"
        sheet['D1'] = "MD5"
        sheet['H1'] = "SHA256"
        
        # Add Data
        for row_index, value in enumerate(data, start = 2):
            md5_hash = Hash.calculate_hash_md5(value)
            sha256_hash = Hash.calculate_sha256(value)
            
            sheet.cell(row = row_index, column = 1, value = value)
            sheet.cell(row = row_index, column = 4, value = md5_hash)
            sheet.cell(row = row_index, column = 8, value = sha256_hash)
            
        #Save to file    
        workbook.save(file_name)
        
class Hash:
    def calculate_hash_md5(input_string):
        md5_hash = hashlib.md5(input_string.encode()).hexdigest()
        return md5_hash
    
    def calculate_sha256(input_string):
        sha256 = hashlib.sha256(input_string.encode()).hexdigest()
        return sha256
        
def main():
    
    # Converts the integers to binary values and slices the last two value to only show the binary values
    binary_strings = [bin(x)[2:] for x in range(10000000, 10000300)]
    
    # File location
    excel_file = "D:\\binary_values.xlsx"
    
    excel = Excel()
    Excel.write_data(excel_file, binary_strings)
    
if __name__ == "__main__":
    main()
