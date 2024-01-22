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
        sheet['P1'] = "MD5 Binary"
        sheet['Q1'] = "SHA256 Binary"
        
        # Add Data
        for row_index, value in enumerate(data, start = 2):
            md5_hash = Hash.calculate_hash_md5(value)
            sha256_hash = Hash.calculate_sha256(value)
            md5_binary = Hash.hash_to_binary(md5_hash)
            sha256_binary = Hash.hash_to_binary(sha256_hash)
            
            sheet.cell(row = row_index, column = 1, value = value)
            sheet.cell(row = row_index, column = 4, value = md5_hash)
            sheet.cell(row = row_index, column = 8, value = sha256_hash)
            sheet.cell(row = row_index, column = 16, value = md5_binary)
            sheet.cell(row = row_index, column = 17, value = sha256_binary)
        #Save to file    
        workbook.save(file_name)
        
        md5_binary_values = []
        sha256_binary_values = []
class Hash:
    def calculate_hash_md5(input_string):
        md5_hash = hashlib.md5(input_string.encode()).hexdigest()
        return md5_hash
    
    def calculate_sha256(input_string):
        sha256 = hashlib.sha256(input_string.encode()).hexdigest()
        return sha256
        
    def hash_to_binary(hash_value):
        binary_hash = bin(int(hash_value, 16))[2:]
        return binary_hash

def main():
    
    # Converts the integers to binary values and slices the last two value to only show the binary values
    binary_strings = [bin(x)[2:] for x in range(10000000, 10000300)]
    
    # File location
    excel_file = "D:\\binary_values.xlsx"
    
    excel = Excel()
    
    Excel.write_data(excel_file, binary_strings)

if __name__ == "__main__":
    main()
