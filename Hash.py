import openpyxl
import hashlib

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
        sheet['P1'] = "MD5 Decimal"
        sheet['Q1'] = "SHA256 Decimal"

        # Add Data
        for row_index, value in enumerate(data, start=2):
            md5_hash = Hash.calculate_hash_md5(value)
            sha256_hash = Hash.calculate_sha256(value)
            md5_decimal = Hash.hash_to_decimal(md5_hash)
            sha256_decimal = Hash.hash_to_decimal(sha256_hash)

            sheet.cell(row=row_index, column=1, value=value)
            sheet.cell(row=row_index, column=4, value=md5_hash)
            sheet.cell(row=row_index, column=8, value=sha256_hash)
            sheet.cell(row=row_index, column=16, value=md5_decimal)
            sheet.cell(row=row_index, column=17, value=sha256_decimal)

        # Save to file
        workbook.save(file_name)

class Hash:
    def calculate_hash_md5(input_string):
        md5_hash = hashlib.md5(input_string.encode()).hexdigest()
        return md5_hash

    def calculate_sha256(input_string):
        sha256 = hashlib.sha256(input_string.encode()).hexdigest()
        return sha256

    def hash_to_decimal(hash_value):
        decimal_value = int(hash_value, 16)
        return decimal_value

class Calculator:
    def calculate_average(data):
        average = sum(data) / len(data)
        return average
    
    def calculate_standard_deviation(data):
        #Place Holder
        return standard_deviation

    def calculate_difference(data):
        #Place Holder
        return difference
def main():
    # Converts the integers to binary values and slices the last two values to only show the binary values
    binary_strings = [bin(x)[2:] for x in range(10000000, 10000300)]

    # File location
    excel_file = "D:\\binary_values.xlsx"

    excel = Excel()
    Excel.write_data(excel_file, binary_strings)

if __name__ == "__main__":
    main()
