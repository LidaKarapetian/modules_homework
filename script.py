#!/usr/bin/python

import xlsxwriter
import argparse

# Function to read the content from data.txt file
def read_content(input_fname):
    try:
        with open(input_fname, "r") as file:
            return file.readlines()
    
    except FileNotFoundError:
        print(f"Input file '{input_fname}' not found")
        return None

def validate_data(data):
    if data is None:
        return False
    return True


# Function tp sanitize the data      
def sanitize_data(ml):
	return [el.strip().split() for el in ml]

# Function to write all data to exsel file
def txt_to_xlsx(data, xlsx_fname):
     workbook = xlsxwriter.Workbook(xlsx_fname)
     worksheet = workbook.add_worksheet()

     header_format = workbook.add_format({'bold' : True, 'bg_color' : 'yellow'})
     green_format = workbook.add_format({'bg_color' : 'green'})

     for row, line in enumerate(data):
          for col, cell_data in enumerate(line):
               if row == 0:
                    worksheet.write(row, col, cell_data, header_format)
               elif col == 2  and int(cell_data) > 20:
                    worksheet.write(row, col, cell_data)
                    worksheet.set_row(row, None, green_format)
               else:
                    worksheet.write(row,col,cell_data)
             
     workbook.close()

# Function to sort colums separately
def sort_by_criteria(data, option):
     if option == 'n':
        sorted_datas = sorted(data[1:], key=lambda x: x[0])
        sorted_column = [row[0] for row in sorted_datas]
        return sorted_column
     if option == 's':
        sorted_datas = sorted(data[1:], key=lambda x: x[1])
        sorted_column = [row[1] for row in sorted_datas]
        return sorted_column
     if option == 'a':
        sorted_datas = sorted(data[1:], key=lambda x: x[2])
        sorted_column = [row[2] for row in sorted_datas]
        return sorted_column
     if option == 'p':
        sorted_datas = sorted(data[1:], key=lambda x: x[3])
        sorted_column = [row[3] for row in sorted_datas]
        return sorted_column

# Function to write the sorted criteria data to a exsel file
def xlsx_for_sorted_data(sorted_column, xlsx_fname):
    workbook = xlsxwriter.Workbook(xlsx_fname)
    worksheet = workbook.add_worksheet()
    for row, value in enumerate(sorted_column):
        worksheet.write(row, 0, value)

    workbook.close()

def main():

    parser = argparse.ArgumentParser()

    parser.add_argument('-f', '--input', required=True, help='Input file name')
    parser.add_argument('-o', '--output', required=True, help='Output xlsx file name')
    parser.add_argument("-t", "--third-option", choices=["n", "s", "a", "p"], help="Third option (n=Name, s=Surname, a=Age, p=Profession)")
    args = parser.parse_args()

    data = read_content(args.input)
    if not validate_data(data):
        return
    
    xlsx_fname = args.output
    
    san_data = sanitize_data(data)

    sorted_column = sort_by_criteria(san_data, args.third_option)

    if args.third_option:
        xlsx_for_sorted_data(sorted_column, xlsx_fname)
    else:
        txt_to_xlsx(data, xlsx_fname)
  

if __name__ == "__main__":
    main()


