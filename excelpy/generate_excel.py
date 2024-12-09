import xlsxwriter


def main(data: dict) -> None:
    workbook = xlsxwriter.Workbook('./dades_borsa.xlsx',{'strings_to_numbers': True})

    header_cell = workbook.add_format(
        {
            "bg_color": "#449dc9",
            "font_name": "Century",
            "font_size": 11,
            "bold": True
        }
    )

    for sheet_name in data.keys():

        nom_worksheet = f"{sheet_name}_ws"
        nom_worksheet = workbook.add_worksheet(sheet_name)

        row = 1

        for row_data in  data[sheet_name]:
            tracta_dades(row_data,nom_worksheet, row)
            row += 1
        
        #Escrivim els headers
        #escriu_headers(data[sheet_name][0],nom_worksheet)
        col = 0
        keys = [ key for key,val in data[sheet_name][0].items() ]
        for key in keys:
            nom_worksheet.write(0,col, key.capitalize(), header_cell)
            col += 1



    workbook.close()

def tracta_dades(row_data, nom_worksheet, row) :
    col = 0

    keys = [ key for key,val in row_data.items() ]

    for key in keys:
        nom_worksheet.write(row,col,row_data[key])
        col += 1

if __name__ == "__main__":
    main(data)