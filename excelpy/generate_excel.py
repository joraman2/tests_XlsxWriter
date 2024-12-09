import xlsxwriter


def main(data: dict) -> None:
    workbook = xlsxwriter.Workbook('./dades_borsa.xlsx',{'strings_to_numbers': True})

    for companyia in data.keys():
        print(companyia)

        #print(data[companyia])

        nom_worksheet = f"{companyia}_ws"
        nom_worksheet = workbook.add_worksheet(companyia)

        row = 1

        for row_data in  data[companyia]:
            tracta_dades(row_data,nom_worksheet, row)
            row += 1
        
        #Escrivim els headers
        escriu_headers(data[companyia][0],nom_worksheet)

    workbook.close()

def tracta_dades(row_data, nom_worksheet, row) :
    col = 0

    keys = [ key for key,val in row_data.items() ]

    for key in keys:
        nom_worksheet.write(row,col,row_data[key])
        col += 1

def escriu_headers(row_data,nom_worksheet):

    col = 0
    keys = [ key for key,val in row_data.items() ]

    for key in keys:
        nom_worksheet.write(0,col,key)
        col += 1


if __name__ == "__main__":
    main(data)