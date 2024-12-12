import xlsxwriter
from itertools import chain


def main(data: dict) -> None:
    '''
    A partir d'un diccionari que conté a cada clau (pestanya del document de sortida) una llista de diccionaris (dades a escriure en cada pestanya del document)
    genera un fitxer excel
    '''

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

        num_row = 1

        #Ens assegurem que totes les files tinguin les mateixes columnes. Si no està definida --> None
        keys = set(chain.from_iterable(data[sheet_name]))
        for row_data in data[sheet_name]:
            row_data_aux = {}   
            row_data_aux = { key:(row_data[key] if key in row_data.keys() else None ) for key in keys }

            escriu_dades(row_data_aux,nom_worksheet,num_row, keys)
            num_row += 1
        
        #Escrivim els headers
        #escriu_headers(data[sheet_name][0],nom_worksheet)
        col = 0

        for key in keys:
            nom_worksheet.write(0,col, key.capitalize(), header_cell)
            col += 1



    workbook.close()

def escriu_dades(row_data, nom_worksheet, num_row,keys) :
    num_col = 0

    for key in keys:
        nom_worksheet.write(num_row,num_col,row_data[key])
        num_col += 1

if __name__ == "__main__":
    main(data)