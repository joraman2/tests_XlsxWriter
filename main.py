#!/usr/bin/env python3.12

from glob import glob
#from rich import print
import csv

from excelpy import generate_excel

def main() -> None:
    '''
    Llegeix dades dels fitxers que hi ha a data_in, les posa en un excel (amb una pestanya per fitxer) i una pestanya extra amb dades reumides  
    '''
    
    dict_aux = { name.split('/')[2].replace('_data.csv',''):name for name in glob('./data_in/*.csv')  }
    #print(dict_aux)

    #Llegir les dades dels csv i posar-les en una llista de diccionaris
    #cada element de la llista serà una pestanya de l'excel

    dict_data : dict = {}

    for nom_cia,path_dades in dict_aux.items():
        dict_data[nom_cia] : list = []
        with open(path_dades, newline='') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                dict_data[nom_cia].append(row)

    
    #print(dict_data)    

    #Generar l'excel
    generate_excel.main(dict_data)
    


if __name__ == "__main__":
    main()

