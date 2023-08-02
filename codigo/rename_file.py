import os
import time
import shutil

def rename_file(folder_path_input, folder_path_output):
    print('Entrando en la funcion rename...')
    print('----------------------------------------------------------------------')
    
    # Variable array
    file_name_list = []
    
    # Bucle para obtener lista de nombre de archivos
    for add_file_list in os.listdir(folder_path_input):
        if add_file_list.endswith(".pdf"):
            file_name_list.append(add_file_list)
    
    print('Cantidad Elem. file_name_list: ', len(file_name_list))
    print('----------------------------------------------------------------------')
    
    # Ordenar lista de archivos por nombre
    new_file_name_list_sort = sorted(file_name_list)
    print('file_name_list_sort: ', new_file_name_list_sort, len(new_file_name_list_sort))
    print('----------------------------------------------------------------------')
    
    # Contador de archivos
    file_count = 0
    
    # Recorrer lista con cada archivo, abrir y extraer numero factura
    for x in range(0, len(new_file_name_list_sort)):
        file_count += 1
        input_file = folder_path_input + new_file_name_list_sort[x]
        print('Archivo PDF', input_file, file_count)
        print('----------------------------------------------------------------------')
        time.sleep(2)

        # Mover a la carpeta output con el nuevo nombre
        source = input_file
        dest = folder_path_output + new_file_name_list_sort[x]
        shutil.copy(source, dest)
        print('Copiando archivo a nuevo destino: ', source, dest)
        print('--------------------------------------------------------------------------')        
                
                
if __name__ == '__main__':
    
    # Obtener en una lista todos los archivos 
    FOLDER_PATH_INPUT = '../input/'
    FOLDER_PATH_OUTPUT = '../output/'
    FOLDER_PATH_CONFIG = '../config/'
    
    rename_file(folder_path_input=FOLDER_PATH_INPUT, 
                folder_path_output=FOLDER_PATH_OUTPUT
                )



