"""Con estas librerias establecemos conexion entre la API de nuestro proyecto de googleApis y el fichero de drive"""
import gspread   #https://gspread.readthedocs.io/en/latest/user-guide.html 
from oauth2client.service_account import ServiceAccountCredentials
"""Con la siguiente libreria accedemos a las funciones del cambio de formatos de la celda, nos permitirá conoce el formato de las celdas y columnas"""
import gspread_formatting as gsf

class Excel:
    def __init__(self): 
        """En el metodo init establecemos el acceso a la hoja de calculo de drive
        Una vez se configuran los credenciales de acceso se guarda en la variable ws 
        la apertura del fichero"""
            
        scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name ('pruebaExcel.json', scope)
        client = gspread.authorize (creds)
                
        self.ws = client.open('sistemas')
     
        
    def acceder_pestaña(self, nombrePestaña):
        """Con este método se accede a la pestaña que recibe por parametro"""
        
        self.nombre_pestaña = self.ws.worksheet(nombrePestaña) 
        return  self.nombre_pestaña
      
        
    def max_filas_hoja(self):
        """Se obtiene el numero total de filas que contienen informacion"""
        
        self.numFilas = len(self.nombre_pestaña.get_all_values())
        return self.numFilas   
    
    
    def max_columnas_hoja(self):
        """Se obtiene el numero total de columnas que contienen informacion,"""
        
        self.numColumnas = len(self.nombre_pestaña.row_values(1)) #Estraemos la logitud del total de datos que se extraen de una de las filas, la 1 en nuestro caso
        return self.numColumnas 
    
    
    def lectura_columna(self, indiceColumna):
        "Extraemos los valores de una columna especifica, por parametros recibe el indice de la columna"
        
        columna=self.nombre_pestaña.col_values(indiceColumna)
        print("\nEl contenido de la columna "+str(indiceColumna)+" es: ", end=" ")
        for i in columna:
            print(i, end=" ")
      
            
    def lectura_fila(self, numeroFila):
        "Extraemos los valores de una fila especifica, por parametros recibe el numero de la fila"
        
        fila=self.nombre_pestaña.row_values(numeroFila)
        print("\n\nEl contenido de la fila "+str(numeroFila)+" es: ", end=" ")
        for i in fila:
            print(i, end=" ")
    
    
    def escritura_fila(self, datos): 
        """Con este metodo agregamos una fila al final y escribimos los datos que recibe como paramtero en forma de lista
        Con el metodo insert_row() especificamos los datos a insertar, el indice donde se inserta la fila """
        
        print("\nVamos a insertar una fila al final: ",end=" ")
        
        self.nombre_pestaña.insert_row(datos, index = self.numFilas + 1) 
           
        print("Fila insertada")
    
    
    def escritura_columna(self, datos, columna):
        """Con este metodo vamos a cambiar los valores de una columna entera en funcion del nombre de la columna recibido por parámetro"""
        
        print("Vamos a modificar el contenido de la columna: "+columna)
        cont=1
        for i in datos:
            """Se concatena un contador a la letra recibida por parámetro para que la celda a modificar sea diferente cada iteracion del bucle
            Con el bucle recorremos los datos de la lista recibida por parametros, estos datos se iran insertando en la celda"""
            
            celdaColumna=columna+str(cont) #CONCATENAMOS A LA LETRA EL CONTADOR
            self.nombre_pestaña.update(celdaColumna, i) #ACTUALIZAMOS DIRECTAMENTE LA CELDA
            cont+=1
        print("Columna insertada")
        
                # OTRA FORMA DE HACERLO

        """ Con este método vamos a cambiar los valores de una columna entera recibiendo por parametro el nombre de la columna
            Con el método gspread.utils.a1_to_rowcol() obtenemos las coordenadas del nombre de la celda que se recibe --> A1 = (1,1)
            Haremos un bucle para modificar celda da celda de la columna los valores que recibe de la lista llamada datos"""
            
        # columna = columna + "1" # Modificamos para pasar la coordenada H1
        # coordenadaColumna = gspread.utils.a1_to_rowcol(columna) # (1,8) num fila, num columna
        # indiceColumna = coordenadaColumna[1] #Extraemos el segundo valor que es el indice de la columna (8)
        
        # print("Vamos a modificar el contenido de la columna: "+columna)
        # for i in range(1, self.numFilas + 1):
        #     self.nombre_pestaña.update_cell( i,indiceColumna, datos[i-1])   
            
        # print("Columna insertada")
     
        
    def lectura_celda(self, coordenada):
        """Este metodo retorna el valor de la celda pasada por parametro"""
        
        valor_celda = self.nombre_pestaña.acell(coordenada).value
        
        return valor_celda
    
    
    def escritura_celda(self, coordenada, nuevoValor):
        """Este método escribe o sobreescribe contenido sobre una celda especifica"""
        
        self.nombre_pestaña.update(coordenada, nuevoValor)
        print("El nuevo valor de la celda "+coordenada+" es: "+nuevoValor)
      
        
    def formato_celda(self, coordenada):
        """Con este método extraemos el formato de una celda específica, es necesario importar la libreria gspread_formatting, previamente instalada"""
        
        formato_celda = gsf.get_user_entered_format(self.nombre_pestaña, coordenada)
        
        return formato_celda
    
    
    def formato_columna(self, col):
        """Con este método estraemos el formato de las celdas de una columna
        El metodo recibe una letra, se le concatena el 1 para escoger la primera celda de esa columna
        Se extraen los indices de la coordenada
        Se guarda el indice correspondiente a la columna
        """
        cont=1
        numeroFilas=self.numFilas
        while(numeroFilas>0):
            coordenada=col+str(cont)
            formatoCelda=gsf.get_user_entered_format(self.nombre_pestaña, coordenada)
            print(coordenada+" -> ",formatoCelda)
            cont+=1
            numeroFilas-=1
            
              
    def lectura_rango_celdas(self, celdaInicial, celdaFinal):
        """ Con este método se hace la lectura de un rango de celdas, recibe por parametros la celda inicial y la celda final
        Se utiliza el metodo get() al que se le pasa por parametros el rango en formato ('A1:B1'), el segundo parametro "major_dimensions" 
        puede ser COLUMN si queremos que junte los valores de A1 Y B1 en subtuplas --> ((A1,B1), (A2,B2)...), o puede ser el valor ROWS si queremos 
        que junte los valores de las columnas en subtuplas independientes ((A1,A2..)(B1,B2...)...)"""
        
        rangoCeldas = celdaInicial+":"+celdaFinal
        print("\nLos valores del rango de filas y columnas recibido por parametro es: ")
        
        resultado = self.nombre_pestaña.get(rangoCeldas, major_dimension='ROWS') 
        print(resultado)
    
    
    def escritura_rango_celdas(self, celdaInicial, celdaFinal, datos): # B3   C10
        """ Con este método actualizamos los valores recibidos por parametros en un rango de celdas
        Utilizamos el metodo update(), en el que se especifica el rango y los datos que se van a actualizar"""
        
        print("\nVamos a modificar los valores del rango de celdas: "+celdaInicial+":"+celdaFinal)
        
        rangoCeldas = celdaInicial+":"+celdaFinal
        self.nombre_pestaña.update(rangoCeldas, datos)  # Actualizamos los datos con el rango de celdas y la lista de valores recibida por parametros

        print("\n MODIFICACIONES HECHAS EN EL RANGO DE CELDAS") 
  
  
if __name__ == '__main__':  
    Hoja = Excel()
    print("Hemos accedido a la hoja excel del drive")

    pestaña = Hoja.acceder_pestaña("Test Maquinas")
    print("Hemos accedido a la pestaña: "+pestaña.title)

    numFilas = Hoja.max_filas_hoja()
    print("El numero de filas de la pestaña "+ pestaña.title +" es: ",numFilas)
    
    numColumnas = Hoja.max_columnas_hoja()
    print("El numero de columnas de la pestaña "+ pestaña.title +" es: ",numColumnas)
    
    Hoja.lectura_columna(3)
    
    Hoja.lectura_fila(3)

    datosFila = ["A","B","C","D","E","F","G","H","I","J","K","L"]
    # Hoja.escritura_fila(datosFila)
    
    datosColumna = ["colorA","colorB","colorC","colorD","colorE","colorF","colorG","colorH","colorI","colorJ","colorK","colorL","colorM","colorN","colorÑ","colorO","colorP","colorQ","colorR","colorS","colorT","colorU","colorV","colorX"]
    # Hoja.escritura_columna(datosColumna, "H")
    
    coordenada_celda = 'H5'
    valor_celda = Hoja.lectura_celda(coordenada_celda)
    print("\nEl contenido de la celda "+coordenada_celda+" es: "+valor_celda)
    
    Hoja.escritura_celda('E10', 'Hoy es jueves')
    
    formato_celda = Hoja.formato_celda(coordenada_celda)
    print("\nEl formato de la celda "+coordenada_celda+" es: ", formato_celda)
    
    Hoja.formato_columna("G")
    
    Hoja.lectura_rango_celdas('B3', 'C10')
    
    datos=[['LUNES','ESPALDA'],['MARTES','PECHO'],['MIERCOLES','PIERNA'],['JUEVES','HOMBRO'],['VIERNES','BRAZO'],['SABADO','CARDIO'],['DOMINGO','ABDOMEN'],['LUNES','BODY COMBAT']]
    Hoja.escritura_rango_celdas('B3', 'C10', datos)
    
    
    print("\n F I N    D E L    P R O G R A M A")