# crear excel, rutas y rutas, abrir 
import xlsxwriter, os, shutil, webbrowser 

# elegir ruta PyQT
from PySide2.QtWidgets import QFileDialog

# convertir excel a pdf
from win32com import client
import win32api

class ClaseManipularPdf():

    def convertirdor_pdf(self, formato, ingreso, salida):

        '''
        formato : orientacion de hoja
        ingreso : ruta excel
        salida : ruta pdf
        '''

        app = client.DispatchEx("Excel.Application")
        app.Interactive = False
        app.Visible = False
        app.DisplayAlerts = False

        Workbook = app.Workbooks.Open(ingreso)

        # ubicar pagina en vertical=1 / horizontal=2
        ws_source = Workbook.Worksheets("Sheet1")    
        ws_source.PageSetup.Orientation = formato
        ws_source.Select()

        try:
            Workbook.ActiveSheet.ExportAsFixedFormat(0,salida)
        except Exception as error:
            print ('No se pudo convertir en formato PDF. Confirme que el entorno cumple ' 
                    'con todos los requisitos y vuelva a intentarlo')
            print(error)

        # salida
        Workbook.Close()
        app.Application.Quit()
        app.Quit()

class ClaseGeneradorExcel():

    def __init__(self):
        """inicializar con las rutas establecidas"""
        
        # RUTA ACTUAL
        raiz = os.getcwd()

        # RUTA IMAGEN
        'ubicacion de la imagen'
        self.ruta_icon = raiz+r'\icon\icono.jpg'

        # DEFINIR RUTA ARCHIVOS
        'necesitamos una ruta de salida y entrada para el excel y pdf'

        # uno
        self.rep_uno_ex = raiz+r'\aux_excel\reporte_uno.xlsx' # entrada excel
        self.rep_uno_pd = raiz+r'\aux_excel\reporte_uno.pdf' # salida pdf

        # dos
        self.rep_dos_ex = raiz+r'\aux_excel\reporte_dos.xlsx' # entrada excel
        self.rep_dos_pd = raiz+r'\aux_excel\reporte_dos.pdf' # salida pdf

        # CONVERTIR PDF
        self.raiz_manip_pdf = ClaseManipularPdf()

    def guardarArchivo(self,ruta_pdf):
        """guardado del archivo"""

        # obtener ruta de guardado
        ruta = QFileDialog.getSaveFileName(None, 'Seleccionar archivo','','Texto (*.pdf)')
        
        if(ruta[0]!=''):
            # mover archivo pdf y cambiar nombre
            shutil.move(ruta_pdf,ruta[0])

            # abrir pdf    
            webbrowser.open(ruta[0], new=2)

        else:
            print('No se guardo archivo.')


    def gen_estilos(self,obj):
        """estos estilos son aplicados a los excel (solo si comparten mismos estilos)"""

        # ESTILOS ------------------------------------- 
            
        self.s_header_titulos = obj.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': '#a6a6a6',
            'color':"white",
            "border_color":"black",
            "text_wrap":True,
            })

        self.s_bar_totales = obj.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': '#808080',
            'color':"white",
            "border_color":"black",
            "text_wrap":True,
            })

        self.s_gris_simple = obj.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': '#e7e6e6',
            "border_color":"black",
            })

        self.s_plomo_simple = obj.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': '#d0cece',
            "border_color":"black",
            })


        self.s_gris_simplev = obj.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': '#e7e6e6',
            "border_color":"black",
            "rotation":90,
            })


        self.s_plomo_simplev = obj.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': '#d0cece',
            "border_color":"black",
            "rotation":90,
            })

        self.s_cover1_simple = obj.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'color':"#595959",
            "border_color": "white",
            "font_size":24,
            "bold":True,

            })
        self.s_cover2_simple = obj.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'color':"#808080",
            "border_color": "white",
            "font_size":15,
            "bold":True,
            })

        self.s_cover3_simple = obj.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'color':"white",
            "border_color": "black",
            'fg_color': '#595959',
            "font_size":15,
            "bold":True,
            })


    # REPORTE EXCEL 1 
    def gene_reporte_uno(self):

        # valores a cargar *-*-*-*-*-*-*-*-**-*-*-*-*-*-*-*-*
        'pueden ser pasados por parametro'

        listaprimero = [    
            # pedido1-pedido2-pedido3-pedido4-pedido5-pedido6
            [8,2,8,1,8,8], 
            [3,0,0,7,80,98],
            [5,0,89,0,0,87],
            [5,0,0,78,89,7],
            [0,0,0,48,6,5], 
            [0,0,98,0,0,0], 
            [5,0,0,78,89,7],
            [0,0,0,48,6,5], 
            [0,0,98,0,0,0], 
            [8,2,8,1,8,8], 
            [3,0,0,7,80,98],
            [5,0,89,0,0,87],
            [5,0,0,78,89,7],
            [0,0,0,48,6,5], 
            [0,0,98,0,0,0], 
            [5,0,0,78,89,7],
            [0,0,0,48,6,5], 
            [0,0,98,0,0,0], 
        ]

        # titulo-encabezado *-*-*-*-*-*-*-*-**-*-*-*-*-*-*-*-*
        mesr = "(Enero - Febrero)"
        ahor = "Año 2020"
        titulor = "Encabezado"

        tituloprinr = "Reporte prueba"

        # crear nuevo excel y agregar worksheet
        workbook = xlsxwriter.Workbook(self.rep_uno_ex)
        worksheet = workbook.add_worksheet()
        worksheet.set_paper(9)  # A4
        worksheet.set_portrait() # vertical

        # ESTILOS *-*-*-*-*-*-*-*-**-*-*-*-*-*-*-*-*
        self.gen_estilos(workbook)


        # PRESENTACION *-*-*-*-*-*-*-*-**-*-*-*-*-*-*-*-*

        # insertar presentacion *-*-*-*-*-*-*-*-*
        worksheet.insert_image('A1', self.ruta_icon, {'x_scale': 0.18, 'y_scale': 0.18})

        worksheet.merge_range('A1:F1', tituloprinr,self.s_cover1_simple)
        worksheet.merge_range('A2:F2', ahor,self.s_cover2_simple)
        worksheet.merge_range('A3:F3', mesr,self.s_cover2_simple)

        # establecer tamaño variado de filas *-*-*-*-*-*-*-*-*
        worksheet.set_row(1-1, 30)
        worksheet.set_row(2-1, 20)
        worksheet.set_row(3-1, 20)
        worksheet.set_row(5-1, 30)
        worksheet.set_row(6-1, 40)

        worksheet.merge_range('A5:F5', titulor,self.s_cover3_simple)

        # CONFIGURACIONES *-*-*-*-*-*-*-*-**-*-*-*-*-*-*-*-*

        # establecer tamaño variado de columnas *-*-*-*-*-*-*-*-*
        for columna,ancho in zip(['A:A','B:B','C:C','D:D','E:E','F:F',],
                            [  13,    14,  13,   14, 13 ,  13,   20, ]):
            worksheet.set_column(columna,ancho)

        # establecer tamaño fijo de filas *-*-*-*-*-*-*-*-*
        '''
        en este caso es la fila es 7, sin embargo en aqui 
        es detectado como index, por se le aplica el -1
        '''

        numerosrecorrer = list(range(7,7+len(listaprimero)))
        numerofinal = numerosrecorrer[-1]

        worksheet.set_row((numerofinal+1)-1, 30)
        worksheet.set_row((numerofinal+2)-1, 35)

        for i in range(7-1,numerofinal):
            worksheet.set_row(i, 25)

        # establecer titulos *-*-*-*-*-*-*-*-*
        lisTitulos = [
            'Pedido\nUno','Pedido\nDos', 'Pedido\nTres',
            'Pedido\nCuatro', 'Pedido\nCinco', 'Pedido\nSeis',
            ]

        for val,let in zip(lisTitulos,"ABCDEF"):
            worksheet.write(f"{let}6", val, self.s_header_titulos)


        # MANEJO DE DATOS *-*-*-*-*-*-*-*-**-*-*-*-*-*-*-*-*

        # establecer valores *-*-*-*-*-*-*-*-*

        '''
        insertados como "int" debido a que se 
        realizara operaciones con dicho datos
        '''

        for lista,numero in zip(listaprimero,numerosrecorrer):
            for item,letra in zip(lista,"ABCDEF"):
                worksheet.write(f'{letra}{numero}', int(item), self.s_gris_simple)



        #obtener totales (operaciones excel) *-*-*-*-*-*-*-*-*
        for letra in "ABCDEF":
            worksheet.write(f'{letra}{numerofinal+1}', 
                            '{'+f'=SUM({letra}{numerosrecorrer[0]}:{letra}{numerofinal})'+'}', self.s_bar_totales)

        worksheet.merge_range(f'A{numerofinal+2}:F{numerofinal+2}', 
                            '{'+f'=SUM(A{numerofinal+1}:F{numerofinal+1})'+'}', self.s_bar_totales)

        workbook.close()

        # conversion pdf *-*-*-*-*-*-*-*-*
        self.raiz_manip_pdf.convertirdor_pdf(
            formato= 1, 
            ingreso= self.rep_uno_ex, 
            salida= self.rep_uno_pd
            )

        # guardar pdf *-*-*-*-*-*-*-*-*
        self.guardarArchivo(self.rep_uno_pd)

    # REPORTE EXCEL 2
    def gene_reporte_dos(self):

        # valores a cargar *-*-*-*-*-*-*-*-*
        'pueden ser pasados por parametro'

        listasegundo = [
            # caso1-caso2-caso3-caso4-caso5-caso6
            [3,6,3,8,8,3],
            [12,0,0,15,0,8],
            [12,0,16,0,0,0],
            [12,0,0,0,6,0],
            [0,0,65,0,30,0], 
            [0,8,84,0,0,60], 
        ]

        listaprimero = [    
            # caso1-caso2-caso3-caso4-caso5-caso6
            [8,2,8,1,8,8], 
            [3,0,0,7,80,98],
            [5,0,89,0,0,87],
            [5,0,0,78,89,7],
            [0,0,0,48,6,5], 
            [0,0,98,0,0,0], 
        ]

        # titulo-encabezado *-*-*-*-*-*-*-*-**-*-*-*-*-*-*-*-*
        mesr = "(Enero - Febrero)"
        ahor = "Año 2020"
        titulor = "Encabezado"

        tituloprinr = "Reporte prueba"

        # crear nuevo excel y agregar worksheet
        workbook = xlsxwriter.Workbook(self.rep_dos_ex)
        worksheet = workbook.add_worksheet()
        worksheet.set_paper(9)  # A4
        worksheet.set_landscape() # horizontal

        # ESTILOS *-*-*-*-*-*-*-*-**-*-*-*-*-*-*-*-*
        self.gen_estilos(workbook)


        # PRESENTACION *-*-*-*-*-*-*-*-**-*-*-*-*-*-*-*-*

        # insertar presentacion *-*-*-*-*-*-*-*-*
        worksheet.insert_image('A1', self.ruta_icon, {'x_scale': 0.18, 'y_scale': 0.18})

        worksheet.merge_range('A1:K1', tituloprinr,self.s_cover1_simple)
        worksheet.merge_range('A2:K2', ahor,self.s_cover2_simple)
        worksheet.merge_range('A3:K3', mesr,self.s_cover2_simple)

        # establecer tamaño variado de filas *-*-*-*-*-*-*-*-*
        worksheet.set_row(1-1, 30)
        worksheet.set_row(2-1, 20)
        worksheet.set_row(3-1, 20)
        worksheet.set_row(5-1, 30)
        worksheet.set_row(6-1, 40)

        worksheet.merge_range('A5:K5', titulor,self.s_cover3_simple)

        # CONFIGURACIONES *-*-*-*-*-*-*-*-**-*-*-*-*-*-*-*-*

        # establecer tamaño variado de columnas *-*-*-*-*-*-*-*-*
        for columna,ancho in zip(['A:A','B:B','C:C','D:D','E:E','F:F','G:G','H:H','I:I','J:J','K:K'],
                            [  5,    20,  10,    12,   12,   12,   10,   10,     10,   10,  10]):
            worksheet.set_column(columna,ancho)

        # establecer tamaño fijo de filas *-*-*-*-*-*-*-*-*
        '''
        en este caso es la fila es 7, sin embargo en aqui 
        es detectado como index, por se le aplica el -1
        '''
        for i in range(7-1,18):
            worksheet.set_row(i, 25)

        # cuadrado unido *-*-*-*-*-*-*-*-*
        worksheet.merge_range('A6:B6', '',self.s_header_titulos)

        # titulos verticales *-*-*-*-*-*-*-*-*
        worksheet.merge_range('A7:A12', 'Primero',self.s_gris_simplev)
        worksheet.merge_range('A13:A18', 'Segundo',self.s_plomo_simplev)

        # establecer titulos *-*-*-*-*-*-*-*-*

        lisTitulos = [
            "Asistencia uno","Asistencia dos","Asistencia tres",
            "Asistencia cuatro","Asistencia cinco","Asistencia seis"
            ]

        for ar,ti in zip(range(7,12+1),lisTitulos):
            worksheet.write(f'B{ar}', ti,self.s_gris_simple)

        for ar,ti in zip(range(13,18+1),lisTitulos):
            worksheet.write(f'B{ar}', ti,self.s_plomo_simple)

        lisTitulos = [
            'Caso\nUno','Caso\nDos', 'Caso\nTres',
            'Caso\nCuatro', 'Caso\nCinco', 'Caso\nSeis',
            ]

        for val,let in zip(lisTitulos,"CDEFGH"):
            worksheet.write(f"{let}6", val, self.s_header_titulos)

        worksheet.merge_range('I6:K6', 'Total', self.s_bar_totales)

        # MANEJO DE DATOS *-*-*-*-*-*-*-*-**-*-*-*-*-*-*-*-*

        # establecer valores *-*-*-*-*-*-*-*-*

        '''
        insertados como "int" debido a que se 
        realizara operaciones con dicho datos
        '''

        for lista,numero in zip(listaprimero,(7,8,9,10,11,12)):
            for item,letra in zip(lista,"CDEFGH"):
                worksheet.write(f'{letra}{numero}', int(item), self.s_gris_simple)

        for lista,numero in zip(listasegundo,(13,14,15,16,17,18)):
            for item,letra in zip(lista,"CDEFGH"):
                worksheet.write(f'{letra}{numero}', int(item), self.s_plomo_simple)

        #obtener totales (operaciones excel) *-*-*-*-*-*-*-*-*

        for num in range(7,18+1):
            worksheet.write_array_formula(f'I{num}:I{num}', '{'+f'=SUM(C{num}:H{num})'+'}', self.s_bar_totales)

        worksheet.merge_range('J7:J12', '{=SUM(C7:H12)}', self.s_bar_totales)
        worksheet.merge_range('J13:J18', '{=SUM(C13:H18)}', self.s_bar_totales)
        worksheet.merge_range('K7:K18', '{=SUM(C7:H18)}',self.s_bar_totales)

        # cerra archivo
        workbook.close()

        # conversion pdf *-*-*-*-*-*-*-*-*
        self.raiz_manip_pdf.convertirdor_pdf(
            formato= 2, 
            ingreso= self.rep_dos_ex, 
            salida= self.rep_dos_pd
            )

        # guardar pdf *-*-*-*-*-*-*-*-*
        self.guardarArchivo(self.rep_dos_pd)

