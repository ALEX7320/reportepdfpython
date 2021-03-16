# Generar reporte PDF Python 3

Al trabajar con pdf en Python se suele usar la libreria reportlab; sin embargo tambien hay otra manera, la cual es trabajando con archivos excel y posteriromente convertirlos a pdf.

Lo importante en este algoritmo es saber el manejo de excel con la libreria xlsxwriter cuya documentación pienso que es muy comprensible, ademas de tener un conocimiento basico de rutas.

**Indice**

  * [Recursos utilizados](#recursos-utilizados)
  * [Documentación](#documentación)
  * [Fuentes](#fuentes)
  * [Previzualización](#previzualización)

# Recursos utilizados

`pip install PySide2` → (Solo para obtener ruta de guardado)

`pip install XlsxWriter`

`pip install pywin32`

# Documentación

Lo importante en trabajar con nuestra ruta actual

```python
raiz = os.getcwd()
```

En este caso utilizamos un icono en el Excel es por ello que tambien tenemos que indicar la ubicación de dicho archivo.

```python
self.ruta_icon = raiz+r'\icon\icono.jpg'
```
Necesitamos trabajar con una carpeta auxiliar, donde se almacenara el excel generado (esta carpeta no estara repleto de archivo, solo de los excel que deseamos trabajar, cabe recalcar que los nuevos no se vuelve a almacenar, simplemente reemplazaran al existente) para ello se creo el aux_excel 

![](https://1.bp.blogspot.com/-dW19TRGwG8w/YFAx9d3ocrI/AAAAAAAAAG4/cFZcUyTzuPkQLJKM4xm8j_45text9oOeACLcBGAsYHQ/s1600/ca.png)

Ahora dependiendo de cuantos PDF deseamos generar, tendremos que tener una ruta de excel y pdf para cada uno, primero haciendo referencia a la ruta auxiliar (aux_excel)

```python
self.rep_uno_ex = raiz+r'\aux_excel\reporte_uno.xlsx' # entrada excel
self.rep_uno_pd = raiz+r'\aux_excel\reporte_uno.pdf' # salida pdf

self.rep_dos_ex = raiz+r'\aux_excel\reporte_dos.xlsx' # entrada excel
self.rep_dos_pd = raiz+r'\aux_excel\reporte_dos.pdf' # salida pdf
```

Ahora tendremos que instanciar el convertidor de PDF

```python
self.raiz_manip_pdf = ClaseManipularPdf()
```

Necesitamos tener estos parametros en cuenta

```python
def convertirdor_pdf(self, formato, ingreso, salida):
    '''
    formato : orientacion de hoja
    ingreso : ruta excel
    salida : ruta pdf
    '''
```

Este algoritmo convertira el excel a pdf, teniendo en cuenta sus respectivas rutas de entrada y salida.

```python
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
```

# Fuentes

 * [TheQuickBlog (Script de conversión)](https://thequickblog.com/convert-an-excel-filexlsx-to-pdf-python/ "TheQuickBlog")
 * [Stackoverflow (Configuraciones)](https://stackoverflow.com/questions/42385563/convert-excel-to-pdf-in-landscape-orientation "Stackoverflow")
 
# Previzualización
 
 **Uno**
 
 ![](https://1.bp.blogspot.com/-vAr0NoNVq9A/YFAx9ddTwOI/AAAAAAAAAG8/MrY2VsOjyegVVHLWY0XELSOAmt11B_XJwCLcBGAsYHQ/s1600/v.jpg)
 
 **Dos**
 
 ![](https://1.bp.blogspot.com/-lvrogixM5_E/YFAx9U36hlI/AAAAAAAAAHA/F39qX3kGQYEFluYwK8kBM4Boj3PV76eYwCLcBGAsYHQ/s1600/h)
 