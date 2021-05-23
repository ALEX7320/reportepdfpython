# Generar reporte PDF Python 3

Al trabajar con pdf en Python se suele usar la libreria `reportlab`; sin embargo tambien hay otra manera, la cual es trabajando con archivos excel y posteriromente convertirlos a pdf.

Lo importante en este algoritmo es saber el manejo de excel con la libreria [`xlsxwriter`](https://xlsxwriter.readthedocs.io/ "xlsxwriter") cuya documentación pienso que es muy comprensible, ademas de tener un conocimiento básico de rutas.

**Nota**
Se publicó una guía sobre 'Como crear PDF en Python con FPDF' guía totalmente gratis en Youtube → [ENLACE](https://github.com/ALEX7320/guia-pdf-python "ENLACE")

**Indice**

  * [Recursos utilizados](#recursos-utilizados)
  * [Documentación](#documentación)
	* [Rutas](#rutas)
	* [Diseño](#diseño)
	* [Conversión](#conversión)
	* [Guardado](#guardado)
  * [Fuentes](#fuentes)
  * [Previsualización](#previsualización)

# Recursos utilizados

`pip install PySide2` → (Solo para obtener ruta de guardado)

`pip install XlsxWriter`

`pip install pywin32`


# Documentación


### Rutas

Lo importante en trabajar con nuestra ruta actual
Modulo: `Fucion_pdf ` / `ClaseGeneradorExcel()` / `init`

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

### Diseño

Enfocandonos en la creación de la hoja necesitamos conocimiento previo en xlsxwriter, para realizar las plantillas. 


[Documentación xlsxwriter](https://xlsxwriter.readthedocs.io/ "Documentación xlsxwriter")

Y es alli donde le pasaremos la ruta auxiliar del 1er excel.

Modulo: `Fucion_pdf ` / `ClaseGeneradorExcel()` / `def gene_reporte_uno(sefl)`


```python
workbook = xlsxwriter.Workbook(self.rep_uno_ex)
worksheet = workbook.add_worksheet()
worksheet.set_paper(9)  # A4
worksheet.set_portrait() # vertical

# configuraciones aqui

workbook.close()
```

Tenemos que tener en cuenta que si tenemos varias hojas que comparten el mismo estilo, tenemos que pasarle el `workbook` esto para que sea el objeto al que van dirijido los estilos. 

```python
self.gen_estilos(workbook)
```

Una vez realizado la plantilla, procedemos a pregarlo en su respectiva función:
```python
    def gene_reporte_uno(self):
		pass
		
    def gene_reporte_dos(self):
		pass
```

### Conversión

Primero se debio instanciar el convertidor de PDF

Modulo: `Fucion_pdf ` / `ClaseGeneradorExcel()` / `init`

```python
self.raiz_manip_pdf = ClaseManipularPdf()
```

Necesitamos tener estos parametros en cuenta

```python
def convertirdor_pdf(self, ingreso, salida):
    '''
    ingreso : ruta excel
    salida : ruta pdf
    '''
```
Recordemos que al final de cada plantilla se realiza la conversión, cons sus respectivas rutas.

```python
        # conversion pdf *-*-*-*-*-*-*-*-*
        self.raiz_manip_pdf.convertirdor_pdf(
            ingreso= self.rep_uno_ex, 
            salida= self.rep_uno_pd
            )
```

Este algoritmo convertira el excel a pdf, teniendo en cuenta sus respectivas rutas de entrada y salida.

```python
app = win32.gencache.EnsureDispatch("Excel.Application")
app.Interactive = False
app.Visible = False
app.DisplayAlerts = False

Workbook = app.Workbooks.Open(ingreso)

# salida
Workbook.ExportAsFixedFormat(0, salida)
Workbook.RefreshAll()
app.Quit()
```

### Guardado

Al igual que en todo lo demas, necesitamos la ruta del pdf, en donde con el `QFileDialog` se elejira la ruta de guardao, y el `shutil.move` movera el archivo con el nombre asignado por el usuario.

Por ultimo el archivo pdf se abrira automaticamente con el `webbrowser`.

```python
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

```




# Fuentes

 * [TheQuickBlog (Script de conversión)](https://thequickblog.com/convert-an-excel-filexlsx-to-pdf-python/ "TheQuickBlog")
 * [Stackoverflow (Configuraciones)](https://stackoverflow.com/questions/42385563/convert-excel-to-pdf-in-landscape-orientation "Stackoverflow")
 
# Previsualización
 
 **Uno**
 
 ![](https://1.bp.blogspot.com/-vAr0NoNVq9A/YFAx9ddTwOI/AAAAAAAAAG8/MrY2VsOjyegVVHLWY0XELSOAmt11B_XJwCLcBGAsYHQ/s1600/v.jpg)
 
 **Dos**
 
 ![](https://1.bp.blogspot.com/-lvrogixM5_E/YFAx9U36hlI/AAAAAAAAAHA/F39qX3kGQYEFluYwK8kBM4Boj3PV76eYwCLcBGAsYHQ/s1600/h)
 