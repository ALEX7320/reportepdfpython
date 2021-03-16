# mosulos de PySide2
from PySide2.QtWidgets import QApplication

# clase generar pdf
from Funcion_pdf import ClaseGeneradorExcel

if __name__ == "__main__":

    import sys
    app = QApplication(sys.argv)
    app.setStyle('Fusion') # opcional

    reportepdf = ClaseGeneradorExcel() # principal
    opcion = input('\n[1] PDF Vertical\n[2] PDF Horizontal\nOpción: ')

    if(opcion=='1'):
        print('Reporte Vertical')
        reportepdf.gene_reporte_uno() # genera

    if(opcion=='2'):
        print('Reporte Horizontal')
        reportepdf.gene_reporte_dos() # genera
    
    else:
        print('Opcion no valida')
        sys.exit()

    print('Fin del proceso')
    
