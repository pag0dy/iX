# iX
Aplicación basada en IfcOpenShell, PyQt5, Pandas y XlsxWriter para extraer la información principal de los modelos BIM en formato [IFC](https://technical.buildingsmart.org/standards/ifc/) en un reporte estilo hoja de cálculo, con extensión xlsx. 

Compatibilidad probada con IFC 4, posiblemente también funciona con IFC 2x3.

![GUI](/ix.png)

# Archivos y Dependencias
Esta aplicación fue creada con Python 3.8. Para replicarla es necesario instalar las librerías PyQt5, Pandas, IfcOpenShell y XlsxWriter. 
Los archivos IfcXcl.py y sobreix.py contienen el código de la interfaz gráfica de la aplicación y la ventana "Sobre la aplicación".
El archivo Funciones contiene el código que extrae la información del archivo IFC y genera la hoja de cálculo. 
La hoja de cálculo generada contiene instrucciones para facilitar la comprensión de la información extraida. 

# Historial de versiones

- 0.1
  - Versión Beta lanzada para realizar pruebas
  
# Meta
Paulina Godoy Del Campo - @pag0dy - pauli@bimfluent.cl
Distribuida bajo licencia GNU General Public License v3.0
