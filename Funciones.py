import ifcopenshell as IfcOs
import pandas as pd
import xlsxwriter as wr
import datetime
import time
import getpass
import os
import os.path
from pathlib import Path

directorio = ''


class ArchIfc:
    def __init__(self, ruta):
        try:
            self.Ifc = IfcOs.open(ruta)
            global directorio
            directorio = Path(ruta).parent
            print(directorio)
        except IOError:
            print('Archivo seleccionado no es válido')

    def infomod(self, Ifc):
        info_mod = {}
        proy = Ifc.by_type('IfcProject')[0]
        info_mod['ProjectName'] = proy.Name
        edi = Ifc.by_type('IfcBuilding')[0]
        info_mod['BuildingName'] = edi.Name
        apli = Ifc.by_type('IfcApplication')[0]
        info_mod['Application'] = apli.ApplicationFullName
        info_mod['Ifc Schema'] = Ifc.schema
        dfInfo = pd.DataFrame.from_dict(info_mod, orient='index', columns=['Resumen'])
        return dfInfo

    def crear_repo(self, Ifc):
        #Crear archivo xlsx base para el reporte
        info_resumen = []
        proy = Ifc.by_type('IfcProject')[0]
        nom_modelo = ''.join(x for x in proy.Name if x.isalnum())# Esta línea elimina los posibles caracteres invalidos del nombre
        global directorio
        reporte = wr.Workbook(str(directorio) + '/' + nom_modelo + '_reporte.xlsx')
        r_portada = reporte.add_worksheet('PORTADA')
        r_instruc = reporte.add_worksheet('INSTRUCCIONES')
        r_resumen = reporte.add_worksheet('RESUMEN')
        # Crear formatos para reporte
        fecha = datetime.date.today()
        t = time.localtime()
        hora = time.strftime("%H:%M:%S", t)
        usuario = getpass.getuser()
        titulo = reporte.add_format(
                 {'bold': True, 'font_name': 'Roboto', 'font_color': 'white', 'font_size': 28,
                  'bg_color': '#244155'})
        subtitulo = reporte.add_format(
                 {'font_name': 'Roboto', 'font_color': 'white', 'font_size': 20, 'bg_color': '#244155'})
        text_norm = reporte.add_format(
                 {'font_name': 'Roboto', 'font_color': 'white', 'font_size': 10, 'bg_color': '#244155'})
        text_bold = reporte.add_format(
                 {'font_name': 'Roboto', 'font_color': 'white', 'font_size': 10, 'bg_color': '#244155', 'bold': True})
        data_text = reporte.add_format(
                 {'font_name': 'Roboto', 'font_color': '#244155', 'font_size': 10, 'align': 'left'})
        fondo = reporte.add_format({'bg_color': '#244155'})

        # Crear portada reporte
        r_portada.set_column('A:A', 21)
        r_portada.set_column('B:B', 60)
        r_portada.write('B1', 'Reporte Modelo IFC', titulo)
        r_portada.write('B2', 'iX - Ifc a Xlsx', subtitulo)
        r_portada.write('B3', 'Usuario: ' + usuario, text_norm)
        r_portada.write('B4', 'Fecha: ' + str(fecha) + ' ' + str(hora), text_norm)
        r_portada.merge_range('A5:B5', None, fondo)
        r_portada.insert_image('A1', 'ix.png')

        # Crear instrucciones reporte
        r_instruc.set_column('A:A', 5)
        r_instruc.set_column('B:B', 120)
        r_instruc.write('B2', 'INSTRUCCIONES:', subtitulo)
        r_instruc.write('B3', 'A continuación encontrarás indicaciones para comprender de mejor forma este reporte.'
                        , text_norm)
        r_instruc.write('B4', 'Pestaña Resumen: ', text_bold)
        r_instruc.write('B5', 'Contiene la información básica del modelo y su origen.', text_norm)
        r_instruc.write('B6', 'Pestañas Proyecto, Edificio y Terreno: ', text_bold)
        r_instruc.write('B7', 'Contienen los atributos más relevantes de estas entidades.', text_norm)
        r_instruc.write('B8', 'Pestaña Entidades: ', text_bold)
        r_instruc.write('B9', 'Contiene todas las entidades del modelo, con su GUId y nombre.', text_norm)
        r_instruc.write('B10', 'Pestañas Atributos: ', text_bold)
        r_instruc.write('B11', 'Contienen los atributos de las entidades del modelo con excepción de Proyecto, '
                               'Edificio, Terreno y Ejes. Estas pestañas contienen la información', text_norm)
        r_instruc.write('B12', 'ordenadade la siguiente manera: ', text_norm)
        r_instruc.write('B13', 'Atributos_1: Elementos constructivos básicos (Muros, Vigas, Columnas, Losas, Techumbres'
                               ', Fundaciones y Estructuras Especiales)', text_bold)
        r_instruc.write('B14', 'Atributos_2: Puertas y Ventanas ', text_bold)
        r_instruc.write('B15', 'Atributos_3: Escaleras y Rampas', text_bold)
        r_instruc.write('B16', 'Atributos_4: Elementos de sistemas MEP', text_bold)
        r_instruc.write('B17', 'Atributos_5: Mobiliario y elementos secundarios (Barandas, Recubrimientos, entre otros)'
                        , text_bold)
        r_instruc.write('B18', 'Atributos_6: Espacios y Zonas', text_bold)
        r_instruc.write('B19', 'Atributos_7: Elementos Geográficos y Civiles', text_bold)
        r_instruc.write('B20', 'Pestaña Propiedades: ', text_bold)
        r_instruc.write('B21', 'Contiene todas las propiedades (que son parte de un Pset) de todas las entidades del '
                               'modelo.', text_norm)
        r_instruc.write('B22', 'Pestaña Cuantías: ', text_bold)
        r_instruc.write('B23', 'Contiene todas las cuantías (que son parte de un Qto) de todas las entidades del '
                               'modelo.', text_norm)

        # Ingresar información en hoja Resumen
        nom_proy = ''
        if proy.LongName is None:
            nom_proy = proy.Name
        else:
            nom_proy = proy.LongName
        info_resumen.append(('PROYECTO', nom_proy))
        edi = Ifc.by_type('IfcBuilding')[0]
        nom_edi = ''
        if edi.LongName is None:
            nom_edi = edi.Name
        else:
            nom_edi = edi.LongName
        info_resumen.append(('EDIFICACIÓN', nom_edi))
        apli = Ifc.by_type('IfcApplication')[0]
        info_resumen.append(('APLICACIÓN DE ORIGEN', apli.ApplicationFullName))
        version = Ifc.schema
        info_resumen.append(('VERSIÓN ESQUEMA IFC', version))
        r_resumen.merge_range('A1:B1', 'RESUMEN DEL MODELO', text_bold)
        r_resumen.set_column('A:A', 25)
        r_resumen.set_column('B:B', 60)

        row = 1
        col = 0

        for i, v in info_resumen:
            r_resumen.write(row, col, i, text_bold)
            r_resumen.write(row, col+1, v, data_text)
            row +=1

        # Crear hoja Proyecto e ingresar información
        r_proyecto = reporte.add_worksheet('PROYECTO')
        r_proyecto.set_column('A:A', 20)
        r_proyecto.set_column('B:B', 60)
        proyatt = []
        proyatt.append(('GLOBAL ID', proy.GlobalId))
        if proy.LongName is None:
            proyatt.append(('LONGNAME', 'No ingresado'))
        else:
            proyatt.append(('LONGNAME', proy.LongName))
        if proy.Name is None:
            proyatt.append(('NAME', 'No ingresado'))
        else:
            proyatt.append(('NAME', proy.Name))
        if proy.Description is None:
            proyatt.append(('DESCRIPTION', 'No ingresada'))
        else:
            proyatt.append(('DESCRIPTION', proy.Description))
        if proy.Phase is None:
            proyatt.append(('PHASE', 'No ingresado'))
        else:
            proyatt.append(('PHASE', proy.Phase))

        r_proyecto.merge_range('A1:B1', 'INFORMACIÓN DEL PROYECTO', text_bold)

        row = 1
        col = 0

        for i, v in proyatt:
            r_proyecto.write(row, col, i, text_bold)
            r_proyecto.write(row, col+1, v, data_text)
            row +=1

        # Crear hoja Edificio e ingresar información

        r_edificio = reporte.add_worksheet('EDIFICIO')
        r_edificio.set_column('A:A', 30)
        r_edificio.set_column('B:B', 60)
        r_edificio.merge_range('A1:B1', 'INFORMACIÓN DEL EDIFICIO', text_bold)

        eInfo = []
        eInfo.append(('GLOBAL ID', edi.GlobalId))

        if edi.LongName is None:
            eInfo.append(('LONGNAME', 'No ingresado'))
        else:
            eInfo.append(('LONGNAME', edi.LongName))

        if edi.Name is None:
            eInfo.append(('NAME', 'No ingresado'))
        else:
            eInfo.append(('NAME', edi.Name))

        if edi.Description is None:
            eInfo.append(('DESCRIPTION', 'No ingresado'))
        else:
            eInfo.append(('DESCRIPTION', edi.Description))

        if edi.ObjectType is None:
            eInfo.append(('OBJECT TYPE', 'No ingresado'))
        else:
            eInfo.append(('OBJECT TYPE', edi.ObjectType))

        if edi.CompositionType is None:
            eInfo.append(('COMPOSITION TYPE', 'No ingresado'))
        else:
            eInfo.append(('COMPOSITION TYPE', edi.CompositionType))

        if edi.ElevationOfRefHeight is None:
            eInfo.append(('ELEVATION OF REF HEIGHT', 'No ingresada'))
        else:
            eInfo.append(('ELEVATION OF REF HEIGHT', edi.ElevationOfRefHeight))

        if edi.ElevationOfTerrain is None:
            eInfo.append(('ELEVATION OF TERRAIN', 'No ingresada'))
        else:
            eInfo.append(('ELEVATION OF TERRAIN', edi.ElevationOfTerrain))

        if edi.BuildingAddress is None:
            eInfo.append(('BUILDING ADDRESS', 'No ingresada'))
        else:
            edirec = ''
            for i in edi.BuildingAddress:
                if i is None:
                    pass
                else:
                    edirec = edirec + str(i) + '|'
            eInfo.append(('BUILDING ADDRESS', edirec))
        pisos = Ifc.by_type('IfcBuildingStorey')
        eInfo.append(('BUILDING STOREYS', str(len(pisos))))

        row = 1
        col = 0

        for i, v in eInfo:
            r_edificio.write(row, col, i, text_bold)
            r_edificio.write(row, col + 1, v, data_text)
            row += 1

        # Crear hoja de Terreno e ingresar información
        r_terreno = reporte.add_worksheet('TERRENO')
        r_terreno.set_column('A:A', 20)
        r_terreno.set_column('B:B', 60)
        r_terreno.merge_range('A1:B1', 'INFORMACIÓN DEL TERRENO', text_bold)

        terreno = Ifc.by_type('IfcSite')[0]
        infote = [('GLOBAL ID', terreno.GlobalId)]
        if terreno.LongName is None:
            infote.append(('LONGNAME', 'No ingresado'))
        else:
            infote.append(('LONGNAME', terreno.LongName))
        if terreno.Name is None:
            infote.append(('NAME', 'No ingresado'))
        else:
            infote.append(('NAME', terreno.Name))
        if terreno.Description is None:
            infote.append(('DESCRIPTION', 'No ingresado'))
        else:
            infote.append(('DESCRIPTION', terreno.Description))
        if terreno.ObjectType is None:
            infote.append(('OBJECT TYPE', 'No ingresado'))
        else:
            infote.append(('OBJECT TYPE', terreno.ObjectType))
        if terreno.CompositionType is None:
            infote.append(('COMPOSITION TYPE', 'No ingresado'))
        else:
            infote.append(('COMPOSITION TYPE', terreno.CompositionType))
        infote.append(('REF LATITUDE', str(terreno.RefLatitude)))
        infote.append(('REF LONGITUDE', str(terreno.RefLongitude)))
        infote.append(('REF ELEVATION', str(terreno.RefElevation)))
        if terreno.LandTitleNumber is None:
            infote.append(('LAND TITLE NUMBER', 'No ingresado'))
        else:
            infote.append(('LAND TITLE NUMBER', terreno.LandTitleNumber))
        if terreno.SiteAddress is None:
            infote.append(('SITE ADDRESS', 'No ingresado'))
        else:
            tedirec = ''
            for i in terreno.SiteAddress:
                if i is None:
                    pass
                else:
                    tedirec = tedirec + str(i) + '|'
            infote.append(('SITE ADDRESS', tedirec))

        row = 1
        col = 0

        for i, v in infote:
            r_terreno.write(row, col, i, text_bold)
            r_terreno.write(row, col + 1, v, data_text)
            row += 1

        # Crear pestaña Entidades
        r_entidades = reporte.add_worksheet('ENTIDADES')
        r_entidades.set_column('A:A', 20)
        r_entidades.set_column('B:B', 30)
        r_entidades.set_column('C:C', 60)
        r_entidades.merge_range('A1:C1', 'ENTIDADES DEL MODELO', text_bold)
        r_entidades.write('A2', 'IFC ENTITY', text_bold)
        r_entidades.write('B2', 'GLOBALID', text_bold)
        r_entidades.write('C2', 'NAME', text_bold)
        id_ent = []
        entidades = Ifc.by_type('IfcObject')
        for en in entidades:
            id_ent.append((str(en.is_a()), str(en.GlobalId), str(en.Name)))

        row = 2
        col = 0

        for e, i, n in id_ent:
            r_entidades.write(row, col, e, data_text)
            r_entidades.write(row, col + 1, i, data_text)
            r_entidades.write(row, col + 2, n, data_text)
            row += 1

        # Crear pestaña Atributos_1
        r_atributos_1 = reporte.add_worksheet('ATRIBUTOS_1')
        r_atributos_1.set_column('A:A', 6)
        r_atributos_1.set_column('B:B', 30)
        r_atributos_1.set_column('C:C', 30)
        r_atributos_1.set_column('D:D', 50)
        r_atributos_1.set_column('E:E', 15)
        r_atributos_1.set_column('F:F', 50)
        r_atributos_1.set_column('G:G', 10)
        r_atributos_1.set_column('H:H', 20)

        r_atributos_1.merge_range('A1:H1', 'ATRIBUTOS DE ELEMENTOS CONSTRUCTIVOS BÁSICOS', text_bold)
        attri_1 = []
        enti_1 = ['IfcWall', 'IfcWallStandardCase', 'IfcColumn', 'IfcBeam', 'IfcSlab', 'IfcFooting', 'IfcPile',
                  'IfcCurtainWall', 'IfcRoof', 'IfcElementAssembly']
        for e in entidades:
            if e.is_a() in enti_1:
                attri_1.append(e.get_info())
        dfAtri = pd.DataFrame(attri_1).drop_duplicates()
        if len(dfAtri) > 0:
            dfA = dfAtri.drop(columns=['OwnerHistory', 'ObjectPlacement', 'Representation'], axis=1)
            dfA = dfA.fillna('No disponible')
            row = 1
            col = 0

            for c in dfA.columns:
                r_atributos_1.write(row, col, c, text_bold)
                col += 1

            row = 2
            col = 0
            x = 0
            while x in range(len(dfA.columns)):
                for i in range(len(dfA)):
                    r_atributos_1.write(row, col, dfA.iloc[i, x], data_text)
                    row +=1
                row = 2
                x += 1
                col +=1

        else:
            row = 1
            col = 0
            r_atributos_1.write(row, col, 'Este proyecto no contiene elementos constructivos básicos.', data_text)

        # Crear pestaña Atributos_2
        r_atributos_2 = reporte.add_worksheet('ATRIBUTOS_2')
        r_atributos_2.set_column('A:A', 6)
        r_atributos_2.set_column('B:B', 10)
        r_atributos_2.set_column('C:C', 30)
        r_atributos_2.set_column('D:D', 50)
        r_atributos_2.set_column('E:E', 15)
        r_atributos_2.set_column('F:F', 50)
        r_atributos_2.set_column('G:G', 10)
        r_atributos_2.set_column('H:H', 15)
        r_atributos_2.set_column('I:I', 15)
        r_atributos_2.set_column('J:J', 15)
        r_atributos_2.set_column('K:K', 25)
        r_atributos_2.set_column('L:L', 25)

        r_atributos_2.merge_range('A1:L1', 'ATRIBUTOS DE PUERTAS Y VENTANAS', text_bold)
        attri_2 = []
        enti_2 = ['IfcWindow', 'IfcDoor']

        for e in entidades:
            if e.is_a() in enti_2:
                attri_2.append(e.get_info())
        dfAtri_2 = pd.DataFrame(attri_2).drop_duplicates()
        if len(dfAtri_2) > 0:
            dfA_2 = dfAtri_2.drop(columns=['OwnerHistory', 'ObjectPlacement', 'Representation'], axis=1)
            dfA_2 = dfA_2.fillna(value='No disponible')
            row = 1
            col = 0

            for c in dfA_2.columns:
                r_atributos_2.write(row, col, c, text_bold)
                col += 1

            row = 2
            col = 0
            x = 0
            while x in range(len(dfA_2.columns)):
                for i in range(len(dfA_2)):
                    r_atributos_2.write(row, col, dfA_2.iloc[i, x], data_text)
                    row += 1
                row = 2
                x += 1
                col += 1

        else:
            row = 1
            col = 0
            r_atributos_2.write(row, col, 'Este proyecto no contiene puertas ni ventanas.', data_text)

        # Crear pestaña Atributos_3
        r_atributos_3 = reporte.add_worksheet('ATRIBUTOS_3')
        r_atributos_3.set_column('A:A', 6)
        r_atributos_3.set_column('B:B', 10)
        r_atributos_3.set_column('C:C', 30)
        r_atributos_3.set_column('D:D', 50)
        r_atributos_3.set_column('E:E', 15)
        r_atributos_3.set_column('F:F', 50)
        r_atributos_3.set_column('G:G', 10)
        r_atributos_3.set_column('H:H', 15)
        r_atributos_3.set_column('I:I', 15)
        r_atributos_3.set_column('J:J', 15)
        r_atributos_3.set_column('K:K', 25)
        r_atributos_3.set_column('L:L', 25)

        r_atributos_3.merge_range('A1:L1', 'ATRIBUTOS DE ESCALERAS Y RAMPAS', text_bold)
        attri_3 = []
        enti_3 = ['IfcStair', 'IfcStairFlight', 'IfcRamp',  'IfcRampFlight']

        for e in entidades:
            if e.is_a() in enti_3:
                attri_3.append(e.get_info())
        dfAtri_3 = pd.DataFrame(attri_3).drop_duplicates()
        if len(dfAtri_3) > 0:
            dfA_3 = dfAtri_3.drop(columns=['OwnerHistory', 'ObjectPlacement', 'Representation'], axis=1)
            dfA_3 = dfA_3.fillna(value='No disponible')
            row = 1
            col = 0

            for c in dfA_3.columns:
                r_atributos_3.write(row, col, c, text_bold)
                col += 1

            row = 2
            col = 0
            x = 0
            while x in range(len(dfA_3.columns)):
                for i in range(len(dfA_3)):
                    r_atributos_3.write(row, col, dfA_3.iloc[i, x], data_text)
                    row += 1
                row = 2
                x += 1
                col += 1
        else:
            row = 1
            col = 0
            r_atributos_3.write(row, col, 'Este proyecto no contiene escaleras ni rampas.', data_text)

            # Crear pestaña Atributos_4
            r_atributos_4 = reporte.add_worksheet('ATRIBUTOS_4')
            r_atributos_4.set_column('A:A', 6)
            r_atributos_4.set_column('B:B', 25)
            r_atributos_4.set_column('C:C', 30)
            r_atributos_4.set_column('D:D', 50)
            r_atributos_4.set_column('E:E', 15)
            r_atributos_4.set_column('F:F', 50)
            r_atributos_4.set_column('G:G', 10)
            r_atributos_4.set_column('H:H', 15)
            r_atributos_4.set_column('I:I', 15)
            r_atributos_4.set_column('J:J', 15)
            r_atributos_4.set_column('K:K', 25)
            r_atributos_4.set_column('L:L', 25)

            r_atributos_4.merge_range('A1:H1', 'ATRIBUTOS DE ELEMENTOS MEP', text_bold)
            attri_4 = []
            entiflow = Ifc.by_type('IfcDistributionElement')
            for e in entiflow:
                attri_4.append(e.get_info())
            dfAtri_4 = pd.DataFrame(attri_4).drop_duplicates()
            if len(dfAtri_4) > 0:
                dfA_4 = dfAtri_4.drop(columns=['OwnerHistory', 'ObjectPlacement', 'Representation'], axis=1)
                dfA_4 = dfA_4.fillna(value='No disponible')
                row = 1
                col = 0

                for c in dfA_4.columns:
                    r_atributos_4.write(row, col, c, text_bold)
                    col += 1

                row = 2
                col = 0
                x = 0
                while x in range(len(dfA_4.columns)):
                    for i in range(len(dfA_4)):
                        r_atributos_4.write(row, col, dfA_4.iloc[i, x], data_text)
                        row += 1
                    row = 2
                    x += 1
                    col += 1
            else:
                row = 1
                col = 0
                r_atributos_4.write(row, col, 'Este proyecto no contiene elementos MEP.', data_text)

        # Crear pestaña Atributos_5
        r_atributos_5 = reporte.add_worksheet('ATRIBUTOS_5')
        r_atributos_5.set_column('A:A', 6)
        r_atributos_5.set_column('B:B', 25)
        r_atributos_5.set_column('C:C', 30)
        r_atributos_5.set_column('D:D', 50)
        r_atributos_5.set_column('E:E', 15)
        r_atributos_5.set_column('F:F', 50)
        r_atributos_5.set_column('G:G', 10)
        r_atributos_5.set_column('H:H', 15)
        r_atributos_5.set_column('I:I', 15)
        r_atributos_5.set_column('J:J', 15)
        r_atributos_5.set_column('K:K', 25)
        r_atributos_5.set_column('L:L', 25)

        r_atributos_5.merge_range('A1:H1', 'ATRIBUTOS DE MOBILIARIO Y ELEMENTOS SECUNDARIOS', text_bold)
        attri_5 = []
        enti_5 = ['IfcFurniture', 'IfcSystemFurnitureElement', 'IfcShadingDevice', 'IfcCovering', 'IfcPlate',
                  'IfcMember', 'IfcRailing', 'IfcBuildingElementProxy']
        for e in entidades:
            if e.is_a() in enti_5:
                attri_5.append(e.get_info())
        dfAtri_5 = pd.DataFrame(attri_5).drop_duplicates()
        if len(dfAtri_5) > 0:
            dfA_5 = dfAtri_5.drop(columns=['OwnerHistory', 'ObjectPlacement', 'Representation'], axis=1)
            dfA_5 = dfA_5.fillna(value='No disponible')
            row = 1
            col = 0

            for c in dfA_5.columns:
                r_atributos_5.write(row, col, c, text_bold)
                col += 1

            row = 2
            col = 0
            x = 0
            while x in range(len(dfA_5.columns)):
                for i in range(len(dfA_5)):
                    r_atributos_5.write(row, col, dfA_5.iloc[i, x], data_text)
                    row += 1
                row = 2
                x += 1
                col += 1
        else:
            row = 1
            col = 0
            r_atributos_5.write(row, col, 'Este proyecto no contiene mobiliario ni elementos secundarios.', data_text)

        # Crear pestaña Atributos_6
        r_atributos_6 = reporte.add_worksheet('ATRIBUTOS_6')
        r_atributos_6.set_column('A:A', 6)
        r_atributos_6.set_column('B:B', 25)
        r_atributos_6.set_column('C:C', 30)
        r_atributos_6.set_column('D:D', 50)
        r_atributos_6.set_column('E:E', 15)
        r_atributos_6.set_column('F:F', 50)
        r_atributos_6.set_column('G:G', 10)
        r_atributos_6.set_column('H:H', 15)
        r_atributos_6.set_column('I:I', 15)
        r_atributos_6.set_column('J:J', 15)
        r_atributos_6.set_column('K:K', 25)
        r_atributos_6.set_column('L:L', 25)

        r_atributos_6.merge_range('A1:H1', 'ATRIBUTOS DE ESPACIOS Y ZONAS', text_bold)
        attri_6 = []
        enti_6 = ['IfcSpace', 'IfcZone, IfcSpatialZone']
        for e in entidades:
            if e.is_a() in enti_6:
                attri_6.append(e.get_info())
        dfAtri_6 = pd.DataFrame(attri_6).drop_duplicates()
        if len(dfAtri_6) > 0:
            dfA_6 = dfAtri_6.drop(columns=['OwnerHistory', 'ObjectPlacement', 'Representation'], axis=1)
            dfA_6 = dfA_6.fillna(value='No disponible')
            row = 1
            col = 0

            for c in dfA_6.columns:
                r_atributos_6.write(row, col, c, text_bold)
                col += 1

            row = 2
            col = 0
            x = 0
            while x in range(len(dfA_6.columns)):
                for i in range(len(dfA_6)):
                    r_atributos_6.write(row, col, dfA_6.iloc[i, x], data_text)
                    row += 1
                row = 2
                x += 1
                col += 1
        else:
            row = 1
            col = 0
            r_atributos_6.write(row, col, 'Este proyecto no contiene espacios ni zonas.', data_text)

        # Crear pestaña Atributos_7
        r_atributos_7 = reporte.add_worksheet('ATRIBUTOS_7')
        r_atributos_7.set_column('A:A', 6)
        r_atributos_7.set_column('B:B', 25)
        r_atributos_7.set_column('C:C', 30)
        r_atributos_7.set_column('D:D', 50)
        r_atributos_7.set_column('E:E', 15)
        r_atributos_7.set_column('F:F', 50)
        r_atributos_7.set_column('G:G', 10)
        r_atributos_7.set_column('H:H', 15)
        r_atributos_7.set_column('I:I', 15)
        r_atributos_7.set_column('J:J', 15)
        r_atributos_7.set_column('K:K', 25)
        r_atributos_7.set_column('L:L', 25)

        r_atributos_7.merge_range('A1:H1', 'ATRIBUTOS DE ELEMENTOS GEOGRÁFICOS Y CIVILES', text_bold)
        attri_7 = []
        enti_7 = ['IfcGeographicElement', 'IfcCivilElement']
        for e in entidades:
            if e.is_a() in enti_7:
                attri_7.append(e.get_info())
        dfAtri_7 = pd.DataFrame(attri_7).drop_duplicates()
        if len(dfAtri_7) > 0:
            dfA_7 = dfAtri_7.drop(columns=['OwnerHistory', 'ObjectPlacement', 'Representation'], axis=1)
            dfA_7 = dfA_7.fillna(value='No disponible')
            row = 1
            col = 0

            for c in dfA_7.columns:
                r_atributos_7.write(row, col, c, text_bold)
                col += 1

            row = 2
            col = 0
            x = 0
            while x in range(len(dfA_7.columns)):
                for i in range(len(dfA_7)):
                    r_atributos_7.write(row, col, dfA_7.iloc[i, x], data_text)
                    row += 1
                row = 2
                x += 1
                col += 1
        else:
            row = 1
            col = 0
            r_atributos_7.write(row, col, 'Este proyecto no contiene Elementos civiles ni geográficos.', data_text)


        # Crear pestaña Propiedades
        r_propiedades = reporte.add_worksheet('PROPIEDADES')
        r_propiedades.set_column('A:A', 25)
        r_propiedades.set_column('B:B', 30)
        r_propiedades.set_column('C:C', 60)
        r_propiedades.set_column('D:D', 25)
        r_propiedades.set_column('E:E', 40)
        r_propiedades.set_column('F:F', 40)

        r_propiedades.merge_range('A1:F1', 'PROPIEDADES DE LOS ELEMENTOS DEL MODELO', text_bold)


        psets = []
        for e in entidades:
            sets = e.IsDefinedBy
            for set in sets:
                if set.is_a('IfcRelDefinesByProperties'):
                    related_data = set.RelatingPropertyDefinition
                    if related_data.is_a('IfcPropertySet'):
                        for data in related_data.HasProperties:
                            if data.is_a('IfcPropertySingleValue'):
                                if data.NominalValue is None:
                                    pass
                                else:
                                    psets.append((e.GlobalId, e.is_a(), e.Name, related_data.Name, data.Name, data.NominalValue.wrappedValue))
                            if data.is_a('IfcPropertyEnumeratedValue'):
                                psets.append((e.GlobalId, e.is_a(), e.Name, related_data.Name, data.Name,
                                                data.EnumerationValues[0].wrappedValue))
        r_propiedades.write('A2', 'GLOBAL ID', text_bold)
        r_propiedades.write('B2', 'IFC ENTITY', text_bold)
        r_propiedades.write('C2', 'NAME', text_bold)
        r_propiedades.write('D2', 'PSET_NAME', text_bold)
        r_propiedades.write('E2', 'PROPERTY', text_bold)
        r_propiedades.write('F2', 'VALUE', text_bold)

        row = 2
        col = 0

        for i, a, n, p, d, v in psets:
            r_propiedades.write(row, col, i, data_text)
            r_propiedades.write(row, col+1, a, data_text)
            r_propiedades.write(row, col+2, n, data_text)
            r_propiedades.write(row, col + 3, p, data_text)
            r_propiedades.write(row, col+4, d, data_text)
            r_propiedades.write(row, col+5, v, data_text)
            row +=1

        # Crear pestaña Cuantías
        r_cuantias = reporte.add_worksheet('CUANTÍAS')
        r_cuantias.set_column('A:A', 25)
        r_cuantias.set_column('B:B', 30)
        r_cuantias.set_column('C:C', 60)
        r_cuantias.set_column('D:D', 25)
        r_cuantias.set_column('E:E', 40)

        r_cuantias.merge_range('A1:E1', 'CUANTÍAS DE LOS ELEMENTOS DEL MODELO', text_bold)
        quant = []
        for e in entidades:
            sets = e.IsDefinedBy
            for set in sets:
                if set.is_a('IfcRelDefinesByProperties'):
                    related_data = set.RelatingPropertyDefinition
                    if related_data.is_a('IfcElementQuantity'):
                        for q in set.RelatingPropertyDefinition.Quantities:
                            quant.append((e.GlobalId, e.is_a(), e.Name, q.Name, q[3]))

        r_cuantias.write('A2', 'GLOBAL ID', text_bold)
        r_cuantias.write('B2', 'IFC ENTITY', text_bold)
        r_cuantias.write('C2', 'NAME', text_bold)
        r_cuantias.write('D2', 'QUANTITY', text_bold)
        r_cuantias.write('E2', 'VALUE', text_bold)

        row = 2
        col = 0

        for i, a, n, d, v in quant:
            r_cuantias.write(row, col, i, data_text)
            r_cuantias.write(row, col+1, a, data_text)
            r_cuantias.write(row, col+2, n, data_text)
            r_cuantias.write(row, col+3, d, data_text)
            r_cuantias.write(row, col+4, v, data_text)
            row +=1

        reporte.close()
        nom_modelo = ''.join(x for x in proy.Name if x.isalnum())
        nombre_repo = (nom_modelo + '_reporte.xlsx')
        ruta_repo = os.path.join(directorio, nombre_repo)
        os.startfile(ruta_repo)