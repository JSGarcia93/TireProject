import pandas as pd
import numpy as np

path = "C:/Users/jonnathan.garcia/Documents/Llantas/"
x  = "Data Ruedata.xlsx"
x1 = "llantas.xlsx"
x2 = "Digitar Desmonte.xlsx"
#dfs: Usuarios, Base Flota, Bodegas, Base de datos llantas y Digitación de movimientos
dfu = pd.read_excel(path+x,sheet_name="Usuarios")
dff = pd.read_excel(path+x,sheet_name="Flota")
dfb = pd.read_excel(path+x,sheet_name="Bodegas")
dfdll = pd.read_excel(path+x1, sheet_name="pyexcel_sheet1", names=['Posición', 'Móvil', 'Placa', 'Ubicación','Código', 'Marca', 'Modelo', 'Dimensión', 'P.exterior','P.centro', 'P.interior', 'Vida', 'Regrabada', 'Marca_reencauche','Banda_reencauche', 'Dimensión_reencauche', 'Catálogo_nuevas','Catálogo_reencauche', 'Fecha_Inspeccion'])
dfdd = pd.read_excel(path+x2, sheet_name="Desmonte", names=['Código', 'Móvil', 'Prof._Exterior','Prof._Centro', 'Prof._Interior', 'Fecha_Inspeccion', 'Actividad', 'Técnico', "OT"])

dfdll = dfdll.set_index(["Código"]) #Base de datos llantas
dfdd1 = dfdd.set_index(["Código"]) #Archivo Digitar
dff = dff.set_index(["Móvil"]) #Base Flota
dfdd2 = dfdd.set_index(["Móvil"]) #Archivo Digitar

#Cruce de datos llantas Vs Digitado
Cruc = pd.merge(dfdd1, dfdll, how="left",right_index=True, left_index=True)

#Asignación de index flota para cruce
Cruc = Cruc.reset_index()
Cruc = Cruc.set_index(["Móvil_x"])

#Cruce de flota Vs Digitado, sumado al cruce entre las dos bases de datos para unificar información. Eliminación de columna duplicada
Cruc2 = pd.merge(dfdd2, dff, how="left",right_index=True, left_index=True)
Cruc3 = pd.merge(Cruc, Cruc2, how="left")
Cruc3.drop("Fecha_Inspeccion",axis=1, inplace=True)

#Asignación de columnas de validación
Crucx = Cruc3.reindex(columns=Cruc3.columns.tolist() +["Fecha_Ultima_Inspección", "Vehiculo","Ultima_Profundidand_Exterior","Ultima_Profundidand_Centro","Ultima_Profundidand_Interior"])

#Asignación de variables para posición de Columnas
Fec1 = Crucx.columns.get_loc("Fecha_Inspeccion_x")
Fec2 = Crucx.columns.get_loc("Fecha_Inspeccion_y")
Fec3 = Crucx.columns.get_loc('Fecha_Ultima_Inspección')
Pl1 = Crucx.columns.get_loc('Placa')
Pl2 = Crucx.columns.get_loc("Activo")
Pl3 = Crucx.columns.get_loc("Vehiculo")
Pe1 = Crucx.columns.get_loc("Prof._Exterior")
Pe2 = Crucx.columns.get_loc("P.exterior")
Pe3 = Crucx.columns.get_loc("Ultima_Profundidand_Exterior")
Pi1 = Crucx.columns.get_loc("Prof._Interior")
Pi2 = Crucx.columns.get_loc("P.interior")
Pi3 = Crucx.columns.get_loc("Ultima_Profundidand_Interior")
Pc1 = Crucx.columns.get_loc("Prof._Centro")
Pc2 = Crucx.columns.get_loc("P.centro")
Pc3 = Crucx.columns.get_loc("Ultima_Profundidand_Centro")

#Validación Fechas
for n in range(0, np.shape(Crucx)[0]):
    if Crucx.iloc[n, Fec1] >= Crucx.iloc[n, Fec2]:
        Crucx.iloc[n, Fec3] = Crucx.iloc[n, Fec1].strftime("%Y-%m-%d")
    else:
        Crucx.iloc[n, Fec3] = "Fecha Mayor Validar: {}".format(Crucx.iloc[n, Fec1].strftime("%Y-%m-%d"))
#Validación Placa
for n in range(0, np.shape(Crucx)[0]):
    if Crucx.iloc[n, Pl1] == Crucx.iloc[n, Pl2]:
        Crucx.iloc[n, Pl3] = Crucx.iloc[n, Pl2]
    else:
        Crucx.iloc[n, Pl3] = "No Coincide Placa: {}".format(Crucx.iloc[n, Pl2])
#Validación Profundidades Externo e Interno
for n in range(0, np.shape(Crucx)[0]):
    if Crucx.iloc[n, Pe1] <= Crucx.iloc[n, Pe2] and Crucx.iloc[n, Pi1] <= Crucx.iloc[n, Pi2]:
        Crucx.iloc[n, Pe3] = Crucx.iloc[n, Pe1]
        Crucx.iloc[n, Pi3] = Crucx.iloc[n, Pi1]
        
    elif Crucx.iloc[n, Pe1] <= Crucx.iloc[n, Pe2] and Crucx.iloc[n, Pi1] >= Crucx.iloc[n, Pi2]:
        if Crucx.iloc[n, Pi1] <= Crucx.iloc[n, Pe2] and Crucx.iloc[n, Pe1] <= Crucx.iloc[n, Pi2]:
            Crucx.iloc[n, Pe3] = Crucx.iloc[n, Pi1]
            Crucx.iloc[n, Pi3] = Crucx.iloc[n, Pe1]
        else:
            if (Crucx.iloc[n, Pe1] - Crucx.iloc[n, Pe2]).round(2).__abs__() <= 0.9 and (Crucx.iloc[n, Pi1] - Crucx.iloc[n, Pi2]).round(2).__abs__() <= 0.9:
                Crucx.iloc[n, Pe3] = Crucx.iloc[n, Pe1]
                Crucx.iloc[n, Pi3] = Crucx.iloc[n, Pi2]
            else:
                Crucx.iloc[n, Pe3] = "Validar: {}".format(Crucx.iloc[n, Pe1])
                Crucx.iloc[n, Pi3] = "Validar: {}".format(Crucx.iloc[n, Pi1])
    elif Crucx.iloc[n, Pe1] >= Crucx.iloc[n, Pe2] and Crucx.iloc[n, Pi1] <= Crucx.iloc[n, Pi2]:
        if Crucx.iloc[n, Pi1] <= Crucx.iloc[n, Pe2] and Crucx.iloc[n, Pe1] <= Crucx.iloc[n, Pi2]:
            Crucx.iloc[n, Pe3] = Crucx.iloc[n, Pi1]
            Crucx.iloc[n, Pi3] = Crucx.iloc[n, Pe1]
        else:
            if (Crucx.iloc[n, Pe1] - Crucx.iloc[n, Pe2]).round(2).__abs__() <= 0.9 and (Crucx.iloc[n, Pi1] - Crucx.iloc[n, Pi2]).round(2).__abs__() <= 0.9:
                Crucx.iloc[n, Pe3] = Crucx.iloc[n, Pe2]
                Crucx.iloc[n, Pi3] = Crucx.iloc[n, Pi1]
            else:
                Crucx.iloc[n, Pe3] = "Validar: {}".format(Crucx.iloc[n, Pe1])
                Crucx.iloc[n, Pi3] = "Validar: {}".format(Crucx.iloc[n, Pi1])
    elif Crucx.iloc[n, Pe1] >= Crucx.iloc[n, Pe2] and Crucx.iloc[n, Pi1] >= Crucx.iloc[n, Pi2]:
        if Crucx.iloc[n, Pi1] <= Crucx.iloc[n, Pe2] and Crucx.iloc[n, Pe1] <= Crucx.iloc[n, Pi2]:
            Crucx.iloc[n, Pe3] = Crucx.iloc[n, Pi1]
            Crucx.iloc[n, Pi3] = Crucx.iloc[n, Pe1]
        else:
            if (Crucx.iloc[n, Pe1] - Crucx.iloc[n, Pe2]).round(2).__abs__() <= 0.9 and (Crucx.iloc[n, Pi1] - Crucx.iloc[n, Pi2]).round(2).__abs__() <= 0.9:
                Crucx.iloc[n, Pe3] = Crucx.iloc[n, Pe2]
                Crucx.iloc[n, Pi3] = Crucx.iloc[n, Pi2]
            else:
                Crucx.iloc[n, Pe3] = "Validar: {}".format(Crucx.iloc[n, Pe1])
                Crucx.iloc[n, Pi3] = "Validar: {}".format(Crucx.iloc[n, Pi1])
#Profundidad Central
for n in range(0, np.shape(Crucx)[0]):
    if Crucx.iloc[n, Pc1] <= Crucx.iloc[n, Pc2]:
        Crucx.iloc[n, Pc3] = Crucx.iloc[n, Pc1]
    else:
        if (Crucx.iloc[n, Pc1] - Crucx.iloc[n, Pc2]).round(2).__abs__() <= 0.9:
            Crucx.iloc[n, Pc3] = Crucx.iloc[n, Pc2]
        else:
            Crucx.iloc[n, Pc3] = "Validar: {}".format(Crucx.iloc[n, Pc1])

#Filtrado de Data
final=Crucx.loc[:,["Fecha_Ultima_Inspección", "Vehiculo","Ultima_Profundidand_Exterior","Ultima_Profundidand_Centro","Ultima_Profundidand_Interior", 'Actividad','Código',"Técnico", "OT"]]
#Adición Columnas
final=final.reindex(columns=final.columns.tolist() +["Hora","Kilometros","Ultima_Presión","Observaciones","Observaciones_Analista"])
#Organización Columnas
final=final[['Fecha_Ultima_Inspección',"Hora",'Vehiculo',"Kilometros","Actividad","Código",'Ultima_Profundidand_Exterior','Ultima_Profundidand_Centro', 'Ultima_Profundidand_Interior',"Ultima_Presión","OT","Observaciones","Observaciones_Analista","Técnico"]]
#Renombrar Columnas
final.columns=['Fecha Ultima Inspección',"Hora",'Vehiculo',"Kilometros","Ubicación","Codigo",'Ultima Profundidand Exterior','Ultima Profundidand Centro', 'Ultima Profundidand Interior',"Ultima Presión","Documento #","Observaciones","Observaciones Analista","Tecnico"]

#Asignación de variables para posición de Columnas
Hr=final.columns.get_loc("Hora")
Bog=final.columns.get_loc("Ubicación")
Tec=final.columns.get_loc("Tecnico")

#Asignación de hora, validación de bodegas y usuarios
hr=str("06:00")
for n in range(0, np.shape(final)[0]):
    final.iloc[n,Hr]=hr
for n in range(0, np.shape(final)[0]):
    if final.iloc[n, Bog] not in dfb.values:
        final.iloc[n, Bog]="Validar"
for n in range(0, np.shape(final)[0]):
    if final.iloc[n, Tec] not in dfu.values:
        final.iloc[n, Tec]="Validar"
        
final.to_excel("C:/Users/jonnathan.garcia/Documents/Python/Desmonte de Llantas/1.Plantilla desmonte.xlsx", index=False)