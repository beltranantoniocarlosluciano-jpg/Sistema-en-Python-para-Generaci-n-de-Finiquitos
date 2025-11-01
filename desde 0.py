import pandas as pd
import datetime as dt
import os
from num2words import num2words

#CARGA DE LA BASE DE DATOS
base_datos_fijos = pd.read_excel(r"C:\Users\beltr\OneDrive\Documentos\YPFB\BASES DE DATOS MADRE\OFICIAL BASE DE DATOS (PERSONAL INDEFINIDO Y FIJO).xlsx", sheet_name = "BASE DE DATOS DE FIJOS", header = 0, index_col = 0)
print('---Para generar 1 o mas finiquitos escriba "ALGUNOS"\n---Para generar todos los finiquitos excepto algunos escriba "EXCEPTO"\n---Para generar todos los finiquitos de la base de datos escriba "TODOS"\nEscriba "SALIR" para salir')

#CREACION DE COLUMNAS EXTRA EN EL DATAFRAME DIFERENCIANDO A LOS CASOS EN QUE NO SE TIENE APELLIDO PATERNO PARA EVITAR ESPACIADOS NO DESEADOS

#CREACION DE NOMBRE APELLIDO
for ind_algunos, series_algunos in base_datos_fijos.iterrows():

    #UNION DE NOMBRES COMPLETOS QUE TIENEN DOS APELLIDOS
    if pd.isna(series_algunos['APELLIDO PATERNO']):
        base_datos_fijos.loc[ind_algunos, 'NOMBRES APELLIDOS'] = str(series_algunos['NOMBRES']).strip() + " " + str(series_algunos['APELLIDO MATERNO']).strip()
    
    #UNION DE NOMBRES CON SOLO UN APELLIDO
    else:
        base_datos_fijos.loc[ind_algunos, 'NOMBRES APELLIDOS'] = str(series_algunos['NOMBRES']).strip() + ' ' + str(series_algunos['APELLIDO PATERNO']).strip() + ' ' + str(series_algunos['APELLIDO MATERNO']).strip()
#CREACION DE APELLIDO NOMBRE
for ind_algunos2, series_algunos2 in base_datos_fijos.iterrows():

    #PARA LO QUE NO TIENE  APELLIDO PATERNO
    if pd.isna(series_algunos2['APELLIDO PATERNO']):
        base_datos_fijos.loc[ind_algunos2, 'APELLIDOS NOMBRES'] = str(series_algunos2['APELLIDO MATERNO']).strip() + ' ' + str(series_algunos2['NOMBRES']).strip()
    else:
        base_datos_fijos.loc[ind_algunos2, 'APELLIDOS NOMBRES'] = str(series_algunos2['APELLIDO PATERNO']).strip() + ' ' + str(series_algunos2['APELLIDO MATERNO']).strip() + ' ' + str(series_algunos2['NOMBRES']).strip()

#MENU
while True:
    decision = input('Usted: ').strip().upper()
    
    #ELECCIONEs
    
    if decision == "ALGUNOS":
        #MENU PARA ALGUNOS FINIQUITOS

        print('---Porfavor escriba los nombres completos "APELLIDO NOMBRE" para la generacion de los finiquitos, escriba "LISTO" para seguir')
        #CREACION DE LISTAS PARA Y DICCIONARIOS PARA MODIFICAR EL DATAFRAME

        lista_nombres_algunos = []
        while True:
            nombres_algunos = input('Usted: ').upper().strip()
            #VERIFICACION DE QUE EL PERSONAL ESTA EN LA BASE DE DATOS

            if nombres_algunos in base_datos_fijos['APELLIDOS NOMBRES'].values:
                #CODIGO QUE NO PERMITE REPETIDOS EN LA LISTA PARA LOS FINIQUITOS

                if nombres_algunos not in lista_nombres_algunos: 
                    lista_nombres_algunos.append(nombres_algunos)
                else:
                    print(f'---El personal "{nombres_algunos}" ya se encuentra registrado para los finiquitos.')
                continue
            #PASAR A LA PRESENTACION DE DATOS

            elif nombres_algunos == "LISTO":
                break
            
            else:
                print(f'---El personal "{nombres_algunos}" no se encontro en la base de datos, verificar porfavor.')
        #PRESENTACION DE DATOS DE LOS FINIQUITOS PARA LA MOFIFICACION O GENERACION
        
        for nombre_personal in lista_nombres_algunos:
            #AGREGACION DE MESES A LAS COLUMNAS DEL DATAFRAME
            orden_meses = {'ENERO': 1,
            'FEBRERO': 2,
            'MARZO': 3,
            'ABRIL': 4,
            'MAYO': 5,
            'JUNIO': 6,
            'JULIO': 7,
            'AGOSTO': 8,
            'SEPTIEMBRE': 9,
            'OCTUBRE': 10,
            'NOVIEMBRE': 11,
            'DICIEMBRE': 12}
            meses_algunos = []
            print(f'---Porfavor mencione los 3 meses de los que se pagara al personal {nombre_personal}')
            for i in range(3):
                meses_algunos_ =input('Usted: ').upper().strip()
                meses_algunos.append(meses_algunos_)

            meses_ordenados_algunos = sorted(meses_algunos, key = lambda mes: orden_meses[mes])

            # Crear (si no existen) las columnas de todos los meses en el DataFrame
            for mes_col in orden_meses.keys():
                if mes_col not in base_datos_fijos.columns:
                    base_datos_fijos[mes_col] = pd.NA

            # Asignar los valores de los 3 meses seleccionados: preguntar primero si hay variación en algún mes
            mask_persona = base_datos_fijos['APELLIDOS NOMBRES'] == nombre_personal

            # Limpiar meses no seleccionados para esta persona
            meses_todos = list(orden_meses.keys())
            meses_no_seleccionados = [m for m in meses_todos if m not in meses_ordenados_algunos]
            if meses_no_seleccionados:
                base_datos_fijos.loc[mask_persona, meses_no_seleccionados] = pd.NA

            # Preguntar si hubo vacaciones o algún descuento en algún mes
            variacion = input('Hubo vacaciones o algun decuento en su sueldo en algun mes? (SI/NO): ').upper().strip()

            valores_meses = {}
            if variacion == 'NO':
                # Asignar el sueldo base a los tres meses seleccionados
                sueldo_base = base_datos_fijos.loc[mask_persona, 'SUELDO'].iloc[0]
                try:
                    sueldo_base_num = float(str(sueldo_base).replace(',', '.'))
                except ValueError:
                    sueldo_base_num = 0.0
                for mes_sel in meses_ordenados_algunos:
                    valores_meses[mes_sel] = sueldo_base_num
            else:
                # Pedir el sueldo específico para cada mes seleccionado y asignar
                for mes_sel in meses_ordenados_algunos:
                    while True:
                        valor_ingresado = input(f'Cual seria el sueldo de la persona {nombre_personal} para el mes de {mes_sel}?: ').strip()
                        # Normalizar separadores decimales
                        valor_normalizado = valor_ingresado.replace(' ', '').replace(',', '.')
                        try:
                            valor_float = float(valor_normalizado)
                            break
                        except ValueError:
                            print('---Valor no válido. Ingrese un número (ejemplo: 2500.50 o 2500,50).')
                    valores_meses[mes_sel] = valor_float

            for mes_sel, val in valores_meses.items():
                base_datos_fijos.loc[mask_persona, mes_sel] = val

            # Calcular TOTAL y PROMEDIO para esta persona
            total_valor = float(sum(valores_meses.get(m, 0.0) for m in meses_ordenados_algunos))
            base_datos_fijos.loc[mask_persona, 'TOTAL'] = total_valor
            base_datos_fijos.loc[mask_persona, 'PROMEDIO'] = (total_valor / 3.0)

            # Pedir correlativos por persona y guardarlos en el DataFrame
            correlativo_finiquito = input('Ingrese el correlativo de finiquito: ').strip()
            correlativo_informe = input('Ingrese el correlativo de informe tecnico: ').strip()
            base_datos_fijos.loc[mask_persona, 'CORRELATIVO DE FINIQUITO'] = correlativo_finiquito
            base_datos_fijos.loc[mask_persona, 'CORRELATIVO DE INFORME TECNICO'] = correlativo_informe

            print(f'----------{nombre_personal}----------\nCargo actual: {base_datos_fijos.loc[base_datos_fijos["APELLIDOS NOMBRES"] == nombre_personal,
            'CARGO ACTUAL'].iloc[0]}\nDireccion de domicilio: {base_datos_fijos.loc[base_datos_fijos['APELLIDOS NOMBRES'] == nombre_personal, 
            'DIRECCION DE DOMICILIO'].iloc[0]}\nEstado Civil: {base_datos_fijos.loc[base_datos_fijos['APELLIDOS NOMBRES'] == nombre_personal, 
            'ESTADO CIVIL'].iloc[0]}\nFecha de ingreso: {base_datos_fijos.loc[base_datos_fijos['APELLIDOS NOMBRES'] == nombre_personal, 
            'FECHA DE INGRESO'].iloc[0].strftime('%d/%m/%Y')}\nFecha de salida: {base_datos_fijos.loc[base_datos_fijos['APELLIDOS NOMBRES'] == nombre_personal, 
            'FECHA DE SALIDA'].iloc[0].strftime('%d/%m/%Y')}\nSueldo: {base_datos_fijos.loc[base_datos_fijos['APELLIDOS NOMBRES'] == nombre_personal, 
            'SUELDO'].iloc[0]}')

            #ESTRUCUTURA PARA LA MODIFICACION DE COLUMNAS
            while True:
                print('---Desea modificar algun dato mostrado?')
                modificacion_algunos = input(f'Usted: ').upper().strip()

                if modificacion_algunos == 'CARGO ACTUAL':
                    cargo_algunos = input(f'Cual seria el cargo actual del personal {nombre_personal}?: ').upper().strip()
                    base_datos_fijos.loc[base_datos_fijos['APELLIDOS NOMBRES'] == nombre_personal, 'CARGO ACTUAL'] = cargo_algunos

                elif modificacion_algunos == 'DIRECCION DE DOMICILIO':
                    direccion_algunos = input(f'Cual seria la nueva direccion de domicilio del personal {nombre_personal}?: ').upper().strip()
                    base_datos_fijos.loc[base_datos_fijos['APELLIDOS NOMBRES'] == nombre_personal, 'DIRECCION DE DOMICILIO'] = direccion_algunos

                elif modificacion_algunos == 'ESTADO CIVIL':
                    estado_civil_algunos = input(f'Cual seria el estado civil del personal {nombre_personal}?: ' ).upper().strip()
                    base_datos_fijos.loc[base_datos_fijos['APELLIDOS NOMBRES'] == nombre_personal, 'ESTADO CIVIL'] = estado_civil_algunos
                
                elif modificacion_algunos == 'FECHA DE INGRESO':
                    try:
                        base_datos_fijos.loc[base_datos_fijos['APELLIDOS NOMBRES'] == nombre_personal, 'FECHA DE INGRESO'] = dt.datetime.strptime(input(f'Cual seria la fecha de ingreso en esta gestion del personal {nombre_personal}?: ').strip(),
                        '%d/%m/%Y')
                    except ValueError:
                        print('---Formato de fecha no valida')

                elif modificacion_algunos == 'FECHA DE SALIDA':
                    try:
                        base_datos_fijos.loc[base_datos_fijos['APELLIDOS NOMBRES'] == nombre_personal, 'FECHA DE SALIDA'] = dt.datetime.strptime(input(f'Cual seria la fecha de salida en esta gestion del personal {nombre_personal}?: ').strip(), 
                        '%d/%m/%Y')
                    except ValueError:
                        print('---Formato de fecha no valida')

                elif modificacion_algunos == 'SUELDO':
                    sueldo_algunos = input(f'Cuanto seria el sueldo del personal {nombre_personal}?: ')
                    base_datos_fijos.loc[base_datos_fijos['APELLIDOS NOMBRES'] == nombre_personal, 'SUELDO'] = sueldo_algunos
                
                #GENERACION DE FINIQUITOS
                elif modificacion_algunos == 'NO':
                    import xlwings as xw
                    apertura_excel_en_blanco = xw.App(visible=True)
                    ruta_plantilla = r'C:\\Users\\beltr\\OneDrive\\Documentos\\YPFB\\PLANTILLAS PARA AUTOMATIZAR\\PLANTILLA EXCEL OFICIAL.xlsx'

                    base_datos_filtrada = base_datos_fijos[base_datos_fijos['APELLIDOS NOMBRES'].isin(lista_nombres_algunos)]
                    for i, fila in base_datos_filtrada.iterrows():
                        plantilla_excel = apertura_excel_en_blanco.books.open(ruta_plantilla)
                        lado1 = plantilla_excel.sheets['LADO1']
                        lado2 = plantilla_excel.sheets['LADO2']
                        informe_tecnico = plantilla_excel.sheets['INFORME']
                        # LADO 1 - DATOS BASE
                        lado1["S12"].value = "PETROLERA"
                        lado1["AB12"].value = "RAMON RIVERO N° 604 ESQ PAPA PAULO"
                        lado1["B25"].value = "OTROS CONCEPTOS PERCIBIDOS"
                        lado1["B26"].value = "B. Antigüedad"
                        lado1["M26"].value = 0
                        lado1["T26"].value = 0
                        lado1["AA26"].value = 0
                        lado1["AH26"].value = 0
                        lado1["B36"].value = "AGUINALDO"
                        lado1["U36"].value = fila["MESES"]
                        lado1["AB36"].value = fila["DIAS"]
                        lado1["B37"].value = "VACACIONES"
                        lado1["AH9"].value = fila.get("CORRELATIVO DE FINIQUITO")
                        lado1["S13"].value = fila.get("NOMBRES APELLIDOS")
                        lado1["J14"].value = fila.get("ESTADO CIVIL")
                        lado1["W14"].value = fila.get("EDAD")
                        lado1["AD14"].value = fila.get("DIRECCION DE DOMICILIO")
                        lado1["M15"].value = fila.get("CARGO ACTUAL")
                        lado1["F16"].value = fila.get("N° DE CARNET DE IDENTIDAD")
                        lado1["W16"].value = fila.get("FECHA DE INGRESO")
                        lado1["AI16"].value = fila.get("FECHA DE SALIDA")
                        lado1["K17"].value = fila.get("MOTIVO DEL RETIRO")
                        lado1["AH17"].value = fila.get("SUELDO")
                        lado1["K18"].value = fila.get("AÑOS")
                        lado1["T18"].value = fila.get("MESES")
                        lado1["AB18"].value = fila.get("DIAS")

                        # Determinar los 3 meses con valores (según selección realizada) y ordenarlos
                        orden_meses_local = {'ENERO': 1, 'FEBRERO': 2, 'MARZO': 3, 'ABRIL': 4, 'MAYO': 5, 'JUNIO': 6, 'JULIO': 7, 'AGOSTO': 8, 'SEPTIEMBRE': 9, 'OCTUBRE': 10, 'NOVIEMBRE': 11, 'DICIEMBRE': 12}
                        meses_presentes = [m for m in orden_meses_local.keys() if pd.notna(fila.get(m))]
                        meses_presentes = sorted(meses_presentes, key=lambda m: orden_meses_local[m])[:3]
                        # Encabezados de los meses seleccionados
                        lado1["L22"].value = meses_presentes[0] if len(meses_presentes) > 0 else ""
                        lado1["S22"].value = meses_presentes[1] if len(meses_presentes) > 1 else ""
                        lado1["Z22"].value = meses_presentes[2] if len(meses_presentes) > 2 else ""

                        # Valores por mes
                        lado1["M23"].value = fila.get(meses_presentes[0]) if len(meses_presentes) > 0 else 0
                        lado1["T23"].value = fila.get(meses_presentes[1]) if len(meses_presentes) > 1 else 0
                        lado1["AA23"].value = fila.get(meses_presentes[2]) if len(meses_presentes) > 2 else 0
                        lado1["AH23"].value = fila.get("TOTAL")
                        lado1["M29"].value = fila.get(meses_presentes[0]) if len(meses_presentes) > 0 else 0
                        lado1["T29"].value = fila.get(meses_presentes[1]) if len(meses_presentes) > 1 else 0
                        lado1["AA29"].value = fila.get(meses_presentes[2]) if len(meses_presentes) > 2 else 0
                        lado1["AH29"].value = fila.get("TOTAL")
                        lado1["AH30"].value = fila.get("PROMEDIO")

                        # Desahucio (actualmente desactivado)
                        lado1["AI32"].value = 0

                        # Tiempo de servicio
                        lado1["U33"].value = fila.get("AÑOS")
                        lado1["U34"].value = fila.get("MESES")
                        lado1["U35"].value = fila.get("DIAS")


                        # Indemnización
                        promedio = fila.get("PROMEDIO") or 0
                        anos = fila.get("AÑOS") or 0
                        meses = fila.get("MESES") or 0
                        dias = fila.get("DIAS") or 0
                        lado1["AC33"].value = promedio * anos
                        lado1["AC34"].value = round((promedio/12) * meses, 2)
                        lado1["AC35"].value = round((promedio/360) * dias, 2)

                        lado1["AI33"].value = lado1["AC33"].value
                        lado1["AI34"].value = lado1["AC34"].value
                        lado1["AI35"].value = lado1["AC35"].value

                        # Aguinaldo y Vacaciones
                        lado1["AI36"].value = (promedio/360) * ((meses * 30) + dias)
                        lado1["AB37"].value = fila.get("VACACIONES") or 0
                        lado1["AI37"].value = (promedio/30) * (fila.get("VACACIONES") or 0)

                        # Total beneficios sociales
                        lado1["AI41"].value = (
                            (lado1["AI32"].value or 0) +
                            (lado1["AI33"].value or 0) +
                            (lado1["AI34"].value or 0) +
                            (lado1["AI35"].value or 0) +
                            (lado1["AI36"].value or 0) +
                            (lado1["AI37"].value or 0)
                        )

                        # Importe líquido a pagar
                        lado1["AI47"].value = lado1["AI41"].value
                        valor_total = round((lado1["AI47"].value or 0), 2)
                        parte_entera = int(valor_total)
                        decimales = int(round((valor_total - parte_entera) * 100, 0))
                        lado2["C6"].value = f"SON: {num2words(parte_entera, lang='es').upper()} {decimales:02d}/100 BOLIVIANOS"
                        # Fecha de impresión en LADO2
                        hoy = pd.Timestamp.today()
                        meses_es_mayus = {1: 'ENERO', 2: 'FEBRERO', 3: 'MARZO', 4: 'ABRIL', 5: 'MAYO', 6: 'JUNIO', 7: 'JULIO', 8: 'AGOSTO', 9: 'SEPTIEMBRE', 10: 'OCTUBRE', 11: 'NOVIEMBRE', 12: 'DICIEMBRE'}
                        lado2["O15"].value = hoy.day
                        lado2["S15"].value = meses_es_mayus[hoy.month]
                        lado2["AA15"].value = hoy.year

                        # INFORME - Fechas y datos (mes en español)
                        fecha_ingreso = pd.to_datetime(fila.get("FECHA DE INGRESO"), errors='coerce')
                        meses_es = {1: 'enero', 2: 'febrero', 3: 'marzo', 4: 'abril', 5: 'mayo', 6: 'junio', 7: 'julio', 8: 'agosto', 9: 'septiembre', 10: 'octubre', 11: 'noviembre', 12: 'diciembre'}
                        if pd.notna(fecha_ingreso):
                            informe_tecnico["J126"].value = f"{fecha_ingreso.day} de {meses_es[fecha_ingreso.month]} del {fecha_ingreso.year}"
                        else:
                            informe_tecnico["J126"].value = "Fecha inválida"

                        fecha_salida = pd.to_datetime(fila.get("FECHA DE SALIDA"), errors='coerce')
                        if pd.notna(fecha_salida):
                            informe_tecnico["T126"].value = f"{fecha_salida.day} de {meses_es[fecha_salida.month]} del {fecha_salida.year}"
                        else:
                            informe_tecnico["T126"].value = "Fecha inválida"

                        informe_tecnico["V230"].value = f"{fila.get('APELLIDO PATERNO','')} {fila.get('APELLIDO MATERNO','')} {fila.get('NOMBRES','')}".strip()
                        informe_tecnico["V231"].value = fila.get("N° DE CARNET DE IDENTIDAD")
                        informe_tecnico["V232"].value = fila.get("CARGO ACTUAL")
                        informe_tecnico["V233"].value = informe_tecnico["J126"].value
                        informe_tecnico["V234"].value = informe_tecnico["T126"].value
                        informe_tecnico["V235"].value = fila.get("MOTIVO DEL RETIRO")
                        informe_tecnico["V236"].value = anos
                        informe_tecnico["AB236"].value = meses
                        informe_tecnico["AH236"].value = dias
                        informe_tecnico["AA240"].value = fila.get("SUELDO")
                        informe_tecnico["AA243"].value = promedio
                        informe_tecnico["AA245"].value = round((lado1["AC33"].value or 0) + (lado1["AC34"].value or 0) + (lado1["AC35"].value or 0), 2)
                        informe_tecnico["AA246"].value = lado1["AI36"].value
                        informe_tecnico["AA248"].value = lado1["AI47"].value
                        # INFORME - Fecha de impresión (W86) con mes en español
                        hoy_impresion = pd.Timestamp.today()
                        informe_tecnico["W86"].value = f"{hoy_impresion.day} de {meses_es[hoy_impresion.month]} del {hoy_impresion.year}"
                        # INFORME - Correlativo y nombres
                        informe_tecnico["N43"].value = fila.get("CORRELATIVO DE INFORME TECNICO")
                        informe_tecnico["K124"].value = fila.get("NOMBRES APELLIDOS")
                        informe_tecnico["L136"].value = fila.get("NOMBRES APELLIDOS")

                        # Eliminar hojas no usadas (si existieran)
                        for hoja in list(plantilla_excel.sheets):
                            if hoja.name not in ["LADO1", "LADO2", "INFORME"]:
                                hoja.delete()

                        # Guardar con nombre Persona-finiquito.xlsx en la misma carpeta de la plantilla
                        nombre_persona = (fila.get("NOMBRES APELLIDOS") or "finiquito").strip()
                        # sanitizar caracteres inválidos en Windows
                        for ch in '<>:"/\\|?*':
                            nombre_persona = nombre_persona.replace(ch, '')
                        ruta_salida = os.path.join(os.path.dirname(ruta_plantilla), f"{nombre_persona}-finiquito.xlsx")
                        plantilla_excel.save(ruta_salida)
                        plantilla_excel.close()
                    apertura_excel_en_blanco.quit()

                    print("listo")
                    break
                else:
                    print(f'El dato "{modificacion_algunos}" no es reconocido')

    elif decision == "EXCEPTO":
        pass

    elif decision == "TODOS":
        pass
    
    elif decision == "SALIR":
        break

    else:
        print(f'===Opcion no valida "{decision}"')