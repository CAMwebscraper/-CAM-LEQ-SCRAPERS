import time
import pandas as pd
from datetime import datetime
import os
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

class LEQScraper:
    def __init__(self):
        self.driver = None
        self.especialidades_seleccionadas_global = []  # Nueva variable para guardar selección global
        self.urls_disponibles = {
            1: {
                'nombre': 'Lista de Espera Quirúrgica por hospital y procesos-patologías',
                'url': 'https://servicioselectronicos.sanidadmadrid.org/LEQ/ConsultaProcesos.aspx'
            },
            2: {
                'nombre': 'Lista de Espera Quirúrgica por hospital y especialidad',
                'url': 'https://servicioselectronicos.sanidadmadrid.org/LEQ/Consulta.aspx'
            },
            3: {
                'nombre': 'Lista de Espera de Consultas Externas por hospital y especialidad',
                'url': 'https://servicioselectronicos.sanidadmadrid.org/LEQ/ConsultaEspecialidades.aspx'
            },
            4: {
                'nombre': 'Lista de Espera de Pruebas Diagnósticas y Terapéuticas',
                'url': 'https://servicioselectronicos.sanidadmadrid.org/LEQ/ConsultaPruebas.aspx'
            }
        }
    
    def mostrar_menu_urls(self):
        """Muestra menú para elegir URL"""
        
        for key, valor in self.urls_disponibles.items():
            print(f"\n\t{key}. {valor['nombre']}")
            print(f"\t   {valor['url']}\n")
        
        print("\t" + "-"*36)
        
        while True:
            try:
                seleccion = input("\n\t¿Qué URL quieres procesar? (número): ").strip()
                opcion = int(seleccion)
                
                if opcion in self.urls_disponibles:
                    url_elegida = self.urls_disponibles[opcion]
                    print(f"\n\tURL seleccionada: {url_elegida['nombre']}")
                    print(f"\tURL: {url_elegida['url']}")
                    
                    return url_elegida
                else:
                    print("\tOpción no válida. Intenta de nuevo.")
            except ValueError:
                print("\tPor favor, introduce un número válido.")
    
    def seleccionar_ano(self):
        """Permite al usuario seleccionar el año o años a filtrar"""
        
        print("\n\tIntroduzca el año o años a filtrar (separados por comas):")
        print("\tEjemplo: 2025,2024  o  2023,2024,2025")
        
        while True:
            try:
                entrada = input("\n\t- Introduzca el año o años a filtrar: ").strip()
                
                if not entrada:
                    print("\tPor favor, introduce al menos un año.")
                    continue
                
                # Limpiar espacios y dividir por comas
                anos_str = entrada.split(',')
                anos_seleccionados = []
                
                for ano_str in anos_str:
                    # Limpiar espacios en blanco
                    ano_limpio = ano_str.strip()
                    
                    # Verificar que sea un año válido
                    if not ano_limpio.isdigit():
                        print(f"\tError: '{ano_limpio}' no es un año válido.")
                        continue
                    
                    ano = int(ano_limpio)
                    
                    ano_actual = datetime.now().year
                    # Validar rango de años (por ejemplo, 2000-2030)
                    if ano < 2015 or ano > ano_actual:
                        print(f"\tAdvertencia: El año {ano} parece fuera del rango esperado.")
                        print("\t¿Continuar de todos modos? (s/n): ", end="")
                        respuesta = input().strip().lower()
                        if respuesta != 's':
                            continue
                    
                    if ano not in anos_seleccionados:
                        anos_seleccionados.append(ano)
                
                if not anos_seleccionados:
                    print("\tNo se ingresaron años válidos. Intenta de nuevo.")
                    continue
                
                # Ordenar años
                anos_seleccionados.sort()
                
                # Mostrar confirmación
                print(f"\n\t ✓ Años seleccionados: {', '.join(map(str, anos_seleccionados))}")
                
                filtrar = ('1' == '1')
                return anos_seleccionados, filtrar
                    
            except ValueError as e:
                print(f"\tError al procesar los años: {e}")
                print("\tPor favor, introduce años válidos separados por comas (ej: 2024,2025)")
            except Exception as e:
                print(f"\tError inesperado: {e}")
    
    def obtener_especialidades(self, driver):
        """Obtiene la lista de especialidades disponibles"""
        try:
            especialidad_dropdown = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "ContenedorContenidoSeccion_ddlEspecialidad"))
            )
            
            select_especialidad = Select(especialidad_dropdown)
            especialidad_options = select_especialidad.options
            
            especialidades = []
            for i, option in enumerate(especialidad_options):
                if option.get_attribute('value') and option.text.strip():
                    especialidades.append({
                        'indice': i,
                        'nombre': option.text.strip(),
                        'valor': option.get_attribute('value')
                    })
            
            return especialidades
            
        except Exception as e:
            print(f"  ✗ No se encontró lista de especialidades o hubo un error: {e}")
            return []
    
    def mostrar_menu_especialidades(self, especialidades):
        """Muestra menú para elegir especialidades"""
        print("\n\tESPECIALIDADES DISPONIBLES:")
        print("\t" + "-"*40)
        
        # Mostrar las primeras 20 especialidades para no saturar la pantalla
        for i, especialidad in enumerate(especialidades[:20], 1):
            print(f"\t{i:3}. {especialidad['nombre'][:40]}")
        
        if len(especialidades) > 20:
            print(f"\t... y {len(especialidades) - 20} más")
        
        print("\n\tOpciones:")
        print("\t  • Un número (ej: 5)")
        print("\t  • Varios números separados por comas (ej: 1,3,5)")
        print("\t  • Rango (ej: 1-5)")
        print("\t  • 'TODOS' para todas las especialidades")
        print("\t  • 'NINGUNO' para procesar sin filtro de especialidad")
        
        return input("\n\t¿Qué especialidad(es) quieres procesar? ").strip()
    
    def procesar_seleccion_especialidades(self, seleccion, especialidades):
        """Procesa la selección del usuario para especialidades"""
        seleccionadas = []
        
        seleccion_upper = seleccion.upper()
        
        if seleccion_upper == 'TODOS':
            return especialidades
        elif seleccion_upper == 'NINGUNO':
            # Devolver lista vacía para indicar "sin filtro"
            return []
        
        # Rango (ej: 1-5)
        if '-' in seleccion:
            try:
                inicio, fin = map(int, seleccion.split('-'))
                for i in range(inicio, fin + 1):
                    if 1 <= i <= len(especialidades):
                        seleccionadas.append(especialidades[i-1])
            except:
                pass
        
        # Múltiples números (ej: 1,3,5)
        elif ',' in seleccion:
            try:
                numeros = [int(n.strip()) for n in seleccion.split(',')]
                for num in numeros:
                    if 1 <= num <= len(especialidades):
                        seleccionadas.append(especialidades[num-1])
            except:
                pass
        
        # Un solo número
        else:
            try:
                num = int(seleccion)
                if 1 <= num <= len(especialidades):
                    seleccionadas.append(especialidades[num-1])
            except:
                pass
        
        return seleccionadas
    
    def es_mes_del_ano(self, texto_mes, anos_seleccionados, filtrar=True):
        """Determina si un mes corresponde a los años seleccionados"""
        texto = texto_mes.lower()
        
        # Buscar patrón de año
        patron_anio = r'(?:19|20)\d{2}'
        matches = re.findall(patron_anio, texto_mes)
        
        if matches:
            año = int(matches[0])
            if filtrar:
                return año in anos_seleccionados
            else:
                return año  # Devuelve el año para marcarlo
        elif not filtrar:
            # Si no estamos filtrando, devolvemos None para meses sin año
            return None
        
        # Si estamos filtrando y no encuentra año, no lo incluye
        return False
    
    def filtrar_meses(self, meses, anos_seleccionados, filtrar=True):
        """Filtra meses según los años seleccionados"""
        if filtrar:
            # Modo filtrado: solo devuelve meses de años seleccionados
            meses_filtrados = []
            for mes in meses:
                if self.es_mes_del_ano(mes['texto'], anos_seleccionados, filtrar):
                    meses_filtrados.append(mes)
            return meses_filtrados
        else:
            # Modo marcado: devuelve todos los meses con información del año
            meses_marcados = []
            for mes in meses:
                año = self.es_mes_del_ano(mes['texto'], anos_seleccionados, filtrar=False)
                mes['año'] = año
                meses_marcados.append(mes)
            return meses_marcados
    
    def mostrar_menu_hospitales(self, hospitales):
        """Muestra menú para elegir hospital"""
        print("\n" )
        
        # Agrupar por primeros caracteres para mejor visualización
        grupos = {}
        for i, hospital in enumerate(hospitales, 1):
            inicial = hospital['nombre'][:3].upper()
            if inicial not in grupos:
                grupos[inicial] = []
            grupos[inicial].append((i, hospital['nombre']))
        
        # Mostrar hospitales agrupados
        contador = 0
        for inicial in sorted(grupos.keys()):
            for num, nombre in grupos[inicial]:
                print(f"\t  {num:3}. {nombre}")
                contador += 1
        
        print("\n\t"+"-"*33)
        print("\n\tOpciones:")
        print("\t  • Un número (ej: 5)")
        print("\t  • Varios números separados por comas (ej: 1,3,5)")
        print("\t  • Rango (ej: 1-5)")
        print("\t  • 'TODOS' para todos los hospitales")
        
        return input("\n\t¿Qué hospital(es) quieres procesar? ").strip()
    
    def procesar_seleccion_hospitales(self, seleccion, hospitales):
        """Procesa la selección del usuario para hospitales"""
        seleccionados = []
        
        if seleccion.upper() == 'TODOS':
            return hospitales
        
        # Rango (ej: 1-5)
        if '-' in seleccion:
            try:
                inicio, fin = map(int, seleccion.split('-'))
                for i in range(inicio, fin + 1):
                    if 1 <= i <= len(hospitales):
                        seleccionados.append(hospitales[i-1])
            except:
                pass
        
        # Múltiples números (ej: 1,3,5)
        elif ',' in seleccion:
            try:
                numeros = [int(n.strip()) for n in seleccion.split(',')]
                for num in numeros:
                    if 1 <= num <= len(hospitales):
                        seleccionados.append(hospitales[num-1])
            except:
                pass
        
        # Un solo número
        else:
            try:
                num = int(seleccion)
                if 1 <= num <= len(hospitales):
                    seleccionados.append(hospitales[num-1])
            except:
                pass
        
        return seleccionados
    
    def crear_estructura_carpetas(self, url_info, anos_seleccionados, hospitales_count):
        """Crea la estructura de carpetas para los resultados"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        # Carpeta principal
        url_nombre = url_info['nombre'].replace(' ', '_').replace('.', '')
        carpeta_principal = f"LEQ_{url_nombre}_{timestamp}"
        os.makedirs(carpeta_principal, exist_ok=True)
        
        print(f"\n\t  Carpeta principal: {carpeta_principal}")
        
        return carpeta_principal
    
    def extraer_datos_span(self, driver, nombre_hospital, texto_mes, nombre_especialidad=None):
        """Extrae datos del span con los indicadores"""
        datos = []
        
        try:
            # Buscar el span específico
            span_element = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "ContenedorContenidoSeccion_lblIndicadores"))
            )
            
            span_text = span_element.get_attribute('innerHTML')
            
            # Extraer los dos valores del span
            pacientes_match = re.search(r'Nº total de pacientes.*?: *([\d.,]+)', span_text)
            demora_match = re.search(r'Demora media.*?: *([\d.,]+)\s*días', span_text)
            
            pacientes = pacientes_match.group(1).replace(',', '.') if pacientes_match else None
            demora = demora_match.group(1).replace(',', '.') if demora_match else None
            
            # Extraer año del texto del mes
            ano, mes = self.extraer_ano_y_mes_del_texto(texto_mes)
                
            if pacientes or demora:
                registro = {
                    'Fecha_Extraccion': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'URL': driver.current_url,
                    'Filtro_Mes': texto_mes,
                    'Filtro_Hospital': nombre_hospital,
                    'Filtro_Especialidad': nombre_especialidad if nombre_especialidad else 'Todas',
                    'Año':  ano,
                    'Mes':  mes,
                    'Pacientes_en_Lista':  pacientes.replace('.', ''),
                    'Demora_Media':  demora,
                    'Texto_Completo': span_text
                }
                
                datos.append(registro)
                return datos, True
            else:
                return [], False
        
        except Exception as e:
            return [], False
    
    def extraer_datos(self, driver, nombre_hospital, texto_mes, nombre_especialidad=None):
        """Función principal para extraer datos según el tipo de contenido"""
        # Primero intentar extraer del span
        datos_span, exito = self.extraer_datos_span(driver, nombre_hospital, texto_mes, nombre_especialidad)
        
        return datos_span
        
    def extraer_ano_y_mes_del_texto(self, texto):
        """Extrae el año y mes del texto en formato 'mes año'"""
        try:
            # Separar por espacio
            partes = texto.strip().split()
            
            if len(partes) >= 2:
                # Mes es la primera parte, año es la segunda
                mes_str = partes[0]
                ano_str = partes[1]
                
                return ano_str, mes_str  # Devuelve (año, mes) como strings
            
            return None, None
            
        except:
            return None, None
    
    def guardar_archivos_consolidados(self, todos_datos, estadisticas, carpeta_principal, anos_seleccionados, filtrar):
        """Guarda archivos consolidados"""
        
        if not todos_datos:
            print("\n⚠ No se extrajeron datos para generar archivos consolidados")
            return
        
        print(f"\n\n\n{'='*60}")
        print("GUARDANDO ARCHIVOS CONSOLIDADOS...")
        print(f"{'='*60}")
        
        df_completo = pd.DataFrame(todos_datos)
        
        # Determinar nombre según modo de operación
        if filtrar and anos_seleccionados:
            nombre_base = f"Datos_Filtrados_{'_'.join(map(str, anos_seleccionados))}"
        else:
            nombre_base = "Datos_Completos"
        
        # 1. EXCEL CONSOLIDADO
        excel_consolidado = os.path.join(carpeta_principal, f"{nombre_base}.xlsx")
        
        # Crear Excel con múltiples hojas
        with pd.ExcelWriter(excel_consolidado, engine='openpyxl') as writer:
            # Hoja principal con todos los datos
            df_completo.to_excel(writer, sheet_name='Todos_Datos', index=False)
            
            # Hoja por hospital (si hay múltiples hospitales)
            if 'Hospital' in df_completo.columns:
                hospitales_unicos = df_completo['Hospital'].unique()
                for hospital in hospitales_unicos:
                    df_hospital = df_completo[df_completo['Hospital'] == hospital]
                    nombre_hoja = hospital[:30]  # Limitar longitud del nombre
                    df_hospital.to_excel(writer, sheet_name=nombre_hoja, index=False)
            
            # Hoja de estadísticas
            if estadisticas:
                df_estadisticas = pd.DataFrame(estadisticas)
                df_estadisticas.to_excel(writer, sheet_name='Estadisticas', index=False)
        
        print(f"\n\t✓ Excel consolidado: {os.path.basename(excel_consolidado)}")
        print(f"\t  - Hojas: {len(writer.sheets)}")
        print(f"\t  - Registros totales: {len(df_completo)}")
        
        # 2. CSV CONSOLIDADO
        csv_consolidado = os.path.join(carpeta_principal, f"{nombre_base}.csv")
        df_completo.to_csv(csv_consolidado, index=False, encoding='utf-8-sig')
        print(f"\t✓ CSV consolidado: {os.path.basename(csv_consolidado)}")
        
    
    def limpiar_nombre_archivo(self, nombre):
        """Limpia el nombre para usarlo como archivo"""
        caracteres_invalidos = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
        nombre_limpio = nombre
        for char in caracteres_invalidos:
            nombre_limpio = nombre_limpio.replace(char, '_')
        
        # Limitar longitud
        if len(nombre_limpio) > 100:
            nombre_limpio = nombre_limpio[:100]
        
        return nombre_limpio
    
    def ejecutar(self):
        """Función principal que ejecuta todo el proceso"""
        
        self.inicio_proceso = datetime.now()
        
        try:
            # 1. SELECCIÓN DE URL
            print("\n\n\n" + "="*60)
            print("PASO 1: SELECCIÓN DE URL")
            print("="*60)
            url_info = self.mostrar_menu_urls()
            self.url_actual = url_info['url']
            
            # 2. SELECCIÓN DE AÑO
            print("\n\n\n" + "="*60)
            print("PASO 2: SELECCIÓN DE AÑO")
            print("="*60)
            anos_seleccionados, filtrar = self.seleccionar_ano()
            
            # 3. INICIAR NAVEGADOR
            print("\n\n\n" + "="*60)
            print("PASO 3: INICIANDO NAVEGADOR")
            print("="*60)
            
            print("\n\tIniciando Chrome...")
            self.driver = webdriver.Chrome()
            self.driver.set_window_size(1400, 1000)
            print(f"\tCargando URL: {url_info['url']}")
            self.driver.get(url_info['url'])
            time.sleep(1)
            
            print(f"\tTítulo página: {self.driver.title}")
            
            # 4. OBTENER HOSPITALES
            print("\n\n\n" + "="*60)
            print("PASO 4: SELECCIÓN DE HOSPITALES")
            print("="*60)
            
            try:
                hospital_dropdown = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.ID, "ContenedorContenidoSeccion_ddlHospital"))
                )
                
                select_hospital = Select(hospital_dropdown)
                hospital_options = select_hospital.options
                
                hospitales = []
                for i, option in enumerate(hospital_options):
                    if option.get_attribute('value') and option.text.strip():
                        hospitales.append({
                            'indice': i,
                            'nombre': option.text.strip(),
                            'valor': option.get_attribute('value')
                        })
                
                print(f"\n\t ✓ {len(hospitales)} hospitales encontrados")
                
                # Selección interactiva de hospitales
                while True:
                    seleccion = self.mostrar_menu_hospitales(hospitales)
                    hospitales_seleccionados = self.procesar_seleccion_hospitales(seleccion, hospitales)
                    
                    if hospitales_seleccionados:
                        print(f"\n\tHospitales seleccionados: {len(hospitales_seleccionados)}")
                        for i, hosp in enumerate(hospitales_seleccionados, 1):
                            print(f"\t  {i:3}. {hosp['nombre']}")
                        break
                    else:
                        print("\n\tSelección no válida. Intenta de nuevo.\n")
                
            except Exception as e:
                print(f"\tError obteniendo hospitales: {e}")
                return
            
			# 5. SELECCIÓN DE ESPECIALIDADES (PARA TODOS LOS HOSPITALES)
            print("\n\n\n" + "="*60)
            print("PASO 5: SELECCIÓN DE ESPECIALIDADES")
            print("="*60)
            
            # Obtener especialidades del primer hospital como referencia
            try:
                # Seleccionar primer hospital para obtener las especialidades disponibles
                hospital_dropdown = self.driver.find_element(By.ID, "ContenedorContenidoSeccion_ddlHospital")
                select_hospital = Select(hospital_dropdown)
                select_hospital.select_by_index(1)  # Primer hospital
#                time.sleep(2)
                
                especialidades = self.obtener_especialidades(self.driver)
                
                if especialidades:
                    print(f"\n\t ✓ {len(especialidades)} especialidades encontradas")
                    
                    # Selección interactiva de especialidades
                    while True:
                        seleccion = self.mostrar_menu_especialidades(especialidades)
                        self.especialidades_seleccionadas_global = self.procesar_seleccion_especialidades(seleccion, especialidades)
                        
                        if self.especialidades_seleccionadas_global is not None:
                            if not self.especialidades_seleccionadas_global:
                                print(f"\n\t  ✓ Procesando SIN filtro de especialidad para TODOS los hospitales")
                                break
                            else:
                                print(f"\n\t  ✓ Especialidades seleccionadas: {len(self.especialidades_seleccionadas_global)}")
                                print(f"\t    Aplicadas a TODOS los hospitales seleccionados")
                                for i, esp in enumerate(self.especialidades_seleccionadas_global[:5], 1):
                                    print(f"\t    {i:2}. {esp['nombre'][:40]}")
                                if len(self.especialidades_seleccionadas_global) > 5:
                                    print(f"\t    ... y {len(self.especialidades_seleccionadas_global)-5} más")
                                break
                        else:
                            print("\n\tSelección no válida. Intenta de nuevo.\n")
                else:
                    print(f"\n\t  ✗ No se encontraron especialidades en este formulario")
                    print(f"\t    Se procesará SIN filtro de especialidad")
                    self.especialidades_seleccionadas_global = []
                    
            except Exception as e:
                print(f"\tError obteniendo especialidades: {e}")
                print(f"\tSe procesará SIN filtro de especialidad")
                self.especialidades_seleccionadas_global = []
            
            # Volver a cargar la página para limpiar selecciones
            print(f"\n\tReiniciando formulario...")
            self.driver.get(self.url_actual)
#            time.sleep(3)
			
			
            # 6. CREAR ESTRUCTURA DE CARPETAS
            print("\n\n\n" + "="*60)
            print("PASO 6: CREAR CARPETA PARA ARCHIVOS FINALES")
            print("="*60)
            carpeta_principal = self.crear_estructura_carpetas(
                url_info, anos_seleccionados, len(hospitales_seleccionados)
            )
            
            # 7. PROCESAR CADA HOSPITAL
            print("\n\n\n" + "="*60)
            print("PASO 7: PROCESANDO HOSPITALES")
            print("="*60)
            
            todos_datos = []
            estadisticas = []
            
            for idx, hospital in enumerate(hospitales_seleccionados):
                print(f"\n\n\t{'-'*60}")
                print(f"\tHOSPITAL {idx+1}/{len(hospitales_seleccionados)}: {hospital['nombre']}")
                print(f"\t{'-'*60}\n")
                
                # Seleccionar hospital
                try:
                    hospital_dropdown = self.driver.find_element(By.ID, "ContenedorContenidoSeccion_ddlHospital")
                    select_hospital = Select(hospital_dropdown)
                    select_hospital.select_by_index(hospital['indice'])
#                    time.sleep(2)
                except Exception as e:
                    print(f"\t✗ Error seleccionando hospital: {e}")
                    continue
                
                # Obtener especialidades para este hospital
                especialidades = self.obtener_especialidades(self.driver)
                
                if not especialidades:
#                    print(f"\t  ✓ {len(especialidades)} especialidades encontradas para este hospital")
                    print(f"\t  ✗ No hay lista de especialidades para este hospital")
                    # Si no hay especialidades, usar una lista vacía para indicar "sin filtro"
                    especialidades_a_procesar = []
                elif not self.especialidades_seleccionadas_global:
                    print(f"\t  ✓ Procesando SIN filtro de especialidad (selección global)")
                    especialidades_a_procesar = []
                else:
                    # Usar las especialidades seleccionadas globalmente
                    especialidades_a_procesar = []
                    
                    # Filtrar solo las especialidades seleccionadas que existan en este hospital
                    for esp_sel in self.especialidades_seleccionadas_global:
                        for esp_hosp in especialidades:
                            if esp_sel['valor'] == esp_hosp['valor']:
                                especialidades_a_procesar.append(esp_hosp)
                                break
                    
#                    print(f"\t  ✓ Procesando {len(especialidades_a_procesar)} especialidades seleccionadas")
                    if len(especialidades_a_procesar) < len(self.especialidades_seleccionadas_global):
                        print(f"\t    (Nota: {len(self.especialidades_seleccionadas_global) - len(especialidades_a_procesar)} especialidades no disponibles en este hospital)")
#                else:
#                    print(f"\t  ✓ {len(especialidades)} especialidades encontradas para este hospital")
#                    
#                    # 8. SELECCIÓN DE ESPECIALIDADES
#                    print(f"\n\t{'='*40}")
#                    print(f"\tSELECCIÓN DE ESPECIALIDADES")
#                    print(f"\t{'='*40}")
#                    
#                    while True:
#                        seleccion_esp = self.mostrar_menu_especialidades(especialidades)
#                        especialidades_a_procesar = self.procesar_seleccion_especialidades(seleccion_esp, especialidades)
#                        
#                        if especialidades_a_procesar is not None:
#                            if not especialidades_a_procesar:
#                                print(f"\n\t  ✓ Procesando SIN filtro de especialidad")
#                                break
#                            else:
#                                print(f"\n\t  ✓ Especialidades seleccionadas: {len(especialidades_a_procesar)}")
#                                for i, esp in enumerate(especialidades_a_procesar[:5], 1):
#                                    print(f"\t    {i:2}. {esp['nombre'][:40]}")
#                                if len(especialidades_a_procesar) > 5:
#                                    print(f"\t    ... y {len(especialidades_a_procesar)-5} más")
#                                break
#                        else:
#                            print("\n\tSelección no válida. Intenta de nuevo.\n")
                
                # Obtener meses disponibles
                try:
                    fecha_dropdown = self.driver.find_element(By.ID, "ContenedorContenidoSeccion_ddlFecha")
                    select_fecha = Select(fecha_dropdown)
                    meses_options = select_fecha.options
                    
                    # Obtener todas las meses
                    todas_meses = []
                    for option in meses_options:
                        if option.get_attribute('value') and option.text.strip():
                            todas_meses.append({
                                'texto': option.text.strip(),
                                'valor': option.get_attribute('value')
                            })
                    
                    # Filtrar meses según selección
                    meses_a_procesar = self.filtrar_meses(todas_meses, anos_seleccionados, filtrar)
                    
#                    print(f"\n\tEstadísticas:")
#                    print(f"\t  • Total meses disponibles: {len(todas_meses)}")
#                    print(f"\t  • Meses a procesar: {len(meses_a_procesar)}")
#                    
                    if not especialidades_a_procesar:
#                        print(f"\t  • Especialidades a procesar: Sin filtro (1 consulta por mes)")
                        total_consultas = len(meses_a_procesar)
                    else:
#                        print(f"\t  • Especialidades a procesar: {len(especialidades_a_procesar)}")
                        total_consultas = len(meses_a_procesar) * len(especialidades_a_procesar)
#                    
#                    print(f"\t  • Total consultas: {total_consultas}")
                    
                    if not meses_a_procesar:
#                        print(f"\t✗ No hay meses para procesar con los criterios seleccionados")
                        estadisticas.append({
                            'Hospital': hospital['nombre'],
                            'Meses_Disponibles': len(todas_meses),
                            #'Especialidades': len(especialidades) if especialidades else 0,
                            'Especialidades_Seleccionadas': len(self.especialidades_seleccionadas_global) if self.especialidades_seleccionadas_global else 0,
                            'Especialidades_Disponibles': len(especialidades) if especialidades else 0,
                            'Consultas_Planificadas': 0,
                            'Consultas_Exitosas': 0,
                            'Registros': 0,
                            'Estado': 'Sin meses para procesar'
                        })
                        continue
                    
                    # Mostrar meses a procesar (solo primeros 5 para no saturar)
#                    print(f"\n\tMESES A PROCESAR (primeros 5):")
#                    for i, mes in enumerate(meses_a_procesar[:5]):
#                        año_info = f" - Año: {mes.get('año', 'N/A')}" if 'año' in mes else ""
#                        print(f"\t  {i+1:2}. {mes['texto'][:30]}...{año_info}")
#                    
#                    if len(meses_a_procesar) > 5:
#                        print(f"\t  ... y {len(meses_a_procesar)-5} más")
                    
                    # Mostrar especialidades (si hay)
#                    if especialidades_a_procesar:
#                        print(f"\n\tESPECIALIDADES A PROCESAR (primeras 5):")
#                        for i, esp in enumerate(especialidades_a_procesar[:5]):
#                            print(f"\t  {i+1:2}. {esp['nombre'][:40]}...")
#                        
#                        if len(especialidades_a_procesar) > 5:
#                            print(f"\t  ... y {len(especialidades_a_procesar)-5} más")
#                    else:
#                        print(f"\n\tMODO: Sin filtro de especialidad")
                    
#                    print(' ')
                    
                except Exception as e:
                    print(f"\t✗ Error obteniendo meses: {e}")
                    continue
                
                # PROCESAR CADA COMBINACIÓN: MES × ESPECIALIDAD
                datos_hospital = []
                consultas_exitosas = 0
                
                # Si no hay especialidades seleccionadas, procesar sin filtro
                if not especialidades_a_procesar:
                    for mes_idx, mes in enumerate(meses_a_procesar):
                        consulta_num = mes_idx + 1
                        
                        ############ LIMITE DE ITERACIONES PARA OPTIMIZAR PRUEBAS ############  
                        # if consulta_num > 15:
                        #     break
                        ############ LIMITE DE ITERACIONES PARA OPTIMIZAR PRUEBAS ############  
                            
                        try:
                            # Seleccionar mes
                            fecha_dropdown = self.driver.find_element(By.ID, "ContenedorContenidoSeccion_ddlFecha")
                            select_fecha = Select(fecha_dropdown)
                            select_fecha.select_by_value(mes['valor'])
#                            time.sleep(1)
                            
                            # Hacer clic en Buscar
                            boton = self.driver.find_element(By.ID, "ContenedorContenidoSeccion_btnEnviar")
                            self.driver.execute_script("arguments[0].click();", boton)
#                            time.sleep(3)
                            
                            # Extraer datos
                            datos = self.extraer_datos(
                                self.driver, 
                                hospital['nombre'], 
                                mes['texto'],
                                None  # Sin especialidad
                            )
                            
                            if datos:
                                datos_hospital.extend(datos)
                                consultas_exitosas += 1
                                # Mostrar datos extraídos
                                if datos and 'Pacientes_en_Lista' in datos[0] and 'Demora_Media' in datos[0]:
                                    pacientes = datos[0]['Pacientes_en_Lista']
                                    demora = datos[0]['Demora_Media']
                                    print(f"\t  [{consulta_num:3}/{total_consultas}] Mes: {mes['texto'][:20]}... - ✓ {pacientes} pacientes, {demora} días")
                                else:
                                    print(f"\t       ✓ {len(datos)} registros")
                            else:
                                print(f"\t       ✗ Sin datos en esta consulta")
                            
                        except Exception as e:
                            print(f"\t       ✗ Error: {str(e)[:80]}")
                            continue
                else:
                    # Procesar con especialidades seleccionadas
                    for mes_idx, mes in enumerate(meses_a_procesar):
                        for esp_idx, especialidad in enumerate(especialidades_a_procesar):
                            consulta_num = mes_idx * len(especialidades_a_procesar) + esp_idx + 1
                            
                            ############ LIMITE DE ITERACIONES PARA OPTIMIZAR PRUEBAS ############  
                            # if consulta_num > 15:
                            #     break
                            ############ LIMITE DE ITERACIONES PARA OPTIMIZAR PRUEBAS ############  
                                
                            try:
                                # Seleccionar especialidad
                                especialidad_dropdown = self.driver.find_element(By.ID, "ContenedorContenidoSeccion_ddlEspecialidad")
                                select_especialidad = Select(especialidad_dropdown)
                                select_especialidad.select_by_value(especialidad['valor'])
#                                time.sleep(1)
                                
                                # Seleccionar mes
                                fecha_dropdown = self.driver.find_element(By.ID, "ContenedorContenidoSeccion_ddlFecha")
                                select_fecha = Select(fecha_dropdown)
                                select_fecha.select_by_value(mes['valor'])
#                                time.sleep(1)
                                
                                # Hacer clic en Buscar
                                boton = self.driver.find_element(By.ID, "ContenedorContenidoSeccion_btnEnviar")
                                self.driver.execute_script("arguments[0].click();", boton)
#                                time.sleep(3)
                                
                                # Extraer datos
                                datos = self.extraer_datos(
                                    self.driver, 
                                    hospital['nombre'], 
                                    mes['texto'],
                                    especialidad['nombre']
                                )
                                
                                if datos:
                                    datos_hospital.extend(datos)
                                    consultas_exitosas += 1
                                    # Mostrar datos extraídos
                                    if datos and 'Pacientes_en_Lista' in datos[0] and 'Demora_Media' in datos[0]:
                                        pacientes = datos[0]['Pacientes_en_Lista']
                                        demora = datos[0]['Demora_Media']
                                        print(f"\t  [{consulta_num:3}/{total_consultas}] Mes: {mes['texto'][:20]} ( {especialidad['nombre'][:20]} )... - ✓ {pacientes} pacientes, {demora} días")
                                    else:
                                        print(f"\t       ✓ {len(datos)} registros")
                                else:
                                    print(f"\t       ✗ Sin datos en esta consulta")
                                
                            except Exception as e:
                                print(f"\t       ✗ Error: {str(e)[:80]}")
                                continue
                
                # GUARDAR DATOS DEL HOSPITAL
                if datos_hospital:
                    todos_datos.extend(datos_hospital)
                     
                    # Estadísticas actualizadas
                    estadisticas.append({
                         'Hospital': hospital['nombre'],
                         'Meses_Disponibles': len(todas_meses),
                         #'Especialidades': len(especialidades) if especialidades else 0,
                         'Especialidades_Seleccionadas': len(self.especialidades_seleccionadas_global) if self.especialidades_seleccionadas_global else 0,
                         'Especialidades_Disponibles': len(especialidades) if especialidades else 0,
                         'Consultas_Planificadas': total_consultas,
                         'Consultas_Exitosas': consultas_exitosas,
                         'Registros': len(datos_hospital),
                         'Estado': 'Completado'
                    })
                    
                else:
                    print(f"\t✗ No se extrajeron datos para este hospital")
                    estadisticas.append({
                        'Hospital': hospital['nombre'],
                        'Meses_Disponibles': len(todas_meses),
                        'Especialidades': len(especialidades) if especialidades else 0,
                        'Consultas_Planificadas': total_consultas,
                        'Consultas_Exitosas': 0,
                        'Registros': 0,
                        'Estado': 'Sin datos extraídos'
                    })
            
            # 9. GUARDAR ARCHIVOS CONSOLIDADOS
            if todos_datos:
                self.guardar_archivos_consolidados(
                    todos_datos, 
                    estadisticas, 
                    carpeta_principal, 
                    anos_seleccionados,
                    filtrar
                )
            else:
                print(f"\n\n\n{'='*60}")
                print("NO SE EXTRAJERON DATOS")
                print(f"{'='*60}")
                print("Posibles causas:")
                print("  1. No hay datos disponibles para los criterios seleccionados")
                print("  2. La estructura de la página ha cambiado")
                print("  3. Problemas de conexión o tiempo de espera")
                print(f"\nArchivos de log guardados en: {carpeta_principal}")
            
        except Exception as e:
            print(f"\n{'='*60}")
            print("ERROR CRÍTICO DURANTE LA EJECUCIÓN")
            print(f"{'='*60}")
            print(f"Error: {e}")
            import traceback
            traceback.print_exc()
            
        finally:
            if self.driver:
                print(f"\n\n\n{'='*60}")
                print("FINALIZANDO EJECUCIÓN")
                print(f"{'='*60}")
                print("\n\tEl navegador se mantendrá abierto.")
                print("\tPuedes cerrarlo manualmente o presionar Enter para cerrarlo automáticamente.")
                
                input("\n\tPresiona Enter para cerrar el navegador y terminar...")
                self.driver.quit()
            
            duracion = datetime.now() - self.inicio_proceso
            print(f"\n\tTiempo total de ejecución: {duracion.total_seconds():.1f} segundos")
            print(f"\tFecha y hora de finalización: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            print(f"\n{'='*60}")
            print("¡PROCESO COMPLETADO!")
            print(f"{'='*60}\n")

def main():
    """Función principal"""
    print("\n\n\n" + "="*60)
    print("\t  SANIDADMADRID.ORG  LEQ  SCRAPER ")
    print("="*60)
    print("\n\tEste script permite extraer datos de listas de espera de hospitales.")
    print("\n\tPara interrumpir la ejecucion del programa presion Ctrl+C.")
    
    print("\n\tCONFIGURACIÓN REQUERIDA:")
    print("\t" +"-"*25)
    print("\t  1. Tener instalado Chrome")
    print("\t  2. Tener chromedriver en el PATH")
    print("\t  3. Conexión a internet estable")
    
    scraper = LEQScraper()
    scraper.ejecutar()

if __name__ == "__main__":
    main()
