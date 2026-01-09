import time
import pandas as pd
from datetime import datetime
import os
import re
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

class LEQScraper:
    def __init__(self):
        self.driver = None
        self.modo_verbose_EXEC = False  # Cambiar entre True/False para modo verbose/silencioso
        self.especialidades_seleccionadas_global = []
        self.modo_verbose = True 
        self.logger = None
        
        # Configurar tiempos de espera (limitados al máximo actualmente)
        self.TIEMPO_ESPERA_CORTO = 0   # 1 segundos
        self.TIEMPO_ESPERA_NORMAL = 0  # 2 segundos
        self.TIEMPO_ESPERA_LARGO = 3   # 3 segundos
        self.TIEMPO_TIMEOUT = 10       # segundos para WebDriverWait
        
        self.urls_disponibles = {
            1: {
                'nombre': 'Lista de Espera Quirúrgica por hospital y procesos-patologías',
                'nombre_file': 'Quirúrgica por procesos-patologías',
                'url': 'https://servicioselectronicos.sanidadmadrid.org/LEQ/ConsultaProcesos.aspx'
            },
            2: {
                'nombre': 'Lista de Espera Quirúrgica por hospital y especialidad',
                'nombre_file': 'Quirúrgica por especialidad',
                'url': 'https://servicioselectronicos.sanidadmadrid.org/LEQ/Consulta.aspx'
            },
            3: {
                'nombre': 'Lista de Espera de Consultas Externas por hospital y especialidad',
                'nombre_file': 'Consultas Externas',
                'url': 'https://servicioselectronicos.sanidadmadrid.org/LEQ/ConsultaEspecialidades.aspx'
            },
            4: {
                'nombre': 'Lista de Espera de Pruebas Diagnósticas y Terapéuticas',
                'nombre_file': 'Pruebas Diagnósticas y Terapéuticas',
                'url': 'https://servicioselectronicos.sanidadmadrid.org/LEQ/ConsultaPruebas.aspx'
            }
        }
    
    def log_info(self, mensaje):
        """Log info si está en modo verbose"""
        if self.modo_verbose:
            print(mensaje)
        
        if self.logger:
            self.logger.info(mensaje)
    
    def log_error(self, mensaje):
        """Log de errores"""
        if self.modo_verbose:
            print(f" {mensaje}")
        
        if self.logger:
            self.logger.error(mensaje)
    
    def log_warning(self, mensaje):
        """Log de advertencias"""
        if self.modo_verbose:
            print(f" {mensaje}")
        
        if self.logger:
            self.logger.warning(mensaje)
    
    def log_success(self, mensaje):
        """Log de éxitos"""
        if self.modo_verbose:
            print(f" {mensaje}")
        
        if self.logger:
            self.logger.info(f" {mensaje}")
    
    def configurar_logging(self, carpeta_principal):
        """Configura logging profesional"""
        log_file = os.path.join(carpeta_principal, 'ejecucion.log')
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler()  # También muestra en consola
            ]
        )
        
        self.logger = logging.getLogger(__name__)
#        self.log_success(f"Log guardado en: {os.path.basename(log_file)}")
    
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
            especialidad_dropdown = WebDriverWait(driver, self.TIEMPO_TIMEOUT).until(
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
            self.log_error(f"No se encontró lista de especialidades o hubo un error: {e}")
            return []
    
    def mostrar_menu_especialidades(self, especialidades):
        """Muestra menú para elegir especialidades"""
        print("\n\tESPECIALIDADES DISPONIBLES:")
        print("\t" + "-"*40)
        
        # Mostrar las primeras 20 especialidades para no saturar la pantalla
        for i, especialidad in enumerate(especialidades[:50], 1):
            print(f"\t{i:3}. {especialidad['nombre'][:40]}")
        
        if len(especialidades) > 50:
            print(f"\t... y {len(especialidades) - 50} más")
        
        print("\n\tOpciones:")
        print("\t  • Un número (ej: 5)")
        print("\t  • Varios números separados por comas (ej: 1,3,5) o puntos (ej: 1.3.5)")
        print("\t  • Rango (ej: 1-5)")
        print("\t  • Vacio (Enter) para todas las especialidades")
        
        return input("\n\t¿Qué especialidad(es) quieres procesar? ").strip()
    
    def validar_y_parsear_entrada(self, entrada, tipo='numeros'):
        """Valida y parsea entradas del usuario"""
        if tipo == 'numeros':
            # Para entradas como "1,3,5" o  "1.3.5" o "1-5"
            numeros = []
            
            if '-' in entrada:
                try:
                    inicio, fin = map(int, entrada.split('-'))
                    numeros = list(range(inicio, fin + 1))
                except:
                    return None
            elif ',' in entrada:
                try:
                    numeros = [int(n.strip()) for n in entrada.split(',')]
                except:
                    return None
            elif '.' in entrada:
                try:
                    numeros = [int(n.strip()) for n in entrada.split('.')]
                except:
                    return None
            else:
                try:
                    numeros = [int(entrada.strip())]
                except:
                    return None
            
            return numeros
        
        return None
    
    def procesar_seleccion_especialidades(self, seleccion, especialidades):
        """Procesa la selección del usuario para especialidades"""
        seleccionadas = []
        
        seleccion_upper = seleccion.upper()
        
        if seleccion_upper == '':
            return especialidades
        elif seleccion_upper == '0':
            # Devolver lista vacía para indicar "sin filtro"
            return especialidades
        
        # Usar función de validación
        numeros = self.validar_y_parsear_entrada(seleccion)
        
        if numeros:
            for num in numeros:
                if 1 <= num <= len(especialidades):
                    seleccionadas.append(especialidades[num-1])
        
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
        print("\n")
        
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
        print("\t  • Varios números separados por comas (ej: 1,3,5) o puntos (ej: 1.3.5)")
        print("\t  • Rango (ej: 1-5)")
        print("\t  • Enter para todos los hospitales")
        
        return input("\n\t¿Qué hospital(es) quieres procesar? ").strip()
    
    def procesar_seleccion_hospitales(self, seleccion, hospitales):
        """Procesa la selección del usuario para hospitales"""
        seleccionados = []
        
        if seleccion.upper() == '':
            return hospitales
        
        # Usar función de validación
        numeros = self.validar_y_parsear_entrada(seleccion)
        
        if numeros:
            for num in numeros:
                if 1 <= num <= len(hospitales):
                    seleccionados.append(hospitales[num-1])
        
        return seleccionados
    
    def seleccionar_elemento_dropdown(self, element_id, valor, usar_index=True):
        """Función genérica para seleccionar elementos dropdown"""
        try:
            elemento = self.driver.find_element(By.ID, element_id)
            select = Select(elemento)
            if usar_index:
                select.select_by_index(valor)
            else:
                select.select_by_value(valor)
            return True
        except Exception as e:
            self.log_error(f"Error seleccionando {element_id}: {e}")
            return False
    
    def hacer_clic_elemento(self, element_id, usar_javascript=True):
        """Función genérica para hacer clic en elementos"""
        try:
            elemento = self.driver.find_element(By.ID, element_id)
            if usar_javascript:
                self.driver.execute_script("arguments[0].click();", elemento)
            else:
                elemento.click()
            return True
        except Exception as e:
            self.log_error(f"Error haciendo clic en {element_id}: {e}")
            return False
    
    def obtener_elemento_con_reintentos(self, by, value, reintentos=3, tiempo_espera=1):
        """Obtiene un elemento con reintentos en caso de fallo"""
        for intento in range(reintentos):
            try:
                elemento = self.driver.find_element(by, value)
                return elemento
            except Exception as e:
                if intento < reintentos - 1:
                    time.sleep(tiempo_espera)
                    continue
                else:
                    raise e
    
    def ejecutar_accion_con_reintentos(self, funcion, *args, reintentos=3, **kwargs):
        """Ejecuta una acción con reintentos en caso de fallo"""
        for intento in range(reintentos):
            try:
                return funcion(*args, **kwargs)
            except Exception as e:
                if intento < reintentos - 1:
                    self.log_info(f"\tReintento {intento + 1}/{reintentos}...")
                    time.sleep(self.TIEMPO_ESPERA_CORTO)
                    continue
                else:
                    raise e
    
    def crear_estructura_carpetas(self, url_info, anos_seleccionados, hospitales_count):
        """Crea la estructura de carpetas para los resultados"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        # Carpeta principal
        url_nombre = url_info['nombre_file'].replace(' ', '_').replace('.', '')
        carpeta_principal = f"LEQ_{url_nombre}_{timestamp}"
        os.makedirs(carpeta_principal, exist_ok=True)
        
        self.log_success(f"Carpeta principal: {carpeta_principal}")
        
        return carpeta_principal
    
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
    
    def extraer_datos_span(self, driver, nombre_hospital, texto_mes, nombre_especialidad=None):
        """Extrae datos del span con los indicadores - Versión mejorada"""
        datos = []
        
        for intento in range(2):  # 2 intentos
            try:
                # Esperar con condiciones más específicas
                span_element = WebDriverWait(driver, self.TIEMPO_TIMEOUT).until(
                    EC.presence_of_element_located((By.ID, "ContenedorContenidoSeccion_lblIndicadores"))
                )
                
                # Esperar a que el texto esté disponible
                WebDriverWait(driver, self.TIEMPO_TIMEOUT).until(
                    lambda d: span_element.text.strip() != ""
                )
                
                span_text = span_element.get_attribute('innerHTML')
                
                # Patrones mejorados para extracción
                patrones_pacientes = [
                    r'Nº total de pacientes.*?: *([\d.,]+)',
                    r'Total pacientes.*?: *([\d.,]+)',
                    r'Pacientes.*?: *([\d.,]+)'
                ]
                
                patrones_demora = [
                    r'Demora media.*?: *([\d.,]+)\s*días',
                    r'Demora.*?: *([\d.,]+)\s*días',
                    r'Media.*?: *([\d.,]+)\s*días'
                ]
                
                pacientes = None
                demora = None
                
                for patron in patrones_pacientes:
                    match = re.search(patron, span_text, re.IGNORECASE)
                    if match:
                        pacientes = match.group(1).replace(',', '.')
                        break
                
                for patron in patrones_demora:
                    match = re.search(patron, span_text, re.IGNORECASE)
                    if match:
                        demora = match.group(1).replace(',', '.')
                        break
                
                # Extraer año y mes con más robustez
                ano, mes = self.extraer_ano_y_mes_del_texto(texto_mes)
                
                # Validar que tenemos datos útiles
                if pacientes is not None or demora is not None:
                    registro = {
                        'Fecha_Extraccion': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        'URL': driver.current_url,
                        'Filtro_Mes': texto_mes,
                        'Filtro_Hospital': nombre_hospital,
                        'Filtro_Especialidad': nombre_especialidad if nombre_especialidad else 'Todas',
                        'Año': ano,
                        'Mes': mes,
                        'Pacientes_en_Lista': pacientes.replace('.', '') if pacientes else '0',
                        'Demora_Media': demora if demora else '0',
                        'Texto_Completo': span_text[:500]  # Limitar longitud
                    }
                    
                    datos.append(registro)
                    return datos, True
                else:
                    # Si no encuentra datos, esperar y reintentar
                    if intento == 0:
                        time.sleep(self.TIEMPO_ESPERA_CORTO)
                        continue
                    
            except Exception as e:
                if intento == 0:
                    time.sleep(self.TIEMPO_ESPERA_CORTO)
                    continue
                else:
                    self.log_error(f"Error extrayendo datos: {str(e)[:100]}")
        
        return [], False
    
    def extraer_datos(self, driver, nombre_hospital, texto_mes, nombre_especialidad=None):
        """Función principal para extraer datos según el tipo de contenido"""
        # Primero intentar extraer del span
        datos_span, exito = self.extraer_datos_span(driver, nombre_hospital, texto_mes, nombre_especialidad)
        
        return datos_span
    
    def mostrar_barra_progreso(self, iteracion, total, longitud=50):
        """Muestra una barra de progreso en consola"""
        porcentaje = (iteracion + 1) / total
        completado = int(longitud * porcentaje)
        restante = longitud - completado
        
        barra = f"[{'#' * completado}{'-' * restante}] {porcentaje*100:6.2f}%"
        print(f"\r{barra}", end='', flush=True)
        
        if iteracion + 1 == total:
            print()  # Nueva línea al completar
    
    def mostrar_progreso_consulta(self, consulta_num, total_consultas, mes, datos, especialidad=None):
        """Muestra el progreso de forma más informativa"""
        if datos and datos[0].get('Pacientes_en_Lista') and datos[0].get('Demora_Media'):
            pacientes = datos[0]['Pacientes_en_Lista']
            demora = datos[0]['Demora_Media']
            
            if especialidad:
                self.log_info(f"\t[{consulta_num:3}/{total_consultas}] ✓ {mes['texto'][:15]:15} | {especialidad['nombre'][:20]:20} | Pac: {pacientes:>6} | Días: {demora:>6}")
            else:
                self.log_info(f"\t[{consulta_num:3}/{total_consultas}] ✓ {mes['texto'][:15]:15} | {'Sin especialidad':20} | Pac: {pacientes:>6} | Días: {demora:>6}")
        else:
            self.log_warning(f"\t[{consulta_num:3}/{total_consultas}] ✗ Sin datos")
    
    def manejar_error_consulta(self, error):
        """Maneja errores en las consultas"""
        self.log_error(f"Error en consulta: {str(error)[:80]}")
    
    def procesar_hospital_optimizado(self, hospital, meses_a_procesar, especialidades_a_procesar, total_consultas):
        """Procesa un hospital de forma optimizada"""
        datos_hospital = []
        consultas_exitosas = 0
        
        # Procesar sin filtro de especialidad
        if not especialidades_a_procesar:
            for mes_idx, mes in enumerate(meses_a_procesar):
                consulta_num = mes_idx + 1
                
                # Límite para pruebas (descomentar si es necesario)
                # if consulta_num > 15: break
                
                try:
                    # Seleccionar mes
                    if not self.seleccionar_elemento_dropdown(
                        "ContenedorContenidoSeccion_ddlFecha", 
                        mes['valor'], 
                        usar_index=False
                    ):
                        continue
                    
                    # Hacer clic en Buscar
                    if not self.hacer_clic_elemento("ContenedorContenidoSeccion_btnEnviar"):
                        continue
                    
                    time.sleep(self.TIEMPO_ESPERA_NORMAL)
                    
                    # Extraer datos
                    datos = self.extraer_datos(
                        self.driver, 
                        hospital['nombre'], 
                        mes['texto'],
                        None
                    )
                    
                    self.mostrar_progreso_consulta(consulta_num, total_consultas, mes, datos)
                    
                    if datos:
                        datos_hospital.extend(datos)
                        consultas_exitosas += 1
                        
                except Exception as e:
                    self.manejar_error_consulta(e)
                    continue
        else:
            # Procesar con especialidades
            for mes_idx, mes in enumerate(meses_a_procesar):
                for esp_idx, especialidad in enumerate(especialidades_a_procesar):
                    consulta_num = mes_idx * len(especialidades_a_procesar) + esp_idx + 1
                    
                    # Límite para pruebas
                    # if consulta_num > 15: break
                    
                    try:
                        # Seleccionar especialidad
                        if not self.seleccionar_elemento_dropdown(
                            "ContenedorContenidoSeccion_ddlEspecialidad",
                            especialidad['valor'],
                            usar_index=False
                        ):
                            continue
                        
                        # Seleccionar mes
                        if not self.seleccionar_elemento_dropdown(
                            "ContenedorContenidoSeccion_ddlFecha",
                            mes['valor'],
                            usar_index=False
                        ):
                            continue
                        
                        # Hacer clic en Buscar
                        if not self.hacer_clic_elemento("ContenedorContenidoSeccion_btnEnviar"):
                            continue
                        
                        time.sleep(self.TIEMPO_ESPERA_NORMAL)
                        
                        # Extraer datos
                        datos = self.extraer_datos(
                            self.driver,
                            hospital['nombre'],
                            mes['texto'],
                            especialidad['nombre']
                        )
                        
                        self.mostrar_progreso_consulta(consulta_num, total_consultas, mes, datos, especialidad)
                        
                        if datos:
                            datos_hospital.extend(datos)
                            consultas_exitosas += 1
                            
                    except Exception as e:
                        self.manejar_error_consulta(e)
                        continue
        
        return datos_hospital, consultas_exitosas
    
    def limpiar_nombre_hoja(self, nombre):
        """Limpia el nombre para usarlo como hoja de Excel"""
        caracteres_invalidos = ['<', '>', ':', '"', '/', '\\', '|', '?', '*', '[', ']']
        nombre_limpio = nombre
        
        for char in caracteres_invalidos:
            nombre_limpio = nombre_limpio.replace(char, '_')
        
        # Limitar longitud para Excel (31 caracteres máximo)
        if len(nombre_limpio) > 31:
            nombre_limpio = nombre_limpio[:31]
        
        return nombre_limpio
    
    def crear_hoja_resumen(self, writer, df_completo, estadisticas):
        """Crea una hoja de resumen en el Excel"""
        try:
            # Resumen general
            resumen_data = {
                'Total Hospitales Procesados': [len(df_completo['Filtro_Hospital'].unique()) if 'Filtro_Hospital' in df_completo.columns else 0],
                'Total Registros': [len(df_completo)],
                'Fecha Inicio': [self.inicio_proceso.strftime('%Y-%m-%d %H:%M:%S')],
                'Fecha Fin': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
                'Tiempo Total (segundos)': [(datetime.now() - self.inicio_proceso).total_seconds()],
                'Especialidades Seleccionadas': [len(self.especialidades_seleccionadas_global) if self.especialidades_seleccionadas_global else 0]
            }
            
            df_resumen = pd.DataFrame(resumen_data)
            df_resumen.to_excel(writer, sheet_name='Resumen', index=False)
            
        except Exception as e:
            self.log_error(f"Error creando hoja de resumen: {e}")
    
    def guardar_excel_completo(self, df_completo, estadisticas, excel_path):
        """Guarda Excel con múltiples hojas organizadas"""
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            # Hoja principal
            df_completo.to_excel(writer, sheet_name='Datos_Completos', index=False)
            
            # Hoja por hospital
            if 'Filtro_Hospital' in df_completo.columns:
                hospitales_unicos = df_completo['Filtro_Hospital'].unique()
                for hospital in hospitales_unicos:
                    df_hospital = df_completo[df_completo['Filtro_Hospital'] == hospital]
                    nombre_hoja = self.limpiar_nombre_hoja(hospital)
                    df_hospital.to_excel(writer, sheet_name=nombre_hoja, index=False)
            
            # Hoja por especialidad (si existe)
            if 'Filtro_Especialidad' in df_completo.columns:
                especialidades_unicas = df_completo['Filtro_Especialidad'].unique()
                for especialidad in especialidades_unicas:
                    df_especialidad = df_completo[df_completo['Filtro_Especialidad'] == especialidad]
                    if len(df_especialidad) > 0:
                        nombre_hoja = f"Esp_{self.limpiar_nombre_hoja(especialidad)}"
                        df_especialidad.to_excel(writer, sheet_name=nombre_hoja, index=False)
            
            # Hoja de estadísticas
            if estadisticas:
                df_estadisticas = pd.DataFrame(estadisticas)
                df_estadisticas.to_excel(writer, sheet_name='Estadisticas', index=False)
            
            # Hoja de resumen
            self.crear_hoja_resumen(writer, df_completo, estadisticas)
        
#        self.log_success(f"Excel: {os.path.basename(excel_path)}")
#        self.log_info(f"  - Hojas creadas: {len(writer.sheets)}")
#        self.log_info(f"  - Registros totales: {len(df_completo):,}")
    
    def guardar_resumen_ejecucion(self, carpeta_principal, df_completo, estadisticas):
        """Guarda un archivo de resumen de la ejecución"""
        resumen_path = os.path.join(carpeta_principal, 'resumen_ejecucion.txt')
        
        with open(resumen_path, 'w', encoding='utf-8') as f:
            f.write("="*60 + "\n")
            f.write("RESUMEN DE EJECUCIÓN\n")
            f.write("="*60 + "\n\n")
            
            f.write(f"Fecha de ejecución: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Duración total: {(datetime.now() - self.inicio_proceso).total_seconds():.1f} segundos\n")
            f.write(f"Total registros extraídos: {len(df_completo):,}\n")
            
            if 'Filtro_Hospital' in df_completo.columns:
                f.write(f"Hospitales procesados: {len(df_completo['Filtro_Hospital'].unique())}\n")
            
            if self.especialidades_seleccionadas_global:
                f.write(f"Especialidades seleccionadas: {len(self.especialidades_seleccionadas_global)}\n")
            else:
                f.write("Modo: Sin filtro de especialidad\n")
            
            f.write("\n" + "="*60 + "\n")
            f.write("ESTADÍSTICAS POR HOSPITAL\n")
            f.write("="*60 + "\n\n")
            
            if estadisticas:
                for estad in estadisticas:
                    f.write(f"Hospital: {estad.get('Hospital', 'N/A')}\n")
                    f.write(f"  - Meses procesados: {estad.get('Consultas_Exitosas', 0)}/{estad.get('Consultas_Planificadas', 0)}\n")
                    f.write(f"  - Registros: {estad.get('Registros', 0)}\n")
                    f.write(f"  - Estado: {estad.get('Estado', 'N/A')}\n\n")
        
        self.log_success(f"Resumen guardado en: {os.path.basename(resumen_path)}")
    
    def guardar_archivos_consolidados(self, todos_datos, estadisticas, carpeta_principal, anos_seleccionados, filtrar):
        """Guarda archivos en múltiples formatos"""
        
        if not todos_datos:
            self.log_warning("No se extrajeron datos para generar archivos consolidados")
            return
        
#        print("\n\n\n")
#        self.log_info(f"{'='*60}")
#        self.log_info("GUARDANDO ARCHIVOS...")
#        self.log_info(f"{'='*60}")
        
        df_completo = pd.DataFrame(todos_datos)
        
        # Nombre base según configuración
        if filtrar and anos_seleccionados:
            nombre_base = f"Datos_Filtrados_{'_'.join(map(str, anos_seleccionados))}"
        else:
            nombre_base = "Datos_Completos"
        
        # 1. EXCEL con múltiples hojas
        excel_path = os.path.join(carpeta_principal, f"{nombre_base}.xlsx")
        self.guardar_excel_completo(df_completo, estadisticas, excel_path)
        
        # 2. CSV principal
        csv_path = os.path.join(carpeta_principal, f"{nombre_base}.csv")
        df_completo.to_csv(csv_path, index=False, encoding='utf-8-sig', sep=';')
#        self.log_success(f"CSV: {os.path.basename(csv_path)} (separador: ;)")
        
        # 3. JSON para fácil consumo
#        json_path = os.path.join(carpeta_principal, f"{nombre_base}.json")
#        df_completo.to_json(json_path, orient='records', force_ascii=False, indent=2)
#        self.log_success(f"JSON: {os.path.basename(json_path)}")
        
        # 4. Archivo de resumen
#        self.guardar_resumen_ejecucion(carpeta_principal, df_completo, estadisticas)
    
    def ejecutar(self):
        """Función principal que ejecuta todo el proceso"""
        
        self.inicio_proceso = datetime.now()
        
        try:
            # 1. SELECCIÓN DE URL
            print("\n\n\n")
            self.log_info(f"{'='*60}")
            self.log_info("PASO 1: SELECCIÓN DE URL")
            self.log_info(f"{'='*60}")
            url_info = self.mostrar_menu_urls()
            self.url_actual = url_info['url']
            
            # 2. SELECCIÓN DE AÑO
            print("\n\n\n")
            self.log_info(f"{'='*60}")
            self.log_info("PASO 2: SELECCIÓN DE AÑO")
            self.log_info(f"{'='*60}")
            anos_seleccionados, filtrar = self.seleccionar_ano()
            
            # 3. INICIAR NAVEGADOR
            print("\n\n\n")
            self.log_info(f"{'='*60}")
            self.log_info("PASO 3: INICIANDO NAVEGADOR")
            self.log_info(f"{'='*60}")
            
            self.log_info("\n\tIniciando Chrome...")
            self.driver = webdriver.Chrome()
            self.driver.set_window_size(1400, 1000)
            self.log_info(f"\tCargando URL: {url_info['url']}")
            self.driver.get(url_info['url'])
            time.sleep(self.TIEMPO_ESPERA_NORMAL)
            
            self.log_info(f"\tTítulo página: {self.driver.title}")
            
            # 4. OBTENER HOSPITALES
            print("\n\n\n")
            self.log_info(f"{'='*60}")
            self.log_info("PASO 4: SELECCIÓN DE HOSPITALES")
            self.log_info(f"{'='*60}")
            
            try:
                hospital_dropdown = WebDriverWait(self.driver, self.TIEMPO_TIMEOUT).until(
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
                
                self.log_success(f"{len(hospitales)} hospitales encontrados")
                
                # Selección interactiva de hospitales
                while True:
                    seleccion = self.mostrar_menu_hospitales(hospitales)
                    hospitales_seleccionados = self.procesar_seleccion_hospitales(seleccion, hospitales)
                    
                    if hospitales_seleccionados:
                        self.log_success(f"Hospitales seleccionados: {len(hospitales_seleccionados)}")
                        for i, hosp in enumerate(hospitales_seleccionados, 1):
                            self.log_info(f"\t  {i:3}. {hosp['nombre']}")
                        break
                    else:
                        self.log_warning("\n\tSelección no válida. Intenta de nuevo.\n")
                
            except Exception as e:
                self.log_error(f"Error obteniendo hospitales: {e}")
                return
            
            # 5. SELECCIÓN DE ESPECIALIDADES (PARA TODOS LOS HOSPITALES)
            print("\n\n\n")
            self.log_info(f"{'='*60}")
            self.log_info("PASO 5: SELECCIÓN DE ESPECIALIDADES")
            self.log_info(f"{'='*60}")
            
            # Obtener especialidades del primer hospital como referencia
            try:
                # Seleccionar primer hospital para obtener las especialidades disponibles
                hospital_dropdown = self.driver.find_element(By.ID, "ContenedorContenidoSeccion_ddlHospital")
                select_hospital = Select(hospital_dropdown)
                select_hospital.select_by_index(1)  # Primer hospital
                
                especialidades = self.obtener_especialidades(self.driver)
                
                if especialidades:
                    self.log_success(f"{len(especialidades)} especialidades encontradas")
                    
                    # Selección interactiva de especialidades
                    while True:
                        seleccion = self.mostrar_menu_especialidades(especialidades)
                        self.especialidades_seleccionadas_global = self.procesar_seleccion_especialidades(seleccion, especialidades)
                        
                        if self.especialidades_seleccionadas_global is not None:
                            if not self.especialidades_seleccionadas_global:
                                self.log_success("Procesando SIN filtro de especialidad para TODOS los hospitales")
                                break
                            else:
                                self.log_success(f"Especialidades seleccionadas: {len(self.especialidades_seleccionadas_global)}")
                                self.log_info(f"\tAplicadas a TODOS los hospitales seleccionados")
                                for i, esp in enumerate(self.especialidades_seleccionadas_global[:5], 1):
                                    self.log_info(f"\t    {i:2}. {esp['nombre'][:40]}")
                                if len(self.especialidades_seleccionadas_global) > 5:
                                    self.log_info(f"\t    ... y {len(self.especialidades_seleccionadas_global)-5} más")
                                break
                        else:
                            self.log_warning("\n\tSelección no válida. Intenta de nuevo.\n")
                else:
                    self.log_warning("No se encontraron especialidades en este formulario")
                    self.log_info("\tSe procesará SIN filtro de especialidad")
                    self.especialidades_seleccionadas_global = []
                    
            except Exception as e:
                self.log_error(f"Error obteniendo especialidades: {e}")
                self.log_info("\tSe procesará SIN filtro de especialidad")
                self.especialidades_seleccionadas_global = []
            
            # Volver a cargar la página para limpiar selecciones
            self.log_info("\n\tReiniciando formulario...")
            self.driver.get(self.url_actual)
            time.sleep(self.TIEMPO_ESPERA_NORMAL)
            
            self.modo_verbose = self.modo_verbose_EXEC
			
            # 6. CREAR ESTRUCTURA DE CARPETAS Y CONFIGURAR LOGGING
            self.log_info(f"\n\n\n{'='*60}")
            self.log_info("PASO 6: PREPARANDO CARPETA Y LOGGING")
            self.log_info(f"{'='*60}")
            
            carpeta_principal = self.crear_estructura_carpetas(
                url_info, anos_seleccionados, len(hospitales_seleccionados)
            )
            
            # Configurar logging
            self.configurar_logging(carpeta_principal)
            
            # 7. PROCESAR CADA HOSPITAL
            print("\n\n\n")
            self.log_info(f"{'='*60}")
            self.log_info("PASO 7: PROCESANDO HOSPITALES")
            self.log_info(f"{'='*60}")
            
            todos_datos = []
            estadisticas = []
            
            for idx, hospital in enumerate(hospitales_seleccionados):
                print("\n\n")
                self.log_info(f"\t{'-'*60}")
                self.log_info(f"\tHOSPITAL {idx+1}/{len(hospitales_seleccionados)}: {hospital['nombre']}")
                self.log_info(f"\t{'-'*60}")
                
                # Seleccionar hospital
                try:
                    if not self.seleccionar_elemento_dropdown(
                        "ContenedorContenidoSeccion_ddlHospital",
                        hospital['indice'],
                        usar_index=True
                    ):
                        continue
                    
                except Exception as e:
                    self.log_error(f"Error seleccionando hospital: {e}")
                    continue
                
                # Obtener especialidades para este hospital
                especialidades = self.obtener_especialidades(self.driver)
                
                if not especialidades:
                    self.log_warning("No hay lista de especialidades para este hospital")
                    especialidades_a_procesar = []
                elif not self.especialidades_seleccionadas_global:
                    self.log_success("Procesando SIN filtro de especialidad (selección global)")
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
                    
                    if len(especialidades_a_procesar) < len(self.especialidades_seleccionadas_global):
                        self.log_warning(f"Nota: {len(self.especialidades_seleccionadas_global) - len(especialidades_a_procesar)} especialidades no disponibles en este hospital")
                
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
                    
                    if not especialidades_a_procesar:
                        total_consultas = len(meses_a_procesar)
                    else:
                        total_consultas = len(meses_a_procesar) * len(especialidades_a_procesar)
                    
                    if not meses_a_procesar:
                        self.log_warning("No hay meses para procesar con los criterios seleccionados")
                        estadisticas.append({
                            'Hospital': hospital['nombre'],
                            'Meses_Disponibles': len(todas_meses),
                            'Especialidades_Seleccionadas': len(self.especialidades_seleccionadas_global) if self.especialidades_seleccionadas_global else 0,
                            'Especialidades_Disponibles': len(especialidades) if especialidades else 0,
                            'Consultas_Planificadas': 0,
                            'Consultas_Exitosas': 0,
                            'Registros': 0,
                            'Estado': 'Sin meses para procesar'
                        })
                        continue
                    
                except Exception as e:
                    self.log_error(f"Error obteniendo meses: {e}")
                    continue
                
                # Procesar hospital con función optimizada
                datos_hospital, consultas_exitosas = self.procesar_hospital_optimizado(
                    hospital, meses_a_procesar, especialidades_a_procesar, total_consultas
                )
                
                # GUARDAR DATOS DEL HOSPITAL
                if datos_hospital:
                    todos_datos.extend(datos_hospital)
                    
                    # Estadísticas actualizadas
                    estadisticas.append({
                        'Hospital': hospital['nombre'],
                        'Meses_Disponibles': len(todas_meses),
                        'Especialidades_Seleccionadas': len(self.especialidades_seleccionadas_global) if self.especialidades_seleccionadas_global else 0,
                        'Especialidades_Disponibles': len(especialidades) if especialidades else 0,
                        'Consultas_Planificadas': total_consultas,
                        'Consultas_Exitosas': consultas_exitosas,
                        'Registros': len(datos_hospital),
                        'Estado': 'Completado'
                    })
                    
                    self.log_success(f"✓ {len(datos_hospital)} registros extraídos de este hospital")
                    
                else:
                    self.log_warning("No se extrajeron datos para este hospital")
                    estadisticas.append({
                        'Hospital': hospital['nombre'],
                        'Meses_Disponibles': len(todas_meses),
                        'Especialidades_Seleccionadas': len(self.especialidades_seleccionadas_global) if self.especialidades_seleccionadas_global else 0,
                        'Especialidades_Disponibles': len(especialidades) if especialidades else 0,
                        'Consultas_Planificadas': total_consultas,
                        'Consultas_Exitosas': 0,
                        'Registros': 0,
                        'Estado': 'Sin datos extraídos'
                    })
            
            # 8. GUARDAR ARCHIVOS CONSOLIDADOS
            if todos_datos:
                self.guardar_archivos_consolidados(
                    todos_datos, 
                    estadisticas, 
                    carpeta_principal, 
                    anos_seleccionados,
                    filtrar
                )
            else:
                self.log_info(f"\n\n\n{'='*60}")
                self.log_info("NO SE EXTRAJERON DATOS")
                self.log_info(f"{'='*60}")
                self.log_info("Posibles causas:")
                self.log_info("  1. No hay datos disponibles para los criterios seleccionados")
                self.log_info("  2. La estructura de la página ha cambiado")
                self.log_info("  3. Problemas de conexión o tiempo de espera")
                self.log_info(f"\nArchivos de log guardados en: {carpeta_principal}")
            
        except Exception as e:
            self.log_info(f"\n{'='*60}")
            self.log_info("ERROR CRÍTICO DURANTE LA EJECUCIÓN")
            self.log_info(f"{'='*60}")
            self.log_error(f"Error: {e}")
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
