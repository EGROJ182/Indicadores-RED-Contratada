import os
import shutil
import logging
import re
import sys
from pathlib import Path
from datetime import datetime
import pandas as pd

class GestorArchivosExcel:
    def __init__(self, ruta_origen, ruta_base_proveedores):
        self.ruta_origen = Path(ruta_origen)
        self.ruta_base_proveedores = Path(ruta_base_proveedores)
        self.archivos_procesados = []
        self.archivos_error = []
        self.carpetas_creadas = []
        self.logger = None
        
        # Configurar logging
        self.setup_logging()
        
    def setup_logging(self):
        """Configura el sistema de logging"""
        log_dir = Path("logs")
        log_dir.mkdir(exist_ok=True)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = log_dir / f"gestion_archivos_{timestamp}.log"
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file),
                logging.StreamHandler(sys.stdout)
            ],
            force=True
        )
        
        self.logger = logging.getLogger(__name__)
        self.logger.info("="*70)
        self.logger.info(f"Iniciando gestion de archivos desde: {self.ruta_origen}")
        self.logger.info("="*70)
        
    def validar_ruta_origen(self):
        """Valida que la ruta origen existe"""
        if not self.ruta_origen.exists():
            self.logger.error(f"La ruta origen no existe: {self.ruta_origen}")
            raise ValueError(f"Ruta no valida: {self.ruta_origen}")
        
        if not self.ruta_origen.is_dir():
            self.logger.error(f"La ruta no es un directorio: {self.ruta_origen}")
            raise ValueError(f"No es un directorio: {self.ruta_origen}")
        
        self.logger.info("Ruta origen validada correctamente")
        
    def validar_archivo_proveedores(self):
        """Valida que el archivo de proveedores existe"""
        if not self.ruta_base_proveedores.exists():
            self.logger.error(f"Archivo de proveedores no encontrado: {self.ruta_base_proveedores}")
            raise FileNotFoundError(f"No existe: {self.ruta_base_proveedores}")
        
        self.logger.info("Archivo de proveedores validado")
        
    def cargar_base_proveedores(self):
        """Carga la base de datos de proveedores"""
        try:
            df = pd.read_excel(self.ruta_base_proveedores)
            
            columnas = df.columns.tolist()
            self.logger.debug(f"Columnas encontradas: {columnas}")
            
            columna_contrato = None
            columna_sucursal = None
            
            for col in columnas:
                col_lower = str(col).lower().strip()
                if 'numero' in col_lower and 'contrato' in col_lower:
                    columna_contrato = col
                # if 'sucursal' in col_lower:
                if 'departamento' in col_lower:
                    columna_sucursal = col
            
            if not columna_contrato or not columna_sucursal:
                raise ValueError(f"Columnas 'numero_contrato' o 'sucursal' no encontradas. Columnas: {columnas}")
            
            df['numero_contrato'] = df[columna_contrato].astype(str).str.strip().str.zfill(4)
            # df['sucursal'] = df[columna_sucursal].astype(str).str.strip()
            df['departamento'] = df[columna_sucursal].astype(str).str.strip()
            
            self.logger.info(f"Base de proveedores cargada: {len(df)} registros")
            return df
        except Exception as e:
            self.logger.error(f"Error al cargar base de proveedores: {str(e)}")
            raise
    
    def extraer_datos_archivo(self, nombre_archivo):
        """
        Extrae numero de contrato, año y NIT del nombre del archivo
        Formato esperado: Anexo X 662-2024 NOMBRE NIT
        """
        try:
            nombre_sin_ext = nombre_archivo
            for ext in ['.xlsx', '.xls', '.docx', '.doc']:
                nombre_sin_ext = nombre_sin_ext.replace(ext, '')
            
            patron = r'(\d+)-(\d{4})\s'
            match = re.search(patron, nombre_sin_ext)
            
            if not match:
                raise ValueError(f"No se encontro patron numero_contrato-ano")
            
            numero_contrato = match.group(1)
            ano_contrato = match.group(2)
            
            numeros = re.findall(r'\d+', nombre_sin_ext)
            if len(numeros) < 2:
                raise ValueError(f"No se pudo extraer NIT del nombre")
            
            nit = numeros[-1]
            
            if len(numero_contrato) > 4:
                raise ValueError(f"Numero de contrato invalido: {numero_contrato} (mas de 4 digitos)")
            
            numero_contrato_formateado = numero_contrato.zfill(4)
            
            self.logger.debug(f"Datos extraidos - Contrato: {numero_contrato_formateado}, Ano: {ano_contrato}, NIT: {nit}")
            
            return {
                'numero_contrato': numero_contrato_formateado,
                'ano_contrato': ano_contrato,
                'nit': nit,
                'numero_original': numero_contrato
            }
        except Exception as e:
            self.logger.error(f"Error al extraer datos de '{nombre_archivo}': {str(e)}")
            raise
    
    def obtener_sucursal(self, numero_contrato, df_proveedores):
        """Obtiene la sucursal de la base de proveedores"""
        try:
            resultado = df_proveedores[df_proveedores['numero_contrato'] == numero_contrato]
            
            if resultado.empty:
                raise ValueError(f"Contrato {numero_contrato} no encontrado en base de proveedores")
            
            # sucursal = resultado.iloc[0]['sucursal']
            sucursal = resultado.iloc[0]['departamento']
            
            if pd.isna(sucursal) or sucursal == '':
                raise ValueError(f"Sucursal vacia para contrato {numero_contrato}")
            
            return str(sucursal).strip()
        except Exception as e:
            self.logger.error(f"Error al obtener sucursal para contrato {numero_contrato}: {str(e)}")
            raise
    
    def crear_carpeta_sucursal(self, nombre_sucursal):
        """Crea la carpeta de sucursal si no existe"""
        try:
            nombre_sucursal = str(nombre_sucursal).strip()
            ruta_sucursal = self.ruta_origen / nombre_sucursal
            
            if not ruta_sucursal.exists():
                ruta_sucursal.mkdir(parents=True, exist_ok=True)
                self.carpetas_creadas.append(str(ruta_sucursal))
                self.logger.info(f"Carpeta creada: {ruta_sucursal}")
            
            return ruta_sucursal
        except Exception as e:
            self.logger.error(f"Error al crear carpeta para sucursal '{nombre_sucursal}': {str(e)}")
            raise
    
    def mover_archivo(self, ruta_origen_archivo, ruta_destino_archivo):
        """Mueve el archivo a su carpeta de destino"""
        try:
            if ruta_destino_archivo.exists():
                self.logger.warning(f"Archivo ya existe en destino: {ruta_destino_archivo}")
                ruta_destino_archivo = self._generar_nombre_unico(ruta_destino_archivo)
            
            shutil.move(str(ruta_origen_archivo), str(ruta_destino_archivo))
            self.logger.info(f"Archivo movido: {ruta_origen_archivo.name} -> {ruta_destino_archivo.parent.name}/")
            return ruta_destino_archivo
        except Exception as e:
            self.logger.error(f"Error al mover archivo '{ruta_origen_archivo.name}': {str(e)}")
            raise
    
    def _generar_nombre_unico(self, ruta_archivo):
        """Genera un nombre unico si el archivo ya existe"""
        ruta = Path(ruta_archivo)
        contador = 1
        
        while ruta.exists():
            nombre = ruta.stem + f"_{contador}"
            ruta = ruta.parent / (nombre + ruta.suffix)
            contador += 1
        
        return ruta
    
    def obtener_archivos_documento(self):
        """Obtiene lista de archivos en la ruta origen"""
        archivos = []
        
        try:
            for archivo in self.ruta_origen.iterdir():
                if archivo.is_file():
                    nombre_lower = archivo.name.lower()
                    if nombre_lower.endswith(('.xlsx', '.xls', '.docx', '.doc')):
                        if not archivo.name.startswith('~'):
                            archivos.append(archivo.name)
        except Exception as e:
            self.logger.error(f"Error al obtener archivos: {str(e)}")
            raise
        
        return archivos
    
    def procesar_archivo(self, nombre_archivo, df_proveedores):
        """Procesa un archivo individual"""
        try:
            ruta_archivo = self.ruta_origen / nombre_archivo
            
            self.logger.info(f"Procesando archivo: {nombre_archivo}")
            
            datos = self.extraer_datos_archivo(nombre_archivo)
            
            sucursal = self.obtener_sucursal(datos['numero_contrato'], df_proveedores)
            
            ruta_sucursal = self.crear_carpeta_sucursal(sucursal)
            
            ruta_destino = ruta_sucursal / nombre_archivo
            self.mover_archivo(ruta_archivo, ruta_destino)
            
            self.archivos_procesados.append({
                'archivo': nombre_archivo,
                'contrato': datos['numero_contrato'],
                'ano': datos['ano_contrato'],
                'nit': datos['nit'],
                'departamento': sucursal,
                'estado': 'Movido exitosamente'
            })
            
        except Exception as e:
            self.archivos_error.append({
                'archivo': nombre_archivo,
                'error': str(e)
            })
    
    def ejecutar(self):
        """Ejecuta el proceso completo"""
        try:
            self.validar_ruta_origen()
            self.validar_archivo_proveedores()
            
            df_proveedores = self.cargar_base_proveedores()
            
            archivos = self.obtener_archivos_documento()
            
            if not archivos:
                self.logger.info("No se encontraron archivos para procesar")
                print("\nNo se encontraron archivos para procesar")
                return
            
            self.logger.info(f"Se encontraron {len(archivos)} archivos")
            print(f"\nProcesando {len(archivos)} archivos...\n")
            
            for idx, archivo in enumerate(archivos, 1):
                print(f"[{idx}/{len(archivos)}] {archivo}")
                self.procesar_archivo(archivo, df_proveedores)
            
            self.generar_resumen()
            
        except Exception as e:
            self.logger.critical(f"Error critico: {str(e)}")
            print(f"\nError critico: {str(e)}")
            sys.exit(1)
    
    def generar_resumen(self):
        """Genera un resumen del proceso"""
        total = len(self.archivos_procesados) + len(self.archivos_error)
        
        resumen = "\n" + "="*80 + "\n"
        resumen += "RESUMEN DE PROCESAMIENTO\n"
        resumen += "="*80 + "\n"
        resumen += f"Total de archivos procesados: {total}\n"
        resumen += f"Exitosos: {len(self.archivos_procesados)}\n"
        resumen += f"Con error: {len(self.archivos_error)}\n"
        resumen += f"Carpetas creadas: {len(self.carpetas_creadas)}\n"
        
        if self.archivos_procesados:
            resumen += "\n" + "-"*80 + "\n"
            resumen += "ARCHIVOS MOVIDOS EXITOSAMENTE:\n"
            resumen += "-"*80 + "\n"
            for item in self.archivos_procesados:
                resumen += f"Archivo: {item['archivo']}\n"
                resumen += f"  Contrato: {item['contrato']} | Ano: {item['ano']} | NIT: {item['nit']}\n"
                resumen += f"  Sucursal: {item['departamento']}\n\n"
        
        if self.archivos_error:
            resumen += "\n" + "-"*80 + "\n"
            resumen += "ARCHIVOS CON ERROR:\n"
            resumen += "-"*80 + "\n"
            for item in self.archivos_error:
                resumen += f"Archivo: {item['archivo']}\n"
                resumen += f"  Error: {item['error']}\n\n"
        
        if self.carpetas_creadas:
            resumen += "\n" + "-"*80 + "\n"
            resumen += "CARPETAS CREADAS:\n"
            resumen += "-"*80 + "\n"
            for carpeta in self.carpetas_creadas:
                resumen += f"  {carpeta}\n"
        
        resumen += "\n" + "="*80 + "\n"
        
        print(resumen)
        self.logger.info(resumen)
        self.logger.info(f"Proceso completado: {len(self.archivos_procesados)} exitosos, {len(self.archivos_error)} errores")


if __name__ == "__main__":
    try:
        print("GESTOR DE ARCHIVOS ANEXO 9")
        print("="*80)
        
        # RUTA_ORIGEN = input("Ingrese la ruta de la carpeta origen: ").strip()
        # RUTA_BASE_PROVEEDORES = input("Ingrese la ruta del archivo base_proveedores.xlsx: ").strip()

        RUTA_ORIGEN = r'D:\Proyectos\Indicadores\Salidas'
        RUTA_BASE_PROVEEDORES = r'D:\Proyectos\Indicadores\Resources\proveedores.xlsx'
        
        gestor = GestorArchivosExcel(RUTA_ORIGEN, RUTA_BASE_PROVEEDORES)
        gestor.ejecutar()
        
    except KeyboardInterrupt:
        print("\n\nProceso cancelado por el usuario")
        sys.exit(0)
    except Exception as e:
        print(f"\nError fatal: {str(e)}")
        sys.exit(1)


# class GestorArchivosExcel:
#     def __init__(self, ruta_origen, ruta_base_proveedores):
#         self.ruta_origen = Path(ruta_origen)
#         self.ruta_base_proveedores = Path(ruta_base_proveedores)
#         self.archivos_procesados = []
#         self.archivos_error = []
#         self.carpetas_creadas = []
        
#         # Configurar logging
#         self.setup_logging()
        
#     def setup_logging(self):
#         """Configura el sistema de logging"""
#         log_dir = Path("logs")
#         log_dir.mkdir(exist_ok=True)
        
#         timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
#         log_file = log_dir / f"gestion_archivos_{timestamp}.log"
        
#         logging.basicConfig(
#             level=logging.INFO,
#             format='%(asctime)s - %(levelname)s - %(message)s',
#             handlers=[
#                 logging.FileHandler(log_file),
#                 logging.StreamHandler(sys.stdout)
#             ]
#         )
        
#         self.logger = logging.getLogger(__name__)
#         self.logger.info(f"Iniciando gestión de archivos desde: {self.ruta_origen}")
        
#     def validar_ruta_origen(self):
#         """Valida que la ruta origen existe"""
#         if not self.ruta_origen.exists():
#             self.logger.error(f"La ruta origen no existe: {self.ruta_origen}")
#             raise ValueError(f"Ruta no válida: {self.ruta_origen}")
        
#         if not self.ruta_origen.is_dir():
#             self.logger.error(f"La ruta no es un directorio: {self.ruta_origen}")
#             raise ValueError(f"No es un directorio: {self.ruta_origen}")
        
#         self.logger.info(f"Ruta origen validada correctamente")
        
#     def validar_archivo_proveedores(self):
#         """Valida que el archivo de proveedores existe"""
#         if not self.ruta_base_proveedores.exists():
#             self.logger.error(f"Archivo de proveedores no encontrado: {self.ruta_base_proveedores}")
#             raise FileNotFoundError(f"No existe: {self.ruta_base_proveedores}")
        
#         self.logger.info(f"Archivo de proveedores validado")
        
#     def cargar_base_proveedores(self):
#         """Carga la base de datos de proveedores"""
#         try:
#             df = pd.read_excel(self.ruta_base_proveedores)
            
#             # Validar que existan las columnas necesarias
#             if 'numero_contrato' not in df.columns or 'sucursal' not in df.columns:
#                 raise ValueError("Columnas 'numero_contrato' o 'sucursal' no encontradas")
            
#             # Convertir número de contrato a string con padding de ceros
#             df['numero_contrato'] = df['numero_contrato'].astype(str).str.zfill(4)
            
#             self.logger.info(f"Base de proveedores cargada: {len(df)} registros")
#             return df
#         except Exception as e:
#             self.logger.error(f"Error al cargar base de proveedores: {str(e)}")
#             raise
    
#     def extraer_datos_archivo(self, nombre_archivo):
#         """
#         Extrae número de contrato, año y NIT del nombre del archivo
#         Formato esperado: Anexo X 662-2024 NOMBRE NIT
#         """
#         try:
#             # Remover extensión
#             nombre_sin_ext = nombre_archivo.replace('.xlsx', '').replace('.xls', '')
            
#             # Patrón para extraer: número-año
#             patron = r'(\d+)-(\d{4})\s'
#             match = re.search(patron, nombre_sin_ext)
            
#             if not match:
#                 raise ValueError(f"No se encontró patrón número_contrato-año")
            
#             numero_contrato = match.group(1)
#             ano_contrato = match.group(2)
            
#             # Extraer NIT (últimos 10 dígitos)
#             numeros = re.findall(r'\d+', nombre_sin_ext)
#             if len(numeros) < 2:
#                 raise ValueError(f"No se pudo extraer NIT del nombre")
            
#             nit = numeros[-1]
            
#             # Validar que número de contrato tenga máximo 4 dígitos
#             if len(numero_contrato) > 4:
#                 raise ValueError(f"Número de contrato inválido: {numero_contrato} (más de 4 dígitos)")
            
#             # Rellenar con ceros a la izquierda
#             numero_contrato_formateado = numero_contrato.zfill(4)
            
#             self.logger.debug(f"Datos extraídos - Contrato: {numero_contrato_formateado}, Año: {ano_contrato}, NIT: {nit}")
            
#             return {
#                 'numero_contrato': numero_contrato_formateado,
#                 'ano_contrato': ano_contrato,
#                 'nit': nit,
#                 'numero_original': numero_contrato
#             }
#         except Exception as e:
#             self.logger.error(f"Error al extraer datos de '{nombre_archivo}': {str(e)}")
#             raise
    
#     def obtener_sucursal(self, numero_contrato, df_proveedores):
#         """Obtiene la sucursal de la base de proveedores"""
#         try:
#             resultado = df_proveedores[df_proveedores['numero_contrato'] == numero_contrato]
            
#             if resultado.empty:
#                 raise ValueError(f"Contrato {numero_contrato} no encontrado en base de proveedores")
            
#             sucursal = resultado.iloc[0]['sucursal']
            
#             if pd.isna(sucursal) or sucursal == '':
#                 raise ValueError(f"Sucursal vacía para contrato {numero_contrato}")
            
#             return str(sucursal).strip()
#         except Exception as e:
#             self.logger.error(f"Error al obtener sucursal para contrato {numero_contrato}: {str(e)}")
#             raise
    
#     def crear_carpeta_sucursal(self, nombre_sucursal):
#         """Crea la carpeta de sucursal si no existe"""
#         try:
#             ruta_sucursal = self.ruta_origen / nombre_sucursal
            
#             if not ruta_sucursal.exists():
#                 ruta_sucursal.mkdir(parents=True, exist_ok=True)
#                 self.carpetas_creadas.append(str(ruta_sucursal))
#                 self.logger.info(f"Carpeta creada: {ruta_sucursal}")
            
#             return ruta_sucursal
#         except Exception as e:
#             self.logger.error(f"Error al crear carpeta para sucursal '{nombre_sucursal}': {str(e)}")
#             raise
    
#     def mover_archivo(self, ruta_origen_archivo, ruta_destino_archivo):
#         """Mueve el archivo a su carpeta de destino"""
#         try:
#             if ruta_destino_archivo.exists():
#                 self.logger.warning(f"Archivo ya existe en destino: {ruta_destino_archivo}")
#                 ruta_destino_archivo = self._generar_nombre_unico(ruta_destino_archivo)
            
#             shutil.move(str(ruta_origen_archivo), str(ruta_destino_archivo))
#             self.logger.info(f"Archivo movido: {ruta_origen_archivo.name} → {ruta_destino_archivo.parent.name}/")
#             return True
#         except Exception as e:
#             self.logger.error(f"Error al mover archivo '{ruta_origen_archivo.name}': {str(e)}")
#             raise
    
#     def _generar_nombre_unico(self, ruta_archivo):
#         """Genera un nombre único si el archivo ya existe"""
#         ruta = Path(ruta_archivo)
#         contador = 1
        
#         while ruta.exists():
#             nombre = ruta.stem + f"_{contador}"
#             ruta = ruta.parent / (nombre + ruta.suffix)
#             contador += 1
        
#         return ruta
    
#     def procesar_archivo(self, nombre_archivo, df_proveedores):
#         """Procesa un archivo individual"""
#         try:
#             ruta_archivo = self.ruta_origen / nombre_archivo
            
#             # Validar que es archivo Excel
#             if not nombre_archivo.lower().endswith(('.xlsx', '.xls')):
#                 raise ValueError("Archivo no es Excel (.xlsx o .xls)")
            
#             self.logger.info(f"Procesando archivo: {nombre_archivo}")
            
#             # Extraer datos del nombre
#             datos = self.extraer_datos_archivo(nombre_archivo)
            
#             # Obtener sucursal
#             sucursal = self.obtener_sucursal(datos['numero_contrato'], df_proveedores)
            
#             # Crear carpeta si no existe
#             ruta_sucursal = self.crear_carpeta_sucursal(sucursal)
            
#             # Mover archivo
#             ruta_destino = ruta_sucursal / nombre_archivo
#             self.mover_archivo(ruta_archivo, ruta_destino)
            
#             # Registrar éxito
#             self.archivos_procesados.append({
#                 'archivo': nombre_archivo,
#                 'contrato': datos['numero_contrato'],
#                 'año': datos['ano_contrato'],
#                 'nit': datos['nit'],
#                 'sucursal': sucursal,
#                 'estado': 'Movido exitosamente'
#             })
            
#         except Exception as e:
#             self.archivos_error.append({
#                 'archivo': nombre_archivo,
#                 'error': str(e)
#             })
#             self.logger.error(f"Falló procesamiento de '{nombre_archivo}': {str(e)}")
    
#     def obtener_archivos_excel(self):
#         """Obtiene lista de archivos Excel en la ruta origen"""
#         archivos = []
        
#         for archivo in self.ruta_origen.iterdir():
#             if archivo.is_file() and archivo.name.lower().endswith(('.xlsx', '.xls')):
#                 # Evitar archivos temporales de Excel
#                 if not archivo.name.startswith('~$'):
#                     archivos.append(archivo.name)
        
#         return archivos
    
#     def ejecutar(self):
#         """Ejecuta el proceso completo"""
#         try:
#             # Validaciones iniciales
#             self.validar_ruta_origen()
#             self.validar_archivo_proveedores()
            
#             # Cargar base de proveedores
#             df_proveedores = self.cargar_base_proveedores()
            
#             # Obtener archivos Excel
#             archivos = self.obtener_archivos_excel()
            
#             if not archivos:
#                 self.logger.info("No se encontraron archivos Excel para procesar")
#                 return
            
#             self.logger.info(f"Se encontraron {len(archivos)} archivos Excel")
            
#             # Procesar cada archivo
#             for archivo in archivos:
#                 self.procesar_archivo(archivo, df_proveedores)
            
#             # Generar resumen
#             self.generar_resumen()
            
#         except Exception as e:
#             self.logger.critical(f"Error crítico: {str(e)}")
#             print(f"\n❌ Error crítico: {str(e)}")
#             sys.exit(1)
    
#     def generar_resumen(self):
#         """Genera un resumen del proceso"""
#         total = len(self.archivos_procesados) + len(self.archivos_error)
        
#         print("\n" + "="*70)
#         print("RESUMEN DE PROCESAMIENTO".center(70))
#         print("="*70)
#         print(f"Total de archivos procesados: {total}")
#         print(f"✅ Exitosos: {len(self.archivos_procesados)}")
#         print(f"❌ Con error: {len(self.archivos_error)}")
#         print(f"📁 Carpetas creadas: {len(self.carpetas_creadas)}")
        
#         if self.archivos_procesados:
#             print("\n" + "-"*70)
#             print("ARCHIVOS MOVIDOS EXITOSAMENTE:")
#             print("-"*70)
#             for item in self.archivos_procesados:
#                 print(f"  📄 {item['archivo']}")
#                 print(f"     Contrato: {item['contrato']} | Año: {item['año']} | NIT: {item['nit']}")
#                 print(f"     → Sucursal: {item['sucursal']}")
        
#         if self.archivos_error:
#             print("\n" + "-"*70)
#             print("ARCHIVOS CON ERROR:")
#             print("-"*70)
#             for item in self.archivos_error:
#                 print(f"  ❌ {item['archivo']}")
#                 print(f"     Error: {item['error']}")
        
#         if self.carpetas_creadas:
#             print("\n" + "-"*70)
#             print("CARPETAS CREADAS:")
#             print("-"*70)
#             for carpeta in self.carpetas_creadas:
#                 print(f"  📁 {carpeta}")
        
#         print("\n" + "="*70)
#         self.logger.info(f"Proceso completado: {len(self.archivos_procesados)} exitosos, {len(self.archivos_error)} errores")


# if __name__ == "__main__":
#     # Configuración
#     # RUTA_ORIGEN = input("Ingrese la ruta de la carpeta origen: ").strip()
#     # RUTA_BASE_PROVEEDORES = input("Ingrese la ruta del archivo base_proveedores.xlsx: ").strip()
#     RUTA_ORIGEN = r'D:\Proyectos\Indicadores\Salidas'
#     RUTA_BASE_PROVEEDORES = r'D:\Proyectos\Indicadores\Resources\proveedores.xlsx'
    
#     # Crear gestor y ejecutar
#     gestor = GestorArchivosExcel(RUTA_ORIGEN, RUTA_BASE_PROVEEDORES)
#     gestor.ejecutar()