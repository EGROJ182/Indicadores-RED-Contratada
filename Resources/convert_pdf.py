import os
import shutil
from pathlib import Path
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

# Intentar importar las librerías de conversión
try:
    from docx2pdf import convert as docx2pdf_convert
    DOCX2PDF_AVAILABLE = True
except ImportError:
    DOCX2PDF_AVAILABLE = False

try:
    import win32com.client as win32
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

# --- Configuración de rutas ---
BASE_DIR_WORD = r'D:\Proyectos\Indicadores\Salidas'  # Carpeta donde están los Word generados
BASE_DIR_PDF = r'D:\Proyectos\Indicadores\Salidas'   # Carpeta donde se guardan los PDF (misma carpeta)
COPIE_DIR = r'C:\Users\JORGEEDILSONVEGAACOS\One Drive Analista Red 11\OneDrive - Positiva Compañia de Seguros S. A\Indicadores\IND New'

def verificar_dependencias():
    """
    Verifica qué métodos de conversión están disponibles.
    """
    metodos = []
    if WIN32_AVAILABLE:
        metodos.append("Word Application (win32com)")
    if DOCX2PDF_AVAILABLE:
        metodos.append("docx2pdf")
    
    if not metodos:
        print("❌ ERROR: No hay métodos de conversión disponibles.")
        print("   Instalar con: pip install docx2pdf pywin32")
        return False
    
    print(f"✅ Métodos disponibles: {', '.join(metodos)}")
    return True

def convertir_word_a_pdf_win32(ruta_word, ruta_pdf):
    """
    Convierte Word a PDF usando la aplicación Word (método más confiable).
    """
    try:
        word_app = win32.Dispatch('Word.Application')
        word_app.Visible = False
        word_app.DisplayAlerts = False  # Evitar diálogos
        
        doc = word_app.Documents.Open(ruta_word)
        doc.SaveAs(ruta_pdf, FileFormat=17)  # 17 = PDF
        doc.Close()
        word_app.Quit()
        
        return True
    except Exception as e:
        print(f"   Error con win32com: {e}")
        return False

def convertir_word_a_pdf_docx2pdf(ruta_word, ruta_pdf):
    """
    Convierte Word a PDF usando docx2pdf.
    """
    try:
        docx2pdf_convert(ruta_word, ruta_pdf)
        return True
    except Exception as e:
        print(f"   Error con docx2pdf: {e}")
        return False

def convertir_archivo(ruta_word, intentar_win32_primero=True):
    """
    Convierte un archivo Word a PDF probando los métodos disponibles.
    """
    nombre_archivo = os.path.basename(ruta_word)
    ruta_pdf = ruta_word.replace('.docx', '.pdf')
    
    print(f"🔄 Convirtiendo: {nombre_archivo}")
    
    # Verificar si ya existe el PDF y es más reciente que el Word
    if os.path.exists(ruta_pdf):
        tiempo_word = os.path.getmtime(ruta_word)
        tiempo_pdf = os.path.getmtime(ruta_pdf)
        if tiempo_pdf > tiempo_word:
            print(f"   ✅ PDF ya existe y es más reciente: {os.path.basename(ruta_pdf)}")
            return ruta_pdf, True
    
    exito = False
    
    # Definir el orden de métodos a probar
    if intentar_win32_primero and WIN32_AVAILABLE:
        metodos = [
            ("Word Application", convertir_word_a_pdf_win32),
            ("docx2pdf", convertir_word_a_pdf_docx2pdf) if DOCX2PDF_AVAILABLE else None
        ]
    else:
        metodos = [
            ("docx2pdf", convertir_word_a_pdf_docx2pdf) if DOCX2PDF_AVAILABLE else None,
            ("Word Application", convertir_word_a_pdf_win32) if WIN32_AVAILABLE else None
        ]
    
    # Filtrar métodos None
    metodos = [m for m in metodos if m is not None]
    
    for nombre_metodo, funcion_conversion in metodos:
        try:
            if funcion_conversion(ruta_word, ruta_pdf):
                print(f"   ✅ Éxito con {nombre_metodo}: {os.path.basename(ruta_pdf)}")
                exito = True
                break
        except Exception as e:
            print(f"   ❌ Falló {nombre_metodo}: {e}")
    
    if not exito:
        print(f"   ❌ FALLÓ: No se pudo convertir {nombre_archivo}")
        return ruta_pdf, False
    
    return ruta_pdf, True

def copiar_pdf_a_onedrive(ruta_pdf):
    """
    Copia un archivo PDF a la carpeta de OneDrive.
    """
    try:
        os.makedirs(COPIE_DIR, exist_ok=True)
        destino = os.path.join(COPIE_DIR, os.path.basename(ruta_pdf))
        shutil.copy2(ruta_pdf, destino)
        return True
    except Exception as e:
        print(f"   ❌ Error copiando a OneDrive: {e}")
        return False

def obtener_archivos_word(carpeta, patron="Anexo 9*.docx"):
    """
    Obtiene todos los archivos Word que coincidan con el patrón.
    """
    archivos = []
    carpeta_path = Path(carpeta)
    
    if not carpeta_path.exists():
        print(f"❌ ERROR: La carpeta {carpeta} no existe")
        return archivos
    
    # Buscar archivos que coincidan con el patrón
    for archivo in carpeta_path.glob(patron):
        if archivo.is_file():
            archivos.append(str(archivo))
    
    return archivos

def procesar_conversion_secuencial(archivos_word, copiar_a_onedrive=True):
    """
    Procesa la conversión de archivos de forma secuencial.
    """
    print(f"\n🔄 Iniciando conversión secuencial de {len(archivos_word)} archivos...")
    
    estadisticas = {
        'total': len(archivos_word),
        'exitosos': 0,
        'fallidos': 0,
        'copiados_onedrive': 0
    }
    
    for i, ruta_word in enumerate(archivos_word, 1):
        print(f"\n[{i}/{len(archivos_word)}]")
        ruta_pdf, exito = convertir_archivo(ruta_word)
        
        if exito:
            estadisticas['exitosos'] += 1
            
            # Copiar a OneDrive si se solicita
            if copiar_a_onedrive and os.path.exists(ruta_pdf):
                if copiar_pdf_a_onedrive(ruta_pdf):
                    estadisticas['copiados_onedrive'] += 1
                    print(f"   📁 Copiado a OneDrive")
        else:
            estadisticas['fallidos'] += 1
    
    return estadisticas

def procesar_conversion_paralela(archivos_word, copiar_a_onedrive=True, max_workers=2):
    """
    Procesa la conversión de archivos en paralelo (más rápido pero puede usar más recursos).
    """
    print(f"\n🔄 Iniciando conversión paralela con {max_workers} workers...")
    
    estadisticas = {
        'total': len(archivos_word),
        'exitosos': 0,
        'fallidos': 0,
        'copiados_onedrive': 0
    }
    
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        # Enviar todas las tareas
        future_to_archivo = {
            executor.submit(convertir_archivo, archivo): archivo 
            for archivo in archivos_word
        }
        
        # Procesar resultados conforme van completándose
        for i, future in enumerate(as_completed(future_to_archivo), 1):
            archivo = future_to_archivo[future]
            print(f"\n[{i}/{len(archivos_word)}]")
            
            try:
                ruta_pdf, exito = future.result()
                
                if exito:
                    estadisticas['exitosos'] += 1
                    
                    # Copiar a OneDrive si se solicita
                    if copiar_a_onedrive and os.path.exists(ruta_pdf):
                        if copiar_pdf_a_onedrive(ruta_pdf):
                            estadisticas['copiados_onedrive'] += 1
                            print(f"   📁 Copiado a OneDrive")
                else:
                    estadisticas['fallidos'] += 1
                    
            except Exception as e:
                print(f"   ❌ Error procesando {os.path.basename(archivo)}: {e}")
                estadisticas['fallidos'] += 1
    
    return estadisticas

def mostrar_estadisticas(estadisticas, tiempo_transcurrido):
    """
    Muestra las estadísticas finales del proceso.
    """
    print(f"\n{'='*50}")
    print(f"📊 ESTADÍSTICAS FINALES")
    print(f"{'='*50}")
    print(f"Total de archivos:      {estadisticas['total']}")
    print(f"Conversiones exitosas:  {estadisticas['exitosos']}")
    print(f"Conversiones fallidas:  {estadisticas['fallidos']}")
    print(f"Copiados a OneDrive:    {estadisticas['copiados_onedrive']}")
    print(f"Tiempo transcurrido:    {tiempo_transcurrido:.2f} segundos")
    
    if estadisticas['total'] > 0:
        porcentaje_exito = (estadisticas['exitosos'] / estadisticas['total']) * 100
        print(f"Porcentaje de éxito:    {porcentaje_exito:.1f}%")

def main():
    """
    Función principal del convertidor.
    """
    print("🚀 CONVERTIDOR DE WORD A PDF")
    print("=" * 50)
    
    # 1. Verificar dependencias
    if not verificar_dependencias():
        return
    
    # 2. Verificar carpetas
    if not os.path.exists(BASE_DIR_WORD):
        print(f"❌ ERROR: La carpeta de archivos Word no existe: {BASE_DIR_WORD}")
        return
    
    print(f"📁 Carpeta Word:     {BASE_DIR_WORD}")
    print(f"📁 Carpeta PDF:      {BASE_DIR_PDF}")
    print(f"📁 OneDrive:         {COPIE_DIR}")
    
    # 3. Buscar archivos Word
    archivos_word = obtener_archivos_word(BASE_DIR_WORD)
    
    if not archivos_word:
        print("❌ No se encontraron archivos Word para convertir")
        return
    
    print(f"\n✅ Encontrados {len(archivos_word)} archivos Word")
    
    # 4. Preguntar tipo de procesamiento
    print("\n¿Cómo deseas procesar los archivos?")
    print("1. Secuencial (uno por uno, más estable, usa Word si está disponible)")
    print("2. Paralelo (más rápido, usa docx2pdf por estabilidad)")
    print("3. Solo mostrar archivos encontrados")
    
    try:
        opcion = input("\nSelecciona una opción (1-3) [1]: ").strip()
        if not opcion:
            opcion = "1"
    except KeyboardInterrupt:
        print("\n❌ Proceso cancelado por el usuario")
        return
    
    if opcion == "3":
        print("\n📋 ARCHIVOS ENCONTRADOS:")
        for i, archivo in enumerate(archivos_word, 1):
            print(f"  {i}. {os.path.basename(archivo)}")
        return
    
    # 5. Iniciar conversión
    inicio = time.time()
    
    if opcion == "2":
        print("\n⚠️  NOTA: El modo paralelo usa docx2pdf por defecto para evitar problemas COM")
        print("   Si necesitas máxima calidad de conversión, usa el modo secuencial")
        estadisticas = procesar_conversion_paralela(archivos_word)
    else:
        estadisticas = procesar_conversion_secuencial(archivos_word)
    
    tiempo_transcurrido = time.time() - inicio
    
    # 6. Mostrar estadísticas
    mostrar_estadisticas(estadisticas, tiempo_transcurrido)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n❌ Proceso interrumpido por el usuario")
    except Exception as e:
        print(f"\n❌ Error inesperado: {e}")
        import traceback
        traceback.print_exc()
