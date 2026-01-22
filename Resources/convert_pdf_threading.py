import os
import shutil
from pathlib import Path
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
from queue import Queue

# Intentar importar las librerías de conversión
try:
    from docx2pdf import convert as docx2pdf_convert
    DOCX2PDF_AVAILABLE = True
except ImportError:
    DOCX2PDF_AVAILABLE = False

try:
    import win32com.client as win32
    import pythoncom
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

# --- Configuración de rutas ---
BASE_DIR_WORD = r'D:\Proyectos\Indicadores\Salidas'  # Carpeta donde están los Word generados
BASE_DIR_PDF = r'D:\Proyectos\Indicadores\Salidas'   # Carpeta donde se guardan los PDF (misma carpeta)
COPIE_DIR = r'C:\Users\JORGEEDILSONVEGAACOS\One Drive Analista Red 11\OneDrive - Positiva Compañia de Seguros S. A\Indicadores\PDF'

# Lock para operaciones de impresión thread-safe
print_lock = threading.Lock()

def thread_safe_print(*args, **kwargs):
    """Función de impresión thread-safe"""
    with print_lock:
        print(*args, **kwargs)

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
        thread_safe_print("❌ ERROR: No hay métodos de conversión disponibles.")
        thread_safe_print("   Instalar con: pip install docx2pdf pywin32")
        return False
    
    thread_safe_print(f"✅ Métodos disponibles: {', '.join(metodos)}")
    return True

def convertir_word_a_pdf_win32_threadsafe(ruta_word, ruta_pdf):
    """
    Convierte Word a PDF usando la aplicación Word con manejo completo de COM.
    """
    word_app = None
    doc = None
    try:
        # Inicializar COM para este thread de manera completa
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        
        # Crear una nueva instancia de Word para este thread
        word_app = win32.gencache.EnsureDispatch('Word.Application')
        word_app.Visible = False
        word_app.DisplayAlerts = False
        word_app.ScreenUpdating = False
        
        # Abrir documento
        doc = word_app.Documents.Open(ruta_word)
        
        # Guardar como PDF
        doc.SaveAs(ruta_pdf, FileFormat=17)  # 17 = PDF
        
        return True
        
    except Exception as e:
        thread_safe_print(f"   Error con win32com: {e}")
        return False
    finally:
        # Cerrar documento y aplicación de manera segura
        try:
            if doc:
                doc.Close(SaveChanges=False)
        except:
            pass
        try:
            if word_app:
                word_app.Quit()
        except:
            pass
        # Limpiar COM
        try:
            pythoncom.CoUninitialize()
        except:
            pass

def convertir_word_a_pdf_docx2pdf_threadsafe(ruta_word, ruta_pdf):
    """
    Convierte Word a PDF usando docx2pdf con manejo COM.
    """
    try:
        # docx2pdf también necesita COM inicializado
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        
        # Realizar conversión
        docx2pdf_convert(ruta_word, ruta_pdf)
        
        return True
    except Exception as e:
        thread_safe_print(f"   Error con docx2pdf: {e}")
        return False
    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass

def convertir_archivo_worker(ruta_word, task_num, total_tasks, usar_solo_docx2pdf=False):
    """
    Worker que convierte un archivo Word a PDF (versión thread-safe).
    """
    nombre_archivo = os.path.basename(ruta_word)
    ruta_pdf = ruta_word.replace('.docx', '.pdf')
    
    thread_safe_print(f"🔄 [{task_num}/{total_tasks}] Convirtiendo: {nombre_archivo}")
    
    # Verificar si ya existe el PDF y es más reciente que el Word
    if os.path.exists(ruta_pdf):
        tiempo_word = os.path.getmtime(ruta_word)
        tiempo_pdf = os.path.getmtime(ruta_pdf)
        if tiempo_pdf > tiempo_word:
            thread_safe_print(f"   ✅ [{task_num}] PDF ya existe y es más reciente: {os.path.basename(ruta_pdf)}")
            return ruta_pdf, True, "skipped"
    
    exito = False
    metodo_usado = ""
    
    # Estrategia de conversión
    if usar_solo_docx2pdf and DOCX2PDF_AVAILABLE:
        # Solo usar docx2pdf (más estable en multithreading)
        try:
            if convertir_word_a_pdf_docx2pdf_threadsafe(ruta_word, ruta_pdf):
                thread_safe_print(f"   ✅ [{task_num}] Éxito con docx2pdf: {os.path.basename(ruta_pdf)}")
                exito = True
                metodo_usado = "docx2pdf"
        except Exception as e:
            thread_safe_print(f"   ❌ [{task_num}] Falló docx2pdf: {e}")
    else:
        # Intentar win32com primero, luego docx2pdf
        if WIN32_AVAILABLE:
            try:
                if convertir_word_a_pdf_win32_threadsafe(ruta_word, ruta_pdf):
                    thread_safe_print(f"   ✅ [{task_num}] Éxito con Word Application: {os.path.basename(ruta_pdf)}")
                    exito = True
                    metodo_usado = "win32"
            except Exception as e:
                thread_safe_print(f"   ❌ [{task_num}] Falló Word Application: {e}")
        
        # Si falló win32com, intentar docx2pdf
        if not exito and DOCX2PDF_AVAILABLE:
            try:
                if convertir_word_a_pdf_docx2pdf_threadsafe(ruta_word, ruta_pdf):
                    thread_safe_print(f"   ✅ [{task_num}] Éxito con docx2pdf: {os.path.basename(ruta_pdf)}")
                    exito = True
                    metodo_usado = "docx2pdf"
            except Exception as e:
                thread_safe_print(f"   ❌ [{task_num}] Falló docx2pdf: {e}")
    
    if not exito:
        thread_safe_print(f"   ❌ [{task_num}] FALLÓ: No se pudo convertir {nombre_archivo}")
        return ruta_pdf, False, "failed"
    
    return ruta_pdf, True, metodo_usado

def copiar_pdf_a_onedrive(ruta_pdf, task_num):
    """
    Copia un archivo PDF a la carpeta de OneDrive (versión thread-safe).
    """
    try:
        os.makedirs(COPIE_DIR, exist_ok=True)
        destino = os.path.join(COPIE_DIR, os.path.basename(ruta_pdf))
        shutil.copy2(ruta_pdf, destino)
        thread_safe_print(f"   📁 [{task_num}] Copiado a OneDrive")
        return True
    except Exception as e:
        thread_safe_print(f"   ❌ [{task_num}] Error copiando a OneDrive: {e}")
        return False

def obtener_archivos_word(carpeta, patron="Anexo 9*.docx"):
    """
    Obtiene todos los archivos Word que coincidan con el patrón.
    """
    archivos = []
    carpeta_path = Path(carpeta)
    
    if not carpeta_path.exists():
        thread_safe_print(f"❌ ERROR: La carpeta {carpeta} no existe")
        return archivos
    
    # Buscar archivos que coincidan con el patrón
    for archivo in carpeta_path.glob(patron):
        if archivo.is_file():
            archivos.append(str(archivo))
    
    return archivos

def procesar_conversion_multihilo(archivos_word, copiar_a_onedrive=True, max_workers=2, usar_solo_docx2pdf=False):
    """
    Procesa la conversión de archivos uno por uno pero usando multiples threads.
    Optimizado para manejar COM correctamente.
    """
    metodo_str = "solo docx2pdf" if usar_solo_docx2pdf else "win32com + docx2pdf fallback"
    thread_safe_print(f"\n🔄 Iniciando conversión multihilo con {max_workers} workers")
    thread_safe_print(f"   📋 Modo: {metodo_str}")
    thread_safe_print("   ⚠️ COM se inicializa/limpia correctamente en cada thread")
    thread_safe_print("   ✅ Orden secuencial garantizado\n")
    
    estadisticas = {
        'total': len(archivos_word),
        'exitosos': 0,
        'fallidos': 0,
        'saltados': 0,
        'copiados_onedrive': 0,
        'metodos_usados': {}
    }
    
    # Procesar archivos de forma secuencial pero con múltiples workers
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        for i, ruta_word in enumerate(archivos_word, 1):
            # Enviar tarea al pool de threads
            future = executor.submit(convertir_archivo_worker, ruta_word, i, len(archivos_word), usar_solo_docx2pdf)
            
            try:
                # Esperar a que termine esta tarea específica antes de continuar
                ruta_pdf, exito, metodo = future.result()
                
                if exito:
                    if metodo == "skipped":
                        estadisticas['saltados'] += 1
                    else:
                        estadisticas['exitosos'] += 1
                        # Contar métodos usados
                        if metodo in estadisticas['metodos_usados']:
                            estadisticas['metodos_usados'][metodo] += 1
                        else:
                            estadisticas['metodos_usados'][metodo] = 1
                    
                    # Copiar a OneDrive si se solicita y el archivo existe
                    if copiar_a_onedrive and os.path.exists(ruta_pdf):
                        if copiar_pdf_a_onedrive(ruta_pdf, i):
                            estadisticas['copiados_onedrive'] += 1
                else:
                    estadisticas['fallidos'] += 1
                    
            except Exception as e:
                thread_safe_print(f"   ❌ [{i}] Error procesando {os.path.basename(ruta_word)}: {e}")
                estadisticas['fallidos'] += 1
    
    return estadisticas

def mostrar_estadisticas(estadisticas, tiempo_transcurrido):
    """
    Muestra las estadísticas finales del proceso.
    """
    print(f"\n{'='*60}")
    print(f"📊 ESTADÍSTICAS FINALES")
    print(f"{'='*60}")
    print(f"Total de archivos:       {estadisticas['total']}")
    print(f"Conversiones exitosas:   {estadisticas['exitosos']}")
    print(f"Archivos ya actualizados: {estadisticas['saltados']}")
    print(f"Conversiones fallidas:   {estadisticas['fallidos']}")
    print(f"Copiados a OneDrive:     {estadisticas['copiados_onedrive']}")
    
    if estadisticas['metodos_usados']:
        print(f"\nMétodos de conversión utilizados:")
        for metodo, count in estadisticas['metodos_usados'].items():
            print(f"  • {metodo}: {count} archivo(s)")
    
    print(f"\nTiempo transcurrido:     {tiempo_transcurrido:.2f} segundos")
    
    if estadisticas['total'] > 0:
        porcentaje_exito = ((estadisticas['exitosos'] + estadisticas['saltados']) / estadisticas['total']) * 100
        print(f"Porcentaje de éxito:     {porcentaje_exito:.1f}%")
        
        if tiempo_transcurrido > 0:
            archivos_por_minuto = (estadisticas['total'] / tiempo_transcurrido) * 60
            print(f"Velocidad promedio:      {archivos_por_minuto:.1f} archivos/minuto")

def main():
    """
    Función principal del convertidor.
    """
    print("🚀 CONVERTIDOR DE WORD A PDF - MULTIHILO CON COM THREAD-SAFE")
    print("=" * 70)
    
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
    
    # 4. Elegir modo de conversión
    print("\n🔧 MODOS DE CONVERSIÓN DISPONIBLES:")
    print("1. 🏆 Híbrido: win32com + docx2pdf fallback (mejor calidad)")
    print("2. ⚡ Solo docx2pdf (más estable, menor calidad)")
    print("3. 🐌 Un solo worker (máxima estabilidad)")
    
    try:
        modo = input("\nSelecciona modo [1]: ").strip()
        if not modo:
            modo = "1"
    except KeyboardInterrupt:
        print("\n❌ Proceso cancelado")
        return
    
    usar_solo_docx2pdf = (modo == "2")
    usar_un_worker = (modo == "3")
    
    # 5. Configurar workers
    if usar_un_worker:
        max_workers = 1
        print("   ✅ Usando 1 worker (modo súper estable)")
    else:
        print(f"\n⚙️ CONFIGURACIÓN WORKERS:")
        print("   💡 Recomendado: 2 workers para balance estabilidad/velocidad")
        
        try:
            workers_input = input("Número de workers [2]: ").strip()
            max_workers = int(workers_input) if workers_input else 2
            max_workers = max(1, min(max_workers, 6))  # Entre 1 y 6 workers
        except (ValueError, KeyboardInterrupt):
            print("\n❌ Proceso cancelado o valor inválido")
            return
    
    # 6. Confirmar procesamiento
    print(f"\n📋 CONFIGURACIÓN FINAL:")
    print(f"   • Archivos a procesar: {len(archivos_word)}")
    print(f"   • Workers: {max_workers}")
    print(f"   • Método: {'Solo docx2pdf' if usar_solo_docx2pdf else 'Híbrido (win32com + docx2pdf)'}")
    print(f"   • Copia a OneDrive: Sí")
    
    confirmar = input("\n¿Proceder con la conversión? [S/n]: ").strip().lower()
    if confirmar == 'n':
        print("❌ Proceso cancelado")
        return
    
    # 7. Iniciar conversión
    inicio = time.time()
    estadisticas = procesar_conversion_multihilo(
        archivos_word, 
        max_workers=max_workers, 
        usar_solo_docx2pdf=usar_solo_docx2pdf
    )
    tiempo_transcurrido = time.time() - inicio
    
    # 8. Mostrar estadísticas
    mostrar_estadisticas(estadisticas, tiempo_transcurrido)
    
    # 9. Recomendaciones si hubo fallos
    if estadisticas['fallidos'] > 0:
        print(f"\n💡 RECOMENDACIONES PARA {estadisticas['fallidos']} ARCHIVOS FALLIDOS:")
        print("   1. Intentar con 1 worker (modo 3)")
        print("   2. Verificar que los archivos Word no estén corruptos")
        print("   3. Ejecutar Word manualmente y cerrar todas las instancias")
        print("   4. Reiniciar el sistema si persisten los problemas COM")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n❌ Proceso interrumpido por el usuario")
    except Exception as e:
        print(f"\n❌ Error inesperado: {e}")
        import traceback
        traceback.print_exc()