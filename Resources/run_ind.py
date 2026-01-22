import subprocess
import time

def ejecutar_otro_script(ruta_script):
    try:
        print("Esperando 2 segundos...")
        time.sleep(2)
        print(f"Ejecutando {ruta_script}...")
        
        resultado = subprocess.run(["python", ruta_script])
        
        if resultado.returncode == 0:
            print(f"✅ {ruta_script} ejecutado exitosamente.")
            print(resultado.stdout)
        else:
            print(f"❌ Error al ejecutar {ruta_script}.")
            print(resultado.stderr)
    
    except Exception as e:
        print(f"⚠️ Error al intentar ejecutar {ruta_script}: {e}")

# Rutas completas a tus scripts
scripts = [
    r"D:\Proyectos\Indicadores\Resources\indicadores_end.py",
    # r"D:\Proyectos\Indicadores\Resources\convert_pdf_threading.py",
]

# Ejecutar los scripts en orden
for script in scripts:
    ejecutar_otro_script(script)
