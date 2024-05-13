import schedule
import time
import subprocess

def ejecutar_archivo():
    try:
        # Ruta del archivo que deseas ejecutar
        archivo_a_ejecutar = "BODY EMAIL TABLE_V0.9.py"
        # Ejecutar el archivo proporcionado como argumento
        subprocess.run(["python", archivo_a_ejecutar], check=True)
    except subprocess.CalledProcessError as e:
        print(f"Error al ejecutar el archivo: {e}")

# Programar la ejecución del archivo cada 5 minutos
schedule.every(5).minutes.do(ejecutar_archivo)

# Ejecutar el programa hasta que finalice la hora de fin
while True:
    # Verificar si hay tareas programadas para ejecutar
    schedule.run_pending()
    time.sleep(60)  # Esperar 1 minuto antes de verificar de nuevo




