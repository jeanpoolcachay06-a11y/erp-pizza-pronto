# ERP Pizza Pronto

Interfaz académica desarrollada con Streamlit y Python. Consulta datos de un libro Excel y permite registrar clientes y materiales cuando se ejecuta en una computadora con el Excel local.

## Ejecutar en VS Code

1. Instala Python 3.10 o posterior.
2. Abre esta carpeta en VS Code.
3. En la terminal, instala las dependencias:

   ```powershell
   py -m pip install -r requirements.txt
   ```

4. Coloca el Excel del proyecto junto al archivo Python, o selecciona el archivo desde la barra lateral de la aplicación.
5. Inicia el sistema:

   ```powershell
   py -m streamlit run erp_pizza_pronto.py
   ```

## Publicar en Streamlit Community Cloud

1. Sube esta carpeta a un repositorio de GitHub.
2. En Streamlit Community Cloud, crea una aplicación desde ese repositorio y selecciona `erp_pizza_pronto.py`.
3. La aplicación publicada leerá automáticamente el Excel incluido en esta carpeta.

El repositorio debe permanecer privado porque el libro contiene datos de clientes y contactos. Los registros agregados desde la versión alojada no se guardan de forma permanente en el Excel del repositorio.

## Archivos

- `erp_pizza_pronto.py`: aplicación y módulos ERP.
- `requirements.txt`: dependencias de Python.
- `proyecto_gestion_avance_t2 (1) (2).xlsx`: libro fuente del sistema.

El libro contiene datos de clientes y contactos. Mantén privado el repositorio de GitHub. La aplicación podrá compartirse con su propio enlace, sin hacer público el repositorio.
