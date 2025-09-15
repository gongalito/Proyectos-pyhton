# Sistema de Gestión de Pagos y Facturas

Proyecto desarrollado para una pequeña empresa que permite **controlar pagos y facturas de clientes específicos** con formatos predefinidos de Excel.  

Cuenta con una **interfaz gráfica amigable** y las siguientes funcionalidades:

- Crear estado de cuenta de un cliente  
- Agregar pagos y estados de cuenta  
- Visualización de facturas:  
  - **Rojo:** no pagadas  
  - **Verde:** pagadas  
  - **Naranja:** diferencia de precio  
- Filtrar facturas no pagadas o con diferencias  
- Marcar facturas específicas como pagadas

---

## Estructura de carpetas

El proyecto requiere las siguientes carpetas para funcionar correctamente:
- Agregar Control/           # Carpeta para agregar nuevos estados de cuenta
- Agregar Pago/              # Carpeta para registrar pagos
- CONTROL/                   # Carpeta principal donde se guardan los archivos de control
- Data/
    - Historial Control/         # Carpeta donde se almacenan los históricos de control
    - Historial Pagos/           # Carpeta donde se almacenan los históricos de pagos
    - FacturasNoPagadas/         # Carpeta para facturas no pagadas
    - Copias de seguridad/       # Carpeta para respaldos de archivos
    - Textos de ayuda/           # Carpeta con documentación y archivos de ayuda

---

## Tecnologías y librerías principales

- **Python 3**  
- **Pandas / Openpyxl** – manejo y formateo de Excel  
- **CustomTkinter / Tkinter (ScrolledText)** – interfaz gráfica  
- **OS, Shutil, Glob, Datetime, Calendar, Locale** – gestión de archivos y fechas  

---


## Uso

```bash
pip install -r requirements.txt
python main.py
