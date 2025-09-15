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
