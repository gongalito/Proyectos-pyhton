# Payment and Invoice Management System

Project developed for a small business to **manage customer payments and invoices** using predefined Excel formats.  

It includes a **user-friendly graphical interface** with the following features:

- Create a customer account statement  
- Add payments and account statements  
- View invoices:  
  - **Red:** unpaid  
  - **Green:** paid  
  - **Orange:** price discrepancy  
- Filter invoices that are unpaid or have discrepancies  
- Mark specific invoices as paid  

*More information is available in the program’s **Help** section.*

---

## Folder Structure

The project requires the following folders to function correctly:

- `Agregar Control/`  # Folder to add new account statements  
- `Agregar Pago/`   # Folder to record payments  
- `CONTROL/`      # Main folder where control files are stored  
- `Data/`  
  - `Historial Control/`   # Stores control history  
  - `Historial Pagos/`    # Stores payment history  
  - `FacturasNoPagadas/`   # Stores unpaid invoices  
  - `Copias de seguridad/`  # Backup folder  
  - `Textos de ayuda/`     # Documentation and help files (see the program’s **Help** section for details)

---

## Key Technologies and Libraries

- **Python 3**  
- **Pandas / Openpyxl** – Excel processing and formatting  
- **CustomTkinter / Tkinter (ScrolledText)** – Graphical interface  
- **OS, Shutil, Glob, Datetime, Calendar, Locale** – File and date management  

---

## Usage

```bash
pip install -r requirements.txt
python main.py

