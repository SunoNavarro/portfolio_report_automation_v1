"""
Interfaz mínima para generar el informe Excel desde un CSV.

Usa solo tkinter (biblioteca estándar de Python): sin dependencias extra de UI.

Ejecutar desde la raíz del proyecto:
    python src/visual_script.py
"""

from __future__ import annotations

import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

from generate_report import generate_report


def main() -> None:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    path = filedialog.askopenfilename(
        parent=root,
        title="Seleccionar CSV de ventas",
        filetypes=[("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")],
    )
    root.attributes("-topmost", False)

    if not path:
        root.destroy()
        return

    input_path = Path(path)
    output_path = input_path.parent / f"{input_path.stem}_report.xlsx"

    try:
        generate_report(input_path, output_path)
    except PermissionError:
        messagebox.showerror(
            "Error de permisos",
            "No se pudo guardar el informe. Cierra el Excel si tienes abierto el fichero "
            "de salida o guarda con otro nombre.",
            parent=root,
        )
        root.destroy()
        sys.exit(1)
    except Exception as exc:  # noqa: BLE001 — mensaje al usuario final
        messagebox.showerror(
            "Error",
            f"No se pudo generar el informe:\n{exc}",
            parent=root,
        )
        root.destroy()
        sys.exit(1)

    messagebox.showinfo(
        "Informe generado",
        "El fichero se ha convertido y guardado en el mismo directorio que el CSV de origen:\n\n"
        f"{output_path}",
        parent=root,
    )
    root.destroy()


if __name__ == "__main__":
    main()
