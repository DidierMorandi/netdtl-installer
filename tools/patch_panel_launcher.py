import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import os
import sys

APP_TITLE = "NetDTL - Générateur tableau de brassage"


def resource_path(filename):
    base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, filename)


def check_python_dependencies():
    missing = []

    try:
        import pandas
    except ImportError:
        missing.append("pandas")

    try:
        import PIL
    except ImportError:
        missing.append("Pillow")

    try:
        import reportlab
    except ImportError:
        missing.append("reportlab")

    if missing:
        raise RuntimeError(
            "Modules Python manquants : " + ", ".join(missing)
        )


def browse_excel():
    path = filedialog.askopenfilename(
        title="Choisir un inventaire Excel",
        filetypes=[
            ("Excel", "*.xlsx *.xls"),
            ("Tous fichiers", "*.*")
        ]
    )

    if path:
        excel_var.set(path)


def browse_output():
    path = filedialog.asksaveasfilename(
        title="Enregistrer le PDF",
        defaultextension=".pdf",
        filetypes=[
            ("PDF", "*.pdf")
        ]
    )

    if path:
        output_var.set(path)


def generate():
    excel = excel_var.get().strip()
    output = output_var.get().strip()
    title = title_var.get().strip()

    if not excel:
        messagebox.showwarning(
            APP_TITLE,
            "Veuillez sélectionner un fichier Excel."
        )
        return

    if not os.path.exists(excel):
        messagebox.showerror(
            APP_TITLE,
            "Le fichier Excel est introuvable."
        )
        return

    if not output:
        messagebox.showwarning(
            APP_TITLE,
            "Veuillez choisir un fichier PDF de destination."
        )
        return

    if not title:
        title = "Tableau de brassage"

    script = resource_path("patch_panel_A3.py")

    if not os.path.exists(script):
        messagebox.showerror(
            APP_TITLE,
            "patch_panel_A3.py introuvable."
        )
        return

    cmd = [
        sys.executable,
        script,
        excel,
        "--output",
        output,
        "--title",
        title
    ]

    try:
        status_var.set("Génération en cours...")
        root.update()

        subprocess.run(
            cmd,
            check=True,
            capture_output=True,
            text=True
        )

        status_var.set("PDF généré avec succès.")

        os.startfile(output)

        messagebox.showinfo(
            APP_TITLE,
            "PDF généré avec succès."
        )

    except subprocess.CalledProcessError as e:
        status_var.set("Erreur.")

        stderr = e.stderr if e.stderr else str(e)

        messagebox.showerror(
            APP_TITLE,
            f"Erreur génération :\n\n{stderr}"
        )

    except Exception as e:
        status_var.set("Erreur.")

        messagebox.showerror(
            APP_TITLE,
            str(e)
        )


# UI
root = tk.Tk()
root.title(APP_TITLE)
root.geometry("720x320")
root.resizable(False, False)

excel_var = tk.StringVar()
output_var = tk.StringVar()
title_var = tk.StringVar(value="Tableau de brassage")
status_var = tk.StringVar(value="Prêt.")

frame = tk.Frame(root, padx=20, pady=20)
frame.pack(fill="both", expand=True)

tk.Label(
    frame,
    text="Inventaire Excel"
).grid(row=0, column=0, sticky="w")

tk.Entry(
    frame,
    textvariable=excel_var,
    width=60
).grid(row=1, column=0, padx=(0, 10))

tk.Button(
    frame,
    text="Parcourir",
    command=browse_excel
).grid(row=1, column=1)

tk.Label(
    frame,
    text="PDF de sortie"
).grid(row=2, column=0, sticky="w", pady=(20, 0))

tk.Entry(
    frame,
    textvariable=output_var,
    width=60
).grid(row=3, column=0, padx=(0, 10))

tk.Button(
    frame,
    text="Parcourir",
    command=browse_output
).grid(row=3, column=1)

tk.Label(
    frame,
    text="Titre"
).grid(row=4, column=0, sticky="w", pady=(20, 0))

tk.Entry(
    frame,
    textvariable=title_var,
    width=60
).grid(row=5, column=0, columnspan=2, sticky="we")

tk.Button(
    frame,
    text="Générer le PDF",
    command=generate,
    height=2
).grid(row=6, column=0, columnspan=2, pady=25)

tk.Label(
    frame,
    textvariable=status_var
).grid(row=7, column=0, columnspan=2)

check_python_dependencies()

root.mainloop()