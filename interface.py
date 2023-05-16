import tkinter as tk
import subprocess
import tkinter.messagebox as mb


def run():
    
    try:
        subprocess.call(['python', 'main2.py'])

        
        # Create a button widget that runs the app when clicked
    
    except Exception as e:
        if str(e) == "There is no excel file in the current directory!":
            mb.showerror("Error", "No hay ningún archivo Excel en el directorio actual.")


root = tk.Tk()
root.geometry("100x100")
button = tk.Button(root, text="Enviar correo", command=run)
button.pack()

# Inicia el loop principal
root.mainloop()