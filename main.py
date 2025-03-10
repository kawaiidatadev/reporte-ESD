import sys
import tkinter as tk
from tkinter import scrolledtext, messagebox
from common.__init__ import *
from obtener_data import generar_data
from archivo.copiar_actualizar import main

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Ejecución del Programa")

        # Área de texto para la salida de la consola
        self.text_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=80, height=20)
        self.text_area.pack(padx=10, pady=10, expand=True, fill="both")

        # Botón de cancelar
        self.cancel_button = tk.Button(root, text="Cancelar", command=self.cancel_execution, bg="red", fg="white")
        self.cancel_button.pack(pady=10)

        # Redirigir stdout y stderr a la interfaz
        sys.stdout = self
        sys.stderr = self

        # Variable de control
        self.running = True

        # Iniciar ejecución en pasos
        self.root.after(100, self.run_step1)

    def write(self, message):
        """Escribe mensajes en la caja de texto y actualiza la interfaz."""
        self.text_area.insert(tk.END, message + "\n")
        self.text_area.yview(tk.END)  # Auto-scroll al final
        self.root.update_idletasks()  # Refresca la UI

    def flush(self):
        """Flush necesario para que los prints funcionen bien."""
        pass

    def run_step1(self):
        """Primer paso: Generar datos."""
        if not self.running:
            self.write("Ejecución cancelada antes de comenzar.")
            return

        self.write("Iniciando copia de seguridad...")
        self.root.after(100, self.execute_generar_data)

    def execute_generar_data(self):
        """Ejecuta la copia de seguridad y pasa al siguiente paso."""
        try:
            generar_data()
            self.write("Copia de seguridad completada.")
            self.root.after(100, self.run_step2)  # Pasa al siguiente paso
        except Exception as e:
            self.write(f"Error en copia de seguridad: {e}")
            messagebox.showerror("Error", f"Se ha producido un error: {e}")
            self.root.quit()

    def run_step2(self):
        """Segundo paso: Copiar y actualizar archivos."""
        if not self.running:
            self.write("Ejecución cancelada antes de actualizar archivos.")
            return

        self.write("Iniciando copia y actualización del archivo...")
        self.root.after(100, self.execute_main)

    def execute_main(self):
        """Ejecuta la actualización y finaliza el proceso."""
        try:
            main()
            self.write("Proceso finalizado exitosamente.")
            messagebox.showinfo("Completado", "El proceso ha finalizado correctamente.")
            self.root.quit()
        except Exception as e:
            self.write(f"Error en la actualización: {e}")
            messagebox.showerror("Error", f"Se ha producido un error: {e}")
            self.root.quit()

    def cancel_execution(self):
        """Cancela la ejecución del script."""
        self.running = False
        self.write("Cancelando el proceso...")
        messagebox.showwarning("Cancelado", "El proceso ha sido cancelado.")
        self.root.quit()

# Iniciar la aplicación gráfica
if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    root.mainloop()
