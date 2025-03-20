import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from report_generator import PDF, process_file

class AsistenciaApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Generador de Informes de Asistencia")
        self.root.geometry("600x400")
        
        self.create_widgets()
        self.configure_styles()

    def create_widgets(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(expand=True, fill=tk.BOTH)

        # Título
        title_label = ttk.Label(main_frame, 
                               text="Generador de Informes PDF",
                               font=('Helvetica', 14, 'bold'))
        title_label.pack(pady=10)

        # Botón para seleccionar archivos
        self.btn_select = ttk.Button(main_frame,
                                    text="Seleccionar archivos Excel",
                                    command=self.select_files)
        self.btn_select.pack(pady=20)

        # Lista de archivos seleccionados
        self.file_list = tk.Listbox(main_frame, 
                                   height=8,
                                   selectmode=tk.EXTENDED)
        self.file_list.pack(expand=True, fill=tk.BOTH, pady=10)

        # Barra de progreso
        self.progress = ttk.Progressbar(main_frame,
                                       orient=tk.HORIZONTAL,
                                       mode='determinate')
        self.progress.pack(fill=tk.X, pady=10)

        # Botón de generar informes
        self.btn_generate = ttk.Button(main_frame,
                                      text="Generar Informes PDF",
                                      command=self.generate_reports,
                                      state=tk.DISABLED)
        self.btn_generate.pack(pady=10)

    def configure_styles(self):
        style = ttk.Style()
        style.configure('TButton', font=('Helvetica', 10))
        style.configure('TLabel', font=('Helvetica', 11))
        style.map('TButton', foreground=[('active', '!disabled', 'blue')])

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="Seleccionar archivos Excel",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if files:
            self.file_list.delete(0, tk.END)
            for f in files:
                self.file_list.insert(tk.END, f)
            self.btn_generate['state'] = tk.NORMAL

    def generate_reports(self):
        files = self.file_list.get(0, tk.END)
        if not files:
            messagebox.showwarning("Advertencia", "No se seleccionaron archivos")
            return

        self.progress['maximum'] = len(files)
        self.progress['value'] = 0
        self.btn_generate['state'] = tk.DISABLED
        self.btn_select['state'] = tk.DISABLED

        success_count = 0
        for i, file_path in enumerate(files, 1):
            try:
                process_file(file_path)
                success_count += 1
            except Exception as e:
                messagebox.showerror("Error", 
                    f"Error procesando {os.path.basename(file_path)}:\n{str(e)}")
            finally:
                self.progress['value'] = i
                self.root.update_idletasks()

        messagebox.showinfo("Completado",
            f"Proceso finalizado\nArchivos procesados correctamente: {success_count}/{len(files)}")
        
        self.btn_generate['state'] = tk.NORMAL
        self.btn_select['state'] = tk.NORMAL
        self.progress['value'] = 0

if __name__ == "__main__":
    root = tk.Tk()
    app = AsistenciaApp(root)
    root.mainloop()