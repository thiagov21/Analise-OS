import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import ttkthemes
from service_order_scraper import ServiceOrderScraper
import threading
from datetime import datetime
import json
import os
from PIL import Image, ImageTk

class SmartAnalyzer:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Smart Analyzer")
        self.root.geometry("1200x800")
        
        # Aplicação do tema
        self.style = ttkthemes.ThemedStyle(self.root)
        self.style.set_theme("arc")
        
        self.scraper = ServiceOrderScraper()
        self.analyzed_results = []
        self.setup_gui()
        
    def setup_gui(self):
        # Configuração do grid
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        
        # container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_frame.grid_columnconfigure(0, weight=1)
        
        # Logo
        header_frame = ttk.Frame(main_frame)
        header_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Titulo 
        title_label = ttk.Label(header_frame, text="Analise OS", font=("Helvetica", 20, "bold"))
        title_label.grid(row=0, column=0, pady=10)
        
        # Login
        login_frame = ttk.LabelFrame(main_frame, text="Login", padding="10")
        login_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        
        # usuario
        username_frame = ttk.Frame(login_frame)
        username_frame.grid(row=0, column=0, padx=5)
        ttk.Label(username_frame, text="Usuario:").grid(row=0, column=0, padx=5)
        self.username_var = tk.StringVar(value="")
        ttk.Entry(username_frame, textvariable=self.username_var).grid(row=0, column=1, padx=5)
        
        # senha
        password_frame = ttk.Frame(login_frame)
        password_frame.grid(row=0, column=1, padx=5)
        ttk.Label(password_frame, text="Senha:").grid(row=0, column=0, padx=5)
        self.password_var = tk.StringVar(value="")
        ttk.Entry(password_frame, textvariable=self.password_var, show="*").grid(row=0, column=1, padx=5)
        
        # Botão login
        self.login_btn = ttk.Button(login_frame, text="Login", command=self.login, style="Accent.TButton")
        self.login_btn.grid(row=0, column=2, padx=20)
        
        # frame de ação
        actions_frame = ttk.LabelFrame(main_frame, text="Ações", padding="10")
        actions_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=10)
        
        # busca de OS
        search_frame = ttk.Frame(actions_frame)
        search_frame.grid(row=0, column=0, padx=10)
        ttk.Label(search_frame, text="Número OS:").grid(row=0, column=0, padx=5)
        self.os_number_var = tk.StringVar()
        ttk.Entry(search_frame, textvariable=self.os_number_var).grid(row=0, column=1, padx=5)
        self.search_btn = ttk.Button(search_frame, text="Buscar OS", command=self.search_os)
        self.search_btn.grid(row=0, column=2, padx=5)
        
        # Batch ação
        batch_frame = ttk.Frame(actions_frame)
        batch_frame.grid(row=0, column=1, padx=10)
        ttk.Button(batch_frame, text="Analisar Planilha", command=self.analyze_excel).grid(row=0, column=0, padx=5)
        ttk.Button(batch_frame, text="Exportar Resultados", command=self.export_results).grid(row=0, column=1, padx=5)
        
        # barra de progresso
        progress_frame = ttk.Frame(main_frame)
        progress_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=5)
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=5)
        self.progress_label = ttk.Label(progress_frame, text="")
        self.progress_label.grid(row=1, column=0, sticky=(tk.W, tk.E), padx=5)
        
        # barra de resultados
        results_frame = ttk.LabelFrame(main_frame, text="Resultados", padding="10")
        results_frame.grid(row=4, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        results_frame.grid_rowconfigure(1, weight=1)
        results_frame.grid_columnconfigure(0, weight=1)
        
        # tabela
        self.results_tree = ttk.Treeview(results_frame, columns=(
            "OS", "Cliente", "Referência", "Garantia", "Status", "Data Compra"
        ), show="headings")
        
        # configuração de colunas
        column_widths = {
            "OS": 100,
            "Cliente": 200,
            "Referência": 150,
            "Garantia": 100,
            "Status": 150,
            "Data Compra": 150
        }
        
        for col, width in column_widths.items():
            self.results_tree.heading(col, text=col)
            self.results_tree.column(col, width=width)
        
        # scrollbars
        y_scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.results_tree.yview)
        x_scrollbar = ttk.Scrollbar(results_frame, orient=tk.HORIZONTAL, command=self.results_tree.xview)
        self.results_tree.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)
        
        # grid do resultado
        self.results_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        y_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        x_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # Status bar
        self.status_var = tk.StringVar()
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=5, column=0, sticky=(tk.W, tk.E))
        
        # estilo config
        self.style.configure("Accent.TButton", background="#007bff")
        
    def update_progress(self, value, message=""):
        self.progress_var.set(value)
        self.progress_label.config(text=message)
        self.root.update_idletasks()
        
    def login(self):
        def login_thread():
            try:
                self.status_var.set("Fazendo login...")
                self.login_btn.state(['disabled'])
                
                success = self.scraper.login(
                    self.username_var.get(),
                    self.password_var.get()
                )
                
                if success:
                    self.status_var.set("Login realizado com sucesso!")
                    messagebox.showinfo("Sucesso", "Login realizado com sucesso!")
                else:
                    self.status_var.set("Falha no login!")
                    messagebox.showerror("Erro", "Falha no login!")
            finally:
                self.login_btn.state(['!disabled'])
        
        thread = threading.Thread(target=login_thread)
        thread.daemon = True
        thread.start()
    
    def search_os(self):
        def search_thread():
            try:
                self.status_var.set("Buscando OS...")
                self.search_btn.state(['disabled'])
                self.update_progress(0, "Iniciando busca...")
                
                # limpar resultados anteriores
                for item in self.results_tree.get_children():
                    self.results_tree.delete(item)
                
                result = self.scraper.analyze_single_os(self.os_number_var.get())
                self.analyzed_results = [result]
                
                # atualizar tabela de resultados
                self.results_tree.insert("", "end", values=(
                    result["numero_os"],
                    result["cliente"],
                    result["referencia"],
                    result["garantia"],
                    result["status"],
                    result["dataCompra"]
                ))
                
                self.update_progress(100, "Busca concluída!")
                self.status_var.set("Busca concluída!")
                
                if result["image_valid"]:
                    messagebox.showinfo("Sucesso", f"Comprovante salvo em: {result['image_path']}")
                
            except Exception as e:
                self.status_var.set(f"Erro: {str(e)}")
                messagebox.showerror("Erro", str(e))
            finally:
                self.search_btn.state(['!disabled'])
                self.update_progress(0, "")
        
        thread = threading.Thread(target=search_thread)
        thread.daemon = True
        thread.start()
    
    def analyze_excel(self):
        def analyze_thread():
            try:
                self.status_var.set("Analisando planilha...")
                self.update_progress(0, "Iniciando análise...")
                
                # limpar resultados anteriores
                for item in self.results_tree.get_children():
                    self.results_tree.delete(item)
                
                results = self.scraper.analyze_excel_file(
                    progress_callback=self.update_progress
                )
                
                self.analyzed_results = results
                
                # atualizar tabela de resultados
                for result in results:
                    if "error" not in result:
                        self.results_tree.insert("", "end", values=(
                            result["numero_os"],
                            result["cliente"],
                            result["referencia"],
                            result["garantia"],
                            result["status"],
                            result["dataCompra"]
                        ))
                
                self.update_progress(100, "Análise concluída!")
                self.status_var.set("Análise concluída!")
                messagebox.showinfo("Sucesso", f"Análise concluída! {len(results)} OS processadas.")
                
            except Exception as e:
                self.status_var.set(f"Erro: {str(e)}")
                messagebox.showerror("Erro", str(e))
            finally:
                self.update_progress(0, "")
        
        thread = threading.Thread(target=analyze_thread)
        thread.daemon = True
        thread.start()
    
    def export_results(self):
        if not self.analyzed_results:
            messagebox.showwarning("Aviso", "Não há resultados para exportar!")
            return
            
        try:
            output_file = self.scraper.export_results(self.analyzed_results, format="excel")
            self.status_var.set(f"Resultados exportados para: {output_file}")
            messagebox.showinfo("Sucesso", f"Resultados exportados para:\n{output_file}")
        except Exception as e:
            self.status_var.set(f"Erro ao exportar: {str(e)}")
            messagebox.showerror("Erro", str(e))
    
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = SmartAnalyzer()
    app.run()