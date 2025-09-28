#Interface Gráfica para o Gerador de Documentos
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
from gerador_rpontes import GeradorRPontes

class InterfaceGerador:
    """Interface gráfica principal do sistema"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.configurar_janela()
        self.criar_widgets()
        self.planilha_selecionada = None
        self.template_selecionado = None
        self.pasta_saida_selecionada = None
        
    def configurar_janela(self):
        """Configura a janela principal"""
        self.root.title("Gerador de Documentos - R Pontes Construtora")
        self.root.geometry("800x600")
        self.root.configure(bg='white')
        self.root.resizable(True, True)
        
        # Centraliza a janela
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (800 // 2)
        y = (self.root.winfo_screenheight() // 2) - (600 // 2)
        self.root.geometry(f"800x600+{x}+{y}")
    
    def criar_widgets(self):
        """Cria todos os elementos da interface"""
        # Título principal
        titulo = tk.Label(
            self.root,
            text="GERADOR DE DOCUMENTOS",
            font=("Arial", 20, "bold"),
            bg='white',
            fg='black'
        )
        titulo.pack(pady=20)
        
        subtitulo = tk.Label(
            self.root,
            text="R Pontes Construtora",
            font=("Arial", 12),
            bg='white',
            fg='gray'
        )
        subtitulo.pack(pady=(0, 30))
        
        # Frame principal
        main_frame = tk.Frame(self.root, bg='white')
        main_frame.pack(expand=True, fill='both', padx=40, pady=20)
        
        # Seção de arquivos
        self.criar_secao_arquivos(main_frame)
        
        # Seção de controles
        self.criar_secao_controles(main_frame)
        
        # Seção de logs
        self.criar_secao_logs(main_frame)
        
        # Barra de status
        self.criar_barra_status()
    
    def criar_secao_arquivos(self, parent):
        """Cria seção de seleção de arquivos"""
        frame = tk.LabelFrame(
            parent,
            text="Arquivos",
            font=("Arial", 12, "bold"),
            bg='white',
            fg='black',
            bd=1,
            relief='solid'
        )
        frame.pack(fill='x', pady=(0, 20))
        
        # Planilha
        tk.Label(frame, text="Planilha Excel:", bg='white', fg='black').grid(row=0, column=0, sticky='w', padx=10, pady=10)
        
        self.label_planilha = tk.Label(
            frame,
            text="Nenhuma planilha selecionada",
            bg='white',
            fg='gray',
            width=50,
            anchor='w'
        )
        self.label_planilha.grid(row=0, column=1, padx=10, pady=10)
        
        btn_planilha = tk.Button(
            frame,
            text="Selecionar",
            command=self.selecionar_planilha,
            bg='black',
            fg='white',
            font=("Arial", 10),
            relief='flat',
            padx=20
        )
        btn_planilha.grid(row=0, column=2, padx=10, pady=10)
        
        # Template
        tk.Label(frame, text="Template Word:", bg='white', fg='black').grid(row=1, column=0, sticky='w', padx=10, pady=10)
        
        self.label_template = tk.Label(
            frame,
            text="Nenhum template selecionado",
            bg='white',
            fg='gray',
            width=50,
            anchor='w'
        )
        self.label_template.grid(row=1, column=1, padx=10, pady=10)
        
        btn_template = tk.Button(
            frame,
            text="Selecionar",
            command=self.selecionar_template,
            bg='black',
            fg='white',
            font=("Arial", 10),
            relief='flat',
            padx=20
        )
        btn_template.grid(row=1, column=2, padx=10, pady=10)
        
        # Pasta de saída
        tk.Label(frame, text="Pasta de Saída:", bg='white', fg='black').grid(row=2, column=0, sticky='w', padx=10, pady=10)
        
        self.label_pasta = tk.Label(
            frame,
            text="Pasta padrão será usada",
            bg='white',
            fg='gray',
            width=50,
            anchor='w'
        )
        self.label_pasta.grid(row=2, column=1, padx=10, pady=10)
        
        btn_pasta = tk.Button(
            frame,
            text="Selecionar",
            command=self.selecionar_pasta,
            bg='black',
            fg='white',
            font=("Arial", 10),
            relief='flat',
            padx=20
        )
        btn_pasta.grid(row=2, column=2, padx=10, pady=10)
    
    def criar_secao_controles(self, parent):
        """Cria seção de controles"""
        frame = tk.LabelFrame(
            parent,
            text="Controles",
            font=("Arial", 12, "bold"),
            bg='white',
            fg='black',
            bd=1,
            relief='solid'
        )
        frame.pack(fill='x', pady=(0, 20))
        
        # Botão principal
        self.btn_gerar = tk.Button(
            frame,
            text="GERAR DOCUMENTOS",
            command=self.iniciar_geracao,
            bg='black',
            fg='white',
            font=("Arial", 14, "bold"),
            relief='flat',
            padx=40,
            pady=15
        )
        self.btn_gerar.pack(pady=20)
        
        # Barra de progresso
        self.progress = ttk.Progressbar(
            frame,
            mode='indeterminate',
            style='Black.Horizontal.TProgressbar'
        )
        self.progress.pack(fill='x', padx=20, pady=(0, 20))
        
        # Configurar estilo da barra de progresso
        style = ttk.Style()
        style.configure('Black.Horizontal.TProgressbar', background='black')
    
    def criar_secao_logs(self, parent):
        """Cria seção de logs"""
        frame = tk.LabelFrame(
            parent,
            text="Log de Execução",
            font=("Arial", 12, "bold"),
            bg='white',
            fg='black',
            bd=1,
            relief='solid'
        )
        frame.pack(fill='both', expand=True)
        
        self.text_log = scrolledtext.ScrolledText(
            frame,
            height=10,
            bg='white',
            fg='black',
            font=("Consolas", 10),
            bd=0,
            relief='flat'
        )
        self.text_log.pack(fill='both', expand=True, padx=10, pady=10)
    
    def criar_barra_status(self):
        """Cria barra de status"""
        self.status_bar = tk.Label(
            self.root,
            text="Pronto",
            bg='black',
            fg='white',
            anchor='w',
            padx=10,
            pady=5
        )
        self.status_bar.pack(side='bottom', fill='x')
    
    def selecionar_planilha(self):
        """Seleciona arquivo Excel"""
        arquivo = filedialog.askopenfilename(
            title="Selecionar Planilha Excel",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if arquivo:
            self.planilha_selecionada = arquivo
            nome = os.path.basename(arquivo)
            self.label_planilha.config(text=nome, fg='black')
            self.log(f"Planilha selecionada: {nome}")
    
    def selecionar_template(self):
        """Seleciona template Word"""
        arquivo = filedialog.askopenfilename(
            title="Selecionar Template Word",
            filetypes=[("Word files", "*.docx")]
        )
        if arquivo:
            self.template_selecionado = arquivo
            nome = os.path.basename(arquivo)
            self.label_template.config(text=nome, fg='black')
            self.log(f"Template selecionado: {nome}")
    
    def selecionar_pasta(self):
        """Seleciona pasta de saída"""
        pasta = filedialog.askdirectory(
            title="Selecionar Pasta para Salvar Documentos"
        )
        if pasta:
            self.pasta_saida_selecionada = pasta
            nome = os.path.basename(pasta) or pasta
            self.label_pasta.config(text=nome, fg='black')
            self.log(f"Pasta de saída: {pasta}")
    
    def log(self, mensagem):
        """Adiciona mensagem ao log"""
        self.text_log.insert(tk.END, f"{mensagem}\n")
        self.text_log.see(tk.END)
        self.root.update_idletasks()
    
    def atualizar_status(self, mensagem):
        """Atualiza barra de status"""
        self.status_bar.config(text=mensagem)
    
    def iniciar_geracao(self):
        """Inicia processo de geração em thread separada"""
        if not self.planilha_selecionada:
            messagebox.showerror("Erro", "Selecione uma planilha Excel")
            return
        
        if not self.template_selecionado:
            messagebox.showerror("Erro", "Selecione um template Word")
            return
        
        # Desabilita botão e inicia progresso
        self.btn_gerar.config(state='disabled')
        self.progress.start()
        self.text_log.delete(1.0, tk.END)
        self.atualizar_status("Processando...")
        
        # Executa em thread separada
        thread = threading.Thread(target=self.executar_geracao)
        thread.daemon = True
        thread.start()
    
    def executar_geracao(self):
        """Executa geração de documentos"""
        try:
            self.log("Iniciando geração de documentos...")
            
            gerador = GeradorRPontes(
                self.planilha_selecionada, 
                self.template_selecionado,
                self.pasta_saida_selecionada
            )
            
            # Redireciona logs para interface
            import sys
            from io import StringIO
            
            old_stdout = sys.stdout
            sys.stdout = StringIO()
            
            documentos = gerador.processar_clientes()
            
            # Restaura stdout
            output = sys.stdout.getvalue()
            sys.stdout = old_stdout
            
            # Mostra resultado na interface
            for linha in output.split('\n'):
                if linha.strip():
                    self.log(linha)
            
            self.log(f"\nConcluído! {len(documentos)} documentos gerados.")
            
            # Mostra mensagem de sucesso
            self.root.after(0, lambda: messagebox.showinfo(
                "Sucesso", 
                f"Processo concluído!\n{len(documentos)} documentos gerados."
            ))
            
        except Exception as e:
            self.log(f"ERRO: {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("Erro", f"Erro durante execução:\n{str(e)}"))
        
        finally:
            # Reabilita interface
            self.root.after(0, self.finalizar_geracao)
    
    def finalizar_geracao(self):
        """Finaliza processo e reabilita interface"""
        self.progress.stop()
        self.btn_gerar.config(state='normal')
        self.atualizar_status("Pronto")
    
    def executar(self):
        """Inicia a interface gráfica"""
        self.root.mainloop()

if __name__ == "__main__":
    app = InterfaceGerador()
    app.executar()