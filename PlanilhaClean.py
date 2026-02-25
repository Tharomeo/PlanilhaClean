import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
import os

class LimpadorFinal:
    def __init__(self, root):
        self.root = root
        self.root.title("Planilha Clean")  # Título da janela alterado
        self.root.geometry("550x550")      # Mais quadrado e compacto

        # --- CORES (DARK MODERN) ---
        self.cor_fundo = "#1e1e1e"
        self.cor_painel = "#252526"
        self.cor_texto_pri = "#ffffff"
        self.cor_texto_sec = "#cccccc"
        self.cor_accent = "#007acc"
        self.cor_sucesso = "#28a745"
        self.cor_aviso = "#e2b93d"
        self.cor_card_lista = "#ffffff"
        
        self.root.configure(bg=self.cor_fundo)
        
        # Estilos
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure("TNotebook", background=self.cor_fundo, borderwidth=0)
        self.style.configure("TNotebook.Tab", background="#333333", foreground="white", padding=[10, 5], font=("Segoe UI", 9))
        self.style.map("TNotebook.Tab", background=[("selected", self.cor_accent)], foreground=[("selected", "white")])

        # Dados
        self.dados_abas = {} 
        self.caminho_atual = ""

        # --- CONTAINER PRINCIPAL ---
        self.container_main = tk.Frame(root, bg=self.cor_fundo)
        self.container_main.pack(fill="both", expand=True)

        self.setup_tela_upload()
        self.setup_tela_editor()

        self.mostrar_tela_upload()

    # =========================================================================
    # TELA 1: UPLOAD (Minimalista)
    # =========================================================================
    def setup_tela_upload(self):
        self.frame_upload = tk.Frame(self.container_main, bg=self.cor_fundo)
        
        # Centraliza tudo perfeitamente
        center_frame = tk.Frame(self.frame_upload, bg=self.cor_fundo)
        center_frame.place(relx=0.5, rely=0.5, anchor="center")

        # Área de Drop Clean
        self.area_drop = tk.Label(center_frame, 
                                  text="\n📂\n\nARRASTE O ARQUIVO AQUI\n\n(Ou clique para buscar)\n", 
                                  bg=self.cor_painel, fg=self.cor_texto_sec, 
                                  font=("Helvetica", 11), relief="ridge", bd=1, width=35)
        self.area_drop.pack(ipady=30) # ipady deixa o quadrado maior internamente
        
        self.area_drop.drop_target_register(DND_FILES)
        self.area_drop.dnd_bind('<<Drop>>', self.soltar_arquivo)
        self.area_drop.bind("<Button-1>", self.clicar_selecionar)

        self.lbl_status_upload = tk.Label(center_frame, text="", bg=self.cor_fundo, fg=self.cor_aviso, font=("Consolas", 9))
        self.lbl_status_upload.pack(pady=15)

    # =========================================================================
    # TELA 2: EDITOR (Abas e Opções)
    # =========================================================================
    def setup_tela_editor(self):
        self.frame_editor = tk.Frame(self.container_main, bg=self.cor_fundo)
        
        # Barra superior compacta
        top_bar = tk.Frame(self.frame_editor, bg="#2d2d30", height=40)
        top_bar.pack(fill="x")
        
        self.btn_voltar = tk.Button(top_bar, text="⬅ Voltar", command=self.mostrar_tela_upload,
                                    bg="#333333", fg="white", bd=0, padx=10, cursor="hand2", font=("Segoe UI", 9))
        self.btn_voltar.pack(side="left", padx=10, pady=5)

        self.lbl_arquivo_nome = tk.Label(top_bar, text="Arquivo.xlsx", bg="#2d2d30", fg="white", font=("Segoe UI", 10, "bold"))
        self.lbl_arquivo_nome.pack(side="left", padx=10)

        # Notebook (Abas) ocupa o espaço restante
        self.notebook = ttk.Notebook(self.frame_editor)
        self.notebook.pack(fill="both", expand=True, padx=15, pady=10)

        # Barra de Ações Inferior
        actions_bar = tk.Frame(self.frame_editor, bg=self.cor_fundo)
        actions_bar.pack(fill="x", padx=15, pady=(0, 15))

        # Botões de Seleção (Esquerda)
        frame_sel = tk.Frame(actions_bar, bg=self.cor_fundo)
        frame_sel.pack(side="left")
        
        tk.Button(frame_sel, text="Marcar Tudo", command=self.marcar_tudo_aba_atual, 
                  bg=self.cor_fundo, fg=self.cor_texto_sec, bd=0, cursor="hand2", font=("Segoe UI", 8)).pack(anchor="w")
        
        tk.Button(frame_sel, text="Desmarcar Tudo", command=self.desmarcar_tudo_aba_atual, 
                  bg=self.cor_fundo, fg=self.cor_texto_sec, bd=0, cursor="hand2", font=("Segoe UI", 8)).pack(anchor="w")

        # Botões de Ação (Direita) - Compactos
        self.btn_processar = tk.Button(actions_bar, text="💾 SALVAR", command=self.processar_real, 
                                       bg=self.cor_sucesso, fg="white", font=("Segoe UI", 10, "bold"), 
                                       relief="flat", padx=15, pady=5, cursor="hand2")
        self.btn_processar.pack(side="right", padx=5)

        self.btn_analisar = tk.Button(actions_bar, text="🔍 SIMULAR", command=self.analisar_simulacao, 
                                      bg=self.cor_accent, fg="white", font=("Segoe UI", 10, "bold"), 
                                      relief="flat", padx=15, pady=5, cursor="hand2")
        self.btn_analisar.pack(side="right", padx=5)

    # =========================================================================
    # NAVEGAÇÃO
    # =========================================================================
    def mostrar_tela_upload(self):
        self.frame_editor.pack_forget()
        self.frame_upload.pack(fill="both", expand=True)
        self.lbl_status_upload.config(text="")
        self.dados_abas = {}
        for tab in self.notebook.tabs(): self.notebook.forget(tab)

    def mostrar_tela_editor(self):
        self.frame_upload.pack_forget()
        self.frame_editor.pack(fill="both", expand=True)

    # =========================================================================
    # LÓGICA (MANTIDA)
    # =========================================================================
    def soltar_arquivo(self, event):
        path = event.data
        if path.startswith('{') and path.endswith('}'): path = path[1:-1]
        self.ler_arquivo(path)

    def clicar_selecionar(self, event):
        path = filedialog.askopenfilename(filetypes=[("Arquivos Excel/CSV", "*.xlsx *.xls *.csv")])
        if path: self.ler_arquivo(path)

    def ler_arquivo(self, caminho):
        try:
            self.lbl_status_upload.config(text="Lendo...", fg=self.cor_aviso)
            self.root.update()

            self.caminho_atual = caminho
            ext = caminho.lower()
            
            if ext.endswith('.csv'):
                df = self.ler_csv_robusto(caminho)
                if df is not None: self.criar_aba_interface("Dados CSV", df)
            else:
                xl = pd.ExcelFile(caminho)
                for sheet_name in xl.sheet_names:
                    df = pd.read_excel(caminho, sheet_name=sheet_name, dtype=str)
                    if df is not None and len(df.columns) > 0:
                        df.columns = df.columns.str.strip().str.replace('\n', '')
                        self.criar_aba_interface(sheet_name, df)

            if len(self.dados_abas) > 0:
                self.lbl_arquivo_nome.config(text=f"📄 {os.path.basename(caminho)}")
                self.mostrar_tela_editor()
            else:
                self.lbl_status_upload.config(text="Erro: Arquivo vazio.", fg="#ff5555")

        except Exception as e:
            self.lbl_status_upload.config(text="Erro fatal.", fg="#ff5555")
            messagebox.showerror("Erro", str(e))

    def ler_csv_robusto(self, caminho):
        try: return pd.read_csv(caminho, sep=';', engine='python', dtype=str)
        except: pass
        try: return pd.read_csv(caminho, sep=',', engine='python', dtype=str)
        except: pass
        return pd.read_csv(caminho, sep=None, engine='python', dtype=str)

    def criar_aba_interface(self, nome_aba, df):
        frame_tab = tk.Frame(self.notebook, bg="#f0f0f0")
        self.notebook.add(frame_tab, text=nome_aba)

        # Regra da Aba (Compacta)
        frame_config = tk.Frame(frame_tab, bg="#e0e0e0", pady=2, padx=5)
        frame_config.pack(fill="x")

        tk.Label(frame_config, text="Regra:", bg="#e0e0e0", font=("Segoe UI", 8, "bold")).pack(side="left")
        var_logica_aba = tk.StringVar(value="AND")
        tk.Radiobutton(frame_config, text="Rigoroso (E)", variable=var_logica_aba, value="AND", bg="#e0e0e0", font=("Segoe UI", 8)).pack(side="left", padx=5)
        tk.Radiobutton(frame_config, text="Flexível (OU)", variable=var_logica_aba, value="OR", bg="#e0e0e0", font=("Segoe UI", 8)).pack(side="left", padx=5)

        # Lista de Colunas
        canvas = tk.Canvas(frame_tab, bg="white", highlightthickness=0)
        scrollbar = ttk.Scrollbar(frame_tab, orient="vertical", command=canvas.yview)
        frame_interno = tk.Frame(canvas, bg="white")

        frame_interno.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=frame_interno, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        scrollbar.pack(side="right", fill="y")

        lista_vars = []
        for col in df.columns:
            var = tk.BooleanVar()
            if any(x in col.lower() for x in ['email', 'cpf', 'cnpj', 'tel', 'cel', 'membro']): var.set(True)
            chk = tk.Checkbutton(frame_interno, text=col, variable=var, bg="white", anchor="w", font=("Segoe UI", 9))
            chk.pack(fill="x", padx=2, pady=0)
            lista_vars.append((col, var))

        self.dados_abas[nome_aba] = {"df": df, "vars": lista_vars, "logic_var": var_logica_aba}

    # =========================================================================
    # PROCESSAMENTO
    # =========================================================================
    def get_aba_ativa(self):
        try: return self.notebook.tab(self.notebook.select(), "text")
        except: return None

    def marcar_tudo_aba_atual(self):
        nome = self.get_aba_ativa()
        if nome:
            for _, var in self.dados_abas[nome]["vars"]: var.set(True)

    def desmarcar_tudo_aba_atual(self):
        nome = self.get_aba_ativa()
        if nome:
            for _, var in self.dados_abas[nome]["vars"]: var.set(False)

    def limpar_dataframe(self, df_entrada, colunas_sel, modo_logica):
        if not colunas_sel: return df_entrada, 0
        df_proc = df_entrada.copy()
        cols_check = []
        for col in colunas_sel:
            nome_temp = f"__temp_{col}"
            df_proc[nome_temp] = df_proc[col].astype(str).str.lower().str.strip()
            cols_check.append(nome_temp)
            
        if modo_logica == "AND": mask = df_proc.duplicated(subset=cols_check, keep='first')
        else:
            mask = pd.Series([False] * len(df_proc), index=df_proc.index)
            for col in cols_check: mask = mask | df_proc.duplicated(subset=[col], keep='first')
        
        df_final = df_proc[~mask].copy()
        df_final = df_final.drop(columns=cols_check)
        removidas = len(df_entrada) - len(df_final)
        return df_final, removidas

    def analisar_simulacao(self):
        nome_aba = self.get_aba_ativa()
        if not nome_aba: return
        dados = self.dados_abas[nome_aba]
        colunas_sel = [col for col, var in dados["vars"] if var.get()]
        
        if not colunas_sel: return messagebox.showinfo("Aviso", "Nenhuma coluna selecionada nesta aba.")

        try:
            total = len(dados["df"])
            _, removidas = self.limpar_dataframe(dados["df"], colunas_sel, dados["logic_var"].get())
            restantes = total - removidas
            
            msg = f"📊 SIMULAÇÃO: {nome_aba}\n--------------------------\n"
            msg += f"Total antes: {total}\nRemovidas: {removidas}\nRestantes: {restantes}"
            messagebox.showinfo("Simulação", msg)
        except Exception as e: messagebox.showerror("Erro", str(e))

    def processar_real(self):
        if not self.dados_abas: return
        save_path = filedialog.asksaveasfilename(initialfile="Arquivo_Limpo.xlsx", defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not save_path: return

        try:
            total_removido = 0
            abas_alteradas = 0
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                for nome_aba, dados in self.dados_abas.items():
                    colunas_sel = [col for col, var in dados["vars"] if var.get()]
                    if colunas_sel:
                        df_limpo, qtd = self.limpar_dataframe(dados["df"], colunas_sel, dados["logic_var"].get())
                        df_limpo.to_excel(writer, sheet_name=nome_aba, index=False)
                        total_removido += qtd
                        abas_alteradas += 1
                    else:
                        dados["df"].to_excel(writer, sheet_name=nome_aba, index=False)

            messagebox.showinfo("Sucesso", f"Concluído!\nLinhas removidas: {total_removido}")
        except Exception as e: messagebox.showerror("Erro", str(e))

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = LimpadorFinal(root)
    root.mainloop()