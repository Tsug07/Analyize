import os
import re
import sys
import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from PIL import Image, ImageTk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import threading
from datetime import datetime
import json

class AnalisadorParcelamentos:
    def __init__(self):
        self.dados_processados = []
        self.empresas_filtradas = set()
        self.empresas_codigos = {}  # Dicionário para mapear CNPJ -> Código
        self.dados_filtrados_atual = []  # Para manter os dados atualmente filtrados
        self.colunas_ordem = {}  # Para rastrear direção de ordenação de cada coluna
        self._assets_dir = self._obter_pasta_assets()
        self.setup_gui()

    def _obter_pasta_assets(self):
        """Retorna o caminho da pasta assets, compatível com PyInstaller e execução direta"""
        if getattr(sys, 'frozen', False):
            base = sys._MEIPASS
        else:
            base = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base, "assets")
        
    def setup_gui(self):
        # Criar janela principal com ttkbootstrap
        self.janela = ttk.Window(themename="superhero")
        self.janela.title("Analyize - Analisador de Parcelamentos v2.0")
        self.janela.geometry("1500x900")
        self.janela.resizable(True, True)

        # Definir ícone da janela (favicon)
        favicon_path = os.path.join(self._assets_dir, "favicon.ico")
        if os.path.exists(favicon_path):
            self.janela.iconbitmap(favicon_path)

        # Notebook para abas
        self.notebook = ttk.Notebook(self.janela)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self.criar_aba_configuracao()
        self.criar_aba_resultados()
        self.criar_aba_dashboard()
        
    def criar_aba_configuracao(self):
        # Aba 1: Configuração
        frame_config = ttk.Frame(self.notebook)
        self.notebook.add(frame_config, text="  Configuração")

        # Header com logo
        header_frame = ttk.Frame(frame_config)
        header_frame.pack(pady=(10, 20))

        logo_path = os.path.join(self._assets_dir, "Analayize_logo.png")
        if os.path.exists(logo_path):
            img = Image.open(logo_path).resize((48, 48), Image.LANCZOS)
            self._logo_img = ImageTk.PhotoImage(img)
            ttk.Label(header_frame, image=self._logo_img).pack(side="left", padx=(0, 10))

        title_label = ttk.Label(header_frame, text="Analyize - Analisador de Parcelamentos v2.0",
                               font=("Arial", 16, "bold"), bootstyle="success")
        title_label.pack(side="left")

        # Frame para inputs
        inputs_frame = ttk.Frame(frame_config)
        inputs_frame.pack(fill="x", padx=20, pady=10)

        # Pasta dos PDFs
        ttk.Label(inputs_frame, text="Pasta com os PDFs:", font=("Arial", 10, "bold")).pack(anchor="w", pady=(5, 2))
        frame_pdfs = ttk.Frame(inputs_frame)
        frame_pdfs.pack(fill="x", pady=(0, 10))
        self.entrada_pasta_pdfs = ttk.Entry(frame_pdfs)
        self.entrada_pasta_pdfs.pack(side="left", fill="x", expand=True, padx=(0, 5))
        ttk.Button(frame_pdfs, text="Selecionar", command=self.selecionar_pasta_pdfs,
                   bootstyle="info-outline").pack(side="right")

        # Excel de empresas (opcional)
        ttk.Label(inputs_frame, text="Excel com empresas filtradas (opcional):", font=("Arial", 10, "bold")).pack(anchor="w", pady=(5, 2))
        frame_excel = ttk.Frame(inputs_frame)
        frame_excel.pack(fill="x", pady=(0, 10))
        self.entrada_excel = ttk.Entry(frame_excel)
        self.entrada_excel.pack(side="left", fill="x", expand=True, padx=(0, 5))
        ttk.Button(frame_excel, text="Selecionar", command=self.selecionar_excel_empresas,
                   bootstyle="info-outline").pack(side="right")

        # Pasta de saída
        ttk.Label(inputs_frame, text="Pasta para salvar resultado:", font=("Arial", 10, "bold")).pack(anchor="w", pady=(5, 2))
        frame_saida = ttk.Frame(inputs_frame)
        frame_saida.pack(fill="x", pady=(0, 10))
        self.entrada_pasta_saida = ttk.Entry(frame_saida)
        self.entrada_pasta_saida.pack(side="left", fill="x", expand=True, padx=(0, 5))
        ttk.Button(frame_saida, text="Selecionar", command=self.selecionar_pasta_saida,
                   bootstyle="info-outline").pack(side="right")

        # Opções avançadas
        options_frame = ttk.LabelFrame(inputs_frame, text="Opções Avançadas", bootstyle="secondary")
        options_frame.pack(fill="x", pady=10)

        self.incluir_detalhes_debitos = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Incluir detalhes de débitos pendentes",
                        variable=self.incluir_detalhes_debitos, bootstyle="round-toggle").pack(anchor="w", padx=10, pady=2)

        self.agrupar_por_empresa = tk.BooleanVar(value=False)
        ttk.Checkbutton(options_frame, text="Agrupar resultados por empresa",
                        variable=self.agrupar_por_empresa, bootstyle="round-toggle").pack(anchor="w", padx=10, pady=2)

        self.salvar_backup_json = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Salvar backup em JSON",
                        variable=self.salvar_backup_json, bootstyle="round-toggle").pack(anchor="w", padx=10, pady=2)

        # Botões de ação
        buttons_frame = ttk.Frame(inputs_frame)
        buttons_frame.pack(pady=20)

        self.btn_processar = ttk.Button(buttons_frame, text="Processar PDFs", command=self.processar_pdfs,
                                        bootstyle="success", padding=(30, 10))
        self.btn_processar.pack(side="left", padx=10)

        ttk.Button(buttons_frame, text="Limpar Resultados", command=self.limpar_resultados,
                   bootstyle="danger-outline", padding=(20, 10)).pack(side="left", padx=10)

        ttk.Button(buttons_frame, text="Salvar Configuração", command=self.salvar_config,
                   bootstyle="info", padding=(20, 10)).pack(side="left", padx=10)

        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(inputs_frame, variable=self.progress_var,
                                            maximum=100, bootstyle="success-striped")
        self.progress_bar.pack(fill="x", pady=10)

        # Status atual
        self.status_label = ttk.Label(inputs_frame, text="Aguardando configuração...",
                                      font=("Arial", 9), bootstyle="secondary")
        self.status_label.pack(pady=5)

        # Área de resultados do processamento
        ttk.Label(inputs_frame, text="Log do Processamento:", font=("Arial", 10, "bold")).pack(anchor="w", pady=(20, 5))
        self.text_resultados = ScrolledText(inputs_frame, height=12, font=("Courier", 9))
        self.text_resultados.pack(fill="both", expand=True, pady=(0, 20))

    def criar_aba_resultados(self):
        # Aba 2: Resultados
        frame_resultados = ttk.Frame(self.notebook)
        self.notebook.add(frame_resultados, text="  Parcelamentos")

        # Toolbar - Linha 1
        toolbar = ttk.Frame(frame_resultados)
        toolbar.pack(fill="x", padx=10, pady=5)

        # Filtros - Linha 1
        ttk.Label(toolbar, text="Filtros:", font=("Arial", 10, "bold"), bootstyle="info").grid(row=0, column=0, padx=5, sticky="w")

        ttk.Label(toolbar, text="Empresa/CNPJ:").grid(row=0, column=1, padx=5)
        self.filtro_empresa = ttk.Entry(toolbar, width=25)
        self.filtro_empresa.grid(row=0, column=2, padx=5)
        self.filtro_empresa.bind("<KeyRelease>", lambda e: self.aplicar_filtros())

        ttk.Label(toolbar, text="Tipo:").grid(row=0, column=3, padx=5)
        self.filtro_tipo = ttk.Combobox(
            toolbar,
            width=15,
            values=["Todos", "PARCMEI", "PARCSN", "SIEFPAR", "SISPAR", "SICOB", "SIPADE", "DÉBITO"]
        )
        self.filtro_tipo.set("Todos")
        self.filtro_tipo.grid(row=0, column=4, padx=5)
        self.filtro_tipo.bind("<<ComboboxSelected>>", lambda e: self.aplicar_filtros())

        ttk.Label(toolbar, text="Status:").grid(row=0, column=5, padx=5)
        self.filtro_status = ttk.Combobox(
            toolbar,
            width=20,
            values=["Todos", "Em Parcelamento", "Exigibilidade Suspensa", "Ativo/Em Dia", "Devedor"]
        )
        self.filtro_status.set("Todos")
        self.filtro_status.grid(row=0, column=6, padx=5)
        self.filtro_status.bind("<<ComboboxSelected>>", lambda e: self.aplicar_filtros())

        # Filtros - Linha 2
        toolbar2 = ttk.Frame(frame_resultados)
        toolbar2.pack(fill="x", padx=10, pady=(0, 5))

        ttk.Label(toolbar2, text="Com Parcelas em Atraso:").grid(row=0, column=0, padx=5)
        self.filtro_atraso = ttk.Combobox(
            toolbar2,
            width=10,
            values=["Todos", "Sim", "Não"]
        )
        self.filtro_atraso.set("Todos")
        self.filtro_atraso.grid(row=0, column=1, padx=5)
        self.filtro_atraso.bind("<<ComboboxSelected>>", lambda e: self.aplicar_filtros())

        ttk.Label(toolbar2, text="Valor Mínimo (R$):").grid(row=0, column=2, padx=5)
        self.filtro_valor_min = ttk.Entry(toolbar2, width=15)
        self.filtro_valor_min.insert(0, "0")
        self.filtro_valor_min.grid(row=0, column=3, padx=5)
        self.filtro_valor_min.bind("<KeyRelease>", lambda e: self.aplicar_filtros())

        ttk.Button(toolbar2, text="Limpar Filtros", command=self.limpar_filtros,
                   bootstyle="warning-outline").grid(row=0, column=4, padx=10)

        # Ações
        ttk.Button(toolbar2, text="Exportar Filtrados", command=self.exportar_filtrados,
                   bootstyle="success-outline").grid(row=0, column=5, padx=5)
        ttk.Button(toolbar2, text="Copiar Selecionados", command=self.copiar_selecionados,
                   bootstyle="secondary").grid(row=0, column=6, padx=5)

        # Contador de resultados
        self.label_contador = ttk.Label(toolbar2, text="Nenhum resultado", font=("Arial", 9), bootstyle="secondary")
        self.label_contador.grid(row=0, column=7, padx=20)

        # Tabela de parcelamentos
        frame_tabela = ttk.Frame(frame_resultados)
        frame_tabela.pack(fill="both", expand=True, padx=10, pady=10)

        # Colunas da tabela
        colunas = ("Código", "Empresa", "CNPJ", "Tipo", "Subtipo", "Conta", "Modalidade", "Status", "Detalhes", "Valor", "Arquivo")
        self.tree_parcelamentos = ttk.Treeview(frame_tabela, columns=colunas, show="headings", height=20, bootstyle="dark")

        # Configurar colunas
        larguras = {"Código": 80, "Empresa": 180, "CNPJ": 120, "Tipo": 80, "Subtipo": 100, "Conta": 100,
                   "Modalidade": 180, "Status": 120, "Detalhes": 200, "Valor": 100, "Arquivo": 150}

        for col in colunas:
            self.tree_parcelamentos.heading(col, text=col, command=lambda c=col: self.ordenar_coluna(c))
            self.tree_parcelamentos.column(col, width=larguras.get(col, 100))

        # Scrollbars
        scroll_y = ttk.Scrollbar(frame_tabela, orient="vertical", command=self.tree_parcelamentos.yview)
        scroll_x = ttk.Scrollbar(frame_tabela, orient="horizontal", command=self.tree_parcelamentos.xview)
        self.tree_parcelamentos.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        self.tree_parcelamentos.pack(side="left", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")

        # Context menu
        self.menu_contexto = tk.Menu(self.janela, tearoff=0)
        self.menu_contexto.add_command(label="Copiar CNPJ", command=self.copiar_cnpj)
        self.menu_contexto.add_command(label="Copiar linha completa", command=self.copiar_linha)
        self.menu_contexto.add_separator()
        self.menu_contexto.add_command(label="Abrir PDF", command=self.abrir_pdf)

        self.tree_parcelamentos.bind("<Button-3>", self.mostrar_menu_contexto)

    def criar_aba_dashboard(self):
        # Aba 3: Dashboard
        frame_dashboard = ttk.Frame(self.notebook)
        self.notebook.add(frame_dashboard, text="  Dashboard")

        # Container principal com scroll
        canvas_dash = tk.Canvas(frame_dashboard, highlightthickness=0)
        scrollbar_dash = ttk.Scrollbar(frame_dashboard, orient="vertical", command=canvas_dash.yview)
        scroll_frame = ttk.Frame(canvas_dash)

        scroll_frame.bind(
            "<Configure>",
            lambda e: canvas_dash.configure(scrollregion=canvas_dash.bbox("all"))
        )

        canvas_dash.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas_dash.configure(yscrollcommand=scrollbar_dash.set)

        canvas_dash.pack(side="left", fill="both", expand=True)
        scrollbar_dash.pack(side="right", fill="y")

        # Estatísticas gerais com cards visuais (mantém tk.Frame/tk.Label para cores customizadas)
        stats_frame = ttk.LabelFrame(scroll_frame, text="Visão Geral", bootstyle="info")
        stats_frame.pack(fill="x", padx=20, pady=10)

        # Grid para estatísticas
        self.stats_labels = {}
        stats_info = [
            ("total_empresas", "Total de Empresas", "#1976D2"),
            ("total_parcelamentos", "Total de Parcelamentos", "#388E3C"),
            ("empresas_com_parcelas_atraso", "Com Parcelas em Atraso", "#D32F2F"),
            ("valor_total_suspenso", "Valor Total", "#F57C00")
        ]

        for i, (key, label, cor) in enumerate(stats_info):
            frame_stat = tk.Frame(stats_frame, bg=cor, relief="raised", bd=2)
            frame_stat.grid(row=0, column=i, padx=10, pady=10, sticky="ew")

            tk.Label(frame_stat, text=label, font=("Arial", 9), bg=cor, fg="white").pack(pady=(5, 0))
            self.stats_labels[key] = tk.Label(frame_stat, text="0", font=("Arial", 16, "bold"), bg=cor, fg="white")
            self.stats_labels[key].pack(pady=(0, 5))

        # Configurar colunas para distribuir igualmente
        for i in range(4):
            stats_frame.grid_columnconfigure(i, weight=1)

        # Empresas Críticas (Top 10 com mais parcelas em atraso)
        criticas_frame = ttk.LabelFrame(scroll_frame, text="Empresas Críticas (Mais Parcelas em Atraso)", bootstyle="danger")
        criticas_frame.pack(fill="x", padx=20, pady=10)

        colunas_criticas = ("Empresa", "CNPJ", "Parcelas em Atraso", "Tipo")
        self.tree_criticas = ttk.Treeview(criticas_frame, columns=colunas_criticas, show="headings", height=10, bootstyle="dark")

        for col in colunas_criticas:
            self.tree_criticas.heading(col, text=col)

        self.tree_criticas.column("Empresa", width=250)
        self.tree_criticas.column("CNPJ", width=130)
        self.tree_criticas.column("Parcelas em Atraso", width=150)
        self.tree_criticas.column("Tipo", width=100)

        self.tree_criticas.pack(fill="x", padx=10, pady=10)

        # Ranking por Valor
        ranking_frame = ttk.LabelFrame(scroll_frame, text="Top 10 Maiores Valores", bootstyle="warning")
        ranking_frame.pack(fill="x", padx=20, pady=10)

        colunas_ranking = ("Empresa", "CNPJ", "Valor Total", "Qtd. Parcelamentos")
        self.tree_ranking = ttk.Treeview(ranking_frame, columns=colunas_ranking, show="headings", height=10, bootstyle="dark")

        for col in colunas_ranking:
            self.tree_ranking.heading(col, text=col)

        self.tree_ranking.column("Empresa", width=250)
        self.tree_ranking.column("CNPJ", width=130)
        self.tree_ranking.column("Valor Total", width=150)
        self.tree_ranking.column("Qtd. Parcelamentos", width=150)

        self.tree_ranking.pack(fill="x", padx=10, pady=10)

        # Resumo por tipo com gráfico visual
        resumo_frame = ttk.LabelFrame(scroll_frame, text="Distribuição por Tipo", bootstyle="success")
        resumo_frame.pack(fill="both", expand=True, padx=20, pady=10)

        # Tabela resumo
        colunas_resumo = ("Tipo", "Quantidade", "Empresas", "Percentual", "Gráfico")
        self.tree_resumo = ttk.Treeview(resumo_frame, columns=colunas_resumo, show="headings", height=8, bootstyle="dark")

        for col in colunas_resumo:
            self.tree_resumo.heading(col, text=col)

        self.tree_resumo.column("Tipo", width=100)
        self.tree_resumo.column("Quantidade", width=100)
        self.tree_resumo.column("Empresas", width=100)
        self.tree_resumo.column("Percentual", width=100)
        self.tree_resumo.column("Gráfico", width=300)

        self.tree_resumo.pack(fill="both", expand=True, padx=10, pady=10)

    # Métodos de interface
    def selecionar_pasta_pdfs(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta com os PDFs")
        if pasta:
            self.entrada_pasta_pdfs.delete(0, tk.END)
            self.entrada_pasta_pdfs.insert(0, pasta)

    def selecionar_excel_empresas(self):
        arquivo = filedialog.askopenfilename(
            title="Selecione o Excel com as empresas filtradas",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
        )
        if arquivo:
            self.entrada_excel.delete(0, tk.END)
            self.entrada_excel.insert(0, arquivo)

    def selecionar_pasta_saida(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta para salvar o resultado")
        if pasta:
            self.entrada_pasta_saida.delete(0, tk.END)
            self.entrada_pasta_saida.insert(0, pasta)

    def normalizar_cnpj(self, cnpj):
        """Remove formatação do CNPJ e retorna apenas números"""
        if not cnpj:
            return ""
        return re.sub(r'\D', '', cnpj)

    def obter_codigo_empresa(self, cnpj_numeros):
        """Retorna o código da empresa a partir do CNPJ, se existir no dicionário"""
        return self.empresas_codigos.get(cnpj_numeros, "")

    def extrair_valor_monetario(self, texto):
        """Extrai valores monetários do texto"""
        pattern = r'[\d\.,]+(?=\s*(?:reais?|R\$|\b))'
        matches = re.findall(pattern, texto)
        if matches:
            # Pega o maior valor encontrado
            valores = []
            for match in matches:
                try:
                    valor = float(match.replace('.', '').replace(',', '.'))
                    valores.append(valor)
                except:
                    continue
            return max(valores) if valores else 0
        return 0

    def extrair_dados_pdf(self, caminho_pdf):
        dados = []
        nome_arquivo = os.path.basename(caminho_pdf)
        
        try:
            with pdfplumber.open(caminho_pdf) as pdf:
                texto = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())

                # Extrair CNPJ e Nome da empresa
                cnpj_match = re.search(r"CNPJ:\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto)
                cnpj_formatado = cnpj_match.group(1) if cnpj_match else "Não encontrado"
                cnpj_numeros = self.normalizar_cnpj(cnpj_formatado)

                # Extrair nome da empresa
                nome_match = re.search(r"CNPJ:\s*\d{2}\.\d{3}\.\d{3}.*?-\s*(.+)", texto)
                nome_empresa = nome_match.group(1).strip() if nome_match else "Não encontrado"

                # Obter código da empresa
                codigo_empresa = self.obter_codigo_empresa(cnpj_numeros)

                # 1) PARCMEI - MEI
                if "MEI - EM PARCELAMENTO" in texto:
                    mei_match = re.search(r"MEI - EM PARCELAMENTO\s+Parcelas em atraso\s*(\d+)", texto)
                    if mei_match:
                        dados.append({
                            "Codigo": codigo_empresa,
                            "CNPJ": cnpj_formatado,
                            "CNPJ_Numeros": cnpj_numeros,
                            "Nome_Empresa": nome_empresa,
                            "Tipo": "PARCMEI",
                            "Subtipo": "MEI",
                            "Conta": "-",
                            "Modalidade": "MEI - Parcelamento",
                            "Detalhes": f"Parcelas em atraso: {mei_match.group(1)}",
                            "Status": "Em Parcelamento",
                            "Valor": 0,
                            "Arquivo": nome_arquivo
                        })

                # 2) PARCSN - Simples Nacional
                if "SIMPLES NACIONAL" in texto and "EM PARCELAMENTO" in texto:
                    # Regex aprimorado para capturar todas as variações (RELP, PERT, etc.)
                    sn_matches = re.findall(r"(SIMPLES NACIONAL(?:\s*-\s*\w+)?\s*-\s*EM PARCELAMENTO)(?:\s+Parcelas em atraso\s*(\d+))?", texto)
                    for match in sn_matches:
                        titulo, parcelas = match
                        parcelas = parcelas if parcelas else "0"
                        dados.append({
                            "Codigo": codigo_empresa,
                            "CNPJ": cnpj_formatado,
                            "CNPJ_Numeros": cnpj_numeros,
                            "Nome_Empresa": nome_empresa,
                            "Tipo": "PARCSN",
                            "Subtipo": "Simples Nacional",
                            "Conta": "-",
                            "Modalidade": titulo.strip(),
                            "Detalhes": f"Parcelas em atraso: {parcelas}",
                            "Status": "Em Parcelamento",
                            "Valor": 0,
                            "Arquivo": nome_arquivo
                        })


                # 3) SIEFPAR - Parcelamento com Exigibilidade Suspensa (Receita Federal)
                if "Pendência – Parcelamento (SIEFPAR)" in texto:
                    pattern = r"Parcelamento:\s*(\d+)\s+Parcelas em Atraso:\s*(\d+)\s+Valor em Atraso:\s*([\d\.,]+)"
                    matches = re.findall(pattern, texto)
                    for conta, parcelas, valor_str in matches:
                        valor = self.extrair_valor_monetario(valor_str)
                        dados.append({
                            "Codigo": codigo_empresa,
                            "Tipo": "SIEFPAR",
                            "Subtipo": "Receita Federal",
                            "Conta": conta.strip(),
                            "Modalidade": "Parcelamento Simplificado",
                            "Detalhes": f"Parcelas em atraso: {parcelas}, Valor em atraso: R$ {valor_str}",
                            "Status": "Exigibilidade Suspensa",
                            "Valor": valor,
                            "Arquivo": nome_arquivo,
                            "CNPJ": cnpj_formatado,
                            "CNPJ_Numeros": cnpj_numeros,
                            "Nome_Empresa": nome_empresa,
                        })

                if "Parcelamento com Exigibilidade Suspensa (SIEFPAR)" in texto:
                    pattern = r"Parcelamento:\s*(\d+)\s+Valor Suspenso:\s*([\d\.,]+)"
                    matches = re.findall(pattern, texto)
                    for conta, valor_str in matches:
                        valor = self.extrair_valor_monetario(valor_str)
                        dados.append({
                            "Codigo": codigo_empresa,
                            "Tipo": "SIEFPAR",
                            "Subtipo": "Receita Federal",
                            "Conta": conta.strip(),
                            "Modalidade": "Parcelamento Simplificado",
                            "Detalhes": f"Valor suspenso: R$ {valor_str}",
                            "Status": "Exigibilidade Suspensa",
                            "Valor": valor,
                            "Arquivo": nome_arquivo,
                            "CNPJ": cnpj_formatado,
                            "CNPJ_Numeros": cnpj_numeros,
                            "Nome_Empresa": nome_empresa,
                        })

                
                # # 4) SISPAR - Parcelamento com Exigibilidade Suspensa (PGFN)
                # if "Parcelamento com Exigibilidade Suspensa (SISPAR)" in texto:
                #     sispar_pattern = r"Conta\s+(\d+)\s+([^\n]+?)\s+Modalidade:\s*([^\n]+)"
                #     matches = re.findall(sispar_pattern, texto)
                    
                #     for conta, tipo_parcela, modalidade in matches:
                #         dados.append({
                #             "CNPJ": cnpj_formatado,
                #             "CNPJ_Numeros": cnpj_numeros,
                #             "Nome_Empresa": nome_empresa,
                #             "Tipo": "SISPAR",
                #             "Subtipo": "PGFN",
                #             "Conta": conta.strip(),
                #             "Modalidade": modalidade.strip(),
                #             "Detalhes": tipo_parcela.strip(),
                #             "Status": "Exigibilidade Suspensa",
                #             "Valor": 0,
                #             "Arquivo": nome_arquivo
                #         })

                if "SISPAR" in texto:
                    sispar_pattern = r"(?:Conta\s*)?(\d+)\s+([^\n]+)\nModalidade:\s*([^\n]+)"
                    matches = re.findall(sispar_pattern, texto)

                    for conta, tipo_parcela, modalidade in matches:
                        dados.append({
                            "Codigo": codigo_empresa,
                            "CNPJ": cnpj_formatado,
                            "CNPJ_Numeros": cnpj_numeros,
                            "Nome_Empresa": nome_empresa,
                            "Tipo": "SISPAR",
                            "Subtipo": "PGFN",
                            "Conta": conta.strip(),
                            "Modalidade": modalidade.strip(),
                            "Detalhes": tipo_parcela.strip(),
                            "Status": "Exigibilidade Suspensa",
                            "Valor": 0,
                            "Arquivo": nome_arquivo
                        })




                # 5) SICOB - Débito com Exigibilidade Suspensa
                if "Débito com Exigibilidade Suspensa (SICOB)" in texto:
                    sicob_pattern = r"Parcelamento:\s*(\d+-\d+)\s+Situação:\s*(\d+\s*-\s*.+)"
                    matches = re.findall(sicob_pattern, texto)

                    for parcela, situacao in matches:
                        dados.append({
                            "Codigo": codigo_empresa,
                            "CNPJ": cnpj_formatado,
                            "CNPJ_Numeros": cnpj_numeros,
                            "Nome_Empresa": nome_empresa,
                            "Tipo": "SICOB",
                            "Subtipo": "Débito Suspenso",
                            "Conta": parcela.strip(),
                            "Modalidade": "RFB LEI 10522/02",
                            "Detalhes": f"Situação: {situacao}",
                            "Status": "Ativo/Em Dia",
                            "Valor": 0,
                            "Arquivo": nome_arquivo
                        })

                    # 6) SIPADE - Pendências SIPADE
                if "Parcelamento com Exigibilidade Suspensa (SIPADE)" in texto:
                    # Localizar a seção SIPADE e extrair os dados linha por linha
                    sipade_inicio = texto.find("Parcelamento com Exigibilidade Suspensa (SIPADE)")
                    sipade_fim = texto.find("Processo Fiscal com Exigibilidade Suspensa (SIEF)", sipade_inicio)
                    if sipade_fim == -1:
                        sipade_fim = texto.find("Parcelamento com Exigibilidade Suspensa (SIEFPAR)", sipade_inicio)
                    if sipade_fim == -1:
                        sipade_fim = len(texto)

                    secao_sipade = texto[sipade_inicio:sipade_fim]

                    # Padrão: número do processo seguido de receita e situação na mesma linha ou próximas
                    sipade_pattern = r"(\d{5}\.\d{3}\.\d{3}/\d{4}-\d{2})\s+(\d{4}-[A-Z]+)\s+(ATIVO|PENDENTE|SUSPENSO|CANCELADO)"
                    matches = re.findall(sipade_pattern, secao_sipade)

                    for processo, receita, situacao in matches:
                        dados.append({
                            "Codigo": codigo_empresa,
                            "CNPJ": cnpj_formatado,
                            "CNPJ_Numeros": cnpj_numeros,
                            "Nome_Empresa": nome_empresa,
                            "Tipo": "SIPADE",
                            "Subtipo": "Parcelamento SIPADE",
                            "Conta": processo.strip(),
                            "Modalidade": f"Receita: {receita.strip()}",
                            "Detalhes": f"Situação: {situacao.strip()}",
                            "Status": "Exigibilidade Suspensa",
                            "Valor": 0,
                            "Arquivo": nome_arquivo
                        })

                # Incluir detalhes de débitos se solicitado
                if self.incluir_detalhes_debitos.get() and "Pendência - Débito (SIEF)" in texto:
                    # Extrair débitos pendentes
                    debito_pattern = r"(\d{4}-\d{2}\s*-\s*.+?)\s+(\d{2}/\d{4})\s+[\d/]+\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)\s+(.+)"
                    debitos = re.findall(debito_pattern, texto)
                    
                    for receita, periodo, dt_vcto, vl_orig, sdo_dev, multa, juros, sdo_cons, situacao in debitos[:5]:  # Limita a 5 débitos
                        valor_total = self.extrair_valor_monetario(sdo_cons) if sdo_cons else 0
                        dados.append({
                            "Codigo": codigo_empresa,
                            "CNPJ": cnpj_formatado,
                            "CNPJ_Numeros": cnpj_numeros,
                            "Nome_Empresa": nome_empresa,
                            "Tipo": "DÉBITO",
                            "Subtipo": "Pendência",
                            "Conta": receita.strip(),
                            "Modalidade": f"Período: {periodo}",
                            "Detalhes": f"Situação: {situacao.strip()}",
                            "Status": "Devedor",
                            "Valor": valor_total,
                            "Arquivo": nome_arquivo
                        })

        except Exception as e:
            self.log(f"Erro ao processar {nome_arquivo}: {str(e)}")
            
        return dados

    def processar_pdfs(self):
        pasta_pdfs = self.entrada_pasta_pdfs.get()
        excel_empresas = self.entrada_excel.get()
        pasta_saida = self.entrada_pasta_saida.get()

        if not pasta_pdfs:
            messagebox.showerror("Erro", "Selecione a pasta com os PDFs!")
            return

        if not pasta_saida:
            messagebox.showerror("Erro", "Selecione a pasta de saída!")
            return

        def processar():
            try:
                self.limpar_resultados()
                inicio = datetime.now()
                
                # Carrega lista de empresas se fornecida
                if excel_empresas:
                    try:
                        df_empresas = pd.read_excel(excel_empresas, dtype=str)
                        if 'CNPJ' in df_empresas.columns:
                            self.empresas_filtradas = set(df_empresas['CNPJ'].apply(self.normalizar_cnpj))

                            # Mapear CNPJ -> Código usando coluna de índice 0 (coluna A)
                            for _, row in df_empresas.iterrows():
                                cnpj_norm = self.normalizar_cnpj(str(row['CNPJ']))
                                # Pega o valor da primeira coluna (índice 0)
                                codigo = str(row.iloc[0]) if pd.notna(row.iloc[0]) else ""
                                self.empresas_codigos[cnpj_norm] = codigo
                            self.log(f"✅ Carregadas {len(self.empresas_filtradas)} empresas com códigos do Excel (coluna A)")
                        else:
                            self.log(f"⚠️ Excel não contém coluna 'CNPJ'")
                    except Exception as e:
                        self.log(f"⚠️ Erro ao carregar Excel: {str(e)}")

                # Lista arquivos PDF
                arquivos_pdf = [f for f in os.listdir(pasta_pdfs) if f.lower().endswith('.pdf')]
                total_arquivos = len(arquivos_pdf)
                
                self.log(f"📁 Encontrados {total_arquivos} PDFs para processar...")
                self.progress_var.set(0)

                todos_dados = []
                for i, arquivo in enumerate(arquivos_pdf, 1):
                    caminho = os.path.join(pasta_pdfs, arquivo)
                    self.status_label.config(text=f"Processando {i}/{total_arquivos}: {arquivo}")
                    self.log(f"[{i}/{total_arquivos}] {arquivo}")
                    
                    dados = self.extrair_dados_pdf(caminho)
                    # Adiciona o código do Excel (se existir) a cada registro
                    for d in dados:
                        cnpj_num = d.get('CNPJ_Numeros', '')
                        d['Codigo'] = self.empresas_codigos.get(cnpj_num, '')
                    
                    # Aplicar filtro de empresas se existe
                    if self.empresas_filtradas:
                        dados_filtrados = [d for d in dados if d['CNPJ_Numeros'] in self.empresas_filtradas]
                        if dados_filtrados:
                            self.log(f"  ✅ {len(dados_filtrados)} parcelamentos encontrados (filtrado)")
                        dados = dados_filtrados
                    else:
                        if dados:
                            self.log(f"  📊 {len(dados)} parcelamentos encontrados")
                    
                    todos_dados.extend(dados)
                    
                    # Atualizar progress bar
                    progresso = (i / total_arquivos) * 100
                    self.progress_var.set(progresso)
                    self.janela.update()

                # Processar e salvar resultados
                if todos_dados:
                    self.dados_processados = todos_dados
                    df = pd.DataFrame(todos_dados)
                    
                    # Agrupar por empresa se solicitado
                    if self.agrupar_por_empresa.get():
                        df = df.sort_values(['Nome_Empresa', 'Tipo'])
                    else:
                        df = df.sort_values(['Tipo', 'Nome_Empresa'])
                    
                    # Salvar Excel
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    caminho_excel = os.path.join(pasta_saida, f"parcelamentos_detalhados_{timestamp}.xlsx")
                    df.to_excel(caminho_excel, index=False)
                    
                    # Salvar backup JSON se solicitado
                    if self.salvar_backup_json.get():
                        caminho_json = os.path.join(pasta_saida, f"parcelamentos_backup_{timestamp}.json")
                        with open(caminho_json, 'w', encoding='utf-8') as f:
                            json.dump(todos_dados, f, ensure_ascii=False, indent=2)
                    
                    # Atualizar interface
                    self.atualizar_tabela()
                    self.atualizar_dashboard()
                    
                    fim = datetime.now()
                    tempo_total = (fim - inicio).total_seconds()
                    
                    self.log(f"\n✅ PROCESSAMENTO CONCLUÍDO!")
                    self.log(f"⏱️ Tempo total: {tempo_total:.1f} segundos")
                    self.log(f"📊 Total de parcelamentos: {len(todos_dados)}")
                    self.log(f"🏢 Empresas processadas: {df['Nome_Empresa'].nunique()}")
                    self.log(f"💾 Arquivo salvo: {caminho_excel}")
                    
                else:
                    self.log(f"\n⚠️ Nenhum parcelamento foi encontrado nos PDFs!")

            except Exception as e:
                self.log(f"\n❌ ERRO: {str(e)}")
            
            finally:
                self.btn_processar.config(state="normal", text="Processar PDFs")
                self.status_label.config(text="Processamento concluído")
                self.progress_var.set(100)

        # Executar em thread separada
        self.btn_processar.config(state="disabled", text="Processando...")
        thread = threading.Thread(target=processar)
        thread.daemon = True
        thread.start()

    def log(self, mensagem):
        """Adiciona mensagem ao log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.text_resultados.insert(tk.END, f"[{timestamp}] {mensagem}\n")
        self.text_resultados.see(tk.END)
        self.janela.update()

    def atualizar_tabela(self, dados=None):
        """Atualiza a tabela de resultados"""
        if dados is None:
            dados = self.dados_processados

        # Salvar os dados filtrados atuais para permitir ordenação
        self.dados_filtrados_atual = dados.copy() if isinstance(dados, list) else dados

        # Limpar tabela
        for item in self.tree_parcelamentos.get_children():
            self.tree_parcelamentos.delete(item)

        # Preencher tabela
        for row in dados:
            valores = (
                row.get('Codigo', ''),
                row['Nome_Empresa'],
                row['CNPJ'],
                row['Tipo'],
                row['Subtipo'],
                row['Conta'],
                row['Modalidade'],
                row['Status'],
                row['Detalhes'],
                f"R$ {row['Valor']:,.2f}" if row['Valor'] > 0 else "-",
                row['Arquivo']
            )
            self.tree_parcelamentos.insert("", tk.END, values=valores)

        self.label_contador.config(text=f"{len(dados)} resultados")

    def atualizar_dashboard(self):
        """Atualiza as estatísticas do dashboard com informações visuais"""
        if not self.dados_processados:
            return

        df = pd.DataFrame(self.dados_processados)

        # Estatísticas gerais
        total_empresas = df['Nome_Empresa'].nunique()
        total_parcelamentos = len(df)
        empresas_atraso = len(df[df['Detalhes'].str.contains('Parcelas em atraso', na=False)]['Nome_Empresa'].unique())
        valor_total = df['Valor'].sum()

        self.stats_labels['total_empresas'].config(text=str(total_empresas))
        self.stats_labels['total_parcelamentos'].config(text=str(total_parcelamentos))
        self.stats_labels['empresas_com_parcelas_atraso'].config(text=str(empresas_atraso))
        self.stats_labels['valor_total_suspenso'].config(text=f"R$ {valor_total:,.2f}")

        # NOVO: Empresas Críticas (Top 10 com mais parcelas em atraso)
        for item in self.tree_criticas.get_children():
            self.tree_criticas.delete(item)

        df_com_atraso = df[df['Detalhes'].str.contains('Parcelas em atraso', na=False)].copy()
        if not df_com_atraso.empty:
            # Extrair número de parcelas em atraso
            df_com_atraso['Num_Atraso'] = df_com_atraso['Detalhes'].str.extract(r'atraso:\s*(\d+)').astype(float)

            # Agrupar por empresa e somar parcelas em atraso
            criticas = df_com_atraso.groupby(['Nome_Empresa', 'CNPJ', 'Tipo']).agg({
                'Num_Atraso': 'sum'
            }).reset_index()

            criticas = criticas.sort_values('Num_Atraso', ascending=False).head(10)

            for _, row in criticas.iterrows():
                self.tree_criticas.insert("", tk.END, values=(
                    row['Nome_Empresa'][:40],
                    row['CNPJ'],
                    int(row['Num_Atraso']),
                    row['Tipo']
                ))

        # NOVO: Ranking por Valor (Top 10)
        for item in self.tree_ranking.get_children():
            self.tree_ranking.delete(item)

        ranking = df.groupby(['Nome_Empresa', 'CNPJ']).agg({
            'Valor': 'sum',
            'Tipo': 'count'
        }).reset_index()

        ranking = ranking.sort_values('Valor', ascending=False).head(10)

        for _, row in ranking.iterrows():
            if row['Valor'] > 0:  # Só mostra se tiver valor
                self.tree_ranking.insert("", tk.END, values=(
                    row['Nome_Empresa'][:40],
                    row['CNPJ'],
                    f"R$ {row['Valor']:,.2f}",
                    int(row['Tipo'])
                ))

        # Resumo por tipo com gráfico visual
        for item in self.tree_resumo.get_children():
            self.tree_resumo.delete(item)

        resumo_tipo = df.groupby('Tipo').agg({
            'Nome_Empresa': 'nunique',
            'CNPJ': 'count'
        }).reset_index()

        resumo_tipo = resumo_tipo.sort_values('CNPJ', ascending=False)

        for _, row in resumo_tipo.iterrows():
            tipo = row['Tipo']
            empresas = row['Nome_Empresa']
            quantidade = row['CNPJ']
            percentual = (quantidade / total_parcelamentos) * 100

            # Criar gráfico de barras simples com caracteres
            barra_tamanho = int(percentual / 2)  # Cada caractere = 2%
            grafico = "█" * barra_tamanho + "░" * (50 - barra_tamanho)

            self.tree_resumo.insert("", tk.END, values=(
                tipo, quantidade, empresas, f"{percentual:.1f}%", grafico
            ))

    def aplicar_filtros(self):
        """Aplica filtros na tabela com busca em tempo real"""
        if not self.dados_processados:
            return

        dados_filtrados = self.dados_processados.copy()

        # Filtro por empresa (busca por nome, CNPJ formatado ou CNPJ numérico)
        filtro_emp = self.filtro_empresa.get().strip()
        if filtro_emp:
            chave_normalizada = re.sub(r"\D", "", filtro_emp)
            def matches_empresa(d):
                nome = (d.get('Nome_Empresa') or "").lower()
                cnpj_fmt = (d.get('CNPJ') or "").lower()
                cnpj_num = (d.get('CNPJ_Numeros') or "")
                if chave_normalizada:
                    return chave_normalizada in cnpj_num
                return filtro_emp.lower() in nome or filtro_emp.lower() in cnpj_fmt
            dados_filtrados = [d for d in dados_filtrados if matches_empresa(d)]

        # Filtro por tipo
        filtro_tip = self.filtro_tipo.get().strip()
        if filtro_tip and filtro_tip.lower() != "todos":
            dados_filtrados = [d for d in dados_filtrados if d.get('Tipo', '').upper() == filtro_tip.upper()]

        # Filtro por status
        filtro_st = self.filtro_status.get().strip()
        if filtro_st and filtro_st.lower() != "todos":
            dados_filtrados = [d for d in dados_filtrados if d.get('Status', '') == filtro_st]

        # Filtro por parcelas em atraso
        filtro_atr = self.filtro_atraso.get().strip()
        if filtro_atr and filtro_atr.lower() != "todos":
            if filtro_atr == "Sim":
                dados_filtrados = [d for d in dados_filtrados if 'Parcelas em atraso' in d.get('Detalhes', '') and 'atraso: 0' not in d.get('Detalhes', '')]
            else:
                dados_filtrados = [d for d in dados_filtrados if 'Parcelas em atraso' not in d.get('Detalhes', '') or 'atraso: 0' in d.get('Detalhes', '')]

        # Filtro por valor mínimo
        try:
            valor_min_str = self.filtro_valor_min.get().strip().replace(',', '.')
            valor_min = float(valor_min_str) if valor_min_str else 0
            if valor_min > 0:
                dados_filtrados = [d for d in dados_filtrados if d.get('Valor', 0) >= valor_min]
        except ValueError:
            pass  # Ignora se o valor não for um número válido

        self.atualizar_tabela(dados_filtrados)

    def limpar_filtros(self):
        """Limpa todos os filtros"""
        self.filtro_empresa.delete(0, tk.END)
        self.filtro_tipo.set("Todos")
        self.filtro_status.set("Todos")
        self.filtro_atraso.set("Todos")
        self.filtro_valor_min.delete(0, tk.END)
        self.filtro_valor_min.insert(0, "0")
        self.atualizar_tabela()

    def limpar_resultados(self):
        """Limpa todos os resultados"""
        self.dados_processados = []
        self.text_resultados.delete(1.0, tk.END)
        self.atualizar_tabela()
        self.progress_var.set(0)
        self.status_label.config(text="Resultados limpos")

    def salvar_config(self):
        """Salva configuração atual"""
        config = {
            "pasta_pdfs": self.entrada_pasta_pdfs.get(),
            "excel_empresas": self.entrada_excel.get(),
            "pasta_saida": self.entrada_pasta_saida.get(),
            "incluir_detalhes_debitos": self.incluir_detalhes_debitos.get(),
            "agrupar_por_empresa": self.agrupar_por_empresa.get(),
            "salvar_backup_json": self.salvar_backup_json.get()
        }
        
        arquivo_config = filedialog.asksaveasfilename(
            title="Salvar configuração",
            defaultextension=".json",
            filetypes=[("JSON", "*.json")]
        )
        
        if arquivo_config:
            with open(arquivo_config, 'w') as f:
                json.dump(config, f, indent=2)
            messagebox.showinfo("Sucesso", "Configuração salva!")

    def exportar_filtrados(self):
        """Exporta dados atualmente filtrados"""
        items = self.tree_parcelamentos.get_children()
        if not items:
            messagebox.showwarning("Aviso", "Nenhum dado para exportar!")
            return
            
        arquivo = filedialog.asksaveasfilename(
            title="Exportar dados filtrados",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")]
        )
        
        if arquivo:
            # Coletar dados da tabela
            dados_exportar = []
            for item in items:
                valores = self.tree_parcelamentos.item(item)['values']
                dados_exportar.append(dict(zip(
                    ["Empresa", "CNPJ", "Tipo", "Subtipo", "Conta", "Modalidade", "Status", "Detalhes", "Valor", "Arquivo"],
                    valores
                )))
            
            df = pd.DataFrame(dados_exportar)
            
            if arquivo.endswith('.xlsx'):
                df.to_excel(arquivo, index=False)
            else:
                df.to_csv(arquivo, index=False, encoding='utf-8-sig')
                
            messagebox.showinfo("Sucesso", f"Dados exportados: {arquivo}")

    def copiar_selecionados(self):
        """Copia itens selecionados para clipboard"""
        selecionados = self.tree_parcelamentos.selection()
        if not selecionados:
            messagebox.showwarning("Aviso", "Selecione pelo menos um item!")
            return
        
        texto = ""
        for item in selecionados:
            valores = self.tree_parcelamentos.item(item)['values']
            texto += "\t".join(str(v) for v in valores) + "\n"
        
        self.janela.clipboard_clear()
        self.janela.clipboard_append(texto)
        messagebox.showinfo("Sucesso", f"{len(selecionados)} itens copiados!")

    def ordenar_coluna(self, coluna):
        """Ordena tabela por coluna clicada"""
        if not self.dados_filtrados_atual:
            return

        # Alternar direção de ordenação
        if coluna not in self.colunas_ordem:
            self.colunas_ordem[coluna] = True  # True = crescente
        else:
            self.colunas_ordem[coluna] = not self.colunas_ordem[coluna]

        reverse = not self.colunas_ordem[coluna]

        # Mapear nomes de colunas para chaves do dicionário
        mapa_colunas = {
            "Código": "Codigo",
            "Empresa": "Nome_Empresa",
            "CNPJ": "CNPJ",
            "Tipo": "Tipo",
            "Subtipo": "Subtipo",
            "Conta": "Conta",
            "Modalidade": "Modalidade",
            "Status": "Status",
            "Detalhes": "Detalhes",
            "Valor": "Valor",
            "Arquivo": "Arquivo"
        }

        chave = mapa_colunas.get(coluna)
        if not chave:
            return

        # Ordenar os dados
        try:
            if chave == "Valor":
                # Ordenação numérica para valores
                self.dados_filtrados_atual.sort(key=lambda x: x.get(chave, 0), reverse=reverse)
            else:
                # Ordenação alfabética para outros campos
                self.dados_filtrados_atual.sort(key=lambda x: str(x.get(chave, "")).lower(), reverse=reverse)
        except Exception as e:
            print(f"Erro ao ordenar: {e}")
            return

        # Atualizar a tabela com dados ordenados
        self.atualizar_tabela(self.dados_filtrados_atual)

    def mostrar_menu_contexto(self, event):
        """Mostra menu de contexto"""
        try:
            item = self.tree_parcelamentos.identify_row(event.y)
            if item:
                self.tree_parcelamentos.selection_set(item)
                self.menu_contexto.post(event.x_root, event.y_root)
        except:
            pass

    def copiar_cnpj(self):
        """Copia CNPJ selecionado"""
        item = self.tree_parcelamentos.selection()[0]
        cnpj = self.tree_parcelamentos.item(item)['values'][1]
        self.janela.clipboard_clear()
        self.janela.clipboard_append(cnpj)

    def copiar_linha(self):
        """Copia linha completa"""
        item = self.tree_parcelamentos.selection()[0]
        valores = self.tree_parcelamentos.item(item)['values']
        texto = "\t".join(str(v) for v in valores)
        self.janela.clipboard_clear()
        self.janela.clipboard_append(texto)

    def abrir_pdf(self):
        """Abre PDF correspondente"""
        item = self.tree_parcelamentos.selection()[0]
        arquivo = self.tree_parcelamentos.item(item)['values'][-1]
        pasta_pdfs = self.entrada_pasta_pdfs.get()
        caminho_pdf = os.path.join(pasta_pdfs, arquivo)
        
        if os.path.exists(caminho_pdf):
            os.startfile(caminho_pdf)  # Windows
        else:
            messagebox.showerror("Erro", "Arquivo PDF não encontrado!")

    def run(self):
        self.janela.mainloop()

# Executar aplicação
if __name__ == "__main__":
    app = AnalisadorParcelamentos()
    app.run()