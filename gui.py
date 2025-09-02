import customtkinter as ctk
from tkinter import ttk, messagebox, TclError, filedialog
from analise import analisar_dados, salvar_excel


# Define a aparência global da aplicação (light, dark, system)
ctk.set_appearance_mode("System")
# Define o tema de cores padrão
ctk.set_default_color_theme("tema_laranja.json")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("COMPARADOR DE PREÇOS - HOSTDIME")
        self.geometry("450x200")

        self.label_produtos = ctk.CTkLabel(self, text="Quantos produtos deseja comparar?", font=("Poppins", 15, 'bold'))
        self.label_produtos.pack(pady=(20, 5))

        self.entry_produtos = ctk.CTkEntry(self, width=140, placeholder_text="Ex: 3")
        self.entry_produtos.pack()

        self.start_button = ctk.CTkButton(self, text="INICIAR COMPARAÇÃO", command=self.abrir_janela_dados, font=("Poppins", 12,'bold'))
        self.start_button.pack(pady=20)
        
    def abrir_janela_dados(self):
        try:
            num_produtos = int(self.entry_produtos.get())
            if num_produtos <= 0: raise ValueError
            self.withdraw()
            DataEntryWindow(self, num_produtos)
        except ValueError:
            messagebox.showerror("Erro de entrada", "Por favor, insira um número inteiro e positivo.")


class DataEntryWindow(ctk.CTkToplevel):
    def __init__(self, parent, num_produtos):
        super().__init__(parent)
        self.parent = parent
        self.title("HOSTDIME - COMPRAS")
        self.geometry("750x550")

        # Estrutura com Scrollbar
        self.scrollable_frame = ctk.CTkScrollableFrame(self, label_text="DADOS DOS PRODUTOS",label_font=("Poppins", 15, "bold"))
        self.scrollable_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.product_widgets = []
        
        for i in range(num_produtos):
            produto_frame = ctk.CTkFrame(self.scrollable_frame)
            produto_frame.pack(padx=10, pady=4, fill="x", expand=True)
            
            ctk.CTkLabel(produto_frame, text="Nome do produto:",font=("Poppins",15,'bold')).grid(row=0, column=0, padx=10, pady=5, sticky='w')
            nome_entry = ctk.CTkEntry(produto_frame, width=200, placeholder_text=f"Nome do produto {i+1}")
            nome_entry.grid(row=0, column=1, padx=10, pady=5, sticky='w')
            
            ctk.CTkLabel(produto_frame, text="Nº de fornecedores:",font=("Poppins",15, 'bold')).grid(row=0, column=2, padx=10, pady=5, sticky='w')
            fornecedores_entry = ctk.CTkEntry(produto_frame, width=70)
            fornecedores_entry.grid(row=0, column=3, padx=10, pady=5, sticky='w')

           
            
            self.product_widgets.append({
            "nome_entry": nome_entry,
            "fornecedores_count_entry": fornecedores_entry,
            "parent_frame": produto_frame, # Guardamos a referência do frame PAI
            "container": None,             # O container ainda não existe
            "fornecedor_entries": []
})
        
        # Frame para os botões de ação fora do scroll
        action_frame = ctk.CTkFrame(self, fg_color="transparent")
        action_frame.pack(pady=10)

        ctk.CTkButton(action_frame, text="CAMPOS PARA PREENCHIMENTO",font=("Poppins",12,'bold'), command=self._gerar_campos_fornecedores).pack(side="left", padx=10)
        ctk.CTkButton(action_frame, text="GERAR COMPARATIVO",font=("Poppins",12,'bold'), command=self.processar_dados).pack(side="left", padx=10)

        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def _gerar_campos_fornecedores(self):
         for product in self.product_widgets:
            # Primeiro, verifica se o container já existe (não é None) antes de tentar limpá-lo
            if product["container"] is not None:
                product["container"].destroy()
            
            # O resto do seu código continua normalmente a partir daqui
            product["fornecedor_entries"] = []
            try:
                num_fornecedores = int(product["fornecedores_count_entry"].get())
                if num_fornecedores > 0:
                    # --- MUDANÇA IMPORTANTE ---
                    # O container AGORA é criado e posicionado
                    parent = product["parent_frame"]
                    container = ctk.CTkFrame(parent, fg_color="transparent")
                    container.grid(row=1, column=0, columnspan=4, pady=(10, 5), sticky="w")
                    product["container"] = container # Armazena a referência do novo container
                   # Cria os cabeçalhos e os campos dentro do novo container
                    ctk.CTkLabel(container, text="Fornecedor", font=('TkDefaultFont', 12, 'bold')).grid(row=0, column=0, padx=5, pady=2)
                    ctk.CTkLabel(container, text="Preço (R$)", font=('TkDefaultFont', 12, 'bold')).grid(row=0, column=1, padx=5, pady=2)
                    ctk.CTkLabel(container, text="Entrega", font=('TkDefaultFont', 12, 'bold')).grid(row=0, column=2, padx=5, pady=2)
               
                for i in range(num_fornecedores):
                    fornecedor = ctk.CTkEntry(product["container"], placeholder_text="Nome do fornecedor")
                    fornecedor.grid(row=i+1, column=0, padx=5, pady=2)
                    preco = ctk.CTkEntry(product["container"], placeholder_text="Ex: 19.90")
                    preco.grid(row=i+1, column=1, padx=5, pady=2)
                    entrega = ctk.CTkEntry(product["container"], placeholder_text="Ex: 5 dias")
                    entrega.grid(row=i+1, column=2, padx=5, pady=2)
                    product["fornecedor_entries"].append((fornecedor, preco, entrega))
            except (ValueError, TclError):
                pass

    def on_closing(self):
        self.parent.deiconify()
        self.destroy()

    def processar_dados(self):
        dados_brutos = []
        for product in self.product_widgets:
            nome_produto = product['nome_entry'].get().strip()
            if not nome_produto: continue
            for forn_entry, preco_entry, entrega_entry in product['fornecedor_entries']:
                fornecedor = forn_entry.get().strip()
                preco_str = preco_entry.get().strip().replace(',', '.')
                entrega = entrega_entry.get().strip()
                if fornecedor and preco_str:
                   dados_brutos.append({"Produto": nome_produto, "Fornecedor": fornecedor, "Preço": preco_str, "Entrega": entrega})
        
        df_final = analisar_dados(dados_brutos)
        
        if df_final.empty:
            messagebox.showwarning("Aviso", "Nenhum dado válido foi inserido. Certifique-se de gerar os campos e preenchê-los.")
            return

        ResultsWindow(self, df_final)


class ResultsWindow(ctk.CTkToplevel):
    def __init__(self, parent, df_final):
        super().__init__(parent)
        self.title("RESULTADO DA COMPARAÇÃO")
        self.geometry("800x400")
        self.df_final = df_final

        frame = ctk.CTkFrame(self)
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        # --- IMPORTANTE: Treeview não existe no CustomTkinter, então mantemos o do ttk ---
        # Para customizar a aparência dele para combinar com o tema, são necessárias mais etapas.
        # Por padrão, ele terá a aparência nativa do seu sistema operacional.
        tree = ttk.Treeview(frame, columns=list(df_final.columns), show="headings")
        tree.pack(fill="both", expand=True)
        for col in df_final.columns:
            tree.heading(col, text=col)
            tree.column(col, width=120, anchor='center')
        for index, row in df_final.iterrows():
            tree.insert("", "end", values=list(row))

        ctk.CTkButton(self, text="Salvar em Excel", command=self.salvar).pack(pady=10)

    def salvar(self):
        nome_arquivo = filedialog.asksaveasfilename(
            title="Salvar comparativo como...",
            defaultextension=".xlsx",
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
        )
        if not nome_arquivo:
            return
        salvar_excel(self.df_final, nome_arquivo)