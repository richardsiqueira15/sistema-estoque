import tkinter as tk
from tkcalendar import DateEntry
from tkinter import ttk, messagebox, Toplevel, simpledialog
from ttkthemes import ThemedTk
from PIL import ImageTk, Image
import tkinter.filedialog as filedialog
import csv
import json
import os
import ctypes
import sys
from datetime import datetime
import pandas as pd
import sqlite3

# print("Executando a versão mais recente do script...")

# --- Importação das Funções de Banco de Dados (funcoes_estoque.py) ---
from funcoes_estoque import (
    adicionar_produto, registrar_saida_log,
    consultar_movimentacoes, consultar_estoque_baixo, excluir_produto,
    consultar_estoque_geral, registrar_entradas_lote, registrar_saidas_lote,
    importar_produtos_excel, limpar_historico_produto, buscar_inventario,
    atualizar_estoque_minimo, atualizar_media_db, obter_lista_producao, listar_cadastros_aux,
    adicionar_cadastro_aux, remover_cadastro_aux, consultar_detalhes_reservado, atualizar_medias_em_lote,
    validar_acesso, configurar_banco_usuarios,
    # NOVAS FUNÇÕES DE PEDIDOS:
    criar_tabelas, registrar_pedido, consultar_pedidos, separar_pedido, consultar_pedido_por_id, 
    excluir_pedido, mover_pedido_para_expedicao, finalizar_pedido, verificar_estoque, atualizar_pedido,
    estornar_pedido, promover_pedido_para_pendente, validar_estoque_rascunhos, buscar_pedidos_por_status, excluir_rascunho_db,
    promover_pedido_com_corte_total, auditar_e_corrigir_reservas, exportar_faltantes_consolidado_pdf
)

criar_tabelas() 


# --- Suporte a telas de alto DPI (Windows) ---
try:
    # Aumenta a clareza em telas de alta resolução no Windows
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except:
    pass

# Variáveis Globais de Controle
USUARIO_ROLE = None
notebook = None

# Variáveis Globais de Widgets (Necessárias para as funções enxergarem os campos)
tree_inventario = None
tree_movimentacoes = None
tree_pedidos_dia = None
tree_pedidos_pendentes = None
tree_pedidos_separados = None
tree_pedidos_expedicao = None
tree_itens_pedido = None
entry_codigo = None
entry_nome = None
entry_estoque = None
entry_estoque_minimo = None

# --- Funcao para determinar o caminho correto para os arquivos de recurso ---
def resource_path(relative_path):
    """Obtém o caminho absoluto para o recurso, seja no script Python ou no EXE compilado."""
    import os
    import sys
    if hasattr(sys, '_MEIPASS'):
        # Caminho dentro do executável PyInstaller
        return os.path.join(sys._MEIPASS, relative_path)
    # Caminho no ambiente de desenvolvimento (pasta atual)
    return os.path.join(os.path.abspath("."), relative_path)

def obter_nomes_aux(tipo, padrao="SELECIONE"):
    """Busca nomes no banco e garante que a lista nunca venha vazia."""
    nomes = listar_cadastros_aux(tipo)
    return nomes if nomes else [padrao]

# --- Listas de Pessoas e Clientes ---
RESPONSAVEIS = ["Célio", "Rafael", "Pedro"]
SEPARADORES = ["David", "Yara", "Reginaldo", "Guilherme", "Rafael", "Caique", "João", "Montagem", "Almoxarifado", "Pet"]
TIPOS_ENTRADA = ["Produção", "Montagem", "Devolução", "Terceiros"]
CLIENTES_FILE = resource_path(os.path.join("arquivos_adicionais", "clientes.json"))
CLIENTES = set()

# Conjuntos para controlar itens de lote
codigos_entrada_lote = set()
codigos_saida_lote = set()

pedido_em_edicao_id = None
entry_pedido_cliente = None
entry_pedido_solicitante = None

# Carrega a lista de clientes do arquivo
def carregar_clientes():
    global CLIENTES
    if os.path.exists(CLIENTES_FILE):
        with open(CLIENTES_FILE, "r") as f:
            CLIENTES = set(json.load(f))
    else:
        CLIENTES = set()

# Salva a lista de clientes no arquivo
def salvar_clientes():
    with open(CLIENTES_FILE, "w") as f:
        json.dump(list(CLIENTES), f)

# --- Funções da Interface ---

def organizar_arvore(tree, col, reverse):
    """
    Função para organizar os itens de uma Treeview.
    """
    data = [(tree.set(k, col), k) for k in tree.get_children('')]
    
    # Ordena os dados
    if col in ('Estoque Atual', 'Estoque Mínimo', 'Código', 'ID'):
        data.sort(key=lambda x: int(x[0]) if str(x[0]).isdigit() else 0, reverse=reverse)
    else:
        data.sort(key=lambda x: x[0], reverse=reverse)
    
    for index, (val, k) in enumerate(data):
        tree.move(k, '', index)
    
    tree.heading(col, command=lambda: organizar_arvore(tree, col, not reverse))


def adicionar_produto_gui():
    codigo = entry_codigo.get()
    nome = entry_nome.get()
    estoque_inicial = entry_estoque.get()
    estoque_minimo = entry_estoque_minimo.get()

    if not all([codigo, nome, estoque_inicial, estoque_minimo]):
        messagebox.showerror("Erro", "Todos os campos devem ser preenchidos!")
        return
    
    try:
        estoque_inicial = int(estoque_inicial)
        estoque_minimo = int(estoque_minimo)
        
        sucesso, mensagem = adicionar_produto(codigo, nome, estoque_inicial, estoque_minimo)
        if sucesso:
            messagebox.showinfo("Sucesso", mensagem)
            limpar_campos_adicionar()
        else:
            messagebox.showerror("Erro", mensagem)
        
    except ValueError:
        messagebox.showerror("Erro", "Campos numéricos inválidos.")

def registrar_entrada_gui():
    codigo = entry_entrada_codigo.get()
    quantidade = entry_entrada_quantidade.get()
    tipo = tipo_entrada_var.get()
    responsavel = responsavel_entrada_var.get()
    
    if not all([codigo, quantidade, tipo, responsavel]):
        messagebox.showerror("Erro", "Todos os campos devem ser preenchidos!")
        return

    try:
        quantidade = int(quantidade)
        sucesso, mensagem = registrar_entrada(codigo, quantidade, tipo, responsavel)
        if sucesso:
            messagebox.showinfo("Sucesso", mensagem)
            limpar_campos_entrada()
        else:
            messagebox.showerror("Erro", mensagem)
    except ValueError:
        messagebox.showerror("Erro", "A quantidade deve ser um número válido.")

def registrar_saida_gui():
    codigo = entry_saida_codigo.get()
    quantidade = entry_saida_quantidade.get()
    cliente = entry_saida_cliente.get()
    responsavel = responsavel_saida_var.get()
    separador = separador_saida_var.get()

    if not all([codigo, quantidade, cliente, responsavel, separador]):
        messagebox.showerror("Erro", "Todos os campos devem ser preenchidos!")
        return

    try:
        quantidade = int(quantidade)
        sucesso, mensagem = registrar_saida(codigo, quantidade, cliente, responsavel, separador)
        if sucesso:
            messagebox.showinfo("Sucesso", mensagem)
            
            # Adiciona o cliente à lista e salva o arquivo
            if cliente not in CLIENTES:
                CLIENTES.add(cliente)
                salvar_clientes()
            
            limpar_campos_saida()
        else:
            messagebox.showerror("Erro", mensagem)

    except ValueError:
        messagebox.showerror("Erro", "A quantidade deve ser um número válido.")

def consultar_movimentacoes_gui():
    codigo = entry_consulta_codigo.get()
    cliente = entry_consulta_cliente.get()
    tipo_mov = tipo_mov_var.get()
    data_inicial_str = entry_data_inicial.get()
    data_final_str = entry_data_final.get()
    # Apenas passa o codigo se o campo não estiver vazio
    movimentacoes = consultar_movimentacoes(
    codigo if codigo else None, 
    cliente if cliente else None, 
    tipo_mov,
    data_inicial_str,
    data_final_str
    )
    
    for i in tree_movimentacoes.get_children():
        tree_movimentacoes.delete(i)
    
    if not movimentacoes:
        messagebox.showinfo("Consulta", "Nenhuma movimentação encontrada para o filtro selecionado.")
        return

    for item in movimentacoes:
        tree_movimentacoes.insert("", "end", values=(item['Data'], item['Código'], item['Produto'], item['Quantidade'], item['Tipo'], item['Detalhe'], item['Responsável']))

def verificar_estoque_baixo_gui():
    estoque_baixo = consultar_estoque_baixo()
    
    for i in tree_estoque_baixo.get_children():
        tree_estoque_baixo.delete(i)
    
    if not estoque_baixo:
        messagebox.showinfo("Alerta de Estoque", "Nenhum produto está com estoque abaixo do mínimo.")
        return

    for produto in estoque_baixo:
        tree_estoque_baixo.insert("", "end", values=(produto['Código'], produto['Nome'], produto['Estoque Atual'], produto['Estoque Mínimo']))

def excluir_produto_gui():
    codigo = entry_excluir_codigo.get()
    if not codigo:
        messagebox.showerror("Erro", "Por favor, insira o código do produto a ser excluído.")
        return

    confirmar = messagebox.askyesno("Confirmar Exclusão", f"Tem certeza que deseja excluir o produto com código '{codigo}'?\nEsta ação é irreversível e apagará todo o histórico do produto.")
    
    if confirmar:
        sucesso, mensagem = excluir_produto(codigo)
        if sucesso:
            messagebox.showinfo("Sucesso", mensagem)
            entry_excluir_codigo.delete(0, tk.END)
        else:
            messagebox.showerror("Erro", mensagem)

def limpar_historico_gui():
    codigo = entry_limpar_historico_codigo.get()
    if not codigo:
        messagebox.showerror("Erro", "Por favor, insira o código do produto para limpar o histórico.")
        return

    confirmar = messagebox.askyesno("Confirmar Limpeza", f"Tem certeza que deseja limpar o histórico de movimentações do produto com código '{codigo}'?\nEsta ação é irreversível.")
    
    if confirmar:
        sucesso, mensagem = limpar_historico_produto(codigo)
        if sucesso:
            messagebox.showinfo("Sucesso", mensagem)
            entry_limpar_historico_codigo.delete(0, tk.END)
            # Opcional: Atualizar a aba Histórico
            notebook.select(3) # muda para a aba Histórico
            consultar_movimentacoes_gui()
        else:
            messagebox.showerror("Erro", mensagem)

def importar_produtos_gui():
    filepath = filedialog.askopenfilename(
        title="Selecione a Planilha de Produtos",
        filetypes=[("Arquivos Excel", "*.xlsx")]
    )
    if not filepath:
        return

    sucesso, mensagem = importar_produtos_excel(filepath)
    if sucesso:
        messagebox.showinfo("Importação Concluída", mensagem)
    else:
        messagebox.showerror("Erro de Importação", mensagem)

def atualizar_lista_clientes(event):
    """
    Atualiza a lista de clientes para o autocompletar.
    Funciona tanto para a aba de Saída Única quanto para a de Saída em Lote.
    """
    carregar_clientes()
    
    # Identifica de qual Entry a função foi chamada
    if event.widget == entry_pedido_cliente:
        filtro = entry_pedido_cliente.get().lower()
        # Não usamos o OptionMenu aqui, apenas a lista de clientes para referência
        cliente_var = cliente_pedido_var
        entry = entry_pedido_cliente
    elif event.widget == entry_saida_lote_cliente: # Adicionada para o Lote
        filtro = entry_saida_lote_cliente.get().lower()
        cliente_var = cliente_saida_lote_var
        entry = entry_saida_lote_cliente
    elif event.widget == entry_pedido_cliente: # Adicionada para Pedido
        filtro = entry_pedido_cliente.get().lower()
        cliente_var = cliente_pedido_var
        entry = entry_pedido_cliente
    else:
        return # Caso seja chamada de outro lugar, ignora

    # Lógica de Autocompletar (usando o OptionMenu ao lado ou similar)
    # Como você tem um OptionMenu ao lado, vou manter apenas a atualização da lista CLIENTES.
    # O OptionMenu com *sorted(list(CLIENTES)) já tenta refletir a lista.
    pass # A lógica de autocompletar com OptionMenu/Entry é complexa. Manter o Entry/OptionMenu existentes é o suficiente por enquanto.


def limpar_campos_adicionar():
    entry_codigo.delete(0, tk.END)
    entry_nome.delete(0, tk.END)
    entry_estoque.delete(0, tk.END)
    entry_estoque_minimo.delete(0, tk.END)

def limpar_campos_entrada():
    entry_entrada_codigo.delete(0, tk.END)
    entry_entrada_quantidade.delete(0, tk.END)
    tipo_entrada_var.set(TIPOS_ENTRADA[0])
    responsavel_entrada_var.set(RESPONSAVEIS[0])

def limpar_campos_saida():
    entry_saida_codigo.delete(0, tk.END)
    entry_saida_quantidade.delete(0, tk.END)
    entry_saida_cliente.delete(0, tk.END)
    cliente_saida_var.set("")
    responsavel_saida_var.set(RESPONSAVEIS[0])
    separador_saida_var.set(SEPARADORES[0])

def carregar_inventario(termo_busca=""):
    inventario = consultar_estoque_geral()
    
    if termo_busca:
        inventario = buscar_inventario(termo_busca)

    for item in tree_inventario.get_children():
        tree_inventario.delete(item)
    
    if not inventario:
        return
        
    for item in inventario:
        estoque = item['Estoque Atual']
        reservado = item['Reservado']
        media = item['Média 45 dias'] if item['Média 45 dias'] else 0
        
        # 1. Calculamos o que realmente está na prateleira livre
        saldo_livre = estoque - reservado
        
        # 2. Lógica da Duração do Estoque
        if media > 0:
            media_diaria = media / 45
            # Cálculo: Saldo Livre / Consumo Diário
            dias_duracao = saldo_livre / media_diaria
            
            if dias_duracao < 0: dias_duracao = 0 # Evita dias negativos
            duracao_exibicao = f"{int(dias_duracao)} dias"
        else:
            duracao_exibicao = "Sem saída"
            dias_duracao = 999 # Valor alto apenas para a lógica de cor

        # 3. Lógica da Quantidade Faltante (que você já tinha)
        if estoque < media:
            qtd_faltante_exibicao = media - estoque
        else:
            qtd_faltante_exibicao = 0

        # 4. Definição de Cores (Tags) baseada no TEMPO
        tags = ()
        if dias_duracao <= 7: 
            tags = ('critico',) 
        elif dias_duracao <= 15: 
            tags = ('alerta',)
        elif dias_duracao <= 30: # Mudamos de 29 para 30 aqui
            tags = ('atencao',)  
        else: # Qualquer coisa acima de 30 cai aqui automaticamente
            tags = ('saudavel',)   

        # 5. Inserção na Treeview (Lembre-se de adicionar a coluna 'Duração' na sua Treeview)
        tree_inventario.insert("", "end", values=(
            item['Código'],
            item['Nome'],
            media,
            estoque,
            reservado,
            qtd_faltante_exibicao,
            duracao_exibicao, # <--- NOVA COLUNA AQUI
            item['Estoque Mínimo'], 
            item['Status']
        ), tags=tags)

    # Não esqueça de configurar as cores das novas tags:
    tree_inventario.tag_configure('critico', foreground='white', background='#8B0000') # Vermelho Escuro
    tree_inventario.tag_configure('alerta', foreground='black', background="#009BAC") # Azuls
    tree_inventario.tag_configure('atencao', foreground='black', background="#FFE100") # Amarelo
    tree_inventario.tag_configure('saudavel', foreground='black', background="#00A108") # Verde

def abrir_janela_producao():
    # 1. Cria a nova janela
    janela_prod = Toplevel()
    janela_prod.title("Lista de Produção Necessária (45 dias)")
    janela_prod.geometry("600x400")
    
    # 2. Busca os dados
    itens_faltantes = obter_lista_producao()
    
    if not itens_faltantes:
        messagebox.showinfo("Produção", "Tudo em dia! Nenhum item faltante para a média de 45 dias.")
        janela_prod.destroy()
        return

    # 3. Cria a tabela na janela nova
    colunas = ("Código", "Nome", "Quantidade a Produzir")
    tree_prod = ttk.Treeview(janela_prod, columns=colunas, show="headings")
    tree_prod.pack(expand=True, fill="both", padx=10, pady=10)
    
    for col in colunas:
        tree_prod.heading(col, text=col)
        tree_prod.column(col, anchor="center")
    tree_prod.column("Nome", width=300, anchor="w")

    # 4. Insere os itens
    for item in itens_faltantes:
        tree_prod.insert("", "end", values=(item['Código'], item['Nome'], item['Faltante']))

def editar_media_double_click(event):
    item_selecionado = tree_inventario.selection()
    if not item_selecionado:
        return
    
    # Pega os dados da linha clicada
    valores = tree_inventario.item(item_selecionado)['values']
    codigo = valores[0]
    nome = valores[1]
    media_atual = valores[2]

    # Abre o balãozinho de pergunta
    nova_media = simpledialog.askinteger("Editar Média", f"Nova Média 45 dias para {nome}:", initialvalue=media_atual)

    if nova_media is not None:
        # 1. Salva no Banco (Chama a função que criamos no funcoes_estoque)
        atualizar_media_db(codigo, nova_media)
        # 2. Atualiza a tela (Chama sua função que recarrega a tabela)
        carregar_inventario()

def importar_planilha_medias_xlsx():
    # 1. Seleciona o arquivo .xlsx
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecionar Planilha de Médias",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    
    if not caminho_arquivo:
        return

    try:
        df = pd.read_excel(caminho_arquivo)
    
        # Padroniza os nomes das colunas para minúsculo e sem espaços
        df.columns = [str(c).strip().lower() for c in df.columns]

        # Verifica se as colunas necessárias existem
        if 'codigo' not in df.columns or 'media' not in df.columns:
            messagebox.showerror("Erro", "A planilha deve ter as colunas 'codigo' e 'media'!")
            return

        dados_para_atualizar = []
        for _, linha in df.iterrows():
            # Pega os dados pelos nomes das colunas
            codigo = str(linha['codigo']).strip()
            media = linha['media']
        
            if pd.notna(codigo) and pd.notna(media):
                dados_para_atualizar.append((float(media), codigo))

        # 3. Manda para a função de banco que já criamos
        if dados_para_atualizar:
            sucesso, total = atualizar_medias_em_lote(dados_para_atualizar)
            if sucesso:
                messagebox.showinfo("Sucesso", f"Importação concluída!\n{total} produtos atualizados.")
                carregar_inventario() # Atualiza a tabela na tela
            else:
                messagebox.showerror("Erro", "Não foi possível atualizar o banco de dados.")
        else:
            messagebox.showwarning("Aviso", "Nenhum dado válido encontrado na planilha.")

    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao ler o arquivo Excel: {e}\nCertifique-se que o arquivo não está aberto.")

def abrir_lista_reservados_gui():
    janela_res = tk.Toplevel(root)
    janela_res.title("Detalhamento de Itens Reservados")
    janela_res.geometry("750x450")
    
    tk.Label(janela_res, text="Produtos Reservados em Pedidos Ativos", 
             font=("Arial", 14, "bold")).pack(pady=10)

    # Colunas conforme você pediu
    colunas = ("Cód. Produto", "Quantidade", "ID Pedido", "Nome do Cliente")
    tree_res = ttk.Treeview(janela_res, columns=colunas, show="headings")
    
    for col in colunas:
        tree_res.heading(col, text=col)
        tree_res.column(col, width=150, anchor="center")
    
    tree_res.pack(expand=True, fill="both", padx=10, pady=10)

    # Busca os dados processados
    dados = consultar_detalhes_reservado()
    
    if not dados:
        messagebox.showinfo("Aviso", "Não há produtos reservados em pedidos pendentes.")
        janela_res.destroy()
        return

    for linha in dados:
        tree_res.insert("", "end", values=linha)

def executar_correcao_reservas():
    sucesso, mensagem = auditar_e_corrigir_reservas()
    if sucesso:
        messagebox.showinfo("Auditoria", mensagem)
        # Atualiza a tabela de produtos para mostrar os novos números
        carregar_inventario() 
    else:
        messagebox.showerror("Erro", mensagem)

def buscar_inventario_gui():
    termo = entry_busca_inventario.get()
    carregar_inventario(termo)

def exportar_inventario_csv():
    try:
        inventario = consultar_estoque_geral()
        if not inventario:
            messagebox.showinfo("Exportar Inventário", "Nenhum produto encontrado para exportar.")
            return

        filepath = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("Arquivos CSV", "*.csv")],
            initialfile="inventario_completo.csv",
            title="Salvar Inventário como CSV"
        )

        if not filepath:
            return

        with open(filepath, 'w', newline='', encoding='cp1252') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['Código', 'Nome', 'Estoque Atual', 'Estoque Mínimo', 'Status'])
            for produto in inventario:
                writer.writerow([
                    produto['Código'],
                    produto['Nome'],
                    produto['Estoque Atual'],
                    produto['Estoque Mínimo'],
                    produto['Status']
                ])
        
        messagebox.showinfo("Exportar Inventário", f"Inventário exportado com sucesso para:\n{filepath}")

    except Exception as e:
        messagebox.showerror("Erro de Exportação", f"Ocorreu um erro ao exportar o inventário: {e}")

# --- Funções para Saída em Lote ---
def adicionar_item_saida_lote_gui():
    codigo = entry_saida_lote_codigo.get().strip()
    quantidade_str = entry_saida_lote_quantidade.get().strip()

    if not codigo or not quantidade_str:
        messagebox.showerror("Erro", "Por favor, preencha o código e a quantidade.")
        return

    if codigo in codigos_saida_lote:
        messagebox.showerror("Erro", f"O código '{codigo}' já foi adicionado à lista.")
        return

    try:
        quantidade = int(quantidade_str)
        if quantidade <= 0:
            messagebox.showerror("Erro", "A quantidade deve ser um número positivo.")
            return
    except ValueError:
        messagebox.showerror("Erro", "A quantidade deve ser um número inteiro válido.")
        return

    codigos_saida_lote.add(codigo)
    tree_saida_lote.insert("", "end", values=(codigo, quantidade))

    entry_saida_lote_codigo.delete(0, tk.END)
    entry_saida_lote_quantidade.delete(0, tk.END)
    entry_saida_lote_codigo.focus_set() # Volta o foco para o primeiro campo

def excluir_item_saida_lote_gui():
    """Exclui o item selecionado da lista de saída em lote."""
    selecionado = tree_saida_lote.selection()
    if not selecionado:
        messagebox.showerror("Erro", "Por favor, selecione um item para excluir.")
        return
    
    item_para_excluir = selecionado[0]
    valores = tree_saida_lote.item(item_para_excluir)['values']
    
    # Converte o valor para string, garantindo a correspondência de tipo
    codigo = str(valores[0]) 
    
    if codigo in codigos_saida_lote:
        codigos_saida_lote.remove(codigo) 
    
    tree_saida_lote.delete(item_para_excluir)
    messagebox.showinfo("Sucesso", f"O item '{codigo}' foi removido da lista.")

def confirmar_saida_lote():
    cliente = entry_saida_lote_cliente.get().strip() # .strip() evita espaços invisíveis
    responsavel = responsavel_saida_lote_var.get()
    separador = separador_saida_lote_var.get()
    
    if not all([cliente, responsavel, separador]):
        messagebox.showerror("Erro", "Por favor, preencha todos os campos da saída em lote.")
        return

    if cliente not in CLIENTES:
        CLIENTES.add(cliente)
        salvar_clientes()

    # 1. Captura os itens e associa ao ID ÚNICO da Treeview
    produtos_lote = []
    # Usaremos uma lista de tuplas para evitar problemas com códigos iguais na mesma lista
    mapeamento_direto = [] 
    
    for item_id in tree_saida_lote.get_children():
        values = tree_saida_lote.item(item_id)['values']
        codigo = str(values[0]).strip() # Garante que é string e sem espaços
        try:
            quantidade = int(values[1])
        except:
            continue
        
        produtos_lote.append({'codigo': codigo, 'quantidade': quantidade})
        # Guardamos o par (codigo, id_visual)
        mapeamento_direto.append((codigo, item_id))

    if not produtos_lote:
        messagebox.showerror("Erro", "Nenhum produto foi adicionado à lista.")
        return
        
    # 2. Processa o lote
    resultados = registrar_saidas_lote(produtos_lote, cliente, responsavel, separador)
    
    # 3. LIMPEZA SEGURA:
    # Vamos percorrer os resultados e o nosso mapeamento ao mesmo tempo
    for i, r in enumerate(resultados):
        if r.get('sucesso'):
            codigo_sucesso = r.get('codigo')
            # Pegamos o ID visual correspondente à posição do item processado
            id_visual = mapeamento_direto[i][1]
            
            # Remove da tela
            if tree_saida_lote.exists(id_visual):
                tree_saida_lote.delete(id_visual)
            
            # Remove do SET de controle (O MAIS IMPORTANTE)
            # Usamos discard() em vez de remove() porque o discard não gera erro se o item já sumiu
            if codigo_sucesso in codigos_saida_lote:
                codigos_saida_lote.discard(codigo_sucesso)

    # 4. Feedback
    sucessos_count = sum(1 for r in resultados if r.get('sucesso'))
    falhas_count = len(resultados) - sucessos_count

    if falhas_count > 0:
        detalhes_falhas = "Itens que permaneceram na lista:\n"
        for r in resultados:
            if not r.get('sucesso'):
                detalhes_falhas += f"- {r.get('codigo')}: {r.get('mensagem')}\n"
        messagebox.showwarning("Atenção", f"{falhas_count} falha(s) detectada(s).\n\n{detalhes_falhas}")
    
    if sucessos_count > 0:
        messagebox.showinfo("Sucesso", f"{sucessos_count} item(ns) processado(s)!")
        if falhas_count == 0:
            entry_saida_lote_cliente.delete(0, tk.END)

# --- Funções para Entrada em Lote ---
def adicionar_item_entrada_lote_gui():
    codigo = entry_entrada_lote_codigo.get().strip()
    quantidade_str = entry_entrada_lote_quantidade.get().strip()

    if not codigo or not quantidade_str:
        messagebox.showerror("Erro", "Por favor, preencha o código e a quantidade.")
        return

    if codigo in codigos_entrada_lote:
        messagebox.showerror("Erro", f"O código '{codigo}' já foi adicionado à lista.")
        return

    try:
        quantidade = int(quantidade_str)
        if quantidade <= 0:
            messagebox.showerror("Erro", "A quantidade deve ser um número positivo.")
            return
    except ValueError:
        messagebox.showerror("Erro", "A quantidade deve ser um número inteiro válido.")
        return

    codigos_entrada_lote.add(codigo)
    tree_entrada_lote.insert("", "end", values=(codigo, quantidade))

    entry_entrada_lote_codigo.delete(0, tk.END)
    entry_entrada_lote_quantidade.delete(0, tk.END)
    entry_entrada_lote_codigo.focus_set() # Volta o foco para o primeiro campo


def excluir_item_entrada_lote_gui():
    """Exclui o item selecionado da lista de entrada em lote."""
    selecionado = tree_entrada_lote.selection()
    if not selecionado:
        messagebox.showerror("Erro", "Por favor, selecione um item para excluir.")
        return
    
    item_para_excluir = selecionado[0]
    valores = tree_entrada_lote.item(item_para_excluir)['values']
    
    # Converte o valor para string para garantir que o tipo de dado corresponda ao do conjunto
    codigo = str(valores[0])
    
    if codigo in codigos_entrada_lote:
        codigos_entrada_lote.remove(codigo)
    
    tree_entrada_lote.delete(item_para_excluir)
    messagebox.showinfo("Sucesso", f"O item '{codigo}' foi removido da lista.")
    
def confirmar_entrada_lote():
    tipo = tipo_entrada_lote_var.get()
    responsavel = responsavel_entrada_lote_var.get()
    
    if not all([tipo, responsavel]):
        messagebox.showerror("Erro", "Por favor, preencha todos os campos da entrada em lote.")
        return

    produtos_lote = []
    for item in tree_entrada_lote.get_children():
        values = tree_entrada_lote.item(item)['values']
        produtos_lote.append({'codigo': values[0], 'quantidade': int(values[1])})
    
    if not produtos_lote:
        messagebox.showerror("Erro", "Nenhum produto foi adicionado à lista para registrar a entrada.")
        return
        
    resultados = registrar_entradas_lote(produtos_lote, tipo, responsavel)
    
    sucessos = sum(1 for r in resultados if r['sucesso'])
    falhas = len(resultados) - sucessos
    
    if sucessos > 0:
        messagebox.showinfo("Entrada em Lote", f"Entrada de {sucessos} produto(s) registrada com sucesso.\nFalhas: {falhas}")

        # --- Chame a nova função de refresh ---
        refresh_historico() 
        # --------------------------------------

    else:
        messagebox.showerror("Erro na Entrada em Lote", f"Nenhuma entrada foi registrada. Falhas: {falhas}")

    # Limpa a lista da entrada em lote
    for item in tree_entrada_lote.get_children():
        tree_entrada_lote.delete(item)
    codigos_entrada_lote.clear()
    
def refresh_historico():
    """Função de segurança para recarregar a Treeview do histórico."""
    try:
        # Chama a função principal de consulta (que já sabemos que existe)
        consultar_movimentacoes_gui() 
    except NameError:
        # Se a função ainda não estiver no escopo, podemos ignorar a falha de refresh
        # Ou, mais seguro, colocar o código de refresh aqui.
        pass
    except Exception as e:
        # Tratar outros erros
        print(f"Erro ao recarregar histórico: {e}")     

# --- NOVA FUNÇÃO PARA ATUALIZAR ESTOQUE MÍNIMO ---
def atualizar_estoque_minimo_gui():
    codigo = entry_estoque_minimo_att_codigo.get()
    novo_estoque = entry_estoque_minimo_novo.get()

    if not all([codigo, novo_estoque]):
        messagebox.showerror("Erro", "Preencha o código do produto e o novo estoque mínimo.")
        return

    try:
        novo_estoque = int(novo_estoque)
        if novo_estoque < 0:
            messagebox.showerror("Erro", "O estoque mínimo deve ser um número inteiro não negativo.")
            return

        sucesso, mensagem = atualizar_estoque_minimo(codigo, novo_estoque)
        if sucesso:
            messagebox.showinfo("Sucesso", mensagem)
            entry_estoque_minimo_att_codigo.delete(0, tk.END)
            entry_estoque_minimo_novo.delete(0, tk.END)
            notebook.select(4) # Muda para a aba de Inventário para o usuário ver a mudança
            carregar_inventario()
        else:
            messagebox.showerror("Erro", mensagem)

    except ValueError:
        messagebox.showerror("Erro", "O novo estoque mínimo deve ser um número inteiro válido.")

# --- NOVAS FUNÇÕES DE PEDIDOS ---

# Lista temporária para armazenar os itens do pedido antes de confirmar
itens_pedido_temp = []

def adicionar_item_pedido_gui():
    # 1. ACESSO ÀS VARIÁVEIS GLOBAIS
    # Precisamos do pedido_em_edicao_id para a verificação inteligente de estoque
    global pedido_em_edicao_id, itens_pedido_temp 
    
    codigo = entry_pedido_codigo.get().strip().upper() 
    quantidade_str = entry_pedido_quantidade.get().strip()

    if not codigo or not quantidade_str:
        messagebox.showerror("Erro", "Por favor, preencha o código e a quantidade.")
        return

    try:
        quantidade_nova = int(quantidade_str)
        if quantidade_nova <= 0:
            messagebox.showerror("Erro", "A quantidade deve ser um número positivo.")
            return
    except ValueError:
        messagebox.showerror("Erro", "A quantidade deve ser um número inteiro válido.")
        return
    
    # 2. LOCALIZAÇÃO E ACUMULO (Lógica de somar se o item já estiver na lista da tela)
    encontrado = False
    item_referencia = None
    quantidade_total_solicitada = quantidade_nova
    
    for item in itens_pedido_temp:
        if item['codigo'] == codigo:
            # Se o item já existe na lista temporária da tela, somamos o que já tinha com o novo
            quantidade_total_solicitada = item['quantidade'] + quantidade_nova
            item_referencia = item # Guarda a referência do objeto para atualizar depois
            encontrado = True
            break
            
    # 3. VERIFICAÇÃO DE PRODUTO E TRAVA CONDICIONAL (Híbrida)
    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()
    
    # Busca existência do produto
    cursor.execute("SELECT nome FROM produtos WHERE codigo = ?", (codigo,))
    produto = cursor.fetchone()
    
    # Identifica o status atual do pedido que está sendo editado
    status_pedido = "Novo" 
    if pedido_em_edicao_id:
        # Usei ID_Pedido conforme sua função verificar_estoque
        cursor.execute("SELECT Status FROM pedidos WHERE ID_Pedido = ?", (pedido_em_edicao_id,))
        res_status = cursor.fetchone()
        if res_status:
            status_pedido = res_status[0]
    
    conn.close()

    if not produto:
        messagebox.showerror("Erro", f"Produto com código {codigo} não encontrado.")
        return 

    # --- A MÁGICA DA TRAVA CONDICIONAL ---
    # Se o pedido já é Pendente, Separado ou Expedição, a trava é ATIVADA
    if status_pedido in ['Pendente', 'Separado', 'Expedição']:
        sucesso, mensagem, saldo = verificar_estoque(
            codigo, 
            quantidade_total_solicitada, 
            pedido_id_edicao=pedido_em_edicao_id
        )
        
        if not sucesso:
            messagebox.showerror("Bloqueio de Estoque", 
                f"Status do Pedido: {status_pedido}\n\n"
                f"Não é permitido adicionar itens sem saldo nesta fase.\n\n"
                f"{mensagem}")
            return 
    # --------------------------------------
    
    nome_produto = produto[0]
    
    # 4. ATUALIZAÇÃO DA LISTA INTERNA
    if encontrado:
        # Se já existia, atualizamos a quantidade total no item que já estava na lista
        item_referencia['quantidade'] = quantidade_total_solicitada
    else:
        # Se é um código novo na lista atual, adicionamos um novo dicionário
        itens_pedido_temp.append({'codigo': codigo, 'quantidade': quantidade_total_solicitada})

    # 5. ATUALIZAÇÃO DA TELA
    atualizar_treeview_itens_pedido()

    # Limpeza de campos para o próximo item
    entry_pedido_codigo.delete(0, tk.END)
    entry_pedido_quantidade.delete(0, tk.END)
    entry_pedido_codigo.focus_set()

def atualizar_treeview_itens_pedido():
    """Recarrega a Treeview de itens do pedido com a lista temporária."""
    for item in tree_itens_pedido.get_children():
        tree_itens_pedido.delete(item)
    
    for item in itens_pedido_temp:
        tree_itens_pedido.insert("", "end", values=(item['codigo'], item['quantidade']))


def excluir_item_pedido_gui():
    global itens_pedido_temp
    
    selecionado = tree_itens_pedido.selection()
    if not selecionado:
        messagebox.showerror("Erro", "Selecione um item.")
        return
    
    item_para_excluir = selecionado[0]
    valores = tree_itens_pedido.item(item_para_excluir)['values']
    codigo = str(valores[0])
    
    # APENAS remove da lista temporária (a memória da tela)
    itens_pedido_temp = [item for item in itens_pedido_temp if str(item['codigo']) != codigo]
    
    atualizar_treeview_itens_pedido()
    # A atualização REAL no banco de dados só vai acontecer quando você clicar em SALVAR.

def mostrar_detalhes_pedido(event):
    """Abre uma janela modal para exibir os itens detalhados de um pedido selecionado."""
    
    # 1. Obtenção do Treeview de origem (mantido)
    if event.widget == tree_pedidos_pendentes:
        tree_alvo = tree_pedidos_pendentes
    elif event.widget == tree_pedidos_separados:
        tree_alvo = tree_pedidos_separados
    elif event.widget == tree_pedidos_expedicao:
        tree_alvo = tree_pedidos_expedicao
    elif event.widget == tree_historico:
        tree_alvo = tree_historico         
    else:
        return

    item_selecionado = tree_alvo.focus()
    if not item_selecionado:
        return

    try:
        # Pega e converte o ID do pedido
        pedido_id_valor_str = tree_alvo.item(item_selecionado, 'values')[0]
        pedido_id = int(pedido_id_valor_str)
    except (IndexError, ValueError):
        return
        
        # --- ADICIONE ESTE PRINT DE DEBBUG ---
    print(f"DEBUG FINAL: ID Lido da Treeview para Consulta: {pedido_id}")
    # ------------------------------------

    pedido = consultar_pedido_por_id(pedido_id) 

    # --- CORREÇÃO AQUI: Chave 'itens' toda em minúsculo ---
    if pedido and 'itens' in pedido:
        try:
            itens_lista = json.loads(pedido['itens']) # <-- Chave 'itens'
        except json.JSONDecodeError:
            messagebox.showerror("Erro de Leitura", "Erro ao ler dados de itens do pedido (JSON corrompido).")
            return

        # --- Criação da Janela Modal (o restante do código pode ser mantido, mas verifique o título) ---
        janela_detalhes = Toplevel(root) 
        janela_detalhes.title(f"Itens do Pedido #{pedido_id}")
        janela_detalhes.transient(root) 
        janela_detalhes.geometry("500x300")
        
        # --- CORREÇÃO NO TÍTULO: Chaves 'cliente' e 'solicitante' em minúsculo ---
        ttk.Label(janela_detalhes, text=f"Pedido: {pedido_id} | Cliente: {pedido['cliente']} | Solicitante: {pedido['solicitante']}", font=('Arial', 10, 'bold')).pack(pady=10)
        # ----------------------------------------------------------------------

        # Configuração do Treeview para os Itens
        tree_itens = ttk.Treeview(janela_detalhes, columns=("Código", "Quantidade"), show="headings")
        tree_itens.heading("Código", text="Código do Produto")
        tree_itens.heading("Quantidade", text="Quantidade Solicitada")
        tree_itens.column("Código", width=150, anchor="center")
        tree_itens.column("Quantidade", width=100, anchor="center")
        tree_itens.pack(padx=10, pady=5, expand=True, fill="both")

        # Inserir os itens (mantido)
        for item in itens_lista:
            tree_itens.insert('', 'end', values=(item.get('codigo', 'N/A'), item.get('quantidade', 'N/A')))
    
    else:
        # Esta mensagem só aparecerá se o pedido ID 13 realmente não existir ou se a coluna 'itens' 
        # for NULL no banco (o que é improvável se o pedido foi criado).
        messagebox.showerror("Erro", f"Pedido {pedido_id} não encontrado ou sem itens no banco de dados.")

def abrir_aba_pedidos_dia():
    # ... código de criação da aba/frame ...
    
    # Botões de Ação
    btn_validar = ttk.Button(frame_acoes, text="VALIDAR TUDO", command=executar_validacao_gui)
    btn_validar.pack(side="left", padx=5)
    
    btn_enviar = ttk.Button(frame_acoes, text="ENVIAR PARA PENDENTES", command=promover_pedido_selecionado)
    btn_enviar.pack(side="left", padx=5)

    # Configuração de Cores na Treeview
    tree_pedidos_dia.tag_configure('falta', foreground='red')
    tree_pedidos_dia.tag_configure('ok', foreground='green')

def executar_validacao_gui():
    res = validar_estoque_rascunhos() # Chama a função que corrigimos
    
    if res["faltantes"] == 0:
        messagebox.showinfo("Validação", f"Tudo OK! Todos os {res['total']} pedidos rascunhos podem ser atendidos.")
    else:
        msg = f"Atenção: {res['faltantes']} pedidos possuem itens sem estoque:\n\n"
        msg += "\n".join(res["logs"][:15]) # Mostra até 15 erros
        messagebox.showwarning("Alerta de Produção", msg)

def promover_pedido_selecionado(event=None):
    # 1. Verifica qual item está selecionado na Treeview
    selecao = tree_pedidos_dia.selection()
    
    if not selecao:
        messagebox.showwarning("Aviso", "Por favor, selecione um pedido na lista para enviar.")
        return

    # 2. Pega o ID do pedido (geralmente a primeira coluna dos valores)
    item_id = tree_pedidos_dia.item(selecao)["values"][0]
    cliente = tree_pedidos_dia.item(selecao)["values"][2]

    # 3. Pergunta para confirmar (segurança nunca é demais!)
    confirmar = messagebox.askyesno("Confirmar Envio", 
                                    f"Deseja enviar o Pedido #{item_id} ({cliente}) para separação?\n\n"
                                    "Isso irá RESERVAR os produtos no estoque.")
    
    if confirmar:
        try:
            # 4. CHAMA A NOVA FUNÇÃO COM LÓGICA DE CORTE
            # Substituímos 'promover_pedido_para_pendente' pela nova função:
            sucesso, msg = promover_pedido_com_corte_total(item_id)
            
            if sucesso:
                # O 'msg' aqui agora trará o relatório de itens cortados, se houver!
                messagebox.showinfo("Resultado da Triagem", msg)
                
                # 5. Atualiza as telas
                carregar_pedidos_dia() 
                carregar_pedidos()
                carregar_historico_faltantes()
                if 'carregar_inventario' in globals():
                    carregar_inventario() 
            else:
                messagebox.showerror("Erro na Triagem", msg)
                
        except Exception as e:
            messagebox.showerror("Erro Crítico", f"Erro ao promover pedido: {e}")

def carregar_pedidos_dia():
    # 1. Limpa a tabela
    tree_pedidos_dia.delete(*tree_pedidos_dia.get_children())
    
    # 2. Busca os rascunhos
    pedidos = buscar_pedidos_por_status("Rascunho") 
    
    # 3. Pega o estoque disponível real (Atual - Reservado)
    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()
    cursor.execute("SELECT codigo, (estoque_atual - reservado) FROM produtos")
    estoque_real = {str(row[0]).upper(): row[1] for row in cursor.fetchall()}
    conn.close()
    
    for p in pedidos:
        id_ped, data, cliente, solicitante, itens_json = p
        
        try:
            itens = json.loads(itens_json)
            estoque_ok = True
            
            for item in itens:
                # Tentamos pegar o código de várias formas para não dar erro de digitação
                cod = str(item.get('codigo') or item.get('Código') or "").upper()
                qtd = int(item.get('quantidade') or item.get('Quantidade') or 0)
                
                disponivel = estoque_real.get(cod, 0)
                
                # DEBUG: Se quiser ver no terminal o que está acontecendo:
                # print(f"Pedido {id_ped} | Item {cod} | Precisa: {qtd} | Tem: {disponivel}")
                
                if qtd > disponivel:
                    estoque_ok = False
                    break 
            
            # 4. Define a TAG
            tag = "ok" if estoque_ok else "falta"
            
            # Insere na Treeview
            tree_pedidos_dia.insert("", "end", values=(id_ped, data, cliente, solicitante), tags=(tag,))
            
        except Exception as e:
            print(f"Erro ao processar pedido {id_ped}: {e}")

    # 5. IMPORTANTE: Garante que as cores existam na Treeview
    tree_pedidos_dia.tag_configure('falta', foreground='white', background='red') # Vermelho vivo
    tree_pedidos_dia.tag_configure('ok', foreground='green')            

def visualizar_itens_rascunho():
    selecao = tree_pedidos_dia.selection()
    if not selecao:
        return

    item_valores = tree_pedidos_dia.item(selecao)["values"]
    pedido_id = item_valores[0]
    cliente = item_valores[2]

    # 1. Busca os itens e o estoque atual para comparar
    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()
    cursor.execute("SELECT itens FROM pedidos WHERE id_pedido = ?", (pedido_id,))
    itens_json = cursor.fetchone()[0]
    
    cursor.execute("SELECT codigo, (estoque_atual - reservado) FROM produtos")
    estoque_real = {row[0]: row[1] for row in cursor.fetchall()}
    conn.close()

    itens = json.loads(itens_json)

    # 2. Cria a Janela Pop-up
    janela_view = tk.Toplevel()
    janela_view.title(f"Itens do Pedido #{pedido_id} - {cliente}")
    janela_view.geometry("600x400")
    janela_view.grab_set() # Bloqueia a janela de trás até fechar essa

    # Cabeçalho na janelinha
    tk.Label(janela_view, text=f"Detalhes do Pedido: {pedido_id}", font=("Arial", 12, "bold")).pack(pady=10)

    # Tabela interna da janelinha
    colunas = ("Código", "Qtd Solicitada", "Disponível", "Status")
    tree_view = ttk.Treeview(janela_view, columns=colunas, show="headings")
    for col in colunas:
        tree_view.heading(col, text=col)
        tree_view.column(col, width=120, anchor="center")
    
    tree_view.pack(expand=True, fill="both", padx=10, pady=10)

    # Configura cores para a janelinha
    tree_view.tag_configure('erro', foreground='red')
    tree_view.tag_configure('ok', foreground='green')

    # 3. Alimenta a janelinha com os itens
    for i in itens:
        cod = i.get('codigo') or i.get('Código')
        qtd = int(i.get('quantidade') or i.get('Quantidade') or 0)
        disponivel = estoque_real.get(cod, 0)
        
        if qtd > disponivel:
            status = "FALTA ESTOQUE"
            tag = 'erro'
        else:
            status = "OK"
            tag = 'ok'
            
        tree_view.insert("", "end", values=(cod, qtd, disponivel, status), tags=(tag,))

    # Botão para fechar
    ttk.Button(janela_view, text="Fechar", command=janela_view.destroy).pack(pady=10)

def acao_excluir_rascunho(event=None):
    selecao = tree_pedidos_dia.selection()
    if not selecao:
        messagebox.showwarning("Aviso", "Selecione um rascunho para excluir.")
        return
    
    # Pega o ID e o Nome do Cliente para confirmar
    valores = tree_pedidos_dia.item(selecao)["values"]
    id_ped = valores[0]
    cliente = valores[2]
    
    confirmar = messagebox.askyesno("Confirmar Exclusão", f"Tem certeza que deseja excluir o rascunho #{id_ped} do cliente {cliente}?\nEsta ação não pode ser desfeita.")
    
    if confirmar:
        sucesso, msg = excluir_rascunho_db(id_ped)
        if sucesso:
            messagebox.showinfo("Sucesso", msg)
            carregar_pedidos_dia() # Atualiza a lista na tela
        else:
            messagebox.showerror("Erro", msg)

def carregar_pedido_para_edicao():
    
    global pedido_em_edicao_id, itens_pedido_temp, entry_pedido_cliente, solicitante_pedido_var, ABA_CRIAR_PEDIDO
    
    import json
    import tkinter as tk
    from tkinter import messagebox
    
    # 1. VERIFICA QUAL TREEVIEW TEM UM ITEM SELECIONADO
    # Cria um dicionário com todas as Treeviews que podem ser editadas
    treeviews_para_edicao = {
        'Rascunho': tree_pedidos_dia,
        'Pendente': tree_pedidos_pendentes,
        'Separado': tree_pedidos_separados,
        'Expedição': tree_pedidos_expedicao
    }
    
    selecionado = None
    tree_origem = None
    
    # Itera sobre todas as Treeviews para encontrar a seleção
    for status, tree in treeviews_para_edicao.items():
        if tree.selection():
            selecionado = tree.selection()
            tree_origem = tree # Guarda qual Treeview foi usada
            break 
            
    if not selecionado:
        messagebox.showwarning("Aviso", "Selecione um pedido nas listas 'Pendentes', 'Separados' ou 'Expedição' para editar.")
        return

    # O ID do pedido
    item_id = selecionado[0]
    # Usa a Treeview de origem para obter o ID
    pedido_id = tree_origem.item(item_id, 'values')[0] 
    
    # 2. Busca os dados do pedido no banco de dados
    dados_pedido = consultar_pedido_por_id(pedido_id) 
    
    if not dados_pedido:
        messagebox.showerror("Erro", f"Não foi possível carregar o pedido ID {pedido_id}.")
        return

    # 3. ATUALIZAÇÃO DA VALIDAÇÃO DE STATUS
    status_atual = dados_pedido['status']
    
    # Status permitidos: Rascunho, Pendente, Separado, Expedição
    status_permitidos = ['Rascunho','Pendente', 'Separado', 'Expedição']

    # Se o status ATUAL NÃO estiver na lista de permitidos, bloqueia a edição
    if status_atual is not None and status_atual not in status_permitidos:
        messagebox.showwarning("Aviso", f"Apenas pedidos com status {', '.join(status_permitidos)} podem ser editados. Status atual: {status_atual}.")
        return
    # O código continua se o status for permitido.

    # 4. Popula a interface de criação de pedido
    itens_pedido_temp.clear()
    
    # Limpa o campo do Cliente (Entry)
    entry_pedido_cliente.delete(0, tk.END)
    
    # LIMPA E PREENCHE O CAMPO SOLICITANTE USANDO A STRINGVAR (Combobox)
    solicitante_pedido_var.set('')
    
    # Insere os dados
    entry_pedido_cliente.insert(0, dados_pedido['cliente'])
    solicitante_pedido_var.set(dados_pedido['solicitante']) 
    
    # Carrega os itens na lista temporária
    itens_lista = json.loads(dados_pedido['itens'])
    itens_pedido_temp.extend(itens_lista)
    
    # 5. Define o modo de edição e atualiza a treeview
    pedido_em_edicao_id = pedido_id # Define o modo de edição!
    atualizar_treeview_itens_pedido()
    
    messagebox.showinfo("Edição", f"Pedido ID {pedido_id} carregado para edição. Altere os itens e clique em 'Salvar Pedido'.")
    notebook.select(5) # Mantenha o índice correto para a ABA_CRIAR_PEDIDO

def registrar_novo_pedido_ou_atualizar_gui():
    try:
        global solicitante_pedido_var, itens_pedido_temp, pedido_em_edicao_id, urgente_var, entry_pedido_cliente, separador_saida_lote_var

        cliente = entry_pedido_cliente.get().strip()
        solicitante = solicitante_pedido_var.get() 
        separador = separador_saida_lote_var.get()
        urgente = urgente_var.get() 

        if not all([cliente, solicitante, separador]):
            messagebox.showerror("Erro", "Preencha o cliente, o Solicitante e o Responsável/Separador.")
            return
            
        if not itens_pedido_temp:
            messagebox.showerror("Erro", "Adicione pelo menos um item ao pedido.")
            return

        produtos_lote = itens_pedido_temp
        import json
        itens_json_string = json.dumps(produtos_lote)

        if pedido_em_edicao_id:
            # Busca o status atual no banco antes de mandar atualizar
            import sqlite3
            conn = sqlite3.connect('estoque.db')
            cursor = conn.cursor()
            cursor.execute("SELECT Status FROM pedidos WHERE ID_Pedido = ?", (pedido_em_edicao_id,))
            status_atual = cursor.fetchone()[0]
            conn.close()

            # AGORA SIM: Enviando os 5 argumentos que a função espera
            sucesso, mensagem = atualizar_pedido(
                pedido_em_edicao_id,
                itens_json_string,
                cliente, 
                solicitante,
                status_atual # <--- O 5º argumento aqui
            )
        else:
            sucesso, mensagem = registrar_pedido(cliente, solicitante, separador, produtos_lote, urgente)
            
        if sucesso:
            messagebox.showinfo("Sucesso", mensagem)
            # ... resto do seu código de limpeza (limpar campos, resetar variáveis) ...
            itens_pedido_temp = []
            entry_pedido_cliente.delete(0, tk.END)
            urgente_var.set("Não")
            pedido_em_edicao_id = None 
            atualizar_treeview_itens_pedido()
            carregar_pedidos_dia()
            carregar_pedidos()
            # Se você já achou o nome da função do inventário, chame ela aqui também:
            carregar_inventario() 

            notebook.select(tab_pedidos_separados)
            notebook_pedidos.select(tab_pedidos_dia)

    except Exception as e:
        messagebox.showerror("Erro Crítico", f"Erro: {str(e)}")

def carregar_pedidos(status_para_recarregar=None):
    """Carrega os pedidos pendentes e separados nas suas respectivas Treeviews."""
    
    # === Lógica para o Histórico (Concluído) ===
    # Se a função for chamada para recarregar o histórico (com ou sem filtros)
    if status_para_recarregar == 'Concluído':
        
        # 1. Limpa a Treeview do Histórico
        for item in tree_historico.get_children():
            tree_historico.delete(item)
            
        # 2. Obtém os valores dos campos de filtro.
        # Estas variáveis (entry_filtro_historico_cliente, etc.) devem ser definidas
        # como 'global' onde você criou os widgets na interface.
        filtro_cliente_val = entry_filtro_historico_cliente.get()
        filtro_data_val = entry_filtro_historico_data.get()
        
        # Otimização: se o valor for uma string vazia ou apenas espaços, envia None
        filtro_cliente_enviar = filtro_cliente_val.strip() if filtro_cliente_val else None
        filtro_data_enviar = filtro_data_val.strip() if filtro_data_val else None
        
        # 3. Consulta apenas o histórico, aplicando os filtros
        pedidos_historico_filtrados = consultar_pedidos(
            status='Concluído', 
            filtro_cliente=filtro_cliente_enviar, # Enviando None se for vazio
            filtro_data_finalizacao=filtro_data_enviar # Enviando None se for vazio
        )
        
        # 4. Insere pedidos filtrados no Histórico
        for pedido in pedidos_historico_filtrados:
            detalhes = f"Cliente: {pedido['Cliente']} / Resp: {pedido['Separador']} / Solicitante: {pedido['Responsavel_Expedicao']}"
            tree_historico.insert("", "end", iid=pedido['ID_Pedido'], values=(
                pedido['ID_Pedido'], 
                pedido['Data_Finalizacao'] if 'Data_Finalizacao' in pedido else pedido['Data_Criacao'],
                detalhes, 
                pedido['Itens']
            ))
            
        # Não processa o restante da função, apenas atualiza o histórico.
        return
    
    # 1. Ajuste na consulta 
    pedidos_pendentes = consultar_pedidos(status='Pendente') 
    pedidos_separados = consultar_pedidos(status='Separado')
    pedidos_expedicao = consultar_pedidos(status='Expedição')
    pedidos_historico = consultar_pedidos(status='Concluído')
    
    # --- CORREÇÃO DE ERRO: IMPLEMENTAÇÃO DA LIMPEZA ---
    # Limpa a Treeview de pedidos PENDENTES
    for item in tree_pedidos_pendentes.get_children():
        tree_pedidos_pendentes.delete(item)
        
    # Limpa a Treeview de pedidos SEPARADOS
    for item in tree_pedidos_separados.get_children():
        tree_pedidos_separados.delete(item)
    # Limpa a Treeview de pedidos na EXPEDIÇÃO
    for item in tree_pedidos_expedicao.get_children():
        tree_pedidos_expedicao.delete(item) 

    for item in tree_historico.get_children():
        tree_historico.delete(item)        
    
    # Insere pedidos pendentes
    for pedido in pedidos_pendentes:
    # Mude para chaves MAIÚSCULAS para carregar a lista:
        detalhes = f"Cliente: {pedido['Cliente']} / Resp: {pedido['Solicitante']}"
    
        tree_pedidos_pendentes.insert("", "end", iid=pedido['ID_Pedido'], values=(
        pedido['ID_Pedido'], # ID
        pedido['Data_Criacao'], # Data
        detalhes, 
        pedido['Itens'] # Itens do Pedido
    ))
    
    # Insere pedidos separados (prontos para expedição)
    for pedido in pedidos_separados:
    # Mude para chaves MAIÚSCULAS
        detalhes = f"Cliente: {pedido['Cliente']} / Resp: {pedido['Separador']}" 
    
        tree_pedidos_separados.insert("", "end", iid=pedido['ID_Pedido'], values=(
        pedido['ID_Pedido'], 
        pedido['Data_Separacao'], 
        detalhes, 
        pedido['Itens']
        ))
        
    for pedido in pedidos_expedicao:
    # 1. Coleta os 3 nomes:
        cliente = pedido['Cliente']
    # Resp: Separador da mercadoria
        separador_nome = pedido.get('Separador', 'N/A') 
    # Solicitante: A pessoa que autorizou o envio (Responsavel_Expedicao)
        solicitante_expedicao = pedido.get('Responsavel_Expedicao', 'N/A') 

    # 2. Constrói a nova string de detalhes com as 3 informações
        detalhes = f"Cliente: {cliente} / Resp. Sep: {separador_nome} / Solicitante: {solicitante_expedicao}"
    
    # 3. Insere na Treeview da Expedição
        tree_pedidos_expedicao.insert("", "end", iid=pedido['ID_Pedido'], values=(
            pedido['ID_Pedido'], 
            pedido['Data_Expedicao'], 
            detalhes,              # <--- A NOVA STRING DETALHADA
            pedido['Itens']
        ))
    for pedido in pedidos_historico:
        detalhes = f"Cliente: {pedido['Cliente']} / Resp: {pedido['Separador']} / Solicitante: {pedido['Responsavel_Expedicao']}"
        tree_historico.insert("", "end", iid=pedido['ID_Pedido'], values=(
            pedido['ID_Pedido'], 
            pedido['Data_Finalizacao'] if 'Data_Finalizacao' in pedido else pedido['Data_Criacao'],
            detalhes, 
            pedido['Itens']
        ))       

def carregar_historico_faltantes():
    # 1. Limpa a Treeview
    for i in tree_faltantes.get_children():
        tree_faltantes.delete(i)

    # 2. Pega os valores e trata espaços em branco
    filtro = entry_busca_faltantes.get().strip()
    data_ini = entry_data_inicio_faltantes.get().strip()
    data_fim = entry_data_fim_faltantes.get().strip()

    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()

    # Query base
    query = """
        SELECT 
            datetime(data_corte, 'localtime'), -- Converte a data salva para o horário local
            id_pedido, 
            cliente, 
            codigo_produto, 
            nome_produto, 
            quantidade_faltante 
        FROM historico_faltantes 
        WHERE 1=1
    """
    params = []

    # Filtro de texto
    if filtro:
        query += " AND (cliente LIKE ? OR codigo_produto LIKE ? OR nome_produto LIKE ?)"
        term = f"%{filtro}%"
        params.extend([term, term, term])

    # Filtro de Período - SÓ ENTRA SE AMBOS TIVEREM ALGO ESCRITO
    if data_ini and data_fim:
        query += " AND DATE(data_corte) BETWEEN ? AND ?"
        params.extend([data_ini, data_fim])

    query += " ORDER BY data_corte DESC"

    try:
        cursor.execute(query, params)
        dados = cursor.fetchall()
        
        # DEBUG: Isso vai mostrar no seu terminal se o banco retornou algo
        print(f"DEBUG: Foram encontrados {len(dados)} registros de faltantes.")

        for linha in dados:
            # Formata a data para o padrão brasileiro (DD/MM/AAAA HH:MM)
            # A data vem do SQLite como 'YYYY-MM-DD HH:MM:SS'
            data_original = linha[0]
            try:
                data_dt = data_original.split(" ")[0] # Pega YYYY-MM-DD
                ano, mes, dia = data_dt.split("-")
                hora = data_original.split(" ")[1][:5] # Pega HH:MM
                data_exibir = f"{dia}/{mes}/{ano} {hora}"
            except:
                data_exibir = data_original

            # Cria uma nova tupla com a data formatada para exibir na tela
            valores_exibir = (data_exibir, linha[1], linha[2], linha[3], linha[4], linha[5])
            tree_faltantes.insert("", "end", values=valores_exibir)

            atualizar_soma_faltantes()
            
    except Exception as e:
        print(f"Erro ao carregar histórico: {e}")
    finally:
        conn.close()

def mover_pedido_gui():
    """Move um pedido de 'Pendente' para 'Separado', solicitando Separador e Solicitante."""
    selecionado = tree_pedidos_pendentes.selection()
    if not selecionado:
        messagebox.showerror("Erro", "Selecione um pedido pendente para mover.")
        return

    pedido_id = tree_pedidos_pendentes.item(selecionado[0], 'values')[0]

    top = tk.Toplevel(root)
    top.title("Confirmar Separação")
    
    frame_campos = ttk.Frame(top)
    frame_campos.pack(padx=10, pady=10)
    
    tk.Label(frame_campos, text=f"Pedido ID: {pedido_id}", font=('Arial', 12, 'bold')).grid(row=0, columnspan=2, pady=5)
    
    # --- CAMPO 1: SEPARADOR (Busca da tabela de RESPONSAVEIS) ---
    tk.Label(frame_campos, text="Separador (Você):").grid(row=1, column=0, sticky="e", padx=5, pady=5)
    
    lista_resps = obter_nomes_aux("RESPONSAVEL") # <--- DINÂMICO
    separador_var_temp = tk.StringVar(top)
    separador_var_temp.set(lista_resps[0])
    
    ttk.OptionMenu(frame_campos, separador_var_temp, lista_resps[0], *lista_resps).grid(row=1, column=1, sticky="ew", padx=5, pady=5)
    
    # --- CAMPO 2: SOLICITANTE (Busca da tabela de SOLICITANTES) ---
    tk.Label(frame_campos, text="Solicitante:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
    
    lista_solics = obter_nomes_aux("SOLICITANTE") # <--- DINÂMICO
    solicitante_var_temp = tk.StringVar(top)
    
    combo_solicitante = ttk.Combobox(
        frame_campos, 
        textvariable=solicitante_var_temp, 
        values=lista_solics, 
        state='readonly'
    )
    combo_solicitante.grid(row=2, column=1, sticky="ew", padx=5, pady=5)
    if lista_solics:
        combo_solicitante.set(lista_solics[0]) 

    def confirmar_movimento():
        separador_nome = separador_var_temp.get()
        solicitante_nome = solicitante_var_temp.get()
        
        if not separador_nome or not solicitante_nome or solicitante_nome == "SELECIONE":
            messagebox.showerror("Erro", "Selecione o Separador e o Solicitante.")
            return
            
        sucesso, mensagem = separar_pedido(pedido_id, separador_nome, solicitante_nome)
        
        if sucesso:
            messagebox.showinfo("Sucesso", mensagem)
            carregar_pedidos()
            top.destroy()
        else:
            messagebox.showerror("Erro", mensagem)
            top.destroy()

    ttk.Button(top, text="Confirmar Separação", command=confirmar_movimento).pack(pady=10)
    
def mover_expedicao_gui():
    item_selecionado = tree_pedidos_separados.focus()
    if not item_selecionado:
        messagebox.showwarning("Atenção", "Selecione um pedido separado para mover para a Expedição.")
        return

    pedido_id = tree_pedidos_separados.item(item_selecionado, 'values')[0]

    janela_responsavel = Toplevel(root)
    janela_responsavel.title("Responsável pela Expedição")
    janela_responsavel.grab_set()

    ttk.Label(janela_responsavel, text=f"Quem solicitou o envio do Pedido {pedido_id}?").pack(padx=10, pady=10)
    
    responsavel_var = tk.StringVar()
    
    # --- BUSCA DINÂMICA DE SOLICITANTES ---
    lista_solics_exp = obter_nomes_aux("SOLICITANTE") 
    
    combo_responsavel = ttk.Combobox(
        janela_responsavel, 
        textvariable=responsavel_var, 
        values=lista_solics_exp, 
        state='readonly', 
        width=30
    )
    combo_responsavel.pack(padx=10, pady=5)
    
    if lista_solics_exp:
        combo_responsavel.set(lista_solics_exp[0]) 

    def confirmar_expedicao():
        responsavel_envio = responsavel_var.get()
        
        # Validação para garantir que não enviou o valor padrão "SELECIONE"
        if not responsavel_envio or responsavel_envio == "SELECIONE":
            messagebox.showerror("Erro", "Selecione um nome válido da lista.")
            return

        sucesso, mensagem = mover_pedido_para_expedicao(pedido_id, responsavel_envio)

        if sucesso:
            messagebox.showinfo("Sucesso", mensagem)
            carregar_pedidos()
            janela_responsavel.destroy()
        else:
            messagebox.showerror("Erro", mensagem)

    ttk.Button(janela_responsavel, text="Confirmar Envio", command=confirmar_expedicao).pack(pady=10)

def atualizar_todos_os_menus():
    """Recarrega os nomes do banco de dados em todos os menus do sistema."""
    # 1. Busca os nomes atualizados
    resps = obter_nomes_aux("RESPONSAVEL")
    solics = obter_nomes_aux("SOLICITANTE")

    # 2. Atualiza os Comboboxes (Aba Pedidos e Pop-ups)
    # Nota: Como alguns combos são locais, o ideal é atualizar os globais aqui
    try:
        combo_solicitante['values'] = solics
    except: pass

    # 3. Atualiza os OptionMenus (Aba Lote)
    # O OptionMenu é mais complexo, o ideal é limpar e reinserir
    try:
        # Exemplo para o menu de Entrada em Lote
        menu = menu_resp_lote["menu"]
        menu.delete(0, "end")
        for nome in resps:
            menu.add_command(label=nome, command=tk._setit(responsavel_entrada_lote_var, nome))
    except: pass
    
    print("Menus atualizados com sucesso!")

def excluir_pedido_gui():
    """Exclui um pedido da lista de 'Separados'."""
    selecionado = tree_pedidos_separados.selection()
    if not selecionado:
        messagebox.showerror("Erro", "Selecione um pedido separado para excluir/finalizar.")
        return

    pedido_id = int(selecionado[0])
    
    confirmar = messagebox.askyesno("Confirmar Finalização", f"Tem certeza que deseja FINALIZAR (excluir) o pedido ID '{pedido_id}' da expedição?\nEsta ação é IRREVERSÍVEL.")
    
    if confirmar:
        sucesso, mensagem = excluir_pedido(pedido_id)
        if sucesso:
            messagebox.showinfo("Sucesso", mensagem)
            carregar_pedidos() # Recarrega a lista
        else:
            messagebox.showerror("Erro", mensagem)

def excluir_e_estornar_pedido():
    try:
        # Assumindo que você pega o pedido selecionado
        item_selecionado = tree_pedidos_separados.focus() 
        valores = tree_pedidos_separados.item(item_selecionado, 'values')
        pedido_id = valores[0]  # Assumindo que o ID do Pedido está na primeira coluna
        
        # Você deve ter uma variável global ou um campo para o nome do usuário logado
        responsavel = responsavel_logado_var.get() 

        if not pedido_id:
             messagebox.showerror("Erro", "Selecione um pedido para estornar.")
             return

        # Confirmação de segurança
        if messagebox.askyesno("Confirmação de Estorno", f"Tem certeza que deseja estornar o Pedido {pedido_id}? Os itens retornarão ao estoque."):
            
            sucesso, mensagem = estornar_pedido(pedido_id, responsavel)
            
            if sucesso:
                messagebox.showinfo("Sucesso", mensagem)
                
                # RECARREGAR A TELA: Recarrega as tabelas afetadas
                carregar_pedidos() # Recarrega a tabela de pedidos
                consultar_movimentacoes_gui() # Recarrega o Histórico
            else:
                messagebox.showerror("Erro", mensagem)

    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao tentar excluir/estornar: {e}")


def abrir_modal_estorno():
    # 1. Obter o ID do Pedido SELECIONADO (Ação Crítica)
    try:
        item_selecionado = tree_pedidos_separados.focus() 
        if not item_selecionado:
            messagebox.showerror("Erro", "Selecione um pedido para estornar.")
            return

        valores = tree_pedidos_separados.item(item_selecionado, 'values')
        pedido_id = valores[0]  # Assumindo que o ID do Pedido está na primeira coluna
        
    except Exception:
        messagebox.showerror("Erro", "Falha ao obter ID do pedido selecionado.")
        return

    # 2. Criar a Janela Toplevel (o Modal)
    modal = tk.Toplevel(root) # 'root' é a sua janela principal
    modal.title(f"Confirmar Estorno Pedido {pedido_id}")
    modal.transient(root) # Faz a janela ficar acima da principal
    modal.grab_set()      # Impede interação com a janela principal

    # 3. Componentes do Modal
    
    # Variável para o nome do responsável
    responsavel_var = tk.StringVar(modal)
    responsavel_var.set(RESPONSAVEIS[0]) # Valor inicial

    # Rótulo e Campo de Seleção (OptionMenu)
    ttk.Label(modal, text="Responsável pelo Estorno:").pack(pady=5, padx=10)
    
    ttk.OptionMenu(modal, responsavel_var, responsavel_var.get(), *RESPONSAVEIS).pack(pady=5, padx=10)

    # 4. Botão de Confirmação
    def confirmar_estorno_modal():
        responsavel_estorno = responsavel_var.get()
        
        # Chama a função principal de estorno com o valor coletado
        sucesso, mensagem = estornar_pedido(pedido_id, responsavel_estorno)
        
        if sucesso:
            messagebox.showinfo("Sucesso", mensagem)
            
            # RECARREGAR A TELA PRINCIPAL
            carregar_pedidos() 
            consultar_movimentacoes_gui() 
            modal.destroy() # Fecha o modal
        else:
            messagebox.showerror("Erro", mensagem)
    
    ttk.Button(modal, text="CONFIRMAR ESTORNO", command=confirmar_estorno_modal).pack(pady=15, padx=10)
    
    # Adicionar um botão Cancelar para boas práticas
    modal.protocol("WM_DELETE_WINDOW", modal.destroy)
    modal.bind('<Return>', lambda e: confirmar_estorno_modal())

    # Foca no modal (importante para acessibilidade)
    root.wait_window(modal)

def finalizar_pedido_gui():
    """Finaliza o pedido selecionado na aba Expedição."""
    
    # 1. Tenta pegar o item selecionado na treeview de Expedição
    item_selecionado = tree_pedidos_expedicao.focus()
    if not item_selecionado:
        messagebox.showwarning("Atenção", "Selecione um pedido na lista de Pedidos na Expedição para finalizar.")
        return

    # 2. Pega o ID
    # Assumindo que o ID está na primeira coluna
    pedido_id = tree_pedidos_expedicao.item(item_selecionado, 'values')[0]

    # 3. Confirmação
    if messagebox.askyesno("Confirmar Finalização", f"Deseja realmente FINALIZAR o Pedido {pedido_id}?"):
        
        # Chama a função do backend
        # Lembre-se de importar finalizar_pedido do funcoes_estoque.py
        sucesso, mensagem = finalizar_pedido(pedido_id)

        if sucesso:
            messagebox.showinfo("Sucesso", mensagem)
            # 4. Recarrega as listas para remover o pedido da Treeview
            carregar_pedidos() 
        else:
            messagebox.showerror("Erro", mensagem)


def ordenar_treeview(tree, col, reverse):
    """
    Ordena a Treeview pelo conteúdo da coluna clicada, com tratamento especial para datas.
    """
    
    # Função auxiliar para conversão de valor.
    def converte_valor(val, col_id):
        # ATENÇÃO: Verifique se '#2' é o ID da sua coluna "Finalizado em"
        if col_id == '#2':
            # 1. Tenta converter DD/MM/AAAA para um objeto datetime ou um formato AAAA-MM-DD
            try:
                # O Treeview exibe "DD/MM/AAAA HH:MM:SS" (ou similar)
                data_str = val.split(' ')[0] # Pega apenas a parte DD/MM/AAAA
                return datetime.strptime(data_str, '%d/%m/%Y')
            except ValueError:
                return datetime.min # Retorna a menor data possível para valores inválidos/vazios
        # Para todas as outras colunas (ID, Cliente, etc.), ordena como string ou número.
        return val


    # 1. Obtém os dados da coluna, aplicando a conversão se for a coluna de data
    # O 'k' é o ID interno da linha
    data = [(converte_valor(tree.set(k, col), col), k) for k in tree.get_children('')]

    # 2. Ordena os dados (agora usando objetos datetime para a coluna #2)
    data.sort(reverse=reverse)

    # 3. Reorganiza os itens na Treeview
    for index, (val, k) in enumerate(data):
        tree.move(k, '', index)

    # 4. Alterna a direção da ordenação
    tree.heading(col, command=lambda: ordenar_treeview(tree, col, not reverse))
    
def limpar_filtros_historico_pedidos():
    global entry_filtro_historico_cliente
    global entry_filtro_historico_data
    
    # 1. Limpa o campo de Cliente
    entry_filtro_historico_cliente.delete(0, tk.END)
    
    # 2. Limpa o campo de Data
    # 2a. Remove o valor atual (se houver)
    entry_filtro_historico_data.delete(0, tk.END)
    
    # 2b. Coloca uma string vazia, garantindo que o .get() retornará "".
    entry_filtro_historico_data.insert(0, "")
    
    # 3. Recarrega a Treeview
    carregar_pedidos('Concluído')

# ----------------- Configuração da Janela Principal -----------------
class TelaLogin:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Login - Sistema de Estoque")
        self.root.geometry("350x450")
        self.root.configure(bg="#1e1e1e")
        self.role_logada = None
        
        # Centralizar a janela de login
        largura = 350
        altura = 450
        largura_tela = self.root.winfo_screenwidth()
        altura_tela = self.root.winfo_screenheight()
        pos_x = (largura_tela // 2) - (largura // 2)
        pos_y = (altura_tela // 2) - (altura // 2)
        self.root.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")

        # Interface Visual do Login
        tk.Label(self.root, text="ACESSO AO SISTEMA", fg="white", bg="#1e1e1e", font=("Arial", 16, "bold")).pack(pady=30)
        
        tk.Label(self.root, text="Usuário:", fg="#aaaaaa", bg="#1e1e1e").pack(anchor="w", padx=40)
        self.ent_user = tk.Entry(self.root, font=("Arial", 12), bg="#2d2d2d", fg="white", insertbackground="white")
        self.ent_user.pack(fill="x", padx=40, pady=5)
        
        tk.Label(self.root, text="Senha:", fg="#aaaaaa", bg="#1e1e1e").pack(anchor="w", padx=40, pady=(10,0))
        self.ent_pass = tk.Entry(self.root, show="*", font=("Arial", 12), bg="#2d2d2d", fg="white", insertbackground="white")
        self.ent_pass.pack(fill="x", padx=40, pady=5)

        btn_login = tk.Button(self.root, text="ENTRAR", bg="#009BAC", fg="white", font=("Arial", 11, "bold"), 
                              command=self.verificar_acesso, cursor="hand2")
        btn_login.pack(fill="x", padx=40, pady=40)
        
        self.root.bind('<Return>', lambda event: self.verificar_acesso())

    def verificar_acesso(self):
        user = self.ent_user.get()
        pw = self.ent_pass.get()
        role = validar_acesso(user, pw) 
        
        if role:
            self.role_logada = role
            self.root.destroy()
        else:
            messagebox.showerror("Erro", "Usuário ou senha incorretos!")

    def rodar(self):
        self.root.mainloop()
        return self.role_logada

def montar_sistema_principal(root):
    global tree_pedidos_pendentes, tree_inventario, tree_pedidos_expedicao, tree_pedidos_separados, tree_faltantes, tree_historico, tree_pedidos_dia, tree_itens_pedido, tree_entrada_lote, tree_saida_lote
    global tree_estoque_baixo, tree_movimentacoes, atualizar_soma_faltantes
    global notebook, notebook_avancadas, notebook_pedidos, frame_acoes, registrar_saida, registrar_entrada, logo, assinatura_logo, USUARIO_ROLE
    global entry_codigo, entry_busca_faltantes, entry_data_inicio_faltantes, entry_data_fim_faltantes, entry_pedido_codigo, entry_pedido_quantidade, entry_pedido_codigo, entry_estoque_minimo_att_codigo
    global entry_saida_lote_codigo, entry_saida_lote_quantidade, entry_saida_lote_codigo, entry_busca_inventario, entry_estoque_minimo_novo, entry_entrada_lote_codigo, entry_entrada_lote_quantidade, entry_entrada_lote_codigo, entry_saida_lote_cliente
    global entry_saida_codigo, entry_saida_quantidade, entry_saida_cliente, entry_entrada_codigo, entry_entrada_quantidade, entry_nome, entry_estoque, entry_estoque_minimo, entry_limpar_historico_codigo
    global entry_excluir_codigo, entry_consulta_codigo, entry_consulta_cliente, entry_data_inicial, entry_data_final, entry_saida_codigo
    global entry_saida_quantidade, entry_saida_cliente
    global tab_pedidos_separados, tab_pedidos_dia, separador_saida_lote_var,  responsavel_entrada_lote_var, tipo_entrada_lote_var, responsavel_logado_var, menu_resp_lote, combo_solicitante, responsavel_saida_lote_var
    global cliente_saida_var, responsavel_saida_var, separador_saida_var, tipo_entrada_var, responsavel_entrada_var, cliente_pedido_var, tipo_mov_var, cliente_saida_lote_var

    # --- LIMPEZA SEGURA (Sem Loop) ---
    # Se for uma tupla/lista, pega o primeiro item. Fazemos isso 2 vezes para garantir.
    if isinstance(USUARIO_ROLE, (tuple, list)) and len(USUARIO_ROLE) > 0:
        USUARIO_ROLE = USUARIO_ROLE
    
    if isinstance(USUARIO_ROLE, (tuple, list)) and len(USUARIO_ROLE) > 0:
        USUARIO_ROLE = USUARIO_ROLE

    # Transforma em string limpa e remove espaços/parênteses extras
    USUARIO_ROLE = str(USUARIO_ROLE).strip("()', ") 
    # ---------------------------------

    print(f"DEBUG FINAL (DEFINITIVO): '{USUARIO_ROLE}'")

    logo_path = resource_path(os.path.join("arquivos_adicionais", "logo.png"))
    assinatura_path = resource_path(os.path.join("arquivos_adicionais", "assinatura.png"))

    try:
        # Carregamento do Logo
        if os.path.exists(logo_path):
            original_image = Image.open(logo_path)
            resized_image = original_image.resize((200, 50), Image.Resampling.LANCZOS)
            logo = ImageTk.PhotoImage(resized_image)

        # Carregamento da Assinatura
        if os.path.exists(assinatura_path):
            original_assinatura = Image.open(assinatura_path)
            resized_assinatura = original_assinatura.resize((150, 50), Image.Resampling.LANCZOS)
            assinatura_logo = ImageTk.PhotoImage(resized_assinatura)

    except Exception as e:
        print(f"AVISO: Falha ao carregar imagens: {e}")
        logo = None
        assinatura_logo = None

    notebook = ttk.Notebook(root)
    notebook.pack(pady=10, expand=True, fill="both")

    style = ttk.Style(root)
    style.configure('TLabel', font=('Arial', 12))
    style.configure('TEntry', font=('Arial', 12))
    style.configure('TButton', font=('Arial', 12, 'bold'))
    style.configure('TMenubutton', font=('Arial', 12))
    style.configure('Treeview', font=('Arial', 11))
    style.configure('Treeview.Heading', font=('Arial', 12, 'bold'))
    style.configure('TNotebook.Tab', font=('Arial', 12))

    # Adiciona o novo estilo para linhas com estoque baixo
    style.configure('Treeview', rowheight=25)
    style.map('Treeview', foreground=[('!disabled', 'white')])
    style.configure('abaixo_minimo.Treeview', foreground='red')

    if USUARIO_ROLE == 'admin':
        print("Acesso de Admin confirmado! Criando abas extras...")
        # --- Aba: Adicionar Produto ---
        tab_adicionar = ttk.Frame(notebook)
        notebook.add(tab_adicionar, text="Adicionar Produto")

        tk.Label(tab_adicionar, text="Adicionar Novo Produto", font=("Arial", 18, "bold")).pack(pady=5)


        frame_adicionar = ttk.Frame(tab_adicionar, padding=(30, 15))
        frame_adicionar.pack(expand=True)

        ttk.Label(frame_adicionar, text="Código:").grid(row=0, column=0, sticky="e", pady=8, padx=5)
        entry_codigo = ttk.Entry(frame_adicionar, width=30)
        entry_codigo.grid(row=0, column=1, pady=8, padx=5)
        entry_codigo.bind("<Return>", lambda event: entry_nome.focus_set())

        ttk.Label(frame_adicionar, text="Nome:").grid(row=1, column=0, sticky="e", pady=8, padx=5)
        entry_nome = ttk.Entry(frame_adicionar, width=30)
        entry_nome.grid(row=1, column=1, pady=8, padx=5)
        entry_nome.bind("<Return>", lambda event: entry_estoque.focus_set())

        ttk.Label(frame_adicionar, text="Estoque Inicial:").grid(row=2, column=0, sticky="e", pady=8, padx=5)
        entry_estoque = ttk.Entry(frame_adicionar, width=30)
        entry_estoque.grid(row=2, column=1, pady=8, padx=5)
        entry_estoque.bind("<Return>", lambda event: entry_estoque_minimo.focus_set())

        ttk.Label(frame_adicionar, text="Estoque Mínimo:").grid(row=3, column=0, sticky="e", pady=8, padx=5)
        entry_estoque_minimo = ttk.Entry(frame_adicionar, width=30)
        entry_estoque_minimo.grid(row=3, column=1, pady=8, padx=5)
        entry_estoque_minimo.bind("<Return>", lambda event: adicionar_produto_gui())

        ttk.Button(frame_adicionar, text="Adicionar Produto", command=adicionar_produto_gui).grid(row=4, columnspan=2, pady=20)

        # --- Aba: Opções Avançadas ---
        tab_avancadas = ttk.Frame(notebook)
        notebook.add(tab_avancadas, text="Avançadas")

        # Sub-Notebook para organizar as sub-abas (Entrada Lote, Saída Lote, Manutenção)
        notebook_avancadas = ttk.Notebook(tab_avancadas)
        notebook_avancadas.pack(pady=10, expand=True, fill="both", padx=10)


        # --- Sub-Aba: Entrada em Lote ---
        tab_entrada_lote = ttk.Frame(notebook_avancadas)
        notebook_avancadas.add(tab_entrada_lote, text="Entrada em Lote")
        tk.Label(tab_entrada_lote, text="Registrar Entrada de Múltiplos Itens", font=("Arial", 16, "bold")).pack(pady=10)

        frame_entrada_lote_controles = ttk.Frame(tab_entrada_lote, padding=(10, 10))
        frame_entrada_lote_controles.pack(fill="x", padx=10)

        ttk.Label(frame_entrada_lote_controles, text="Cód:").grid(row=0, column=0, sticky="e", padx=5)
        entry_entrada_lote_codigo = ttk.Entry(frame_entrada_lote_controles, width=15)
        entry_entrada_lote_codigo.grid(row=0, column=1, padx=5)
        entry_entrada_lote_codigo.bind("<Return>", lambda event: entry_entrada_lote_quantidade.focus_set())

        ttk.Label(frame_entrada_lote_controles, text="Qtd:").grid(row=0, column=2, sticky="e", padx=5)
        entry_entrada_lote_quantidade = ttk.Entry(frame_entrada_lote_controles, width=10)
        entry_entrada_lote_quantidade.grid(row=0, column=3, padx=5)
        entry_entrada_lote_quantidade.bind("<Return>", lambda event: adicionar_item_entrada_lote_gui())

        ttk.Button(frame_entrada_lote_controles, text="Adicionar Item", command=adicionar_item_entrada_lote_gui).grid(row=0, column=4, padx=10)

        # Treeview para Entrada em Lote
        columns_entrada_lote = ("Código", "Quantidade")
        tree_entrada_lote = ttk.Treeview(tab_entrada_lote, columns=columns_entrada_lote, show="headings")
        tree_entrada_lote.pack(pady=10, expand=True, fill="both", padx=10)
        for col in columns_entrada_lote:
            tree_entrada_lote.heading(col, text=col)
            tree_entrada_lote.column(col, width=100, anchor="center")

        # Controles de Confirmação da Entrada em Lote
        frame_entrada_lote_confirmar = ttk.Frame(tab_entrada_lote, padding=(10, 10))
        frame_entrada_lote_confirmar.pack(fill="x", padx=10)

        ttk.Label(frame_entrada_lote_confirmar, text="Tipo:").grid(row=0, column=0, sticky="e", padx=5)
        tipo_entrada_lote_var = tk.StringVar(root)
        tipo_entrada_lote_var.set(TIPOS_ENTRADA[0])
        menu_resp_lote = ttk.OptionMenu(frame_entrada_lote_confirmar, tipo_entrada_lote_var, TIPOS_ENTRADA[0], *TIPOS_ENTRADA).grid(row=0, column=1, padx=5, sticky="ew")

        ttk.Label(frame_entrada_lote_confirmar, text="Resp:").grid(row=0, column=2, sticky="e", padx=5)
        lista_resps_entrada = obter_nomes_aux("RESPONSAVEL")
        responsavel_entrada_lote_var = tk.StringVar(root)
        responsavel_entrada_lote_var.set(lista_resps_entrada[0])
        menu_resp_lote = ttk.OptionMenu(frame_entrada_lote_confirmar, responsavel_entrada_lote_var, lista_resps_entrada[0], *lista_resps_entrada)
        menu_resp_lote.grid(row=0, column=3, padx=5, sticky="ew")

        ttk.Button(frame_entrada_lote_confirmar, text="Excluir Item", command=excluir_item_entrada_lote_gui).grid(row=1, column=0, columnspan=2, pady=10)
        ttk.Button(frame_entrada_lote_confirmar, text="CONFIRMAR ENTRADA", command=confirmar_entrada_lote).grid(row=1, column=2, columnspan=2, pady=10)


        # --- Sub-Aba: Saída em Lote ---
        tab_saida_lote = ttk.Frame(notebook_avancadas)
        notebook_avancadas.add(tab_saida_lote, text="Saída em Lote")
        tk.Label(tab_saida_lote, text="Registrar Saída de Múltiplos Itens", font=("Arial", 16, "bold")).pack(pady=10)

        frame_saida_lote_controles = ttk.Frame(tab_saida_lote, padding=(10, 10))
        frame_saida_lote_controles.pack(fill="x", padx=10)

        ttk.Label(frame_saida_lote_controles, text="Cód:").grid(row=0, column=0, sticky="e", padx=5)
        entry_saida_lote_codigo = ttk.Entry(frame_saida_lote_controles, width=15)
        entry_saida_lote_codigo.grid(row=0, column=1, padx=5)
        entry_saida_lote_codigo.bind("<Return>", lambda event: entry_saida_lote_quantidade.focus_set())

        ttk.Label(frame_saida_lote_controles, text="Qtd:").grid(row=0, column=2, sticky="e", padx=5)
        entry_saida_lote_quantidade = ttk.Entry(frame_saida_lote_controles, width=10)
        entry_saida_lote_quantidade.grid(row=0, column=3, padx=5)
        entry_saida_lote_quantidade.bind("<Return>", lambda event: adicionar_item_saida_lote_gui())

        ttk.Button(frame_saida_lote_controles, text="Adicionar Item", command=adicionar_item_saida_lote_gui).grid(row=0, column=4, padx=10)

        # Treeview para Saída em Lote
        columns_saida_lote = ("Código", "Quantidade")
        tree_saida_lote = ttk.Treeview(tab_saida_lote, columns=columns_saida_lote, show="headings")
        tree_saida_lote.pack(pady=10, expand=True, fill="both", padx=10)
        for col in columns_saida_lote:
            tree_saida_lote.heading(col, text=col)
            tree_saida_lote.column(col, width=100, anchor="center")

        # Controles de Confirmação da Saída em Lote
        frame_saida_lote_confirmar = ttk.Frame(tab_saida_lote, padding=(10, 10))
        frame_saida_lote_confirmar.pack(fill="x", padx=10)

        ttk.Label(frame_saida_lote_confirmar, text="Cliente:").grid(row=0, column=0, sticky="e", padx=5)
        cliente_saida_lote_var = tk.StringVar(root)
        entry_saida_lote_cliente = ttk.Entry(frame_saida_lote_confirmar, width=20, textvariable=cliente_saida_lote_var)
        entry_saida_lote_cliente.grid(row=0, column=1, padx=5, sticky="ew")
        # Não adiciono o OptionMenu para simplificar a interface em lote.
        entry_saida_lote_cliente.bind("<KeyRelease>", atualizar_lista_clientes)

        ttk.Label(frame_saida_lote_confirmar, text="Responsável:").grid(row=1, column=0, sticky="e", padx=5)
        lista_resps_saida = obter_nomes_aux("RESPONSAVEL")
        responsavel_saida_lote_var = tk.StringVar(root)
        responsavel_saida_lote_var.set(lista_resps_saida[0])
        menu_resp_saida_lote = ttk.OptionMenu(frame_saida_lote_confirmar, responsavel_saida_lote_var, lista_resps_saida[0], *lista_resps_saida)
        menu_resp_saida_lote.grid(row=1, column=1, padx=5, sticky="ew")

        ttk.Label(frame_saida_lote_confirmar, text="Solicitante:").grid(row=2, column=0, sticky="e", padx=5)
        lista_solics_saida = obter_nomes_aux("SOLICITANTE")
        separador_saida_lote_var = tk.StringVar(root)
        separador_saida_lote_var.set(lista_solics_saida[0])

        menu_solic_saida_lote = ttk.OptionMenu(frame_saida_lote_confirmar, separador_saida_lote_var, lista_solics_saida[0], *lista_solics_saida)
        menu_solic_saida_lote.grid(row=2, column=1, padx=5, sticky="ew")

        ttk.Button(frame_saida_lote_confirmar, text="Excluir Item", command=excluir_item_saida_lote_gui).grid(row=3, column=0, pady=10)
        ttk.Button(frame_saida_lote_confirmar, text="CONFIRMAR SAÍDA", command=confirmar_saida_lote).grid(row=3, column=1, pady=10)


        # --- Sub-Aba: Manutenção ---
        tab_manutencao = ttk.Frame(notebook_avancadas)
        notebook_avancadas.add(tab_manutencao, text="Manutenção")

        tk.Label(tab_manutencao, text="Manutenção de Dados", font=("Arial", 16, "bold")).pack(pady=10)

        # Excluir Produto
        frame_excluir = ttk.LabelFrame(tab_manutencao, text="Excluir Produto (IRREVERSÍVEL)", padding=(10, 10))
        frame_excluir.pack(pady=10, padx=20, fill="x")
        ttk.Label(frame_excluir, text="Código do Produto:").pack(side="left", padx=5)
        entry_excluir_codigo = ttk.Entry(frame_excluir, width=20)
        entry_excluir_codigo.pack(side="left", padx=5)
        ttk.Button(frame_excluir, text="Excluir Produto", command=excluir_produto_gui).pack(side="left", padx=10)

        # Limpar Histórico
        frame_limpar_historico = ttk.LabelFrame(tab_manutencao, text="Limpar Histórico (IRREVERSÍVEL)", padding=(10, 10))
        frame_limpar_historico.pack(pady=10, padx=20, fill="x")
        ttk.Label(frame_limpar_historico, text="Código do Produto:").pack(side="left", padx=5)
        entry_limpar_historico_codigo = ttk.Entry(frame_limpar_historico, width=20)
        entry_limpar_historico_codigo.pack(side="left", padx=5)
        ttk.Button(frame_limpar_historico, text="Limpar Histórico", command=limpar_historico_gui).pack(side="left", padx=10)

        # Importar Produtos
        frame_importar = ttk.LabelFrame(tab_manutencao, text="Importação de Produtos", padding=(10, 10))
        frame_importar.pack(pady=10, padx=20, fill="x")
        ttk.Label(frame_importar, text="Importar produtos de planilha Excel (.xlsx):").pack(side="left", padx=5)
        ttk.Button(frame_importar, text="Selecionar Arquivo", command=importar_produtos_gui).pack(side="left", padx=10)

        # --- GERENCIAR NOMES (ADICIONAR/REMOVER RESPONSÁVEIS) ---
        frame_auxiliar = ttk.LabelFrame(tab_manutencao, text="Gerenciar Solicitantes / Responsáveis", padding=(10, 10))
        frame_auxiliar.pack(pady=10, padx=20, fill="x")

        # Controles de Entrada
        frame_inputs_aux = ttk.Frame(frame_auxiliar)
        frame_inputs_aux.pack(fill="x")

        ttk.Label(frame_inputs_aux, text="Tipo:").grid(row=0, column=0, padx=5, pady=5)
        combo_tipo_aux = ttk.Combobox(frame_inputs_aux, values=["SOLICITANTE", "RESPONSAVEL", "SEPARADOR"], state="readonly", width=15)
        combo_tipo_aux.grid(row=0, column=1, padx=5, pady=5)
        combo_tipo_aux.set("SOLICITANTE")

        ttk.Label(frame_inputs_aux, text="Nome:").grid(row=0, column=2, padx=5, pady=5)
        entry_nome_aux = ttk.Entry(frame_inputs_aux, width=25)
        entry_nome_aux.grid(row=0, column=3, padx=5, pady=5)

        # Lista para exibir os nomes cadastrados
        lista_nomes_aux = tk.Listbox(frame_auxiliar, height=4)
        lista_nomes_aux.pack(side="left", fill="both", expand=True, padx=(0, 10), pady=5)

        # Funções internas para os botões
        def atualizar_lista_interface():
            tipo = combo_tipo_aux.get()
            lista_nomes_aux.delete(0, 'end')
            # Certifique-se de que esta função está no seu arquivo de funções
            for n in listar_cadastros_aux(tipo):
                lista_nomes_aux.insert('end', n)

        def add_aux_gui():
            nome = entry_nome_aux.get().strip()
            if nome:
                if adicionar_cadastro_aux(combo_tipo_aux.get(), nome):
                    entry_nome_aux.delete(0, 'end')
                    atualizar_lista_interface() # Atualiza a listinha da própria manutenção
                    
                    # --- O TRUQUE PARA ATUALIZAR OS MENUS DE LOTE NA HORA ---
                    atualizar_menus_dinamicos() 
                    
                    messagebox.showinfo("Sucesso", f"'{nome}' cadastrado! Os menus já foram atualizados.")
                else:
                    messagebox.showerror("Erro", "Este nome já existe.")

        def del_aux_gui():
            try:
                selecionado = lista_nomes_aux.get(lista_nomes_aux.curselection())
                if messagebox.askyesno("Confirmar", f"Deseja remover '{selecionado}'?"):
                    remover_cadastro_aux(combo_tipo_aux.get(), selecionado)
                    atualizar_lista_interface()
                    atualizar_menus_dinamicos() 
            except tk.TclError:
                messagebox.showwarning("Aviso", "Selecione um nome na lista para excluir.")

        def atualizar_menus_dinamicos():
            """Recarrega os nomes do banco e força os OptionMenus a se atualizarem."""
            novos_resps = obter_nomes_aux("RESPONSAVEL")
            novos_solics = obter_nomes_aux("SOLICITANTE")

            # --- ATUALIZAÇÃO DA ABA SEPARAR PEDIDO ---
            
            # 1. Responsável (OptionMenu)
            if 'menu_resp_separar_pedido' in globals() and menu_resp_separar_pedido is not None:
                menu_ped = menu_resp_separar_pedido["menu"]
                menu_ped.delete(0, "end")
                for nome in novos_resps:
                    menu_ped.add_command(label=nome, command=tk._setit(responsavel_pedido_var, nome))

            # 2. Solicitante (Combobox) - MUITO MAIS SIMPLES!
            if 'combo_solicitante_pedido' in globals() and combo_solicitante_pedido is not None:
                combo_solicitante_pedido['values'] = novos_solics
                if novos_solics:
                    combo_solicitante_pedido.set(novos_solics[0])

            # 1. Atualiza o menu de Responsável na ENTRADA EM LOTE
            menu_ent = menu_resp_lote["menu"]
            menu_ent.delete(0, "end")
            for nome in novos_resps:
                menu_ent.add_command(label=nome, command=tk._setit(responsavel_entrada_lote_var, nome))

            # 2. Atualiza o menu de Responsável na SAÍDA EM LOTE
            menu_sai_resp = menu_resp_saida_lote["menu"] # Certifique-se de dar esse nome ao widget
            menu_sai_resp.delete(0, "end")
            for nome in novos_resps:
                menu_sai_resp.add_command(label=nome, command=tk._setit(responsavel_saida_lote_var, nome))

            # 3. Atualiza o menu de Solicitante na SAÍDA EM LOTE
            menu_sai_solic = menu_solic_saida_lote["menu"]
            menu_sai_solic.delete(0, "end")
            for nome in novos_solics:
                menu_sai_solic.add_command(label=nome, command=tk._setit(separador_saida_lote_var, nome))

        def atualizar_soma_faltantes():
            total = 0
            itens = tree_faltantes.get_children()

            for item in itens:
                valores = tree_faltantes.item(item, 'values')
                try:
                    # Pegamos o valor da coluna 5
                    valor_texto = valores[5]
                    # Convertemos para float (caso tenha decimais) e depois int
                    qtd = float(valor_texto) 
                    total += qtd
                except Exception as e:
                    print(f"DEBUG: Erro ao somar valor '{valores[5]}': {e}")
                    continue
                    
            # Atualiza a variável que está no Label
            var_total_faltantes.set(f"{int(total)}")

        def acao_exportar_pdf():
            # 1. Pega os dados da Treeview
            itens_da_tela = []
            for item in tree_faltantes.get_children():
                itens_da_tela.append(tree_faltantes.item(item, 'values'))
            
            if not itens_da_tela:
                messagebox.showwarning("Aviso", "Não há dados na tabela para exportar.")
                return

            # 2. Abre o Modal para escolher onde salvar
            caminho_arquivo = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("Arquivos PDF", "*.pdf")],
                initialfile=f"Relatorio_Faltantes_{datetime.now().strftime('%d_%m_%Y')}.pdf",
                title="Escolha onde salvar o relatório"
            )

            # 3. Se o usuário não cancelou a janela, gera o PDF
            if caminho_arquivo:
                sucesso, erro = exportar_faltantes_consolidado_pdf(itens_da_tela, caminho_arquivo)
                if sucesso:
                    messagebox.showinfo("Sucesso", "Relatório PDF gerado com sucesso!")
                    import os
                    os.startfile(caminho_arquivo) # Abre o PDF automaticamente
                else:
                    messagebox.showerror("Erro", f"Falha ao gerar PDF: {erro}")  

        # Botões de Ação
        frame_botoes_aux = ttk.Frame(frame_auxiliar)
        frame_botoes_aux.pack(side="right", fill="y")

        ttk.Button(frame_botoes_aux, text="Adicionar", command=add_aux_gui).pack(fill="x", pady=2)
        ttk.Button(frame_botoes_aux, text="Remover", command=del_aux_gui).pack(fill="x", pady=2)

        # Evento para trocar a lista ao mudar o tipo no Combo
        combo_tipo_aux.bind("<<ComboboxSelected>>", lambda e: atualizar_lista_interface())

        # Carrega a lista inicial
        atualizar_lista_interface()

        # --- Aba: Histórico de Movimentação ---
        tab_historico = ttk.Frame(notebook)
        notebook.add(tab_historico, text="Histórico")
        tk.Label(tab_historico, text="Histórico de Movimentações", font=("Arial", 18, "bold")).pack(pady=15)
        frame_filtro = ttk.Frame(tab_historico, padding=(20, 10))
        frame_filtro.pack(pady=10)

        # Novo menu suspenso para o tipo de movimentação
        ttk.Label(frame_filtro, text="Tipo de Mov.:").pack(side="left", padx=5)
        tipo_mov_var = tk.StringVar(root)
        tipo_mov_var.set("Ambas")
        option_tipo_mov = ttk.OptionMenu(frame_filtro, tipo_mov_var, "Ambas", "Entrada", "Saída", "Ambas")
        option_tipo_mov.pack(side="left", padx=5)

        ttk.Label(frame_filtro, text="Código:").pack(side="left", padx=5)
        entry_consulta_codigo = ttk.Entry(frame_filtro, width=15)
        entry_consulta_codigo.pack(side="left", padx=5)
        entry_consulta_codigo.bind("<Return>", lambda event: entry_consulta_cliente.focus_set())

        ttk.Label(frame_filtro, text="Cliente:").pack(side="left", padx=5)
        entry_consulta_cliente = ttk.Entry(frame_filtro, width=15)
        entry_consulta_cliente.pack(side="left", padx=5)
        entry_consulta_cliente.bind("<Return>", lambda event: consultar_movimentacoes_gui())

        # Campo Data Inicial (Substituído de ttk.Entry para DateEntry)
        ttk.Label(frame_filtro, text="Data Inicial:").pack(side=tk.LEFT, padx=5)
        entry_data_inicial = DateEntry(
            frame_filtro, 
            width=12, 
            date_pattern='dd/mm/yyyy', # Garante que o formato seja DD/MM/AAAA
            locale='pt_BR'
        )
        entry_data_inicial.pack(side=tk.LEFT, padx=5)

        # Campo Data Final (Substituído de ttk.Entry para DateEntry)
        ttk.Label(frame_filtro, text="Data Final:").pack(side=tk.LEFT, padx=5)
        entry_data_final = DateEntry(
            frame_filtro, 
            width=12, 
            date_pattern='dd/mm/yyyy', # Garante que o formato seja DD/MM/AAAA
            locale='pt_BR'
        )
        entry_data_final.pack(side=tk.LEFT, padx=5)

        ttk.Button(frame_filtro, text="Consultar", command=consultar_movimentacoes_gui).pack(side="left", padx=5)

        columns = ("Data", "Código", "Produto", "Quantidade", "Tipo", "Detalhe", "Responsável")
        tree_movimentacoes = ttk.Treeview(tab_historico, columns=columns, show="headings")
        tree_movimentacoes.pack(pady=10, expand=True, fill="both", padx=10)
        for col in columns:
            tree_movimentacoes.heading(col, text=col)
            tree_movimentacoes.column(col, width=100)
            # Alinha as colunas de Quantidade e Tipo ao centro
            if col in ("Quantidade", "Tipo", "Data", "Código", "Responsável"):
                tree_movimentacoes.column(col, anchor="center")

        tree_movimentacoes.column("Data", width=80)       # Reduzido de 150 para 120
        tree_movimentacoes.column("Código", width=20)      # Novo ajuste
        tree_movimentacoes.column("Produto", width=200)    # Ajuste para manter o nome visível
        tree_movimentacoes.column("Quantidade", width=20)  # Novo ajuste
        tree_movimentacoes.column("Tipo", width=10)        # Novo ajuste
        tree_movimentacoes.column("Detalhe", width=400)    # Reduzido de 150 para 120
        tree_movimentacoes.column("Responsável", width=20) # Reduzido de 150 para 100

    # --- Aba: Alertas & Relatórios ---
    tab_alertas = ttk.Frame(notebook)
    notebook.add(tab_alertas, text="Alertas & Relatórios")
    tk.Label(tab_alertas, text="Alertas de Estoque", font=("Arial", 18, "bold")).pack(pady=15)
    frame_alertas = ttk.Frame(tab_alertas, padding=(20, 10))
    frame_alertas.pack(pady=10)
    ttk.Button(frame_alertas, text="Verificar Estoque Baixo", command=verificar_estoque_baixo_gui).pack(pady=10)
    columns_alertas = ("Código", "Nome", "Estoque Atual", "Estoque Mínimo")
    tree_estoque_baixo = ttk.Treeview(tab_alertas, columns=columns_alertas, show="headings")
    tree_estoque_baixo.pack(pady=10, expand=True, fill="both", padx=10)
    for col in columns_alertas:
        tree_estoque_baixo.heading(col, text=col)
        # Alinha a coluna do estoque e do estoque mínimo ao centro
        if col in ("Estoque Atual", "Estoque Mínimo", "Código"):
            tree_estoque_baixo.column(col, width=50, anchor="center")
        else:
            tree_estoque_baixo.column(col, width=50)  
    tree_estoque_baixo.column("Nome", width=650)       

    # --- Aba: Inventário ---
    tab_inventario = ttk.Frame(notebook)
    notebook.add(tab_inventario, text="Inventário")
    tk.Label(tab_inventario, text="Inventário de Produtos", font=("Arial", 18, "bold")).pack(pady=15)

    # Novo frame para o campo de busca
    frame_busca_inventario = ttk.Frame(tab_inventario, padding=(5, 5))
    frame_busca_inventario.pack(pady=5, padx=10, fill="x")

    ttk.Label(frame_busca_inventario, text="Buscar por Código ou Nome:").pack(side="left", padx=(0, 5))
    entry_busca_inventario = ttk.Entry(frame_busca_inventario, width=25)
    entry_busca_inventario.pack(side="left", padx=5, fill="x", expand=True)
    entry_busca_inventario.bind("<Return>", lambda event: buscar_inventario_gui())

    ttk.Button(frame_busca_inventario, text="Buscar", command=buscar_inventario_gui).pack(side="left", padx=5)
    ttk.Button(frame_busca_inventario, text="Carregar Tudo", command=lambda: carregar_inventario("")).pack(side="left", padx=5)
    ttk.Button(frame_busca_inventario, text="Exportar CSV", command=exportar_inventario_csv).pack(side="left", padx=5)

    # 1. Define o estado (Pode usar a variável que criamos antes ou fazer direto)
    estado_producao = "normal" if USUARIO_ROLE == 'admin' else "disabled"

    # 2. Cria o botão com o bloqueio aplicado
    btn_producao = ttk.Button(
        frame_busca_inventario, 
        text="Lista de Produção", 
        command=abrir_janela_producao,
        state=estado_producao # <--- Bloqueia se não for admin
    )
    btn_producao.pack(side="left", padx=10)

    estado_acesso = "normal" if USUARIO_ROLE == 'admin' else "disabled"

    # Cria o botão usando esse estado
    ttk.Button(
        frame_busca_inventario, 
        text="Lista de Reservado", 
        command=abrir_lista_reservados_gui,
        state=estado_acesso  # Aplica o bloqueio se não for admin
    ).pack(side="left", padx=5)

    # 1. Definimos o estado uma única vez no início para facilitar
    estado_acesso = "normal" if USUARIO_ROLE == 'admin' else "disabled"

    # Frame para atualizar estoque mínimo
    frame_att_minimo = ttk.Frame(tab_inventario, padding=(5, 5))
    frame_att_minimo.pack(pady=5, padx=10, fill="x")

    # Criamos os elementos aplicando o 'state=estado_acesso'
    ttk.Label(frame_att_minimo, text="Atualizar Estoque Mínimo:", state=estado_acesso).pack(side="left", padx=(0, 5))

    entry_estoque_minimo_att_codigo = ttk.Entry(frame_att_minimo, width=10, state=estado_acesso)
    entry_estoque_minimo_att_codigo.pack(side="left", padx=5)
    entry_estoque_minimo_att_codigo.insert(0, "Código")
    entry_estoque_minimo_att_codigo.bind("<FocusIn>", lambda event: entry_estoque_minimo_att_codigo.delete(0, tk.END) if entry_estoque_minimo_att_codigo.get() == "Código" else None)
    entry_estoque_minimo_att_codigo.bind("<Return>", lambda event: entry_estoque_minimo_novo.focus_set())

    entry_estoque_minimo_novo = ttk.Entry(frame_att_minimo, width=10, state=estado_acesso)
    entry_estoque_minimo_novo.pack(side="left", padx=5)
    entry_estoque_minimo_novo.insert(0, "Novo Mín.")
    entry_estoque_minimo_novo.bind("<FocusIn>", lambda event: entry_estoque_minimo_novo.delete(0, tk.END) if entry_estoque_minimo_novo.get() == "Novo Mín." else None)
    entry_estoque_minimo_novo.bind("<Return>", lambda event: atualizar_estoque_minimo_gui())

    # Botão Atualizar
    ttk.Button(frame_att_minimo, text="Atualizar", command=atualizar_estoque_minimo_gui, state=estado_acesso).pack(side="left", padx=5)

    # Botão Importar
    btn_importar = ttk.Button(frame_att_minimo, text="IMPORTAR MÉDIAS (XLSX)", command=importar_planilha_medias_xlsx, state=estado_acesso)
    btn_importar.pack(side="left", padx=5)

    # Botão de Corrigir reservados
    btn_auditoria = ttk.Button(frame_att_minimo, text="🔄 Corrigir Reservas", command=executar_correcao_reservas, state=estado_acesso)
    btn_auditoria.pack(side="left", pady=5)

    # --- BLOCO DE LEGENDA ---
    frame_legenda = tk.Frame(tab_inventario) # Certifique-se que o nome da aba está correto
    frame_legenda.pack(fill="x", padx=10, pady=5)

    tk.Label(frame_legenda, text="Legenda de Duração:", font=("Arial", 9, "bold")).pack(side="left", padx=5)

    # Crítico
    lbl_critico = tk.Label(frame_legenda, text=" CRÍTICO (Até 7 dias) ", bg="#8B0000", fg="white", font=("Arial", 8, "bold"))
    lbl_critico.pack(side="left", padx=5)

    # Alerta
    lbl_alerta = tk.Label(frame_legenda, text=" ALERTA (7 a 15 dias) ", bg="#009BAC", fg="black", font=("Arial", 8, "bold"))
    lbl_alerta.pack(side="left", padx=5)

    # Atenção
    lbl_minimo = tk.Label(frame_legenda, text=" ATENÇÃO (16 a 29 dias) ", bg="#FFE100", fg="white", font=("Arial", 8, "bold"))
    lbl_minimo.pack(side="left", padx=5)

    # Saudável
    lbl_normal = tk.Label(frame_legenda, text=" SAUDÁVEL (30 ou mais dias) ", bg="#00A108", fg="black", relief="sunken", font=("Arial", 8, "bold"))
    lbl_normal.pack(side="left", padx=5)

    # Treeview do Inventário
    columns_inventario = ("Código", "Nome", "Média (45d)", "Estoque Atual", "Reservado", "Quantidade Faltante (Produção)", "Duração (Dias)", "Estoque Mínimo", "Status")

    tree_inventario = ttk.Treeview(tab_inventario, columns=columns_inventario, show="headings")
    tree_inventario.pack(pady=10, expand=True, fill="both", padx=10)

    tree_inventario.heading("Duração (Dias)", text="Duração (Dias)")
    tree_inventario.column("Duração (Dias)", width=100, anchor="center")

    for col in columns_inventario:
        tree_inventario.heading(col, text=col, command=lambda c=col: organizar_arvore(tree_inventario, c, False))
        # Ajuste de largura: Nome maior, números centralizados
        if col == "Nome":
            tree_inventario.column(col, width=250, anchor="w")
        else:
            tree_inventario.column(col, width=120, anchor="center")

    # VINCULAR O DUPLO CLIQUE
    tree_inventario.bind("<Double-1>", editar_media_double_click)

    if USUARIO_ROLE == 'admin':
        # --------------------------------------------------------
        #               NOVAS ABAS DE PEDIDOS
        # --------------------------------------------------------

        # --- Aba: Separar Pedido (Criação de Pedidos) ---
        tab_pedido = ttk.Frame(notebook)
        notebook.add(tab_pedido, text="Separar Pedido")
        tk.Label(tab_pedido, text="Criação e Separação de Pedidos", font=("Arial", 18, "bold")).pack(pady=10)

        frame_pedido_superior = ttk.Frame(tab_pedido, padding=(10, 10))
        frame_pedido_superior.pack(fill="x", padx=10)

        # Controles do Pedido (Cliente, Responsável, Solicitante)
        frame_pedido_info = ttk.LabelFrame(frame_pedido_superior, text="Dados do Pedido", padding=(10, 10))
        frame_pedido_info.pack(side="left", padx=10, fill="y")

        global menu_resp_separar_pedido, combo_solicitante_pedido, entry_pedido_cliente, cliente_pedido_var, solicitante_pedido_var

        # Cliente
        ttk.Label(frame_pedido_info, text="Cliente:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        cliente_pedido_var = tk.StringVar(root)
        entry_pedido_cliente = ttk.Entry(frame_pedido_info, width=20, textvariable=cliente_pedido_var)
        entry_pedido_cliente.grid(row=0, column=1, padx=5, pady=5)
        entry_pedido_cliente.bind("<KeyRelease>", atualizar_lista_clientes)

        # Responsável (MANTIDO, assumindo que este é quem está criando/separando o pedido)
        ttk.Label(frame_pedido_info, text="Responsável:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        lista_resps_ped = obter_nomes_aux("RESPONSAVEL") 
        responsavel_pedido_var = tk.StringVar(root)
        responsavel_pedido_var.set(lista_resps_ped[0])
        # Atualizado para usar a lista do banco (*lista_resps_ped)
        menu_resp_separar_pedido = ttk.OptionMenu(frame_pedido_info, responsavel_pedido_var, lista_resps_ped[0], *lista_resps_ped)
        menu_resp_separar_pedido.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        # --- NOVO CAMPO SOLICITANTE ---
        ttk.Label(frame_pedido_info, text="Solicitante:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
        solicitante_pedido_var = tk.StringVar(root)

        # BUSCA DINÂMICA DO BANCO
        lista_solics_ped = obter_nomes_aux("SOLICITANTE")

        combo_solicitante_pedido = ttk.Combobox(
            frame_pedido_info, 
            textvariable=solicitante_pedido_var, 
            values=lista_solics_ped,
            state='readonly',
            width=18
        )
        combo_solicitante_pedido.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

        if lista_solics_ped:
            combo_solicitante_pedido.set(lista_solics_ped[0])
        # -----------------------------

        # --- NOVO CAMPO URGÊNCIA ---
        global urgente_var # Importante para a função de salvar enxergar
        urgente_var = tk.StringVar(value="Não")

        check_urgente = tk.Checkbutton(
            frame_pedido_info, 
            text="⚠️ PEDIDO URGENTE", 
            variable=urgente_var,
            onvalue="Sim", 
            offvalue="Não",
            font=("Arial", 10, "bold"),
            fg="red",
            cursor="hand2"
        )
        # Colocamos na linha 3, e o botão de confirmar desce para a linha 4
        check_urgente.grid(row=3, column=0, columnspan=2, pady=5)

        # AJUSTE O BOTÃO DE CONFIRMAR PARA A LINHA 4 (ele estava na 3)
        ttk.Button(frame_pedido_info, text="CONFIRMAR PEDIDO", command=registrar_novo_pedido_ou_atualizar_gui).grid(row=4, column=0, columnspan=2, pady=10)


        # Controles de Itens (Código, Qtd)
        frame_pedido_item = ttk.LabelFrame(frame_pedido_superior, text="Adicionar Itens", padding=(10, 10))
        frame_pedido_item.pack(side="left", padx=10, fill="y")

        ttk.Label(frame_pedido_item, text="Cód:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        entry_pedido_codigo = ttk.Entry(frame_pedido_item, width=15)
        entry_pedido_codigo.grid(row=0, column=1, padx=5, pady=5)
        entry_pedido_codigo.bind("<Return>", lambda event: entry_pedido_quantidade.focus_set())

        ttk.Label(frame_pedido_item, text="Qtd:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        entry_pedido_quantidade = ttk.Entry(frame_pedido_item, width=10)
        entry_pedido_quantidade.grid(row=1, column=1, padx=5, pady=5)
        entry_pedido_quantidade.bind("<Return>", lambda event: adicionar_item_pedido_gui())

        ttk.Button(frame_pedido_item, text="Adicionar Item", command=adicionar_item_pedido_gui).grid(row=2, column=0, columnspan=2, pady=5)

        btn_remover = tk.Button(frame_pedido_item, text="❌ Remover Selecionado", 
                                command=excluir_item_pedido_gui, 
                                bg="#ffffff", fg="black", font=("Arial", 9))
        btn_remover.grid(row=3, column=0, columnspan=2, pady=5, sticky="ew")

        # Treeview para Itens do Pedido
        columns_itens_pedido = ("Código", "Quantidade")
        tree_itens_pedido = ttk.Treeview(tab_pedido, columns=columns_itens_pedido, show="headings")
        tree_itens_pedido.pack(pady=10,expand=True, fill="both", padx=10)
        for col in columns_itens_pedido:
            tree_itens_pedido.heading(col, text=col)
            tree_itens_pedido.column(col, width=150, anchor="center")
            


        # --- Aba: Pedidos Separados (Acompanhamento) ---
        tab_pedidos_separados = ttk.Frame(notebook)
        notebook.add(tab_pedidos_separados, text="Pedidos")
        notebook.bind("<<NotebookTabChanged>>", lambda event: carregar_pedidos() if notebook.tab(notebook.select(), "text") == "Pedidos" else None)
        tk.Label(tab_pedidos_separados, text="Gerenciamento de Pedidos", font=("Arial", 18, "bold")).pack(pady=10)

        # Sub-Notebook para organizar as listas de Pendentes e Separados
        notebook_pedidos = ttk.Notebook(tab_pedidos_separados)
        notebook_pedidos.pack(pady=10, expand=True, fill="both", padx=10)

        # --- Sub-Aba: Pedidos do Dia (Triagem/Rascunho) ---
        tab_pedidos_dia = ttk.Frame(notebook_pedidos)
        # Inserimos na posição 0 para ser a primeira aba
        notebook_pedidos.add(tab_pedidos_dia, text="Pedidos do Dia (Triagem)")

        # Frame para os botões de ação da Triagem
        frame_acoes_dia = ttk.Frame(tab_pedidos_dia)
        frame_acoes_dia.pack(fill="x", padx=10, pady=5)

        # Botões da nova aba
        btn_validar = ttk.Button(frame_acoes_dia, text="Validar Estoque (Tudo)", command=executar_validacao_gui)
        btn_validar.pack(side=tk.LEFT, padx=5)

        btn_enviar_pendente = ttk.Button(frame_acoes_dia, text="Enviar Selecionado para Separação", command=promover_pedido_selecionado)
        btn_enviar_pendente.pack(side=tk.LEFT, padx=5)

        btn_excluir = ttk.Button(frame_acoes_dia, text="Excluir Rascunho", command=acao_excluir_rascunho)
        btn_excluir.pack(side="left", padx=5)

        btn_atualizar_triagem = ttk.Button(
            frame_acoes_dia, 
            text="🔄 Atualizar Lista", 
            command=carregar_pedidos_dia
        )
        btn_atualizar_triagem.pack(side="left", padx=5)

        # Treeview da Triagem
        columns_dia = ("ID", "Data", "Cliente", "Solicitante")
        tree_pedidos_dia = ttk.Treeview(tab_pedidos_dia, columns=columns_dia, show="headings")
        tree_pedidos_dia.bind("<Double-1>", lambda event: visualizar_itens_rascunho())
        tree_pedidos_dia.bind("<Delete>", acao_excluir_rascunho)
        tree_pedidos_dia.bind("<F5>", lambda event: carregar_pedidos_dia())
        tree_pedidos_dia.bind("<Return>", promover_pedido_selecionado)
        tree_pedidos_dia.pack(pady=5, expand=True, fill="both", padx=10)

        for col in columns_dia:
            tree_pedidos_dia.heading(col, text=col)
            tree_pedidos_dia.column(col, width=150, anchor="center")

        # Configuração de cores (Tags)
        tree_pedidos_dia.tag_configure('falta', foreground='red')
        tree_pedidos_dia.tag_configure('ok', foreground='green')

        # --- Sub-Aba: Pendentes para Separação ---
        tab_pendentes = ttk.Frame(notebook_pedidos)
        notebook_pedidos.add(tab_pendentes, text="Pendentes para Separação")

        columns_pedidos = ("ID", "Data", "Detalhes", "Itens do Pedido")
        tree_pedidos_pendentes = ttk.Treeview(tab_pendentes, columns=columns_pedidos, show="headings")
        tree_pedidos_pendentes.bind("<Double-1>", mostrar_detalhes_pedido)
        tree_pedidos_pendentes.pack(pady=10, expand=True, fill="both", padx=10)

        for col in columns_pedidos:
            tree_pedidos_pendentes.heading(col, text=col, command=lambda c=col: organizar_arvore(tree_pedidos_pendentes, c, False))
            tree_pedidos_pendentes.column(col, width=150)
            
        tree_pedidos_pendentes.column("ID", width=50, anchor="center")
        tree_pedidos_pendentes.column("Data", width=120, anchor="center")
        tree_pedidos_pendentes.column("Detalhes", width=200)
        tree_pedidos_pendentes.column("Itens do Pedido", width=300)

        frame_botoes_pendentes = ttk.Frame(tab_pedidos_separados)
        frame_botoes_pendentes.pack(pady=10, fill='x')

        # Botão EDITAR PEDIDO
        btn_editar_pedido = ttk.Button(
            frame_botoes_pendentes, 
            text="Editar Pedido", 
            command=carregar_pedido_para_edicao # Liga à função que carrega os dados
        )
        btn_editar_pedido.pack(side=tk.LEFT, padx=10)
        ttk.Button(tab_pendentes, text="Marcar como SEPARADO (Baixar Estoque)", command=mover_pedido_gui).pack(pady=10)

        # --- Sub-Aba: Separados / Expedição ---
        tab_separados = ttk.Frame(notebook_pedidos)
        notebook_pedidos.add(tab_separados, text="Separados / Mesanino")

        tree_pedidos_separados = ttk.Treeview(tab_separados, columns=columns_pedidos, show="headings")
        tree_pedidos_separados.bind("<Double-1>", mostrar_detalhes_pedido)
        tree_pedidos_separados.pack(pady=10, expand=True, fill="both", padx=10)

        # --- CÓDIGO DO NOVO BOTÃO PARA MOVER PARA EXPEDIÇÃO ---
        frame_botoes_separados = ttk.Frame(tab_separados) # Cria um frame para organizar o botão
        frame_botoes_separados.pack(pady=10)

        for col in columns_pedidos:
            tree_pedidos_separados.heading(col, text=col, command=lambda c=col: organizar_arvore(tree_pedidos_separados, c, False))
            tree_pedidos_separados.column(col, width=150)
            
        tree_pedidos_separados.column("ID", width=50, anchor="center")
        tree_pedidos_separados.column("Data", width=120, anchor="center")
        tree_pedidos_separados.column("Detalhes", width=200)
        tree_pedidos_separados.column("Itens do Pedido", width=300)


        # Botão para mover o pedido para Expedição
        btn_mover_expedicao = ttk.Button(
            frame_botoes_separados, 
            text="Enviar para Expedição", 
            command=mover_expedicao_gui
        )
        btn_mover_expedicao.pack(side=tk.LEFT, padx=10)

        botao_excluir_pedido = ttk.Button(
            frame_botoes_separados, 
            text="Excluir Pedido (Estornar)", 
            command=abrir_modal_estorno  # <--- CHAMA O NOVO MODAL!
        )
        botao_excluir_pedido.pack(pady=10)

        # Aba 3: Pedidos na Expedição
        tab_expedicao = ttk.Frame(notebook_pedidos)
        notebook_pedidos.add(tab_expedicao, text='Pedidos na Expedição') 

        # Configuração e exibição da Treeview na aba de Expedição
        tree_pedidos_expedicao = ttk.Treeview(tab_expedicao, columns=columns_pedidos, show="headings")
        tree_pedidos_expedicao.bind("<Double-1>", mostrar_detalhes_pedido)
        tree_pedidos_expedicao.pack(expand=True, fill='both')

        frame_botoes_expedicao = ttk.Frame(tab_expedicao)
        frame_botoes_expedicao.pack(pady=10)

        for col in columns_pedidos:
            tree_pedidos_expedicao.heading(col, text=col, command=lambda c=col: organizar_arvore(tree_pedidos_expedicao, c, False))
            tree_pedidos_expedicao.column(col, width=150)
            
        tree_pedidos_expedicao.column("ID", width=50, anchor="center")
        tree_pedidos_expedicao.column("Data", width=120, anchor="center")
        tree_pedidos_expedicao.column("Detalhes", width=200)
        tree_pedidos_expedicao.column("Itens do Pedido", width=300)

        # Botão para finalizar o pedido
        btn_finalizar_pedido = ttk.Button(
            frame_botoes_expedicao, 
            text="Finalizar Pedido", 
            command=finalizar_pedido_gui
        )
        btn_finalizar_pedido.pack(side=tk.LEFT, padx=10)

        # Aba 4: Histórico (Concluídos)
        tab_historico = ttk.Frame(notebook_pedidos)
        notebook_pedidos.add(tab_historico, text='Histórico de Pedidos') 

        # --- Adicionar uma área para os Filtros ---
        frame_filtros_historico_pedidos = ttk.Frame(tab_historico)
        frame_filtros_historico_pedidos.pack(fill='x', padx=10, pady=5)

        # 1. Filtro por Cliente
        ttk.Label(frame_filtros_historico_pedidos, text="Cliente:").pack(side=tk.LEFT, padx=5)
        # Use a declaração global ANTES de criar o widget
        global entry_filtro_historico_cliente 
        entry_filtro_historico_cliente = ttk.Entry(frame_filtros_historico_pedidos, width=20)
        entry_filtro_historico_cliente.pack(side=tk.LEFT, padx=5)

        # 2. Filtro por Data de Finalização
        ttk.Label(frame_filtros_historico_pedidos, text="Finalizado Em:").pack(side=tk.LEFT, padx=5)
        # Use a declaração global ANTES de criar o widget
        global entry_filtro_historico_data 
        entry_filtro_historico_data = DateEntry(
            frame_filtros_historico_pedidos,
            width=12,
            date_pattern='dd/mm/yyyy',
            locale='pt_BR'
        )
        entry_filtro_historico_data.pack(side=tk.LEFT, padx=5)

        # 3. Botão de Consultar/Recarregar
        # Certifique-se que o status passado é 'Concluído', se é a aba que lista pedidos concluídos
        ttk.Button(frame_filtros_historico_pedidos, text="Aplicar Filtros", command=lambda: carregar_pedidos('Concluído')).pack(side=tk.LEFT, padx=10)

        # --- NOVO: Botão Limpar Filtros ---
        ttk.Button(frame_filtros_historico_pedidos, text="Limpar Filtros", command=limpar_filtros_historico_pedidos).pack(side=tk.LEFT, padx=5)

        # Criar a Treeview para o Histórico
        tree_historico = ttk.Treeview(tab_historico, columns=columns_pedidos, show="headings")
        # Note: Você pode querer adicionar mais colunas aqui, como 'Data de Finalização'
        tree_historico.heading('#2', text='Finalizado em', 
                            command=lambda: ordenar_treeview(tree_historico, '#2', False)) 
        # ... (Adicionar as outras headings/cabeçalhos) ...
        tree_historico.pack(expand=True, fill='both')

        # Vincular o clique duplo para ver detalhes no histórico
        tree_historico.bind("<Double-1>", mostrar_detalhes_pedido)


        # Aba Histórico Faltantes
        # 1. CRIAR O CONTAINER DA ABA
        aba_historico_faltantes = ttk.Frame(notebook)
        notebook.add(aba_historico_faltantes, text=" Histórico de Faltantes ")

        # 2. ÁREA DE FILTROS (No topo da aba)
        frame_filtros_faltantes = ttk.LabelFrame(aba_historico_faltantes, text=" Filtros de Pesquisa ")
        frame_filtros_faltantes.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_filtros_faltantes, text="Buscar:").grid(row=0, column=0, padx=5, pady=5)
        entry_busca_faltantes = ttk.Entry(frame_filtros_faltantes, width=30)
        entry_busca_faltantes.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(frame_filtros_faltantes, text="Data Início:").grid(row=0, column=2, padx=5, pady=5)
        entry_data_inicio_faltantes = ttk.Entry(frame_filtros_faltantes, width=12)
        entry_data_inicio_faltantes.grid(row=0, column=3, padx=5, pady=5)

        ttk.Label(frame_filtros_faltantes, text="Data Fim:").grid(row=0, column=4, padx=5, pady=5)
        entry_data_fim_faltantes = ttk.Entry(frame_filtros_faltantes, width=12)
        entry_data_fim_faltantes.grid(row=0, column=5, padx=5, pady=5)

        btn_filtrar_faltantes = ttk.Button(frame_filtros_faltantes, text="Pesquisar", command=carregar_historico_faltantes)
        btn_filtrar_faltantes.grid(row=0, column=6, padx=10, pady=5)

        # 3. TABELA (TREEVIEW)
        colunas_faltantes = ("Data", "Pedido", "Cliente", "Código", "Produto", "Qtd. Perdida")
        tree_faltantes = ttk.Treeview(aba_historico_faltantes, columns=colunas_faltantes, show="headings")

        # Cabeçalhos
        for col in colunas_faltantes:
            tree_faltantes.heading(col, text=col)

        # Ajuste de largura das colunas
        tree_faltantes.column("Data", width=140, anchor="center")
        tree_faltantes.column("Pedido", width=80, anchor="center")
        tree_faltantes.column("Cliente", width=200)
        tree_faltantes.column("Código", width=100, anchor="center")
        tree_faltantes.column("Produto", width=250)
        tree_faltantes.column("Qtd. Perdida", width=100, anchor="center")

        tree_faltantes.pack(fill="both", expand=True, padx=10, pady=10)


        # Frame para o rodapé da tabela de histórico
        frame_rodape_historico = ttk.Frame(aba_historico_faltantes)
        frame_rodape_historico.pack(fill=tk.X, padx=10, pady=5)

        # Label de texto fixo
        lbl_texto_total = ttk.Label(frame_rodape_historico, text="Total de Itens Faltantes:", font=("Arial", 10, "bold"))
        lbl_texto_total.pack(side=tk.LEFT)

        # Label que vai receber o número (começa com 0)
        var_total_faltantes = tk.StringVar(value="0")
        lbl_valor_total = ttk.Label(frame_rodape_historico, textvariable=var_total_faltantes, font=("Arial", 10, "bold"), foreground="red")
        lbl_valor_total.pack(side=tk.LEFT, padx=5)

        btn_pdf = ttk.Button(frame_rodape_historico, text="Exportar PDF Consolidado", command=acao_exportar_pdf)
        btn_pdf.pack(side=tk.RIGHT, padx=10)

# -------------------------------------

# ----------------- Fim da Aplicação -----------------

if __name__ == "__main__":
    configurar_banco_usuarios() 
    
    app_login = TelaLogin()
    role_recebida = app_login.rodar()
    
    if role_recebida:
        # --- ESTA É A PARTE CRUCIAL ---
        # Se vier ('admin',), pegamos o primeiro item. 
        # Se já vier 'admin', mantemos como está.
        if isinstance(role_recebida, (tuple, list)):
            USUARIO_ROLE = role_recebida
        else:
            USUARIO_ROLE = role_recebida
        # ------------------------------
        
        root = ThemedTk(theme="equilux")
        root.title("Sistema de Controle de Estoque")
        
        montar_sistema_principal(root)
        
        root.mainloop()