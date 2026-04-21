import sqlite3
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from datetime import datetime
import os
import pandas as pd
import unicodedata
import json # NOVO: Importado para serializar os itens do pedido
import csv

# --- Funções do Banco de Dados ---
def criar_tabelas():
    """Cria as tabelas necessárias no banco de dados se elas não existirem."""
    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()
    
    # Tabela de Pedidos (NOVA TABELA)
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS pedidos (
        id_pedido INTEGER PRIMARY KEY AUTOINCREMENT,
        data_criacao TEXT NOT NULL,
        cliente TEXT NOT NULL,
        solicitante TEXT,
        separador TEXT,
        status TEXT NOT NULL, -- Valores: Mesanino, Expedição
        itens TEXT NOT NULL, -- Lista de produtos/quantidades serializada em JSON
        urgente TEXT DEFAULT 'Não',
        data_expedicao TEXT -- Data em que o pedido saiu do mesanino
    )
    ''')
    # Tabela de cadastros
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS cadastros_auxiliares (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tipo TEXT NOT NULL, -- 'SOLICITANTE', 'RESPONSAVEL', 'SEPARADOR'
        nome TEXT NOT NULL,
        UNIQUE(tipo, nome) -- Impede nomes duplicados no mesmo tipo
    )
    ''')

    # Tabela de Produtos (EXISTENTE)
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS produtos (
        id_produto INTEGER PRIMARY KEY AUTOINCREMENT,
        codigo TEXT NOT NULL UNIQUE,
        nome TEXT NOT NULL,
        estoque_atual INTEGER NOT NULL,
        estoque_minimo INTEGER,
        media_vendas_mensal REAL,
        media_45_dias INTEGER DEFAULT 0,
        data_ultima_atualizacao TEXT,
        reservado INTEGER DEFAULT 0
    )
    ''')

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS historico_faltantes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data_corte TEXT,
            id_pedido INTEGER,
            cliente TEXT,
            codigo_produto TEXT,
            nome_produto TEXT,
            quantidade_faltante INTEGER
        )
    """)      

    # Verifica se a coluna 'reservado' existe, caso contrário, adiciona
    cursor.execute("PRAGMA table_info(produtos)")
    colunas = [coluna[1] for coluna in cursor.fetchall()]

    if 'reservado' not in colunas:
        cursor.execute("ALTER TABLE produtos ADD COLUMN reservado INTEGER DEFAULT 0")
        print("Coluna 'reservado' adicionada com sucesso!")

    if 'media_45_dias' not in colunas:
        try:
            cursor.execute("ALTER TABLE produtos ADD COLUMN media_45_dias INTEGER DEFAULT 0")
            print("Coluna 'media_45_dias' adicionada com sucesso!")
        except sqlite3.OperationalError:
            pass    

    # Tabela de Entradas (EXISTENTE)
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS entradas (
        id_entrada INTEGER PRIMARY KEY AUTOINCREMENT,
        id_produto INTEGER,
        quantidade INTEGER NOT NULL,
        tipo_entrada TEXT,
        responsavel TEXT,
        data TEXT NOT NULL,
        FOREIGN KEY(id_produto) REFERENCES produtos(id_produto)
    )
    ''')

    # Tabela de Saídas (EXISTENTE)
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS saidas (
        id_saida INTEGER PRIMARY KEY AUTOINCREMENT,
        id_produto INTEGER,
        quantidade INTEGER NOT NULL,
        cliente TEXT,
        responsavel TEXT,
        separador TEXT,
        data TEXT NOT NULL,
        FOREIGN KEY(id_produto) REFERENCES produtos(id_produto)
    )
    ''')
    
    # --- NOVO: Criação da Tabela de Histórico de Movimentação ---
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS historico (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data_registro TEXT,
            codigo_produto TEXT,
            tipo TEXT,         -- Ex: ENTRADA, SAIDA - SEPARACAO
            quantidade REAL,
            responsavel TEXT,
            observacao TEXT,
            FOREIGN KEY (codigo_produto) REFERENCES produtos(codigo)
        )
    """)
    # -----------------------------------------------------------
    
    try:
        # 1. Coluna de Finalização (Se ainda não foi adicionada)
        cursor.execute("ALTER TABLE pedidos ADD COLUMN Data_Finalizacao TEXT")
    except sqlite3.OperationalError:
        pass

    try:
    # 2. COLUNA DE RESPONSÁVEL PELA EXPEDIÇÃO (NOVA CORREÇÃO)
        cursor.execute("ALTER TABLE pedidos ADD COLUMN Responsavel_Expedicao TEXT")
    except sqlite3.OperationalError:
        pass
    try:
    # --- NOVO BLOCO: Adiciona a coluna Data_Separacao ---
        cursor.execute("ALTER TABLE pedidos ADD COLUMN Data_Separacao TEXT")
    except sqlite3.OperationalError:
        pass
    try:
        cursor.execute("ALTER TABLE pedidos ADD COLUMN urgente TEXT DEFAULT 'Não'")
        conn.commit()
    except:
        pass        
    
    conn.commit()
    conn.close()

def configurar_banco_usuarios():
    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario TEXT UNIQUE NOT NULL,
            senha TEXT NOT NULL,
            role TEXT NOT NULL
        )
    ''')
    # Insira aqui os usuários padrão (Richard como admin)
    cursor.execute("INSERT OR IGNORE INTO usuarios (usuario, senha, role) VALUES (?, ?, ?)", 
                   ('celio', 'celio1234', 'admin'))
    cursor.execute("INSERT OR IGNORE INTO usuarios (usuario, senha, role) VALUES (?, ?, ?)", 
                   ('julio', '123456', 'visualizador'))
    conn.commit()
    conn.close()

def validar_acesso(usuario, senha):
    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()
    # Procuramos a role (permissão) do utilizador
    cursor.execute("SELECT role FROM usuarios WHERE usuario=? AND senha=?", (usuario, senha))
    resultado = cursor.fetchone()
    conn.close()
    
    # Se 'resultado' existe, ele vem como ('admin',). 
    # Usamos para pegar apenas o primeiro elemento da tupla: 'admin'
    return resultado if resultado else None    

# --- Funções de Manipulação do Estoque ---
# (Manter todas as funções existentes aqui: adicionar_produto, registrar_entrada, registrar_saida,
# consultar_movimentacoes, consultar_estoque_baixo, excluir_produto, limpar_historico_produto,
# consultar_estoque_geral, buscar_inventario, importar_produtos_excel,
# registrar_entradas_lote, registrar_saidas_lote, atualizar_estoque_minimo)

# ... (Seu código original de 'adicionar_produto' até 'registrar_saidas_lote' vai aqui)
# Para economizar espaço, mantenho as funções existentes (registrar_saida, etc) intactas.

def listar_cadastros_aux(tipo):
    """tipo pode ser 'SOLICITANTE', 'RESPONSAVEL' ou 'SEPARADOR'"""
    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()
    cursor.execute("SELECT nome FROM cadastros_auxiliares WHERE tipo = ? ORDER BY nome", (tipo,))
    nomes = [linha[0] for linha in cursor.fetchall()]
    conn.close()
    return nomes

def adicionar_cadastro_aux(tipo, nome):
    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO cadastros_auxiliares (tipo, nome) VALUES (?, ?)", (tipo, nome.upper()))
        conn.commit()
        return True
    except:
        return False
    finally:
        conn.close()

def remover_cadastro_aux(tipo, nome):
    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()
    cursor.execute("DELETE FROM cadastros_auxiliares WHERE tipo = ? AND nome = ?", (tipo, nome))
    conn.commit()
    conn.close()

def adicionar_produto(codigo, nome, estoque_inicial, estoque_minimo):
    """
    Adiciona um novo produto à tabela 'produtos'.
    """
    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()
    data_atual = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    try:
        cursor.execute('''
        INSERT INTO produtos (codigo, nome, estoque_atual, estoque_minimo, media_vendas_mensal, data_ultima_atualizacao)
        VALUES (?, ?, ?, ?, ?, ?)
        ''', (codigo, nome, estoque_inicial, estoque_minimo, 0.0, data_atual))
        
        conn.commit()
        return True, f"Produto '{nome}' adicionado com sucesso!"
    except sqlite3.IntegrityError:
        return False, f"Erro: O código '{codigo}' já existe no banco de dados. O produto não foi adicionado."
    except Exception as e:
        return False, f"Ocorreu um erro: {e}"
    finally:
        conn.close()
        
def registrar_saida_log(cursor, codigo_produto, quantidade, cliente, solicitante, responsavel_expedicao):
    """
    Registra a LOG de saída na tabela 'historico' refletindo a saída física na Expedição.
    """
    from datetime import datetime
    
    try:
        # 1. Obter o novo saldo após a baixa
        cursor.execute("SELECT estoque_atual FROM produtos WHERE codigo = ?", (codigo_produto,))
        produto_info = cursor.fetchone()
        
        if not produto_info:
             print(f"AVISO: Produto {codigo_produto} não encontrado para log.")
             return False 

        # Trata o retorno dependendo de como o cursor está configurado
        novo_saldo = produto_info['estoque_atual'] if hasattr(produto_info, 'keys') else produto_info[0]

        # 2. Preparar dados
        timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        tipo_para_filtro = "Saída" 
        
        # No histórico, o 'responsável' pela saída física agora é quem moveu para expedição
        # Mas mantemos os outros nomes na observação para rastreabilidade total
        observacao_detalhada = (f"Saída Expedição | Cliente: {cliente} | "
                                f"Solicitante: {solicitante} | Novo Saldo: {novo_saldo}")
        
        # 3. Executar o INSERT
        cursor.execute("""
            INSERT INTO historico (data_registro, codigo_produto, tipo, quantidade, responsavel, observacao)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (timestamp, codigo_produto, tipo_para_filtro, quantidade, responsavel_expedicao, observacao_detalhada))
        
        return True
        
    except Exception as e:
        print(f"Erro interno ao registrar LOG: {e}")
        raise # Mantém o raise para garantir o rollback na função principal
        
def registrar_saida(codigo, quantidade, cliente, responsavel, separador=None, observacao_adicional=""):
    """
    Registra a saída de um único item no estoque e atualiza o saldo.
    Retorna (sucesso: bool, mensagem: str)
    """
    
    # 1. Validação
    if not isinstance(quantidade, int) or quantidade <= 0:
        return False, "Quantidade deve ser um número inteiro positivo."
    if not codigo:
        return False, "Código do produto não pode ser vazio."

    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()
    
    try:
        # A. Verificar se o produto existe e obter o estoque
        cursor.execute("SELECT estoque_atual FROM produtos WHERE codigo = ?", (codigo,))
        resultado = cursor.fetchone()
        
        if not resultado:
            return False, f"Produto com código '{codigo}' não encontrado."

        estoque_atual = resultado[0]
        
        # B. Checar se há estoque suficiente
        if estoque_atual < quantidade:
            return False, f"Estoque insuficiente para o produto {codigo}. Saldo atual: {estoque_atual}."
            
        novo_estoque = estoque_atual - quantidade # <--- SUBTRAÇÃO

        # C. Atualizar o saldo de estoque
        cursor.execute("""
            UPDATE produtos 
            SET estoque_atual = ? 
            WHERE codigo = ?
        """, (novo_estoque, codigo))
        
        # D. Registrar no histórico de movimentações (Onde está a correção para o filtro)
        timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        # Padroniza o tipo para o filtro da GUI
        tipo_para_filtro = "Saída" 
        
        # Cria a observação detalhada (importante para Saídas de Pedidos)
        detalhe_cliente = f"Cliente: {cliente}" if cliente else "Sem Cliente"
        detalhe_separador = f" | Solicitante: {separador}" if separador else ""
        
        observacao_detalhada = f"{detalhe_cliente}{detalhe_separador} | Observação: {observacao_adicional}. Novo Saldo: {novo_estoque}"


        cursor.execute("""
            INSERT INTO historico (data_registro, codigo_produto, tipo, quantidade, responsavel, observacao)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (timestamp, codigo, tipo_para_filtro, quantidade, responsavel, observacao_detalhada))
        #                                          ^^^^^^^^^^^^^^^^
        #                                          Agora salva "Saída"

        conn.commit()
        return True, f"Saída de {quantidade} unidades do produto {codigo} registrada. Novo saldo: {novo_estoque}."

    except sqlite3.Error as e:
        conn.rollback()
        return False, f"Erro de banco de dados ao registrar saída: {e}"
        
    finally:
        conn.close()

def consultar_movimentacoes(codigo=None, cliente=None, tipo_mov='Ambas', data_inicial_str=None, data_final_str=None):
    """
    Consulta o histórico de movimentações (entradas, saídas, etc) da tabela 'historico'.
    Inclui filtragem por Código, Cliente/Responsável, Tipo e Intervalo de Datas.
    """
    import sqlite3
    # Importar datetime é crucial para a conversão de formato do filtro!
    from datetime import datetime 
    
    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()
    
    # *** 1. EXPRESSÃO SQL PARA CONVERTER DATA PARA FORMATO ORDENÁVEL (AAAA/MM/DD) ***
    # Isso é necessário porque o seu banco armazena em DD/MM/AAAA.
    # Usamos o formato AAAAMMDDHHMMSS (sem hífens) para garantir a ordenação cronológica da string.
    DATA_CONVERTIDA_SQL = "SUBSTR(T1.data_registro, 7, 4) || SUBSTR(T1.data_registro, 4, 2) || SUBSTR(T1.data_registro, 1, 2) || SUBSTR(T1.data_registro, 11)"
    
    # Monta a Query Base
    query = '''
    SELECT 
        T1.data_registro, T1.codigo_produto, T2.nome, T1.quantidade, 
        T1.tipo, T1.observacao, T1.responsavel
    FROM historico AS T1
    INNER JOIN produtos AS T2 ON T1.codigo_produto = T2.codigo
    '''
    
    where_clauses = []
    params = []

    # 2. Adiciona Filtros
    
    if tipo_mov != 'Ambas':
        where_clauses.append('T1.tipo = ?')
        params.append(tipo_mov)
        
    if codigo:
        where_clauses.append('T1.codigo_produto = ?') 
        params.append(codigo)
        
    if cliente:
        where_clauses.append('T1.observacao LIKE ?') 
        params.append(f'%{cliente}%')
    
    # ----------------- NOVO FILTRO POR DATA (CRONOLOGICAMENTE CORRETO) -----------------
    
    if data_inicial_str and data_inicial_str.strip():
        try:
            # Converte o filtro do usuário para o formato AAAA/MM/DD para comparação
            data_str_filtrada = data_inicial_str[6:10] + data_inicial_str[3:5] + data_inicial_str[0:2]
            
            # Usa a expressão SQL convertida para comparar com o filtro (AAAAMMDD 00:00:00)
            where_clauses.append(f"{DATA_CONVERTIDA_SQL} >= ?")
            params.append(f"{data_str_filtrada} 00:00:00")
            
        except Exception:
             # Ignora se a data for inválida ou o slicing falhar
             pass
        
    if data_final_str and data_final_str.strip():
        try:
            # Converte o filtro do usuário para o formato AAAA/MM/DD para comparação
            data_str_filtrada = data_final_str[6:10] + data_final_str[3:5] + data_final_str[0:2]
            
            # AÇÃO PRINCIPAL: Usa a expressão SQL convertida (DATA_CONVERTIDA_SQL)
            where_clauses.append(f"{DATA_CONVERTIDA_SQL} <= ?")
            params.append(f"{data_str_filtrada} 23:59:59")
            
        except Exception:
             # Ignora se a data for inválida ou o slicing falhar
             pass
    
    # ------------------------------------------------------------------------------------
        
    # 3. Finaliza a Query
    if where_clauses:
        query += ' WHERE ' + ' AND '.join(where_clauses)
        
    # AÇÃO CRÍTICA: ORDENA PELA DATA CONVERTIDA para garantir a ordem cronológica
    query += f" ORDER BY {DATA_CONVERTIDA_SQL} DESC"
    
    try:
        cursor.execute(query, params)
        movimentacoes = cursor.fetchall()
        
        conn.close()
        
        # 4. Formata o Resultado (mantido o seu código original)
        resultado = []
        for item in movimentacoes:
            resultado.append({
                "Data": item[0], 
                "Código": item[1], 
                "Produto": item[2], 
                "Quantidade": item[3],
                "Tipo": item[4], 
                "Detalhe": item[5], 
                "Responsável": item[6] 
            })
            
        return resultado
        
    except sqlite3.OperationalError as e:
        print(f"Erro ao consultar histórico: {e}")
        conn.close()
        return []
    except Exception as e:
        print(f"Erro geral: {e}")
        conn.close()
        return []
        
def consultar_estoque_baixo():
    """
    Consulta produtos com estoque abaixo do nível mínimo.
    Retorna uma lista de dicionários.
    """
    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()
    cursor.execute('''
    SELECT codigo, nome, estoque_atual, estoque_minimo
    FROM produtos
    WHERE estoque_atual <= estoque_minimo
    ''')
    produtos_baixo_estoque = cursor.fetchall()
    conn.close()
    resultado = []
    for item in produtos_baixo_estoque:
        resultado.append({
            "Código": item[0], "Nome": item[1], "Estoque Atual": item[2], "Estoque Mínimo": item[3]
        })
    return resultado

def excluir_produto(codigo_produto):
    """
    Exclui um produto e todas as suas movimentações associadas.
    """
    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT id_produto, nome FROM produtos WHERE codigo = ?", (codigo_produto,))
        produto = cursor.fetchone()
        if not produto:
            return False, "Produto não encontrado."
        
        id_produto, nome_produto = produto
        
        cursor.execute("DELETE FROM entradas WHERE id_produto = ?", (id_produto,))
        cursor.execute("DELETE FROM saidas WHERE id_produto = ?", (id_produto,))
        
        # NOVO: Se o produto for excluído, também removemos pedidos associados (se houver)
        cursor.execute("DELETE FROM pedidos WHERE instr(itens, ?) > 0", (f'"codigo": "{codigo_produto}"',))
        
        cursor.execute("DELETE FROM produtos WHERE id_produto = ?", (id_produto,))
        
        conn.commit()
        return True, f"Produto '{nome_produto}' e suas movimentações excluídos com sucesso."
    except Exception as e:
        conn.rollback()
        return False, f"Ocorreu um erro ao excluir o produto: {e}"
    finally:
        conn.close()

def limpar_historico_produto(codigo_produto):
    """
    Exclui todas as movimentações (entradas e saídas) de um produto,
    mantendo-o no inventário.
    """
    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT id_produto, nome FROM produtos WHERE codigo = ?", (codigo_produto,))
        produto = cursor.fetchone()
        if not produto:
            return False, "Produto não encontrado."
        
        id_produto, nome_produto = produto
        
        cursor.execute("DELETE FROM entradas WHERE id_produto = ?", (id_produto,))
        cursor.execute("DELETE FROM saidas WHERE id_produto = ?", (id_produto,))
        
        conn.commit()
        return True, f"Histórico de movimentações do produto '{nome_produto}' limpo com sucesso."
    except Exception as e:
        conn.rollback()
        return False, f"Ocorreu um erro ao limpar o histórico do produto: {e}"
    finally:
        conn.close()

def consultar_estoque_geral():
    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()
    
    cursor.execute('''
    SELECT codigo, nome, media_45_dias, estoque_atual, reservado, estoque_minimo
    FROM produtos
    ORDER BY nome
    ''')
    
    produtos = cursor.fetchall()
    conn.close()
    
    inventario = []
    for item in produtos:
        codigo, nome, media_45, atual, reservado, minimo = item
        
        # Garante que valores nulos virem 0 para o cálculo não quebrar
        atual_val = atual or 0
        media_val = media_45 or 0
        minimo_val = minimo or 0
        
        # --- APLICAÇÃO DA SUA NOVA REGRA ---
        if atual_val < media_val:
            faltante = media_val - atual_val
        else:
            faltante = 0
        # -----------------------------------
        
        status = "Abaixo do mínimo" if atual_val <= minimo_val else "Ok"
        
        inventario.append({
            'Código': codigo,
            'Nome': nome,
            'Média 45 dias': media_val,
            'Estoque Atual': atual_val,
            'Reservado': reservado or 0,
            'Quantidade Faltante': faltante,
            'Estoque Mínimo': minimo_val,
            'Status': status
        })
    return inventario

def carregar_inventario_com_duracao():
    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()
    # Buscamos o estoque atual e a média de 45 dias
    cursor.execute("SELECT codigo, nome, estoque_atual, reservado, media_45_dias FROM produtos")
    dados = cursor.fetchall()
    conn.close()
    return dados

def atualizar_media_db(codigo, nova_media):
    """Atualiza o valor da média 45 dias no banco de dados."""
    import sqlite3
    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()
    try:
        cursor.execute("UPDATE produtos SET media_45_dias = ? WHERE codigo = ?", (nova_media, codigo))
        conn.commit()
    except Exception as e:
        print(f"Erro ao atualizar média no banco: {e}")
    finally:
        conn.close()

def atualizar_medias_em_lote(lista_dados):
    """
    Recebe uma lista de tuplas [(media, codigo), (media, codigo)...]
    e atualiza o banco de dados.
    """
    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()
    
    try:
        # Usamos executemany para ser ultra rápido
        cursor.executemany('''
            UPDATE produtos 
            SET media_45_dias = ? 
            WHERE codigo = ?
        ''', lista_dados)
        
        conn.commit()
        return True, cursor.rowcount # Retorna sucesso e quantas linhas mudaram
    except Exception as e:
        print(f"Erro ao atualizar lote: {e}")
        return False, 0
    finally:
        conn.close()

def obter_lista_producao():
    """Retorna apenas os produtos onde a Média 45 dias é maior que o Estoque Atual."""
    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()
    
    cursor.execute('''
        SELECT codigo, nome, media_45_dias, estoque_atual 
        FROM produtos 
        WHERE media_45_dias > estoque_atual
        ORDER BY (media_45_dias - estoque_atual) DESC
    ''')
    
    dados = cursor.fetchall()
    conn.close()
    
    lista_producao = []
    for item in dados:
        codigo, nome, media, atual = item
        lista_producao.append({
            'Código': codigo,
            'Nome': nome,
            'Faltante':  media - atual
        })
    return lista_producao

def auditar_e_corrigir_reservas():
    import sqlite3
    import json
    
    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()

    try:
        # 1. Zera todas as reservas de todos os produtos para começar do zero
        cursor.execute("UPDATE produtos SET reservado = 0")

        # 2. Busca itens APENAS dos pedidos que estão em 'Pendente'
        # (Já que Rascunho não reserva e Separado já deu baixa)
        cursor.execute("SELECT itens FROM pedidos WHERE status = 'Pendente'")
        pedidos_pendentes = cursor.fetchall()

        for (itens_json,) in pedidos_pendentes:
            try:
                itens = json.loads(itens_json)
                for item in itens:
                    # Tenta pegar a chave com 'C' maiúsculo ou minúsculo
                    codigo = item.get('Código') or item.get('codigo')
                    qtd = int(item.get('Quantidade') or item.get('quantidade') or 0)
                    
                    if codigo and qtd > 0:
                        # 3. Aloca a reserva real
                        cursor.execute("""
                            UPDATE produtos 
                            SET reservado = reservado + ? 
                            WHERE codigo = ?
                        """, (qtd, str(codigo)))
            except Exception as e:
                print(f"Erro ao processar JSON de um pedido pendente: {e}")

        conn.commit()
        return True, "Reservas recalculadas com sucesso! Fantasmas removidos."
    except Exception as e:
        conn.rollback()
        return False, f"Erro na auditoria: {e}"
    finally:
        conn.close()

def consultar_detalhes_reservado():
    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()
    
    # Filtramos APENAS pedidos que estão no Mesanino (Ainda reservados)
    # Se você tiver outros status intermediários, adicione-os no IN ('Mesanino', 'Outro')
    cursor.execute('''
        SELECT id_pedido, cliente, itens 
        FROM pedidos 
        WHERE status = 'Pendente'
    ''')
    
    pedidos = cursor.fetchall()
    conn.close()
    
    lista_detalhada = []
    
    for id_ped, cliente, itens_json in pedidos:
        try:
            itens_lista = json.loads(itens_json)
            for item in itens_lista:
                # Pegando as chaves conforme o padrão do seu sistema
                codigo = item.get('codigo') or item.get('Código')
                # Somamos a quantidade (garantindo que seja número)
                qtd = item.get('quantidade') or item.get('Quantidade') or 0
                
                lista_detalhada.append((codigo, qtd, id_ped, cliente))
        except Exception as e:
            print(f"Erro ao ler itens do pedido {id_ped}: {e}")
            continue
            
    return lista_detalhada

def buscar_inventario(termo_busca=""):
    """
    Busca por produtos no inventário com suporte às novas colunas
    (Média e Quantidade Faltante).
    """
    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()

    termo_busca = f"%{termo_busca}%"

    # 1. Atualizamos o SELECT para trazer os novos campos
    cursor.execute('''
    SELECT codigo, nome, media_45_dias, estoque_atual, reservado, estoque_minimo
    FROM produtos
    WHERE codigo LIKE ? OR nome LIKE ?
    ORDER BY nome
    ''', (termo_busca, termo_busca))

    produtos = cursor.fetchall()
    conn.close()

    inventario = []
    for item in produtos:
        # 2. Desempacotamos todos os 6 campos vindos do banco
        codigo, nome, media_45, atual, reservado, minimo = item
        
        # 3. Refazemos os cálculos para manter o padrão
        media_45 = media_45 or 0
        atual = atual or 0
        if atual < media_45:
            faltante = media_45 - atual
        else:
            faltante = 0
        # ---------------------------------
        status = "Abaixo do mínimo" if atual <= (minimo or 0) else "Ok"

        # 4. Montamos o dicionário com as 8 chaves exatas da interface
        inventario.append({
            'Código': codigo,
            'Nome': nome,
            'Média 45 dias': media_45,
            'Estoque Atual': atual,
            'Reservado': reservado or 0,
            'Quantidade Faltante': faltante,
            'Estoque Mínimo': minimo or 0,
            'Status': status
        })
    return inventario

def importar_produtos_excel(filepath):
    """
    Importa produtos de uma planilha Excel e os adiciona ao banco de dados.
    Esperam-se colunas: 'codigo', 'nome', 'estoque_atual' e 'estoque_minimo'.
    """
    try:
        df = pd.read_excel(filepath)
        
        # Remove espaços, converte para minúsculas, remove acentos e substitui espaços por sublinhados
        df.columns = (
            df.columns.str.strip()
            .str.lower()
            .str.normalize('NFKD')
            .str.encode('ascii', errors='ignore')
            .str.decode('utf-8')
            .str.replace(' ', '_')
        )
        
        colunas_necessarias = ['codigo', 'nome', 'estoque_atual', 'estoque_minimo']
        if not all(col in df.columns for col in colunas_necessarias):
            colunas_encontradas = df.columns.tolist()
            return False, f"As colunas necessárias {colunas_necessarias} não foram encontradas na planilha. Colunas encontradas: {colunas_encontradas}"
        
        # Preenche valores ausentes (NaN) com 0 para evitar erros de conversão
        df['estoque_atual'] = df['estoque_atual'].fillna(0)
        df['estoque_minimo'] = df['estoque_minimo'].fillna(0)
        
        resultados = []
        for index, row in df.iterrows():
            codigo = str(row['codigo']).strip()
            nome = str(row['nome']).strip()
            
            try:
                estoque_atual = int(row['estoque_atual'])
                estoque_minimo = int(row['estoque_minimo'])
            except ValueError:
                resultados.append(f"Falha ao adicionar: {nome} ({codigo}) - 'estoque_atual' ou 'estoque_minimo' não são números válidos.")
                continue

            sucesso, mensagem = adicionar_produto(codigo, nome, estoque_atual, estoque_minimo)
            if not sucesso:
                resultados.append(mensagem)
            
        if resultados:
            return False, f"Importação concluída com algumas falhas:\n" + "\n".join(resultados)
        
        return True, "Importação de produtos concluída com sucesso!"

    except FileNotFoundError:
        return False, "Arquivo não encontrado."
    except pd.errors.ParserError:
        return False, "Erro ao ler o arquivo. Verifique se o formato está correto."
    except Exception as e:
        return False, f"Ocorreu um erro inesperado durante a importação: {e}"
        
def registrar_entrada(codigo, quantidade, tipo_entrada, responsavel):
    """
    Registra a entrada de um único item no estoque e atualiza o saldo.
    Retorna (sucesso: bool, mensagem: str)
    """
    
    # 1. Validação
    if not isinstance(quantidade, int) or quantidade <= 0:
        return False, "Quantidade deve ser um número inteiro positivo."
    if not codigo:
        return False, "Código do produto não pode ser vazio."

    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()
    
    try:
        # A. Verificar se o produto existe e obter o estoque
        cursor.execute("SELECT estoque_atual FROM produtos WHERE codigo = ?", (codigo,))
        resultado = cursor.fetchone()
        
        if not resultado:
            return False, f"Produto com código '{codigo}' não encontrado."

        # O resultado vem como uma tupla, o estoque é o primeiro elemento (índice 0)
        estoque_atual = resultado[0] 
        novo_estoque = estoque_atual + quantidade

        # B. Atualizar o saldo de estoque
        cursor.execute("""
            UPDATE produtos 
            SET estoque_atual = ? 
            WHERE codigo = ?
        """, (novo_estoque, codigo))
        
        # C. Registrar no histórico de movimentações 
        # AQUI FOI CORRIGIDO O PROBLEMA DO TIPO E O USO DA DATA/HORA
        timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        # Padroniza o tipo para o filtro da GUI (DEVE ser "Entrada" ou "Saída" para o seu filtro funcionar!)
        tipo_para_filtro = "Entrada" 
        #                 ^^^^^^^^^ MUDANÇA: AGORA SALVA O TERMO EXATO ESPERADO PELO FILTRO

        # Adiciona o detalhe do tipo_entrada original na observação
        # O valor "Produção" (que é o detalhe) pode ir para a observação.
        observacao_detalhada = f"Entrada ({tipo_entrada}). Estoque: {novo_estoque}"

        cursor.execute("""
            INSERT INTO historico (data_registro, codigo_produto, tipo, quantidade, responsavel, observacao)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (timestamp, codigo, tipo_para_filtro, quantidade, responsavel, observacao_detalhada))

        #                                          ^^^^^^^^^^^^^^^^^^
        #                                          Agora salva "Entrada"

        conn.commit()
        return True, f"Entrada de {quantidade} unidades do produto {codigo} registrada. Novo saldo: {novo_estoque}."

    except sqlite3.Error as e:
        conn.rollback()
        return False, f"Erro de banco de dados ao registrar entrada: {e}"
        
    finally:
        conn.close()  

# --- Funções de Movimentações em Lote ---
def registrar_entradas_lote(produtos_lote, tipo_entrada, responsavel):
    """
    Registra múltiplas entradas de produtos em uma única operação.
    produtos_lote deve ser uma lista de dicionários, ex: [{'codigo': 'ABC-101', 'quantidade': 10}]
    """
    resultados = []
    for item in produtos_lote:
        codigo = item.get('codigo')
        quantidade = item.get('quantidade')
        
        if not codigo or not quantidade:
            resultados.append({'codigo': codigo, 'sucesso': False, 'mensagem': 'Dados incompletos.'})
            continue

        try:
            quantidade = int(quantidade)
            sucesso, mensagem = registrar_entrada(codigo, quantidade, tipo_entrada, responsavel)
             # --- ADICIONE ESTE PRINT DE DEBBUG ---
            if not sucesso:
                print(f"DEBUG - ENTRADA LOTE FALHOU: Produto {codigo}. Mensagem: {mensagem}")
            # ------------------------------------
            resultados.append({'codigo': codigo, 'sucesso': sucesso, 'mensagem': mensagem})
        except ValueError:
            resultados.append({'codigo': codigo, 'sucesso': False, 'mensagem': f'Quantidade inválida para o produto {codigo}.'})
            
    return resultados

def registrar_saidas_lote(produtos_lote, cliente, responsavel, separador):
    """
    Registra saídas individualmente. 
    Se o produto 79 falhar, os outros que têm estoque serão processados normalmente.
    """
    resultados = []
    
    # Processamos item por item da lista enviada pela interface
    for item in produtos_lote:
        codigo = item.get('codigo')
        quantidade = item.get('quantidade')

        if not codigo or not quantidade:
            resultados.append({'codigo': codigo, 'sucesso': False, 'mensagem': 'Dados incompletos.'})
            continue
        
        try:
            quantidade = int(quantidade)
            
            # Chamamos sua função unitária que você postou acima.
            # Ela mesma verificará o estoque e dará o COMMIT se estiver tudo certo.
            sucesso, mensagem = registrar_saida(codigo, quantidade, cliente, responsavel, separador)
            
            # Guardamos o resultado (seja Sucesso ou Erro de estoque)
            resultados.append({
                'codigo': codigo, 
                'sucesso': sucesso, 
                'mensagem': mensagem
            })
            
        except ValueError:
            resultados.append({
                'codigo': codigo, 
                'sucesso': False, 
                'mensagem': 'Quantidade deve ser um número válido.'
            })
        except Exception as e:
            resultados.append({
                'codigo': codigo, 
                'sucesso': False, 
                'mensagem': f'Erro inesperado: {e}'
            })
            
    # Retornamos a lista completa de sucessos e falhas para a interface tratar
    return resultados

def excluir_rascunho_db(id_pedido):
    """Remove permanentemente um pedido com status Rascunho."""
    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()
    try:
        # Por segurança, garantimos que só deleta se for Rascunho
        cursor.execute("DELETE FROM pedidos WHERE id_pedido = ? AND status = 'Rascunho'", (id_pedido,))
        conn.commit()
        return True, f"Rascunho #{id_pedido} excluído com sucesso."
    except Exception as e:
        return False, f"Erro ao excluir: {e}"
    finally:
        conn.close()

# --- Funções para a Nova Funcionalidade: Separação de Pedidos ---

def buscar_pedidos_por_status(status_alvo):
    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()
    # Verifique se a ordem das colunas bate com o que a sua função 'carregar' espera
    cursor.execute("SELECT id_pedido, data_criacao, cliente, solicitante, itens FROM pedidos WHERE status = ?", (status_alvo,))
    dados = cursor.fetchall()
    conn.close()
    return dados # Importante: Verifique se tem essa linha!

def validar_estoque_rascunhos():
    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()
    
    # 1. Busca todos os rascunhos
    cursor.execute("SELECT id_pedido, cliente, itens FROM pedidos WHERE status = 'Rascunho'")
    rascunhos = cursor.fetchall()
    
    # 2. Busca o estoque disponível real (Estoque Atual - Reservado)
    cursor.execute("SELECT codigo, (estoque_atual - reservado) as disponivel FROM produtos")
    estoque = {row[0]: row[1] for row in cursor.fetchall()}
    
    resultados = []
    pedidos_com_falta = 0
    
    for id_ped, cliente, itens_json in rascunhos:
        try:
            itens = json.loads(itens_json)
            erro_no_pedido = False
            detalhes_falta = []
            
            for item in itens:
                cod = item.get('codigo') or item.get('Código')
                qtd = int(item.get('quantidade') or item.get('Quantidade') or 0)
                disponivel = estoque.get(cod, 0)
                
                if qtd > disponivel:
                    erro_no_pedido = True
                    detalhes_falta.append(f"{cod} (Faltam {qtd - disponivel})")
            
            if erro_no_pedido:
                pedidos_com_falta += 1
                resultados.append(f"Pedido #{id_ped} ({cliente}): {', '.join(detalhes_falta)}")
        except:
            continue
            
    conn.close()
    
    # RETORNO CORRIGIDO:
    return {
        "total": len(rascunhos),
        "faltantes": pedidos_com_falta,
        "logs": resultados
    }

def promover_pedido_para_pendente(id_pedido):
    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()
    
    try:
        # Busca os itens do pedido rascunho
        cursor.execute("SELECT itens FROM pedidos WHERE id_pedido = ?", (id_pedido,))
        resultado = cursor.fetchone()
        
        if not resultado:
            return False, "Pedido não encontrado."

        itens = json.loads(resultado[0])
        
        # PASSO 1: Reservar cada item no estoque
        for item in itens:
            cod = item.get('codigo') or item.get('Código')
            qtd = int(item.get('quantidade') or item.get('Quantidade') or 0)
            
            # Incrementa o valor na coluna 'reservado' do produto
            cursor.execute('''
                UPDATE produtos 
                SET reservado = reservado + ? 
                WHERE codigo = ?
            ''', (qtd, cod))

        # PASSO 2: Mudar o status para 'Pendente' (ou o status que sua aba de separação usa)
        cursor.execute("UPDATE pedidos SET status = 'Pendente' WHERE id_pedido = ?", (id_pedido,))
        
        conn.commit()
        return True, f"Pedido #{id_pedido} enviado com sucesso e estoque reservado!"
        
    except Exception as e:
        conn.rollback()
        return False, str(e)
    finally:
        conn.close()


def registrar_pedido(cliente, solicitante, separador, produtos_lote, urgente):
    """
    REGISTRA O PEDIDO no DB como 'Rascunho'. 
    Nesta etapa, NÃO reserva estoque ainda.
    """
    import sqlite3
    from datetime import datetime
    import json

    pedido_id = None
    data_criacao = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    
    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()
    
    try:
        # 1. Serializa a lista de itens
        itens_json = json.dumps(produtos_lote)
        
        # 2. Insere o pedido na tabela de pedidos com status 'Rascunho'
        cursor.execute('''
            INSERT INTO pedidos (data_criacao, cliente, solicitante, separador, status, itens, urgente)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (data_criacao, cliente, solicitante, separador, "Rascunho", itens_json, urgente))
        
        pedido_id = cursor.lastrowid 
        
        # --- O PASSO 3 FOI REMOVIDO DAQUI ---
        # A reserva agora só acontece quando você clica em "ENVIAR PARA SEPARAÇÃO"
        # através da função promover_pedido_para_pendente.

        conn.commit()
        return True, f"Pedido {pedido_id} registrado como Rascunho com sucesso."
        
    except Exception as e:
        if conn:
            conn.rollback()
        return False, f"Erro ao registrar o rascunho: {e}"
        
    finally:
        if conn:
            conn.close()

# Modifique a definição da função para aceitar 2 argumentos
def separar_pedido(pedido_id, separador_nome, solicitante_nome):
    import sqlite3 
    from datetime import datetime
    import json
    
    conn = None 
    try:
        conn = sqlite3.connect('estoque.db')
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        # 1. Busca dados do pedido
        cursor.execute("SELECT Itens, Status, Cliente FROM pedidos WHERE ID_Pedido = ?", (pedido_id,))
        pedido_data = cursor.fetchone()
        
        if not pedido_data or pedido_data['Status'] != 'Pendente':
            return False, "Pedido não encontrado ou já processado."

        itens_lista = json.loads(pedido_data['Itens'])
        cliente_do_pedido = pedido_data['Cliente']
        data_separacao = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        # 2. BAIXA FÍSICA + LIMPEZA DE RESERVA + LOG
        for item in itens_lista:
            codigo = item['codigo']
            qtd = item['quantidade']
            
            # Tira do físico e limpa a reserva
            cursor.execute("""
                UPDATE produtos 
                SET estoque_atual = estoque_atual - ?,
                    reservado = MAX(0, reservado - ?)
                WHERE codigo = ?
            """, (qtd, qtd, codigo))
            
            # Registra no histórico (Saída)
            registrar_saida_log(cursor, codigo, qtd, cliente_do_pedido, solicitante_nome, separador_nome)

        # 3. Atualiza status do pedido
        cursor.execute("""
            UPDATE pedidos 
            SET Status = ?, Separador = ?, Solicitante = ?, Data_Separacao = ?
            WHERE ID_Pedido = ?
        """, ('Separado', separador_nome, solicitante_nome, data_separacao, pedido_id))

        conn.commit()
        return True, f"Pedido {pedido_id} separado e estoque baixado."
        
    except Exception as e:
        if conn: conn.rollback()
        return False, f"Erro na separação: {e}"
    finally:
        if conn: conn.close()
            
def verificar_estoque(codigo_produto, quantidade_solicitada, pedido_id_edicao=None):
    """
    Verifica se o saldo disponível (Físico - Reservado) é suficiente.
    Considera a reserva do próprio pedido em caso de edição.
    """
    import sqlite3
    import json
    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()
    
    try:
        # 1. Busca nome, estoque físico e o quanto já está reservado no total
        cursor.execute("SELECT nome, estoque_atual, reservado FROM produtos WHERE codigo = ?", (codigo_produto,))
        produto = cursor.fetchone()
        
        if not produto:
            return False, f"Produto com código '{codigo_produto}' não encontrado.", 0
        
        nome_produto, estoque_fisico, total_reservado = produto
        
        # 2. TRATAMENTO PARA EDIÇÃO:
        # Se estamos editando, precisamos saber quanto este pedido já 'segura' de reserva
        # para não barrarmos o próprio usuário de manter ou alterar o item.
        qtd_ja_reservada_neste_pedido = 0
        if pedido_id_edicao:
            cursor.execute("SELECT Itens FROM pedidos WHERE ID_Pedido = ?", (pedido_id_edicao,))
            pedido = cursor.fetchone()
            if pedido and pedido[0]:
                itens_originais = json.loads(pedido[0])
                for item in itens_originais:
                    if item['codigo'] == codigo_produto:
                        qtd_ja_reservada_neste_pedido = item['quantidade']
                        break

        # 3. CÁLCULO DO SALDO DISPONÍVEL REAL
        # Disponível é o que está livre na prateleira menos o que outros pedidos já reservaram
        saldo_disponivel = (estoque_fisico - total_reservado) + qtd_ja_reservada_neste_pedido
        
        if quantidade_solicitada > saldo_disponivel:
            mensagem = (f"Estoque insuficiente para '{nome_produto}'.\n"
                        f"Saldo Disponível: {saldo_disponivel}\n"
                        f"(Físico: {estoque_fisico} / Reservado: {total_reservado})")
            return False, mensagem, saldo_disponivel
        
        return True, "Estoque OK.", saldo_disponivel
        
    except Exception as e:
        return False, f"Erro ao consultar estoque: {e}", 0
        
    finally:
        conn.close()

def promover_pedido_com_corte_total(pedido_id):
    import sqlite3
    import json
    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()
    
    try:
        # 1. Busca os itens E o nome do cliente antes de tudo
        cursor.execute("SELECT Itens, Cliente FROM pedidos WHERE ID_Pedido = ?", (pedido_id,))
        resultado = cursor.fetchone()
        
        if not resultado:
            return False, "Pedido não encontrado."

        itens_rascunho = json.loads(resultado[0])
        cliente_nome = resultado[1] # <--- DEFINIDO AQUI PARA NÃO DAR ERRO
        
        itens_para_promover = []
        itens_cortados = []

        # 2. Verifica estoque
        for item in itens_rascunho:
            codigo = item['codigo']
            qtd_pedida = item['quantidade']
            
            cursor.execute("SELECT nome, (estoque_atual - reservado) FROM produtos WHERE codigo = ?", (codigo,))
            res_prod = cursor.fetchone()
            
            if res_prod:
                nome_prod, saldo_disponivel = res_prod
                if saldo_disponivel >= qtd_pedida:
                    itens_para_promover.append(item)
                else:
                    # SALVA NO HISTÓRICO (6 valores para 6 colunas, excluindo o ID)
                    cursor.execute("""
                        INSERT INTO historico_faltantes 
                        (data_corte, id_pedido, cliente, codigo_produto, nome_produto, quantidade_faltante)
                        VALUES (datetime('now', 'localtime'), ?, ?, ?, ?, ?)
                    """, (pedido_id, cliente_nome, codigo, nome_prod, qtd_pedida))
                    
                    itens_cortados.append(f"{nome_prod} (Pediu: {qtd_pedida})")
            else:
                itens_cortados.append(f"Cód {codigo} não encontrado.")

        # 3. Finaliza a promoção
        if not itens_para_promover:
            return False, "Corte Total: Nenhum item disponível."

        novos_itens_json = json.dumps(itens_para_promover)
        cursor.execute("UPDATE pedidos SET Itens = ?, Status = 'Pendente' WHERE ID_Pedido = ?", (novos_itens_json, pedido_id))

        for item in itens_para_promover:
            cursor.execute("UPDATE produtos SET reservado = reservado + ? WHERE codigo = ?", (item['quantidade'], item['codigo']))

        conn.commit()
        return True, "Pedido processado! Verifique a aba de Faltantes se houveram cortes."

    except Exception as e:
        conn.rollback()
        return False, f"Erro: {e}"
    finally:
        conn.close()

def atualizar_pedido(pedido_id, itens_json_string, cliente, solicitante, status_atual):
    import sqlite3
    import json
    conn = None
    
    try:
        conn = sqlite3.connect('estoque.db')
        cursor = conn.cursor()
        
        # 1. Busca os itens antigos para poder estornar o estoque/reserva
        cursor.execute("SELECT Itens FROM pedidos WHERE ID_Pedido = ?", (pedido_id,))
        resultado = cursor.fetchone()
        itens_antigos = json.loads(resultado[0]) if resultado and resultado[0] else []
        itens_novos = json.loads(itens_json_string)

        # --- LÓGICA DE ESTOQUE ---
        if status_atual in ('Separado', 'Expedição'):
            # Ajusta Estoque Físico
            for item in itens_antigos:
                cursor.execute("UPDATE produtos SET estoque_atual = estoque_atual + ? WHERE codigo = ?", 
                               (item.get('quantidade', 0), item.get('codigo')))
            for item in itens_novos:
                cursor.execute("UPDATE produtos SET estoque_atual = estoque_atual - ? WHERE codigo = ?", 
                               (item.get('quantidade', 0), item.get('codigo')))

        elif status_atual == 'Pendente':
            # Ajusta apenas Reserva
            for item in itens_antigos:
                cursor.execute("UPDATE produtos SET reservado = MAX(0, reservado - ?) WHERE codigo = ?", 
                               (item.get('quantidade', 0), item.get('codigo')))
            for item in itens_novos:
                cursor.execute("UPDATE produtos SET reservado = reservado + ? WHERE codigo = ?", 
                               (item.get('quantidade', 0), item.get('codigo')))
        
        # Se for 'Rascunho', o código pula os IFs acima e não mexe em nenhum estoque!

        # 2. Atualiza os dados do pedido
        cursor.execute("""
            UPDATE pedidos
            SET Itens = ?, Cliente = ?, Solicitante = ?
            WHERE ID_Pedido = ?
        """, (itens_json_string, cliente, solicitante, pedido_id))
        
        conn.commit()
        return True, "Pedido atualizado com sucesso!"
        
    except Exception as e:
        if conn: conn.rollback()
        return False, f"Erro interno: {e}"
    finally:
        if conn: conn.close()
            
# Função para mover pedidos para aba expedição        
def mover_pedido_para_expedicao(pedido_id, responsavel_expedicao): 
    import sqlite3
    from datetime import datetime
    
    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()
    
    try:
        data_expedicao = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        cursor.execute("""
            UPDATE pedidos 
            SET Status = ?, Data_Expedicao = ?, Responsavel_Expedicao = ?
            WHERE ID_Pedido = ?
        """, ('Expedição', data_expedicao, responsavel_expedicao, pedido_id))

        conn.commit()
        return True, f"Pedido {pedido_id} movido para Expedição."
    except Exception as e:
        if conn: conn.rollback()
        return False, f"Erro: {e}" 
    finally:
        conn.close()

# Função para consultar pedidos, filtrando por status ('Pendente', 'Mesanino' ou 'Expedição'
def consultar_pedidos(status=None, filtro_cliente=None, filtro_data_finalizacao=None):
    """
    Consulta pedidos na tabela 'pedidos', filtrando por status, cliente e data de finalização.
    """
    conn = sqlite3.connect('estoque.db')
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    # 1. Monta a Query Base
    query = """
    SELECT 
        id_pedido, data_criacao, cliente, solicitante, separador, status, itens, 
        data_expedicao, data_separacao, responsavel_expedicao, data_finalizacao 
    FROM pedidos
    """
    
    where_clauses = []
    params = []
    
    # --- LÓGICA DE FILTRAGEM ---
    
    # Filtro 1: Status (Mesanino ou Expedição)
    if status:
        where_clauses.append("status = ?")
        params.append(status)
        
    # Filtro 2: Cliente (Busca parcial)
    if filtro_cliente: # Agora filtro_cliente será None se o campo estiver vazio
        where_clauses.append("cliente LIKE ?")
        params.append(f'%{filtro_cliente}%')
        
    # Filtro 3: Data de Finalização (Busca Exata para um Dia)
    if filtro_data_finalizacao:
        try:
            # Converte a data do DateEntry (DD/MM/AAAA) para um formato seguro no SQL (AAAA-MM-DD)
            data_obj = datetime.strptime(filtro_data_finalizacao, '%d/%m/%Y')
            data_iso_inicial = data_obj.strftime('%Y-%m-%d 00:00:00')
            data_iso_final = data_obj.strftime('%Y-%m-%d 23:59:59')
            
            # **NOTA CRÍTICA:** Se a sua coluna Data_Finalizacao está em DD/MM/AAAA,
            # precisamos que o SQL compare como string entre o início e fim do dia.
            
            # Usaremos o formato de comparação de strings que você usa no banco (DD/MM/AAAA)
            where_clauses.append("data_finalizacao >= ? AND data_finalizacao <= ?")
            params.append(f"{filtro_data_finalizacao} 00:00:00")
            params.append(f"{filtro_data_finalizacao} 23:59:59")
            
        except ValueError:
            print(f"Aviso: Formato de data de finalização inválido: {filtro_data_finalizacao}")
            # Não adiciona o filtro se a data for inválida
            pass
    
    # --- FIM DA LÓGICA DE FILTRAGEM ---
    
    # Adiciona a cláusula WHERE
    if where_clauses:
        query += " WHERE " + " AND ".join(where_clauses)
        
    query += " ORDER BY data_criacao DESC"
    
    cursor.execute(query, params)
    pedidos_db = cursor.fetchall()
    conn.close()
    
    resultado = []
    for p in pedidos_db:
        # --- 2. ADICIONE AS VARIÁVEIS NO DESEMPACOTAMENTO ---
        # ATENÇÃO: Adicionei data_finalizacao também, caso você a use
        id_pedido, data_criacao, cliente, solicitante, separador, status_atual, itens_json, data_expedicao, data_separacao, responsavel_expedicao, data_finalizacao = p
        
        try:
            itens = json.loads(itens_json)
        except json.JSONDecodeError:
            itens = [{"codigo": "ERRO", "quantidade": 0}]
            
        resultado.append({
            "ID_Pedido": id_pedido,
            "Data_Criacao": data_criacao,
            "Cliente": cliente,
            "Solicitante": solicitante,
            "Separador": separador,
            "Status": status_atual,
            "Itens": itens,
            "Data_Expedicao": data_expedicao,
            "Data_Separacao": data_separacao,
            # --- 3. ADICIONE A CHAVE NO DICIONÁRIO DE RETORNO ---
            "Responsavel_Expedicao": responsavel_expedicao,
            "Data_Finalizacao": data_finalizacao
        })
        
    return resultado
    
# Função para consultar o pedido por id
def consultar_pedido_por_id(pedido_id):
    """Consulta um pedido específico pelo seu ID e retorna os dados como um dicionário Python."""
    import sqlite3 
    
    try:
        # 1. Garante que o ID é tratado como INT (essencial para a consulta)
        pedido_id_int = int(pedido_id) 
    except (ValueError, TypeError):
        return None

    conn = sqlite3.connect('estoque.db')
    conn.row_factory = sqlite3.Row  # Mantemos o Row para acessar por nome no backend
    cursor = conn.cursor()
    
    try:
        # 2. Usando o nome da coluna CORRETO: 'id_pedido'
        cursor.execute("SELECT * FROM pedidos WHERE id_pedido = ?", (pedido_id_int,)) 
        
        pedido_row = cursor.fetchone()
        
        if pedido_row:
            # 3. SOLUÇÃO FINAL: Converte o objeto Row para um DICTIONARY Python Puro.
            # Isso garante que as chaves (minúsculas) sejam acessadas sem erro no frontend.
            pedido_dict = {k: pedido_row[k] for k in pedido_row.keys()} 
            return pedido_dict
        else:
            return None
            
    except Exception as e:
        print(f"ERRO DE BANCO DE DADOS na consulta: {e}")
        return None
        
    finally:
        conn.close()
        
# Função para excluir pedidos        
def excluir_pedido(pedido_id):
    """Exclui um pedido da tabela 'pedidos' pelo seu ID."""
    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()
    
    try:
        cursor.execute("DELETE FROM pedidos WHERE ID_Pedido = ?", (pedido_id,))
        
        if cursor.rowcount > 0:
            conn.commit()
            return True, f"Pedido {pedido_id} excluído com sucesso."
        else:
            return False, f"Erro: Pedido {pedido_id} não encontrado."
            
    except Exception as e:
        conn.rollback()
        return False, f"Erro ao excluir pedido: {e}"
        
    finally:
        conn.close()

# Função para estornar pedidos (pedidos cancelados)
def estornar_pedido(pedido_id, responsavel_estorno):
    import sqlite3
    import json
    from datetime import datetime
    
    conn = None
    try:
        conn = sqlite3.connect('estoque.db')
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        cursor.execute("SELECT Itens, Status, Cliente FROM pedidos WHERE ID_Pedido = ?", (pedido_id,))
        pedido_data = cursor.fetchone()
        
        if not pedido_data:
            return False, f"Erro: Pedido {pedido_id} não encontrado."
            
        status_atual = pedido_data['Status']
        itens_lista = json.loads(pedido_data['Itens'])
        cliente_do_pedido = pedido_data['Cliente']

        if status_atual == 'Cancelado':
            return False, "Este pedido já está cancelado."

        # --- AJUSTE NA LÓGICA DE DEVOLUÇÃO ---
        for item in itens_lista:
            codigo = item['codigo']
            qtd = item['quantidade']

            # Se o status for Separado OU Expedição, o físico JÁ SAIU.
            # Então precisamos DEVOLVER ao estoque físico.
            if status_atual in ('Separado', 'Expedição'):
                cursor.execute("""
                    UPDATE produtos 
                    SET estoque_atual = estoque_atual + ? 
                    WHERE codigo = ?
                """, (qtd, codigo))
                tipo_log = "Estorno Físico"
            
            # Se for Pendente, o físico ainda está lá, só existia a RESERVA.
            # Então apenas LIMPAMOS a reserva.
            else:
                cursor.execute("""
                    UPDATE produtos 
                    SET reservado = MAX(0, reservado - ?) 
                    WHERE codigo = ?
                """, (qtd, codigo))
                tipo_log = "Estorno Reserva"

            # 3. REGISTRO NO HISTÓRICO
            timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            obs = f"ESTORNO ({tipo_log}): Pedido {pedido_id} ({status_atual}). Cliente: {cliente_do_pedido}."
            
            cursor.execute("""
                INSERT INTO historico (data_registro, codigo_produto, tipo, quantidade, responsavel, observacao)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (timestamp, codigo, "Cancelamento", qtd, responsavel_estorno, obs))

        # 4. ATUALIZA STATUS DO PEDIDO
        cursor.execute("""
            UPDATE pedidos 
            SET Status = 'Cancelado', Data_Finalizacao = ? 
            WHERE ID_Pedido = ?
        """, (datetime.now().strftime("%d/%m/%Y %H:%M:%S"), pedido_id))

        conn.commit()
        return True, f"Pedido {pedido_id} cancelado. Estoque devolvido (Status era {status_atual})."

    except Exception as e:
        if conn: conn.rollback()
        return False, f"Erro ao estornar: {e}"
    finally:
        if conn: conn.close()

def finalizar_pedido(pedido_id):
    """
    Apenas conclui o pedido. 
    A baixa física e limpeza de reserva já foram feitas na SEPARAÇÃO.
    """
    import sqlite3
    from datetime import datetime
    
    conn = sqlite3.connect('estoque.db')
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    try:
        # 1. Busca o status atual para evitar finalização dupla
        cursor.execute("SELECT Status FROM pedidos WHERE ID_Pedido = ?", (pedido_id,))
        pedido = cursor.fetchone()
        
        if not pedido:
            return False, "Pedido não encontrado."
        
        if pedido['Status'] == 'Concluído':
            return False, "Este pedido já foi finalizado anteriormente."

        data_finalizacao = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        # 2. APENAS MUDA O STATUS
        # Não mexemos mais na tabela 'produtos' aqui, pois o estoque já saiu na separação.
        cursor.execute("""
            UPDATE pedidos 
            SET Status = ?, Data_Finalizacao = ?
            WHERE ID_Pedido = ?
        """, ('Concluído', data_finalizacao, pedido_id))

        conn.commit()
        return True, f"Pedido {pedido_id} finalizado com sucesso e movido para o histórico."
        
    except Exception as e:
        if conn:
            conn.rollback()
        return False, f"Erro ao finalizar pedido: {e}"
        
    finally:
        if conn:
            conn.close()    

# --- Adicionada a nova função para atualizar o estoque mínimo (EXISTENTE) ---
def atualizar_estoque_minimo(codigo, novo_estoque_minimo):
    """
    Atualiza o estoque mínimo de um produto no banco de dados.
    """
    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()
    
    try:
        # Verifica se o produto existe
        cursor.execute("SELECT id_produto FROM produtos WHERE codigo = ?", (codigo,))
        produto = cursor.fetchone()
        
        if not produto:
            return False, "Erro: Produto com este código não encontrado."
        
        id_produto = produto[0]
        
        cursor.execute("UPDATE produtos SET estoque_minimo = ? WHERE id_produto = ?", (novo_estoque_minimo, id_produto))
        conn.commit()
        return True, "Estoque mínimo atualizado com sucesso!"
    except Exception as e:
        conn.rollback()
        return False, f"Erro ao atualizar o estoque mínimo: {e}"
    finally:
        conn.close()

def exportar_faltantes_consolidado_pdf(dados_treeview, caminho_salvar):
    try:
        # Importações no início da função para melhor performance
        import re
        from datetime import datetime
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
        from reportlab.lib.pagesizes import A4
        from reportlab.lib import colors
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

        doc = SimpleDocTemplate(caminho_salvar, pagesize=A4)
        elementos = []
        styles = getSampleStyleSheet()

        # Criamos um estilo específico para a célula do produto permitir quebra de linha
        estilo_celula_produto = ParagraphStyle(
            'EstiloProduto',
            parent=styles['Normal'],
            fontSize=9,
            leading=11,  # Espaçamento entre linhas dentro da célula
            wordWrap='LTR'
        )

        # Cabeçalho
        elementos.append(Paragraph("Relatório de Produtos Faltantes (Consolidado)", styles['Title']))
        elementos.append(Paragraph(f"Emitido em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", styles['Normal']))
        elementos.append(Spacer(1, 20))

        consolidado = {}
        total_perdido = 0

        for item in dados_treeview:
            # Proteção: Verifica se a linha não está vazia
            if not item or len(item) < 6:
                continue
                
            try:
                # Pegando os dados pelos índices: 3=Código, 4=Produto, 5=Qtd. Perdida
                cod = str(item[3]).strip()
                nome = str(item[4]).strip()
                
                # Captura e limpeza da Quantidade
                raw_value = str(item[5]).strip()
                clean_value = re.sub(r'[^\d,.]', '', raw_value)
                
                if not clean_value:
                    continue
                
                # Tratamento de formato brasileiro para float
                if ',' in clean_value:
                    clean_value = clean_value.replace('.', '').replace(',', '.')
                
                qtd = float(clean_value)
                
                if qtd > 0:
                    total_perdido += qtd
                    if cod in consolidado:
                        consolidado[cod]['qtd'] += qtd
                    else:
                        consolidado[cod] = {'nome': nome, 'qtd': qtd}
            except Exception as e:
                print(f"Erro ao converter linha: {item} -> {e}")
                continue

        if not consolidado:
            return False, "Nenhum dado numérico válido encontrado nas colunas da tabela."

        # Montagem da Tabela
        # O cabeçalho permanece como String simples
        dados_tabela = [["Código", "Produto", "Qtd. Total"]]
        
        # Ordenação por nome
        for cod in sorted(consolidado.keys(), key=lambda x: consolidado[x]['nome']):
            # Transformamos o nome em Paragraph para aplicar a quebra de linha automática
            nome_com_quebra = Paragraph(consolidado[cod]['nome'], estilo_celula_produto)
            
            dados_tabela.append([
                cod, 
                nome_com_quebra, 
                f"{int(consolidado[cod]['qtd'])}"
            ])

        # Linha de Totalizador Final
        dados_tabela.append(["", "TOTAL GERAL", f"{int(total_perdido)}"])

        # Estilo da Tabela
        # O alinhamento vertical 'MIDDLE' ajuda a manter o texto centralizado se a linha crescer
        tabela = Table(dados_tabela, colWidths=[80, 310, 80])
        tabela.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#333333')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), # Alinhamento vertical centralizado
            ('ALIGN', (1, 0), (1, -1), 'LEFT'), 
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.white, colors.HexColor('#F2F2F2')]),
            ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
        ]))
        
        elementos.append(tabela)
        doc.build(elementos)
        return True, None

    except Exception as e:
        return False, f"Erro inesperado: {str(e)}"


if __name__ == '__main__':
    # Remover o banco de dados para iniciar do zero
    if os.path.exists('estoque.db'):
        os.remove('estoque.db')
    
    print("Criando tabelas...")
    criar_tabelas()
    
    print("\nAdicionando produtos de teste...")
    adicionar_produto('ABC-101', 'Fone de Ouvido', 50, 10)
    adicionar_produto('XYZ-202', 'Carregador Portátil', 5, 20)
    
    print("\n--- Teste de Saída em Lote (Novo Pedido) ---")
    produtos_para_separar = [
        {'codigo': 'ABC-101', 'quantidade': 10},
        {'codigo': 'XYZ-202', 'quantidade': 3}
    ]
    
    sucesso, mensagem = registrar_pedido("Empresa Teste SA", "João", "David", produtos_para_separar)
    print(f"Registro do Pedido: {mensagem}")

    print("\n--- Produtos com estoque baixo ---")
    estoque_baixo = consultar_estoque_baixo()
    for produto in estoque_baixo:
        print(f"Produto: {produto['Nome']} ({produto['Código']}) - Estoque Atual: {produto['Estoque Atual']} (Mínimo: {produto['Estoque Mínimo']})")
    
    print("\n--- Pedidos no Mesanino (Após Registro) ---")
    pedidos_mesanino = consultar_pedidos("Mesanino")
    for pedido in pedidos_mesanino:
        print(f"ID: {pedido['ID_Pedido']} | Cliente: {pedido['Cliente']} | Itens: {pedido['Itens']}")
        
    # Teste de movimentação
    if pedidos_mesanino:
        id_teste = pedidos_mesanino[0]['ID_Pedido']
        sucesso, mensagem = mover_pedido_para_expedicao(id_teste)
        print(f"\nMovimentação do Pedido {id_teste}: {mensagem}")
        
    print("\n--- Pedidos na Expedição (Após Movimentação) ---")
    pedidos_expedicao = consultar_pedidos("Expedição")
    for pedido in pedidos_expedicao:
        print(f"ID: {pedido['ID_Pedido']} | Cliente: {pedido['Cliente']} | Data Expedição: {pedido['Data_Expedicao'].split(' ')[0]}")