import sqlite3

def criar_tabelas():
    conn = sqlite3.connect('estoque.db')
    cursor = conn.cursor()

    # Tabela de Produtos
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS produtos (
        id_produto INTEGER PRIMARY KEY AUTOINCREMENT,
        codigo TEXT NOT NULL UNIQUE,
        nome TEXT NOT NULL,
        estoque_atual INTEGER NOT NULL,
        estoque_minimo INTEGER,
        media_vendas_mensal REAL,
        preco_compra REAL,
        preco_venda REAL,
        data_ultima_atualizacao TEXT
    )
    ''')

    # Tabela de Entradas
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

    # Tabela de Saídas
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

    conn.commit()
    conn.close()

if __name__ == '__main__':
    criar_tabelas()
    print("Banco de dados 'estoque.db' e tabelas criados com sucesso!")