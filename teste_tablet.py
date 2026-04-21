from flask import Flask, render_template_string, redirect, url_for, request
import sqlite3
import json

app = Flask(__name__)

# CSS para uma aparência de aplicativo moderno
ESTILO_CSS = '''
<style>
    body { font-family: sans-serif; background: #f4f7f6; margin: 0; padding: 15px; }
    .card { background: white; padding: 15px; border-radius: 10px; margin-bottom: 15px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); border-left: 8px solid #3498db; }
    .urgente { border-left-color: #e74c3c; background: #fff5f5; }
    .btn { display: block; width: 100%; padding: 12px; margin-top: 10px; border: none; border-radius: 5px; font-weight: bold; text-align: center; text-decoration: none; cursor: pointer; }
    .btn-azul { background: #3498db; color: white; }
    .btn-verde { background: #27ae60; color: white; }
    .btn-voltar { background: #95a5a6; color: white; margin-top: 20px; }
    .item-linha { display: flex; justify-content: space-between; padding: 10px 0; border-bottom: 1px solid #eee; }
    .qtd { font-weight: bold; color: #2c3e50; font-size: 1.2em; }
</style>
'''

# --- TELAS (HTML) ---

TELA_PRINCIPAL = ESTILO_CSS + '''
<h1>📦 Gestão de Fluxo</h1>

<div style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 20px;">
    <a href="/novo_pedido" class="btn btn-verde" style="text-decoration: none; text-align: center;">➕ Novo Pedido</a>
    <a href="/triagem" class="btn btn-azul" style="text-decoration: none; text-align: center; background-color: #8e44ad;">📋 Triagem</a>
    <a href="/mezanino" class="btn btn-azul" style="text-decoration: none; text-align: center;">🏗️ Pedidos no Mezanino</a>
    <a href="/tela_expedicao" class="btn btn-azul" style="text-decoration: none; text-align: center; background-color: #2c3e50;">🚚 Pedidos na Expedição</a>
</div>

<h2 style="border-bottom: 2px solid #eee; padding-bottom: 10px;">📋 Pendentes para Separar</h2>

{% for pedido in pedidos %}
<div class="card {{ 'urgente' if pedido[3] == 'Sim' else '' }}">
    <div style="display: flex; justify-content: space-between; align-items: center;">
        <div>
            <strong>Pedido #{{ pedido[0] }}</strong> - {{ pedido[1] }}<br>
            <span style="font-size: 0.9em; color: #7f8c8d;">Status: {{ pedido[2] }}</span>
        </div>
        
        {% if pedido[3] == 'Sim' %}
        <span style="background: #e74c3c; color: white; padding: 5px 10px; border-radius: 5px; font-size: 0.8em; font-weight: bold;">🔥 URGENTE</span>
        {% endif %}
    </div>

    <div style="margin-top: 10px;">
        <a href="/pedido/{{ pedido[0] }}" class="btn btn-azul" style="display: block; text-align: center; text-decoration: none;">🔍 Abrir Lista de Itens</a>
    </div>
</div>
{% endfor %}
'''

TELA_DETALHES = ESTILO_CSS + '''
<h1>📋 Lista de Separação</h1>
<div class="card" style="border-left-color: #2ecc71;">
    <strong>Pedido:</strong> #{{ id_pedido }}<br>
    <strong>Cliente:</strong> {{ cliente }}
</div>
<div class="card" style="background: #fff3cd; border: 1px solid #ffeeba; margin-bottom: 15px;">
    <label style="font-weight: bold; display: block; margin-bottom: 5px;">👤 Selecione o Responsável da Separação:</label>
    <select name="responsavel_separacao" required style="width: 100%; padding: 12px; border-radius: 5px;">
        <option value="" disabled selected>Clique para selecionar...</option>
        {% for nome in responsaveis %}
            <option value="{{ nome }}">{{ nome }}</option>
        {% endfor %}
    </select>
</div>

{% for item in itens %}
<div class="card" style="border-left-width: 5px;">
    <div style="font-size: 1.1em; color: #2c3e50; margin-bottom: 5px;">
        <strong>{{ item.nome }}</strong>
    </div>
    <div style="display: flex; justify-content: space-between; color: #7f8c8d;">
        <span>Cód: {{ item.codigo }}</span>
        <span style="font-size: 1.2em; color: #e67e22; font-weight: bold;">Qtd: {{ item.quantidade }}</span>
    </div>
</div>
{% endfor %}

<form action="/finalizar_separacao/{{ id_pedido }}" method="POST">
    <button type="submit" class="btn btn-verde">✅ Finalizar Separação</button>
</form>

<a href="/" class="btn btn-voltar">⬅ Voltar para a Lista</a>
'''

# --- TELA DE CRIAÇÃO DE PEDIDO (HTML) ---
TELA_NOVO_PEDIDO = ESTILO_CSS + '''
<div class="container">
    <h1>📝 Novo Pedido</h1>
    
    <div class="card">
        <input type="text" id="cliente" placeholder="Nome do Cliente" style="width: 100%; padding: 12px; margin-bottom: 15px; border: 1px solid #ccc; border-radius: 5px;">
        
        <label style="font-weight: bold; display: block; margin-bottom: 5px;">Solicitante:</label>
        <select id="solicitante" style="width: 100%; padding: 12px; margin-bottom: 15px; border: 1px solid #ccc; border-radius: 5px; background: white; font-size: 1em;">
            <option value="" disabled selected>Selecione quem solicitou...</option>
            {% for nome in solicitantes %}
                <option value="{{ nome }}">{{ nome }}</option>
            {% endfor %}
        </select>
        
        <label style="font-size: 1.1em; cursor: pointer;">
            <input type="checkbox" id="urgente" style="transform: scale(1.5); margin-right: 10px;"> 🔥 Pedido Urgente
        </label>
    </div>

    <div class="card" style="background: #f0f4f8; border: 1px solid #d1d9e0;">
        <h3 style="margin-top: 0;">📦 Adicionar Itens</h3>
        <div id="nome_prod_preview" style="font-weight: bold; color: #2c3e50; margin-bottom: 8px; min-height: 20px;">-</div>
        
        <div style="display: flex; gap: 10px;">
            <input type="text" id="input_codigo" placeholder="Código/Bip" style="flex: 2; padding: 12px;">
            <input type="number" id="input_qtd" value="1" min="1" style="flex: 1; padding: 12px;">
        </div>
        <button type="button" onclick="adicionarItem()" class="btn btn-azul" style="width: 100%; margin-top: 15px; font-weight: bold;">➕ ADICIONAR ITEM</button>
    </div>

    <h3>Itens no Pedido</h3>
    <div id="lista_carrinho"></div>

    <button onclick="finalizarPedido()" class="btn btn-verde" style="width: 100%; height: 70px; font-size: 1.4em; margin-top: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
        💾 SALVAR PEDIDO
    </button>
    
    <a href="/" class="btn btn-voltar" style="display: block; text-align: center; margin-top: 20px;">⬅ Voltar</a>
</div>

<script>
let itensPedido = [];

// Busca automática do nome ao digitar/bipar o código
document.getElementById('input_codigo').addEventListener('change', function() {
    let cod = this.value;
    if(cod) {
        fetch('/buscar_produto/' + cod)
            .then(res => res.json())
            .then(data => {
                document.getElementById('nome_prod_preview').innerText = data.nome;
            });
    }
});

function adicionarItem() {
    const inputCod = document.getElementById('input_codigo');
    const inputQtd = document.getElementById('input_qtd');
    const previewNome = document.getElementById('nome_prod_preview').innerText;
    
    if(!inputCod.value || !inputQtd.value || previewNome.includes("não encontrado") || previewNome === "-") {
        alert("Por favor, insira um código válido e a quantidade.");
        return;
    }

    itensPedido.push({
        codigo: inputCod.value,
        nome: previewNome,
        quantidade: parseInt(inputQtd.value)
    });
    
    atualizarTabela();
    
    // Limpa para o próximo
    inputCod.value = '';
    inputQtd.value = '1';
    document.getElementById('nome_prod_preview').innerText = '-';
    inputCod.focus();
}

function atualizarTabela() {
    let html = "";
    itensPedido.forEach((item, index) => {
        html += `
        <div class="card" style="display: flex; justify-content: space-between; align-items: center; padding: 10px; margin-bottom: 5px;">
            <div>
                <strong>${item.nome}</strong><br>
                <small>Cod: ${item.codigo} | Qtd: ${item.quantidade}</small>
            </div>
            <button onclick="removerItem(${index})" style="background:#ff4444; color:white; border:none; border-radius:5px; padding: 8px 12px;">X</button>
        </div>`;
    });
    document.getElementById('lista_carrinho').innerHTML = html;
}

function removerItem(index) {
    itensPedido.splice(index, 1);
    atualizarTabela();
}

function finalizarPedido() {
    const cliente = document.getElementById('cliente').value;
    const solicitante = document.getElementById('solicitante').value;
    const urgente = document.getElementById('urgente').checked ? "Sim" : "Não";

    if(!cliente || !solicitante || itensPedido.length === 0) {
        alert("Preencha o cliente, selecione o solicitante e adicione itens!");
        return;
    }

    fetch('/salvar_pedido_novo', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({
            cliente: cliente,
            solicitante: solicitante,
            urgente: urgente,
            itens: itensPedido
        })
    }).then(() => window.location.href = "/");
}
</script>
'''

# --- TELA MEZANINO (HTML) ---
TELA_MEZANINO = ESTILO_CSS + '''
<div class="container">
    <h1>📦 Mezanino / Separados</h1>
    <p>Pedidos aguardando envio para a expedição</p>
    
    {% for pedido in pedidos %}
    <div class="card" style="border-left: 5px solid #f1c40f;"> <strong>Pedido #{{ pedido[0] }}</strong> - {{ pedido[1] }}<br>
        <span style="font-size: 0.9em; color: #666;">Aguardando conferência</span>
        
        <div style="margin-top: 10px;">
            <a href="/enviar_expedicao/{{ pedido[0] }}" class="btn btn-verde" 
               style="display: block; text-align: center; text-decoration: none; font-weight: bold;">
               🚚 ENVIAR PARA EXPEDIÇÃO
            </a>
        </div>
    </div>
    {% endfor %}
    
    <a href="/" class="btn btn-voltar">⬅ Voltar</a>
</div>
'''

# --- TELA EXPEDIÇÃO (HTML) ---
TELA_EXPEDICAO = ESTILO_CSS + '''
<div class="container">
    <h1>🚚 Pedidos na Expedição</h1>
    <p>Prontos para entrega ao cliente ou carregamento.</p>
    
    {% for pedido in pedidos %}
    <div class="card" style="border-left: 5px solid #2ecc71;">
        <div style="display: flex; justify-content: space-between; align-items: flex-start;">
            <div>
                <strong>Pedido #{{ pedido[0] }}</strong><br>
                <span style="font-size: 1.1em;">👤 {{ pedido[1] }}</span>
            </div>
            <a href="/pedido/{{ pedido[0] }}" class="btn btn-azul" style="padding: 5px 15px;">🔍 Itens</a>
        </div>
        
        <div style="margin-top: 15px; display: grid; grid-template-columns: 1fr;">
            <a href="/dar_baixa/{{ pedido[0] }}" 
               onclick="return confirm('Confirmar entrega do pedido #{{ pedido[0] }}?')"
               class="btn" 
               style="background-color: #34495e; color: white; text-align: center; text-decoration: none; padding: 15px; border-radius: 8px; font-weight: bold;">
               ✅ DAR BAIXA (ENTREGUE)
            </a>
        </div>
    </div>
    {% endfor %}
    
    <a href="/" class="btn btn-voltar" style="margin-top: 20px;">⬅ Voltar ao Início</a>
</div>
'''

# --- TELA DE TRIAGEM (HTML) ---
TELA_TRIAGEM = ESTILO_CSS + '''
<div class="container">
    <h1>📋 Triagem de Pedidos</h1>
    <p>Selecione um pedido para iniciar a separação.</p>
    
    {% for pedido in pedidos %}
    <div class="card {{ 'urgente' if pedido[3] == 'Sim' else '' }}">
        <div style="display: flex; justify-content: space-between; align-items: center;">
            <div>
                <strong>Pedido #{{ pedido[0] }}</strong> - {{ pedido[1] }}<br>
                <span style="font-size: 0.9em; color: #666;">Criado em: {{ pedido[2] }}</span>
            </div>
            {% if pedido[3] == 'Sim' %}
                <span style="background: red; color: white; padding: 5px; border-radius: 5px; font-weight: bold;">🔥 URGENTE</span>
            {% endif %}
        </div>
        
        <div style="margin-top: 15px;">
            <a href="/conferir_triagem/{{ pedido[0] }}" class="btn btn-verde" 
                style="display: block; text-align: center; text-decoration: none; font-weight: bold; font-size: 1.1em;">
                🔍 CONFERIR DISPONIBILIDADE
            </a>
        </div>
    </div>
    {% endfor %}
    
    <a href="/" class="btn btn-voltar">⬅ Voltar</a>
</div>
'''

TELA_CHECKLIST_TRIAGEM = ESTILO_CSS + '''
<div class="container">
    <h1>📋 Conferência: Pedido #{{ id_pedido }}</h1>
    <p>Cliente: <strong>{{ cliente }}</strong></p>
    <form action="/confirmar_envio_separacao/{{ id_pedido }}" method="POST">
        {% for item in itens %}
        <div class="card" style="display: flex; justify-content: space-between; align-items: center; padding: 15px;">
            <div>
                <strong>{{ item.nome }}</strong><br>
                Cod: {{ item.codigo }} | Qtd: {{ item.quantidade }}
            </div>
            <div style="text-align: center; color: red;">
                <label style="font-size: 0.8em; font-weight: bold;">FALTA?</label><br>
                <input type="checkbox" name="falta" value="{{ item.codigo }}" style="transform: scale(2);">
            </div>
        </div>
        {% endfor %}
        
        <button type="submit" class="btn btn-verde" style="width: 100%; height: 70px; font-size: 1.3em; margin-top: 20px;">
            🚀 ENVIAR PARA SEPARAÇÃO
        </button>
    </form>
    <a href="/triagem" class="btn btn-voltar">⬅ Cancelar</a>
</div>
'''

# --- FUNÇÕES DE BANCO ---

def obter_conexao():
    return sqlite3.connect('estoque.db')

def descobrir_tabelas():
    conn = obter_conexao()
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    print("Tabelas encontradas no seu banco:", cursor.fetchall())
    conn.close()

descobrir_tabelas()

@app.route('/')
def index():
    conn = obter_conexao()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT id_pedido, cliente, data_criacao, Urgente 
        FROM pedidos 
        WHERE status = 'Pendente' 
        ORDER BY Urgente DESC, id_pedido ASC
    """)
    pedidos = cursor.fetchall()
    conn.close()
    return render_template_string(TELA_PRINCIPAL, pedidos=pedidos)

@app.route('/pedido/<int:id_pedido>')
def detalhes(id_pedido):
    conn = obter_conexao()
    cursor = conn.cursor()
    
    # 1. Busca os dados do pedido
    cursor.execute("SELECT itens, cliente FROM pedidos WHERE id_pedido = ?", (id_pedido,))
    resultado = cursor.fetchone()
    
    if not resultado:
        conn.close()
        return "Pedido não encontrado."

    texto_json = resultado[0]
    cliente = resultado[1]
    
    # 2. BUSCA A LISTA DE RESPONSÁVEIS PARA O DROPDOWN
    cursor.execute("SELECT nome FROM cadastros_auxiliares WHERE tipo = 'RESPONSAVEL' ORDER BY nome ASC")
    lista_responsaveis = [linha[0] for linha in cursor.fetchall()]

    try:
        lista_itens_bruta = json.loads(texto_json)
    except:
        conn.close()
        return "Erro ao ler itens do pedido."

    lista_final = []
    for item in lista_itens_bruta:
        codigo = item.get('codigo') or item.get('Código')
        quantidade = item.get('quantidade') or item.get('Quantidade')
        
        cursor.execute("SELECT nome FROM produtos WHERE codigo = ?", (str(codigo),))
        prod_resultado = cursor.fetchone()
        
        nome_produto = prod_resultado[0] if prod_resultado else "PRODUTO NÃO ENCONTRADO"
        
        lista_final.append({
            'codigo': codigo,
            'nome': nome_produto,
            'quantidade': quantidade
        })

    conn.close()
    
    # 3. PASSAMOS 'responsaveis' PARA O TEMPLATE
    return render_template_string(TELA_DETALHES, 
                                  id_pedido=id_pedido, 
                                  cliente=cliente, 
                                  itens=lista_final,
                                  responsaveis=lista_responsaveis)

@app.route('/finalizar_separacao/<int:id_pedido>', methods=['POST'])
def finalizar_separacao(id_pedido):
    import json
    from datetime import datetime
    
    # CAPTURA O NOME SELECIONADO NO DROPDOWN
    responsavel_selecionado = request.form.get('responsavel_separacao')
    if not responsavel_selecionado:
        responsavel_selecionado = "Separador Desconhecido"

    conn = obter_conexao()
    cursor = conn.cursor()
    
    cursor.execute("SELECT itens, cliente FROM pedidos WHERE id_pedido = ?", (id_pedido,))
    res = cursor.fetchone()
    
    if res:
        itens = json.loads(res[0])
        cliente = res[1]
        data_hoje = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        for item in itens:
            cod = str(item.get('codigo') or item.get('Código'))
            qtd = int(item.get('quantidade') or item.get('Quantidade'))
            
            # Baixa Física e Limpeza de Reserva
            cursor.execute("""
                UPDATE produtos 
                SET quantidade = quantidade - ?,
                    reservado = CASE WHEN COALESCE(reservado, 0) >= ? THEN reservado - ? ELSE 0 END
                WHERE codigo = ?
            """, (qtd, qtd, qtd, cod))

            # GRAVAÇÃO NO HISTÓRICO COM O NOME DO RESPONSÁVEL SELECIONADO
            cursor.execute("""
                INSERT INTO historico (data_registro, codigo_produto, tipo, quantidade, responsavel, observacao)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (data_hoje, cod, 'Saída', qtd, responsavel_selecionado, f"Pedido: {cliente}"))

    # Atualiza o pedido com o status e o nome de quem separou
    cursor.execute("UPDATE pedidos SET status = 'Separado', responsavel = ? WHERE id_pedido = ?", 
                   (responsavel_selecionado, id_pedido))
    
    conn.commit()
    conn.close()
    return redirect('/')

# Tela da Expedição

@app.route('/novo_pedido')
def novo_pedido():
    conn = obter_conexao()
    cursor = conn.cursor()
    
    # Buscamos apenas quem está cadastrado como SOLICITANTE
    cursor.execute("SELECT nome FROM cadastros_auxiliares WHERE tipo = 'SOLICITANTE' ORDER BY nome ASC")
    lista_solicitantes = [linha[0] for linha in cursor.fetchall()]
    
    conn.close()
    # Passamos a lista para o HTML através da variável 'solicitantes'
    return render_template_string(TELA_NOVO_PEDIDO, solicitantes=lista_solicitantes)

@app.route('/salvar_pedido_novo', methods=['POST'])
def salvar_pedido_novo():
    from datetime import datetime
    dados = request.get_json()
    
    cliente = dados.get('cliente')
    solicitante = dados.get('solicitante') # Recebe o solicitante
    urgente = dados.get('urgente')
    itens = dados.get('itens')
    
    data_hoje = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    conn = obter_conexao()
    cursor = conn.cursor()
    
    # Adicionamos 'solicitante' na lista de colunas do INSERT
    cursor.execute("""
        INSERT INTO pedidos (cliente, solicitante, data_criacao, itens, status, Urgente)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (cliente, solicitante, data_hoje, json.dumps(itens), 'Rascunho', urgente))
    
    conn.commit()
    conn.close()
    return {"status": "sucesso"}

@app.route('/buscar_produto/<codigo>')
def buscar_produto(codigo):
    conn = obter_conexao()
    cursor = conn.cursor()
    cursor.execute("SELECT nome FROM produtos WHERE codigo = ?", (str(codigo),))
    res = cursor.fetchone()
    conn.close()
    if res:
        return {"nome": res[0]}
    return {"nome": "⚠️ Produto não encontrado"}
    
# Tela Mezanino

@app.route('/mezanino')
def mezanino():
    conn = obter_conexao()
    cursor = conn.cursor()
    # Busca pedidos com status 'Separado' ou 'Mezanino' (ajuste conforme seu banco)
    cursor.execute("SELECT id_pedido, cliente FROM pedidos WHERE status = 'Separado'")
    pedidos_mezanino = cursor.fetchall()
    conn.close()
    return render_template_string(TELA_MEZANINO, pedidos=pedidos_mezanino)

@app.route('/enviar_expedicao/<int:id_pedido>')
def enviar_expedicao(id_pedido):
    conn = obter_conexao()
    cursor = conn.cursor()
    
    # Atualiza o status para 'Expedicao'
    cursor.execute("UPDATE pedidos SET status = 'Expedicao' WHERE id_pedido = ?", (id_pedido,))
    
    conn.commit()
    conn.close()
    return redirect('/mezanino') # Volta para a tela de mezanino

# Tela Expedição

@app.route('/tela_expedicao')
def tela_expedicao():
    conn = obter_conexao()
    cursor = conn.cursor()
    # Busca pedidos que já foram enviados para a expedição
    cursor.execute("SELECT id_pedido, cliente FROM pedidos WHERE status = 'Expedição'")
    pedidos_exp = cursor.fetchall()
    conn.close()
    return render_template_string(TELA_EXPEDICAO, pedidos=pedidos_exp)

@app.route('/dar_baixa/<int:id_pedido>')
def dar_baixa(id_pedido):
    from datetime import datetime
    conn = obter_conexao()
    cursor = conn.cursor()
    
    # Captura a data e hora atual
    data_conclusao = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    
    # Atualiza o status E a data de finalização
    cursor.execute("""
        UPDATE pedidos 
        SET status = 'Concluído', 
            Data_Finalizacao = ? 
        WHERE id_pedido = ?
    """, (data_conclusao, id_pedido))
    
    conn.commit()
    conn.close()
    return redirect('/tela_expedicao')    

# Tela de Triagem

@app.route('/triagem')
def triagem():
    conn = obter_conexao()
    cursor = conn.cursor()
    
    # O SEGREDO DA PRIORIDADE: 
    # Ordenamos primeiro por 'Urgente' (descendente faz o 'Sim' vir antes do 'Não')
    # e depois pela data de criação
    cursor.execute("""
        SELECT id_pedido, cliente, data_criacao, Urgente 
        FROM pedidos 
        WHERE status = 'Rascunho' 
        ORDER BY Urgente DESC, id_pedido ASC
    """)
    pedidos_triagem = cursor.fetchall()
    conn.close()
    return render_template_string(TELA_TRIAGEM, pedidos=pedidos_triagem)

@app.route('/iniciar_separacao/<int:id_pedido>')
def iniciar_separacao(id_pedido):
    conn = obter_conexao()
    cursor = conn.cursor()
    
    # Muda o status para 'Em Separação' ou outro que você use na tela principal
    cursor.execute("UPDATE pedidos SET status = 'Pendente' WHERE id_pedido = ?", (id_pedido,))
    
    conn.commit()
    conn.close()
    # Após iniciar, ele já cai na tela principal onde estão os pedidos para bipar
    return redirect('/')

@app.route('/conferir_triagem/<int:id_pedido>')
def conferir_triagem(id_pedido):
    conn = obter_conexao()
    cursor = conn.cursor()
    cursor.execute("SELECT itens, cliente FROM pedidos WHERE id_pedido = ?", (id_pedido,))
    resultado = cursor.fetchone()
    
    if not resultado:
        conn.close()
        return "Pedido não encontrado."

    itens = json.loads(resultado[0])
    cliente = resultado[1]
    
    # Busca nomes dos produtos para exibir
    itens_com_nome = []
    for item in itens:
        cod = item.get('codigo') or item.get('Código')
        qtd = item.get('quantidade') or item.get('Quantidade')
        cursor.execute("SELECT nome FROM produtos WHERE codigo = ?", (str(cod),))
        res_nome = cursor.fetchone()
        nome = res_nome[0] if res_nome else "Não cadastrado"
        itens_com_nome.append({'codigo': cod, 'nome': nome, 'quantidade': qtd})

    conn.close()
    return render_template_string(TELA_CHECKLIST_TRIAGEM, id_pedido=id_pedido, cliente=cliente, itens=itens_com_nome)

@app.route('/confirmar_envio_separacao/<int:id_pedido>', methods=['POST'])
def confirmar_envio_separacao(id_pedido):
    from datetime import datetime
    conn = obter_conexao()
    cursor = conn.cursor()
    
    # 1. Busca os dados do pedido
    cursor.execute("SELECT itens, cliente FROM pedidos WHERE id_pedido = ?", (id_pedido,))
    res = cursor.fetchone()
    itens_originais = json.loads(res[0])
    cliente = res[1]
    
    codigos_falta = request.form.getlist('falta')
    itens_que_ficam = []
    data_hoje = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    for item in itens_originais:
        cod = str(item.get('codigo') or item.get('Código'))
        qtd = int(item.get('quantidade') or item.get('Quantidade'))
        
        if cod in codigos_falta:
            # --- LOGICA DE FALTA (Já estava pronta) ---
            cursor.execute("SELECT nome FROM produtos WHERE codigo = ?", (cod,))
            res_p = cursor.fetchone()
            nome_produto = res_p[0] if res_p else "Produto não identificado"
            
            cursor.execute("""
                INSERT INTO historico_faltantes (data_corte, id_pedido, cliente, codigo_produto, nome_produto, quantidade_faltante)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (data_hoje, id_pedido, cliente, cod, nome_produto, qtd))
        else:
            # AUMENTA O RESERVADO (Não mexe no estoque físico ainda)
            cursor.execute("""
                UPDATE produtos 
                SET reservado = COALESCE(reservado, 0) + ? 
                WHERE codigo = ?
            """, (qtd, cod))
    
            itens_que_ficam.append(item)

    # 2. Atualiza o pedido com os itens restantes e muda o status
    novo_json = json.dumps(itens_que_ficam)
    cursor.execute("UPDATE pedidos SET itens = ?, status = 'Pendente' WHERE id_pedido = ?", (novo_json, id_pedido))
    
    conn.commit()
    conn.close()
    return redirect('/')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)