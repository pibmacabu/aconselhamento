import sqlite3
import json
import os
import secrets
import shutil
import io
from datetime import datetime, timedelta
from functools import wraps

from flask import Flask, request, jsonify, g, send_file
from flask_cors import CORS
import bcrypt
import jwt
from dotenv import load_dotenv
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from werkzeug.utils import secure_filename
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.interval import IntervalTrigger

load_dotenv()

app = Flask(__name__)
CORS(app)

# ======= CONFIGURAÇÕES =======
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'segredo_super_seguro_aconselhamento')
app.config['UPLOAD_FOLDER'] = './uploads'
app.config['BACKUP_FOLDER'] = './backups'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB

# Criar pastas necessárias
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['BACKUP_FOLDER'], exist_ok=True)

PORT = int(os.getenv('PORT', 5000))
DATABASE = './database/aconselhamento.db'
os.makedirs(os.path.dirname(DATABASE), exist_ok=True)   # <--- ESSENCIAL PARA O RENDER

ALLOWED_EXTENSIONS = {'pdf', 'png', 'jpg', 'jpeg', 'gif', 'doc', 'docx', 'txt'}

# ======= BANCO DE DADOS =======
def get_db():
    db = getattr(g, '_database', None)
    if db is None:
        db = g._database = sqlite3.connect(DATABASE)
        db.row_factory = sqlite3.Row
    return db

@app.teardown_appcontext
def close_connection(exception):
    db = getattr(g, '_database', None)
    if db is not None:
        db.close()

def init_db():
    """Cria tabelas e adiciona colunas necessárias (migração)."""
    with app.app_context():
        db = get_db()
        cursor = db.cursor()
        cursor.executescript('''
            CREATE TABLE IF NOT EXISTS conselheiros (
                id INTEGER PRIMARY KEY,
                nome TEXT NOT NULL,
                senha_hash TEXT NOT NULL,
                obs TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS pessoas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                conselheiro_id INTEGER NOT NULL,
                nome TEXT NOT NULL,
                obs TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (conselheiro_id) REFERENCES conselheiros(id) ON DELETE CASCADE
            );
            CREATE TABLE IF NOT EXISTS casais (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                conselheiro_id INTEGER NOT NULL,
                nome_casal TEXT,
                id_homem INTEGER NOT NULL,
                id_mulher INTEGER NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (conselheiro_id) REFERENCES conselheiros(id) ON DELETE CASCADE,
                FOREIGN KEY (id_homem) REFERENCES pessoas(id) ON DELETE CASCADE,
                FOREIGN KEY (id_mulher) REFERENCES pessoas(id) ON DELETE CASCADE
            );
            CREATE TABLE IF NOT EXISTS sessoes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                conselheiro_id INTEGER NOT NULL,
                data TEXT NOT NULL,
                caso_num TEXT NOT NULL,
                sessao_num INTEGER NOT NULL,
                duracao TEXT,
                is_casal INTEGER NOT NULL,
                id_pessoa INTEGER,
                id_casal INTEGER,
                versiculo TEXT,
                anotacao_sessao TEXT,
                tasks TEXT,
                tarefas_anteriores TEXT,
                assuntos_nesta TEXT,
                assuntos_proximas TEXT,
                status TEXT DEFAULT 'realizada',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (conselheiro_id) REFERENCES conselheiros(id) ON DELETE CASCADE
            );
            CREATE TABLE IF NOT EXISTS anexos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sessao_id INTEGER NOT NULL,
                conselheiro_id INTEGER NOT NULL,
                nome_original TEXT NOT NULL,
                nome_arquivo TEXT NOT NULL,
                caminho TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (sessao_id) REFERENCES sessoes(id) ON DELETE CASCADE,
                FOREIGN KEY (conselheiro_id) REFERENCES conselheiros(id) ON DELETE CASCADE
            );
            CREATE TABLE IF NOT EXISTS lembretes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                conselheiro_id INTEGER NOT NULL,
                sessao_id INTEGER,
                titulo TEXT NOT NULL,
                descricao TEXT,
                data_lembrete TEXT NOT NULL,
                concluido INTEGER DEFAULT 0,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (conselheiro_id) REFERENCES conselheiros(id) ON DELETE CASCADE,
                FOREIGN KEY (sessao_id) REFERENCES sessoes(id) ON DELETE CASCADE
            );
            CREATE TABLE IF NOT EXISTS logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                conselheiro_id INTEGER NOT NULL,
                acao TEXT NOT NULL,
                detalhes TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (conselheiro_id) REFERENCES conselheiros(id) ON DELETE CASCADE
            );
            CREATE TABLE IF NOT EXISTS tokens_blacklist (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                token TEXT NOT NULL,
                expira_em TIMESTAMP NOT NULL
            );
            CREATE TABLE IF NOT EXISTS configuracoes (
                chave TEXT PRIMARY KEY,
                valor TEXT NOT NULL,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
        ''')
        # Migrações para tabela sessoes
        colunas_sessoes = [row['name'] for row in cursor.execute("PRAGMA table_info(sessoes)")]
        if 'status' not in colunas_sessoes:
            cursor.execute("ALTER TABLE sessoes ADD COLUMN status TEXT DEFAULT 'realizada'")
        if 'tarefas_anteriores' not in colunas_sessoes:
            cursor.execute("ALTER TABLE sessoes ADD COLUMN tarefas_anteriores TEXT")
        if 'created_at' not in colunas_sessoes:
            cursor.execute("ALTER TABLE sessoes ADD COLUMN created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP")
        if 'updated_at' not in colunas_sessoes:
            cursor.execute("ALTER TABLE sessoes ADD COLUMN updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP")

        # Migrações para lembretes
        colunas_lembretes = [row['name'] for row in cursor.execute("PRAGMA table_info(lembretes)")]
        if 'sessao_id' not in colunas_lembretes:
            cursor.execute("ALTER TABLE lembretes ADD COLUMN sessao_id INTEGER REFERENCES sessoes(id) ON DELETE CASCADE")

        # Configurações padrão
        cursor.execute("INSERT OR IGNORE INTO configuracoes (chave, valor) VALUES ('cor_primaria', '#b58b4b')")
        cursor.execute("INSERT OR IGNORE INTO configuracoes (chave, valor) VALUES ('cor_secundaria', '#2d6a4f')")
        cursor.execute("INSERT OR IGNORE INTO configuracoes (chave, valor) VALUES ('cor_fundo', '#f5f1eb')")
        cursor.execute("INSERT OR IGNORE INTO configuracoes (chave, valor) VALUES ('cor_texto', '#2e2b28')")
        cursor.execute("INSERT OR IGNORE INTO configuracoes (chave, valor) VALUES ('logo_url', '')")
        cursor.execute("INSERT OR IGNORE INTO configuracoes (chave, valor) VALUES ('nome_igreja', 'Igreja Batista')")

        db.commit()

init_db()

# ======= BACKUP AUTOMÁTICO =======
def fazer_backup():
    if not os.path.exists(DATABASE):
        return
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_name = f"acompanhamento_backup_{timestamp}.db"
    backup_path = os.path.join(app.config['BACKUP_FOLDER'], backup_name)
    shutil.copy2(DATABASE, backup_path)
    backups = sorted([f for f in os.listdir(app.config['BACKUP_FOLDER']) if f.endswith('.db')])
    if len(backups) > 30:
        for f in backups[:-30]:
            os.remove(os.path.join(app.config['BACKUP_FOLDER'], f))
    print(f"Backup realizado: {backup_path}")

scheduler = BackgroundScheduler()
scheduler.add_job(fazer_backup, trigger=IntervalTrigger(days=1), next_run_time=datetime.now() + timedelta(hours=1))
scheduler.start()

# ======= AUTENTICAÇÃO =======
def gerar_token(conselheiro_id, refresh=False):
    exp = datetime.utcnow() + (timedelta(days=7) if refresh else timedelta(hours=1))
    payload = {'id': conselheiro_id, 'refresh': refresh, 'exp': exp}
    return jwt.encode(payload, app.config['SECRET_KEY'], algorithm='HS256')

def autenticar_token(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        auth_header = request.headers.get('Authorization')
        if not auth_header:
            return jsonify({'erro': 'Token não fornecido'}), 401
        try:
            token = auth_header.split(' ')[1]
            db = get_db()
            cursor = db.cursor()
            cursor.execute('SELECT 1 FROM tokens_blacklist WHERE token = ?', (token,))
            if cursor.fetchone():
                return jsonify({'erro': 'Token revogado'}), 403
            payload = jwt.decode(token, app.config['SECRET_KEY'], algorithms=['HS256'])
            if payload.get('refresh'):
                return jsonify({'erro': 'Use token de acesso'}), 403
            request.user_id = payload['id']
        except jwt.ExpiredSignatureError:
            return jsonify({'erro': 'Token expirado'}), 403
        except jwt.InvalidTokenError:
            return jsonify({'erro': 'Token inválido'}), 403
        return f(*args, **kwargs)
    return decorated

def log_acao(conselheiro_id, acao, detalhes=None):
    db = get_db()
    cursor = db.cursor()
    cursor.execute('INSERT INTO logs (conselheiro_id, acao, detalhes) VALUES (?, ?, ?)',
                   (conselheiro_id, acao, detalhes))
    db.commit()

def obter_proximo_sessao_num(conselheiro_id, caso_num):
    db = get_db()
    cursor = db.cursor()
    cursor.execute('SELECT MAX(sessao_num) as max_num FROM sessoes WHERE conselheiro_id = ? AND caso_num = ?',
                   (conselheiro_id, caso_num))
    row = cursor.fetchone()
    return (row['max_num'] or 0) + 1

# ======= ROTAS PÚBLICAS =======
@app.route('/api/registrar', methods=['POST'])
def registrar():
    data = request.json
    id_ = data.get('id')
    nome = data.get('nome')
    senha = data.get('senha')
    obs = data.get('obs', '')
    if not id_ or not nome or not senha:
        return jsonify({'erro': 'ID, nome e senha são obrigatórios'}), 400
    if not isinstance(id_, int) or id_ <= 0:
        return jsonify({'erro': 'ID deve ser um número inteiro positivo'}), 400
    db = get_db()
    cursor = db.cursor()
    cursor.execute('SELECT id FROM conselheiros WHERE id = ?', (id_,))
    if cursor.fetchone():
        return jsonify({'erro': 'Este ID já está em uso'}), 400
    senha_hash = bcrypt.hashpw(senha.encode('utf-8'), bcrypt.gensalt(12)).decode('utf-8')
    cursor.execute('INSERT INTO conselheiros (id, nome, senha_hash, obs) VALUES (?, ?, ?, ?)',
                   (id_, nome, senha_hash, obs))
    db.commit()
    log_acao(id_, 'registro', f'Conselheiro {nome} cadastrado')
    return jsonify({'mensagem': 'Conselheiro registrado com sucesso'}), 201

@app.route('/api/login', methods=['POST'])
def login():
    data = request.json
    id_ = data.get('id')
    senha = data.get('senha')
    if not id_ or not senha:
        return jsonify({'erro': 'ID e senha são obrigatórios'}), 400
    db = get_db()
    cursor = db.cursor()
    cursor.execute('SELECT * FROM conselheiros WHERE id = ?', (id_,))
    conselheiro = cursor.fetchone()
    if not conselheiro:
        return jsonify({'erro': 'ID ou senha inválidos'}), 401
    if not bcrypt.checkpw(senha.encode('utf-8'), conselheiro['senha_hash'].encode('utf-8')):
        return jsonify({'erro': 'ID ou senha inválidos'}), 401
    token = gerar_token(conselheiro['id'], refresh=False)
    refresh = gerar_token(conselheiro['id'], refresh=True)
    log_acao(conselheiro['id'], 'login', 'Login realizado')
    return jsonify({
        'token': token,
        'refresh': refresh,
        'conselheiro': {
            'id': conselheiro['id'],
            'nome': conselheiro['nome'],
            'obs': conselheiro['obs']
        }
    })

@app.route('/api/refresh', methods=['POST'])
def refresh_token():
    data = request.json
    refresh_token = data.get('refresh')
    if not refresh_token:
        return jsonify({'erro': 'Refresh token não fornecido'}), 401
    try:
        payload = jwt.decode(refresh_token, app.config['SECRET_KEY'], algorithms=['HS256'])
        if not payload.get('refresh'):
            return jsonify({'erro': 'Token inválido'}), 403
        db = get_db()
        cursor = db.cursor()
        cursor.execute('SELECT 1 FROM tokens_blacklist WHERE token = ?', (refresh_token,))
        if cursor.fetchone():
            return jsonify({'erro': 'Token revogado'}), 403
        novo_token = gerar_token(payload['id'], refresh=False)
        return jsonify({'token': novo_token})
    except jwt.ExpiredSignatureError:
        return jsonify({'erro': 'Refresh token expirado'}), 403
    except jwt.InvalidTokenError:
        return jsonify({'erro': 'Refresh token inválido'}), 403

@app.route('/api/logout', methods=['POST'])
@autenticar_token
def logout():
    token = request.headers.get('Authorization').split(' ')[1]
    db = get_db()
    cursor = db.cursor()
    exp = datetime.utcnow() + timedelta(days=1)
    cursor.execute('INSERT INTO tokens_blacklist (token, expira_em) VALUES (?, ?)', (token, exp))
    db.commit()
    log_acao(request.user_id, 'logout', 'Logout realizado')
    return jsonify({'mensagem': 'Deslogado com sucesso'})

# ======= CONFIGURAÇÕES (CORES E BRANDING) =======
@app.route('/api/configuracoes', methods=['GET'])
@autenticar_token
def get_configuracoes():
    db = get_db()
    cursor = db.cursor()
    cursor.execute('SELECT chave, valor FROM configuracoes')
    rows = cursor.fetchall()
    return jsonify({row['chave']: row['valor'] for row in rows})

@app.route('/api/configuracoes', methods=['POST'])
@autenticar_token
def update_configuracoes():
    data = request.json
    db = get_db()
    cursor = db.cursor()
    for chave, valor in data.items():
        cursor.execute('UPDATE configuracoes SET valor = ?, updated_at = CURRENT_TIMESTAMP WHERE chave = ?',
                       (valor, chave))
    db.commit()
    log_acao(request.user_id, 'atualizar_configuracoes', 'Configurações atualizadas')
    return jsonify({'mensagem': 'Configurações salvas'})

# ======= PESSOAS =======
@app.route('/api/pessoas', methods=['GET'])
@autenticar_token
def listar_pessoas():
    db = get_db()
    cursor = db.cursor()
    cursor.execute('SELECT * FROM pessoas WHERE conselheiro_id = ? ORDER BY nome', (request.user_id,))
    rows = cursor.fetchall()
    return jsonify([dict(row) for row in rows])

@app.route('/api/pessoas', methods=['POST'])
@autenticar_token
def criar_pessoa():
    data = request.json
    nome = data.get('nome')
    obs = data.get('obs', '')
    if not nome:
        return jsonify({'erro': 'Nome é obrigatório'}), 400
    db = get_db()
    cursor = db.cursor()
    cursor.execute('INSERT INTO pessoas (conselheiro_id, nome, obs) VALUES (?, ?, ?)',
                   (request.user_id, nome, obs))
    db.commit()
    log_acao(request.user_id, 'criar_pessoa', f'Pessoa {nome} criada')
    return jsonify({'id': cursor.lastrowid, 'nome': nome, 'obs': obs}), 201

@app.route('/api/pessoas/<int:id>', methods=['DELETE'])
@autenticar_token
def deletar_pessoa(id):
    db = get_db()
    cursor = db.cursor()
    cursor.execute('DELETE FROM pessoas WHERE id = ? AND conselheiro_id = ?', (id, request.user_id))
    db.commit()
    return jsonify({'deletado': cursor.rowcount})

# ======= CASAIS =======
@app.route('/api/casais', methods=['GET'])
@autenticar_token
def listar_casais():
    db = get_db()
    cursor = db.cursor()
    cursor.execute('SELECT * FROM casais WHERE conselheiro_id = ? ORDER BY nome_casal', (request.user_id,))
    rows = cursor.fetchall()
    return jsonify([dict(row) for row in rows])

@app.route('/api/casais', methods=['POST'])
@autenticar_token
def criar_casal():
    data = request.json
    nome_casal = data.get('nome_casal', '')
    id_homem = data.get('id_homem')
    id_mulher = data.get('id_mulher')
    if not id_homem or not id_mulher:
        return jsonify({'erro': 'Selecione marido e esposa'}), 400
    db = get_db()
    cursor = db.cursor()
    cursor.execute('INSERT INTO casais (conselheiro_id, nome_casal, id_homem, id_mulher) VALUES (?, ?, ?, ?)',
                   (request.user_id, nome_casal, id_homem, id_mulher))
    db.commit()
    log_acao(request.user_id, 'criar_casal', f'Casal {nome_casal} criado')
    return jsonify({'id': cursor.lastrowid}), 201

@app.route('/api/casais/<int:id>', methods=['DELETE'])
@autenticar_token
def deletar_casal(id):
    db = get_db()
    cursor = db.cursor()
    cursor.execute('DELETE FROM casais WHERE id = ? AND conselheiro_id = ?', (id, request.user_id))
    db.commit()
    return jsonify({'deletado': cursor.rowcount})

# ======= SESSÕES =======
@app.route('/api/sessoes', methods=['GET'])
@autenticar_token
def listar_sessoes():
    db = get_db()
    cursor = db.cursor()
    page = int(request.args.get('page', 1))
    per_page = int(request.args.get('per_page', 20))
    offset = (page - 1) * per_page
    status_filter = request.args.get('status')
    data_inicio = request.args.get('data_inicio')
    data_fim = request.args.get('data_fim')
    assunto = request.args.get('assunto')

    query = 'SELECT * FROM sessoes WHERE conselheiro_id = ?'
    params = [request.user_id]
    if status_filter:
        query += ' AND status = ?'
        params.append(status_filter)
    if data_inicio:
        query += ' AND data >= ?'
        params.append(data_inicio)
    if data_fim:
        query += ' AND data <= ?'
        params.append(data_fim)
    if assunto:
        query += ' AND assuntos_nesta LIKE ?'
        params.append(f'%{assunto}%')
    query += ' ORDER BY data DESC LIMIT ? OFFSET ?'
    params.extend([per_page, offset])

    cursor.execute(query, params)
    rows = cursor.fetchall()
    resultado = []
    for row in rows:
        d = dict(row)
        d['is_casal'] = bool(d['is_casal'])
        d['tasks'] = json.loads(d['tasks'] or '[]')
        d['tarefas_anteriores'] = json.loads(d['tarefas_anteriores'] or '[]')
        d['assuntosNestaSessao'] = json.loads(d['assuntos_nesta'] or '[]')
        d['assuntosProximasSessoes'] = json.loads(d['assuntos_proximas'] or '[]')
        del d['assuntos_nesta']
        del d['assuntos_proximas']
        resultado.append(d)

    count_query = 'SELECT COUNT(*) as total FROM sessoes WHERE conselheiro_id = ?'
    count_params = [request.user_id]
    if status_filter:
        count_query += ' AND status = ?'
        count_params.append(status_filter)
    if data_inicio:
        count_query += ' AND data >= ?'
        count_params.append(data_inicio)
    if data_fim:
        count_query += ' AND data <= ?'
        count_params.append(data_fim)
    if assunto:
        count_query += ' AND assuntos_nesta LIKE ?'
        count_params.append(f'%{assunto}%')
    cursor.execute(count_query, count_params)
    total = cursor.fetchone()['total']

    return jsonify({
        'items': resultado,
        'total': total,
        'page': page,
        'per_page': per_page,
        'total_pages': (total + per_page - 1) // per_page
    })

@app.route('/api/sessoes', methods=['POST'])
@autenticar_token
def criar_sessao():
    data = request.json
    data_sessao = data.get('data')
    caso_num = data.get('casoNum')
    if not data_sessao or not caso_num:
        return jsonify({'erro': 'Data e caso são obrigatórios'}), 400

    is_casal = data.get('isCasal', False)
    id_pessoa = data.get('idPessoa')
    id_casal = data.get('idCasal')
    if is_casal and not id_casal:
        return jsonify({'erro': 'Casal não selecionado'}), 400
    if not is_casal and not id_pessoa:
        return jsonify({'erro': 'Pessoa não selecionada'}), 400

    sessao_num = obter_proximo_sessao_num(request.user_id, caso_num)
    duracao = data.get('duracao', '')
    versiculo = data.get('versiculo', '')
    anotacao = data.get('anotacaoSessao', '')
    status = data.get('status', 'realizada')
    tasks = json.dumps(data.get('tasks', []))
    assuntos_nesta = json.dumps(data.get('assuntosNestaSessao', []))
    assuntos_proximas = json.dumps(data.get('assuntosProximasSessoes', []))

    tarefas_anteriores = []
    db = get_db()
    cursor = db.cursor()
    if 'tarefas_anteriores' in data and data['tarefas_anteriores']:
        tarefas_anteriores = data['tarefas_anteriores']
    else:
        cursor.execute('''
            SELECT id, tasks FROM sessoes
            WHERE conselheiro_id = ? AND caso_num = ? AND id != (SELECT MAX(id) FROM sessoes WHERE conselheiro_id = ? AND caso_num = ?)
            ORDER BY sessao_num DESC LIMIT 1
        ''', (request.user_id, caso_num, request.user_id, caso_num))
        ultima = cursor.fetchone()
        if ultima:
            tasks_ant = json.loads(ultima['tasks'] or '[]')
            for idx, t in enumerate(tasks_ant):
                if not t.get('avaliacao', '').strip():
                    tarefas_anteriores.append({
                        'sessao_origem_id': ultima['id'],
                        'indice': idx,
                        'descricao': t.get('descricao', '')
                    })

    tarefas_anteriores_json = json.dumps(tarefas_anteriores)
    cursor.execute('''
        INSERT INTO sessoes (
            conselheiro_id, data, caso_num, sessao_num, duracao,
            is_casal, id_pessoa, id_casal, versiculo, anotacao_sessao,
            tasks, tarefas_anteriores, assuntos_nesta, assuntos_proximas, status
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        request.user_id, data_sessao, caso_num, sessao_num, duracao,
        1 if is_casal else 0, id_pessoa, id_casal,
        versiculo, anotacao,
        tasks, tarefas_anteriores_json, assuntos_nesta, assuntos_proximas, status
    ))
    db.commit()
    log_acao(request.user_id, 'criar_sessao', f'Sessão {caso_num}-{sessao_num} criada')
    return jsonify({'id': cursor.lastrowid, 'sessao_num': sessao_num}), 201

@app.route('/api/sessoes/<int:id>', methods=['PUT'])
@autenticar_token
def atualizar_sessao(id):
    data = request.json
    data_sessao = data.get('data')
    caso_num = data.get('casoNum')
    sessao_num = data.get('sessaoNum')
    if not data_sessao or not caso_num or sessao_num is None:
        return jsonify({'erro': 'Data, caso e número da sessão são obrigatórios'}), 400

    is_casal = data.get('isCasal', False)
    id_pessoa = data.get('idPessoa')
    id_casal = data.get('idCasal')
    if is_casal and not id_casal:
        return jsonify({'erro': 'Casal não selecionado'}), 400
    if not is_casal and not id_pessoa:
        return jsonify({'erro': 'Pessoa não selecionada'}), 400

    duracao = data.get('duracao', '')
    versiculo = data.get('versiculo', '')
    anotacao = data.get('anotacaoSessao', '')
    status = data.get('status', 'realizada')
    tasks = json.dumps(data.get('tasks', []))
    tarefas_anteriores = json.dumps(data.get('tarefas_anteriores', []))
    assuntos_nesta = json.dumps(data.get('assuntosNestaSessao', []))
    assuntos_proximas = json.dumps(data.get('assuntosProximasSessoes', []))

    db = get_db()
    cursor = db.cursor()
    cursor.execute('''
        UPDATE sessoes SET
            data = ?, caso_num = ?, sessao_num = ?, duracao = ?,
            is_casal = ?, id_pessoa = ?, id_casal = ?,
            versiculo = ?, anotacao_sessao = ?,
            tasks = ?, tarefas_anteriores = ?, assuntos_nesta = ?, assuntos_proximas = ?,
            status = ?, updated_at = CURRENT_TIMESTAMP
        WHERE id = ? AND conselheiro_id = ?
    ''', (
        data_sessao, caso_num, sessao_num, duracao,
        1 if is_casal else 0, id_pessoa, id_casal,
        versiculo, anotacao,
        tasks, tarefas_anteriores, assuntos_nesta, assuntos_proximas,
        status, id, request.user_id
    ))
    db.commit()
    if cursor.rowcount == 0:
        return jsonify({'erro': 'Sessão não encontrada'}), 404
    log_acao(request.user_id, 'atualizar_sessao', f'Sessão {id} atualizada')
    return jsonify({'atualizado': cursor.rowcount})

@app.route('/api/sessoes/<int:id>', methods=['DELETE'])
@autenticar_token
def deletar_sessao(id):
    db = get_db()
    cursor = db.cursor()
    cursor.execute('DELETE FROM sessoes WHERE id = ? AND conselheiro_id = ?', (id, request.user_id))
    db.commit()
    return jsonify({'deletado': cursor.rowcount})

# ======= ÚLTIMA SESSÃO (tarefas pendentes) =======
@app.route('/api/ultima_sessao', methods=['GET'])
@autenticar_token
def ultima_sessao():
    pessoa_id = request.args.get('pessoa_id', type=int)
    casal_id = request.args.get('casal_id', type=int)
    if not pessoa_id and not casal_id:
        return jsonify({'erro': 'Informe pessoa_id ou casal_id'}), 400

    db = get_db()
    cursor = db.cursor()
    if pessoa_id:
        cursor.execute('''
            SELECT id, caso_num, tasks FROM sessoes
            WHERE conselheiro_id = ? AND id_pessoa = ?
            ORDER BY sessao_num DESC LIMIT 1
        ''', (request.user_id, pessoa_id))
    else:
        cursor.execute('''
            SELECT id, caso_num, tasks FROM sessoes
            WHERE conselheiro_id = ? AND id_casal = ?
            ORDER BY sessao_num DESC LIMIT 1
        ''', (request.user_id, casal_id))

    row = cursor.fetchone()
    if not row:
        return jsonify({'tarefas': []})

    tasks = json.loads(row['tasks'] or '[]')
    pendentes = []
    for idx, t in enumerate(tasks):
        if not t.get('avaliacao', '').strip():
            pendentes.append({
                'sessao_origem_id': row['id'],
                'indice': idx,
                'descricao': t.get('descricao', '')
            })
    return jsonify({'tarefas': pendentes})

# ======= AVALIAR TAREFA ANTERIOR =======
@app.route('/api/avaliar_tarefa_anterior', methods=['POST'])
@autenticar_token
def avaliar_tarefa_anterior():
    data = request.json
    sessao_origem_id = data.get('sessao_origem_id')
    indice = data.get('indice')
    avaliacao = data.get('avaliacao', '')

    if sessao_origem_id is None or indice is None:
        return jsonify({'erro': 'Parâmetros inválidos'}), 400

    db = get_db()
    cursor = db.cursor()
    cursor.execute('SELECT tasks FROM sessoes WHERE id = ? AND conselheiro_id = ?', (sessao_origem_id, request.user_id))
    row = cursor.fetchone()
    if not row:
        return jsonify({'erro': 'Sessão não encontrada'}), 404

    tasks = json.loads(row['tasks'] or '[]')
    if indice >= len(tasks):
        return jsonify({'erro': 'Índice inválido'}), 400

    tasks[indice]['avaliacao'] = avaliacao
    cursor.execute('UPDATE sessoes SET tasks = ? WHERE id = ?', (json.dumps(tasks), sessao_origem_id))
    db.commit()
    log_acao(request.user_id, 'avaliar_tarefa', f'Tarefa {indice} da sessão {sessao_origem_id} avaliada')
    return jsonify({'mensagem': 'Avaliação salva'})

# ======= ANEXOS =======
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/api/anexos/<int:sessao_id>', methods=['POST'])
@autenticar_token
def upload_anexo(sessao_id):
    if 'file' not in request.files:
        return jsonify({'erro': 'Nenhum arquivo enviado'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'erro': 'Nome de arquivo vazio'}), 400
    if not allowed_file(file.filename):
        return jsonify({'erro': 'Tipo de arquivo não permitido'}), 400

    db = get_db()
    cursor = db.cursor()
    cursor.execute('SELECT id FROM sessoes WHERE id = ? AND conselheiro_id = ?', (sessao_id, request.user_id))
    if not cursor.fetchone():
        return jsonify({'erro': 'Sessão não encontrada'}), 404

    filename = secure_filename(file.filename)
    unique_name = f"{secrets.token_hex(8)}_{filename}"
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], unique_name)
    file.save(filepath)

    cursor.execute(
        'INSERT INTO anexos (sessao_id, conselheiro_id, nome_original, nome_arquivo, caminho) VALUES (?, ?, ?, ?, ?)',
        (sessao_id, request.user_id, filename, unique_name, filepath)
    )
    db.commit()
    log_acao(request.user_id, 'upload_anexo', f'Anexo {filename} para sessão {sessao_id}')
    return jsonify({'id': cursor.lastrowid, 'mensagem': 'Arquivo enviado'}), 201

@app.route('/api/anexos/<int:sessao_id>', methods=['GET'])
@autenticar_token
def listar_anexos(sessao_id):
    db = get_db()
    cursor = db.cursor()
    cursor.execute('SELECT * FROM anexos WHERE sessao_id = ? AND conselheiro_id = ?', (sessao_id, request.user_id))
    rows = cursor.fetchall()
    return jsonify([dict(row) for row in rows])

@app.route('/api/anexos/<int:anexo_id>', methods=['DELETE'])
@autenticar_token
def deletar_anexo(anexo_id):
    db = get_db()
    cursor = db.cursor()
    cursor.execute('SELECT caminho FROM anexos WHERE id = ? AND conselheiro_id = ?', (anexo_id, request.user_id))
    row = cursor.fetchone()
    if not row:
        return jsonify({'erro': 'Anexo não encontrado'}), 404
    filepath = row['caminho']
    if os.path.exists(filepath):
        os.remove(filepath)
    cursor.execute('DELETE FROM anexos WHERE id = ?', (anexo_id,))
    db.commit()
    log_acao(request.user_id, 'deletar_anexo', f'Anexo {anexo_id} removido')
    return jsonify({'deletado': 1})

@app.route('/api/anexos/download/<int:anexo_id>', methods=['GET'])
@autenticar_token
def download_anexo(anexo_id):
    db = get_db()
    cursor = db.cursor()
    cursor.execute('SELECT nome_original, caminho FROM anexos WHERE id = ? AND conselheiro_id = ?', (anexo_id, request.user_id))
    row = cursor.fetchone()
    if not row:
        return jsonify({'erro': 'Anexo não encontrado'}), 404
    return send_file(row['caminho'], download_name=row['nome_original'])

# ======= LEMBRETES =======
@app.route('/api/lembretes', methods=['GET'])
@autenticar_token
def listar_lembretes():
    db = get_db()
    cursor = db.cursor()
    cursor.execute('SELECT * FROM lembretes WHERE conselheiro_id = ? ORDER BY data_lembrete', (request.user_id,))
    rows = cursor.fetchall()
    return jsonify([dict(row) for row in rows])

@app.route('/api/lembretes', methods=['POST'])
@autenticar_token
def criar_lembrete():
    data = request.json
    titulo = data.get('titulo')
    descricao = data.get('descricao', '')
    data_lembrete = data.get('data_lembrete')
    sessao_id = data.get('sessao_id')
    if not titulo or not data_lembrete:
        return jsonify({'erro': 'Título e data são obrigatórios'}), 400
    db = get_db()
    cursor = db.cursor()
    if sessao_id:
        cursor.execute('SELECT id FROM sessoes WHERE id = ? AND conselheiro_id = ?', (sessao_id, request.user_id))
        if not cursor.fetchone():
            return jsonify({'erro': 'Sessão não encontrada'}), 404
    cursor.execute(
        'INSERT INTO lembretes (conselheiro_id, sessao_id, titulo, descricao, data_lembrete) VALUES (?, ?, ?, ?, ?)',
        (request.user_id, sessao_id, titulo, descricao, data_lembrete)
    )
    db.commit()
    log_acao(request.user_id, 'criar_lembrete', f'Lembrete {titulo}')
    return jsonify({'id': cursor.lastrowid}), 201

@app.route('/api/lembretes/<int:id>', methods=['PUT'])
@autenticar_token
def atualizar_lembrete(id):
    data = request.json
    concluido = data.get('concluido', 0)
    db = get_db()
    cursor = db.cursor()
    cursor.execute(
        'UPDATE lembretes SET concluido = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ? AND conselheiro_id = ?',
        (concluido, id, request.user_id)
    )
    db.commit()
    return jsonify({'atualizado': cursor.rowcount})

@app.route('/api/lembretes/<int:id>', methods=['DELETE'])
@autenticar_token
def deletar_lembrete(id):
    db = get_db()
    cursor = db.cursor()
    cursor.execute('DELETE FROM lembretes WHERE id = ? AND conselheiro_id = ?', (id, request.user_id))
    db.commit()
    return jsonify({'deletado': cursor.rowcount})

# ======= ESTATÍSTICAS =======
@app.route('/api/estatisticas', methods=['GET'])
@autenticar_token
def estatisticas():
    db = get_db()
    cursor = db.cursor()
    cursor.execute('SELECT COUNT(*) as total FROM sessoes WHERE conselheiro_id = ?', (request.user_id,))
    total_sessoes = cursor.fetchone()['total']
    cursor.execute('SELECT COUNT(*) as total FROM pessoas WHERE conselheiro_id = ?', (request.user_id,))
    total_pessoas = cursor.fetchone()['total']
    cursor.execute('SELECT COUNT(*) as total FROM casais WHERE conselheiro_id = ?', (request.user_id,))
    total_casais = cursor.fetchone()['total']
    cursor.execute('''
        SELECT COUNT(*) as total FROM sessoes, json_each(sessoes.tasks) as task
        WHERE sessoes.conselheiro_id = ? AND json_extract(task.value, '$.avaliacao') IS NULL
    ''', (request.user_id,))
    tarefas_pendentes = cursor.fetchone()['total']
    cursor.execute('SELECT COUNT(*) as total FROM lembretes WHERE conselheiro_id = ? AND concluido = 0', (request.user_id,))
    lembretes_pendentes = cursor.fetchone()['total']
    cursor.execute('''
        SELECT strftime('%Y-%m', data) as mes, COUNT(*) as total
        FROM sessoes
        WHERE conselheiro_id = ? AND data >= date('now', '-12 months')
        GROUP BY mes
        ORDER BY mes
    ''', (request.user_id,))
    sessoes_por_mes = [{'mes': row['mes'], 'total': row['total']} for row in cursor.fetchall()]
    return jsonify({
        'total_sessoes': total_sessoes,
        'total_pessoas': total_pessoas,
        'total_casais': total_casais,
        'tarefas_pendentes': tarefas_pendentes,
        'lembretes_pendentes': lembretes_pendentes,
        'sessoes_por_mes': sessoes_por_mes
    })

# ======= RELATÓRIOS =======
@app.route('/api/relatorios/tarefas-pendentes', methods=['GET'])
@autenticar_token
def relatorio_tarefas_pendentes():
    db = get_db()
    cursor = db.cursor()
    cursor.execute('''
        SELECT id, caso_num, sessao_num, data, tasks FROM sessoes
        WHERE conselheiro_id = ? AND tasks IS NOT NULL AND tasks != '[]'
        ORDER BY data DESC
    ''', (request.user_id,))
    rows = cursor.fetchall()
    resultado = []
    for row in rows:
        tasks = json.loads(row['tasks'] or '[]')
        pendentes = [{'descricao': t.get('descricao', ''), 'avaliacao': t.get('avaliacao', '')}
                     for t in tasks if not t.get('avaliacao', '').strip()]
        if pendentes:
            resultado.append({
                'sessao_id': row['id'],
                'caso': row['caso_num'],
                'sessao': row['sessao_num'],
                'data': row['data'],
                'tarefas_pendentes': pendentes
            })
    return jsonify(resultado)

@app.route('/api/relatorios/sessoes-por-periodo', methods=['GET'])
@autenticar_token
def relatorio_sessoes_periodo():
    data_inicio = request.args.get('data_inicio')
    data_fim = request.args.get('data_fim')
    if not data_inicio or not data_fim:
        return jsonify({'erro': 'Informe data_inicio e data_fim'}), 400
    db = get_db()
    cursor = db.cursor()
    cursor.execute('''
        SELECT * FROM sessoes
        WHERE conselheiro_id = ? AND data BETWEEN ? AND ?
        ORDER BY data DESC
    ''', (request.user_id, data_inicio, data_fim))
    rows = cursor.fetchall()
    resultado = []
    for row in rows:
        d = dict(row)
        d['is_casal'] = bool(d['is_casal'])
        d['tasks'] = json.loads(d['tasks'] or '[]')
        d['tarefas_anteriores'] = json.loads(d['tarefas_anteriores'] or '[]')
        d['assuntosNestaSessao'] = json.loads(d['assuntos_nesta'] or '[]')
        d['assuntosProximasSessoes'] = json.loads(d['assuntos_proximas'] or '[]')
        del d['assuntos_nesta']
        del d['assuntos_proximas']
        resultado.append(d)
    return jsonify(resultado)

@app.route('/api/relatorios/pessoas-atendidas', methods=['GET'])
@autenticar_token
def relatorio_pessoas_atendidas():
    db = get_db()
    cursor = db.cursor()
    cursor.execute('''
        SELECT DISTINCT p.id, p.nome, p.obs,
            (SELECT COUNT(*) FROM sessoes WHERE sessoes.id_pessoa = p.id AND sessoes.conselheiro_id = ?) as total_sessoes
        FROM pessoas p
        WHERE p.conselheiro_id = ?
        ORDER BY total_sessoes DESC
    ''', (request.user_id, request.user_id))
    rows = cursor.fetchall()
    return jsonify([dict(row) for row in rows])

# ======= EXPORTAÇÃO EXCEL =======
@app.route('/api/exportar/excel', methods=['GET'])
@autenticar_token
def exportar_excel():
    db = get_db()
    cursor = db.cursor()
    data_inicio = request.args.get('data_inicio')
    data_fim = request.args.get('data_fim')
    query = 'SELECT * FROM sessoes WHERE conselheiro_id = ?'
    params = [request.user_id]
    if data_inicio:
        query += ' AND data >= ?'
        params.append(data_inicio)
    if data_fim:
        query += ' AND data <= ?'
        params.append(data_fim)
    cursor.execute(query, params)
    rows = cursor.fetchall()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Sessões'
    headers = ['ID', 'Data', 'Caso', 'Sessão', 'Duração', 'Tipo', 'Pessoa/Casal', 'Versículo', 'Anotação', 'Tarefas', 'Assuntos Nesta', 'Assuntos Próx', 'Status']
    ws.append(headers)
    for col in range(1, len(headers)+1):
        cell = ws.cell(row=1, column=col)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
        cell.alignment = Alignment(horizontal='center')

    for row in rows:
        tipo = 'Casal' if row['is_casal'] else 'Individual'
        pessoa_casal = ''
        if row['is_casal']:
            cursor.execute('SELECT nome_casal FROM casais WHERE id = ?', (row['id_casal'],))
            c = cursor.fetchone()
            pessoa_casal = c['nome_casal'] if c else ''
        else:
            cursor.execute('SELECT nome FROM pessoas WHERE id = ?', (row['id_pessoa'],))
            p = cursor.fetchone()
            pessoa_casal = p['nome'] if p else ''
        tasks = json.loads(row['tasks'] or '[]')
        tasks_str = '; '.join([f"{t.get('descricao','')} [{t.get('avaliacao','')}]" for t in tasks])
        assuntos_nesta = json.loads(row['assuntos_nesta'] or '[]')
        assuntos_proximas = json.loads(row['assuntos_proximas'] or '[]')
        ws.append([
            row['id'], row['data'], row['caso_num'], row['sessao_num'], row['duracao'],
            tipo, pessoa_casal, row['versiculo'] or '', row['anotacao_sessao'] or '',
            tasks_str, ', '.join(assuntos_nesta), ', '.join(assuntos_proximas), row['status'] or 'realizada'
        ])

    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[col_letter].width = adjusted_width

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, download_name='sessoes.xlsx', as_attachment=True, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# ======= BACKUP MANUAL =======
@app.route('/api/backup', methods=['GET'])
@autenticar_token
def baixar_backup():
    return send_file(DATABASE, as_attachment=True, download_name=f'backup_{datetime.now().strftime("%Y%m%d")}.db')

# ======= LOGS =======
@app.route('/api/logs', methods=['GET'])
@autenticar_token
def listar_logs():
    db = get_db()
    cursor = db.cursor()
    cursor.execute('SELECT * FROM logs WHERE conselheiro_id = ? ORDER BY created_at DESC LIMIT 100', (request.user_id,))
    rows = cursor.fetchall()
    return jsonify([dict(row) for row in rows])

# ======= PING =======
@app.route('/api/ping', methods=['GET'])
def ping():
    return jsonify({'status': 'ok', 'timestamp': datetime.utcnow().isoformat()})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=PORT, debug=False)