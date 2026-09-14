from flask import Flask, render_template, g, redirect, url_for, request, Response, flash, jsonify, send_file
import sqlite3
import os
import math
from datetime import date, timedelta, datetime
from openpyxl import Workbook, load_workbook
import calendar
import pandas as pd
import webbrowser
from threading import Timer
import sys
import urllib.request
import json
import shutil
import glob
import time 
import threading
from pathlib import Path

app = Flask(__name__)
@app.get("/favicon.ico")
def favicon():
    return ("", 204)


PER_PAGE = 50

# -------------------------------------------------------------------
# Caminho base do app (funciona tanto rodando .py quanto .exe)
# -------------------------------------------------------------------
if getattr(sys, "frozen", False):
    # Modo "congelado" (PyInstaller): pasta onde está o .exe
    BASE_DIR = Path(sys.executable).parent
else:
    # Modo normal (python app.py): pasta onde está o app.py
    BASE_DIR = Path(__file__).resolve().parent

# Caminhos do banco de dados e do Excel (sempre na mesma pasta do app/exe)
DATABASE = str(BASE_DIR / "financeiro_escola.db")
EXCEL_NOTAS = str(BASE_DIR / "notas_emitidas.xlsx")

def data_br(value):
    if not value:
        return ""
    return datetime.strptime(value, "%Y-%m-%d").strftime("%d/%m/%Y")
app.jinja_env.filters["data_br"] = data_br

from datetime import date, datetime
# ...

def limpar_digitos(valor: str) -> str:
    """Remove tudo que não for dígito."""
    if not valor:
        return ""
    return "".join(c for c in valor if c.isdigit())

def data_br_excel(value):
    """Converte 'YYYY-MM-DD' → 'DD/MM/YYYY' para saída no Excel."""
    if not value:
        return ""
    y, m, d = value.split("-")
    return f"{d}/{m}/{y}"


def validar_cpf(cpf: str) -> bool:
    """
    Validação oficial do CPF:
    - 11 dígitos
    - rejeita sequências iguais
    - confere dígitos verificadores
    """
    cpf = limpar_digitos(cpf)

    # Precisa ter 11 dígitos
    if len(cpf) != 11:
        return False

    # Não pode ser sequência repetida (11111111111 etc.)
    if cpf == cpf[0] * 11:
        return False

    # Dígito 1
    soma = sum(int(cpf[i]) * (10 - i) for i in range(9))
    dig1 = (soma * 10) % 11
    dig1 = 0 if dig1 == 10 else dig1
    if dig1 != int(cpf[9]):
        return False

    # Dígito 2
    soma = sum(int(cpf[i]) * (11 - i) for i in range(10))
    dig2 = (soma * 10) % 11
    dig2 = 0 if dig2 == 10 else dig2
    if dig2 != int(cpf[10]):
        return False

    return True


def validar_campo_cpf_cnpj(cpf_cnpj_raw: str) -> bool:
    """
    Regras para o campo cpf_cnpj:
    - vazio  → ok
    - 11 dígitos → valida como CPF
    - outro tamanho → deixa passar (pode ser CNPJ)
    """
    digitos = limpar_digitos(cpf_cnpj_raw)

    if not digitos:
        return True  # campo em branco, aceita

    if len(digitos) != 11:
        return True  # provavelmente CNPJ, não validamos aqui

    return validar_cpf(digitos)


def add_months(date_str, months):
    """Recebe 'YYYY-MM-DD' e devolve a mesma data + N meses."""
    if not date_str:
        return None
    y, m, d = map(int, date_str.split("-"))
    m += months
    y += (m - 1) // 12
    m = (m - 1) % 12 + 1
    last_day = calendar.monthrange(y, m)[1]
    d = min(d, last_day)
    return f"{y:04d}-{m:02d}-{d:02d}"


def get_db():
    """Abre conexão com o SQLite e guarda em g (contexto da requisição)."""
    if "db" not in g:
        g.db = sqlite3.connect(DATABASE)
        g.db.row_factory = sqlite3.Row  # permite acessar colunas por nome
    return g.db
def append_nf_to_excel(row_dict):
    """
    row_dict deve conter todas as colunas da NF (chaves iguais às do header abaixo).
    Se o arquivo ainda não existir, ele é criado com cabeçalho.
    """
    headers = [
        "numero_nf",
        "data_emissao",
        "valor_servicos",
        "discriminacao",
        "cpf_cnpj_tomador",
        "nome_tomador",
        "endereco_rua",
        "endereco_numero",
        "endereco_complemento",
        "bairro",
        "cidade",
        "uf",
        "cep",
        "email_tomador",
        "contrato_id",
        "parcela_id",
    ]

    if os.path.exists(EXCEL_NOTAS):
        wb = load_workbook(EXCEL_NOTAS)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.append(headers)

    ws.append([row_dict.get(h, "") for h in headers])
    wb.save(EXCEL_NOTAS)


def export_nf_to_excel_for_parcela(db, parcela_id):
    """
    Busca os dados completos da NF ligada à parcela (parcela_id)
    e grava uma linha no Excel.
    """
    nota = db.execute(
        """
        SELECT
            nf.numero_nf,
            nf.data_emissao,
            nf.valor,
            a.nome  AS aluno,
            r.nome  AS responsavel,
            r.cpf_cnpj       AS cpf_cnpj_tomador,
            r.endereco_rua   AS rua_tomador,
            r.endereco_numero       AS num_tomador,
            r.endereco_complemento  AS compl_tomador,
            r.endereco_bairro       AS bairro_tomador,
            r.endereco_cidade       AS cidade_tomador,
            r.endereco_uf           AS uf_tomador,
            r.endereco_cep          AS cep_tomador,
            r.email                 AS email_tomador,
            c.id AS contrato_id,
            p.id AS parcela_id
        FROM notas_fiscais nf
        JOIN parcelas p  ON nf.parcela_id = p.id
        JOIN contratos c ON p.contrato_id = c.id
        JOIN alunos a    ON c.aluno_id = a.id
        JOIN responsaveis r ON c.responsavel_id = r.id
        WHERE nf.parcela_id = ?;
        """,
        (parcela_id,),
    ).fetchone()

    if nota is None:
        return

    valor_servicos = nota["valor"] or 0.0
    tributos_aprox = round(valor_servicos * 0.1619, 2)

    discriminacao = (
        f"Parcela referente ao curso de programação e robótica, contrato nº {nota['contrato_id']}. "
        f"Conf lei 12.741-2012, o valor aproximado dos tributos é de R$ {tributos_aprox:.2f}, "
        f"correspondente ao percentual de 16,19%. Fonte IBPT."
    )

    row = {
        "numero_nf": nota["numero_nf"],
        "data_emissao": nota["data_emissao"],
        "valor_servicos": valor_servicos,
        "discriminacao": discriminacao,
        "cpf_cnpj_tomador": nota["cpf_cnpj_tomador"],
        "nome_tomador": nota["responsavel"],
        "endereco_rua": nota["rua_tomador"],
        "endereco_numero": nota["num_tomador"],
        "endereco_complemento": nota["compl_tomador"],
        "bairro": nota["bairro_tomador"],
        "cidade": nota["cidade_tomador"],
        "uf": nota["uf_tomador"],
        "cep": nota["cep_tomador"],
        "email_tomador": nota["email_tomador"],
        "contrato_id": nota["contrato_id"],
        "parcela_id": nota["parcela_id"],
    }

    append_nf_to_excel(row)
def fazer_backup():
    """
    Faz backup do banco de dados.
    Retorna (sucesso, mensagem, caminho_arquivo)
    """
    try:
        # Cria pasta backups se não existir
        backup_dir = BASE_DIR / "backups"
        backup_dir.mkdir(exist_ok=True)
        
        # Nome do arquivo com timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_file = backup_dir / f"financeiro_backup_{timestamp}.db"
        
        # Copia o banco
        shutil.copy2(DATABASE, backup_file)
        
        # Limpa backups antigos (mantém últimos 7 dias)
        limpar_backups_antigos(backup_dir)
        
        # Log
        print(f"✅ Backup criado: {backup_file.name}")
        return True, f"Backup criado com sucesso: {backup_file.name}", str(backup_file)
        
    except Exception as e:
        error_msg = f"Erro ao criar backup: {str(e)}"
        print(f"❌ {error_msg}")
        return False, error_msg, None

def limpar_backups_antigos(backup_dir, dias=7):
    """Remove backups com mais de X dias."""
    limite = datetime.now() - timedelta(days=dias)
    removidos = 0
    
    for backup in backup_dir.glob("financeiro_backup_*.db"):
        try:
            # Extrai data do nome do arquivo
            nome = backup.stem
            data_str = nome.replace("financeiro_backup_", "")
            
            # Tenta parsear a data do formato YYYYMMDD_HHMMSS
            data_backup = datetime.strptime(data_str[:15], "%Y%m%d_%H%M%S")
            
            if data_backup < limite:
                backup.unlink()
                removidos += 1
                print(f"  🗑️ Removido backup antigo: {backup.name}")
                
        except Exception as e:
            print(f"  ⚠️ Não pôde processar {backup.name}: {e}")
            continue
    
    if removidos > 0:
        print(f"  ✅ {removidos} backups antigos removidos")

def fazer_backup_periodico(intervalo_horas=24):
    """
    Agenda backups periódicos.
    Executa a cada X horas enquanto o app estiver rodando.
    SEM backup inicial - apenas periódico.
    """
    # Verifica se já está rodando (evita duplicação no reload do debug)
    if hasattr(fazer_backup_periodico, '_timer_rodando'):
        print("⏸️ Backup periódico já está configurado (ignorando chamada dupla)")
        return
    
    fazer_backup_periodico._timer_rodando = True
    
    def executar_backup_e_agendar_proximo():
        """Executa backup e agenda o próximo."""
        fazer_backup()
        
        # Agenda próximo (após intervalo completo)
        timer = threading.Timer(intervalo_horas * 3600, executar_backup_e_agendar_proximo)
        timer.daemon = True
        timer.start()
        print(f"⏰ Próximo backup agendado para daqui a {intervalo_horas} horas")
    
    # NÃO faz backup imediato - agenda o primeiro para daqui a X horas
    primeiro_timer = threading.Timer(intervalo_horas * 3600, executar_backup_e_agendar_proximo)
    primeiro_timer.daemon = True
    primeiro_timer.start()
    
    print(f"✅ Backup automático configurado (primeiro em {intervalo_horas} horas, depois a cada {intervalo_horas} horas)")

def listar_backups():
    """Lista todos os backups disponíveis."""
    backup_dir = BASE_DIR / "backups"
    backup_dir.mkdir(exist_ok=True)
    
    backups = []
    for arquivo in sorted(backup_dir.glob("financeiro_backup_*.db"), reverse=True):
        try:
            stats = arquivo.stat()
            # Extrai data do nome
            nome = arquivo.stem
            data_str = nome.replace("financeiro_backup_", "")
            data_formatada = "Data inválida"
            
            try:
                data_obj = datetime.strptime(data_str[:15], "%Y%m%d_%H%M%S")
                data_formatada = data_obj.strftime("%d/%m/%Y %H:%M:%S")
            except:
                data_formatada = datetime.fromtimestamp(stats.st_mtime).strftime("%d/%m/%Y %H:%M")
            
            backups.append({
                'nome': arquivo.name,
                'tamanho': stats.st_size,
                'tamanho_formatado': formatar_tamanho(stats.st_size),
                'data': data_formatada,
                'data_original': data_str[:15] if len(data_str) >= 15 else "N/A",
                'caminho': str(arquivo),
                'idade_dias': (datetime.now() - datetime.fromtimestamp(stats.st_mtime)).days
            })
        except Exception as e:
            print(f"⚠️ Erro ao processar {arquivo.name}: {e}")
            continue
    
    return backups

def formatar_tamanho(bytes_size):
    """Formata tamanho em bytes para KB/MB/GB."""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if bytes_size < 1024.0 or unit == 'GB':
            return f"{bytes_size:.2f} {unit}"
        bytes_size /= 1024.0
    return f"{bytes_size:.2f} GB"

@app.teardown_appcontext
def close_db(exception):
    """Fecha a conexão ao final de cada requisição."""
    db = g.pop("db", None)
    if db is not None:
        db.close()


# =========================
# PÁGINA INICIAL (PAINEL)
# =========================
@app.post("/notas/gerar/<int:parcela_id>")
def gerar_nf(parcela_id):
    db = get_db()

    # Busca infos da parcela
    parcela = db.execute(
        """
        SELECT
            p.id,
            p.valor_final
        FROM parcelas p
        WHERE p.id = ?;
        """,
        (parcela_id,),
    ).fetchone()

    if parcela is None:
        return "Parcela não encontrada", 404

    # Valor e data da NF
    valor = parcela["valor_final"]
    hoje = date.today().isoformat()

    # Número de NF fictício (você pode depois editar esse campo no DB Browser ou criar tela própria)
    numero_nf = f"NF-{parcela_id}-{hoje}"

    db.execute(
        """
        INSERT INTO notas_fiscais (parcela_id, numero_nf, data_emissao, valor, status)
        VALUES (?, ?, ?, ?, 'emitida');
        """,
        (parcela_id, numero_nf, hoje, valor),
    )
    db.commit()
    export_nf_to_excel_for_parcela(db, parcela_id)

    # Voltar para a página de onde veio
    ref = request.headers.get("Referer")
    if ref:
        return redirect(ref)
    return redirect(url_for("notas_a_emitir"))
from flask import jsonify
import urllib.request
import json
@app.route("/admin/backup")
def pagina_backup():
    """Página principal de backup."""
    backups = listar_backups()
    total_size = sum(b['tamanho'] for b in backups)
    
    # Estatísticas
    hoje = datetime.now().date()
    backups_hoje = len([b for b in backups if datetime.strptime(b['data_original'], "%Y%m%d_%H%M%S").date() == hoje])
    
    return render_template(
        "backup.html",
        backups=backups,
        total_backups=len(backups),
        total_size=formatar_tamanho(total_size),
        backups_hoje=backups_hoje,
        database_size=os.path.getsize(DATABASE) if os.path.exists(DATABASE) else 0,
        database_size_fmt=formatar_tamanho(os.path.getsize(DATABASE)) if os.path.exists(DATABASE) else "N/A"
    )

@app.post("/admin/backup/criar")
def criar_backup():
    """Cria um backup manualmente."""
    sucesso, mensagem, caminho = fazer_backup()
    
    if sucesso:
        return jsonify({
            'sucesso': True,
            'mensagem': mensagem,
            'caminho': caminho
        })
    else:
        return jsonify({
            'sucesso': False,
            'mensagem': mensagem
        }), 500

@app.route("/admin/backup/download/<nome>")
def download_backup(nome):
    """Faz download de um backup."""
    backup_path = BASE_DIR / "backups" / nome
    
    if not backup_path.exists():
        return "Arquivo não encontrado", 404
    
    return send_file(
        backup_path,
        as_attachment=True,
        download_name=nome
    )

@app.post("/admin/backup/restaurar")
def restaurar_backup():
    """Restaura um backup."""
    nome_arquivo = request.form.get("arquivo")
    
    if not nome_arquivo:
        return "Nome do arquivo não especificado", 400
    
    backup_path = BASE_DIR / "backups" / nome_arquivo
    
    if not backup_path.exists():
        return "Arquivo de backup não encontrado", 404
    
    try:
        # 1. Fecha conexões atuais
        close_db(None)
        time.sleep(1)  # Aguarda liberação
        
        # 2. Faz backup do banco atual antes de restaurar
        fazer_backup()
        
        # 3. Restaura o backup
        shutil.copy2(backup_path, DATABASE)
        
        return """
        <h1>✅ Backup restaurado com sucesso!</h1>
        <p>O banco de dados foi restaurado para o estado do backup.</p>
        <p>O sistema será reiniciado automaticamente em 5 segundos...</p>
        <script>
            setTimeout(function() {
                window.location.href = "/";
            }, 5002);
        </script>
        """
        
    except Exception as e:
        return f"""
        <h1>❌ Erro ao restaurar backup</h1>
        <p>{str(e)}</p>
        <p>Tente novamente ou contate o suporte.</p>
        <a href="/admin/backup">Voltar</a>
        """, 500

@app.post("/admin/backup/excluir")
def excluir_backup():
    """Exclui um arquivo de backup."""
    nome_arquivo = request.form.get("arquivo")
    
    if not nome_arquivo:
        return jsonify({'sucesso': False, 'mensagem': 'Nome do arquivo não especificado'}), 400
    
    backup_path = BASE_DIR / "backups" / nome_arquivo
    
    if not backup_path.exists():
        return jsonify({'sucesso': False, 'mensagem': 'Arquivo não encontrado'}), 404
    
    try:
        backup_path.unlink()
        return jsonify({'sucesso': True, 'mensagem': f'Backup {nome_arquivo} excluído'})
    except Exception as e:
        return jsonify({'sucesso': False, 'mensagem': f'Erro ao excluir: {str(e)}'}), 500
@app.route("/notas/preview_lote", methods=["GET", "POST"])
def preview_lote():
    """Pré-visualização das notas a serem emitidas em lote."""
    db = get_db()
    
    if request.method == "GET":
        # Redireciona se acessado diretamente sem parcelas
        return redirect(url_for("notas_a_emitir"))
    
    # POST - mostra pré-visualização
    parcelas_ids = request.form.getlist("parcelas_ids")
    
    if not parcelas_ids:
        return redirect(url_for("notas_a_emitir"))
    
    # Busca informações detalhadas das parcelas selecionadas
    placeholders = ",".join(["?"] * len(parcelas_ids))
    
    parcelas = db.execute(f"""
        SELECT
            p.id,
            p.numero_parcela,
            p.valor_final,
            p.data_vencimento,
            a.nome AS aluno,
            r.nome AS responsavel,
            r.cpf_cnpj,
            r.email AS email_tomador,
            c.id AS contrato_id,
            c.numero_contrato
        FROM parcelas p
        JOIN contratos c    ON p.contrato_id    = c.id
        JOIN alunos a       ON c.aluno_id       = a.id
        JOIN responsaveis r ON c.responsavel_id = r.id
        WHERE p.id IN ({placeholders})
        ORDER BY p.data_vencimento, p.numero_parcela;
    """, parcelas_ids).fetchall()
    
    # Calcula totais
    valor_total = sum(p['valor_final'] for p in parcelas)
    quantidade = len(parcelas)
    
    # Agrupa por responsável (para mostrar quantas notas serão emitidas)
    responsaveis = {}
    for parcela in parcelas:
        resp_id = parcela['responsavel']
        if resp_id not in responsaveis:
            responsaveis[resp_id] = {
                'nome': parcela['responsavel'],
                'cpf_cnpj': parcela['cpf_cnpj'],
                'parcelas': [],
                'valor_total': 0
            }
        responsaveis[resp_id]['parcelas'].append(parcela)
        responsaveis[resp_id]['valor_total'] += parcela['valor_final']
    
    return render_template(
        "preview_lote.html",
        parcelas=parcelas,
        parcelas_ids=parcelas_ids,
        valor_total=valor_total,
        quantidade=quantidade,
        responsaveis=responsaveis,
        quantidade_responsaveis=len(responsaveis)
    )
@app.get("/api/cep/<cep>")
def api_cep(cep):
    cep = "".join([c for c in cep if c.isdigit()])

    if len(cep) != 8:
        return jsonify({"erro": True, "msg": "CEP inválido"}), 400

    url = f"https://viacep.com.br/ws/{cep}/json/"
    try:
        with urllib.request.urlopen(url, timeout=8) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except Exception:
        return jsonify({"erro": True, "msg": "Falha ao consultar ViaCEP"}), 502

    return jsonify(data)
@app.post("/aluno/<int:aluno_id>/excluir")
def excluir_aluno(aluno_id):
    """Desativa aluno e seus contratos (soft delete)."""
    db = get_db()
    
    try:
        # 1. Marca aluno como excluído
        db.execute(
            """
            UPDATE alunos 
            SET ativo = 0, excluido_em = CURRENT_TIMESTAMP
            WHERE id = ?;
            """,
            (aluno_id,)
        )
        
        # 2. Desativa contratos do aluno
        db.execute(
            """
            UPDATE contratos 
            SET ativo = 0, excluido_em = CURRENT_TIMESTAMP
            WHERE aluno_id = ?;
            """,
            (aluno_id,)
        )
        
        # 3. Cancela parcelas pendentes (sem NF emitida)
        db.execute(
            """
            UPDATE parcelas 
            SET cancelada = 1, excluido_em = CURRENT_TIMESTAMP
            WHERE contrato_id IN (
                SELECT id FROM contratos WHERE aluno_id = ?
            )
            AND id NOT IN (
                SELECT parcela_id FROM notas_fiscais WHERE parcela_id IS NOT NULL
            );
            """,
            (aluno_id,)
        )
        
        db.commit()
        return redirect(url_for("listar_alunos"))
        
    except Exception as e:
        db.rollback()
        return f"Erro ao excluir aluno: {str(e)}", 500

@app.post("/responsavel/<int:resp_id>/excluir")
def excluir_responsavel(resp_id):
    """Desativa responsável e seus dependentes."""
    db = get_db()
    
    try:
        # 1. Marca responsável como excluído
        db.execute(
            """
            UPDATE responsaveis 
            SET ativo = 0, excluido_em = CURRENT_TIMESTAMP
            WHERE id = ?;
            """,
            (resp_id,)
        )
        
        # 2. Desativa alunos desse responsável
        db.execute(
            """
            UPDATE alunos 
            SET ativo = 0, excluido_em = CURRENT_TIMESTAMP
            WHERE responsavel_id = ?;
            """,
            (resp_id,)
        )
        
        # 3. Desativa contratos desse responsável
        db.execute(
            """
            UPDATE contratos 
            SET ativo = 0, excluido_em = CURRENT_TIMESTAMP
            WHERE responsavel_id = ?;
            """,
            (resp_id,)
        )
        
        # 4. Cancela parcelas pendentes
        db.execute(
            """
            UPDATE parcelas 
            SET cancelada = 1, excluido_em = CURRENT_TIMESTAMP
            WHERE contrato_id IN (
                SELECT id FROM contratos WHERE responsavel_id = ?
            )
            AND id NOT IN (
                SELECT parcela_id FROM notas_fiscais WHERE parcela_id IS NOT NULL
            );
            """,
            (resp_id,)
        )
        
        db.commit()
        return redirect(url_for("listar_alunos"))
        
    except Exception as e:
        db.rollback()
        return f"Erro ao excluir responsável: {str(e)}", 500

@app.post("/contrato/<int:contrato_id>/excluir")
def excluir_contrato(contrato_id):
    """Desativa um contrato específico."""
    db = get_db()
    
    # Verifica se há notas fiscais emitidas
    nfs_emitidas = db.execute(
        """
        SELECT COUNT(*) as total
        FROM notas_fiscais nf
        JOIN parcelas p ON nf.parcela_id = p.id
        WHERE p.contrato_id = ?;
        """,
        (contrato_id,)
    ).fetchone()
    
    if nfs_emitidas['total'] > 0:
        return """
        <h1>Não é possível excluir este contrato</h1>
        <p>Já existem notas fiscais emitidas para este contrato.</p>
        <p>Se necessário, cancele apenas as parcelas pendentes.</p>
        <a href="{{ url_for('detalhe_contrato', contrato_id=contrato_id) }}">Voltar</a>
        """, 400
    
    try:
        # 1. Marca contrato como excluído
        db.execute(
            """
            UPDATE contratos 
            SET ativo = 0, excluido_em = CURRENT_TIMESTAMP
            WHERE id = ?;
            """,
            (contrato_id,)
        )
        
        # 2. Cancela todas as parcelas (já que não há NFs)
        db.execute(
            """
            UPDATE parcelas 
            SET cancelada = 1, excluido_em = CURRENT_TIMESTAMP
            WHERE contrato_id = ?;
            """,
            (contrato_id,)
        )
        
        db.commit()
        
        # Redireciona para o aluno
        contrato = db.execute(
            "SELECT aluno_id FROM contratos WHERE id = ?;",
            (contrato_id,)
        ).fetchone()
        
        if contrato:
            return redirect(url_for('detalhe_aluno', aluno_id=contrato['aluno_id']))
        else:
            return redirect(url_for('listar_alunos'))
            
    except Exception as e:
        db.rollback()
        return f"Erro ao excluir contrato: {str(e)}", 500

@app.post("/parcela/<int:parcela_id>/cancelar")
def cancelar_parcela(parcela_id):
    """Cancela uma parcela específica (se não tiver NF)."""
    db = get_db()
    
    # Verifica se já tem NF
    nf = db.execute(
        "SELECT id FROM notas_fiscais WHERE parcela_id = ?;",
        (parcela_id,)
    ).fetchone()
    
    if nf:
        return """
        <h1>Não é possível cancelar esta parcela</h1>
        <p>Já existe uma nota fiscal emitida para esta parcela.</p>
        <a href="javascript:history.back()">Voltar</a>
        """, 400
    
    try:
        db.execute(
            """
            UPDATE parcelas 
            SET cancelada = 1, excluido_em = CURRENT_TIMESTAMP
            WHERE id = ?;
            """,
            (parcela_id,)
        )
        
        db.commit()
        
        # Redireciona de volta
        parcela = db.execute(
            "SELECT contrato_id FROM parcelas WHERE id = ?;",
            (parcela_id,)
        ).fetchone()
        
        if parcela:
            return redirect(url_for('detalhe_contrato', contrato_id=parcela['contrato_id']))
        else:
            return redirect(url_for('listar_parcelas'))
            
    except Exception as e:
        db.rollback()
        return f"Erro ao cancelar parcela: {str(e)}", 500
    
@app.post("/notas/gerar_lote")
def gerar_nf_lote():
    print("="*70)
    print("🚀 INICIANDO gerar_nf_lote")
    print("="*70)
    
    db = get_db()
    
    parcelas_ids = request.form.getlist("parcelas_ids")
    print(f"📋 Parcelas selecionadas: {parcelas_ids}")
    
    if not parcelas_ids:
        print("❌ Nenhuma parcela selecionada")
        return redirect(url_for("notas_a_emitir"))
    
    hoje = date.today().isoformat()
    emitidas = 0
    erros = []
    
    linhas_excel = []
    
    for pid in parcelas_ids:
        try:
            print(f"\n📦 Processando parcela {pid}...")
            
            # Busca dados da parcela
            parcela = db.execute(
                """
                SELECT
                    p.id,
                    p.numero_parcela,
                    p.valor_final,
                    p.data_vencimento,
                    a.nome AS aluno,
                    r.nome AS responsavel,
                    r.cpf_cnpj,
                    r.endereco_rua,
                    r.endereco_numero,
                    r.endereco_complemento,
                    r.endereco_bairro,
                    r.endereco_cidade,
                    r.endereco_uf,
                    r.endereco_cep,
                    r.email,
                    c.id AS contrato_id,
                    c.numero_contrato
                FROM parcelas p
                JOIN contratos c    ON p.contrato_id = c.id
                JOIN alunos a       ON c.aluno_id = a.id
                JOIN responsaveis r ON c.responsavel_id = r.id
                WHERE p.id = ?;
                """,
                (pid,),
            ).fetchone()
            
            if parcela is None:
                erro = f"Parcela {pid} não encontrada"
                print(f"❌ {erro}")
                erros.append(erro)
                continue
            
            print(f"✅ Parcela encontrada: {parcela['aluno']} - R$ {parcela['valor_final']}")
            
            # Cria a nota fiscal
            cur = db.execute(
                """
                INSERT INTO notas_fiscais (
                    parcela_id,
                    data_emissao,
                    valor,
                    status
                ) VALUES (?, ?, ?, 'emitida');
                """,
                (parcela["id"], hoje, parcela["valor_final"]),
            )
            nf_id = cur.lastrowid
            
            # Gera número da NF
            numero_nf = f"NF-{nf_id}-{date.today().strftime('%Y%m%d')}"
            db.execute(
                "UPDATE notas_fiscais SET numero_nf = ? WHERE id = ?;",
                (numero_nf, nf_id),
            )
            
            print(f"📝 NF criada: {numero_nf} (ID: {nf_id})")
            
            # Adiciona ao Excel
            linhas_excel.append({
                "ID_NF": nf_id,
                "ID_Parcela": parcela["id"],
                "Contrato": parcela["numero_contrato"],
                "Numero_Parcela": parcela["numero_parcela"],
                "Data_Emissao": data_br_excel(hoje),
                "Valor_Servicos": float(parcela["valor_final"] or 0),
                "Aluno": parcela["aluno"],
                "CPF_CNPJ_Tomador": parcela["cpf_cnpj"],
                "Nome_Tomador": parcela["responsavel"],
                "Endereco_Tomador": parcela["endereco_rua"],
                "Numero_Tomador": parcela["endereco_numero"],
                "Compl_Tomador": parcela["endereco_complemento"],
                "Bairro_Tomador": parcela["endereco_bairro"],
                "Cidade_Tomador": parcela["endereco_cidade"],
                "UF_Tomador": parcela["endereco_uf"],
                "CEP_Tomador": parcela["endereco_cep"],
                "Email_Tomador": parcela["email"],
            })
            
            emitidas += 1
            
        except Exception as e:
            erro = f"Erro na parcela {pid}: {str(e)}"
            print(f"❌ {erro}")
            import traceback
            traceback.print_exc()
            erros.append(erro)
            continue
    
    db.commit()
    print(f"\n✅ {emitidas} notas fiscais criadas")
    print(f"⚠️  {len(erros)} erros")
    
    # Cria planilha Excel
    caminho = None
    nome_arquivo = None
    
    if linhas_excel:
        print("\n📊 Criando planilha Excel...")
        
        try:
            import pandas as pd
            
            df = pd.DataFrame(linhas_excel)
            print(f"✅ DataFrame criado: {len(df)} linhas")
            
            pasta = os.path.join(BASE_DIR, "exports_nf")
            print(f"📂 Pasta destino: {pasta}")
            
            # Cria pasta se não existir
            if not os.path.exists(pasta):
                os.makedirs(pasta)
                print(f"✅ Pasta criada: {pasta}")
            else:
                print(f"✅ Pasta já existe")
            
            # Gera nome do arquivo
            nome_arquivo = f"nf_lote_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            caminho = os.path.join(pasta, nome_arquivo)
            print(f"💾 Salvando como: {nome_arquivo}")
            print(f"📍 Caminho completo: {caminho}")
            
            # Salva Excel
            df.to_excel(caminho, index=False)
            
            # Verifica se foi salvo
            if os.path.exists(caminho):
                tamanho = os.path.getsize(caminho)
                print(f"✅ Arquivo salvo! Tamanho: {tamanho} bytes")
            else:
                print(f"❌ ERRO: Arquivo não foi criado!")
                
        except Exception as e:
            print(f"❌ ERRO ao criar Excel: {str(e)}")
            import traceback
            traceback.print_exc()
    
    print("="*70)
    print("🎯 FINALIZANDO gerar_nf_lote")
    print("="*70)
    
    # Página de confirmação
    return render_template(
        "confirmacao_lote.html",
        emitidas=emitidas,
        erros=erros,
        caminho_arquivo=caminho if linhas_excel else None,
        nome_arquivo=nome_arquivo if linhas_excel else None
    )

@app.route("/notas/download_excel/<nome>")
def download_excel_lote(nome):
    """Faz download da planilha Excel do lote."""
    print("="*70)
    print("📥 INICIANDO download_excel_lote")
    print("="*70)
    print(f"📄 Nome solicitado: {nome}")
    
    # Verifica BASE_DIR
    print(f"📁 BASE_DIR: {BASE_DIR}")
    
    # Constrói caminho
    pasta = os.path.join(BASE_DIR, "exports_nf")
    caminho = os.path.join(pasta, nome)
    
    print(f"📂 Pasta exports_nf: {pasta}")
    print(f"📍 Caminho completo: {caminho}")
    print(f"✅ Pasta existe? {os.path.exists(pasta)}")
    
    if not os.path.exists(pasta):
        print("❌ ERRO: Pasta exports_nf não existe!")
        # Tenta listar o que tem no BASE_DIR
        print("\n📋 Conteúdo do BASE_DIR:")
        try:
            for item in os.listdir(BASE_DIR):
                print(f"   - {item}")
        except Exception as e:
            print(f"   Erro ao listar: {e}")
        return "Pasta de exportação não encontrada", 404
    
    print(f"✅ Arquivo existe? {os.path.exists(caminho)}")
    
    if not os.path.exists(caminho):
        print("❌ ERRO: Arquivo não encontrado!")
        
        # Lista o que tem na pasta
        print("\n📋 Arquivos na pasta exports_nf:")
        try:
            arquivos = os.listdir(pasta)
            if arquivos:
                for arquivo in arquivos:
                    caminho_arquivo = os.path.join(pasta, arquivo)
                    tamanho = os.path.getsize(caminho_arquivo)
                    print(f"   - {arquivo} ({tamanho} bytes)")
            else:
                print("   (pasta vazia)")
        except Exception as e:
            print(f"   Erro ao listar: {e}")
        
        return "Arquivo não encontrado", 404
    
    # Arquivo encontrado
    tamanho = os.path.getsize(caminho)
    print(f"✅ Arquivo encontrado! Tamanho: {tamanho} bytes")
    print("="*70)
    print("🚀 Enviando arquivo para download...")
    
    try:
        return send_file(
            caminho,
            as_attachment=True,
            download_name=nome
        )
    except Exception as e:
        print(f"❌ ERRO no send_file: {str(e)}")
        import traceback
        traceback.print_exc()
        return f"Erro ao enviar arquivo: {str(e)}", 500


@app.route("/notas")
def listar_notas():
    db = get_db()

    notas = db.execute(
        """
        SELECT
            nf.id,
            nf.numero_nf,
            nf.data_emissao,
            nf.valor,
            nf.status,
            nf.link_arquivo,
            a.nome AS aluno,
            r.nome AS responsavel,
            p.numero_parcela
        FROM notas_fiscais nf
        LEFT JOIN parcelas p ON nf.parcela_id = p.id
        LEFT JOIN contratos c ON p.contrato_id = c.id
        LEFT JOIN alunos a ON c.aluno_id = a.id
        LEFT JOIN responsaveis r ON c.responsavel_id = r.id
        ORDER BY nf.data_emissao DESC;
        """
    ).fetchall()

    return render_template("notas.html", notas=notas)


# ═══════════════════════════════════════════════════════════════
# PATCH — notas_a_emitir
# Substitua a função notas_a_emitir() existente por esta versão
# A única diferença é que agora passa todos_ids para o template
# ═══════════════════════════════════════════════════════════════

@app.route("/notas-a-emitir")
def notas_a_emitir():
    try:
        db = get_db()

        de_str  = (request.args.get("de")  or "").strip()
        ate_str = (request.args.get("ate") or "").strip()

        def normaliza_data(s: str):
            if not s:
                return None
            s = s.strip()
            if len(s) == 10 and s[4] == "-" and s[7] == "-":
                y, m, d = s.split("-")
                if y.isdigit() and m.isdigit() and d.isdigit():
                    return s
                return None
            if "/" in s:
                partes = s.split("/")
                if len(partes) != 3:
                    return None
                d, m, y = partes
                if not (d.isdigit() and m.isdigit() and y.isdigit()):
                    return None
                return f"{y}-{m.zfill(2)}-{d.zfill(2)}"
            return None

        de_iso  = normaliza_data(de_str)
        ate_iso = normaliza_data(ate_str)

        page   = request.args.get("page", 1, type=int)
        offset = (page - 1) * PER_PAGE

        sql_base = """
            SELECT
                p.id,
                p.numero_parcela,
                p.valor_final,
                p.data_vencimento,
                a.nome AS aluno,
                r.nome AS responsavel,
                c.id   AS contrato_id,
                c.numero_contrato
            FROM parcelas p
            JOIN contratos    c ON p.contrato_id    = c.id
            JOIN alunos       a ON c.aluno_id       = a.id
            JOIN responsaveis r ON c.responsavel_id = r.id
            LEFT JOIN notas_fiscais nf ON nf.parcela_id = p.id
            WHERE nf.id IS NULL
              AND (p.cancelada = 0 OR p.cancelada IS NULL)
              AND (c.ativo = 1 OR c.ativo IS NULL)
              AND (a.ativo = 1 OR a.ativo IS NULL)
              AND (r.ativo = 1 OR r.ativo IS NULL)
        """

        params = []
        if de_iso:
            sql_base += " AND p.data_vencimento >= ?"
            params.append(de_iso)
        if ate_iso:
            sql_base += " AND p.data_vencimento <= ?"
            params.append(ate_iso)
        sql_base += " ORDER BY p.data_vencimento"

        # Busca TODOS os resultados do filtro
        todas = db.execute(sql_base, params).fetchall()
        total = len(todas)
        total_pages = max(1, math.ceil(total / PER_PAGE)) if total else 1

        # IDs de TODAS as parcelas do filtro (não só da página atual)
        todos_ids = [row["id"] for row in todas]

        # Só a página atual para exibição
        parcelas = todas[offset: offset + PER_PAGE]

        # Marca críticas (vencimento <= 30 dias)
        hoje   = date.today()
        limite = hoje + timedelta(days=30)
        criticas = set()
        for p in parcelas:
            if p["data_vencimento"]:
                y, m, d = map(int, p["data_vencimento"].split("-"))
                if date(y, m, d) <= limite:
                    criticas.add(p["id"])

        return render_template(
            "notas_a_emitir.html",
            parcelas=parcelas,
            criticas=criticas,
            de=de_str,
            ate=ate_str,
            page=page,
            total_pages=total_pages,
            todos_ids=todos_ids,       # NOVO — todos os IDs do filtro
            total_filtro=total,        # NOVO — total real do filtro
        )

    except Exception as e:
        import traceback
        error_traceback = traceback.format_exc()
        print("=" * 80)
        print("ERRO em notas_a_emitir:")
        print(error_traceback)
        print("=" * 80)
        return f"<h1>Erro</h1><pre>{error_traceback}</pre>"

@app.route("/nota/<int:nota_id>")
def detalhe_nota(nota_id):
    db = get_db()

    nota = db.execute(
        """
        SELECT
            nf.id,
            nf.numero_nf,
            nf.data_emissao,
            nf.valor,
            nf.status,

            -- aluno (só para referência)
            a.nome  AS aluno,

            -- tomador = responsável
            r.nome                  AS responsavel,
            r.email                 AS email_tomador,
            r.cpf_cnpj              AS cpf_cnpj_tomador,
            r.endereco_rua          AS rua_tomador,
            r.endereco_numero       AS numero_endereco_tomador,
            r.endereco_complemento  AS complemento_endereco_tomador,
            r.endereco_bairro       AS bairro_tomador,
            r.endereco_cidade       AS cidade_tomador,
            r.endereco_uf           AS uf_tomador,
            r.endereco_cep          AS cep_tomador,

            c.id AS contrato_id
        FROM notas_fiscais nf
        LEFT JOIN parcelas p     ON nf.parcela_id   = p.id
        LEFT JOIN contratos c    ON p.contrato_id   = c.id
        LEFT JOIN alunos a       ON c.aluno_id      = a.id
        LEFT JOIN responsaveis r ON c.responsavel_id = r.id
        WHERE nf.id = ?;
        """,
        (nota_id,),
    ).fetchone()

    if nota is None:
        return "Nota não encontrada", 404

    valor_servicos = nota["valor"] or 0.0
    tributos_aprox = round(valor_servicos * 0.1619, 2)

    discriminacao = (
        f"Parcela referente ao curso de programação e robótica, contrato nº {nota['contrato_id']}. "
        f"Conf lei 12.741-2012, o valor aproximado dos tributos é de R$ {tributos_aprox:.2f}, "
        f"correspondente ao percentual de 16,19%. Fonte IBPT."
    )

    return render_template(
        "nota_detalhe.html",
        nota=nota,
        valor_servicos=valor_servicos,
        tributos_aprox=tributos_aprox,
        discriminacao=discriminacao,
    )



# ═══════════════════════════════════════════════════════════════
# PATCH — modulo_atual e stats da index
# Cole este bloco no app3.py em dois lugares:
#
# 1) A função inject_modulo_atual: ANTES do "if __name__..."
# 2) Substitua a rota def index() existente pela versão abaixo
# ═══════════════════════════════════════════════════════════════

# ── 1) Cole antes do "if __name__ == '__main__':" ─────────────

@app.context_processor
def inject_modulo_atual():
    """
    Injeta a variável 'modulo_atual' em todos os templates
    com base na rota atual, para que o base.html saiba
    qual barra de módulo exibir.
    """
    endpoint = request.endpoint or ""
    if endpoint in ("notas_a_emitir", "listar_notas", "listar_parcelas",
                    "detalhe_nota", "editar_parcela", "preview_lote",
                    "confirmacao_lote", "parcelas_atrasadas"):
        modulo = "notas"
    elif endpoint in ("pagamentos", "registrar_pagamento", "alterar_status_parcela"):
        modulo = "financeiro"
    elif endpoint in ("listar_rematriculas", "detalhe_turma_rematricula",
                      "listar_turmas", "listar_professores",
                      "nova_rematricula", "editar_rematricula",
                      "excluir_rematricula", "nova_turma", "editar_turma",
                      "novo_professor", "editar_professor"):
        modulo = "rematriculas"
    else:
        modulo = ""
    return dict(modulo_atual=modulo)
def gerar_avisos(db):
    """Gera lista de avisos para o quadro da página inicial."""
    from datetime import date, timedelta
    hoje    = date.today()
    em_2sem = (hoje + timedelta(days=14)).isoformat()
    hoje_iso = hoje.isoformat()
 
    avisos = []
 
    # Turmas encerrando em até 2 semanas com alunos de renovação automática
    encerramento = db.execute("""
        SELECT t.id AS turma_id, t.tipo, t.numero, t.horario,
               t.data_fim, p.nome AS professor,
               COUNT(r.id) AS alunos_auto
        FROM turmas t
        LEFT JOIN professores p ON t.professor_id = p.id
        JOIN rematriculas r ON r.turma_id = t.id
            AND r.renovacao_automatica = 1
            AND r.status = 'automatico'
        WHERE t.data_fim IS NOT NULL
          AND t.data_fim <= ?
          AND t.data_fim >= ?
          AND (t.ativa = 1 OR t.ativa IS NULL)
        GROUP BY t.id
        HAVING alunos_auto > 0;
    """, (em_2sem, hoje_iso)).fetchall()
 
    for t in encerramento:
        avisos.append({
            "tipo":     "urgente",
            "titulo":   f"{t['tipo']} {t['numero']} encerra em breve",
            "detalhe":  f"{t['alunos_auto']} aluno(s) com renovação automática precisam ser colocados na próxima turma.",
            "subtexto": f"Encerramento: {t['data_fim']} · Prof. {t['professor'] or '—'}",
            "link":     f"/rematriculas/turma/{t['turma_id']}",
            "link_txt": "Ver turma",
        })
 
    # Alunos com rematrícula pendente de decisão
    pendentes = db.execute("""
        SELECT COUNT(*) AS total FROM rematriculas
        WHERE status = 'pendente'
           OR status = 'em_analise';
    """).fetchone()["total"]
 
    if pendentes:
        avisos.append({
            "tipo":     "atencao",
            "titulo":   f"{pendentes} aluno(s) pendentes de decisão de rematrícula",
            "detalhe":  "Alunos sem status definido de renovação.",
            "subtexto": "",
            "link": "/rematriculas?aba=aluno&status_filtro=pendente",
            "link_txt": "Ver rematrículas",
        })
 
    # Alunos sem turma vinculada
    sem_turma = db.execute("""
        SELECT COUNT(DISTINCT a.id) AS total
        FROM alunos a
        LEFT JOIN rematriculas r ON r.aluno_id = a.id
        WHERE (a.ativo = 1 OR a.ativo IS NULL)
          AND r.id IS NULL;
    """).fetchone()["total"]
 
    if sem_turma:
        avisos.append({
            "tipo":     "info",
            "titulo":   f"{sem_turma} aluno(s) sem turma vinculada",
            "detalhe":  "Alunos ativos que ainda não foram associados a nenhuma turma.",
            "subtexto": "",
            "link":     "/alunos?aba=ativos&sem_turma=1",
            "link_txt": "Ver alunos",
        })
 
    # Parcelas atrasadas
    atrasadas = db.execute("""
        SELECT COUNT(*) AS total FROM parcelas p
        JOIN contratos c ON p.contrato_id = c.id
        WHERE p.status_pagamento = 'atrasado'
          AND (p.cancelada = 0 OR p.cancelada IS NULL)
          AND (c.ativo = 1 OR c.ativo IS NULL);
    """).fetchone()["total"]
 
    if atrasadas:
        avisos.append({
            "tipo":     "atencao",
            "titulo":   f"{atrasadas} parcela(s) em atraso",
            "detalhe":  "Parcelas vencidas sem registro de pagamento.",
            "subtexto": "",
            "link":     "/pagamentos?aba=mes&status_filtro=atrasado",
            "link_txt": "Ver parcelas",
        })
 
    # Notas a emitir
    import calendar as _cal
    ultimo_dia_mes = hoje.replace(
        day=_cal.monthrange(hoje.year, hoje.month)[1]
    ).isoformat()

    notas = db.execute("""
        SELECT COUNT(*) AS total FROM parcelas p
        JOIN contratos c ON p.contrato_id = c.id
        WHERE NOT EXISTS (SELECT 1 FROM notas_fiscais nf WHERE nf.parcela_id = p.id)
        AND (p.cancelada = 0 OR p.cancelada IS NULL)
        AND (c.ativo = 1 OR c.ativo IS NULL)
        AND p.data_vencimento <= ?;
    """, (ultimo_dia_mes,)).fetchone()["total"]
 
    if notas:
        avisos.append({
            "tipo":     "info",
            "titulo":   f"{notas} nota(s) fiscal(is) a emitir",
            "detalhe":  "Parcelas com vencimento passado sem nota emitida.",
            "subtexto": "",
            "link":     "/notas-a-emitir",
            "link_txt": "Ir para notas",
        })
 
    return avisos
@app.route("/")
def index():
    db = get_db()
    hoje     = date.today()
    hoje_iso = hoje.isoformat()
    primeiro = hoje.replace(day=1).isoformat()
 
    # Atualiza atrasadas automaticamente
    db.execute("""
        UPDATE parcelas SET status_pagamento = 'atrasado'
        WHERE (status_pagamento = 'pendente' OR status_pagamento IS NULL)
          AND data_vencimento < ?
          AND (cancelada = 0 OR cancelada IS NULL);
    """, (hoje_iso,))
    db.commit()
 
    stats = {
        "total_alunos": db.execute(
            "SELECT COUNT(*) FROM alunos WHERE ativo = 1 OR ativo IS NULL;"
        ).fetchone()[0],
        "parcelas_pagas_mes": db.execute("""
            SELECT COUNT(*) FROM parcelas
            WHERE status_pagamento = 'pago'
              AND pago_em >= ? AND pago_em <= ?
              AND (cancelada = 0 OR cancelada IS NULL);
        """, (primeiro, hoje_iso)).fetchone()[0],
        "parcelas_atrasadas": db.execute("""
            SELECT COUNT(*) FROM parcelas
            WHERE status_pagamento = 'atrasado'
              AND (cancelada = 0 OR cancelada IS NULL);
        """).fetchone()[0],
        "notas_a_emitir": db.execute("""
            SELECT COUNT(*) FROM parcelas p
            JOIN contratos c ON p.contrato_id = c.id
            WHERE NOT EXISTS (SELECT 1 FROM notas_fiscais nf WHERE nf.parcela_id = p.id)
              AND (p.cancelada = 0 OR p.cancelada IS NULL)
              AND (c.ativo = 1 OR c.ativo IS NULL)
              AND p.data_vencimento <= ?;
        """, (hoje_iso,)).fetchone()[0],
    }
 
    avisos = gerar_avisos(db)
 
    return render_template("index.html", stats=stats, avisos=avisos)
# =========================
# LISTA DE ALUNOS
# =========================
@app.post("/responsavel/<int:resp_id>/excluir_completo")
def excluir_responsavel_completo(resp_id):
    db = get_db()

    resp = db.execute(
        "SELECT id, nome FROM responsaveis WHERE id = ?;",
        (resp_id,),
    ).fetchone()

    if resp is None:
        return "Responsável não encontrado", 404

    # 1) Desativar responsável
    db.execute(
        "UPDATE responsaveis SET ativo = 0 WHERE id = ?;",
        (resp_id,),
    )

    # 2) Desativar alunos desse responsável
    db.execute(
        "UPDATE alunos SET ativo = 0 WHERE responsavel_id = ?;",
        (resp_id,),
    )

    # 3) Desativar contratos desse responsável
    db.execute(
        "UPDATE contratos SET ativo = 0 WHERE responsavel_id = ?;",
        (resp_id,),
    )

    # 4) Cancelar APENAS parcelas ainda sem NF
    db.execute(
        """
        UPDATE parcelas
        SET cancelada = 1
        WHERE contrato_id IN (
            SELECT id FROM contratos WHERE responsavel_id = ?
        )
        AND id NOT IN (
            SELECT parcela_id FROM notas_fiscais
            WHERE parcela_id IS NOT NULL
        );
        """,
        (resp_id,),
    )

    db.commit()
    return redirect(url_for("listar_alunos"))


from flask import request, render_template

@app.route("/alunos")
def listar_alunos():
    db  = get_db()
    aba = request.args.get("aba", "ativos")

    # ← adicione estas linhas aqui
    sort         = request.args.get("sort", "aluno")
    direction    = request.args.get("dir",  "asc")
    sem_turma    = ""
    total_sem_turma = 0
    sort_map     = {"id": "a.id", "aluno": "a.nome"}
    sort_sql     = sort_map.get(sort, "a.nome")
    dir_sql      = "DESC" if direction.lower() == "desc" else "ASC"
    order_by     = f"{sort_sql} {dir_sql}, a.id ASC"
 
    if aba == "historico":
        sem_turma = ""
        total_sem_turma = 0
        alunos = db.execute(f"""
            SELECT
                a.id,
                a.nome          AS aluno,
                r.nome          AS responsavel,
                a.ativo,
                (SELECT cl.data_cancelamento
                 FROM cancelamentos cl
                 JOIN contratos c2 ON cl.contrato_id = c2.id
                 WHERE c2.aluno_id = a.id AND cl.reativado_em IS NULL
                 ORDER BY cl.data_cancelamento DESC LIMIT 1
                ) AS data_cancelamento,
                (SELECT cl.motivo
                 FROM cancelamentos cl
                 JOIN contratos c2 ON cl.contrato_id = c2.id
                 WHERE c2.aluno_id = a.id AND cl.reativado_em IS NULL
                 ORDER BY cl.data_cancelamento DESC LIMIT 1
                ) AS motivo_cancelamento,
                (SELECT cl.houve_ressarcimento
                 FROM cancelamentos cl
                 JOIN contratos c2 ON cl.contrato_id = c2.id
                 WHERE c2.aluno_id = a.id AND cl.reativado_em IS NULL
                 ORDER BY cl.data_cancelamento DESC LIMIT 1
                ) AS houve_ressarcimento,
                (SELECT cl.valor_devolvido
                 FROM cancelamentos cl
                 JOIN contratos c2 ON cl.contrato_id = c2.id
                 WHERE c2.aluno_id = a.id AND cl.reativado_em IS NULL
                 ORDER BY cl.data_cancelamento DESC LIMIT 1
                ) AS valor_devolvido
            FROM alunos a
            LEFT JOIN responsaveis r ON a.responsavel_id = r.id
            WHERE a.ativo = 0
            ORDER BY {order_by};
        """).fetchall()
    else:
        sem_turma = request.args.get("sem_turma", "")
        if sem_turma:
            alunos = db.execute(f"""
                SELECT a.id, a.nome AS aluno, r.nome AS responsavel, a.ativo
                FROM alunos a
                LEFT JOIN responsaveis r ON a.responsavel_id = r.id
                WHERE a.ativo = 1
                  AND NOT EXISTS (
                      SELECT 1 FROM rematriculas rm WHERE rm.aluno_id = a.id
                  )
                ORDER BY {order_by};
            """).fetchall()
        else:
            sem_turma = ""
            alunos = db.execute(f"""
                SELECT a.id, a.nome AS aluno, r.nome AS responsavel, a.ativo
                FROM alunos a
                LEFT JOIN responsaveis r ON a.responsavel_id = r.id
                WHERE a.ativo = 1
                ORDER BY {order_by};
            """).fetchall()
 
    total_ativos    = db.execute("SELECT COUNT(*) FROM alunos WHERE ativo = 1;").fetchone()[0]
    total_historico = db.execute("SELECT COUNT(*) FROM alunos WHERE ativo = 0;").fetchone()[0]
    total_sem_turma = db.execute("""
        SELECT COUNT(*) FROM alunos a
        WHERE (a.ativo = 1 OR a.ativo IS NULL)
          AND NOT EXISTS (SELECT 1 FROM rematriculas rm WHERE rm.aluno_id = a.id);
    """).fetchone()[0]
 
    return render_template(
        "alunos.html",
        alunos=alunos,
        aba=aba,
        sort=sort,
        dir=direction,
        sem_turma=sem_turma,
        total_ativos=total_ativos,
        total_historico=total_historico,
        total_sem_turma=total_sem_turma,
    )



@app.route("/aluno/<int:aluno_id>")
def detalhe_aluno(aluno_id):
    db = get_db()
 
    aluno = db.execute("""
        SELECT a.id, a.nome AS aluno, a.data_nascimento, a.ativo,
               a.responsavel_id, r.nome AS responsavel,
               r.telefone, r.email
        FROM alunos a
        LEFT JOIN responsaveis r ON a.responsavel_id = r.id
        WHERE a.id = ?;
    """, (aluno_id,)).fetchone()
 
    if aluno is None:
        return "Aluno não encontrado", 404
 
    # Contratos ativos
    contratos_ativos = db.execute("""
        SELECT c.id, c.numero_contrato, c.data_inicio, c.data_fim,
               c.valor_total, c.qtd_parcelas, c.observacoes, c.ativo
        FROM contratos c
        WHERE c.aluno_id = ? AND c.ativo = 1
        ORDER BY c.data_inicio DESC;
    """, (aluno_id,)).fetchall()
 
    # Contratos inativos
    contratos_inativos = db.execute("""
        SELECT c.id, c.numero_contrato, c.data_inicio, c.data_fim,
               c.valor_total, c.qtd_parcelas, c.observacoes, c.ativo,
               c.excluido_em, c.cancelado,
               CASE
                   WHEN c.cancelado = 1               THEN 'cancelado'
                   WHEN c.excluido_em IS NOT NULL      THEN 'excluido'
                   WHEN c.data_fim < date('now')       THEN 'vencido'
                   ELSE 'inativo'
               END as status
        FROM contratos c
        WHERE c.aluno_id = ? AND c.ativo = 0
        ORDER BY COALESCE(c.excluido_em, c.data_fim) DESC;
    """, (aluno_id,)).fetchall()
 
    # ── REMATRÍCULAS do aluno ─────────────────────────────────
    rematriculas_raw = db.execute("""
        SELECT r.id, r.status, r.tipo_contrato, r.desconto_anterior,
               r.motivo_nao_renovou, r.observacoes, r.periodo,
               r.renovacao_automatica,
               r.turma_id,
               t.tipo      AS turma_tipo,
               t.numero    AS turma_numero,
               t.horario,  t.dia_semana,
               t.data_inicio, t.data_fim,
               t.periodo   AS turma_periodo,
               p.nome      AS professor
        FROM rematriculas r
        LEFT JOIN turmas t      ON r.turma_id      = t.id
        LEFT JOIN professores p ON t.professor_id  = p.id
        WHERE r.aluno_id = ?
        ORDER BY r.id DESC;
    """, (aluno_id,)).fetchall()
 
    # Calcula percentual do curso para cada rematrícula
    from datetime import date as _date
    rematriculas = []
    for r in rematriculas_raw:
        row = dict(r)
        try:
            if r["data_inicio"] and r["data_fim"]:
                ini = _date.fromisoformat(r["data_inicio"])
                fim = _date.fromisoformat(r["data_fim"])
                hoje_d = _date.today()
                total_dias = (fim - ini).days or 1
                dias_passados = (hoje_d - ini).days
                pct = round(max(0, min(100, (dias_passados / total_dias) * 100)))
            else:
                pct = 0
        except Exception:
            pct = 0
        row["pct_curso"] = pct
        rematriculas.append(row)
 
    # ── Parcelas recentes (últimas 6) ─────────────────────────
    parcelas_recentes = db.execute("""
        SELECT p.id, p.numero_parcela, p.valor_final,
               p.data_vencimento, p.status_pagamento,
               p.pago_em, p.canal_pagamento,
               c.numero_contrato
        FROM parcelas p
        JOIN contratos c ON p.contrato_id = c.id
        WHERE c.aluno_id = ?
          AND (p.cancelada = 0 OR p.cancelada IS NULL)
        ORDER BY p.data_vencimento DESC
        LIMIT 6;
    """, (aluno_id,)).fetchall()
 
    # ── Turmas disponíveis para vínculo ──────────────────────
    turmas_disponiveis = db.execute("""
        SELECT t.id, t.tipo, t.numero, t.horario, t.dia_semana,
               t.periodo, p.nome AS professor
        FROM turmas t
        LEFT JOIN professores p ON t.professor_id = p.id
        WHERE (t.ativa = 1 OR t.ativa IS NULL)
        ORDER BY t.tipo, t.numero, t.horario;
    """).fetchall()
 
    return render_template(
        "aluno_detalhe.html",
        aluno=aluno,
        contratos_ativos=contratos_ativos,
        contratos_inativos=contratos_inativos,
        rematriculas=rematriculas,
        parcelas_recentes=parcelas_recentes,
        turmas_disponiveis=turmas_disponiveis,
        hoje=date.today().isoformat(),
    )
@app.post("/contrato/<int:contrato_id>/reativar")
def reativar_contrato(contrato_id):
    """Reativa um contrato inativo."""
    db = get_db()
    
    contrato = db.execute(
        "SELECT aluno_id FROM contratos WHERE id = ?;",
        (contrato_id,)
    ).fetchone()
    
    if contrato:
        db.execute(
            "UPDATE contratos SET ativo = 1, excluido_em = NULL WHERE id = ?;",
            (contrato_id,)
        )
        
        # Reativa parcelas canceladas (mas não as que têm NF)
        db.execute("""
            UPDATE parcelas 
            SET cancelada = 0, excluido_em = NULL 
            WHERE contrato_id = ? 
            AND id NOT IN (
                SELECT parcela_id FROM notas_fiscalis WHERE parcela_id IS NOT NULL
            );
        """, (contrato_id,))
        
        db.commit()
        
    return redirect(url_for('detalhe_aluno', aluno_id=contrato['aluno_id']))


# =========================
# DETALHES DE UM CONTRATO
# =========================
@app.route("/contrato/<int:contrato_id>")
def detalhe_contrato(contrato_id):
    db = get_db()
    
    contrato = db.execute(
        """
        SELECT
            c.id,
            c.data_inicio,
            c.data_fim,
            c.valor_total,
            c.qtd_parcelas,
            c.observacoes,
            c.aluno_id,
            c.ativo,
            a.nome AS aluno,
            r.nome AS responsavel
        FROM contratos c
        JOIN alunos a ON c.aluno_id = a.id
        JOIN responsaveis r ON c.responsavel_id = r.id
        WHERE c.id = ?;
        """,
        (contrato_id,),
    ).fetchone()
    
    if contrato is None:
        return "Contrato não encontrado", 404
    
    parcelas = db.execute(
        """
        SELECT
            p.id,
            p.numero_parcela,
            p.valor_final,
            p.data_vencimento,
            nf.id AS nf_id,
            nf.status AS nf_status,
            nf.numero_nf,
            CASE
                WHEN nf.id IS NOT NULL AND nf.status = 'emitida'
                THEN 1
                ELSE 0
            END AS nf_emitida
        FROM parcelas p
        LEFT JOIN notas_fiscais nf ON nf.parcela_id = p.id
        WHERE p.contrato_id = ?
        ORDER BY p.numero_parcela;
        """,
        (contrato_id,),
    ).fetchall()
    
    # VERIFICA SE TEM ALGUMA NF EMITIDA NESTE CONTRATO
    tem_nf_emitida = False
    for parcela in parcelas:
        if parcela['nf_id'] is not None:
            tem_nf_emitida = True
            break
    
    return render_template(
        "contrato_detalhe.html",
        contrato=contrato,
        parcelas=parcelas,
        tem_nf_emitida=tem_nf_emitida,  # <- NOVA VARIÁVEL
    )
# =========================
# TODAS AS PARCELAS
# =========================
@app.route("/parcelas")
def listar_parcelas():
    db = get_db()

    page = request.args.get("page", 1, type=int)
    offset = (page - 1) * PER_PAGE

    # total de registros (para calcular número de páginas)
    total = db.execute("SELECT COUNT(*) FROM parcelas;").fetchone()[0]
    total_pages = max(1, math.ceil(total / PER_PAGE))

    parcelas = db.execute(
        """
        SELECT
            p.id,
            p.numero_parcela,
            p.valor_final,
            p.data_vencimento,
            p.cancelada,
            a.nome AS aluno,
            r.nome AS responsavel
        FROM parcelas p
        JOIN contratos c    ON p.contrato_id    = c.id
        JOIN alunos a       ON c.aluno_id       = a.id
        JOIN responsaveis r ON c.responsavel_id = r.id
        ORDER BY p.data_vencimento
        LIMIT ? OFFSET ?;
        """,
        (PER_PAGE, offset),
    ).fetchall()

    return render_template(
        "parcelas.html",
        parcelas=parcelas,
        page=page,
        total_pages=total_pages,
    )

# =========================
# PARCELAS ATRASADAS
# =========================
'''
@app.route("/parcelas/atrasadas")
def parcelas_atrasadas():
    db = get_db()
    parcelas = db.execute(
        """
        SELECT
            p.id,
            p.numero_parcela,
            p.valor_final,
            p.data_vencimento,
            p.data_pagamento,
            p.status,
            a.nome AS aluno,
            r.nome AS responsavel
        FROM parcelas p
        JOIN contratos c ON p.contrato_id = c.id
        JOIN alunos a ON c.aluno_id = a.id
        JOIN responsaveis r ON c.responsavel_id = r.id
        WHERE p.status = 'atrasada'
        ORDER BY p.data_vencimento;
        """
    ).fetchall()

    return render_template("parcelas_atrasadas.html", parcelas=parcelas)

@app.route("/export/parcelas_atrasadas.csv")
def export_parcelas_atrasadas():
    db = get_db()

    rows = db.execute(
        """
        SELECT
            p.id,
            a.nome AS aluno,
            r.nome AS responsavel,
            p.numero_parcela,
            p.valor_final,
            p.data_vencimento,
            p.data_pagamento,
            p.status,
            p.forma_pagamento
        FROM parcelas p
        JOIN contratos c ON p.contrato_id = c.id
        JOIN alunos a ON c.aluno_id = a.id
        JOIN responsaveis r ON c.responsavel_id = r.id
        WHERE p.status = 'atrasada'
        ORDER BY p.data_vencimento;
        """
    ).fetchall()

    # Monta CSV em memória
    lines = []
    header = [
        "id_parcela",
        "aluno",
        "responsavel",
        "numero_parcela",
        "valor_final",
        "data_vencimento",
        "data_pagamento",
        "status",
        "forma_pagamento",
    ]
    lines.append(";".join(header))

    for r in rows:
        line = [
            str(r["id"]),
            r["aluno"],
            r["responsavel"],
            str(r["numero_parcela"]),
            f"{r['valor_final']:.2f}",
            r["data_vencimento"] or "",
            r["data_pagamento"] or "",
            r["status"] or "",
            r["forma_pagamento"] or "",
        ]
        # usa ; como separador (brasileiro)
        lines.append(";".join(line))

    csv_content = "\n".join(lines)

    return Response(
        csv_content,
        mimetype="text/csv",
        headers={"Content-Disposition": "attachment; filename=parcelas_atrasadas.csv"},
    )
'''

# =========================
# MARCAR PARCELA COMO PAGA
# =========================
@app.post("/parcelas/<int:parcela_id>/pagar")
def marcar_parcela_paga(parcela_id):
    db = get_db()
    hoje = date.today().isoformat()  # 'YYYY-MM-DD'

    db.execute(
        """
        UPDATE parcelas
        SET status = 'paga',
            data_pagamento = ?
        WHERE id = ?;
        """,
        (hoje, parcela_id),
    )
    db.commit()

    # Volta para a página de onde veio, se possível
    ref = request.headers.get("Referer")
    if ref:
        return redirect(ref)
    return redirect(url_for("listar_parcelas"))

# ═══════════════════════════════════════════════════════════════
# PATCH — novo_cadastro
# Substitua a função novo_cadastro() existente no app3.py
# Adiciona vínculo de turma e renovação automática ao cadastro
# ═══════════════════════════════════════════════════════════════

@app.route("/novo-cadastro", methods=["GET", "POST"])
def novo_cadastro():
    db = get_db()

    # Turmas disponíveis para o select
    turmas_disponiveis = db.execute("""
        SELECT t.id, t.tipo, t.numero, t.horario, t.dia_semana,
               t.periodo, p.nome AS professor
        FROM turmas t
        LEFT JOIN professores p ON t.professor_id = p.id
        WHERE (t.ativa = 1 OR t.ativa IS NULL)
        ORDER BY t.tipo, t.numero, t.horario;
    """).fetchall()

    if request.method == "POST":
        # ---------- RESPONSÁVEL ----------
        r_nome   = request.form.get("r_nome")
        r_tel    = request.form.get("r_telefone")
        r_email  = request.form.get("r_email")
        r_cpf    = request.form.get("r_cpf_cnpj")
        r_rua    = request.form.get("r_endereco_rua")
        r_num    = request.form.get("r_endereco_numero")
        r_comp   = request.form.get("r_endereco_complemento")
        r_bairro = request.form.get("r_endereco_bairro")
        r_cidade = request.form.get("r_endereco_cidade")
        r_uf     = request.form.get("r_endereco_uf")
        r_cep    = request.form.get("r_endereco_cep")

        if not validar_campo_cpf_cnpj(r_cpf):
            erro = "CPF inválido. Confira o número digitado."
            return render_template("cadastro_novo.html", erro=erro,
                                   turmas_disponiveis=turmas_disponiveis,
                                   hoje=date.today().isoformat())

        # ---------- CONTRATO ----------
        c_numero = request.form.get("c_numero_contrato", "").strip()
        if not c_numero:
            erro = "Número do contrato é obrigatório!"
            return render_template("cadastro_novo.html", erro=erro,
                                   turmas_disponiveis=turmas_disponiveis,
                                   hoje=date.today().isoformat())

        existe = db.execute(
            "SELECT id FROM contratos WHERE numero_contrato = ?;", (c_numero,)
        ).fetchone()
        if existe:
            erro = f"Já existe um contrato com o número '{c_numero}'!"
            return render_template("cadastro_novo.html", erro=erro,
                                   turmas_disponiveis=turmas_disponiveis,
                                   hoje=date.today().isoformat())

        # ---------- TURMA / REMATRÍCULA ----------
        turma_id             = request.form.get("turma_id", "").strip()
        renovacao_automatica = 1 if request.form.get("renovacao_automatica") == "1" else 0
        periodo_rem          = request.form.get("periodo_rem", "").strip()

        # ---------- INSERE RESPONSÁVEL ----------
        cur = db.execute("""
            INSERT INTO responsaveis (
                nome, telefone, email, cpf_cnpj,
                endereco_rua, endereco_numero, endereco_complemento,
                endereco_bairro, endereco_cidade, endereco_uf, endereco_cep
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
        """, (r_nome, r_tel, r_email, r_cpf,
              r_rua, r_num, r_comp, r_bairro, r_cidade, r_uf, r_cep))
        responsavel_id = cur.lastrowid

        # ---------- ALUNO ----------
        a_nome = request.form.get("a_nome")
        a_nasc = request.form.get("a_data_nascimento")
        cur = db.execute(
            "INSERT INTO alunos (nome, data_nascimento, responsavel_id) VALUES (?, ?, ?);",
            (a_nome, a_nasc, responsavel_id)
        )
        aluno_id = cur.lastrowid

        # ---------- CONTRATO ----------
        c_inicio      = request.form.get("c_data_inicio")
        c_fim         = request.form.get("c_data_fim")
        c_valor_total = request.form.get("c_valor_total")
        c_qtd_parc    = request.form.get("c_qtd_parcelas")

        valor_total = float(c_valor_total.replace(",", ".")) if c_valor_total else 0.0
        qtd_parc    = int(c_qtd_parc) if c_qtd_parc else 0

        cur = db.execute("""
            INSERT INTO contratos (
                aluno_id, responsavel_id, numero_contrato,
                data_inicio, data_fim, valor_total, qtd_parcelas
            ) VALUES (?, ?, ?, ?, ?, ?, ?);
        """, (aluno_id, responsavel_id, c_numero,
              c_inicio, c_fim, valor_total, qtd_parc))
        contrato_id = cur.lastrowid

        # ---------- PARCELAS ----------
        if contrato_id and qtd_parc > 0:
            valor_base = round(valor_total / qtd_parc, 2)
            acumulado  = 0.0
            for i in range(1, qtd_parc + 1):
                if i < qtd_parc:
                    valor_parcela = valor_base
                    acumulado += valor_parcela
                else:
                    valor_parcela = round(valor_total - acumulado, 2)
                data_venc = add_months(c_inicio, i - 1)
                db.execute("""
                    INSERT INTO parcelas (
                        contrato_id, numero_parcela,
                        valor_original, valor_final, data_vencimento
                    ) VALUES (?, ?, ?, ?, ?);
                """, (contrato_id, i, valor_parcela, valor_parcela, data_venc))

        # ---------- REMATRÍCULA (opcional) ----------
        if turma_id or renovacao_automatica:
            status_rem = "automatico" if renovacao_automatica else "pendente"
            db.execute("""
                INSERT INTO rematriculas (
                    aluno_id, turma_id, periodo,
                    tipo_contrato, status, renovacao_automatica
                ) VALUES (?, ?, ?, ?, ?, ?);
            """, (
                aluno_id,
                turma_id or None,
                periodo_rem or None,
                "automatico" if renovacao_automatica else "renovar",
                status_rem,
                renovacao_automatica,
            ))

        db.commit()
        return redirect(url_for("detalhe_aluno", aluno_id=aluno_id))

    # GET
    return render_template("cadastro_novo.html",
                           hoje=date.today().isoformat(),
                           turmas_disponiveis=turmas_disponiveis)
@app.route("/buscar")
def buscar():
    termo = request.args.get("q", "").strip()
    db = get_db()
    
    if not termo:
        return render_template("buscar.html", resultados=[], termo="")
    
    try:
        # Query COMPLETA que busca em alunos, responsáveis E número do contrato
        resultados = db.execute("""
            SELECT DISTINCT
                a.id as aluno_id,
                a.nome AS aluno,
                r.nome AS responsavel,
                c.numero_contrato,
                c.id as contrato_id,
                c.data_fim,
                CASE 
                    WHEN c.data_fim < date('now') THEN 'vencido'
                    WHEN c.data_fim BETWEEN date('now') AND date('now', '+30 days') THEN 'a_vencer'
                    ELSE 'ativo'
                END as status_contrato
            FROM alunos a
            JOIN responsaveis r ON a.responsavel_id = r.id
            LEFT JOIN contratos c ON a.id = c.aluno_id
            WHERE a.ativo = 1
              AND (
                -- Busca no nome do aluno
                LOWER(a.nome) LIKE LOWER(?)
                -- Busca no nome do responsável
                OR LOWER(r.nome) LIKE LOWER(?)
                -- Busca no número do contrato (CORRIGIDO)
                OR c.numero_contrato LIKE ?
                -- Busca no telefone do responsável (opcional)
                OR r.telefone LIKE ?
              )
            ORDER BY a.nome
            LIMIT 50;
        """, (f"%{termo}%", f"%{termo}%", f"%{termo}%", f"%{termo}%")).fetchall()
        
        # Se não encontrou pelo LEFT JOIN, tenta buscar direto nos contratos
        if not resultados:
            resultados = db.execute("""
                SELECT 
                    c.id as contrato_id,
                    c.numero_contrato,
                    c.data_fim,
                    a.id as aluno_id,
                    a.nome AS aluno,
                    r.nome AS responsavel,
                    CASE 
                        WHEN c.data_fim < date('now') THEN 'vencido'
                        WHEN c.data_fim BETWEEN date('now') AND date('now', '+30 days') THEN 'a_vencer'
                        ELSE 'ativo'
                    END as status_contrato
                FROM contratos c
                JOIN alunos a ON c.aluno_id = a.id
                JOIN responsaveis r ON c.responsavel_id = r.id
                WHERE c.numero_contrato LIKE ?
                ORDER BY c.numero_contrato
                LIMIT 20;
            """, (f"%{termo}%",)).fetchall()
        
        return render_template("buscar.html", 
                             resultados=resultados, 
                             termo=termo)
                             
    except Exception as e:
        import traceback
        error_msg = f"Erro na busca: {str(e)}"
        print(f"ERRO: {error_msg}")
        print(traceback.format_exc())
        
        # Fallback: busca apenas em alunos e responsáveis
        resultados = db.execute("""
            SELECT 
                a.id as aluno_id,
                a.nome AS aluno,
                r.nome AS responsavel
            FROM alunos a
            JOIN responsaveis r ON a.responsavel_id = r.id
            WHERE a.ativo = 1
              AND (a.nome LIKE ? OR r.nome LIKE ?)
            ORDER BY a.nome
            LIMIT 20;
        """, (f"%{termo}%", f"%{termo}%")).fetchall()
        
        return render_template("buscar.html", 
                             resultados=resultados, 
                             termo=termo,
                             erro_tecnico=str(e))
@app.route("/contrato/buscar/<numero>")
def buscar_contrato(numero):
    db = get_db()
    
    contrato = db.execute("""
        SELECT c.*, a.nome as aluno, r.nome as responsavel
        FROM contratos c
        JOIN alunos a ON c.aluno_id = a.id
        JOIN responsaveis r ON c.responsavel_id = r.id
        WHERE c.numero_contrato = ?;
    """, (numero,)).fetchone()
    
    if not contrato:
        return "Contrato não encontrado", 404
    
    return redirect(url_for('detalhe_contrato', contrato_id=contrato['id']))
@app.route("/debug/colunas")
def debug_colunas():
    db = get_db()
    
    # Verifica colunas da tabela contratos
    colunas = db.execute("PRAGMA table_info(contratos);").fetchall()
    
    resultado = "<h1>Colunas da tabela 'contratos':</h1><ul>"
    for col in colunas:
        resultado += f"<li>{col[1]} ({col[2]}) - {'NOT NULL' if col[3] else 'NULLABLE'}</li>"
    resultado += "</ul>"
    
    return resultado
@app.route("/responsavel/<int:resp_id>/editar", methods=["GET", "POST"])
def editar_responsavel(resp_id):
    db = get_db()

    if request.method == "POST":
        nome   = request.form.get("nome")
        telefone = request.form.get("telefone")
        email  = request.form.get("email")
        cpf_cnpj = request.form.get("cpf_cnpj")
        endereco_rua = request.form.get("endereco_rua")
        endereco_numero = request.form.get("endereco_numero")
        endereco_complemento = request.form.get("endereco_complemento")
        endereco_bairro = request.form.get("endereco_bairro")
        endereco_cidade = request.form.get("endereco_cidade")
        endereco_uf = request.form.get("endereco_uf")
        endereco_cep = request.form.get("endereco_cep")

        # >>> VALIDA CPF/CNPJ <<<
        if not validar_campo_cpf_cnpj(cpf_cnpj):
            erro = "CPF inválido. Confira o número digitado."

            # monta um "objeto" simples pra preencher o formulário de volta
            r = {
                "id": resp_id,
                "nome": nome,
                "telefone": telefone,
                "email": email,
                "cpf_cnpj": cpf_cnpj,
                "endereco_rua": endereco_rua,
                "endereco_numero": endereco_numero,
                "endereco_complemento": endereco_complemento,
                "endereco_bairro": endereco_bairro,
                "endereco_cidade": endereco_cidade,
                "endereco_uf": endereco_uf,
                "endereco_cep": endereco_cep,
            }
            return render_template("responsavel_editar.html", r=r, erro=erro)

        db.execute(
            """
            UPDATE responsaveis
            SET nome = ?, telefone = ?, email = ?, cpf_cnpj = ?,
                endereco_rua = ?, endereco_numero = ?, endereco_complemento = ?,
                endereco_bairro = ?, endereco_cidade = ?, endereco_uf = ?, endereco_cep = ?
            WHERE id = ?;
            """,
            (
                nome, telefone, email, cpf_cnpj,
                endereco_rua, endereco_numero, endereco_complemento,
                endereco_bairro, endereco_cidade, endereco_uf, endereco_cep,
                resp_id,
            ),
        )
        db.commit()
        return redirect(url_for("listar_alunos"))

    # GET
    resp = db.execute(
        "SELECT * FROM responsaveis WHERE id = ?;",
        (resp_id,),
    ).fetchone()

    if resp is None:
        return "Responsável não encontrado", 404

@app.route("/parcela/<int:parcela_id>/editar", methods=["GET", "POST"])
def editar_parcela(parcela_id):
    db = get_db()

    # GET - carregar dados da parcela
    if request.method == "GET":
        parcela = db.execute(
            """
            SELECT
                p.id,
                p.numero_parcela,
                p.valor_final,
                p.data_vencimento,
                p.contrato_id,
                nf.id AS nf_id
            FROM parcelas p
            LEFT JOIN notas_fiscais nf
                ON nf.parcela_id = p.id
                AND nf.status = 'emitida'
            WHERE p.id = ?;
            """,
            (parcela_id,)
        ).fetchone()

        if parcela is None:
            return "Parcela não encontrada", 404

        if parcela["nf_id"] is not None:
            return "Esta parcela já possui nota emitida e não pode ser alterada.", 403

        return render_template("editar_parcela.html", parcela=parcela)

    # POST - atualizar dados
    elif request.method == "POST":
        valor = request.form["valor_final"]
        vencimento = request.form["data_vencimento"]

        # Verifica se tem NF emitida primeiro
        parcela = db.execute(
            """
            SELECT p.contrato_id, nf.id AS nf_id
            FROM parcelas p
            LEFT JOIN notas_fiscais nf ON nf.parcela_id = p.id
            WHERE p.id = ?;
            """,
            (parcela_id,)
        ).fetchone()

        if parcela is None:
            return "Parcela não encontrada", 404

        if parcela["nf_id"] is not None:
            return "Esta parcela já possui nota emitida e não pode ser alterada.", 403

        # Atualiza a parcela
        db.execute(
            """
            UPDATE parcelas
            SET valor_final = ?, data_vencimento = ?
            WHERE id = ?;
            """,
            (valor, vencimento, parcela_id)
        )

        # Recalcula valor total do contrato
        total = db.execute(
            """
            SELECT SUM(valor_final)
            FROM parcelas
            WHERE contrato_id = ? AND (cancelada = 0 OR cancelada IS NULL);
            """,
            (parcela["contrato_id"],)
        ).fetchone()[0] or 0

        # Atualiza o contrato
        db.execute(
            """
            UPDATE contratos
            SET valor_total = ?
            WHERE id = ?;
            """,
            (total, parcela["contrato_id"])
        )

        db.commit()

        return redirect(url_for("detalhe_contrato", contrato_id=parcela["contrato_id"]))
@app.route("/aluno/<int:aluno_id>/novo_contrato", methods=["GET", "POST"])
def novo_contrato_aluno(aluno_id):
    db = get_db()
    
    aluno = db.execute(
        """
        SELECT a.id, a.nome, a.responsavel_id, r.nome AS responsavel
        FROM alunos a
        JOIN responsaveis r ON a.responsavel_id = r.id
        WHERE a.id = ?;
        """,
        (aluno_id,)
    ).fetchone()
    
    if aluno is None:
        return "Aluno não encontrado", 404
    
    if request.method == "POST":
        numero_contrato = request.form.get("numero_contrato", "").strip()
        data_inicio = request.form.get("data_inicio")
        data_fim = request.form.get("data_fim")
        valor_total = request.form.get("valor_total")
        qtd_parcelas = request.form.get("qtd_parcelas")
        observacoes = request.form.get("observacoes")
        
        # VALIDA NÚMERO DO CONTRATO
        if not numero_contrato:
            return render_template(
                "novo_contrato.html",
                aluno=aluno,
                erro="Número do contrato é obrigatório!",
                hoje=date.today().isoformat()
            )
        
        # Verifica se número já existe
        existe = db.execute(
            "SELECT id FROM contratos WHERE numero_contrato = ?;",
            (numero_contrato,)
        ).fetchone()
        
        if existe:
            return render_template(
                "novo_contrato.html",
                aluno=aluno,
                erro=f"Já existe um contrato com o número '{numero_contrato}'!",
                hoje=date.today().isoformat()
            )
        
        valor_total_float = float(valor_total.replace(",", ".")) if valor_total else 0.0
        qtd_parcelas_int = int(qtd_parcelas) if qtd_parcelas else 0
        
        cur = db.execute(
            """
            INSERT INTO contratos (
                aluno_id, responsavel_id, numero_contrato,
                data_inicio, data_fim, valor_total, qtd_parcelas, observacoes
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?);
            """,
            (
                aluno_id,
                aluno["responsavel_id"],
                numero_contrato,
                data_inicio,
                data_fim,
                valor_total_float,
                qtd_parcelas_int,
                observacoes,
            ),
        )
        contrato_id = cur.lastrowid
        
        if contrato_id and qtd_parcelas_int > 0:
            valor_base = round(valor_total_float / qtd_parcelas_int, 2)
            acumulado = 0.0
            
            for i in range(1, qtd_parcelas_int + 1):
                if i < qtd_parcelas_int:
                    valor_parcela = valor_base
                    acumulado += valor_parcela
                else:
                    valor_parcela = round(valor_total_float - acumulado, 2)
                
                data_venc = add_months(data_inicio, i - 1)
                
                db.execute(
                    """
                    INSERT INTO parcelas (
                        contrato_id,
                        numero_parcela,
                        valor_original,
                        valor_final,
                        data_vencimento
                    )
                    VALUES (?, ?, ?, ?, ?);
                    """,
                    (
                        contrato_id,
                        i,
                        valor_parcela,
                        valor_parcela,
                        data_venc,
                    ),
                )
        
        db.commit()
        return redirect(url_for("detalhe_aluno", aluno_id=aluno_id))
    
    return render_template(
        "novo_contrato.html",
        aluno=aluno,
        hoje=date.today().isoformat(),
    )
# ═══════════════════════════════════════════════════════════════
# ROTAS DE PAGAMENTOS — adicionar no app3.py
# Cole este bloco antes do "if __name__ == '__main__':"
# ═══════════════════════════════════════════════════════════════

# ═══════════════════════════════════════════════════════════════
# ROTAS DE PAGAMENTOS — adicionar no app3.py
# Cole este bloco antes do "if __name__ == '__main__':"
# ═══════════════════════════════════════════════════════════════

@app.route("/pagamentos")
def pagamentos():
    db = get_db()
    hoje = date.today().isoformat()
    aba = request.args.get("aba", "mes")

    # ── atualiza status atrasado automaticamente ──────────────
    db.execute("""
        UPDATE parcelas
        SET status_pagamento = 'atrasado'
        WHERE (status_pagamento = 'pendente' OR status_pagamento IS NULL)
          AND data_vencimento < ?
          AND (cancelada = 0 OR cancelada IS NULL);
    """, (hoje,))
    db.commit()

    # parâmetros comuns
    por_pagina_raw = request.args.get("por_pagina", "50")
    try:
        por_pagina = int(por_pagina_raw)
    except ValueError:
        por_pagina = 50

    # ── ABA: parcelas por período ─────────────────────────────
    if aba == "mes":
        de_str       = (request.args.get("de")  or "").strip()
        ate_str      = (request.args.get("ate") or "").strip()
        status_filtro = request.args.get("status_filtro", "").strip()
        page          = request.args.get("page", 1, type=int)

        def normaliza_data(s):
            if not s:
                return None
            s = s.strip()
            if len(s) == 10 and s[4] == "-" and s[7] == "-":
                y, m, d = s.split("-")
                if y.isdigit() and m.isdigit() and d.isdigit():
                    return s
            if "/" in s:
                partes = s.split("/")
                if len(partes) == 3:
                    d, m, y = partes
                    if d.isdigit() and m.isdigit() and y.isdigit():
                        return f"{y}-{m.zfill(2)}-{d.zfill(2)}"
            return None

        de_iso  = normaliza_data(de_str)
        ate_iso = normaliza_data(ate_str)

        # se não houver filtro de data, mostra o mês atual
        if not de_iso and not ate_iso:
            import calendar as cal_mod
            hoje_obj = date.today()
            primeiro = hoje_obj.replace(day=1).isoformat()
            ultimo   = hoje_obj.replace(
                day=cal_mod.monthrange(hoje_obj.year, hoje_obj.month)[1]
            ).isoformat()
            de_iso   = primeiro
            ate_iso  = ultimo
            de_str   = primeiro
            ate_str  = ultimo

        sql_base = """
            SELECT
                p.id,
                p.numero_parcela,
                p.valor_final,
                p.data_vencimento,
                p.status_pagamento,
                p.pago_em,
                p.valor_pago,
                p.canal_pagamento,
                p.observacao_pagamento,
                p.registrado_por,
                a.nome         AS aluno,
                r.nome         AS responsavel,
                c.numero_contrato
            FROM parcelas p
            JOIN contratos    c ON p.contrato_id    = c.id
            JOIN alunos       a ON c.aluno_id       = a.id
            JOIN responsaveis r ON c.responsavel_id = r.id
            WHERE (p.cancelada = 0 OR p.cancelada IS NULL)
              AND (c.ativo = 1 OR c.ativo IS NULL)
              AND (a.ativo = 1 OR a.ativo IS NULL)
        """
        params = []

        if de_iso:
            sql_base += " AND p.data_vencimento >= ?"
            params.append(de_iso)
        if ate_iso:
            sql_base += " AND p.data_vencimento <= ?"
            params.append(ate_iso)
        if status_filtro:
            sql_base += " AND p.status_pagamento = ?"
            params.append(status_filtro)

        sql_base += " ORDER BY p.data_vencimento, a.nome"

        # totais para o resumo (sem paginação)
        todas = db.execute(sql_base, params).fetchall()
        total_registros = len(todas)
        total_pagas     = sum(1 for p in todas if p["status_pagamento"] == "pago")
        total_pendentes = sum(1 for p in todas if p["status_pagamento"] in ("pendente", None))
        total_atrasadas = sum(1 for p in todas if p["status_pagamento"] == "atrasado")
        valor_recebido  = sum(float(p["valor_pago"] or 0) for p in todas if p["status_pagamento"] == "pago")
        valor_aberto    = sum(float(p["valor_final"] or 0) for p in todas if p["status_pagamento"] != "pago")

        # paginação
        if por_pagina >= 99999:
            parcelas    = todas
            total_pages = 1
        else:
            total_pages = max(1, math.ceil(total_registros / por_pagina))
            offset      = (page - 1) * por_pagina
            parcelas    = db.execute(
                sql_base + " LIMIT ? OFFSET ?",
                params + [por_pagina, offset]
            ).fetchall()

        return render_template(
            "pagamentos.html",
            aba=aba,
            parcelas=parcelas,
            de=de_str, ate=ate_str,
            status_filtro=status_filtro,
            por_pagina=por_pagina,
            page=page,
            total_pages=total_pages,
            total_registros=total_registros,
            total_pagas=total_pagas,
            total_pendentes=total_pendentes,
            total_atrasadas=total_atrasadas,
            valor_recebido=valor_recebido,
            valor_aberto=valor_aberto,
            q="",
        )

    # ── ABA: busca por nome ───────────────────────────────────
    else:
        q = request.args.get("q", "").strip()
        resultados = []

        if q:
            termo = f"%{q}%"
            resultados = db.execute("""
                SELECT
                    p.id,
                    p.numero_parcela,
                    p.valor_final,
                    p.data_vencimento,
                    p.status_pagamento,
                    p.pago_em,
                    p.valor_pago,
                    p.canal_pagamento,
                    p.observacao_pagamento,
                    p.registrado_por,
                    a.nome         AS aluno,
                    r.nome         AS responsavel,
                    c.numero_contrato
                FROM parcelas p
                JOIN contratos    c ON p.contrato_id    = c.id
                JOIN alunos       a ON c.aluno_id       = a.id
                JOIN responsaveis r ON c.responsavel_id = r.id
                WHERE (p.cancelada = 0 OR p.cancelada IS NULL)
                  AND (c.ativo = 1 OR c.ativo IS NULL)
                  AND (a.ativo = 1 OR a.ativo IS NULL)
                  AND (a.nome LIKE ? OR r.nome LIKE ?)
                ORDER BY p.data_vencimento, a.nome
                LIMIT ? OFFSET 0
            """, (termo, termo, por_pagina if por_pagina < 99999 else 99999)).fetchall()

        return render_template(
            "pagamentos.html",
            aba=aba,
            q=q,
            resultados=resultados,
            por_pagina=por_pagina,
            de="", ate="",
            status_filtro="",
            page=1, total_pages=1,
            total_registros=0,
            total_pagas=0, total_pendentes=0, total_atrasadas=0,
            valor_recebido=0, valor_aberto=0,
        )


@app.post("/pagamentos/registrar")
def registrar_pagamento():
    db = get_db()
    parcela_id     = request.form.get("parcela_id")
    canal          = request.form.get("canal")
    valor_pago     = request.form.get("valor_pago")
    data_pagamento = request.form.get("data_pagamento")
    registrado_por = request.form.get("registrado_por", "").strip()
    observacao     = request.form.get("observacao", "").strip()

    # origem para redirecionar de volta
    origem_aba = request.form.get("origem_aba", "mes")
    origem_de  = request.form.get("origem_de", "")
    origem_ate = request.form.get("origem_ate", "")
    origem_q   = request.form.get("origem_q", "")
    por_pagina = request.form.get("por_pagina", "50")

    if not all([parcela_id, canal, valor_pago, data_pagamento]):
        return "Dados incompletos", 400

    try:
        # 1. Atualiza a parcela
        db.execute("""
            UPDATE parcelas
            SET status_pagamento   = 'pago',
                pago_em            = ?,
                valor_pago         = ?,
                canal_pagamento    = ?,
                observacao_pagamento = ?,
                registrado_por     = ?
            WHERE id = ?;
        """, (data_pagamento, float(valor_pago), canal,
              observacao or None, registrado_por or None, parcela_id))

        # 2. Registra no histórico de pagamentos
        db.execute("""
            INSERT INTO pagamentos
                (parcela_id, valor_pago, data_pagamento, canal, origem, registrado_por, observacao)
            VALUES (?, ?, ?, ?, 'manual', ?, ?);
        """, (parcela_id, float(valor_pago), data_pagamento,
              canal, registrado_por or None, observacao or None))

        db.commit()

    except Exception as e:
        db.rollback()
        return f"Erro ao registrar pagamento: {str(e)}", 500

    # redireciona de volta para a aba de origem
    return redirect(url_for(
        "pagamentos",
        aba=origem_aba,
        de=origem_de,
        ate=origem_ate,
        q=origem_q,
        por_pagina=por_pagina,
    ))

@app.post("/pagamentos/<int:parcela_id>/estornar")
def estornar_pagamento(parcela_id):
    db = get_db()
    hoje      = date.today().isoformat()
    estornado_por  = request.form.get("estornado_por", "").strip()
    motivo_estorno = request.form.get("motivo_estorno", "").strip()
    origem_aba = request.form.get("origem_aba", "mes")
    origem_de  = request.form.get("origem_de", "")
    origem_ate = request.form.get("origem_ate", "")
    origem_q   = request.form.get("origem_q", "")
    por_pagina = request.form.get("por_pagina", "50")
 
    try:
        # Volta parcela para pendente ou atrasado
        novo_status = "atrasado" if db.execute(
            "SELECT data_vencimento FROM parcelas WHERE id = ?;",
            (parcela_id,)
        ).fetchone()["data_vencimento"] < hoje else "pendente"
 
        db.execute("""
            UPDATE parcelas SET
                status_pagamento     = ?,
                pago_em              = NULL,
                valor_pago           = NULL,
                canal_pagamento      = NULL,
                observacao_pagamento = NULL,
                registrado_por       = NULL
            WHERE id = ?;
        """, (novo_status, parcela_id))
 
        # Marca o pagamento como estornado no histórico
        db.execute("""
            UPDATE pagamentos SET
                estornado      = 1,
                estornado_em   = ?,
                estornado_por  = ?,
                motivo_estorno = ?
            WHERE parcela_id = ?
              AND estornado   = 0;
        """, (hoje, estornado_por or None, motivo_estorno or None, parcela_id))
 
        db.commit()
 
    except Exception as e:
        db.rollback()
        return f"Erro ao estornar: {str(e)}", 500
 
    return redirect(url_for(
        "pagamentos",
        aba=origem_aba, de=origem_de, ate=origem_ate,
        q=origem_q, por_pagina=por_pagina,
    ))
 
 
# ── EDITAR PAGAMENTO ──────────────────────────────────────────
 
@app.post("/pagamentos/<int:parcela_id>/editar")
def editar_pagamento(parcela_id):
    db = get_db()
    canal          = request.form.get("canal", "").strip()
    data_pagamento = request.form.get("data_pagamento", "").strip()
    valor_pago     = request.form.get("valor_pago", "").strip()
    registrado_por = request.form.get("registrado_por", "").strip()
    observacao     = request.form.get("observacao", "").strip()
    origem_aba = request.form.get("origem_aba", "mes")
    origem_de  = request.form.get("origem_de", "")
    origem_ate = request.form.get("origem_ate", "")
    origem_q   = request.form.get("origem_q", "")
    por_pagina = request.form.get("por_pagina", "50")
 
    if not canal or not data_pagamento or not valor_pago:
        return "Dados incompletos", 400
 
    try:
        db.execute("""
            UPDATE parcelas SET
                canal_pagamento      = ?,
                pago_em              = ?,
                valor_pago           = ?,
                registrado_por       = ?,
                observacao_pagamento = ?
            WHERE id = ? AND status_pagamento = 'pago';
        """, (
            canal, data_pagamento, float(valor_pago),
            registrado_por or None, observacao or None,
            parcela_id,
        ))
 
        # Atualiza também o registro mais recente no histórico
        db.execute("""
            UPDATE pagamentos SET
                canal          = ?,
                data_pagamento = ?,
                valor_pago     = ?,
                registrado_por = ?,
                observacao     = ?
            WHERE parcela_id = ?
              AND estornado   = 0
              AND id = (
                  SELECT id FROM pagamentos
                  WHERE parcela_id = ? AND estornado = 0
                  ORDER BY criado_em DESC LIMIT 1
              );
        """, (
            canal, data_pagamento, float(valor_pago),
            registrado_por or None, observacao or None,
            parcela_id, parcela_id,
        ))
 
        db.commit()
 
    except Exception as e:
        db.rollback()
        return f"Erro ao editar: {str(e)}", 500
 
    return redirect(url_for(
        "pagamentos",
        aba=origem_aba, de=origem_de, ate=origem_ate,
        q=origem_q, por_pagina=por_pagina,
    ))
@app.post("/pagamentos/alterar_status")
def alterar_status_parcela():
    db = get_db()
    parcela_id  = request.form.get("parcela_id")
    novo_status = request.form.get("novo_status")

    origem_aba = request.form.get("origem_aba", "mes")
    origem_de  = request.form.get("origem_de", "")
    origem_ate = request.form.get("origem_ate", "")
    origem_q   = request.form.get("origem_q", "")
    por_pagina = request.form.get("por_pagina", "50")

    if novo_status not in ("pendente", "atrasado"):
        return "Status inválido", 400

    try:
        db.execute("""
            UPDATE parcelas
            SET status_pagamento = ?
            WHERE id = ? AND status_pagamento != 'pago';
        """, (novo_status, parcela_id))
        db.commit()
    except Exception as e:
        db.rollback()
        return f"Erro: {str(e)}", 500

    return redirect(url_for(
        "pagamentos",
        aba=origem_aba,
        de=origem_de,
        ate=origem_ate,
        q=origem_q,
        por_pagina=por_pagina,
    ))


# ── REGISTRO EM LOTE ─────────────────────────────────────────

@app.post("/pagamentos/registrar-lote")
def registrar_pagamento_lote():
    db = get_db()
    parcelas_ids   = request.form.getlist("parcelas_ids")
    canal          = request.form.get("canal")
    data_pagamento = request.form.get("data_pagamento")
    registrado_por = request.form.get("registrado_por", "").strip()
    observacao     = request.form.get("observacao", "").strip()
    origem_q       = request.form.get("origem_q", "")
    origem_aba     = request.form.get("origem_aba", "busca")

    if not parcelas_ids or not canal or not data_pagamento:
        return "Dados incompletos", 400

    try:
        for parcela_id in parcelas_ids:
            # Busca o valor original da parcela
            parcela = db.execute(
                "SELECT id, valor_final FROM parcelas WHERE id = ? AND (cancelada = 0 OR cancelada IS NULL);",
                (parcela_id,)
            ).fetchone()
            if not parcela:
                continue

            # Atualiza a parcela
            db.execute("""
                UPDATE parcelas SET
                    status_pagamento     = 'pago',
                    pago_em              = ?,
                    valor_pago           = ?,
                    canal_pagamento      = ?,
                    observacao_pagamento = ?,
                    registrado_por       = ?
                WHERE id = ?;
            """, (data_pagamento, parcela["valor_final"], canal,
                  observacao or None, registrado_por or None, parcela_id))

            # Registra no histórico
            db.execute("""
                INSERT INTO pagamentos
                    (parcela_id, valor_pago, data_pagamento, canal, origem, registrado_por, observacao)
                VALUES (?, ?, ?, ?, 'manual', ?, ?);
            """, (parcela_id, parcela["valor_final"], data_pagamento,
                  canal, registrado_por or None, observacao or None))

        db.commit()

    except Exception as e:
        db.rollback()
        return f"Erro ao registrar pagamentos: {str(e)}", 500

    return redirect(url_for("pagamentos", aba=origem_aba, q=origem_q))

# ── PROFESSORES ──────────────────────────────────────────────

@app.route("/professores")
def listar_professores():
    db = get_db()
    professores = db.execute("""
        SELECT p.id, p.nome, p.contato, p.ativo,
               COUNT(t.id) AS total_turmas
        FROM professores p
        LEFT JOIN turmas t ON t.professor_id = p.id AND (t.ativa = 1 OR t.ativa IS NULL)
        GROUP BY p.id
        ORDER BY p.nome;
    """).fetchall()
    return render_template("professores.html", professores=professores)


@app.post("/professores/novo")
def novo_professor():
    db = get_db()
    nome    = request.form.get("nome", "").strip()
    contato = request.form.get("contato", "").strip()
    if not nome:
        return redirect(url_for("listar_professores"))
    db.execute(
        "INSERT INTO professores (nome, contato, ativo) VALUES (?, ?, 1);",
        (nome, contato or None)
    )
    db.commit()
    return redirect(url_for("listar_professores"))


@app.post("/professores/<int:prof_id>/editar")
def editar_professor(prof_id):
    db = get_db()
    nome    = request.form.get("nome", "").strip()
    contato = request.form.get("contato", "").strip()
    if not nome:
        return redirect(url_for("listar_professores"))
    db.execute(
        "UPDATE professores SET nome = ?, contato = ? WHERE id = ?;",
        (nome, contato or None, prof_id)
    )
    db.commit()
    return redirect(url_for("listar_professores"))

# ═══════════════════════════════════════════════════════════════
# PATCH — Módulo de relatórios
# Cole este bloco antes do "if __name__ == '__main__':"
# Também adicione no menu base.html e no inject_modulo_atual
# ═══════════════════════════════════════════════════════════════

import io
import calendar as cal_mod
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter


def _buscar_dados_relatorio(db, de_iso, ate_iso):
    """Busca todos os dados necessários para o relatório."""

    # ── Pagamentos recebidos no período ──────────────────────
    pagamentos = db.execute("""
        SELECT
            a.nome        AS aluno,
            r.nome        AS responsavel,
            c.numero_contrato,
            p.numero_parcela,
            p.valor_final,
            p.valor_pago,
            p.pago_em,
            p.canal_pagamento,
            p.registrado_por
        FROM parcelas p
        JOIN contratos    c ON p.contrato_id    = c.id
        JOIN alunos       a ON c.aluno_id       = a.id
        JOIN responsaveis r ON c.responsavel_id = r.id
        WHERE p.status_pagamento = 'pago'
          AND p.pago_em >= ? AND p.pago_em <= ?
          AND (p.cancelada = 0 OR p.cancelada IS NULL)
        ORDER BY p.pago_em, a.nome;
    """, (de_iso, ate_iso)).fetchall()

    # ── Inadimplência ────────────────────────────────────────
    inadimplentes = db.execute("""
        SELECT
            a.nome        AS aluno,
            r.nome        AS responsavel,
            r.telefone,
            c.numero_contrato,
            p.numero_parcela,
            p.valor_final,
            p.data_vencimento,
            julianday('now') - julianday(p.data_vencimento) AS dias_atraso
        FROM parcelas p
        JOIN contratos    c ON p.contrato_id    = c.id
        JOIN alunos       a ON c.aluno_id       = a.id
        JOIN responsaveis r ON c.responsavel_id = r.id
        WHERE p.status_pagamento = 'atrasado'
          AND (p.cancelada = 0 OR p.cancelada IS NULL)
          AND (c.ativo = 1 OR c.ativo IS NULL)
        ORDER BY dias_atraso DESC, a.nome;
    """).fetchall()

    # ── Parcelas em aberto no período (não pagas, não atrasadas) ─
    em_aberto = db.execute("""
        SELECT
            a.nome        AS aluno,
            r.nome        AS responsavel,
            c.numero_contrato,
            p.numero_parcela,
            p.valor_final,
            p.data_vencimento,
            p.status_pagamento
        FROM parcelas p
        JOIN contratos    c ON p.contrato_id    = c.id
        JOIN alunos       a ON c.aluno_id       = a.id
        JOIN responsaveis r ON c.responsavel_id = r.id
        WHERE p.status_pagamento IN ('pendente')
          AND p.data_vencimento >= ? AND p.data_vencimento <= ?
          AND (p.cancelada = 0 OR p.cancelada IS NULL)
          AND (c.ativo = 1 OR c.ativo IS NULL)
        ORDER BY p.data_vencimento, a.nome;
    """, (de_iso, ate_iso)).fetchall()

    # ── Notas fiscais emitidas no período ────────────────────
    notas = db.execute("""
        SELECT
            a.nome        AS aluno,
            r.nome        AS responsavel,
            nf.numero_nf AS numero_nota,
            nf.valor,
            nf.data_emissao,
            c.numero_contrato,
            p.numero_parcela
        FROM notas_fiscais nf
        JOIN parcelas     p  ON nf.parcela_id    = p.id
        JOIN contratos    c  ON p.contrato_id    = c.id
        JOIN alunos       a  ON c.aluno_id       = a.id
        JOIN responsaveis r  ON c.responsavel_id = r.id
        WHERE nf.data_emissao >= ? AND nf.data_emissao <= ?
        ORDER BY nf.data_emissao, a.nome;
    """, (de_iso, ate_iso)).fetchall()

    # ── Rematrículas (independente do período) ───────────────
    rematriculas = db.execute("""
        SELECT
            a.nome        AS aluno,
            t.tipo        AS turma_tipo,
            t.numero      AS turma_numero,
            t.horario,
            p.nome        AS professor,
            t.periodo,
            rm.status,
            rm.renovacao_automatica,
            rm.motivo_nao_renovou
        FROM rematriculas rm
        JOIN alunos       a  ON rm.aluno_id     = a.id
        LEFT JOIN turmas  t  ON rm.turma_id     = t.id
        LEFT JOIN professores p ON t.professor_id = p.id
        WHERE (a.ativo = 1 OR a.ativo IS NULL)
        ORDER BY t.tipo, t.numero, a.nome;
    """).fetchall()

    # ── Resumo ───────────────────────────────────────────────
    total_recebido  = sum(float(p["valor_pago"] or p["valor_final"]) for p in pagamentos)
    total_notas     = sum(float(n["valor"] or 0) for n in notas)
    total_aberto    = sum(float(p["valor_final"]) for p in em_aberto)
    total_atrasado  = sum(float(p["valor_final"]) for p in inadimplentes)

    resumo = {
        "total_recebido":   total_recebido,
        "qtd_pagamentos":   len(pagamentos),
        "total_notas":      total_notas,
        "qtd_notas":        len(notas),
        "total_aberto":     total_aberto,
        "qtd_aberto":       len(em_aberto),
        "total_atrasado":   total_atrasado,
        "qtd_inadimplentes": len(set(p["aluno"] for p in inadimplentes)),
        "qtd_parcelas_atr": len(inadimplentes),
    }

    return {
        "resumo":        resumo,
        "pagamentos":    [dict(p) for p in pagamentos],
        "inadimplentes": [dict(p) for p in inadimplentes],
        "em_aberto":     [dict(p) for p in em_aberto],
        "notas":         [dict(n) for n in notas],
        "rematriculas":  [dict(r) for r in rematriculas],
    }


def _fmt_brl(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _fmt_data(s):
    if not s:
        return "—"
    try:
        from datetime import date
        d = date.fromisoformat(str(s)[:10])
        return d.strftime("%d/%m/%Y")
    except Exception:
        return str(s)[:10]


def _canal_label(c):
    m = {"rede": "Maquininha Rede", "link_rede": "Link Rede", "asaas": "Asaas",
         "c6": "C6", "pix_manual": "PIX Manual", "dinheiro": "Dinheiro",
         "boleto": "Boleto", "outro": "Outro"}
    return m.get(c, c or "—")


def _status_label(s):
    m = {"renovou": "Renovou", "automatico": "Automático",
         "em_analise": "Em análise", "nao_vai_renovar": "Não vai renovar",
         "pendente": "Pendente"}
    return m.get(s, s or "—")


# ── GERADOR DE PDF ────────────────────────────────────────────

def gerar_pdf_relatorio(dados, de_str, ate_str, titulo):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=1.8*cm, rightMargin=1.8*cm,
        topMargin=2*cm, bottomMargin=2*cm,
        title=titulo,
    )

    # Paleta
    AZUL_ESC  = colors.HexColor("#1a1d23")
    AZUL_MED  = colors.HexColor("#3b5bdb")
    CINZA_CLR = colors.HexColor("#f4f5f7")
    CINZA_BRD = colors.HexColor("#e2e4e9")
    VERDE     = colors.HexColor("#1a7f4e")
    VERM      = colors.HexColor("#c0392b")
    AMARELO   = colors.HexColor("#c07a00")

    styles = getSampleStyleSheet()
    s_titulo = ParagraphStyle("titulo", fontName="Helvetica-Bold",
                              fontSize=18, textColor=AZUL_ESC, spaceAfter=4)
    s_sub    = ParagraphStyle("sub",    fontName="Helvetica",
                              fontSize=10, textColor=colors.HexColor("#6b7280"), spaceAfter=16)
    s_h2     = ParagraphStyle("h2",     fontName="Helvetica-Bold",
                              fontSize=12, textColor=AZUL_ESC, spaceBefore=16, spaceAfter=8)
    s_label  = ParagraphStyle("label",  fontName="Helvetica",
                              fontSize=8,  textColor=colors.HexColor("#9099a8"))
    s_valor  = ParagraphStyle("valor",  fontName="Helvetica-Bold",
                              fontSize=14, textColor=AZUL_ESC)
    s_obs    = ParagraphStyle("obs",    fontName="Helvetica",
                              fontSize=8,  textColor=colors.HexColor("#9099a8"),
                              spaceBefore=4)
    s_normal = ParagraphStyle("norm",   fontName="Helvetica", fontSize=9)

    def th_style(col_widths, header, rows, zebra=True):
        """Cria uma tabela estilizada."""
        data = [header] + rows
        t = Table(data, colWidths=col_widths, repeatRows=1)
        ts = [
            # Cabeçalho
            ("BACKGROUND",  (0,0), (-1,0), AZUL_ESC),
            ("TEXTCOLOR",   (0,0), (-1,0), colors.white),
            ("FONTNAME",    (0,0), (-1,0), "Helvetica-Bold"),
            ("FONTSIZE",    (0,0), (-1,0), 8),
            ("TOPPADDING",  (0,0), (-1,0), 7),
            ("BOTTOMPADDING",(0,0), (-1,0), 7),
            ("LEFTPADDING", (0,0), (-1,0), 6),
            # Linhas
            ("FONTNAME",    (0,1), (-1,-1), "Helvetica"),
            ("FONTSIZE",    (0,1), (-1,-1), 8),
            ("TOPPADDING",  (0,1), (-1,-1), 5),
            ("BOTTOMPADDING",(0,1), (-1,-1), 5),
            ("LEFTPADDING", (0,1), (-1,-1), 6),
            ("LINEBELOW",   (0,0), (-1,-1), 0.3, CINZA_BRD),
            ("VALIGN",      (0,0), (-1,-1), "MIDDLE"),
        ]
        if zebra:
            for i in range(1, len(data)):
                if i % 2 == 0:
                    ts.append(("BACKGROUND", (0,i), (-1,i), CINZA_CLR))
        t.setStyle(TableStyle(ts))
        return t

    story = []
    r = dados["resumo"]
    W = A4[0] - 3.6*cm  # largura útil

    # ── CAPA ─────────────────────────────────────────────────
    story.append(Spacer(1, 1*cm))
    story.append(Paragraph("CTRLPlay", s_titulo))
    story.append(Paragraph(f"Relatório financeiro · {de_str} a {ate_str}", s_sub))
    story.append(HRFlowable(width="100%", thickness=0.5, color=CINZA_BRD))
    story.append(Spacer(1, 0.5*cm))

    # Cards de resumo
    card_data = [
        [
            Paragraph("TOTAL RECEBIDO", s_label),
            Paragraph("NOTAS EMITIDAS", s_label),
            Paragraph("EM ABERTO", s_label),
            Paragraph("INADIMPLÊNCIA", s_label),
        ],
        [
            Paragraph(_fmt_brl(r["total_recebido"]),
                      ParagraphStyle("v", fontName="Helvetica-Bold", fontSize=13, textColor=VERDE)),
            Paragraph(_fmt_brl(r["total_notas"]),
                      ParagraphStyle("v", fontName="Helvetica-Bold", fontSize=13, textColor=AZUL_MED)),
            Paragraph(_fmt_brl(r["total_aberto"]),
                      ParagraphStyle("v", fontName="Helvetica-Bold", fontSize=13, textColor=AMARELO)),
            Paragraph(_fmt_brl(r["total_atrasado"]),
                      ParagraphStyle("v", fontName="Helvetica-Bold", fontSize=13, textColor=VERM)),
        ],
        [
            Paragraph(f"{r['qtd_pagamentos']} parcelas", s_obs),
            Paragraph(f"{r['qtd_notas']} notas", s_obs),
            Paragraph(f"{r['qtd_aberto']} parcelas", s_obs),
            Paragraph(f"{r['qtd_inadimplentes']} alunos · {r['qtd_parcelas_atr']} parcelas", s_obs),
        ],
    ]
    cw = W / 4
    ct = Table(card_data, colWidths=[cw]*4)
    ct.setStyle(TableStyle([
        ("BOX",         (0,0), (0,-1), 0.5, CINZA_BRD),
        ("BOX",         (1,0), (1,-1), 0.5, CINZA_BRD),
        ("BOX",         (2,0), (2,-1), 0.5, CINZA_BRD),
        ("BOX",         (3,0), (3,-1), 0.5, CINZA_BRD),
        ("BACKGROUND",  (0,0), (-1,-1), colors.white),
        ("TOPPADDING",  (0,0), (-1,-1), 10),
        ("BOTTOMPADDING",(0,0),(-1,-1), 8),
        ("LEFTPADDING", (0,0), (-1,-1), 10),
        ("VALIGN",      (0,0), (-1,-1), "TOP"),
    ]))
    story.append(ct)
    story.append(Spacer(1, 0.4*cm))

    # ── 1. PAGAMENTOS RECEBIDOS ───────────────────────────────
    story.append(PageBreak())
    story.append(Paragraph("1. Pagamentos recebidos no período", s_h2))
    if dados["pagamentos"]:
        cols = [3.5*cm, 3.5*cm, 2*cm, 1.5*cm, 2*cm, 2*cm, 2.5*cm]
        header = ["Aluno", "Responsável", "Contrato", "Parc.", "Valor", "Pago em", "Canal"]
        rows = []
        for p in dados["pagamentos"]:
            rows.append([
                p["aluno"], p["responsavel"], p["numero_contrato"],
                str(p["numero_parcela"]),
                _fmt_brl(float(p["valor_pago"] or p["valor_final"])),
                _fmt_data(p["pago_em"]), _canal_label(p["canal_pagamento"]),
            ])
        story.append(th_style(cols, header, rows))
        total = sum(float(p["valor_pago"] or p["valor_final"]) for p in dados["pagamentos"])
        story.append(Paragraph(
            f"<b>Total recebido: {_fmt_brl(total)}</b>",
            ParagraphStyle("tot", fontName="Helvetica-Bold", fontSize=9,
                           textColor=VERDE, spaceBefore=6)
        ))
    else:
        story.append(Paragraph("Nenhum pagamento registrado no período.", s_normal))

    # ── 2. INADIMPLÊNCIA ─────────────────────────────────────
    story.append(PageBreak())
    story.append(Paragraph("2. Inadimplência", s_h2))
    if dados["inadimplentes"]:
        cols = [3.5*cm, 3*cm, 2.5*cm, 2*cm, 1.5*cm, 2*cm, 2*cm]
        header = ["Aluno", "Responsável", "Telefone", "Contrato", "Parc.", "Valor", "Dias atraso"]
        rows = []
        for p in dados["inadimplentes"]:
            rows.append([
                p["aluno"], p["responsavel"], p["telefone"] or "—",
                p["numero_contrato"], str(p["numero_parcela"]),
                _fmt_brl(float(p["valor_final"])),
                str(int(p["dias_atraso"] or 0)) + " dias",
            ])
        story.append(th_style(cols, header, rows))
        total = sum(float(p["valor_final"]) for p in dados["inadimplentes"])
        story.append(Paragraph(
            f"<b>Total em atraso: {_fmt_brl(total)}</b>",
            ParagraphStyle("tot", fontName="Helvetica-Bold", fontSize=9,
                           textColor=VERM, spaceBefore=6)
        ))
    else:
        story.append(Paragraph("Nenhuma parcela em atraso.", s_normal))

    # ── 3. EM ABERTO ─────────────────────────────────────────
    story.append(PageBreak())
    story.append(Paragraph("3. Parcelas em aberto no período", s_h2))
    if dados["em_aberto"]:
        cols = [3.5*cm, 3.5*cm, 2*cm, 1.5*cm, 2.5*cm, 2.5*cm]
        header = ["Aluno", "Responsável", "Contrato", "Parc.", "Valor", "Vencimento"]
        rows = []
        for p in dados["em_aberto"]:
            rows.append([
                p["aluno"], p["responsavel"], p["numero_contrato"],
                str(p["numero_parcela"]),
                _fmt_brl(float(p["valor_final"])),
                _fmt_data(p["data_vencimento"]),
            ])
        story.append(th_style(cols, header, rows))
    else:
        story.append(Paragraph("Nenhuma parcela em aberto para o período.", s_normal))

    # ── 4. NOTAS FISCAIS ─────────────────────────────────────
    story.append(PageBreak())
    story.append(Paragraph("4. Notas fiscais emitidas", s_h2))
    if dados["notas"]:
        cols = [3.5*cm, 3*cm, 2.5*cm, 2*cm, 2*cm, 2.5*cm]
        header = ["Aluno", "Responsável", "Nº Nota", "Valor", "Emitida em", "Contrato"]
        rows = []
        for n in dados["notas"]:
            rows.append([
                n["aluno"], n["responsavel"],
                str(n["numero_nota"] or "—"),
                _fmt_brl(float(n["valor"] or 0)),
                _fmt_data(n["data_emissao"]),
                n["numero_contrato"],
            ])
        story.append(th_style(cols, header, rows))
        total = sum(float(n["valor"] or 0) for n in dados["notas"])
        story.append(Paragraph(
            f"<b>Total emitido: {_fmt_brl(total)}</b>",
            ParagraphStyle("tot", fontName="Helvetica-Bold", fontSize=9,
                           textColor=AZUL_MED, spaceBefore=6)
        ))
    else:
        story.append(Paragraph("Nenhuma nota fiscal emitida no período.", s_normal))

    # ── 5. REMATRÍCULAS ──────────────────────────────────────
    story.append(PageBreak())
    story.append(Paragraph("5. Situação de rematrículas", s_h2))
    if dados["rematriculas"]:
        cols = [3.5*cm, 2.5*cm, 3*cm, 2*cm, 2.5*cm, 2*cm]
        header = ["Aluno", "Turma", "Professor", "Período", "Status", "Auto"]
        rows = []
        for r_ in dados["rematriculas"]:
            turma = f"{r_['turma_tipo']} {r_['turma_numero']}" if r_["turma_tipo"] else "—"
            rows.append([
                r_["aluno"], turma,
                r_["professor"] or "—",
                r_["periodo"] or "—",
                _status_label(r_["status"]),
                "Sim" if r_["renovacao_automatica"] else "Não",
            ])
        story.append(th_style(cols, header, rows))
    else:
        story.append(Paragraph("Nenhuma rematrícula registrada.", s_normal))

    doc.build(story)
    buf.seek(0)
    return buf


# ── GERADOR DE EXCEL ──────────────────────────────────────────

def gerar_excel_relatorio(dados, de_str, ate_str):
    wb = openpyxl.Workbook()

    AZUL    = "1A1D23"
    AZUL_MD = "3B5BDB"
    CINZA   = "F4F5F7"
    VERDE   = "1A7F4E"
    VERM    = "C0392B"
    AMARELO = "C07A00"
    BRANCO  = "FFFFFF"

    def hdr_fill(hex_): return PatternFill("solid", fgColor=hex_)
    def hdr_font(bold=True, color=BRANCO, size=9):
        return Font(name="Arial", bold=bold, color=color, size=size)
    def cell_font(bold=False, color="1A1D23", size=9):
        return Font(name="Arial", bold=bold, color=color, size=size)
    def thin_border():
        s = Side(style="thin", color="E2E4E9")
        return Border(left=s, right=s, top=s, bottom=s)

    def escrever_secao(ws, titulo, headers, rows, cor_total=None, total_col=None):
        """Escreve cabeçalho + dados em uma aba."""
        # Título
        ws.append([titulo])
        t_cell = ws.cell(ws.max_row, 1)
        t_cell.font = Font(name="Arial", bold=True, size=12, color=AZUL)
        ws.append([])

        # Cabeçalhos
        ws.append(headers)
        hr = ws.max_row
        for c in range(1, len(headers)+1):
            cell = ws.cell(hr, c)
            cell.fill   = hdr_fill(AZUL)
            cell.font   = hdr_font()
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border()

        # Dados
        for i, row in enumerate(rows):
            ws.append(row)
            rr = ws.max_row
            bg = CINZA if i % 2 == 0 else BRANCO
            for c in range(1, len(row)+1):
                cell = ws.cell(rr, c)
                cell.font   = cell_font()
                cell.fill   = hdr_fill(bg)
                cell.border = thin_border()
                cell.alignment = Alignment(vertical="center")

        # Total
        if cor_total and total_col and rows:
            total = sum(r[total_col] for r in rows if isinstance(r[total_col], (int, float)))
            ws.append([])
            ws.append([""] * total_col + [f"Total: R$ {total:,.2f}".replace(",","X").replace(".",",").replace("X",".")])
            tc = ws.cell(ws.max_row, total_col+1)
            tc.font = Font(name="Arial", bold=True, size=9, color=cor_total)

        # Ajusta largura das colunas
        for col in ws.columns:
            max_w = 0
            for cell in col:
                try:
                    max_w = max(max_w, len(str(cell.value or "")))
                except Exception:
                    pass
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_w + 4, 40)

    r = dados["resumo"]

    # ── ABA: Resumo ──────────────────────────────────────────
    ws_res = wb.active
    ws_res.title = "Resumo"

    ws_res.append(["CTRLPlay — Relatório Financeiro"])
    ws_res.cell(1,1).font = Font(name="Arial", bold=True, size=16, color=AZUL)
    ws_res.append([f"Período: {de_str} a {ate_str}"])
    ws_res.cell(2,1).font = Font(name="Arial", size=10, color="6B7280")
    ws_res.append([])

    metricas = [
        ("Total recebido",    _fmt_brl(r["total_recebido"]),   VERDE),
        ("Notas emitidas",    _fmt_brl(r["total_notas"]),      AZUL_MD),
        ("Em aberto",         _fmt_brl(r["total_aberto"]),     AMARELO),
        ("Total inadimplente",_fmt_brl(r["total_atrasado"]),   VERM),
        ("Qtd. pagamentos",   str(r["qtd_pagamentos"]),        AZUL),
        ("Qtd. notas",        str(r["qtd_notas"]),             AZUL),
        ("Alunos inadimplentes", str(r["qtd_inadimplentes"]), VERM),
        ("Parcelas em atraso",str(r["qtd_parcelas_atr"]),      VERM),
    ]
    for label, valor, cor in metricas:
        ws_res.append([label, valor])
        rr = ws_res.max_row
        ws_res.cell(rr,1).font = Font(name="Arial", size=10, color="6B7280")
        ws_res.cell(rr,2).font = Font(name="Arial", bold=True, size=11, color=cor)

    ws_res.column_dimensions["A"].width = 28
    ws_res.column_dimensions["B"].width = 22

    # ── ABA: Pagamentos ──────────────────────────────────────
    ws_pag = wb.create_sheet("Pagamentos recebidos")
    headers = ["Aluno","Responsável","Contrato","Parcela","Valor pago","Pago em","Canal","Registrado por"]
    rows = [[
        p["aluno"], p["responsavel"], p["numero_contrato"],
        p["numero_parcela"],
        float(p["valor_pago"] or p["valor_final"]),
        _fmt_data(p["pago_em"]),
        _canal_label(p["canal_pagamento"]),
        p["registrado_por"] or "—",
    ] for p in dados["pagamentos"]]
    escrever_secao(ws_pag, "Pagamentos recebidos no período",
                   headers, rows, cor_total=VERDE, total_col=4)

    # ── ABA: Inadimplência ───────────────────────────────────
    ws_inad = wb.create_sheet("Inadimplência")
    headers = ["Aluno","Responsável","Telefone","Contrato","Parcela","Valor","Vencimento","Dias em atraso"]
    rows = [[
        p["aluno"], p["responsavel"], p["telefone"] or "—",
        p["numero_contrato"], p["numero_parcela"],
        float(p["valor_final"]),
        _fmt_data(p["data_vencimento"]),
        int(p["dias_atraso"] or 0),
    ] for p in dados["inadimplentes"]]
    escrever_secao(ws_inad, "Inadimplência atual",
                   headers, rows, cor_total=VERM, total_col=5)

    # ── ABA: Em aberto ───────────────────────────────────────
    ws_ab = wb.create_sheet("Em aberto")
    headers = ["Aluno","Responsável","Contrato","Parcela","Valor","Vencimento"]
    rows = [[
        p["aluno"], p["responsavel"], p["numero_contrato"],
        p["numero_parcela"], float(p["valor_final"]),
        _fmt_data(p["data_vencimento"]),
    ] for p in dados["em_aberto"]]
    escrever_secao(ws_ab, "Parcelas em aberto no período",
                   headers, rows, cor_total=AMARELO, total_col=4)

    # ── ABA: Notas fiscais ───────────────────────────────────
    ws_nf = wb.create_sheet("Notas fiscais")
    headers = ["Aluno","Responsável","Nº Nota","Valor","Emitida em","Contrato","Parcela"]
    rows = [[
        n["aluno"], n["responsavel"],
        n["numero_nota"] or "—",
        float(n["valor"] or 0),
        _fmt_data(n["data_emissao"]),
        n["numero_contrato"],
        n["numero_parcela"],
    ] for n in dados["notas"]]
    escrever_secao(ws_nf, "Notas fiscais emitidas no período",
                   headers, rows, cor_total=AZUL_MD, total_col=3)

    # ── ABA: Rematrículas ────────────────────────────────────
    ws_rem = wb.create_sheet("Rematrículas")
    headers = ["Aluno","Turma","Professor","Período","Status","Renovação automática","Motivo"]
    rows = [[
        r_["aluno"],
        f"{r_['turma_tipo']} {r_['turma_numero']}" if r_["turma_tipo"] else "—",
        r_["professor"] or "—",
        r_["periodo"] or "—",
        _status_label(r_["status"]),
        "Sim" if r_["renovacao_automatica"] else "Não",
        r_["motivo_nao_renovou"] or "—",
    ] for r_ in dados["rematriculas"]]
    escrever_secao(ws_rem, "Situação de rematrículas", headers, rows)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ── ROTAS ─────────────────────────────────────────────────────

@app.route("/relatorios")
def relatorios():
    """Página de seleção de período para geração de relatório."""
    from datetime import date
    hoje     = date.today()
    primeiro = hoje.replace(day=1).isoformat()
    ultimo   = hoje.replace(
        day=cal_mod.monthrange(hoje.year, hoje.month)[1]
    ).isoformat()
    return render_template("relatorios.html",
                           de_default=primeiro, ate_default=ultimo)


@app.route("/relatorios/gerar")
def gerar_relatorio():
    """Gera e devolve o relatório no formato escolhido."""
    from datetime import date
    from flask import send_file

    de_str  = request.args.get("de",      "").strip()
    ate_str = request.args.get("ate",     "").strip()
    fmt     = request.args.get("formato", "pdf").strip()

    if not de_str or not ate_str:
        return redirect(url_for("relatorios"))

    db    = get_db()
    dados = _buscar_dados_relatorio(db, de_str, ate_str)

    # Formata datas para nome do arquivo e título
    def fmt_d(s):
        try:
            from datetime import date as _d
            return _d.fromisoformat(s).strftime("%d/%m/%Y")
        except Exception:
            return s

    de_fmt  = fmt_d(de_str)
    ate_fmt = fmt_d(ate_str)
    titulo  = f"Relatório CTRLPlay · {de_fmt} a {ate_fmt}"
    slug    = f"ctrlplay_{de_str}_{ate_str}"

    if fmt == "excel":
        buf = gerar_excel_relatorio(dados, de_fmt, ate_fmt)
        return send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=f"{slug}.xlsx",
        )
    else:
        buf = gerar_pdf_relatorio(dados, de_fmt, ate_fmt, titulo)
        return send_file(
            buf,
            mimetype="application/pdf",
            as_attachment=True,
            download_name=f"{slug}.pdf",
        )
# ═══════════════════════════════════════════════════════════════
# PATCH — Cancelamento de contratos
# Cole antes do "if __name__ == '__main__':"
# ═══════════════════════════════════════════════════════════════

@app.route("/contrato/<int:contrato_id>/cancelar", methods=["GET", "POST"])
def cancelar_contrato(contrato_id):
    db = get_db()

    contrato = db.execute("""
        SELECT c.*, a.nome AS aluno, a.id AS aluno_id,
               r.nome AS responsavel
        FROM contratos c
        JOIN alunos       a ON c.aluno_id       = a.id
        JOIN responsaveis r ON c.responsavel_id = r.id
        WHERE c.id = ?;
    """, (contrato_id,)).fetchone()

    if not contrato:
        return "Contrato não encontrado", 404

    if contrato["cancelado"]:
        return redirect(url_for("detalhe_contrato", contrato_id=contrato_id))

    # Parcelas pagas
    pagas = db.execute("""
        SELECT id, numero_parcela, valor_final, valor_pago,
               data_vencimento, status_pagamento, pago_em
        FROM parcelas
        WHERE contrato_id = ?
          AND status_pagamento = 'pago'
          AND (cancelada = 0 OR cancelada IS NULL)
        ORDER BY numero_parcela;
    """, (contrato_id,)).fetchall()

    # Parcelas não pagas (pendentes + atrasadas)
    nao_pagas = db.execute("""
        SELECT id, numero_parcela, valor_final,
               data_vencimento, status_pagamento
        FROM parcelas
        WHERE contrato_id = ?
          AND (status_pagamento != 'pago' OR status_pagamento IS NULL)
          AND (cancelada = 0 OR cancelada IS NULL)
        ORDER BY numero_parcela;
    """, (contrato_id,)).fetchall()

    valor_canc = sum(float(p["valor_final"]) for p in nao_pagas)
    total_parc = len(pagas) + len(nao_pagas)

    if request.method == "POST":
        # Confirma que não é o GET de preview
        confirmado = request.form.get("confirmado") == "1"
        if not confirmado:
            return redirect(url_for("cancelar_contrato", contrato_id=contrato_id))

        data_canc       = request.form.get("data_cancelamento", date.today().isoformat())
        motivo          = request.form.get("motivo", "").strip()
        autorizado_por  = request.form.get("autorizado_por", "").strip()
        aulas_dadas     = request.form.get("aulas_dadas", "0").strip()
        total_aulas     = request.form.get("total_aulas", "0").strip()
        houve_res       = 1 if request.form.get("houve_ressarcimento") == "1" else 0
        valor_dev       = request.form.get("valor_devolvido", "0").strip()
        observacoes     = request.form.get("observacoes", "").strip()

        aulas_dadas_int = int(aulas_dadas) if aulas_dadas.isdigit() else 0
        total_aulas_int = int(total_aulas) if total_aulas.isdigit() else 0
        passou_75 = 0
        if total_aulas_int > 0:
            passou_75 = 1 if (aulas_dadas_int / total_aulas_int) >= 0.75 else 0

        try:
            valor_dev_f = float(valor_dev.replace(",", ".")) if valor_dev else 0.0
        except ValueError:
            valor_dev_f = 0.0

        try:
            # 1. Registra o cancelamento
            cur = db.execute("""
                INSERT INTO cancelamentos (
                    contrato_id, aluno_id, data_cancelamento,
                    motivo, autorizado_por,
                    aulas_dadas, total_aulas, passou_75_pct,
                    valor_parcelas_canc, houve_ressarcimento,
                    valor_devolvido, observacoes
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
            """, (
                contrato_id, contrato["aluno_id"], data_canc,
                motivo or None, autorizado_por or None,
                aulas_dadas_int, total_aulas_int, passou_75,
                valor_canc, houve_res, valor_dev_f,
                observacoes or None,
            ))
            cancelamento_id = cur.lastrowid

            # 2. Cancela parcelas não pagas
            for p in nao_pagas:
                db.execute("""
                    UPDATE parcelas SET
                        cancelada        = 1,
                        cancelamento_id  = ?,
                        excluido_em      = ?
                    WHERE id = ?;
                """, (cancelamento_id, data_canc, p["id"]))

            # 3. Marca parcelas pagas com o cancelamento (sem alterar status)
            for p in pagas:
                db.execute("""
                    UPDATE parcelas SET cancelamento_id = ?
                    WHERE id = ?;
                """, (cancelamento_id, p["id"]))

            # 4. Cancela o contrato
            db.execute("""
                UPDATE contratos SET
                    ativo      = 0,
                    cancelado  = 1,
                    excluido_em = ?
                WHERE id = ?;
            """, (data_canc, contrato_id))

            # 5. Desativa o aluno
            db.execute("""
                UPDATE alunos SET ativo = 0 WHERE id = ?;
            """, (contrato["aluno_id"],))

            # 6. Atualiza rematrícula para não vai renovar
            db.execute("""
                UPDATE rematriculas SET
                    status = 'nao_vai_renovar',
                    motivo_nao_renovou = COALESCE(motivo_nao_renovou, 'Contrato cancelado')
                WHERE aluno_id = ? AND status NOT IN ('nao_vai_renovar');
            """, (contrato["aluno_id"],))

            db.commit()

        except Exception as e:
            db.rollback()
            return f"Erro ao registrar cancelamento: {str(e)}", 500

        return redirect(url_for("detalhe_aluno", aluno_id=contrato["aluno_id"]))

    # GET — mostra formulário de confirmação
    return render_template(
        "cancelar_contrato.html",
        contrato=contrato,
        pagas=pagas,
        nao_pagas=nao_pagas,
        valor_canc=valor_canc,
        total_parc=total_parc,
        hoje=date.today().isoformat(),
    )


@app.post("/contrato/<int:contrato_id>/reativar-cancelado")
def reativar_contrato_cancelado(contrato_id):
    """Reverte um cancelamento — só para casos de erro."""
    db = get_db()

    contrato = db.execute(
        "SELECT aluno_id, cancelado FROM contratos WHERE id = ?;",
        (contrato_id,)
    ).fetchone()

    if not contrato or not contrato["cancelado"]:
        return "Contrato não está cancelado", 400

    reativado_por = request.form.get("reativado_por", "").strip()
    hoje_iso      = date.today().isoformat()

    try:
        # Reativa contrato
        db.execute("""
            UPDATE contratos SET
                ativo = 1, cancelado = 0, excluido_em = NULL
            WHERE id = ?;
        """, (contrato_id,))

        # Reativa parcelas canceladas por este contrato
        db.execute("""
            UPDATE parcelas SET
                cancelada = 0, excluido_em = NULL, cancelamento_id = NULL
            WHERE contrato_id = ?
              AND cancelada = 1;
        """, (contrato_id,))

        # Reativa o aluno
        db.execute(
            "UPDATE alunos SET ativo = 1 WHERE id = ?;",
            (contrato["aluno_id"],)
        )

        # Marca o cancelamento como revertido
        db.execute("""
            UPDATE cancelamentos SET
                reativado_em  = ?,
                reativado_por = ?
            WHERE contrato_id = ?
              AND reativado_em IS NULL;
        """, (hoje_iso, reativado_por or "Sistema", contrato_id))

        db.commit()

    except Exception as e:
        db.rollback()
        return f"Erro ao reativar: {str(e)}", 500

    return redirect(url_for("detalhe_aluno", aluno_id=contrato["aluno_id"]))
@app.post("/aluno/<int:aluno_id>/reativar")
def reativar_aluno(aluno_id):
    db = get_db()
    db.execute("UPDATE alunos SET ativo = 1 WHERE id = ?;", (aluno_id,))
    db.commit()
    return redirect(url_for("detalhe_aluno", aluno_id=aluno_id))

@app.post("/professores/<int:prof_id>/toggle")
def toggle_professor(prof_id):
    db = get_db()
    db.execute("""
        UPDATE professores
        SET ativo = CASE WHEN ativo = 1 THEN 0 ELSE 1 END
        WHERE id = ?;
    """, (prof_id,))
    db.commit()
    return redirect(url_for("listar_professores"))
 
 
@app.post("/professores/<int:prof_id>/excluir")
def excluir_professor(prof_id):
    db = get_db()
    turmas_ativas = db.execute(
        "SELECT COUNT(*) FROM turmas WHERE professor_id = ? AND (ativa = 1 OR ativa IS NULL);",
        (prof_id,)
    ).fetchone()[0]
    if turmas_ativas > 0:
        return "Não é possível excluir: professor tem turmas ativas. Reatribua as turmas antes.", 400
    db.execute("DELETE FROM professores WHERE id = ?;", (prof_id,))
    db.commit()
    return redirect(url_for("listar_professores"))


# ── TURMAS ───────────────────────────────────────────────────

@app.route("/turmas")
def listar_turmas():
    db = get_db()
    turmas = db.execute("""
        SELECT t.id, t.tipo, t.numero, t.horario, t.dia_semana,
               t.data_inicio, t.data_fim, t.codigo_astro,
               t.ativa, t.periodo,
               p.nome AS professor,
               p.id   AS professor_id,
               COUNT(r.id) AS total_alunos
        FROM turmas t
        LEFT JOIN professores p ON t.professor_id = p.id
        LEFT JOIN rematriculas r ON r.turma_id = t.id
        GROUP BY t.id
        ORDER BY t.tipo, t.numero, t.horario;
    """).fetchall()
    professores = db.execute(
        "SELECT id, nome FROM professores WHERE ativo = 1 ORDER BY nome;"
    ).fetchall()
    return render_template("turmas.html",
                           turmas=turmas, professores=professores)


@app.post("/turmas/nova")
def nova_turma():
    db = get_db()
    tipo        = request.form.get("tipo", "").strip()
    numero      = request.form.get("numero", "").strip()
    professor_id= request.form.get("professor_id", "").strip()
    horario     = request.form.get("horario", "").strip()
    dia_semana  = request.form.get("dia_semana", "").strip()
    data_inicio = request.form.get("data_inicio", "").strip()
    data_fim    = request.form.get("data_fim", "").strip()
    codigo_astro= request.form.get("codigo_astro", "").strip()
    periodo     = request.form.get("periodo", "").strip()
 
    if not tipo or not numero:
        return redirect(url_for("listar_turmas"))
 
    db.execute("""
        INSERT INTO turmas
            (tipo, numero, professor_id, horario, dia_semana,
             data_inicio, data_fim, codigo_astro, periodo, ativa)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 1);
    """, (
        tipo,
        numero,
        professor_id or None,
        horario or None,
        dia_semana or None,
        data_inicio or None,
        data_fim or None,
        codigo_astro or None,
        periodo or None,
    ))
    db.commit()
    return redirect(url_for("listar_turmas"))


@app.post("/turmas/<int:turma_id>/editar")
def editar_turma(turma_id):
    db = get_db()
    db.execute("""
        UPDATE turmas SET
            tipo         = ?,
            numero       = ?,
            professor_id = ?,
            horario      = ?,
            dia_semana   = ?,
            data_inicio  = ?,
            data_fim     = ?,
            codigo_astro = ?,
            periodo      = ?
        WHERE id = ?;
    """, (
        request.form.get("tipo", "").strip(),
        request.form.get("numero", "").strip(),
        request.form.get("professor_id") or None,
        request.form.get("horario") or None,
        request.form.get("dia_semana") or None,
        request.form.get("data_inicio") or None,
        request.form.get("data_fim") or None,
        request.form.get("codigo_astro") or None,
        request.form.get("periodo") or None,
        turma_id,
    ))
    db.commit()
    return redirect(url_for("listar_turmas"))


@app.post("/turmas/<int:turma_id>/toggle")
def toggle_turma(turma_id):
    db = get_db()
    db.execute("""
        UPDATE turmas
        SET ativa = CASE WHEN ativa = 1 THEN 0 ELSE 1 END
        WHERE id = ?;
    """, (turma_id,))
    db.commit()
    return redirect(url_for("listar_turmas"))
 
 
@app.post("/turmas/<int:turma_id>/excluir")
def excluir_turma(turma_id):
    db = get_db()
    alunos = db.execute(
        "SELECT COUNT(*) FROM rematriculas WHERE turma_id = ?;",
        (turma_id,)
    ).fetchone()[0]
    if alunos > 0:
        return "Não é possível excluir: turma tem alunos vinculados. Mova ou cancele os alunos antes.", 400
    db.execute("DELETE FROM turmas WHERE id = ?;", (turma_id,))
    db.commit()
    return redirect(url_for("listar_turmas"))


# ── API: busca alunos (autocomplete) ─────────────────────────

@app.get("/api/alunos/buscar")
def api_buscar_alunos():
    q = request.args.get("q", "").strip()
    if len(q) < 2:
        return jsonify([])
    termo = f"%{q}%"
    alunos = get_db().execute("""
        SELECT a.id, a.nome, r.nome AS responsavel
        FROM alunos a
        LEFT JOIN responsaveis r ON a.responsavel_id = r.id
        WHERE a.nome LIKE ? AND (a.ativo = 1 OR a.ativo IS NULL)
        ORDER BY a.nome LIMIT 10;
    """, (termo,)).fetchall()
    return jsonify([dict(a) for a in alunos])


# ── REMATRÍCULAS ─────────────────────────────────────────────

@app.route("/rematriculas")
def listar_rematriculas():
    db  = get_db()
    aba = request.args.get("aba", "turmas")
 
    # Filtros comuns
    periodo_filtro  = request.args.get("periodo", "").strip()
    professor_filtro= request.args.get("professor_id", "").strip()
 
    # Listas para os selects
    professores = db.execute(
        "SELECT id, nome FROM professores WHERE ativo = 1 ORDER BY nome;"
    ).fetchall()
    periodos = db.execute(
        "SELECT DISTINCT periodo FROM turmas WHERE periodo IS NOT NULL ORDER BY periodo DESC;"
    ).fetchall()
 
    # ── ABA: por turma ───────────────────────────────────────
    if aba == "turmas":
        sql = """
            SELECT t.id AS turma_id,
                   t.tipo, t.numero, t.horario, t.dia_semana,
                   t.data_inicio, t.data_fim, t.periodo,
                   p.nome AS professor,
                   COUNT(r.id) AS total,
                   SUM(CASE WHEN r.status IN ('renovou','automatico') THEN 1 ELSE 0 END) AS renovados,
                   SUM(CASE WHEN r.status = 'nao_vai_renovar'         THEN 1 ELSE 0 END) AS nao_renovaram,
                   SUM(CASE WHEN r.status = 'em_analise'              THEN 1 ELSE 0 END) AS em_analise,
                   SUM(CASE WHEN r.status = 'pendente'                THEN 1 ELSE 0 END) AS pendentes
            FROM turmas t
            LEFT JOIN professores p  ON t.professor_id = p.id
            LEFT JOIN rematriculas r ON r.turma_id = t.id
            WHERE (t.ativa = 1 OR t.ativa IS NULL)
        """
        params = []
        if periodo_filtro:
            sql += " AND t.periodo = ?"
            params.append(periodo_filtro)
        if professor_filtro:
            sql += " AND t.professor_id = ?"
            params.append(professor_filtro)
        sql += " GROUP BY t.id ORDER BY t.tipo, t.numero, t.horario"
        turmas = db.execute(sql, params).fetchall()
 
        return render_template("rematriculas.html",
            aba=aba, turmas=turmas,
            professores=professores, periodos=periodos,
            periodo_filtro=periodo_filtro,
            professor_filtro=professor_filtro,
        )
 
    # ── ABA: por professor ───────────────────────────────────
    elif aba == "professor":
        sql = """
            SELECT p.id AS professor_id, p.nome AS professor,
                   COUNT(DISTINCT t.id) AS total_turmas,
                   COUNT(r.id) AS total_alunos,
                   SUM(CASE WHEN r.status IN ('renovou','automatico') THEN 1 ELSE 0 END) AS renovados,
                   SUM(CASE WHEN r.status = 'nao_vai_renovar'         THEN 1 ELSE 0 END) AS nao_renovaram,
                   SUM(CASE WHEN r.status = 'em_analise'              THEN 1 ELSE 0 END) AS em_analise,
                   SUM(CASE WHEN r.status = 'pendente'                THEN 1 ELSE 0 END) AS pendentes
            FROM professores p
            LEFT JOIN turmas t       ON t.professor_id = p.id
            LEFT JOIN rematriculas r ON r.turma_id = t.id
            WHERE p.ativo = 1
        """
        params = []
        if periodo_filtro:
            sql += " AND t.periodo = ?"
            params.append(periodo_filtro)
        if professor_filtro:
            sql += " AND p.id = ?"
            params.append(professor_filtro)
        sql += " GROUP BY p.id ORDER BY p.nome"
        por_professor = db.execute(sql, params).fetchall()
 
        return render_template("rematriculas.html",
            aba=aba, por_professor=por_professor,
            professores=professores, periodos=periodos,
            periodo_filtro=periodo_filtro,
            professor_filtro=professor_filtro,
        )
 
    # ── ABA: por aluno ───────────────────────────────────────
    else:
        q             = request.args.get("q", "").strip()
        status_filtro = request.args.get("status_filtro", "").strip()
        alunos_rem    = []
 
        if q or status_filtro:
            sql = """
                SELECT r.id, r.status, r.tipo_contrato,
                       r.desconto_anterior, r.motivo_nao_renovou,
                       r.professor, r.observacoes, r.periodo,
                       a.id AS aluno_id, a.nome AS aluno,
                       t.tipo AS turma_tipo, t.numero AS turma_numero,
                       t.horario, t.dia_semana, t.periodo AS turma_periodo,
                       t.data_inicio, t.data_fim,
                       p.nome AS professor_nome,
                       pt.tipo AS proxima_tipo, pt.numero AS proxima_numero,
                       pt.horario AS proxima_horario
                FROM rematriculas r
                JOIN alunos a           ON r.aluno_id       = a.id
                LEFT JOIN turmas t      ON r.turma_id       = t.id
                LEFT JOIN turmas pt     ON r.proxima_turma_id = pt.id
                LEFT JOIN professores p ON t.professor_id   = p.id
                WHERE (a.ativo = 1 OR a.ativo IS NULL)
            """
            params = []
            if q:
                sql += " AND a.nome LIKE ?"
                params.append(f"%{q}%")
            if status_filtro:
                sql += " AND r.status = ?"
                params.append(status_filtro)
            sql += " ORDER BY a.nome, r.periodo DESC;"
            from datetime import date as _date
            rows = db.execute(sql, params).fetchall()
            alunos_rem = []
            hoje_d = _date.today()
            for row in rows:
                r_ = dict(row)
                try:
                    if row["data_inicio"] and row["data_fim"]:
                        ini = _date.fromisoformat(row["data_inicio"])
                        fim = _date.fromisoformat(row["data_fim"])
                        total = (fim - ini).days or 1
                        passados = (hoje_d - ini).days
                        r_["pct_curso"] = max(0, min(100, round((passados / total) * 100)))
                    else:
                        r_["pct_curso"] = 0
                except Exception:
                    r_["pct_curso"] = 0
                alunos_rem.append(r_)
 
        return render_template("rematriculas.html",
            aba=aba, alunos_rem=alunos_rem, q=q,
            status_filtro=status_filtro,
            professores=professores, periodos=periodos,
            periodo_filtro=periodo_filtro,
            professor_filtro=professor_filtro,
        )


@app.route("/rematriculas/turma/<int:turma_id>")
def detalhe_turma_rematricula(turma_id):
    db = get_db()
    turma = db.execute("""
        SELECT t.*, p.nome AS professor
        FROM turmas t
        LEFT JOIN professores p ON t.professor_id = p.id
        WHERE t.id = ?;
    """, (turma_id,)).fetchone()
 
    if not turma:
        return "Turma não encontrada", 404
 
    alunos = db.execute("""
        SELECT r.id, r.status, r.tipo_contrato,
               r.desconto_anterior, r.motivo_nao_renovou,
               r.observacoes, r.periodo,
               a.id AS aluno_id, a.nome AS aluno,
               pt.id AS proxima_id,
               pt.tipo AS proxima_tipo, pt.numero AS proxima_numero,
               pt.horario AS proxima_horario
        FROM rematriculas r
        JOIN alunos a
          ON r.aluno_id = a.id
        LEFT JOIN turmas pt
          ON r.proxima_turma_id = pt.id
        WHERE r.turma_id = ?
        ORDER BY a.nome;
    """, (turma_id,)).fetchall()
 
    todas_turmas = db.execute("""
        SELECT id, tipo, numero, horario, dia_semana
        FROM turmas
        WHERE ativa = 1 AND id != ?
        ORDER BY tipo, numero, horario;
    """, (turma_id,)).fetchall()
 
    return render_template("detalhe_turma.html",
        turma=turma, alunos=alunos, todas_turmas=todas_turmas)


# ── 1) nova_rematricula ────────────────────────────────────────

@app.post("/rematriculas/nova")
def nova_rematricula():
    db = get_db()
    aluno_id         = request.form.get("aluno_id", "").strip()
    turma_id         = request.form.get("turma_id", "").strip()
    proxima_turma_id = request.form.get("proxima_turma_id", "").strip()
    tipo_contrato    = request.form.get("tipo_contrato", "renovar")
    desconto         = request.form.get("desconto_anterior", "0").strip()
    status           = request.form.get("status", "pendente")
    motivo           = request.form.get("motivo_nao_renovou", "").strip()
    observacoes      = request.form.get("observacoes", "").strip()
    periodo          = request.form.get("periodo", "").strip()
    origem           = request.form.get("origem", "")
 
    if not aluno_id or not turma_id:
        return "Aluno e turma são obrigatórios", 400
 
    # Garante que aluno existe no cadastro
    aluno = db.execute(
        "SELECT id FROM alunos WHERE id = ? AND (ativo = 1 OR ativo IS NULL);",
        (aluno_id,)
    ).fetchone()
    if not aluno:
        return "Aluno não encontrado no cadastro", 400
 
    # Se tem próxima turma, status é automático
    if proxima_turma_id:
        status = "automatico"
 
    try:
        db.execute("""
            INSERT INTO rematriculas
                (aluno_id, turma_id, proxima_turma_id, periodo,
                 tipo_contrato, desconto_anterior, status,
                 motivo_nao_renovou, observacoes)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?);
        """, (
            aluno_id,
            turma_id,
            proxima_turma_id or None,
            periodo or None,
            tipo_contrato,
            float(desconto) if desconto else 0.0,
            status,
            motivo or None,
            observacoes or None,
        ))
        db.commit()
    except Exception as e:
        db.rollback()
        return f"Erro: {str(e)}", 500
 
    if origem:
        return redirect(url_for("detalhe_turma_rematricula", turma_id=origem))
    return redirect(url_for("listar_rematriculas"))


# ── 2) editar_rematricula ─────────────────────────────────────

@app.post("/rematriculas/<int:rem_id>/editar")
def editar_rematricula(rem_id):
    db = get_db()
    proxima_turma_id = request.form.get("proxima_turma_id", "").strip()
    status           = request.form.get("status", "pendente")
    motivo           = request.form.get("motivo_nao_renovou", "").strip()
    observacoes      = request.form.get("observacoes", "").strip()
    desconto         = request.form.get("desconto_anterior", "0").strip()
    origem_turma     = request.form.get("origem_turma", "").strip()
 
    if proxima_turma_id:
        status = "automatico"
 
    db.execute("""
        UPDATE rematriculas SET
            proxima_turma_id  = ?,
            status            = ?,
            motivo_nao_renovou= ?,
            observacoes       = ?,
            desconto_anterior = ?,
            atualizado_em     = CURRENT_TIMESTAMP
        WHERE id = ?;
    """, (
        proxima_turma_id or None,
        status,
        motivo or None,
        observacoes or None,
        float(desconto) / 100 if desconto else 0.0,
        rem_id,
    ))
    db.commit()
 
    if origem_turma:
        return redirect(url_for("detalhe_turma_rematricula", turma_id=origem_turma))
    return redirect(url_for("listar_rematriculas"))

@app.post("/rematriculas/<int:rem_id>/excluir")
def excluir_rematricula(rem_id):
    db = get_db()
    rem = db.execute(
        "SELECT turma_id FROM rematriculas WHERE id = ?;", (rem_id,)
    ).fetchone()
    db.execute("DELETE FROM rematriculas WHERE id = ?;", (rem_id,))
    db.commit()
    origem = request.form.get("origem_turma", "")
    if origem:
        return redirect(url_for("detalhe_turma_rematricula", turma_id=origem))
    return redirect(url_for("listar_rematriculas"))

@app.route("/aluno/<int:aluno_id>/editar", methods=["GET", "POST"])
def editar_aluno(aluno_id):
    db = get_db()
    
    if request.method == "POST":
        nome = request.form.get("nome")
        data_nascimento = request.form.get("data_nascimento")
        
        db.execute(
            """
            UPDATE alunos
            SET nome = ?, data_nascimento = ?
            WHERE id = ?;
            """,
            (nome, data_nascimento, aluno_id),
        )
        db.commit()
        return redirect(url_for("detalhe_aluno", aluno_id=aluno_id))
    
    aluno = db.execute(
        "SELECT * FROM alunos WHERE id = ?;",
        (aluno_id,),
    ).fetchone()
    
    if aluno is None:
        return "Aluno não encontrado", 404
    
    return render_template("aluno_editar.html", aluno=aluno)
# Inicia backup automático (a cada 24 horas)


#<!-- def open_browser():
    # isso é só pro servidor abrir o navegador localmente
#    webbrowser.open_new("http://127.0.0.1:5001/")

# ... seu código atual do app3.py ...

if __name__ == "__main__":
    import socket
    import time
    import sys
    
    # PORTA FIXA 5000
    PORT = 5000
    
    print("\n" + "="*60)
    print("🚀 SISTEMA FINANCEIRO - PORTA 5000")
    print("="*60)
    
    # TENTA LIBERAR A PORTA
    def liberar_porta(porta):
        try:
            # Tenta matar processos usando a porta
            import subprocess
            import os
            if os.name == 'nt':  # Windows
                result = subprocess.run(
                    f'netstat -ano | findstr :{porta}',
                    shell=True,
                    capture_output=True,
                    text=True
                )
                if result.stdout:
                    for line in result.stdout.strip().split('\n'):
                        if 'LISTENING' in line:
                            parts = line.split()
                            pid = parts[-1]
                            print(f"🔫 Matando processo {pid} na porta {porta}...")
                            subprocess.run(f'taskkill /F /PID {pid}', shell=True)
            time.sleep(2)
            return True
        except:
            return False
    
    # Tenta liberar a porta
    print(f"🔧 Preparando porta {PORT}...")
    liberar_porta(PORT)
    
    # Testa se porta está livre
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        sock.bind(('0.0.0.0', PORT))
        sock.close()
        print(f"✅ Porta {PORT} está livre e pronta!")
    except OSError as e:
        print(f"❌ ERRO: Não foi possível usar porta {PORT}")
        print(f"   Detalhe: {e}")
        print("\n📌 Soluções:")
        print("   1. Feche outros programas usando porta 5000")
        print("   2. Reinicie o computador")
        print("   3. Execute 'limpar_portas.bat' como Administrador")
        input("\nPressione Enter para sair...")
        sys.exit(1)
    
    # Obtém IP local
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        local_ip = s.getsockname()[0]
        s.close()
    except:
        local_ip = "127.0.0.1"
    
    print(f"\n🌐 SERVIDOR PRONTO!")
    print(f"   • URL Local: http://localhost:{PORT}")
    print(f"   • URL Rede:  http://{local_ip}:{PORT}")
    print(f"   • Computador: {socket.gethostname()}")
    print("="*60)
    print("\n⏳ Iniciando... (Ctrl+C para parar)\n")
    
    # INICIA SERVIDOR NA PORTA 5001
    app.run(
        host="0.0.0.0",
        port=PORT,  # SEMPRE 5001
        debug=True,
        threaded=True,
        use_reloader=False
    )