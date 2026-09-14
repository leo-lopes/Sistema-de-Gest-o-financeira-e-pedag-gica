"""
CRIAR BANCO DE DEMONSTRAÇÃO
============================
Gera um banco SQLite com dados fictícios para demonstração no GitHub.
Execute: python criar_banco_demo.py
Isso cria o arquivo: demo.db
"""

import sqlite3
import random
from datetime import date, timedelta
from pathlib import Path

DB = "demo.db"

# Remove banco anterior se existir
Path(DB).unlink(missing_ok=True)

conn = sqlite3.connect(DB)
cur  = conn.cursor()

# ── SCHEMA ───────────────────────────────────────────────────
cur.executescript("""
CREATE TABLE responsaveis (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nome TEXT NOT NULL,
    telefone TEXT,
    email TEXT,
    cpf_cnpj TEXT,
    endereco_rua TEXT,
    endereco_numero TEXT,
    endereco_bairro TEXT,
    endereco_cidade TEXT,
    endereco_uf TEXT,
    ativo INTEGER DEFAULT 1
);

CREATE TABLE alunos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nome TEXT NOT NULL,
    data_nascimento TEXT,
    responsavel_id INTEGER REFERENCES responsaveis(id),
    ativo INTEGER DEFAULT 1
);

CREATE TABLE professores (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nome TEXT NOT NULL,
    contato TEXT,
    ativo INTEGER DEFAULT 1,
    criado_em TEXT
);

CREATE TABLE turmas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nome TEXT,
    tipo TEXT,
    numero TEXT,
    professor_id INTEGER REFERENCES professores(id),
    horario TEXT,
    dia_semana TEXT,
    data_inicio TEXT,
    data_fim TEXT,
    codigo_astro TEXT,
    periodo TEXT,
    ativa INTEGER DEFAULT 1
);

CREATE TABLE contratos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    aluno_id INTEGER REFERENCES alunos(id),
    responsavel_id INTEGER REFERENCES responsaveis(id),
    numero_contrato TEXT,
    data_inicio TEXT,
    data_fim TEXT,
    valor_total REAL,
    qtd_parcelas INTEGER,
    ativo INTEGER DEFAULT 1,
    cancelado INTEGER DEFAULT 0,
    excluido_em TEXT,
    modulos TEXT,
    asaas_id TEXT
);

CREATE TABLE parcelas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    contrato_id INTEGER REFERENCES contratos(id),
    numero_parcela INTEGER,
    valor_original REAL,
    valor_final REAL,
    data_vencimento TEXT,
    cancelada INTEGER DEFAULT 0,
    status_pagamento TEXT DEFAULT 'pendente',
    pago_em TEXT,
    valor_pago REAL,
    canal_pagamento TEXT,
    observacao_pagamento TEXT,
    registrado_por TEXT,
    cancelamento_id INTEGER,
    excluido_em TEXT
);

CREATE TABLE notas_fiscais (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    parcela_id INTEGER REFERENCES parcelas(id),
    numero_nf TEXT,
    data_emissao TEXT,
    valor REAL,
    status TEXT DEFAULT 'emitida'
);

CREATE TABLE rematriculas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    aluno_id INTEGER REFERENCES alunos(id),
    turma_id INTEGER REFERENCES turmas(id),
    periodo TEXT,
    tipo_contrato TEXT,
    desconto_anterior REAL DEFAULT 0,
    status TEXT DEFAULT 'pendente',
    motivo_nao_renovou TEXT,
    observacoes TEXT,
    renovacao_automatica INTEGER DEFAULT 0,
    criado_em TEXT,
    atualizado_em TEXT
);

CREATE TABLE pagamentos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    parcela_id INTEGER REFERENCES parcelas(id),
    valor_pago REAL,
    data_pagamento TEXT,
    canal TEXT,
    origem TEXT DEFAULT 'manual',
    registrado_por TEXT,
    observacao TEXT,
    criado_em TEXT,
    estornado INTEGER DEFAULT 0,
    estornado_em TEXT,
    estornado_por TEXT,
    motivo_estorno TEXT
);

CREATE TABLE cancelamentos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    contrato_id INTEGER REFERENCES contratos(id),
    aluno_id INTEGER REFERENCES alunos(id),
    data_cancelamento TEXT,
    motivo TEXT,
    autorizado_por TEXT,
    aulas_dadas INTEGER,
    total_aulas INTEGER,
    passou_75_pct INTEGER DEFAULT 0,
    valor_parcelas_canc REAL DEFAULT 0,
    houve_ressarcimento INTEGER DEFAULT 0,
    valor_devolvido REAL DEFAULT 0,
    observacoes TEXT,
    criado_em TEXT,
    reativado_em TEXT,
    reativado_por TEXT
);
""")

# ── DADOS FICTÍCIOS ──────────────────────────────────────────

NOMES_RESP = [
    "Carlos Silva", "Ana Oliveira", "Roberto Santos", "Maria Costa",
    "João Ferreira", "Luciana Alves", "Pedro Rodrigues", "Fernanda Lima",
    "Marcos Souza", "Patricia Mendes", "Ricardo Carvalho", "Juliana Martins",
    "Bruno Nascimento", "Camila Rocha", "Felipe Araujo", "Beatriz Gomes",
    "Thiago Pereira", "Vanessa Barbosa", "Leonardo Costa", "Cristina Dias",
]
NOMES_ALUNOS = [
    "Gabriel Silva", "Sofia Oliveira", "Miguel Santos", "Isabella Costa",
    "Arthur Ferreira", "Valentina Alves", "Heitor Rodrigues", "Laura Lima",
    "Davi Souza", "Manuela Mendes", "Lorenzo Carvalho", "Julia Martins",
    "Mateus Nascimento", "Alice Rocha", "Rafael Araujo", "Mariana Gomes",
    "Nicolas Pereira", "Giovanna Barbosa", "Guilherme Costa", "Leticia Dias",
    "Lucas Ribeiro", "Ana Clara Cunha", "Pedro Henrique Lopes", "Isabela Castro",
    "Enzo Monteiro", "Yasmin Cardoso", "Gabriel Teixeira", "Clara Moreira",
    "Samuel Correia", "Helena Pinto",
]
PROFESSORES = ["Leo", "John", "Novato", "Ana Paula", "Rafael"]
CANAIS = ["rede", "link_rede", "asaas", "c6", "pix_manual", "dinheiro"]
TIPOS  = ["Kids", "Teens", "Young"]
DIAS   = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado"]
STATUS_REM = ["renovou", "renovou", "renovou", "automatico", "pendente", "em_analise", "nao_vai_renovar"]

hoje = date.today()

# Professores
for i, nome in enumerate(PROFESSORES, 1):
    cur.execute("INSERT INTO professores (nome, ativo) VALUES (?, 1);", (nome,))

# Turmas
turma_ids = []
for tipo in TIPOS:
    for num in ["1", "2", "3", "4"]:
        for _ in range(random.randint(1, 2)):
            prof_id = random.randint(1, len(PROFESSORES))
            dia     = random.choice(DIAS)
            h_ini   = random.choice(["08:00", "10:00", "13:00", "15:00", "18:30", "19:00"])
            inicio  = date(2026, 2, 1).isoformat()
            fim     = date(2026, 8, 31).isoformat()
            cur.execute("""
                INSERT INTO turmas (nome, tipo, numero, professor_id, horario,
                                    dia_semana, data_inicio, data_fim, periodo, ativa)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, '02-2026', 1);
            """, (f"{tipo} {num}", tipo, num, prof_id, h_ini, dia, inicio, fim))
            turma_ids.append(cur.lastrowid)

# Responsáveis, alunos, contratos, parcelas
for i, nome_aluno in enumerate(NOMES_ALUNOS):
    resp_nome = NOMES_RESP[i % len(NOMES_RESP)]
    cur.execute("""
        INSERT INTO responsaveis (nome, telefone, email, ativo)
        VALUES (?, ?, ?, 1);
    """, (resp_nome, f"(31) 9{random.randint(1000,9999)}-{random.randint(1000,9999)}",
          f"resp{i+1}@email.com"))
    resp_id = cur.lastrowid

    nasc = date(random.randint(2012, 2020), random.randint(1,12), random.randint(1,28))
    cur.execute("""
        INSERT INTO alunos (nome, data_nascimento, responsavel_id, ativo)
        VALUES (?, ?, ?, 1);
    """, (nome_aluno, nasc.isoformat(), resp_id))
    aluno_id = cur.lastrowid

    # Contrato
    qtd_parc   = random.choice([6, 12])
    valor_parc = round(random.uniform(280, 480), 2)
    valor_total = round(valor_parc * qtd_parc, 2)
    inicio_c   = date(2025, random.randint(7,12), random.randint(1,28))
    fim_c      = date(inicio_c.year + 1, inicio_c.month, inicio_c.day)
    num_cont   = f"CT-{100 + i}"

    cur.execute("""
        INSERT INTO contratos (aluno_id, responsavel_id, numero_contrato,
                               data_inicio, data_fim, valor_total, qtd_parcelas, ativo)
        VALUES (?, ?, ?, ?, ?, ?, ?, 1);
    """, (aluno_id, resp_id, num_cont, inicio_c.isoformat(),
          fim_c.isoformat(), valor_total, qtd_parc))
    cont_id = cur.lastrowid

    # Parcelas
    for p in range(1, qtd_parc + 1):
        venc = date(inicio_c.year + (inicio_c.month + p - 2) // 12,
                    (inicio_c.month + p - 1) % 12 or 12,
                    min(inicio_c.day, 28))
        status = "pendente"
        pago_em = None
        valor_pago = None
        canal = None

        if venc < hoje:
            if random.random() < 0.75:
                status   = "pago"
                pago_em  = (venc + timedelta(days=random.randint(0,5))).isoformat()
                valor_pago = valor_parc
                canal    = random.choice(CANAIS)
            else:
                status = "atrasado"

        cur.execute("""
            INSERT INTO parcelas (contrato_id, numero_parcela, valor_original,
                                  valor_final, data_vencimento, cancelada,
                                  status_pagamento, pago_em, valor_pago, canal_pagamento)
            VALUES (?, ?, ?, ?, ?, 0, ?, ?, ?, ?);
        """, (cont_id, p, valor_parc, valor_parc, venc.isoformat(),
              status, pago_em, valor_pago, canal))
        parc_id = cur.lastrowid

        # Nota fiscal para pagas com vencimento passado
        if status == "pago" and venc < hoje and random.random() < 0.8:
            cur.execute("""
                INSERT INTO notas_fiscais (parcela_id, numero_nf, data_emissao, valor, status)
                VALUES (?, ?, ?, ?, 'emitida');
            """, (parc_id, f"NF-{parc_id}-{pago_em}", pago_em, valor_parc))

    # Rematrícula
    turma_id = random.choice(turma_ids)
    status_rem = random.choice(STATUS_REM)
    cur.execute("""
        INSERT INTO rematriculas (aluno_id, turma_id, periodo, tipo_contrato,
                                  desconto_anterior, status, renovacao_automatica)
        VALUES (?, ?, '02-2026', 'renovar', ?, ?, ?);
    """, (aluno_id, turma_id, round(random.uniform(0, 0.35), 2),
          status_rem, 1 if status_rem == 'automatico' else 0))

conn.commit()
conn.close()
print(f"✅ Banco demo criado: {DB}")
print(f"   {len(NOMES_ALUNOS)} alunos · {len(PROFESSORES)} professores · {len(turma_ids)} turmas")
