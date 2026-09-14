"""
MIGRATION V2 - Sistema Financeiro CTRLPlay
==========================================
Execute este script UMA ÚNICA VEZ no ambiente de testes (teste_financeiro_v2).

O que este script faz:
  - Adiciona campos novos em tabelas existentes (sem apagar nada)
  - Cria as tabelas novas: turmas, modulos_contrato, pagamentos
  - É seguro: verifica antes de alterar, nunca destrói dados

Como executar:
  1. Certifique-se de que o servidor Flask NÃO está rodando
  2. Abra o terminal na pasta teste_financeiro_v2
  3. Execute: python migration_v2.py
"""

import sqlite3
import os
import sys
from pathlib import Path
from datetime import datetime

# ─────────────────────────────────────────
# LOCALIZA O BANCO NA MESMA PASTA DO SCRIPT
# ─────────────────────────────────────────
BASE_DIR = Path(__file__).resolve().parent
DATABASE = str(BASE_DIR / "financeiro_escola.db")

if not os.path.exists(DATABASE):
    print(f"❌ ERRO: Banco de dados não encontrado em:")
    print(f"   {DATABASE}")
    print(f"\nCertifique-se de que este script está na pasta teste_financeiro_v2")
    input("\nPressione Enter para sair...")
    sys.exit(1)

print("\n" + "="*60)
print("   MIGRATION V2 - Sistema Financeiro CTRLPlay")
print("="*60)
print(f"\n📂 Banco: {DATABASE}")
print(f"🕐 Iniciado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
print("\nIniciando migration...\n")

conn = sqlite3.connect(DATABASE)
conn.row_factory = sqlite3.Row
cur = conn.cursor()

erros = []
alteracoes = []


def coluna_existe(tabela, coluna):
    """Verifica se uma coluna já existe em uma tabela."""
    cur.execute(f"PRAGMA table_info({tabela});")
    colunas = [row["name"] for row in cur.fetchall()]
    return coluna in colunas


def tabela_existe(tabela):
    """Verifica se uma tabela já existe no banco."""
    cur.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name=?;",
        (tabela,)
    )
    return cur.fetchone() is not None


def add_coluna(tabela, coluna, definicao):
    """Adiciona coluna se não existir. Retorna True se adicionou."""
    if not coluna_existe(tabela, coluna):
        cur.execute(f"ALTER TABLE {tabela} ADD COLUMN {coluna} {definicao};")
        alteracoes.append(f"  ✅ {tabela}.{coluna} adicionado")
        return True
    else:
        print(f"  ⏭️  {tabela}.{coluna} já existe, pulando")
        return False


# ═══════════════════════════════════════════════════════
# BLOCO 1 — CAMPOS NOVOS EM TABELAS EXISTENTES
# ═══════════════════════════════════════════════════════
print("─" * 50)
print("BLOCO 1: Campos novos em tabelas existentes")
print("─" * 50)

try:
    # ── CONTRATOS ──
    # Módulos do contrato (ex: "CT1 + CT2") - texto descritivo
    add_coluna("contratos", "modulos", "TEXT")

    # Vínculo com sistema externo (loja.ctrlplay / Asaas) - para uso futuro
    add_coluna("contratos", "asaas_id", "TEXT")
    add_coluna("contratos", "loja_contrato_id", "TEXT")

    # ── PARCELAS ──
    # Status do pagamento desta parcela
    add_coluna("parcelas", "status_pagamento",
               "TEXT DEFAULT 'pendente' CHECK(status_pagamento IN ('pendente','pago','atrasado','cancelado'))")

    # Dados do pagamento quando ocorrer
    add_coluna("parcelas", "pago_em", "TEXT")          # data do pagamento (YYYY-MM-DD)
    add_coluna("parcelas", "valor_pago", "REAL")        # pode diferir do valor_final
    add_coluna("parcelas", "canal_pagamento",
               "TEXT CHECK(canal_pagamento IN ('rede','link_rede','asaas','c6','pix_manual','dinheiro','boleto','outro'))")
    add_coluna("parcelas", "observacao_pagamento", "TEXT")
    add_coluna("parcelas", "registrado_por", "TEXT")    # quem lançou o pagamento

    conn.commit()
    print()

except Exception as e:
    erros.append(f"Bloco 1: {str(e)}")
    print(f"  ❌ ERRO no Bloco 1: {e}")


# ═══════════════════════════════════════════════════════
# BLOCO 2 — TABELA DE TURMAS
# ═══════════════════════════════════════════════════════
print("─" * 50)
print("BLOCO 2: Tabela de turmas")
print("─" * 50)

try:
    if not tabela_existe("turmas"):
        cur.execute("""
            CREATE TABLE turmas (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                nome            TEXT NOT NULL,
                codigo_astro    TEXT,
                professor       TEXT,
                horario         TEXT,
                modulo          TEXT,
                data_termino    TEXT,
                data_proxima    TEXT,
                ativa           INTEGER DEFAULT 1,
                criado_em       TEXT DEFAULT CURRENT_TIMESTAMP,
                observacoes     TEXT
            );
        """)
        alteracoes.append("  ✅ Tabela 'turmas' criada")
        conn.commit()
    else:
        print("  ⏭️  Tabela 'turmas' já existe, pulando")

    print()

except Exception as e:
    erros.append(f"Bloco 2: {str(e)}")
    print(f"  ❌ ERRO no Bloco 2: {e}")


# ═══════════════════════════════════════════════════════
# BLOCO 3 — TABELA DE REMATRÍCULAS
# ═══════════════════════════════════════════════════════
print("─" * 50)
print("BLOCO 3: Tabela de rematrículas")
print("─" * 50)

try:
    if not tabela_existe("rematriculas"):
        cur.execute("""
            CREATE TABLE rematriculas (
                id                  INTEGER PRIMARY KEY AUTOINCREMENT,
                aluno_id            INTEGER NOT NULL REFERENCES alunos(id),
                turma_id            INTEGER REFERENCES turmas(id),
                periodo             TEXT NOT NULL,
                tipo_contrato       TEXT CHECK(tipo_contrato IN ('renovar','automatico')),
                desconto_anterior   REAL DEFAULT 0,
                status              TEXT DEFAULT 'pendente'
                                    CHECK(status IN (
                                        'pendente',
                                        'renovou',
                                        'automatico',
                                        'em_analise',
                                        'nao_vai_renovar'
                                    )),
                motivo_nao_renovou  TEXT,
                professor           TEXT,
                criado_em           TEXT DEFAULT CURRENT_TIMESTAMP,
                atualizado_em       TEXT DEFAULT CURRENT_TIMESTAMP,
                observacoes         TEXT
            );
        """)
        alteracoes.append("  ✅ Tabela 'rematriculas' criada")
        conn.commit()
    else:
        print("  ⏭️  Tabela 'rematriculas' já existe, pulando")

    print()

except Exception as e:
    erros.append(f"Bloco 3: {str(e)}")
    print(f"  ❌ ERRO no Bloco 3: {e}")


# ═══════════════════════════════════════════════════════
# BLOCO 4 — TABELA DE HISTÓRICO DE PAGAMENTOS
# (registro detalhado de cada transação)
# ═══════════════════════════════════════════════════════
print("─" * 50)
print("BLOCO 4: Tabela de histórico de pagamentos")
print("─" * 50)

try:
    if not tabela_existe("pagamentos"):
        cur.execute("""
            CREATE TABLE pagamentos (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                parcela_id      INTEGER NOT NULL REFERENCES parcelas(id),
                valor_pago      REAL NOT NULL,
                data_pagamento  TEXT NOT NULL,
                canal           TEXT NOT NULL
                                CHECK(canal IN (
                                    'rede',
                                    'link_rede',
                                    'asaas',
                                    'c6',
                                    'pix_manual',
                                    'dinheiro',
                                    'boleto',
                                    'outro'
                                )),
                origem          TEXT DEFAULT 'manual'
                                CHECK(origem IN ('manual','asaas_webhook','importacao')),
                registrado_por  TEXT,
                observacao      TEXT,
                criado_em       TEXT DEFAULT CURRENT_TIMESTAMP,
                asaas_payment_id TEXT
            );
        """)
        alteracoes.append("  ✅ Tabela 'pagamentos' criada")
        conn.commit()
    else:
        print("  ⏭️  Tabela 'pagamentos' já existe, pulando")

    print()

except Exception as e:
    erros.append(f"Bloco 4: {str(e)}")
    print(f"  ❌ ERRO no Bloco 4: {e}")


# ═══════════════════════════════════════════════════════
# RESULTADO FINAL
# ═══════════════════════════════════════════════════════
print("=" * 60)
print("RESULTADO DA MIGRATION")
print("=" * 60)

if alteracoes:
    print(f"\n✅ {len(alteracoes)} alteração(ões) aplicada(s):\n")
    for a in alteracoes:
        print(a)
else:
    print("\n⏭️  Nenhuma alteração necessária (tudo já estava atualizado)")

if erros:
    print(f"\n❌ {len(erros)} erro(s) encontrado(s):\n")
    for e in erros:
        print(f"  • {e}")
    print("\n⚠️  Verifique os erros acima antes de continuar.")
else:
    print("\n✅ Migration concluída sem erros!")

conn.close()

print(f"\n🕐 Finalizado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
print("=" * 60)
print("\nPode iniciar o servidor de testes normalmente agora.")
input("\nPressione Enter para fechar...")
