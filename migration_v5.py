"""
MIGRATION V5 — Tabela de cancelamentos
=======================================
Execute na pasta teste_financeiro_v2 com o servidor parado.
python migration_v5.py
"""

import sqlite3
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DATABASE = str(BASE_DIR / "financeiro_escola.db")

conn = sqlite3.connect(DATABASE)
conn.row_factory = sqlite3.Row
cur  = conn.cursor()

print("\n" + "="*55)
print("   MIGRATION V5 — Cancelamentos")
print("="*55 + "\n")

def tab_existe(t):
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?;", (t,))
    return cur.fetchone() is not None

def col_existe(tabela, coluna):
    cur.execute(f"PRAGMA table_info({tabela});")
    return coluna in [r["name"] for r in cur.fetchall()]

ok = []

# ── Tabela cancelamentos ──────────────────────────────────────
if not tab_existe("cancelamentos"):
    cur.execute("""
        CREATE TABLE cancelamentos (
            id                  INTEGER PRIMARY KEY AUTOINCREMENT,
            contrato_id         INTEGER NOT NULL REFERENCES contratos(id),
            aluno_id            INTEGER NOT NULL REFERENCES alunos(id),
            data_cancelamento   TEXT NOT NULL,
            motivo              TEXT,
            autorizado_por      TEXT,
            aulas_dadas         INTEGER,
            total_aulas         INTEGER,
            passou_75_pct       INTEGER DEFAULT 0,
            valor_parcelas_canc REAL DEFAULT 0,
            houve_ressarcimento INTEGER DEFAULT 0,
            valor_devolvido     REAL DEFAULT 0,
            observacoes         TEXT,
            criado_em           TEXT DEFAULT CURRENT_TIMESTAMP,
            reativado_em        TEXT,
            reativado_por       TEXT
        );
    """)
    ok.append("Tabela 'cancelamentos' criada")
    print("  ✅ Tabela cancelamentos criada")
else:
    print("  ⏭  cancelamentos já existe")

# ── Campo cancelamento_id nas parcelas (para rastrear) ───────
if not col_existe("parcelas", "cancelamento_id"):
    cur.execute("ALTER TABLE parcelas ADD COLUMN cancelamento_id INTEGER;")
    ok.append("parcelas.cancelamento_id adicionado")
    print("  ✅ parcelas.cancelamento_id adicionado")
else:
    print("  ⏭  parcelas.cancelamento_id já existe")

# ── Campo cancelado nos contratos (além de ativo=0) ──────────
if not col_existe("contratos", "cancelado"):
    cur.execute("ALTER TABLE contratos ADD COLUMN cancelado INTEGER DEFAULT 0;")
    ok.append("contratos.cancelado adicionado")
    print("  ✅ contratos.cancelado adicionado")
else:
    print("  ⏭  contratos.cancelado já existe")

conn.commit()
conn.close()

print(f"\n✅ Migration V5 concluída — {len(ok)} alteração(ões)")
input("\nPressione Enter para fechar...")
