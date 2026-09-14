"""
MIGRATION V3 — Professores e ajustes de turmas/rematrículas
============================================================
Execute na pasta teste_financeiro_v2 com o servidor parado.
python migration_v3.py
"""

import sqlite3, sys
from pathlib import Path
from datetime import datetime

BASE_DIR = Path(__file__).resolve().parent
DATABASE = str(BASE_DIR / "financeiro_escola.db")

print("\n" + "="*55)
print("   MIGRATION V3")
print("="*55)

conn = sqlite3.connect(DATABASE)
conn.row_factory = sqlite3.Row
cur  = conn.cursor()

def col_existe(tabela, coluna):
    cur.execute(f"PRAGMA table_info({tabela});")
    return coluna in [r["name"] for r in cur.fetchall()]

def tab_existe(tabela):
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?;", (tabela,))
    return cur.fetchone() is not None

ok = []

# ── Tabela professores ────────────────────────────────────────
if not tab_existe("professores"):
    cur.execute("""
        CREATE TABLE professores (
            id       INTEGER PRIMARY KEY AUTOINCREMENT,
            nome     TEXT NOT NULL,
            contato  TEXT,
            ativo    INTEGER DEFAULT 1,
            criado_em TEXT
        );
    """)
    ok.append("Tabela 'professores' criada")
else:
    print("  ⏭  professores já existe")

# ── Ajustes na tabela turmas ──────────────────────────────────
novos_turmas = {
    "tipo":          "TEXT",
    "numero":        "TEXT",
    "professor_id":  "INTEGER",
    "dia_semana":    "TEXT",
    "periodo":       "TEXT",
}
for col, defn in novos_turmas.items():
    if not col_existe("turmas", col):
        cur.execute(f"ALTER TABLE turmas ADD COLUMN {col} {defn};")
        ok.append(f"turmas.{col} adicionado")
    else:
        print(f"  ⏭  turmas.{col} já existe")

# ── Ajustes na tabela rematriculas ───────────────────────────
novos_rem = {
    "proxima_turma_id": "INTEGER",
}
for col, defn in novos_rem.items():
    if not col_existe("rematriculas", col):
        cur.execute(f"ALTER TABLE rematriculas ADD COLUMN {col} {defn};")
        ok.append(f"rematriculas.{col} adicionado")
    else:
        print(f"  ⏭  rematriculas.{col} já existe")

conn.commit()
conn.close()

print(f"\n✅ {len(ok)} alteração(ões):")
for a in ok: print(f"  • {a}")
print("\nPronto! Pode subir o servidor.")
input("\nPressione Enter para fechar...")
