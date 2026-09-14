"""
MIGRATION V6 — Campos de estorno na tabela pagamentos
======================================================
Execute na pasta teste_financeiro_v2 com o servidor parado.
python migration_v6.py
"""

import sqlite3
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DATABASE = str(BASE_DIR / "financeiro_escola.db")

conn = sqlite3.connect(DATABASE)
conn.row_factory = sqlite3.Row
cur  = conn.cursor()

print("\n" + "="*55)
print("   MIGRATION V6 — Estorno de pagamentos")
print("="*55 + "\n")

def col_existe(tabela, coluna):
    cur.execute(f"PRAGMA table_info({tabela});")
    return coluna in [r["name"] for r in cur.fetchall()]

ok = []

campos = {
    "estornado":       "INTEGER DEFAULT 0",
    "estornado_em":    "TEXT",
    "estornado_por":   "TEXT",
    "motivo_estorno":  "TEXT",
}

for col, defn in campos.items():
    if not col_existe("pagamentos", col):
        cur.execute(f"ALTER TABLE pagamentos ADD COLUMN {col} {defn};")
        ok.append(f"pagamentos.{col} adicionado")
        print(f"  ✅ {col} adicionado")
    else:
        print(f"  ⏭  {col} já existe")

conn.commit()
conn.close()

print(f"\n✅ Migration V6 concluída — {len(ok)} alteração(ões)")
input("\nPressione Enter para fechar...")
