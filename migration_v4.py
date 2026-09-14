"""
MIGRATION V4 — Simplificação de rematrículas
=============================================
Remove dependência de proxima_turma_id (mantém a coluna para não quebrar,
mas adiciona renovacao_automatica como boolean).
Execute na pasta teste_financeiro_v2 com o servidor parado.
"""
import sqlite3
from pathlib import Path
from datetime import datetime

BASE_DIR = Path(__file__).resolve().parent
DATABASE = str(BASE_DIR / "financeiro_escola.db")

conn = sqlite3.connect(DATABASE)
conn.row_factory = sqlite3.Row
cur  = conn.cursor()

print("\n" + "="*55)
print("   MIGRATION V4")
print("="*55 + "\n")

def col_existe(tabela, coluna):
    cur.execute(f"PRAGMA table_info({tabela});")
    return coluna in [r["name"] for r in cur.fetchall()]

ok = []

# Adiciona campo renovacao_automatica (1=sim, 0=não)
if not col_existe("rematriculas", "renovacao_automatica"):
    cur.execute("ALTER TABLE rematriculas ADD COLUMN renovacao_automatica INTEGER DEFAULT 0;")
    ok.append("rematriculas.renovacao_automatica adicionado")
    print("  ✅ renovacao_automatica adicionado")
else:
    print("  ⏭  renovacao_automatica já existe")

# Migra dados existentes: tipo_contrato='automatico' → renovacao_automatica=1
cur.execute("""
    UPDATE rematriculas
    SET renovacao_automatica = 1
    WHERE tipo_contrato = 'automatico'
      AND renovacao_automatica = 0;
""")
migrados = cur.rowcount
if migrados:
    ok.append(f"{migrados} registro(s) migrados para renovacao_automatica=1")
    print(f"  ✅ {migrados} registro(s) migrados")

conn.commit()
conn.close()

print(f"\n✅ Migration V4 concluída — {len(ok)} alteração(ões)")
input("\nPressione Enter para fechar...")
