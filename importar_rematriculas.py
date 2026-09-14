"""
IMPORTAR REMATRÍCULAS — Período Fevereiro a Julho 2026
======================================================
Execute este script UMA ÚNICA VEZ na pasta teste_financeiro_v2.

O que faz:
  - Lê a planilha Rematrículas_CTRLPlay.xlsx
  - Cria as turmas do período no banco
  - Vincula cada aluno à sua turma com status de renovação
  - Pula alunos não encontrados no banco e avisa no final

Como executar:
  1. Servidor Flask parado
  2. Coloque este script na pasta teste_financeiro_v2
  3. Coloque a planilha na mesma pasta (ou ajuste CAMINHO_XLSX abaixo)
  4. Execute: python importar_rematriculas.py
"""

import sqlite3
import os
import sys
import pandas as pd
from pathlib import Path
from datetime import datetime

# ─────────────────────────────────────────────────────
# CONFIGURAÇÃO
# ─────────────────────────────────────────────────────
BASE_DIR     = Path(__file__).resolve().parent
DATABASE     = str(BASE_DIR / "financeiro_escola.db")
CAMINHO_XLSX = str(BASE_DIR / "Rematrículas_CTRLPlay.xlsx")
ABA          = "Período - Fevereiro a julho de "
PERIODO      = "Fevereiro a Julho 2026"

# ─────────────────────────────────────────────────────
# VERIFICAÇÕES INICIAIS
# ─────────────────────────────────────────────────────
if not os.path.exists(DATABASE):
    print(f"❌ Banco não encontrado: {DATABASE}")
    input("\nPressione Enter para sair...")
    sys.exit(1)

if not os.path.exists(CAMINHO_XLSX):
    print(f"❌ Planilha não encontrada: {CAMINHO_XLSX}")
    input("\nPressione Enter para sair...")
    sys.exit(1)

print("\n" + "="*60)
print("   IMPORTAÇÃO DE REMATRÍCULAS — Fev a Jul 2026")
print("="*60)
print(f"\n📂 Banco:    {DATABASE}")
print(f"📊 Planilha: {CAMINHO_XLSX}")
print(f"🕐 Iniciado: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")

conn = sqlite3.connect(DATABASE)
conn.row_factory = sqlite3.Row
cur  = conn.cursor()

# ─────────────────────────────────────────────────────
# MAPEAMENTO DE STATUS
# ─────────────────────────────────────────────────────
STATUS_MAP = {
    "renovou":          "renovou",
    "automático":       "automatico",
    "automatico":       "automatico",
    "em análise":       "em_analise",
    "em analise":       "em_analise",
    "não vai renovar":  "nao_vai_renovar",
    "nao vai renovar":  "nao_vai_renovar",
    "não continuará":   "nao_vai_renovar",
    "nao continuará":   "nao_vai_renovar",
}

TIPO_MAP = {
    "renovar":          "renovar",
    "rema automática":  "automatico",
    "rema automatica":  "automatico",
}

def normaliza(s):
    """Remove espaços extras e converte para minúsculas."""
    if not s or str(s).strip() == "nan":
        return ""
    return str(s).strip().lower()

def formata_data(val):
    """Converte datetime do pandas para string YYYY-MM-DD."""
    if pd.isna(val):
        return None
    if hasattr(val, 'strftime'):
        return val.strftime("%Y-%m-%d")
    return str(val)[:10]

# ─────────────────────────────────────────────────────
# BLOCO 1 — LER TURMAS DA PLANILHA
# ─────────────────────────────────────────────────────
print("─"*50)
print("BLOCO 1: Lendo turmas da planilha")
print("─"*50)

df = pd.read_excel(CAMINHO_XLSX, sheet_name=ABA, header=None)

# Turmas ficam nas linhas 16 a 22 (índices), colunas 1,2,3,4,5
# Col 1=nome, 2=código astro, 3=término, 4=próxima turma, 5=nº alunos
LINHA_TURMAS_INICIO = 16
LINHA_TURMAS_FIM    = 23  # exclusive

turmas_planilha = []
for i in range(LINHA_TURMAS_INICIO, LINHA_TURMAS_FIM):
    row = df.iloc[i]
    nome = str(row[1]).strip() if not pd.isna(row[1]) else ""
    if not nome or nome == "nan":
        continue
    turmas_planilha.append({
        "nome":         nome,
        "codigo_astro": str(int(row[2])) if not pd.isna(row[2]) else "",
        "data_termino": formata_data(row[3]),
        "data_proxima": formata_data(row[4]),
    })
    print(f"  📋 Turma encontrada: {nome}")

print(f"\n  Total: {len(turmas_planilha)} turma(s)\n")

# ─────────────────────────────────────────────────────
# BLOCO 2 — CRIAR TURMAS NO BANCO
# ─────────────────────────────────────────────────────
print("─"*50)
print("BLOCO 2: Criando turmas no banco")
print("─"*50)

turma_id_map = {}  # nome_turma -> id no banco

for t in turmas_planilha:
    # Verifica se já existe pelo código Astro ou nome
    existente = cur.execute(
        "SELECT id FROM turmas WHERE codigo_astro = ? OR nome = ?;",
        (t["codigo_astro"], t["nome"])
    ).fetchone()

    if existente:
        turma_id_map[t["nome"]] = existente["id"]
        print(f"  ⏭️  Turma já existe: {t['nome']} (id={existente['id']})")
    else:
        cur.execute("""
            INSERT INTO turmas (nome, codigo_astro, data_termino, data_proxima, ativa)
            VALUES (?, ?, ?, ?, 1);
        """, (t["nome"], t["codigo_astro"], t["data_termino"], t["data_proxima"]))
        turma_id_map[t["nome"]] = cur.lastrowid
        print(f"  ✅ Turma criada: {t['nome']} (id={cur.lastrowid})")

conn.commit()
print()

# ─────────────────────────────────────────────────────
# BLOCO 3 — LER ALUNOS DA PLANILHA
# ─────────────────────────────────────────────────────
print("─"*50)
print("BLOCO 3: Lendo alunos da planilha")
print("─"*50)

# Alunos começam na linha 31 (índice), colunas:
# 1=nome, 2=tipo, 3=turma, 4=desconto, 5=status, 6=motivo, 8=professor
LINHA_ALUNOS_INICIO = 31
LINHA_ALUNOS_FIM    = 81  # exclusive (até linha 80)

alunos_planilha = []
for i in range(LINHA_ALUNOS_INICIO, LINHA_ALUNOS_FIM):
    row = df.iloc[i]
    nome = str(row[1]).strip() if not pd.isna(row[1]) else ""
    if not nome or nome == "nan":
        continue

    # Pega só os campos das colunas principais (1-8)
    # ignora colunas 11+ que são tabela de conecta embutida
    turma_nome = str(row[3]).strip() if not pd.isna(row[3]) else ""
    tipo_raw   = normaliza(row[2])
    status_raw = normaliza(row[5])
    motivo     = str(row[6]).strip() if not pd.isna(row[6]) else ""
    professor  = str(row[8]).strip() if not pd.isna(row[8]) else ""
    desconto   = float(row[4]) if not pd.isna(row[4]) else 0.0

    # Ignora linhas que são cabeçalhos da tabela de conecta
    if nome.upper() in ("ALUNOS", "ALUNOS CONECTA"):
        continue

    tipo_contrato = TIPO_MAP.get(tipo_raw, "renovar")
    status        = STATUS_MAP.get(status_raw, "pendente")

    alunos_planilha.append({
        "nome":           nome,
        "turma_nome":     turma_nome,
        "tipo_contrato":  tipo_contrato,
        "desconto":       desconto,
        "status":         status,
        "motivo":         motivo if motivo and motivo != "nan" else None,
        "professor":      professor if professor and professor != "nan" else None,
    })

print(f"  Total de alunos na planilha: {len(alunos_planilha)}\n")

# ─────────────────────────────────────────────────────
# BLOCO 4 — IMPORTAR REMATRÍCULAS
# ─────────────────────────────────────────────────────
print("─"*50)
print("BLOCO 4: Importando rematrículas")
print("─"*50)

importados  = []
nao_encontrados = []
ja_existia  = []

for aluno in alunos_planilha:
    # Busca aluno no banco pelo nome (case-insensitive)
    aluno_db = cur.execute(
        "SELECT id, nome FROM alunos WHERE LOWER(TRIM(nome)) = LOWER(TRIM(?)) AND (ativo = 1 OR ativo IS NULL);",
        (aluno["nome"],)
    ).fetchone()

    if not aluno_db:
        nao_encontrados.append(aluno["nome"])
        print(f"  ⚠️  NÃO ENCONTRADO: {aluno['nome']}")
        continue

    # Resolve turma_id
    turma_id = None
    for nome_turma, tid in turma_id_map.items():
        # Casa pelo início do nome (planilha pode ter variações)
        if aluno["turma_nome"].lower()[:20] in nome_turma.lower() or \
           nome_turma.lower()[:20] in aluno["turma_nome"].lower():
            turma_id = tid
            break

    # Verifica se já existe rematrícula para esse aluno nesse período
    existente = cur.execute(
        "SELECT id FROM rematriculas WHERE aluno_id = ? AND periodo = ?;",
        (aluno_db["id"], PERIODO)
    ).fetchone()

    if existente:
        ja_existia.append(aluno["nome"])
        print(f"  ⏭️  Já importado: {aluno['nome']}")
        continue

    # Insere a rematrícula
    cur.execute("""
        INSERT INTO rematriculas (
            aluno_id, turma_id, periodo,
            tipo_contrato, desconto_anterior,
            status, motivo_nao_renovou, professor
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?);
    """, (
        aluno_db["id"],
        turma_id,
        PERIODO,
        aluno["tipo_contrato"],
        aluno["desconto"],
        aluno["status"],
        aluno["motivo"],
        aluno["professor"],
    ))
    importados.append(aluno["nome"])
    print(f"  ✅ {aluno['nome']} ({aluno['status']})")

conn.commit()
conn.close()

# ─────────────────────────────────────────────────────
# RESULTADO FINAL
# ─────────────────────────────────────────────────────
print("\n" + "="*60)
print("RESULTADO DA IMPORTAÇÃO")
print("="*60)
print(f"\n✅ Importados com sucesso: {len(importados)}")
print(f"⏭️  Já existiam no banco:  {len(ja_existia)}")
print(f"⚠️  Não encontrados:       {len(nao_encontrados)}")

if nao_encontrados:
    print("\n" + "─"*50)
    print("ALUNOS NÃO ENCONTRADOS — corrija o nome no Flask:")
    print("─"*50)
    for nome in nao_encontrados:
        print(f"  • {nome}")
    print("\nDica: o nome na planilha precisa ser idêntico ao")
    print("cadastrado no sistema (incluindo acentos e maiúsculas).")

print(f"\n🕐 Finalizado: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
print("="*60)
input("\nPressione Enter para fechar...")
