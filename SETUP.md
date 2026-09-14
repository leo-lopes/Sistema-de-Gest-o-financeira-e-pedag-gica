# Guia de configuração — CTRLPlay

## Passos para subir o projeto do zero

### 1. Dependências Python

```bash
pip install flask reportlab openpyxl
```

### 2. Banco de dados

**Para demonstração (dados fictícios):**
```bash
python criar_banco_demo.py
mv demo.db financeiro_escola.db
```

**Para produção (banco novo):**
```bash
python migration_v2.py
python migration_v3.py
python migration_v4.py
python migration_v5.py
python migration_v6.py
```

### 3. Iniciar o servidor

```bash
python app3.py
```

O servidor sobe na porta `5000` por padrão.

### 4. Estrutura esperada de pastas

```
pasta_do_projeto/
├── app3.py
├── financeiro_escola.db   ← não vai para o git
├── templates/
│   └── *.html
└── backups/               ← não vai para o git
```

## Variáveis de configuração

No topo do `app3.py`:

| Variável | Padrão | Descrição |
|----------|--------|-----------|
| `PORT` | `5000` | Porta do servidor |
| `DATABASE` | `financeiro_escola.db` | Nome do banco |
| `BACKUP_DIR` | `backups/` | Pasta de backups |
| `PER_PAGE` | `50` | Itens por página nas listagens |
