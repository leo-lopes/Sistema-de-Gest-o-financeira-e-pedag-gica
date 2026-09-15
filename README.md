# CTRLPlay — Sistema de Gestão Financeira e Pedagógica

Sistema web completo desenvolvido para gestão financeira e pedagógica de uma escola de programação. Construído com Python/Flask e SQLite, com foco em praticidade para secretaria e coordenação.

---

## Demo ao vivo

**[https://sistema-de-gest-o-financeira-e-pedag-gica.onrender.com](https://sistema-de-gest-o-financeira-e-pedag-gica.onrender.com)**

> O servidor pode levar ~30 segundos para acordar na primeira visita (plano gratuito do Render).
> Os dados exibidos são fictícios — gerados automaticamente para demonstração.

---

## Contexto

Este projeto nasceu da necessidade real de integrar três fluxos distintos que operavam de forma desconectada: emissão de notas fiscais, controle de pagamentos e acompanhamento de rematrículas. O desafio central foi criar uma solução leve, segura (rodando em rede privada via Tailscale) e fácil de usar por uma equipe não técnica.

---

## Funcionalidades

### Módulo Financeiro
- Registro de pagamentos por múltiplos canais (maquininha, PIX, boleto, C6, Asaas)
- Baixa de múltiplas parcelas simultaneamente
- Estorno e edição de pagamentos com histórico de auditoria
- Detecção automática de inadimplência
- Relatório mensal em PDF e Excel (por período)

### Módulo de Notas Fiscais
- Emissão individual e em lote
- Seleção por período com filtro de datas
- Seleção de todas as parcelas do filtro (não apenas da página atual)

### Módulo de Rematrículas
- Cadastro de turmas por tipo (Kids, Teens, Young) e semestre
- Vinculação de professores a turmas
- Acompanhamento de renovações por turma, professor e aluno
- Barra de progresso do curso baseada em datas de início e fim
- Renovação automática para alunos com dois módulos assinados
- Taxa de renovação por turma

### Gestão de Alunos
- Cadastro completo com responsável, contrato e turma em uma única tela
- Histórico de alunos inativos e cancelados
- Ficha do aluno com contratos, parcelas recentes e status de rematrícula

### Cancelamentos
- Registro formal de cancelamentos com cálculo da regra dos 75% de aulas
- Controle de ressarcimento
- Reversão de cancelamentos por erro com trilha de auditoria

### Dashboard
- Quadro de avisos automático (turmas encerrando, inadimplência, pendências)
- Estatísticas em tempo real
- Dark mode com preferência salva no navegador

---

## Stack

| Camada | Tecnologia |
|--------|-----------|
| Backend | Python 3.13 + Flask |
| Banco de dados | SQLite |
| Frontend | HTML + CSS (variáveis CSS, dark mode nativo) + JavaScript vanilla |
| Relatórios | ReportLab (PDF) + OpenPyXL (Excel) |
| Infraestrutura | Tailscale (rede privada) |
| Deploy | Render |

---

## Estrutura do projeto

```
├── app3.py              # Aplicação Flask principal (~2400 linhas)
├── templates/           # Templates Jinja2
│   ├── base.html        # Layout base com dark mode e navegação
│   ├── index.html       # Dashboard com quadro de avisos
│   ├── pagamentos.html  # Módulo financeiro
│   ├── rematriculas.html
│   ├── turmas.html
│   ├── professores.html
│   ├── notas_a_emitir.html
│   └── ...
├── migrations/          # Scripts de migration do banco
│   ├── migration_v2.py  # Tabelas de pagamentos e rematrículas
│   ├── migration_v3.py  # Professores e turmas
│   ├── migration_v4.py  # Renovação automática
│   ├── migration_v5.py  # Cancelamentos
│   └── migration_v6.py  # Estorno de pagamentos
├── criar_banco_demo.py  # Gera banco com dados fictícios para teste
├── requirements.txt     # Dependências Python
├── render.yaml          # Configuração de deploy no Render
└── .gitignore           # Exclui bancos reais e dados sensíveis
```

---

## Como acessar

### Opção 1 — Demo online (mais rápido)
Acesse diretamente:
**[https://sistema-de-gest-o-financeira-e-pedag-gica.onrender.com](https://sistema-de-gest-o-financeira-e-pedag-gica.onrender.com)**

### Opção 2 — Rodar localmente

```bash
# 1. Clone o repositório
git clone https://github.com/leo-lopes/Sistema-de-Gest-o-financeira-e-pedag-gica.git
cd Sistema-de-Gest-o-financeira-e-pedag-gica

# 2. Instale as dependências
pip install -r requirements.txt

# 3. Crie o banco de demonstração
python criar_banco_demo.py

# 4. Renomeie para o nome esperado pelo sistema
# Linux/Mac:
mv demo.db financeiro_escola.db
# Windows:
rename demo.db financeiro_escola.db

# 5. Inicie o servidor
python app3.py
```

Acesse: `http://localhost:5000`

### Opção 3 — Deploy próprio no Render (gratuito)

1. Faça fork deste repositório
2. Acesse [render.com](https://render.com) e crie uma conta com GitHub
3. Clique em **New → Web Service**
4. Selecione o repositório forkado
5. Configure:
   - **Build Command:** `pip install -r requirements.txt && python criar_banco_demo.py && mv demo.db financeiro_escola.db`
   - **Start Command:** `gunicorn app3:app --bind 0.0.0.0:$PORT`
   - **Instance Type:** Free
6. Clique em **Deploy web service**

Em ~5 minutos o site estará no ar com sua própria URL.

---

## Decisões técnicas relevantes

**SQLite em produção**
Escolha consciente para um sistema com ~70 usuários e acesso concorrente baixo. Elimina a necessidade de um servidor de banco de dados separado e simplifica backups (cópia de um único arquivo).

**Sem framework JS**
Todo o frontend é JavaScript vanilla. Reduz dependências, facilita manutenção por uma pessoa só e carrega instantaneamente.

**Dark mode via CSS custom properties**
O tema escuro é implementado inteiramente via variáveis CSS (`--bg-surface`, `--text-primary`, etc.) no elemento raiz, sem JavaScript além do toggle. A preferência é salva no `localStorage` e o sistema respeita a preferência do sistema operacional via `prefers-color-scheme`.

**Migrations incrementais**
Cada versão adiciona colunas ou tabelas sem tocar nos dados existentes, permitindo atualizar o sistema em produção sem downtime ou perda de dados.

**Rede privada (Tailscale)**
O sistema nunca é exposto à internet pública. Outros membros da equipe acessam via IP da rede Tailscale, eliminando a necessidade de autenticação no nível da aplicação.

---

## Aspectos de Data Science

- **Análise de retenção** por turma e professor (taxa de rematrícula)
- **Detecção de inadimplência** com cálculo automático de dias em atraso
- **Progresso do aluno** calculado por interpolação linear entre datas de início e fim
- **Relatórios financeiros** com agregações por período (recebido, em aberto, emitido)
- **Quadro de avisos preditivo** que identifica turmas prestes a encerrar com alunos de renovação automática

---

## Próximas etapas planejadas

- [ ] Integração com API do Asaas (confirmação automática de pagamentos)
- [ ] Webhook para receber eventos de pagamento em tempo real
- [ ] Dashboard com gráficos de evolução mensal (Recharts ou Chart.js)
- [ ] Publicação segura com Tailscale Funnel

---

## Licença

MIT — uso livre com atribuição.
