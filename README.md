# 🦺 Controle de EPI / EPC / Ferramentas — Excel + VBA

Sistema de controle de almoxarifado para **Equipamentos de Proteção Individual (EPI)**, **Equipamentos de Proteção Coletiva (EPC)** e **Ferramentas**, desenvolvido inteiramente em Excel com macros VBA.

---

## 📋 Sobre o Projeto

Centraliza o controle de movimentação de EPIs, EPCs e ferramentas, permitindo rastrear entradas, saídas, baixas, vencimentos e gerar relatórios gerenciais — tudo sem necessidade de banco de dados externo ou software adicional.

---

## ✨ Funcionalidades

- **Dashboard gerencial** com indicadores visuais de:
  - Total de entradas e saídas em R$
  - TOP 10 itens por quantidade de entrada e saída
  - TOP 10 itens por valor financeiro
  - TOP 10 itens por vida útil média
  - TOP 10 funcionários por volume de saídas

- **Registro de movimentações** (Lançamentos) com:
  - ID único por item
  - Data, matrícula e nome do colaborador
  - Tipo de movimentação (Entrada / Saída)
  - Justificativa e motivo de baixa
  - Categoria (EPI / EPC / FER / DIV)
  - Código e descrição do item
  - Quantidade, valor unitário e valor total
  - Data de vencimento, CA e número de série
  - Controle de baixa (data, motivo, vida útil em dias)

- **Emissão de formulários** via VBA:
  - `frm_Cautela` — Formulário de Cautela individual
  - `frm_Vale` — Formulário de Vale para descontos em caso de perda/extravio pelo colaborador
  - `frm_Movimentação` — Formulário para Movimentação/transferência
  - `frm_Requisição` — Formulário para Requisição de compra

- **Gestão de vencimentos** com alertas de itens próximos ao vencimento (vencido, 30/60/90 dias)

- **Análise de compra** com sugestão de reposição baseada em consumo e saldo atual

- **Catálogo de produtos** com código, descrição, categoria, exigência de CA/validade/série e saldo em estoque

- **Cadastro de colaboradores** com matrícula e função

---

## 🗂️ Estrutura das Abas

| Aba | Descrição |
|---|---|
| `Dashboard` | Painel de indicadores e gráficos gerenciais |
| `Lançamentos` | Base de dados principal com todas as movimentações |
| `Auxiliar` | Tabelas de referência (justificativas, categorias, bases, catálogo de produtos, cadastro de funcionários) |
| `Vencimentos` | Controle de itens com data de vencimento |
| `Análise de Compra` | Sugestão de compras baseada em consumo e estoque |
| `frm_Cautela` | Layout para impressão de ficha de cautela |
| `frm_Vale` | Layout para impressão de vale de material |
| `frm_Movimentação` | Layout para impressão de relatório de movimentação |
| `frm_Requisição` | Layout para impressão de requisição de compra |

---

## 🏷️ Categorias de Itens

| Código | Categoria | Exemplos |
|---|---|---|
| `EPI` | Equipamento de Proteção Individual | Capacete, luvas, óculos, botas, cinto paraquedista, manga isolante |
| `EPC` | Equipamento de Proteção Coletiva | Cones, fitas de sinalização |
| `FER` | Ferramentas | Alicates, chaves, multímetro, carretilha |
| `DIV` | Diversos / Acessórios | Bolsas, maletas, cordas de serviço, cadeados |

---

## 📊 Tipos de Movimentação

**Entradas:**
- Saldo Inicial
- Reposição Normal
- Reposição de Urgência
- Retorno ao Estoque
- Transferência
- Mobilização

**Saídas:**
- Admissão
- Desgaste Natural
- Mau Uso / Extravio
- Manutenção
- Troca
- Vencimento
- Desligamento

---

## ⚙️ Requisitos

- Microsoft Excel 2016 ou superior
- Habilitar macros (VBA) e Controles ActiveX ao abrir o arquivo
- Extensão do arquivo: `.xlsm` (Excel com macros habilitadas)

---

## 🔒 Observações

- O sistema utiliza fórmulas avançadas e macros VBA para automação dos formulários
- Faça backup regularmente, pois toda a base de dados está contida no próprio arquivo `.xlsm`

---

## 📄 Licença

Desenvolvido por: Pedro Ricardo Moreno.
