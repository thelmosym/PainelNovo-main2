# Manual Técnico e Contratual: Apuração dos Indicadores FDM, TSPE, TSPR e IAPFARQ

> **Documento Oficial de Governança e Regras de Negócio de Medição Petrobras**  
> **Contratos:** Lote 1 - IRON MOUNTAIN (SAP 4600686987 | Jurídico 5900.0133210.25.2) e Lote 2 - PA (SAP 4600687006 | Jurídico 5900.0133214.25.2)  
> **Amparo Contratual:** Itens 8.4.2 e 8.4.3 do Edital de Guarda Externa e Gerenciamento de Acervo Documental  
> **Planilha de Auditoria Vinculada:** [`MEDIÇÃO/Demonstrativo_Calculo_FDM_TSPE_IAPFARQ.xlsx`](../MEDIÇÃO/Demonstrativo_Calculo_FDM_TSPE_IAPFARQ.xlsx)

---

## 1. Visão Geral e Objetivos do Documento

Este manual consolida todas as premissas, fórmulas matemáticas, regras de negócio e fluxos sistêmicos implementados no **Painel de Controle de Medição** para o cálculo dos indicadores de qualidade e nível de serviço (SLA) da Petrobras:

* **TSPE:** Total de Solicitações Padrão Esperado (denominador de demanda elegível).
* **TSPR:** Total de Solicitações Padrão Realizado no Prazo (numerador de eficácia).
* **IAPFARQ:** Índice de Atendimento aos Prazos de Arquivo (percentual de pontualidade).
* **FDM:** Fator de Desempenho Mensal (redutor financeiro de fatura por descumprimento de SLA).

O objetivo principal destes indicadores é assegurar a excelência operacional nas entregas, coletas, digitalizações e pesquisas de acervo, ajustando o pagamento mensal da contratada em função de seu cumprimento real dos prazos contratuais.

---

## 2. Conceitos Fundamentais e Definições Contratuais

```
┌───────────────────────────────────────────────────────────────────────────────────────┐
│                              SISTEMÁTICA DE QUALIDADE PETROBRAS                       │
│                                                                                       │
│   ┌────────────────────┐          ┌────────────────────┐          ┌─────────────────┐ │
│   │ Demanda Elegível   │          │ Eficácia no Prazo  │          │ Índice Global   │ │
│   │     ( TSPE )       │ ───────► │     ( TSPR )       │ ───────► │  ( IAPFARQ % )  │ │
│   │   Total (NP + FP)  │          │    Apenas (NP)     │          │  (TSPR / TSPE)  │ │
│   └────────────────────┘          └────────────────────┘          └────────┬────────┘ │
│                                                                            │          │
│                                                                            ▼          │
│                                                                   ┌─────────────────┐ │
│                                                                   │ Fator Financeiro│ │
│                                                                   │     ( FDM )     │ │
│                                                                   │ (0,9500 a 1,00) │ │
│                                                                   └─────────────────┘ │
└───────────────────────────────────────────────────────────────────────────────────────┘
```

### 2.1 TSPE (Total de Solicitações Padrão Esperado)
* **Conceito:** Representa o volume total de ordens de serviço / atendimentos concluídos (`Situação = "CO"`) durante o ciclo de faturamento (dia 26 ao 25) que possuem meta temporal avaliável.
* **Critério de Elegibilidade:** São computadas todas as linhas que receberam a avaliação de prazo `Status Prazo = "NP"` (No Prazo) ou `Status Prazo = "FP"` (Fora do Prazo).
* **Fórmula:**
  $$\text{TSPE} = \sum (\text{Ordens com } StatusPrazo = \text{"NP"}) + \sum (\text{Ordens com } StatusPrazo = \text{"FP"})$$

### 2.2 TSPR (Total de Solicitações Padrão Realizado)
* **Conceito:** Quantidade de solicitações atendidas pela contratada estritamente dentro da data e hora limites estipuladas pelo SLA ou pelo prazo combinado pactuado com a fiscalização.
* **Critério de Elegibilidade:** Registros concluídos onde `Status Prazo = "NP"`.
* **Fórmula:**
  $$\text{TSPR} = \sum (\text{Ordens com } StatusPrazo = \text{"NP"})$$

### 2.3 IAPFARQ (Índice de Atendimento aos Prazos de Arquivo)
* **Conceito:** Indicador percentual de conformidade de prazos, calculado pela razão entre os atendimentos pontuais e o total elegível de demandas.
* **Fórmula Geral (Item 8.4.2 do Contrato):**
  $$\text{IAPFARQ} = \left( \frac{\text{TSPR}}{\text{TSPE}} \right) \times 100$$
* **Casos Especiais e Robustez Matemática:**
  * Se $\text{TSPE} = 0$: $\text{IAPFARQ} = 100,00\%$ (a ausência de demandas no período confere desempenho pleno por presunção).
  * Se $\text{TSPR} = 0$ e $\text{TSPE} > 0$: $\text{IAPFARQ} = 0,00\%$.
  * Clamping defensivo: garante que $\text{TSPR} \le \text{TSPE}$ em qualquer cenário de base.
  * Precisão: arredondado a **2 casas decimais** (`Number.Round(..., 2)`).

### 2.4 FDM (Fator de Desempenho Mensal)
* **Conceito:** Coeficiente multiplicador aplicado diretamente sobre o valor total faturado dos serviços na Memória de Cálculo (`MC`).
* **Função Financeira:**
  * $\text{FDM} = 1,0000 \implies$ Pagamento integral (100% da medição de serviços).
  * $\text{FDM} < 1,0000 \implies$ Retenção contratual (glosa) proporcional ao volume de atrasos apurado.

---

## 3. Segregação por Linha de Serviço Contratual (`Ref. FDM`)

O Contrato Petrobras determina que a qualidade seja avaliada de forma independente para as duas macroatividades de prestação de serviços:

| Linha de Serviço Contratual | Código `Ref. FDM` | Descrição das Atividades Abrangidas | Indicadores Vinculados |
| :--- | :---: | :--- | :--- |
| **Serviço 1: Guarda de Documentos e Informações** | **`1`** | Coleta de embalagem, Entrega de embalagem (normal e expressa), Devolução de empréstimo, Inserção de item avulso, Pesquisa de item avulso (normal e expressa), Destruição segura de documentos e Migração de acervo. | $\text{TSPE}_1$, $\text{TSPR}_1$, $\text{IAPFARQ}_1$, $\text{FDM}_1$ |
| **Serviço 2: Gerenciamento de Documentos e Informações** | **`2`** | Digitalização de documentos (administrativo, contratual, técnico, financeiro, microfilme, microficha), Catalogação de embalagem e item avulso, Organização analítica/simples, Higienização, Indexação, Cópia/Conversão/Gravação de mídias e vídeos. | $\text{TSPE}_2$, $\text{TSPR}_2$, $\text{IAPFARQ}_2$, $\text{FDM}_2$ |
| **Atividades Isentas de FDM** | **`0`** | Armazenamento físico de acervo documental (custódia mensal), Transferência de acervo, Baixa permanente (expurgo sem prazo operacional). | Não pontuam em TSPE nem TSPR. |

> [!NOTE]
> O de/para entre a descrição de cada atividade e seu respectivo `Ref. FDM` é governado pela consulta [`tabRefPrazoFDM.m`](../Codes/CodigosM/tabRefPrazoFDM.m) (tabela estruturada `tabRefPrazoFDM` na aba `tabelaA`).

---

## 4. Fluxo e Algoritmos de Cálculo Detalhados

A apuração completa dos indicadores percorre uma esteira técnica composta por funções em Linguagem M (Power Query) e módulos VBA:

### 4.1 Qualificação do Chamado e Unicidade de OS (`fnClassificarFDM.m`)
Antes de avaliar o prazo, cada linha da base bruta é submetida às seguintes regras de elegibilidade:
1. **Filtro de Situação:** Apenas atendimentos com `Situação = "CO"` (Concluído) são processados. Registros cancelados (`CA`) ou em aberto (`AB`/`EA`) são desprezados.
2. **Prevenção de Falsas Penalidades (Múltiplos Itens por OS):**
   * Se um chamado operacional contiver 3 itens físicos faturados separadamente (ex.: entrega de 3 caixas na mesma OS), a primeira linha é avaliada para SLA e as linhas subsequentes recebem `CalculoFDM = "S"`.
   * Linhas com `CalculoFDM = "S"` são classificadas como **`"FDM definido"`** com `Status Prazo = null`.
   * **Regra de Negócio Petrobras:** Um atraso em uma OS pontua apenas **1 vez** no FDM, impedindo penalidades multiplicadas artificialmente para um único evento logístico.
3. **Atividades Isentas (`Ref. FDM = 0`):**
   * Classificadas como **`"Isento"`** com `Status Prazo = null`, sem impacto no TSPE.

### 4.2 Apuração de Prazos Úteis e SLA (`fnCalcularPrazoSLA.m`)
Para ordens elegíveis regidas por tabela (`Ref. Prazo = "Tabela"`):
1. **Expediente Útil:** A contagem considera a jornada comercial das **08:00 às 17:00** (8 horas úteis por dia útil).
2. **Intervalo Intrajornada:** O intervalo das **12:00 às 13:00** é desconsiderado.
3. **Entrada de Chamados:**
   * Abertura antes das 08:00 $\to$ início às 08:00 do mesmo dia.
   * Abertura após as 17:00 $\to$ início às 08:00 do próximo dia útil.
   * Abertura em sábado, domingo ou feriado $\to$ início às 08:00 do primeiro dia útil subsequente.
4. **Feriados Multilocais:** O sistema realiza cruzamento automático com a base de feriados (`Feriados_2026_2030.csv`), considerando a cidade e UF de atendimento da Petrobras.
5. **Classificação do Status do Prazo:**
   * Se $\text{Data/Hora Fechamento} \le \text{Data/Hora Limite SLA} \implies \text{Status Prazo} = \text{"NP"}$ (No Prazo).
   * Se $\text{Data/Hora Fechamento} > \text{Data/Hora Limite SLA} \implies \text{Status Prazo} = \text{"FP"}$ (Fora do Prazo).
6. **Exceções de Aplicação:**
   * **`ABF` (Autorizado Bloqueio Fiscal)** ou **`RFT` (Reagendamento Fiscal Técnico):** Se a ordem atrasou por solicitação ou impedimento da própria Petrobras, o status é convertido para **`"NP"`**, não prejudicando a contratada.
   * **`IPF` (Impossibilidade Prática Factual pela Contratada):** Mantém o status **`"FP"`**.
7. **Ordens com Prazo Combinado (`Ref. Prazo = "Fiscalização"`):**
   * A data de fechamento é comparada diretamente com o `Prazo combinado`. Se `Prazo combinado` não for informado pela fiscalização, assume-se por padrão `Status Prazo = "NP"`.

---

## 5. Escala Contratual do FDM (Item 8.4.3)

O índice percentual apurado ($\text{IAPFARQ}$) determina o enquadramento na tabela oficial de desempenho:

| Faixa de Cumprimento de Prazos | Fator FDM Bruto | Percentual Faturado | Glosa / Retenção | Classificação Contratual | Consequência Operacional |
| :---: | :---: | :---: | :---: | :--- | :--- |
| **$\text{IAPFARQ} \ge 99,00\%$** | **`1,0000`** | **100,0%** | **0,0%** | **Meta Cumprida (Excelência)** | Pagamento integral sem qualquer retenção. |
| **$95,00\% \le \text{IAPFARQ} < 99,00\%$** | **`0,9900`** | **99,0%** | **1,0%** | **Nível Aceitável / Atenção** | Notificação de alerta para mitigação de filas/rotas. |
| **$90,00\% \le \text{IAPFARQ} < 95,00\%$** | **`0,9800`** | **98,0%** | **2,0%** | **Nível Crítico / Insatisfatório** | Abertura de plano de ação formal para correção de prazos. |
| **$\text{IAPFARQ} < 90,00\%$** | **`0,9700`** | **97,0%** | **3,0%** | **Nível Não Conforme** | Aplicação da penalidade máxima mensal e sujeição à cascata. |

---

## 6. Regra de Degradação em Cascata por Reincidência (`fnRecalcularFDM.m`)

Para coibir a permanência continuada em níveis insatisfatórios de atendimento, o **Item 8.4.3 (Parágrafo 2º)** estipula a degradação progressiva do FDM quando houver reincidência de notas na faixa mais baixa:

```
                  ┌──────────────────────────────────────────────┐
                  │  FDM Mês Vigente (M) < 90% (FDM Bruto = 0,97) │
                  └──────────────────────┬───────────────────────┘
                                         │
             ┌───────────────────────────┴───────────────────────────┐
             ▼                                                       ▼
┌──────────────────────────┐                            ┌──────────────────────────┐
│  FDM(M-1) >= 0,98        │                            │  FDM(M-1) <= 0,97        │
│  (1º mês de insuficiência)│                           │  (Reincidência no 2º mês)│
└────────────┬─────────────┘                            └────────────┬─────────────┘
             │                                                       │
             ▼                                                       ▼
┌──────────────────────────┐                            ┌──────────────────────────┐
│ Mantém FDM = 0,9700      │                            │ Analisa FDM(M-2)         │
│ (Glosa padrão de 3,0%)   │                            └────────────┬─────────────┘
└──────────────────────────┘                                         │
                                             ┌───────────────────────┴───────────────────────┐
                                             ▼                                               ▼
                                ┌──────────────────────────┐                    ┌──────────────────────────┐
                                │ FDM(M-2) > 0,97          │                    │ FDM(M-2) <= 0,97         │
                                │ (Reincidência em 2 meses)│                    │ (3 meses consecutivos)   │
                                └────────────┬─────────────┘                    └────────────┬─────────────┘
                                             │                                               │
                                             ▼                                               ▼
                                ┌──────────────────────────┐                    ┌──────────────────────────┐
                                │ Degrada para FDM = 0,9600│                    │ Degrada para FDM = 0,9500│
                                │ (Glosa total de 4,0%)    │                    │ (PISO MÁXIMO: Glosa 5,0%)│
                                └──────────────────────────┘                    └──────────────────────────┘
```

### Regras Formais de Transição:
1. Se $\text{FDM} > 0,97 \implies$ Mantém o FDM apurado sem degradação.
2. Se $\text{FDM} = 0,97$ e $\text{FDM}_{M-1} \ge 0,98 \implies$ Mantém **`0,97`** (1ª ocorrência).
3. Se $\text{FDM} = 0,97$ e $\text{FDM}_{M-1} = 0,97$ e $\text{FDM}_{M-2} > 0,97 \implies$ Degrada para **`0,96`** (glosa de 4,0%).
4. Se $\text{FDM} = 0,97$ e $\text{FDM}_{M-1} = 0,97$ e $\text{FDM}_{M-2} \le 0,97 \implies$ Degrada para **`0,95`** (glosa de 5,0%).
5. Se $\text{FDM} = 0,97$ e $\text{FDM}_{M-1} \le 0,96 \implies$ Mantém no piso contratual de **`0,95`** (glosa de 5,0% e emissão de notificação formal de descumprimento de contrato).

---

## 7. Apuração Prática e Valores Reais do Ciclo Vigente

A base analítica de 583 linhas consolidada na consulta [`Painel T2M`](../Codes/CodigosM/Painel%20T2M.m) gerou os seguintes valores exatos de faturamento e qualidade:

### 7.1 Tabela Consolidada por Contrato e Serviço (`TabelaFDMPorContrato`)

| Contrato | Linha de Serviço | Ref. FDM | Total Concluído (CO) | Isentos / Duplicatas | TSPE (Elegíveis) | TSPR (No Prazo) | Desvios (Fora do Prazo) | IAPFARQ Calculado | Meta Contratual | FDM Bruto | FDM Final | Glosa Aplicada na Medição |
| :--- | :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **IRON-LT1** | **Serviço 1:** Guarda | **1** | 271 | 40 | **231** | **75** | 156 | **32,47%** | 99,00% | **0,9700** | **0,9700** | Retenção de 3,0% |
| **IRON-LT1** | **Serviço 2:** Gerenciamento | **2** | 140 | 44 | **96** | **71** | 25 | **73,96%** | 99,00% | **0,9700** | **0,9700** | Retenção de 3,0% |
| **PA-LT2** | **Serviço 1:** Guarda | **1** | 97 | 14 | **83** | **49** | 34 | **59,04%** | 99,00% | **0,9700** | **0,9700** | Retenção de 3,0% |
| **PA-LT2** | **Serviço 2:** Gerenciamento | **2** | 65 | 14 | **51** | **44** | 7 | **86,27%** | 99,00% | **0,9700** | **0,9700** | Retenção de 3,0% |
| **TOTAL** | **Consolidado Geral** | — | **582** | **121** | **461** | **239** | **222** | **51,84%** | 99,00% | **0,9700** | **0,9700** | **Glosa Geral de 3,0%** |

### 7.2 Diagnóstico e Ranking dos Ofensores de SLA

A abertura dos dados demonstra quais atividades operacionais causaram a retenção de faturamento:

#### A. Contrato IRON-LT1 (Lote 1 - RJ / SP / ES):
* **Pesquisa de Item Avulso Normal (Ofensor Crítico):** 61 chamados elegíveis, sendo **59 atendidos fora do prazo** e apenas 2 no prazo (**3,28% de pontualidade**).
* **Coleta de Embalagem:** 79 chamados elegíveis, sendo 48 fora do prazo (**39,24% de pontualidade**).
* **Entrega de Embalagem Normal:** 52 chamados elegíveis, sendo 32 fora do prazo (**38,46% de pontualidade**).
* **Digitalização de Documento Administrativo:** 44 chamados elegíveis, sendo 25 fora do prazo (**43,18% de pontualidade**).
* **Destaques Positivos:** Indexação de Documento (**100,00% no prazo** com 51 atendimentos) e Entrega de Embalagem Expresso (**100,00% no prazo** com 2 atendimentos).

#### B. Contrato PA-LT2 (Lote 2 - Bahia):
* **Entrega de Embalagem Normal e Material p/ Arquivamento:** Ambas fecharam com **0,00% de pontualidade** (todas as solicitações entregues em atraso).
* **Coleta de Embalagem:** 49 chamados elegíveis, sendo 16 em atraso (**67,35% de pontualidade**).
* **Pesquisa de Item Avulso Normal:** 27 chamados elegíveis, sendo 11 em atraso (**59,26% de pontualidade**).
* **Destaques Positivos:** Indexação de Documento (**100,00% no prazo** com 26 atendimentos) e Catalogação de Item Avulso (**100,00% no prazo** com 2 atendimentos).

---

## 8. Aplicação na Memória de Cálculo (`MC`)

A aplicação contábil e fiscal do FDM ocorre na aba `MC` dos cadernos oficiais gerados em [`MEDIÇÃO/`](../MEDIÇÃO/):

$$\text{Faturamento Líquido de Serviços} = \left(\sum \text{PPU}_{\text{Guarda}} \times QExec\right) \times FDM_1 + \left(\sum \text{PPU}_{\text{Gerenciamento}} \times QExec\right) \times FDM_2$$

> [!IMPORTANT]
> A armazenagem física mensal (`ARM`) **NÃO sofre aplicação de FDM**, pois refere-se à custódia e guarda estática de caixas, fitas e tubos em galpão, sendo faturada integralmente a 100% do PPU contratual.

---

## 9. Relação de Arquivos de Código Responsáveis

| Arquivo / Script | Linguagem | Finalidade no Projeto |
| :--- | :---: | :--- |
| [`Codes/CodigosM/fnCalcularIAPFARQ.m`](../Codes/CodigosM/fnCalcularIAPFARQ.m) | Power Query M | Função atômica que implementa a fórmula percentual do IAPFARQ com tratamento de divisões por zero e arredondamento. |
| [`Codes/CodigosM/fnCalcularTSPE_TSPR.m`](../Codes/CodigosM/fnCalcularTSPE_TSPR.m) | Power Query M | Agregação vetorizada das linhas concluídas por contrato e por `Ref. FDM` (1 e 2). |
| [`Codes/CodigosM/fnClassificarFDM.m`](../Codes/CodigosM/fnClassificarFDM.m) | Power Query M | Classifica as ordens em Normal/Expresso/Isento/FDM definido e avalia conformidade de SLA (`NP`/`FP`). |
| [`Codes/CodigosM/fnCalcularPrazoSLA.m`](../Codes/CodigosM/fnCalcularPrazoSLA.m) | Power Query M | Motor de cálculo de horas úteis, expediente comercial e integração com calendário multilocacional de feriados. |
| [`Codes/CodigosM/fnCalcularFDMMensal.m`](../Codes/CodigosM/fnCalcularFDMMensal.m) | Power Query M | Enquadra o percentual IAPFARQ nas faixas da tabela contratual (0,97 a 1,00). |
| [`Codes/CodigosM/fnRecalcularFDM.m`](../Codes/CodigosM/fnRecalcularFDM.m) | Power Query M | Avalia reincidência histórica de até 3 ciclos consecutivos e aplica degradação em cascata (0,96 e 0,95). |
| [`Codes/CodigosM/TabelaFDMPorContrato.m`](../Codes/CodigosM/TabelaFDMPorContrato.m) | Power Query M | Orquestração central que consolida os indicadores por contrato para exibição no painel e cadernos de saída. |
| [`Codes/VBA/FuncoesAntigas/03fncEditalGuarda.bas`](../Codes/VBA/FuncoesAntigas/03fncEditalGuarda.bas) | VBA | Módulo legado de referência com as rotinas originais `calcular_TSPE_TSPR`, `calcular_IAPFARQ` e `calcular_FDM`. |
| [`MEDIÇÃO/Demonstrativo_Calculo_FDM_TSPE_IAPFARQ.xlsx`](../MEDIÇÃO/Demonstrativo_Calculo_FDM_TSPE_IAPFARQ.xlsx) | Excel `.xlsx` | Planilha limpa e oficial demonstrando os valores detalhados de cada cálculo em 3 abas estruturadas. |
