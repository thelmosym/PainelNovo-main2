# Códigos Power Query M (Pipeline ETL)

## Overview
Conjunto de 83 consultas e funções desenvolvidas na linguagem Power Query M que materializam o pipeline central de Extração, Transformação e Carga (ETL) da medição contratual da Petrobras. Realiza a ingestão de planilhas individuais, consolidação das ordens concluídas, enriquecimento dimensional com a Tabela B, cálculo de Quantidade Executada (QExec), apuração dos prazos de SLA, apuração de FDM, faturamento de frete e memória de armazenagem (ARM).

## Interface
- `[Painel T2M.m:Painel T2M]` - Pipeline principal de consolidação, relacionamentos dimensionais, cálculo de SLA/FDM, QExec e deduplicação de KM
- `[Painel otimizado.m:Painel otimizado]` - Variante enxuta de alta performance da esteira de transformação de atendimentos
- `[DADOS.m:DADOS]` - Tabela mãe operacional enriquecida com colunas de auditoria manual (`Agrupamento`, `FA`, `Obs. Isenção`) e bufferizada em RAM
- `[ARM.m:ARM]` - Consolidação de armazenagem unindo dados de ES, RJ, SP e PA com apuração mensal de QUA
- `[FRETE.m:FRETE]` - Filtro especializado de serviços de frete normal (`FRE-NRM`) e frete expresso (`FRE-EXP`)
- `[TabelaFDMPorContrato.m:TabelaFDMPorContrato]` - Matriz de apuração dos indicadores mensais TSPE, TSPR, IAPFARQ e FDM por contrato
- `[fnCalcularDataSLA.m]` / `[fnCalcularPrazoSLA.m]` - Funções de apuração de cumprimento de SLA em horário comercial útil
- `[fnCalcularTSPE_TSPR.m]` / `[fnCalcularIAPFARQ.m]` / `[fnRecalcularFDM.m]` - Funções matemáticas de apuração dos índices contratuais de FDM
- `[Contrato.m]`, `[Centro.m]`, `[Origem_Destino.m]`, `[PPU.m]` - Consultas de referência dimensional carregadas dinamicamente via `localTabelaB`
- Consultas de particionamento por contrato e UF: `DADOS_IRON-LT1-*.m`, `ARM_IRON-LT1-*.m`, `FRETE_IRON-LT1-*.m`, `DADOS_PA-LT2.m`, `ARM_PA-LT2.m`

## Invariants
- `[Painel T2M.m:FiltrarSituacaoEContrato]` Apenas ordens com `[Situação] = "CO"` e pertencentes aos contratos `IRON-LT1` ou `PA-LT2` integram o faturamento.
- `[Painel T2M.m:FiltrarPeriodoValido]` O corte temporal das ordens obedece rigidamente ao ciclo mensal contratual (dia 26 do mês anterior ao dia 25 do mês de referência).
- `[ARM.m:QUA]` O saldo final de armazenagem obedece estritamente à fórmula: $\text{QUA}_m = \text{QUA}_{m-1} - (\text{QB}_m + \text{QD}_m) + \text{QN}_m$, com tratamento compulsório de nulos para zero.
- `[FRETE.m:Linhas Filtradas]` Apenas linhas com serviço `FRE-EXP` ou `FRE-NRM` integram as tabelas de faturamento de transporte.
- `[Painel T2M.m:AdicionarKMAdicional]` Deduz a franquia padrão lida do parâmetro escalar `valKmAdicional` sobre a distância total da rota.
- `[TabelaFDMPorContrato.m:ComFDM1Final]` Aplica fallback seguro `try fnRecalcularFDM(...) otherwise [FDM1_Bruto]` quando períodos históricos estão indisponíveis.

## Side Effects
- `[Arquivo de Amostra.m:Fonte]` Lê recursivamente o sistema de arquivos no diretório apontado pelo parâmetro `localMonitIndividualT2M`.
- `[Centro.m:Fonte]` / `[Contrato.m:ArquivoTabelaB]` Realiza leitura de pastas de trabalho externas no disco através de `File.Contents`.
- `[DADOS.m:ResultadoFinalBuffer]` / `[ARM.m:ResultadoFinal]` Aloca e fixa estruturas de dados na memória RAM do processo via `Table.Buffer` e `Binary.Buffer`.

## Verification
- Run: `<test-command>`
