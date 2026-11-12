# VBA Funções Antigas

## Overview
Biblioteca legada de rotinas e funções VBA que originariamente implementavam o pipeline de medição contratual da Petrobras. Abrange o cálculo de SLA em horário comercial útil, apuração de FDM, faturamento de frete agrupado e individual, cálculo de franquia de quilometragem e custódia de armazenagem (ARM).

## Interface
- `[Planilha7.cls:btCodigoDeAgrupamento]` - Gera códigos de agrupamento AGP para ordens de frete selecionadas
- `[Planilha7.cls:preenchimentoPrazoFDM]` - Processa linhas concluídas ('CO'), calculando prazos SLA e classificações FDM
- `[Planilha7.cls:importacaoMonitoramentoInd]` - Importa dados brutos das planilhas operacionais individuais para o ciclo mensal
- `[A_Calc_Medicao.bas:Inicial_Importar_e_calcular_dados_Atendimentos]` - Fluxo central de importação, ordenação e consolidação de atendimentos
- `[A_Calc_Medicao.bas:Organizar_dados]` - Limpa colunas de resultado e ordena os atendimentos pelas colunas S, V, K e B
- `[A_Calc_Medicao.bas:CalcularQuilometragemAdicional]` - Calcula o KM excedente à franquia contratual entre centros operacionais
- `[A_Calc_Medicao.bas:CalcularFrete]` - Agrupa atendimentos de frete por data de fechamento, rota e linha PPU com divisor de volume 10
- `[fncTabelaB.bas:pesquisar_centro]` - Consulta o Centro de custos correspondente à Localidade na base Tabela B
- `[fncTabelaB.bas:calcular_KMAdicional]` - Aplica a regra de franquia de 100 km (RAIO = 100) para transporte interestadual/intermunicipal
- `[fncTabelaB.bas:recalcular_FDM]` - Recalcula os índices de Fator de Desempenho Mensal com base nas medições anteriores
- `[fncGeral.bas:calcular_dataSLA]` - Algoritmo de tempo de atendimento em horas comerciais úteis com suporte a feriados nacionais e estaduais

## Invariants
- `[Planilha7.cls:btCodigoDeAgrupamento]` Exige no mínimo 2 células selecionadas no intervalo de linhas 3 a 2003, coluna 17 e Situação 'CO'.
- `[Planilha7.cls:importacaoMonitoramentoInd]` Considera apenas ordens com fechamento preenchido dentro do ciclo contratual (dia 26 ao 25), contratos 'IRON-LT1' ou 'PA-LT2' e Situação 'CO'.
- `[fncTabelaB.bas:calcular_KMAdicional]` Franquia estabelecida em 100 km (RAIO = 100); retorna 0 km se origem e destino pertencerem à mesma região metropolitana ou se a distância total for $\le 100$ km.
- `[fncGeral.bas:calcular_dataSLA]` Contagem estrita em horário comercial das 08:00 às 17:00, com pausa entre 12:00 e 13:00 para almoço; ignora sábados, domingos e feriados validados via `pesquisar_EFeriado`.
- `[A_Calc_Medicao.bas:CalcularFrete]` Ordens com soma de quantidade $\le 10$ geram volume de frete unitário (1); quantitativos superiores são faturados na fração inteira/decimal $\text{soma} / 10$.

## Side Effects
- `[Planilha7.cls:importacaoMonitoramentoInd]` Abre pastas de trabalho de monitoramento externas em segundo plano via `Workbooks.Open` e fecha sem salvar.
- `[Planilha7.cls:preenchimentoPrazoFDM]` Limpa intervalos `AG3:AG2003` e `R3:S2003`, alterando cores de fundo e gravando códigos SLA e diagnósticos de erro.
- `[A_Calc_Medicao.bas:Inicial_Importar_e_calcular_dados_Atendimentos]` Modifica e limpa os intervalos `Dados!B3:AF`, colando novos dados importados e salvando a pasta de trabalho via `ActiveWorkbook.Save`.
- `[A_Calc_Medicao.bas:CopiarDadosParaFretes]` Preenche e formata a planilha `Fretes` com os registros elegíveis de transporte e persiste as alterações.
- `[A_Calc_Medicao.bas:preencher_valores_MC2]` Grava preços unitários nas células `MC!L13:L52` e oculta a planilha `Tabelas`.

## Verification
- Run: `<test-command>`
