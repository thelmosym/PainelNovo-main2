# Findings & Decisions

## Requirements
- Fazer análise completa do projeto.
- Elaborar posteriormente um dashboard no arquivo `Painel de controle MemoriaPetrobras V.8.6.xlsm`.
- Criar backup da planilha original antes de qualquer alteração.
- Preservar compatibilidade de VBA/Power Query conforme as regras da skill.

## Research Findings
- O arquivo original está na raiz do repositório e possui aproximadamente 1,27 MB.
- Backup criado em `Backups\Painel de controle MemoriaPetrobras V.8.6_original_20261110_003012.xlsm`.
- SHA-256 do original e do backup: `1911360CCA5953984D02F520AB0067807500B3A147FC834D06CA916D844FA405`; hashes coincidem.
- A raiz contém diretórios de código (`Codes`), documentação (`Docs`), medições, monitoramento, tabelas específicas, scripts e backups existentes.
- O README descreve uma arquitetura híbrida: Power Query para ingestão/ETL/cálculos e VBA para atualização, auditoria, exportação e dashboard.
- O fluxo de negócio documentado usa ciclo de medição do dia 26 ao dia 25, filtra situação `CO`, calcula SLA/FDM/QExec/frete e exporta cadernos por contrato.
- O inventário encontrou 40 módulos `.bas`, 10 classes `.cls`, 85 consultas/funções `.m`, 5 arquivos `.vba` e múltiplos scripts/documentos; a inspeção automática do workbook está disponível em `scripts/inspect_workbook.py`.
- O workbook contém 31 abas, incluindo `Painel`, `Painel T2M`, `ARM`, `DADOS`, `Frete`, `LOG_CRITICAS`, `TabelaFDMPorContrato` e uma aba `DASHBOARD` atualmente oculta.
- Foram detectadas 44 tabelas estruturadas, 87 conexões Power Query/OLEDB e projeto VBA embutido (`xl/vbaProject.bin`); o pacote Mashup também foi detectado.
- Já existe código de dashboard em `Codes\VBA\Funções Novas\modDashboard.bas` e uma cópia em `Backup_PreRefatoracao`, indicando trabalho anterior que precisa ser auditado antes de reutilização.
- O `modDashboard.bas` atual está desalinhado com o workbook: procura a aba `01_DASHBOARD`, mas a aba real é `DASHBOARD` (oculta); também usa `.Activate`/`.Select` e `On Error Resume Next` em `RefreshAll`, contrariando os padrões de robustez definidos.
- O módulo de atualização de consultas percorre 87 conexões, desativa `BackgroundQuery` apenas para OLEDB, registra status em `Painel!F9` e restaura estados da aplicação; há tratamento tolerante de erros por conexão.
- O módulo de auditoria usa `DADOS`/`Painel_T2M`, mapeamento dinâmico de colunas, processamento em memória e gravação em `LOG_CRITICAS`, sendo uma fonte importante para os indicadores de qualidade do dashboard.
- A consulta `DADOS.m` produz a base operacional final com contrato, origem/destino, quantidades, FDM, FC, QExec e campos manuais; `TabelaFDMPorContrato.m` consolida TSPE/TSPR, IAPFARQ e FDM por contrato com fallback quando não há histórico válido.
- Os arquivos VBA visualizados exibem mojibake em acentos (`M�DULO`, `Atualiza��o`), evidência de possível corrupção de CP1252/UTF-8 que deve ser corrigida somente nos arquivos de código e validada antes de importação.
- A aba `DASHBOARD` já existe no workbook, está oculta, ocupa `A1:N40` e possui 10 desenhos/controles, mas não há gráficos Excel, tabelas dinâmicas ou segmentações; portanto o trabalho parece ser de correção/integração da dashboard existente, não de criação de uma aba do zero.
- O `modDashboard.bas` em `Backup_PreRefatoracao` está mais completo: constrói cabeçalho, ações, filtros, cards, resumo por contrato, alertas e rodapé. Porém contém dependências a nomes de cabeçalhos possivelmente inexistentes (`Data SLA`/`Prazo SLA`) e usa `ActiveWindow`, `Activate` e limpeza destrutiva da aba.
- O workbook tem desenhos na aba `Painel` e na aba `DASHBOARD`, mas nenhum gráfico nativo; os indicadores atuais são provavelmente formas/células atualizadas por VBA.
- O script `scripts\inspect_dashboard.py` está desatualizado: usa o nome legado `Painel de controle MemoriaPetrobras V6.xlsm` e falha no arquivo atual; o script de verificação também referencia abas inexistentes (`01_DASHBOARD`, `04_CALCULOS`, `00_GUIA`).
- A leitura COM inicial mostrou `DASHBOARD` oculta, período Setembro/2026, 583 registros e valores zerados para quantidade/QExec/KM/SLA; a coleta foi interrompida apenas por erro de formatação do diagnóstico ao enumerar shapes, sem escrita no arquivo.
- A validação COM corrigida confirmou `DASHBOARD` oculta, `A1:N40`, 5 shapes de botões e 1 gráfico; os botões apontam para `AtualizarTodasConsultasPowerQuery`, `GerarArquivosPorContrato`, `AtualizarDashboard`, `AbrirPastaMedicao` e `MostrarResumoDashboard`.
- Os cards atuais mostram 583 registros, mas `Qtd. Solicitada`, `Qtd. Atendida`, `QExec`, `Fora do Prazo` estão em zero; não foram encontrados erros de fórmula por `SpecialCells`. Isso indica falha de ligação/cálculo, não erro de fórmula Excel.
- A aba `DADOS` tem cabeçalhos na linha 2, com quebras de linha em nomes como `Qtd.\nSolicitada` e `Qtd.\nAtendida`; a macro de dashboard que busca `Qtd. Solicitada`/`Qtd. Atendida` sem quebra de linha tende a não localizar essas colunas.
- O período exibido na dashboard é Setembro/2026, enquanto a documentação e a base revisada registram cenários de Agosto/2026; o dashboard precisa derivar o período dos parâmetros atuais, sem valores hardcoded.
- A validação final confirmou que o SHA-256 do original continua igual ao do backup; o workbook não foi salvo nem alterado.

## Technical Decisions
| Decision | Rationale |
|----------|-----------|

## Issues Encountered
| Issue | Resolution |
|-------|------------|

## Resources
- `Painel de controle MemoriaPetrobras V.8.6.xlsm`
- `Backups\Painel de controle MemoriaPetrobras V.8.6_original_20261110_003012.xlsm`
