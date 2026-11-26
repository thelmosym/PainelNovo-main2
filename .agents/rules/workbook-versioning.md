---
name: workbook-versioning
version: 1.0.0
priority: P0
trigger: always_on
description: Regra mandatória de backup prévio e incremento de versão ao alterar planilhas principais (Painel de Controle ou Monitoramento).
---

# Regra de Backup e Versionamento de Planilhas Principais (P0)

> **MANDATÓRIO (SEM EXCEÇÕES):** Sempre que houver qualquer alteração em uma das planilhas principais do projeto (**Painel de Controle** ou **Monitoramento Individual**), você DEVE executar rigorosamente o protocolo abaixo.

---

## 📂 Escopo das Planilhas Principais

1. **Painel Principal (Raiz):** `Painel de controle MemoriaPetrobras *.xlsm` (ex.: `V.8.6` → `V.8.7` / `V.9`)
2. **Monitoramento Individual (`Monitoramento/`):** `Monitoramento individual_*.xlsm` (`A4UU`, `DPBR`, `GPZ1`, `GQ6S`, `S2IJ`, etc., ex.: `V.8` → `V.9`)

---

## 🔒 Protocolo Obrigatório de Alteração (3 Etapas)

1. **Backup Prévio Obrigatório (`Backups/`):**
   - **ANTES** de aplicar qualquer modificação (módulos VBA, consultas M embarcadas, abas, tabelas, validações, formatações ou dados estruturais), copie o arquivo original intacto para a pasta `Backups/` (preservando sua versão atual, ex.: `Backups/Monitoramento individual_A4UU_V.8_backup_YYYYMMDD_HHMMSS.xlsm` ou movendo o exemplar anterior intacto para `Backups/`).
2. **Criação de Novo Arquivo com Versão Incrementada:**
   - Gere a planilha alterada em um **novo arquivo aumentando o valor da versão no final do nome** (ex.: `_V.8.xlsm` → `_V.9.xlsm`; `V.8.6.xlsm` → `V.8.7.xlsm`).
   - **Atenção ao Power Query (`Monitoramento/`):** Como o pipeline Power Query ingere todos os `.xlsm` presentes na pasta `Monitoramento/` via `Folder.Files`, **não deixe duas versões ativas da mesma frente simultaneamente dentro de `Monitoramento/`** — mantenha apenas a nova versão incrementada em `Monitoramento/` e arquive a versão anterior intacta em `Backups/`.
3. **Integridade de Botões e Macros (`Shape.OnAction`):**
   - Ao salvar com o novo número de versão, valide/atualize a propriedade `OnAction` dos botões de formulário (`Shapes`) para que apontem para o nome da nova versão do arquivo (sem referências residuais à versão anterior).
