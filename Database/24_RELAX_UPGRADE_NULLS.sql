-- ============================================================================
-- GBS Inventory - Migration 24
-- Corrige ORA-01400 em PACK_UPGRADE.PROC_INSERT_UPGRADE: RAM_ANTERIOR_GB,
-- RAM_NOVA_GB, STORAGE_ANTERIOR_GB e STORAGE_NOVO_GB sao NOT NULL, mas o par
-- que nao se aplica ao tipo de componente escolhido (BATTERY, SCREEN,
-- KEYBOARD, COVER, GPU, OTHER) sempre foi NULL na chamada — clsGravacaoUpgrade.vb
-- e UpgradeRepository.cs (GBS_Web) ja mandam NULL corretamente pro par nao
-- aplicavel; o problema e so a constraint da coluna.
--
-- Uma tentativa anterior de corrigir isso na camada de chamada (mandando 0 em
-- vez de NULL) causou corrupcao silenciosa: o cascade de PROC_INSERT_UPGRADE
-- pra TBL_EQUIPAMENTO usa "IS NOT NULL" pra decidir se atualiza RAM_GB/
-- STORAGE_GB, e 0 satisfaz esse teste — zerando o campo que nao deveria ser
-- tocado. Essa tentativa foi revertida; o fix correto e no schema.
--
-- Idempotente: MODIFY para NULL nao gera erro se a coluna ja permitir NULL.
-- ============================================================================

ALTER TABLE TBL_EQUIPAMENTO_UPGRADE
    MODIFY (RAM_ANTERIOR_GB     NULL,
            RAM_NOVA_GB         NULL,
            STORAGE_ANTERIOR_GB NULL,
            STORAGE_NOVO_GB     NULL);

-- ----------------------------------------------------------------------------
-- Verificacao final
-- ----------------------------------------------------------------------------
SELECT COLUMN_NAME, NULLABLE
  FROM USER_TAB_COLUMNS
 WHERE TABLE_NAME = 'TBL_EQUIPAMENTO_UPGRADE'
   AND COLUMN_NAME IN ('RAM_ANTERIOR_GB','RAM_NOVA_GB','STORAGE_ANTERIOR_GB','STORAGE_NOVO_GB')
 ORDER BY COLUMN_NAME;
