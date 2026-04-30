-- ============================================================================
-- GBS Inventory - Limpeza de dados de teste
-- Mantem estrutura (tabelas, packages, sequences, constraints)
-- Apaga apenas os dados
-- Execute como GBS_OWNER no SQL Developer com F5
-- ============================================================================

-- ----------------------------------------------------------------------------
-- PASSO 1: Desabilitar constraints de FK temporariamente
-- para permitir DELETE em qualquer ordem
-- ----------------------------------------------------------------------------
ALTER TABLE TBL_INVOICE_ITEM    DISABLE CONSTRAINT FK_INV_ITEM_INVOICE;
ALTER TABLE TBL_INVOICE         DISABLE CONSTRAINT FK_INVOICE_CLIENTE;
ALTER TABLE TBL_REMESSA_ITEM    DISABLE CONSTRAINT FK_REM_ITEM_REM;
ALTER TABLE TBL_EQUIPAMENTO_UPGRADE DISABLE CONSTRAINT FK_UPGRADE_EQUIP;

-- ----------------------------------------------------------------------------
-- PASSO 2: Apagar dados de todas as tabelas (ordem segura)
-- ----------------------------------------------------------------------------
DELETE FROM TBL_INVOICE_ITEM;
DELETE FROM TBL_INVOICE;
DELETE FROM TBL_REMESSA_ITEM;
DELETE FROM TBL_REMESSA;
DELETE FROM TBL_EQUIPAMENTO_UPGRADE;
DELETE FROM TBL_EQUIPAMENTO;
DELETE FROM TBL_CLIENTE;
DELETE FROM TBL_IMPORTACAO_LOG;
DELETE FROM TBL_WHATSAPP_LOG;

COMMIT;

-- ----------------------------------------------------------------------------
-- PASSO 3: Reabilitar constraints de FK
-- ----------------------------------------------------------------------------
ALTER TABLE TBL_INVOICE_ITEM    ENABLE CONSTRAINT FK_INV_ITEM_INVOICE;
ALTER TABLE TBL_INVOICE         ENABLE CONSTRAINT FK_INVOICE_CLIENTE;
ALTER TABLE TBL_REMESSA_ITEM    ENABLE CONSTRAINT FK_REM_ITEM_REM;
ALTER TABLE TBL_EQUIPAMENTO_UPGRADE ENABLE CONSTRAINT FK_UPGRADE_EQUIP;

-- ----------------------------------------------------------------------------
-- PASSO 4: Resetar sequences para começar do 1
-- ----------------------------------------------------------------------------
DROP SEQUENCE SEQ_CLIENTE;
CREATE SEQUENCE SEQ_CLIENTE     START WITH 1 INCREMENT BY 1 NOCACHE NOCYCLE;

DROP SEQUENCE SEQ_INVOICE;
CREATE SEQUENCE SEQ_INVOICE     START WITH 1 INCREMENT BY 1 NOCACHE NOCYCLE;

DROP SEQUENCE SEQ_INVOICE_ITEM;
CREATE SEQUENCE SEQ_INVOICE_ITEM START WITH 1 INCREMENT BY 1 NOCACHE NOCYCLE;

DROP SEQUENCE SEQ_REMESSA;
CREATE SEQUENCE SEQ_REMESSA     START WITH 1 INCREMENT BY 1 NOCACHE NOCYCLE;

DROP SEQUENCE SEQ_REMESSA_ITEM;
CREATE SEQUENCE SEQ_REMESSA_ITEM START WITH 1 INCREMENT BY 1 NOCACHE NOCYCLE;

DROP SEQUENCE SEQ_EQUIPAMENTO;
CREATE SEQUENCE SEQ_EQUIPAMENTO  START WITH 1 INCREMENT BY 1 NOCACHE NOCYCLE;

-- ----------------------------------------------------------------------------
-- PASSO 5: Verificacao final - todas as tabelas devem ter 0 linhas
-- ----------------------------------------------------------------------------
SELECT 'TBL_CLIENTE'             AS TABELA, COUNT(*) AS TOTAL FROM TBL_CLIENTE            UNION ALL
SELECT 'TBL_EQUIPAMENTO'         AS TABELA, COUNT(*) AS TOTAL FROM TBL_EQUIPAMENTO        UNION ALL
SELECT 'TBL_EQUIPAMENTO_UPGRADE' AS TABELA, COUNT(*) AS TOTAL FROM TBL_EQUIPAMENTO_UPGRADE UNION ALL
SELECT 'TBL_INVOICE'             AS TABELA, COUNT(*) AS TOTAL FROM TBL_INVOICE            UNION ALL
SELECT 'TBL_INVOICE_ITEM'        AS TABELA, COUNT(*) AS TOTAL FROM TBL_INVOICE_ITEM       UNION ALL
SELECT 'TBL_REMESSA'             AS TABELA, COUNT(*) AS TOTAL FROM TBL_REMESSA            UNION ALL
SELECT 'TBL_REMESSA_ITEM'        AS TABELA, COUNT(*) AS TOTAL FROM TBL_REMESSA_ITEM       UNION ALL
SELECT 'TBL_IMPORTACAO_LOG'      AS TABELA, COUNT(*) AS TOTAL FROM TBL_IMPORTACAO_LOG     UNION ALL
SELECT 'TBL_WHATSAPP_LOG'        AS TABELA, COUNT(*) AS TOTAL FROM TBL_WHATSAPP_LOG
ORDER BY 1;
