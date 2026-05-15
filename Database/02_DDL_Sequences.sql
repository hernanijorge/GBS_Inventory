-- ============================================================================
-- GBS - Sequences para PK autoincrement
-- Nomes alinhados com o que existe no banco de producao.
-- Sequences de outros modulos criadas junto com seus DDLs:
--   SEQ_CLIENTE / SEQ_INVOICE / SEQ_INVOICE_ITEM  -> 07_DDL_CADASTROS_INVOICE.sql
--   SEQ_COMPONENT                                 -> 15_DDL_TBL_COMPONENT.sql
-- ============================================================================

CREATE SEQUENCE SEQ_EQUIPAMENTO
    START WITH 1 INCREMENT BY 1 NOCACHE NOCYCLE;

CREATE SEQUENCE SEQ_EQUIPAMENTO_UPGRADE
    START WITH 1 INCREMENT BY 1 NOCACHE NOCYCLE;

CREATE SEQUENCE SEQ_REMESSA
    START WITH 1 INCREMENT BY 1 NOCACHE NOCYCLE;

CREATE SEQUENCE SEQ_REMESSA_ITEM
    START WITH 1 INCREMENT BY 1 NOCACHE NOCYCLE;

CREATE SEQUENCE SEQ_IMPORTACAO_LOG
    START WITH 1 INCREMENT BY 1 NOCACHE NOCYCLE;

CREATE SEQUENCE SEQ_WHATSAPP_LOG
    START WITH 1 INCREMENT BY 1 NOCACHE NOCYCLE;

COMMIT;
