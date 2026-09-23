-- ============================================================================
-- GBS Inventory - Migration 31
-- TBL_CLIENT_ALIAS: mapeia variacoes do destinatario (TBL_REMESSA.DESTINATARIO)
-- para um nome canonico, usado no "Report by Recipient" (modo Custom mapping).
-- ALIAS_NAME / CANONICAL_NAME sao gravados normalizados (UPPER(TRIM())) pela app.
-- Idempotente.
-- ============================================================================

DECLARE V_N NUMBER;
BEGIN
    SELECT COUNT(*) INTO V_N FROM USER_TABLES WHERE TABLE_NAME = 'TBL_CLIENT_ALIAS';
    IF V_N = 0 THEN
        EXECUTE IMMEDIATE
            'CREATE TABLE TBL_CLIENT_ALIAS (' ||
            '    ID_ALIAS        NUMBER(10)     NOT NULL,' ||
            '    CANONICAL_NAME  VARCHAR2(200)  NOT NULL,' ||
            '    ALIAS_NAME      VARCHAR2(200)  NOT NULL,' ||
            '    CREATED_AT      DATE           DEFAULT SYSDATE,' ||
            '    CREATED_BY      VARCHAR2(100),' ||
            '    CONSTRAINT PK_CLIENT_ALIAS PRIMARY KEY (ID_ALIAS),' ||
            '    CONSTRAINT UQ_CLIENT_ALIAS UNIQUE (ALIAS_NAME)' ||
            ')';
        DBMS_OUTPUT.PUT_LINE('OK: TBL_CLIENT_ALIAS criada');
    ELSE
        DBMS_OUTPUT.PUT_LINE('SKIP: TBL_CLIENT_ALIAS ja existe');
    END IF;

    SELECT COUNT(*) INTO V_N FROM USER_SEQUENCES WHERE SEQUENCE_NAME = 'SEQ_CLIENT_ALIAS';
    IF V_N = 0 THEN
        EXECUTE IMMEDIATE 'CREATE SEQUENCE SEQ_CLIENT_ALIAS START WITH 1 INCREMENT BY 1 NOCACHE NOCYCLE';
        DBMS_OUTPUT.PUT_LINE('OK: SEQ_CLIENT_ALIAS criada');
    ELSE
        DBMS_OUTPUT.PUT_LINE('SKIP: SEQ_CLIENT_ALIAS ja existe');
    END IF;
END;
/

-- ----------------------------------------------------------------------------
-- Verificacao final
-- ----------------------------------------------------------------------------
SELECT COLUMN_NAME, DATA_TYPE, DATA_LENGTH, NULLABLE
  FROM USER_TAB_COLUMNS
 WHERE TABLE_NAME = 'TBL_CLIENT_ALIAS'
 ORDER BY COLUMN_ID;
