-- ============================================================================
-- GBS Inventory - Migration 35
-- TBL_CLIENTE.ALIASES VARCHAR2(500): nomes alternativos do cliente, separados
-- por ";" (gravados normalizados pela app: UPPER/TRIM, ex. "MACIR;MACIR LLC").
-- Usado no auto-link de TBL_REMESSA.ID_CLIENTE pelo DESTINATARIO (aba Shipments).
-- PACK_CLIENTE le C.* -> recompilado aqui para passar a devolver ALIASES.
-- Idempotente.
-- ============================================================================

DECLARE
    V_N NUMBER;
BEGIN
    SELECT COUNT(*) INTO V_N FROM USER_TAB_COLUMNS
     WHERE TABLE_NAME = 'TBL_CLIENTE' AND COLUMN_NAME = 'ALIASES';
    IF V_N = 0 THEN
        EXECUTE IMMEDIATE 'ALTER TABLE TBL_CLIENTE ADD (ALIASES VARCHAR2(500))';
        DBMS_OUTPUT.PUT_LINE('OK: TBL_CLIENTE.ALIASES criada');
    ELSE
        DBMS_OUTPUT.PUT_LINE('SKIP: TBL_CLIENTE.ALIASES ja existe');
    END IF;
END;
/

ALTER PACKAGE PACK_CLIENTE COMPILE
/

ALTER PACKAGE PACK_CLIENTE COMPILE BODY
/

-- ----------------------------------------------------------------------------
-- Verificacao final
-- ----------------------------------------------------------------------------
SELECT COLUMN_NAME, DATA_TYPE, DATA_LENGTH, NULLABLE
  FROM USER_TAB_COLUMNS
 WHERE TABLE_NAME = 'TBL_CLIENTE' AND COLUMN_NAME = 'ALIASES'
/
