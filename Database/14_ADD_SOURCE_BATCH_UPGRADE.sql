-- ============================================================================
-- GBS Inventory - Remove constraint CK_SOURCE de SOURCE_ORIGEM
-- Permite armazenar o SOURCE_BATCH do equipamento como origem da peça
-- Pré-requisito: nenhum (idempotente).
-- ============================================================================

DECLARE V_N NUMBER;
BEGIN
    SELECT COUNT(*) INTO V_N FROM USER_CONSTRAINTS
     WHERE TABLE_NAME      = 'TBL_EQUIPAMENTO_UPGRADE'
       AND CONSTRAINT_NAME = 'CK_SOURCE';
    IF V_N > 0 THEN
        EXECUTE IMMEDIATE 'ALTER TABLE TBL_EQUIPAMENTO_UPGRADE DROP CONSTRAINT CK_SOURCE';
        DBMS_OUTPUT.PUT_LINE('OK: CK_SOURCE removida');
    ELSE
        DBMS_OUTPUT.PUT_LINE('SKIP: CK_SOURCE nao encontrada');
    END IF;
END;
/
