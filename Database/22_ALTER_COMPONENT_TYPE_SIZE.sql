-- ============================================================================
-- GBS Inventory - Migration 22
-- Corrige ORA-12899: COMPONENT_TYPE era VARCHAR2(10), mas 'MINI_DESKTOP'
-- (adicionado na migration 20) tem 12 caracteres. A CHECK constraint foi
-- atualizada na epoca, a coluna nao. Amplia para VARCHAR2(20) (mesmo
-- tamanho ja usado em INTERNAL_UID).
-- Idempotente.
-- ============================================================================

DECLARE V_LEN NUMBER;
BEGIN
    SELECT DATA_LENGTH INTO V_LEN FROM USER_TAB_COLUMNS
     WHERE TABLE_NAME = 'TBL_COMPONENT' AND COLUMN_NAME = 'COMPONENT_TYPE';
    IF V_LEN < 20 THEN
        EXECUTE IMMEDIATE 'ALTER TABLE TBL_COMPONENT MODIFY (COMPONENT_TYPE VARCHAR2(20))';
        DBMS_OUTPUT.PUT_LINE('OK: COMPONENT_TYPE ampliado para VARCHAR2(20)');
    ELSE
        DBMS_OUTPUT.PUT_LINE('SKIP: COMPONENT_TYPE ja tem tamanho >= 20');
    END IF;
END;
/

-- ----------------------------------------------------------------------------
-- Verificacao final
-- ----------------------------------------------------------------------------
SELECT column_name, data_type, data_length
  FROM user_tab_columns
 WHERE table_name = 'TBL_COMPONENT' AND column_name = 'COMPONENT_TYPE';
