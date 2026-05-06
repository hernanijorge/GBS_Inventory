-- ============================================================================
-- GBS Inventory - Adiciona campo BATTERY_CHECK + corrige PROC_INSERT
-- Pré-requisito: executar 11_FIX_SCHEMA.sql antes, se ainda não rodou.
-- Idempotente.
-- ============================================================================

-- ----------------------------------------------------------------------------
-- 1. Adiciona coluna BATTERY_CHECK (se não existir)
-- ----------------------------------------------------------------------------
DECLARE V_N NUMBER;
BEGIN
    SELECT COUNT(*) INTO V_N FROM USER_TAB_COLUMNS
     WHERE TABLE_NAME = 'TBL_EQUIPAMENTO' AND COLUMN_NAME = 'BATTERY_CHECK';
    IF V_N = 0 THEN
        EXECUTE IMMEDIATE 'ALTER TABLE TBL_EQUIPAMENTO ADD BATTERY_CHECK VARCHAR2(20)';
        DBMS_OUTPUT.PUT_LINE('OK: coluna BATTERY_CHECK adicionada');
    ELSE
        DBMS_OUTPUT.PUT_LINE('SKIP: BATTERY_CHECK já existe');
    END IF;
END;
/

-- ----------------------------------------------------------------------------
-- 2. Recompila spec e body (arquivos 03 e 04 devem ser executados antes disto)
-- ----------------------------------------------------------------------------
ALTER PACKAGE PACK_EQUIPAMENTO COMPILE;
ALTER PACKAGE PACK_EQUIPAMENTO COMPILE BODY;

-- ----------------------------------------------------------------------------
-- 3. Verificação final
-- ----------------------------------------------------------------------------
SELECT OBJECT_TYPE, STATUS, LAST_DDL_TIME
  FROM USER_OBJECTS
 WHERE OBJECT_NAME = 'PACK_EQUIPAMENTO'
 ORDER BY OBJECT_TYPE;

-- Deve retornar 0 linhas se compilou sem erros
SELECT LINE, POSITION, TEXT
  FROM USER_ERRORS
 WHERE NAME = 'PACK_EQUIPAMENTO'
 ORDER BY SEQUENCE;
