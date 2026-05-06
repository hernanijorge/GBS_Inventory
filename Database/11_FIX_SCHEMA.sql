-- ============================================================================
-- GBS Inventory - Migração: alinha colunas TBL_EQUIPAMENTO com PACK_EQUIPAMENTO
-- Execute como GBS_OWNER antes de recompilar os packages.
-- Idempotente: verifica existência antes de cada ALTER.
-- ============================================================================

-- ----------------------------------------------------------------------------
-- 1. MODELO → MODEL
-- ----------------------------------------------------------------------------
DECLARE V_N NUMBER;
BEGIN
    SELECT COUNT(*) INTO V_N FROM USER_TAB_COLUMNS
     WHERE TABLE_NAME = 'TBL_EQUIPAMENTO' AND COLUMN_NAME = 'MODELO';
    IF V_N > 0 THEN
        EXECUTE IMMEDIATE 'ALTER TABLE TBL_EQUIPAMENTO RENAME COLUMN MODELO TO MODEL';
        DBMS_OUTPUT.PUT_LINE('OK: MODELO → MODEL');
    ELSE
        DBMS_OUTPUT.PUT_LINE('SKIP: MODEL já existe');
    END IF;
END;
/

-- ----------------------------------------------------------------------------
-- 2. CPU_MODEL → PROCESSADOR
-- ----------------------------------------------------------------------------
DECLARE V_N NUMBER;
BEGIN
    SELECT COUNT(*) INTO V_N FROM USER_TAB_COLUMNS
     WHERE TABLE_NAME = 'TBL_EQUIPAMENTO' AND COLUMN_NAME = 'CPU_MODEL';
    IF V_N > 0 THEN
        EXECUTE IMMEDIATE 'ALTER TABLE TBL_EQUIPAMENTO RENAME COLUMN CPU_MODEL TO PROCESSADOR';
        DBMS_OUTPUT.PUT_LINE('OK: CPU_MODEL → PROCESSADOR');
    ELSE
        DBMS_OUTPUT.PUT_LINE('SKIP: PROCESSADOR já existe');
    END IF;
END;
/

-- ----------------------------------------------------------------------------
-- 3. NOTES → OBSERVACAO
-- ----------------------------------------------------------------------------
DECLARE V_N NUMBER;
BEGIN
    SELECT COUNT(*) INTO V_N FROM USER_TAB_COLUMNS
     WHERE TABLE_NAME = 'TBL_EQUIPAMENTO' AND COLUMN_NAME = 'NOTES';
    IF V_N > 0 THEN
        EXECUTE IMMEDIATE 'ALTER TABLE TBL_EQUIPAMENTO RENAME COLUMN NOTES TO OBSERVACAO';
        DBMS_OUTPUT.PUT_LINE('OK: NOTES → OBSERVACAO');
    ELSE
        DBMS_OUTPUT.PUT_LINE('SKIP: OBSERVACAO já existe');
    END IF;
END;
/

-- ----------------------------------------------------------------------------
-- 4. DATA_ALTERACAO → DATA_ATUALIZACAO
-- ----------------------------------------------------------------------------
DECLARE V_N NUMBER;
BEGIN
    SELECT COUNT(*) INTO V_N FROM USER_TAB_COLUMNS
     WHERE TABLE_NAME = 'TBL_EQUIPAMENTO' AND COLUMN_NAME = 'DATA_ALTERACAO';
    IF V_N > 0 THEN
        EXECUTE IMMEDIATE 'ALTER TABLE TBL_EQUIPAMENTO RENAME COLUMN DATA_ALTERACAO TO DATA_ATUALIZACAO';
        DBMS_OUTPUT.PUT_LINE('OK: DATA_ALTERACAO → DATA_ATUALIZACAO');
    ELSE
        DBMS_OUTPUT.PUT_LINE('SKIP: DATA_ATUALIZACAO já existe');
    END IF;
END;
/

-- ----------------------------------------------------------------------------
-- 5. RAM_GB VARCHAR2 → NUMBER
--    Remove sufixos ("8GB" → 8) e converte o tipo.
-- ----------------------------------------------------------------------------
DECLARE V_TYPE VARCHAR2(30);
BEGIN
    SELECT DATA_TYPE INTO V_TYPE FROM USER_TAB_COLUMNS
     WHERE TABLE_NAME = 'TBL_EQUIPAMENTO' AND COLUMN_NAME = 'RAM_GB';

    IF V_TYPE = 'VARCHAR2' THEN
        -- Extrai a parte numérica (ex: "8 GB" → "8", "16" → "16")
        EXECUTE IMMEDIATE
            'UPDATE TBL_EQUIPAMENTO
                SET RAM_GB = REGEXP_SUBSTR(RAM_GB, ''\d+'')
              WHERE RAM_GB IS NOT NULL';
        -- NULL em valores que ficaram sem dígitos
        EXECUTE IMMEDIATE
            'UPDATE TBL_EQUIPAMENTO
                SET RAM_GB = NULL
              WHERE TRIM(RAM_GB) IS NULL';
        COMMIT;
        EXECUTE IMMEDIATE 'ALTER TABLE TBL_EQUIPAMENTO MODIFY (RAM_GB NUMBER)';
        DBMS_OUTPUT.PUT_LINE('OK: RAM_GB VARCHAR2 → NUMBER');
    ELSE
        DBMS_OUTPUT.PUT_LINE('SKIP: RAM_GB já é ' || V_TYPE);
    END IF;
END;
/

-- ----------------------------------------------------------------------------
-- 6. STORAGE_GB VARCHAR2 → NUMBER
--    Mesma lógica de RAM_GB.
-- ----------------------------------------------------------------------------
DECLARE V_TYPE VARCHAR2(30);
BEGIN
    SELECT DATA_TYPE INTO V_TYPE FROM USER_TAB_COLUMNS
     WHERE TABLE_NAME = 'TBL_EQUIPAMENTO' AND COLUMN_NAME = 'STORAGE_GB';

    IF V_TYPE = 'VARCHAR2' THEN
        EXECUTE IMMEDIATE
            'UPDATE TBL_EQUIPAMENTO
                SET STORAGE_GB = REGEXP_SUBSTR(STORAGE_GB, ''\d+'')
              WHERE STORAGE_GB IS NOT NULL';
        EXECUTE IMMEDIATE
            'UPDATE TBL_EQUIPAMENTO
                SET STORAGE_GB = NULL
              WHERE TRIM(STORAGE_GB) IS NULL';
        COMMIT;
        EXECUTE IMMEDIATE 'ALTER TABLE TBL_EQUIPAMENTO MODIFY (STORAGE_GB NUMBER)';
        DBMS_OUTPUT.PUT_LINE('OK: STORAGE_GB VARCHAR2 → NUMBER');
    ELSE
        DBMS_OUTPUT.PUT_LINE('SKIP: STORAGE_GB já é ' || V_TYPE);
    END IF;
END;
/

-- ----------------------------------------------------------------------------
-- 7. Recompila PACK_EQUIPAMENTO (spec + body)
-- ----------------------------------------------------------------------------
ALTER PACKAGE PACK_EQUIPAMENTO COMPILE;
ALTER PACKAGE PACK_EQUIPAMENTO COMPILE BODY;

-- ----------------------------------------------------------------------------
-- 8. Verificação final
-- ----------------------------------------------------------------------------
SELECT OBJECT_TYPE, STATUS, LAST_DDL_TIME
  FROM USER_OBJECTS
 WHERE OBJECT_NAME = 'PACK_EQUIPAMENTO'
 ORDER BY OBJECT_TYPE;

-- Erros de compilação (deve retornar 0 linhas se tudo OK)
SELECT LINE, POSITION, TEXT
  FROM USER_ERRORS
 WHERE NAME = 'PACK_EQUIPAMENTO'
 ORDER BY SEQUENCE;
