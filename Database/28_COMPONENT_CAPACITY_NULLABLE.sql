-- ============================================================================
-- GBS Inventory - Migration 28
-- CAPACITY_GB (RAM para MINI_DESKTOP) passa a aceitar NULL. STORAGE_GB ja e
-- nullable (migration 21). PACK_COMPONENT.PROC_INSERT ja repassa NULL direto,
-- nao precisa recompilar. Idempotente.
-- ============================================================================

DECLARE V_NULLABLE VARCHAR2(1);
BEGIN
    SELECT NULLABLE INTO V_NULLABLE FROM USER_TAB_COLUMNS
     WHERE TABLE_NAME = 'TBL_COMPONENT' AND COLUMN_NAME = 'CAPACITY_GB';
    IF V_NULLABLE = 'N' THEN
        EXECUTE IMMEDIATE 'ALTER TABLE TBL_COMPONENT MODIFY (CAPACITY_GB NULL)';
        DBMS_OUTPUT.PUT_LINE('OK: CAPACITY_GB agora aceita NULL');
    ELSE
        DBMS_OUTPUT.PUT_LINE('SKIP: CAPACITY_GB ja aceita NULL');
    END IF;
END;
/

-- ----------------------------------------------------------------------------
-- Verificacao final
-- ----------------------------------------------------------------------------
SELECT COLUMN_NAME, NULLABLE
  FROM USER_TAB_COLUMNS
 WHERE TABLE_NAME = 'TBL_COMPONENT' AND COLUMN_NAME IN ('CAPACITY_GB', 'STORAGE_GB');

SELECT OBJECT_NAME, OBJECT_TYPE, STATUS
  FROM USER_OBJECTS
 WHERE OBJECT_NAME IN ('PACK_COMPONENT', 'PACK_UPGRADE')
 ORDER BY OBJECT_NAME, OBJECT_TYPE;
