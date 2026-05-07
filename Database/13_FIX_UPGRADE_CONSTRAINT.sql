-- ============================================================================
-- GBS Inventory - Expande constraint CK_UPGRADE_TIPO
-- Permite os novos tipos de componente adicionados ao formulário:
-- BATTERY, SCREEN, KEYBOARD, COVER, GPU, OTHER
-- Pré-requisito: nenhum (idempotente).
-- ============================================================================

-- ----------------------------------------------------------------------------
-- 1. Remove a constraint antiga (apenas se existir)
-- ----------------------------------------------------------------------------
DECLARE V_N NUMBER;
BEGIN
    SELECT COUNT(*) INTO V_N FROM USER_CONSTRAINTS
     WHERE TABLE_NAME = 'TBL_EQUIPAMENTO_UPGRADE'
       AND CONSTRAINT_NAME = 'CK_UPGRADE_TIPO';
    IF V_N > 0 THEN
        EXECUTE IMMEDIATE 'ALTER TABLE TBL_EQUIPAMENTO_UPGRADE DROP CONSTRAINT CK_UPGRADE_TIPO';
        DBMS_OUTPUT.PUT_LINE('OK: CK_UPGRADE_TIPO removida');
    ELSE
        DBMS_OUTPUT.PUT_LINE('SKIP: CK_UPGRADE_TIPO não encontrada');
    END IF;
END;
/

-- ----------------------------------------------------------------------------
-- 2. Recria a constraint com todos os tipos suportados pelo aplicativo
-- ----------------------------------------------------------------------------
ALTER TABLE TBL_EQUIPAMENTO_UPGRADE
    ADD CONSTRAINT CK_UPGRADE_TIPO
    CHECK (TIPO_UPGRADE IN ('RAM','SSD','HDD','BATTERY','SCREEN','KEYBOARD','COVER','GPU','OTHER'));

DBMS_OUTPUT.PUT_LINE('OK: CK_UPGRADE_TIPO recriada com todos os tipos');

-- ----------------------------------------------------------------------------
-- 3. Verificação
-- ----------------------------------------------------------------------------
SELECT CONSTRAINT_NAME, SEARCH_CONDITION, STATUS
  FROM USER_CONSTRAINTS
 WHERE TABLE_NAME = 'TBL_EQUIPAMENTO_UPGRADE'
   AND CONSTRAINT_NAME = 'CK_UPGRADE_TIPO';
