-- ============================================================================
-- GBS Inventory - Migration 18
-- Corrige constraints de CONDITION_STATUS que divergiam do DDL versionado.
--
-- Problema encontrado em 2026-05-15:
--   CK_COMP_COND  tinha 'NO_BATTERY' adicionado manualmente por engano
--                 (deveria estar apenas em TBL_EQUIPAMENTO).
--   CK_EQUIP_CONDITION tinha 'NO_BATTERY' e 'N/A' adicionados manualmente
--                 sem migration — este script os oficializa.
--
-- Idempotente: usa DROP IF EXISTS via bloco PL/SQL.
-- ============================================================================

-- ----------------------------------------------------------------------------
-- 1. TBL_COMPONENT: remove 'NO_BATTERY' da constraint (não pertence aqui)
--    Valores corretos: GOOD | FAIR | POOR | UNTESTED
-- ----------------------------------------------------------------------------
DECLARE
    V_N NUMBER;
BEGIN
    SELECT COUNT(*) INTO V_N FROM USER_CONSTRAINTS
     WHERE TABLE_NAME = 'TBL_COMPONENT' AND CONSTRAINT_NAME = 'CK_COMP_COND';
    IF V_N > 0 THEN
        EXECUTE IMMEDIATE 'ALTER TABLE TBL_COMPONENT DROP CONSTRAINT CK_COMP_COND';
    END IF;
    EXECUTE IMMEDIATE
        'ALTER TABLE TBL_COMPONENT ADD CONSTRAINT CK_COMP_COND
         CHECK (CONDITION_STATUS IN (''GOOD'',''FAIR'',''POOR'',''UNTESTED''))';
    DBMS_OUTPUT.PUT_LINE('OK: CK_COMP_COND recriada sem NO_BATTERY');
END;
/

-- ----------------------------------------------------------------------------
-- 2. TBL_EQUIPAMENTO: oficializa 'NO_BATTERY' e 'N/A' que já estavam no banco
--    Valores corretos: EXCELLENT | GOOD | FAIR | POOR | NO_BATTERY | N/A
-- ----------------------------------------------------------------------------
DECLARE
    V_N NUMBER;
BEGIN
    SELECT COUNT(*) INTO V_N FROM USER_CONSTRAINTS
     WHERE TABLE_NAME = 'TBL_EQUIPAMENTO' AND CONSTRAINT_NAME = 'CK_EQUIP_CONDITION';
    IF V_N > 0 THEN
        EXECUTE IMMEDIATE 'ALTER TABLE TBL_EQUIPAMENTO DROP CONSTRAINT CK_EQUIP_CONDITION';
    END IF;
    EXECUTE IMMEDIATE
        'ALTER TABLE TBL_EQUIPAMENTO ADD CONSTRAINT CK_EQUIP_CONDITION
         CHECK (CONDITION_STATUS IN (''EXCELLENT'',''GOOD'',''FAIR'',''POOR'',''NO_BATTERY'',''N/A''))';
    DBMS_OUTPUT.PUT_LINE('OK: CK_EQUIP_CONDITION recriada com NO_BATTERY e N/A');
END;
/

-- ----------------------------------------------------------------------------
-- 3. Verificação final
-- ----------------------------------------------------------------------------
SELECT TABLE_NAME, CONSTRAINT_NAME, SEARCH_CONDITION
  FROM USER_CONSTRAINTS
 WHERE CONSTRAINT_NAME IN ('CK_COMP_COND', 'CK_EQUIP_CONDITION')
 ORDER BY TABLE_NAME;
