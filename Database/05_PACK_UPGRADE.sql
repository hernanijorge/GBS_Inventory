-- ============================================================================
-- PACK_UPGRADE - Specification
-- Controle de upgrades (RAM, SSD, etc) em equipamentos
-- ============================================================================

CREATE OR REPLACE PACKAGE PACK_UPGRADE IS

    FUNCTION FUNC_PROXIMO_UPGRADE
        RETURN NUMBER;

    FUNCTION FUNC_CUSTO_TOTAL_EQUIPAMENTO (V_ID_EQUIPAMENTO IN NUMBER)
        RETURN NUMBER;

    PROCEDURE PROC_INSERT_UPGRADE (
        V_ID_EQUIPAMENTO   NUMBER,
        V_INTERNAL_UID     VARCHAR2,
        V_COMPONENT_TYPE   VARCHAR2,
        V_VALUE_BEFORE     VARCHAR2,
        V_VALUE_AFTER      VARCHAR2,
        V_PART_SERIAL      VARCHAR2,
        V_SOURCE_ORIGEM    VARCHAR2,
        V_COST_USD         NUMBER,
        V_TECHNICIAN       VARCHAR2,
        V_NOTES            VARCHAR2,
        V_ID               OUT NUMBER);

    PROCEDURE PROC_DELETE_UPGRADE (V_ID IN NUMBER);

    PROCEDURE PROC_SELECT_UPGRADES_EQUIP (
        V_ID_EQUIPAMENTO IN  NUMBER,
        V_CURSOR         OUT SYS_REFCURSOR);

    PROCEDURE PROC_SELECT_UPGRADES_RECENTES (
        V_DIAS   IN  NUMBER,
        V_CURSOR OUT SYS_REFCURSOR);

END PACK_UPGRADE;
/

-- ============================================================================
-- PACK_UPGRADE - Body
-- ============================================================================

CREATE OR REPLACE PACKAGE BODY PACK_UPGRADE IS

-- ----------------------------------------------------------------------------
FUNCTION FUNC_PROXIMO_UPGRADE
RETURN NUMBER

AS

V_ID NUMBER;

BEGIN

       SELECT SEQ_UPGRADE.NEXTVAL INTO V_ID FROM DUAL;

       RETURN V_ID;

END;

-- ----------------------------------------------------------------------------
FUNCTION FUNC_CUSTO_TOTAL_EQUIPAMENTO (V_ID_EQUIPAMENTO IN NUMBER)
RETURN NUMBER

AS

V_CUSTO NUMBER(10,2);

BEGIN

       SELECT NVL(SUM(COST_USD), 0)
       INTO   V_CUSTO
       FROM   TBL_EQUIPAMENTO_UPGRADE
       WHERE  ID_EQUIPAMENTO = V_ID_EQUIPAMENTO;

       RETURN V_CUSTO;

EXCEPTION
       WHEN OTHERS THEN
            RETURN 0;
END;

-- ----------------------------------------------------------------------------
PROCEDURE PROC_INSERT_UPGRADE (
    V_ID_EQUIPAMENTO   NUMBER,
    V_INTERNAL_UID     VARCHAR2,
    V_COMPONENT_TYPE   VARCHAR2,
    V_VALUE_BEFORE     VARCHAR2,
    V_VALUE_AFTER      VARCHAR2,
    V_PART_SERIAL      VARCHAR2,
    V_SOURCE_ORIGEM    VARCHAR2,
    V_COST_USD         NUMBER,
    V_TECHNICIAN       VARCHAR2,
    V_NOTES            VARCHAR2,
    V_ID               OUT NUMBER)

AS

V_IDUPGRADE NUMBER;

BEGIN
         V_IDUPGRADE := FUNC_PROXIMO_UPGRADE;
         V_ID := V_IDUPGRADE;

         INSERT INTO TBL_EQUIPAMENTO_UPGRADE
                      (ID_UPGRADE,
                       ID_EQUIPAMENTO,
                       INTERNAL_UID,
                       COMPONENT_TYPE,
                       VALUE_BEFORE,
                       VALUE_AFTER,
                       PART_SERIAL,
                       SOURCE_ORIGEM,
                       COST_USD,
                       TECHNICIAN,
                       NOTES,
                       DATA_UPGRADE)
               VALUES
                      (V_IDUPGRADE,
                       V_ID_EQUIPAMENTO,
                       V_INTERNAL_UID,
                       UPPER(V_COMPONENT_TYPE),
                       V_VALUE_BEFORE,
                       V_VALUE_AFTER,
                       V_PART_SERIAL,
                       UPPER(V_SOURCE_ORIGEM),
                       V_COST_USD,
                       V_TECHNICIAN,
                       V_NOTES,
                       SYSDATE);

         -- Atualiza o campo do equipamento conforme o componente
         IF UPPER(V_COMPONENT_TYPE) = 'RAM' THEN
            UPDATE TBL_EQUIPAMENTO SET RAM_GB = V_VALUE_AFTER, DATA_ALTERACAO = SYSDATE
            WHERE  ID_EQUIPAMENTO = V_ID_EQUIPAMENTO;
         ELSIF UPPER(V_COMPONENT_TYPE) IN ('SSD','HDD') THEN
            UPDATE TBL_EQUIPAMENTO SET STORAGE_GB = V_VALUE_AFTER, DATA_ALTERACAO = SYSDATE
            WHERE  ID_EQUIPAMENTO = V_ID_EQUIPAMENTO;
         END IF;

END;

-- ----------------------------------------------------------------------------
PROCEDURE PROC_DELETE_UPGRADE (V_ID IN NUMBER)

AS

BEGIN

       DELETE FROM TBL_EQUIPAMENTO_UPGRADE
       WHERE  ID_UPGRADE = V_ID;

END;

-- ----------------------------------------------------------------------------
PROCEDURE PROC_SELECT_UPGRADES_EQUIP (
    V_ID_EQUIPAMENTO IN  NUMBER,
    V_CURSOR         OUT SYS_REFCURSOR)

AS

BEGIN

        OPEN V_CURSOR FOR
             SELECT U.ID_UPGRADE,
                    U.ID_EQUIPAMENTO,
                    U.INTERNAL_UID,
                    U.COMPONENT_TYPE,
                    U.VALUE_BEFORE,
                    U.VALUE_AFTER,
                    U.PART_SERIAL,
                    U.SOURCE_ORIGEM,
                    U.COST_USD,
                    U.TECHNICIAN,
                    U.NOTES,
                    U.DATA_UPGRADE
             FROM   TBL_EQUIPAMENTO_UPGRADE U
             WHERE  U.ID_EQUIPAMENTO = V_ID_EQUIPAMENTO
             ORDER  BY U.DATA_UPGRADE DESC;

END;

-- ----------------------------------------------------------------------------
PROCEDURE PROC_SELECT_UPGRADES_RECENTES (
    V_DIAS   IN  NUMBER,
    V_CURSOR OUT SYS_REFCURSOR)

AS

BEGIN

        OPEN V_CURSOR FOR
             SELECT U.ID_UPGRADE,
                    U.INTERNAL_UID,
                    E.MANUFACTURER,
                    E.MODEL,
                    U.COMPONENT_TYPE,
                    U.VALUE_BEFORE,
                    U.VALUE_AFTER,
                    U.SOURCE_ORIGEM,
                    U.COST_USD,
                    U.TECHNICIAN,
                    U.DATA_UPGRADE
             FROM   TBL_EQUIPAMENTO_UPGRADE U,
                    TBL_EQUIPAMENTO         E
             WHERE  U.ID_EQUIPAMENTO = E.ID_EQUIPAMENTO
             AND    U.DATA_UPGRADE  >= TRUNC(SYSDATE) - NVL(V_DIAS, 30)
             ORDER  BY U.DATA_UPGRADE DESC;

END;

END PACK_UPGRADE;
/
