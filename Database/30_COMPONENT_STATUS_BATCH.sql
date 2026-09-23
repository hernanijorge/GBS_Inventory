-- ============================================================================
-- GBS Inventory - Migration 30
-- Change Status em lote na aba Components:
--   * CK_COMP_STAT ganha RESERVED, SHIPPED, IN_REPAIR (INSTALLED continua,
--     usado pelo fluxo de Upgrade)
--   * PROC_UPDATE_STATUS ganha P_NOTE (opcional): anexa "[YYYY-MM-DD] nota"
--     ao NOTES existente
-- Idempotente.
-- ============================================================================

-- ----------------------------------------------------------------------------
-- 1. CK_COMP_STAT
-- ----------------------------------------------------------------------------
DECLARE V_N NUMBER;
BEGIN
    SELECT COUNT(*) INTO V_N FROM USER_CONSTRAINTS
     WHERE TABLE_NAME = 'TBL_COMPONENT' AND CONSTRAINT_NAME = 'CK_COMP_STAT';
    IF V_N > 0 THEN
        EXECUTE IMMEDIATE 'ALTER TABLE TBL_COMPONENT DROP CONSTRAINT CK_COMP_STAT';
    END IF;
    EXECUTE IMMEDIATE 'ALTER TABLE TBL_COMPONENT ADD CONSTRAINT CK_COMP_STAT ' ||
                      'CHECK (STATUS IN (''IN_STOCK'',''INSTALLED'',''RESERVED'',''SHIPPED'',' ||
                      '''SOLD'',''SCRAPPED'',''IN_REPAIR''))';
    DBMS_OUTPUT.PUT_LINE('OK: CK_COMP_STAT recriada');
END;
/

-- ----------------------------------------------------------------------------
-- 2. PACK_COMPONENT: PROC_UPDATE_STATUS com P_NOTE (resto igual a migration 29)
-- ----------------------------------------------------------------------------
CREATE OR REPLACE PACKAGE PACK_COMPONENT AS

    PROCEDURE PROC_INSERT(
        P_COMPONENT_TYPE IN VARCHAR2,
        P_CAPACITY_GB    IN NUMBER,
        P_SPEED_MHZ      IN NUMBER,
        P_GENERATION     IN VARCHAR2,
        P_BRAND          IN VARCHAR2,
        P_PART_NUMBER    IN VARCHAR2,
        P_CONDITION      IN VARCHAR2,
        P_STATUS         IN VARCHAR2,
        P_SOURCE_BATCH   IN VARCHAR2,
        P_NOTES          IN VARCHAR2,
        P_CPU            IN VARCHAR2,
        P_STORAGE_GB     IN NUMBER,
        P_MODEL          IN VARCHAR2,
        P_UID_OUT        OUT VARCHAR2
    );

    PROCEDURE PROC_SELECT_FILTER(
        P_FILTER   IN  VARCHAR2,
        P_TYPE     IN  VARCHAR2,
        P_STATUS   IN  VARCHAR2,
        P_CURSOR   OUT SYS_REFCURSOR
    );

    PROCEDURE PROC_SELECT_ID(
        P_ID     IN  NUMBER,
        P_CURSOR OUT SYS_REFCURSOR
    );

    PROCEDURE PROC_UPDATE_STATUS(
        P_ID     IN NUMBER,
        P_STATUS IN VARCHAR2,
        P_NOTE   IN VARCHAR2 DEFAULT NULL
    );

END PACK_COMPONENT;
/

CREATE OR REPLACE PACKAGE BODY PACK_COMPONENT AS

-- ----------------------------------------------------------------------------
PROCEDURE PROC_INSERT(
    P_COMPONENT_TYPE IN VARCHAR2, P_CAPACITY_GB IN NUMBER, P_SPEED_MHZ IN NUMBER,
    P_GENERATION IN VARCHAR2, P_BRAND IN VARCHAR2, P_PART_NUMBER IN VARCHAR2,
    P_CONDITION IN VARCHAR2, P_STATUS IN VARCHAR2, P_SOURCE_BATCH IN VARCHAR2,
    P_NOTES IN VARCHAR2, P_CPU IN VARCHAR2, P_STORAGE_GB IN NUMBER, P_MODEL IN VARCHAR2,
    P_UID_OUT OUT VARCHAR2
) IS
    V_ID NUMBER; V_PFX VARCHAR2(1);
BEGIN
    SELECT SEQ_COMPONENT.NEXTVAL INTO V_ID FROM DUAL;
    V_PFX := CASE UPPER(TRIM(P_COMPONENT_TYPE))
                 WHEN 'RAM' THEN 'R' WHEN 'SSD' THEN 'S' WHEN 'HDD' THEN 'H'
                 WHEN 'MINI_DESKTOP' THEN 'M' WHEN 'DESKTOP' THEN 'D' ELSE 'C'
             END;
    P_UID_OUT := V_PFX || LPAD(TO_CHAR(V_ID), 6, '0');
    INSERT INTO TBL_COMPONENT
               (ID_COMPONENT, INTERNAL_UID, COMPONENT_TYPE, CAPACITY_GB, SPEED_MHZ,
                GENERATION, BRAND, MODEL, PART_NUMBER, CONDITION_STATUS, STATUS,
                SOURCE_BATCH, NOTES, CPU, STORAGE_GB, DATE_CREATED)
        VALUES (V_ID, P_UID_OUT, UPPER(TRIM(P_COMPONENT_TYPE)), P_CAPACITY_GB,
                NULLIF(P_SPEED_MHZ, 0), NULLIF(TRIM(P_GENERATION), ''), NULLIF(TRIM(P_BRAND), ''),
                NULLIF(TRIM(P_MODEL), ''),
                NULLIF(TRIM(P_PART_NUMBER), ''), NVL(NULLIF(TRIM(P_CONDITION), ''), 'GOOD'),
                NVL(NULLIF(TRIM(P_STATUS), ''), 'IN_STOCK'), NULLIF(TRIM(P_SOURCE_BATCH), ''),
                NULLIF(TRIM(P_NOTES), ''), NULLIF(TRIM(P_CPU), ''), NULLIF(P_STORAGE_GB, 0), SYSDATE);
END;

-- ----------------------------------------------------------------------------
PROCEDURE PROC_SELECT_FILTER(
    P_FILTER IN VARCHAR2, P_TYPE IN VARCHAR2, P_STATUS IN VARCHAR2, P_CURSOR OUT SYS_REFCURSOR
) IS
    V_PESQ VARCHAR2(200);
BEGIN
    V_PESQ := '%' || UPPER(TRIM(NVL(P_FILTER, ''))) || '%';
    OPEN P_CURSOR FOR
        SELECT ID_COMPONENT, INTERNAL_UID, COMPONENT_TYPE, CAPACITY_GB, SPEED_MHZ,
               GENERATION, BRAND, MODEL, PART_NUMBER, CONDITION_STATUS, STATUS,
               SOURCE_BATCH, NOTES, CPU, STORAGE_GB, DATE_CREATED, DATE_UPDATED
          FROM TBL_COMPONENT
         WHERE (P_FILTER IS NULL
                OR UPPER(INTERNAL_UID) LIKE V_PESQ OR UPPER(BRAND) LIKE V_PESQ
                OR UPPER(PART_NUMBER)  LIKE V_PESQ OR UPPER(GENERATION) LIKE V_PESQ
                OR UPPER(CPU) LIKE V_PESQ OR UPPER(MODEL) LIKE V_PESQ)
           AND (P_TYPE   IS NULL OR UPPER(COMPONENT_TYPE) = UPPER(P_TYPE))
           AND (P_STATUS IS NULL OR UPPER(STATUS)         = UPPER(P_STATUS))
         ORDER BY DATE_CREATED DESC;
END;

-- ----------------------------------------------------------------------------
PROCEDURE PROC_SELECT_ID(P_ID IN NUMBER, P_CURSOR OUT SYS_REFCURSOR) IS
BEGIN
    OPEN P_CURSOR FOR
        SELECT ID_COMPONENT, INTERNAL_UID, COMPONENT_TYPE, CAPACITY_GB, SPEED_MHZ,
               GENERATION, BRAND, MODEL, PART_NUMBER, CONDITION_STATUS, STATUS,
               SOURCE_BATCH, NOTES, CPU, STORAGE_GB, DATE_CREATED, DATE_UPDATED
          FROM TBL_COMPONENT WHERE ID_COMPONENT = P_ID;
END;

-- ----------------------------------------------------------------------------
PROCEDURE PROC_UPDATE_STATUS(P_ID IN NUMBER, P_STATUS IN VARCHAR2, P_NOTE IN VARCHAR2 DEFAULT NULL) IS
    V_NOTE VARCHAR2(4000);
BEGIN
    V_NOTE := NULLIF(TRIM(P_NOTE), '');
    IF V_NOTE IS NOT NULL THEN
        V_NOTE := '[' || TO_CHAR(SYSDATE, 'YYYY-MM-DD') || '] ' || V_NOTE;
    END IF;

    UPDATE TBL_COMPONENT
       SET STATUS       = P_STATUS,
           DATE_UPDATED = SYSDATE,
           NOTES        = CASE
                              WHEN V_NOTE IS NULL  THEN NOTES
                              WHEN NOTES  IS NULL  THEN V_NOTE
                              ELSE NOTES || CHR(10) || V_NOTE
                          END
     WHERE ID_COMPONENT = P_ID;

    -- Batch callers run this per id inside one transaction: a missing id must fail the batch
    IF SQL%ROWCOUNT = 0 THEN
        RAISE_APPLICATION_ERROR(-20010, 'Component not found: ID ' || P_ID);
    END IF;
END;

END PACK_COMPONENT;
/

-- ----------------------------------------------------------------------------
-- 3. Verificacao final
-- ----------------------------------------------------------------------------
SELECT OBJECT_TYPE, STATUS, LAST_DDL_TIME
  FROM USER_OBJECTS
 WHERE OBJECT_NAME = 'PACK_COMPONENT'
 ORDER BY OBJECT_TYPE;

SELECT LINE, POSITION, TEXT
  FROM USER_ERRORS
 WHERE NAME = 'PACK_COMPONENT'
 ORDER BY SEQUENCE;
