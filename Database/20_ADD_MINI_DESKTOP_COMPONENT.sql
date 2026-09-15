-- ============================================================================
-- GBS Inventory - Migration 20
-- Adiciona tipo MINI_DESKTOP em TBL_COMPONENT + coluna CPU (modelo/familia).
-- Idempotente.
-- ============================================================================

-- ----------------------------------------------------------------------------
-- 1. Adiciona coluna CPU (se nao existir)
-- ----------------------------------------------------------------------------
DECLARE V_N NUMBER;
BEGIN
    SELECT COUNT(*) INTO V_N FROM USER_TAB_COLUMNS
     WHERE TABLE_NAME = 'TBL_COMPONENT' AND COLUMN_NAME = 'CPU';
    IF V_N = 0 THEN
        EXECUTE IMMEDIATE 'ALTER TABLE TBL_COMPONENT ADD CPU VARCHAR2(100)';
        DBMS_OUTPUT.PUT_LINE('OK: coluna CPU adicionada');
    ELSE
        DBMS_OUTPUT.PUT_LINE('SKIP: CPU ja existe');
    END IF;
END;
/

-- ----------------------------------------------------------------------------
-- 2. CK_COMP_TYPE: inclui MINI_DESKTOP no enum de COMPONENT_TYPE
-- ----------------------------------------------------------------------------
DECLARE V_N NUMBER;
BEGIN
    SELECT COUNT(*) INTO V_N FROM USER_CONSTRAINTS
     WHERE TABLE_NAME = 'TBL_COMPONENT' AND CONSTRAINT_NAME = 'CK_COMP_TYPE';
    IF V_N > 0 THEN
        EXECUTE IMMEDIATE 'ALTER TABLE TBL_COMPONENT DROP CONSTRAINT CK_COMP_TYPE';
    END IF;
    EXECUTE IMMEDIATE 'ALTER TABLE TBL_COMPONENT ADD CONSTRAINT CK_COMP_TYPE ' ||
                       'CHECK (COMPONENT_TYPE IN (''RAM'',''SSD'',''HDD'',''MINI_DESKTOP''))';
    DBMS_OUTPUT.PUT_LINE('OK: CK_COMP_TYPE atualizada com MINI_DESKTOP');
END;
/

-- ----------------------------------------------------------------------------
-- 3. PACK_COMPONENT: PROC_INSERT ganha P_CPU; selects passam a trazer CPU
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
        P_STATUS IN VARCHAR2
    );

END PACK_COMPONENT;
/

CREATE OR REPLACE PACKAGE BODY PACK_COMPONENT AS

-- ----------------------------------------------------------------------------
PROCEDURE PROC_INSERT(
    P_COMPONENT_TYPE IN VARCHAR2, P_CAPACITY_GB IN NUMBER, P_SPEED_MHZ IN NUMBER,
    P_GENERATION IN VARCHAR2, P_BRAND IN VARCHAR2, P_PART_NUMBER IN VARCHAR2,
    P_CONDITION IN VARCHAR2, P_STATUS IN VARCHAR2, P_SOURCE_BATCH IN VARCHAR2,
    P_NOTES IN VARCHAR2, P_CPU IN VARCHAR2, P_UID_OUT OUT VARCHAR2
) IS
    V_ID NUMBER; V_PFX VARCHAR2(1);
BEGIN
    SELECT SEQ_COMPONENT.NEXTVAL INTO V_ID FROM DUAL;
    V_PFX := CASE UPPER(TRIM(P_COMPONENT_TYPE))
                 WHEN 'RAM' THEN 'R' WHEN 'SSD' THEN 'S' WHEN 'HDD' THEN 'H'
                 WHEN 'MINI_DESKTOP' THEN 'M' ELSE 'C'
             END;
    P_UID_OUT := V_PFX || LPAD(TO_CHAR(V_ID), 6, '0');
    INSERT INTO TBL_COMPONENT
               (ID_COMPONENT, INTERNAL_UID, COMPONENT_TYPE, CAPACITY_GB, SPEED_MHZ,
                GENERATION, BRAND, PART_NUMBER, CONDITION_STATUS, STATUS,
                SOURCE_BATCH, NOTES, CPU, DATE_CREATED)
        VALUES (V_ID, P_UID_OUT, UPPER(TRIM(P_COMPONENT_TYPE)), P_CAPACITY_GB,
                NULLIF(P_SPEED_MHZ, 0), NULLIF(TRIM(P_GENERATION), ''), NULLIF(TRIM(P_BRAND), ''),
                NULLIF(TRIM(P_PART_NUMBER), ''), NVL(NULLIF(TRIM(P_CONDITION), ''), 'GOOD'),
                NVL(NULLIF(TRIM(P_STATUS), ''), 'IN_STOCK'), NULLIF(TRIM(P_SOURCE_BATCH), ''),
                NULLIF(TRIM(P_NOTES), ''), NULLIF(TRIM(P_CPU), ''), SYSDATE);
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
               GENERATION, BRAND, PART_NUMBER, CONDITION_STATUS, STATUS,
               SOURCE_BATCH, NOTES, CPU, DATE_CREATED, DATE_UPDATED
          FROM TBL_COMPONENT
         WHERE (P_FILTER IS NULL
                OR UPPER(INTERNAL_UID) LIKE V_PESQ OR UPPER(BRAND) LIKE V_PESQ
                OR UPPER(PART_NUMBER)  LIKE V_PESQ OR UPPER(GENERATION) LIKE V_PESQ
                OR UPPER(CPU) LIKE V_PESQ)
           AND (P_TYPE   IS NULL OR UPPER(COMPONENT_TYPE) = UPPER(P_TYPE))
           AND (P_STATUS IS NULL OR UPPER(STATUS)         = UPPER(P_STATUS))
         ORDER BY DATE_CREATED DESC;
END;

-- ----------------------------------------------------------------------------
PROCEDURE PROC_SELECT_ID(P_ID IN NUMBER, P_CURSOR OUT SYS_REFCURSOR) IS
BEGIN
    OPEN P_CURSOR FOR
        SELECT ID_COMPONENT, INTERNAL_UID, COMPONENT_TYPE, CAPACITY_GB, SPEED_MHZ,
               GENERATION, BRAND, PART_NUMBER, CONDITION_STATUS, STATUS,
               SOURCE_BATCH, NOTES, CPU, DATE_CREATED, DATE_UPDATED
          FROM TBL_COMPONENT WHERE ID_COMPONENT = P_ID;
END;

-- ----------------------------------------------------------------------------
PROCEDURE PROC_UPDATE_STATUS(P_ID IN NUMBER, P_STATUS IN VARCHAR2) IS
BEGIN
    UPDATE TBL_COMPONENT SET STATUS = P_STATUS, DATE_UPDATED = SYSDATE
     WHERE ID_COMPONENT = P_ID;
END;

END PACK_COMPONENT;
/

-- ----------------------------------------------------------------------------
-- 4. Verificacao final
-- ----------------------------------------------------------------------------
SELECT OBJECT_TYPE, STATUS, LAST_DDL_TIME
  FROM USER_OBJECTS
 WHERE OBJECT_NAME = 'PACK_COMPONENT'
 ORDER BY OBJECT_TYPE;

SELECT LINE, POSITION, TEXT
  FROM USER_ERRORS
 WHERE NAME = 'PACK_COMPONENT'
 ORDER BY SEQUENCE;
