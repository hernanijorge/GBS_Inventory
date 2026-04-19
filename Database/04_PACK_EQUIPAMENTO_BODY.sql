-- ============================================================================
-- PACK_EQUIPAMENTO - Body
-- Implementação das procedures e functions
-- ============================================================================

CREATE OR REPLACE PACKAGE BODY PACK_EQUIPAMENTO IS

-- ----------------------------------------------------------------------------
FUNCTION FUNC_PROXIMO_EQUIPAMENTO
RETURN NUMBER

AS

V_ID NUMBER;

BEGIN

       SELECT SEQ_EQUIPAMENTO.NEXTVAL INTO V_ID FROM DUAL;

       RETURN V_ID;

END;

-- ----------------------------------------------------------------------------
FUNCTION FUNC_EXISTE_UID (V_INTERNAL_UID IN VARCHAR2)
RETURN NUMBER

AS

V_CONTADOR NUMBER;

BEGIN

       SELECT COUNT(*)
       INTO   V_CONTADOR
       FROM   TBL_EQUIPAMENTO
       WHERE  INTERNAL_UID = V_INTERNAL_UID;

       IF V_CONTADOR > 0 THEN
          RETURN 1;
       ELSE
          RETURN 0;
       END IF;

EXCEPTION
       WHEN OTHERS THEN
            RETURN 0;
END;

-- ----------------------------------------------------------------------------
FUNCTION FUNC_STATUS_DESCRICAO (V_STATUS IN VARCHAR2)
RETURN VARCHAR2

AS

BEGIN

       RETURN CASE V_STATUS
                  WHEN 'IN_STOCK'  THEN 'Em Estoque'
                  WHEN 'LISTED'    THEN 'Anunciado'
                  WHEN 'SOLD'      THEN 'Vendido'
                  WHEN 'SHIPPED'   THEN 'Enviado'
                  WHEN 'UPGRADING' THEN 'Em Upgrade'
                  WHEN 'BROKEN'    THEN 'Quebrado'
                  ELSE 'Desconhecido'
              END;

END;

-- ----------------------------------------------------------------------------
PROCEDURE PROC_INSERT_EQUIPAMENTO (
    V_INTERNAL_UID         VARCHAR2,
    V_MANUFACTURER         VARCHAR2,
    V_MODEL                VARCHAR2,
    V_SERIAL_NUMBER        VARCHAR2,
    V_CPU_FAMILY           VARCHAR2,
    V_CPU_MODEL            VARCHAR2,
    V_CPU_SPEED_GHZ        NUMBER,
    V_RAM_GB               VARCHAR2,
    V_STORAGE_GB           VARCHAR2,
    V_HARD_DRIVE_TYPE      VARCHAR2,
    V_RESOLUTION           VARCHAR2,
    V_GRAPHICS             VARCHAR2,
    V_DEVICE_TYPE          VARCHAR2,
    V_CONDITION_STATUS     VARCHAR2,
    V_NOTES                VARCHAR2,
    V_SOURCE_BATCH         VARCHAR2,
    V_STATUS               VARCHAR2,
    V_ID_EMPRESA           NUMBER,
    V_ID_USUARIO_CADASTRO  NUMBER,
    V_ID                   OUT NUMBER)

AS

V_IDEQUIPAMENTO NUMBER;

BEGIN
         V_IDEQUIPAMENTO := FUNC_PROXIMO_EQUIPAMENTO;

         V_ID := V_IDEQUIPAMENTO;

         INSERT INTO TBL_EQUIPAMENTO
                      (ID_EQUIPAMENTO,
                       INTERNAL_UID,
                       MANUFACTURER,
                       MODEL,
                       SERIAL_NUMBER,
                       CPU_FAMILY,
                       CPU_MODEL,
                       CPU_SPEED_GHZ,
                       RAM_GB,
                       STORAGE_GB,
                       HARD_DRIVE_TYPE,
                       RESOLUTION,
                       GRAPHICS,
                       DEVICE_TYPE,
                       CONDITION_STATUS,
                       NOTES,
                       SOURCE_BATCH,
                       STATUS,
                       ID_EMPRESA,
                       DATA_CADASTRO,
                       ID_USUARIO_CADASTRO)
               VALUES
                      (V_IDEQUIPAMENTO,
                       V_INTERNAL_UID,
                       UPPER(V_MANUFACTURER),
                       V_MODEL,
                       UPPER(V_SERIAL_NUMBER),
                       V_CPU_FAMILY,
                       V_CPU_MODEL,
                       V_CPU_SPEED_GHZ,
                       V_RAM_GB,
                       V_STORAGE_GB,
                       V_HARD_DRIVE_TYPE,
                       V_RESOLUTION,
                       V_GRAPHICS,
                       NVL(UPPER(V_DEVICE_TYPE), 'LAPTOP'),
                       NVL(UPPER(V_CONDITION_STATUS), 'GOOD'),
                       V_NOTES,
                       V_SOURCE_BATCH,
                       NVL(UPPER(V_STATUS), 'IN_STOCK'),
                       V_ID_EMPRESA,
                       TRUNC(SYSDATE),
                       V_ID_USUARIO_CADASTRO);

END;

-- ----------------------------------------------------------------------------
PROCEDURE PROC_UPDATE_EQUIPAMENTO (
    V_MANUFACTURER         VARCHAR2,
    V_MODEL                VARCHAR2,
    V_CPU_FAMILY           VARCHAR2,
    V_CPU_MODEL            VARCHAR2,
    V_CPU_SPEED_GHZ        NUMBER,
    V_RAM_GB               VARCHAR2,
    V_STORAGE_GB           VARCHAR2,
    V_HARD_DRIVE_TYPE      VARCHAR2,
    V_RESOLUTION           VARCHAR2,
    V_GRAPHICS             VARCHAR2,
    V_DEVICE_TYPE          VARCHAR2,
    V_CONDITION_STATUS     VARCHAR2,
    V_NOTES                VARCHAR2,
    V_STATUS               VARCHAR2,
    V_ID                   NUMBER)

AS

BEGIN

         UPDATE TBL_EQUIPAMENTO SET
                  MANUFACTURER     = UPPER(V_MANUFACTURER),
                  MODEL            = V_MODEL,
                  CPU_FAMILY       = V_CPU_FAMILY,
                  CPU_MODEL        = V_CPU_MODEL,
                  CPU_SPEED_GHZ    = V_CPU_SPEED_GHZ,
                  RAM_GB           = V_RAM_GB,
                  STORAGE_GB       = V_STORAGE_GB,
                  HARD_DRIVE_TYPE  = V_HARD_DRIVE_TYPE,
                  RESOLUTION       = V_RESOLUTION,
                  GRAPHICS         = V_GRAPHICS,
                  DEVICE_TYPE      = UPPER(V_DEVICE_TYPE),
                  CONDITION_STATUS = UPPER(V_CONDITION_STATUS),
                  NOTES            = V_NOTES,
                  STATUS           = UPPER(V_STATUS),
                  DATA_ALTERACAO   = SYSDATE
          WHERE   ID_EQUIPAMENTO = V_ID;

END;

-- ----------------------------------------------------------------------------
PROCEDURE PROC_UPDATE_STATUS (V_STATUS VARCHAR2, V_ID NUMBER)

AS

BEGIN

         UPDATE TBL_EQUIPAMENTO SET
                  STATUS         = UPPER(V_STATUS),
                  DATA_ALTERACAO = SYSDATE
          WHERE   ID_EQUIPAMENTO = V_ID;

END;

-- ----------------------------------------------------------------------------
PROCEDURE PROC_DELETE_EQUIPAMENTO (V_ID IN NUMBER)

AS

BEGIN

       DELETE FROM TBL_REMESSA_ITEM
       WHERE  ID_EQUIPAMENTO = V_ID;

       DELETE FROM TBL_EQUIPAMENTO_UPGRADE
       WHERE  ID_EQUIPAMENTO = V_ID;

       DELETE FROM TBL_EQUIPAMENTO
       WHERE  ID_EQUIPAMENTO = V_ID;

END;

-- ----------------------------------------------------------------------------
PROCEDURE PROC_UPSERT_EQUIPAMENTO (
    V_INTERNAL_UID         VARCHAR2,
    V_MANUFACTURER         VARCHAR2,
    V_MODEL                VARCHAR2,
    V_SERIAL_NUMBER        VARCHAR2,
    V_CPU_FAMILY           VARCHAR2,
    V_CPU_MODEL            VARCHAR2,
    V_CPU_SPEED_GHZ        NUMBER,
    V_RAM_GB               VARCHAR2,
    V_STORAGE_GB           VARCHAR2,
    V_HARD_DRIVE_TYPE      VARCHAR2,
    V_CONDITION_STATUS     VARCHAR2,
    V_NOTES                VARCHAR2,
    V_SOURCE_BATCH         VARCHAR2,
    V_RESULTADO            OUT VARCHAR2)

AS

V_IDEXISTE NUMBER;

BEGIN

         SELECT COUNT(*) INTO V_IDEXISTE
         FROM   TBL_EQUIPAMENTO
         WHERE  INTERNAL_UID = V_INTERNAL_UID;

         IF V_IDEXISTE = 0 THEN

            INSERT INTO TBL_EQUIPAMENTO
                        (ID_EQUIPAMENTO,
                         INTERNAL_UID,
                         MANUFACTURER,
                         MODEL,
                         SERIAL_NUMBER,
                         CPU_FAMILY,
                         CPU_MODEL,
                         CPU_SPEED_GHZ,
                         RAM_GB,
                         STORAGE_GB,
                         HARD_DRIVE_TYPE,
                         CONDITION_STATUS,
                         NOTES,
                         SOURCE_BATCH,
                         STATUS,
                         ID_EMPRESA,
                         DATA_CADASTRO)
                 VALUES
                        (SEQ_EQUIPAMENTO.NEXTVAL,
                         V_INTERNAL_UID,
                         UPPER(V_MANUFACTURER),
                         V_MODEL,
                         UPPER(V_SERIAL_NUMBER),
                         V_CPU_FAMILY,
                         V_CPU_MODEL,
                         V_CPU_SPEED_GHZ,
                         V_RAM_GB,
                         V_STORAGE_GB,
                         V_HARD_DRIVE_TYPE,
                         NVL(UPPER(V_CONDITION_STATUS), 'GOOD'),
                         V_NOTES,
                         V_SOURCE_BATCH,
                         'IN_STOCK',
                         1,
                         TRUNC(SYSDATE));

            V_RESULTADO := 'INSERTED';

         ELSE

            UPDATE TBL_EQUIPAMENTO SET
                      MANUFACTURER     = UPPER(V_MANUFACTURER),
                      MODEL            = V_MODEL,
                      CPU_FAMILY       = V_CPU_FAMILY,
                      CPU_MODEL        = V_CPU_MODEL,
                      CPU_SPEED_GHZ    = V_CPU_SPEED_GHZ,
                      RAM_GB           = V_RAM_GB,
                      STORAGE_GB       = V_STORAGE_GB,
                      HARD_DRIVE_TYPE  = V_HARD_DRIVE_TYPE,
                      CONDITION_STATUS = NVL(UPPER(V_CONDITION_STATUS), CONDITION_STATUS),
                      NOTES            = V_NOTES,
                      SOURCE_BATCH     = V_SOURCE_BATCH,
                      DATA_ALTERACAO   = SYSDATE
             WHERE    INTERNAL_UID = V_INTERNAL_UID;

            V_RESULTADO := 'UPDATED';

         END IF;

EXCEPTION
        WHEN OTHERS THEN
             V_RESULTADO := 'ERROR: ' || SQLERRM;

END;

-- ----------------------------------------------------------------------------
PROCEDURE PROC_SELECT_EQUIPAMENTO (
    V_ID      IN  NUMBER,
    V_CURSOR  OUT SYS_REFCURSOR)

AS

BEGIN

        OPEN V_CURSOR FOR
             SELECT E.*,
                    FUNC_STATUS_DESCRICAO(E.STATUS) AS STATUS_DESCRICAO
             FROM   TBL_EQUIPAMENTO E
             WHERE  ID_EQUIPAMENTO = V_ID;

END;

-- ----------------------------------------------------------------------------
PROCEDURE PROC_SELECT_EQUIPAMENTO_UID (
    V_INTERNAL_UID IN  VARCHAR2,
    V_CURSOR       OUT SYS_REFCURSOR)

AS

BEGIN

        OPEN V_CURSOR FOR
             SELECT E.*,
                    FUNC_STATUS_DESCRICAO(E.STATUS) AS STATUS_DESCRICAO,
                    (SELECT COUNT(*) FROM TBL_EQUIPAMENTO_UPGRADE U
                     WHERE U.ID_EQUIPAMENTO = E.ID_EQUIPAMENTO) AS TOTAL_UPGRADES
             FROM   TBL_EQUIPAMENTO E
             WHERE  UPPER(TRIM(INTERNAL_UID)) = UPPER(TRIM(V_INTERNAL_UID));

END;

-- ----------------------------------------------------------------------------
PROCEDURE PROC_SELECT_EQUIPAMENTO_SERIAL (
    V_SERIAL IN  VARCHAR2,
    V_CURSOR OUT SYS_REFCURSOR)

AS

BEGIN

        OPEN V_CURSOR FOR
             SELECT E.*,
                    FUNC_STATUS_DESCRICAO(E.STATUS) AS STATUS_DESCRICAO
             FROM   TBL_EQUIPAMENTO E
             WHERE  UPPER(TRIM(SERIAL_NUMBER)) = UPPER(TRIM(V_SERIAL));

END;

-- ----------------------------------------------------------------------------
PROCEDURE PROC_SELECT_EQUIPAMENTOS (
    V_EMPRESA IN  NUMBER,
    V_CURSOR  OUT SYS_REFCURSOR)

AS

BEGIN

        OPEN V_CURSOR FOR
             SELECT E.ID_EQUIPAMENTO,
                    E.INTERNAL_UID,
                    E.MANUFACTURER,
                    E.MODEL,
                    E.SERIAL_NUMBER,
                    E.CPU_MODEL,
                    E.CPU_SPEED_GHZ,
                    E.RAM_GB,
                    E.STORAGE_GB,
                    E.CONDITION_STATUS,
                    E.STATUS,
                    FUNC_STATUS_DESCRICAO(E.STATUS) AS STATUS_DESCRICAO,
                    E.SOURCE_BATCH,
                    E.NOTES,
                    E.DATA_CADASTRO
             FROM   TBL_EQUIPAMENTO E
             WHERE  (E.ID_EMPRESA = V_EMPRESA OR V_EMPRESA IS NULL)
             ORDER  BY E.DATA_CADASTRO DESC, E.INTERNAL_UID DESC;

END;

-- ----------------------------------------------------------------------------
PROCEDURE PROC_SELECT_POR_FILTRO (
    V_PESQUISA     IN  VARCHAR2,
    V_MANUFACTURER IN  VARCHAR2,
    V_CONDITION    IN  VARCHAR2,
    V_STATUS       IN  VARCHAR2,
    V_CURSOR       OUT SYS_REFCURSOR)

AS

V_PESQ VARCHAR2(100);

BEGIN

        V_PESQ := '%' || UPPER(TRIM(NVL(V_PESQUISA, ''))) || '%';

        OPEN V_CURSOR FOR
             SELECT E.ID_EQUIPAMENTO,
                    E.INTERNAL_UID,
                    E.MANUFACTURER,
                    E.MODEL,
                    E.SERIAL_NUMBER,
                    E.CPU_MODEL,
                    E.CPU_SPEED_GHZ,
                    E.RAM_GB,
                    E.STORAGE_GB,
                    E.CONDITION_STATUS,
                    E.STATUS,
                    FUNC_STATUS_DESCRICAO(E.STATUS) AS STATUS_DESCRICAO,
                    E.SOURCE_BATCH,
                    E.NOTES,
                    E.DATA_CADASTRO
             FROM   TBL_EQUIPAMENTO E
             WHERE  (V_PESQUISA IS NULL
                     OR UPPER(E.INTERNAL_UID)   LIKE V_PESQ
                     OR UPPER(E.SERIAL_NUMBER)  LIKE V_PESQ
                     OR UPPER(E.MODEL)          LIKE V_PESQ
                     OR UPPER(E.CPU_MODEL)      LIKE V_PESQ
                     OR UPPER(E.MANUFACTURER)   LIKE V_PESQ)
             AND    (V_MANUFACTURER IS NULL OR UPPER(E.MANUFACTURER) = UPPER(V_MANUFACTURER))
             AND    (V_CONDITION    IS NULL OR UPPER(E.CONDITION_STATUS) = UPPER(V_CONDITION))
             AND    (V_STATUS       IS NULL OR UPPER(E.STATUS) = UPPER(V_STATUS))
             ORDER  BY E.DATA_CADASTRO DESC;

END;

-- ----------------------------------------------------------------------------
PROCEDURE PROC_DASHBOARD_TOTAIS (
    V_EMPRESA IN  NUMBER,
    V_CURSOR  OUT SYS_REFCURSOR)

AS

BEGIN

        OPEN V_CURSOR FOR
             SELECT
                (SELECT COUNT(*) FROM TBL_EQUIPAMENTO
                 WHERE (ID_EMPRESA = V_EMPRESA OR V_EMPRESA IS NULL))                     AS TOTAL_UNIDADES,
                (SELECT COUNT(*) FROM TBL_EQUIPAMENTO
                 WHERE STATUS = 'IN_STOCK'
                 AND (ID_EMPRESA = V_EMPRESA OR V_EMPRESA IS NULL))                       AS EM_ESTOQUE,
                (SELECT COUNT(*) FROM TBL_EQUIPAMENTO
                 WHERE CONDITION_STATUS IN ('EXCELLENT','GOOD')
                 AND (ID_EMPRESA = V_EMPRESA OR V_EMPRESA IS NULL))                       AS CONDICAO_BOA,
                (SELECT COUNT(*) FROM TBL_EQUIPAMENTO_UPGRADE U, TBL_EQUIPAMENTO E
                 WHERE U.ID_EQUIPAMENTO = E.ID_EQUIPAMENTO
                 AND   U.DATA_UPGRADE >= TRUNC(SYSDATE) - 30
                 AND   (E.ID_EMPRESA = V_EMPRESA OR V_EMPRESA IS NULL))                   AS UPGRADES_30D,
                (SELECT COUNT(*) FROM TBL_REMESSA
                 WHERE STATUS_REMESSA NOT IN ('DELIVERED','RETURNED'))                    AS REMESSAS_ATIVAS
             FROM DUAL;

END;

-- ----------------------------------------------------------------------------
PROCEDURE PROC_RESUMO_WHATSAPP (
    V_EMPRESA IN  NUMBER,
    V_CURSOR  OUT SYS_REFCURSOR)

AS

BEGIN

        OPEN V_CURSOR FOR
             SELECT MANUFACTURER,
                    MODEL,
                    COUNT(*) AS QTD,
                    SUM(DECODE(CONDITION_STATUS,'EXCELLENT',1,'GOOD',1,0)) AS QTD_BOAS
             FROM   TBL_EQUIPAMENTO
             WHERE  STATUS = 'IN_STOCK'
             AND    (ID_EMPRESA = V_EMPRESA OR V_EMPRESA IS NULL)
             GROUP  BY MANUFACTURER, MODEL
             ORDER  BY COUNT(*) DESC;

END;

END PACK_EQUIPAMENTO;
/
