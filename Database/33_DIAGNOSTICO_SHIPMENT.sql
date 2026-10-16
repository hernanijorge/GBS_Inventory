-- ============================================================================
-- GBS Inventory - Migration 33
-- PACK_DIAGNOSTICO.PROC_LIST_BY_SHIPMENT: certificados de teste por remessa
-- (botao "Test Certificates" da aba Shipments).
--
-- Recria o pacote inteiro (spec + body): procedures do script 32 sem alteracao
-- + PROC_LIST_BY_SHIPMENT. Sem DDL de tabela. Idempotente.
-- Lembrete XE 10.2 + ODP.NET 21.x: nada de LOB em cursor (TTC Error); o PDF
-- de cada maquina continua sendo lido por PROC_GET_PDF (parametro OUT).
-- ============================================================================

CREATE OR REPLACE PACKAGE PACK_DIAGNOSTICO AS

    -- Grava um diagnostico. Resolve ID_EQUIPAMENTO por INTERNAL_UID (ignorando
    -- zeros a esquerda) e, se nao achar, por SERIAL_NUMBER = SERVICE_TAG.
    -- P_ID_EQUIPAMENTO_OUT volta NULL quando o equipamento nao foi encontrado.
    PROCEDURE PROC_INSERT(
        P_INTERNAL_UID       IN  VARCHAR2,
        P_SERVICE_TAG        IN  VARCHAR2,
        P_MANUFACTURER       IN  VARCHAR2,
        P_MODEL              IN  VARCHAR2,
        P_CPU                IN  VARCHAR2,
        P_RAM_GIB            IN  NUMBER,
        P_STORAGE            IN  VARCHAR2,
        P_VERDICT            IN  VARCHAR2,
        P_BATTERY_HEALTH     IN  NUMBER,
        P_BATTERY_CYCLES     IN  NUMBER,
        P_SSD_LIFE           IN  NUMBER,
        P_FAIL_ITEMS         IN  VARCHAR2,
        P_WARN_ITEMS         IN  VARCHAR2,
        P_NOT_TESTED_ITEMS   IN  VARCHAR2,
        P_TECHNICIAN         IN  VARCHAR2,
        P_ORIGIN             IN  VARCHAR2,
        P_TOOL_VERSION       IN  VARCHAR2,
        P_STAMP_UTC          IN  VARCHAR2,
        P_RESULT_JSON        IN  CLOB,
        P_PDF_FILENAME       IN  VARCHAR2,
        P_PDF_BLOB           IN  BLOB,
        P_ID_OUT             OUT NUMBER,
        P_ID_EQUIPAMENTO_OUT OUT NUMBER
    );

    -- Lista sem colunas LOB. Filtros opcionais (NULL = todos):
    -- UID/tag por "contem", verdict exato, periodo por DATA_TESTE (dias inteiros).
    PROCEDURE PROC_LIST(
        P_INTERNAL_UID IN  VARCHAR2,
        P_SERVICE_TAG  IN  VARCHAR2,
        P_VERDICT      IN  VARCHAR2,
        P_DATA_INI     IN  DATE,
        P_DATA_FIM     IN  DATE,
        P_CURSOR       OUT SYS_REFCURSOR
    );

    PROCEDURE PROC_GET_PDF(
        P_ID           IN  NUMBER,
        P_PDF_FILENAME OUT VARCHAR2,
        P_PDF_BLOB     OUT BLOB
    );

    PROCEDURE PROC_GET_JSON(
        P_ID          IN  NUMBER,
        P_RESULT_JSON OUT CLOB
    );


    -- Uma linha por equipamento da remessa com o diagnostico MAIS RECENTE
    -- (LEFT JOIN: equipamento sem teste aparece com ID_DIAGNOSTICO nulo).
    -- Vinculo: TBL_DIAGNOSTICO.ID_EQUIPAMENTO; para diagnosticos gravados sem
    -- vinculo (maquina testada antes do cadastro) usa INTERNAL_UID (sem zeros a
    -- esquerda) ou SERVICE_TAG = SERIAL_NUMBER. Ordem: UID numerica.
    PROCEDURE PROC_LIST_BY_SHIPMENT(
        P_ID_REMESSA IN  NUMBER,
        P_CURSOR     OUT SYS_REFCURSOR
    );

END PACK_DIAGNOSTICO;
/

CREATE OR REPLACE PACKAGE BODY PACK_DIAGNOSTICO AS

-- ----------------------------------------------------------------------------
FUNCTION FUNC_RESOLVE_EQUIPAMENTO(P_INTERNAL_UID IN VARCHAR2, P_SERVICE_TAG IN VARCHAR2)
    RETURN NUMBER
IS
    V_ID  NUMBER;
    V_UID VARCHAR2(80);
    V_TAG VARCHAR2(100);
BEGIN
    V_UID := LTRIM(UPPER(TRIM(P_INTERNAL_UID)), '0');
    V_TAG := UPPER(TRIM(P_SERVICE_TAG));

    IF V_UID IS NOT NULL THEN
        SELECT MAX(ID_EQUIPAMENTO) INTO V_ID
          FROM TBL_EQUIPAMENTO
         WHERE LTRIM(UPPER(TRIM(INTERNAL_UID)), '0') = V_UID;
    END IF;

    IF V_ID IS NULL AND V_TAG IS NOT NULL THEN
        SELECT MAX(ID_EQUIPAMENTO) INTO V_ID
          FROM TBL_EQUIPAMENTO
         WHERE UPPER(TRIM(SERIAL_NUMBER)) = V_TAG;
    END IF;

    RETURN V_ID;
END;

-- ----------------------------------------------------------------------------
PROCEDURE PROC_INSERT(
    P_INTERNAL_UID IN VARCHAR2, P_SERVICE_TAG IN VARCHAR2, P_MANUFACTURER IN VARCHAR2,
    P_MODEL IN VARCHAR2, P_CPU IN VARCHAR2, P_RAM_GIB IN NUMBER, P_STORAGE IN VARCHAR2,
    P_VERDICT IN VARCHAR2, P_BATTERY_HEALTH IN NUMBER, P_BATTERY_CYCLES IN NUMBER,
    P_SSD_LIFE IN NUMBER, P_FAIL_ITEMS IN VARCHAR2, P_WARN_ITEMS IN VARCHAR2,
    P_NOT_TESTED_ITEMS IN VARCHAR2, P_TECHNICIAN IN VARCHAR2, P_ORIGIN IN VARCHAR2,
    P_TOOL_VERSION IN VARCHAR2, P_STAMP_UTC IN VARCHAR2, P_RESULT_JSON IN CLOB,
    P_PDF_FILENAME IN VARCHAR2, P_PDF_BLOB IN BLOB,
    P_ID_OUT OUT NUMBER, P_ID_EQUIPAMENTO_OUT OUT NUMBER
) IS
BEGIN
    P_ID_EQUIPAMENTO_OUT := FUNC_RESOLVE_EQUIPAMENTO(P_INTERNAL_UID, P_SERVICE_TAG);

    SELECT SEQ_DIAGNOSTICO.NEXTVAL INTO P_ID_OUT FROM DUAL;

    INSERT INTO TBL_DIAGNOSTICO
               (ID_DIAGNOSTICO, ID_EQUIPAMENTO, INTERNAL_UID, SERVICE_TAG, MANUFACTURER,
                MODEL, CPU, RAM_GIB, STORAGE, VERDICT, BATTERY_HEALTH, BATTERY_CYCLES,
                SSD_LIFE, FAIL_ITEMS, WARN_ITEMS, NOT_TESTED_ITEMS, TECHNICIAN, ORIGIN,
                TOOL_VERSION, STAMP_UTC, DATA_TESTE, RESULT_JSON, PDF_FILENAME, PDF_BLOB)
        VALUES (P_ID_OUT, P_ID_EQUIPAMENTO_OUT,
                SUBSTR(NULLIF(TRIM(P_INTERNAL_UID), ''), 1, 80),
                SUBSTR(NULLIF(UPPER(TRIM(P_SERVICE_TAG)), ''), 1, 100),
                SUBSTR(NULLIF(TRIM(P_MANUFACTURER), ''), 1, 100),
                SUBSTR(NULLIF(TRIM(P_MODEL), ''), 1, 150),
                SUBSTR(NULLIF(TRIM(P_CPU), ''), 1, 200),
                P_RAM_GIB,
                SUBSTR(NULLIF(TRIM(P_STORAGE), ''), 1, 500),
                UPPER(TRIM(P_VERDICT)),
                P_BATTERY_HEALTH, P_BATTERY_CYCLES, P_SSD_LIFE,
                SUBSTR(NULLIF(TRIM(P_FAIL_ITEMS), ''), 1, 1000),
                SUBSTR(NULLIF(TRIM(P_WARN_ITEMS), ''), 1, 1000),
                SUBSTR(NULLIF(TRIM(P_NOT_TESTED_ITEMS), ''), 1, 1000),
                SUBSTR(NULLIF(TRIM(P_TECHNICIAN), ''), 1, 100),
                NULLIF(UPPER(TRIM(P_ORIGIN)), ''),
                SUBSTR(NULLIF(TRIM(P_TOOL_VERSION), ''), 1, 30),
                SUBSTR(NULLIF(TRIM(P_STAMP_UTC), ''), 1, 20),
                SYSDATE, P_RESULT_JSON,
                SUBSTR(NULLIF(TRIM(P_PDF_FILENAME), ''), 1, 200),
                P_PDF_BLOB);
END;

-- ----------------------------------------------------------------------------
PROCEDURE PROC_LIST(
    P_INTERNAL_UID IN VARCHAR2, P_SERVICE_TAG IN VARCHAR2, P_VERDICT IN VARCHAR2,
    P_DATA_INI IN DATE, P_DATA_FIM IN DATE, P_CURSOR OUT SYS_REFCURSOR
) IS
    V_UID VARCHAR2(100);
    V_TAG VARCHAR2(120);
BEGIN
    V_UID := '%' || UPPER(TRIM(P_INTERNAL_UID)) || '%';
    V_TAG := '%' || UPPER(TRIM(P_SERVICE_TAG)) || '%';

    OPEN P_CURSOR FOR
        SELECT D.ID_DIAGNOSTICO, D.DATA_TESTE, D.INTERNAL_UID, D.SERVICE_TAG,
               D.MANUFACTURER, D.MODEL, D.CPU, D.RAM_GIB, D.STORAGE, D.VERDICT,
               D.BATTERY_HEALTH, D.BATTERY_CYCLES, D.SSD_LIFE,
               D.FAIL_ITEMS, D.WARN_ITEMS, D.NOT_TESTED_ITEMS,
               D.TECHNICIAN, D.ORIGIN, D.TOOL_VERSION, D.STAMP_UTC,
               D.PDF_FILENAME,
               CASE WHEN D.PDF_BLOB IS NULL THEN 0
                    ELSE DBMS_LOB.GETLENGTH(D.PDF_BLOB) END AS PDF_BYTES,
               D.ID_EQUIPAMENTO
          FROM TBL_DIAGNOSTICO D
         WHERE (TRIM(P_INTERNAL_UID) IS NULL OR UPPER(D.INTERNAL_UID) LIKE V_UID)
           AND (TRIM(P_SERVICE_TAG)  IS NULL OR UPPER(D.SERVICE_TAG)  LIKE V_TAG)
           AND (TRIM(P_VERDICT)      IS NULL OR D.VERDICT = UPPER(TRIM(P_VERDICT)))
           AND (P_DATA_INI IS NULL OR D.DATA_TESTE >= TRUNC(P_DATA_INI))
           AND (P_DATA_FIM IS NULL OR D.DATA_TESTE <  TRUNC(P_DATA_FIM) + 1)
         ORDER BY D.DATA_TESTE DESC, D.ID_DIAGNOSTICO DESC;
END;

-- ----------------------------------------------------------------------------
PROCEDURE PROC_GET_PDF(P_ID IN NUMBER, P_PDF_FILENAME OUT VARCHAR2, P_PDF_BLOB OUT BLOB) IS
BEGIN
    SELECT PDF_FILENAME, PDF_BLOB INTO P_PDF_FILENAME, P_PDF_BLOB
      FROM TBL_DIAGNOSTICO WHERE ID_DIAGNOSTICO = P_ID;
EXCEPTION
    WHEN NO_DATA_FOUND THEN
        P_PDF_FILENAME := NULL;
        P_PDF_BLOB     := NULL;
END;

-- ----------------------------------------------------------------------------
PROCEDURE PROC_GET_JSON(P_ID IN NUMBER, P_RESULT_JSON OUT CLOB) IS
BEGIN
    SELECT RESULT_JSON INTO P_RESULT_JSON
      FROM TBL_DIAGNOSTICO WHERE ID_DIAGNOSTICO = P_ID;
EXCEPTION
    WHEN NO_DATA_FOUND THEN
        P_RESULT_JSON := NULL;
END;


-- ----------------------------------------------------------------------------
PROCEDURE PROC_LIST_BY_SHIPMENT(P_ID_REMESSA IN NUMBER, P_CURSOR OUT SYS_REFCURSOR) IS
BEGIN
    OPEN P_CURSOR FOR
        SELECT ID_EQUIPAMENTO, INTERNAL_UID, SERVICE_TAG, MANUFACTURER, MODEL,
               ID_DIAGNOSTICO, VERDICT, BATTERY_HEALTH, SSD_LIFE, DATA_TESTE,
               PDF_FILENAME, PDF_BYTES
          FROM (SELECT E.ID_EQUIPAMENTO, E.INTERNAL_UID, E.SERIAL_NUMBER AS SERVICE_TAG,
                       E.MARCA AS MANUFACTURER, E.MODEL,
                       D.ID_DIAGNOSTICO, D.VERDICT, D.BATTERY_HEALTH, D.SSD_LIFE, D.DATA_TESTE,
                       D.PDF_FILENAME,
                       CASE WHEN D.PDF_BLOB IS NULL THEN 0
                            ELSE DBMS_LOB.GETLENGTH(D.PDF_BLOB) END AS PDF_BYTES,
                       ROW_NUMBER() OVER (PARTITION BY E.ID_EQUIPAMENTO
                                          ORDER BY D.DATA_TESTE DESC NULLS LAST,
                                                   D.ID_DIAGNOSTICO DESC NULLS LAST) AS RN
                  FROM TBL_REMESSA_ITEM RI
                  JOIN TBL_EQUIPAMENTO E ON E.ID_EQUIPAMENTO = RI.ID_EQUIPAMENTO
                  LEFT JOIN TBL_DIAGNOSTICO D
                         ON (D.ID_EQUIPAMENTO = E.ID_EQUIPAMENTO
                             OR (D.ID_EQUIPAMENTO IS NULL
                                 AND (LTRIM(UPPER(TRIM(D.INTERNAL_UID)), '0') = LTRIM(UPPER(TRIM(E.INTERNAL_UID)), '0')
                                      OR D.SERVICE_TAG = UPPER(TRIM(E.SERIAL_NUMBER)))))
                 WHERE RI.ID_REMESSA = P_ID_REMESSA)
         WHERE RN = 1
         ORDER BY LENGTH(LTRIM(INTERNAL_UID, '0')), LTRIM(INTERNAL_UID, '0');
END;

END PACK_DIAGNOSTICO;
/

-- ----------------------------------------------------------------------------
-- Verificacao final
-- ----------------------------------------------------------------------------
SELECT OBJECT_TYPE, OBJECT_NAME, STATUS
  FROM USER_OBJECTS
 WHERE OBJECT_NAME = 'PACK_DIAGNOSTICO'
 ORDER BY OBJECT_TYPE
/
