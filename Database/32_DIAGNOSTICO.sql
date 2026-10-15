-- ============================================================================
-- GBS Inventory - Migration 32
-- TBL_DIAGNOSTICO + SEQ_DIAGNOSTICO + PACK_DIAGNOSTICO
-- Resultados do diagnostico de laptops (pen drive Ubuntu / PXE / Windows),
-- recebidos pelo GBS_DiagReceiver (POST /api/diagnostico) e exibidos na aba
-- "Diagnostics" do Desktop.
--
-- ATENCAO (Oracle XE 10.2 + ODP.NET Managed 21.x):
--   colunas LOB em SELECT/cursor falham no cliente com "TTC Error".
--   Por isso PROC_LIST nao retorna RESULT_JSON nem PDF_BLOB, e o PDF/JSON
--   sao lidos por parametro OUT (PROC_GET_PDF / PROC_GET_JSON), que funciona.
-- Idempotente (tabela/sequence/indices so sao criados se nao existirem).
-- ============================================================================

DECLARE
    V_N NUMBER;

    PROCEDURE CRIA_INDICE(P_NOME VARCHAR2, P_DDL VARCHAR2) IS
        V_X NUMBER;
    BEGIN
        SELECT COUNT(*) INTO V_X FROM USER_INDEXES WHERE INDEX_NAME = P_NOME;
        IF V_X = 0 THEN
            EXECUTE IMMEDIATE P_DDL;
            DBMS_OUTPUT.PUT_LINE('OK: ' || P_NOME || ' criado');
        END IF;
    END;
BEGIN
    SELECT COUNT(*) INTO V_N FROM USER_SEQUENCES WHERE SEQUENCE_NAME = 'SEQ_DIAGNOSTICO';
    IF V_N = 0 THEN
        EXECUTE IMMEDIATE 'CREATE SEQUENCE SEQ_DIAGNOSTICO START WITH 1 INCREMENT BY 1 NOCACHE NOCYCLE';
        DBMS_OUTPUT.PUT_LINE('OK: SEQ_DIAGNOSTICO criada');
    ELSE
        DBMS_OUTPUT.PUT_LINE('SKIP: SEQ_DIAGNOSTICO ja existe');
    END IF;

    SELECT COUNT(*) INTO V_N FROM USER_TABLES WHERE TABLE_NAME = 'TBL_DIAGNOSTICO';
    IF V_N = 0 THEN
        EXECUTE IMMEDIATE
            'CREATE TABLE TBL_DIAGNOSTICO (' ||
            '    ID_DIAGNOSTICO    NUMBER          NOT NULL,' ||
            '    ID_EQUIPAMENTO    NUMBER,' ||
            '    INTERNAL_UID      VARCHAR2(80),' ||
            '    SERVICE_TAG       VARCHAR2(100),' ||
            '    MANUFACTURER      VARCHAR2(100),' ||
            '    MODEL             VARCHAR2(150),' ||
            '    CPU               VARCHAR2(200),' ||
            '    RAM_GIB           NUMBER,' ||
            '    STORAGE           VARCHAR2(500),' ||
            '    VERDICT           VARCHAR2(20)    NOT NULL,' ||
            '    BATTERY_HEALTH    NUMBER,' ||
            '    BATTERY_CYCLES    NUMBER,' ||
            '    SSD_LIFE          NUMBER,' ||
            '    FAIL_ITEMS        VARCHAR2(1000),' ||
            '    WARN_ITEMS        VARCHAR2(1000),' ||
            '    NOT_TESTED_ITEMS  VARCHAR2(1000),' ||
            '    TECHNICIAN        VARCHAR2(100),' ||
            '    ORIGIN            VARCHAR2(20),' ||
            '    TOOL_VERSION      VARCHAR2(30),' ||
            '    STAMP_UTC         VARCHAR2(20),' ||
            '    DATA_TESTE        DATE            DEFAULT SYSDATE NOT NULL,' ||
            '    RESULT_JSON       CLOB,' ||
            '    PDF_FILENAME      VARCHAR2(200),' ||
            '    PDF_BLOB          BLOB,' ||
            '    CONSTRAINT PK_DIAGNOSTICO PRIMARY KEY (ID_DIAGNOSTICO),' ||
            '    CONSTRAINT FK_DIAG_EQUIPAMENTO FOREIGN KEY (ID_EQUIPAMENTO)' ||
            '        REFERENCES TBL_EQUIPAMENTO (ID_EQUIPAMENTO),' ||
            '    CONSTRAINT CK_DIAG_VERDICT CHECK (VERDICT IN' ||
            '        (''PASS'',''PASS_WITH_WARNINGS'',''FAIL'',''INCOMPLETE'')),' ||
            '    CONSTRAINT CK_DIAG_ORIGIN CHECK (ORIGIN IS NULL OR ORIGIN IN' ||
            '        (''UBUNTU-USB'',''UBUNTU-PXE'',''WINDOWS''))' ||
            ')';
        DBMS_OUTPUT.PUT_LINE('OK: TBL_DIAGNOSTICO criada');
    ELSE
        DBMS_OUTPUT.PUT_LINE('SKIP: TBL_DIAGNOSTICO ja existe');
    END IF;

    CRIA_INDICE('IDX_DIAG_EQUIPAMENTO', 'CREATE INDEX IDX_DIAG_EQUIPAMENTO ON TBL_DIAGNOSTICO (ID_EQUIPAMENTO)');
    CRIA_INDICE('IDX_DIAG_UID',         'CREATE INDEX IDX_DIAG_UID ON TBL_DIAGNOSTICO (INTERNAL_UID)');
    CRIA_INDICE('IDX_DIAG_TAG',         'CREATE INDEX IDX_DIAG_TAG ON TBL_DIAGNOSTICO (SERVICE_TAG)');
    CRIA_INDICE('IDX_DIAG_DATA',        'CREATE INDEX IDX_DIAG_DATA ON TBL_DIAGNOSTICO (DATA_TESTE)');
END;
/

-- ============================================================================
-- PACK_DIAGNOSTICO
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

END PACK_DIAGNOSTICO;
/

-- ----------------------------------------------------------------------------
-- Verificacao final
-- ----------------------------------------------------------------------------
SELECT OBJECT_TYPE, OBJECT_NAME, STATUS
  FROM USER_OBJECTS
 WHERE OBJECT_NAME IN ('TBL_DIAGNOSTICO', 'SEQ_DIAGNOSTICO', 'PACK_DIAGNOSTICO')
 ORDER BY OBJECT_TYPE
/
