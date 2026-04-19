-- ============================================================================
-- PACK_EQUIPAMENTO - Specification
-- Pacote de procedures e functions para equipamentos (laptops/desktops)
-- ============================================================================

CREATE OR REPLACE PACKAGE PACK_EQUIPAMENTO IS

    -- Function para gerar próximo ID
    FUNCTION FUNC_PROXIMO_EQUIPAMENTO
        RETURN NUMBER;

    -- Function para verificar se UID já existe
    FUNCTION FUNC_EXISTE_UID (V_INTERNAL_UID IN VARCHAR2)
        RETURN NUMBER;

    -- Function para retornar status em texto legível
    FUNCTION FUNC_STATUS_DESCRICAO (V_STATUS IN VARCHAR2)
        RETURN VARCHAR2;

    -- CRUD Equipamento
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
        V_ID                   OUT NUMBER);

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
        V_ID                   NUMBER);

    PROCEDURE PROC_UPDATE_STATUS (
        V_STATUS   VARCHAR2,
        V_ID       NUMBER);

    PROCEDURE PROC_DELETE_EQUIPAMENTO (V_ID IN NUMBER);

    -- Import em lote (usado pela importação de planilha)
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
        V_RESULTADO            OUT VARCHAR2);

    -- Consultas
    PROCEDURE PROC_SELECT_EQUIPAMENTO (
        V_ID      IN  NUMBER,
        V_CURSOR  OUT SYS_REFCURSOR);

    PROCEDURE PROC_SELECT_EQUIPAMENTO_UID (
        V_INTERNAL_UID IN  VARCHAR2,
        V_CURSOR       OUT SYS_REFCURSOR);

    PROCEDURE PROC_SELECT_EQUIPAMENTO_SERIAL (
        V_SERIAL IN  VARCHAR2,
        V_CURSOR OUT SYS_REFCURSOR);

    PROCEDURE PROC_SELECT_EQUIPAMENTOS (
        V_EMPRESA IN  NUMBER,
        V_CURSOR  OUT SYS_REFCURSOR);

    PROCEDURE PROC_SELECT_POR_FILTRO (
        V_PESQUISA     IN  VARCHAR2,
        V_MANUFACTURER IN  VARCHAR2,
        V_CONDITION    IN  VARCHAR2,
        V_STATUS       IN  VARCHAR2,
        V_CURSOR       OUT SYS_REFCURSOR);

    PROCEDURE PROC_DASHBOARD_TOTAIS (
        V_EMPRESA IN  NUMBER,
        V_CURSOR  OUT SYS_REFCURSOR);

    PROCEDURE PROC_RESUMO_WHATSAPP (
        V_EMPRESA IN  NUMBER,
        V_CURSOR  OUT SYS_REFCURSOR);

END PACK_EQUIPAMENTO;
/
