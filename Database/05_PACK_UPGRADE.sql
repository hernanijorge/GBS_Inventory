-- ============================================================================
-- GBS Inventory - PACK_UPGRADE
-- Schema: colunas numéricas (RAM_NOVA_GB, STORAGE_NOVO_GB).
-- Os SELECTs derivam VALUE_BEFORE/VALUE_AFTER por TIPO para alinhar
-- com os nomes esperados pelo frmHistoricoEquipamento.
-- ============================================================================

CREATE OR REPLACE PACKAGE PACK_UPGRADE AS

    PROCEDURE PROC_INSERT_UPGRADE(
        P_ID_EQUIPAMENTO      IN NUMBER,
        P_TIPO_UPGRADE        IN VARCHAR2,
        P_RAM_ANTERIOR_GB     IN NUMBER,
        P_RAM_NOVA_GB         IN NUMBER,
        P_STORAGE_ANTERIOR_GB IN NUMBER,
        P_STORAGE_NOVO_GB     IN NUMBER,
        P_TECNICO             IN VARCHAR2,
        P_OBSERVACAO          IN VARCHAR2
    );

    PROCEDURE PROC_SELECT_UPGRADES_EQUIP(
        V_ID_EQUIPAMENTO IN  NUMBER,
        V_CURSOR         OUT SYS_REFCURSOR
    );

    PROCEDURE PROC_SELECT_UPGRADES_RECENTES(
        V_DIAS   IN  NUMBER,
        V_CURSOR OUT SYS_REFCURSOR
    );

    PROCEDURE PROC_DELETE_UPGRADE(
        V_ID IN NUMBER
    );

END PACK_UPGRADE;
/

CREATE OR REPLACE PACKAGE BODY PACK_UPGRADE AS

    -- -------------------------------------------------------------------------

    PROCEDURE PROC_INSERT_UPGRADE(
        P_ID_EQUIPAMENTO      IN NUMBER,
        P_TIPO_UPGRADE        IN VARCHAR2,
        P_RAM_ANTERIOR_GB     IN NUMBER,
        P_RAM_NOVA_GB         IN NUMBER,
        P_STORAGE_ANTERIOR_GB IN NUMBER,
        P_STORAGE_NOVO_GB     IN NUMBER,
        P_TECNICO             IN VARCHAR2,
        P_OBSERVACAO          IN VARCHAR2
    ) IS
    BEGIN
        INSERT INTO TBL_EQUIPAMENTO_UPGRADE (
            ID_UPGRADE, ID_EQUIPAMENTO, TIPO_UPGRADE,
            RAM_ANTERIOR_GB, RAM_NOVA_GB,
            STORAGE_ANTERIOR_GB, STORAGE_NOVO_GB,
            TECNICO, OBSERVACAO, DATA_UPGRADE
        ) VALUES (
            SEQ_EQUIPAMENTO_UPGRADE.NEXTVAL, P_ID_EQUIPAMENTO, P_TIPO_UPGRADE,
            P_RAM_ANTERIOR_GB, P_RAM_NOVA_GB,
            P_STORAGE_ANTERIOR_GB, P_STORAGE_NOVO_GB,
            P_TECNICO, P_OBSERVACAO, SYSDATE
        );

        UPDATE TBL_EQUIPAMENTO
           SET RAM_GB           = CASE WHEN P_RAM_NOVA_GB     > 0 THEN P_RAM_NOVA_GB     ELSE RAM_GB     END,
               STORAGE_GB       = CASE WHEN P_STORAGE_NOVO_GB > 0 THEN P_STORAGE_NOVO_GB ELSE STORAGE_GB END,
               STATUS           = CASE WHEN STATUS IN ('IN_STOCK', 'AVAILABLE') THEN 'IN_REPAIR' ELSE STATUS END,
               DATA_ATUALIZACAO = SYSDATE
         WHERE ID_EQUIPAMENTO = P_ID_EQUIPAMENTO;
    END;

    -- -------------------------------------------------------------------------
    -- Retorna upgrades de um equipamento com colunas alinhadas ao frmHistorico.
    -- VALUE_BEFORE/VALUE_AFTER são derivados por TIPO_UPGRADE:
    --   RAM  → RAM_ANTERIOR_GB / RAM_NOVA_GB
    --   SSD, HDD → STORAGE_ANTERIOR_GB / STORAGE_NOVO_GB
    --   outros   → NULL (campo ainda não mapeado no schema antigo)
    -- -------------------------------------------------------------------------
    PROCEDURE PROC_SELECT_UPGRADES_EQUIP(
        V_ID_EQUIPAMENTO IN  NUMBER,
        V_CURSOR         OUT SYS_REFCURSOR
    ) IS
    BEGIN
        OPEN V_CURSOR FOR
            SELECT
                ID_UPGRADE,
                ID_EQUIPAMENTO,
                DATA_UPGRADE,
                TIPO_UPGRADE                                                       AS COMPONENT_TYPE,
                CASE TIPO_UPGRADE
                    WHEN 'RAM' THEN CASE WHEN RAM_ANTERIOR_GB     > 0 THEN TO_CHAR(RAM_ANTERIOR_GB)     || ' GB' END
                    WHEN 'SSD' THEN CASE WHEN STORAGE_ANTERIOR_GB > 0 THEN TO_CHAR(STORAGE_ANTERIOR_GB) || ' GB' END
                    WHEN 'HDD' THEN CASE WHEN STORAGE_ANTERIOR_GB > 0 THEN TO_CHAR(STORAGE_ANTERIOR_GB) || ' GB' END
                    ELSE NULL
                END                                                                AS VALUE_BEFORE,
                CASE TIPO_UPGRADE
                    WHEN 'RAM' THEN CASE WHEN RAM_NOVA_GB     > 0 THEN TO_CHAR(RAM_NOVA_GB)     || ' GB' END
                    WHEN 'SSD' THEN CASE WHEN STORAGE_NOVO_GB > 0 THEN TO_CHAR(STORAGE_NOVO_GB) || ' GB' END
                    WHEN 'HDD' THEN CASE WHEN STORAGE_NOVO_GB > 0 THEN TO_CHAR(STORAGE_NOVO_GB) || ' GB' END
                    ELSE NULL
                END                                                                AS VALUE_AFTER,
                CAST(NULL AS VARCHAR2(30))                                         AS SOURCE_ORIGEM,
                CAST(NULL AS NUMBER(10,2))                                         AS COST_USD,
                CAST(NULL AS VARCHAR2(80))                                         AS PART_SERIAL,
                TECNICO                                                            AS TECHNICIAN,
                OBSERVACAO                                                         AS NOTES
              FROM TBL_EQUIPAMENTO_UPGRADE
             WHERE ID_EQUIPAMENTO = V_ID_EQUIPAMENTO
             ORDER BY DATA_UPGRADE DESC;
    END;

    -- -------------------------------------------------------------------------
    -- Upgrades recentes (últimos V_DIAS dias), mesma lógica de derivação.
    -- -------------------------------------------------------------------------
    PROCEDURE PROC_SELECT_UPGRADES_RECENTES(
        V_DIAS   IN  NUMBER,
        V_CURSOR OUT SYS_REFCURSOR
    ) IS
    BEGIN
        OPEN V_CURSOR FOR
            SELECT
                ID_UPGRADE,
                ID_EQUIPAMENTO,
                DATA_UPGRADE,
                TIPO_UPGRADE                                                       AS COMPONENT_TYPE,
                CASE TIPO_UPGRADE
                    WHEN 'RAM' THEN CASE WHEN RAM_ANTERIOR_GB     > 0 THEN TO_CHAR(RAM_ANTERIOR_GB)     || ' GB' END
                    WHEN 'SSD' THEN CASE WHEN STORAGE_ANTERIOR_GB > 0 THEN TO_CHAR(STORAGE_ANTERIOR_GB) || ' GB' END
                    WHEN 'HDD' THEN CASE WHEN STORAGE_ANTERIOR_GB > 0 THEN TO_CHAR(STORAGE_ANTERIOR_GB) || ' GB' END
                    ELSE NULL
                END                                                                AS VALUE_BEFORE,
                CASE TIPO_UPGRADE
                    WHEN 'RAM' THEN CASE WHEN RAM_NOVA_GB     > 0 THEN TO_CHAR(RAM_NOVA_GB)     || ' GB' END
                    WHEN 'SSD' THEN CASE WHEN STORAGE_NOVO_GB > 0 THEN TO_CHAR(STORAGE_NOVO_GB) || ' GB' END
                    WHEN 'HDD' THEN CASE WHEN STORAGE_NOVO_GB > 0 THEN TO_CHAR(STORAGE_NOVO_GB) || ' GB' END
                    ELSE NULL
                END                                                                AS VALUE_AFTER,
                CAST(NULL AS VARCHAR2(30))                                         AS SOURCE_ORIGEM,
                CAST(NULL AS NUMBER(10,2))                                         AS COST_USD,
                CAST(NULL AS VARCHAR2(80))                                         AS PART_SERIAL,
                TECNICO                                                            AS TECHNICIAN,
                OBSERVACAO                                                         AS NOTES
              FROM TBL_EQUIPAMENTO_UPGRADE
             WHERE DATA_UPGRADE >= SYSDATE - V_DIAS
             ORDER BY DATA_UPGRADE DESC;
    END;

    -- -------------------------------------------------------------------------

    PROCEDURE PROC_DELETE_UPGRADE(
        V_ID IN NUMBER
    ) IS
    BEGIN
        DELETE FROM TBL_EQUIPAMENTO_UPGRADE
         WHERE ID_UPGRADE = V_ID;
    END;

END PACK_UPGRADE;
/

-- Verificação pós-compilação
SELECT OBJECT_TYPE, STATUS, LAST_DDL_TIME
  FROM USER_OBJECTS
 WHERE OBJECT_NAME = 'PACK_UPGRADE'
 ORDER BY OBJECT_TYPE;

SELECT LINE, POSITION, TEXT
  FROM USER_ERRORS
 WHERE NAME = 'PACK_UPGRADE'
 ORDER BY SEQUENCE;
