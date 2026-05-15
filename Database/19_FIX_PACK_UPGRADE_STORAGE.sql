-- ============================================================================
-- GBS Inventory - Migration 19
-- Corrige PROC_INSERT_UPGRADE: STORAGE_GB nao era atualizado quando
-- o novo valor era 0 (ex: SSD extraido da maquina).
--
-- Causa: UPDATE usava > 0 como condicao, bloqueando valor 0.
-- Fix:   IS NOT NULL — o VB agora envia NULL para campos nao aplicaveis
--        e o valor real (incluindo 0) para o campo do tipo em questao.
-- ============================================================================

CREATE OR REPLACE PACKAGE BODY PACK_UPGRADE AS

    PROCEDURE PROC_INSERT_UPGRADE(
        P_ID_EQUIPAMENTO      IN NUMBER,
        P_TIPO_UPGRADE        IN VARCHAR2,
        P_RAM_ANTERIOR_GB     IN NUMBER,
        P_RAM_NOVA_GB         IN NUMBER,
        P_STORAGE_ANTERIOR_GB IN NUMBER,
        P_STORAGE_NOVO_GB     IN NUMBER,
        P_TECNICO             IN VARCHAR2,
        P_OBSERVACAO          IN VARCHAR2,
        P_ID_COMPONENT        IN NUMBER,
        P_ACTION_TYPE         IN VARCHAR2,
        P_COMP_NEW_STATUS     IN VARCHAR2
    ) IS
    BEGIN
        INSERT INTO TBL_EQUIPAMENTO_UPGRADE (
            ID_UPGRADE, ID_EQUIPAMENTO, TIPO_UPGRADE,
            RAM_ANTERIOR_GB, RAM_NOVA_GB,
            STORAGE_ANTERIOR_GB, STORAGE_NOVO_GB,
            TECNICO, OBSERVACAO, DATA_UPGRADE,
            ID_COMPONENT, ACTION_TYPE
        ) VALUES (
            SEQ_EQUIPAMENTO_UPGRADE.NEXTVAL, P_ID_EQUIPAMENTO, P_TIPO_UPGRADE,
            P_RAM_ANTERIOR_GB, P_RAM_NOVA_GB,
            P_STORAGE_ANTERIOR_GB, P_STORAGE_NOVO_GB,
            P_TECNICO, P_OBSERVACAO, SYSDATE,
            P_ID_COMPONENT, P_ACTION_TYPE
        );

        UPDATE TBL_EQUIPAMENTO
           SET RAM_GB           = CASE WHEN P_RAM_NOVA_GB     IS NOT NULL THEN P_RAM_NOVA_GB     ELSE RAM_GB     END,
               STORAGE_GB       = CASE WHEN P_STORAGE_NOVO_GB IS NOT NULL THEN P_STORAGE_NOVO_GB ELSE STORAGE_GB END,
               STATUS           = CASE WHEN STATUS IN ('IN_STOCK', 'AVAILABLE') THEN 'IN_REPAIR' ELSE STATUS END,
               DATA_ATUALIZACAO = SYSDATE
         WHERE ID_EQUIPAMENTO = P_ID_EQUIPAMENTO;

        IF P_ID_COMPONENT IS NOT NULL THEN
            IF P_ACTION_TYPE = 'INSTALL' THEN
                UPDATE TBL_COMPONENT SET STATUS = 'INSTALLED' WHERE ID_COMPONENT = P_ID_COMPONENT;
            ELSIF P_ACTION_TYPE = 'REMOVE' THEN
                UPDATE TBL_COMPONENT SET STATUS = NVL(P_COMP_NEW_STATUS, 'IN_STOCK') WHERE ID_COMPONENT = P_ID_COMPONENT;
            END IF;
        END IF;
    END;

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

    PROCEDURE PROC_SELECT_UPGRADES_RECENTES(
        V_DIAS   IN  NUMBER,
        V_CURSOR OUT SYS_REFCURSOR
    ) IS
    BEGIN
        OPEN V_CURSOR FOR
            SELECT
                U.ID_UPGRADE,
                E.INTERNAL_UID,
                U.DATA_UPGRADE,
                U.TIPO_UPGRADE                                                          AS COMPONENT_TYPE,
                CASE U.TIPO_UPGRADE
                    WHEN 'RAM' THEN CASE WHEN U.RAM_ANTERIOR_GB     > 0 THEN TO_CHAR(U.RAM_ANTERIOR_GB)     || ' GB' END
                    WHEN 'SSD' THEN CASE WHEN U.STORAGE_ANTERIOR_GB > 0 THEN TO_CHAR(U.STORAGE_ANTERIOR_GB) || ' GB' END
                    WHEN 'HDD' THEN CASE WHEN U.STORAGE_ANTERIOR_GB > 0 THEN TO_CHAR(U.STORAGE_ANTERIOR_GB) || ' GB' END
                    ELSE NULL
                END                                                                     AS VALUE_BEFORE,
                CASE U.TIPO_UPGRADE
                    WHEN 'RAM' THEN CASE WHEN U.RAM_NOVA_GB     > 0 THEN TO_CHAR(U.RAM_NOVA_GB)     || ' GB' END
                    WHEN 'SSD' THEN CASE WHEN U.STORAGE_NOVO_GB > 0 THEN TO_CHAR(U.STORAGE_NOVO_GB) || ' GB' END
                    WHEN 'HDD' THEN CASE WHEN U.STORAGE_NOVO_GB > 0 THEN TO_CHAR(U.STORAGE_NOVO_GB) || ' GB' END
                    ELSE NULL
                END                                                                     AS VALUE_AFTER,
                CAST(NULL AS VARCHAR2(30))                                              AS SOURCE_ORIGEM,
                CAST(NULL AS NUMBER(10,2))                                              AS COST_USD,
                CAST(NULL AS VARCHAR2(80))                                              AS PART_SERIAL,
                U.TECNICO                                                               AS TECHNICIAN,
                U.OBSERVACAO                                                            AS NOTES
              FROM TBL_EQUIPAMENTO_UPGRADE U
              JOIN TBL_EQUIPAMENTO E ON E.ID_EQUIPAMENTO = U.ID_EQUIPAMENTO
             WHERE U.DATA_UPGRADE >= SYSDATE - V_DIAS
             ORDER BY U.DATA_UPGRADE DESC;
    END;

    PROCEDURE PROC_DELETE_UPGRADE(
        V_ID IN NUMBER
    ) IS
    BEGIN
        DELETE FROM TBL_EQUIPAMENTO_UPGRADE
         WHERE ID_UPGRADE = V_ID;
    END;

END PACK_UPGRADE;
/

SELECT OBJECT_TYPE, STATUS FROM USER_OBJECTS WHERE OBJECT_NAME = 'PACK_UPGRADE' ORDER BY OBJECT_TYPE;
SELECT LINE, TEXT FROM USER_ERRORS WHERE NAME = 'PACK_UPGRADE' ORDER BY SEQUENCE;
