-- ============================================================================
-- GBS Inventory - TBL_BACKUP_LOG + SEQ_BACKUP_LOG
-- Historico de execucoes do backup_gbs_weekly.ps1 (manual ou agendado).
-- O proprio script grava aqui ao final de cada execucao; a aba "Backup"
-- do app so le esta tabela para exibir o grid de historico.
-- ============================================================================

CREATE SEQUENCE SEQ_BACKUP_LOG START WITH 1 INCREMENT BY 1 NOCACHE NOCYCLE;

CREATE TABLE TBL_BACKUP_LOG (
    ID_BACKUP_LOG  NUMBER        NOT NULL,
    DATA_EXECUCAO  DATE          DEFAULT SYSDATE NOT NULL,
    ARQUIVO_DMP    VARCHAR2(100),
    STATUS         VARCHAR2(20)  NOT NULL,   -- SUCCESS, ERROR
    MENSAGEM       VARCHAR2(4000),
    CONSTRAINT PK_BACKUP_LOG    PRIMARY KEY (ID_BACKUP_LOG),
    CONSTRAINT CK_BACKUP_STATUS CHECK (STATUS IN ('SUCCESS','ERROR'))
);
