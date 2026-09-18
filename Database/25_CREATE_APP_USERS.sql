-- ============================================================================
-- GBS Inventory - Migration 25
-- Cria a tabela de usuarios da aplicacao GBS_Web (login real via hash, substitui
-- a checagem de senha fixa em variavel de ambiente por PBKDF2 armazenado no
-- proprio Oracle). Puramente aditiva: nao toca em nenhuma tabela existente.
--
-- NOTA: este banco e Oracle 10g Express Edition (10.2.0.1.0), sem suporte a
-- "GENERATED ALWAYS AS IDENTITY" (Oracle 12c+). Segue o padrao ja usado pelas
-- demais tabelas do schema: SEQUENCE + NEXTVAL fornecido por quem insere
-- (nao ha triggers BEFORE INSERT em nenhuma TBL_ existente).
--
-- Idempotente: cria tabela/sequence so se ainda nao existirem.
-- ============================================================================

BEGIN
    EXECUTE IMMEDIATE '
        CREATE TABLE TBL_APP_USER (
            ID_USER         NUMBER PRIMARY KEY,
            USERNAME        VARCHAR2(100) NOT NULL,
            PASSWORD_HASH   VARCHAR2(500) NOT NULL,
            IS_ACTIVE       CHAR(1) DEFAULT ''Y'' NOT NULL CHECK (IS_ACTIVE IN (''Y'',''N'')),
            CREATED_AT      DATE DEFAULT SYSDATE NOT NULL,
            LAST_LOGIN_AT   DATE,
            CONSTRAINT UQ_APP_USER_USERNAME UNIQUE (USERNAME)
        )';
EXCEPTION
    WHEN OTHERS THEN
        IF SQLCODE != -955 THEN -- ORA-00955: name already used by an existing object
            RAISE;
        END IF;
END;
/

BEGIN
    EXECUTE IMMEDIATE 'CREATE SEQUENCE SEQ_APP_USER START WITH 1 INCREMENT BY 1';
EXCEPTION
    WHEN OTHERS THEN
        IF SQLCODE != -955 THEN
            RAISE;
        END IF;
END;
/

-- ----------------------------------------------------------------------------
-- Verificacao
-- ----------------------------------------------------------------------------
SELECT TABLE_NAME FROM USER_TABLES WHERE TABLE_NAME = 'TBL_APP_USER';
SELECT COLUMN_NAME, DATA_TYPE, NULLABLE FROM USER_TAB_COLUMNS
 WHERE TABLE_NAME = 'TBL_APP_USER' ORDER BY COLUMN_ID;
SELECT SEQUENCE_NAME FROM USER_SEQUENCES WHERE SEQUENCE_NAME = 'SEQ_APP_USER';
