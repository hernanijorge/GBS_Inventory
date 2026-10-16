-- ============================================================================
-- GBS Inventory - Migration 34
-- TBL_REMESSA.ID_CLIENTE (FK -> TBL_CLIENTE, nullable): cliente da remessa.
-- Definido na aba Shipments (combo Customer + "Set Customer"); o e-mail dos
-- "Test Certificates" vem desse cliente. Nao altera status nem PACK_REMESSA.
-- Idempotente.
-- ============================================================================

DECLARE
    V_N NUMBER;
BEGIN
    SELECT COUNT(*) INTO V_N FROM USER_TAB_COLUMNS
     WHERE TABLE_NAME = 'TBL_REMESSA' AND COLUMN_NAME = 'ID_CLIENTE';
    IF V_N = 0 THEN
        EXECUTE IMMEDIATE 'ALTER TABLE TBL_REMESSA ADD (ID_CLIENTE NUMBER)';
        DBMS_OUTPUT.PUT_LINE('OK: TBL_REMESSA.ID_CLIENTE criada');
    ELSE
        DBMS_OUTPUT.PUT_LINE('SKIP: TBL_REMESSA.ID_CLIENTE ja existe');
    END IF;

    SELECT COUNT(*) INTO V_N FROM USER_CONSTRAINTS WHERE CONSTRAINT_NAME = 'FK_REMESSA_CLIENTE';
    IF V_N = 0 THEN
        EXECUTE IMMEDIATE 'ALTER TABLE TBL_REMESSA ADD CONSTRAINT FK_REMESSA_CLIENTE ' ||
                          'FOREIGN KEY (ID_CLIENTE) REFERENCES TBL_CLIENTE (ID_CLIENTE)';
        DBMS_OUTPUT.PUT_LINE('OK: FK_REMESSA_CLIENTE criada');
    END IF;

    SELECT COUNT(*) INTO V_N FROM USER_INDEXES WHERE INDEX_NAME = 'IDX_REMESSA_CLIENTE';
    IF V_N = 0 THEN
        EXECUTE IMMEDIATE 'CREATE INDEX IDX_REMESSA_CLIENTE ON TBL_REMESSA (ID_CLIENTE)';
        DBMS_OUTPUT.PUT_LINE('OK: IDX_REMESSA_CLIENTE criado');
    END IF;
END;
/

-- ----------------------------------------------------------------------------
-- Verificacao final
-- ----------------------------------------------------------------------------
SELECT C.COLUMN_NAME, C.DATA_TYPE, C.NULLABLE,
       (SELECT STATUS FROM USER_CONSTRAINTS WHERE CONSTRAINT_NAME = 'FK_REMESSA_CLIENTE') AS FK_STATUS
  FROM USER_TAB_COLUMNS C
 WHERE C.TABLE_NAME = 'TBL_REMESSA' AND C.COLUMN_NAME = 'ID_CLIENTE'
/
