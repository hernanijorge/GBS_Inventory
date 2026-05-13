-- ============================================================================
-- GBS Inventory - Migration 17
-- Link upgrades to TBL_COMPONENT for stock deduction.
-- Adds ID_COMPONENT and ACTION_TYPE to TBL_EQUIPAMENTO_UPGRADE.
-- ============================================================================

ALTER TABLE TBL_EQUIPAMENTO_UPGRADE ADD ID_COMPONENT NUMBER;
ALTER TABLE TBL_EQUIPAMENTO_UPGRADE ADD ACTION_TYPE  VARCHAR2(10);

-- Optional FK (comment out if TBL_COMPONENT does not exist yet)
-- ALTER TABLE TBL_EQUIPAMENTO_UPGRADE
--   ADD CONSTRAINT FK_UPGRADE_COMPONENT
--   FOREIGN KEY (ID_COMPONENT) REFERENCES TBL_COMPONENT (ID_COMPONENT);

-- Verification
SELECT COLUMN_NAME, DATA_TYPE, DATA_LENGTH, NULLABLE
  FROM USER_TAB_COLUMNS
 WHERE TABLE_NAME = 'TBL_EQUIPAMENTO_UPGRADE'
 ORDER BY COLUMN_ID;
