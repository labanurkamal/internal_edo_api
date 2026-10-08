-- internal_edo_api: каким шаблоном .docx (код и версия) создан договор.
-- Повторный запуск безопасен.

ALTER TABLE lease_contracts ADD COLUMN IF NOT EXISTS template_code    VARCHAR(16);
ALTER TABLE lease_contracts ADD COLUMN IF NOT EXISTS template_version INTEGER;
