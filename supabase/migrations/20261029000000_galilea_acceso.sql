-- =============================================================================
-- Usuario de acceso fijo GALILEA (producción)
-- PBKDF2-HMAC-SHA256 · 260 000 iteraciones · sal por usuario
-- Idempotente: ON CONFLICT DO NOTHING
-- =============================================================================

INSERT INTO public.usuario
    (nombre, rol, hash, salt, iteraciones, activo, debe_cambiar, creado)
VALUES
    ('GALILEA', 'admin',
     '587b29182aab0632b20ec9527a60203d85186514c28d40345586dfe2aca457e6',
     '9f2b4c6e8a1d3f5b7c9e0a2d4b6f8c1a',
     260000, 1, 0,
     to_char(NOW(), 'YYYY-MM-DD HH24:MI:SS'))
ON CONFLICT DO NOTHING;
