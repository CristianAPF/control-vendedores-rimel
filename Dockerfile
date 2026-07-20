# Actualización v2026.07.20.5

Este paquete es el repositorio completo listo para reemplazar el contenido de la rama `main`.

## Seguridad de datos
- No incluye `rimel.db`.
- No modifica `DATABASE_URL`.
- No ejecuta `DROP TABLE`, `DELETE` ni reinicios de esquema.
- `db.create_all()` conserva tablas y datos existentes.
- Las ampliaciones del módulo de quejas se aplican con `ALTER TABLE ... ADD COLUMN` solo si falta la columna.

## Verificación
Después del despliegue, abrir `/health`. Debe mostrar la versión `2026.07.20.5` y la zona `America/Montevideo`.
