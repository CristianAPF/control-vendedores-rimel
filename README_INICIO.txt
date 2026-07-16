CONTROL DE VENDEDORES RIMEL - INICIO LOCAL CORREGIDO

1. Descomprima completamente el ZIP.
2. Ejecute INICIAR_APLICACION.bat.
3. Espere a que se abra http://127.0.0.1:8000/login

USUARIO GERENCIAL
Usuario: gerencia
Contraseña: Rimel2026!

VENDEDORES
gerson / Gerson2026!
eduardo / Eduardo2026!
victoria / Victoria2026!

CORRECCION DE ESTA VERSION
- La instalacion local ya no intenta instalar PostgreSQL ni psycopg2.
- Utiliza SQLite, que no requiere componentes externos.
- Verifica e instala las dependencias aunque la carpeta .venv ya exista.

SI QUEDO UNA INSTALACION INCOMPLETA
Ejecute LIMPIAR_Y_REINSTALAR.bat.
Ese archivo elimina solamente .venv. No borra rimel.db.

IMPORTANTE
Para uso desde diferentes celulares fuera de la misma red, la aplicacion debe publicarse en un servidor HTTPS con base central PostgreSQL. El archivo requirements-production.txt queda reservado para esa etapa.
