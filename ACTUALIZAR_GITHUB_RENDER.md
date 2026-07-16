# Actualizar GitHub y Render a V6

## 1. Actualizar GitHub

1. Descomprime el ZIP V6.
2. Entra al repositorio `control-vendedores-rimel`.
3. Pulsa **Add file** y luego **Upload files**.
4. Arrastra el contenido de la carpeta `rimel_v6`, no la carpeta completa y no el ZIP.
5. Confirma el reemplazo de los archivos existentes.
6. Pulsa **Commit changes**.

Deben quedar en la raíz del repositorio, entre otros:

- `app.py`
- `requirements.txt`
- `templates/`
- `static/`

No subas `.venv`, `rimel.db` ni `__pycache__`.

## 2. Actualizar Render

Render normalmente despliega automáticamente el nuevo commit. Si no lo hace:

1. Entra al servicio `vendedores-rimel`.
2. Pulsa **Manual Deploy**.
3. Elige **Deploy latest commit**.
4. Si continúa apareciendo la versión anterior, usa **Clear build cache & deploy**.

No cambies `DATABASE_URL`; es la conexión con la base central que ya contiene la información.

## 3. Verificar

Cuando el despliegue muestre **Live**:

1. Abre la URL pública.
2. Inicia sesión como Gerencia.
3. Debe aparecer **Clientes y rutas** en el menú.
4. Optimiza una ruta de prueba.
5. Ingresa con el vendedor correspondiente y selecciona ese día.
6. La lista debe aparecer con el orden optimizado guardado.
