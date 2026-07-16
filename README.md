# Control de Vendedores RIMEL — V6

## Novedades

- Las rutas optimizadas se guardan en la base central.
- Al seleccionar un día, el vendedor recibe directamente la secuencia guardada por Gerencia.
- Nueva sección **Clientes y rutas**, visible solamente para usuarios gerenciales.
- Gerencia puede agregar, modificar, reasignar, dar de baja y reactivar clientes.
- Gerencia puede optimizar cada combinación de vendedor y día desde el navegador.
- El orden optimizado queda disponible inmediatamente para todos los celulares.
- Los clientes dados de baja no se eliminan: se conserva su historial de visitas y quejas.

## Optimización inicial

La optimización necesita geocodificar las direcciones. Por esa razón Gerencia debe realizarla una vez para cada combinación de vendedor y día:

1. Ingresar con `gerencia`.
2. Abrir **Clientes y rutas**.
3. Seleccionar vendedor y día.
4. Confirmar el punto de partida.
5. Pulsar **Optimizar y guardar**.

Después de guardarla, el vendedor verá ese orden automáticamente. Cuando se agregue, cambie de día, cambie de vendedor o cambie de dirección un cliente, conviene volver a optimizar esa ruta.

## Actualización de la versión publicada

No crees otro servicio ni otra base de datos. Reemplaza los archivos del repositorio de GitHub y conserva las variables actuales de Render.

En Render deben mantenerse:

- Build Command: `pip install -r requirements.txt`
- Start Command: `gunicorn app:app`
- Variables: `DATABASE_URL`, `SECRET_KEY`, `PYTHON_VERSION`

La aplicación añade automáticamente a la base existente las columnas nuevas necesarias, sin borrar visitas, usuarios ni clientes.
