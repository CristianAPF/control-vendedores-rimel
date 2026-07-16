# Control de Vendedores RIMEL — versión multiusuario

Esta versión usa una base de datos central. Los vendedores ingresan desde sus celulares y Gerencia ve los registros cargados por todos ellos.

## Usuarios iniciales

- Gerencia: `gerencia` / `Rimel2026!`
- Gerson: `gerson` / `Gerson2026!`
- Eduardo: `eduardo` / `Eduardo2026!`
- Victoria: `victoria` / `Victoria2026!`

Todos deben cambiar la contraseña en el primer ingreso.

## Funciones

- Inicio de sesión individual.
- Perfil vendedor: solo puede ver sus clientes, rutas, visitas y quejas.
- Perfil gerencial: ve todos los vendedores, rutas, KPI, historial y quejas.
- Creación de nuevos usuarios desde el perfil gerencial.
- Restablecimiento de contraseñas.
- Registro de visitas sin importe de venta.
- Resultados permitidos: pedido concretado, sin necesidad de reposición, cliente ausente, cliente cerrado, visita reprogramada y cliente inactivo.
- Registro independiente de quejas.
- Captura de GPS.
- Encabezados de tablas fijos.
- Aplicación instalable como PWA en Android y iPhone.
- Actualización central inmediata; el tablero gerencial consulta cambios automáticamente cada 30 segundos.

## Importante: cómo usarla en varios celulares

No debe abrirse como un archivo HTML. Debe publicarse en un servidor con HTTPS. Se incluyen Dockerfile y render.yaml para desplegarla en Render u otro proveedor compatible con Docker y PostgreSQL.

### Despliegue recomendado

1. Subir esta carpeta a un repositorio privado de GitHub.
2. Crear un servicio en Render usando el archivo `render.yaml`.
3. Esperar a que Render cree el servicio web y la base PostgreSQL.
4. Abrir la dirección HTTPS entregada por Render.
5. En cada celular, abrir la dirección y elegir “Agregar a pantalla de inicio” o “Instalar aplicación”.

## Prueba en una sola computadora y la misma red Wi-Fi

En Windows, ejecutar `run_local_windows.bat`. Luego consultar la IP local de la computadora y abrir en cada celular:

`http://IP-DE-LA-COMPUTADORA:8000`

Esta modalidad sirve para pruebas. Para GPS, instalación PWA y uso fuera de la red local se necesita HTTPS y un servidor publicado.

## Base de datos

- Producción: PostgreSQL mediante la variable `DATABASE_URL`.
- Pruebas locales: SQLite, creada automáticamente como `rimel.db`.

## Seguridad

Antes de publicar, configurar una variable `SECRET_KEY` segura. No compartir las contraseñas iniciales y exigir que cada usuario las cambie.
