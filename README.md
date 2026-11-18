# Canje Mayorista AJE

Sitio: https://aje-canje.holafree.com

- `index.html` · portal con dos accesos.
- `app.html` · app de canjistas (un usuario por mercado + supervisor): asistencia con selfie y GPS, bebida del día, canjes (paquetes → drop → premio, con foto de boleta y del premio) e incidencias. Funciona sin señal: guarda en el celular y envía cuando vuelve la conexión.
- `dashboard.html` · dashboard de la campaña con el histórico de la hoja "IMPULSO Y CANJES MAYORISTA LIMA 2026" (17/06–08/11/2026).
- `config.js` · URL del backend, mercados, tabla de drops, bebidas. `API_URL` vacío = modo demo (usuarios `productores`, `caqueta`, `unicachi`, `huamantanga`, `supervisor` · clave `aje2026`).
- `apps-script/Code.gs` · backend (Google Sheet + Drive). Hojas: Canjes, Asistencia, Incidencias, Equipo (accesos), Drops (mecánica), Config.

## Poner en marcha el backend

1. script.google.com → Proyecto nuevo "Canje Mayorista AJE Backend" → pegar `Code.gs`.
2. Ejecutar `setup()` y aceptar permisos. Crea el Sheet, la carpeta de fotos y los accesos (pestaña Equipo).
3. Implementar → Nueva implementación → Aplicación web · Ejecutar como: yo · Acceso: cualquier persona.
4. Copiar la URL `/exec` en `config.js` → `API_URL`, subir `sw.js` con un CACHE nuevo y hacer push.
