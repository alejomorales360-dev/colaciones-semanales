# Colaciones Semanales

App para controlar la inscripción semanal de los trabajadores al menú de
colaciones: evita que alguien quede sin anotarse y evita anotaciones dobles.
Backend en Google Sheets + Google Apps Script (gratis, sin servidores
propios). Frontend estático (`index.html`), publicado con **GitHub Pages**
en vez de Netlify.

## Cómo funciona

- **Trabajadores** entran con su **RUT** (sin clave), ven el menú de la
  semana vigente y eligen una opción por día (o la quitan tocándola de
  nuevo). Pasado el plazo de cierre, queda bloqueado para ellos; solo
  administración puede seguir editando.
- **Administración** entra con una clave y gestiona trabajadores, arma el
  menú semanal, marca días de "almuerzo mejorado", configura el cierre de
  inscripciones, y tiene reportes (semanales y por rango de fechas) con
  exportación a Excel y PDF.

## 1. Backend: Google Sheets + Apps Script

1. Crea una planilla de Google Sheets nueva.
2. **Extensiones → Apps Script**, borra el contenido de ejemplo y pega el
   contenido de [`gas/Colaciones.gs`](gas/Colaciones.gs) de este repo.
3. Ejecuta una vez la función **`crearHojasIniciales`** (crea las hojas y
   una clave de admin por defecto: `cambiar123`).
4. Cambia `admin_password` en la hoja **Config** por una clave real.
5. Carga a tus trabajadores en la hoja **Trabajadores**.
6. **Implementar → Nueva implementación → Aplicación web** (Ejecutar como:
   Yo · Acceso: Cualquier persona) y copia la URL que termina en `/exec`.

## 2. Conectar el frontend

1. Abre `index.html` en este repo (botón de lápiz para editar, o clona el
   repo).
2. Busca la línea `const GAS_URL = '...'` y reemplázala por tu URL `/exec`.
3. Guarda los cambios (commit directo a `main`).

## 3. Publicar con GitHub Pages (gratis, sin límites de "créditos")

1. En este repo, ve a **Settings → Pages**.
2. En **Build and deployment → Source**, elige **Deploy from a branch**.
3. **Branch**: `main`, carpeta **/ (root)** → **Save**.
4. Espera 1-2 minutos. La app queda disponible en:
   `https://alejomorales360-dev.github.io/colaciones-semanales/`

Cada vez que edites `index.html` en este repo (o hagas push), GitHub Pages
actualiza el sitio solo, en un par de minutos — no requiere nada de
Netlify.

## Volver a desplegar el Apps Script tras editar `Colaciones.gs`

Editar el código no actualiza la URL `/exec` ya publicada. Hay que ir a
**Implementar → Gestionar implementaciones**, editar (ícono de lápiz) la
implementación de tipo "Aplicación web", elegir **Nueva versión** y
presionar **Implementar**. La URL `/exec` no cambia.
