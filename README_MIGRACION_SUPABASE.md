# Migración a Supabase (para que la app vuele 🚀)

Esta versión mantiene tu app igual (misma UI), pero el dataset deja de vivir en Google Sheets.
Ahora los datos están en **Supabase (Postgres)** y Apps Script queda como “puente” (backend) rápido.

---

## Qué vas a tener al final

- **Supabase**: base de datos rápida (tablas: estudiantes, catálogo, estado por ciclo, etc.)
- **Apps Script Web App**: backend con las mismas acciones que tu frontend ya usa.
- **GitHub Pages**: tu frontend (index/app/styles) igual, pero con una opción para pegar la URL del backend desde la pantalla de ingreso (sin editar código).

---

## Paso 1 — Crear el proyecto Supabase

1. Entrá a https://supabase.com/ y creá un proyecto.
2. Cuando termine, andá a:
   - **SQL Editor** → “New query”

---

## Paso 2 — Crear tablas (copiar y pegar SQL)

1. Abrí el archivo `supabase_setup.sql`
2. Copiá TODO el contenido
3. Pegalo en el SQL Editor de Supabase
4. Ejecutá (“Run”)

---

## Paso 3 — Importar tus datos (CSV)

En Supabase:
1. **Table Editor** → abrís una tabla (por ejemplo `estudiantes`)
2. Botón **Import data** (o “Import CSV”, depende de la UI)
3. Subís el CSV correspondiente:

- `estudiantes.csv` → tabla `estudiantes`
- `materias_catalogo.csv` → tabla `materias_catalogo`
- `estado_por_ciclo.csv` → tabla `estado_por_ciclo`
- (opcional) `egresados.csv` → tabla `egresados`
- `auditoria.csv` está vacío a propósito (puede importarse o no)

⚠️ Importá primero `estudiantes` y `materias_catalogo` antes que `estado_por_ciclo`.

---

## Paso 4 — Backend (Apps Script)

1. Entrá a https://script.google.com/
2. Creá un proyecto nuevo.
3. Abrí el archivo `Code.gs` del ZIP y **reemplazá TODO** el contenido del archivo del proyecto por este `Code.gs`.
4. En Apps Script:
   - **Project Settings** → “Script properties”
   - Agregá estas 3 propiedades:

**TRAYECTORIAS_API_KEY**
- poné cualquier clave que quieras (ej: una frase larga).
- Esa misma clave la vas a pegar luego en la app (pantalla “Ingresá la clave de acceso”).

**SUPABASE_URL**
- la URL del proyecto (ej: `https://xxxx.supabase.co`)

**SUPABASE_SERVICE_KEY**
- la “service_role key”:
  - Supabase → Settings → API → “service_role”

⚠️ IMPORTANTE: la service_role key es secreta. Está bien acá porque queda guardada en Apps Script, no en el frontend.

---

## Paso 5 — Deploy del Web App (URL /exec)

1. En Apps Script:
   - Deploy → **New deployment**
   - Type: **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
2. Deploy
3. Copiá la URL que termina en `/exec`

---

## Paso 6 — Frontend (GitHub Pages)

En tu repo (GitHub Pages), reemplazá estos archivos por los del ZIP:

- `index.html`
- `app.js`
- `styles.css`
- `config.js`

👉 Con esta versión **NO necesitás editar** la URL en código:
- Abrís tu app
- En la pantalla de la API Key, abrís **“Configurar backend”**
- Pegás la URL `/exec` del paso anterior
- Guardás ✅

---

## Paso 7 — Probar

1. Abrís la app
2. Pegás:
   - URL del backend (/exec) (solo 1 vez)
   - API Key (la que pusiste en Script properties)
3. Entrás y listo.

---

## Si algo falla

- Error “Falta configurar la URL del backend”:
  - Abrí “Configurar backend” y pegá la URL /exec.

- Error “No autorizado: API Key inválida”:
  - La API Key del frontend no coincide con `TRAYECTORIAS_API_KEY` del Apps Script.

- Error “Supabase error 401/403”:
  - Revisá `SUPABASE_URL` y `SUPABASE_SERVICE_KEY`.

---
