# Gestión de morosos — puesta en marcha

Esto ya no guarda nada en tu navegador: guarda todo en una base de datos
compartida (Supabase) y se publica como página web (Vercel). Cualquiera
que entre al link ve y carga sobre los mismos datos, en tiempo real.

Las dos partes son gratis en el uso de un club. Vas a necesitar:
una cuenta de [supabase.com](https://supabase.com) y una de
[vercel.com](https://vercel.com) — con Google o GitHub alcanza para las dos.

No hace falta que sepas programar para seguir estos pasos, pero sí que
sigas el orden: la base de datos primero, la página después.

---

## Paso 1 — Crear la base de datos (Supabase)

1. Entrá a [supabase.com](https://supabase.com) → **Start your project** → creá una cuenta.
2. **New project**. Ponele un nombre (ej. "club-morosos"), elegí una
   contraseña para la base — guardala, no la vas a necesitar de nuevo pero
   por las dudas — y la región más cercana (**South America (São Paulo)**
   si está disponible). Esperá 1-2 minutos a que se cree.
3. En el menú de la izquierda, andá a **SQL Editor** → **New query**.
4. Abrí el archivo `esquema.sql` de esta carpeta, copiá todo su contenido,
   pegalo ahí, y apretá **Run**. Esto crea las tablas donde vive todo:
   saldos, gestión, historial y comentarios.
5. Andá a **Project Settings** (el engranaje) → **API**. Vas a necesitar
   dos datos de esta pantalla en el paso 3:
   - **Project URL**
   - **anon public** (la clave pública, no la `service_role`)

Con esto la base ya existe, pero todavía está vacía. Se llena sola la
primera vez que subas el archivo de saldos desde la página.

---

## Paso 2 — Subir el proyecto a GitHub

Vercel despliega leyendo un repositorio de GitHub. Si nunca usaste GitHub:

1. Creá una cuenta en [github.com](https://github.com) si no tenés.
2. Creá un repositorio nuevo (**New repository**), privado, con el nombre
   que quieras (ej. `gestion-morosos`). No marques ninguna opción de
   inicializarlo con archivos.
3. Subí el contenido de esta carpeta a ese repositorio. Si no usás git
   desde la terminal, la forma más simple es arrastrar todos los archivos
   de esta carpeta a la página del repositorio en GitHub
   ("uploading an existing file") — **excepto** `node_modules` y `.env`
   si llegaste a crearlos, esos no se suben nunca.

---

## Paso 3 — Publicar la página (Vercel)

1. Entrá a [vercel.com](https://vercel.com) → **Add New** → **Project**.
2. Elegí **Continue with GitHub** y autorizá el acceso si te lo pide.
3. Buscá el repositorio que subiste (`gestion-morosos`) y hacé **Import**.
   Vercel detecta solo que es un proyecto Vite; no hace falta tocar nada
   en "Build settings".
4. Antes de desplegar, abrí **Environment Variables** y cargá las dos
   claves que copiaste en el paso 1:

   | Name | Value |
   |---|---|
   | `VITE_SUPABASE_URL` | tu Project URL de Supabase |
   | `VITE_SUPABASE_ANON_KEY` | tu clave anon public de Supabase |

5. **Deploy**. En un minuto te da un link (algo como
   `gestion-morosos.vercel.app`). Ese es el link que compartís con la
   comisión y con quien vaya a cargar reclamos.

Listo. La primera vez que alguien entre va a ver la pantalla de
"Subir archivo de saldos" — eso pasa una sola vez, para todos: en cuanto
uno lo suba, ya va a estar cargado para el resto.

---

## Cómo se actualiza después

- **Cambios en el diseño o el código**: si volvés a pedirme ajustes, te
  paso el `App.jsx` actualizado. Lo reemplazás en tu repositorio de
  GitHub (subiendo el archivo nuevo encima del viejo) y Vercel vuelve a
  publicar solo, automáticamente, en 1-2 minutos.
- **Saldos del mes**: se actualizan desde el botón del menú
  "Datos y archivos → Actualizar saldos", dentro de la página. No hace
  falta tocar nada acá.

---

## Qué es gratis y hasta dónde

- **Vercel** (plan Hobby): gratis para este uso, sin límite de tiempo.
- **Supabase** (plan Free): gratis, con 500 MB de base de datos —
  muchísimo más de lo que un club con algunos cientos de socios va a usar
  — y el proyecto se pausa solo si nadie lo usa durante 7 días seguidos.
  Si eso pasa, entrás una vez a supabase.com y lo reactivás con un botón;
  no se pierde nada.

---

## Qué falta para que sea un sistema con usuarios reales

Ahora mismo el botón "Secretaría / Comisión" arriba a la derecha es sólo
una forma de mostrar la pantalla distinta — **no es un login**. Cualquiera
con el link puede tocar cualquier botón, incluido borrar la gestión.

Eso es aceptable para empezar (es el mismo nivel de confianza que tenían
con el Excel compartido por mail), pero si más adelante querés que cada
persona entre con su usuario y contraseña, y que sólo la secretaría pueda
editar, avisame: se agrega con Supabase Auth y es un paso más, no hay que
volver a armar nada de lo que ya está.
