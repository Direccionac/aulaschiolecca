# Ambientes Chio Lecca — portal de ocupación 2026-2

Portal de una sola página con la ocupación de ambientes, aulas libres, alertas de aforo y secciones homologadas con Q10.
No contiene información económica: tarifas, costos y planilla se quedan en los archivos de Dirección Académica.

- **Archivo único:** `index.html`. No necesita servidor, base de datos ni compilación.
- **Fuente:** libro *Auditoría de Aulas 2026 · v10*, hoja PROGRAMACIÓN EDITABLE. Matrícula del export «Estudiantes» de Q10 del 05-09-2026 12:30.

---

## Publicar en GitHub Pages (10 minutos, sin instalar nada)

Todo se hace desde el navegador, con la cuenta de GitHub creada con gpacheco@chio-lecca.edu.pe.

1. Entra a **github.com** → botón **+** arriba a la derecha → **New repository**.
2. **Repository name:** `ambientes-chio-lecca`. Marca **Public**. No marques «Add a README». → **Create repository**.
3. En la pantalla siguiente, haz clic en **uploading an existing file**.
4. Arrastra `index.html` (y este `README.md` si quieres) → **Commit changes**.
5. Pestaña **Settings** → menú lateral **Pages**.
6. En *Build and deployment* → *Source*: **Deploy from a branch**. *Branch*: **main**, carpeta **/ (root)** → **Save**.
7. Espera 1–2 minutos y recarga la página de Pages. Aparecerá la dirección:

   `https://TU-USUARIO.github.io/ambientes-chio-lecca/`

Esa es la dirección que se comparte con Gerencia. Es pública: cualquiera con el enlace la ve, no requiere cuenta ni permisos.

### Si prefieres que no sea pública

GitHub Pages en repositorios privados solo funciona con plan de pago. Dos alternativas gratuitas:

- **Netlify Drop** (netlify.com/drop): arrastras la carpeta y te da un enlace al instante. Permite proteger con contraseña en el plan gratuito por tiempo limitado.
- **Cloudflare Pages**: sube la carpeta y activa *Cloudflare Access* para restringir por correo institucional.

---

## Actualizar el portal cuando cambie el horario o la matrícula

El portal es una foto de los datos del libro. Para refrescarlo:

1. Actualiza el libro **Auditoría de Aulas** como siempre.
2. Pide la regeneración del `index.html` con el libro nuevo.
3. En GitHub: entra al repositorio → clic en `index.html` → icono del lápiz → borra el contenido, pega el nuevo → **Commit changes**.
   O bien **Add file → Upload files** y sube el `index.html` nuevo encima.

El sitio se actualiza solo, uno o dos minutos después del commit.

---

## Qué ve el usuario

| Vista | Para qué sirve |
|---|---|
| **Ocupación** | Rejilla ambiente × turno para el día elegido. Cada bloque abre una ficha con docente, grupo Q10, alumnos y aforo. |
| **Aulas libres** | Qué ambientes no tienen dictado, por día y turno, más el catálogo con aforo y capacidad física. |
| **Alertas** | Bloques sobre aforo, cruces de ambiente y bloques desactivados. |
| **Secciones** | Las secciones homologadas con Q10, con su grupo o campaña y sus matriculados. |

El buscador filtra por sección, docente, curso o aula en todas las vistas.
Los ambientes marcados en magenta son excepciones que requieren decisión; el cian indica uso normal.

## Colores corporativos

`--gris-tinta #2B2F36` · `--gris-medio #787F89` · `--gris-linea #DCE0E5` · `--gris-fondo #F2F4F6` · `--cian #00A3C4` · `--magenta #C6007E` · blanco.
Están definidos como variables CSS al inicio de `index.html`; cambiar un valor ahí cambia todo el sitio.
