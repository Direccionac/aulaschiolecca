# Auditoría de Ambientes — Instituto Chio Lecca · Ciclo 2026-II

Plataforma de una sola página para la Dirección Académica: ocupación de aulas, carga docente,
detección de cruces y reprogramación de dictados con aprobación.

- **`index.html`** — la plataforma completa. Sin servidor, sin compilación, sin dependencias.
- **`Codigo.gs`** — el servicio de Google Apps Script que guarda los cambios en una hoja de Google.
- **`isotipo.png`** — el isotipo que aparece en la cabecera (opcional).

No contiene información económica. El libro de Auditoría de Aulas sigue siendo el registro oficial.

---

## Qué hace cada pestaña

| Pestaña | Para qué sirve |
|---|---|
| **Panorama** | Mapa de calor de la semana: 18 ambientes × 18 turnos. Intensidad de cian = ocupación; magenta = cruce. Clic en una celda abre ese día. |
| **Grilla del día** | Ambientes en filas, horas en columnas a escala real. Clic en un bloque para ver su ficha o reprogramarlo. |
| **Ambientes** | Catálogo con aforo, capacidad física, dictados y horas de uso por semana. |
| **Aulas libres** | Matriz día × turno, más un buscador de huecos por duración y número de alumnos. |
| **Docentes** | Carga semanal de los 64 docentes. Clic para ver el horario completo de uno. |
| **Alertas** | Cruces de ambiente, secciones sobre aforo, dictados sin ambiente y bajas. Se recalcula solo. |
| **Cambios** | Propuestas pendientes e historial con autor, fecha y motivo. |
| **Datos** | Conexión a la sala compartida y exportación a CSV. |
| **Guía** | Los doce pasos para alguien que entra nuevo al equipo. |

---

## Parte 1 · Publicar la página

El repositorio es `Direccionac/aulaschiolecca` y GitHub Pages ya está activo en `main / (root)`.
Para actualizar: **Add file → Upload files**, sube el `index.html` nuevo encima y **Commit changes**.
Uno o dos minutos después, `https://direccionac.github.io/aulaschiolecca/` sirve la versión nueva.

Sube también `isotipo.png` a la raíz si quieres el logo en la cabecera. Si el archivo no está, la
cabecera funciona igual.

---

## Parte 2 · Montar la sala compartida (15 minutos, una sola vez)

Sin esto la plataforma sirve para consultar, pero los cambios que proponga alguien no los ve nadie más.

1. Crea una **hoja de cálculo de Google** nueva. Llámala, por ejemplo, `Ambientes — Cambios 2026-II`.
2. Dentro de la hoja: **Extensiones → Apps Script**.
3. Borra lo que haya y pega el contenido de **`Codigo.gs`**. Guarda.
4. En el desplegable de funciones elige **`preparar`** y pulsa **Ejecutar**. Google pedirá autorizar; acepta.
   Esto crea la pestaña `Cambios` con sus encabezados.
5. Rueda dentada **Configuración del proyecto → Propiedades del script → Añadir propiedad**. Crea dos:

   | Propiedad | Valor |
   |---|---|
   | `CLAVE_ADMIN` | la clave de quien aprueba y rechaza |
   | `CLAVE_EDITOR` | la clave de quien solo propone |

   Estas claves viven en el servidor de Google, no en el `index.html`. Son las únicas que valen.
6. **Implementar → Nueva implementación → Aplicación web**:
   - *Ejecutar como*: **Yo**
   - *Quién tiene acceso*: **Cualquier usuario**

   Copia la URL que termina en **`/exec`**.
7. Abre el portal, ve a **Datos**, pega la URL en «Dirección del servicio» y pulsa **Conectar**.
   El punto de la cabecera se pone cian y dice «Sala conectada».

Cada persona pega la URL una vez en su navegador. Si prefieres que venga puesta de fábrica,
edita en `index.html` la línea `var URL_SALA = "";` y pon la URL entre las comillas.

### Editores

En `index.html`, al inicio del bloque `<script>`, está la lista `EDITORES`. Cambia correos y nombres
por los reales. El campo `rol` solo decide qué botones se muestran; quien manda de verdad es la clave
que valida el servidor.

---

## Cómo funciona la aprobación

1. Un editor abre un bloque en la grilla, cambia ambiente, día u hora, y escribe el motivo.
2. Antes de enviar, la ficha valida en vivo: ambiente ocupado, docente ya dictando a esa hora, aforo
   insuficiente. **Avisa, no bloquea.**
3. La propuesta llega a la hoja de Google con estado `pendiente`.
4. Quien tenga la `CLAVE_ADMIN` la aprueba o la rechaza desde la pestaña Cambios.
5. Al aprobarla, el cambio se aplica sobre la programación base y todas las vistas se recalculan:
   cruces, aulas libres, carga docente.
6. Todo queda en la hoja con autor, fecha, motivo y quién resolvió.

La plataforma sincroniza cada 15 segundos.

---

## Lo que hay que tener claro

**La página es pública.** Cualquiera con el enlace puede consultar la ocupación, los nombres de los
docentes y los números de matrícula. La clave protege la escritura, no la lectura. Si algún día eso
deja de servir, la ruta es Cloudflare Pages con acceso restringido al dominio institucional.

**La programación base viene incrustada.** Sale del libro de Auditoría de Aulas, hoja
PROGRAMACIÓN EDITABLE. Cuando cambie el horario madre hay que regenerar el `index.html` y volver a
subirlo. El historial de cambios no se pierde: vive en la hoja de Google, aparte.

**El registro oficial sigue siendo el Excel.** Esta plataforma sirve para decidir y dejar rastro.
Lo aprobado aquí hay que trasladarlo al libro.

---

## Colores y tipografía

Sistema editorial: fondo `#f3f2f2`, superficie `#eae9e9`, tinta `#201e1d`, cian `#0088b0`,
magenta `#d6006c`. Tipografía Source Serif 4 en 400 y 600. Todo está en las variables CSS del
`:root`, al inicio del archivo: cambiar un valor ahí cambia el sitio entero.
