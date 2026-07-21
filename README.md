# Subvantia — App funcional (prototipo real)

Esta es una aplicación web real (no una maqueta): tiene backend en Flask, base de
datos propia (SQLite) y el motor de cruce se ejecuta automáticamente cada vez que
se importa una cartera de clientes. Ya viene precargada con las 20 convocatorias
reales verificadas y una cartera de ejemplo de 15 clientes para que puedas probarla
sin hacer nada más.

## Cómo arrancarla en tu ordenador

Necesitas tener Python 3.10+ instalado.

```bash
cd subvantia_app
pip install -r requirements.txt
python app.py
```

Abre `http://localhost:5000` en el navegador.

**Hay dos tipos de cuenta, con permisos distintos:**

- **Cuenta de gestoría** (lo que vería un cliente que paga la suscripción): `demo@subvantia.test` / `subvantia2026`
- **Cuenta de administrador** (solo para Subvantia / Pablo): `admin@subvantia.test` / `subvantia2026admin`

### Por qué existen dos roles

Una gestoría puede ver la base de convocatorias (con enlace directo a la fuente
oficial de cada una) e importar y gestionar su propia cartera de clientes, pero
**no puede añadir ni editar convocatorias**. Esa base de datos compartida —
buscada, verificada y mantenida al día por Subvantia — es precisamente lo que
una gestoría paga mensualmente por usar. Si cualquier gestoría pudiera
modificarla libremente, no habría diferencia entre pagar la suscripción y no
pagarla. Solo la cuenta `admin` puede añadir o editar convocatorias, y cada
cambio recalcula automáticamente el cruce con todas las carteras de todas las
gestorías.

## Qué puedes hacer ya

- **Dashboard**: KPIs en vivo (clientes, convocatorias activas, dinero potencial, matches).
- **Clientes**: importar tu propio Excel de cartera (o descargar la plantilla desde la propia pantalla). Al importar, el cruce se recalcula automáticamente para los clientes nuevos.
- **Convocatorias**: las 20 convocatorias reales cargadas, con su estado, su nivel de verificación y un enlace directo a la fuente oficial de cada una. Solo la cuenta admin puede añadir o editar convocatorias.
- **Cruce**: todos los resultados, filtrables por banda (alta/media/baja) o por nombre de cliente. Desde ahí puedes generar el email personalizado de cada match y crear un expediente.
- **Expedientes**: kanban con las 5 fases (detectada → informada → en preparación → presentada → resuelta). Cambiar el estado desde el desplegable actualiza la base de datos al instante.
- **Exportar**: descarga de los resultados del cruce en Excel, con el mismo formato que ya conoces.

## Desplegarla en Render (para que otros la usen desde internet)

Ya está preparada para producción: usa `gunicorn` como servidor (no el servidor de
desarrollo de Flask), respeta el puerto que le asigne el hosting (variable `PORT`)
y desactiva el modo debug si se define `SUBVANTIA_DEBUG=0`. Incluye un `Procfile`
(`web: gunicorn app:app`) que Render reconoce automáticamente.

Pasos resumidos (ver la guía paso a paso `Subvantia_Guia_Despliegue.docx` para el
detalle sin dar nada por hecho):
1. Sube esta carpeta a un repositorio de GitHub.
2. En Render: New → Web Service → conectar ese repositorio.
3. Build command: `pip install -r requirements.txt`. Start command: `gunicorn app:app`.
4. Variable de entorno recomendada: `SUBVANTIA_SECRET` con una clave larga cualquiera.
5. Deploy. Render da una URL pública tipo `https://subvantia.onrender.com`.

**Importante sobre los datos en el plan gratuito de Render**: el disco no es
permanente y se borra en cada reinicio (el servicio se "duerme" tras 15 minutos
sin tráfico y se despierta —reinicia— en la siguiente visita, así que pasa a
menudo, no solo en despliegues). Para que las convocatorias añadidas a mano y
las carteras de clientes reales no se pierdan, la app ya soporta un segundo
modo de persistencia, sin coste:

- Si no defines la variable de entorno `DATABASE_URL`, usa SQLite local (el
  fichero `subvantia_data.db`) — cómodo para desarrollo, pero efímero en Render.
- Si defines `DATABASE_URL` apuntando a una base PostgreSQL (p. ej. un proyecto
  gratuito de [Neon](https://neon.tech)), la app guarda ahí todos los datos de
  forma permanente, sin coste. No hace falta tocar ni una línea de código: basta
  con añadir esa variable de entorno en Render.

## Seguimiento de pagos de las gestorías

Desde la pestaña "Gestorías" (solo admin) cada gestoría tiene un estado de
suscripción: **activa**, **impago** o **pausada**. Cambiarlo es instantáneo
(un desplegable, sin recargar formularios largos) y no borra nada: si una
gestoría deja de pagar un mes, la marcas como "impago" y no podrá iniciar
sesión (verá un aviso pidiéndole que contacte contigo), pero su cartera de
clientes y sus matches se quedan guardados intactos. En cuanto vuelva a pagar,
la marcas otra vez como "activa" y recupera el acceso exactamente donde lo
dejó. El rol `admin` nunca se ve afectado por este control.

Esto es seguimiento **manual** — de momento no hay cobro automático ni
recordatorios de pago integrados. Para automatizarlo cuando tengas varias
gestorías de pago, lo natural es conectar un procesador de pagos con
suscripciones recurrentes (Stripe Billing, o GoCardless si prefieres
domiciliación SEPA, muy habitual en el sector). Eso implica crear tú la cuenta
en el procesador elegido (KYC, cuenta bancaria) — es una decisión de negocio y
una cuenta que debes abrir tú mismo — pero una vez la tengas, la app ya está
lista para conectar ese estado de suscripción a los webhooks de pago en vez de
cambiarlo a mano.

## Qué NO incluye todavía (y qué haría falta para tenerlo)

- **Cobro automático de suscripciones**: ver la sección de arriba. Requiere conectar Stripe Billing o GoCardless.
- **Alta de nuevas gestorías desde la propia app sin pasar por el admin**: hoy solo tú (admin) puedes dar de alta una gestoría desde la pestaña "Gestorías". Un formulario de auto-registro público (con pago incluido en el propio alta) sería el siguiente paso si quieres que las gestorías se den de alta solas.
- **Envío real de emails**: el email se genera y se muestra en pantalla para copiar y pegar, pero no se envía solo. Conectar un proveedor de email (p. ej. SendGrid o el propio Gmail vía SMTP) es un paso posterior.
- **Checklist de documentación por convocatoria**: existe como generador en Excel (`checklist_y_kanban.py`) pero todavía no está integrado dentro de la app.
- **Persistencia de datos en producción**: ver el aviso de Render arriba — para el piloto real conviene pasar a disco persistente o Postgres antes de cargar datos reales de una gestoría.

## Estructura del proyecto

```
subvantia_app/
  app.py            -> rutas y lógica de la aplicación (Flask)
  database.py       -> esquema SQLite, siembra inicial de convocatorias/clientes demo
  matching.py        -> motor de cruce (mismas reglas que el prototipo: CNAE 40 / CCAA 25 / forma jurídica 15 / antigüedad 15 / plazo 5)
  templates/         -> las pantallas (HTML con Jinja2)
  static/style.css    -> estilos, mismos colores de marca que el resto de materiales
  requirements.txt
  subvantia_data.db  -> la base de datos (se genera sola al arrancar si no existe)
```
