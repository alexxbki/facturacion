# Facturación

Aplicación web privada para llevar **facturas emitidas, gastos y conciliación bancaria** dentro de tu propio Google Workspace, pensada para autónomos y pequeñas sociedades españolas con pocas facturas al mes.

No hay servidor ni base de datos que mantener: es un proyecto de Google Apps Script. Los datos se guardan en una hoja de cálculo de tu Drive y los PDF y justificantes, en carpetas de tu Drive. Nada sale de tu cuenta, salvo que actives el autocompletado opcional con Gemini.

## Qué hace

- **Facturas**: borradores, emisión con numeración correlativa por serie y año, IVA por línea, retención de IRPF, PDF con logo y facturas rectificativas.
- **Gastos**: registro con justificante adjunto e IVA deducible o no. Opcionalmente, lectura automática del justificante con Gemini.
- **Banco**: importación de extractos en Excel, CSV o Norma 43 y conciliación con facturas y gastos por importe, fecha y concepto.
- **Trimestres**: resumen de IVA repercutido y soportado, retenciones y exportación de libros en CSV.
- **Varias empresas**, dos roles (administrador y asistente) y copia de seguridad diaria automática.

## Aviso importante

Este proyecto se ofrece **tal cual, sin garantía de ningún tipo**. No es un programa de facturación homologado:

- **No cumple Verifactu** (Real Decreto 1007/2023) **ni genera factura electrónica**. Antes de usarlo para emitir facturas, comprueba con tu asesoría si tu actividad está obligada y desde qué fecha.
- Los resúmenes fiscales son orientativos y no sustituyen a una gestoría.
- Cada persona que lo instala es responsable de su uso, de sus datos y del cumplimiento de sus obligaciones fiscales.

## Requisitos

- Una cuenta de Google. Está pensado para **Google Workspace con dominio propio**; con una cuenta de Gmail personal funciona para un solo usuario.
- Unos 10 minutos para instalarlo. No hace falta saber programar: se copian y pegan tres archivos.

| Archivo | Qué es |
|---|---|
| `Code.gs` | Lógica del servidor (datos, numeración, PDF, banco, copias) |
| `Index.html` | La interfaz |
| `appsscript.json` | Configuración: zona horaria de Madrid y acceso restringido a tu dominio |

## Instalación (unos 10 minutos)

1. Entra en <https://script.google.com> con la cuenta de Google que será la propietaria de los datos y pulsa **Nuevo proyecto**. Ponle de nombre «Facturación».
2. En el archivo que aparece (`Código.gs`), borra todo y pega el contenido de **`Code.gs`**.
3. Pulsa **+ → HTML**, llámalo exactamente **`Index`** y pega el contenido de **`Index.html`** (sustituyendo lo que haya).
4. Ve a **Configuración del proyecto** (el engranaje) y marca *Mostrar el archivo de manifiesto «appsscript.json» en el editor*. Vuelve al editor, abre `appsscript.json` y sustituye su contenido por el de **`appsscript.json`**.
5. Guarda. En el desplegable de funciones elige **`setup`** y pulsa **Ejecutar**. Google te pedirá autorizar el acceso a Drive y Hojas de cálculo: acepta.
   En el *Registro de ejecución* verás los enlaces a la carpeta y a la hoja de datos creadas.
6. Pulsa **Implementar → Nueva implementación**, pulsa el engranaje junto a «Seleccionar tipo» y elige **Aplicación web**:
   - *Ejecutar como*: **Yo**
   - *Quién tiene acceso*: **Cualquier usuario de tu dominio** (con una cuenta de Gmail personal, elige **Solo yo**)
   
   Pulsa **Implementar** y copia la URL. Esa es tu aplicación: guárdala en favoritos o añádela a la pantalla de inicio del móvil.
7. Abre la URL y crea la empresa emisora. Después sube el logo desde **Ajustes**.

## Qué crea `setup` en tu Drive

```
Facturación/
├── Facturación · Datos          ← hoja de cálculo (la base de datos)
├── Facturas emitidas/<empresa>/<año>/T1…T4/   ← PDF de cada factura
├── Gastos/<empresa>/<año>/T1…T4/              ← justificantes
├── Copias de seguridad/                       ← copia diaria de la hoja
└── Logos/
```

`setup` se puede volver a ejecutar sin riesgo: no borra ni duplica nada.

## Seguridad

- Solo se entra con sesión iniciada en Google, así que aplica la verificación en dos pasos de tu cuenta.
- Aunque la URL esté abierta a tu dominio, la aplicación solo responde a los correos de la lista **`ALLOWED_EMAILS`** (al instalar, solo el tuyo).
  Para dar acceso a alguien más: *Configuración del proyecto → Propiedades de la secuencia de comandos → `ALLOWED_EMAILS`* y añade su correo separado por comas.
- La aplicación se ejecuta con tu cuenta: los demás usuarios autorizados no necesitan (ni deben tener) acceso directo a la hoja ni a las carpetas.
- No compartas la carpeta «Facturación» de Drive con nadie que no deba ver los datos.

## Roles: administrador y asistente

| | Administrador (`ALLOWED_EMAILS`) | Asistente (`ASSISTANT_EMAILS`) |
|---|---|---|
| Ver todo | Sí | Sí |
| Crear clientes, facturas y gastos; emitir facturas | Sí | Sí |
| Importar extractos, conciliar, marcar cobrado o pagado, ignorar movimientos | Sí | Sí |
| Editar o eliminar clientes y gastos; rectificar facturas | Sí | No |
| Deshacer cobros, pagos y conciliaciones; recuperar o eliminar movimientos | Sí | No |
| Empresas, logo, copias de seguridad | Sí | No |

- Para dar de alta a un asistente: *Configuración del proyecto → Propiedades de la secuencia de comandos*, añade **`ASSISTANT_EMAILS`** con su correo (varios, separados por comas) y pásale la URL. No hay que volver a desplegar.
- Donde el administrador ve «Editar» o «Eliminar», el asistente ve **Pedir cambio**: escribe qué hay que corregir y la solicitud aparece en el Panel del administrador, con un aviso en el menú. El administrador hace el cambio y pulsa **Hecho**; el asistente ve el estado en su Panel.
- El asistente sí puede terminar de editar un borrador de factura, porque todavía no es un documento emitido.
- Las restricciones se aplican en el servidor, no solo ocultando botones.
- Para que el asistente pueda abrir los PDF y los justificantes, comparte con su cuenta, como **Lector**, las carpetas «Facturas emitidas» y «Gastos» de Drive. No compartas la hoja de datos.

## Copias de seguridad

- Cada noche, entre las 3:00 y las 4:00, se copia la hoja de datos a «Copias de seguridad».
- Se conservan las 30 últimas copias diarias y una mensual durante 24 meses. Las que sobran van a la papelera de Drive.
- Desde **Ajustes** puedes hacer una copia en cualquier momento.
- Los PDF y justificantes no se duplican: ya están en Drive, con su historial de versiones y su papelera.

**Para restaurar una copia:** abre la copia que quieras en «Copias de seguridad», haz *Archivo → Hacer una copia*, copia el identificador de la nueva hoja (lo que va entre `/d/` y `/edit` en su URL) y pégalo en la propiedad `SPREADSHEET_ID` (*Configuración del proyecto → Propiedades de la secuencia de comandos*).

## Uso

- **Facturas**: se guardan como borrador y, al **emitir**, reciben número definitivo (`F2026-0001`) y se genera el PDF. Una factura emitida ya no se puede modificar ni borrar; se corrige con una **rectificativa** (serie `R`). La numeración es correlativa por serie y año, sin huecos, y no deja emitir con fecha anterior a la última factura.
- **Gastos**: con Gemini activado, los datos se rellenan solos al adjuntar el justificante. También puedes escribir la base o directamente el total (calcula la base), adjuntar el PDF o la foto y marcar si el IVA es deducible.
- **Banco**: descarga el extracto de tu banco en Excel, CSV o Norma 43 e impórtalo. Detecta las columnas solo; si falla, las eliges en el propio diálogo. Los movimientos ya importados se omiten, así que los extractos se pueden solapar.
  - *Coincide*: mismo importe y además el número de factura o el nombre aparece en el concepto. Se concilian solos al importar.
  - *Posible*: mismo importe pero sin más pistas. Se confirma con un clic.
  - Lo demás: **Buscar…** para elegir el documento, **Crear gasto** (comisiones, recibos sin registrar) o **Ignorar** (traspasos, impuestos…).
- **Trimestres**: IVA repercutido y soportado por tipo, retenciones de IRPF y exportación de los libros en CSV para la gestoría.

## Autocompletado de gastos con Gemini (opcional)

Al adjuntar el PDF o la foto de una factura recibida, la aplicación la envía a Gemini y rellena proveedor, NIF, número, fecha, concepto, categoría, base, IVA y retención. Nada se guarda hasta que pulsas **Guardar**.

Para activarlo:

1. Crea una clave de API en <https://aistudio.google.com/apikey>.
2. En el proyecto de Apps Script: *Configuración del proyecto → Propiedades de la secuencia de comandos → Añadir propiedad*, con nombre **`GEMINI_API_KEY`** y la clave como valor.
3. En el editor, elige la función **`probarGemini`** y pulsa **Ejecutar**. Google pedirá un permiso nuevo (conectarse a un servicio externo): acéptalo. En el registro debe aparecer «Gemini responde correctamente».

Detalles:

- Modelo por defecto: `gemini-3.8-flash`. Para usar otro, añade la propiedad `GEMINI_MODEL`.
- La clave se guarda solo en el servidor y nunca llega al navegador.
- El documento sale de tu Workspace hacia la API de Gemini. Con una clave del nivel gratuito, Google puede usar los datos enviados para mejorar sus modelos; con facturación activada en el proyecto de la clave, no. Compruébalo en las condiciones vigentes de la API.
- Si la factura tiene varios tipos de IVA o los importes no cuadran, aparece un aviso en rojo: cada gasto admite un solo tipo de IVA.

## Actualizar el código

Si cambias `Code.gs` o `Index.html`: *Implementar → Gestionar implementaciones → lápiz → Versión: Nueva versión → Implementar*. La URL no cambia.

## Límites conocidos

- No lleva contabilidad (asientos, balances) ni conexión directa con el banco.
- Los cobros y pagos parciales no se dividen: un movimiento se concilia con un solo documento.
- El resumen trimestral es orientativo: no contempla prorrata, recargo de equivalencia, criterio de caja, inversión del sujeto pasivo ni operaciones intracomunitarias.
- No genera registros Verifactu ni factura electrónica (Facturae). Comprueba los plazos vigentes de ambas obligaciones antes de usarlo para facturar.
- Pensado para volúmenes pequeños (cientos de facturas al año). Con miles de filas se volvería lento.

## Privacidad

El código no envía datos a ningún sitio salvo a los servicios de Google de tu propia cuenta. Las claves, los correos autorizados y los identificadores de tus archivos se guardan en las propiedades de tu proyecto de Apps Script, nunca en el código.

## Licencia

MIT. Consulta el archivo [LICENSE](LICENSE).
