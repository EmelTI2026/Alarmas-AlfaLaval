# Clariot — Informes de alarmas de vibración

Convierte las alertas de vibración de **Clariot / Alfa Laval**, que llegan por correo
en inglés y con un PDF adjunto, en **borradores de Outlook en español, listos para
revisar y enviar al cliente**.

Hacerlo a mano toma unos diez minutos por alerta: leerla, traducirla, agruparla con
las anteriores del mismo equipo, decidir si es urgente y redactar el correo. El
programa hace todo eso. La persona solo revisa, completa y envía.

> [!IMPORTANT]
> **El programa nunca envía correos.** Solo crea borradores. En el código no existe
> ninguna llamada a `Send` y no hay forma de configurarlo para que envíe. Una persona
> siempre lee y aprueba antes de que algo salga.

---

## Contenido

1. [Cómo funciona](#1-cómo-funciona)
2. [Inicio rápido](#2-inicio-rápido)
3. [Requisitos](#3-requisitos)
4. [Instalación paso a paso](#4-instalación-paso-a-paso)
5. [Configurar Outlook](#5-configurar-outlook)
6. [Primera prueba](#6-primera-prueba)
7. [Uso diario](#7-uso-diario)
8. [Comandos](#8-comandos)
9. [Cómo decide el programa](#9-cómo-decide-el-programa)
10. [Configuración](#10-configuración)
11. [Archivos que genera](#11-archivos-que-genera)
12. [Solución de problemas](#12-solución-de-problemas)
13. [Actualizar a una versión nueva](#13-actualizar-a-una-versión-nueva)
14. [Desarrollo](#14-desarrollo)
15. [Límites conocidos](#15-límites-conocidos)
16. [Otros documentos](#16-otros-documentos)

---

## 1. Cómo funciona

```
  Alfa Laval  (support@aliotportal.com)
       │  correo en inglés con un PDF
       ▼
  Regla de Outlook ──►  Bandeja de entrada / Clariot
       │
       ▼  doble clic en "Procesar alertas.bat"   (python -m clariot)
  ┌────────────────────────────────────────────────────────────┐
  │ 1. Lee el PDF y extrae los datos                           │
  │ 2. Descarta los correos y eventos que ya procesó           │
  │ 3. Agrupa con las alarmas anteriores del mismo equipo      │
  │ 4. Clasifica la gravedad: normal, urgente o crítica        │
  │ 5. Traduce las frases técnicas al español                  │
  │ 6. Redacta el informe con destinatarios, firma y PDF       │
  └────────────────────────────────────────────────────────────┘
       │
       ├──►  Borradores / Urgencias          urgente o crítico: enviar HOY
       ├──►  Borradores / Por enviar         rutina
       └──►  Bandeja de entrada / Procesados el correo original, archivado
```

| Tarea | Quién la hace |
| --- | --- |
| Detectar las alertas nuevas en la carpeta | Programa |
| Leer el PDF, agrupar, clasificar y traducir | Programa |
| Redactar el informe y adjuntar los PDF | Programa |
| Cargar destinatarios y copias de los clientes conocidos | Programa (desde `config/clients.yaml`) |
| Agregar la firma de Outlook | Programa |
| Dejar el borrador en `Urgencias` o `Por enviar` | Programa |
| Poner el destinatario si el cliente **no** está en `clients.yaml` | **Persona** |
| Reemplazar `[NOMBRE]` por el nombre del contacto | **Persona** |
| Escribir la recomendación de Emeltec donde dice `[COMPLETAR]` | **Persona** |
| Revisar y enviar | **Persona** |

El programa **no es un servicio** que vigila el buzón: se ejecuta, procesa lo que
encuentra y se cierra. Si nadie lo ejecuta, no se pierde nada. Las alertas esperan en
la carpeta y se procesan todas en la próxima corrida.

## 2. Inicio rápido

Esta es la versión corta, para quien ya tiene Python 3.12 y Outlook clásico. Si es la
primera vez, sigue la [sección 4](#4-instalación-paso-a-paso), donde cada paso viene
explicado y con su verificación.

Ejecuta todo en **CMD**, no en PowerShell ([por qué](#4-instalación-paso-a-paso)):

```cmd
cd /d C:\
git clone https://github.com/Nlira23/Automatizacion-Alarmas.git clariot
cd clariot
python -m venv .venv
.venv\Scripts\activate.bat
pip install -r requirements.txt
pip install -e .
copy config\env.example .env
notepad .env
python -m clariot --setup-folders
python -m clariot --self-check
```

En `.env` pega la clave de Gemini ([paso 5](#paso-5--la-clave-de-gemini)). Después
crea la regla de Outlook ([sección 5.2](#52-la-regla-de-outlook)) y haz la
[primera prueba](#6-primera-prueba).

## 3. Requisitos

| Requisito | Detalle |
| --- | --- |
| Sistema operativo | Windows 10 u 11 |
| Correo | **Outlook clásico de escritorio**, abierto y con la cuenta que recibe las alertas ya configurada |
| Python | 3.10 o superior. **Se recomienda 3.12**: es la versión probada con `pywin32` y `pymupdf` |
| Git | Opcional. Sin Git se puede descargar el ZIP del repositorio |
| Clave de IA | Una clave de [Google AI Studio](https://aistudio.google.com/apikey), gratuita y sin tarjeta, para traducir las frases |
| Sesión de Windows | Iniciada. El programa necesita el escritorio interactivo y no corre como servicio |

**El "nuevo Outlook" no sirve.** No expone COM, que es la interfaz por la que el
programa se comunica con el correo. Si arriba a la derecha de Outlook aparece un
interruptor "Nuevo Outlook", apágalo. Microsoft retira el Outlook clásico de las
suscripciones empresariales en 2028. Migrar a Microsoft Graph implica cambiar **un solo
archivo**: `src/clariot/adapters/outlook.py`.

**No hace falta** registrar una aplicación en Azure AD, ni permisos de administrador,
ni OAuth, ni Microsoft Graph. El programa usa la sesión de Outlook que ya está abierta
en el escritorio. Fue una restricción de diseño desde el inicio: en muchas empresas es
imposible conseguir un registro de aplicación.

### Dependencias

| Archivo | Qué instala | Cuándo hace falta |
| --- | --- | --- |
| `requirements.txt` | `pywin32` (Outlook por COM), `pymupdf` (lectura del PDF), `Jinja2` (plantillas del correo), `PyYAML`, `python-dotenv` | Siempre |
| `requirements-optional.txt` | `google-cloud-translate`, `deepl` | Solo si se activa la traducción del PDF adjunto (`translation.provider`) |
| `requirements-dev.txt` | `pytest`, `pytest-cov` | Solo para correr los tests |

`pymupdf` tiene licencia AGPL-3.0. Para uso interno no hay problema; se vuelve
relevante solo si la herramienta se distribuye o se ofrece como servicio a terceros.

## 4. Instalación paso a paso

> [!WARNING]
> **Usa CMD, no PowerShell.** Para abrirlo: `Windows + R`, escribe `cmd`, Enter.
>
> En PowerShell, `activate.bat` **falla en silencio**. Un `.bat` corre en un proceso
> hijo y muere sin devolverle las variables a PowerShell. No da error, no aparece
> `(.venv)` en el prompt, y los `pip install` siguientes se instalan en el Python del
> sistema sin avisar. Si necesitas PowerShell, el equivalente es
> `.\.venv\Scripts\Activate.ps1`.

Para instalar en el equipo de un técnico que no tiene nada instalado, usa
**[DESPLIEGUE.md](DESPLIEGUE.md)**: es este mismo camino, con checklist y con las
pantallas del instalador de Python explicadas una por una.

### Antes de empezar

| | Qué | Dónde se consigue |
| --- | --- | --- |
| 1 | **Outlook clásico** con la cuenta que recibe las alertas | Viene con Microsoft 365. Se verifica con `reg query HKEY_CLASSES_ROOT\Outlook.Application\CLSID` |
| 2 | **Acceso al repositorio** | Una cuenta de GitHub con permiso, o el ZIP descargado |
| 3 | **La clave de Gemini** | [aistudio.google.com/apikey](https://aistudio.google.com/apikey). **No viene en el repositorio** |
| 4 | **Un PDF de alerta real** | Para la primera prueba, si todavía no llegó ninguna alerta al buzón |

**Solo dos cosas se hacen a mano.** Todo lo demás es un comando:

1. **Pegar la clave de Gemini en `.env`** ([paso 5](#paso-5--la-clave-de-gemini)). Es un
   secreto y nunca viaja en el repositorio.
2. **Crear la regla de Outlook** ([sección 5.2](#52-la-regla-de-outlook)). Se crea en el
   servidor, desde el navegador, para que funcione con el equipo apagado.

Si algo falla al terminar, revisa esas dos primero: son el origen de casi todos los
problemas. Sin la clave, el programa no arranca. Sin la regla, las alertas nunca
llegan a la carpeta y el programa corre sobre una carpeta vacía sin quejarse, porque
para él no tener alertas es normal.

### Paso 1 — Instalar Python

```cmd
winget install Python.Python.3.12 --accept-package-agreements --accept-source-agreements
```

Instala sin pantallas y deja el PATH configurado. Se elige 3.12 y no la última versión
porque es con la que están probadas `pywin32` y `pymupdf`.

Si `winget` no está disponible, descarga Python desde
[python.org](https://www.python.org/downloads/windows/) y, en la **primera pantalla**
del instalador, marca **"Add python.exe to PATH"**. **No lo instales desde la
Microsoft Store**: esa versión corre aislada y da problemas con `pywin32`.

Cierra CMD, abre uno nuevo (los cambios de PATH solo aplican a ventanas nuevas) y
verifica:

```cmd
python --version
```

✅ Debe responder `Python 3.12.x`. Si dice "no se reconoce como un comando interno o
externo", el PATH no quedó configurado.

### Paso 2 — Descargar el proyecto

```cmd
cd /d C:\
git clone https://github.com/Nlira23/Automatizacion-Alarmas.git clariot
cd clariot
```

El repositorio es privado: Git abre el navegador para iniciar sesión en GitHub. Si en
cambio pide usuario y contraseña en la consola, la contraseña de GitHub no sirve. Hay
que crear un token en **Settings → Developer settings → Personal access tokens**, con
permiso `repo`.

Sin Git: descarga el ZIP del repositorio, descomprímelo y muévelo a `C:\clariot`.

> [!CAUTION]
> **Dónde NO dejar la carpeta:**
> - **Directamente en `C:\Users`**: es una carpeta protegida y el clon falla con
>   `fatal: could not create work tree dir`. Sí funciona dentro del perfil
>   (`C:\Users\<usuario>\clariot`).
> - **En el Escritorio, Documentos o cualquier ruta con `OneDrive`**: OneDrive sube
>   los archivos mientras cambian, y el programa escribe dos bases SQLite en `state\`
>   en cada corrida. Subir una base abierta puede bloquearla o corromperla, y perder
>   `ledger.db` es perder el registro de qué correos ya se procesaron. El acceso
>   directo del Escritorio sí puede quedar ahí: es solo un puntero al `.bat`.

### Paso 3 — Crear el entorno virtual

```cmd
python -m venv .venv
.venv\Scripts\activate.bat
```

✅ El prompt debe quedar con `(.venv)` adelante:

```
(.venv) C:\clariot>
```

**Si no aparece, no sigas**: lo que instales a continuación iría al Python del
sistema. El entorno se activa en cada ventana nueva de CMD. Los lanzadores `.bat` no
lo necesitan, porque usan directamente `.venv\Scripts\python.exe`.

### Paso 4 — Instalar las dependencias

```cmd
pip install -r requirements.txt
pip install -e .
```

**Son las dos líneas.** La segunda registra el paquete `clariot`; sin ella, cualquier
comando responde `No module named clariot`.

> [!NOTE]
> `pip install -e .` guarda la ruta absoluta de la carpeta dentro de `.venv`. Si
> después mueves la carpeta del proyecto, hay que recrear el entorno (ver
> [sección 12](#12-solución-de-problemas)).

### Paso 5 — La clave de Gemini

Crea una clave en [aistudio.google.com/apikey](https://aistudio.google.com/apikey).
Es gratuita y no pide tarjeta. Después:

```cmd
copy config\env.example .env
notepad .env
```

Busca la línea `GEMINI_API_KEY=` y pega la clave justo después del signo igual, sin
espacios ni comillas:

```
GEMINI_API_KEY=AIzaSyA...
```

Guarda con `Ctrl + S`. Las demás líneas quedan vacías: son para traducir el PDF
adjunto, que está desactivado. El archivo `.env` va en la **raíz del proyecto** (no en
`config\`), está en `.gitignore` y nunca se versiona.

### Paso 6 — Crear las carpetas en Outlook

Con Outlook abierto:

```cmd
python -m clariot --setup-folders
```

✅ Crea las cuatro carpetas definidas en `config/settings.yaml` e imprime `OK` por
cada una. Para usar otros nombres, ver la [sección 5.1](#51-las-carpetas).

### Paso 7 — Verificar que todo esté conectado

```cmd
python -m clariot --self-check
```

✅ Una salida sana se ve así:

```
Destinatarios precargados: 4 cliente(s) en clients.yaml
Outlook               : conectado
Carpeta 'Clariot': OK (0 correo(s) por revisar)
PDF adjunto           : sin traducir (translation.provider: none)
Frases (gemini)       : modelo gemini-flash-lite-latest, cuota gratuita de AI Studio

Todo listo.
```

El número de clientes depende de `config/clients.yaml`. Si algo falta, el comando lo
dice con nombre y apellido. **No sigas hasta ver `Todo listo`.**

### Paso 8 — La regla de Outlook

Ver la [sección 5.2](#52-la-regla-de-outlook). Sin ella, las alertas nunca llegan a la
carpeta que lee el programa.

### Paso 9 — La firma

Ver la [sección 5.3](#53-la-firma). Basta con confirmar que la cuenta tiene una firma
predeterminada para mensajes nuevos.

### Paso 10 — La primera prueba

Ver la [sección 6](#6-primera-prueba).

### Paso 11 — El acceso directo

Clic derecho sobre `Procesar alertas.bat` → **Enviar a** → **Escritorio (crear acceso
directo)**. Con eso, el uso diario es un doble clic, sin consola y sin comandos que
recordar.

## 5. Configurar Outlook

### 5.1 Las carpetas

Son cuatro, y **los nombres deben coincidir entre Outlook y `config/settings.yaml`**.
Los nombres en sí no importan; importa que sean iguales en los dos lados.

```
Bandeja de entrada
├── Clariot          ← acá llegan las alertas (la regla las mueve)
└── Procesados       ← acá se archivan los originales ya procesados

Borradores
├── Urgencias        ← lo que hay que enviar HOY
└── Por enviar       ← rutina
```

| Carpeta | Clave en `settings.yaml` | Vive bajo |
| --- | --- | --- |
| `Clariot` | `outlook.source_folders` | Bandeja de entrada |
| `Procesados` | `outlook.processed_folder` | Bandeja de entrada |
| `Por enviar` | `outlook.draft_folder` | Borradores |
| `Urgencias` | `outlook.urgent_draft_folder` | Borradores |

**Para cambiar los nombres**, edita primero `config/settings.yaml` y después ejecuta
`python -m clariot --setup-folders`. Si renombras las carpetas solo en Outlook, el
programa buscará carpetas que ya no existen, y `--self-check` lo va a reportar.

Dos detalles que no son arbitrarios:

- **Los borradores van bajo `Borradores`**, no bajo la Bandeja de entrada. Un mensaje
  sin enviar solo se abre en modo redacción, con el botón Enviar activo, si está bajo
  Borradores. Lo controla `outlook.draft_folder_parent: "drafts"`.
- **Se pueden usar subcarpetas** dentro de la de origen, por ejemplo
  `Clariot/ClienteA`. Con `include_subfolders: true` se leen todas, y el original se
  archiva en la subcarpeta equivalente: `Procesados/ClienteA`.

### 5.2 La regla de Outlook

La regla de Outlook mueve las alertas a la carpeta de origen. Se crea **una sola vez**
y conviene hacerla desde el navegador ([outlook.office.com](https://outlook.office.com)
→ engranaje → Correo → Reglas), para que corra en el servidor y funcione con Outlook
cerrado.

| Campo | Valor |
| --- | --- |
| Nombre | `Alertas Clariot` |
| Condición | **De** → `support@aliotportal.com` |
| Acción | **Mover a** → `Clariot` |

**Por remitente, no por asunto.** El asunto de las alertas es
`Event notification report - VX-0000000.`, donde la última parte es el número de serie
del equipo y cambia en cada alerta. Además, el PDF se llama
`Event Report - 1st Notification`, y ese "1st" indica que hay seguimientos con otra
redacción. Una regla por asunto fallaría **en silencio** justo con las alarmas
repetidas, que son las que más importan.

Al guardar, Outlook ofrece dos opciones:

- **"Detener el procesamiento de otras reglas"**: márcala si ninguna otra regla
  necesita ver estos correos.
- **"Ejecutar la regla ahora"**: márcala solo si quieres mover también las alertas que
  ya están en la bandeja.

### 5.3 La firma

Cada borrador lleva al final la **firma de Outlook de la cuenta**, con su logo e
imágenes. Ya no hay que insertarla a mano.

El programa la lee de `%APPDATA%\Microsoft\Signatures` y elige así:

1. La firma **predeterminada para mensajes nuevos** configurada en Outlook
   (Archivo → Opciones → Correo → Firmas).
2. Si no hay una predeterminada: la firma cuyo nombre contiene el correo definido en
   `email.sender_account`.
3. Si hay una sola firma: esa.

> [!WARNING]
> Si hay **varias firmas y ninguna es predeterminada**, el borrador sale **sin firma**
> y solo queda un aviso en `logs/clariot.log`. Para evitarlo, define una firma
> predeterminada para mensajes nuevos en Outlook.

Un problema con la firma nunca impide crear el borrador. Para desactivarla, pon
`email.use_outlook_signature: false` en `config/settings.yaml`.

## 6. Primera prueba

No hace falta esperar una alerta real. La prueba avanza en cuatro pasos, de menor a
mayor impacto:

**1. ¿Lee bien el PDF?** Sin tocar Outlook:

```cmd
python -m clariot --preview "ruta\al\reporte.pdf"
```

Muestra los datos extraídos y escribe `preview.html`, que es exactamente el cuerpo del
correo que se generaría. Ábrelo en el navegador. Si el PDF se leyó mal, se ve acá
antes de hacer nada.

**2. ¿Está todo conectado?**

```cmd
python -m clariot --self-check
```

**3. ¿Qué haría, sin hacerlo?**

```cmd
python -m clariot --dry-run
```

Recorre la carpeta y muestra qué haría. No crea borradores, no mueve correos, no
archiva PDF y no marca nada como procesado. Lo que sí puede escribir es la caché de
frases traducidas y una línea en `audit.csv` por cada alarma que quedaría retenida.

**4. Una alerta de prueba de punta a punta:**

```cmd
python -m clariot --simulate "ruta\al\reporte.pdf"
```

Esto deja un correo "PRUEBA CLARIOT", sin leer y con el PDF adjunto, en la carpeta de
origen. No envía ni recibe nada. Después, doble clic en **`Probar una alerta.bat`**,
que procesa un solo correo (equivale a `python -m clariot --limit 1`).

> [!NOTE]
> `--limit 1` procesa el correo **más antiguo** de la carpeta. Si ya hay alertas
> reales esperando, procesará una de esas y no la de prueba.

Revisa el borrador en Outlook:

- [ ] El PDF original está adjunto.
- [ ] El informe en español se lee bien.
- [ ] El destinatario y las copias son los esperados para ese cliente (ver
      [sección 9.3](#93-destinatarios-y-copias)).
- [ ] Aparecen `[NOMBRE]` y `[COMPLETAR]`.
- [ ] Al final está la firma.

Al terminar, borra el borrador de prueba y el mensaje de prueba de `Procesados`.

> [!CAUTION]
> Usa el PDF de una alerta **que ya fue atendida**. El evento queda registrado: si
> después llega esa misma alerta de verdad, el programa la tomará como un reenvío y no
> creará borrador.

## 7. Uso diario

### Los lanzadores

| Archivo | Qué hace | Cuándo usarlo |
| --- | --- | --- |
| `Procesar alertas.bat` | Procesa todas las alertas pendientes y muestra qué revisar | **Todos los días** (es el del acceso directo) |
| `Probar una alerta.bat` | Procesa **un solo** correo (`--limit 1`) | Para probar después de instalar o de un cambio |
| `run.bat` | Ejecuta `python -m clariot` sin mensajes ni pausa, y acepta parámetros | Para el Programador de tareas |

### La rutina

1. Abre Outlook.
2. Doble clic en el acceso directo **Procesar alertas**.
3. Lee el resumen que aparece al final, por ejemplo:

   ```
   3 en la carpeta | 3 eventos nuevos | 1 urgentes | 2 normales
   ```

   Presta atención a las partes que dicen **ATENCION**: alarmas retenidas, PDF sin
   traducir o correos con más de un PDF. Ver la [sección 12](#12-solución-de-problemas).
4. Revisa en este orden:
   1. `Borradores / Urgencias` → enviar **hoy**.
   2. `Borradores / Por enviar` → rutina.

### Qué revisar en cada borrador

- [ ] **Destinatario.** Viene cargado si el cliente está en `config/clients.yaml`; si
      no, agrégalo. Las copias internas de Emeltec ya vienen en CC.
- [ ] Reemplaza **`[NOMBRE]`** por el nombre del contacto.
- [ ] Escribe la recomendación de Emeltec donde dice **`[COMPLETAR]`**.
- [ ] Si el asunto empieza con **`[TRADUCIR]`** o hay un aviso rojo "NO ENVIAR ASÍ",
      corrige el texto que quedó en inglés.
- [ ] Revisa los PDF adjuntos.
- [ ] Envía.

> [!IMPORTANT]
> Para los clientes de `clients.yaml` el borrador **ya tiene destinatario**, así que
> Outlook permite enviarlo tal cual. Revísalo siempre antes de enviar.

Mientras un borrador conserve `[COMPLETAR]` y no le agregues destinatarios, el
programa lo sigue **actualizando** con las alarmas nuevas del mismo equipo. Apenas lo
editas, deja de tocarlo y las alarmas siguientes van a un borrador nuevo. Ver la
[sección 9.2](#92-un-borrador-por-equipo-que-va-creciendo).

### Automatizarlo (opcional)

Se puede ejecutar solo con el **Programador de tareas de Windows**:

| Opción | Valor |
| --- | --- |
| Programa | `C:\clariot\run.bat` |
| Desencadenador | **Al iniciar sesión**, repitiendo cada 15 minutos |
| Seguridad | **"Ejecutar solo cuando el usuario haya iniciado sesión"** |

La opción "Ejecutar tanto si el usuario inició sesión como si no" **no funciona**:
corre sin escritorio interactivo, Outlook no responde y la tarea falla siempre.

## 8. Comandos

Con el entorno activado (`.venv\Scripts\activate.bat`), o usando
`.venv\Scripts\python.exe` en lugar de `python`:

| Comando | Qué hace | ¿Modifica algo? |
| --- | --- | --- |
| `python -m clariot` | Operación normal: procesa las alertas de la carpeta | Sí: crea borradores, mueve correos, escribe en `state/` y `archive/` |
| `--self-check` | Verifica Outlook, carpetas y traductor | No |
| `--setup-folders` | Crea las carpetas de Outlook definidas en `settings.yaml` | Solo crea carpetas |
| `--preview PDF` | Muestra los datos de un PDF y escribe `preview.html` con el correo resultante | Solo `preview.html` y la caché de frases |
| `--dry-run` | Muestra qué haría, sin hacerlo | No crea borradores ni mueve correos ([detalle](#6-primera-prueba)) |
| `--simulate PDF` | Deja una alerta de prueba sin leer en la carpeta de origen | Un correo de prueba en Outlook |
| `--simulate-id ID` | Junto con `--simulate`: fija el Message-ID, para probar la protección contra duplicados | — |
| `--limit N` | Procesa como máximo N correos (los omitidos también cuentan) | Igual que la operación normal |
| `--pending` | Lista las alarmas retenidas y las que esperan informe semanal | Solo la caché de frases |
| `--report` | Crea el informe semanal consolidado por cliente | Sí: crea borradores |
| `--revisar-frases` | Lista todas las frases que tradujo la IA, para auditarlas (alias: `--review-phrases`) | No |
| `--verbose` | Log detallado (nivel DEBUG) | — |

**Códigos de salida**, útiles para el Programador de tareas:

| Código | Significado |
| --- | --- |
| `0` | Todo bien |
| `1` | Terminó, pero algún correo falló (o `--self-check` encontró problemas) |
| `2` | No pudo arrancar: configuración inválida, falta la clave de Gemini u Outlook no disponible |

## 9. Cómo decide el programa

### 9.1 La gravedad

| Situación | Asunto | Carpeta |
| --- | --- | --- |
| 2 o más alarmas del mismo equipo **el mismo día** | `[CRITICO]` | `Urgencias` |
| 2 o más alarmas del mismo equipo **en 7 días** | `[URGENTE]` | `Urgencias` |
| El reporte declara urgencia alta (`urgency.critical_values`) | `[URGENTE]` | `Urgencias` |
| Alarma aislada | `Reporte de monitoreo - …` | `Por enviar` |

- **La ventana de 7 días se cuenta hacia atrás desde la fecha de la alarma**, no desde
  el lunes. Un par domingo–lunes con ocho horas de diferencia cuenta como repetición;
  con semanas de calendario quedaría partido en dos.
- **El equipo se identifica por número de serie**, no por nombre: los nombres varían
  entre reportes, los números de serie no.
- **Solo sale un informe nuevo si la gravedad sube.** Dentro de los 7 días
  (`urgency_cooldown_days`) siguientes a un informe urgente o crítico, una alarma del
  mismo nivel o menor se informa como normal. Sin esa regla, la cuarta alarma del día
  generaría otro `[CRITICO]` y la palabra dejaría de significar algo.

### 9.2 Un borrador por equipo, que va creciendo

La segunda alarma de un equipo **no crea otro correo**: actualiza el borrador abierto,
suma su PDF y lo mueve de carpeta si la gravedad subió. La gravedad de un borrador es
**la peor que alcanzó el equipo**, nunca la de la última alarma.

El cuerpo detalla **cada condición distinta**, no solo la última. Si la bomba repite
la misma holgura tres veces, sale un bloque con esa condición y sus tres fechas. Si
reporta holgura y después un rodamiento, salen dos bloques.

Al actualizar un borrador, el programa reescribe el asunto, las copias, el cuerpo y
los adjuntos, pero **nunca toca el campo Para**.

**Cuándo deja de actualizarlo y crea uno nuevo:**

| Situación | Cómo lo detecta |
| --- | --- |
| El borrador ya se envió | Outlook lo marca como enviado |
| Alguien lo editó | Ya no tiene `[COMPLETAR]`, o tiene más destinatarios (Para o CC) que los configurados para ese cliente |
| El borrador se borró | Ya no existe en Outlook |
| La primera alarma tiene más de 7 días | `grouping.window_days` |

El trabajo de la persona nunca se pierde: un borrador editado no se vuelve a tocar.

Cuando el borrador anterior **ya se envió**, el nuevo detalla solo las alarmas
nuevas, adjunta solo sus PDF y explica de dónde viene la gravedad: *"Este equipo
acumula 4 alarmas desde el 22 de julio. Las 3 anteriores ya fueron informadas."*

### 9.3 Destinatarios y copias

Los destinatarios salen de `config/clients.yaml`:

```yaml
clients:
  - match: ["nombre del cliente"]     # se busca dentro del campo COMPANY del PDF
    display_name: "Nombre del cliente"
    to: ["contacto@cliente.cl"]
    cc: ["copia-propia-del-cliente@emeltec.cl"]

critical_cc: []        # se agrega solo a los borradores de nivel crítico
always_cc:             # se agrega a TODOS los borradores
  - "archivo@emeltec.cl"
```

- **Cómo encuentra al cliente:** busca cada texto de `match` dentro del campo
  **COMPANY** del PDF, sin distinguir mayúsculas ni tildes. Si coinciden varios, gana
  el más largo.
- **Clientes cargados hoy:** Prolesur, Vital Jugos, Soprole y Surlat. `always_cc`
  lleva 4 direcciones internas de Emeltec.
- **Copias:** se arman con el `cc` del cliente, más `always_cc`, más `critical_cc` si
  el nivel es crítico. No se repiten direcciones, y lo que ya está en Para no se
  copia.
- **Cliente que no está en la lista:** el borrador sale con el campo Para **vacío** y
  el resumen lo informa como "N esperan destinatario". Outlook no permite enviar un
  correo sin destinatario, así que no puede salir por accidente.

> [!NOTE]
> `tests/test_clients_config.py` verifica las listas exactas de `clients.yaml`. Si
> agregas o cambias un destinatario, actualiza también ese test.

### 9.4 La traducción, en tres capas

```
Frase en inglés
   ├─ 1. config/glossary.yaml   → la redacción aprobada. Siempre gana
   ├─ 2. caché en events.db     → ya se tradujo una vez. Gratis e idéntica
   └─ 3. Gemini la traduce      → SE GUARDA → desde ahí, es capa 2
```

**La capa 2 es la que hace esto confiable.** Un servicio de traducción consultado dos
veces puede responder distinto, y un cliente que compara el informe de enero con el
de marzo vería la misma condición descrita de dos formas. Guardar la primera
respuesta lo elimina, y un test lo verifica.

**A Gemini solo viaja la frase técnica**: sin cliente, planta, equipo ni número de
serie. También hay un test que lo verifica. El modelo es
`glossary.gemini_model: "gemini-flash-lite-latest"`, un alias que Google mantiene
actualizado.

**Si una frase no se puede traducir** (por ejemplo, Gemini no responde):

| Nivel de la alarma | Qué pasa |
| --- | --- |
| Normal | **Queda retenida, sin borrador** (`glossary.on_missing: "hold"`). Se reintenta sola en la próxima corrida. El cliente nunca recibe texto en inglés |
| Urgente o crítica | **Nunca se retiene**: el borrador sale con `[TRADUCIR]` en el asunto y un aviso rojo "NO ENVIAR ASÍ — FALTA TRADUCIR" |

Para ver lo retenido: `python -m clariot --pending`. Para destrabarlo de inmediato,
agrega la frase a `config/glossary.yaml` y vuelve a ejecutar el programa.

`python -m clariot --revisar-frases` lista todo lo que tradujo la IA, para auditarlo.
Lo que no guste se corrige en `glossary.yaml`, que siempre gana.

**El PDF adjunto queda en su idioma original.** Traducirlo está construido pero
desactivado: cuesta unos USD 0,08 por página y requiere facturación en Google Cloud.
Se activa con `translation.provider: "google"` (o `"deepl"`), instalando
`requirements-optional.txt` y completando las variables correspondientes en `.env`.

### 9.5 Duplicados y respaldo

Dos protecciones, en capas distintas:

- **`state/ledger.db`: ¿ya procesé este correo?** Usa el `Message-ID` de internet, que
  se mantiene aunque el correo cambie de carpeta (el `EntryID` de Outlook no). El
  indicador de leído de Outlook **no** se usa como fuente de verdad: bastaría con que
  alguien abriera el correo para que esa alerta quedara invisible.
- **`state/events.db`: ¿ya conté este evento?** La clave es el equipo más la fecha y
  hora del reporte. Los sistemas de alerta reenvían notificaciones, y contar correos
  dispararía una urgencia falsa por un solo problema real.

Cada alerta deja tres respaldos: el correo completo en `Procesados`, el PDF en
`archive/AAAA-MM/` y la decisión tomada, con su razón, en `archive/audit.csv`.

### 9.6 El informe semanal (`--report`)

`python -m clariot --report` agrupa las alarmas que todavía no se informaron y crea
**un borrador por cliente** en `Por enviar`.

Con la configuración actual (`reports.single_alarm: "draft"`), cada alarma genera su
propio borrador en el momento. Por eso `--report` en la práctica solo recoge las
alarmas **retenidas**, y ningún lanzador lo ejecuta. Si se cambia a
`single_alarm: "weekly"`, las alarmas normales se acumulan y salen todas juntas con
`--report` (por ejemplo, los viernes); las urgentes y críticas siguen saliendo al
momento.

Antes de enviar un informe semanal, revisa que no queden frases en inglés.

## 10. Configuración

| Archivo | Para qué sirve |
| --- | --- |
| `config/settings.yaml` | Carpetas, umbrales de gravedad, asuntos, traducción, firma |
| `config/clients.yaml` | Destinatarios y copias por cliente ([sección 9.3](#93-destinatarios-y-copias)) |
| `config/glossary.yaml` | Frases aprobadas y vocabulario técnico |
| `config/pdf_labels.yaml` | Etiquetas que se buscan en el PDF (inglés y español) |
| `templates/*.html.j2` | Cuerpo de los correos (urgencia y semanal) |
| `.env` (raíz del proyecto) | La clave de Gemini. Nunca se versiona |

Todos los YAML están comentados línea por línea. Los ajustes que más se tocan:

| Clave | Valor actual | Qué controla |
| --- | --- | --- |
| `grouping.window_days` | `7` | Días en que dos alarmas del mismo equipo cuentan como el mismo problema |
| `grouping.urgent_threshold` | `2` | Alarmas necesarias para escalar a urgente o crítico |
| `grouping.same_day_is_critical` | `true` | Si 2 alarmas el mismo día son nivel crítico |
| `grouping.urgency_cooldown_days` | `7` | Días sin repetir un informe urgente para el mismo equipo |
| `grouping.retention_days` | `180` | Días que se guardan los eventos ya informados |
| `glossary.gemini_model` | `gemini-flash-lite-latest` | Modelo que traduce las frases |
| `glossary.on_missing` | `hold` | `hold` retiene la alarma; `mark` crea el borrador con `[TRADUCIR]` |
| `reports.single_alarm` | `draft` | `draft`: un borrador por alarma; `weekly`: se acumulan para `--report` |
| `reports.brand_color` | `#005f7f` | Color corporativo de los encabezados del correo |
| `reports.include_emeltec_note` | `true` | Incluye el recuadro "Recomendación de Emeltec" con `[COMPLETAR]` |
| `email.use_outlook_signature` | `true` | Agrega la firma de Outlook |
| `email.sender_account` | `""` | Cuenta desde la que se envía; vacío usa la predeterminada |
| `translation.provider` | `none` | Traducción del PDF adjunto: `none`, `google` o `deepl` |

> [!WARNING]
> Con `reports.include_emeltec_note: false` desaparece `[COMPLETAR]`, y el programa
> trata **todos** los borradores como editados: dejan de actualizarse con las alarmas
> nuevas y cada alarma crea un borrador aparte.

**El color corporativo es una sola línea** (`reports.brand_color`): las dos
plantillas derivan de ahí. Todos los bloques del cuerpo declaran el mismo ancho como
atributo HTML, porque el motor de Word que usa Outlook ignora `max-width`.

**Claves sin efecto:** `email.subject_template`, `email.critical_subject_prefix` y
`email.include_report` se leen pero no se usan. Los asuntos se definen en
`reports.*_subject`.

## 11. Archivos que genera

Todo se crea en la primera corrida y está en `.gitignore`: no hay que crear nada a mano
y nada de esto se sube al repositorio.

| Ruta | Contenido | ¿Se puede borrar? |
| --- | --- | --- |
| `state/ledger.db` | Qué correos ya se procesaron | **No.** Es el registro que evita procesar dos veces |
| `state/events.db` | Alarmas, caché de traducciones y borradores abiertos | **No.** Perderlo borra el historial de gravedad y las traducciones guardadas |
| `archive/AAAA-MM/` | Los PDF originales, con su nombre intacto | Conviene conservarlos |
| `archive/audit.csv` | Una fila por decisión: fecha, equipo, estado y motivo (separado por `;`) | Conviene conservarlo |
| `logs/clariot.log` | Log de cada corrida | Sí |
| `.work/` | Archivos temporales de cada correo | Sí, se limpia solo |
| `preview.html` | Resultado de `--preview` | Sí |

Respalda `state/` y `archive/` si el equipo se va a formatear o reemplazar.

Estados que aparecen en `audit.csv`:

| Estado | Significado |
| --- | --- |
| `informe_single`, `informe_urgent`, `informe_critical` | Se creó un borrador nuevo |
| `acoplado_*` | La alarma se sumó a un borrador abierto |
| `retenido` | Sin borrador: falta traducir una frase |
| `acumulado` | Espera el informe semanal |
| `reenvio` | Evento ya registrado; no se contó de nuevo |
| `ignorado` | El correo no traía PDF |
| `revisar` | El correo traía más de un PDF; requiere revisión manual |
| `informe_semanal` | Incluido en un informe semanal |
| `error` | Falló; el detalle está en la columna `detail` y en el log |

## 12. Solución de problemas

| Síntoma | Causa probable | Qué hacer |
| --- | --- | --- |
| `No module named clariot` | Faltó `pip install -e .` o el entorno no está activado | Activa `.venv` y repite el [paso 4](#paso-4--instalar-las-dependencias) |
| No aparece `(.venv)` al activar | Se usó PowerShell | Usa CMD, o `.\.venv\Scripts\Activate.ps1` |
| `--self-check` no conecta con Outlook | Outlook cerrado, o es el "nuevo Outlook" | Abre el Outlook clásico y desactiva "Nuevo Outlook" |
| `Outlook folder 'Clariot' not found` | Los nombres no coinciden con `settings.yaml` | Ejecuta `--setup-folders` o corrige el YAML |
| El programa termina con código `2` al arrancar | Falta `GEMINI_API_KEY` en `.env`, o el YAML tiene un error | Revisa `.env` (en la raíz) y el mensaje en pantalla |
| Llegó una alerta y no hay borrador | Quedó **retenida** por una frase sin traducir | `python -m clariot --pending`; agrega la frase a `glossary.yaml` o espera la próxima corrida |
| El log muestra HTTP `503` o `429` de Gemini | Gemini saturado o cuota agotada | Nada: se reintenta solo en la próxima corrida |
| "El modelo '…' fue retirado por Google" | Google dio de baja ese modelo | Cambia `glossary.gemini_model` en `settings.yaml` (por ejemplo, a `gemini-flash-latest`) |
| "La clave de Gemini es invalida" | La clave en `.env` está mal copiada o fue revocada | Vuelve a pegarla ([paso 5](#paso-5--la-clave-de-gemini)) |
| La segunda alarma creó otro borrador en vez de actualizar | El anterior se editó, se envió o tiene más de 7 días | Es lo esperado ([sección 9.2](#92-un-borrador-por-equipo-que-va-creciendo)) |
| El borrador no tiene firma | Hay varias firmas y ninguna es predeterminada | Define una firma predeterminada ([sección 5.3](#53-la-firma)) |
| Un correo queda en la carpeta y se avisa en cada corrida | Trae más de un PDF | Revísalo a mano ([límites](#15-límites-conocidos)) |
| La carpeta de origen siempre está vacía | Falta la regla de Outlook, o filtra por asunto | Revisa la regla ([sección 5.2](#52-la-regla-de-outlook)) |
| La tarea programada falla siempre | Está configurada para correr sin sesión iniciada | "Ejecutar solo cuando el usuario haya iniciado sesión" |
| Moví la carpeta del proyecto y nada funciona | `.venv` guarda la ruta antigua | `rmdir /s /q .venv` y repite los pasos 3 y 4 |

Si nada de esto aplica: `python -m clariot --verbose` y revisa `logs/clariot.log`.

## 13. Actualizar a una versión nueva

```cmd
cd /d C:\clariot
git pull
.venv\Scripts\activate.bat
pip install -r requirements.txt
pip install -e .
python -m clariot --self-check
```

`.env`, `state/`, `archive/` y `logs/` no se tocan: están fuera del repositorio. Los
archivos de `config/` **sí** están versionados. Si los editaste en este equipo,
`git pull` puede avisar de un conflicto: guarda tus cambios con `git stash` antes del
`pull` y recupéralos con `git stash pop` después.

## 14. Desarrollo

### Arquitectura

La regla es simple: **todo lo que decide algo es puro y testeable**, y todo lo que se
comunica con el exterior vive en `adapters/`. Por eso los tests corren sin Outlook,
sin red y sin claves.

```
src/clariot/
├── __main__.py            línea de comandos y arranque
├── config.py              carga y valida settings.yaml, clients.yaml y .env
│
│  --- núcleo puro, testeable sin Outlook ni red ---
├── models.py              AlertReport, ClientRoute
├── textutils.py           normalización sin tildes, nombres de archivo seguros
├── pdf_parser.py          lee el PDF por coordenadas (su docstring explica por qué)
├── glossary.py            las frases aprobadas
├── resolver.py            las tres capas de traducción
├── store.py               eventos, caché de frases, borradores abiertos
├── ledger.py              qué correos ya se procesaron
├── classifier.py          normal / urgente / crítico
├── aggregate.py           agrupa por equipo y por cliente, fechas en español
├── report_builder.py      arma los borradores, destinatarios y adjuntos
├── audit.py               el CSV de trazabilidad
├── ingest.py              el flujo principal de captura
├── weekly.py              el informe semanal consolidado
│
│  --- adaptadores al mundo exterior ---
└── adapters/
    ├── outlook.py             COM: carpetas, borradores, firma. Único archivo que conoce Outlook
    ├── gemini_text.py         frases, por HTTP sin SDK
    ├── google_text.py         frases, alternativa con Google Cloud
    ├── google_translator.py   PDF adjunto, con Google Cloud
    └── deepl_translator.py    PDF adjunto, con DeepL

templates/
├── urgency_report.html.j2     correo por equipo (normal, urgente, crítico)
└── weekly_report.html.j2      informe semanal por cliente
```

### Tests

```cmd
pip install -r requirements-dev.txt
pytest
```

Son **235 tests**. Cubren el parseo por coordenadas, el glosario, las tres capas de
traducción, la gravedad, la agregación, el armado de informes, la idempotencia, los
destinatarios y la firma. Outlook se reemplaza por objetos falsos, así que no hace
falta tenerlo abierto.

Unos 20 tests se saltan solos en una copia recién clonada, y `pytest` queda en verde
igual:

- **Reportes reales de Alfa Laval** (`test_real_report.py`,
  `test_real_english_report.py`). Esos PDF no están en el repositorio porque llevan
  empresa, planta, equipo y número de serie de un cliente real; la carpeta `Imagenes/`
  está en `.gitignore` por eso.
- **Traductores opcionales**, cuando `deepl` o `google-cloud-translate` no están
  instalados.

Contra el Outlook real, la verificación es `--self-check` y `--simulate`.

### Convenciones

- Commits con [Conventional Commits](https://www.conventionalcommits.org/):
  `feat(outlook): ...`, `fix(ingest): ...`, `docs: ...`.
- El código y sus comentarios van en inglés. Los textos que ve el usuario y la
  documentación van en español.

## 15. Límites conocidos

- **Nunca envía correos.** Por diseño, y no es configurable.
- **Requiere Outlook clásico abierto** y una sesión de Windows iniciada. No corre como
  servicio.
- **Un correo con dos PDF no se procesa.** Un reporte real trae uno solo. Con dos no
  hay forma segura de elegir, y tomar el primero descartaría una alarma en silencio.
  El correo queda en la carpeta de origen y la consola lo avisa en cada corrida hasta
  que alguien lo revise.
- **La hora del reporte se copia tal cual, en UTC.** La etiqueta del PDF es
  `Time Of Event (UTC)` y ese valor pasa al correo sin convertir. En Chile (UTC−4), un
  correo que dice "24 de julio, 21:05" corresponde a las 17:05 locales, y la regla de
  "dos alarmas el mismo día" compara el día UTC. Es una decisión explícita: el correo
  dice exactamente lo que dice el reporte del fabricante.
- **Las fechas se leen como día-mes-año**, verificado con un reporte real. Si alguna
  vez llega uno con formato mes-día, hay que ajustar `_DATE_FORMATS` en
  `src/clariot/store.py`.
- **Una bomba que alarma cada 10 días no se agrupa**: la ventana es de 7 días,
  configurable en `grouping.window_days`.
- **El PDF adjunto queda en inglés**, y los gráficos de vibración no se traducen
  nunca: son imágenes.
- **Máximo 20 adjuntos por borrador**, para que un equipo con cincuenta alarmas no
  genere un correo que el servidor rechace.
- **Firma:** con varias firmas y ninguna predeterminada, el borrador sale sin firma.

## 16. Otros documentos

| Documento | Para qué sirve |
| --- | --- |
| [DESPLIEGUE.md](DESPLIEGUE.md) | Instalación desde cero en el equipo del técnico, con checklist para seguir en pantalla |
| [INSTALACION.md](INSTALACION.md) | Manual extendido de instalación y operación |
| [ESTADO.md](ESTADO.md) | Registro de decisiones: qué se decidió, por qué y qué se aprendió. Se lee antes de **cambiar** el sistema |
| [PROPUESTA-INFORMES.md](PROPUESTA-INFORMES.md) | Diseño original de los informes (documento histórico) |
