# Instalación en el equipo del técnico, desde cero

Para seguir en pantalla el día de la instalación, en un equipo **donde no hay nada
instalado**. Cada paso se verifica antes de pasar al siguiente.

**Todos los comandos van en CMD**, el símbolo del sistema de Windows. No en
PowerShell. Se eligió CMD a propósito: PowerShell bloquea por defecto la ejecución
de scripts y eso obliga a un permiso extra que en CMD no hace falta.

**Equipo destino:** el del técnico que atiende las alertas, con su Outlook y la
cuenta que recibe los correos de `support@aliotportal.com`.

---

## Cómo abrir CMD

Apretá **`Windows + R`**, escribí `cmd` y Enter. Se abre una ventana negra: ahí van
todos los comandos de este documento.

Si en algún paso Windows dice que hacen falta permisos de administrador: cerrá esa
ventana, andá al menú Inicio, escribí `cmd`, clic derecho sobre **Símbolo del
sistema** → **Ejecutar como administrador**.

---

## Antes de ir

- [ ] Confirmar que en ese equipo hay **Outlook clásico de escritorio**. Si solo
      tiene el nuevo, hay que instalar el clásico primero, y eso puede necesitar a
      TI. Sin Outlook clásico **nada de esto funciona**.
- [ ] Tener a mano el instalador de **Python 3.12** descargado, por si el equipo no
      tiene internet cómodo: [python.org/downloads/windows](https://www.python.org/downloads/windows/)
- [ ] Tener a mano **la clave de Gemini** (la de AI Studio) para copiarla en el paso 5.
- [ ] Tener a mano **un PDF de reporte real** en un pendrive, para la prueba del
      paso 8 si todavía no llegó ninguna alerta al buzón.

---

## 1. Outlook clásico

### Primero: ¿ya está instalado?

Outlook clásico **no se descarga como un archivo suelto ni se elige dónde
instalarlo**: viene con Microsoft 365 y se instala solo en `C:\Program Files`. Lo
más probable es que ya esté en el equipo.

- [ ] Comprobarlo con un comando, sin abrir nada:

```cmd
reg query HKEY_CLASSES_ROOT\Outlook.Application\CLSID
```

- **Si imprime una clave** (algo como `HKEY_CLASSES_ROOT\Outlook.Application\CLSID`),
  Outlook clásico está instalado. No hay que descargar nada.
- **Si dice que no puede encontrar la clave del registro**, en el equipo solo está
  el Outlook nuevo y hay que instalar el clásico.

Ese es exactamente el chequeo que hace el programa por dentro: el Outlook nuevo
(`olk.exe`, que es una app de la Tienda) no registra servidor COM, y sin COM no hay
forma de hablarle.

### Si hace falta instalarlo

- [ ] Entrar a [portal.office.com](https://portal.office.com) con la cuenta de la
      empresa → **Instalar aplicaciones** → las aplicaciones de escritorio de
      Microsoft 365.
- [ ] Eso instala Word, Excel y **Outlook clásico** juntos, en la ruta que Microsoft
      decide. **No se elige la carpeta y no hace falta.**
- [ ] Si el portal no ofrece la instalación, la licencia no la habilita: eso lo
      resuelve TI, no se puede saltar.

> La regla de no usar OneDrive **es solo para la carpeta del proyecto**, porque ahí
> se escriben las bases de datos. Los programas que se instalan —Outlook, Python—
> van donde su instalador los ponga. No hay nada que decidir ahí.

### Después, en los dos casos

- [ ] Abrir Outlook y confirmar que la cuenta configurada es la que recibe las
      alertas.
- [ ] Si arriba a la derecha hay un interruptor **"Probar el nuevo Outlook"**,
      dejarlo **apagado**.
- [ ] Dejar Outlook **abierto**. El programa le habla a la sesión que está abierta;
      con Outlook cerrado, falla.

## 2. Instalar Python

**Lo más rápido, desde la consola.** Windows 11 ya trae `winget`:

```cmd
winget install Python.Python.3.12 --accept-package-agreements --accept-source-agreements
```

Instala sin pantallas y deja el PATH puesto. Después hay que **cerrar la ventana de
CMD y abrir una nueva**, y verificar con `python --version`.

Se elige 3.12 y no la última: es la versión con la que están probadas `pywin32` y
`pymupdf`, las dos librerías que hacen el trabajo pesado. El equipo de producción no
es el lugar para estrenar una versión.

**Node.js no hace falta.** El proyecto es Python puro: ni un archivo JavaScript, ni
`package.json`. Y no podría haberlo, porque Outlook renderiza con el motor de Word,
que ignora JavaScript.

### Si winget no está o falla

Puede pasar con un App Installer viejo o por política de la empresa.

- [ ] Ejecutar el instalador de [python.org](https://www.python.org/downloads/windows/).
- [ ] En la **primera pantalla**, antes de apretar Install, marcar la casilla de
      abajo que dice **"Add python.exe to PATH"**. Es la más importante de toda la
      instalación: sin ella, ningún comando de acá funciona.
- [ ] Si pide permisos de administrador y no los hay: **Customize installation** y
      desmarcar *"Install for all users"*. Instala solo para ese usuario y sirve
      igual.
- [ ] **No instalarlo desde la Microsoft Store.** Esa versión corre aislada y da
      problemas justo con `pywin32`, que es la librería que le habla a Outlook.
- [ ] Terminar la instalación.
- [ ] **Cerrar la ventana de CMD y abrir una nueva** (los cambios de PATH solo
      aplican a ventanas nuevas).
- [ ] Verificar:

```cmd
python --version
```

Tiene que responder `Python 3.12.x` o similar. Si dice **"no se reconoce como un
comando interno o externo"**, es que la casilla del PATH no quedó marcada:
desinstalar Python y repetir el paso.

## 3. Copiar el proyecto

El proyecto va a estar en un repositorio privado de Git. Hay dos formas de bajarlo,
y **la primera no requiere instalar nada**:

### Dónde va la carpeta: `C:\clariot`

Directo en la raíz del disco C. **No en el Escritorio, ni en Documentos, ni en
Descargas.**

El motivo es concreto: en los equipos de la empresa, OneDrive tiene activada la
copia de seguridad de carpetas, así que **el Escritorio y Documentos están adentro
de OneDrive**. Se comprueba mirando la ruta: si dice `OneDrive`, está sincronizada.

```
C:\Users\<usuario>\OneDrive\Escritorio     <- sincronizada, NO usar
C:\Users\<usuario>\OneDrive\Documentos     <- sincronizada, NO usar
C:\clariot                                 <- no se sincroniza, ACÁ VA
```

Y el problema con eso es que el programa guarda dos bases de datos SQLite en
`state\` —`ledger.db` y `events.db`— y las **escribe mientras corre**. OneDrive
sube los archivos justo cuando cambian, y subir una base SQLite abierta puede
bloquearla o dejarla corrupta. Perder `ledger.db` significa perder el registro de
qué correos ya se procesaron, y ahí sí aparecen borradores duplicados.

De paso, cada PDF que se archiva en `archive\` se subiría a la nube y se comería la
cuota sin necesidad: el correo original ya queda respaldado en `Procesados`.

> El **acceso directo del Escritorio sí puede quedar en OneDrive.** Es un archivo
> de dos líneas que solo apunta al `.bat`; lo que no puede estar sincronizado es la
> carpeta del proyecto.

Crear `C:\clariot` **no necesita permisos de administrador**. Si igual Windows los
pide, usá `C:\Users\<usuario>\clariot`, que tampoco se sincroniza.

### Cómo bajarlo

**Opción A — el ZIP (recomendada, no requiere instalar Git).**

- [ ] En la página del repositorio: botón verde **Code** → **Download ZIP**.
- [ ] Se descarga en `Descargas`. Clic derecho → **Extraer todo**.
- [ ] Abrir la carpeta extraída. **Ojo con esto:** el ZIP de GitHub trae una carpeta
      adentro de otra, con un nombre tipo `clariot-main`. Lo que hay que mover es
      **la de adentro**, la que contiene `config`, `src` y `requirements.txt`.
- [ ] Mover esa carpeta a `C:\` y renombrarla `clariot`.
- [ ] Verificar que la ruta final sea `C:\clariot\config`, y **no**
      `C:\clariot\clariot-main\config`. Si quedó anidada, nada va a funcionar.

> **Ojo con `C:\Users`.** Clonar ahí falla con
> `fatal: could not create work tree dir`, porque es una carpeta protegida de
> Windows y no se puede escribir sin administrador. Pasó en una instalación
> real. `C:\Users` guarda perfiles; la carpeta va en `C:\clariot`, o dentro
> del perfil del usuario (`C:\Users\<usuario>\clariot`), que sí es escribible.

**Opción B — con Git**, si el equipo ya lo tiene:

```cmd
cd /d C:\
git clone <URL-DEL-REPO> clariot
```

El nombre al final renombra la carpeta al clonar; sin él queda con el nombre del
repositorio. Siendo un repositorio privado, Git va a pedir autenticación: abre una
ventana del navegador para entrar con la cuenta de GitHub. Si en cambio pide usuario
y contraseña en la consola, la contraseña de GitHub ya no sirve desde 2021 y hay que
generar un token en Settings → Developer settings → Personal access tokens, con
permiso `repo`.

## 4. Instalar el programa

- [ ] Uno por uno, en la ventana de CMD:

```cmd
cd /d C:\clariot
python -m venv .venv
.venv\Scripts\activate.bat
pip install -r requirements.txt
pip install -e .
```

Después de la tercera línea, el texto de la izquierda tiene que empezar con
**`(.venv)`**, así:

```
(.venv) C:\clariot>
```

**Si no aparece `(.venv)`, no sigas**: los comandos siguientes se instalarían en el
Python del sistema y después nada va a funcionar. Repetí la línea
`.venv\Scripts\activate.bat`.

Las dos líneas de `pip` son necesarias. La segunda registra el paquete; sin ella,
todo comando responde `No module named clariot`.

- [ ] Verificar:

```cmd
python -m clariot --help
```

Tiene que imprimir la lista de opciones.

## 5. La clave de Gemini

- [ ] Crear el archivo de configuración:

```cmd
copy config\env.example .env
notepad .env
```

- [ ] En el Notepad que se abre, buscar la línea `GEMINI_API_KEY=` y pegar la clave
      pegada al signo igual, sin espacios y sin comillas:

```
GEMINI_API_KEY=AIzaSyA...
```

- [ ] Guardar (`Ctrl + S`) y cerrar el Notepad.

Las demás líneas del archivo se dejan vacías: son para traducir el PDF adjunto, que
hoy está desactivado.

## 6. Las carpetas de Outlook

- [ ] Con Outlook abierto:

```cmd
python -m clariot --setup-folders
```

- [ ] Confirmar en Outlook que quedaron estas cuatro:

```
Bandeja de entrada / Clariot
Bandeja de entrada / Procesados
Borradores / Urgencias
Borradores / Por enviar
```

Los nombres tienen que coincidir con `config\settings.yaml`. Si alguien las
renombra a mano en Outlook, hay que renombrarlas también en ese archivo.

## 7. Verificar que todo esté conectado

```cmd
python -m clariot --self-check
```

- [ ] Tiene que terminar diciendo **`Todo listo`**:

```
Destinatarios precargados: 4 cliente(s) en clients.yaml
Outlook               : conectado
Carpeta 'Clariot': OK (0 correo(s) por revisar)
PDF adjunto           : sin traducir (translation.provider: none)
Frases (gemini)       : modelo gemini-flash-lite-latest, cuota gratuita de AI Studio

Todo listo.
```

Si no lo dice, el mensaje nombra exactamente qué falta. **No sigas hasta ver
`Todo listo`.**

## 8. La regla del servidor

- [ ] Abrir un navegador y entrar a `outlook.office.com` **con la cuenta que recibe
      las alertas**.
- [ ] Engranaje de Configuración → **Correo** → **Reglas** → **Agregar nueva regla**.
- [ ] Nombre: `Alertas Clariot`
- [ ] Condición: **El remitente es** → `support@aliotportal.com`
- [ ] Acción: **Mover a** → `Clariot`
- [ ] Guardar.

> Desde el navegador, **no** desde el Outlook de escritorio: así la regla corre en
> el servidor y funciona incluso con el equipo apagado.

Por remitente y no por asunto: el asunto lleva el número de serie del equipo al
final y cambia en cada alerta, así que una regla por asunto fallaría en silencio.

Si al guardar Outlook ofrece **"Ejecutar la regla ahora"**, marcala solo si querés
que las alertas que ya están en la bandeja se muevan también.

## 9. La primera prueba

**Si ya hay alertas acumuladas en `Clariot`**, procesar una sola:

```cmd
python -m clariot --limit 1
```

**Si todavía no llegó ninguna**, usar un reporte real del pendrive. Copialo a
`C:\clariot` y después:

```cmd
python -m clariot --simulate "reporte.pdf" --simulate-id "<prueba-1@aliot>"
python -m clariot
```

La primera línea deja una alerta de prueba en `Clariot`; la segunda la procesa.

- [ ] Abrir el borrador en `Borradores\Por enviar` y revisar:
  - [ ] El PDF original está adjunto
  - [ ] El informe en español se lee bien
  - [ ] El destinatario es el esperado: cargado si el cliente está en
        `config/clients.yaml` (Prolesur, Vital Jugos, Soprole, Surlat), **vacío** si no
  - [ ] Las copias internas de Emeltec están en CC
  - [ ] Aparecen `[NOMBRE]` y `[COMPLETAR]`
  - [ ] Al final está la firma de Outlook (ver la sección 5.3 del README si falta)
- [ ] Anotar las frases que salgan marcadas `[EN]`: hay que agregarlas al glosario.
- [ ] **Borrar el mensaje de prueba** de `Procesados` y el borrador de prueba, si se
      usó `--simulate`.

## 10. El acceso directo

- [ ] Abrir `C:\clariot` en el Explorador de archivos.
- [ ] Clic derecho en **`Procesar alertas.bat`** → **Enviar a** → **Escritorio
      (crear acceso directo)**.
- [ ] Renombrarlo a **"Procesar alertas Clariot"**.
- [ ] Probarlo con doble clic, con Outlook abierto.

Desde acá el técnico no vuelve a escribir un comando nunca más.

## 11. Mostrarle el flujo al técnico

Explicarle estos seis puntos, en este orden:

- [ ] Abre Outlook y lo deja abierto.
- [ ] Doble clic en el acceso directo. Espera. Aprieta una tecla para cerrar.
- [ ] Revisa **`Urgencias`** primero: eso se manda hoy.
- [ ] Después **`Por enviar`**, que es rutina.
- [ ] En cada borrador: revisa el destinatario (agrégalo si viene vacío), reemplaza `[NOMBRE]`, completa la
      recomendación donde dice `[COMPLETAR]`, lee y envía.
- [ ] Si ve texto marcado `[EN]`, lo traduce y **avisa para agregarlo al glosario**,
      así no vuelve a aparecer nunca más.

Y lo más importante, decirlo explícitamente: **el programa nunca envía nada solo.**
Solo deja borradores. Él siempre lee y aprueba antes de que algo salga.

---

## Opcional: que corra solo cada 15 minutos

Con el Programador de tareas de Windows:

- [ ] Menú Inicio → **Programador de tareas** → **Crear tarea básica**.
- [ ] Nombre: `Clariot`
- [ ] Desencadenador: **Diariamente**, y después en Propiedades marcar
      **"Repetir cada 15 minutos"** durante 1 día.
- [ ] Acción: **Iniciar un programa** → `C:\clariot\Procesar alertas.bat`
- [ ] En la pestaña **General**, dejar marcado **"Ejecutar solo cuando el usuario
      haya iniciado sesión"**.

> La otra opción, **"Ejecutar aunque el usuario no haya iniciado sesión", NO
> funciona.** Corre sin escritorio interactivo y Outlook no responde: la tarea
> falla siempre. No es algo que se pueda parchear.

---

## Después de la primera semana

- [ ] Abrir `C:\clariot\archive\audit.csv` en Excel: cuántas alertas llegaron, de
      qué equipos, qué gravedad tuvieron. **Ese es el dato de volumen real** que hoy
      no existe.
- [ ] Correr `python -m clariot --revisar-frases` y revisar lo que tradujo la IA.
- [ ] Juntar las frases marcadas `[EN]` y agregarlas a `config\glossary.yaml`.
- [ ] Revisar `logs\clariot.log` por si hay errores repetidos.
- [ ] Con el volumen real a la vista, decidir si conviene:
  - cambiar `reports.single_alarm` a `"weekly"` si son demasiados correos,
  - ajustar `grouping.window_days` si aparecen equipos que alarman cada 10 días.

---

## Si algo se rompe

| Síntoma | Causa más probable |
| --- | --- |
| `python` no se reconoce como comando | No se marcó "Add python.exe to PATH". Reinstalar Python |
| No aparece `(.venv)` en el prompt | Falta correr `.venv\Scripts\activate.bat` |
| `No module named clariot` | Falta `pip install -e .` |
| `Classic Outlook ... is not installed` | Es el Outlook nuevo. Apagar el interruptor |
| `Folder 'Clariot' not found` | Falta correr `--setup-folders` |
| Las alertas no llegan a `Clariot` | La regla no está, o se creó en el Outlook de escritorio en vez del navegador |
| Los borradores salen con mucho `[EN]` | El glosario todavía tiene pocas frases; es esperable al principio |
| `ATENCION: ... con mas de un PDF` | Un correo trajo dos reportes. No se procesa a propósito: hay que mirarlo a mano |

El detalle de cada caso está en `INSTALACION.md`, sección 12.
