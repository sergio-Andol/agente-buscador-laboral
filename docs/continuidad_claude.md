# Continuidad del proyecto — Agente Buscador Laboral

> Handoff para una nueva sesión de Claude Code. Última actualización: 2026-07-13.
> Regla de mantenimiento: actualizar esta sección al tope con cada cambio funcional
> importante (qué se cambió / por qué / qué se probó / qué quedó pendiente /
> advertencias de seguridad). No hace falta por cambios de estilo o comentarios.

## 1. Estado actual del proyecto

Agente en Python que busca ofertas en Bumeran y Computrabajo, filtra, evita repetidos,
clasifica (POSTULAR/REVISAR/DESCARTAR), genera fichas de postulación y Excel con hoja
RESUMEN + hoja ACCIONES. Última capa agregada: **postulación asistida/confirmada con
click real**, gateada por doble confirmación humana. Funcional y probado (con red real,
sin disparar un envío real todavía — ver sección 7 y 9).

## 2. Ruta exacta del proyecto

**Carpeta definitiva (repo git, portfolio, donde se trabaja de acá en adelante):**
```
C:\Users\sergi\Documentos\Buscador de trabajo\Agente-Buscador-Laboral\
```

Sincronizada el 12/07: se copió el script y este mismo handoff desde la carpeta raíz
(`C:\Users\sergi\Documentos\Buscador de trabajo\`, sin subcarpeta) hacia acá, se limpió
`_PERFIL_SERGIO_CONTENIDO_BASE` de datos personales reales (quedó un perfil de
ejemplo genérico), y se pusheó a `origin/main` — commit `dc18b91`. `git status` limpio
al momento del push.

⚠️ La carpeta raíz (`...\Buscador de trabajo\`, un nivel arriba) sigue existiendo con
su propia copia del script y sus propios `vistos.json`/`historial_trabajos.xlsx`
reales — es la copia "de trabajo diario" de antes del sync, ahora **superada** por
esta. No editarla pensando que se propaga sola para acá: si hace falta traer algo de
ahí, hay que copiarlo a mano de nuevo, explícitamente.

⚠️ `perfil_sergio.txt` real (con nombre completo de Sergio) ya estaba commiteado en
`origin/main` de ANTES de esta sesión (commit `52c8c3d`) y sigue en el historial de
git aunque ahora esté en `.gitignore` — el `.gitignore` solo evita que se vuelva a
subir si cambia, no lo saca de commits viejos. No se tocó ese historial; es decisión
de Sergio si quiere reescribirlo.

## 3. Archivo principal

```
buscador_trabajos_v2.py
```
Todo el agente vive en este único archivo (~2900 líneas a esta fecha).

## 4. Configuración actual relevante

Panel de interruptores al principio del archivo. Valores **actuales en el archivo**
(no necesariamente los defaults recomendados — Sergio está en medio de pruebas):

```python
MODO_EJECUCION = "MANUAL"            # preparado para la 1ra prueba real supervisada
ACTIVAR_INDEED = False
GUARDAR_DEBUG_INDEED = True
DESCARTAR_BUMERAN_EXTERNOS = True
CT_DIAS = 2
SONIDO_FINAL = "SUAVE"

# auto-configuradas por MODO_EJECUCION=MANUAL (ver bloque if/elif):
NAVEGADOR_VISIBLE = False
MOSTRAR_CARTEL_FINAL = True
ABRIR_EXCEL_AL_FINAL = True
MODO_TEST_GENERAR_EXCEL_AUNQUE_SEAN_REPETIDAS = False
ANALIZAR_DESCRIPCION_DETALLE = True

# postulacion:
MODO_POSTULACION = "CONFIRMADA"
DRY_RUN_POSTULACION = True           # ⚠️ A PROPOSITO sigue en True: el envio real
                                      # NO esta habilitado todavia. Cambiar a False
                                      # es la accion que dispara la prueba real, y
                                      # la tiene que hacer Sergio explicitamente.
MAX_POSTULACIONES_POR_CORRIDA = 1
PROBAR_POSTULACION_EN_TEST = False

FUENTES = [bumeran, computrabajo]    # Indeed opcional via ACTIVAR_INDEED
```

Esta es EXACTAMENTE la config para la primera prueba real supervisada (sección 11),
salvo por `DRY_RUN_POSTULACION`, que queda en `True` a propósito hasta que Sergio
decida disparar la prueba de verdad.

## 5. Funcionalidades implementadas

- Scraping: Bumeran (Playwright, con fallback zona+fecha si el combinado da 0),
  Computrabajo (requests+BeautifulSoup), Indeed (opcional, detecta bloqueo Cloudflare).
- Filtros: fecha (`CT_DIAS`), zona (CABA/GBA), palabras clave en título, lista negra.
- Historial: `vistos.json` (dedup por link) + `historial_trabajos.xlsx` (acumulado).
- Clasificación `POSTULAR/REVISAR/DESCARTAR` por reglas + `categoria_detectada`
  (Data/BI, Soporte IT, QA/Testing, Desarrollo, Analista Funcional, Supply Chain,
  Administrativo/Procesos, Técnico/Producción, Otro).
- Análisis opcional de descripción del aviso (alertas: seniority, inglés alto,
  monotributo, posible estafa) — `ANALIZAR_DESCRIPCION_DETALLE`.
- Fichas de postulación (`porque_conviene`, `encaje_perfil`, `mensaje_sugerido`,
  `accion_recomendada`) usando `perfil_sergio.txt` como referencia.
- Excel: hoja RESUMEN (con bloque de acciones sugeridas), hoja TODOS, una hoja por
  fuente, hoja ACCIONES — todo con formato (freeze, autofiltro, colores por decisión,
  hipervínculos). Guardado en `resultados/YYYY-MM/`.
- Modos: `AUTO` / `MANUAL` / `TEST` / `DEMO` (este último con datos simulados, sin
  internet, sin tocar historial).
- Lock (`buscador.lock`) contra doble ejecución, con detección de lock viejo/abandonado.
- **Bandeja de acciones**: `accion_sugerida` (POSTULAR HOY / REVISAR ANTES DE POSTULAR /
  REVISAR MANUALMENTE / NO ACCIONAR) por reglas sobre alertas graves/moderadas.
- **Postulación asistida/automática segura** — ver sección 7.

## 6. Cambios recientes (más nuevo primero)

- **Fix: `clasificar_decision()` llegaba a POSTULAR solo con señales de
  contexto (ubicación/modalidad/rubro), sin ningún skill técnico real.**
  - *Por qué*: la corrida MANUAL del 13/07 (19:16) marcó "Administrativo de
    Programa de Seguimiento de Pacientes" (Bumeran) como POSTULAR HOY. Al
    diagnosticar la fila: `motivo_decision` mostraba "match fuerte: ai,
    híbrido, capital federal, administrativo" — 4 matches contra
    `DECISION_POSTULAR_MIN_MATCHES=2` de esa época. Ninguno de los 4 era un
    skill técnico real. Peor: `"ai"` era un **falso positivo de substring**,
    matcheaba dentro de **"Buenos Aires"** (bu-en-os-**ai**-res), no
    inteligencia artificial. Sin ese bug, igual habría llegado a POSTULAR
    solo con `híbrido` + `capital federal` + `administrativo` — puro
    contexto, cero técnico.
  - *Fix aplicado*: se separó `DECISION_POSTULAR_KEYWORDS` en dos listas —
    `DECISION_POSTULAR_KEYWORDS_FUERTES` (skills/roles técnicos reales: sql,
    python, power bi, data analyst, soporte it, mesa de ayuda, help desk, qa,
    tester, testing, desarrollador/programador junior, developer, trainee,
    junior, analista funcional, ai trainer, prompt evaluator, data
    annotation, annotator, automation, back office sistemas, administrativo
    sistemas) y `DECISION_POSTULAR_KEYWORDS_CONTEXTO` (híbrido, remoto, caba,
    capital federal, buenos aires, administrativo — solo suman prioridad y
    aparecen en el motivo, nunca deciden POSTULAR por sí solas). Nueva regla
    en `clasificar_decision()`: POSTULAR requiere
    `DECISION_POSTULAR_MIN_MATCHES_FUERTES=2` **fuertes**; con 1 fuerte o con
    0 fuertes (aunque haya contexto) queda en REVISAR, nunca POSTULAR.
  - **Bug de substring corregido**: nuevo helper `_matchea_keyword()` +
    `DECISION_KEYWORDS_PALABRA_ENTERA = {"ai","ia","it","bi","qa","sql"}` —
    esas 6 keywords cortas matchean SOLO como palabra entera (regex `\b`),
    nunca como substring. Antes "ai" matcheaba "Buenos Aires", "it" matcheaba
    "digital", etc.
  - *Qué se probó*: `py -3.14 -m py_compile` → compila. Tests unitarios del
    helper (`ai` en "buenos aires" → False; `ai` en "ai trainer" → True; `it`
    en "digital" → False; `it` en "analista it" → True) → los 4 correctos.
    Tests de `clasificar_decision()` sobre 4 títulos sintéticos (admin+
    híbrido+CABA → REVISAR; Analista de Datos Junior → POSTULAR; Soporte IT
    Junior Mesa de Ayuda → POSTULAR; Administrativo contable Senior →
    DESCARTAR por "senior") → los 4 correctos. **Reprocesada la fila real**
    de "Administrativo de Programa de Seguimiento de Pacientes" con sus datos
    exactos del Excel → `decision_sugerida` pasó de POSTULAR a **REVISAR**
    (motivo: solo contexto — híbrido, capital federal, buenos aires,
    administrativo — 0 fuertes), por lo tanto ya no puede llegar a POSTULAR
    HOY.
  - *No se tocó*: postulación real, `DRY_RUN_POSTULACION`, `truststore`,
    historial. No se corrió el buscador completo todavía con este fix (queda
    pendiente antes de revisar las otras 2 ofertas POSTULAR HOY de esa
    corrida).
  - *Pendiente*: revisar si las otras 2 ofertas que dieron POSTULAR HOY en la
    corrida del 13/07 19:16 ("Administrativo Contable" y "Administrativo/a
    Contable") siguen en POSTULAR con la regla nueva, o si eran el mismo
    patrón de falso positivo de contexto.

- **Mejora de calidad de búsqueda: más foco IT/Data/Soporte/Junior, menos ruido
  industrial.** Motivo: la corrida MANUAL del 13/07 (ver entrada de validación abajo)
  trajo 134 avisos crudos (Bumeran 48 + Computrabajo 86) pero solo 2 nuevas, ambas
  DESCARTAR — 0 POSTULAR, 0 REVISAR. Diagnóstico: `BUSQUEDAS`/`PERFIL_KEYWORDS`/
  `TITULO_DEBE_CONTENER` tenían mucho peso industrial (mecánico, operario,
  mantenimiento, producción) que diluía el volumen real de perfil, y un bug real:
  `TITULO_DEBE_CONTENER` no tenía ninguna palabra de la familia AI/prompt/trainee
  (`ai`, `trainer`, `prompt`, `annotation`, `junior`, `trainee`, etc.) — un aviso
  titulado literalmente "AI Trainer" se perdía en ese filtro antes de llegar a
  relevancia, aunque estuviera en `BUSQUEDAS`.
  - **`BUSQUEDAS`**: agregados 11 términos de prioridad (SQL junior, Power BI junior,
    soporte IT junior, help desk, QA tester junior, trainee IT, analista funcional
    junior, administrativo sistemas, back office sistemas, automation trainee,
    prompt evaluator). Recortado el bloque industrial explícito (tecnico mecanico,
    operario de producción, mantenimiento industrial, supervisor de producción, jefe
    de almacén, y sueltas operario/mantenimiento/supervisor/tecnico/produccion/
    almacen/mechanic/maintenance/production operator).
  - **`PERFIL_KEYWORDS`**: quitadas mecanico/produccion/operario/mantenimiento/
    industrial/maquinas (fuera de perfil). Agregadas ai/annotator/evaluator/
    helpdesk/back office/automation.
  - **`TITULO_DEBE_CONTENER`** (fix del bug de arriba): agregadas ai/ia/trainer/
    prompt/annotation/annotator/evaluator/junior/trainee/helpdesk. Quitadas
    mecanico/mantenimiento/produccion/operario/electromecanico.
  - **`TITULO_NO_DEBE_CONTENER`**: agregadas ventas, "capacitación y desarrollo",
    mantenimiento/operario/producción/mecánico/electromecánico/"técnico industrial".
    A propósito NO se agregó "senior" acá (sigue viva en `DECISION_DESCARTAR_KEYWORDS`,
    para que quede visible como DESCARTAR en el Excel en vez de desaparecer antes de
    tiempo — mismo criterio de diseño del proyecto: nunca eliminar, solo marcar).
  - **`DECISION_POSTULAR_KEYWORDS`**: agregadas help desk/ai/annotation/annotator/
    trainer/prompt/back office/automation — antes esas categorías no sumaban ningún
    match para POSTULAR/REVISAR aunque sobrevivieran los filtros de arriba.
  - **`DECISION_POSTULAR_MIN_MATCHES`**: bajado de `3` a `2` (título corto rara vez
    junta 3 keywords; REVISAR sigue en 1+ match, sin cambios ahí).
  - *Qué se probó*: `py -3.14 -m py_compile buscador_trabajos_v2.py` → compila sin
    errores. Todavía NO se corrió una MANUAL completa con las listas nuevas (queda
    pendiente, ver sección 10).
  - *No se tocó*: postulación real, `DRY_RUN_POSTULACION`, `truststore`, historial
    (`vistos.json`/`historial_trabajos.xlsx`), lock, Programador de tareas de
    Windows.

- **Validación del fix SSL en corrida MANUAL completa real (2026-07-13 10:27:30).**
  Corrida `MODO_EJECUCION="MANUAL"` de punta a punta contra Bumeran y Computrabajo
  reales, con `DRY_RUN_POSTULACION=True` sin tocar y respuesta automática "n" a
  cualquier posible prompt de confirmación (no hizo falta: 0 candidatas).
  - *Resultado*: **Computrabajo funcionó en el flujo completo** — 86 resultados
    crudos (antes del fix: 0, siempre `CERTIFICATE_VERIFY_FAILED`). Bumeran: 48
    resultados crudos. Sin ningún SSLError durante toda la corrida.
  - *Clasificación*: POSTULAR: 0 | REVISAR: 0 | DESCARTAR: 2. Acciones sugeridas:
    POSTULAR HOY: 0 (el resto en 0 también). Como no hubo POSTULAR HOY, no se activó
    el flujo de postulación (`ejecutar_postulaciones` no tuvo nada que procesar, no
    se llamó `input()`), y no se generó `postulaciones_log.xlsx`.
  - *Excel generado*: `resultados\2026-07\trabajos_2026-07-13_10-27-30.xlsx` (2
    ofertas nuevas: "Analista de Capacitación y Desarrollo" y "Administrativo
    contable Senior", ambas DESCARTAR). Duración total: 34m 54s.
  - *Seguridad*: no se envió nada real, `DRY_RUN_POSTULACION` siguió en `True` sin
    tocarse, no se hizo commit en esta corrida.
  - *Intento previo (mismo día, descartado)*: una corrida anterior se cortó por una
    caída real de conectividad a mitad de ejecución (`ERR_INTERNET_DISCONNECTED` /
    `NameResolutionError` en ambas fuentes) — no relacionado al fix SSL, solo
    confirma que sin red ninguna fuente funciona. Se repitió una vez restablecida la
    conexión, con el resultado de arriba.

- **Fix SSL: Computrabajo devolvía 0 resultados por `CERTIFICATE_VERIFY_FAILED`.**
  Agregado `truststore` (nuevo en `requirements.txt`), inyectado con
  `truststore.inject_into_ssl()` justo después de `import requests`, envuelto en
  `try/except ImportError` (si no está instalado, sigue con certifi/default sin
  romper el script, solo avisa por consola).
  - *Por qué*: en esta máquina, **Norton Antivirus intercepta HTTPS** ("SSL/TLS
    scanning") y re-firma los certificados con su propia CA
    (`Norton Web/Mail Shield Root`). Windows confía en esa root (por eso
    Playwright/Bumeran nunca tuvo problema), pero el bundle propio de certifi que
    usa `requests` (solo lo usa `computrabajo()`) no la incluye → cada request
    fallaba con `SSLError: CERTIFICATE_VERIFY_FAILED`.
  - *Cómo se diagnosticó*: se probó request mínimo con `verify` default y con
    `verify=certifi.where()` explícito → mismo error en ambos (descartaba certifi
    desactualizado, certifi ya estaba en última versión). Se inspeccionó la cadena
    de certificado real del server con `openssl s_client -connect
    ar.computrabajo.com:443` → confirmó el certificado re-firmado por Norton.
  - *Por qué esta solución y no otra*: se descartó `verify=False` (desactiva
    verificación real). `truststore` hace que el `ssl` de Python use el almacén de
    certificados del sistema operativo — el mismo que ya usa el navegador/Playwright
    — en vez del bundle aislado de certifi. Sigue verificando de verdad, solo iguala
    el trust anchor.
  - *Qué se probó*: `pip install -r requirements.txt` (truststore 0.10.4 instalado
    sin conflictos); request mínimo a Computrabajo con truststore inyectado → ya sin
    SSLError; `computrabajo("analista")` corrida real → devolvió 1 oferta real (antes
    0); `py -3.14 -m py_compile buscador_trabajos_v2.py` → compila sin errores.
  - *No se tocó*: Bumeran/Playwright (no pasan por `requests`, no les afecta),
    postulación real, `DRY_RUN_POSTULACION`. No se agregó `CT_VERIFY_SSL=False` ni
    ningún switch para desactivar verificación.
  - *Pendiente*: confirmar en la próxima corrida `MANUAL` completa que Computrabajo
    vuelve a aportar candidatos reales al Excel (esta prueba fue solo con
    `computrabajo()` aislado, no corrida completa del script).

- **Fix: la bandeja de ACCIONES enterraba REVISAR en NO ACCIONAR por culpa de la
  palabra "senior".** `determinar_accion_sugerida()` usaba `_ALERTAS_GRAVES`
  (compartida con `_postulacion_es_segura`, que SÍ incluye `"senior"`/`"5 años"` etc.
  como graves) para decidir `NO ACCIONAR` — eso pisaba la rama `REVISAR MANUALMENTE`
  antes de que se evaluara. Se separaron 2 listas nuevas, propias de esta capa:
  `_ALERTAS_GRAVES_ACCION` (estafa/portal externo real: pagar para postular, curso
  pago, zonajobs, google forms, hiringroom, linkedin externo, etc. — siempre
  NO ACCIONAR) y `_ALERTAS_SENIORITY_ACCION` (senior, ssr, contractor, monotributo,
  inglés avanzado, excluyente, etc. — en POSTULAR bajan a "REVISAR ANTES DE
  POSTULAR", en REVISAR ya no fuerzan nada, cae en "REVISAR MANUALMENTE" como
  corresponde).
  - *Por qué*: la corrida MANUAL real (12/07) dio 2 ofertas REVISAR (Copilot
    Developer, Técnico de mantenimiento) y las 2 terminaron en NO ACCIONAR en vez de
    REVISAR MANUALMENTE, dejando esa categoría en 0 — invisibles para Sergio aunque
    fueran legítimamente revisables.
  - *Qué se probó*: 5 casos unitarios de `determinar_accion_sugerida()` (POSTULAR sin
    alertas → POSTULAR HOY; POSTULAR con senior → REVISAR ANTES DE POSTULAR; REVISAR
    con senior → REVISAR MANUALMENTE; REVISAR con "pagar para postular" → NO ACCIONAR;
    POSTULAR con zonajobs → NO ACCIONAR) — los 5 correctos. Se simularon las 2 filas
    reales del Excel de la corrida (`Copilot Developer` con
    `senior; contractor; salario bajo`, `Técnico de mantenimiento` con `senior`) → las
    2 ahora dan `REVISAR MANUALMENTE`, ninguna pasó a `POSTULAR HOY` (a propósito, no
    se aflojó el clasificador principal).
  - *No se tocó*: `_ALERTAS_GRAVES` original (la usa `_postulacion_es_segura`, el gate
    del envío real), `clasificar_decision`, `ajustar_decision_por_descripcion`,
    `DRY_RUN_POSTULACION`, ni el flujo de postulación real.
  - *Pendiente*: el falso positivo de `"salario bajo (~$11.111)"` en Copilot Developer
    (probablemente el regex de salario agarró un número que no es un sueldo) — no se
    tocó, queda para revisar aparte.

- **Config preparada para la 1ra prueba real supervisada, sin dispararla.**
  `MODO_EJECUCION` cambiado de `"TEST"` a `"MANUAL"` en el archivo; el resto
  (`MODO_POSTULACION="CONFIRMADA"`, `MAX_POSTULACIONES_POR_CORRIDA=1`,
  `PROBAR_POSTULACION_EN_TEST=False`) ya estaba correcto.
  - *Por qué*: dejar todo listo para que Sergio dispare la primera postulación real
    supervisada apenas decida hacerlo, sin tener que tocar config en el momento.
  - *Qué se probó*: solo se verificó que compile (`py -3.14 -m py_compile`). No se
    corrió el script.
  - *Qué queda pendiente*: `DRY_RUN_POSTULACION` sigue en `True` a propósito — no se
    tocó, no se ejecutó el buscador, no se hizo commit. El paso de cambiar
    `DRY_RUN_POSTULACION=False` y correrlo es explícitamente de Sergio.
  - *Advertencia de seguridad*: ninguna, este cambio no habilita nada nuevo — solo
    prepara la config. La barrera real (`DRY_RUN_POSTULACION=True`) sigue intacta.

- **Conectado el click real de postulación**, SOLO para
  `MANUAL + CONFIRMADA + DRY_RUN_POSTULACION=False + confirmación humana doble`.
  Nuevas funciones: `detectar_boton_postulacion`, `detectar_preguntas_o_bloqueos`,
  `enviar_postulacion_confirmada`, `_verificar_exito_postulacion`.
  - *Por qué*: Sergio pidió avanzar de "solo simular" a poder enviar de verdad, pero
    exclusivamente con supervisión humana en el momento, nunca desatendido.
  - *Qué se probó*: en vivo contra Bumeran/Computrabajo reales (solo lectura, sin
    click), y con mocks para la lógica de branching completa. Ver sección 9.
  - *Bugs encontrados y corregidos en el proceso*:
    1. La lista de textos de botón no incluía el texto real de Bumeran
       (`"Postulación rápida"`) — se agregó, verificado en vivo.
    2. `detectar_preguntas_o_bloqueos` daba falso positivo SIEMPRE por el link de
       footer "Preguntas frecuentes" (contiene la palabra "pregunta") — corregido
       excluyendo esa frase específica antes de buscar keywords.
  - *Advertencia de seguridad*: nunca se probó un click real de verdad (ver sección 9
    para el razonamiento). La primera corrida con `DRY_RUN_POSTULACION=False` real
    debe hacerla Sergio en persona, mirando la consola.

- Agregado `MODO_EJECUCION="TEST"` + `PROBAR_POSTULACION_EN_TEST` para poder probar el
  flujo de confirmación sin esperar ofertas nuevas reales. Nunca envía si
  `MODO_EJECUCION` no es `MANUAL` (ni siquiera con este switch en `True` — solo permite
  *probar* con `DRY_RUN=True`, nunca enviar de verdad desde TEST).

- Agregado `MODO_POSTULACION="CONFIRMADA"`: pregunta por consola, aviso por aviso,
  antes de intentar nada. Solo funciona con `MODO_EJECUCION="MANUAL"` (guard en 2
  capas: `ejecutar_postulaciones` y `_postular_oferta_logica`, para que `input()`
  nunca se llame en una corrida desatendida).

- Agregada infraestructura de postulación asistida/automática (`MODO_POSTULACION`:
  `OFF`/`ASISTIDA`/`AUTO_SEGURO`), `postular_oferta()`, `ejecutar_postulaciones()`,
  `postulaciones_log.xlsx`. En esa primera versión, el click real NO estaba conectado
  en ningún modo (decisión deliberada de seguridad, ver sección 8).

- Hoja ACCIONES en el Excel (`construir_acciones`, `determinar_accion_sugerida`),
  bloque "Acciones sugeridas" en RESUMEN y en el mensaje final.

- `categoria_detectada`, formato visual del Excel (RESUMEN, colores, freeze/autofiltro
  por hoja calculado dinámicamente — importante porque ACCIONES tiene columnas
  distintas a TODOS).

- `MODO_EJECUCION="DEMO"` con `generar_datos_demo()` (6 ofertas simuladas, sin
  internet, sin tocar historial).

- **Sync + push a `origin/main` (commit `dc18b91`)**: script y `docs/continuidad_claude.md`
  copiados desde la carpeta raíz hacia acá (repo git), con backup del script viejo
  (`buscador_trabajos_v2_BACKUP_antes_sync.py`, gitignoreado). `.gitignore` ampliado
  (`postulaciones_log.xlsx`, `perfil_sergio.txt`, el backup). `_PERFIL_SERGIO_CONTENIDO_BASE`
  despersonalizado (nombre real → "Perfil de ejemplo", sin universidad). `git status`
  quedó limpio tras el push. Ver advertencias en sección 2 sobre la carpeta raíz
  (ahora superada) y sobre `perfil_sergio.txt` ya presente en commits viejos.

- README.md, requirements.txt, .gitignore, docs/guion_presentacion.md,
  docs/estructura_proyecto.md — creados originalmente en esta misma carpeta (ya era
  el repo git desde el principio; la raíz nunca los tuvo).

## 7. Estado de la postulación asistida/confirmada

| Modo | Envía de verdad? |
|---|---|
| `OFF` | No, no hace nada. |
| `ASISTIDA` | No, nunca. Abre y verifica, se detiene antes de cualquier click. |
| `AUTO_SEGURO` | No, nunca (deliberado — ver sección 8). Con `DRY_RUN=True` simula. |
| `CONFIRMADA` + `MANUAL` + `DRY_RUN=False` + usuario confirma 2 veces | **Sí, esta es la única combinación que puede hacer click real.** |
| `CONFIRMADA` + cualquier otro `MODO_EJECUCION` | No, nunca (ni con `PROBAR_POSTULACION_EN_TEST=True`). |

El click real, cuando se dispara, sigue este flujo (`enviar_postulacion_confirmada`):
detecta botón único → revisa bloqueos/preguntas → pide **última confirmación** por
consola → click → revisa bloqueos otra vez → si aparece un segundo botón, pide
confirmación de nuevo → busca texto de éxito explícito → si no lo encuentra, marca
`"requiere revisión manual" / "estado incierto"` (nunca asume éxito).

## 8. Switches de seguridad

- `DRY_RUN_POSTULACION` (default `True`): mientras esté en `True`, nunca se hace un
  click real en ningún modo.
- `MODO_POSTULACION` debe ser exactamente `"CONFIRMADA"` para que exista *algún*
  camino hacia un click real. `AUTO_SEGURO` con `DRY_RUN=False` sigue devolviendo
  `"requiere revisión manual"` — el envío automático desatendido NO está implementado
  a propósito (irreversible hacia un tercero, requiere confirmación humana en el
  momento).
- `MODO_EJECUCION` debe ser exactamente `"MANUAL"` en el momento del click real —
  chequeado en 2 lugares (`_confirmada_permitida()` y de nuevo dentro de
  `_postular_oferta_logica`).
- `MAX_POSTULACIONES_POR_CORRIDA`: tope duro de cuántas ofertas se procesan.
- Lista de dominios propios (`_DOMINIOS_POSTULACION_PROPIOS`): solo bumeran.com.ar y
  computrabajo.com. Cualquier otro dominio → no se procesa.
- `_ALERTAS_GRAVES`: si el aviso tiene alguna, nunca se postula (ni siquiera se
  pregunta).
- `detectar_boton_postulacion`: si hay 0 o 2+ botones candidatos, devuelve `None` —
  nunca adivina cuál clickear.
- `detectar_preguntas_o_bloqueos`: si aparece cualquier señal de pregunta
  obligatoria/adjunto/CAPTCHA, frena — nunca inventa una respuesta.
- El agente **nunca** decidió por sí solo hacer un click real durante esta sesión de
  desarrollo — todo lo que se "envió" en las pruebas fue con botones/páginas mockeados
  en Python, no con Playwright real.

## 9. Pruebas realizadas y resultados

Todas hechas en esta sesión, contra Bumeran/Computrabajo reales cuando fue posible:

- `detectar_boton_postulacion` contra oferta real de Bumeran → encontró
  `"Postulación rápida"` como único candidato (tras el fix del bug #1 de sección 6).
- `detectar_preguntas_o_bloqueos` contra la misma página → `False` (sin bloqueo), tras
  el fix del bug #2.
- `_confirmada_permitida()`: 5 combinaciones de `MODO_EJECUCION`/
  `PROBAR_POSTULACION_EN_TEST` probadas, las 5 correctas.
- `input()` nunca se llama fuera de `MANUAL` (o `TEST` con el switch) — verificado
  reemplazando `input` por una función que explota si se invoca.
- `_get_page()` nunca se llama cuando `MODO_POSTULACION="OFF"` o cuando
  `CONFIRMADA` no está permitida — verificado igual, con AssertionError si se llama.
- Flujo `MANUAL + CONFIRMADA + DRY_RUN=True + "s"` → `"simulado"`, sin click.
- Flujo `TEST + DRY_RUN=False + "s"` → `"requiere revisión manual"`, sin llamar a
  `enviar_postulacion_confirmada` (verificado con assert).
- Flujo `AUTO_SEGURO + DRY_RUN=False` → sigue bloqueado (regresión, sin cambios).
- `enviar_postulacion_confirmada` con botón y página **mockeados en Python** (nunca
  Playwright real): responder "n" → sin click; responder "s" sin texto de éxito →
  `"requiere revisión manual"/"estado incierto"`; responder "s" con texto de éxito →
  `"postulado"`.
- Formato de Excel (RESUMEN primero, ACCIONES con columnas propias, colores por
  decisión, hipervínculos) verificado reabriendo archivos reales generados.
- Los 4 `MODO_EJECUCION` (AUTO/MANUAL/TEST/DEMO) probados con sus combinaciones
  correctas de interruptores auto-configurados.

**Lo que NUNCA se probó**: un click real de verdad contra una oferta real con
`DRY_RUN_POSTULACION=False`. Decisión deliberada — ver sección 8 del razonamiento en
el resumen de la tarea anterior. Esa primera prueba la debe hacer Sergio en persona.

## 10. Qué quedó pendiente

- Probar el primer envío real (`DRY_RUN_POSTULACION=False`, `MODO_EJECUCION="MANUAL"`,
  `MODO_POSTULACION="CONFIRMADA"`) con una oferta real elegida por Sergio, mirando la
  consola.
- La lista de `_TEXTOS_EXITO_POSTULACION` no fue verificada contra una postulación
  real exitosa (no hay forma de verificarla sin postular de verdad) — puede necesitar
  ajuste cuando se vea el mensaje real de éxito de Bumeran/Computrabajo.
- `_TEXTOS_BOTON_POSTULACION` solo se verificó en una oferta de cada fuente — otros
  tipos de aviso (ej. con formulario extendido, o de empresas que usan un flujo
  distinto) podrían usar textos distintos todavía no vistos.
- Decidir si vale la pena reescribir el historial de git para sacar el
  `perfil_sergio.txt` real del commit `52c8c3d` (con nombre completo), o dejarlo
  como está.
- La carpeta raíz (`...\Buscador de trabajo\`) quedó con su propia copia desactualizada
  del script post-sync — no es un problema en sí, pero si se sigue usando por hábito
  hay que acordarse de que ya no es la fuente de verdad.

## 11. Próximo paso recomendado

**Todo queda preparado para la primera prueba real supervisada, pero SIN dispararla.**
La config del archivo (sección 4) ya tiene `MODO_EJECUCION="MANUAL"`,
`MODO_POSTULACION="CONFIRMADA"`, `MAX_POSTULACIONES_POR_CORRIDA=1`,
`PROBAR_POSTULACION_EN_TEST=False` — todo listo salvo `DRY_RUN_POSTULACION`, que
sigue en `True` a propósito.

El paso que falta, y que tiene que decidir y ejecutar Sergio explícitamente:
1. Cambiar `DRY_RUN_POSTULACION` a `False` en el archivo.
2. Correr el script sobre UNA oferta elegida a mano.
3. Responder los 2 prompts de confirmación por consola.
4. Ver qué pasa de verdad contra `detectar_boton_postulacion`/
   `enviar_postulacion_confirmada` en el sitio real.

Recién después de esa prueba conviene seguir iterando sobre esa capa (ej. ajustar
`_TEXTOS_EXITO_POSTULACION` con el mensaje real que aparezca — ver sección 10).

Ninguna sesión de Claude Code debe cambiar `DRY_RUN_POSTULACION` a `False` ni correr
el script en este estado sin que Sergio lo pida explícitamente en el momento.

## 12. Qué NO se debe tocar

Estas partes están estables y probadas — no modificar salvo pedido explícito:

- Scraping base de Bumeran (`bumeran()`, `_bumeran_cargar_listado`,
  `_clasificar_h3_bumeran`) y Computrabajo (`computrabajo()`, `_parse_tarjeta_ct`).
- Indeed (`indeed()`) — opcional, bloqueo Cloudflare ya manejado.
- `CT_DIAS = 2` (o el valor que tenga en cada momento — no cambiarlo sin que lo pida
  Sergio explícitamente).
- Historial (`_cargar_vistos`, `_guardar_vistos`, `_actualizar_historial_xlsx`,
  `vistos.json`, `historial_trabajos.xlsx`).
- Lock (`adquirir_lock`, `liberar_lock`, `buscador.lock`).
- Carpetas `resultados/YYYY-MM/` (`construir_ruta_excel`).
- Los 4 modos de ejecución y su bloque `if/elif` de auto-configuración.
- Clasificación (`clasificar_decision`), fichas (`generar_ficha_postulacion`),
  categoría (`detectar_categoria`).
- Programador de tareas de Windows (`run_buscador_trabajos.bat`, tarea
  "BuscadorTrabajos", 10:00 diario, `MODO_EJECUCION="AUTO"`).
- El punto de diseño de seguridad de la sección 8 en general: no conectar envío real
  desatendido bajo ninguna circunstancia sin que Sergio lo pida explícitamente y con
  plena conciencia de las implicancias.

## 13. Comandos útiles para probar

```powershell
# Verificar sintaxis
py -3.14 -m py_compile buscador_trabajos_v2.py

# Correr el script tal cual esta configurado en el panel
py -3.14 buscador_trabajos_v2.py

# Probar una función puntual sin correr todo el script (ejemplo)
py -3.14 -c "import buscador_trabajos_v2 as b; print(b.MODO_POSTULACION)"

# Ver botones reales de una oferta de Bumeran (solo lectura, sin click)
py -3.14 -c "
import buscador_trabajos_v2 as b
b.NAVEGADOR_VISIBLE = False
page = b._get_page()
page.goto('URL_DE_UNA_OFERTA_REAL', wait_until='domcontentloaded')
print(b.detectar_boton_postulacion(page, 'Bumeran'))
b._cerrar_navegador()
"
```

Antes de correr con `MODO_POSTULACION="CONFIRMADA"` y `DRY_RUN_POSTULACION=False`:
confirmar que `MODO_EJECUCION="MANUAL"` y tener a mano la consola para responder los
2 prompts de confirmación.
