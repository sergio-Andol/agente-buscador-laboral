# Continuidad del proyecto — Agente Buscador Laboral

> Handoff para una nueva sesión de Claude Code. Última actualización: 2026-07-12.
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

**Carpeta activa (donde se trabaja, la que tiene todo lo último):**
```
C:\Users\sergi\Documentos\Buscador de trabajo\
```

⚠️ **Advertencia importante:** existe una subcarpeta `Agente-Buscador-Laboral\` dentro
de la carpeta activa, con su propio repo git (`.git`) y una copia **desactualizada**
del script (95 KB, del 11/07 ~16:03 — sin toda la capa de postulación confirmada/click
real). Esa subcarpeta parece ser una copia "presentable" para portfolio/GitHub, armada
por Sergio a mano. **No confundir las dos copias.** La carpeta raíz (130+ KB) es la
única que se debe editar. Si en algún momento hay que sincronizar la copia de
portfolio, eso lo decide Sergio explícitamente — no asumir que hay que copiar cambios
para allá.

## 3. Archivo principal

```
buscador_trabajos_v2.py
```
Todo el agente vive en este único archivo (~2900 líneas a esta fecha).

## 4. Configuración actual relevante

Panel de interruptores al principio del archivo. Valores **actuales en el archivo**
(no necesariamente los defaults recomendados — Sergio está en medio de pruebas):

```python
MODO_EJECUCION = "TEST"              # Sergio esta probando; default recomendado: "AUTO"
ACTIVAR_INDEED = False
GUARDAR_DEBUG_INDEED = True
DESCARTAR_BUMERAN_EXTERNOS = True
CT_DIAS = 2
SONIDO_FINAL = "SUAVE"

# auto-configuradas por MODO_EJECUCION (ver bloque if/elif):
NAVEGADOR_VISIBLE = False
MOSTRAR_CARTEL_FINAL = False
ABRIR_EXCEL_AL_FINAL = False
MODO_TEST_GENERAR_EXCEL_AUNQUE_SEAN_REPETIDAS = False
ANALIZAR_DESCRIPCION_DETALLE = False

# postulacion:
MODO_POSTULACION = "CONFIRMADA"      # Sergio probando; default recomendado: "OFF"
DRY_RUN_POSTULACION = True           # IMPORTANTE: sigue en True, no se envio nada real todavia
MAX_POSTULACIONES_POR_CORRIDA = 1
PROBAR_POSTULACION_EN_TEST = False

FUENTES = [bumeran, computrabajo]    # Indeed opcional via ACTIVAR_INDEED
```

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

- README.md, requirements.txt, .gitignore, docs/guion_presentacion.md,
  docs/estructura_proyecto.md — **creados dentro de la subcarpeta
  `Agente-Buscador-Laboral/`**, no en la raíz (ver advertencia en sección 2).

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
- README.md / requirements.txt / docs/ están en la subcarpeta de portfolio, no en la
  raíz activa — decidir si conviene tener una copia (o symlink) también acá, o dejarlo
  así a propósito.
- No se sincronizó nunca la copia de `Agente-Buscador-Laboral/` con los cambios
  recientes (sigue en la versión del 11/07 ~16:03, sin la capa de postulación
  confirmada).

## 11. Próximo paso recomendado

Antes de nada nuevo: que Sergio corra manualmente
`MODO_EJECUCION="MANUAL"`, `MODO_POSTULACION="CONFIRMADA"`,
`DRY_RUN_POSTULACION=False`, sobre UNA oferta elegida a mano, para validar que
`detectar_boton_postulacion`/`enviar_postulacion_confirmada` funcionan de punta a
punta contra el sitio real. Recién después de eso conviene seguir iterando sobre esa
capa (ej. ajustar `_TEXTOS_EXITO_POSTULACION` con el mensaje real que aparezca).

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
