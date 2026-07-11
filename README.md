# Agente Buscador Laboral

Agente en Python que busca ofertas laborales en Bumeran y Computrabajo, filtra por criterios personalizados, evita repetidos usando historial, clasifica oportunidades y genera reportes en Excel.

## Versión actual

v1.0 - Agente buscador laboral con búsqueda automática, historial, clasificación, reportes Excel y modo demo.

## Demo

El proyecto incluye `MODO_EJECUCION = "DEMO"`, que permite generar un reporte de demostración sin depender de internet ni modificar el historial real.

## Funcionalidades principales

- Búsqueda automática en portales laborales (Bumeran, Computrabajo; Indeed opcional).
- Filtro por fecha máxima de publicación (`CT_DIAS`).
- Filtro por zona (CABA / GBA).
- Filtro por palabras clave en el título.
- Exclusión de rubros no deseados (lista negra).
- Historial por URL con `vistos.json` — no repite avisos ya vistos.
- Historial acumulado en `historial_trabajos.xlsx`.
- Clasificación automática: **POSTULAR / REVISAR / DESCARTAR**.
- Detección de categoría laboral (Data/BI, Soporte IT, QA/Testing, Desarrollo, Analista Funcional, Supply Chain, Administrativo/Procesos, Técnico/Producción, Otro).
- Análisis opcional de la descripción completa del aviso (alertas de seniority, inglés alto, monotributo, posibles estafas, etc.).
- Generación de Excel en carpetas mensuales: `resultados/YYYY-MM/`.
- Modos de ejecución **AUTO / MANUAL / TEST / DEMO**.
- Lock por archivo para evitar dos ejecuciones simultáneas.
- Ejecución automática con el Programador de tareas de Windows.

## Modos de ejecución

Se controlan con una sola variable: `MODO_EJECUCION` en el panel de interruptores del script.

| Modo | Uso | Comportamiento |
|---|---|---|
| **AUTO** | Uso diario automático (Programador de tareas) | Sin ventanas, sin abrir Excel, sin análisis de descripción — corre rápido y desatendido. |
| **MANUAL** | Cuando lo corrés vos a mano | Abre el Excel al terminar, muestra cartel/resumen final, análisis de descripción activado. |
| **TEST** | Probar mejoras sin esperar avisos nuevos | Genera Excel aunque las ofertas ya estén vistas (usa las repetidas), **no modifica el historial**. |
| **DEMO** | Presentaciones y pruebas de formato | Genera un Excel con 6 ofertas simuladas (`generar_datos_demo()`), **sin conexión a internet** (no llama a Bumeran, Computrabajo ni Indeed) y **sin tocar `vistos.json` ni `historial_trabajos.xlsx`**. Pasa por la misma clasificación, categorización y fichas que un aviso real. |

### Sobre el MODO DEMO

Pensado para mostrar el proyecto (entrevistas, portfolio, verificar que el formato del Excel se vea bien) sin depender de que Bumeran/Computrabajo estén arriba o de tener conexión. Las 6 ofertas simuladas cubren a propósito los tres veredictos: 2 POSTULAR (Data/BI y Soporte IT), 2 REVISAR (Analista Funcional y Supply Chain) y 2 DESCARTAR (una por seniority alta, otra por rubro no deseado — call center). El Excel se guarda en `resultados/YYYY-MM/trabajos_DEMO_YYYY-MM-DD_HH-MM-SS.xlsx` y la hoja RESUMEN aclara explícitamente que es una corrida de demostración.

## Archivos importantes

| Archivo | Descripción |
|---|---|
| `buscador_trabajos_v2.py` | Script principal — todo el agente vive acá. |
| `requirements.txt` | Dependencias externas del proyecto. |
| `perfil_sergio.txt` | Resumen de perfil usado como referencia para clasificación y fichas de postulación. Se autogenera la primera vez si no existe. |
| `vistos.json` | Historial de links ya vistos (clave = URL del aviso). |
| `historial_trabajos.xlsx` | Historial acumulado de todas las ofertas nuevas detectadas a lo largo del tiempo. |
| `log_buscador.txt` | Log de consola de las corridas (si se redirige la salida al ejecutar). |
| `resultados/YYYY-MM/*.xlsx` | Reportes Excel del día, organizados por mes. |
| `run_buscador_trabajos.bat` | Lanzador usado por el Programador de tareas de Windows. |

## Cómo ejecutar manualmente

Instalar dependencias (una sola vez):

```
py -3.14 -m pip install -r requirements.txt
py -3.14 -m playwright install chromium
```

Correr el script:

```
py -3.14 buscador_trabajos_v2.py
```

Recomendado: cambiar `MODO_EJECUCION = "MANUAL"` antes de correrlo así, para ver el resumen y que abra el Excel solo.

## Cómo se ejecuta automáticamente

Tarea programada de Windows llamada **BuscadorTrabajos**, configurada para correr todos los días a las **10:00**, usando `run_buscador_trabajos.bat` con `MODO_EJECUCION = "AUTO"`.

## Arquitectura general

```
Configuración + perfil
    → búsqueda en portales
        → extracción de ofertas
            → filtros (zona, relevancia, título, lista negra)
                → historial / deduplicación (vistos.json)
                    → clasificación (POSTULAR / REVISAR / DESCARTAR)
                        → análisis de detalle (opcional)
                            → reporte Excel (resultados/YYYY-MM/)
                                → alerta / log
```

## Limitaciones

- Depende de la estructura HTML de los portales — si Bumeran o Computrabajo cambian su sitio, los selectores pueden romperse.
- Indeed queda opcional por defecto porque suele bloquear con Cloudflare.
- La clasificación es por reglas simples (keywords), no por IA generativa.
- El análisis de perfil usa `perfil_sergio.txt` como resumen de texto — no lee ni parsea un CV real (PDF/Word) automáticamente.
- Algunas páginas pueden bloquear el scraping (CAPTCHA, rate limiting).

## Próximas mejoras posibles

- Integración con CV real (lectura de PDF/Word).
- Generación de cartas de postulación más específicas por oferta.
- Panel web simple para revisar resultados sin abrir Excel.
- Base SQLite en lugar de JSON/Excel para el historial.
- Revisión manual de postulaciones (marcar como "ya postulé").
