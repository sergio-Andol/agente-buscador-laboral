# Guion de presentación — Agente Buscador Laboral

## Guion corto (1 minuto)

> Desarrollé un agente automatizado en Python para optimizar mi búsqueda laboral. El sistema consulta portales como Bumeran y Computrabajo, filtra ofertas por fecha, zona y palabras clave, evita repetidos con historial por URL, clasifica oportunidades según reglas personalizadas y genera reportes en Excel organizados por mes. También cuenta con modos AUTO, MANUAL, TEST y DEMO, lock para evitar doble ejecución y ejecución automática diaria mediante el Programador de tareas de Windows.

## Qué problema resuelve

Buscar trabajo a mano en varios portales todos los días es repetitivo: hay que repetir las mismas búsquedas, revisar avisos que ya viste ayer, y decidir manualmente cuáles valen la pena. Este agente automatiza esa rutina diaria y deja solo la parte que requiere criterio humano: decidir si postular.

## Qué hace el agente

1. Busca ofertas en Bumeran y Computrabajo (Indeed queda disponible como fuente opcional).
2. Filtra por fecha de publicación, zona geográfica y palabras clave del título.
3. Descarta rubros que no interesan (lista negra).
4. Compara contra un historial de avisos ya vistos para no repetir.
5. Clasifica cada oferta nueva como **POSTULAR / REVISAR / DESCARTAR**, con un motivo y una prioridad.
6. Detecta la categoría del puesto (Data/BI, Soporte IT, QA/Testing, Desarrollo, Analista Funcional, Supply Chain, Administrativo/Procesos, Técnico/Producción).
7. Opcionalmente entra al detalle de cada aviso para leer la descripción completa, resumirla y detectar alertas (seniority alta, inglés excluyente, posibles estafas).
8. Genera una ficha de postulación para cada oferta marcada como POSTULAR: por qué conviene, cómo encaja con mi perfil, un mensaje sugerido y una acción recomendada.
9. Arma un reporte Excel con formato (colores por decisión, hoja RESUMEN, hipervínculos) guardado en `resultados/YYYY-MM/`.
10. Avisa al terminar (sonido + cartel) y puede abrir el Excel automáticamente.

## Qué tecnologías usa

- **Python** como lenguaje principal.
- **Playwright** para scraping de sitios con contenido dinámico (Bumeran, Indeed).
- **requests + BeautifulSoup** para scraping de Computrabajo (HTML estático).
- **pandas** para el procesamiento tabular de las ofertas.
- **openpyxl** para generar y formatear el Excel final (colores, hipervínculos, hoja de resumen).
- **JSON** como almacenamiento simple de historial (`vistos.json`).
- **Programador de tareas de Windows** para la ejecución automática diaria.

## Qué automatiza

La búsqueda repetitiva en múltiples portales, la comparación contra lo ya visto, la clasificación de relevancia, y la generación de un reporte legible — todo con un solo comando (o de forma completamente desatendida, vía tarea programada).

## Cómo evita repetidos

Cada aviso tiene un link único que se usa como clave. Al finalizar cada corrida, los links nuevos se guardan en `vistos.json` (lookup rápido) y también se agregan a `historial_trabajos.xlsx` (registro acumulado, legible). En la siguiente corrida, cualquier aviso cuyo link ya esté en el historial se marca como repetido y no vuelve a aparecer en el reporte del día.

## Qué modos tiene

- **AUTO**: para la tarea programada — sin ventanas, corre rápido y desatendido.
- **MANUAL**: para correrlo a mano — muestra el resumen final y abre el Excel.
- **TEST**: genera un Excel de prueba reusando avisos ya vistos, sin tocar el historial — útil para probar cambios.
- **DEMO**: genera un Excel de ejemplo con datos simulados, sin conexión a internet ni impacto en el historial — pensado para mostrar el proyecto.

## Qué salida genera

Un archivo Excel por corrida (`resultados/YYYY-MM/trabajos_YYYY-MM-DD_HH-MM-SS.xlsx`) con:
- Hoja **RESUMEN**: fecha, modo, duración, conteos por decisión, top 5 recomendadas.
- Hoja **TODOS** y una hoja por fuente, con formato visual (fila congelada, autofiltro, colores por decisión, links clickeables).
- Columnas de clasificación, categoría, y ficha de postulación lista para usar.

Además, un aviso final (sonido + cartel) con el resumen de la corrida.

## Qué aprendí / desarrollé con este proyecto

- Scraping robusto contra sitios que cambian de estructura seguido (selectores en cascada, detección de bloqueos/CAPTCHA sin romper el flujo).
- Diseño de un sistema de decisión basado en reglas explicables (no una caja negra) — cada clasificación viene con su motivo.
- Manejo de estado persistente entre corridas (historial, deduplicación) sin una base de datos formal.
- Automatización real de un proceso de la vida diaria, con ejecución desatendida vía tarea programada, lock para evitar condiciones de carrera, y distintos modos de operación según el contexto de uso.
- Generación de reportes Excel con formato profesional usando openpyxl directamente (más allá de lo que da pandas `to_excel` por default).

## Limitaciones actuales

- Depende de la estructura HTML de los portales — si Bumeran o Computrabajo cambian su sitio, algunos selectores pueden dejar de funcionar.
- Indeed queda opcional por defecto porque suele bloquear con Cloudflare.
- La clasificación es por reglas simples (keywords), no por IA generativa — es explicable pero no entiende contexto como lo haría un modelo de lenguaje.
- El análisis de perfil usa `perfil_sergio.txt` como resumen de texto — no lee ni parsea un CV real (PDF/Word) automáticamente.
- Algunas páginas pueden bloquear el scraping (CAPTCHA, rate limiting).

## Próximas mejoras

- Integración con CV real (lectura de PDF/Word).
- Generación de cartas de postulación más específicas por oferta.
- Panel web simple para revisar resultados sin abrir Excel.
- Base SQLite en lugar de JSON/Excel para el historial.
- Revisión manual de postulaciones (marcar como "ya postulé").
