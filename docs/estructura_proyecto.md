# Estructura del proyecto

```
Agente-Buscador-Laboral/
├─ buscador_trabajos_v2.py
├─ README.md
├─ requirements.txt
├─ perfil_sergio.txt
├─ run_buscador_trabajos.bat
├─ .gitignore
├─ docs/
│  ├─ guion_presentacion.md
│  └─ estructura_proyecto.md
├─ resultados/
│  └─ YYYY-MM/
└─ logs / archivos locales (vistos.json, historial_trabajos.xlsx, log_buscador.txt, buscador.lock)
```

## Archivos presentables (portfolio / repositorio público)

Estos son seguros para mostrar o subir a un repo:

- `buscador_trabajos_v2.py` — el script principal.
- `README.md`
- `requirements.txt`
- `run_buscador_trabajos.bat`
- `docs/` (esta carpeta, con el guion y esta misma estructura documentada).
- `perfil_sergio.txt` — **revisar antes de subir**: hoy contiene nombre completo y resumen de perfil, sin datos de contacto (mail, teléfono, dirección) ni documentos. Si se agrega ese tipo de dato más adelante, sacarlo antes de compartir el archivo.

## Archivos que conviene NO subir (locales / privados)

Estos son estado de ejecución o datos generados — no aportan al portfolio y algunos pueden llenarse de información personal (empresas a las que se evaluó postular, historial de búsqueda):

- `vistos.json` — historial de links vistos.
- `historial_trabajos.xlsx` — historial acumulado de ofertas.
- `log_buscador.txt` — log de corridas.
- `resultados/` — reportes Excel reales (usar el **MODO DEMO** para generar un ejemplo presentable en su lugar).
- `buscador.lock` — archivo de lock, solo existe mientras el script corre.
- `debug_indeed_cloudflare.png` / `debug_indeed_cloudflare.html` — evidencia de debug de bloqueos, se genera solo si Indeed bloquea.
- `*.pyc` y carpetas `__pycache__/` — artefactos de compilación de Python.
- `.env` — si en algún momento se agregan credenciales o claves.

Todos estos ya están cubiertos por `.gitignore`, así que si el proyecto se sube a un repositorio Git no se van a incluir por accidente.

## Cómo generar una demo presentable

En vez de mostrar `resultados/` real (con datos de búsquedas reales), correr el script con `MODO_EJECUCION = "DEMO"` — genera un Excel de ejemplo completo (con las mismas hojas, formato y clasificación) sin tocar internet ni el historial real. Ese archivo sí es presentable.
