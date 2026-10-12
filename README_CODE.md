# Categorización de farmacias — código

Repositorio de fuentes Python y React para calcular categorías y visualizar
resultados en una aplicación local. `master` contiene la base revisada;
`develop` parte de ella para los siguientes cambios.

## Contenido

- `pharmacy/scripts/`: procesamiento, geocodificación, cálculo y exportación.
- `pharmacy/tests/`: pruebas del modelo.
- `pharmacy/pyproject.toml` y `uv.lock`: dependencias Python.
- `pharmacy/categorization_config.json`: configuración del modelo histórico.
- `visualizacion/server.py`: servidor de uso local.
- `visualizacion/scripts/`: preparación de datos y compilación.
- `visualizacion/app/`: fuentes React, dependencias y documentación técnica.
- `scripts/audit_publication.py`: auditoría del contenido preparado en Git.

## Datos privados y funcionamiento

Este repositorio no incluye Excel, CRM, IQVIA, datos derivados, cartografía,
informes, capturas, memoria interna ni compilados. También se excluyen
generadores de informes y pruebas que contienen resultados comerciales reales
escritos en el código. Todos esos ficheros permanecen en el equipo original.

La copia de GitHub requiere aportar localmente las fuentes autorizadas y los
artefactos necesarios antes de ejecutar el procesamiento o abrir la aplicación.
El repositorio por sí solo no proporciona una aplicación con datos. No sustituir
el snapshot privado `visualizacion/app/src/data.json` por datos vacíos.

Python requiere la versión indicada en `pharmacy/pyproject.toml`. El frontend
utiliza las dependencias de `visualizacion/app/package-lock.json`; respetar
`visualizacion/app/AGENTS.md` al modificarlo. El script de compilación requiere
el compilador local del plugin Data Analytics. El arranque existente utiliza
el entorno Python y los artefactos compilados locales, sin descargar datos ni
dependencias y con el servidor ligado a loopback.

## Antes de publicar cambios

El `.gitignore` permite solo las carpetas de fuentes revisadas y excluye formatos
de datos y credenciales. Las exclusiones no protegen secretos escritos dentro
de código ni archivos añadidos con `git add -f`; revisar siempre el diff.

```powershell
git add .
python scripts/audit_publication.py
git diff --cached --stat
git diff --cached
```

La auditoría no sustituye la revisión manual. No publicar datos de IQVIA o CRM.
