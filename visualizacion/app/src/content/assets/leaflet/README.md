# Leaflet local · 1.9.4

Fuente: https://unpkg.com/leaflet@1.9.4/dist/leaflet-src.esm.js y CSS/recursos de la misma versión. Licencia BSD de dos cláusulas: `LICENSE`.

Esta copia se sirve localmente; no carga bibliotecas desde una CDN. Cambios de integración:

- Se elimina la detección automática de iconos mediante elementos temporales en `document.body`. Este mapa utiliza puntos Canvas y marcadores vectoriales; los iconos de imagen futuros requieren URL local explícita.
- La selección de texto, arrastre de imágenes y clase de cursor durante el arrastre se aplican al contenedor de mapa, sin alterar el documento global ni otros controles de la aplicación.
- El foco del mapa emplea `preventScroll` y conserva el desplazamiento sin modificar la posición global de la ventana.
- Se eliminan las reglas VML y filtros propietarios `progid` de Internet Explorer antiguo. No son CSS estándar ni se ejecutan en los navegadores modernos de esta aplicación.

Estas modificaciones delimitan el comportamiento real del mapa dentro de la aplicación. No modifican la infraestructura ni los controles de integridad del runtime.
