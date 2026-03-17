# Accesorios - módulo de compras

## Qué hace

Genera archivos de Excel para:

- Heat Transfer
- Size Strip
- Color Tag
- Hang Tag

## Cambios incluidos

- `Color Tag` usa el diseño nuevo basado en `color_tag2.xlsx`
- Selector de color por estilo: `Pink` o `Grey`
- La descripción del Excel cambia automáticamente según el color elegido
- Se inserta la imagen correcta del tag según el color seleccionado
- `Color Tag` pide cantidad manual por estilo
- `Heat Transfer`, `Size Strip` y `Hang Tag` mantienen su funcionamiento
- La interfaz ahora usa el estilo visual unificado del proyecto

## Archivos principales

- `index.html`
- `JS/kms.js`
- `plantillas/`
- `assets/`
- `../shared/accessories.css`
- `../shared/theme.css`

## Nota

Para que carguen bien las imágenes del Excel, abre el proyecto desde un servidor local como Live Server o `localhost`, no directo con `file://`.
