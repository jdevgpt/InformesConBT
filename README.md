# Informe BT

Webapp para generar el **INFORME BT - XXXXX** (centros con Baja Tensión) en Word y PDF desde el móvil o el ordenador.

## Archivos

| Archivo | Para qué sirve |
| --- | --- |
| `index.html` | La app |
| `motor.js` | Rellena la plantilla Word y crea el PDF |
| `plantilla.docx` | Plantilla del informe (con los campos `[NOMBRE]`, `[FECHA]`, `[CENTRO]`, `[DIRECCION]`, `[POBLACION]` y la fila de la tabla `[CODIGO_APOYO] [CENTRO] [COD_DEF] [DES_DEF] [OBSERV]`) |
| `direcciones.json` | Base de datos de direcciones de centros (61.021 centros) |

## Subirla a GitHub Pages

1. Crea un repositorio nuevo (por ejemplo `InformeBT`).
2. Sube los 4 archivos de arriba a la raíz del repositorio.
3. En **Settings › Pages**, elige la rama `main` y la carpeta `/ (root)`.
4. La app queda en `https://TU_USUARIO.github.io/InformeBT/`.

## Actualizar la base de direcciones

`direcciones.json` es un objeto `{ "CENTRO": "DIRECCIÓN" }`. Si cambia el Excel, hay que regenerarlo y volver a subirlo.
