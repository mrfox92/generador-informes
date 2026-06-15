# App Informes de Notas - v7 formato institucional

Generador local de informes de notas en Word y PDF desde planillas Excel flexibles.

## Cambios v7

- Se ajustó el ancho de la tabla de notas para volver al formato institucional original.
- La tabla ya no usa `100%` de ancho automático en Word.
- Se dejaron anchos fijos por columna, basados en `TEMPLATE_INFORME.docx`:
  - Asignatura: 5102 dxa
  - N1 a N6: 748/749 dxa
  - Promedio: 1188 dxa
- Se mantiene el botón `Nueva carga / limpiar`.
- Se mantienen logos oficiales en `assets/`.
- PDF sigue generándose sin LibreOffice.

## Instalación

```bash
npm install
```

## Ejecutar

```bash
npm run dev
```

Luego abrir:

```text
http://localhost:3000
```

## Logos

Los logos están en:

```text
assets/logo-bicentenario.png
assets/logo-liceo.png
```

Para ajustar tamaños o posición:

```text
src/reportGenerator.js
```

Buscar:

```javascript
// LOGOS WORD
// LOGOS PDF
```

## Ajuste principal de formato

En `src/reportGenerator.js` buscar:

```javascript
const W_SUBJECT = 5102;
const W_GRADE = 749;
const W_AVG = 1188;
```

Y también:

```javascript
width: { size: 10783, type: WidthType.DXA },
columnWidths: [5102, 748, 749, 749, 749, 749, 749, 1188],
layout: TableLayoutType.FIXED,
alignment: AlignmentType.CENTER,
```

Esas líneas controlan que la tabla no se expanda de forma innecesaria en Word.

## Nota de seguridad

`npm audit` puede seguir mostrando una vulnerabilidad asociada a `xlsx`. Para uso local dentro del establecimiento no debería ser crítico, pero esta app no debe exponerse a internet sin cambiar la librería de lectura Excel o aplicar validaciones adicionales.
# generador-informes
