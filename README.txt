IST · Filtros en cascada desde Excel
====================================

Estructura:
- src/
  - template/
    - portada.html
    - inicial.html
  - static/
    - css/
      - portada.css
      - inicial.css
    - img/
      - logo_ist.png   (coloca aquí tu logo)
  - js/
    - portada.js
    - inicial.js

Uso:
1) Abre `src/template/portada.html` y navega a **Cargar Excel**.
2) En `inicial.html`, selecciona tu archivo Excel (.xlsx).
   - Se leerá la hoja llamada “inicial”/“inicio” (ignorando mayúsculas) o, si no existe, la primera hoja.
   - Desde la **fila 3**, se toman las columnas:
       B = Área
       C = Puesto de trabajo
       D = Tareas del puesto de trabajo
   - Filas con datos vacíos o incompletos se descartan.
3) Usa los filtros en cascada para explorar.

Tecnologías:
- Bootstrap 5 + Bootstrap Icons.
- SheetJS (xlsx) para leer Excel en el navegador.

Colores y estilo: paleta corporativa IST (púrpuras y fucsia) aplicada a botones, barras y tablas.


Carga automática por defecto
----------------------------
- La página `inicial.html` intentará cargar automáticamente:
  src/source/4. LO MIRANDA MATRIZ VOTME 2024 PLANTA CERDOS.XLSX
- Si el navegador bloquea la lectura por abrir archivos con `file://`, usa un servidor local (por ejemplo, VS Code Live Server) o haz clic en **Desde ruta fija**.
- Siempre puedes cargar un Excel manualmente con el selector de archivo.
