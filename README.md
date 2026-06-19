# Sistema de Pruebas de Bombeo - Irrisal Consulting Ltda.

Versión 2: informe PDF profesional.

## Mejoras principales

- Portada más profesional.
- Datos institucionales en portada y pie de página.
- Sección de ubicación y habilitación.
- Campos de cribas, tubería ciega, profundidad de bomba y tubería de extracción.
- Estratigrafía editable.
- Equipos y metodología.
- Gráficos insertados en PDF.
- Tabla de gasto constante.
- Tabla de recuperación.
- Conclusiones automáticas más cuidadosas.
- Recomendaciones automáticas.
- Exportación Excel ampliada.

## Ejecutar localmente

```bash
pip install -r requirements.txt
streamlit run app/main.py
```

## Nota técnica

El sistema no rellena datos faltantes ni presenta valores estimados como medidos.


## Versión 2.2 - mejoras de formato

- Logo pequeño en el encabezado de todas las páginas.
- Logo de portada sin deformación.
- Imágenes subidas al informe con proporción conservada.
- Croquis/ubicación sin achatamiento.
- Gráficos insertados sin deformación.
- Márgenes ajustados para encabezado y pie de página.


## Versión 2.3 - exportación Word y PDF

- Se elimina la exportación Excel de la interfaz.
- Se agrega exportación Word editable (.docx).
- Se mantiene exportación PDF formal.
- Ambos informes incluyen datos del proyecto, captación, estratigrafía, equipos, metodología, resultados, gráficos, tablas, advertencias, conclusiones y recomendaciones.
