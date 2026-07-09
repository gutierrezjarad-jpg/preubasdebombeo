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


## Versión 2.4 - formato de exportación

- Cuerpo de informe en Arial 11 e interlineado 1,5 en Word.
- Títulos desde Introducción en Arial 13.
- PDF con tamaño de cuerpo equivalente y mejor ajuste de tablas.
- Tabla de resultados con texto largo ajustable.
- Porcentaje de recuperación con dos decimales.
- Se elimina la sección Advertencias técnicas del Word y PDF exportados.
- Gráfico de prueba constante más grande.


## Versión 2.4.1 - gráfico Word

- Aumenta el tamaño del gráfico de prueba constante solo en la exportación Word.
- Mantiene el PDF sin cambios.
- Ajusta levemente los márgenes laterales del Word para evitar recorte del gráfico.


## Versión 2.4.2 - firma profesional

- Agrega cargador de firma PNG en la barra lateral.
- Word y PDF muestran solo una firma al final.
- Elimina firma del beneficiario/cliente.
- Inserta la firma del profesional si se sube imagen PNG.
- Texto final: David Gutiérrez Jara, Ingeniero Agrónomo.


## Versión 2.5 - guardar y cargar datos

- Agrega guardado de ficha de datos en formato JSON.
- Permite cargar una ficha JSON para autocompletar empresa, proyecto, captación, equipos, metodología, estratigrafía, bombeo y recuperación.
- La ficha sirve para recuperar información después de actualizar la aplicación.
- Por seguridad, la imagen de firma no se guarda dentro de la ficha JSON; debe subirse nuevamente si se requiere insertar en informes.


## Versión 2.5.1 - texto justificado y títulos no huérfanos

- Justifica el texto del cuerpo en Word.
- Justifica el texto del cuerpo en PDF.
- Mantiene los títulos junto al contenido siguiente en Word.
- Mantiene los títulos junto al contenido siguiente en PDF mediante keepWithNext.
- Evita que títulos como Introducción, Resultados, Conclusiones o Recomendaciones queden solos al final de una página.


## Versión 2.5.2 - metodología narrativa y antecedentes fluidos

- Automatiza una introducción más narrativa.
- Agrega sección 1.1 Antecedentes generales.
- Agrega sección 1.2 Metodología de la prueba de bombeo con texto generado según datos reales de la prueba.
- Agrega sección 1.3 Características generales de la captación en texto continuo, sin tabla.
- Mantiene las tablas como síntesis técnica posterior.
- Renombra la antigua tabla de metodología como Parámetros metodológicos registrados para evitar duplicidad.


## Versión 2.5.3 - corrección de campos numéricos

- Corrige error de Streamlit cuando una ficha guardada trae campos numéricos como texto.
- Normaliza profundidad total, nivel estático, cribas y profundidad de bomba antes de crear los widgets.
- Si criba desde/hasta viene vacía o como 'No informado', la app la transforma en 0.0.
- Mantiene la metodología narrativa de la versión 2.5.2.
