# Memoria Explicativa DGA - Riego

Aplicación Streamlit para completar automáticamente una memoria explicativa DGA para solicitudes de derechos de aprovechamiento de aguas subterráneas destinadas a riego, con porcentaje opcional para uso doméstico de subsistencia.

## Funciones

- Exporta Word (.docx).
- Exporta PDF (.pdf).
- Autocompleta datos desde un informe técnico BLA en Word (.docx).
- Permite cargar firma digital PNG.
- Permite cargar croquis/mapa opcional.
- Guarda y carga ficha JSON.
- Considera por defecto:
  - aguas subterráneas;
  - derecho consuntivo;
  - ejercicio permanente y continuo;
  - sin derechos constituidos asociados;
  - sin derechos en trámite asociados;
  - uso riego y uso doméstico de subsistencia.

## Ejecutar localmente

```bash
pip install -r requirements.txt
streamlit run app/main.py
```


## Versión 1.1 - autocompletado corregido

- Corrige extractor de informes Word BLA.
- Ya no toma frases narrativas como nombre del solicitante.
- Prioriza campos estructurados: Solicitante, RUT, Domicilio, Comuna, Provincia, Teléfono, Correo, Coordenadas, Datum, Huso y Caudal solicitado.
- Corrige firmante para que use el nombre y RUT reales detectados.


## Versión 1.2 - autocompletar predio desde avalúo fiscal y Conservador

- Agrega cargador para Avalúo Fiscal SII en PDF, JPG o PNG.
- Agrega cargador para inscripción del Conservador de Bienes Raíces en PDF, JPG o PNG.
- Intenta detectar Rol SII, comuna, predio, superficie, fojas, número, año y conservador.
- Agrega OCR mediante Tesseract para documentos escaneados.
- Actualiza Dockerfile para instalar tesseract-ocr y tesseract-ocr-spa.


## Versión 1.3 - lectura de predio mejorada

- Mejora extracción de fojas, número, año y conservador desde inscripciones CBR.
- Agrega campos opcionales para pegar texto del avalúo fiscal o inscripción CBR si el OCR falla.
- Agrega visualización del texto detectado para diagnosticar documentos escaneados.
- Mantiene soporte para PDF, JPG y PNG.


## Versión 1.4 - predio autocompletado corregido

- Corrige el mapeo interno de campos del avalúo y Conservador.
- Ahora los datos detectados se aplican correctamente a:
  - rol_sii
  - fojas
  - numero_inscripcion
  - anio_inscripcion
  - conservador
- Fuerza actualización visual inmediata después de autocompletar datos del predio.


## Versión 1.5 - coordenadas UTM corregidas

- Mejora detección de coordenadas desde informes BLA.
- Soporta formatos como:
  - Huso 18, coordenadas 753.811 m E, 5.821.082 m S
  - Este 753.811 m y Norte 5.821.082 m
  - UTM Este / UTM Norte
- Aplica automáticamente los valores a UTM Norte, UTM Este y Datum/Huso.


## Versión 2.0 - formato oficial DGA exacto

- El PDF se genera usando como fondo el formulario oficial DGA cargado en assets.
- Se conservan tamaño de página, tablas, líneas, cuadros, tipografías y distribución del formulario original.
- La app solo estampa los datos sobre los espacios del formulario.
- Se exportan las páginas 1 a 10 del formulario, excluyendo las instrucciones.
- El Word se genera como páginas-imagen a partir del PDF completado para conservar apariencia idéntica.
- Nota: el Word prioriza fidelidad visual; su contenido no queda editable como texto.


## Versión 2.1 - formulario oficial calibrado

- Corrige coordenadas de escritura sobre la plantilla oficial DGA.
- Ajusta marcas X para que calcen dentro de los casilleros.
- Completa correctamente identificación del peticionario.
- Completa naturaleza, tipo/ejercicio, caudal, volumen anual y captación.
- Agrega descripción complementaria referencial de ubicación.
- Completa 3.1 con resumen del proyecto desde el informe BLA.
- Marca NO en derechos constituidos y NO en derechos en trámite.
- Marca correctamente riego y uso doméstico de subsistencia.
- Exporta solo páginas aplicables: 1 a 5 y página de información adicional/firma.
- En 4.2 y 4.3 muestra la región por nombre, no por número.
- Genera información adicional relacionada con el proyecto si el campo está vacío.


## Versión 2.2 - peticionario Irrisal

- El punto 1 Identificación del peticionario se completa siempre con:
  - Irrisal Consulting Ltda.
  - San Martín 553 oficina 901
  - RUT 78.271.963-7
  - +56 9 6796 0884
  - Irrisalconsulting@gmail.com
- El autocompletado desde informe BLA ya no reemplaza esos datos con el beneficiario.
- Los datos del beneficiario se conservan solo como antecedente del proyecto/predio.


## Versión 2.2.1 - corrección peticionario Irrisal

- Corrige error NameError: PETICIONARIO_EMPRESA is not defined.
- Define los datos de Irrisal Consulting Ltda. antes de DEFAULTS.
- Mantiene el punto 1 del formulario con los datos de la empresa.


## Versión 2.3 - corrección integral formulario DGA

- Recalibra las marcas X de 2.1, 2.2, 2.3, 3.2 y 3.3.
- Corrige ubicación de textos en 2.4, 3.1, 4.2, 4.3 y 5.
- El punto 1 usa siempre los datos de Irrisal Consulting Ltda. como peticionario.
- Región se exporta como nombre (Biobío) y no como número ni frase larga.
- Limpia el campo Conservador para eliminar frases de certificación del documento CBR.
- La descripción de ubicación incluye coordenadas, comuna, predio y tipo de captación si existen.
- Información adicional se genera vinculada al proyecto, caudal, superficie, riego, subsistencia y antecedentes adjuntos.


## Versión 2.4 - ajuste fino formulario DGA

- Corrige posición de UTM Norte, UTM Este y Datum/Huso en 2.4.
- Corrige posición de volumen anual en 2.3.
- Ajusta 4.2 y 4.3 para que región/provincia/comuna, predio, rol, hectáreas, fojas, número, año y conservador caigan en sus celdas.
- Si falta caudal o volumen, no exporta 0,00 como dato válido; deja el espacio en blanco.
- Mejora autocompletado de caudal y coordenadas desde informe BLA.


## Versión 2.5 - firma calibrada

- Reubica la firma PNG dentro del espacio oficial de firma, centrada sobre la línea.
- Elimina el texto adicional bajo la línea de firma para evitar que aparezca fuera del espacio asignado.
- Ajusta el bloque de información adicional para que no invada la zona de firma.
- Mantiene la exportación con plantilla oficial DGA.


## Versión 2.6 - UTM desde informe reforzado

- El autocompletado del informe técnico ahora acepta Word, PDF, JPG y PNG.
- Agrega campo para pegar texto del informe si el archivo no se lee bien.
- Refuerza la detección de UTM Norte y UTM Este desde frases como:
  - coordenadas 753.811 m E, 5.821.082 m S
  - Este 753.811 m y Norte 5.821.082 m
  - UTM Este / UTM Norte
- Normaliza coordenadas ingresadas con puntos de miles antes de exportar.


## Versión 2.7 - UTM por casillas y secciones 4.2/4.3 calibradas

- En el punto 2.4, UTM Norte y UTM Este se escriben con un dígito por casilla.
- Redibuja las secciones 4.2 y 4.3 para evitar descalces al omitir usos no aplicables.
- Corrige ubicación de Región, Provincia, Comuna, Predio, Rol SII, Hectáreas, Fojas, Número, Año y Conservador.
- Mantiene el formulario enfocado en uso doméstico de subsistencia y riego.


## Versión 2.8 - punto 4 limpio, UTM calibradas e información adicional alineada

- Ajusta la escritura de UTM Norte y UTM Este para que cada dígito caiga dentro de su casilla.
- Mueve Datum/Huso al espacio correcto del punto de captación.
- Elimina visualmente el bloque 4.1 Agua Potable en la exportación, dejando solo 4.2 y 4.3.
- Agrega títulos visibles para 4.2 y 4.3.
- Redibuja tablas 4.2 y 4.3 con posiciones calibradas.
- Alinea el texto del punto 5 con las líneas del formulario oficial.
- Mejora posición de textos en el punto 1 Identificación del peticionario.
- Evita frases con caudal o hectáreas en cero cuando el dato no está disponible.
