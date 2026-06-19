# Sistema de Pruebas de Bombeo - Irrisal Consulting Ltda.

Aplicación mínima funcional en Streamlit para registrar datos de pruebas de bombeo, calcular indicadores técnicos y exportar Excel/PDF.

## Funciones incluidas

- Datos institucionales de Irrisal Consulting Ltda.
- Registro de proyecto, captación y equipos.
- Tabla editable de prueba de gasto constante.
- Tabla editable de recuperación.
- Cálculo de:
  - duración
  - abatimiento
  - caudal promedio
  - caudal específico
  - volumen bombeado
  - pendiente final en cm/h
  - porcentaje de recuperación
- Advertencias técnicas.
- Exportación Excel.
- Exportación PDF con ReportLab.
- Dockerfile compatible con Cloud Run.

## Ejecutar localmente

```bash
pip install -r requirements.txt
streamlit run app/main.py
```

## Ejecutar con Docker

```bash
docker build -t pruebas-bombeo .
docker run -p 8080:8080 pruebas-bombeo
```

## Nota técnica

La aplicación no inventa ni rellena datos faltantes.  
Si la prueba dura menos de 24 horas en pozo profundo, el informe advierte que no corresponde declarar cumplimiento formal de una prueba estándar de 24 h.
