# Conciliador de carteras

App Streamlit para conciliar carteras a partir de tres ficheros:

1. Excel del depositario.
2. Excel de Bloomberg.
3. Excel plantilla/modelo de salida.

La app conserva el formato del modelo, permite introducir efectivo en EUR y genera el Excel conciliado con gráfico por tipo de activo y reporte PDF.

## Ejecución local

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Despliegue en Streamlit Cloud

Main file path:

```text
streamlit_app.py
```

No subir Excels reales, PDFs de salida ni datos de carteras al repositorio.
