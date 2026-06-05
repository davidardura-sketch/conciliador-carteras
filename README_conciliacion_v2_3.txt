# Conciliador de carteras - v2.3

Cambios incluidos:
- La interfaz, el PDF y el Excel utilizan "Efectivo" en lugar de "Efectivo manual".
- El Excel de salida genera un gráfico de tarta con la distribución por tipo de activo.
- La tabla auxiliar del gráfico queda en la hoja `Hoja1`, para preservar el formato de `Worksheet`.
- La app sigue usando el Excel modelo como plantilla obligatoria para mantener el formato requerido.

Ejecución local en Mac:

1. Guarda en la misma carpeta:
   - app_conciliacion_carteras_v2_3.py
   - requirements_conciliacion.txt

2. Abre Terminal y entra en la carpeta, por ejemplo:
   cd Downloads

3. Instala dependencias:
   pip3 install -r requirements_conciliacion.txt

4. Ejecuta:
   python3 -m streamlit run app_conciliacion_carteras_v2_3.py

Inputs de la app:
- Excel depositario.
- Excel Bloomberg.
- Excel plantilla/modelo de salida.
- Efectivo en EUR.

Outputs:
- Excel conciliado en formato de plantilla, con gráfico de tarta por tipo de activo.
- PDF de reporte de conciliación.
