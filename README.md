# Cisco Price Manager — versión inicial

Aplicación local para:
- Importar cotizaciones Cisco en formato `.xlsx`.
- Conservar los campos de la tabla Cisco: Part Number, Smart Account Mandatory, Description, Service Duration, Estimated Lead Time, Unit List Price, Pricing Term, Qty, Unit Net Price, Disc(%), Extended Net Price.
- Ingresar SKU manualmente, uno por línea.
- Editar cantidades, precios de lista y descuentos manualmente por SKU.
- Calcular precio neto y extendido.
- Exportar una nueva cotización Excel.
- Comparar dos cotizaciones y detectar cambios de precio por SKU.

## Ejecutar en Windows

1. Instalar Python 3.10 o superior.
2. Descomprimir esta carpeta.
3. Abrir PowerShell dentro de la carpeta.
4. Ejecutar:

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
streamlit run app.py
```

Se abrirá la aplicación en el navegador, normalmente en `http://localhost:8501`.

## Importante sobre los precios Cisco

Esta versión NO se conecta todavía a Cisco Commerce Workspace. Los precios de lista deben venir del Excel cargado o ser introducidos por el usuario. El descuento de partner es editable manualmente por línea.

La integración real requiere que la organización tenga acceso autorizado a la API de Cisco y que se confirme qué endpoint devuelve el precio de lista aplicable al catálogo/país y qué reglas comerciales se deben usar. No se deben guardar credenciales de Cisco en el archivo Excel ni en el código fuente.

## Recomendaciones para la siguiente fase

- Añadir conexión segura a la API de Cisco autorizada por la organización.
- Programar una sincronización diaria.
- Guardar histórico de precios y fechas de consulta.
- Mantener separado precio de lista, descuento manual, precio neto calculado y precio devuelto por Cisco.
- Añadir validación para SKU sin precio, SKU discontinuados y diferencias de moneda.
