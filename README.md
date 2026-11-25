# Cisco Price Manager

Aplicación desarrollada en Python y Streamlit para automatizar el análisis de cotizaciones de Cisco, calcular costos de equipos y generar una calculadora financiera en Excel.

## Descripción

Cisco Price Manager permite cargar cotizaciones exportadas desde Cisco Commerce, identificar automáticamente los equipos principales y sus componentes asociados, procesar sus costos y generar un archivo Excel con los cálculos necesarios para evaluar diferentes escenarios comerciales.

El objetivo es reducir el trabajo manual, agilizar el análisis de cotizaciones y facilitar la evaluación financiera de proyectos Cisco.

## Funcionalidades principales

- **Carga de cotizaciones:** lectura de archivos Excel exportados desde Cisco Commerce.
- **Reconocimiento automático de equipos:** identificación de equipos principales sin depender de una lista fija de referencias o modelos.
- **Clasificación de componentes:** agrupación de componentes tangibles e intangibles asociados a cada equipo.
- **Cálculo de Equipment Cost:** consolidación de los costos de los equipos y sus componentes.
- **Cálculo de Aduanas:** aplicación de tarifas según el país seleccionado, considerando únicamente el costo físico del hardware.
- **Selección de plazo:** análisis de escenarios de 12, 24, 36 y 48 meses.
- **Parámetros financieros editables:** NRC, Shipping & Handling, Margin, Return Rate y Markup.
- **Cálculo de MRC:** cálculo de cuotas mensuales mediante fórmulas financieras de Excel.
- **Generación de Excel:** creación automática de un archivo con una única hoja llamada `Calculadora`.
- **Revisión de productos ambiguos:** posibilidad de clasificar manualmente las líneas que no puedan identificarse con suficiente confianza.

## Tecnologías utilizadas

- Python
- Streamlit
- Pandas
- OpenPyXL
- Excel

## Estructura del repositorio

```text
cisco-price-manager/
│
├── Cisco_Price_Manager_app.py
├── requirements.txt
└── README.md
```

## Instalación y ejecución local

### 1. Clonar el repositorio

```bash
git clone <URL_DEL_REPOSITORIO>
cd cisco-price-manager
```

### 2. Crear un entorno virtual (opcional)

```bash
python -m venv venv
```

Activar en Windows:

```bash
venv\Scripts\activate
```

### 3. Instalar dependencias

```bash
pip install -r requirements.txt
```

### 4. Ejecutar la aplicación

```bash
streamlit run Cisco_Price_Manager_app.py
```

La aplicación estará disponible normalmente en:

```text
http://localhost:8501
```

## Dependencias

El archivo `requirements.txt` contiene las librerías necesarias para ejecutar la aplicación.

Ejemplo:

```text
streamlit
pandas
openpyxl
```

## Flujo de funcionamiento

1. El usuario carga una cotización de Cisco Commerce.
2. La aplicación analiza las líneas de productos.
3. Identifica los equipos principales.
4. Agrupa los componentes tangibles e intangibles asociados a cada equipo.
5. Presenta los resultados para su revisión.
6. El usuario configura los parámetros financieros y comerciales.
7. La aplicación genera el archivo Excel con los cálculos correspondientes.

## Consideraciones financieras

- El **Equipment Cost** contempla el equipo principal y sus componentes asociados.
- Los costos de aduanas se calculan sobre el costo físico del hardware.
- Los componentes con duración de servicio se ajustan al plazo de análisis seleccionado.
- El MRC se calcula utilizando la función financiera `PMT` de Excel.
- El Return Rate anual se convierte a una tasa mensual dividiéndolo entre 12.
- El plazo seleccionado determina el número de cuotas mensuales.

## Estado del proyecto

**Versión actual: Beta**

La aplicación se encuentra en fase de validación y revisión funcional. Se esperan comentarios de usuarios y posibles mejoras antes de publicar una versión estable.

## Próximas mejoras

- Incorporar nuevas validaciones para distintos formatos de cotizaciones Cisco.
- Mejorar continuamente el reconocimiento automático de equipos y componentes.
- Optimizar la experiencia de usuario.
- Incorporar nuevas funcionalidades de análisis financiero y comercial.
- Evaluar futuras integraciones con fuentes de información de Cisco.

## Autoría

Proyecto desarrollado para automatizar y optimizar el análisis de cotizaciones y costos comerciales de soluciones Cisco.

---

**Cisco Price Manager — Simplificando el análisis de cotizaciones Cisco.**
