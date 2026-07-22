# Unit Commitment con generación eólica y BESS en el sistema IEEE de 24 barras

Este repositorio contiene el código, los datos de entrada y los resultados del trabajo de titulación **«Modelado de UC multi-período con integración de generación renovable y manejo de recursos flexibles»**.

El proyecto integra dos módulos:

1. Predicción exploratoria de velocidad del viento mediante una red neuronal BiLSTM desarrollada en Python.
2. Modelo determinista simplificado de *Unit Commitment* (UC), implementado en Julia/JuMP y resuelto con HiGHS para el sistema IEEE RTS de 24 barras.

El modelo UC evalúa cuatro escenarios durante un horizonte de 24 horas:

- Tradicional: generación térmica y red de transmisión.
- Tradicional + Eólico: incorpora generación eólica pronosticada.
- Tradicional + BESS: incorpora almacenamiento de energía.
- Tradicional + Eólico + BESS: operación conjunta de ambos recursos.

## Estructura del repositorio

```text
.
├── LSTM PREDICTION/
│   ├── INPUT/
│   │   └── TURBINAS/
│   │       ├── Turb_15m1.csv
│   │       ├── ...
│   │       └── Turb_15m12.csv
│   ├── SCRIPTS/
│   │   └── LSTM_NEURONAL_FORECASTING.ipynb
│   └── OUTPUT/
│       ├── FIGURES/
│       ├── METRICS/
│       │   └── metricas_LSTM_resumen.csv
│       ├── PREDICTIONS/
│       │   └── Prediccion_m*.csv
│       └── WIND_FORECAST/
│           └── velocidad_viento.xlsx
│
├── UC MODEL/
│   ├── INPUT/
│   │   ├── 24 NODE MODEL/
│   │   │   ├── datos_generadores.xlsx
│   │   │   ├── demanda_sistema.xlsx
│   │   │   ├── distribucion_de_carga.xlsx
│   │   │   └── lineas_transmision.xlsx
│   │   ├── BESS/
│   │   │   └── datos_bess.xlsx
│   │   └── Wind power generation forecast/
│   │       ├── datos_eolicos.xlsx
│   │       └── velocidad_viento.xlsx
│   ├── SCRIPTS/
│   │   ├── run_ubicacion.ipynb
│   │   ├── run_casos.ipynb
│   │   └── notebooks/
│   │       ├── UnitCommitment_CompleteBasic.ipynb
│   │       ├── UnitCommitment_CompleteEolico.ipynb
│   │       ├── UnitCommitment_CompleteBESS.ipynb
│   │       ├── UnitCommitment_CompleteFinal.ipynb
│   │       └── UnitCommitment_BestNodeEolBess.ipynb
│   └── OUTPUT/
│       ├── LOCATION/
│       │   ├── OUT_sweep.csv
│       │   └── top10.csv
│       └── CASES_DATA/
│           ├── Basic_*.csv
│           ├── Eolico_*.csv
│           ├── BESS_*.csv
│           ├── Final_*.csv
│           └── resumen_casos.csv
│
├── requirements.txt
├── Project.toml
├── Manifest.toml
└── README.md
```

Los directorios mostrados corresponden a los archivos necesarios para ejecutar, documentar y reproducir los módulos del repositorio. Las gráficas y tablas utilizadas en el documento de tesis se generan a partir de los archivos de salida y no forman parte de la estructura requerida para ejecutar los modelos.

## 1. Módulo de predicción BiLSTM

El notebook `LSTM PREDICTION/SCRIPTS/LSTM_NEURONAL_FORECASTING.ipynb` procesa doce series mensuales de velocidad del viento y construye ventanas cronológicas de seis meses de entrada y 24 horas de salida.

Las cinco primeras ventanas se utilizan para el ajuste del modelo. La ventana:

```text
m6...m11 → m12
```

corresponde a la única evaluación estrictamente fuera de muestra. Esta ventana también genera el perfil de 24 horas utilizado como entrada del modelo UC.

### Resultados de la ventana externa

| Métrica | Valor |
|---|---:|
| MAE | 5.8589 m/s |
| RMSE | 7.7604 m/s |
| R² | -14.8505 |

Debido al desempeño limitado de la ventana externa y al reducido número de ventanas disponibles, la BiLSTM se utiliza como una herramienta exploratoria para construir un perfil de entrada. Los resultados no se presentan como evidencia de capacidad general de pronóstico.

### Archivos generados

- `OUTPUT/METRICS/metricas_LSTM_resumen.csv`: métricas por ventana temporal.
- `OUTPUT/PREDICTIONS/Prediccion_m*.csv`: valores observados y predichos.
- `OUTPUT/FIGURES/`: curva de pérdida y evaluación gráfica de las ventanas.
- `OUTPUT/WIND_FORECAST/velocidad_viento.xlsx`: pronóstico externo de 24 horas para el modelo UC.

## 2. Acoplamiento entre la BiLSTM y el modelo UC

El archivo:

```text
LSTM PREDICTION/OUTPUT/WIND_FORECAST/velocidad_viento.xlsx
```

es la salida del módulo BiLSTM. Su columna de viento contiene la predicción inversamente escalada correspondiente a la ventana externa `m6...m11 → m12`.

Para ejecutar el UC de manera independiente se incluye una copia del perfil en:

```text
UC MODEL/INPUT/Wind power generation forecast/velocidad_viento.xlsx
```

Si se genera un nuevo pronóstico, el archivo de salida de la BiLSTM debe copiarse a esa ubicación antes de ejecutar nuevamente los escenarios UC. Esto puede modificar el despacho, los flujos y los costos.

## 3. Modelo de Unit Commitment

El modelo UC considera:

- Doce unidades térmicas.
- Estados iniciales de las unidades leídos desde `datos_generadores.xlsx`.
- Costos variables, fijos y de arranque.
- Balance de potencia por barra.
- Flujo de potencia DC con una base de 100 MVA.
- Límites de transmisión.
- Generación eólica disponible y utilizada.
- Vertimiento eólico penalizado.
- BESS de operación bidireccional con eficiencias de carga y descarga.
- Límites de potencia, energía y variación horaria del BESS.
- Exclusión de carga y descarga simultáneas.
- Condición terminal de energía almacenada.

El modelo es determinista y simplificado. No incluye reservas operativas, rampas térmicas de las unidades, tiempos mínimos de encendido y apagado, contingencias N−1 ni incertidumbre explícita.

### Datos de ubicación

`run_casos.ipynb` utiliza directamente las barras especificadas en los archivos de entrada:

- `datos_eolicos.xlsx`: barra del parque eólico.
- `datos_bess.xlsx`: barra del BESS.

Los archivos incluidos establecen la barra 1 para ambos recursos. Por tanto, `run_casos.ipynb` puede ejecutarse sin ejecutar previamente el barrido de ubicación.

## Instalación

### Entorno Python

Se recomienda utilizar Python 3.11 y un entorno virtual:

```bash
python -m venv .venv
```

En Windows:

```powershell
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Después, iniciar Jupyter Notebook:

```powershell
jupyter notebook
```

### Entorno Julia

Desde la raíz del repositorio, iniciar Julia con el proyecto local:

```powershell
julia --project=.
```

En la consola de Julia:

```julia
using Pkg
Pkg.instantiate()
```

Los archivos `Project.toml` y `Manifest.toml` permiten reconstruir el entorno utilizado.

## Ejecución

### Predicción de viento

Abrir y ejecutar todas las celdas de:

```text
LSTM PREDICTION/SCRIPTS/LSTM_NEURONAL_FORECASTING.ipynb
```

### Barrido opcional de ubicaciones

Ejecutar:

```text
UC MODEL/SCRIPTS/run_ubicacion.ipynb
```

Este notebook evalúa las 576 combinaciones posibles de conexión del BESS y del parque eólico en las 24 barras. Genera:

```text
UC MODEL/OUTPUT/LOCATION/OUT_sweep.csv
UC MODEL/OUTPUT/LOCATION/top10.csv
```

El barrido es un análisis independiente y no constituye un requisito previo para `run_casos.ipynb`.

### Ejecución de los cuatro escenarios

Ejecutar:

```text
UC MODEL/SCRIPTS/run_casos.ipynb
```

El notebook resuelve los cuatro escenarios utilizando las barras definidas en los archivos de entrada y almacena los resultados en:

```text
UC MODEL/OUTPUT/CASES_DATA/
```

Para cada escenario se exportan archivos de despacho, flujo máximo por hora, flujos de las líneas, desglose de la función objetivo y dimensión del modelo.

## Resultados actuales

| Escenario | Costo total (USD) | Ahorro frente al tradicional (USD) | Ahorro (%) |
|---|---:|---:|---:|
| Tradicional | 429357.91 | — | — |
| Tradicional + Eólico | 424303.74 | 5054.17 | 1.18 |
| Tradicional + BESS | 429235.88 | 122.03 | 0.03 |
| Tradicional + Eólico + BESS | 423869.43 | 5488.48 | 1.28 |

El escenario conjunto eólico–BESS presentó el menor costo operativo. La totalidad de los 364.6713 MWh de energía eólica disponible fue utilizada sin vertimiento. En los escenarios con almacenamiento, el BESS cargó 63.1579 MWh, descargó 57 MWh y finalizó con un estado de carga del 80 %.

### Barrido de ubicación

Las 576 configuraciones evaluadas presentaron costos entre aproximadamente 423869.4296 y 423869.4302 USD. La diferencia máxima, cercana a 0.0006 USD, es económicamente despreciable y se encuentra al nivel de la tolerancia numérica del solver. Por ello, no se identificó una ubicación económicamente única y las configuraciones se consideran coóptimas con la precisión de reporte adoptada.

La barra 1 se utiliza como configuración de referencia porque es la ubicación indicada en los archivos de entrada, no porque el barrido haya demostrado una ventaja económica exclusiva de ese nodo.

## Reproducibilidad

| Componente | Versión o configuración |
|---|---|
| Python | 3.11.7 |
| NumPy | 1.26.4 |
| Pandas | 2.1.4 |
| Matplotlib | 3.8.0 |
| Scikit-learn | 1.2.2 |
| TensorFlow | 2.19.0 |
| Keras | 3.10.0 |
| Julia | 1.10.2 |
| JuMP | 1.26.0 |
| HiGHS.jl | 1.18.1 |
| CSV.jl | 0.10.15 |
| DataFrames.jl | 1.7.0 |
| XLSX.jl | 0.10.4 |
| Potencia base | 100 MVA |
| Tolerancia relativa MIP | `1e-4` |
| Semilla aleatoria | 42 |

## Consideraciones para interpretar los resultados

- Los costos corresponden a un horizonte determinista de 24 horas y no incluyen costos de inversión.
- El beneficio económico del BESS es operativo y específico para los parámetros evaluados; no constituye un análisis de rentabilidad.
- No se cuantifica el efecto del error de pronóstico sobre el despacho UC.
- Los datos de viento se utilizan con fines académicos y no representan mediciones locales del parque Villonaco.
- La ausencia de un óptimo espacial único corresponde a las capacidades y condiciones de red evaluadas y no debe generalizarse a sistemas congestionados.

## Autor

Jonathan Alberto Añazco Gallardo  
Carrera de Electricidad — Universidad Nacional de Loja
