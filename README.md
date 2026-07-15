[README.md](https://github.com/user-attachments/files/30067292/README.md)
# Unit Commitment con generación eólica y BESS (IEEE 24 barras)

Este repositorio contiene el código, los datos de entrada y los resultados del proyecto de
titulación **«Modelado de UC multi-período con integración de generación renovable y manejo de
recursos flexibles»**. Está organizado en dos módulos independientes y **portables**:

```
├── LSTM PREDICTION/          → Pronóstico de velocidad de viento con red LSTM bidireccional
│   ├── SCRIPTS/              → Notebook de la LSTM (Python / TensorFlow-Keras)
│   ├── INPUT/
│   │   └── TURBINAS/         → Series mensuales de viento (Turb_15m1.csv … Turb_15m12.csv)
│   └── OUTPUT/
│       ├── METRICS/          → metricas_LSTM_resumen.csv (MAE, RMSE, R² por ventana)
│       ├── PREDICTIONS/      → Prediccion_m*.csv (real vs. predicción por ventana)
│       ├── FIGURES/          → curva de pérdida y gráficas de predicción
│       └── WIND_FORECAST/    → velocidad_viento.xlsx  (entrada de viento para el modelo UC)
│
└── UC MODEL/                 → Modelo de Unit Commitment (Julia / JuMP + HiGHS)
    ├── SCRIPTS/
    │   ├── run_ubicacion.ipynb   → Barrido de 576 ubicaciones BESS/eólico (todas equivalentes; se adopta el nodo 1)
    │   ├── run_casos.ipynb       → 4 escenarios (Básico, Eólico, BESS, Completo)
    │   └── notebooks/            → Los 5 notebooks por escenario (documentación)
    ├── INPUT/
    │   ├── BESS/                         → datos_bess.xlsx
    │   ├── Wind power generation forecast/ → datos_eolicos.xlsx, velocidad_viento.xlsx
    │   └── 24 NODE MODEL/                → datos_generadores, demanda_sistema,
    │                                       lineas_transmision, distribucion_de_carga
    └── OUTPUT/
        ├── LOCATION/    → OUT_sweep.csv (576 costos) y top10.csv
        └── CASES_DATA/  → CSV por escenario (despacho, flujos, flujo máximo, objetivo, dimensión)
```

## Flujo de trabajo (cómo se conectan los dos módulos)

1. **LSTM PREDICTION** entrena la red con las series de las turbinas (`INPUT/TURBINAS`) y, al
   ejecutarse, genera el pronóstico de 24 horas y el archivo **`velocidad_viento.xlsx`**
   (columna de viento = predicción de la ventana `m1…m6 → m7`). Ese archivo es la entrada de viento
   del modelo UC.
2. **UC MODEL** usa `velocidad_viento.xlsx` (junto con los demás datos) para resolver el
   despacho, encontrar la ubicación óptima del BESS y del parque eólico, y comparar los 4 escenarios.

> **Nota sobre `velocidad_viento.xlsx` (acople entre módulos).** Este archivo es la **salida** del
> módulo LSTM (la predicción de 24 h de la ventana `m1…m6 → m7`) y, a la vez, la **entrada de viento**
> del modelo UC. Se incluye una copia ya generada en `UC MODEL/INPUT/Wind power generation forecast/`
> por dos razones: (1) los dos módulos corren en entornos distintos —Python para la LSTM y Julia para
> el UC—, por lo que entregar el perfil ya materializado permite ejecutar el UC de forma autónoma sin
> re-entrenar la red; y (2) el entrenamiento de la LSTM es estocástico y puede variar ligeramente entre
> corridas, de modo que se **fija un único perfil de viento pronosticado como escenario de entrada del
> UC**, garantizando que los resultados del despacho reportados sean reproducibles. Si se desea usar un
> nuevo pronóstico, basta con re-entrenar la LSTM y copiar el `OUTPUT/WIND_FORECAST/velocidad_viento.xlsx`
> resultante a esa carpeta.

## Cómo ejecutar

### 1) Predicción LSTM  (Python 3.11)
Requisitos: `tensorflow`, `scikit-learn`, `pandas`, `numpy`, `matplotlib`, `openpyxl`.
Abrir y ejecutar el notebook:
```
LSTM PREDICTION/SCRIPTS/LSTM_NEURONAL_FORECASTING.ipynb
```
Genera métricas, predicciones, figuras y `velocidad_viento.xlsx` en `LSTM PREDICTION/OUTPUT/`.

### 2) Modelo UC  (Julia 1.10)
Requisitos: `JuMP`, `HiGHS`, `XLSX`, `CSV`, `DataFrames`.
Ejecutar los notebooks de `UC MODEL/SCRIPTS/` en este orden:
```
run_ubicacion.ipynb   # ~15 min: barrido de 576 combinaciones → OUTPUT/LOCATION
run_casos.ipynb       # 4 escenarios en el nodo de referencia → OUTPUT/CASES_DATA
```
`run_casos.ipynb` requiere que `run_ubicacion.ipynb` se haya ejecutado antes
(lee `OUTPUT/LOCATION/OUT_sweep.csv`).

## Reproducibilidad

| Componente | Versión / valor |
|---|---|
| Python | 3.11.9 |
| TensorFlow / Keras | 2.19 (Keras incluido: `tensorflow.keras`) |
| Scikit-learn | 1.2.2 |
| Julia | 1.10.2 |
| JuMP | 1.26.0 |
| HiGHS (solver) | 1.11.0 |
| HiGHS.jl (interfaz Julia) | 1.18.1 |
| XLSX.jl / CSV.jl / DataFrames.jl | 0.10.4 / 0.10.15 / 1.7.0 |
| Tolerancia de optimalidad (MIP gap) | Por defecto de HiGHS (`mip_rel_gap = 1e-4`) |
| Semilla aleatoria (seed) | 42 (`np.random.seed(42)`, `tf.random.set_seed(42)`) |

## Resultado principal
El barrido de las **576 combinaciones** de ubicación (`OUTPUT/LOCATION/OUT_sweep.csv`) muestra que
todas registran el **mismo costo total**: la red del sistema de prueba no presenta congestión a los
niveles de inyección considerados (30 MW), por lo que la ubicación del BESS y del parque eólico
resulta económicamente indiferente y se adopta el **nodo 1 como referencia**. El escenario combinado
(eólica + BESS) obtiene el menor costo de operación — **429,870.38 USD frente a 434,096.02 USD del
caso base — con un ahorro de 0.97 % (4,225.64 USD)**.
