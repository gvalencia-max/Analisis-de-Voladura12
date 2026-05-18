# ⛏️ SPC Blasting Analysis — Aplicación Web Streamlit

**Análisis Estadístico de Resultados de Voladura y Control de Calidad mediante SPC**

> Universidad Nacional del Altiplano · Facultad de Ingeniería de Minas  
> Estudiante: **Giovany Valencia Huanca** · Puno, Perú — 2026

---

## 📋 Descripción

Aplicación web interactiva que integra cuatro herramientas de ingeniería de voladura:

| Módulo | Modelo | Referencia |
|--------|--------|-----------|
| Fragmentación | **Kuz-Ram** (Kuznetsov–Cunningham) | Cunningham, 1983/1987 |
| Vibraciones | **Sadovsky PPV** (distancia escalada) | Sadovsky, 1959 |
| Control de proceso | **Cartas X̄-R** + Reglas de Nelson | Shewhart, 1931 / Montgomery, 2013 |
| Riesgo probabilístico | **Monte Carlo** (10 000 simulaciones) | ISO 22514-2 |

---

## 🚀 Instalación y ejecución rápida

### Requisitos previos
- **Python 3.9 o superior** ([descargar](https://www.python.org/downloads/))
- Conexión a internet (para instalar librerías la primera vez)

### Paso 1 — Descargar los archivos

Descarga los tres archivos en una misma carpeta:

```
mi_proyecto/
├── app.py
├── requirements.txt
└── README.md
```

### Paso 2 — Instalar dependencias

Abre una terminal (CMD, PowerShell o Terminal) dentro de la carpeta y ejecuta:

```bash
pip install -r requirements.txt
```

> 💡 **Tip:** Si tienes varias versiones de Python, usa `pip3` en lugar de `pip`.

### Paso 3 — Ejecutar la aplicación

**Windows:**
```bash
py -m streamlit run app.py
```

**Linux / macOS:**
```bash
streamlit run app.py
```

La aplicación abrirá automáticamente en tu navegador en:
```
http://localhost:8501
```

---

## 🖥️ Estructura de la aplicación

```
Panel lateral (Sidebar)          Área principal
─────────────────────────        ──────────────────────────────────────────
Sección 1: Kuz-Ram               Diagrama técnico de taladros (portada)
  · Factor de roca A              ↓
  · Powder factor K              KPIs (6 indicadores clave)
  · Carga Q, RWS, d              ↓
  · B, S, H, W, BCL, CCL        Tab 1 │ Kuz-Ram + Curva Rosin-Rammler
                                  Tab 1 │ Sadovsky PPV + Curva de atenuación
Sección 2: Sadovsky              Tab 2 │ Cartas de control X̄-R
  · K_s, α, R, Qd, PPV_máx      Tab 3 │ Capacidad Cp, Cpk + distribución
                                  Tab 4 │ Monte Carlo histogramas + tornado
Sección 3: SPC                   Tab 5 │ Tablas de datos y estadísticas
  · X50 nominal, ±tol            ↓
  · Tamaño subgrupo n            Botón: Descargar Excel (.xlsx)
  · N° simulaciones MC

Datos históricos X50
  · Entrada libre (texto)
  · ≥ 3 subgrupos requeridos
```

---

## 📊 Figuras generadas

| # | Figura | Descripción |
|---|--------|-------------|
| 0 | Diagrama taladros | Vista frontal + planta + tabla parámetros + semáforo |
| 1 | Curva Kuz-Ram | Rosin-Rammler con X₅₀ marcado |
| 2 | Atenuación PPV | Curva Sadovsky en log-log con punto de diseño |
| 3 | Cartas X̄-R | Con puntos fuera de control marcados en rojo |
| 4 | Capacidad | Distribución normal del proceso vs. especificaciones |
| 5 | MC X50 | Histograma + FDA con percentiles P5/P50/P95 |
| 6 | MC PPV | Histograma + FDA con límite permisible |
| 7 | Tornado X50 | Correlaciones de Pearson — sensibilidad a inputs |
| 8 | Tornado PPV | Correlaciones de Pearson — sensibilidad a inputs |
| 9 | Dispersión | Scatter plots: A, K, Q, B vs. X50 |

---

## 📥 Reporte Excel exportado

El archivo `SPC_Voladura_Giovany_Valencia.xlsx` contiene **4 hojas**:

| Hoja | Contenido |
|------|-----------|
| `1_Datos_Entrada` | Todos los parámetros ingresados con valores típicos de referencia |
| `2_Resultados` | Interpretación técnica con semáforo de colores |
| `3_SPC_Calculos` | Tabla de subgrupos + límites de control UCL/CL/LCL |
| `4_MonteCarlo` | Estadísticas de X50, PPV y correlaciones de Pearson |

---

## 🔧 Parámetros de entrada — Valores típicos

### Kuz-Ram

| Parámetro | Símbolo | Unidad | Rango típico |
|-----------|---------|--------|--------------|
| Factor de roca | A | — | 0.6 (blanda) – 22 (muy dura) |
| Powder factor | K | kg/m³ | 0.20 – 0.80 |
| Carga por barreno | Q | kg | 50 – 500 |
| RWS (vs ANFO=100) | RWS | — | 75 – 130 |
| Diámetro barreno | d | mm | 76 – 311 |
| Burden | B | m | 2 – 8 |
| Espaciamiento | S | m | 2 – 10 |
| Altura banco | H | m | 5 – 20 |
| Desv. perforación | W | m | 0.01 – 0.06 |
| Carga de fondo | BCL | m | 1 – 5 |
| Carga columna | CCL | m | 3 – 15 |

### Sadovsky PPV

| Parámetro | Símbolo | Unidad | Rango típico |
|-----------|---------|--------|--------------|
| Coef. de sitio | K_s | — | 100 – 1200 |
| Exponente atenuación | α | — | 1.0 – 2.0 |
| Distancia | R | m | 50 – 2000 |
| Carga por retardo | Qd | kg | 20 – 300 |
| PPV máximo | PPV_máx | mm/s | NTP Perú: 15 / USBM: 19–50.8 |

---

## 📐 Fórmulas implementadas

**X50 — Kuznetsov modificado por Cunningham (1987):**
```
X50 = A · K^(-0.8) · Q^(1/6) · (115/RWS)^(19/20)    [resultado en cm × 10 = mm]
```

**Índice de uniformidad n — Cunningham (1987):**
```
n = (2.2 - 14·B/d) · √((1+S/B)/2) · (1-W/B) · (L/H)^0.3 · (BCL/L + CCL/L)
    [d en mm, B/S/H/W/BCL/CCL en metros]
```

**PPV — Sadovsky (1959):**
```
SD  = R / √Qd
PPV = K_s · SD^(-α)
```

**Límites de control SPC (n=5):**
```
UCL_X̄ = X̄̄ + A₂·R̄  (A₂ = 0.577)
LCL_X̄ = X̄̄ - A₂·R̄
UCL_R  = D₄·R̄        (D₄ = 2.114)
σ̂      = R̄ / d₂      (d₂ = 2.326)
```

**Índices de capacidad:**
```
Cp  = (USL - LSL) / (6·σ̂)
Cpu = (USL - X̄̄) / (3·σ̂)
Cpl = (X̄̄ - LSL) / (3·σ̂)
Cpk = min(Cpu, Cpl)
```

---

## 🐛 Solución de problemas comunes

| Problema | Solución |
|----------|----------|
| `ModuleNotFoundError: streamlit` | Ejecutar `pip install -r requirements.txt` |
| Navegador no abre automáticamente | Ir manualmente a `http://localhost:8501` |
| Error en datos históricos | Verificar que cada fila tenga exactamente **n** valores separados por comas |
| `pip` no reconocido en Windows | Usar `py -m pip install -r requirements.txt` |
| Puerto 8501 ocupado | Añadir `--server.port 8502` al comando de streamlit |

---

## 📚 Referencias bibliográficas

1. Kuznetsov, V.M. (1973). *Soviet Mining Science*, 9(2): 144–148.
2. Cunningham, C.V.B. (1983). *1st Int. Symp. Rock Fragmentation by Blasting*, Luleå, pp. 439–454.
3. Cunningham, C.V.B. (1987). *2nd Int. Symp. Rock Fragmentation by Blasting*, Keystone, pp. 475–487.
4. Sadovsky, M.A. (1959). *Mechanical action of air blast*. Academy of Sciences, Moscow.
5. Shewhart, W.A. (1931). *Economic Control of Quality of Manufactured Product*. Van Nostrand.
6. Montgomery, D.C. (2013). *Introduction to Statistical Quality Control*, 7th Ed. Wiley.
7. ISO 22514-2:2017. *Statistical methods in process management — Capability and performance*.
8. Nelson, L.S. (1984). *Journal of Quality Technology*, 16(4): 237–239.

---

## 📄 Licencia

Proyecto académico — Universidad Nacional del Altiplano, Puno, Perú.  
Uso libre para fines educativos con atribución al autor.

---

*Giovany Valencia Huanca · Ingeniería de Minas · UNA Puno · 2026*
