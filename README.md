# AxDashBoard-Set-Controls (v2.1.0)

Graphics Labels suite for creating modern, flat-style Dashboards in **Visual Basic 6** using GDI+ rendering.

---

## 📦 Controles Incluidos

### 1. AxDashSmallLabel
Label compacto con 3 Captions y un IcoFont enmarcado que sobresale del recuadro principal para dar mayor realce visual.
- **Novedades v2.1**: Barra de progreso integrada (`ProgressVisible`, `ProgressValue`, `ProgressColor`, `ProgressHeight`), soporte de sombra suave (`DropShadow`), efecto cristal (`GlassEffect`), evento `DblClick` y propiedad `Tag`.

![](https://user-images.githubusercontent.com/61160830/129249022-c1c2885d-cea3-4c73-8de0-8f2dcefecb1a.png)

---

### 2. AxDashBigLabel
Label de formato amplio con 3 Captions con alineación, color y fuente independientes, y un IcoFont enmarcado a uno de los costados.
- **Novedades v2.1**: Soporte de distribución horizontal (`LayoutStyle` = `lsHorizontal`), sombra suave (`DropShadow`), efecto cristal (`GlassEffect`), evento `DblClick` y propiedad `Tag`.

![](https://user-images.githubusercontent.com/61160830/129249371-80e34b2d-16e6-4ad4-8f5e-182114f46a70.png)

---

### 3. AxDashGraphLabel
Label con 3 Captions, IcoFont superior y gráfico de fondo personalizable (línea recta, curva suavizada o barras) a partir de matriz de valores (0 a 100).
- **Novedades v2.1**: Opacidad de relleno configurable (`GraphFillOpacity`), líneas de cuadrícula (`GraphGridLines`, `GraphGridColor`), sombra suave (`DropShadow`), efecto cristal (`GlassEffect`), evento `DblClick` y propiedad `Tag`.

![](https://user-images.githubusercontent.com/61160830/129249549-e8be2fce-83fe-46b6-a3ce-88a0acbd99f2.png)

---

### 4. AxDashGraphLabel2
Label con 2 Captions, IcoFont enmarcado (derecha o izquierda) y gráfico en la sección inferior (líneas, curvas, barras o gráfico circular / Pie).
- **Novedades v2.1**: Modo **Donut Chart** moderno mediante tamaño de orificio interior (`PieHoleSize` de 0 a 85%), porcentaje en slices (`PieShowPercent`), opacidad de relleno (`GraphFillOpacity`), cuadrícula de fondo (`GraphGridLines`), sombra (`DropShadow`) y efecto cristal (`GlassEffect`).

![](https://user-images.githubusercontent.com/61160830/129249663-5309b309-3cf9-41fe-987a-a4be1eff3bd5.png)

---

### 5. AxDashGaugeLabel ⭐ *(Nuevo en v2.1)*

<img width="949" height="650" alt="image" src="https://github.com/user-attachments/assets/e7c499a5-ab54-4aa2-bf54-9a74c7a85e3b" />

Medidor / Gauge circular animado con renderizado GDI+ y **Double Buffering** (sin parpadeo).
- **Estilos de Dial (`GaugeStyle`)**:
  - `gsArcDial` (270° estilo dial moderno - default)
  - `gsSemicircle` (180° estilo velocímetro)
  - `gsFullCircle` (360° círculo completo)
- **Modos de Color (`GaugeColorMode`)**:
  - `gcmSolid`: Color único configurable
  - `gcmGradient`: Degradado interpolado continuo entre `GaugeColor1` y `GaugeColor2`
  - `gcmThreshold`: Zonas semafóricas automáticas (Verde, Amarillo por `GaugeThresholdWarning` y Rojo por `GaugeThresholdDanger`)
- **Indicadores / Agujas (`GaugeNeedle`)**:
  - `gnNone` (solo arco de relleno), `gnDot` (punto en el extremo del arco), `gnLine` (aguja central delgada).
- **Animación**: Relleno fluido con aceleración/desaceleración (*exponential easing*) mediante `GaugeAnimated` y `GaugeAnimSpeed`.
- **Textos**: 3 Captions independientes (Título inferior, Valor central, Subtítulo/Unidad) más Icono central superior.

---

### 6. AxDashAnimLabel ⭐ *(Nuevo en v2.1)*

<img width="1033" height="441" alt="image" src="https://github.com/user-attachments/assets/e0368c19-f95d-46d7-9661-5394b1f3b82a" />

Gráfico de datos con **animación en tiempo real / continua** y **Double Buffering** en memoria RAM para un desplazamiento 100% fluido y libre de *flickering*.
- **Modos de Animación (`AnimMode`)**:
  - `amLoop`: Reproduce y cicla de forma continua e infinita la matriz de puntos (`GraphMatrix`) para presentaciones dinámicas.
  - `amShift`: Buffer FIFO para telemetría e ingesta de datos en tiempo real mediante el método `PushValue(NewValue)`.
- **Tipos de Gráfico (`GraphLine`)**: Línea recta (`egRectLine`), Curva Bezier suavizada (`egCurvedLine`) o Barras verticales (`egBars`).
- **Personalización**: Control de velocidad (`AnimSpeed`), paso de píxeles (`AnimStep`), líneas de cuadrícula (`GraphGridLines`), degradados y badges.

---

## 🛠️ Estructura del Proyecto

- **`Grupo Desarrollo.vbg`**: Grupo de proyectos de VB6 para abrir y probar todo el conjunto en tiempo de diseño y ejecución.
  - **`AxFramework.vbp`**: Proyecto de la librería de Controles ActiveX (`.ocx`).
  - **`TestFrameWork.vbp`**: Proyecto ejecutable de prueba con formularios de demostración:
    - `Form2.frm`: Demo de `AxDashSmallLabel`
    - `Form3.frm`: Demo de `AxDashBigLabel`
    - `Form4.frm`: Demo de `AxDashGraphLabel`
    - `Form5.frm`: Demo de `AxDashGraphLabel2`
    - `Form6.frm`: Demo de `AxDashGaugeLabel` (Gauges en vivo con simulación de CPU, RAM, Disco y GPU)
    - `Form7.frm`: Demo de `AxDashAnimLabel` (Gráficos en loop y alimentación en tiempo real vía `PushValue`)

---

## 👨‍💻 Autor y Licencia
- **Desarrollador**: David Rojas ([AxioUK](https://github.com/AxioUK))
- **Lenguaje**: Visual Basic 6.0
- **Motor Gráfico**: GDI+ (GdiPlus.dll)
