# STATCAL ONLINE - Categorical Frequency Distribution and Bivariate Bar Chart Analyzer

This R Shiny application analyzes categorical variables such as Pendidikan, Jenis Kelamin, Golongan Darah, and Pekerjaan.

## Features

- Upload Excel data or use sample `data variabel kategori.xlsx`.
- Univariate frequency and percentage tables with flexible decimal digits.
- Bivariate frequency and percentage tables with one dependent variable and multiple independent variables.
- Univariate frequency/percentage bar charts with flexible panels, colors, text sizes, orientation, and publication-style backgrounds.
- Bivariate frequency/percentage bar charts with stacked or grouped bars, flexible panels, colors, text sizes, orientation, and publication-style backgrounds.
- Export univariate and bivariate charts as high-resolution PNG with 300, 600, 900, 1200, and 1500 DPI.
- Export tables and chart data to Excel.

## Install packages

```r
install.packages(c(
  "shiny", "shinydashboard", "DT", "readxl", "dplyr", "tidyr",
  "ggplot2", "shinycssloaders", "scales", "openxlsx"
))
```

## Run app

```r
shiny::runApp(".")
```

Training data URL:
https://drive.google.com/drive/folders/1x623-vbvaCj7-YA-HljZXTPU7LSdgJmW?usp=sharing
