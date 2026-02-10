#Instalar paquetes necesarios
if(!require(GGally)) 
if(!require(PerformanceAnalytics)) 
if(!require(kableExtra))
  
# Cargar librerías necesarias
library(tidyverse)
library(corrplot)
library(GGally)
library(PerformanceAnalytics)
library(kableExtra)
library(patchwork)

# Carga de datos
library(readxl)
datos <- read_excel("/Users/yonshualxinico/Documentos/2025/URL/Estudio de caso/Datos arroz.xlsx")
View(datos)
attach(datos)

#### Tabla de resumen medidas de tendencia central y dispersión ###
resumen <- datos %>%
  group_by(Tratamiento) %>%
  summarise(
    n = n(),
    
    # 1. Panículas/m²
    Paniculas_media = mean(`#Paniculas/m2`, na.rm = TRUE),
    Paniculas_de = sd(`#Paniculas/m2`, na.rm = TRUE),
    Paniculas_med = median(`#Paniculas/m2`, na.rm = TRUE),
    
    # 2. Granos/panícula
    Granos_media = mean(`#Granos/panicula`, na.rm = TRUE),
    Granos_de = sd(`#Granos/panicula`, na.rm = TRUE),
    Granos_med = median(`#Granos/panicula`, na.rm = TRUE),
    
    # 3. % Llenos
    Llenos_media = mean(`%Llenos`, na.rm = TRUE),
    Llenos_de = sd(`%Llenos`, na.rm = TRUE),
    Llenos_med = median(`%Llenos`, na.rm = TRUE),
    
    # 4. Peso de 1000 granos llenos (g)
    Peso1000_media = mean(`Peso de 1000 granos llenos (g)`, na.rm = TRUE),
    Peso1000_de = sd(`Peso de 1000 granos llenos (g)`, na.rm = TRUE),
    Peso1000_med = median(`Peso de 1000 granos llenos (g)`, na.rm = TRUE),
    
    # 5. Rendimiento (kg/ha)
    Rendimiento_media = mean(`Rendimiento (kg/ha)`, na.rm = TRUE),
    Rendimiento_de = sd(`Rendimiento (kg/ha)`, na.rm = TRUE),
    Rendimiento_med = median(`Rendimiento (kg/ha)`, na.rm = TRUE)
  ) %>%
  mutate(
    # Formatear cada variable en una sola columna legible
    `#Panículas/m²` = sprintf("%.1f ± %.1f\n(%d)", 
                                  Paniculas_media, Paniculas_de, round(Paniculas_med)),
    
    `#Granos/panícula` = sprintf("%.2f ± %.2f\n(%.2f)", 
                                 Granos_media, Granos_de, Granos_med),
    
    `% Granos llenos` = sprintf("%.1f%% ± %.1f%%\n(%.1f%%)", 
                                Llenos_media, Llenos_de, Llenos_med),
    
    `Peso 1000 granos (g)` = sprintf("%.1f ± %.1f\n(%.1f)", 
                                     Peso1000_media, Peso1000_de, Peso1000_med),
    
    `Rendimiento (kg/ha)` = sprintf("%.3f ± %.3f\n(%.3f)", 
                                 Rendimiento_media, Rendimiento_de, Rendimiento_med)
  )

# Seleccionar solo las columnas 
tabla_final <- resumen %>%
  select(Tratamiento, n,
         `#Panículas/m²`,
         `#Granos/panícula`,
         `% Granos llenos`,
         `Peso 1000 granos (g)`,
         `Rendimiento (kg/ha)`)

# Mostrar tabla
kbl(tabla_final, align = "c", 
    caption = "Tabla 1. Medidas de tendencia central y dispersión para cinco variables en seis tratamientos") %>%
  kable_classic(full_width = FALSE, 
                html_font = "Arial",
                font_size = 12) %>%
  column_spec(1, bold = TRUE, width = "1.5cm") %>%
  column_spec(2, width = "1cm") %>%
  column_spec(3:7, width = "2.5cm") %>%
  footnote(
    general = "Valores presentados como: Media ± Desviación Estándar (Mediana)",
    general_title = "Nota: "
  ) %>%
  row_spec(0, bold = TRUE, background = "#f2f2f2")

# Para Word (usando flextable)
library(flextable)
ft <- flextable(tabla_final) %>%
  theme_box() %>%
  autofit()
save_as_docx(ft, path = "resumen_estadistico.docx")

#################### BOXPLOT ##########################

datos_long <- datos  %>%
  pivot_longer(
    cols = c(
      `#Paniculas/m2`,
      `#Granos/panicula`,
      `%Llenos`,
      `Peso de 1000 granos llenos (g)`,
      `Rendimiento (kg/ha)`
    ),
    names_to = "Variable",
    values_to = "Valor"
  )

ggplot(datos_long, aes(x = Tratamiento, y = Valor)) +
  geom_boxplot(
    fill = "#a6cee3",
    color = "black",
    alpha = 0.85,
    outlier.shape = 21
  ) +
  facet_wrap(
    ~ Variable,
    scales = "free_y",
    labeller = as_labeller(c(
      "#Paniculas/m2" = "Panículas por metro cuadrado",
      "#Granos/panicula" = "Granos por panícula",
      "%Llenos" = "Granos llenos (%)",
      "Peso de 1000 granos llenos (g)" = "Peso de 1,000 granos (g)",
      "Rendimiento (kg/ha)" = "Rendimiento (kg/ha)"
    ))
  ) +
  labs(
    x = "Tratamientos",
    y = "Valor observado",
    title = "Distribución de variables agronómicas por tratamiento"
  ) +
  theme_bw() +
  theme(
    legend.position = "none",
    strip.text = element_text(face = "bold", size = 14),
    axis.text.x = element_text(angle = 45, hjust = 1),
    axis.title.x = element_text(size = 14),
    axis.title.y = element_text(size = 14),
  )

############# Análisis de correlación ####################

# Seleccionar las variables de interés
variables <- datos[, c("#Paniculas/m2", 
                       "#Granos/panicula", 
                       "%Llenos", 
                       "Peso de 1000 granos llenos (g)", 
                       "Rendimiento (kg/ha)")]

# Función para el panel superior (correlaciones)
panel.cor <- function(x, y, digits = 2, prefix = "", cex.cor, ...) {
  usr <- par("usr")
  on.exit(par(usr))
  par(usr = c(0, 1, 0, 1))
  r <- cor(x, y, use = "complete.obs")
  txt <- format(c(r, 0.123456789), digits = digits)[1]
  txt <- paste0(prefix, txt)
  
  # Significancia con asteriscos
  p <- cor.test(x, y)$p.value
  stars <- ifelse(p < 0.001, "***", 
                  ifelse(p < 0.01, "**", 
                         ifelse(p < 0.1, "*", "")))
  
  # Color según magnitud
  color <- ifelse(abs(r) > 0.7, "red",
                  ifelse(abs(r) > 0.4, "black", "gray20"))
  
  if(missing(cex.cor)) cex.cor <- 0.8/strwidth(txt)
  text(0.5, 0.5, txt, cex = cex.cor * abs(r) * 2, col = color)
  text(0.5, 0.3, stars, cex = 4, col = "red")
}

# Función para el panel diagonal (histogramas)
panel.hist <- function(x, ...) {
  usr <- par("usr")
  on.exit(par(usr))
  par(usr = c(usr[1:2], 0, 1.5))
  h <- hist(x, plot = FALSE)
  breaks <- h$breaks
  nB <- length(breaks)
  y <- h$counts
  y <- y/max(y)
  rect(breaks[-nB], 0, breaks[-1], y, col = "lightblue", border = "black")
  
  # Añadir curva de densidad
  d <- density(x, na.rm = TRUE)
  d$y <- d$y/max(d$y)
  lines(d, col = "red", lwd = 2)
}

# Crear la matriz personalizada
pairs(variables,
      lower.panel = panel.smooth,    # Dispersión con línea suavizada
      upper.panel = panel.cor,        # Correlaciones
      diag.panel = panel.hist,        # Histogramas
      pch = 20,
      col = rgb(0, 0, 0, 0.5),       # Puntos semi-transparentes
      main = "Matriz de Correlación y Dispersión")

# Exportar la gráfica
png("correlograma.png", width = 3000, height = 3000, res = 300)
pairs(variables,
      lower.panel = panel.smooth,
      upper.panel = panel.cor,
      diag.panel = panel.hist,
      pch = 20,
      col = rgb(0, 0, 0, 0.5),
      main = "Matriz de Correlación y Dispersión")
dev.off()

############## ANÁLISIS DE LA VARIANZA ####################

# 1. Cargar librerías necesarias
library(agricolae) # Para comparación de medias (Tukey)
library(car)       # Para prueba de homogeneidad de varianza (Levene)

# 2. Verificación de los factores
datos$Tratamiento <- as.factor(datos$Tratamiento)
datos$Bloque <- as.factor(datos$Bloque)

# 3. Lista de las variables
variables <- c("#Paniculas/m2", "#Granos/panicula", "%Llenos", 
               "Peso de 1000 granos llenos (g)", "Rendimiento (kg/ha)")

# 4. Bucle de análisis conjuto de las variables
for (var in variables) {
  cat("\n==============================================================\n")
  cat("ANÁLISIS TÉCNICO PARA:", var, "\n")
  cat("==============================================================\n")
    
  # Crear la fórmula protegida con backticks
  formula_text <- paste0("`", var, "` ~ Bloque + Tratamiento")
  formula_modelo <- as.formula(formula_text)
    
  # ANOVA
  modelo_aov <- aov(formula_modelo, data = datos)
  print(summary(modelo_aov))
    
  # Supuesto 1: Normalidad
  shapiro <- shapiro.test(residuals(modelo_aov))
  cat("\n1. Normalidad (p-valor):", round(shapiro$p.value, 4))
    
  # Supuesto 2: Homocedasticidad (Levene corregido para Tratamiento)
  formula_levene <- reformulate(termlabels = "Tratamiento", response = paste0("`", var, "`"))
  levene <- leveneTest(formula_levene, data = datos)
  cat("\n2. Homocedasticidad (p-valor):", round(levene$`Pr(>F)`[1], 4), "\n")
    
  # Prueba de Tukey si hay significancia (p < 0.05)
  resumen <- summary(modelo_aov)
  p_valor_trat <- resumen[[1]]["Tratamiento", "Pr(>F)"]
    
  if (!is.na(p_valor_trat) && p_valor_trat < 0.05) {
    cat("\n[!] Diferencias significativas. Realizando Tukey...\n")
    prueba_tukey <- HSD.test(modelo_aov, "Tratamiento", group = TRUE)
    print(prueba_tukey$groups)
  } else {
    cat("\n[ ] Sin diferencias significativas.\n")
  }
}

########## ANÁLISIS DE COMPONENETES PRINCIPALES ################

# 1. Cargar librerías necesarias
if(!require(tidyverse)) install.packages("tidyverse")
if(!require(FactoMineR)) install.packages("FactoMineR")
if(!require(factoextra)) install.packages("factoextra")
library(tidyverse)
library(FactoMineR)
library(factoextra)

# Definir el objeto
variables_analisis <- c("#Paniculas/m2", 
                        "#Granos/panicula", 
                        "%Llenos", 
                        "Peso de 1000 granos llenos (g)", 
                        "Rendimiento (kg/ha)")

# Agrupar datos por Tratamiento 
# Calcular el promedio de cada variable para cada tratamiento.
# Esto devolverá una tabla con solo 6 filas (T1 a T6).
datos_promedio <- datos %>%
  group_by(Tratamiento) %>%
  summarise(across(all_of(variables_analisis), mean)) %>%
  column_to_rownames(var = "Tratamiento")

# 2. Ejecución del PCA sobre los datos promediados
# scale.unit = TRUE para normalizar las diferentes unidades
res_pca_promedio <- PCA(datos_promedio, scale.unit = TRUE, graph = FALSE)


# 3. Generación del Biplot
fviz_pca_biplot(res_pca_promedio,
                # --- Configuración de INDIVIDUOS (Tratamientos) ---
                # Mostrar puntos y texto para identificar cada tratamiento (T1-T6)
                geom.ind = c("point", "text"),
                # Color de los puntos (azul sólido)
                col.ind = "blue",
                pointshape = 21,   # Forma de círculo relleno
                pointsize = 3,     # Tamaño del punto
                fill.ind = "blue", # Color de relleno
                
                # --- Configuración de VARIABLES (Vectores) ---
                # Color de las flechas y texto de las variables
                col.var = "dimgray",
                # Usamos repel = TRUE para que las etiquetas no se solapen
                repel = TRUE,
                labelsize = 6,     # Tamaño de letra para variables
                
                # --- Configuración General ---
                # Desactivar las elipses para evitar el desorden visual
                addEllipses = FALSE,
                title = "Biplot de Promedios por Tratamiento (Densidades)"
) +
  # Tema minimalista para un fondo limpio
  theme_minimal() +
  theme(
    plot.title = element_text(hjust = 0.5, face = "bold"), # Título centrado y negrita
    panel.grid.major = element_blank(), # Quitar la cuadrícula mayor
    panel.grid.minor = element_blank(), # Quitar la cuadrícula menor
    axis.line = element_line(color = "black"), # Añadimos líneas a los ejes
    axis.title.x = element_text(size = 15),
    axis.title.y = element_text(size = 15)
  )
