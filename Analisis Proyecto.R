# Jeyner Arango
# Elisa Samayoa
# Oscar Méndez

library(tidyverse)
library(readxl)
library(forecast)


#=======================#
#       Análisis        #
#=======================#

df <- read_excel("precios_energia.xlsx", sheet="precios tidy")
summary(df)

#Serie de tiempo de precios
precios <- ts(df$Precio, frequency = 24, start=c(14, 0))
plot.ts(precios,
        main="Serie de tiempo del precio",
        xlab = "Días",
        ylab = "Precio",
        col = "#088F8F")

#Componentes de la serie de tiempo
componentes <- decompose(precios)
plot(componentes,
     xlab = "Días",
     col = "#088F8F")


#=========================#
#      Proyecciones       #
#=========================#

num_pred <- 24

#Modelo Exponencial
pred_exp <- forecast(precios, num_pred)
plot(pred_exp,
     main="Proyecciones del modelo Exponencial",
     xlab = "Días",
     ylab = "Precio")

#Modelo ARIMA
ar  <- auto.arima(precios)
pred_ar <- forecast(ar, num_pred)
plot(pred_ar,
     main="Proyecciones del modelo ARIMA",
     xlab = "Días",
     ylab = "Precio")

#Modelo Holt-Winters
hw <- HoltWinters(precios)
pred_hw <- forecast(hw, num_pred)
plot(pred_hw,
     main="Proyecciones del modelo Holt-Winters",
     xlab = "Días",
     ylab = "Precio")

#Evaluacion de modelos
evaluacion <- rbind(accuracy(pred_exp),
                    accuracy(pred_ar),
                    accuracy(pred_hw))
rownames(evaluacion) <- c("Exponencial",  "ARIMA", "Holt-Winters")
evaluacion
