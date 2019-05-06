# saldos_ICV
#--------------------------------
#---------- Codigo ICV ----------

library(readxl)
library(zoo)
library(dplyr)
library(reshape2)
library(xlsx)
library(sqldf)

### leer el directorio en el que están los archivos:
# archivo descargado de la superintendencia financiera, donde las 8 primeras filas tienen el encabezado
# MODALIDAD_HOMOLOGACION.xlsx
#CODIGO_NOMBRE_BANCOS.xlsx
setwd("~/trabajo1_base")


MODALIDAD_HOMOLOGACION<- read_excel("MODALIDAD_HOMOLOGACION_SALDOS.xlsx")
CODIGO_NOMBRE_BANCOS <- read_excel("CODIGO_NOMBRE_BANCOS.xlsx")



Desembolsos_interno <- function(periodo,tipo=list("Banco","CCF","CF"),nombre_arch){
  MODALIDAD_HOMOLOGACION <- read_excel("MODALIDAD_HOMOLOGACION_SALDOS.xlsx")
  CODIGO_NOMBRE_BANCOS <- read_excel("CODIGO_NOMBRE_BANCOS.xlsx")
  if(tipo=="Banco" | tipo=="CCF" | tipo=="CF"){
    hoja <- ifelse(tipo=="Banco",1,ifelse(tipo=="CCF",4,3))

saldo_112018 <- read_excel(nombre_arch, skip = 8,sheet=hoja)
saldo_112018_1 <- saldo_112018
saldo_112018_1$MODALIDAD <- na.locf(saldo_112018$MODALIDAD)
saldo_112018_1$PRODUCTO <- na.locf(saldo_112018$PRODUCTO)
saldo_112018_1[is.na(saldo_112018_1)]=0


saldo_112018_2 <- saldo_112018_1%>%melt(id=c("MODALIDAD","PRODUCTO","CONCEPTO"),value.name = "Valor")%>%
  mutate(cod_ent=substr(variable,start=1,stop=6))%>%mutate(cod_ent= case_when( cod_ent=="999 TO"~"999",TRUE~cod_ent))%>%
  select(MODALIDAD,PRODUCTO,CONCEPTO,cod_ent,Valor,-variable)

periodo="201811"

saldo_112018_3 <- saldo_112018_2%>%mutate(Corte=rep(periodo,dim(saldo_112018_2)[1]))%>%
  mutate(Modalidad_Tipo_Pto=case_when(
    MODALIDAD=="CONSUMO"~ paste("CO",PRODUCTO,sep = "_"),
    MODALIDAD=="LEASING"~ paste("LE",PRODUCTO,sep = "_"),
    MODALIDAD=="VIVIENDA"~ paste("VI",PRODUCTO,sep = "_"),
    MODALIDAD=="COMERCIAL"~ paste("COM",PRODUCTO,sep = "_"),
    MODALIDAD=="MICROCREDITO"~ paste("MI",PRODUCTO,sep = "_")
  ))%>%mutate(Modalidad_Tipo_Pto=trimws(Modalidad_Tipo_Pto))
saldo_112018_3_1 <- merge(x=saldo_112018_3,y=MODALIDAD_HOMOLOGACION,by="Modalidad_Tipo_Pto",all.x=TRUE);colnames(saldo_112018_3_1) <- c(colnames(saldo_112018_3_1)[1:7],"Tipo_Final_Pdto")
saldo_112018_3_2 <- merge(x=saldo_112018_3_1,y=CODIGO_NOMBRE_BANCOS[,-2],by.x="cod_ent",by.y="Cod_Banco",all.x=TRUE);colnames(saldo_112018_3_2) <- c(colnames(saldo_112018_3_2)[1:8],"Entidad")
saldo_112018_3 <- saldo_112018_3_2  
saldo_112018_3$CONCEPTO[saldo_112018_3$CONCEPTO=="Monto"]="Valor"; saldo_112018_3$CONCEPTO[saldo_112018_3$CONCEPTO=="Creditos"]="N_Creditos"
saldo_112018_4 <- dcast(saldo_112018_3,MODALIDAD+PRODUCTO+cod_ent+Corte+Modalidad_Tipo_Pto+Tipo_Final_Pdto+Entidad~CONCEPTO, value.var="Valor",fill = 0)%>%
  mutate(Tipo_Ent=tipo)
saldo_112018_4$ICV=saldo_112018_4$`Cartera Vencida`/saldo_112018_4$`Saldo Cartera`
saldo_112018_4 <- saldo_112018_4[as.logical(1-is.na(saldo_112018_4$ICV)),]
return(saldo_112018_4)
}
else return(cat("El Valor en Tipo no es admisible"))  
}
