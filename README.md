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



Saldos_final <- function(tipo=1,retur1,arch1,nombre_inicial,nombre_final=paste("tbl_saldos",periodo,format(Sys.time(), "%Y%m%d"),".csv",sep = "_")){
  # TIPO=1 Es para la tabla  de un solo mes, TIPO=2  es para la tabla de varios meses
  # RETUR1= 2 devuelve la tabla  para seguir trabajando en R, RETUR<>2 no devuelve la tabla
  # ARCH1=1 Genera la tabla en excel
  # nombre inicial: 
  #             si es TIPO=1 entonces es nombre es de tipo Ej."022019saldosindicadorpormora.xls" DEBE TENER EL PERIODO AL INICIO EN ESE ORDEN
  #             si es TIPO=2 entonces es nombre es de tipo Ej.'*saldosindicadorpormora.xls' y tomará todos los excel con este tipo de nombre que esten en la carpeta
  #nombre_final: No es obligatorio ponerlo. Será de la forma Ej. "nombre_archivo.csv"
  
  


  Saldos_interno <- function(periodo,tipo=list("Banco","CCF","CF"),nombre_arch){
    MODALIDAD_HOMOLOGACION <- read_excel("MODALIDAD_HOMOLOGACION_SALDOS.xlsx")
    CODIGO_NOMBRE_BANCOS <- read_excel("CODIGO_NOMBRE_BANCOS.xlsx")
    if(tipo=="Banco" | tipo=="CCF" | tipo=="CF"){
      hoja <- ifelse(tipo=="Banco",1,ifelse(tipo=="CCF",4,3))
      
      saldo_112018 <- read_excel(nombre_arch, skip = 8,sheet=hoja)
      saldo_112018_1 <- saldo_112018
      saldo_112018_1$MODALIDAD <- na.locf(saldo_112018$MODALIDAD)
      saldo_112018_1$PRODUCTO <- na.locf(saldo_112018$PRODUCTO)
      saldo_112018_1[is.na(saldo_112018_1)]=0
      
      
      saldo_112018_2 <- saldo_112018_1%>%mutate(PRODUCTO=trimws(PRODUCTO))%>%melt(id=c("MODALIDAD","PRODUCTO","CONCEPTO"),value.name = "Valor")%>%
        mutate(cod_ent=substr(variable,start=1,stop=6))%>%mutate(cod_ent= case_when( cod_ent=="999 TO"~"999",TRUE~cod_ent))%>%
        select(MODALIDAD,PRODUCTO,CONCEPTO,cod_ent,Valor,-variable)
      
      
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
      saldo_112018_4 <- dcast(saldo_112018_3,MODALIDAD+PRODUCTO+cod_ent+Corte+Modalidad_Tipo_Pto+Tipo_Final_Pdto+Entidad~CONCEPTO, value.var="Valor",fill = 0,fun.aggregate=sum)%>%
        mutate(Tipo_Ent=tipo)
      saldo_112018_4$ICV=saldo_112018_4$`Cartera Vencida`/saldo_112018_4$`Saldo Cartera`
      index <- as.logical(1-is.na(saldo_112018_4$ICV))
      saldo_112018_4_1 <- saldo_112018_4[index,]
      return(saldo_112018_4_1)
    }
    else return(cat("El Valor en Tipo no es admisible"))  
  }
  
  
  
  ### parte 2 coger el libro completo
  
  Saldos_completo <- function(retur=1,arch=1,nombre_arch,nombre_final=paste("tbl_saldos",periodo,format(Sys.time(), "%Y%m%d"),".csv",sep = "_")){
    
    ## con arch=1 genera 2 excel, el 1 es la tabla y el 2 es la tabla agrupada   
    ## retur=2 retorna la tabla larga  
    periodo=paste(substr(nombre_arch,start = 3,stop = 6),substr(nombre_arch,start = 1,stop = 2),sep = "")
    peri=periodo
    Bancos <- Saldos_interno(periodo = peri,tipo = "Banco",nombre_arch)
    CCF <- Saldos_interno(periodo = peri,tipo = "CCF",nombre_arch)
    CF <- Saldos_interno(periodo = peri,tipo = "CF",nombre_arch)
    completo <- rbind(Bancos,CCF,CF)
    if(arch==1){
      write.csv(completo, file = nombre_final,row.names=FALSE)
    }
    
    if(retur==2){ return(completo)}
    else return()
  }
  
  
  Saldos_varios <- function(retur=1,arch=1,nombre_comun,nombre_final=paste("tbl_saldos_varios",format(Sys.time(), "%Y%m%d"),".csv",sep = "_")){
    # el nombre comun debe ser de la forma  Ej= '*desembolsoaprobaciones.xls'
    
    file.list <- list.files(pattern=nombre_comun)
    n <- length(file.list);tabla_c <- rep(0,28)
    for( i in 1:n){
      tabla <- Saldos_completo(arch=2,retur=2,nombre_arch = file.list[i])
      #cat("i=",i,"dim_t",dim(tabla),"\n")
      tabla_c <- rbind(tabla_c,tabla)
    }
    tabla_c <- tabla_c[-1,]
    
    if(arch==1){
      write.csv(tabla_c, file = nombre_final,row.names=FALSE)
    }
    
    if(retur==2){ return(tabla_c )}
    else return()
  } 




if(missing(nombre_inicial)){ return( cat("ERROR: Ingrese un nombre inicial válido"))}
else{  
  if(tipo==1){
    if(missing(nombre_final)){
      if(missing(retur1) & missing(arch1)){return(Saldos_completo(nombre_arch = nombre_inicial))}
      if(missing(retur1)==FALSE & missing(arch1)){return(Saldos_completo(retur = retur1,nombre_arch = nombre_inicial))}
      if(missing(retur1) & missing(arch1)==FALSE){return(Saldos_completo(arch = arch1,nombre_arch = nombre_inicial))}
      if(missing(retur1)==FALSE & missing(arch1)==FALSE){return(Saldos_completo(retur = retur1,arch = arch1,nombre_arch = nombre_inicial))}
    }
    else{
      if(missing(retur1) & missing(arch1)){return(Saldos_completo(nombre_arch = nombre_inicial,nombre_final = nombre_final))}
      if(missing(retur1)==FALSE & missing(arch1)){return(Saldos_completo(retur = retur1,nombre_arch = nombre_inicial,nombre_final = nombre_final))}
      if(missing(retur1) & missing(arch1)==FALSE){return(Saldos_completo(arch = arch1,nombre_arch = nombre_inicial,nombre_final = nombre_final))}
      if(missing(retur1)==FALSE & missing(arch1)==FALSE){return(Saldos_completo(retur = retur1,arch = arch1,nombre_arch = nombre_inicial,nombre_final = nombre_final))}
    }
  }
  
  if(tipo==2){
    if(missing(nombre_final)){
      if(missing(retur1) & missing(arch1)){return(Saldos_varios(nombre_comun = nombre_inicial))}
      if(missing(retur1)==FALSE & missing(arch1)){return(Saldos_varios(retur = retur1,nombre_comun = nombre_inicial))}
      if(missing(retur1) & missing(arch1)==FALSE){return(Saldos_varios(arch = arch1,nombre_comun = nombre_inicial))}
      if(missing(retur1)==FALSE & missing(arch1)==FALSE){return(Saldos_varios(retur = retur1,arch = arch1,nombre_comun = nombre_inicial))}
    }
    else{
      if(missing(retur1) & missing(arch1)){return(Saldos_varios(nombre_arch = nombre_comun,nombre_final = nombre_final))}
      if(missing(retur1)==FALSE & missing(arch1)){return(Saldos_varios(retur = retur1,nombre_comun = nombre_inicial,nombre_final = nombre_final))}
      if(missing(retur1) & missing(arch1)==FALSE){return(Saldos_varios(arch = arch1,nombre_comun= nombre_inicial,nombre_final = nombre_final))}
      if(missing(retur1)==FALSE & missing(arch1)==FALSE){return(Saldos_varios(retur = retur1,arch = arch1,nombre_comun= nombre_inicial,nombre_final = nombre_final))}
    }
  }  
  if(tipo>2){cat("ERROR:Tipo solo puede tener valores 1 o 2")} 
 }

}


Saldos_final(tipo=1,nombre_inicial = '102016saldosindicadorpormora.xls')

