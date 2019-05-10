#----------------------------------------
#---------- Código desembolsos ----------

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


Desembolsos_final <- function(tipo=1,retur1,arch1,nombre_inicial,nombre_final=paste("tbl_desembolsos",periodo,format(Sys.time(), "%Y%m%d"),".csv",sep = "_")){
# TIPO=1 Es para la tabla  de un solo mes, TIPO=2  es para la tabla de varios meses
# RETUR1= 2 devuelve la tabla larga para seguir trabajando en R, RETUR<>2 no devuelve la tabla larga
# ARCH1=1 Genera las tablas en excel la larga y la que esta agrupada
# nombre inicial: 
#             si es TIPO=1 entonces es nombre es de tipo Ej."022019desembolsoaprobaciones.xls" DEBE TENER EL PERIODO AL INICIO EN ESE ORDEN
#             si es TIPO=2 entonces es nombre es de tipo Ej.'*desembolsoaprobaciones.xls' y tomará todos los excel con este tipo de nombre que esten en la carpeta
  #nombre_final: No es obligatorio ponerlo. Será de la forma Ej. "nombre_archivo.csv"


      
  Desembolsos_interno <- function(periodo,tipo=list("Banco","CCF","CF"),nombre_arch){
    MODALIDAD_HOMOLOGACION <- read_excel("MODALIDAD_HOMOLOGACION.xlsx")
    CODIGO_NOMBRE_BANCOS <- read_excel("CODIGO_NOMBRE_BANCOS.xlsx")
    if(tipo=="Banco" | tipo=="CCF" | tipo=="CF"){
      hoja <- ifelse(tipo=="Banco",1,ifelse(tipo=="CCF",3,2))
      tbl_201902<- read_excel(nombre_arch, skip = 8,sheet=hoja)
      tbl_201902_1 <- tbl_201902 
      tbl_201902_1$MODALIDAD <- na.locf(tbl_201902$MODALIDAD)
      tbl_201902_1$PRODUCTO <- na.locf(tbl_201902$PRODUCTO)
      tbl_201902_1[is.na(tbl_201902_1)]=0
      
      tbl_201902_2 <- tbl_201902_1%>%mutate(PRODUCTO=trimws(PRODUCTO))%>%melt(id=c("MODALIDAD","PRODUCTO","CONCEPTO"),value.name = "Valor")%>%
        mutate(cod_ent=substr(variable,start=1,stop=6))%>%mutate(cod_ent= case_when( cod_ent=="999 TO"~"999",TRUE~cod_ent))%>%
        select(MODALIDAD,PRODUCTO,CONCEPTO,cod_ent,Valor,-variable)
      tbl_201902_3 <- tbl_201902_2%>%mutate(Corte=rep(periodo,dim(tbl_201902_2)[1]))%>%
        mutate(Modalidad_Tipo_Pto=case_when(
          MODALIDAD=="CONSUMO"~ paste("CO",PRODUCTO,sep = "_"),
          MODALIDAD=="LEASING"~ paste("LE",PRODUCTO,sep = "_"),
          MODALIDAD=="VIVIENDA"~ paste("VI",PRODUCTO,sep = "_"),
          MODALIDAD=="MICROCREDITO"~ paste("MI",PRODUCTO,sep = "_")
        ))%>%mutate(Modalidad_Tipo_Pto=trimws(Modalidad_Tipo_Pto))
      tbl_201902_3_1 <- merge(x=tbl_201902_3,y=MODALIDAD_HOMOLOGACION,by="Modalidad_Tipo_Pto",all.x=TRUE);colnames(tbl_201902_3_1) <- c(colnames(tbl_201902_3_1)[1:7],"Tipo_Final_Pdto")
      tbl_201902_3_2 <- merge(x=tbl_201902_3_1,y=CODIGO_NOMBRE_BANCOS[,-2],by.x="cod_ent",by.y="Cod_Banco",all.x=TRUE);colnames(tbl_201902_3_2) <- c(colnames(tbl_201902_3_2)[1:8],"Entidad")
      tbl_201902_3 <- tbl_201902_3_2
      tbl_201902_3$CONCEPTO[tbl_201902_3$CONCEPTO=="Monto"]="Valor"; tbl_201902_3$CONCEPTO[tbl_201902_3$CONCEPTO=="Creditos"]="N_Creditos"
      tbl_201902_4 <- dcast(tbl_201902_3,MODALIDAD+PRODUCTO+cod_ent+Corte+Modalidad_Tipo_Pto+Tipo_Final_Pdto+Entidad~CONCEPTO, value.var="Valor")%>%
        mutate(Tipo_Ent=tipo)
      return(tbl_201902_4)
    }
    else return(cat("El Valor en Tipo no es admisible"))  
  }
  
  
  Desembolsos_completo <- function(retur=1,arch=1,nombre_arch,nombre_final=paste("tbl_desembolsos",periodo,format(Sys.time(), "%Y%m%d"),".csv",sep = "_")){
    
    ## con arch=1 genera 2 excel, el 1 es la tabla y el 2 es la tabla agrupada   
    ## retur=2 retorna la tabla larga  
    periodo=paste(substr(nombre_arch,start = 3,stop = 6),substr(nombre_arch,start = 1,stop = 2),sep = "")
    peri=periodo
    Bancos <- Desembolsos_interno(periodo = peri,tipo = "Banco",nombre_arch)
    CCF <- Desembolsos_interno(periodo = peri,tipo = "CCF",nombre_arch)
    CF <- Desembolsos_interno(periodo = peri,tipo = "CF",nombre_arch)
    completo <- rbind(Bancos,CCF,CF)
    completo_2 <- completo%>%mutate(AG_Trimestre=case_when(
      substr(Corte,start=5,stop=6)=="01" |  substr(Corte,start=5,stop=6)=="02" | substr(Corte,start=5,stop=6)=="03"~paste(substr(Corte,start=1,stop=4),"I",sep = "_"),
      substr(Corte,start=5,stop=6)=="04" |  substr(Corte,start=5,stop=6)=="05" | substr(Corte,start=5,stop=6)=="06"~paste(substr(Corte,start=1,stop=4),"II",sep = "_"),
      substr(Corte,start=5,stop=6)=="07" |  substr(Corte,start=5,stop=6)=="08" | substr(Corte,start=5,stop=6)=="09"~paste(substr(Corte,start=1,stop=4),"III",sep = "_"),
      substr(Corte,start=5,stop=6)=="10" |  substr(Corte,start=5,stop=6)=="11"| substr(Corte,start=5,stop=6)=="12"~ paste(substr(Corte,start=1,stop=4),"IV",sep = "_")
    ))%>% mutate(AG_Semestre=case_when(
      substr(Corte,start=5,stop=6)=="01" |substr(Corte,start=5,stop=6)=="02" |substr(Corte,start=5,stop=6)=="03" |substr(Corte,start=5,stop=6)=="04" |substr(Corte,start=5,stop=6)=="05" | substr(Corte,start=5,stop=6)<="06"~paste(substr(Corte,start=1,stop=4),"I",sep = "_"),
      substr(Corte,start=5,stop=6)=="07" |substr(Corte,start=5,stop=6)=="08" |substr(Corte,start=5,stop=6)=="09" |substr(Corte,start=5,stop=6)=="10" |substr(Corte,start=5,stop=6)=="11" | substr(Corte,start=5,stop=6)<="12"~paste(substr(Corte,start=1,stop=4),"II",sep = "_")
    ))%>%mutate(M_Corte_Trim=case_when(
      substr(Corte,start=5,stop=6)=="03"~paste(substr(Corte,start=1,stop=4),"I",sep = "_"),
      substr(Corte,start=5,stop=6)=="06"~paste(substr(Corte,start=1,stop=4),"II",sep = "_"),
      substr(Corte,start=5,stop=6)=="09"~paste(substr(Corte,start=1,stop=4),"III",sep = "_"),
      substr(Corte,start=5,stop=6)=="12"~ paste(substr(Corte,start=1,stop=4),"IV",sep = "_"),
      TRUE~"NA"))%>% mutate(M_Corte_Sem=case_when(
        substr(Corte,start=5,stop=6)=="06"~paste(substr(Corte,start=1,stop=4),"I",sep = "_"),
        substr(Corte,start=5,stop=6)=="12"~ paste(substr(Corte,start=1,stop=4),"II",sep = "_"),
        TRUE~"NA"))%>% mutate(Clas_pdto_MIS=ifelse(substr(Tipo_Final_Pdto,start=1,stop=2)=="CO"&Tipo_Final_Pdto!="CO_T_Credito","CO_SIN_TDC",Tipo_Final_Pdto)) #
    
    indice_2 <- completo_2$N_Creditos!=0 & completo_2$Valor!=0
    completo_3 <- completo_2[indice_2,]%>%select(MODALIDAD,PRODUCTO,cod_ent,N_Creditos,Valor,Corte,Modalidad_Tipo_Pto,Tipo_Final_Pdto,Entidad,AG_Trimestre,AG_Semestre,M_Corte_Trim,M_Corte_Sem,Clas_pdto_MIS,Tipo_Ent)
    
    
    agrupado <- sqldf('select Corte,AG_Trimestre,AG_Semestre,M_Corte_Trim,M_Corte_Sem,Tipo_Ent,Entidad,MODALIDAD,Tipo_Final_Pdto,Clas_pdto_MIS,
         sum(N_Creditos) as Suma_de_N_Creditos,(sum(Valor)/sum(N_Creditos)) as Suma_de_Prom_Desem, sum(Valor) as Suma_de_Valor
                  from completo_3
                  group by Corte,AG_Trimestre,AG_Semestre,M_Corte_Trim,M_Corte_Sem,Tipo_Ent,Entidad,MODALIDAD,Tipo_Final_Pdto,Clas_pdto_MIS')
     if(arch==1){
      write.csv(completo_3, file = nombre_final,row.names=FALSE)
      write.csv(agrupado, file = paste("TD_BaseInsumo",peri,".csv"),row.names=FALSE)
    }
    
    if(retur==2){ return(completo_3)}
    else return()
  }  
  
  Desembolsos_varios <- function(retur=1,arch=1,nombre_comun,nombre_final=paste("tbl_desembolsos_varios",format(Sys.time(), "%Y%m%d"),".csv",sep = "_")){
    # el nombre comun debe ser de la forma  Ej= '*desembolsoaprobaciones.xls'
    
    file.list <- list.files(pattern=nombre_comun)
    n <- length(file.list);tabla_c <- rep(0,14)
    for( i in 1:n){
      tabla <- Desembolsos_completo(arch=2,retur=2,nombre_arch = file.list[i])
      tabla_c <- rbind(tabla_c,tabla)
    }
    tabla_c <- tabla_c[-1,]
    
    agrupado <- sqldf('select Corte,AG_Trimestre,AG_Semestre,M_Corte_Trim,M_Corte_Sem,Tipo_Ent,Entidad,Modalidad,Tipo_Final_Pdto,Clas_pdto_MIS,
sum(N_Creditos) as Suma_de_N_Creditos,(sum(Valor)/sum(N_Creditos)) as Suma_de_Prom_Desem, sum(Valor) as Suma_de_Valor
                      from tabla_c
                      group by Corte,AG_Trimestre,AG_Semestre,M_Corte_Trim,M_Corte_Sem,Tipo_Ent,Entidad,Modalidad,Tipo_Final_Pdto,Clas_pdto_MIS')
    if(arch==1){
      write.csv(tabla_c[,-1], file = nombre_final,row.names=FALSE)
      write.csv(agrupado, file = paste("TD_BaseInsumo",format(Sys.time(), "%Y%m%d"),".csv"),row.names=FALSE)
    }
    
    if(retur==2){ return(tabla_c )}
    else return()
  }  
  
  
if(missing(nombre_inicial)){ return( cat("ERROR: Ingrese un nombre inicial válido"))}
else{  
  if(tipo==1){
    if(missing(nombre_final)){
      if(missing(retur1) & missing(arch1)){return(Desembolsos_completo(nombre_arch = nombre_inicial))}
      if(missing(retur1)==FALSE & missing(arch1)){return(Desembolsos_completo(retur = retur1,nombre_arch = nombre_inicial))}
      if(missing(retur1) & missing(arch1)==FALSE){return(Desembolsos_completo(arch = arch1,nombre_arch = nombre_inicial))}
      if(missing(retur1)==FALSE & missing(arch1)==FALSE){return(Desembolsos_completo(retur = retur1,arch = arch1,nombre_arch = nombre_inicial))}
       }
    else{
      if(missing(retur1) & missing(arch1)){return(Desembolsos_completo(nombre_arch = nombre_inicial,nombre_final = nombre_final))}
      if(missing(retur1)==FALSE & missing(arch1)){return(Desembolsos_completo(retur = retur1,nombre_arch = nombre_inicial,nombre_final = nombre_final))}
      if(missing(retur1) & missing(arch1)==FALSE){return(Desembolsos_completo(arch = arch1,nombre_arch = nombre_inicial,nombre_final = nombre_final))}
      if(missing(retur1)==FALSE & missing(arch1)==FALSE){return(Desembolsos_completo(retur = retur1,arch = arch1,nombre_arch = nombre_inicial,nombre_final = nombre_final))}
       }
     }

  if(tipo==2){
    if(missing(nombre_final)){
      if(missing(retur1) & missing(arch1)){return(Desembolsos_varios(nombre_comun = nombre_inicial))}
      if(missing(retur1)==FALSE & missing(arch1)){return(Desembolsos_varios(retur = retur1,nombre_comun = nombre_inicial))}
      if(missing(retur1) & missing(arch1)==FALSE){return(Desembolsos_varios(arch = arch1,nombre_comun = nombre_inicial))}
      if(missing(retur1)==FALSE & missing(arch1)==FALSE){return(Desembolsos_varios(retur = retur1,arch = arch1,nombre_comun = nombre_inicial))}
    }
    else{
      if(missing(retur1) & missing(arch1)){return(Desembolsos_varios(nombre_comun= nombre_inicial,nombre_final = nombre_final))}
      if(missing(retur1)==FALSE & missing(arch1)){return(Desembolsos_varios(retur = retur1,nombre_comun = nombre_inicial,nombre_final = nombre_final))}
      if(missing(retur1) & missing(arch1)==FALSE){return(Desembolsos_varios(arch = arch1,nombre_comun= nombre_inicial,nombre_final = nombre_final))}
      if(missing(retur1)==FALSE & missing(arch1)==FALSE){return(Desembolsos_varios(retur = retur1,arch = arch1,nombre_comun= nombre_inicial,nombre_final = nombre_final))}
      }
    }  
  if(tipo>2){cat("ERROR:Tipo solo puede tener valores 1 o 2")} 
   }
}

setwd("~/trabajo1_base/desembolsos")
Desembolsos_final(tipo = 2,nombre_inicial= "*desembolsoaprobaciones.xls",nombre_final = "desem_completo_hoy.csv")
Desembolsos_final(tipo = 1,nombre_inicial= "012015desembolsoaprobaciones.xls")


