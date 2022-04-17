list.of.packages = c("pdftools","pdfsearch","readxl","openxlsx","stringr")
new.packages2=list.of.packages[!(list.of.packages %in% installed.packages()[,"Package"])]
if(length(new.packages2)!=0) install.packages(new.packages2)

library(pdftools) #
library(pdfsearch) #
library(readxl) #
library(openxlsx) #
library(stringr) #



setwd("C:\\Users\\52811\\Desktop\\RProjects\\XXXXX\\DesarrolloFuentes\\EstadoDeCuenta")
archivo=list.files()
archivo=archivo[substr(archivo,nchar(archivo)-3,nchar(archivo))==".pdf"]
estado_de_cuenta=pdf_text(paste("C:\\Users\\52811\\Desktop\\RProjects\\XXXXX\\DesarrolloFuentes\\EstadoDeCuenta\\",archivo,sep=""))

### data para ligar
asegu_concept=read_excel("C:\\Users\\52811\\Desktop\\RProjects\\XXXXX\\DesarrolloFuentes\\BaseDatos\\ZZZZZData.xlsx",sheet = "Cuentas")
asegu_concepto=asegu_concept
asegu_concept=as.data.frame(asegu_concepto[,c(1,3)])
asegu_concept=asegu_concept[!is.na(asegu_concept[,2]),]
asegu_concept=asegu_concept[!is.na(asegu_concept[,1]),]
asegu_concept=asegu_concept[(asegu_concept[,1])!="-",]
asegu_concept1=asegu_concept[nchar(asegu_concept[,2])>6,]
asegu_concept2=asegu_concept[nchar(asegu_concept[,2])<=6,]
rfc_info=asegu_concepto[,c(1,4)]
rfc_info=na.omit(rfc_info)
rfc_info=rfc_info[rfc_info[,2]!="---------",]
rfc_info=as.data.frame(rfc_info)
#rfc_info=rfc_info[!duplicated(rfc_info[,2]),]
claveEGR=na.omit(asegu_concepto[,c(1,2,2)])
###saber paginas de EDC con mov
aqui_para=grepl("REGIONOMINA",estado_de_cuenta)
aqui_para2=grepl("FOLIO FISCAL",estado_de_cuenta)
aqui_para3=grepl("Transaccional",estado_de_cuenta)
alto=F
alto2=F
alto3=F
edc=0
edc2=0
edc3=0
for(gg in 1:length(estado_de_cuenta)){
  if(alto==F&&aqui_para[gg]){alto=T
  edc=gg
  }
  if(alto2==F&&aqui_para2[gg]){alto2=T
  edc2=gg
  }
  if(alto3==F&&aqui_para3[gg]){alto3=T
  edc3=gg
  }
}
if(edc==0) edc=1000
if(edc2==0) edc2=1000
if(edc3==0) edc2=1000
edc=min(edc,edc2,edc3)
info2=NA
for(qq in 1:edc){
  por_pagina=estado_de_cuenta[qq]
  
  por_renglon=str_split(por_pagina,"\r\n")
  por_renglon=trimws(por_renglon[[1]])
  
  se.va=0
  se.va2=0
  for(i in 1:length(por_renglon)){
    if(substr(por_renglon[i],1,4)=="Page" | substr(por_renglon[i],1,4)=="Pági") se.va2=i
  }
  page1=por_renglon[se.va2]
  page1=gsub("\\/.*","",page1)
  page1=substr(page1,1,nchar("Page xx of"))
  page1=trimws(str_remove_all(page1,"[ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzáÁ]"))
  
  if(page1==1){
    fecha1=por_renglon[substr(por_renglon,1,6)=="UCCALY"]
    fecha1=trimws(str_remove(fecha1,"UCCALY"))
    fecha1=trimws(substr(fecha1,nchar("del 01 al 30 de "),nchar(fecha1)))
    anio=substr(fecha1,nchar(fecha1)-3,nchar(fecha1))
    mes=trimws(substr(fecha1,1,nchar(fecha1)-nchar(anio)))
  }
  
  
  if(any(page1==2:qq)&&qq!=1){
    
    #datos
    Comisiones_Efectivamente_Cobradas =por_renglon[grepl("Comisiones Efectivamente Cobradas",x = por_renglon)]
    saldo_inicial=por_renglon[grepl("Saldo Inicial",x = por_renglon)]
    if(page1==2)saldo_anterior=trimws(str_remove_all(saldo_inicial,"[ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz]"))
    #limpia de info
    se.va=0
    se.va2=0
    for(i in 1:length(por_renglon)){
      if(substr(por_renglon[i],1,3)=="DIA") se.va=i-1
      if(substr(por_renglon[i],1,4)=="Page"|| substr(por_renglon[i],1,4)=="Pági") se.va2=i
    }
    por_renglon=por_renglon[-c(1:se.va)]
    por_renglon=por_renglon[-c((se.va2-se.va):length(por_renglon))]
    
    
    info=matrix(nrow=length(por_renglon),ncol=6)
    dia=NA
    for(i in 1:nrow(info)){
      bb=F
      if(is.na(as.numeric(substr(por_renglon[i],1,2)))==FALSE &&substr(por_renglon[i],3,3)==" "&&nchar(por_renglon[i])>=25){
        data_spt=strsplit(por_renglon[i],"      ")
        son.data=NA
        for(ff in 1:length(data_spt[[1]])){
          if(data_spt[[1]][ff]!="")son.data=rbind(na.omit(son.data),ff)
        }
        
        saldo=trimws(data_spt[[1]][son.data[nrow(son.data)]])
        
        if(son.data[nrow(son.data)]-son.data[nrow(son.data)-1]>=2) {cargo= trimws(data_spt[[1]][son.data[nrow(son.data)-1]])
        abono=NA} else{ abono= trimws(data_spt[[1]][son.data[nrow(son.data)-1]]) 
        cargo=NA}
        
        dia=substr(por_renglon[i],1,2)
        info[i,1]=paste(dia,"/",substr(mes,1,3),"/",anio,sep="")
        info[i,3]=saldo_anterior
        info[i,4]=cargo
        info[i,5]=abono
        info[i,6]=saldo
        bb=T
        if(as.numeric(page1)==2&&i==1) saldo_anterior=trimws(str_remove_all(saldo_inicial,"[ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz]")) else saldo_anterior=saldo
      }
      if(bb){                                 
        data=trimws(substr(por_renglon[i],4,99),which = "right")} else{ data=trimws(substr(por_renglon[i],1,99),which = "right")
        info[i,1]="A"}
      info[i,2]=data
      
    }
    info=info[-1,]
    colnames(info)=c("Dia","Datos","SaldoInicial","Cargo","Abono","SaldoFinal")
  }
  if(qq!=1){
    info2=rbind(info2,info)}
}

info2=info2[-1,]
info3=info2
info4=matrix(nrow = sum(info2[,1]!="A"),ncol=ncol(info3))
rr=0
por_pago_info=NA
for(i in 1:nrow(info3)){
  if(info3[i,1]!="A"){
    if(is.na(por_pago_info)==F){
      info4[rr,2]=por_pago_info
      por_pago_info=NA
    }
    rr=rr+1
    info4[rr,1]=info3[i,1]
    info4[rr,2]=info3[i,2]
    info4[rr,3]=info3[i,3]
    info4[rr,4]=info3[i,4]
    info4[rr,5]=info3[i,5]
    info4[rr,6]=info3[i,6]
    por_pago_info=(info3[i,2])} else {
      #por_pago_info=cbind(por_pago_info,info3[i,2])
      por_pago_info=paste(por_pago_info,info3[i,2])
    }
  
}

aseguradoras=c("AA","BB","CC","DD","EE","FF","GG","HH","II","JJ","KK","LL","MM","NN","OO","PP","QQ","RR","SS","TT","UU","VV","WW","XX","YY","ZZ","OTRAS") #HERE GO THE INSURANCE NAMES
#GRUPO MEXICANO DE SE GUROS - GMX
info5=matrix(nrow = nrow(info4),ncol=2)
RFC_duplicado=matrix(nrow=nrow(info4))
for(i in 1:nrow(info4)){ 
  if(!is.na(info4[i,4])){ #Cargos
    if(grepl(info4[i,2],pattern="RFC:",ignore.case = T)&&grepl(info4[i,2],pattern="IVA",ignore.case = T)){
      cual_es=sapply(rfc_info[,2],grepl,info4[i,2]) 
      if(sum(cual_es)==1){
        info5[i,1]=rfc_info[cual_es,1]
        concepto=info4[i,2]
        #concepto=str_remove(concepto,"-")
        concepto=gsub(".*-","",concepto)
        info5[i,2]=concepto} else if (sum(cual_es)==0) {
          concepto=info4[i,2]
          #concepto=str_remove(concepto,"-")
          concepto=gsub(".*-","",concepto)
          info5[i,1]=concepto
          info5[i,2]=info4[i,2]} else if(sum(cual_es)>1){
            concepto=info4[i,2]
            #concepto=str_remove(concepto,"-")
            concepto=gsub(".*-","",concepto)
            info5[i,1]=concepto
            info5[i,2]=info4[i,2]
            RFC_duplicado[i,]="RFC Dup"
            
          }
    }else if(grepl(info4[i,2],pattern = "CHEQUE")|grepl(info4[i,2],pattern = "Cheque")){
      cheque= gsub("\\-.*","",info4[i,2])
      cheque=str_replace(cheque,"DOC","CHEQUE")
      info5[i,1]="Preguntar Proveedor/Compañia"
      info5[i,2]=cheque
    }else if(str_count(info4[i,2],pattern = ",")==6){
      cuenta33=trimws(strsplit(info4[i,2],",")[[1]][3]) #proveedor
      cual_es=sapply(asegu_concept1[,2],grepl,cuenta33)
      if(sum(cual_es)!=0){
        cuenta= asegu_concept1[cual_es,1]
      } else {cual_es=sapply(asegu_concept2[,2],grepl,cuenta33)
      if(sum(cual_es)==1){
        cuenta= asegu_concept2[cual_es,1]} else cuenta = trimws(strsplit(info4[i,2],",")[[1]][4])}
      info5[i,1]=cuenta
      info5[i,2]=strsplit(info4[i,2],",")[[1]][7] #concepto
    } else if(str_count(info4[i,2],pattern = ",")==7){
      
      cuenta33=trimws(strsplit(info4[i,2],",")[[1]][3]) #proveedor
      cual_es=sapply(asegu_concept1[,2],grepl,cuenta33)
      if(sum(cual_es)!=0){
        cuenta= asegu_concept1[cual_es,1]
      } else {cual_es=sapply(asegu_concept2[,2],grepl,cuenta33)
      if(sum(cual_es)==1){
        cuenta= asegu_concept2[cual_es,1]} else cuenta = trimws(strsplit(info4[i,2],",")[[1]][4])}
      info5[i,1]=cuenta
      info5[i,2]=strsplit(info4[i,2],",")[[1]][6] #concepto
    }else if(grepl(info4[i,2],pattern="PAGO SERVICIO",ignore.case = T)){
      mov=strsplit(info4[i,2],"-")[[1]]
      info5[i,1]=mov[grepl("PAGO SERVICIO",mov,ignore.case = T)]
      if(length(mov)>=3){
        dd=!grepl("PAGO SERVICIO",mov,ignore.case = T)
        for(s in 2:sum(dd)){
          if(all(!grepl("PAGO SERVICIO",mov,ignore.case = T)[c(1:s)])){
            mov[1]=paste(mov[1],mov[s],sep = "-")}
        }
      }
      info5[i,2]=str_remove(mov[1],"TRA")
    }else if(any(sapply(asegu_concept[,2],grepl,info4[i,2]))){
      cual_es=sapply(asegu_concept1[,2],grepl,info4[i,2])
      if(sum(cual_es)!=0){
        cuenta= asegu_concept1[cual_es,1]
      } else {cual_es=sapply(asegu_concept2[,2],grepl,info4[i,2])
      if(sum(cual_es)==1){
        cuenta= asegu_concept2[cual_es,1]} else cuenta = info4[i,2]}
      info5[i,1]=cuenta
      concepto=info4[i,2]
      concepto=gsub(".*-","",concepto)
      info5[i,2]=concepto #str_remove(concepto,"TRA")
    }else if(nchar(info4[i,2])<=80){
      if(substr(info4[i,2],1,3)=="INT"){
        div=strsplit(info4[i,2],"-")[[1]]
        info5[i,1]=div[2]
        info5[i,2]=div[1]
      }else if(substr(trimws(info4[i,2]),1,3)=="TRA"){
        div=strsplit(info4[i,2],"-")[[1]]
        div[1]=str_remove(div[1],"TRA")
        if(grepl(info4[i,2],pattern = "TRASPASO A CUENTA",ignore.case = T)){
          div[2]=strsplit(info4[i,2],"-")[[1]][2]
        }else  div[2]=str_remove(info4[i,2],"TRA")
        info5[i,1]=div[1]
        info5[i,2]=div[2]
      }
    } 
    
  } ## cerrar cargos
  if(is.na(info4[i,4])){ #Abonos
    if(grepl(info4[i,2],pattern = "CHEQUE",ignore.case = T)|grepl(info4[i,2],pattern = "Cheque",ignore.case = T)){
      cheque= gsub("\\-.*","",info4[i,2])
      cheque=str_replace(cheque,"DOC","CHEQUE")
      info5[i,1]="Preguntar Proveedor/Compañia"
      info5[i,2]=cheque
    }else if(str_count(info4[i,2],pattern = ",")==6){
      info5[i,1]=strsplit(info4[i,2],",")[[1]][4] #proveedor
      info5[i,2]=strsplit(info4[i,2],",")[[1]][5] #concepto
    }else if(str_count(info4[i,2],pattern = ",")==7){
      info5[i,1]=strsplit(info4[i,2],",")[[1]][4] #proveedor
      info5[i,2]=strsplit(info4[i,2],",")[[1]][6] #concepto
    }else if(any(sapply(asegu_concept[,2],grepl,info4[i,2]))){  
      cual_es=sapply(asegu_concept1[,2],grepl,info4[i,2])
      if(sum(cual_es)!=0){
        cuenta= asegu_concept1[cual_es,1]
      } else {cual_es=sapply(asegu_concept2[,2],grepl,info4[i,2])
      if(sum(cual_es)==1){
        cuenta= asegu_concept2[cual_es,1]} else cuenta = info4[i,2]}
      info5[i,1]=cuenta
      concepto=info4[i,2]
      concepto=gsub(".*-","",concepto)
      info5[i,2]=concepto
    }else if(grepl("PAGO ELECTRONICO",info4[i,2])){
      info5[i,1]="BANORTE"
      info5[i,2]=info4[i,2]
      
    }else if(nchar(info4[i,2])<=80){
      if(substr(info4[i,2],1,3)=="INT"){
        div=strsplit(info4[i,2],"-")[[1]]
        info5[i,1]=div[2]
        info5[i,2]=div[1]
      } else if(grepl(info4[i,2],pattern="PAGO SERVICIO",ignore.case = T)){
        mov=strsplit(info4[i,2],"-")[[1]]
        info5[i,1]=mov[grepl("PAGO SERVICIO",mov,ignore.case = T)]
        if(length(mov)>=3){
          dd=!grepl("PAGO SERVICIO",mov,ignore.case = T)
          for(s in 2:sum(dd)){
            if(all(!grepl("PAGO SERVICIO",mov,ignore.case = T)[c(1:s)])){
              mov[1]=paste(mov[1],mov[s],sep = "-")}
          }
        }
        info5[i,2]=str_remove(mov[1],"TRA")
      }else if(substr(trimws(info4[i,2]),1,3)=="TRA"){
        div=strsplit(info4[i,2],"-")[[1]]
        div[1]=str_remove(div[1],"TRA")
        if(grepl(info4[i,2],pattern = "TRASPASO A CUENTA",ignore.case = T)){
          div[2]=strsplit(info4[i,2],"-")[[1]][2]
        }else  div[2]=str_remove(info4[i,2],"TRA")
        info5[i,1]=div[1]
        info5[i,2]=div[2]
      }
    }
    
    
  }
  print(i)
}

info6=cbind(info4[,1],info5,info4[,-c(1,2)])
colnames(info6)=c("Fecha","Proveedor/Compañia","Concepto","SaldoInicial","Cargo","Abono","SaldoFinal")
info6=trimws(info6)

info7=info6
## poner claves ingreso
clave_ingreso=read_excel("C:\\Users\\52811\\Desktop\\RProjects\\XXXXX\\DesarrolloFuentes\\BaseDatos\\ZZZZZData.xlsx",sheet = "ClaveIngreso")
clave_ingreso=as.data.frame(clave_ingreso)
clave_ingreso=clave_ingreso[!is.na(clave_ingreso[,1]),]
info_clave_ing_egr=matrix(ncol=2,nrow=nrow(info7))
for(ci in 1:nrow(clave_ingreso)){
  clave_nombre=grepl(clave_ingreso[ci,1],info7[,2],ignore.case = T)
  if(sum(na.omit(clave_nombre))==0)clave_nombre=grepl(clave_ingreso[ci,1],info7[,3],ignore.case = T)
  if(sum(clave_nombre)!=0){
    info7[clave_nombre,2]=clave_ingreso[ci,2]
    info_clave_ing_egr[clave_nombre,1]=clave_ingreso[ci,3]
  }}

for(ci in 1:nrow(info6)){
  if(!is.na(info7[ci,6]) && is.na(info_clave_ing_egr[ci,1])) info_clave_ing_egr[ci,1]="DatoManual"
}

clave_egreso=as.data.frame(claveEGR)
clave_egreso=clave_egreso[clave_egreso[,1]!="-",]
clave_egreso=clave_egreso[clave_egreso[,2]!="-",]
clave_egreso[,2]=round(as.numeric(clave_egreso[,2]),2)
clave_egreso[,3]=round(as.numeric(clave_egreso[,2]),2)

for(ci in 1:nrow(clave_egreso)){
  if(clave_egreso[ci,1]=="PAGO DE IMSS E INFONAVIT") clave_nombre=info7[,3]=="PAGO DE IMSS E INFONAVIT" else if(grepl("PAGO REFERENCIADO",clave_egreso[ci,1],ignore.case = T)){
    clave_nombre=grepl(clave_egreso[ci,1],info7[,2],ignore.case = T)} else {
      clave_nombre=grepl(clave_egreso[ci,1],info7[,3],ignore.case = T)}
  if(sum(clave_nombre)!=0){
    info_clave_ing_egr[clave_nombre,2]=clave_egreso[ci,3]
  }}

for(ci in 1:nrow(clave_egreso)){
  if(sum(grepl(clave_egreso[ci,1],info7[,2],ignore.case = T)!=0)) info_clave_ing_egr[grepl(clave_egreso[ci,1],info7[,2],ignore.case = T),2]=clave_egreso[ci,3]
}

#### OVERRIDE
#info_clave_ing_egr[(grepl("COMISION FEDERAL DE",info7[,2])),2]=4.02
override_cfe=(grepl("COMISION FEDERAL DE",info7[,2]))
info7[override_cfe,2]="CFE"
info_clave_ing_egr[override_cfe,2]=clave_egreso[grepl("CFE",clave_egreso[,1]),2]

#######



for(ci in 1:nrow(info6)){
  if(!is.na(info7[ci,5]) && is.na(info_clave_ing_egr[ci,2])) info_clave_ing_egr[ci,2]="DatoManual"
  if(!is.na(info7[ci,5]) && (info_clave_ing_egr[ci,2])=="-") info_clave_ing_egr[ci,2]="DatoManual"
}



colnames(info_clave_ing_egr)=c("ClaveING","ClaveEGR")

for(ci in 1:nrow(info6)){
  if(!is.na(info7[ci,6]) && is.na(info_clave_ing_egr[ci,1])) info_clave_ing_egr[ci,1]="DatoManual"
}

##RFC Dup
if(sum(grepl("RFC Dup",RFC_duplicado))!=0){
  info_clave_ing_egr[grepl("RFC Dup",RFC_duplicado),2]="RFCDup"
}
YYYYY=cbind(info7[,1],info_clave_ing_egr,info7[,c(2:ncol(info7))])

mes33=c("ENE","JAN","FEB","MAR","ABR","APR","MAY","JUN","JUL","AGO","AUG","SEP","OCT","NOV","DIC","DEC")
mes_num33=c("01","01","02","03","04","04","05","06","07","08","08","09","10","11","12","12")

meses11=cbind(mes33,mes_num33)

que_mes=strsplit(matrix(YYYYY[1,1]),"/")[[1]][2]

for(i in 1:nrow(meses11)){
  if(grepl(que_mes,meses11[i,1],ignore.case = TRUE)) mes_num22=meses11[i,2]
}

YYYYY[,1]=sub(que_mes,mes_num22,YYYYY[,1])



colnames(YYYYY)=c("Fecha","ClaveING","ClaveEGR","Proveedor/Compañia","Concepto","SaldoInicial","Cargo","Abono","SaldoFinal")
write.xlsx(YYYYY,"C:\\Users\\52811\\Desktop\\RProjects\\XXXXX\\DesarrolloFuentes\\Resultado\\Prueba.xlsx",colWidths="auto")
shell.exec("C:\\Users\\52811\\Desktop\\RProjects\\XXXXX\\DesarrolloFuentes\\Resultado\\Prueba.xlsx")