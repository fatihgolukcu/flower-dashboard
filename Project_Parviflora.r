### A) ENVIROMENT
## 1- LIBRARY IMPORT
library(tidyverse)
library(readxl)
library(stringr)
#---------------------# End A #---------------------#
## B) IMPORT DATA
## 1- Daffodils2020.xls
# 1-1 Function for reading all sheets of excel file 
read_excel_allsheets <- function(filename, tibble = TRUE) 
{
  sheets <- readxl::excel_sheets(filename)
  x <- lapply(sheets, function(X) readxl::read_excel(filename, sheet = X))
  if(!tibble) x <- lapply(x, as.data.frame)
  names(x) <- sheets
  x
}
# 1-2 Months = List of Daffodils Months Dataframe
Months=read_excel_allsheets('Daffodils2020.xls')
# 1-3 Creating some list for importing properly
Monthly_daffodil<-vector('list',length = length(Months))         # vector for monthly daffodil
Daffodils<-vector('list',length = length(Months))                # vector for Daf
Submitting_Location<-vector('list',length = length(Months))      # vector for Submitting Location
Totals<-vector('list',length = length(Months))                   # vector for Totals
Count<-vector('list',length = length(Months))                    # vector for Count
# 1-4 Loop for filling vectors just before created and, combine them in "Daffodils"
for(i in 1:length(Months)){
  Monthly_daffodil[[i]]=Months[[i]][Months[[i]]$`Parviflora Store Financial Report`=="Submitting Location:"|
                                      Months[[i]]$`Parviflora Store Financial Report`=="Totals"|
                                      Months[[i]]$`Parviflora Store Financial Report`=="Summary Totals",]
  Submitting_Location[[i]]<-Monthly_daffodil[[i]][which(Monthly_daffodil[[i]]$`Parviflora Store Financial Report`=="Submitting Location:"),3]
  Totals[[i]]<-Monthly_daffodil[[i]][which(Monthly_daffodil[[i]]$`Parviflora Store Financial Report`=="Totals"),2]                   
  Count[[i]]<-Monthly_daffodil[[i]][which(Monthly_daffodil[[i]]$`Parviflora Store Financial Report`=="Totals"),3]                   
  Daffodils[[i]]<-data.frame(cbind('Store ID'=Submitting_Location[[i]],'Totals'=Totals[[i]]),'Count'=Count[[i]])
  names(Daffodils[[i]])[1]<-"Store ID"
  names(Daffodils[[i]])[2]<-'DAFFODIL'
  names(Daffodils[[i]])[3]<-paste0('COUNT','_',3)
} 
Daffodils
#---------------------# End B 1 #---------------------#
## 2- Stores = Stores.xlsx
Stores <- read_excel("Stores.xlsx")
Stores <- Stores %>% add_row(`Store ID` = c(671,345), `Store Name` = c("Unkown Place","Parviflora"))
#---------------------# End B 2 #---------------------#
## 3- tbl = Monthly Csv files  
# 3-1 files = directory SELECTING all csv files in folder
files <- list.files(path = "/home/Will/Lecture/Data Processing Securıity/Project/", pattern = "*.csv", full.names = T)
# 3-2 tbl = list of monthly reports dataframe`s
tbl <- sapply(files, read_csv, simplify=FALSE)
#---------------------# End B #---------------------#
### C) Cleaning
## 1- Monthly Reports
# 1-1 ID correction
for (g in 1:length(tbl)){
  for (n in 1:nrow(tbl[[g]])){
    if (tbl[[g]]$`STORE #`[n]==5514870) {
      tbl[[g]]$`STORE #`[n]= 5487123340
    }else if (tbl[[g]]$`STORE #`[n]==6366583 | tbl[[g]]$`STORE #`[n]==6366584 | tbl[[g]]$`STORE #`[n]==6366582) {
      tbl[[g]]$`STORE #`[n]= 6366584590
    }else if (tbl[[g]]$`STORE #`[n]==5890646) {
      tbl[[g]]$`STORE #`[n]= 5487123110
    }else if (tbl[[g]]$`STORE #`[n]==5891382) {
      tbl[[g]]$`STORE #`[n]= 5487123260
    }else if (tbl[[g]]$`STORE #`[n]==5890958) {
      tbl[[g]]$`STORE #`[n]= 5487123240
    }else if (tbl[[g]]$`STORE #`[n]==5889308) {
      tbl[[g]]$`STORE #`[n]= 5487123360
    }else if (tbl[[g]]$`STORE #`[n]==5890909) {
      tbl[[g]]$`STORE #`[n]= 5487123200
    }else if (tbl[[g]]$`STORE #`[n]==5890933) {
      tbl[[g]]$`STORE #`[n]= 5487123150
    }else if (tbl[[g]]$`STORE #`[n]==6237566 | tbl[[g]]$`STORE #`[n]==6226380 | tbl[[g]]$`STORE #`[n]==6237241) {
      tbl[[g]]$`STORE #`[n]= 6237566510
    }else if (tbl[[g]]$`STORE #`[n]==6366593 | tbl[[g]]$`STORE #`[n]==6366592 | tbl[[g]]$`STORE #`[n]==6366594) {
      tbl[[g]]$`STORE #`[n]= 6366594540
    }else if (tbl[[g]]$`STORE #`[n]==6369492 | tbl[[g]]$`STORE #`[n]==6369489 | tbl[[g]]$`STORE #`[n]==6369490) {
      tbl[[g]]$`STORE #`[n]= 6369492550
    }else if (tbl[[g]]$`STORE #`[n]==6287467) {
      tbl[[g]]$`STORE #`[n]= 6287467530
    }else if (tbl[[g]]$`STORE #`[n]==6046893) {
      tbl[[g]]$`STORE #`[n]= 5487123500
    }else if (tbl[[g]]$`STORE #`[n]==5890887) {
      tbl[[g]]$`STORE #`[n]= 5890887270
    }else if (tbl[[g]]$`STORE #`[n]==5890646) {
      tbl[[g]]$`STORE #`[n]= 5890646110
    }else if (tbl[[g]]$`STORE #`[n]==6366598 | tbl[[g]]$`STORE #`[n]==6366597 | tbl[[g]]$`STORE #`[n]==6366599) {
      tbl[[g]]$`STORE #`[n]= 6366599580
    }else if (tbl[[g]]$`STORE #`[n]==6366121 | tbl[[g]]$`STORE #`[n]==6366123 ) {
      tbl[[g]]$`STORE #`[n]= 6366123570
    }else if (tbl[[g]]$`STORE #`[n]==6356299) {
      tbl[[g]]$`STORE #`[n]= 6356299345
    }
  }
} ### loop for ID correction
# 1-2 function for select last characters function
substrRight <- function(x, n){
  substr(x, nchar(x)-n+1, nchar(x))
}
# 1-3 adding daffodils to raw tbl
for (i in 1:length(tbl)){
  tbl[[i]] <- mutate(tbl[[i]], `Store ID`=as.numeric(substrRight(tbl[[i]]$`STORE #`,3)))
  tbl[[i]] <- tbl[[i]] %>% left_join(Stores, by="Store ID")
} # to add daffodil some adjustments needed.
for (i in 1:length(Daffodils)){
  Daffodils[[i]][,] <- sapply( Daffodils[[i]][,], as.numeric)
  tbl[[i]] <- tbl[[i]] %>% left_join(Daffodils[[i]], by="Store ID")
  tbl[[i]] <- tbl[[i]] %>% select(`STORE NAME`:`CARNATION`,`COUNT_3`=`COUNT_3.y`,`DAFFODIL`=`DAFFODIL.y`,`TOTAL`)
  tbl[[i]][is.na(tbl[[i]])] <- 0
  tbl[[i]] <- tbl[[i]] %>% mutate(`TOTAL`= (`AZALEA`+`BEGONIA`+`CARNATION`+`DAFFODIL`),`COUNT_TOTAL`=(`COUNT`+`COUNT_1`+`COUNT_2`+`COUNT_3`))
  tbl[[i]] <- tbl[[i]] %>% select(`STORE NAME`:`DAFFODIL`,`COUNT_TOTAL`,`TOTAL`)
} # ADDING DAFFODIL TO TBL
# 1-4 create ID_list file which is all cleaned data
ID<-function(x){
  ID_list<-vector('list',length = length(x))
  for (i in 1:length(x)){
    ID_list[[i]] <- mutate(x[[i]], `Store ID`=as.numeric(substrRight(x[[i]]$`STORE #`,3)))
    ID_list[[i]] <- ID_list[[i]] %>% left_join(Stores, by="Store ID")
    for (j in 1:nrow(ID_list[[i]])){ 
      if(is.na(ID_list[[i]]$`Store Name`[j])){
        ID_list[[i]]$`Store Name`[j]=ID_list[[i]]$`STORE NAME`[j]}
    }      
    ID_list[[i]] <-aggregate(ID_list[[i]][,3:12], by=list(`Store Name`=ID_list[[i]]$`Store Name`,`Store ID`=ID_list[[i]]$`Store ID`), FUN=sum)     
    
  }
  ID_list
} # Adjust All other function
ID_list=ID(tbl) # applying the function on tbl
for (i in 1:length(ID_list)) {
  ID_list[[i]] <-ID_list[[i]] %>% select(`Store Name`:`TOTAL`)
} # select neccessary columns

## 2- Creating Quarterly Report
# 2-1 Join All dataframes in list
Q1Store <- reduce(ID_list, left_join, by = "Store ID")
# 2-2 correct names
names(Q1Store)[which(str_detect(names(Q1Store),'\\.'))]=
  str_extract(names(Q1Store[which(str_detect(names(Q1Store),'\\.'))]),"^[^.]+(?=.)")


# 2-3 Sum corrected names
Sum_of_quarter <- cbind(Q1Store[1],'Store ID'=Q1Store[,!duplicated(colnames(Q1Store)) & 
                                                        !duplicated(colnames(Q1Store), fromLast = TRUE)],
                        sapply(unique(colnames(Q1Store)[duplicated(colnames(Q1Store))])[-1], function(x)
                          rowSums(Q1Store[,grepl(paste(x, "$", sep=""), colnames(Q1Store))],na.rm = T)))
# Sum_of_quarter <- Sum_of_quarter[!is.na(Sum_of_quarter$TOTAL),] 
row.names(Sum_of_quarter) <- 1:nrow(Sum_of_quarter)
row=nrow(Sum_of_quarter)+1

#### this loop if irregular store comes it wont work.
for (i in 1:ncol(Sum_of_quarter)){
  
  for(j in 1:3){
    if(names(Sum_of_quarter[i])==names(Daffodils[[3]][j])){
      Sum_of_quarter[row,i]<-Daffodils[[3]][which(Daffodils[[3]][1]==6563027671),j]
      Sum_of_quarter[row,'TOTAL']=sum(Sum_of_quarter[row,c('AZALEA',"BEGONIA","CARNATION","DAFFODIL")],na.rm=T)
      Sum_of_quarter[row,'COUNT_TOTAL']=sum(Sum_of_quarter[row,c('COUNT',"COUNT_1","COUNT_2","COUNT_3")],na.rm=T)
    }
  }
}
#---------------------# End C 2 #---------------------#

