library(psych)
library(fpp3)
#Read Files
Company <- read_excel("~/Downloads/Company%20Profile235444093/TRD_Co.xlsx")
Cash_Flow <- read_excel("~/Downloads/Cash%20Flow153735271/FI_T6.xlsx")
Income <- read_excel("~/Downloads/Income%20Statement153234253/FS_Comins.xlsx")
Stock_trading <- read_excel("~/Downloads/Monthly%20Stock%20Price%20%20Returns151737723/TRD_Mnth.xlsx")
Modified_balance <- read_excel("~/Downloads/Balance%20Sheet162201901/FS_Combas.xlsx")
#Change the columns' name
colnames(Income) <- c('code','date','type','profit','RD','name')
colnames(Cash_Flow) <- c('code','date','type','amor','name')
colnames(Modified_balance) <- c('code','date','type','otherasset','totalasset','name')
colnames(Stock_trading) <- c('code','date','day','close','TMV','returnwith','returnwithout','type')
#Construct quarterly data grougby the data with code and time
Q_Cashflow <- Cash_Flow %>% mutate(Time=yearquarter(date)) %>% select(-date) %>% group_by(code,Time) %>% summarise(amor=mean(amor,na.rm=TRUE))
Q_Cashflow <- Q_Cashflow %>% filter(startsWith(code,'00') == TRUE | startsWith(code,'600') == TRUE | startsWith(code,'601') == TRUE | startsWith(code,'603') == TRUE | startsWith(code,'605') == TRUE)
Q_Income <- Income %>% mutate(Time=yearquarter(date)) %>% select(-date) %>%  group_by(code,Time) %>% summarise(profit = mean(profit,na.rm=TRUE),RD = mean(RD,na.rm=TRUE))
Q_Income <- Q_Income %>% filter(startsWith(code,'00') == TRUE | startsWith(code,'600') == TRUE | startsWith(code,'601') == TRUE | startsWith(code,'603') == TRUE | startsWith(code,'605') == TRUE)
Q_Balance <- Modified_balance %>% mutate(Time=yearquarter(date)) %>% select(-date) %>%  group_by(code,Time) %>% summarise(otherasset = mean(otherasset,na.rm=TRUE),totalasset = mean(totalasset,na.rm=TRUE))
Q_Balance <- Q_Balance %>% filter(startsWith(code,'00') == TRUE | startsWith(code,'600') == TRUE | startsWith(code,'601') == TRUE | startsWith(code,'603') == TRUE | startsWith(code,'605') == TRUE)
Q_Stock <- Stock_trading %>% mutate(Time=yearquarter(date)) %>% select(-date,-day,-returnwith)
Q_Stock <- Q_Stock %>% mutate(Q_close = 1) %>% mutate(Q_return = 1) 
Q_Stock[is.na(Q_Stock)] <-0
Q_Stock <- Q_Stock %>% filter(startsWith(code,'00') == TRUE | startsWith(code,'600') == TRUE | startsWith(code,'601') == TRUE | startsWith(code,'603') == TRUE | startsWith(code,'605') == TRUE)
#Change the monthly stock return and the closing price to the quarterly data
i=578934
while (i >0){
  ret = 1
  for (j in i:0){
    if (Q_Stock[j,6] == Q_Stock[j-1,6]){
      ret = ret * (1+Q_Stock[j,4])
      Q_Stock[j,7] = Q_Stock[i,2]
      j=j-1
    }else{
      ret = ret * (1+Q_Stock[j,4])
      Q_Stock[j,7] = Q_Stock[i,2]
      break
    }
  }
  Q_Stock[j:i,8] = ret-1
  i=j-1
  j=i
}
Q_Stock2 <- Q_Stock %>% group_by(code,Time) %>% summarise(Close=mean(Q_close),return=mean(Q_return),type=mean(type),code=max(code),TMV=mean(TMV)*1000)

for (i in 1:187484 ){
  if (Q_Balance[i,3]=='NaN'){
    Q_Balance[i,3]=0
  }
}
for (i in 1:187875 ){
  if (Q_Income[i,3]=='NaN'){
    Q_Income[i,3]=0
  }
  if (Q_Income[i,4]=='NaN'){
    Q_Income[i,4]=0
  }
}
for (i in 1:190617 ){
  if (Q_Cashflow[i,3]=='NaN'){
    Q_Cashflow[i,3]=0
  }
}
cash_balance <- left_join(Q_Cashflow,Q_Balance)
cash_income <- left_join(cash_balance,Q_Income)
joinstock <- left_join(cash_income,Q_Stock2)
company <- Company %>% select(`Stock Code`,`Establishment Date`,`Market Type`)
colnames(company) <- c('code','esdate','type')
total <- left_join(company,joinstock,by='code')
total <- total %>% filter(startsWith(code,'00') == TRUE | startsWith(code,'600') == TRUE | startsWith(code,'601') == TRUE | startsWith(code,'603') == TRUE | startsWith(code,'605') == TRUE)
#GEM board
Q_Cashflow <- Cash_Flow %>% mutate(Time=yearquarter(date)) %>% select(-date) %>% group_by(code,Time) %>% summarise(amor=mean(amor,na.rm=TRUE))
Q_Cashflow <- Q_Cashflow %>% filter(startsWith(code,'002') == TRUE)
Q_Income <- Income %>% mutate(Time=yearquarter(date)) %>% select(-date) %>%  group_by(code,Time) %>% summarise(profit = mean(profit,na.rm=TRUE),RD = mean(RD,na.rm=TRUE))
Q_Income <- Q_Income %>% filter(startsWith(code,'002') == TRUE )
Q_Balance <- Modified_balance %>% mutate(Time=yearquarter(date)) %>% select(-date) %>%  group_by(code,Time) %>% summarise(otherasset = mean(otherasset,na.rm=TRUE),totalasset = mean(totalasset,na.rm=TRUE))
Q_Balance <- Q_Balance %>% filter(startsWith(code,'002') == TRUE)
Q_Stock <- Stock_trading %>% mutate(Time=yearquarter(date)) %>% select(-date,-day,-returnwith)
Q_Stock <- Q_Stock %>% mutate(Q_close = 1) %>% mutate(Q_return = 1) 
Q_Stock[is.na(Q_Stock)] <-0
Q_Stock <- Q_Stock %>% filter(startsWith(code,'002') == TRUE)

i=131815
while (i >0){
  ret = 1
  for (j in i:0){
    if (Q_Stock[j,6] == Q_Stock[j-1,6]){
      ret = ret * (1+Q_Stock[j,4])
      Q_Stock[j,7] = Q_Stock[i,2]
      j=j-1
    }else{
      ret = ret * (1+Q_Stock[j,4])
      Q_Stock[j,7] = Q_Stock[i,2]
      break
    }
  }
  Q_Stock[j:i,8] = ret-1
  i=j-1
  j=i
}

Q_Stock3 <- Q_Stock %>% group_by(code,Time) %>% summarise(Close=mean(Q_close),return=mean(Q_return),type=mean(type),code=max(code),TMV=mean(TMV)*1000)

for (i in 1:44994 ){
  if (Q_Balance[i,3]=='NaN'){
    Q_Balance[i,3]=0
  }
}
for (i in 1:45233 ){
  if (Q_Income[i,3]=='NaN'){
    Q_Income[i,3]=0
  }
  if (Q_Income[i,4]=='NaN'){
    Q_Income[i,4]=0
  }
}
for (i in 1:48046 ){
  if (Q_Cashflow[i,3]=='NaN'){
    Q_Cashflow[i,3]=0
  }
}
cash_balance2 <- left_join(Q_Cashflow,Q_Balance)
cash_income2 <- left_join(cash_balance2,Q_Income)
joinstock2 <- left_join(cash_income2,Q_Stock3)
total2 <- left_join(company,joinstock2,by='code')
total2 <- total2 %>% filter(startsWith(code,'002') == TRUE)


#SME board
Q_Cashflow <- Cash_Flow %>% mutate(Time=yearquarter(date)) %>% select(-date) %>% group_by(code,Time) %>% summarise(amor=mean(amor,na.rm=TRUE))
Q_Cashflow <- Q_Cashflow %>% filter(startsWith(code,'300') == TRUE)
Q_Income <- Income %>% mutate(Time=yearquarter(date)) %>% select(-date) %>%  group_by(code,Time) %>% summarise(profit = mean(profit,na.rm=TRUE),RD = mean(RD,na.rm=TRUE))
Q_Income <- Q_Income %>% filter(startsWith(code,'300') == TRUE )
Q_Balance <- Modified_balance %>% mutate(Time=yearquarter(date)) %>% select(-date) %>%  group_by(code,Time) %>% summarise(otherasset = mean(otherasset,na.rm=TRUE),totalasset = mean(totalasset,na.rm=TRUE))
Q_Balance <- Q_Balance %>% filter(startsWith(code,'300') == TRUE)
Q_Stock <- Stock_trading %>% mutate(Time=yearquarter(date)) %>% select(-date,-day,-returnwith)
Q_Stock <- Q_Stock %>% mutate(Q_close = 1) %>% mutate(Q_return = 1) 
Q_Stock[is.na(Q_Stock)] <-0
Q_Stock <- Q_Stock %>% filter(startsWith(code,'300') == TRUE)

total[is.na(total)] <- 0
total2[is.na(total2)] <- 0
total3[is.na(total3)] <- 0
i=87371
while (i >0){
  ret = 1
  for (j in i:0){
    if (Q_Stock[j,6] == Q_Stock[j-1,6]){
      ret = ret * (1+Q_Stock[j,4])
      Q_Stock[j,7] = Q_Stock[i,2]
      j=j-1
    }else{
      ret = ret * (1+Q_Stock[j,4])
      Q_Stock[j,7] = Q_Stock[i,2]
      break
    }
  }
  Q_Stock[j:i,8] = ret-1
  i=j-1
  j=i
}
Q_Stock4 <- Q_Stock %>% group_by(code,Time) %>% summarise(Close=mean(Q_close),return=mean(Q_return),type=mean(type),code=max(code),TMV=mean(TMV)*1000)

for (i in 1:29776){
  if (Q_Balance[i,3]=='NaN'){
    Q_Balance[i,3]=0
  }
}
for (i in 1:29767){
  if (Q_Income[i,3]=='NaN'){
    Q_Income[i,3]=0
  }
  if (Q_Income[i,4]=='NaN'){
    Q_Income[i,4]=0
  }
}
for (i in 1:32985){
  if (Q_Cashflow[i,3]=='NaN'){
    Q_Cashflow[i,3]=0
  }
}
cash_balance3 <- left_join(Q_Cashflow,Q_Balance)
cash_income3 <- left_join(cash_balance3,Q_Income)
joinstock3 <- left_join(cash_income3,Q_Stock4)
total3 <- left_join(company,joinstock3,by='code')
total3 <- total3 %>% filter(startsWith(code,'300') == TRUE)
total <- total %>% mutate(es = yearquarter(esdate)) %>% select(-age) %>% mutate(age = (Time-es)/4) %>% select(-type.y,-esdate)
total2 <- total2 %>% mutate(es = yearquarter(esdate))  %>% mutate(age = (Time-es)/4) %>% select(-type.y,-esdate)
total3 <- total3 %>% mutate(es = yearquarter(esdate)) %>% mutate(age = (Time-es)/4) %>% select(-type.y,-esdate)
rtotal <- rtotal %>% mutate(pe=TMV/profit,pb=TMV/(totalasset-otherasset-amor))
rtotal2 <- rtotal2 %>% mutate(pe=TMV/profit,pb=TMV/(totalasset-otherasset-amor))
rtotal3 <- rtotal3 %>% mutate(pe=TMV/profit,pb=TMV/(totalasset-otherasset-amor))

rtotal <- rtotal %>% filter(yearquarter(Time)<=yearquarter('2022Q4'))
rtotal2<-rtotal2 %>% filter(yearquarter(Time)<=yearquarter('2022Q4'))
rtotal3<-rtotal3 %>% filter(yearquarter(Time)<=yearquarter('2022Q4'))
des1 <- rtotal %>% select(return,TMV,age,totalasset,RD,income,pe,pb)
des2 <- rtotal2 %>% select(return,TMV,age,totalasset,RD,income,pe,pb)
des3 <- rtotal3 %>% select(return,TMV,age,totalasset,RD,income,pe,pb)
des1[is.infinite(des1$pe),7]<-0
des1[is.infinite(des1$pb),8]<-0
des2[is.infinite(des2$pe),7]<-0
des2[is.infinite(des2$pb),8]<-0
des3[is.infinite(des3$pe),7]<-0
des3[is.infinite(des3$pb),8]<-0
des1[is.na(des1)]<-0
des2[is.na(des2)]<-0
des3[is.na(des3)]<-0
describe(des1)
#to summarise the statistics
summar <- function(x, na.omit=FALSE){
  Median <- median(x,na.rm=TRUE)
  Mean <- mean(x,na.rm=TRUE)
  q1 <- as.numeric(quantile(x,0.25,na.rm=TRUE))
  q3 <- as.numeric(quantile(x,0.75,na.rm=TRUE))
  Min <- min(x,na.rm=TRUE)
  Max <- max(x,na.rm=TRUE)
  n <- length(x)
  s <- sd(x,na.rm=TRUE)
  re <- c(Min=Min,Max=Max,Count=n,Median=Median,Mean=Mean,p25 = q1,p75=q3,SD=s)
  par(mfrow=c(1,3))
  par(pin=c(3,5))
  return(re)
}
summar(des1$return)
summar(des1$TMV)
summar(des1$age)
summar(des1$totalasset)
summar(des1$RD)
summar(des1$income)
summar(des1$pe)
summar(des1$pb)

summar(des2$return)
summar(des2$TMV)
summar(des2$age)
summar(des2$totalasset)
summar(des2$RD)
summar(des2$income)
summar(des2$pe)
summar(des2$pb)

summar(des3$return)
summar(des3$TMV)
summar(des3$age)
summar(des3$totalasset)
summar(des3$RD)
summar(des3$income)
summar(des3$pe)
summar(des3$pb)

d1 <- rtotal %>% group_by(Time) %>% summarise(pe=median(pe,na.rm =TRUE)/100)
d2 <- rtotal2 %>% group_by(Time) %>% summarise(pe=median(pe,na.rm =TRUE)/100)
d3 <- rtotal3 %>% group_by(Time) %>% summarise(pe=median(pe,na.rm =TRUE)/100)

d1[is.na(d1)] <-0
d2[is.na(d2)] <-0
d3[is.na(d3)] <-0
#Draw the graph
ggplot()+geom_line(data=d1,aes(x=Time,y=pe),color='red')+geom_line(data=d2,aes(x=Time,y=pe),color='blue')+geom_line(data=d3,aes(x=Time,y=pe),color='green')
