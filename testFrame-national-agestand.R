library(tidyverse)
library(haven)
library(openxlsx)
Sys.setlocale(locale = "persian")
library(srvyr)
library(survey)
library(knitr)
library(readr)
library(gdata)
library(ggplot2)
library(gridExtra)
library(maptools)
library(grid)
library(RColorBrewer)
library(colorspace)
library(foreign)
library(readstata13)
library(grDevices)
library(png)
library(xfun)

#---- Read the data and variable lists 
mainData <- read_dta("../data/steps_2020_7.11.2021_01_p3_final_4.1.2_exported.dta")
varlist = read.xlsx("../resource/table A variable list.xlsx", sheet = "mean")
var_mis <- varlist$varlist[which(varlist$mis_status==1)]
var_zero <- varlist$varlist[which(varlist$mis_status==0)]

# Some Cleanings
mainData <- mainData %>% mutate_at(.vars = all_of(var_mis) ,funs(na_if(., -555)))
mainData <- mainData %>% mutate_at(.vars = "a19" ,funs(na_if(., 2)))
mainData <- mainData %>% mutate_if(is.numeric ,funs(ifelse(.==-555, 0, .)))
mainData$W_Anthropometry[which(is.na(mainData$W_Anthropometry))] <- 0
mainData <- mainData %>% arrange(age_cat)


# ingredients of age-standardization later
pop=read_dta("../data/pop_province_80-92_single_age.dta")
pop$age=as.numeric(pop$age)
pop=pop[pop$year==1395 & pop$age>=18 ,]

pop$age_cat=NA
pop$age_cat[pop$age<25]=18
pop$age_cat[pop$age>=25 & pop$age<35]=25
pop$age_cat[pop$age>=35 & pop$age<45]=35
pop$age_cat[pop$age>=45 & pop$age<55]=45
pop$age_cat[pop$age>=55 & pop$age<65]=55
pop$age_cat[pop$age>=65 & pop$age<75]=65
pop$age_cat[pop$age>=75 ]=75

pop=aggregate(pop$pred2pop~pop$sex_name+pop$area_name+pop$age_cat,FUN = sum)
colnames(pop)=c("c1","area","age_cat","pop")
pop= pop %>% rename(pop18=pop) %>%  mutate(pop25=if_else(age_cat>=25,pop18,0))

p25=pop %>%group_by(age_cat) %>% summarise(p=sum(pop18)) %>% filter(age_cat>=25)
p25_s=pop %>%group_by(age_cat,c1) %>% summarise(p=sum(pop18)) %>% filter(age_cat>=25)
p25_a=pop %>%group_by(age_cat,area) %>% summarise(p=sum(pop18)) %>% filter(age_cat>=25)
p25_sa=pop %>%group_by(age_cat,c1,area) %>% summarise(p=sum(pop18)) %>% filter(age_cat>=25)

p18=pop %>%group_by(age_cat) %>% summarise(p=sum(pop18)) 
p18_s=pop %>%group_by(age_cat,c1) %>% summarise(p=sum(pop18)) 
p18_a=pop %>%group_by(age_cat,area) %>% summarise(p=sum(pop18)) 
p18_sa=pop %>%group_by(age_cat,c1,area) %>% summarise(p=sum(pop18)) 

#---- 
# new indexes
# adequate  Dairy
mainData$adqDairy=mainData$d9a %>% recode("1"=0,"2"=0,"3"=0,"4"=0,"5"=1,"6"=1)
#second hand
mainData$secHand=mainData$t17 | mainData$t18
# daily processed meat consumption
mainData$weeklySausage=mainData$d8n==3 | mainData$d8n==4 |mainData$d8n==5 | mainData$d8n==6
# 
mainData$dailyCigUser=mainData$t5ad>0
mainData$dailyCigUser[is.na(mainData$dailyCigUser)]=0

mainData$prevDailyTob=mainData$s5ad |mainData$s5dd|mainData$s5cd|mainData$s5ed|mainData$s5fd
mainData$prevDailyTob=as.numeric(mainData$prevDailyTob)
mainData$prevDailyTob[is.na(mainData$prevDailyTob)]=0

mainData$prevCigSmoker=mainData$s5a_count>0
mainData$prevCigSmoker[is.na(mainData$prevCigSmoker)]=0
mainData$prevDailyCigSmoker=mainData$s5ad>0
mainData$prevDailyCigSmoker[is.na(mainData$prevDailyCigSmoker)]=0
#---- salt
mainData$extSalt=mainData$salt_24>5 
mainData$goodSalt=mainData$salt_24<=5 


#----
#Creating main list of new variables
lst=c("salt_24","fruveg","adqDairy","extSalt","htn_ecare12080","t17","t18","goodSalt")
lst=c("secHand","weeklySausage","dailyCigUser","prevDailyTob")
lst=c("prevDailyCigSmoker","prevDailyTob")
#"pre_HTN",
#surveyData= mainData %>% select(age_cat, c1, i07, i20,area, WI_National, W_Questionnaire, all_of(lst))  %>% as_survey_design(strata= i07, weights = W_Questionnaire)
surveyData= mainData %>% select(age_cat, c1, i07, i20,area, WI_National, W_Questionnaire, all_of(lst))  %>% as_survey_design(strata= i07, weights = W_Questionnaire)

provs=mainData['i07'] %>% unique()

#svystandardize(by = ~age_cat+c1+area,over = ~i07,population =p18_sa$p )

combs=expand.grid(c("tot","joz"),c("tot","joz"),c("tot","joz"),c("tot","joz","stn"),c("tot","joz"))
combs=combs %>% filter(!((Var4=="stn" & Var3=="joz")|(Var4=="stn" &Var5=="joz")))
vec2=c("c1","area","i20","age_cat","WI_National") 
  resTot=data.frame()
  for (vr in lst) {
    res=data.frame()
    variable <- as.name(vr)
    cons <- varlist$x[which(varlist$varlist==vr)]
    desc1 <- varlist$X5[which(varlist$varlist==vr)]

    
    
    for (i in 1:nrow(combs)){
     # for (i in 16:21){          
      
              #vec=c(sex,place,edu,age,WI)
              vec=combs[i,]
              
              # calculations with age standardization
              if (vec$Var4=="stn"){
                vecJoz=vec2[vec=="joz"]
                vecTot=vec2[vec=="tot"] 
                # standard for age only
                if (length(vecJoz)==0){
                  b=  surveyData %>%  svystandardize(by = ~age_cat,over = ~i07,population =p18$p )
                  a =  b %>% group_by_at(c("i07",vecJoz))  %>% summarise(round(survey_mean(eval(parse(text=vr))*100, na.rm = T, vartype = "ci"),2))
                }
                
                #standardize for c1, area and both
                else if (length(vecJoz)==1){
                  if (vecJoz==c("c1")){
                    b=  surveyData %>%svystandardize(by = ~age_cat+c1,over = ~i07,population =p18_s$p )
                    a =  b %>% group_by_at(c("i07",vecJoz))  %>% summarise(round(survey_mean(eval(parse(text=vr))*100, na.rm = T, vartype = "ci"),2))  
                    
                  }
                  else if (vecJoz==c("area")){
                    b=  surveyData %>%svystandardize(by = ~age_cat+area,over = ~i07,population =p18_a$p )
                    a =  b %>% group_by_at(c("i07",vecJoz))  %>% summarise(round(survey_mean(eval(parse(text=vr))*100, na.rm = T, vartype = "ci"),2))  
                    
                  }
                  }
                  else if (length(vecJoz)==2){
                    b=  surveyData %>%svystandardize(by = ~age_cat+c1+area,over = ~i07,population =p18_sa$p )
                    a =   b %>% group_by_at(c("i07",vecJoz))  %>% summarise(round(survey_mean(eval(parse(text=vr))*100, na.rm = T, vartype = "ci"),2))  
                    
                  }  
                if (vr %in% c("salt_24")){
                a[c("coef","_low","_upp")]=a[c("coef","_low","_upp")]/100}
                if (!("age_cat" %in% names(a))){
                  a$age_cat="Age Standard"
                }
                }
             

     
    # normal calculations
    else{
      vecJoz=vec2[vec=="joz"]
      vecTot=vec2[vec=="tot"]
      
      # multiply by 100 for indexes that are prevalence and not value ex. diabetes prevalence vs daily salt
      if (vr %in% c("salt_24")){
        a = surveyData %>% group_by_at(c("i07",vecJoz))  %>% summarise(round(survey_mean(eval(parse(text=vr)), na.rm = T, vartype = "ci"),2))
      }else{
        a = surveyData %>% group_by_at(c("i07",vecJoz))  %>% summarise(round(survey_mean(eval(parse(text=vr))*100, na.rm = T, vartype = "ci"),2)) 
      }
    }
              
              
              if ("WI_National" %in% names(a)){
                a=a %>% mutate(WI_National=as.character(WI_National)) %>% drop_na(WI_National)}

    
              
              a= a %>% mutate_at(names(a),as.character)
              
              # add count in each group  showing number 
              a['id']=do.call(paste0,a[c("i07",vecJoz)])
              ns=mainData %>% group_by_at(c("i07",vecJoz)) %>% summarize(n=n(),.groups = "drop")
              ns['id']=do.call(paste0,ns[c("i07",vecJoz)])
              a=left_join(a,ns[,c("id","n")],by="id",suffix=c("","_n"))
              
              
              for( n in vecTot){
                a[n]="Total"}
              
              
              # create the first dataframe if not currently present o.w. bind rows to it 
              if (length(res)==0){
                res=a 
              }else
              {
                res=rbind(res,a) 
              }
              print(i)
              
              
          }
    
    res['index']=vr
    if (length(resTot)==0){
      resTot=res 
    }else
    {
      resTot=rbind(resTot,res) 
    }
    
    
    
  }

# rename columns and ready data for export
  resTot2=resTot
  resTot[resTot['_low']<0,'_low']=0
  
  resTot[['WI_National']]=recode(resTot[['WI_National']], "1"="Poorest(1)","5"="Richest(5)","Total"="Toatal")
  resTot[['c1']]=recode(resTot[['c1']],"0"="Female","1"="Male","Total"="Both sex")
  resTot[['age_cat']]=recode(resTot[['age_cat']],"Total"="All ages",.missing="Age Standard")
  resTot[['area']]=recode(resTot[['area']] ,"0"="Rural","1"="Urban","Total"="Both area")
  resTot[['i20']]=recode(resTot[['i20']] ,"1"="1-7","2"="7-11","3"="12+")
  
  resTot=  resTot %>%  drop_na(i20)
  
  resTot=resTot %>% rename("p"="coef" , "p_low"="_low",  "p_upp"="_upp") %>% mutate(Measure="Prevalence")
  resTot [resTot$index=="salt_24","Measure"]="Mean"
  resTot=resTot[c("Measure","index","c1","area","i07","age_cat","WI_National","i20","n","p","p_low","p_upp")]
  for (prov in 0:31){
  write.csv(resTot %>% filter(i07==prov),paste0('test-17-std-',prov,'.csv'))
  
}


