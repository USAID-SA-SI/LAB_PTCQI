# Title: LAB_PCTQI MER Reporting Script
# Author: C. Trapence
# Purpose: Automating the process of Reporting!
# Updated:Rosaline & Clement
# Date:2022-10-20
# Modified:2023-10-30

#Load Required libraries
library(tidyverse)
library(readxl)
library(lubridate)
library(readr)
library(excel.link)
library(openxlsx)
library(data.table)
library(here)

current_quarter<-"FY23Q4"

#Lab_PTQI

#Reference :2022-10-06_Lab PTQI Reporting meeting .
# Loading data elements from data sets,Element and Combos
# Use cumulative MER Partner reported HTS numbers to report on test volume 
# USAID prepare their own import file
#We will replace this genie file once Partner have completed manual entry of the Non-Tier HTS numbers.

genie<-list.files(here("Data/genie"),pattern="SITE")

DATIM<-read.delim2(here("data/genie",genie))

#Data_elements2<-read.csv("C:/Users/rpineteh/Documents/data/Data sets, elements and combos paramaterized.csv")

DE<-list.files(here("data"),pattern = "elements")

70
#genie<-read.delim2("C:/Users/rpineteh/Documents/data/Genie-SITE_IM-South Africa-Daily-2023-10-30.txt")

Data_elements3<-read.csv(here("Data", DE))

MC<- list.files(here("data"),pattern= "Mechanism")

mechanisms<- read.csv(here("Data", MC)) %>% filter(ou=="South Africa") %>% rename(AttributeOptionCombo=uid,mech_code=code) %>% mutate(mech_code=as.character(mech_code))


LAB_PTQI<-DATIM %>% filter(indicator=="HTS_TST",fiscal_year==2023 ,funding_agency=="USAID",standardizeddisaggregate %like% "Total Numerator")
write.xlsx(LAB_PTQI,"Data_view4.xlsx")


LAB_PTQIv1<-LAB_PTQI %>%   select(facility,orgunituid,psnu,psnuuid,cumulative,mech_code,mech_name,categoryoptioncomboname,qtr1,qtr2,qtr3,qtr4,categoryoptioncombouid,dataelementuid) %>% 
  group_by(facility,orgunituid,psnu,psnuuid,mech_code,mech_name,categoryoptioncomboname,categoryoptioncombouid,dataelementuid)  %>%  mutate(cumulative=as.integer(cumulative),qtr1=as.integer(qtr1),qtr2=as.integer(qtr2),qtr3=as.integer(qtr3),qtr4=as.integer(qtr4)) %>% 
  summarise_at(vars(qtr1,qtr2,qtr3,qtr4,cumulative),sum,na.rm=TRUE) %>%  mutate(Value=qtr1+qtr2+qtr3+qtr4) %>% 
mutate(CategoryOptionCombo="LAB_PTCQI_N_NoApp_POCT_TestVolume",dataelementuid ="KMtAtCRNZl8",categoryoptioncombouid="oCr3aOvULR9",Period="2023Q3") %>%   rename(OrgUnit="orgunituid")



#Joining this file with the Data sets, elements and combos paramaterized to get the UID's 
#LAB_PTQIv2<-left_join(LAB_PTQIv1,Data_elements,by=c("dataelementuid","categoryoptioncombouid"))

LAB_PTQIv2.1<-LAB_PTQIv1 %>%  select(dataelementuid,Period,OrgUnit,mech_code,categoryoptioncombouid,Value) %>% filter(mech_code %in% c("70310","70290","70287","81902", "70301"))%>%  
  rename(Dataelement=dataelementuid,CategoryOptionCombo=categoryoptioncombouid)


# #Lining with the mechanism to get the  Mechanism Attribute Combo Option UIDs

LAB_PTQIv2.2<-left_join(LAB_PTQIv2.1,mechanisms,by="mech_code") %>%  select(Dataelement ,Period,OrgUnit,CategoryOptionCombo,AttributeOptionCombo, Value) 


write.xlsx(LAB_PTQIv2.2,"Data_view9.xlsx")

LAB_PTQIv2.3<-LAB_PTQIv2.2%>% group_by(Dataelement ,Period,OrgUnit,CategoryOptionCombo,AttributeOptionCombo) %>% summarise_at(vars(Value),sum,na.rm=TRUE) 

write.xlsx(LAB_PTQIv2.3,"Data_view10.xlsx")
# Export CSV in readiness for import

#shell("taskkill /im EXCEL.exe /f /t")
filename<-paste0(current_quarter,"_LAB_PTCQI_",Sys.Date(),".csv")

write.csv(LAB_PTQIv2.3,here("dataout",filename),row.names = FALSE)
