# Title: LAB_PCTQI MER Reporting Script
# Author: C. Trapence
# Purpose: Automating the process of Reporting!
# Date:2022-10-20
# Modified:2023-09-08

#Load Required libraries
library(tidyverse)
library(readxl)
library(lubridate)
library(readr)
library(excel.link)
library(openxlsx)
library(data.table)

current_quarter<-"FY23Q4"

#Lab_PTQI

#Reference :2022-10-06_Lab PTQI Reporting meeting .
# Loading data elements from data sets,Element and Combos
# Use cumulative MER Partner reported HTS numbers to report on test volume 
# USAID prepare their own import file
#We will replace this genie file once Partner have completed manual entry of the Non-Tier HTS numbers.

genie<-list.files(here("Data/genie"),pattern="SITE")

DATIM<-read.delim2(here("data/genie",genie))


Data_elements<-read.csv("Data/Data sets, elements and combos paramaterized.csv")

LAB_PTQI<-DATIM %>% filter(sitetype=="Facility",indicator=="HTS_TST",fiscal_year==2023 ,funding_agency=="USAID",standardizeddisaggregate %like% "Total Numerator")

LAB_PTQIv1<-LAB_PTQI %>%   select(facility,orgunituid,psnu,psnuuid,cumulative,mech_code,mech_name,categoryoptioncomboname,qtr1,qtr2,qtr3,qtr4,categoryoptioncombouid,dataelementuid) %>%
  group_by(facility,orgunituid,psnu,psnuuid,mech_code,mech_name,categoryoptioncomboname,categoryoptioncombouid,dataelementuid) %>%  mutate(cumulative=as.integer(cumulative),qtr1=as.integer(qtr1),qtr2=as.integer(qtr2),qtr3=as.integer(qtr3),qtr4=as.integer(qtr4)) %>% 
  summarise_at(vars(qtr1,qtr2,qtr3,qtr4,cumulative),sum,na.rm=TRUE) %>%  mutate(Value=cumulative) %>% 
mutate(CategoryOptionCombo="LAB_PTCQI_N_NoApp_POCT_TestVolume",dataelementuid ="KMtAtCRNZl8",categoryoptioncombouid="oCr3aOvULR9",Period="2023Q3") %>% 
  
  rename(OrgUnit="orgunituid")

#Joining this file with the Data sets, elements and combos paramaterized to get the UID's 
LAB_PTQIv2<-left_join(LAB_PTQIv1,Data_elements,by=c("dataelementuid","categoryoptioncombouid"))

LAB_PTQIv2.1<-LAB_PTQIv2 %>%  select(dataelementuid,Period,OrgUnit,mech_code,categoryoptioncombouid,Value) %>% filter(mech_code %in% c("70310","70290","70287","81902"))%>%  
  rename(Dataelement=dataelementuid,CategoryOptionCombo=categoryoptioncombouid)

# #Lining with the mechanism to get the  Mechanism Attribute Combo Option UIDs

mechanisms<-read.csv("data/Mechanisms partners agencies OUS Start End.csv") %>% filter(ou=="South Africa") %>% rename(AttributeOptionCombo=uid,mech_code=code) %>% mutate(mech_code=as.character(mech_code))

LAB_PTQIv2.2<-left_join(LAB_PTQIv2.1,mechanisms,by="mech_code") %>%  select(Dataelement ,Period,OrgUnit,CategoryOptionCombo,AttributeOptionCombo, Value) 



LAB_PTQIv2.3<-LAB_PTQIv2.2%>% group_by(Dataelement ,Period,OrgUnit,CategoryOptionCombo,AttributeOptionCombo) %>% summarise_at(vars(Value),sum,na.rm=TRUE) 


# Export CSV in readiness for import

shell("taskkill /im EXCEL.exe /f /t")
filename<-paste(Sys.Date(),current_quarter,"LAB_PTCQI.csv")

write.csv(LAB_PTQIv2.3,filename,row.names = FALSE)
