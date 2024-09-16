# Title: LAB_PCTQI MER Reporting Script
# Author: C. Trapence
# Purpose: Automating the process of Reporting!
# Updated:Rosaline & Clement
# Date:2022-10-20
# Modified:2023-10-30
# Modified_v2 2024-09-12

#Load Required libraries
if(!require(pacman)) install.packages("pacman")
pacman::p_load(glitr, gophr, extrafont, scales,data.table,tierdrop,lubridate,tidyverse,here, tidytext, anytime, here, patchwork, ggtext, glue, readxl, googlesheets4,glamr, tidyverse, maditr, openxlsx,readr,stringr,sqldf,janitor,excel.link, readr)

#load global variables
current_quarter<-"FY24Q4"

load_secrets()

#Lab_PTQI

#Reference :2022-10-06_Lab PTQI Reporting meeting .
# Loading data elements from data sets,Element and Combos
# Use cumulative MER Partner reported HTS numbers to report on test volume 
# USAID prepare their own import file
#We will replace this genie file once Partner have completed manual entry of the Non-Tier HTS numbers.

Genie<-list.files(here("Data/Genie"),pattern="SITE")

#DATIM<-read.delim2(here("data/genie",genie))

DATIM<-read_psd(here("Data/Genie",Genie))


DE<-list.files(here("data"),pattern = "elements")


Data_elements<-read.csv(here("Data", DE))

MC<- list.files(here("data"),pattern= "Mechanism")

mechanisms<- read.csv(here("Data", MC)) %>% filter(ou=="South Africa") %>% rename(AttributeOptionCombo=uid,mech_code=code) %>% mutate(mech_code=as.character(mech_code))


LAB_PTCQI<-DATIM %>% filter(indicator=="HTS_TST",fiscal_year==2024 ,funding_agency=="USAID",standardizeddisaggregate %like% "Total Numerator")

#write.xlsx(LAB_PTCQI,"Data_view.xlsx")

LAB_PTCQIv1 <- LAB_PTCQI %>%   select(facility,orgunituid,psnu,psnuuid,cumulative,mech_code,mech_name,qtr1,qtr2,qtr3,qtr4) %>% 
  group_by(facility,orgunituid,psnu,psnuuid,mech_code,mech_name)  %>%  mutate(cumulative=as.integer(cumulative),qtr1=as.integer(qtr1),qtr2=as.integer(qtr2),qtr3=as.integer(qtr3),qtr4=as.integer(qtr4)) %>% 
  summarise_at(vars(qtr1,qtr2,qtr3,qtr4,cumulative),sum,na.rm=TRUE) %>%  mutate(Value=qtr1+qtr2+qtr3+qtr4) %>% 
  mutate(CategoryOptionCombo="LAB_PTCQI_N_NoApp_POCT_TestVolume",dataelementuid ="KMtAtCRNZl8",categoryoptioncombouid="oCr3aOvULR9",Period="2024Q3") %>%   rename(OrgUnit="orgunituid")

#Joining this file with the Data sets, elements and combos paramaterized to get the UID's 
LAB_PTCQIv2<-left_join(LAB_PTCQIv1,Data_elements,by=c("dataelementuid","categoryoptioncombouid")) %>% 
  mutate(mech_code= if_else(mech_code %in% c("81902", "87576"), "87576", mech_code )) %>% 
  mutate(mech_code =if_else(psnu %in% c("lp Capricorn District Municipality", "lp Mopani District Municipality"), "87577", mech_code )) %>% 
mutate(mech_code =if_else(psnu %in% c("kz King Cetshwayo District Municipality","kz Harry Gwala District Municipality", "kz Ugu District Municipality"), "87575", mech_code ))

LAB_PTCQIv2.1<-LAB_PTCQIv1 %>%  select(dataelementuid,Period,OrgUnit,mech_code,categoryoptioncombouid,Value) %>% filter(mech_code %in% c("70310","70290","70287","87575", "87576", "87577","70301"))%>%  
  rename(Dataelement=dataelementuid,CategoryOptionCombo=categoryoptioncombouid)


# #Lining with the mechanism to get the  Mechanism Attribute Combo Option UIDs

LAB_PTCQIv2.2<-left_join(LAB_PTCQIv2.1,mechanisms,by="mech_code") %>%  select(Dataelement ,Period,OrgUnit,CategoryOptionCombo,AttributeOptionCombo, Value) 


#write.xlsx(LAB_PTQIv2.2,"Data_view9.xlsx")

LAB_PTQQIv2.2%>% group_by(Dataelement ,Period,OrgUnit,CategoryOptionCombo,AttributeOptionCombo) %>% summarise_at(vars(Value),sum,na.rm=TRUE) 

#write.xlsx(LAB_PTQIv2.3,"Data_view10.xlsx")
# Export CSV in readiness for import

#shell("taskkill /im EXCEL.exe /f /t")

# Grab relevant HTS data elements from the Data elements reference file and store in "cols" object

cols <-  c("HTS_INDEX (N, DSD, Index/Age Aggregated/Sex/Contacts): Number of contacts",
"HTS_INDEX (N, DSD, Index/Age/Sex/CasesAccepted): Number of index cases accepted",
"HTS_INDEX (N, DSD, Index/Age/Sex/CasesOffered): Number of index cases offered",
"HTS_INDEX (N, DSD, Index/Age/Sex/Result): HTS Result",
"HTS_TST (N, DSD, OtherPITC/Age/Sex/Result): HTS received results",
"HTS_TST (N, DSD, PMTCT PostANC1 Pregnant-L&D/Age/Sex/Result): HTS received results",
"PMTCT_STAT (D, DSD, Age/Sex): New ANC clients",
"HTS_INDEX (N, DSD, IndexMod/Age Aggregated/Sex/Contacts): Number of contacts",
"HTS_INDEX (N, DSD, IndexMod/Age/Sex/CasesAccepted): Number of index cases accepted",
"HTS_INDEX (N, DSD, IndexMod/Age/Sex/CasesOffered): Number of index cases offered",
"HTS_INDEX (N, DSD, IndexMod/Age/Sex/Result): HTS Result",
"HTS_TST (N, DSD, MobileMod/Age/Sex/Result): HTS received results",
"HTS_TST (N, TA, OtherPITC/Age/Sex/Result): HTS received results",
"HTS_TST (N, TA, PMTCT PostANC1 Pregnant-L&D/Age/Sex/Result): HTS received results",
"PMTCT_STAT (D, TA, Age/Sex): New ANC clients")

# Bring in the non-Tier import files for all partners "70310","70290","70287","87575", "87576", "87577","70301"
 #Download and save each IPs non-Tier file locally
  
Match_87576 <- read_excel("C:/Users/rpineteh/Documents/Github/LAB_PTCQI/MatCH_FY24Q3_87576_Non Tier_V1_20240802 - Farzana Alli.xlsx", sheet = "import") 

Match_87576_v2 <- Match_87576 %>% clean_names() %>% filter(data_element_name %in% cols ) %>% 
  select(org_unit_name,org_unit_uid , mech_uid ,value) %>% 
  mutate(Dataelement = "KMtAtCRNZl8", CategoryOptionCombo = "oCr3aOvULR9" , Period = "2024Q3") %>% 
  rename(AttributeOptionCombo = mech_uid, OrgUnit =org_unit_uid ) %>% 
  select(Dataelement ,Period,OrgUnit ,CategoryOptionCombo,AttributeOptionCombo, value) %>% 
  group_by_if(is.character) %>% summarise(value= sum(value)) %>% 
  ungroup()
  

#Append all partner non-tier files

#Then Append combined partner non-tier file with DATIM data for q1-q3

filename<-paste0(current_quarter,"_LAB_PTCQI_",Sys.Date(),".csv")

write.csv(LAB_PTQIv2.3,here("dataout",filename),row.names = FALSE)
