########################    Afsan Sami     ########################
## Don't delete any codes from Rcode Master File Analysis Codes ##
################################################################################
# 0. Loading necessary libraries
install.packages(c("readxl","tidyr","tidyverse","dplyr","table1","expss",
                   "labelled","gtsummary","gt","nnet","knitr","kableExtra","car"))
library(readxl)
library(tidyr)
library(tidyverse)
library(dplyr)
library(table1)
library(expss) 
library(labelled)
library(gtsummary)
library(gt)
## Multinomial Logistic ##
# Basic data cleaning and visualization packages
library(tidyverse)
library(ggplot2) 
# Package for multinomial logistic regression
library(nnet) 
# To transpose output results
library(broom) 
# To create html tables
library(knitr)
library(kableExtra)
# To exponentiate results to get RRR
library(gtsummary) 
# To plot predicted values
library(ggeffects) 
# To get average marginal effect
library(marginaleffects) 
# For the Wald Test
library(car) 
#install.packages(c("nnet","broom","knitr","kableExtra","ggeffects","marginaleffects","car"))

################################################################################
setwd("E:\\Project")
mydata<-read_excel("main.xlsx",sheet="Coded")

########################### Labeling Variables #################################
A1=mydata %>% 
  select(FOEL,Sex,Age,RE,MFI,AF,CAS,PEA,ICA,FIU,ADHSNP,OPJ,FAWI,IICMR,OFL,RFL,UIEL,EAMHI,IUAAP,IUISQ,HSPD,WMCWF)
A1<-mutate(A1,
           FOEL=factor(FOEL,c(1,2,3,4,5),
                       labels=c("Never","Rarely","Sometimes","Often","Always")),
           Sex=factor(Sex,c(1,2),labels=c("Male","Female")),
           Age=factor(Age,c(1,2),
                      labels=c("age under 22", "greater than or equal to 22")),
           RE=factor(RE,c(1,2,3),labels=c("Hall","Mess","With family")),
           MFI=factor(MFI,c(1,2,3,4),
                      labels=c("Less than BDT 18,000 per month","BDT 18,000 to BDT 45,000 per month", 
                               "BDT 45,000 to BDT 135,000 per month","More than BDT 135,000 per month")),
           AF=factor(AF,c(1,2),
                     labels=c("Science and Engineering","Non-Science and Engineering")),
           CAS=factor(CAS,c(1,2),
                      labels=c("Undergraduate","Graduated")),
           PEA=factor(PEA,c(1,2),
                      labels=c("Yes","No")),
           ICA=factor(ICA,c(1,2),
                      labels=c("Yes","No")),
           FIU=factor(FIU,c(1,2),
                      labels=c("Yes","No")),
           ADHSNP=factor(ADHSNP,c(1,2,3,4),
                         labels=c("Less than 2 hours","2-4 hours","4-6 hours","Above 6 hours")),
           OPJ=factor(OPJ,c(1,2),
                      labels=c("Yes","No")),
           FAWI=factor(FAWI,c(1,2,3),
                       labels=c("Never","Sometimes","Always")),
           IICMR=factor(IICMR,c(1,2,3,4),
                        labels=c("Very importan","Importan","Slightly importan","Not importan")),
           OFL=factor(OFL,c(1,2),
                      labels=c("Yes","No")),
           RFL=factor(RFL,c(1,2,3,4,5,6),
                      labels=c("Introvert","Few friends","Family problem","Personal Relationships","Distance from Home","Others")),
           UIEL=factor(UIEL,c(1,2,3),
                       labels=c("Yes","No","Rarely")),
           EAMHI=factor(EAMHI,c(1,2),
                        labels=c("Yes","No")),
           IUAAP=factor(IUAAP,c(1,2,3),
                        labels=c("Yes","No","Not sure")),
           IUISQ=factor(IUISQ,c(1,2,3,4),
                        labels=c("Significantly","Moderately","Slightly","Not at all")),
           HSPD=factor(HSPD,c(1,2,3,4),
                       labels=c("Less than 4 hours","4-6 hours","6-8 hours","More than 8 hours")),
           WMCWF=factor(WMCWF,c(1,2,3,4,5),
                        labels=c("Never","Once","2-3 times","4-5 times","More than 5 times"))
)
#
var_label(A1)<-list(
  FOEL="Frequency of experiencing loneliness",
  Sex= "Sex",
  Age="Age",
  RE="Residence",
  MFI="Monthly family income",
  AF="Academic faculty",
  CAS="Current academic status",
  PEA="Participation in extracurricular activities",
  ICA="Internet connection availability",
  FIU="Frequent internet use",
  ADHSNP="Average daily hours spent online for non-academic purposes",
  OPJ="Online part-time job",
  FAWI="Feel anxious without internet access for a few hours",
  IICMR="Importance of internet connectivity for maintaining relationships with friends and family",
  OFL="Often feel lonely",
  RFL="Reason for your loneliness",
  UIEL="Using internet to escape loneliness",
  EAMHI="Experienced any mental health issues",
  IUAAP="Internet usage affecting academic performance",
  IUISQ="Internet usage impact on sleep quality",
  HSPD="Hours slept per day",
  WMCWF="Weekly meetups and chats with friends outside of class"
)

############################### Descriptive Table ##############################
MyStat <- list(all_continuous() ~ "{mean} ± {sd}", all_categorical() ~ "{n} ({p})")
MyDigit <- list( all_categorical() ~ c(0, 2), all_continuous() ~ c(2,2) )

DT=A1%>%
  tbl_summary(statistic = MyStat, digits = MyDigit)%>%
  bold_labels()
DT %>%  
  as_gt() %>% 
  gtsave(DT,filename="Descriptive_Table.docx",path="E:\\Project")

############################### Association Table ##############################
ATable<-A1%>%
  select(FOEL,Sex,Age,RE,MFI,AF,CAS,PEA,ICA,FIU,ADHSNP,OPJ,FAWI,IICMR,OFL,RFL,UIEL,EAMHI,IUAAP,IUISQ,HSPD,WMCWF)%>%
  tbl_summary(by=FOEL,missing="no",statistic=MyStat,digits = MyDigit)%>%
  add_overall(last=TRUE)%>%
  add_p(list(all_continuous()~"t.test"),pvalue_fun=~style_pvalue(.x,digits=3))%>%
  modify_caption("General characteristic of respondents")%>%
  add_stat_label()%>%
  bold_labels()
ATable
ATable %>%  
  as_gt() %>% 
  gtsave(ATable,filename="011 Association_Table.docx",path="E:\\Project")
########################### Simple Logistic regression #########################
################
mydata<-read_excel("main.xlsx",sheet="Coded")
A1=mydata %>% 
  select(FOEL,Sex,Age,RE,MFI,AF,CAS,PEA,ICA,FIU,ADHSNP,OPJ,FAWI,IICMR,OFL,RFL,UIEL,EAMHI,IUAAP,IUISQ,HSPD,WMCWF)
A1<-mutate(A1,
           FOEL=factor(FOEL,c(1,2,3,4,5),
                       labels=c("Never","Rarely","Sometimes","Often","Always")),
           Sex=factor(Sex,c(1,2),labels=c("Male","Female")),
           Age=factor(Age,c(1,2),
                      labels=c("age under 22", "greater than or equal to 22")),
           RE=factor(RE,c(1,2,3),labels=c("Hall","Mess","With family")),
           MFI=factor(MFI,c(1,2,3,4),
                      labels=c("Less than BDT 18,000 per month","BDT 18,000 to BDT 45,000 per month", 
                               "BDT 45,000 to BDT 135,000 per month","More than BDT 135,000 per month")),
           AF=factor(AF,c(1,2),
                     labels=c("Science and Engineering","Non-Science and Engineering")),
           CAS=factor(CAS,c(1,2),
                      labels=c("Undergraduate","Graduated")),
           PEA=factor(PEA,c(1,2),
                      labels=c("Yes","No")),
           ICA=factor(ICA,c(1,2),
                      labels=c("Yes","No")),
           FIU=factor(FIU,c(1,2),
                      labels=c("Yes","No")),
           ADHSNP=factor(ADHSNP,c(1,2,3,4),
                         labels=c("Less than 2 hours","2-4 hours","4-6 hours","Above 6 hours")),
           OPJ=factor(OPJ,c(1,2),
                      labels=c("Yes","No")),
           FAWI=factor(FAWI,c(1,2,3),
                       labels=c("Never","Sometimes","Always")),
           IICMR=factor(IICMR,c(1,2,3,4),
                        labels=c("Very importan","Importan","Slightly importan","Not importan")),
           OFL=factor(OFL,c(1,2),
                      labels=c("Yes","No")),
           RFL=factor(RFL,c(1,2,3,4,5,6),
                      labels=c("Introvert","Few friends","Family problem","Personal Relationships","Distance from Home","Others")),
           UIEL=factor(UIEL,c(1,2,3),
                       labels=c("Yes","No","Rarely")),
           EAMHI=factor(EAMHI,c(1,2),
                        labels=c("Yes","No")),
           IUAAP=factor(IUAAP,c(1,2,3),
                        labels=c("Yes","No","Not sure")),
           IUISQ=factor(IUISQ,c(1,2,3,4),
                        labels=c("Significantly","Moderately","Slightly","Not at all")),
           HSPD=factor(HSPD,c(1,2,3,4),
                       labels=c("Less than 4 hours","4-6 hours","6-8 hours","More than 8 hours")),
           WMCWF=factor(WMCWF,c(1,2,3,4,5),
                        labels=c("Never","Once","2-3 times","4-5 times","More than 5 times")),
           ## 
           FOEL = relevel(FOEL, ref = "Never"),
           Sex = relevel(Sex, ref = "Female"),
           Age = relevel(Age, ref = "age under 22"),
           RE = relevel(RE, ref = "Hall"),
           MFI = relevel(MFI, ref = "More than BDT 135,000 per month"),
           AF = relevel(AF, ref = "Science and Engineering"),
           CAS = relevel(CAS, ref = "Graduated"),
           PEA = relevel(PEA, ref = "No"),
           ICA = relevel(ICA, ref = "No"),
           FIU = relevel(FIU, ref = "No"),
           ADHSNP = relevel(ADHSNP, ref = "Less than 2 hours"),
           OPJ = relevel(OPJ, ref = "No"),
           FAWI = relevel(FAWI, ref = "Never"),
           IICMR = relevel(IICMR, ref = "Not importan"),
           OFL = relevel(OFL, ref = "No"),
           RFL = relevel(RFL, ref = "Introvert"),
           UIEL = relevel(UIEL, ref = "No"),
           EAMHI = relevel(EAMHI, ref = "No"),
           IUAAP = relevel(IUAAP, ref = "No"),
           IUISQ = relevel(IUISQ, ref = "Not at all"),
           HSPD = relevel(HSPD, ref = "More than 8 hours"),
           WMCWF = relevel(WMCWF, ref = "Never")
)
var_label(A1)<-list(
  FOEL="Frequency of experiencing loneliness",
  Sex= "Sex",
  Age="Age",
  RE="Residence",
  MFI="Monthly family income",
  AF="Academic faculty",
  CAS="Current academic status",
  PEA="Participation in extracurricular activities",
  ICA="Internet connection availability",
  FIU="Frequent internet use",
  ADHSNP="Average daily hours spent online for non-academic purposes",
  OPJ="Online part-time job",
  FAWI="Feel anxious without internet access for a few hours",
  IICMR="Importance of internet connectivity for maintaining relationships with friends and family",
  OFL="Often feel lonely",
  RFL="Reason for your loneliness",
  UIEL="Using internet to escape loneliness",
  EAMHI="Experienced any mental health issues",
  IUAAP="Internet usage affecting academic performance",
  IUISQ="Internet usage impact on sleep quality",
  HSPD="Hours slept per day",
  WMCWF="Weekly meetups and chats with friends outside of class"
)

OR1<- A1%>%
  select(FOEL,Sex,Age,RE,MFI,AF,CAS,PEA,ICA,FIU,ADHSNP,OPJ,FAWI,IICMR,OFL,RFL,UIEL,EAMHI,IUAAP,IUISQ,HSPD,WMCWF)%>%
  tbl_uvregression(method = glm, y=FOEL, method.args = list(family= binomial(link= "logit")),
                   exponentiate = TRUE, pvalue_fun = ~style_pvalue(.x,digits=3))%>%
  modify_column_merge(pattern = "{estimate} ({ci})", rows= !is.na(estimate))%>%
  modify_header(estimate~"**OR (95% CI)**")
#OR1

OR2<-glm(FOEL~Sex+Age+RE+MFI+AF+CAS+PEA+ICA+FIU+ADHSNP+OPJ+FAWI+IICMR+OFL+RFL+UIEL+EAMHI+IUAAP+IUISQ+HSPD+WMCWF,
         data=A1,family= binomial(link= "logit"))%>%
  tbl_regression(exponentiate = TRUE, pvalue_fun = ~style_pvalue(.x,digits=3))%>%
  modify_column_merge(pattern = "{estimate} ({ci})", rows= !is.na(estimate))%>%
  modify_header(estimate~"**OR (95% CI)**")
#OR2

Tab2<- tbl_merge(
  tbls = list(OR1, OR2), tab_spanner = c("**Unadjusted**","**Adjusted**")
)
#Tab2

Tab2%>%
  as_gt()%>%
  gtsave(filename = "Mytable.docx", path="E:\\Project")

################################################################################
# Chi-Square Test for Associations and Saving Results to Word
################################################################################

# Install necessary packages if not already installed
install.packages(c("readxl", "dplyr", "officer", "flextable"))

# Load libraries
library(readxl)
library(dplyr)
library(officer)
library(flextable)

# Set working directory and load data
setwd("E:\\Project")
mydata <- read_excel("main.xlsx", sheet = "Coded")

# Prepare data and perform Chi-Square tests
data_chi <- mydata %>%
  select(FOEL, FIU, ADHSNP, FAWI) %>%
  mutate(
    FOEL = factor(FOEL, levels = c(1,2,3,4,5), labels = c("Never","Rarely","Sometimes","Often","Always")),               # Often Feel Lonely
    FIU = factor(FIU, levels = c(1, 2), labels = c("Yes", "No")),               # Frequent Internet Use
    ADHSNP = factor(ADHSNP, levels = c(1,2,3,4),labels = c("Less than 2 hours","2-4 hours","4-6 hours","Above 6 hours")),
    FAWI = factor(FAWI, levels = c(1,2,3), labels = c("Never","Sometimes","Always"))             # Mental Health Issues Experienced
    
  )

# Perform Chi-Square tests and store results
chi_foel_fiu <- chisq.test(table(data_chi$FOEL, data_chi$FIU))
chi_foel_adhsnp <- chisq.test(table(data_chi$FOEL, data_chi$ADHSNP))
chi_foel_fawi <- chisq.test(table(data_chi$FOEL, data_chi$FAWI))

# Create a data frame of results for Word export
chi_results <- data.frame(
  Test = c("Frequency of experiencing loneliness vs Frequent Internet Use",
           "Frequency of experiencing loneliness vs Average Daily Hours Spent Online for Non-Academic Purposes",
           "Frequency of experiencing loneliness vs Feelings of Anxiety Without Internet Access"),
  Chi_Square = c(chi_foel_fiu$statistic,
                 chi_foel_adhsnp$statistic,
                 chi_foel_fawi$statistic),
  df = c(chi_foel_fiu$parameter,
         chi_foel_adhsnp$parameter,
         chi_foel_fawi$parameter),
  p_value = c(chi_foel_fiu$p.value,
              chi_foel_adhsnp$p.value,
              chi_foel_fawi$p.value)
)

# Convert the results to a flextable for Word
chi_table <- flextable(chi_results) %>%
  set_header_labels(
    Test = "Test",
    Chi_Square = "Chi-Square Statistic",
    df = "Degrees of Freedom",
    p_value = "P-Value"
  ) %>%
  autofit()

# Create Word document and add the flextable
doc <- read_docx() %>%
  body_add_par("Chi-Square Test Results for Associations", style = "heading 1") %>%
  body_add_flextable(chi_table)

# Save the document
print(doc, target = "E:\\Project\\ChiSquaretest.docx")









################################################################################
# Chi-Square Test for Specified Associations and Saving Results to Word
################################################################################

# Install necessary packages if not already installed
install.packages(c("readxl", "dplyr", "officer", "flextable"))

# Load libraries
library(readxl)
library(dplyr)
library(officer)
library(flextable)

# Set working directory and load data
setwd("E:\\Project")
mydata <- read_excel("main.xlsx", sheet = "Coded")

# Selecting and preparing variables for analysis
data_chi <- mydata %>%
  select(FIU, FOEL, FOEL_freq, UIEL, FAWI, ADHSNP, IUISQ, WMCWF, EAMHI, RE) %>%
  mutate(
    FIU = factor(FIU, levels = c(1, 2), labels = c("Yes", "No")),                     # Frequent Internet Use
    FOEL = factor(FOEL, levels = c(1, 2), labels = c("Yes", "No")),                   # Often Feel Lonely
    FOEL_freq = factor(FOEL_freq, levels = c(1, 2, 3, 4, 5),                          # Frequency of Experiencing Loneliness
                       labels = c("Never", "Rarely", "Sometimes", "Often", "Always")),
    UIEL = factor(UIEL, levels = c(1, 2, 3), labels = c("Yes", "No", "Rarely")),      # Using Internet to Escape Loneliness
    FAWI = factor(FAWI, levels = c(1, 2, 3), labels = c("Never", "Sometimes", "Always")), # Feelings of Anxiety Without Internet Access
    ADHSNP = factor(ADHSNP, levels = c(1, 2, 3, 4),                                   # Average Daily Hours Spent Online for Non-Academic Purposes
                    labels = c("Less than 2 hours", "2-4 hours", "4-6 hours", "Above 6 hours")),
    IUISQ = factor(IUISQ, levels = c(1, 2, 3, 4), labels = c("Significantly", "Moderately", "Slightly", "Not at all")), # Internet Usage Impact on Sleep Quality
    WMCWF = factor(WMCWF, levels = c(1, 2, 3, 4, 5),                                 # Weekly Meetups with Friends Outside of Class
                   labels = c("Never", "Once", "2-3 times", "4-5 times", "More than 5 times")),
    EAMHI = factor(EAMHI, levels = c(1, 2), labels = c("Yes", "No")),                # Experienced Any Mental Health Issues
    RE = factor(RE, levels = c(1, 2, 3), labels = c("Hall", "Mess", "With family"))  # Residence Type
  )

# Perform Chi-Square tests and store results
chi_results <- data.frame(
  Test = c("Frequent Internet Use vs Often Feel Lonely",
           "Frequent Internet Use vs Frequency of Experiencing Loneliness",
           "Using Internet to Escape Loneliness vs Feelings of Anxiety Without Internet Access",
           "Average Daily Hours Spent Online for Non-Academic Purposes vs Often Feel Lonely",
           "Internet Usage Impact on Sleep Quality vs Often Feel Lonely",
           "Weekly Meetups with Friends Outside of Class vs Often Feel Lonely",
           "Experienced Any Mental Health Issues vs Using Internet to Escape Loneliness",
           "Residence Type vs Frequency of Experiencing Loneliness"),
  Chi_Square = c(
    chisq.test(table(data_chi$FIU, data_chi$FOEL))$statistic,
    chisq.test(table(data_chi$FIU, data_chi$FOEL_freq))$statistic,
    chisq.test(table(data_chi$UIEL, data_chi$FAWI))$statistic,
    chisq.test(table(data_chi$ADHSNP, data_chi$FOEL))$statistic,
    chisq.test(table(data_chi$IUISQ, data_chi$FOEL))$statistic,
    chisq.test(table(data_chi$WMCWF, data_chi$FOEL))$statistic,
    chisq.test(table(data_chi$EAMHI, data_chi$UIEL))$statistic,
    chisq.test(table(data_chi$RE, data_chi$FOEL_freq))$statistic
  ),
  df = c(
    chisq.test(table(data_chi$FIU, data_chi$FOEL))$parameter,
    chisq.test(table(data_chi$FIU, data_chi$FOEL_freq))$parameter,
    chisq.test(table(data_chi$UIEL, data_chi$FAWI))$parameter,
    chisq.test(table(data_chi$ADHSNP, data_chi$FOEL))$parameter,
    chisq.test(table(data_chi$IUISQ, data_chi$FOEL))$parameter,
    chisq.test(table(data_chi$WMCWF, data_chi$FOEL))$parameter,
    chisq.test(table(data_chi$EAMHI, data_chi$UIEL))$parameter,
    chisq.test(table(data_chi$RE, data_chi$FOEL_freq))$parameter
  ),
  p_value = c(
    chisq.test(table(data_chi$FIU, data_chi$FOEL))$p.value,
    chisq.test(table(data_chi$FIU, data_chi$FOEL_freq))$p.value,
    chisq.test(table(data_chi$UIEL, data_chi$FAWI))$p.value,
    chisq.test(table(data_chi$ADHSNP, data_chi$FOEL))$p.value,
    chisq.test(table(data_chi$IUISQ, data_chi$FOEL))$p.value,
    chisq.test(table(data_chi$WMCWF, data_chi$FOEL))$p.value,
    chisq.test(table(data_chi$EAMHI, data_chi$UIEL))$p.value,
    chisq.test(table(data_chi$RE, data_chi$FOEL_freq))$p.value
  )
)

# Convert the results to a flextable for Word
chi_table <- flextable(chi_results) %>%
  set_header_labels(
    Test = "Test",
    Chi_Square = "Chi-Square Statistic",
    df = "Degrees of Freedom",
    p_value = "P-Value"
  ) %>%
  autofit()

# Create Word document and add the flextable
doc <- read_docx() %>%
  body_add_par("Chi-Square Test Results for Specified Associations", style = "heading 1") %>%
  body_add_flextable(chi_table)

# Save the document
print(doc, target = "E:\\Project\\ChiSquareResults_Specified.docx")

