#### Set working directory --------------------------
setwd("C:\\Users\\USER\\Downloads\\project")

library(dplyr)
library(tidyverse)
library(readxl)

#### Load data --------------------------------------
data <- read_xlsx("raw_17666.xlsx")

#### Select variables -------------------------------
data1 <- data %>% 
  select("퇴원일자", "년도", "퇴원과", "성별", "나이", "퇴원병실",
         "입원일수", "보험유형", "보험보조유형", "입원경로", "퇴원주치의",
         "ADRG", "청구주상병코드", "CCS", "CCS구분", "분류")
  # 새로받은 데이터에는 퇴원병실은 없음

#### Classification ---------------------------------
data2 <- data1 %>% 
  mutate(class = ifelse(보험보조유형=="CA","암환자",
                        ifelse(퇴원과=="OG"&substr(청구주상병코드,1,1)=="O","산과",
                               ifelse(as.numeric(substr(ADRG,2,3))<=49,"외과계",
                                      ifelse(CCS구분=="심호흡계","심호흡계",
                                             ifelse(CCS구분=="심혈관계","심혈관계",
                                                    ifelse(CCS구분=="신경계","신경계","기타내과계")))))))

data2$class[is.na(data2$class)] <- "기타내과계"

data3 <- data2 %>% 
  mutate(class = factor(class, levels = c("암환자", "산과", "외과계", "심호흡계", "심혈관계", "신경계", "기타내과계"))) %>% 
  mutate(disease = substr(data2$ADRG,1,3))
  
summary(data3$class)
summary(data3$disease)

data4 <- data3 %>% 
  mutate(dis = case_when(
    class == "암환자" & disease == "E67" ~ "E67",
    class == "암환자" & disease == "G60" ~ "G60",
    class == "암환자" & disease == "J63" ~ "J63",
    class == "암환자" & disease == "R63" ~ "R63",
    
    class == "산과" & disease == "O04" ~ "O04",
    class == "산과" & disease == "O05" ~ "O05",
    class == "산과" & disease == "O12" ~ "O12",
    class == "산과" & disease == "O62" ~ "O62",
    class == "산과" & disease == "O64" ~ "O64",
    
    class == "외과계" & disease == "I19" ~ "I19",
    class == "외과계" & disease == "I18" ~ "I18",
    class == "외과계" & disease == "H10" ~ "H10",
    class == "외과계" & disease == "I32" ~ "I32",
    class == "외과계" & disease == "C04" ~ "C04",
    class == "외과계" & disease == "L08" ~ "L08",
    class == "외과계" & disease == "D06" ~ "D06",
    class == "외과계" & disease == "E01" ~ "E01",
    class == "외과계" & disease == "F12" ~ "F12",
    class == "외과계" & disease == "B01" ~ "B01",
    
    class == "심호흡계" & disease == "E62" ~ "E62",
    class == "심호흡계" & disease == "E72" ~ "E72",
    class == "심호흡계" & disease == "E73" ~ "E73",
    class == "심호흡계" & disease == "E74" ~ "E74",
    class == "심호흡계" & disease == "F63" ~ "F63",
    
    class == "심혈관계" & disease == "F50" ~ "F50",
    class == "심혈관계" & disease == "F74" ~ "F74",
    class == "심혈관계" & disease == "F76" ~ "F76",
    class == "심혈관계" & disease == "F78" ~ "F78",
    
    class == "신경계" & disease == "B68" ~ "B68",
    class == "신경계" & disease == "B69" ~ "B69",
    class == "신경계" & disease == "B71" ~ "B71",
    class == "신경계" & disease == "B77" ~ "B77",
    class == "신경계" & disease == "B79" ~ "B79",
    
    class == "기타내과계" & disease == "G52" ~ "G52",
    class == "기타내과계" & disease == "G67" ~ "G67",
    class == "기타내과계" & disease == "I68" ~ "I68",
    class == "기타내과계" & disease == "I75" ~ "I75"
  )) %>% 
  mutate(dis = factor(dis, levels = c("E67", "G60", "J63", "R63",
                                      "O04", "O05", "O12", "O62", "O64",
                                      "I19", "I18", "H10", "I32", "C04", "L08", "D06", "E01", "F12", "B01",
                                      "E62", "E72", "E73", "E74", "F63",
                                      "F50", "F74", "F76", "F78",
                                      "B68", "B69", "B71", "B77", "B79",
                                      "G52", "G67", "I68", "I75"))) %>%  
  mutate(년도 = factor(년도))

summary(data4$dis)
summary(data4$dis)
summary(data4$년도)

#### Data preprocessing -----------------------------
data5 <- data4 %>%
  mutate(date = as.Date(as.character(퇴원일자), format = "%Y%m%d")) %>% 
  mutate(month = factor(as.numeric(substr(퇴원일자, 5, 6)))) %>%
  mutate(보험유형 = factor(보험유형, levels = c("건강보험", "의료급여1종", "의료급여2종", "차상위1종", "차상위2종"))) %>% 
  mutate(입원경로 = factor(입원경로)) %>% 
  mutate(퇴원주치의 = factor(퇴원주치의)) %>% 
  select(year = 년도, month, date, class, dis, days = 입원일수, 
         sex = 성별, age = 나이, insurance = 보험유형, adm = 입원경로, 
         charge = 퇴원주치의, div = 퇴원과, ward = 퇴원병실) %>% 
  mutate(year = factor(year)) %>% 
  mutate(yearmonth = floor_date(date, unit = "month")) %>% 
  mutate(sex = factor(sex)) %>% 
  mutate(age = as.numeric(age)) %>% 
  mutate(div = factor(div)) %>% 
  mutate(ward = factor(as.numeric(ward)))

data5$age[is.na(data5$age)] <- 0

#summary(data5)
#skim(data5)

#### For R shiny ------------------------------------
data5 %>% 
  group_by(class) %>% 
  summarise(mean_days = mean(days),
            median_days = median(days),
            n = n())

summary(data5$div)
summary(data5$age)

unique(data5$insurance)
unique(data5$charge)

writexl::write_xlsx(data5,"data.xlsx",col_names = TRUE)

#### Exercise ---------------------------------------

data5 %>% 
  group_by(class, yearmonth) %>%  # Group by Year-Month
  summarise(Monthlydays = mean(days)) %>% # Calculate the mean for each month
  ggplot(aes(x = yearmonth, y = Monthlydays, group = class)) +
  geom_line() +
  labs(title = "Time Series Data", x = "Date", y = "Value")

#### Random data generation -------------------------

set.seed(100)

data5$year <- sample(unique(data5$year), 32992, replace = TRUE)

data5$month <- sample(unique(data5$month), 32992, replace = TRUE)

data5$date <- sample(as.Date("2020-01-01"):as.Date("2022-12-31"), 32992, replace = TRUE)

data5$class <- sample(unique(data5$class), 32992, replace = TRUE)

data5$dis <- sample(unique(data5$dis), 32992, replace = TRUE)

data5$days <- sample(0:100, 32992, replace = TRUE)

data5$sex <- sample(unique(data5$sex), 32992, replace = TRUE)

data5$age <- sample(0:100, 32992, replace = TRUE)

data5$insurance <- sample(unique(data5$insurance), 32992, replace = TRUE)

data5$adm <- sample(unique(data5$adm), 32992, replace = TRUE)

data5$charge <- sample(unique(data5$charge), 32992, replace = TRUE)

data5$div <- sample(unique(data5$div), 32992, replace = TRUE)

data5$ward <- sample(unique(data5$ward), 32992, replace = TRUE)

sample = data5[sample(nrow(data5), 10000, replace = TRUE),]

sample$date = as.Date(sample$date, origin = "1970-01-01")

writexl::write_xlsx(sample,"data.xlsx",col_names = TRUE)
