#### Set working directory --------------------------
library(tidyverse)
library(readxl)

#### Load data --------------------------------------
data <- read_xlsx("입원일수 데이터_(2023년)_최종.xlsx")

#### 명세서 우선순위 --------------------------------
length(unique(data$접수번호))

data[duplicated(data$접수번호), ]

# 명세서 우선순위를 결정하는 알고리즘 작성
data_pre1 <- data %>% 
  mutate(계통 = case_when(
    as.numeric(substr(질병군,2,3)) <= 49 ~ "외과계",
    (as.numeric(substr(질병군,2,3)) >= 50 & as.numeric(substr(질병군,2,3)) <= 59) ~ "내과계시술",
    (as.numeric(substr(질병군,2,3)) >= 60 & as.numeric(substr(질병군,2,3)) <= 99) ~ "내과계"
  )) %>% 
  mutate(계통 = factor(계통, levels = c("외과계", "내과계시술", "내과계")))

# 그대로 처리하는 경우
data_pre2 <- data_pre1 %>% 
  group_by(접수번호) %>%   # Group by the columns that you want to check for duplicates
  mutate(Dup = n()) %>%  # Count the number of duplicates
  ungroup() %>% # Remove grouping
  filter(Dup == 1)

# 1. 질병군코드 알파벳 뒤 두자리 분류 
# - 외과계 (1순위)
# - 내과계시술 (2순위)
# - 내과계 (3순위)

# 중복되는 경우 별도로 처리
data_with_counts <- data_pre1 %>% 
  group_by(접수번호) %>%   # Group by the columns that you want to check for duplicates
  mutate(Dup = n()) %>%  # Count the number of duplicates
  ungroup() %>% # Remove grouping
  filter(Dup >= 2)

data_with_counts <- data_with_counts %>% 
  arrange(접수번호, 계통, desc(명세서총비용), 입원일자)

# 중복데이터 제거
data_with_counts2 = data_with_counts[-which(duplicated(data_with_counts$접수번호)),]

# 데이터 다시 결합
data_pre3 <- rbind(data_pre2, data_with_counts2)

#### AHRQ_CCS_진단군분류 ----------------------------
ahrq <- read_xlsx("입원일수 데이터_(2023년)_최종.xlsx", sheet = 2) %>% 
  select(명세서청구주상병코드, 분류)

length(unique(ahrq$명세서청구주상병코드))

ahrq <- na.omit(ahrq)

data1 <- left_join(data_pre3, ahrq, by='명세서청구주상병코드')

#### Classification (7개 진료군) --------------------
data2 <- data1 %>% 
  mutate(class = case_when(
    보조유형 == "CA" ~ "암환자",
    퇴원과 == "OG" & substr(명세서청구주상병코드, 1, 1) == "O" ~ "산과계",
    if_else(is.na(질병군), FALSE, as.numeric(substr(질병군, 2, 3))) <= 49 ~ "외과계",
    분류 %in% c("심호흡계", "심혈관계", "신경계") ~ 분류,
    TRUE ~ "기타내과계"
  )) %>% 
  mutate(class = factor(class, levels = c("암환자", "산과계", "외과계", "심호흡계", "심혈관계", "신경계", "기타내과계")))

library(skimr)
skim(data2$class)
skim(data2$질병군)

#### Classification (37개 질병군) --------------------
data3 <- data2 %>%
  mutate(disease = substr(data2$질병군,1,3))

# disases 질병군 코드는 missing 461개  
skim(data3$disease)

data4 <- data3 %>% 
  mutate(dis = case_when(
    class == "암환자" & disease == "E67" ~ "E67",
    class == "암환자" & disease == "G60" ~ "G60",
    class == "암환자" & disease == "J63" ~ "J63",
    class == "암환자" & disease == "R63" ~ "R63",
    
    class == "산과계" & disease == "O04" ~ "O04",
    class == "산과계" & disease == "O05" ~ "O05",
    class == "산과계" & disease == "O12" ~ "O12",
    class == "산과계" & disease == "O62" ~ "O62",
    class == "산과계" & disease == "O64" ~ "O64",
    
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
                                      "G52", "G67", "I68", "I75")))

# 30% 만이 37개 질병군에 해당
skim(data4$dis)

#### Data preprocessing -----------------------------
data5 <- data4 %>%
  mutate(date = as.Date(as.character(퇴원일자), format = "%Y%m%d")) %>% 
  mutate(month = factor(as.numeric(substr(퇴원일자, 5, 6)))) %>%
  mutate(년도 = factor(as.numeric(substr(퇴원일자, 1, 4))))  %>% 
  mutate(보험유형 = factor(보험유형, levels = c("건강보험", "의료급여1종", "의료급여2종", "차상위1종", "차상위2종"))) %>% 
  mutate(입원경로 = factor(입원경로)) %>% 
  mutate(퇴원주치의 = factor(퇴원주치의)) %>%
  mutate(명세서청구주상병코드 = factor(substr(명세서청구주상병코드,1,3))) %>% 
  select(접수번호,등록번호,환자명,year = 년도, month, date, class, dis, ICD10 = 명세서청구주상병코드,
         days = 입원일수, sex = 성별, age = 나이, insurance = 보험유형, adm = 입원경로, 
         charge = 퇴원주치의, div = 퇴원과, ward = 퇴원병실) %>% 
  mutate(year = factor(year)) %>% 
  mutate(yearmonth = floor_date(date, unit = "month")) %>% 
  mutate(sex = factor(sex)) %>% 
  mutate(age = as.numeric(age)) %>% 
  mutate(div = factor(div)) %>% 
  mutate(ward = factor(as.numeric(ward)))

skim(data5)

#### For R shiny ------------------------------------
data5 %>% 
  group_by(class) %>% 
  summarise(mean_days = mean(days),
            median_days = median(days),
            n = n())

writexl::write_xlsx(data5,"data_2023.xlsx",col_names = TRUE)

#### Exercise ---------------------------------------

data5 %>% 
  group_by(class, yearmonth) %>%  # Group by Year-Month
  summarise(Monthlydays = mean(days)) %>% # Calculate the mean for each month
  ggplot(aes(x = yearmonth, y = Monthlydays, group = class)) +
  geom_line() +
  labs(title = "Time Series Data", x = "Date", y = "Value")
