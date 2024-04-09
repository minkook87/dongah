# 필요한 패키지 로드
library(tidyverse)
library(readxl)
library(openxlsx)

# 데이터 로드
data5 <- read_xlsx("data/data_202401.xlsx") # 2024년 자료는 여기에 넣습니다
data6 <- read_xlsx("data/data_2023.xlsx")

# 해당진료과
data5 <- data5 %>%
  mutate(year = factor(year)) %>%
  mutate(month = factor(month)) %>%
  mutate(date = as.Date(date)) %>%
  mutate(yearmonth = floor_date(date, unit = "month")) %>% 
  mutate(class = factor(class, levels = c("암환자", "산과계", "외과계", "심호흡계", "심혈관계", "신경계", "기타내과계"))) %>% 
  mutate(dis = factor(dis, levels = c("E67", "G60", "J63", "R63",
                                      "O04", "O05", "O12", "O62", "O64",
                                      "I19", "I18", "H10", "I32", "C04", "L08", "D06", "E01", "F12", "B01",
                                      "E62", "E72", "E73", "E74", "F63",
                                      "F50", "F74", "F76", "F78",
                                      "B68", "B69", "B71", "B77", "B79",
                                      "G52", "G67", "I68", "I75"))) %>% 
  mutate(ICD10 = factor(ICD10)) %>% 
  mutate(sex = factor(sex, levels = c("M","F"))) %>%
  mutate(insurance = factor(insurance, levels = c("건강보험", "의료급여1종", "의료급여2종", "차상위1종", "차상위2종"))) %>% 
  mutate(adm = factor(adm)) %>% 
  mutate(charge = factor(charge)) %>%
  mutate(div = factor(div)) %>% 
  mutate(ward = factor(ward))

data6 <- data6 %>%
  mutate(year = factor(year)) %>%
  mutate(month = factor(month)) %>%
  mutate(date = as.Date(date)) %>%
  mutate(yearmonth = floor_date(date, unit = "month")) %>% 
  mutate(class = factor(class, levels = c("암환자", "산과계", "외과계", "심호흡계", "심혈관계", "신경계", "기타내과계"))) %>% 
  mutate(dis = factor(dis, levels = c("E67", "G60", "J63", "R63",
                                      "O04", "O05", "O12", "O62", "O64",
                                      "I19", "I18", "H10", "I32", "C04", "L08", "D06", "E01", "F12", "B01",
                                      "E62", "E72", "E73", "E74", "F63",
                                      "F50", "F74", "F76", "F78",
                                      "B68", "B69", "B71", "B77", "B79",
                                      "G52", "G67", "I68", "I75"))) %>% 
  mutate(ICD10 = factor(ICD10)) %>% 
  mutate(sex = factor(sex, levels = c("M","F"))) %>%
  mutate(insurance = factor(insurance, levels = c("건강보험", "의료급여1종", "의료급여2종", "차상위1종", "차상위2종"))) %>% 
  mutate(adm = factor(adm)) %>% 
  mutate(charge = factor(charge)) %>%
  mutate(div = factor(div)) %>% 
  mutate(ward = factor(ward))

# 진료과목 별로 데이터 처리
unique_divisions <- unique(data5$div)

# 전체 데이터의 자료를 추출 (Total)
all_summary <- data5 %>%
  summarise(
    all_mean_days = round(mean(days), 1),
    .groups = 'drop'
  )

all_summary_2023 <- data6 %>%
  summarise(
    all_mean_days_2023 = round(mean(days), 1),
    .groups = 'drop'
  )

# 전체 데이터의 자료를 추출 (진료군)
class_summary <- data5 %>% 
  group_by(class) %>% 
  summarise(
    class_mean_days = round(mean(days), 1),
    .groups = 'drop'
  )

# 2023년 데이터의 자료를 추출 (진료군)
class_summary_2023 <- data6 %>% 
  group_by(class) %>% 
  summarise(
    class_mean_days_2023 = round(mean(days), 1),
    .groups = 'drop'
  )

# 전체 데이터의 자료를 추출 (질병군)
total_summary <- data5 %>% 
  group_by(class, dis) %>% 
  summarise(
    total_mean_days = round(mean(days), 1),
    .groups = 'drop'
  ) %>% 
  filter(!is.na(dis))

total_summary

# 2023년 데이터의 자료를 추출 (질병군)
total_summary_2023 <- data6 %>% 
  group_by(class, dis) %>% 
  summarise(
    total_mean_days_2023 = round(mean(days), 1),
    .groups = 'drop'
  ) %>% 
  filter(!is.na(dis))

total_summary_2023

# 질병 코드와 설명 매핑
disease_descriptions <- tibble(
  dis = c("E67", "G60", "J63", "R63", "O04", "O05", "O12", "O62", "O64",
          "I19", "I18", "H10", "I32", "C04", "L08" ,"E01", "B01", "F12", "E62",
          "E72", "E73", "E74", "F63", "F50", "F74", "F76", "F78", "B68",
          "B69", "B71", "B77", "B79", "G52", "I68", "G67", "I75"),
  description = c("호흡기신생물", "소화기악성종양", "악성유방질환", "화학요법",
                  "질식분만[초산]", "질식분만[경산]", "자궁소파술 및 흡인소파술", "절박유산", "기타 산전 질환",
                  "슬부 수술", "견부 수술", "담낭절제술", "복잡 관절수술", "망막 및 유리체 수술", "요로결석 수술",
                  "주요 흉부 수술", "뇌동맥류 수술", "급성 심근경색증의 경피적 관상동맥 수술", "폐렴",
                  "만성폐색성폐질환", "기관지염", "천식", "심부전 및 쇼크", "진단 목적의 경피적 심혈관 시술",
                  "주요 부정맥", "협심증", "흉통", "뇌졸중", "뇌 및 두경부 혈관 질환", "뇌신경 및 말초신경 장애",
                  "발작 및 뇌전증", "외상성 혼미 및 혼수", "결장경 시술", "비외과적 경부 및 척추 상태",
                  "식도염, 위장관염 및 기타 위장관 증상", "견부, 상지, 주 관절, 슬부 하지 및 족관절의 손상")
)

# 전체
data7 <- data5 %>% 
  filter((surgery=='N'& ICD10%in%c("C18", "C34", "C22", "I63", "C50", "C16", "C25", "I20", "I50",
                                   "C20", "K63", "I21", "R07", "J18", "K80" ,"N18", "U07", "J96", "R57",
                                   "N40")) | (surgery=='Y'&ICD10%in%c("C34","I20","I63","C22","J18","I67","C16","N10","I21","K80")))  %>% 
  mutate(surgery = factor(surgery, levels = c("N", "Y"), labels = c("비수술", "수술"))) %>%
  arrange(surgery)


data8 <- data6 %>% 
  filter((surgery=='N'& ICD10%in%c("C18", "C34", "C22", "I63", "C50", "C16", "C25", "I20", "I50",
                                   "C20", "K63", "I21", "R07", "J18", "K80" ,"N18", "U07", "J96", "R57",
                                   "N40")) | (surgery=='Y'&ICD10%in%c("C34","I20","I63","C22","J18","I67","C16","N10","I21","K80")))  %>% 
  mutate(surgery = factor(surgery, levels = c("N", "Y"), labels = c("비수술", "수술"))) %>%
  arrange(surgery)

# 반복문, 과별로 정리
for (division in unique_divisions) {
  # 진료과별로 엑셀 시트 생성
  wb <- createWorkbook()
  
  # 과별로 데이터 추출
  division_data <- filter(data5, div == division)
  division_data_2023 <- filter(data6, div == division)
  
  #### Total 계산 (sheet1) ----
  # total 통계 계산
  all_stats <- division_data %>%
    summarise(
      all_patient_count = n(),
      all_div_mean_days = round(mean(days), 1),
      .groups = 'drop'
    ) %>% 
    mutate(class = "의과 입원 전체")
  
  # total 통계 계산 2023
  all_stats_2023 <- division_data_2023 %>%
    summarise(
      all_div_mean_days_2023 = round(mean(days), 1),
      .groups = 'drop'
    )
  
  # total_class_summary 출발
  combined_stats_t <- data.frame(all_stats,all_summary,all_summary_2023,all_stats_2023)
  
  #차 계산
  combined_stats_t <- combined_stats_t %>%
    mutate(
      all_diff_mean_days = as.numeric(all_div_mean_days) - as.numeric(all_div_mean_days_2023)
    ) %>% 
    mutate(etc = NA) %>% 
    select(class,all_patient_count,all_div_mean_days,all_mean_days,
           all_mean_days_2023,all_div_mean_days_2023,all_diff_mean_days,etc)
  
  #열이름 변경
  names(combined_stats_t) <- c("대상영역", "조회기간\n해당진료과\n환자수(명)", "조회기간\n해당진료과\n입원일수(일)",
                               "[참고1] 조회기간\n전체진료과\n입원일수(일)", "[참고2]\n2023년 전체진료과\n입원일수(일)", "[참고3] 2023년 해당진료과 입원일수(일)",
                               "[조회기간 해당진료과]-\n[참고3] 차이값(일)", "비고(극단값)")
  
  #### 진료군 계산 (sheet1) -----
  division_class_summary <- data.frame(class = c("암환자","산과계","외과계","심호흡계","심혈관계",
                                                 "신경계","기타내과계"))
  
  division_class_summary <- division_class_summary %>% 
    mutate(class = factor(class, levels=c("암환자","산과계","외과계","심호흡계","심혈관계",
                                          "신경계","기타내과계")))
  
  # division별 통계 계산
  division_stats <- division_data %>%
    group_by(class) %>%
    summarise(
      class_patient_count = n(),
      class_div_mean_days = round(mean(days), 1),
      .groups = 'drop'
    )
  
  # division별 통계 계산 2023
  division_stats_2023 <- division_data_2023 %>%
    group_by(class) %>%
    summarise(
      class_div_mean_days_2023 = round(mean(days), 1),
      .groups = 'drop'
    )
  
  # 이상치 계산
  class_outliers <- division_data %>%
    group_by(class) %>%
    filter(abs(scale(days)) > 3) %>%
    summarise(
      outlier_count = n(),
      outlier_values = paste(days, collapse = ", "),
      .groups = 'drop'
    )
  
  # division_class_summary 출발
  combined_stats7 <- left_join(division_class_summary, division_stats, by = c("class"))
  combined_stats7 <- left_join(combined_stats7, class_summary, by = c("class"))
  combined_stats7 <- left_join(combined_stats7, class_summary_2023, by = c("class"))
  combined_stats7 <- left_join(combined_stats7, division_stats_2023, by = c("class"))
  
  #차 계산
  combined_stats7 <- combined_stats7 %>%
    mutate(
      class_diff_mean_days = as.numeric(class_div_mean_days) - as.numeric(class_div_mean_days_2023)
    )
  
  combined_stats7 <- left_join(combined_stats7, class_outliers, by = c("class"))
  
  # 이상치가 있는 경우에만 이상치 값을 문자열로 변환하여 비고(극단값) 열에 추가
  combined_stats7 <- combined_stats7 %>%
    group_by(class) %>%
    mutate(
      outlier_values = ifelse(outlier_count > 0, {
        outlier_values <- paste(na.omit(outlier_values), collapse = ", ")
        outlier_values <- gsub(", ", "일, ", outlier_values)
        paste(outlier_count, "명(", outlier_values, "일)", sep = "")
      }, "")
    ) %>% 
    ungroup() %>%
    select(-outlier_count)
  
  # 열 이름 변경
  combined_stats7 <- data.frame(combined_stats7)
  
  names(combined_stats7) <- c("대상영역", "조회기간\n해당진료과\n환자수(명)", "조회기간\n해당진료과\n입원일수(일)",
                              "[참고1] 조회기간\n전체진료과\n입원일수(일)", "[참고2]\n2023년 전체진료과\n입원일수(일)", "[참고3] 2023년 해당진료과 입원일수(일)",
                              "[조회기간 해당진료과]-\n[참고3] 차이값(일)", "비고(극단값)")
  
  # 전체 + 7개 진료군 데이터 합치기
  combined_stats7 <- rbind(combined_stats_t, combined_stats7)
  
  #원하는대로 이름 바꾸기
  combined_stats7 <- combined_stats7 %>% 
    mutate(대상영역 = factor(c("의과 입원 전체","1) 암환자","2) 산과계","3) 외과계","4) 심호흡계","5) 심혈관계",
                           "6) 신경계","7) 기타내과계"),
                         levels = c("의과 입원 전체","1) 암환자","2) 산과계","3) 외과계","4) 심호흡계","5) 심혈관계",
                                    "6) 신경계","7) 기타내과계")))
  
  subjects <- data.frame("진료과" = division)
  
  # 진료과별 엑셀 시트 생성
  addWorksheet(wb, "sheet1")
  
  # 진료과별 통계 정보 출력
  writeDataTable(wb, sheet = "sheet1", subjects, startCol = 2, startRow = 1, tableStyle = "TableStyleLight8")
  writeData(wb, sheet = 1, x = "<의과 입원 전체 & 7개 진료군별 입원일수>", startRow = 4, startCol = 2)
  writeDataTable(wb, sheet = "sheet1", combined_stats7, startCol = 2, startRow = 6, tableStyle = "TableStyleMedium7")
  writeData(wb, sheet = 1, x = paste0("※조회기간: ", year(min(data5$yearmonth)), "년 ", month(min(data5$yearmonth)),
                                      "월 ~ ", year(max(data5$yearmonth)), "년 ", month(max(data5$yearmonth)), "월"), startRow = 5, startCol = 3)
  
  # sheet1 styling
  conditionalFormatting(wb, sheet = "sheet1", cols = 8, rows = 7:(nrow(combined_stats7) + 6), rule = ">0", style = createStyle(fontColour = "#FF0000"))
  conditionalFormatting(wb, sheet = "sheet1", cols = 2:9, rows = 6, rule = ">0", style = createStyle(fontColour = "black", textDecoration = "bold"))
  addStyle(wb, sheet = 1, 
           style = createStyle(border = "Left", borderColour = "#E26B0A", borderStyle = "thick"), 
           rows = 6:14, cols = 2)
  addStyle(wb, sheet = 1, 
           style = createStyle(border = "Right", borderColour = "#E26B0A", borderStyle = "thick"), 
           rows = 6:14, cols = 4)
  addStyle(wb, sheet = 1, 
           style = createStyle(border = "Bottom", borderColour = "#E26B0A", borderStyle = "thick"), 
           rows = 14, cols = 2:4, stack = TRUE)
  addStyle(wb, sheet = 1, 
           style = createStyle(border = "Top", borderColour = "#E26B0A", borderStyle = "thick"), 
           rows = 6, cols = 2:4, stack = TRUE)
  addStyle(wb, sheet = 1, style = createStyle(fgFill = "#D8E4BC"), rows = 6, cols = c(4, 7, 8), stack = TRUE)
  addStyle(wb, sheet = 1, style = createStyle(textDecoration = "bold"), rows = 4, cols = 2, stack = TRUE)
  addStyle(wb, sheet = 1, style = createStyle(wrapText = TRUE), rows = 6, cols = 3:8, stack = TRUE)
  setColWidths(wb, sheet = 1, cols = 8, widths = 21)
  setColWidths(wb, sheet = 1, cols = 2:7, widths = 18)
  
  ####bar plot1
  plot1 <- combined_stats7 %>%
    pivot_longer(cols = c(`[참고1] 조회기간\n전체진료과\n입원일수(일)`, `[참고2]\n2023년 전체진료과\n입원일수(일)`), 
                 names_to = "참고", values_to = "입원일수") %>%
    mutate(참고 = ifelse(참고 == "[참고2]\n2023년 전체진료과\n입원일수(일)", "23년 전체진료과", "조회기간 전체진료과")) %>%
    ggplot() +
    geom_bar(aes(x = 대상영역, y = 입원일수, fill = 참고), 
             stat = "identity", position = position_dodge(width = 0.9), width = 0.8, na.rm = TRUE) +
    geom_text(aes(x = 대상영역, y = 0, label = 입원일수, group = 참고),
              color = "white", fontface = "bold", vjust = -3, size = 4, 
              position = position_dodge(width = 0.9), na.rm = TRUE) +
    labs(title = "전체진료과 2023년 대비 입원일수 비교 그래프", x = NULL,
         y = NULL, fill = NULL) +
    scale_fill_manual(values = c("#17375E", "#953735"), 
                      labels = c("23년 전체진료과[참고2]", "조회기간 전체진료과[참고1]")) +
    scale_y_continuous(breaks = seq(0, 10, by = 2)) +
    theme_minimal() +
    theme(legend.position = "right",
          plot.title = element_text(hjust = 0.5, face = "bold"),
          axis.text.x = element_text(face = "bold"),
          legend.text = element_text(face = "bold"))
  
  print(plot1)
  insertPlot(wb, sheet = 1, width = 10, height = 2, xy=c(2, 16), fileType = "png",
             units = "in",dpi = 300)
  
  ####bar plot2
  plot2 <- combined_stats7 %>%
    pivot_longer(cols = c(`[참고3] 2023년 해당진료과 입원일수(일)`, `조회기간\n해당진료과\n입원일수(일)`), 
                 names_to = "참고", values_to = "입원일수") %>%
    mutate(참고 = ifelse(참고 == "조회기간\n해당진료과\n입원일수(일)", "23년 해당진료과", "조회기간 해당진료과")) %>%
    ggplot() +
    geom_bar(aes(x = 대상영역, y = 입원일수, fill = 참고), 
             stat = "identity", position = position_dodge(width = 0.9), width = 0.8, na.rm = TRUE) +
    geom_text(aes(x = factor(대상영역), y = 0, label = 입원일수, group = 참고), 
              color = "white", fontface = "bold", vjust = -2, size = 4,
              position = position_dodge(width = 0.9), na.rm = TRUE) +
    labs(title = "해당진료과 2023년 대비 입원일수 비교 그래프", x = NULL,
         y = NULL, fill = NULL) +
    scale_fill_manual(values = c("#4F81BD", "#C00000"), 
                      labels = c("23년 해당진료과[참고3]", "조회기간 해당진료과")) +
    scale_y_continuous(breaks = seq(0, 10, by = 2)) +
    theme_minimal() +
    theme(legend.position = "right",
          plot.title = element_text(hjust = 0.5, face = "bold"),
          axis.text.x = element_text(face = "bold"),
          legend.text = element_text(face = "bold"))
  
  print(plot2)
  insertPlot(wb, sheet = 1, width = 10, height = 2, xy=c(2, 26), fileType = "png",
             units = "in",dpi = 300)
  
  ####질병군 계산(sheet2)----
  # 과에 해당하는 질병만 추출
  division_summary <- division_data %>%
    group_by(class, dis) %>%
    summarise(
      .groups = 'drop'
    ) %>% 
    filter(!is.na(dis))
  
  # 담당자(교수)별 통계 계산
  professor_stats <- division_data %>%
    group_by(class, dis, charge) %>%
    summarise(
      patient_count = n(),
      mean_days = round(mean(days), 1),
      .groups = 'drop'
    ) %>% 
    filter(!is.na(dis))
  
  professor_stats_2023 <- division_data_2023 %>%
    group_by(class, dis, charge) %>%
    summarise(
      mean_days_2023 = round(mean(days), 1),
      .groups = 'drop'
    ) %>% 
    filter(!is.na(dis))
  
  # 이상치 계산
  outliers <- division_data %>%
    group_by(class, dis, charge) %>%
    filter(abs(scale(days)) > 3) %>%
    summarise(
      outlier_count = n(),
      outlier_values = paste(days, collapse = ", "),
      .groups = 'drop'
    )  %>% 
    filter(!is.na(dis))
  
  # division_summary와 professor_stats 합치기
  combined_stats <- left_join(division_summary, disease_descriptions, by = c("dis"))
  combined_stats <- left_join(combined_stats, professor_stats, by = c("class", "dis"))
  combined_stats <- left_join(combined_stats, total_summary, by = c("class", "dis"))
  combined_stats <- left_join(combined_stats, total_summary_2023, by = c("class", "dis"))
  combined_stats <- left_join(combined_stats, professor_stats_2023, by = c("class", "dis", "charge"))
  
  #차 계산
  combined_stats <- combined_stats %>%
    mutate(
      diff_mean_days = as.numeric(mean_days) - as.numeric(mean_days_2023)
    )
  
  combined_stats <- left_join(combined_stats, outliers, by = c("class", "dis", "charge"))
  
  # 이상치가 있는 경우에만 이상치 값을 문자열로 변환하여 비고(극단값) 열에 추가
  combined_stats <- combined_stats %>%
    group_by(class, dis, charge) %>%
    mutate(
      outlier_values = ifelse(outlier_count > 0, {
        outlier_values <- paste(na.omit(outlier_values), collapse = ", ")
        outlier_values <- gsub(", ", "일, ", outlier_values)
        paste(outlier_count, "명(", outlier_values, "일)", sep = "")
      }, "")
    ) %>% 
    ungroup() %>%
    select(-outlier_count) %>% 
    arrange(class, dis, desc(patient_count))
  
  # 진료군과 질병군이 중복되는 경우 진료군과 질병군 칸의 값을 지우고 빈 문자열로 대체
  combined_stats <- combined_stats %>%
    mutate(
      class = ifelse(duplicated(class), "", as.character(class)),
      description = ifelse(duplicated(description), "", as.character(description)),    
      total_mean_days = ifelse(duplicated(dis), "", as.numeric(total_mean_days)),
      dis = ifelse(duplicated(dis), "", as.character(dis))
    )
  
  # 열 이름 변경
  combined_stats <- data.frame(combined_stats)
  
  names(combined_stats) <- c("7개 진료군", "37개 질병군", "37개 질병군명", "교수", "조회기간\n교수별\n환자수(명)", 
                             "조회기간\n교수별\n입원일수(명)", "[참고1]\n조회기간\n전체진료과\n입원일수(일)",
                             "[참고2]\n2023년\n전체진료과\n입원일수(일)", "[참고3]\n2023년\n교수별\n입원일수(일)",
                             "[조회기간 교수별]-\n[참고3] 차이값(일)", "비고(극단값)")
  
  subjects <- data.frame("진료과" = division)
  
  # 진료과별 엑셀 시트 생성
  addWorksheet(wb, "sheet2")
  
  # 진료과별 통계 정보 출력
  writeDataTable(wb, sheet = "sheet2", subjects, startCol = 2, startRow = 1, tableStyle = "TableStyleLight8")
  writeData(wb, sheet = 2, x = "<37개 질병군별 교수별 입원일수>", startRow = 2, startCol = 4)
  writeDataTable(wb, sheet = "sheet2", combined_stats, startCol = 2, startRow = 4, tableStyle = "TableStyleLight14")
  conditionalFormatting(wb, sheet = "sheet2", cols = 11, rows = 5:ifelse(nrow(combined_stats)>0, nrow(combined_stats)+4, 5), rule = ">0", style = createStyle(fontColour = "#FF0000"))
  conditionalFormatting(wb, sheet = "sheet2", cols = 2:12, rows = 4, rule = ">0", style = createStyle(fontColour = "black", textDecoration = "bold"))
  addStyle(wb, sheet = 2, 
           style = createStyle(border = "Left", borderColour = "#E26B0A", borderStyle = "thick"), 
           rows = 4:(nrow(combined_stats) + 4), cols = 4)
  addStyle(wb, sheet = 2, 
           style = createStyle(border = "Right", borderColour = "#E26B0A", borderStyle = "thick"), 
           rows = 4:(nrow(combined_stats) + 4), cols = 6)
  addStyle(wb, sheet = 2, 
           style = createStyle(border = "Bottom", borderColour = "#E26B0A", borderStyle = "thick"), 
           rows = (nrow(combined_stats) + 4), cols = 4:6, stack = TRUE)
  addStyle(wb, sheet = 2, 
           style = createStyle(border = "Top", borderColour = "#E26B0A", borderStyle = "thick"), 
           rows = 4, cols = 4:6, stack = TRUE)
  addStyle(wb, sheet = 2, style = createStyle(fgFill = "#D8E4BC"), rows = 4, cols = c(6, 9, 10), stack = TRUE)
  addStyle(wb, sheet = 2, style = createStyle(textDecoration = "bold"), rows = 2, cols = 4, stack = TRUE)
  addStyle(wb, sheet = 2, style = createStyle(wrapText = TRUE), rows = 4, cols = 6:11, stack = TRUE)
  setColWidths(wb, sheet = 2, cols = 4, widths = 35)
  setColWidths(wb, sheet = 2, cols = 6:11, widths = 18)
  # 진료군 별로 테두리 추가
  num_rows <- nrow(combined_stats)
  num_cols <- ncol(combined_stats)
  last_division_row <- 3
  if(nrow(combined_stats)!=0) {
    for (col in 1:num_cols+1) {
      style <- createStyle(fgFill = "#FDE9D9")
      addStyle(wb, sheet = 2, rows = 5, cols = col, style = style, stack = TRUE)
    }
    for (i in 2:num_rows) {
      if (!is.na(combined_stats[i, "37개 질병군"]) && 
          combined_stats[i, "37개 질병군"] != "" &&
          combined_stats[i, "37개 질병군"] != combined_stats[i-1, "37개 질병군"]&&
          nrow(combined_stats)!=1)  {
        # 새로운 질병군이 등장하면 해당 행 색 변경
        for (col in 1:num_cols+1) {
          style <- createStyle(fgFill = "#FDE9D9")
          addStyle(wb, sheet = 2, rows = i + 4, cols = col, style = style, stack = TRUE)
        }
        last_division_row <- i + 3 # 새로운 질병군의 행 인덱스 업데이트
      }
    }
  }
  
  #### 전체 다빈도상병(sheet3)----
  # 각 코드 앞 번호 추가
  # 번호 매핑을 위한 데이터 프레임 생성 (비수술 그룹)
  number_non_surgery <- tibble(
    ICD10 = c("C18", "C34", "C22", "I63", "C50", "C16", "C25", "I20", "I50", 
              "C20", "K63", "I21", "R07", "J18", "K80", "N18", "U07", "J96", 
              "R57", "N40"),
    surgery = rep("비수술", 20),
    No = 1:20
  )
  
  # 번호 매핑을 위한 데이터 프레임 생성 (수술 그룹)
  number_surgery <- tibble(
    ICD10 = c("C34", "I20", "I63", "C22", "J18", "I67", "C16", "N10", "I21", "K80"),
    surgery = rep("수술", 10),
    No = 1:10
  )
  
  # 두 매핑 데이터 프레임 결합
  code_with_no <- bind_rows(number_non_surgery, number_surgery)
  
  code_with_no <- code_with_no %>% 
    select(surgery,No,ICD10)
  
  
  # 'ICD10' 열 필터링링
  all_p_number <- data7 %>%
    group_by(surgery,ICD10) %>%
    summarise(count = n(), .groups = 'drop') 
  
  all_p_day <- data7 %>% 
    group_by(surgery,ICD10) %>%
    summarise(mean_days = round(mean(days), 1),
              .groups = 'drop') 
  
  
  all_p_day_2023 <- data8 %>% 
    group_by(surgery,ICD10) %>%
    summarise(mean_days_2023 = round(mean(days), 1),
              .groups = 'drop') 
  
  disease_descriptions_icd10 <- tibble(
    ICD10 = c("C18", "C34", "C22", "I63", "C50", "C16", "C25", "I20", "I50",
              "C20", "K63", "I21", "R07", "J18", "K80" ,"N18", "U07", "J96", "R57",
              "N40", "I67", "N10"),
    description = c("결장의 악성 신생물", "기관지 및 폐의 악성 신생물", "간 및 간내 담관의 악성 신생물", "뇌경색증",
                    "유방의 악성 신생물", "위의 악성 신생물", "췌장의 악성 신생물", "협심증", "심부전",
                    "직장의 악성 신생물", "장의 기타 질환", "급성 심근경색증", "목구멍 및 가슴의 통증", "상세불명 병원체의 폐렴", "담석증",
                    "만성 신장병", "코로나 바이러스", "달리 분류되지 않은 호흡부전", "달리 분류되지 않은 쇼크",
                    "전립선 증식증", "기타 뇌혈관질환", "급성 세뇨관-간질신장염")
  )
  
  outliers <- data7 %>%
    group_by(surgery,ICD10) %>%
    filter(abs(scale(days)) > 3) %>%
    summarise(
      outlier_count = n(),
      outlier_values = paste(days, collapse = ", "),
      .groups = 'drop'
    )
  
  combined_stats_3 <- code_with_no %>%
    left_join(disease_descriptions_icd10, by = "ICD10") %>%
    left_join(all_p_number, by = c("ICD10", "surgery")) %>%
    left_join(all_p_day, by = c("ICD10", "surgery")) %>%
    left_join(all_p_day_2023, by = c("ICD10", "surgery"))
  
  table<- combined_stats_3 %>% 
    mutate(day_difference = mean_days - mean_days_2023)
  
  
  # 진료군과 질병군이 중복되는 경우 진료군과 질병군 칸의 값을 지우고 빈 문자열로 대체
  table <- table %>%
    mutate(surgery = ifelse(duplicated(surgery), "", as.character(surgery)))
 
  table <- left_join(table, outliers, by = c("surgery", "ICD10"))
   
  # 이상치가 있는 경우에만 이상치 값을 문자열로 변환하여 비고(극단값) 열에 추가
  table <- table %>%
    group_by(surgery, ICD10) %>%
    mutate(
      outlier_values = ifelse(outlier_count > 0, {
        outlier_values <- paste(na.omit(outlier_values), collapse = ", ")
        outlier_values <- gsub(", ", "일, ", outlier_values)
        paste(outlier_count, "명(", outlier_values, "일)", sep = "")
      }, "")
    ) %>% 
    ungroup() %>%
    select(-outlier_count)
  
  # 열 이름 변경
  table <- data.frame(table)
  
  names(table) <- c("구분", "No", "주상병", "상병명", "조회기간 전체진료과 환자수(명)", "조회기간 전체진료과 입원일수(일)",
                    "[참고] 23년 전체진료과 입원일수(일)", "[조회기관 전체]-[참고]차이값(일)",  "비고(극단값)")
  
  subjects <- data.frame("진료과" = "전체")
  addWorksheet(wb, "sheet3")  
  writeDataTable(wb, sheet = 3, subjects, startCol = 3, startRow = 1, tableStyle = "TableStyleLight8")
  writeData(wb, sheet = 3, x = "<2023년 본원 다빈도 상병(비수술 20개, 수술 10개) 입원일수>", startRow = 4, startCol = 3)
  writeDataTable(wb, sheet = 3, table, startCol = 3, startRow = 6, tableStyle = "TableStyleMedium6")
  conditionalFormatting(wb, sheet = 3, cols = 10, rows = 7:ifelse(nrow(combined_stats)>0, nrow(combined_stats)+6, 7), rule = ">0", style = createStyle(fontColour = "#FF0000"))
  addStyle(wb, sheet = 3, style = createStyle(textDecoration = "bold"), rows = 4, cols = 3, stack = TRUE)
  setColWidths(wb, sheet = 3, cols = 6, widths = 35)
  setColWidths(wb, sheet = 3, cols = c(3:5,7:11), widths = 18)
  
  # 엑셀 파일 저장
  saveWorkbook(wb, file = paste0("out/", division, ".xlsx"), overwrite = TRUE)
  removeWorksheet(wb, "sheet1")
  removeWorksheet(wb, "sheet2")
  removeWorksheet(wb, "sheet3")
  rm(wb)
}