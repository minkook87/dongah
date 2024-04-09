# 필요한 패키지 로드
library(tidyverse)
library(readxl)
library(openxlsx)

data5 <- read_xlsx("data_2023.xlsx") # 2024년 자료는 여기에 넣습니다
data6 <- read_xlsx("data_2023.xlsx")

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

# 진료과별로 엑셀 시트 생성
wb <- createWorkbook()

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

# 반복문, 과별로 정리
for (division in unique_divisions) {
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
  names(combined_stats_t) <- c("대상영역", "조회기간 해당진료과 환자수(명)", "조회기간 해당진료과 입원일수(명)", 
                              "[참고1] 조회기간 전체진료과 입원일수(일)", "[참고2] 2023년 전체진료과 입원일수(일)", "[참고3] 2023년 해당진료과 입원일수(일)",
                              "[조회기간 해당진료과]-[참고3] 차이값(일)", "비고(극단값)")
  
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
  
  names(combined_stats7) <- c("대상영역", "조회기간 해당진료과 환자수(명)", "조회기간 해당진료과 입원일수(명)",
                             "[참고1] 조회기간 전체진료과 입원일수(일)", "[참고2] 2023년 전체진료과 입원일수(일)", "[참고3] 2023년 해당진료과 입원일수(일)",
                             "[조회기간 해당진료과]-[참고3] 차이값(일)", "비고(극단값)")
  
  # 중간에 새로운 행을 추가
  combined_stats7 <- rbind(combined_stats_t, combined_stats7)
  
  #### 질병군 계산 (sheet2) -----
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
  
  # 담당자(교수)별 통계 계산 2023
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
    ) %>% 
    filter(!is.na(dis))
  
  # division_summary와 professor_stats 합치기
  # 2023년 자료도 붙인다
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
  
  names(combined_stats) <- c("7개 진료군", "37개 질병군", "37개 질병군명", "교수", "조회기간 교수별 환자수(명)", "조회기간 교수별 입원일수(명)",
                             "[참고1] 조회기간 전체진료과 입원일수(일)", "[참고2] 2023년 전체진료과 입원일수(일)", "[참고3] 2023년 교수별 입원일수(일)",
                            "[조회기간 교수별]-[참고3] 차이값(일)", "비고(극단값)")
  
  subjects <- data.frame("진료과" = division)
  
  # 진료과별 엑셀 시트 생성
  addWorksheet(wb, sheetName = "")
  
  # 진료과별 통계 정보 출력
  writeDataTable(wb, sheet = division, subjects, startCol = 2, startRow = 1, tableStyle = "TableStyleMedium7")
  writeDataTable(wb, sheet = division, combined_stats, startCol = 2, startRow = 4, tableStyle = "TableStyleMedium7")
  conditionalFormatting(wb, sheet = division, cols = 11, rows = 5:(nrow(combined_stats) + 4), rule = ">0", style = createStyle(fontColour = "#FF0000"))

    # 진료군 별로 테두리 추가
  num_rows <- sum(data5$div == division)
  num_cols <- ncol(combined_stats)
  last_division_row <- 4
  for (i in 2:num_rows) {
    if (i > 1 && 
        !is.na(combined_stats[i, "7개 진료군"]) && 
        !is.na(combined_stats[i - 1, "7개 진료군"]) &&
        combined_stats[i, "7개 진료군"] != "" &&
        combined_stats[i, "7개 진료군"] != combined_stats[i - 1, "7개 진료군"]) {
      # 새로운 진료군이 등장하면 해당 행 바로 위에 테두리 추가
      for (col in 1:num_cols+1) {
        style <- createStyle(border = "bottom")
        addStyle(wb, sheet = division, rows = i + 3, cols = col, style = style)
      }
      last_division_row <- i + 3 # 새로운 진료군의 행 인덱스 업데이트
    }
  }
}  

# 엑셀 파일 저장
saveWorkbook(wb, "group_stats.xlsx", overwrite = TRUE)
