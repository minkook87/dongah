# 과별로 데이터 추출
division_data <- filter(data5, div == 'CS')
division_data_2023 <- filter(data6, div == 'CS')

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
