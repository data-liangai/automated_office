library(dplyr)
# 仅读取动态数据
df <- read.csv(file.choose()) %>% subset(status == 1)
df["query_date"] <- as.Date(df$query_time) %>% format("%Y%m")

# 对vin列进行去重,确保同一个vin在相同月份里只计算一次
df_unique <- df %>%
  distinct(vin, query_date, .keep_all = TRUE)

rm(df)

# 对不同因素下的数据进行计数
count_data <- df_unique %>%
  group_by(query_date) %>%
  summarise(Count = n())

# 定义阶梯单价函数
calculate_amount <- function(vin_billed_number) {
  brackets <- c(2,4,6,7)*10^5
  rates <- c(0.5189,0.4717,0.4245,0,0.4245)
  amount <- 0
  previous_bracket <- 0
  
  for (i in seq_along(brackets)) {
    if (vin_billed_number > brackets[i]) {
      amount <- amount + (brackets[i] - previous_bracket)*rates[i]*1.06
      previous_bracket <- brackets[i]
    } else {
      amount <- amount + (vin_billed_number - previous_bracket)*rates[i]*1.06
      return(amount) # 执行语句落到else即return
    }
  }
  
  amount <- amount + (vin_billed_number - previous_bracket)*rates[length(rates)]*1.06
  return(amount) # 计费车辆数超过70万时,计算超出部分再累加金额之后再return
}

# 计费车辆数
vin_billed_number <- sum(count_data$Count) - 10^5

# 计算总金额(含税)
amount <- calculate_amount(vin_billed_number)

# 打印结果
print("各月份动态数据计数:")
print(count_data)
print("计费车辆数:")
print(vin_billed_number)
print("总金额:")
print(amount)
