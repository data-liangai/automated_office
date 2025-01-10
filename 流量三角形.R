library(readxl);library(dplyr);library(openxlsx)
# 定义读取Excel文件的函数
read_all_sheets <- function(file_path) {
  # 获取所有工作表的名称
  sheet_names <- excel_sheets(file_path)
  
  # 读取所有工作表并存储在一个列表中
  sheets <- lapply(sheet_names,
                   function(sheet) read_excel(file_path, sheet = sheet))
  
  # 为列表中的每个元素命名
  names(sheets) <- sheet_names
  
  return(sheets)
}

# 使用示例
file_path <- file.choose()
all_sheets <- read_all_sheets(file_path)
# 创建空列表来保存要输出的三角形
mylist <- list()
################赔款################
payout <- all_sheets[["赔款"]]
payout <- payout[,c("服务中心名称","评估险类名称","事故发生月度",
                    "核赔通过月度","再保前赔款")]
payout$事故发生月度 <- as.factor(payout$事故发生月度)
payout$核赔通过月度 <- as.factor(payout$核赔通过月度)
payout <- lapply(payout,
                 function(x) if(is.character(x)) as.factor(x) else x) %>% 
  as.data.frame()
payout_triangle <- xtabs(再保前赔款~事故发生月度+核赔通过月度,data = payout)
################追偿################
recovery <- all_sheets[["追偿"]]
recovery <- recovery[,c("服务中心名称","评估险类名称","事故发生月度",
                        "费用审核完成时间","再保前赔款")]
recovery$事故发生月度 <- as.factor(recovery$事故发生月度)
recovery$费用审核完成时间 <- as.factor(recovery$费用审核完成时间)
recovery <- lapply(recovery,
                   function(x) if(is.character(x)) as.factor(x) else x) %>% 
  as.data.frame()
recovery_triangle <- xtabs(再保前赔款~事故发生月度+费用审核完成时间,
                           data = recovery)
################赔偿+追偿累计三角形################
# 单加第一列为0
recovery_triangle <- cbind('202405' = 0,recovery_triangle)
payout_recovery <- matrix(0,nrow = dim(payout_triangle)[1],
                          ncol = dim(payout_triangle)[2])
row.names(payout_recovery) <- rownames(payout_triangle)
colnames(payout_recovery) <- 0:(dim(payout_triangle)[2]-1)

# 计算累计三角形
# payout_recovery[,"0"] <- diag(payout_triangle)+diag(recovery_triangle)
for (i in 1:dim(payout_recovery)[1]) {
  for (j in 1:dim(payout_recovery)[2]) {
    if ((i+j) <= (dim(payout_triangle)[1]+1)){
      payout_recovery[i,j] <- sum(payout_triangle[i,i:(i+j-1)]) + 
        sum(recovery_triangle[i,i:(i+j-1)])
    }
  }
}
################未决################
undecided <- all_sheets[["未决"]]
undecided <- undecided[,c("服务中心名称","评估险类名称","事故发生月度",
                          "未决时点","再保前赔款(包含费用)")]
undecided$事故发生月度 <- as.factor(undecided$事故发生月度)
undecided$未决时点 <- as.factor(undecided$未决时点)
undecided <- lapply(undecided,
                    function(x) if(is.character(x)) as.factor(x) else x) %>% 
  as.data.frame()
undecided_triangle <- xtabs(再保前赔款.包含费用.~事故发生月度+未决时点,
                            data = undecided)

undecided_triangle2 <- matrix(0,nrow = dim(undecided_triangle)[1],
                              ncol = dim(undecided_triangle)[2])
row.names(undecided_triangle2) <- rownames(undecided_triangle)
colnames(undecided_triangle2) <- 0:(dim(undecided_triangle)[2]-1)

# 计算未决累计三角形
for (i in 1:dim(undecided_triangle2)[1]) {
  for (j in 1:dim(undecided_triangle2)[2]) {
    if ((i+j) <= (dim(undecided_triangle)[1]+1)){
      undecided_triangle2[i,j] <- undecided_triangle[i,(i+j-1)]
    }
  }
}
# 已报告赔款累计三角形
reported_payout_triangle = payout_recovery + undecided_triangle2
################满期保费################
maturity_premium <- all_sheets[["已赚清单"]]
maturity_premium <- maturity_premium[-1,]
maturity_premium1 <- maturity_premium[,c("服务中心名称","评估险类名称")]
maturity_premium2 <- maturity_premium[,(ncol(maturity_premium)-
                                          ncol(undecided_triangle)+1):ncol(maturity_premium)]
maturity_premium2 <- lapply(maturity_premium2,
                            function(x) if(is.character(x)) as.numeric(x) else x) %>% 
  as.data.frame()
maturity_premium <- cbind(maturity_premium1,maturity_premium2)
maturity_premium$服务中心名称 <- as.factor(maturity_premium$服务中心名称)
maturity_premium$评估险类名称 <- as.factor(maturity_premium$评估险类名称)
# 满期保费累计三角形
maturity_premium_triangle <- matrix(0,nrow = dim(undecided_triangle)[1],
                                    ncol = dim(undecided_triangle)[2])
row.names(maturity_premium_triangle) <- rownames(undecided_triangle)
colnames(maturity_premium_triangle) <- 0:(dim(undecided_triangle)[2]-1)

for (i in 1:dim(maturity_premium_triangle)[1]) {
  for (j in 1:dim(maturity_premium_triangle)[2]) {
    if ((i+j) <= (dim(maturity_premium_triangle)[1]+1)){
      maturity_premium_triangle[j,i] <- sum(maturity_premium2[,j])
    }
  }
}
################案件数################
cases <- all_sheets[[length(all_sheets)]]
cases["出险日期"] <- as.Date(cases$出险时间) %>% format("%Y%m")
cases["已报告赔款"] <- cases$已决赔款 + cases$未决赔款
#------------------已决案件------------------
decided_cases <- cases[,c("报案号","已决赔款","出险日期","评估时点")]
decided_cases <- aggregate(已决赔款 ~ 报案号+出险日期+评估时点,
                           data = decided_cases, FUN=sum)
# 已决案件数三角形
decided_cases_triangle <- matrix(0,nrow = dim(undecided_triangle)[1],
                                 ncol = dim(undecided_triangle)[2])
row.names(decided_cases_triangle) <- rownames(undecided_triangle)
colnames(decided_cases_triangle) <- 0:(dim(undecided_triangle)[2]-1)

result_decided_cases <- xtabs(已决赔款 != 0 ~ 出险日期+评估时点,
                              data = decided_cases)

# 将每一列放置在相应的副对角线上
for (i in 1:nrow(result_decided_cases)) {
  for (j in 1:ncol(result_decided_cases)) {
    if ((i + j) <= nrow(result_decided_cases) + 1) {
      decided_cases_triangle[i,j] <- result_decided_cases[i,i+j-1]
    }
  }
}

#------------------已报告案件------------------
reported_cases <- cases[,c("报案号","已报告赔款","出险日期","评估时点")]
reported_cases <- aggregate(已报告赔款 ~ 报案号+出险日期+评估时点,
                            data = reported_cases, FUN=sum)
# 已决案件数三角形
reported_cases_triangle <- matrix(0,nrow = dim(undecided_triangle)[1],
                                  ncol = dim(undecided_triangle)[2])
row.names(reported_cases_triangle) <- rownames(undecided_triangle)
colnames(reported_cases_triangle) <- 0:(dim(undecided_triangle)[2]-1)

result_reported_cases <- xtabs(已报告赔款 != 0 ~ 出险日期+评估时点,
                               data = reported_cases)

# 将每一列放置在相应的副对角线上
for (i in 1:nrow(result_reported_cases)) {
  for (j in 1:ncol(result_reported_cases)) {
    if ((i + j) <= nrow(result_reported_cases) + 1) {
      reported_cases_triangle[i,j] <- result_reported_cases[i,i+j-1]
    }
  }
}

# 已报告赔付率三角形
reported_payout_percent_triangle <- reported_payout_triangle / maturity_premium_triangle
# 结案率(件数)三角形
closure_rate_triangle1 <- decided_cases_triangle / reported_cases_triangle
# 结案率(金额)三角形
closure_rate_triangle2 <- payout_recovery / reported_payout_triangle
# 已决案均赔款三角形
decided_average_payout_triangle <- payout_recovery / decided_cases_triangle
# 已报告案均赔款三角形
reported_average_payout_triangle <- reported_payout_triangle / reported_cases_triangle
# 已决赔付率三角形
decided_payout_percent_triangle <- payout_recovery / maturity_premium_triangle

# 定义一个函数,将数据框中的小数转换为百分比形式,并保留2位小数,忽略NA和NaN
# convert_to_percentage <- function(df) {
#   df_percent <- apply(df, 1:2, function(x) {
#     if (is.na(x) | is.nan(x)) {
#       return(x)
#     } else {
#       return(sprintf("%.2f%%", x * 100))
#     }
#   })
#   return(df_percent)
# }

uptri_matrix <- function(matrix_example){
  # 将矩阵的右下三角部分设置为空
  for (i in 1:nrow(matrix_example)) {
    for (j in 1:ncol(matrix_example)) {
      if (i + j > nrow(matrix_example) + 1) {
        matrix_example[i, j] <- NA
      }
    }
  }
  return(matrix_example)
}
################输出流量三角形################
mylist[["赔款+追偿累计三角形"]] <- uptri_matrix(payout_recovery)
mylist[["未决累计三角形"]] <- uptri_matrix(undecided_triangle2)
mylist[["已报告赔款累计三角形"]] <- uptri_matrix(reported_payout_triangle)
mylist[["满期保费累计三角形"]] <- uptri_matrix(maturity_premium_triangle)
mylist[["已决案件数三角形"]] <- uptri_matrix(decided_cases_triangle)
mylist[["已报告案件数三角形"]] <- uptri_matrix(reported_cases_triangle)
mylist[["已报告赔付率三角形"]] <- uptri_matrix(reported_payout_percent_triangle)
mylist[["结案率(件数)三角形"]] <- closure_rate_triangle1 %>% uptri_matrix()
mylist[["结案率(金额)三角形"]] <- closure_rate_triangle2 %>% uptri_matrix()
mylist[["已决案均赔款三角形"]] <- uptri_matrix(decided_average_payout_triangle)
mylist[["已报告案均赔款三角形"]] <- uptri_matrix(reported_average_payout_triangle)
mylist[["已决赔付率三角形"]] <- uptri_matrix(decided_payout_percent_triangle)

#-----------------在不同sheet输出-----------------
# write.xlsx(mylist,"流量三角形.xlsx",rowNames = T,colNames = T)

#-----------------在同一个sheet输出-----------------
# 创建一个新的工作簿
wb <- createWorkbook()

# 添加一个工作表
addWorksheet(wb, "Sheet1")

# 定义一个函数，将数据框写入工作表
write_df_with_title <- function(wb, sheet, df, title, start_row) {
  writeData(wb, sheet, title, startRow = start_row, colNames = F)
  writeData(wb, sheet, df, startRow = start_row + 1, rowNames = T)
}

# 写入数据框及其标题
for (n in 1:length(mylist)) {
  write_df_with_title(wb, "Sheet1", mylist[[n]], names(mylist)[n], 
                      nrow(payout_recovery)*(n-1)+2*n-1)
}

# 保存工作簿
saveWorkbook(wb, "流量三角形.xlsx", overwrite = T)
