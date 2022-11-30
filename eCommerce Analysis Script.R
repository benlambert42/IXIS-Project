library(tidyverse)
library(readr)
library(lubridate)
library(openxlsx)

# Load both csv files
addsToCart <- read_csv("IXIS/DataAnalyst_Ecom_data_addsToCart.csv")
seshData <- read_csv("IXIS/DataAnalyst_Ecom_data_sessionCounts.csv")

# Check Classes of columns, nrows, ncols
glimpse(addsToCart)
glimpse(seshData)

# Check top and bottom of data sets to make check if it looks as it should
head(addsToCart)
head(seshData)
tail(addsToCart)
tail(seshData)

# Check for Duplicates, Check for NA's
sum(duplicated(addsToCart))
sum(duplicated(seshData))

sum(is.na(addsToCart))
sum(is.na(seshData))

#Both data sets have the wrong class for date
# The addsToCart data has year and month in separate cols, need to combine
addsToCart_date <- addsToCart %>% mutate(date = ym(paste(dim_year,"-",dim_month)))  
glimpse(addsToCart_date)

# SeshData has date but not in correct format
seshData_date <- seshData %>% 
  mutate(date = mdy(dim_date))
glimpse((seshData_date))

# Calc percent change from 2012 to 2013 on all metrics
aveMetrics2012 <- seshData_date %>% 
  filter(date < '2013-01-01') %>%
  summarize(mean_ses = mean(sessions),
            mean_tra = mean(transactions),
            mean_qty = mean(QTY))
aveMetrics2012

aveMetrics2013 <- seshData_date %>% 
  filter(date >= '2013-01-01') %>%
  summarize(mean_ses = mean(sessions),
            mean_tra = mean(transactions),
            mean_qty = mean(QTY))
aveMetrics2013

percentChange_12to13 <-rbind(aveMetrics2012,aveMetrics2013) %>%
  mutate(Session_perc = 100*(mean_ses - lag(mean_ses))/lag(mean_ses),
         Trans_perc = 100*(mean_tra - lag(mean_tra))/lag(mean_tra),
         QTY_perc = 100*(mean_qty - lag(mean_qty))/lag(mean_qty))

percentChange_12to13

aveAddsToCart2012 <- addsToCart_date %>% 
  filter(date < '2013-01-01') %>%
  summarize(mean_atc = mean(addsToCart))
aveAddsToCart2012

aveAddsToCart2013 <- addsToCart_date %>% 
  filter(date >= '2013-01-01') %>%
  summarize(mean_atc = mean(addsToCart))
aveAddsToCart2013

percentChangeAtc_12to13 <-rbind(aveAddsToCart2012,aveAddsToCart2013) %>%
  mutate(atc_perc = 100*(mean_atc - lag(mean_atc))/lag(mean_atc))
percentChangeAtc_12to13

# Start testing out plots
linePlot <- function(df,x,y) {
  df %>%
    ggplot(aes({{x}},{{y}})) + 
    geom_line()
  }
linePlot(seshData_date,date,transactions)
linePlot(seshData_date,date,sessions)
linePlot(seshData_date,date,QTY)

scatPlot <- function(df,x,y) {
  df %>%
    ggplot(aes({{x}},{{y}})) + 
    geom_point() +
    geom_smooth(se = FALSE)
}
scatPlot(seshData_date,date,sessions)
scatPlot(seshData_date,date,transactions)
scatPlot(seshData_date,date,QTY)

# These plots are tough to see any trends/insight. Will group by month and run again
seshData_month <- seshData_date %>%
  group_by(month = floor_date(date, "month")) %>%
  summarize(sum_ses = sum(sessions),
            sum_tra = sum(transactions),
            sum_qty = sum(QTY))
seshData_month

scatPlot(seshData_month, month, sum_ses)
scatPlot(seshData_month, month, sum_tra)
scatPlot(seshData_month, month, sum_qty)
scatPlot(addsToCart_date, date, addsToCart)

# This time we see a clear positive trend line 
# In addition to breaking out by month it may be useful to further break out by device
seshData_mthDev <- seshData_date %>%
  group_by(month = floor_date(date, "month"), dim_deviceCategory) %>%
  summarize(sum_ses = sum(sessions),
            sum_tra = sum(transactions),
            sum_qty = sum(QTY))
seshData_mthDev

scatPlot_dev <- function(df,x,y) {
  df %>%
    ggplot(aes({{x}},{{y}},color = dim_deviceCategory)) + 
    geom_point() +
    geom_smooth(se = FALSE)
}
scatPlot_dev(seshData_mthDev, month, sum_ses)
scatPlot_dev(seshData_mthDev, month, sum_tra)
scatPlot_dev(seshData_mthDev, month, sum_qty)

# With the latest plots we are still seeing the positive trendline starting at 
# the beginning of 2013, but we are also seeing the rate of change of the 
# different device categories are quite different with desktop having the 
# fastest rate of change

##### Start of creating xlsx
# Sheet 1 Month*Device aggregation of Session, Transactions, QTY, ECR
MonthDevice_Agg <- seshData_mthDev %>%
  mutate(ECR = sum_tra/sum_ses)
MonthDevice_Agg

scatPlot_dev(MonthDevice_Agg, month, ECR) + geom_abline()

# Sheet 2 Month over Month Comparison; Sessions, Transactions, QTY, AddsToCart
MoM_compare <- seshData_month %>%
  filter(month >= '2013-05-01') %>%
  left_join(addsToCart_date, by = c("month"="date")) %>%
  transmute(
    month, sum_ses, sum_tra, 
    sum_qty, addsToCart, 
    ecr = sum_tra/sum_ses) %>%
  mutate(Session_Diff_MoM = (sum_ses - lag(sum_ses)),
         Transaction_Diff_MoM = (sum_tra - lag(sum_tra)),
         QTY_Diff_MoM = (sum_qty - lag(sum_qty)),
         AddsToCart_Diff_MoM = (addsToCart - lag(addsToCart)),
         ECR_Diff_MoM = (ecr - lag(ecr)),
         Session_AbsDiff_MoM = abs(sum_ses - lag(sum_ses)),
         Transaction_AbsDiff_MoM = abs(sum_tra - lag(sum_tra)),
         QTY_AbsDiff_MoM = abs(sum_qty - lag(sum_qty)),
         AddsToCart_AbsDiff_MoM = abs(addsToCart - lag(addsToCart)),
         ECR_AbsDiff_MoM = abs(ecr - lag(ecr)),
         Session_Perc_MoM = 100*(sum_ses - lag(sum_ses)) / lag(sum_ses),
         Transaction_Perc_MoM = 100*(sum_tra - lag(sum_tra))/lag(sum_tra),
         QTY_Perc_MoM = 100*(sum_qty - lag(sum_qty))/lag(sum_qty),
         AddsToCart_Perc_MoM = 100*(addsToCart - lag(addsToCart))/lag(addsToCart),
         ECR_Perc_MoM = 100*(ecr - lag(ecr))/lag(ecr))
MoM_compare

# Create xlsx for export of tables
Ecom_RefTables <- createWorkbook()
addWorksheet(Ecom_RefTables, "Month by Device Aggregation")
addWorksheet(Ecom_RefTables, "Month by Month Comparison")
writeData(Ecom_RefTables, sheet = "Month by Device Aggregation", x = MonthDevice_Agg)
writeData(Ecom_RefTables, sheet = "Month by Month Comparison", x = MoM_compare)
saveWorkbook(Ecom_RefTables, "IXIS/Ecommerce Reference Tables.xlsx")