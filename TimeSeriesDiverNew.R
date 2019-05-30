library(xts)
library(dplyr)
library(tidyverse)
library(readxl)
library(astsa)
library(forecast)
library(lubridate)
library(writexl)
library(timeDate)
library(rlist)


#Script Parameters and Tools
RunNeuralNet <- FALSE
NeuralnetIterations <- 20
rollperiod <- 6
forecastperiod <- 12
numberyearsOPDdata <- 8
filteroutspecials <- c("Pallitive Care", "Diabetic Day Centre", "Chiropody", "Neurophysiology", "General Medical", "OMNITEST", "Oncology Radiation")
'%!in%' <- function(x,y)!('%in%'(x,y))
ReferralFile <- "OPD Referrals.xlsx"
ReferralsSheet <- "Referrals"
Activityfile <- "OPDActivityfromDiver.xlsx"

#Pull in Referrals grouped by Specialty
Referrals <- read_excel(ReferralFile, 
                        sheet = ReferralsSheet)
Referrals <- Referrals[-1,]
Referrals$WeekEnding <- str_remove(Referrals$WeekEnding, "Week Ending ")
Referrals$WeekEnding <- ymd(Referrals$WeekEnding)
Referrals <- select(Referrals, WeekEnding, BookSpecialty, BookConsultant, Outcome, Referrals)

#set the end period by rolling back to previous month end from the max week
fullmonthmax <- rollback(max(Referrals$WeekEnding))

#filter out unwanted
Referrals <- Referrals %>%
  filter(Outcome == "Seen",
         year(WeekEnding) != 2011,
         WeekEnding <= fullmonthmax) %>%
  filter(BookSpecialty %!in% filteroutspecials) %>%
  group_by( WeekEnding, BookSpecialty) %>%
  summarise(Referrals = sum(Referrals))

#OPD dataframe
OPD <- lapply(1:numberyearsOPDdata, function(i) read_excel(Activityfile, sheet = i))
OPD <- bind_rows(OPD)


#Tidy up OPD Dataframe
OPD$WeekEnding <- str_remove(OPD$WeekEnding, "Week Ending ")

OPD$NewReturn <- str_remove(OPD$NewReturn, " PATIENTS :")

OPD$WeekEnding <- ymd(OPD$WeekEnding)

OPD <- select(OPD, WeekEnding, BookSpecialty, NewReturn, OPD = `O/P Count`, DNA = 'DNA Count') %>%
  filter(WeekEnding <= fullmonthmax) %>%
  filter(BookSpecialty %!in% filteroutspecials)

OPD$NewReturn <- as.factor(OPD$NewReturn)


OPDWIP <- NULL

#Get the News
OPDWIP$News <- OPD %>%
  filter(NewReturn == "NEW") %>%
  group_by(WeekEnding, BookSpecialty) %>%
  summarise(NewOPD  = sum(OPD), NewDNA = sum(DNA))

#Get the Retunrs
OPDWIP$Returns <- OPD %>%
  filter(NewReturn == "RETURN") %>%
  group_by(WeekEnding, BookSpecialty) %>%
  summarise(ReturnOPD = sum(OPD), ReturnDNA = sum(DNA))

#Join them Together
OPD <- full_join(OPDWIP$News, OPDWIP$Returns, by = c("WeekEnding", "BookSpecialty"))
remove(OPDWIP)

#Join them to the Referrals
OPD <- full_join(OPD, Referrals, by = c("WeekEnding", "BookSpecialty"))
remove(Referrals)

#Set the Weekend to be month end dates
OPD$WeekEnding <- as.Date(timeLastDayInMonth(as.character(OPD$WeekEnding)))

#Group and Sum by new date and specialty
OPD <- OPD %>%
  group_by(WeekEnding, BookSpecialty) %>%
  summarise(NewOPD = sum(NewOPD, na.rm = TRUE),
            NewDNA = sum(NewDNA, na.rm = TRUE),
            ReturnOPD = sum(ReturnOPD, na.rm = TRUE),
            ReturnDNA = sum(ReturnDNA, na.rm = TRUE),
            Referrals = sum(Referrals, na.rm = TRUE))


#Split the OPD data based on Specialty
OPD <- split.data.frame(OPD, OPD$BookSpecialty)

#Function to Convert Specialties to TS on Demand
converttoMxts <- function(x) {
  x_matrix <- as.matrix(x$Referrals)
  na.fill(as.xts(x_matrix,order.by = x$WeekEnding, frequency = 12), fill = 0)
}

#Create 80% Confidence Forecast Function
CreateDemandForecasts <- function(x) {
  y <- converttoMxts(x)
  Z <- forecast(auto.arima(y, stepwise = FALSE, approximation = FALSE), h = forecastperiod)
  mean <- as.data.frame(Z$mean)
  upper <- as.data.frame(Z$upper)$'80%'
  lower <- as.data.frame(Z$lower)$'80%'
  combo <- cbind.data.frame(mean,upper,lower)
}

#test <- CreateDemandForecasts(OPD$`Breast Surgery`)

#Get The Forecasts
ForecastOutput <- lapply(OPD, CreateDemandForecasts)

#Create The Forecast Dates
forecastindex <- fullmonthmax %m+% months(1:forecastperiod)
names(forecastindex) <- "WeekEnding"
combineforecastsandindex <- function(x) {
  cbind.data.frame(forecastindex,x)
}

#Fix Column Names function
fixnames <- function(x) {
  x <- setNames(x, c("WeekEnding", "Forecast", "80% Upper", "80% Lower"))
}

#Combine the forecasts with the index and fix the names, then make it a Data Frame
ForecastOutput <- lapply(ForecastOutput, combineforecastsandindex) 
ForecastOutput <- lapply(ForecastOutput, fixnames)

combinedforecasts <- bind_rows(ForecastOutput, .id = "BookSpecialty")

#Create function for rolling means
CreateRollingMeans <- function(x,y) {
  DNAAvg <- rollmean(x$NewDNA, y, fill = "extend", align = "right")
  ApptsAvg <- rollmean(x$NewOPD, y, fill = "extend", align = "right")
  OvCapAvg <- rollmean(x$NewDNA, y, fill = "extend", align = "right") + rollmean(x$NewOPD, y, fill = "extend", align = "right")
  cbind.data.frame(x,DNAAvg, ApptsAvg, OvCapAvg)
}


#Create Rolling Means for set period and make a data frame
OPD <- lapply(OPD, CreateRollingMeans, y = rollperiod)
combinedOPD <- bind_rows(OPD)

#combine the rolling means data frame and forecasts
combined <- bind_rows(combinedOPD, combinedforecasts)


#Arrange the combined data frame and rename the date to Month Ending
combined <- arrange(combined, BookSpecialty, WeekEnding)
names(combined)[names(combined) == "WeekEnding"] <- "MonthEnding"


#Roll last available rolling average forward with the predictions
OPDtest <- split.data.frame(combined, combined$BookSpecialty)

OPDtestfunc  <- function(x) {
  DNAAvg <- x$DNAAvg[(nrow(x)-(forecastperiod-1)):nrow(x)] <- x$DNAAvg[(nrow(x)-forecastperiod)]
  ApptsAvg <- x$ApptsAvg[(nrow(x)-(forecastperiod-1)):nrow(x)] <- x$ApptsAvg[(nrow(x)-forecastperiod)]
  OvCapAvg <- x$OvCapAvg[(nrow(x)-(forecastperiod-1)):nrow(x)] <- x$DNAAvg[(nrow(x)-forecastperiod)] + x$ApptsAvg[(nrow(x)-forecastperiod)]
  cbind.data.frame(x[(nrow(x)-(forecastperiod-1)):nrow(x),], DNAAvg, ApptsAvg, OvCapAvg)
}
OPDtest <- lapply(OPDtest, OPDtestfunc)
OPDtest <- bind_rows(OPDtest)

#Filter out the incomplete averages and replace with the last period rolled forward
combined <- combined %>%
  filter(is.na(DNAAvg) == FALSE) %>%
  bind_rows(OPDtest) %>%
  arrange(BookSpecialty, MonthEnding)


#Output it all to an Excel file
write_xlsx(combined, "DiverForecasts.xlsx", col_names = TRUE)




#Create 80% Confidence Forecast Function
CreateNNDemandForecasts <- function(x) {
  y <- converttoMxts(x)
  Z <- forecast(nnetar(y, repeats = NeuralnetIterations), PI = TRUE, h = forecastperiod)
  mean <- as.data.frame(Z$mean)
  upper <- as.data.frame(Z$upper)$'90%'
  lower <- as.data.frame(Z$lower)$'10%'
  combo <- cbind.data.frame(mean,upper,lower)
}

if(RunNeuralNet == TRUE) {
ForecastNNOutput <- lapply(OPD, CreateNNDemandForecasts)

ForecastNNOutput <- lapply(ForecastNNOutput, combineforecastsandindex) 
ForecastNNOutput <- lapply(ForecastNNOutput, fixnames)
 
combinedNNforecasts <- bind_rows(ForecastNNOutput, .id = "BookSpecialty")
 
combinedNN <- bind_rows(combinedOPD, combinedNNforecasts)
 
combinedNN <- arrange(combinedNN, BookSpecialty, WeekEnding)
names(combinedNN)[names(combinedNN) == "WeekEnding"] <- "MonthEnding"

#Roll last available rolling average forward with the predictions
OPDtestNN <- split.data.frame(combinedNN, combinedNN$BookSpecialty)
OPDtestNN <- lapply(OPDtestNN, OPDtestfunc)
OPDtestNN <- bind_rows(OPDtestNN)

combinedNN <- combinedNN %>%
  filter(is.na(DNAAvg) == FALSE) %>%
  bind_rows(OPDtestNN) %>%
  arrange(BookSpecialty, MonthEnding)
# 
write_xlsx(combinedNN, "DiverForecastsNN.xlsx", col_names = TRUE)
}
