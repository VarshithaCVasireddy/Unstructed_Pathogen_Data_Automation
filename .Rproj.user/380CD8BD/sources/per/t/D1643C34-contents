rm(list = ls())

library(tidyverse)
library(readxl)
library(psych)
library(lubridate)

(inp_files <- list.files("inputs", pattern="*.xlsx", full.names=TRUE))
#From inputs folder, files with .xlsx pattern are pulled

for (i in 1:length(inp_files)) 
{
  
  org_data <- read_excel(inp_files[i]) %>%
    select(ID, Name, `Sample Date`, Population, `Virus/L`) %>%
    drop_na(ID, Name) %>%
    mutate(`Sample Date`=as.Date(`Sample Date`, format="%m/%d/%y")) %>%
    mutate(`Sun Week Start` = floor_date(`Sample Date`, "week", 7))  # 7 means Sunday
  #Calculated Sunday date associated with the Sample date
  
  
  geo_means <- org_data %>%
    group_by(ID, `Sun Week Start`) %>% 
    summarise(`Virus/L` = psych::geometric.mean(`Virus/L`), 
              `LowestDate` = min(`Sample Date`))
  
  get_weekly_geomean_citywise <- function(i, o_geo_means) {
    return(
      o_geo_means$`Virus/L`[
        o_geo_means$ID==org_data$ID[i] &
          o_geo_means$LowestDate==org_data$`Sample Date`[i]][1])
  }
  #Geomean is calculated on Virus/L field for samples with same ID values
  
  org_data$`Weekly Geomean` <- sapply(1:nrow(org_data), 
                                      function(i) 
                                        get_weekly_geomean_citywise(i, geo_means))
  
  
  
  # Type = 0 - Sunday to Saturday
  get_weekly_geomean <- function(org_tbl, type = 0) {
    # By Default: Week - Sunday to Saturday
    week_tagged_data <- org_tbl %>%
      mutate(`Week Start` = floor_date(`Sample Date`, "week", 7)) %>% # 7 means Sunday
      mutate(`Week End` = `Week Start` + days(6)) %>%
      mutate(`Week Mid Date` = `Week Start` + days(3))
    # View(week_tagged_data)
    # Week start date, end date and middle date are calculated
    
    return(week_tagged_data %>%
             drop_na(`Weekly Geomean`) %>%
             group_by(`Week Start`, `Week End`, `Week Mid Date`) %>%
             summarise(`Population Avg` = (sum(`Weekly Geomean`*Population) / sum(Population))) %>%
             ungroup() %>%
             mutate(`Week` = sprintf("%s - %s",
                                     format(`Week Start`, "%m/%d/%Y"),
                                     format(`Week End`, "%m/%d/%Y")),
                    `Week Mid Date` = format(`Week Mid Date`, "%m/%d/%Y")) %>%
             select(`Population Avg`, `Week Mid Date`, `Week`))
  }
  # Population Average is calculated for the samples within a week (Sunday - Saturday)
  

  weekly_pop_avg <- rbind(get_weekly_geomean(org_data, 0)) #For Sun - Sat
  diff <- nrow(org_data) - nrow(weekly_pop_avg)
  weekly_pop_avg[(nrow(weekly_pop_avg) + 1):(nrow(weekly_pop_avg) + diff), ] <- NA
  
  org_data <- org_data %>%
    # arrange(`Sample Date`, ID) %>%
    select(ID, Name, `Sample Date`, `Population`, `Virus/L`, `Weekly Geomean`)
  
  org_data <- cbind(org_data, weekly_pop_avg)
  
  
  out_name <- sub(".xlsx", sprintf("_%s.csv", "stats_processing"), inp_files[i])
  out_name <- sub("inputs", "outputs", out_name)
  # csv file is created after executing the file
  
  write.csv(org_data, out_name, row.names = F)
}
rm(list = ls())