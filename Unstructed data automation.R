library(tidyverse)
library(readxl)
library(openxlsx)
# install.packages("Dict")
library(Dict)

# Function to assign ID based on the abbreviation
Abbr_to_ID <- function (Abbr){
  return(recode(Abbr, .default = NaN, 
                `OKC SC` = 2,`OKC NC` = 1,`OKC DC` = 4,`OKC CC` = 3,
                TUN = 8, TUS = 9,TUH = 7,MWC = 6, NMI = 5, ANA = 10,TUT = 14,
                SAL = 15,ENID = 16,PRY = 17,DUR = 18,ALT = 19,SEM = 20,BAR = 21,
                MIA = 22,MUSK = 23,CLA = 24,ADA = 25,LAW = 26,HUGO = 27,
                HUG = 27,ELK = 28,CUSH = 29,TAL = 30,CLINT = 31,STILL = 32,
                CLNT = 31, ANT = 33, WOOD = 34, GUY = 35))
}

# Function to assign Name based on the abbreviation
Abbr_to_Name <- function (Abbr) {
  return(recode(Abbr, .missing = NULL,
                'OKC SC' = 'OKC South Canadian','OKC NC' = 'OKC North Canadian',
                'OKC DC' = 'OKC Deer Creek','OKC CC' = 'OKC Chisholm Creek',
                TUN = 'Tulsa North' ,TUS = 'Tulsa South' ,TUH = 'Tulsa Haikey Creek' ,
                MWC = 'Midwest City' ,NMI = 'Norman' ,ANA = 'Anadarko' ,
                TUT = 'Tuttle' ,SAL = 'Sallisaw' ,ENID = 'Enid' ,PRY = 'Pryor' ,
                DUR = 'Durant' ,ALT = 'Altus' ,SEM = 'Seminole' ,BAR = 'Bartlesville' ,
                MIA = 'Miami' ,MUSK = 'Muskogee' ,CLA = 'Claremore' ,ADA = 'Ada' ,
                LAW = 'Lawton' ,ELK = 'Elk City' ,CUSH = 'Cushing' , HUGO = 'Hugo', HUG = 'Hugo',
                TAL = 'Tahlequah' ,CLINT = 'Clinton', CLNT = 'Clinton' ,STILL = 'Stillwater', 
                STIL = 'Stillwater',ANT = 'Antlers',WOOD = 'Woodward', GUY = 'Guymon'))
}

# Function to assign Population based on the abbreviation
Abbr_to_Pop <- function (Abbr) {
  return(recode(Abbr, .default = NaN,
                `OKC SC` = 40017,`OKC NC` = 442168,`OKC DC` = 107331,`OKC CC` = 78106,
                TUN = 181670,TUS = 147101,TUH = 112025,MWC = 57288,NMI = 122837,ANA = 6623,
                TUT = 7294,ENID = 49583,PRY = 9253,SAL = 8410,DUR = 18589,ALT = 19813,
                SEM = 7488,BAR = 36495,MIA = 13176,MUSK = 36790,CLA = 19419,ADA = 16738,
                LAW = 91055,HUGO = 5174,HUG = 5174, ELK = 11561,CUSH = 8184,
                TAL = 16463,CLINT = 8380, CLNT = 8380,STILL = 48134, ANT = 2175, WOOD = 11998,
                GUY = 12561))
}

#Dictionary is created with key as Region names and values as the city names
Regions <- Dict$new(
  `Region 1` = c('Enid','Guymon','Clinton ','Elk City','Woodward'),
  `Region 2` = c('Bartlesville','Stillwater','Cushing','Miami','Pryor','Claremore'),
  `Region 3` = c('Tuttle','Anadarko','Lawton','Ardmore','Ada','Altus'),
  `Region 4` = c('Sallisaw','Tahlequah','Muskogee'),
  `Region 5` = c('Seminole','McAlester','Talihina','Durant','Hugo','Antlers'),
  `Region 6` = c('Yukon','Norman'),
  `Region 7` = c('Tulsa North','Tulsa South','Tulsa Haikey Creek'),
  `Region 8` = c('OKC South Canadian','OKC North Canadian',
                 'OKC Deer Creek','OKC Chisolm Creek','Midwest City'),
  `Small Cities` = c('Antlers','Hugo','Anadarko','Tuttle','Seminole','Cushing','Clinton','Sallisaw','Pryor'),
  `Medium Cities` = c('Elk City','Woodward','Guymon','Miami','Tahlequah','Ada','McAlester','Durant',
                      'Claremore','Altus','Yukon','Ardmore','Bartlesville','Muskogee','OKC South Canadian','Stillwater','Enid'),
  `Large Cities` = c('Midwest City','OKC Chisolm Creek','Lawton','OKC Deer Creek','Tulsa Haikey Creek','Norman','Tulsa South',
                     'Tulsa North','OKC North Canadian')
)


pathogens <- c("N1")
#COVID pathogen values are to be pulled

(inp_files <- list.files("inputs", pattern="*.xls$", full.names=TRUE))
#From inputs folder, files with .xls pattern are pulled

for (i in 1:length(inp_files)) {
  raw_data <- read_excel(inp_files[i])
  
  for (j in 1:length(pathogens)) {
    data <- raw_data %>%
      select(`Sample Name`, `Target Name`, `Geo Mean`) %>%
      na.omit() %>%
      filter(`Target Name` == pathogens[j], !grepl("passive", `Sample Name`)) %>% #Sample
      #names with passive are to be removed
      mutate(`Abbr`= trimws(str_extract(toupper(`Sample Name`), "[A-Z0-9]* ?[A-Z]*")),
             # Abbreviation of Sample name are extracted
             `Sample Date`= trimws(str_extract(`Sample Name`, " \\d{1,2}/\\d{1,2}/\\d{2,4}")),
             # Sample date in sample name are extracted
             `Virus/L` = `Geo Mean`,
             `Type`= trimws(str_extract(`Sample Name`, " REDO ")),
             `Pathogen` = pathogens[j]) %>%
      mutate(`ID`= Abbr_to_ID(`Abbr`),
             `Name` = Abbr_to_Name(`Abbr`),
             `Population` = Abbr_to_Pop(`Abbr`),
             `Sample Date`=as.Date(`Sample Date`, format="%m/%d/%y")) %>%
      select(ID, Name, `Sample Date`, `Population`, `Virus/L`, `Type`, `Pathogen`) %>%
      filter(!is.na(`ID`)) %>%
      arrange(`Sample Date`) %>%
      as_data_frame()
    
    #data <- data[order(as.Date(data$`Sample Date`, format="%m/%d/%Y")), ]
    data <- data[order(data$`Sample Date`),]
    # Data is kept in ascending order of sample date
    
    if(nrow(data) > 0)
    {
      out_name <- sub(".xls", sprintf("_%s.xlsx", pathogens[j]), inp_files[i])
      out_name <- sub("inputs", "outputs", out_name)
      #write.csv(data, out_name, row.names = F)
      sheets.in.file <- 0
      wb <- openxlsx::createWorkbook("Varshitha C Vasireddy")
      openxlsx::addWorksheet(wb, "All Regions")
      sheets.in.file <- sheets.in.file + 1
      openxlsx::writeData(wb, sheet = sheets.in.file, data, rowNames = FALSE)
      
      # Dividing data into Sheets
      for (i in 1:length(Regions$keys)) {
        regionName <- Regions$keys[i]
        regionSubNames <- Regions[regionName]
        
        regionData <- data %>%
          filter(Name %in% regionSubNames) %>%
          as_data_frame()
        
        if(nrow(regionData) > 0) {
          openxlsx::addWorksheet(wb, regionName)
          sheets.in.file <- sheets.in.file + 1
          openxlsx::writeData(wb, sheet = sheets.in.file, regionData, rowNames = FALSE)
        }
      }
      
      openxlsx::saveWorkbook(wb, out_name, overwrite = TRUE)
    }
  }
}

rm(list = ls())