# Unstructed Data Automation
## Author: Varshitha Choudary Vasireddy
## Description of the project:
Wastewater or sewage water is treated to get the concentrations of different pathogens like COVID, Norovirus, Influenza A, Influenza B, Camplobacteria etc. These concentrations are to be reported on Arc Gis dashboards. For statistical infographic presentation on dashboard the large unstructed pathogens data that is received from the lab is to be processed to get structed data and then statistical processing is to be done. So the operations on the file is divided into 2 tasks  
1. Large Unstructed pathogens data file to be converted into structed pathogen specific data file
2. Structured pathogen specific data file is to be converted into statistically processed data file

These tasks are to be performed on daily basis onto 9 pathogens and 18 different files, so to reduce the manual work an automation program is developed in R programming language.

## Automation Programs description
Two automation programs are written to handle the tasks. They are as below  
1. Unstructured data automation.R ---> Used to get the structed pathogen specific data file
2. Statistical processing automation.R ---> Used to get statistically processed data file

### 1. Unstructed data automation
#### **Functionality**
From the large unstructed pathogens data file, ID, Name, Sample Date and Virus/L are to be pulled. From the Sample Name of unstructed data the abbreviation of sampled city is collected and sample date is collected. From the abbreviation the sampled city name is determined and for that specific city the ID is assigned. Geo Mean column of the first sample in the given 3 samples is consider as Virus/L concentration for the specific pathogen. The unstructed data consists of many pathogens data, so excels are created for all the pathogens that we want to pull the data. And for each region for every pathogen excel sheets are created.
#### **Working process of the code**
The unstructured data file is stored in inputs folder. When the program is executed all .xls files in the inputs are executed and for the given pathogens .xlsx excel files with different region sheets are created. This can be seen in outputs folder. In the outputs folder the excel file is named as the same file for which it is processed additionally added pathogen name to it's name.  
For example N1_BCV.xls is named as N1_BCV_N1.xlsx for N1 pathogen file
#### **Advantages**
- Multiple files can be processed at same time and it reduces manual work hours.
- Used this program at my Research Assistantship work and this reduced 9 hours per week manual work.

### 2. Statistical processing automation
#### **Functionality**
From the structed pathogens data file, statistical operations like weekly average is to be calculated for all the cities and for cities divided by region for statistical infographical presentation on the dashboard. The file consists of ID, Name, population and Virus/L so statewide average i.e average for all cities is calculated for every week (Sunday to Saturday) along with regional averages i.e averages for cities present in that particular region. Mid date of the week is used to represent the averages in the dashboard. So averages, mid date of the week and the week start and end date are calculated and inserted into the file using R programming. Weekly average for each city is calculated by taking geomean of that particular city that occured in a week, a city can be sampled twice a week, so geomean of the 2 samples is taken. 

$Statewide Average$ = $\frac {sum(cities weekly average \times population of cities)} {sum(population of cities in that week)}$

#### **Working process of the code**
The structed data file is stored in inputs_structed folder. When the program is executed all .xlsx files in the inputs_structed folder are executed and .csv files are created with the statistical information. This can be seen in outputs_structed folder. In the outputs folder the excel file is named as the same file for which it is processed additionally added stats_processing to it's name.  
For example COVID region1.xlsx is named as COVID region1_stats_processing.csv

#### **Advantages**
- Multiple files can be processed at the same time and it reduces manual work hours.
- Used this program at my Research Assistantship work and this reduced 6 hours per week manual work.
- For files which take 12 hours of manual work for statistical processing can now be prepared in seconds.

### Dashboard Link:
https://uok.maps.arcgis.com/apps/dashboards/51657c21386d4f1a962b1853c76ec589