library(tidyverse)
library(tidyxl)
library(unpivotr)
library(readxl)

setwd("D:\\Personal\\2024\\Training-HSP\\NIS-Rscript-04042024")
# metadata will be used in join operation and served as input for data computation
metadata.path <- "input/metadata.xlsx"
metadata = map(excel_sheets(metadata.path), 
               read_excel, path = metadata.path)
names(metadata) = excel_sheets(metadata.path)

# 2.2 Read health service indicator data from sheet 2 into single hsp.indicator variable
file_path = "input/MoH indicator data 2018-2022.xlsx"
hsp.indicator = xlsx_cells(file_path, sheets = 2)

# 3.1 extract hs indicators from untidy data
hsp = hsp.indicator %>% 
  filter(row >= 5) %>% 
  behead("up", year) %>% 
  behead("left", code) %>%
  behead("left", hspindicators)

# 3.2 transform unstructured data into database format
hsp = hsp %>% 
  rename( value = numeric) %>% 
  select(code,hspindicators,year,value) %>% 
  filter(hspindicators != "NA" & value != "NA") %>% 
  right_join(metadata$hspindicator, join_by(code == code)) %>% 
  select(`Health and Nutrition`,camstat_status,camstat_name,Indicator_eng,year,value,agegroup:sex)

# 5.1 Split HSP into list by indicators from its column name
hsp.ind = hsp %>% group_split(`Health and Nutrition`)

# 5.2 Pull unique name of each indicator from the list
name.indicator = hsp.ind %>%
  map(~pull(.,`Health and Nutrition`)) %>%
  map(~unique(.))

# 5.3 Assign the list name with the pulled names
names(hsp.ind) = name.indicator

##### 6. Compute HSP indicators for CAMSTAT website within the list #####
hsp.indicator = hsp.ind %>% 
  map(~ mutate(.,
               DATAFLOW = str_c("KH_NIS:DF_", str_replace_all(str_to_upper(unique(`Health and Nutrition`)),"\\s","_"), "(1.0)"),
               `INDICATOR: Indicator` = str_c(str_replace_all(str_to_upper(camstat_name),"\\s","_"),": ", str_to_title(camstat_name)),
               `REF_AREA: Reference area` = "ASIKHM: Cambodia",
               `SEX: Sex` = sex,
               `LOCATION: Degree of Urbanisation` = "_Z: NA",
               `AGE_GROUP: Age group` = agegroup,
               `VACCINE: Vaccine` = "_Z: NA", 
               `HEALTH_CARE_PROVIDER: Health Care Provider` = "_Z: NA",
               `CONTRACEPTION: Contraception` = "_Z: NA",
               `WEALTH_QUINTILE: Wealth Quintile` = "_Z: NA",
               `NUMBER_OF_VISITS: Number of Visits` = "_Z: NA",
               `UNIT_MEASURE: Unit of measure` = unitmeasure,
               `FREQ: Frequency` = "A: Annual",
               `TIME_PERIOD: Time period` = year,
               OBS_VALUE = value,
               `UNIT_MULTIPLIER: Unit multiplier` = "0: Units",
               `RESPONSIBLE_AGENCY: Responsible agency` = "MOH: Ministry of Health",
               `DATA_SOURCE: Data source` = "") %>% 
        select(DATAFLOW:`DATA_SOURCE: Data source`))

# 7.1 Write list file to multiple CSV files
pmap(list(data = hsp.indicator, names = names(hsp.indicator)),
     function(data, names) { 
       write_csv(data, paste0(
         "output/","KH_NIS_DF_", 
         str_replace_all(str_to_upper(names),"\\s","_"),
         "_1.0_all", ".csv"))})





