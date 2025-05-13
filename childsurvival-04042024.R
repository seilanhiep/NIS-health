##### 1. Call Library #####
library(tidyverse)
library(tidyxl)
library(unpivotr)
library(readxl)

##### 2. Read Excel files into R #####
# 2.1 Read Multiple sheets from metadata into metadata variable
# metadata will be used in join operation and served as input for data computation
metadata.path <- "input/metadata.xlsx"
metadata = metadata.path %>% 
  excel_sheets() %>%
  set_names() %>% 
  map(read_excel, path = metadata.path)

# 2.2 Read child survival data from sheet 1 into childsurvival variable
child.path = "input/MoH indicator data 2018-2022.xlsx"
child.indicator = xlsx_cells(child.path, sheets = 1)

##### 3. Extract unstructured dataset with behead #####
# 3.1 extract child survival from untidy data
children = child.indicator %>% 
  filter(!is_blank) %>% 
  behead("up", year) %>% 
  behead("left", indicatorkh) %>% 
  behead("left", indicatoreng) %>% 
  rename(value = numeric) %>% 
  select(indicatoreng,year,value) %>% 
  filter(year != "2018-2022")

##### 4. Detect and compute data #####

##### 5 transform them into needed format #####
# 5.1 Split HSP into list by indicators from its column name
cs.ind = children %>% group_split(indicatoreng)

# 5.2 Pull unique name of each indicator from the list
name.cs = cs.ind %>%
  map(~pull(.,indicatoreng)) %>%
  map(~unique(.))

# 5.3 Assign the list name with the pulled names
names(cs.ind) = name.cs

##### 6. Compute delivery data for CAMSTAT website #####
childsurvival = cs.ind %>%  
  map(~ mutate(., DATAFLOW = "KH_NIS:DF_CHILD_SURVIVAL(1.0)",
          `INDICATOR: Indicator` = str_c(str_replace_all(
            str_to_upper(indicatoreng),"\\s","_"),": ", str_to_title(indicatoreng)),
          `REF_AREA: Reference area` = "ASIKHM: Cambodia",
          `SEX: Sex` = case_when(str_detect(str_to_lower(indicatoreng),"mother") ~ "F: Female",
                                 .default = "_T:Total"),
          `LOCATION: Degree of Urbanisation` = "_Z: NA",
          `AGE_GROUP: Age group` = case_when(str_detect(str_to_lower(indicatoreng),"mother") ~ "Y15T49: 15-49 years",
                                              .default = "M0: <01 month"),
          `VACCINE: Vaccine` = "_Z: NA", 
          `HEALTH_CARE_PROVIDER: Health Care Provider` = "DELIVERIES: Deliveries",
          `CONTRACEPTION: Contraception` = "_Z: NA",
          `WEALTH_QUINTILE: Wealth Quintile` = "_Z: NA",
          `NUMBER_OF_VISITS: Number of Visits` = "_Z: NA",
          `UNIT_MEASURE: Unit of measure` = "PERCENT: Percent",
          `FREQ: Frequency` = "A: Annual",
          `TIME_PERIOD: Time period` = year,
          OBS_VALUE = value,
          `UNIT_MULTIPLIER: Unit multiplier` = "0: Units",
          `RESPONSIBLE_AGENCY: Responsible agency` = "MOH: Ministry of Health",
          `DATA_SOURCE: Data source` = "") %>% 
        select(DATAFLOW:`DATA_SOURCE: Data source`))

##### 7. Write HSP indicators to directory #####
pmap(list(data = childsurvival, names = names(childsurvival)),
     function(data, names) { 
       write_csv(data, paste0(
         "output/childsurvival/","KH_NIS_DF_", 
         str_replace_all(str_to_upper(names),"\\s","_"),
         "_1.0_all", ".csv")) } )
