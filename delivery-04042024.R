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

# 2.2 Read delivery data from sheet 1 into delivery variable
file_path = "input/Delivery 2018-2023.xlsx"
delivery.1823 = xlsx_cells(file_path, sheets = 1)

##### 3. Extract unstructured dataset with behead #####
# 3.1 extract delivery from untidy data
delivery.1823 = delivery.1823 %>% 
  filter(!is_blank) %>%  
  filter(row >= 2) %>% 
  behead("up-left", indicator) %>% 
  behead("up", year) %>%
  behead("left", code) %>% 
  behead("left", province) %>%  
  mutate ( indicator = "Births Attended by Skilled Health Personnel") %>% 
  select(code,indicator,province,year,value = numeric) %>% 
  left_join(metadata$referencearea, by = join_by(code == province_code)) %>% 
  select(code,indicator,nis_province,year:value)

##### 4. Detect and compute data #####

##### 5 transform them into needed format #####

##### 6. Compute delivery data for CAMSTAT website #####
delivery = delivery.1823 %>%  
  mutate( DATAFLOW = "KH_NIS:DF_ACCESS_TO_HEALTH_SERVICES(1.0)",
          `INDICATOR: Indicator` = str_c(str_replace_all(
            str_to_upper(indicator),"\\s","_"),": ", str_to_title(indicator)),
          `REF_AREA: Reference area` = nis_province,
          `SEX: Sex` = "F: Female",
          `LOCATION: Degree of Urbanisation` = "_Z: NA",
          `AGE_GROUP: Age group` = "_Z: NA",
          `VACCINE: Vaccine` = "_Z: NA", 
          `HEALTH_CARE_PROVIDER: Health Care Provider` = "_Z: NA",
          `CONTRACEPTION: Contraception` = "_Z: NA",
          `WEALTH_QUINTILE: Wealth Quintile` = "_Z: NA",
          `NUMBER_OF_VISITS: Number of Visits` = "_Z: NA",
          `UNIT_MEASURE: Unit of measure` = "NUMBER: Number",
          `FREQ: Frequency` = "A: Annual",
          `TIME_PERIOD: Time period` = year,
          OBS_VALUE = value,
          `UNIT_MULTIPLIER: Unit multiplier` = "0: Units",
          `RESPONSIBLE_AGENCY: Responsible agency` = "MOH: Ministry of Health",
          `DATA_SOURCE: Data source` = "") %>%
  select(DATAFLOW:`DATA_SOURCE: Data source`)

##### 7. Write HSP indicators to directory #####
write_csv(delivery, paste0("output/delivery/", "Births_Attended_by_Skilled_Health_Personnel(1.0)", ".csv"))


