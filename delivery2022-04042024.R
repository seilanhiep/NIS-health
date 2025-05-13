##### 1. Call Library #####
library(tidyverse)
library(tidyxl)
library(unpivotr)
library(readxl)

setwd("D:/Personal/2024/Progress report/Final report/Final-NIS-R script")
##### 2. Read Excel files into R #####
# 2.1 Read Multiple sheets from metadata into metadata variable
# metadata will be used in join operation and served as input for data computation
metadata.path <- "input/metadata.xlsx"
metadata = metadata.path %>% 
  excel_sheets() %>%
  set_names() %>% 
  map(read_excel, path = metadata.path)

# 2.2 Read delivery data from sheet 1 into delivery variable
file_path = "input/delivery2022.xlsx"
nation.delivery.2022 = xlsx_cells(file_path, sheets = 1)

##### 3. Extract unstructured dataset with behead #####
# 3.1 extract delivery from untidy data
delivery2022 = nation.delivery.2022 %>% 
  filter(!is_blank) %>%  
  filter(row >= 4) %>% 
  behead("up-left", n1) %>% 
  behead("up-left", n2) %>%
  behead("up", n3) %>%
  behead("left", code) %>% 
  behead("left", province) %>%  
  select(n1,n2,n3,code,province,delivery = numeric) %>% 
  filter(delivery != "NA")

# 3.2 transform unstructured data into database format
delivery2022 = delivery2022 %>% 
  mutate ( name = case_when( 
    !is.na(n1) & !is.na(n2) & !is.na(n3) ~ str_c(n1,n2,n3, sep = "-"),
    !is.na(n1) & !is.na(n2) ~ str_c(n1,n2, sep = "-"),
    !is.na(n1) ~ n1)) %>% 
  select(name,code,province,delivery)

##### 4. Detect and compute data #####
# 4.1 Turn one column into multiple columns
pw.delivery = delivery2022 %>% 
  pivot_wider( names_from = name, values_from = delivery) %>% 
  select(code,province, `EPW`= `Expected Pregnant Women`, starts_with("No. of")) %>%
  replace(is.na(.),0) %>% 
  mutate(`birth attended by health skilled personnel` = rowSums(across(c(4,6,7)), na.rm=TRUE),
         `delivery place` = rowSums(across(c(4,5,6,7)), na.rm=TRUE),
         `Deliveries At Home` = rowSums(across(c(4,5)), na.rm=TRUE),
         `Deliveries At Health Center` = `No. of delivery-at HC`,
         `birth attended by TBAs` = `No. of delivery-at Home by-TBAs`,
         `delivery received  ARV` = `No. of delivery received  ARV`,
         `delivery referred to` = `No. of referred to`,
         `% of birth attended by health skilled personnel` = round((`birth attended by health skilled personnel`/`EPW`)*100,digits = 2),
         `% of delivery place` = round((`delivery place`/`EPW`)*100,digits = 2),
         `% of Deliveries At Home` = round((`Deliveries At Home`/`EPW`)*100,digits = 2),
         `% of Deliveries At Health Center` = round((`Deliveries At Health Center`/`EPW`)*100,
                              digits = 2),
         `% of birth attended by TBAs` = round((`birth attended by TBAs`/`EPW`)*100,digits = 2),
         `% of delivery received  ARV` = round((`delivery received  ARV`/`EPW`)*100,digits = 2),
         `% of delivery referred to` = round((`delivery referred to`/`EPW`)*100,
                                                    digits = 2)) %>% 
  select(code,province,`EPW`,`birth attended by health skilled personnel`:`% of delivery referred to`)

##### 5. transform them into needed format #####
# 5.1 Turn many columns into database format
pl.delivery = pw.delivery %>% 
pivot_longer(!c(code,province), 
             names_to = "indicator_delivery",
             values_to = "value") %>% 
  filter(value != "NA" & indicator_delivery != "EPW") %>% 
  mutate(unitmeasure = case_when(
    str_sub(indicator_delivery,1,1) == "%" ~ "PERCENT: Percent",
    .default = "NUMBER: Number"))

# 5.2 Join dataset
delivery22 = pl.delivery %>% 
  left_join(metadata$referencearea, by = join_by(
    code == province_code)) %>% 
  select(nis_province,indicator_delivery:unitmeasure) %>% 
  filter(nis_province != "NA")

# 5.3 Split HSP into list by indicators from its column name
delivery.ind = delivery22 %>% group_split(indicator_delivery)

# 5.4 Pull unique name of each indicator from the list
name.indicator = delivery.ind %>%
  map(~pull(.,indicator_delivery)) %>%
  map(~unique(.))

# 5.5 Assign the list name with the pulled names
names(delivery.ind) = name.indicator

##### 6. Compute delivery data for CAMSTAT website #####
delivery.2022 = delivery.ind %>%  
  map(~ mutate(., DATAFLOW = "KH_NIS:DF_ACCESS_TO_HEALTH_SERVICES(1.0)",
          `INDICATOR: Indicator` = str_c(str_replace_all(
            str_to_upper(indicator_delivery),"\\s","_"),": ", str_to_title(indicator_delivery)),
          `REF_AREA: Reference area` = nis_province,
          `SEX: Sex` = "_T:Total",
          `LOCATION: Degree of Urbanisation` = "_Z: NA",
          `AGE_GROUP: Age group` = "_Z: NA",
          `VACCINE: Vaccine` = "_Z: NA", 
          `HEALTH_CARE_PROVIDER: Health Care Provider` = "DELIVERIES: Deliveries",
          `CONTRACEPTION: Contraception` = "_Z: NA",
          `WEALTH_QUINTILE: Wealth Quintile` = "_Z: NA",
          `NUMBER_OF_VISITS: Number of Visits` = "_Z: NA",
          `UNIT_MEASURE: Unit of measure` = unitmeasure,
          `FREQ: Frequency` = "A: Annual",
          `TIME_PERIOD: Time period` = 2022,
          OBS_VALUE = value,
          `UNIT_MULTIPLIER: Unit multiplier` = "0: Units",
          `RESPONSIBLE_AGENCY: Responsible agency` = "MOH: Ministry of Health",
          `DATA_SOURCE: Data source` = "") %>% 
        select(DATAFLOW:`DATA_SOURCE: Data source`))

##### 7. Write Delivery 2022 to directory #####
pmap(list(data = delivery.2022, names = names(delivery.2022)),
     function(data, names) { 
       write_csv(data, paste0(
         "output/","KH_NIS_DF_", 
         str_replace_all(str_to_upper(names),"\\s","_"),
         "_1.0_all", ".csv"))})
