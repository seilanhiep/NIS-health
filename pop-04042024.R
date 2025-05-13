##### 1. Call Library #####
library(readxl)
library(tidyverse)
setwd("D:\\Personal\\2024\\Training-HSP\\NIS-Rscript-04042024")
##### 2. Read Excel files into R #####
# 2.1 Read Multiple sheets from metadata into metadata variable
# metadata will be used in join operation and served as input for data computation
metadata.path <- "input/metadata.xlsx"
metadata = map(excel_sheets(metadata.path), 
               read_excel, path = metadata.path)
names(metadata) = excel_sheets(metadata.path)

# 2.2 Read population data from all three sheets into single pop variable
pop.path = "input/population.xlsx"
pop = map(excel_sheets(pop.path), read_excel, path = pop.path)
names(pop) = excel_sheets(pop.path)

##### 3. Extract ready-to-use population data #####
# 3.1 population by province data
# Remove dot sign from the province column and extract number for code column and then subtract texts for province name
clean.pop = pop$clean.pop %>% 
  mutate(pro.vince = str_remove(Province, "\\."),
         code = as.numeric(str_extract_all(Province,"[0-9]+")),
         province = str_trim(str_sub(pro.vince,
                                     str_length(code)+1,str_length(Province)), 
                             side = c("both")))%>% 
  select(code, province, Year, Age, Gender, Population)

# 3.2 Total population data
# Remove dot sign from the province column and extract number for code column and then subtract texts for province name
total.pop = pop$total.pop %>% 
  mutate(pro.vince = str_remove(Province, "\\."),
         code = as.numeric(str_extract_all(Province,"[0-9]+")),
         province = str_trim(str_sub(pro.vince,str_length(code)+1,
                                     str_length(Province)), side = c("both")))%>% 
  select(code, province, Gender, Year, agegroup = Age, Population)

##### 4. Detect and compute data #####
clan.pop.byage = clean.pop %>% 
  mutate(agegroup = case_when( Age <= 4 ~ "Y0T4: <=04 years",
                               Age >= 5 & Age <= 9 ~ "Y5T9: 05-09 years",
                               Age >= 10 & Age <= 14 ~ "Y10T14: 10-14 years",
                               Age >= 15 & Age <= 19 ~ "Y15T19: 15-19 years",
                               Age >= 20 & Age <= 24 ~ "Y20T24: 20-24 years",
                               Age >= 25 & Age <= 29 ~ "Y25T29: 25-29 years",
                               Age >= 30 & Age <= 34 ~ "Y30T34: 30-34 years",
                               Age >= 35 & Age <= 39 ~ "Y35T39: 35-39 years",
                               Age >= 40 & Age <= 44 ~ "Y40T44: 40-44 years",
                               Age >= 45 & Age <= 49 ~ "Y45T49: 45-49 years",
                               Age >= 50 & Age <= 54 ~ "Y50T54: 50-54 years",
                               Age >= 55 & Age <= 59 ~ "Y55T59: 55-59 years",
                               Age >= 60 & Age <= 64 ~ "Y60T64: 60-64 years",
                               Age >= 65 & Age <= 69 ~ "Y65T69: 65-69 years",
                               Age >= 70 & Age <= 74 ~ "Y70T74: 70-74 years",
                               Age >= 75 & Age <= 79 ~ "Y75T79: 75-79 years",
                               Age >= 80 ~ "YGE80: 80+ years")) %>% 
  select(!c(Age))

##### 5 transform them into needed format #####
## 5.1 Transform population by province data
tran.clean.pop = clan.pop.byage %>% 
group_by(code,province,Gender,Year,agegroup) %>% 
  summarise(pop=sum (Population),.groups='drop_last')

## 5.2 Transform total population data
tran.total.pop = total.pop %>% 
  group_by(code,province,Gender,Year) %>% 
  summarise( agegroup = "TOTAL: Total", 
             pop = sum (Population), .groups = 'drop_last')

## 5.3 Combine population by province and total population data
# 5.3.1 integrate both data into one pop file
all.pop = rbind(tran.clean.pop, tran.total.pop)

# 5.3.2 Link CAMSTAT NIS-province format to final pop data
link.pop = inner_join(all.pop, metadata$referencearea, 
                      join_by(code == province_code)) %>% 
  ungroup() %>% 
  select(code,nis_province,Gender:pop)

##### 6. Compute population data for CAMSTAT website #####
final.pop = link.pop %>%  
  mutate( DATAFLOW = "KH_DCX:DF_POPULATION(1.0)",
          `INDICATOR: Indicator` = str_c("URBAN POPULATION: ",
                                         "Urban Population"),
          `REF_AREA: Reference area` = nis_province,
          `SEX: Sex` = str_replace_all(Gender, 
                                       c("Female" = "F: Female",
                                         "Male" = "M: Male", "Total" = "_T: Total")),
          `LOCATION: Degree of Urbanisation` = "_Z: NA",
          `AGE_GROUP: Age group` = agegroup,
          `LAND_USE: Land use` = "_Z: NA",
          `RELIGION: Religion` = "_Z: NA",
          `FREQ: Frequency` = "A: Annual",
          `TIME_PERIOD: Time period` = Year,
          OBS_VALUE = pop,
          `UNIT_MULTIPLIER: Unit multiplier` = "NUMBER: Number",
          `RESPONSIBLE_AGENCY: Responsible agency` = paste0("MOP: Ministry of Planning", "NIS: National Institute of Statistics, ","MOH: Ministry of Health"),
          `DATA_SOURCE: Data source` = paste0("ds_health: MoH_National Health Statistics ", Year,"_", Year+1)) %>%   
  select(DATAFLOW:`DATA_SOURCE: Data source`)

##### 7. Write pop data to directory #####
write_csv(final.pop, paste0("output/", "KH_NIS_DF_POPULATION_1.1_all", ".csv"))


