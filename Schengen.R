library(tidyverse)
library(readxl)
library(imputeTS)
library(openxlsx)

Schengen_2011 <- read_excel("Cluster/synthese_2011_with_filters_en.xls")
Schengen_2012 <- read_excel("Cluster/synthese_2012_with_filters_en.xlsx")
Schengen_2013 <- read_excel("Cluster/synthese_2013_with_filters_en.xls")
Schengen_2014 <- read_excel("Cluster/2014_global_schengen_visa_stats_compilation_consulates_-_final_en.xlsx")
Schengen_2015 <- read_excel("Cluster/2015_consulates_schengen_visa_stats_en.xlsx")
Schengen_2016 <- read_excel("Cluster/2016_consulates_schengen_visa_stats_en.xlsx")
Schengen_2017 <- read_excel("Cluster/2017-consulates-schengen-visa-stats.xlsx")
Schengen_2018 <- read_excel("Cluster/2018-consulates-schengen-visa-stats.xlsx")
Schengen_2019 <- read_excel("Cluster/2019-consulates-schengen-visa-stats.xlsx")

Schengen_2011$year <- 2011
Schengen_2012$year <- 2012
Schengen_2013$year <- 2013
Schengen_2014$year <- 2014
Schengen_2015$year <- 2015
Schengen_2016$year <- 2016
Schengen_2017$year <- 2017
Schengen_2018$year <- 2018
Schengen_2019$year <- 2019


# 2011 
Schengen_2011 <- Schengen_2011[- c(1, 2072:2085), - 2]

Schengen_2011$new <- NA

order <- c(1:3, 6, 4, 5, 7, 21, 11, 8, 9, 10, 17, 12, 21, 14, 21, 15, 16, 20)
Schengen_2011 <- Schengen_2011[,order]
names(Schengen_2011)[1:20] <- names(Schengen_2018)

Schengen_2011 <- Schengen_2011[order(Schengen_2011$`Schengen State`, Schengen_2011$`Country where consulate is located`, Schengen_2011$Consulate), ] 

# 2012

Schengen_2012 <- Schengen_2012[-c(1, 2025:2036), - 2]

Schengen_2012$new <- NA

order <- c(1:3, 6, 4, 5, 7, 21, 11, 8, 9, 10, 17, 12, 21, 14, 21, 15, 16, 20)

Schengen_2012 <- Schengen_2012[,order]
names(Schengen_2012) <- names(Schengen_2018)
Schengen_2012 <- Schengen_2012[order(Schengen_2012$`Schengen State`, Schengen_2012$`Country where consulate is located`, Schengen_2012$Consulate), ] 

#2013
Schengen_2013 <- Schengen_2013[- c(1992:1998), ]
names(Schengen_2013) <- names(Schengen_2018)
Schengen_2013 <- Schengen_2013[order(Schengen_2013$`Schengen State`, Schengen_2013$`Country where consulate is located`, Schengen_2013$Consulate), ] 

# 2014
Schengen_2014 <- Schengen_2014[- c(1963:1973), ]
Schengen_2014 <- Schengen_2014 %>% rename("Multiple entry uniform visas (MEVs) issued" = "Multiple entry uniform visas (MEVs) issued 1)") %>% rename("Total LTVs issued" = "Total LTVs issued 2)")

Schengen_2014 <- Schengen_2014[order(Schengen_2014$`Schengen State`, Schengen_2014$`Country where consulate is located`, Schengen_2014$Consulate), ] 

# 2015
Schengen_2015 <- Schengen_2015[- c(1943:1952), ]
Schengen_2015 <- Schengen_2015 %>% rename("Multiple entry uniform visas (MEVs) issued" = "Multiple entry uniform visas (MEVs) issued 1)")

#2016-2018

Schengen_2016 <- Schengen_2016[- c(1881:1887), ]
Schengen_2017 <- Schengen_2017[- c(1872:1878), ]
Schengen_2018 <- Schengen_2018[- c(1901:1907), ]
Schengen_2019 <- Schengen_2019[- c(1835:1841), ]


df_Schengen <- dplyr::bind_rows(Schengen_2011, Schengen_2012, Schengen_2013, Schengen_2014, Schengen_2015, Schengen_2016, Schengen_2017, Schengen_2018, Schengen_2019)


df_Schengen <- df_Schengen[,c(20, 1:19 )]

names(df_Schengen) <- c("year",	"schengen_state",	"consulate_country",	"consulate",	"ATV_applications",	"ATV_issued",	"ATV_multiple_issued",	"ATV_not_issued",	"ATV_not_issued_rate",	"uniform_applications",	"uniform_issued",	"uniform_MEV_issued",	"share_MEV_uniform_issued",	"LTV_issued1",	"uniform_not_issued",	"uniform_not_issued_rate",	"ATV_uniform_applications",	"ATV_unifrom_issued_including_LTV",	"ATV_uniform_not_issued",	"ATV_uniform_not_issued_rate")


write.csv(df_Schengen, "Cluster/df_Schengen.csv")
write.xlsx(df_Schengen, "Cluster/df_Schengen_final.xlsx")


df_Schengen <- read.xlsx("Cluster/df_Schengen_final.xlsx")
###Identifying possible outliers / miss-reported observations
a <- df_Schengen %>% 
  filter(share_MEV_uniform_issued > 1 |uniform_not_issued_rate > 1 |ATV_not_issued_rate > 1|ATV_uniform_not_issued_rate > 1)

b <- df_Schengen %>% 
  filter(uniform_not_issued_rate > 1)
c <- df_Schengen %>% 
  filter(ATV_not_issued_rate > 1)
d <- df_Schengen %>% 
  filter(ATV_uniform_not_issued_rate > 1)

write.xlsx(a, "Cluster/df_Schengen_missings.xlsx")


  
