############################################################################# 
# Oil and Gas Monthly Tracking Reports 
##############################################################################

#Use this program to produce monthly tracking reports
#First, set the working directory
setwd("//dfssrv2/ACD/DataAnalytics/Projects/County volume report")

#Load libraries to read and write .xlsx files. Make sure that data is saved as .xlsx
library(readxl)
library(writexl)
library(openxlsx)

##################################################################################################################################################
#read data
#Download this file each month and change the name below
#Change file name
TAP <- read_xlsx("//dfssrv2/ACD/DataAnalytics/Projects/County volume report/County Vol_Val By Filing Period 2023_10_17 CY14-CY23.xlsx")
Legacy_ONGARD <- read_xlsx("//dfssrv2/ACD/DataAnalytics/Projects/County volume report/HISTORICAL ONGARD LEGACY DATA.xlsx")


# merge two data frames by ID
County_vol_val <- rbind(Legacy_ONGARD,TAP)

#Format Filing Period Variable as dates
County_vol_val$`Filing Period` <- as.Date(County_vol_val$`Filing Period`, origin = "1960-01-01")

#################################################################################################################################################

#################################################################################################################################################
#filter data by basin
Permian <- subset(County_vol_val, County_vol_val$Basin=="Permian")
San_Juan <- subset(County_vol_val, County_vol_val$Basin=="San Juan")
Raton <- subset(County_vol_val, County_vol_val$Basin=="Raton")
Bravo_Dome <- subset(County_vol_val, County_vol_val$Basin=="Bravo Dome")
##################################################################################################################################################


##################################################################################################################################################
################################################
# FILTER BY FEDERAL LAND FOR TRACKING PURPOSES #
################################################

#filter by federal land
Fed_Land <- subset(County_vol_val, County_vol_val$`Land Type`=="Federal")
Fed_Oil <- subset(Fed_Land, Fed_Land$`Product Code`=="01"
                  | Fed_Land$`Product Code`=="02" | Fed_Land$`Product Code`=="14"
                  | Fed_Land$`Product Code`=="05")

fed_oil_vol <- aggregate(x = Fed_Oil[c("Volume")],
                         FUN = sum,
                         by = list("Filing Period" = Fed_Oil$`Filing Period`))
fed_oil_val <- aggregate(x = Fed_Oil[c("Gross Value")],
                         FUN = sum,
                         by = list("Filing Period" = Fed_Oil$`Filing Period`))
fed_oil_dedct <- aggregate(x = Fed_Oil[c("Total Deductions")],
                           FUN = sum,
                           by = list("Filing Period" = Fed_Oil$`Filing Period`))
fed_oil_Transdedct <- aggregate(x = Fed_Oil[c("Transportation Deduction")],
                                FUN = sum,
                                by = list("Filing Period" = Fed_Oil$`Filing Period`))
fed_oil_Processdedct <- aggregate(x = Fed_Oil[c("Processing Deduction")],
                                  FUN = sum,
                                  by = list("Filing Period" = Fed_Oil$`Filing Period`))
fed_oil_Royaltydedct <- aggregate(x = Fed_Oil[c("Royalty Deduction")],
                                  FUN = sum,
                                  by = list("Filing Period" = Fed_Oil$`Filing Period`))


Fed_Oil_comb <- Reduce(function(x, y) merge(x, y, all=TRUE), list(fed_oil_vol, fed_oil_val, fed_oil_dedct, fed_oil_Transdedct, fed_oil_Processdedct, fed_oil_Royaltydedct))
Fed_Oil_comb$`Deduction Percent` <- Fed_Oil_comb$`Total Deductions`/Fed_Oil_comb$`Gross Value`
Fed_Oil_comb$`Gross Price` <- Fed_Oil_comb$`Gross Value`/Fed_Oil_comb$Volume
Fed_Oil_comb$`Net Value` <- Fed_Oil_comb$`Gross Value`-Fed_Oil_comb$`Total Deductions`

#####################################################################################################################


########################################################################################################################################################################################
##############################################
# FILTER BY STATE LAND FOR TRACKING PURPOSES #
##############################################
#filter by State land
State_Land <- subset(County_vol_val, County_vol_val$`Land Type`=="State")
State_Oil <- subset(State_Land, State_Land$`Product Code`=="01"
                    | State_Land$`Product Code`=="02" | State_Land$`Product Code`=="14"
                    | State_Land$`Product Code`=="05")

state_oil_vol <- aggregate(x = State_Oil[c("Volume")],
                           FUN = sum,
                           by = list("Filing Period" = State_Oil$`Filing Period`))
state_oil_val <- aggregate(x = State_Oil[c("Gross Value")],
                           FUN = sum,
                           by = list("Filing Period" = State_Oil$`Filing Period`))
state_oil_dedct <- aggregate(x = State_Oil[c("Total Deductions")],
                             FUN = sum,
                             by = list("Filing Period" = State_Oil$`Filing Period`))
state_oil_Transdedct <- aggregate(x = State_Oil[c("Transportation Deduction")],
                                  FUN = sum,
                                  by = list("Filing Period" = State_Oil$`Filing Period`))
state_oil_Processdedct <- aggregate(x = State_Oil[c("Processing Deduction")],
                                    FUN = sum,
                                    by = list("Filing Period" = State_Oil$`Filing Period`))
state_oil_Royaltydedct <- aggregate(x = State_Oil[c("Royalty Deduction")],
                                    FUN = sum,
                                    by = list("Filing Period" = State_Oil$`Filing Period`))

state_Oil_comb <- Reduce(function(x, y) merge(x, y, all=TRUE), list(state_oil_vol, state_oil_val, state_oil_dedct, state_oil_Transdedct, state_oil_Processdedct, state_oil_Royaltydedct))
state_Oil_comb$`Deduction Percent` <- state_Oil_comb$`Total Deductions`/state_Oil_comb$`Gross Value`
state_Oil_comb$`Gross Price` <- state_Oil_comb$`Gross Value`/state_Oil_comb$Volume
state_Oil_comb$`Net Value` <- state_Oil_comb$`Gross Value`-state_Oil_comb$`Total Deductions`

#######################################################################################################################################################################################################

## Natural gas by federal and state lands
State_gas <- subset(State_Land, State_Land$`Product Code`=="07"
                    | State_Land$`Product Code`=="03" | State_Land$`Product Code`=="04")
state_gas_vol <- aggregate(x = State_gas[c("Volume")],
                           FUN = sum,
                           by = list("Filing Period" = State_gas$`Filing Period`))
state_gas_val <- aggregate(x = State_gas[c("Gross Value")],
                           FUN = sum,
                           by = list("Filing Period" = State_gas$`Filing Period`))
state_gas_dedct <- aggregate(x = State_gas[c("Total Deductions")],
                             FUN = sum,
                             by = list("Filing Period" = State_gas$`Filing Period`))
state_gas_Transdedct <- aggregate(x = State_gas[c("Transportation Deduction")],
                                  FUN = sum,
                                  by = list("Filing Period" = State_gas$`Filing Period`))
state_gas_Processdedct <- aggregate(x = State_gas[c("Processing Deduction")],
                                    FUN = sum,
                                    by = list("Filing Period" = State_gas$`Filing Period`))
state_gas_Royaltydedct <- aggregate(x = State_gas[c("Royalty Deduction")],
                                    FUN = sum,
                                    by = list("Filing Period" = State_gas$`Filing Period`))

state_gas_comb <- Reduce(function(x, y) merge(x, y, all=TRUE), list(state_gas_vol, state_gas_val, state_gas_dedct, state_gas_Transdedct, state_gas_Processdedct, state_gas_Royaltydedct))
state_gas_comb$`Deduction Percent` <- state_gas_comb$`Total Deductions`/state_gas_comb$`Gross Value`
state_gas_comb$`Gross Price` <- state_gas_comb$`Gross Value`/state_gas_comb$Volume
state_gas_comb$`Net Value` <- state_gas_comb$`Gross Value`-state_gas_comb$`Total Deductions`

#####################################################################################################################################################################################################
# Federal land natural gas tracking

fed_gas <- subset(Fed_Land, Fed_Land$`Product Code`=="07"
                  | Fed_Land$`Product Code`=="03" | Fed_Land$`Product Code`=="04")
fed_gas_vol <- aggregate(x = fed_gas[c("Volume")],
                         FUN = sum,
                         by = list("Filing Period" = fed_gas$`Filing Period`))
fed_gas_val <- aggregate(x = fed_gas[c("Gross Value")],
                         FUN = sum,
                         by = list("Filing Period" = fed_gas$`Filing Period`))
fed_gas_dedct <- aggregate(x = fed_gas[c("Total Deductions")],
                           FUN = sum,
                           by = list("Filing Period" = fed_gas$`Filing Period`))
fed_gas_Transdedct <- aggregate(x = fed_gas[c("Transportation Deduction")],
                                FUN = sum,
                                by = list("Filing Period" = fed_gas$`Filing Period`))
fed_gas_Processdedct <- aggregate(x = fed_gas[c("Processing Deduction")],
                                  FUN = sum,
                                  by = list("Filing Period" = fed_gas$`Filing Period`))
fed_gas_Royaltydedct <- aggregate(x = fed_gas[c("Royalty Deduction")],
                                  FUN = sum,
                                  by = list("Filing Period" = fed_gas$`Filing Period`))

fed_gas_comb <- Reduce(function(x, y) merge(x, y, all=TRUE), list(fed_gas_vol, fed_gas_val, fed_gas_dedct, fed_gas_Transdedct, fed_gas_Processdedct, fed_gas_Royaltydedct))
fed_gas_comb$`Deduction Percent` <- fed_gas_comb$`Total Deductions`/fed_gas_comb$`Gross Value`
fed_gas_comb$`Gross Price` <- fed_gas_comb$`Gross Value`/fed_gas_comb$Volume
fed_gas_comb$`Net Value` <- fed_gas_comb$`Gross Value`-fed_gas_comb$`Total Deductions`


############################################################################################################################################

#filter by commodity
#OIL
Oil_State <- subset(County_vol_val, County_vol_val$`Product Code`=="01"
                    | County_vol_val$`Product Code`=="02" | County_vol_val$`Product Code`=="14"
                    | County_vol_val$`Product Code`=="05")
#NATURAL GAS
NG_State <- subset(County_vol_val, County_vol_val$`Product Code`=="07"
                   | County_vol_val$`Product Code`=="03"
                   | County_vol_val$`Product Code`=="04")
#C02
CO2_State <- subset(County_vol_val, County_vol_val$`Product Code`=="17")
#HELIUM
Helium_State <- subset(County_vol_val, County_vol_val$`Product Code`=="08")

##############################################################################################################################################

# filter data by basin and product type
# Permian Basin
Permian_Oil <- subset(Permian, Permian$`Product Code`=="01"
                      | Permian$`Product Code`=="02" | Permian$`Product Code`=="14"
                      | Permian$`Product Code`=="05")
Permian_NatGas <- subset(Permian, Permian$`Product Code`=="07" | 
                           Permian$`Product Code`=="03" | 
                           Permian$`Product Code`=="04" )
Permian_CO2 <- subset(Permian, Permian$`Product Code`=="17")
Permian_Helium <- subset(Permian, Permian$`Product Code`=="08")
#San Juan Basin
San_Juan_Oil <- subset(San_Juan, San_Juan$`Product Code`=="01"
                       | San_Juan$`Product Code`=="02" | San_Juan$`Product Code`=="14"
                       | San_Juan$`Product Code`=="05")
San_Juan_NatGas <- subset(San_Juan, San_Juan$`Product Code`=="07" | 
                            San_Juan$`Product Code`=="03" | 
                            San_Juan$`Product Code`=="04")
San_Juan_CO2 <- subset(San_Juan, San_Juan$`Product Code`=="17")
San_Juan_Helium <- subset(San_Juan, San_Juan$`Product Code`=="08")
#Raton Basin
Raton_Oil <- subset(Raton, Raton$`Product Code`=="01"
                    | Raton$`Product Code`=="02" | Raton$`Product Code`=="14"
                    | Raton$`Product Code`=="05")
Raton_NatGas <- subset(Raton, Raton$`Product Code`=="07" | 
                         Raton$`Product Code`=="03" | 
                         Raton$`Product Code`=="04")
Raton_CO2 <- subset(Raton, Raton$`Product Code`=="17")
Raton_Helium <- subset(Raton, Raton$`Product Code`=="08")
#Bravo Dome Basin
Bravo_Dome_Oil <- subset(Bravo_Dome, Bravo_Dome$`Product Code`=="01"
                         | Bravo_Dome$`Product Code`=="02" | Bravo_Dome$`Product Code`=="14"
                         | Bravo_Dome$`Product Code`=="05")
Raton_NatGas <- subset(Bravo_Dome, Bravo_Dome$`Product Code`=="07" | 
                         Bravo_Dome$`Product Code`=="03" | 
                         Bravo_Dome$`Product Code`=="04")
Bravo_Dome_CO2 <- subset(Bravo_Dome, Bravo_Dome$`Product Code`=="17")
Bravo_Dome_Helium <- subset(Bravo_Dome, Bravo_Dome$`Product Code`=="08")

# sum data by filing period and produce excel tracking worksheets
#### OIL
#Permian Basin
Perm_oil_vol_sum <- aggregate(x = Permian_Oil[c("Volume")],
                              FUN = sum,
                              by = list("Filing Period" = Permian_Oil$`Filing Period`))
Perm_oil_val_sum <- aggregate(x = Permian_Oil[c("Gross Value")],
                              FUN = sum,
                              by = list("Filing Period" = Permian_Oil$`Filing Period`))
Perm_oil_dedct_sum <- aggregate(x = Permian_Oil[c("Total Deductions")],
                                FUN = sum,
                                by = list("Filing Period" = Permian_Oil$`Filing Period`))
Perm_oil_Transdedct_sum <- aggregate(x = Permian_Oil[c("Transportation Deduction")],
                                     FUN = sum,
                                     by = list("Filing Period" = Permian_Oil$`Filing Period`))
Perm_oil_Processdedct_sum <- aggregate(x = Permian_Oil[c("Processing Deduction")],
                                       FUN = sum,
                                       by = list("Filing Period" = Permian_Oil$`Filing Period`))
Perm_oil_Royaltydedct_sum <- aggregate(x = Permian_Oil[c("Royalty Deduction")],
                                       FUN = sum,
                                       by = list("Filing Period" = Permian_Oil$`Filing Period`))

Perm_Oil_comb <- Reduce(function(x, y) merge(x, y, all=TRUE), list(Perm_oil_vol_sum, Perm_oil_val_sum, Perm_oil_dedct_sum))
Perm_Oil_comb$`Deduction Percent` <- Perm_Oil_comb$`Total Deductions`/Perm_Oil_comb$`Gross Value`
Perm_Oil_comb$`Gross Price` <- Perm_Oil_comb$`Gross Value`/Perm_Oil_comb$Volume
Perm_Oil_comb$`Net Value` <- Perm_Oil_comb$`Gross Value`-Perm_Oil_comb$`Total Deductions`

#rename columns
names(Perm_Oil_comb)[names(Perm_Oil_comb) == "Volume"] <- "Permian Volume"
names(Perm_Oil_comb)[names(Perm_Oil_comb) == "Gross Value"] <- "Permian Gross Value"
names(Perm_Oil_comb)[names(Perm_Oil_comb) == "Total Deductions"] <- "Permian Total Deductions"
names(Perm_Oil_comb)[names(Perm_Oil_comb) == "Deduction Percent"] <- "Permian Deduction Percent"
names(Perm_Oil_comb)[names(Perm_Oil_comb) == "Gross Price"] <- "Permian Gross Price"
names(Perm_Oil_comb)[names(Perm_Oil_comb) == "Net Value"] <- "Permian Net Value"

#San Juan Basin
SJ_oil_vol_sum <- aggregate(x = San_Juan_Oil[c("Volume")],
                            FUN = sum,
                            by = list("Filing Period" = San_Juan_Oil$`Filing Period`))
SJ_oil_val_sum <- aggregate(x = San_Juan_Oil[c("Gross Value")],
                            FUN = sum,
                            by = list("Filing Period" = San_Juan_Oil$`Filing Period`))
SJ_oil_dedct_sum <- aggregate(x = San_Juan_Oil[c("Total Deductions")],
                              FUN = sum,
                              by = list("Filing Period" = San_Juan_Oil$`Filing Period`))
SJ_oil_Transdedct_sum <- aggregate(x = San_Juan_Oil[c("Transportation Deduction")],
                                   FUN = sum,
                                   by = list("Filing Period" = San_Juan_Oil$`Filing Period`))
SJ_oil_Processdedct_sum <- aggregate(x = San_Juan_Oil[c("Processing Deduction")],
                                     FUN = sum,
                                     by = list("Filing Period" = San_Juan_Oil$`Filing Period`))
SJ_oil_Royaltydedct_sum <- aggregate(x = San_Juan_Oil[c("Royalty Deduction")],
                                     FUN = sum,
                                     by = list("Filing Period" = San_Juan_Oil$`Filing Period`))

SJ_Oil_comb <- Reduce(function(x, y) merge(x, y, all=TRUE), list(SJ_oil_vol_sum, SJ_oil_val_sum, SJ_oil_dedct_sum))
SJ_Oil_comb$`Deduction Percent` <- SJ_Oil_comb$`Total Deductions`/SJ_Oil_comb$`Gross Value`
SJ_Oil_comb$`Gross Price` <- SJ_Oil_comb$`Gross Value`/SJ_Oil_comb$Volume
SJ_Oil_comb$`Net Value` <- SJ_Oil_comb$`Gross Value`-SJ_Oil_comb$`Total Deductions`

#rename columns
names(SJ_Oil_comb)[names(SJ_Oil_comb) == "Volume"] <- "SJ Volume"
names(SJ_Oil_comb)[names(SJ_Oil_comb) == "Gross Value"] <- "SJ Gross Value"
names(SJ_Oil_comb)[names(SJ_Oil_comb) == "Total Deductions"] <- "SJ Total Deductions"
names(SJ_Oil_comb)[names(SJ_Oil_comb) == "Deduction Percent"] <- "SJ Deduction Percent"
names(SJ_Oil_comb)[names(SJ_Oil_comb) == "Gross Price"] <- "SJ Gross Price"
names(SJ_Oil_comb)[names(SJ_Oil_comb) == "Net Value"] <- "SJ Net Value"

#NM
NM_oil_vol_sum <- aggregate(x = Oil_State[c("Volume")],
                            FUN = sum,
                            by = list("Filing Period" = Oil_State$`Filing Period`))
NM_oil_val_sum <- aggregate(x = Oil_State[c("Gross Value")],
                            FUN = sum,
                            by = list("Filing Period" = Oil_State$`Filing Period`))
NM_oil_dedct_sum <- aggregate(x = Oil_State[c("Total Deductions")],
                              FUN = sum,
                              by = list("Filing Period" = Oil_State$`Filing Period`))
NM_oil_Transdedct_sum <- aggregate(x = Oil_State[c("Transportation Deduction")],
                                   FUN = sum,
                                   by = list("Filing Period" = Oil_State$`Filing Period`))
NM_oil_Processdedct_sum <- aggregate(x = Oil_State[c("Processing Deduction")],
                                     FUN = sum,
                                     by = list("Filing Period" = Oil_State$`Filing Period`))
NM_oil_Royaltydedct_sum <- aggregate(x = Oil_State[c("Royalty Deduction")],
                                     FUN = sum,
                                     by = list("Filing Period" = Oil_State$`Filing Period`))

NM_Oil_comb <- Reduce(function(x, y) merge(x, y, all=TRUE), list(NM_oil_vol_sum, NM_oil_val_sum, NM_oil_dedct_sum))
NM_Oil_comb$`Deduction Percent` <- NM_Oil_comb$`Total Deductions`/NM_Oil_comb$`Gross Value`
NM_Oil_comb$`Gross Price` <- NM_Oil_comb$`Gross Value`/NM_Oil_comb$Volume
NM_Oil_comb$`Net Value` <- NM_Oil_comb$`Gross Value`-NM_Oil_comb$`Total Deductions`

#rename columns
names(NM_Oil_comb)[names(NM_Oil_comb) == "Volume"] <- "NM Volume"
names(NM_Oil_comb)[names(NM_Oil_comb) == "Gross Value"] <- "NM Gross Value"
names(NM_Oil_comb)[names(NM_Oil_comb) == "Total Deductions"] <- "NM Total Deductions"
names(NM_Oil_comb)[names(NM_Oil_comb) == "Deduction Percent"] <- "NM Deduction Percent"
names(NM_Oil_comb)[names(NM_Oil_comb) == "Gross Price"] <- "NM Gross Price"
names(NM_Oil_comb)[names(NM_Oil_comb) == "Net Value"] <- "NM Net Value"


#### Natural Gas
#Permian Basin
Perm_NG_vol_sum <- aggregate(x = Permian_NatGas[c("Volume")],
                             FUN = sum,
                             by = list("Filing Period" = Permian_NatGas$`Filing Period`))
Perm_NG_val_sum <- aggregate(x = Permian_NatGas[c("Gross Value")],
                             FUN = sum,
                             by = list("Filing Period" = Permian_NatGas$`Filing Period`))
Perm_NG_dedct_sum <- aggregate(x = Permian_NatGas[c("Total Deductions")],
                               FUN = sum,
                               by = list("Filing Period" = Permian_NatGas$`Filing Period`))
Perm_NG_Transdedct_sum <- aggregate(x = Permian_NatGas[c("Transportation Deduction")],
                                    FUN = sum,
                                    by = list("Filing Period" = Permian_NatGas$`Filing Period`))
Perm_NG_Processdedct_sum <- aggregate(x = Permian_NatGas[c("Processing Deduction")],
                                      FUN = sum,
                                      by = list("Filing Period" = Permian_NatGas$`Filing Period`))
Perm_NG_Royaltydedct_sum <- aggregate(x = Permian_NatGas[c("Royalty Deduction")],
                                      FUN = sum,
                                      by = list("Filing Period" = Permian_NatGas$`Filing Period`))

Perm_NG_comb <- Reduce(function(x, y) merge(x, y, all=TRUE), list(Perm_NG_vol_sum, Perm_NG_val_sum, Perm_NG_dedct_sum))
Perm_NG_comb$`Deduction Percent` <- Perm_NG_comb$`Total Deductions`/Perm_NG_comb$`Gross Value`
Perm_NG_comb$`Gross Price` <- Perm_NG_comb$`Gross Value`/Perm_NG_comb$Volume
Perm_NG_comb$`Net Value` <- Perm_NG_comb$`Gross Value`-Perm_NG_comb$`Total Deductions`

#rename columns
names(Perm_NG_comb)[names(Perm_NG_comb) == "Volume"] <- "Permian Volume"
names(Perm_NG_comb)[names(Perm_NG_comb) == "Gross Value"] <- "Permian Gross Value"
names(Perm_NG_comb)[names(Perm_NG_comb) == "Total Deductions"] <- "Permian Total Deductions"
names(Perm_NG_comb)[names(Perm_NG_comb) == "Deduction Percent"] <- "Permian Deduction Percent"
names(Perm_NG_comb)[names(Perm_NG_comb) == "Gross Price"] <- "Permian Gross Price"
names(Perm_NG_comb)[names(Perm_NG_comb) == "Net Value"] <- "Permian Net Value"


#San Juan Basin
SJ_NG_vol_sum <- aggregate(x = San_Juan_NatGas[c("Volume")],
                           FUN = sum,
                           by = list("Filing Period" = San_Juan_NatGas$`Filing Period`))
SJ_NG_val_sum <- aggregate(x = San_Juan_NatGas[c("Gross Value")],
                           FUN = sum,
                           by = list("Filing Period" = San_Juan_NatGas$`Filing Period`))
SJ_NG_dedct_sum <- aggregate(x = San_Juan_NatGas[c("Total Deductions")],
                             FUN = sum,
                             by = list("Filing Period" = San_Juan_NatGas$`Filing Period`))
SJ_NG_Transdedct_sum <- aggregate(x = San_Juan_NatGas[c("Transportation Deduction")],
                                  FUN = sum,
                                  by = list("Filing Period" = San_Juan_NatGas$`Filing Period`))
SJ_NG_Processdedct_sum <- aggregate(x = San_Juan_NatGas[c("Processing Deduction")],
                                    FUN = sum,
                                    by = list("Filing Period" = San_Juan_NatGas$`Filing Period`))
SJ_NG_Royaltydedct_sum <- aggregate(x = San_Juan_NatGas[c("Royalty Deduction")],
                                    FUN = sum,
                                    by = list("Filing Period" = San_Juan_NatGas$`Filing Period`))

SJ_NG_comb <- Reduce(function(x, y) merge(x, y, all=TRUE), list(SJ_NG_vol_sum, SJ_NG_val_sum, SJ_NG_dedct_sum))
SJ_NG_comb$`Deduction Percent` <- SJ_NG_comb$`Total Deductions`/SJ_NG_comb$`Gross Value`
SJ_NG_comb$`Gross Price` <- SJ_NG_comb$`Gross Value`/SJ_NG_comb$Volume
SJ_NG_comb$`Net Value` <- SJ_NG_comb$`Gross Value`-SJ_NG_comb$`Total Deductions`

#rename columns
names(SJ_NG_comb)[names(SJ_NG_comb) == "Volume"] <- "SJ Volume"
names(SJ_NG_comb)[names(SJ_NG_comb) == "Gross Value"] <- "SJ Gross Value"
names(SJ_NG_comb)[names(SJ_NG_comb) == "Total Deductions"] <- "SJ Total Deductions"
names(SJ_NG_comb)[names(SJ_NG_comb) == "Deduction Percent"] <- "SJ Deduction Percent"
names(SJ_NG_comb)[names(SJ_NG_comb) == "Gross Price"] <- "SJ Gross Price"
names(SJ_NG_comb)[names(SJ_NG_comb) == "Net Value"] <- "SJ Net Value"


#NM
NM_NG_vol_sum <- aggregate(x = NG_State[c("Volume")],
                           FUN = sum,
                           by = list("Filing Period" = NG_State$`Filing Period`))
NM_NG_val_sum <- aggregate(x = NG_State[c("Gross Value")],
                           FUN = sum,
                           by = list("Filing Period" = NG_State$`Filing Period`))
NM_NG_dedct_sum <- aggregate(x = NG_State[c("Total Deductions")],
                             FUN = sum,
                             by = list("Filing Period" = NG_State$`Filing Period`))
NM_NG_Transdedct_sum <- aggregate(x = NG_State[c("Transportation Deduction")],
                                  FUN = sum,
                                  by = list("Filing Period" = NG_State$`Filing Period`))
NM_NG_Processdedct_sum <- aggregate(x = NG_State[c("Processing Deduction")],
                                    FUN = sum,
                                    by = list("Filing Period" = NG_State$`Filing Period`))
NM_NG_Royaltydedct_sum <- aggregate(x = NG_State[c("Royalty Deduction")],
                                    FUN = sum,
                                    by = list("Filing Period" = NG_State$`Filing Period`))

NM_NG_comb <- Reduce(function(x, y) merge(x, y, all=TRUE), list(NM_NG_vol_sum, NM_NG_val_sum, NM_NG_dedct_sum))
NM_NG_comb$`Deduction Percent` <- NM_NG_comb$`Total Deductions`/NM_NG_comb$`Gross Value`
NM_NG_comb$`Gross Price` <- NM_NG_comb$`Gross Value`/NM_NG_comb$Volume
NM_NG_comb$`Net Value` <- NM_NG_comb$`Gross Value`-NM_NG_comb$`Total Deductions`

#rename columns
names(NM_NG_comb)[names(NM_NG_comb) == "Volume"] <- "NM Volume"
names(NM_NG_comb)[names(NM_NG_comb) == "Gross Value"] <- "NM Gross Value"
names(NM_NG_comb)[names(NM_NG_comb) == "Total Deductions"] <- "NM Total Deductions"
names(NM_NG_comb)[names(NM_NG_comb) == "Deduction Percent"] <- "NM Deduction Percent"
names(NM_NG_comb)[names(NM_NG_comb) == "Gross Price"] <- "NM Gross Price"
names(NM_NG_comb)[names(NM_NG_comb) == "Net Value"] <- "NM Net Value"


# Deduction Tracking
NM_NG_Deductions <- Reduce(function(x, y) merge(x, y, all=TRUE), list(NM_NG_Transdedct_sum, NM_NG_Processdedct_sum, NM_NG_Royaltydedct_sum))
SJ_NG_Deductions <- Reduce(function(x, y) merge(x, y, all=TRUE), list(SJ_NG_Transdedct_sum, SJ_NG_Processdedct_sum, SJ_NG_Royaltydedct_sum))
Perm_NG_Deductions <- Reduce(function(x, y) merge(x, y, all=TRUE), list(Perm_NG_Transdedct_sum, Perm_NG_Processdedct_sum, Perm_NG_Royaltydedct_sum))

Perm_Oil_Deductions <- Reduce(function(x, y) merge(x, y, all=TRUE), list(Perm_oil_Transdedct_sum, Perm_oil_Processdedct_sum, Perm_oil_Royaltydedct_sum))
NM_Oil_Deductions <- Reduce(function(x, y) merge(x, y, all=TRUE), list(NM_oil_Transdedct_sum, NM_oil_Processdedct_sum, NM_oil_Royaltydedct_sum))
SJ_Oil_Deductions <- Reduce(function(x, y) merge(x, y, all=TRUE), list(SJ_oil_Transdedct_sum, SJ_oil_Processdedct_sum, SJ_oil_Royaltydedct_sum))

#rename columns
names(NM_NG_Deductions)[names(NM_NG_Deductions) == "Transportation Deduction"] <- "NM NatGas Transportation Deduction"
names(NM_NG_Deductions)[names(NM_NG_Deductions) == "Processing Deduction"] <- "NM NatGas Processing Deduction"
names(NM_NG_Deductions)[names(NM_NG_Deductions) == "Royalty Deduction"] <- "NM NatGas Royalty Deduction"
names(NM_Oil_Deductions)[names(NM_Oil_Deductions) == "Transportation Deduction"] <- "NM Oil Transportation Deduction"
names(NM_Oil_Deductions)[names(NM_Oil_Deductions) == "Processing Deduction"] <- "NM Oil Processing Deduction"
names(NM_Oil_Deductions)[names(NM_Oil_Deductions) == "Royalty Deduction"] <- "NM Oil Royalty Deduction"

names(Perm_NG_Deductions)[names(Perm_NG_Deductions) == "Transportation Deduction"] <- "Perm NatGas Transportation Deduction"
names(Perm_NG_Deductions)[names(Perm_NG_Deductions) == "Processing Deduction"] <- "Perm NatGas Processing Deduction"
names(Perm_NG_Deductions)[names(Perm_NG_Deductions) == "Royalty Deduction"] <- "Perm NatGas Royalty Deduction"
names(Perm_Oil_Deductions)[names(Perm_Oil_Deductions) == "Transportation Deduction"] <- "Perm Oil Transportation Deduction"
names(Perm_Oil_Deductions)[names(Perm_Oil_Deductions) == "Processing Deduction"] <- "Perm Oil Processing Deduction"
names(Perm_Oil_Deductions)[names(Perm_Oil_Deductions) == "Royalty Deduction"] <- "Perm Oil Royalty Deduction"

names(SJ_NG_Deductions)[names(SJ_NG_Deductions) == "Transportation Deduction"] <- "SJ NatGas Transportation Deduction"
names(SJ_NG_Deductions)[names(SJ_NG_Deductions) == "Processing Deduction"] <- "SJ NatGas Processing Deduction"
names(SJ_NG_Deductions)[names(SJ_NG_Deductions) == "Royalty Deduction"] <- "SJ NatGas Royalty Deduction"
names(SJ_Oil_Deductions)[names(SJ_Oil_Deductions) == "Transportation Deduction"] <- "SJ Oil Transportation Deduction"
names(SJ_Oil_Deductions)[names(SJ_Oil_Deductions) == "Processing Deduction"] <- "SJ Oil Processing Deduction"
names(SJ_Oil_Deductions)[names(SJ_Oil_Deductions) == "Royalty Deduction"] <- "SJ Oil Royalty Deduction"


#######################################################################################################################################################
#Oil Tracking by Month
Oil_Tracking <- Reduce(function(x, y) merge(x, y, all=TRUE), list(Perm_Oil_comb, SJ_Oil_comb, NM_Oil_comb))
#Natural Gas Tracking by Month
NG_Tracking <- Reduce(function(x, y) merge(x, y, all=TRUE), list(Perm_NG_comb, SJ_NG_comb, NM_NG_comb))
#Deductions Tracking
Oil_Deductions_Tracking <- Reduce(function(x, y) merge(x, y, all=TRUE), list(Perm_Oil_Deductions, SJ_Oil_Deductions, NM_Oil_Deductions))
NG_Deductions_Tracking <- Reduce(function(x, y) merge(x, y, all=TRUE), list(Perm_NG_Deductions, SJ_NG_Deductions, NM_NG_Deductions))


#########################################################################################################################################################
#########################################################################################################################################################
#########################################################################################################################################################
#########################################################################################################################################################
#########################################################################################################################################################
#########################################################################################################################################################
#########################################################################################################################################################

# Plots and formatting

# Create Time Series Plots
# Libraries
library(ggplot2)
library(dplyr)
library(plotly)
library(hrbrthemes)
hrbrthemes::import_roboto_condensed()

# Oil Time Series Plot
#NM_OilCh <- Oil_Tracking %>%
# ggplot( aes(x=`Filing Period`, y=`NM Volume`)) +
#scale_y_continuous(name="NM Oil Volume (in barrels)", labels = scales::comma) +
#geom_area(fill="#69b3a2", alpha=0.5) +
#geom_line(color="#69b3a2") +
#ylab("NM Oil Volume (in barrels)") +
#ggtitle("NM Oil Volume") +
#theme_ipsum()
#NM_OilCh <- ggplotly(NM_OilCh)
#NM_OilCh

# Oil Time Series Plot
NM_OilCh <- ggplot(Oil_Tracking, aes(x=`Filing Period`, y=`NM Volume`)) +
  scale_y_continuous(name= "NM Oil Volume (in barrels)", labels = scales::comma) +
  geom_area(fill="#69b3a2", alpha=0.5) +
  geom_line(color="#69b3a2") +
  xlab("Filing Period")+
  ylab("NM Oil Volume (in barrels")+
  ggtitle("NM Oil Volume") +
  theme_ipsum()
NM_OilCh

# Nat Gas Time Series Plot
#NM_GasCh <- NG_Tracking %>%
# ggplot( aes(x=`Filing Period`, y=`NM Volume`)) + 
#scale_y_continuous(name="NM Natural Gas Volume (in MCF)", labels = scales::comma) +
#geom_area(fill="#69b3a2", alpha=0.5) +
#geom_line(color="#69b3a2") +
#ylab("NM Natural Gas Volume (in MCF)") +
#ggtitle("NM Natural Gas Volume") +
#theme_ipsum()
#NM_GasCh <- ggplotly(NM_GasCh)
#NM_GasCh

# Nat Gas Time Series Plot
NM_GasCh <- ggplot(NG_Tracking, aes(x=`Filing Period`, y=`NM Volume`)) +
  scale_y_continuous(name= "NM Natural Gas Volume (in MCF)", labels = scales::comma) +
  geom_area(fill="#69b3a2", alpha=0.5) +
  geom_line(color="#69b3a2") +
  xlab("Filing Period")+
  ylab("NM Natural Gas Volume (in MCF)")+
  ggtitle("NM Natural Gas Volume") +
  theme_ipsum()
NM_GasCh

# Oil Price
#NM_OilPr <- Oil_Tracking %>%
# ggplot(aes(x=`Filing Period`, y=`NM Gross Price`)) +
#scale_y_continuous(name= "NM Oil Price per Barrel", labels = scales::dollar) +
#geom_area(fill="#69b3a2", alpha=0.5) +
#geom_line(color="#69b3a2") +
#ylab("NM Oil Price per Barrel")+
#ggtitle("NM Oil Price") +
#theme_ipsum()
#NM_OilPr <- ggplotly(NM_OilPr)
#NM_OilPr

# Oil Price
NM_OilPr <- ggplot(Oil_Tracking, aes(x=`Filing Period`, y=`NM Gross Price`)) +
  scale_y_continuous(name= "NM Gross Price", labels = scales::dollar) +
  geom_line() +
  xlab("Filing Period")+
  ylab("NM Oil Price per Barrel")+
  ggtitle("NM Oil Price per Barrel")
NM_OilPr

# Natural Gas Price
NM_GasPr <- ggplot(NG_Tracking, aes(x=`Filing Period`, y=`NM Gross Price`)) +
  scale_y_continuous(name= "NM Natural Gas Price per MCF", labels = scales::dollar) +
  geom_line() +
  xlab("Filing Period")+
  ylab("NM Natural Gas Price per MCF")+
  ggtitle("NM Natural Gas Price")
NM_GasPr

##########################################################################################################################################################
# CREATE WORKBOOKS WITH SEPARATE SHEETS #

# create workbook
wb <- createWorkbook()
# add worksheets
addWorksheet(wb, sheetName = "Charts")
addWorksheet(wb, sheetName = "Oil Tracker")
addWorksheet(wb, sheetName = "Nat Gas Tracker")
addWorksheet(wb, sheetName = "Fed Land Oil Tracker")
addWorksheet(wb, sheetName = "Fed Land Nat Gas Tracker")
addWorksheet(wb, sheetName = "State Land Oil Tracker")
addWorksheet(wb, sheetName = "State Land Nat Gas Tracker")
addWorksheet(wb, sheetName = "Oil Deductions Tracker")
addWorksheet(wb, sheetName = "Nat Gas Deductions Tracker")

# Write data to worksheets
print(NM_OilCh)
wb %>% insertPlot(sheet = "Charts", startCol = "A", startRow = 1)
print(NM_GasCh)
wb %>% insertPlot(sheet = "Charts", startCol = "I", startRow = 1)
print(NM_OilPr)
wb %>% insertPlot(sheet = "Charts", startCol = "A", startRow = 24)
print(NM_GasPr)
wb %>% insertPlot(sheet = "Charts", startCol = "I", startRow = 24)

# format values in worksheets #

writeData(wb = wb, sheet = "Oil Tracker", x = Oil_Tracking,
          startCol = 1, startRow = 1, colNames = TRUE)

LabelStyle <- createStyle(halign = "center",
                          border = c("bottom", "right"), 
                          borderStyle = "thin", 
                          textDecoration = "bold", 
                          fgFill = "#0491A1", 
                          fontColour = "white")
BorderStyle <- createStyle(halign = "right", border = c("bottom", "right", "left"))
NumStyle <- createStyle(halign = "right", numFmt = "COMMA")
PrcStyle <- createStyle(halign = "right", numFmt = "CURRENCY")
PrcntgStyle <- createStyle(halign = "right", numFmt = "PERCENTAGE")
TextStyle <- createStyle(halign = "center", 
                         border = "bottom", 
                         borderStyle = "thin")

DateStyle <- createStyle(halign = "center", numFmt = "mm/dd/yyyy")

BkGrdStyle <- createStyle(fgFill = "#FFFFFF")

addStyle(wb, sheet = "Oil Tracker", style = LabelStyle, rows = 1, cols = 1:19, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Oil Tracker", style = DateStyle, rows = 2:500, cols = 1, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Oil Tracker", style = NumStyle, rows = 2:500, cols = 2, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Oil Tracker", style = PrcStyle, rows = 2:500, cols = 3, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Oil Tracker", style = PrcStyle, rows = 2:500, cols = 4, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Oil Tracker", style = PrcntgStyle, rows = 2:500, cols = 5, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Oil Tracker", style = PrcStyle, rows = 2:500, cols = 6, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Oil Tracker", style = PrcStyle, rows = 2:500, cols = 7, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Oil Tracker", style = NumStyle, rows = 2:500, cols = 8, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Oil Tracker", style = PrcStyle, rows = 2:500, cols = 9, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Oil Tracker", style = PrcStyle, rows = 2:500, cols = 10, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Oil Tracker", style = PrcntgStyle, rows = 2:500, cols = 11, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Oil Tracker", style = PrcStyle, rows = 2:500, cols = 12, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Oil Tracker", style = PrcStyle, rows = 2:500, cols = 13, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Oil Tracker", style = NumStyle, rows = 2:500, cols = 14, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Oil Tracker", style = PrcStyle, rows = 2:500, cols = 15, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Oil Tracker", style = PrcStyle, rows = 2:500, cols = 16, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Oil Tracker", style = PrcntgStyle, rows = 2:500, cols = 17, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Oil Tracker", style = PrcStyle, rows = 2:500, cols = 18, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Oil Tracker", style = PrcStyle, rows = 2:500, cols = 19, 
         gridExpand = FALSE, stack = FALSE)


#Append style information to a multiple rows and columns without overwriting
#the current style
addStyle(wb, sheet = "Oil Tracker", style = BkGrdStyle, rows = 2:500, cols = 1:19,
         gridExpand = TRUE, stack = TRUE)

addStyle(wb, sheet = "Oil Tracker", style = BorderStyle, rows = 2:500, cols = 1:19,
         gridExpand = TRUE, stack = TRUE)
###########################

writeData(wb = wb, sheet = "Nat Gas Tracker", x = NG_Tracking,
          startCol = 1, startRow = 1, colNames = TRUE)

LabelStyle <- createStyle(halign = "center",
                          border = c("bottom", "right"), 
                          borderStyle = "thin", 
                          textDecoration = "bold", 
                          fgFill = "#0491A1", 
                          fontColour = "white")
BorderStyle <- createStyle(halign = "right", border = c("bottom", "right", "left"))
NumStyle <- createStyle(halign = "right", numFmt = "COMMA")
PrcStyle <- createStyle(halign = "right", numFmt = "CURRENCY")
PrcntgStyle <- createStyle(halign = "right", numFmt = "PERCENTAGE")
TextStyle <- createStyle(halign = "center", 
                         border = "bottom", 
                         borderStyle = "thin")

DateStyle <- createStyle(halign = "center", numFmt = "mm/dd/yyyy")

BkGrdStyle <- createStyle(fgFill = "#FFFFFF")

addStyle(wb, sheet = "Nat Gas Tracker", style = LabelStyle, rows = 1, cols = 1:19, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Nat Gas Tracker", style = DateStyle, rows = 2:500, cols = 1, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Nat Gas Tracker", style = NumStyle, rows = 2:500, cols = 2, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Nat Gas Tracker", style = PrcStyle, rows = 2:500, cols = 3, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Nat Gas Tracker", style = PrcStyle, rows = 2:500, cols = 4, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Nat Gas Tracker", style = PrcntgStyle, rows = 2:500, cols = 5, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Nat Gas Tracker", style = PrcStyle, rows = 2:500, cols = 6, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Nat Gas Tracker", style = PrcStyle, rows = 2:500, cols = 7, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Nat Gas Tracker", style = NumStyle, rows = 2:500, cols = 8, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Nat Gas Tracker", style = PrcStyle, rows = 2:500, cols = 9, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Nat Gas Tracker", style = PrcStyle, rows = 2:500, cols = 10, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Nat Gas Tracker", style = PrcntgStyle, rows = 2:500, cols = 11, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Nat Gas Tracker", style = PrcStyle, rows = 2:500, cols = 12, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Nat Gas Tracker", style = PrcStyle, rows = 2:500, cols = 13, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Nat Gas Tracker", style = NumStyle, rows = 2:500, cols = 14, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Nat Gas Tracker", style = PrcStyle, rows = 2:500, cols = 15, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Nat Gas Tracker", style = PrcStyle, rows = 2:500, cols = 16, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Nat Gas Tracker", style = PrcntgStyle, rows = 2:500, cols = 17, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Nat Gas Tracker", style = PrcStyle, rows = 2:500, cols = 18, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Nat Gas Tracker", style = PrcStyle, rows = 2:500, cols = 19, 
         gridExpand = FALSE, stack = FALSE)

#Append style information to a multiple rows and columns without overwriting
#the current style
addStyle(wb, sheet = "Nat Gas Tracker", style = BkGrdStyle, rows = 2:500, cols = 1:19,
         gridExpand = TRUE, stack = TRUE)

addStyle(wb, sheet = "Nat Gas Tracker", style = BorderStyle, rows = 2:500, cols = 1:19,
         gridExpand = TRUE, stack = TRUE)

############################################################################

writeData(wb = wb, sheet = "Fed Land Oil Tracker", x = Fed_Oil_comb,
          startCol = 1, startRow = 1, colNames = TRUE)

LabelStyle <- createStyle(halign = "center",
                          border = c("bottom", "right"), 
                          borderStyle = "thin", 
                          textDecoration = "bold", 
                          fgFill = "#0491A1", 
                          fontColour = "white")
BorderStyle <- createStyle(halign = "right", border = c("bottom", "right", "left"))
NumStyle <- createStyle(halign = "right", numFmt = "COMMA")
PrcStyle <- createStyle(halign = "right", numFmt = "CURRENCY")
PrcntgStyle <- createStyle(halign = "right", numFmt = "PERCENTAGE")
TextStyle <- createStyle(halign = "center", 
                         border = "bottom", 
                         borderStyle = "thin")

DateStyle <- createStyle(halign = "center", numFmt = "mm/dd/yyyy")

BkGrdStyle <- createStyle(fgFill = "#FFFFFF")

addStyle(wb, sheet = "Fed Land Oil Tracker", style = LabelStyle, rows = 1, cols = 1:10, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Fed Land Oil Tracker", style = DateStyle, rows = 2:500, cols = 1, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Fed Land Oil Tracker", style = NumStyle, rows = 2:500, cols = 2, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Fed Land Oil Tracker", style = PrcStyle, rows = 2:500, cols = 3, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Fed Land Oil Tracker", style = PrcStyle, rows = 2:500, cols = 4, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Fed Land Oil Tracker", style = PrcStyle, rows = 2:500, cols = 5, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Fed Land Oil Tracker", style = PrcStyle, rows = 2:500, cols = 6, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Fed Land Oil Tracker", style = PrcStyle, rows = 2:500, cols = 7, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Fed Land Oil Tracker", style = PrcntgStyle, rows = 2:500, cols = 8, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Fed Land Oil Tracker", style = PrcStyle, rows = 2:500, cols = 9, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Fed Land Oil Tracker", style = PrcStyle, rows = 2:500, cols = 10, 
         gridExpand = FALSE, stack = FALSE)

#Append style information to a multiple rows and columns without overwriting
#the current style
addStyle(wb, sheet = "Fed Land Oil Tracker", style = BkGrdStyle, rows = 2:500, cols = 1:10,
         gridExpand = TRUE, stack = TRUE)

addStyle(wb, sheet = "Fed Land Oil Tracker", style = BorderStyle, rows = 2:500, cols = 1:10,
         gridExpand = TRUE, stack = TRUE)

#########################################################################################

writeData(wb = wb, sheet = "Fed Land Nat Gas Tracker", x = fed_gas_comb,
          startCol = 1, startRow = 1, colNames = TRUE)

LabelStyle <- createStyle(halign = "center",
                          border = c("bottom", "right"), 
                          borderStyle = "thin", 
                          textDecoration = "bold", 
                          fgFill = "#0491A1", 
                          fontColour = "white")
BorderStyle <- createStyle(halign = "right", border = c("bottom", "right", "left"))
NumStyle <- createStyle(halign = "right", numFmt = "COMMA")
PrcStyle <- createStyle(halign = "right", numFmt = "CURRENCY")
PrcntgStyle <- createStyle(halign = "right", numFmt = "PERCENTAGE")
TextStyle <- createStyle(halign = "center", 
                         border = "bottom", 
                         borderStyle = "thin")

DateStyle <- createStyle(halign = "center", numFmt = "mm/dd/yyyy")

BkGrdStyle <- createStyle(fgFill = "#FFFFFF")

addStyle(wb, sheet = "Fed Land Nat Gas Tracker", style = LabelStyle, rows = 1, cols = 1:10, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Fed Land Nat Gas Tracker", style = DateStyle, rows = 2:500, cols = 1, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Fed Land Nat Gas Tracker", style = NumStyle, rows = 2:500, cols = 2, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Fed Land Nat Gas Tracker", style = PrcStyle, rows = 2:500, cols = 3, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Fed Land Nat Gas Tracker", style = PrcStyle, rows = 2:500, cols = 4, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Fed Land Nat Gas Tracker", style = PrcStyle, rows = 2:500, cols = 5, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Fed Land Nat Gas Tracker", style = PrcStyle, rows = 2:500, cols = 6, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Fed Land Nat Gas Tracker", style = PrcStyle, rows = 2:500, cols = 7, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Fed Land Nat Gas Tracker", style = PrcntgStyle, rows = 2:500, cols = 8, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Fed Land Nat Gas Tracker", style = PrcStyle, rows = 2:500, cols = 9, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Fed Land Nat Gas Tracker", style = PrcStyle, rows = 2:500, cols = 10, 
         gridExpand = FALSE, stack = FALSE)

#Append style information to a multiple rows and columns without overwriting
#the current style
addStyle(wb, sheet = "Fed Land Nat Gas Tracker", style = BkGrdStyle, rows = 2:500, cols = 1:10,
         gridExpand = TRUE, stack = TRUE)

addStyle(wb, sheet = "Fed Land Nat Gas Tracker", style = BorderStyle, rows = 2:500, cols = 1:10,
         gridExpand = TRUE, stack = TRUE)


#########################################################################################

writeData(wb = wb, sheet = "State Land Oil Tracker", x = state_Oil_comb,
          startCol = 1, startRow = 1, colNames = TRUE)

LabelStyle <- createStyle(halign = "center",
                          border = c("bottom", "right"), 
                          borderStyle = "thin", 
                          textDecoration = "bold", 
                          fgFill = "#0491A1", 
                          fontColour = "white")
BorderStyle <- createStyle(halign = "right", border = c("bottom", "right", "left"))
NumStyle <- createStyle(halign = "right", numFmt = "COMMA")
PrcStyle <- createStyle(halign = "right", numFmt = "CURRENCY")
PrcntgStyle <- createStyle(halign = "right", numFmt = "PERCENTAGE")
TextStyle <- createStyle(halign = "center", 
                         border = "bottom", 
                         borderStyle = "thin")

DateStyle <- createStyle(halign = "center", numFmt = "mm/dd/yyyy")

BkGrdStyle <- createStyle(fgFill = "#FFFFFF")

addStyle(wb, sheet = "State Land Oil Tracker", style = LabelStyle, rows = 1, cols = 1:10, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "State Land Oil Tracker", style = DateStyle, rows = 2:500, cols = 1, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "State Land Oil Tracker", style = NumStyle, rows = 2:500, cols = 2, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "State Land Oil Tracker", style = PrcStyle, rows = 2:500, cols = 3, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "State Land Oil Tracker", style = PrcStyle, rows = 2:500, cols = 4, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "State Land Oil Tracker", style = PrcStyle, rows = 2:500, cols = 5, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "State Land Oil Tracker", style = PrcStyle, rows = 2:500, cols = 6, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "State Land Oil Tracker", style = PrcStyle, rows = 2:500, cols = 7, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "State Land Oil Tracker", style = PrcntgStyle, rows = 2:500, cols = 8, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "State Land Oil Tracker", style = PrcStyle, rows = 2:500, cols = 9, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "State Land Oil Tracker", style = PrcStyle, rows = 2:500, cols = 10, 
         gridExpand = FALSE, stack = FALSE)

#Append style information to a multiple rows and columns without overwriting
#the current style
addStyle(wb, sheet = "State Land Oil Tracker", style = BkGrdStyle, rows = 2:500, cols = 1:10,
         gridExpand = TRUE, stack = TRUE)

addStyle(wb, sheet = "State Land Oil Tracker", style = BorderStyle, rows = 2:500, cols = 1:10,
         gridExpand = TRUE, stack = TRUE)

#########################################################################################################


writeData(wb = wb, sheet = "State Land Nat Gas Tracker", x = state_gas_comb,
          startCol = 1, startRow = 1, colNames = TRUE)

LabelStyle <- createStyle(halign = "center",
                          border = c("bottom", "right"), 
                          borderStyle = "thin", 
                          textDecoration = "bold", 
                          fgFill = "#0491A1", 
                          fontColour = "white")
BorderStyle <- createStyle(halign = "right", border = c("bottom", "right", "left"))
NumStyle <- createStyle(halign = "right", numFmt = "COMMA")
PrcStyle <- createStyle(halign = "right", numFmt = "CURRENCY")
PrcntgStyle <- createStyle(halign = "right", numFmt = "PERCENTAGE")
TextStyle <- createStyle(halign = "center", 
                         border = "bottom", 
                         borderStyle = "thin")

DateStyle <- createStyle(halign = "center", numFmt = "mm/dd/yyyy")

BkGrdStyle <- createStyle(fgFill = "#FFFFFF")

addStyle(wb, sheet = "State Land Nat Gas Tracker", style = LabelStyle, rows = 1, cols = 1:10, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "State Land Nat Gas Tracker", style = DateStyle, rows = 2:500, cols = 1, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "State Land Nat Gas Tracker", style = NumStyle, rows = 2:500, cols = 2, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "State Land Nat Gas Tracker", style = PrcStyle, rows = 2:500, cols = 3, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "State Land Nat Gas Tracker", style = PrcStyle, rows = 2:500, cols = 4, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "State Land Nat Gas Tracker", style = PrcStyle, rows = 2:500, cols = 5, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "State Land Nat Gas Tracker", style = PrcStyle, rows = 2:500, cols = 6, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "State Land Nat Gas Tracker", style = PrcStyle, rows = 2:500, cols = 7, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "State Land Nat Gas Tracker", style = PrcntgStyle, rows = 2:500, cols = 8, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "State Land Nat Gas Tracker", style = PrcStyle, rows = 2:500, cols = 9, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "State Land Nat Gas Tracker", style = PrcStyle, rows = 2:500, cols = 10, 
         gridExpand = FALSE, stack = FALSE)

#Append style information to a multiple rows and columns without overwriting
#the current style
addStyle(wb, sheet = "State Land Nat Gas Tracker", style = BkGrdStyle, rows = 2:500, cols = 1:10,
         gridExpand = TRUE, stack = TRUE)

addStyle(wb, sheet = "State Land Nat Gas Tracker", style = BorderStyle, rows = 2:500, cols = 1:10,
         gridExpand = TRUE, stack = TRUE)

#################################################################################################################


writeData(wb = wb, sheet = "Oil Deductions Tracker", x = Oil_Deductions_Tracking,
          startCol = 1, startRow = 1, colNames = TRUE)

LabelStyle <- createStyle(halign = "center",
                          border = c("bottom", "right"), 
                          borderStyle = "thin", 
                          textDecoration = "bold", 
                          fgFill = "#0491A1", 
                          fontColour = "white")
BorderStyle <- createStyle(halign = "right", border = c("bottom", "right", "left"))
NumStyle <- createStyle(halign = "right", numFmt = "COMMA")
PrcStyle <- createStyle(halign = "right", numFmt = "CURRENCY")
PrcntgStyle <- createStyle(halign = "right", numFmt = "PERCENTAGE")
TextStyle <- createStyle(halign = "center", 
                         border = "bottom", 
                         borderStyle = "thin")

DateStyle <- createStyle(halign = "center", numFmt = "mm/dd/yyyy")

BkGrdStyle <- createStyle(fgFill = "#FFFFFF")

addStyle(wb, sheet = "Oil Deductions Tracker", style = LabelStyle, rows = 1, cols = 1:10, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Oil Deductions Tracker", style = DateStyle, rows = 2:500, cols = 1, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Oil Deductions Tracker", style = PrcStyle, rows = 2:500, cols = 2, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Oil Deductions Tracker", style = PrcStyle, rows = 2:500, cols = 3, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Oil Deductions Tracker", style = PrcStyle, rows = 2:500, cols = 4, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Oil Deductions Tracker", style = PrcStyle, rows = 2:500, cols = 5, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Oil Deductions Tracker", style = PrcStyle, rows = 2:500, cols = 6, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Oil Deductions Tracker", style = PrcStyle, rows = 2:500, cols = 7, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Oil Deductions Tracker", style = PrcStyle, rows = 2:500, cols = 8, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Oil Deductions Tracker", style = PrcStyle, rows = 2:500, cols = 9, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Oil Deductions Tracker", style = PrcStyle, rows = 2:500, cols = 10, 
         gridExpand = FALSE, stack = FALSE)


#Append style information to a multiple rows and columns without overwriting
#the current style
addStyle(wb, sheet = "Oil Deductions Tracker", style = BkGrdStyle, rows = 2:500, cols = 1:10,
         gridExpand = TRUE, stack = TRUE)

addStyle(wb, sheet = "Oil Deductions Tracker", style = BorderStyle, rows = 2:500, cols = 1:10,
         gridExpand = TRUE, stack = TRUE)


#################################################################################################################


writeData(wb = wb, sheet = "Nat Gas Deductions Tracker", x = NG_Deductions_Tracking,
          startCol = 1, startRow = 1, colNames = TRUE)

LabelStyle <- createStyle(halign = "center",
                          border = c("bottom", "right"), 
                          borderStyle = "thin", 
                          textDecoration = "bold", 
                          fgFill = "#0491A1", 
                          fontColour = "white")
BorderStyle <- createStyle(halign = "right", border = c("bottom", "right", "left"))
NumStyle <- createStyle(halign = "right", numFmt = "COMMA")
PrcStyle <- createStyle(halign = "right", numFmt = "CURRENCY")
PrcntgStyle <- createStyle(halign = "right", numFmt = "PERCENTAGE")
TextStyle <- createStyle(halign = "center", 
                         border = "bottom", 
                         borderStyle = "thin")

DateStyle <- createStyle(halign = "center", numFmt = "mm/dd/yyyy")

BkGrdStyle <- createStyle(fgFill = "#FFFFFF")

addStyle(wb, sheet = "Nat Gas Deductions Tracker", style = LabelStyle, rows = 1, cols = 1:10, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Nat Gas Deductions Tracker", style = DateStyle, rows = 2:500, cols = 1, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Nat Gas Deductions Tracker", style = PrcStyle, rows = 2:500, cols = 2, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Nat Gas Deductions Tracker", style = PrcStyle, rows = 2:500, cols = 3, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Nat Gas Deductions Tracker", style = PrcStyle, rows = 2:500, cols = 4, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Nat Gas Deductions Tracker", style = PrcStyle, rows = 2:500, cols = 5, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Nat Gas Deductions Tracker", style = PrcStyle, rows = 2:500, cols = 6, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Nat Gas Deductions Tracker", style = PrcStyle, rows = 2:500, cols = 7, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Nat Gas Deductions Tracker", style = PrcStyle, rows = 2:500, cols = 8, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Nat Gas Deductions Tracker", style = PrcStyle, rows = 2:500, cols = 9, 
         gridExpand = FALSE, stack = FALSE)

addStyle(wb, sheet = "Nat Gas Deductions Tracker", style = PrcStyle, rows = 2:500, cols = 10, 
         gridExpand = FALSE, stack = FALSE)


#Append style information to a multiple rows and columns without overwriting
#the current style
addStyle(wb, sheet = "Nat Gas Deductions Tracker", style = BkGrdStyle, rows = 2:500, cols = 1:10,
         gridExpand = TRUE, stack = TRUE)

addStyle(wb, sheet = "Nat Gas Deductions Tracker", style = BorderStyle, rows = 2:500, cols = 1:10,
         gridExpand = TRUE, stack = TRUE)

####################################################################################################################################

# Write to Excel
saveWorkbook(wb, "Oil and Gas Tracker 10.17.23.xlsx", overwrite = TRUE)

