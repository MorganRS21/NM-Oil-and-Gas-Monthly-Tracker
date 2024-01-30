############################################################################# 
# Oil and Gas Monthly Tracking Reports 
# Date edited:  01/30/2024
##############################################################################

# Use this program to produce monthly tracking reports

# Install and load required packages using 'pacman'
if (!require(pacman)) install.packages("pacman")
pacman::p_load(dplyr, tidyr, openxlsx, ggplot2, plotly, hrbrthemes)

# Import roboto condensed font for plotting
hrbrthemes::import_roboto_condensed()

################################################################################
# Read data - data should be placed in a folder in your directory so it's easy to navigate to
# Read data and create CSV file
TAP <- read.xlsx('data/County Vol_Val By Filing Period 202401110828.xlsx', detectDates = TRUE)
csv_file_name <- sprintf("data/County_Vol_Val_%s.csv", format(Sys.Date(), "%Y%m%d"))
write.csv(TAP, file = csv_file_name, row.names = FALSE)

TAP <- read.csv(csv_file_name)
Legacy <- read.csv('data/Legacy_data.csv')

# merge two data frames
County_vol_val <- rbind(Legacy,TAP)

# Format filing period variable as dates
County_vol_val$Filing.Period <- as.Date(County_vol_val$Filing.Period, origin = '1960-01-01')

#################################################################################################################################################
# Filter data by basin
Permian <- subset(County_vol_val, County_vol_val$Basin=="Permian")
San_Juan <- subset(County_vol_val, County_vol_val$Basin=="San Juan")
Raton <- subset(County_vol_val, County_vol_val$Basin=="Raton")
Bravo_Dome <- subset(County_vol_val, County_vol_val$Basin=="Bravo Dome")
##################################################################################################################################################

# Combined function to filter data by optional land type, basin, and product codes
filter_data_combined <- function(data, product_codes, land_type = NULL, basin = NULL) {
  if (!is.null(land_type)) {
    data <- subset(data, Land.Type == land_type)
  }
  if (!is.null(basin)) {
    data <- subset(data, Basin == basin)
  }
  filtered_data <- subset(data, Product.Code %in% product_codes)
  return(filtered_data)
}

# Function to perform aggregation on selected columns
aggregate_data <- function(data) {
  vol <- aggregate(x = data[c("Volume")], FUN = sum, by = list("Filing.Period" = data$Filing.Period))
  val <- aggregate(x = data[c("Gross.Value")], FUN = sum, by = list("Filing.Period" = data$Filing.Period))
  dedct <- aggregate(x = data[c("Total.Deductions")], FUN = sum, by = list("Filing.Period" = data$Filing.Period))
  Transdedct <- aggregate(x = data[c("Transportation.Deduction")], FUN = sum, by = list("Filing.Period" = data$Filing.Period))
  Processdedct <- aggregate(x = data[c("Processing.Deduction")], FUN = sum, by = list("Filing.Period" = data$Filing.Period))
  Royaltydedct <- aggregate(x = data[c("Royalty.Deduction")], FUN = sum, by = list("Filing.Period" = data$Filing.Period))
  
  combined_data <- Reduce(function(x, y) merge(x, y, all=TRUE), list(vol, val, dedct, Transdedct, Processdedct, Royaltydedct))
  combined_data$Deduction.Percent <- combined_data$Total.Deductions / combined_data$Gross.Value
  combined_data$Gross.Price <- combined_data$Gross.Value / combined_data$Volume
  combined_data$Net.Value <- combined_data$Gross.Value - combined_data$Total.Deductions
  
  return(combined_data)
}

# Define product codes for different commodities
oil_codes <- c("1", "2", "14", "5")
ng_codes <- c("7", "3", "4")
co2_code <- "17"
helium_code <- "8"

# Filter and aggregate data for Federal land oil
Fed_Oil <- filter_data(County_vol_val, "Federal", oil_codes)
fed_oil_comb <- aggregate_data(Fed_Oil)

# Filter and aggregate data for State land oil
State_Oil <- filter_data(County_vol_val, "State", oil_codes)
state_oil_comb <- aggregate_data(State_Oil)

# Filter and aggregate data for State land gas
state_gas_comb <- aggregate_data(filter_data(County_vol_val, "State", ng_codes))

# Filter and aggregate data for Federal land gas
fed_gas_comb <- aggregate_data(filter_data(County_vol_val, "Federal", ng_codes))

################################################################################
#################################################################################
# Filter by commodity - Apply same function from lines 39-48 to filter 
Oil_State <-  filter_data_combined(County_vol_val, oil_codes)
NG_State <- filter_data_combined(County_vol_val, ng_codes)
CO2_State <- filter_data_combined(County_vol_val, co2_code)
Helium_State <- filter_data_combined(County_vol_val, helium_code)

###############################################################################

# Filtering by basin and product type
# Apply same filter function from line 39 to the basins
basins <- list(Permian = Permian, San_Juan = San_Juan, Raton = Raton, Bravo_Dome = Bravo_Dome)
for (basin_name in names(basins)) {
  assign(paste0(basin_name, "_Oil"), filter_data_combined(basins[[basin_name]], oil_codes))
  assign(paste0(basin_name, "_NatGas"), filter_data_combined(basins[[basin_name]], ng_codes))
  assign(paste0(basin_name, "_CO2"), filter_data_combined(basins[[basin_name]], co2_code))
  assign(paste0(basin_name, "_Helium"), filter_data_combined(basins[[basin_name]], helium_code))
}

##############################################################################

# Sum data by filing period and produce excel tracking worksheets
# Function to aggregate data
aggregate_data <- function(data, cols) {
  lapply(cols, function(col) {
    aggregate(x = data[col],
              FUN = sum,
              by = list(Filing.Period = data$Filing.Period))
  })
}

# Function to process data for each basin and resource
process_data <- function(data, prefix, cols_to_sum) {
  # Aggregate data
  aggregated_data <- aggregate_data(data, cols_to_sum)
  
  # Combine aggregated data
  combined_data <- Reduce(function(x, y) merge(x, y, by = "Filing.Period", all = TRUE), aggregated_data)
  
  # Calculate additional metrics if needed
  if (all(c("Volume", "Gross.Value", "Total.Deductions") %in% cols_to_sum)) {
    combined_data[paste0(prefix, "_Deduction_Percent")] <- combined_data$Total.Deductions / combined_data$Gross.Value
    combined_data[paste0(prefix, "_Gross_Price")] <- combined_data$Gross.Value / combined_data$Volume
    combined_data[paste0(prefix, "_Net_Value")] <- combined_data$Gross.Value - combined_data$Total.Deductions
  }
  
  # Rename columns
  colnames(combined_data) <- gsub("^(Volume|Gross.Value|Total.Deductions|Transportation.Deduction|Processing.Deduction|Royalty.Deduction)$", 
                                  paste0(prefix, "_", "\\1"), 
                                  colnames(combined_data))
  
  # Return combined data
  combined_data
}

# Oil data processing
oil_cols <- c("Volume", "Gross.Value", "Total.Deductions", "Transportation.Deduction", "Processing.Deduction", "Royalty.Deduction")
Perm_Oil_comb <- process_data(Permian_Oil, "Permian", oil_cols)
SJ_Oil_comb <- process_data(San_Juan_Oil, "SJ", oil_cols)
NM_Oil_comb <- process_data(Oil_State, "NM", oil_cols)

# Natural Gas data processing
NG_cols <- c("Volume", "Gross.Value", "Total.Deductions", "Transportation.Deduction", "Processing.Deduction", "Royalty.Deduction")
Perm_NG_comb <- process_data(Permian_NatGas, "Permian", NG_cols)
SJ_NG_comb <- process_data(San_Juan_NatGas, "SJ", NG_cols)
NM_NG_comb <- process_data(NG_State, "NM", NG_cols)

# Combine the data for Oil and Natural Gas Tracking
Oil_Tracking <- Reduce(function(x, y) merge(x, y, by = "Filing.Period", all = TRUE), list(Perm_Oil_comb, SJ_Oil_comb, NM_Oil_comb))
NG_Tracking <- Reduce(function(x, y) merge(x, y, by = "Filing.Period", all = TRUE), list(Perm_NG_comb, SJ_NG_comb, NM_NG_comb))

# Deductions data processing
deduction_cols <- c("Transportation.Deduction", "Processing.Deduction", "Royalty.Deduction")
Perm_Oil_Deductions <- process_data(Permian_Oil, "Permian_Oil", deduction_cols)
SJ_Oil_Deductions <- process_data(San_Juan_Oil, "SJ_Oil", deduction_cols)
NM_Oil_Deductions <- process_data(Oil_State, "NM_Oil", deduction_cols)
Perm_NG_Deductions <- process_data(Permian_NatGas, "Permian_NG", deduction_cols)
SJ_NG_Deductions <- process_data(San_Juan_NatGas, "SJ_NG", deduction_cols)
NM_NG_Deductions <- process_data(NG_State, "NM_NG", deduction_cols)

# Combine the data for Deductions Tracking
Oil_Deductions_Tracking <- Reduce(function(x, y) merge(x, y, by = "Filing.Period", all = TRUE), list(Perm_Oil_Deductions, SJ_Oil_Deductions, NM_Oil_Deductions))
NG_Deductions_Tracking <- Reduce(function(x, y) merge(x, y, by = "Filing.Period", all = TRUE), list(Perm_NG_Deductions, SJ_NG_Deductions, NM_NG_Deductions))


###############################################################################
###############################################################################

# Plots and formatting

# Oil Time Series Plot
NM_OilCh <- ggplot(Oil_Tracking, aes(x=Filing.Period, y=NM_Volume)) +
  scale_y_continuous(name= "NM Oil Volume (in barrels)", labels = scales::comma) +
  geom_area(fill="#69b3a2", alpha=0.5) +
  geom_line(color="#69b3a2") +
  xlab("Filing Period")+
  ylab("NM Oil Volume (in barrels")+
  ggtitle("NM Oil Volume") +
  theme_ipsum()
NM_OilCh

# Nat Gas Time Series Plot
NM_GasCh <- ggplot(NG_Tracking, aes(x=Filing.Period, y=NM_Volume)) +
  scale_y_continuous(name= "NM Natural Gas Volume (in MCF)", labels = scales::comma) +
  geom_area(fill="#69b3a2", alpha=0.5) +
  geom_line(color="#69b3a2") +
  xlab("Filing Period")+
  ylab("NM Natural Gas Volume (in MCF)")+
  ggtitle("NM Natural Gas Volume") +
  theme_ipsum()
NM_GasCh

# Oil Price
NM_OilPr <- ggplot(Oil_Tracking, aes(x=Filing.Period, y=NM_Gross_Price)) +
  scale_y_continuous(name= "NM Gross Price", labels = scales::dollar) +
  geom_line() +
  xlab("Filing Period")+
  ylab("NM Oil Price per Barrel")+
  ggtitle("NM Oil Price per Barrel")
NM_OilPr

# Natural Gas Price
NM_GasPr <- ggplot(NG_Tracking, aes(x=Filing.Period, y=NM_Gross_Price)) +
  scale_y_continuous(name= "NM Natural Gas Price per MCF", labels = scales::dollar) +
  geom_line() +
  xlab("Filing Period")+
  ylab("NM Natural Gas Price per MCF")+
  ggtitle("NM Natural Gas Price")
NM_GasPr

################################################################################
# CREATE WORKBOOKS WITH SEPARATE SHEETS #

# Remove and order columns 
Oil_Tracking <- Oil_Tracking %>%
  select(Filing.Period, Permian_Volume, Permian_Gross.Value, Permian_Total.Deductions,
         Permian_Deduction_Percent, Permian_Gross_Price, Permian_Net_Value, SJ_Volume,
         SJ_Gross.Value, SJ_Total.Deductions, SJ_Deduction_Percent, SJ_Gross_Price,
         SJ_Net_Value, NM_Volume, NM_Gross.Value, NM_Total.Deductions, NM_Deduction_Percent,
         NM_Gross_Price, NM_Net_Value)

NG_Tracking <- NG_Tracking %>% 
  select(Filing.Period, Permian_Volume, Permian_Gross.Value, Permian_Total.Deductions,
         Permian_Deduction_Percent, Permian_Gross_Price, Permian_Net_Value, SJ_Volume,
         SJ_Gross.Value, SJ_Total.Deductions, SJ_Deduction_Percent, SJ_Gross_Price,
         SJ_Net_Value, NM_Volume, NM_Gross.Value, NM_Total.Deductions, NM_Deduction_Percent,
         NM_Gross_Price, NM_Net_Value)


# Create workbook and add worksheets
wb <- createWorkbook()
sheetNames <- c("Charts", "Oil Tracker", "Nat Gas Tracker", "Fed Land Oil Tracker",
                "Fed Land Nat Gas Tracker", "State Land Oil Tracker", 
                "State Land Nat Gas Tracker", "Oil Deductions Tracker", 
                "Nat Gas Deductions Tracker")

# Add worksheets in a loop
for (sheet in sheetNames) {
  addWorksheet(wb, sheetName = sheet)
}

# Write data to worksheets
# Insert plots into "Charts" worksheet
chartData <- list(NM_OilCh, NM_GasCh, NM_OilPr, NM_GasPr)
plotPositions <- list(c("A", 1), c("I", 1), c("A", 24), c("I", 24))

for (i in 1:length(chartData)) {
  print(chartData[[i]])
  wb %>% insertPlot(sheet = "Charts", startCol = plotPositions[[i]][1], startRow = plotPositions[[i]][2])
}

# Create styles once
LabelStyle <- createStyle(halign = "center", border = c("bottom", "right"), 
                          borderStyle = "thin", textDecoration = "bold", 
                          fgFill = "#0491A1", fontColour = "white")
BorderStyle <- createStyle(halign = "right", 
                           border = c("bottom", "right", "left"), borderStyle = 'thin')
NumStyle <- createStyle(halign = "right", numFmt = "COMMA", 
                        border = c("bottom", "right", "left", "top"), borderStyle = "thin")
PrcStyle <- createStyle(halign = "right", numFmt = "CURRENCY", 
                        border = c("bottom", "right", "left", "top"), borderStyle = "thin")
PrcntgStyle <- createStyle(halign = "right", numFmt = "0.00%", 
                           border = c("bottom", "right", "left", "top"), borderStyle = "thin")
TextStyle <- createStyle(halign = "center", 
                         border = c("bottom", "right", "left", "top"), borderStyle = "thin")
DateStyle <- createStyle(halign = "center", numFmt = "mm/dd/yyyy", 
                         border = c("bottom", "right", "left", "top"), borderStyle = "thin")
BkGrdStyle <- createStyle(fgFill = "#FFFFFF", 
                          border = c("bottom", "right", "left", "top"), borderStyle = "thin")
DollarStyle <- createStyle(halign = 'right', numFmt = "$#,##0.00", 
                           border = c("bottom", "right", "left", "top"), borderStyle = "thin")


# Define a function to add styles with dynamic range setting based on data
applyStyles <- function(sheetName, data) {
  numCols <- ncol(data)
  numRows <- nrow(data) + 1 # plus one for the header row
  
  # Apply general styles
  addStyle(wb, sheet = sheetName, style = LabelStyle, rows = 1, cols = 1:numCols, gridExpand = FALSE, stack = FALSE)
  addStyle(wb, sheet = sheetName, style = BorderStyle, rows = 1:numRows, cols = 1:numCols, gridExpand = TRUE, stack = TRUE)
  addStyle(wb, sheet = sheetName, style = DateStyle, rows = 2:numRows, cols = 1, gridExpand = FALSE, stack = FALSE)
  addStyle(wb, sheet = sheetName, style = TextStyle, rows = 2:numRows, cols = 1, gridExpand = FALSE, stack = FALSE)
  addStyle(wb, sheet = sheetName, style = BkGrdStyle, rows = 2:numRows, cols = 1, gridExpand = FALSE, stack = FALSE)
  # Apply number styles based on column names
  for (col in 2:numCols) {
    colName <- colnames(data)[col]
    if (grepl("Volume", colName)) {
      addStyle(wb, sheet = sheetName, style = NumStyle, rows = 2:numRows, cols = col, gridExpand = FALSE, stack = FALSE)
    } else if (grepl("Percent", colName)) {
      addStyle(wb, sheet = sheetName, style = PrcntgStyle, rows = 2:numRows, cols = col, gridExpand = FALSE, stack = FALSE)
    } else {
      addStyle(wb, sheet = sheetName, style = PrcStyle, rows = 2:numRows, cols = col, gridExpand = FALSE, stack = FALSE)
    }
  }
}

# Change column names to remove . and _
format_column_names <- function(df) {
  colnames(df) <- gsub("[_.]", " ", colnames(df))
  return(df)
}

# Apply format_column_names to all dataframes in the environment
list_of_dataframes <- Filter(function(x) is.data.frame(get(x)), ls())
for (df_name in list_of_dataframes) {
  assign(df_name, format_column_names(get(df_name)))
}


# Create a mapping between sheet names and corresponding dataframe names
sheet_to_df_mapping <- c(
  "Oil Tracker" = 'Oil_Tracking',
  "Nat Gas Tracker" = 'NG_Tracking',
  "Fed Land Oil Tracker" = 'fed_oil_comb',
  "Fed Land Nat Gas Tracker" = 'fed_gas_comb',
  "State Land Oil Tracker" = 'state_oil_comb',
  "State Land Nat Gas Tracker" = 'state_gas_comb',
  "Oil Deductions Tracker"= 'Oil_Deductions_Tracking',
  "Nat Gas Deductions Tracker" = 'NG_Deductions_Tracking'
)

# Apply styles to each sheet
for (sheet in sheetNames[-1]) { # Excluding "Charts"
  if (sheet %in% names(sheet_to_df_mapping)) {
    df_name <- sheet_to_df_mapping[sheet]
    if (exists(df_name)) {
      df <- get(df_name)
      applyStyles(sheet, df)
    } else {
      warning(paste("Dataframe", df_name, "not found for sheet", sheet))
    }
  } else {
    warning(paste("No mapping found for sheet", sheet))
  }
}

# Write data to each sheet
writeData(wb, "Oil Tracker", Oil_Tracking, startCol = 1, startRow = 1, colNames = TRUE)
writeData(wb, "Nat Gas Tracker", NG_Tracking, startCol = 1, startRow = 1, colNames = TRUE)
writeData(wb, "Fed Land Oil Tracker", fed_oil_comb, startCol = 1, startRow = 1, colNames = TRUE)
writeData(wb, "Fed Land Nat Gas Tracker", fed_gas_comb, startCol = 1, startRow = 1, colNames = TRUE)
writeData(wb, "State Land Oil Tracker", state_oil_comb, startCol = 1, startRow = 1, colNames = TRUE)
writeData(wb, "State Land Nat Gas Tracker", state_gas_comb, startCol = 1, startRow = 1, colNames = TRUE)
writeData(wb, "Oil Deductions Tracker", Oil_Deductions_Tracking, startCol = 1, startRow = 1, colNames = TRUE)
writeData(wb, "Nat Gas Deductions Tracker", NG_Deductions_Tracking, startCol = 1, startRow = 1, colNames = TRUE)


# Get current date
current_date <- Sys.Date()

# Format the date as MM.DD.YY
formatted_date <- format(current_date, "%m.%d.%y")

# Create the dynamic filename
filename <- paste("results/Oil and Gas Tracker ", formatted_date, ".xlsx", sep = "")

# Save the workbook
saveWorkbook(wb, filename, overwrite = TRUE)

