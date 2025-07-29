import pandas as pd
import numpy as np
pd.set_option('display.max_columns', None)


finbo_Opp=pd.read_excel("CX Americas BCS FY26 Prelim Forecast_L2_Third Pass (1).xlsx",sheet_name="FinBO")
finbo_Opp=finbo_Opp[finbo_Opp["Sales Level 2"]=="US PS Market Segment"]

finbo_quarterly_sums_by_sales_level_3 = finbo_Opp.groupby(["Sales Level 3", "Fiscal Quarter ID"])["Annual Bookings Net"].sum()

# Reset the index to make the grouped data easier to work with
finbo_quarterly_sums_by_sales_level_3 = finbo_quarterly_sums_by_sales_level_3.reset_index()

# Pivot the data to get a wide format with quarters as columns
finbo_quarterly_sums_pivot = finbo_quarterly_sums_by_sales_level_3.pivot(
    index="Sales Level 3",
    columns="Fiscal Quarter ID",
    values="Annual Bookings Net"
)

# Optionally fill NaN values with 0 (if there are missing quarters for some Sales Level 3)
finbo_quarterly_sums_pivot = finbo_quarterly_sums_pivot.fillna(0)

finbo_rows_to_drop = ["US PS Market Segment-MISCL3", "US Public Sector Misc."]
finbo_quarterly_sums_pivot = finbo_quarterly_sums_pivot.drop(index=finbo_rows_to_drop)
# Remove the "Q" from the column names
finbo_quarterly_sums_pivot.columns = finbo_quarterly_sums_pivot.columns.str.replace("Q", "", regex=False)

#reorder to prevent issues when performing operations on the tables
finbo_quarterly_sums_pivot = finbo_quarterly_sums_pivot.sort_values(by="Sales Level 3")

# Print the resulting DataFrame
print(finbo_quarterly_sums_pivot)


#CRBO
CRBO_Opp=pd.read_excel("CX Americas BCS FY26 Prelim Forecast_L2_Third Pass (1).xlsx",sheet_name="CRBO")
CRBO_Opp=CRBO_Opp[CRBO_Opp["DRR Level 2 Name"]=="US PS Market Segment"]

CRBO_quarterly_sums_by_sales_level_3 = CRBO_Opp.groupby(["DRR Level 3 Name", "Service Contract End Quarter"])["Opportunity - Annual"].sum()

# Reset the index to make the grouped data easier to work with
CRBO_quarterly_sums_by_sales_level_3 = CRBO_quarterly_sums_by_sales_level_3.reset_index()

# Pivot the data to get a wide format with quarters as columns
CRBO_quarterly_sums_pivot = CRBO_quarterly_sums_by_sales_level_3.pivot(
    index="DRR Level 3 Name",
    columns="Service Contract End Quarter",
    values="Opportunity - Annual"
)

# Optionally fill NaN values with 0 (if there are missing quarters for some Sales Level 3)
CRBO_quarterly_sums_pivot = CRBO_quarterly_sums_pivot.fillna(0)

CRBO_rows_to_drop = ["US Public Sector Misc."]
CRBO_quarterly_sums_pivot = CRBO_quarterly_sums_pivot.drop(index=CRBO_rows_to_drop)

# Print the resulting DataFrame
print(CRBO_quarterly_sums_pivot)



#Conversions

Y6Dropped_CRBO = CRBO_quarterly_sums_pivot.iloc[:, :-4]

#I need to align the two tables before dividing them

Y6Dropped_CRBO.index.name = "Sales Level 3"
Y6Dropped_CRBO.columns = Y6Dropped_CRBO.columns.astype(str)

#divide FinBo by CRBO
Conversions = finbo_quarterly_sums_pivot / Y6Dropped_CRBO

#add in blank columns for 20261-20264
# Add blank columns for 20261, 20262, 20263, and 20264 with NaN as default value
Conversions["20261"] = float("nan")
Conversions["20262"] = float("nan")
Conversions["20263"] = float("nan")
Conversions["20264"] = float("nan")

#define a function for forcasting the quarters conversion rate



# Define the logic for forecasting
def forecast_value(row, previous_quarters):
    # Get the values from the previous quarters
    previous_values = [row[quarter] for quarter in previous_quarters]

    # Calculate the average of the previous values
    avg = np.mean(previous_values)

    # Apply the logic based on the average
    if avg > 1:  # Greater than 100% (assuming proportions)
        return min(previous_values)  # Use the minimum value
    else:  # Less than or equal to 100%
        return avg  # Use the average value


# Define the mapping for each placeholder column and its previous quarters
forecast_quarters = {
    "20261": ["20231", "20241", "20251"],
    "20262": ["20232", "20242", "20252"],
    "20263": ["20233", "20243", "20253"],
    "20264": ["20234", "20244", "20254"]
}

# Apply the forecasting logic to each placeholder column
for current_quarter, previous_quarters in forecast_quarters.items():
    Conversions[current_quarter] = Conversions.apply(
        forecast_value, axis=1, previous_quarters=previous_quarters)

# Print the updated DataFrame
print(Conversions.round(2))

#Go back to FinBo to do conversion * crbo to get 20261-20264

#create place holders
finbo_quarterly_sums_pivot["20261"] = float("nan")
finbo_quarterly_sums_pivot["20262"] = float("nan")
finbo_quarterly_sums_pivot["20263"] = float("nan")
finbo_quarterly_sums_pivot["20264"] = float("nan")

#Change CRBO cols to str to match finbo and conversions
CRBO_quarterly_sums_pivot.columns = CRBO_quarterly_sums_pivot.columns.astype(str)

#insert vaules into those placeholders
columns_to_update = ["20261", "20262", "20263", "20264"]

for col in columns_to_update:
    finbo_quarterly_sums_pivot[col] = Conversions[col] * CRBO_quarterly_sums_pivot[col]

#theirs are inflated thats why im off by a few million
finbo_quarterly_sums_pivot["20261"].sum()




##################################################### L4 #############################################################





finbo_quarterly_sums_by_sales_level_4 = finbo_Opp.groupby(["Sales Level 4", "Fiscal Quarter ID"])["Annual Bookings Net"].sum()

# Reset the index to make the grouped data easier to work with
finbo_quarterly_sums_by_sales_level_4 = finbo_quarterly_sums_by_sales_level_4.reset_index()

# Pivot the data to get a wide format with quarters as columns
finbo_quarterly_sums_pivot_level_4 = finbo_quarterly_sums_by_sales_level_4.pivot(
    index="Sales Level 4",
    columns="Fiscal Quarter ID",
    values="Annual Bookings Net"
)

# Optionally fill NaN values with 0 (if there are missing quarters for some Sales Level 3)
finbo_quarterly_sums_pivot_level_4 = finbo_quarterly_sums_pivot_level_4.fillna(0)

#finbo_rows_to_drop = ["US PS Market Segment-MISCL3", "US Public Sector Misc."]
#finbo_quarterly_sums_pivot = finbo_quarterly_sums_pivot.drop(index=finbo_rows_to_drop)

# Remove the "Q" from the column names
finbo_quarterly_sums_pivot_level_4.columns = finbo_quarterly_sums_pivot_level_4.columns.str.replace("Q", "", regex=False)

#reorder to prevent issues when performing operations on the tables
finbo_quarterly_sums_pivot_level_4 = finbo_quarterly_sums_pivot_level_4.sort_values(by="Sales Level 4")

# Print the resulting DataFrame
print(finbo_quarterly_sums_pivot_level_4)

###### CRBO

CRBO_Opp=pd.read_excel("CX Americas BCS FY26 Prelim Forecast_L2_Third Pass (1).xlsx",sheet_name="CRBO")
CRBO_Opp=CRBO_Opp[CRBO_Opp["DRR Level 2 Name"]=="US PS Market Segment"]

CRBO_quarterly_sums_by_sales_level_4 = CRBO_Opp.groupby(["DRR Level 4 Name", "Service Contract End Quarter"])["Opportunity - Annual"].sum()

# Reset the index to make the grouped data easier to work with
CRBO_quarterly_sums_by_sales_level_4 = CRBO_quarterly_sums_by_sales_level_4.reset_index()

# Pivot the data to get a wide format with quarters as columns
CRBO_quarterly_sums_pivot_level_4 = CRBO_quarterly_sums_by_sales_level_4.pivot(
    index="DRR Level 4 Name",
    columns="Service Contract End Quarter",
    values="Opportunity - Annual"
)

# Optionally fill NaN values with 0 (if there are missing quarters for some Sales Level 3)
CRBO_quarterly_sums_pivot_level_4 = CRBO_quarterly_sums_pivot_level_4.fillna(0)

#CRBO_rows_to_drop = ["US Public Sector Misc."]
#CRBO_quarterly_sums_pivot = CRBO_quarterly_sums_pivot.drop(index=CRBO_rows_to_drop)

# Print the resulting DataFrame
print(CRBO_quarterly_sums_pivot_level_4)

### stopping point. rows not lining up, dont know if i can drop them. Can i just keep the 4 fed and 8 sled
CRBO_quarterly_sums_pivot_level_4 = pd.DataFrame(CRBO_quarterly_sums_pivot_level_4)
CRBO_quarterly_sums_pivot_level_4[20221].sum()
