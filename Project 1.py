
import pandas as pd
import numpy as np
pd.options.display.float_format = '{:,.0f}'.format
import matplotlib.pyplot as plt
plt.switch_backend('agg')

##################################################################
prefix = "L5"  # Define the prefix you want to match
variables_to_delete = [var for var in globals() if var.startswith(prefix)]

# Delete each variable with the matching prefix
for var in variables_to_delete:
    del globals()[var]
#################################################################


#To excel template######################

excel_file_name = 'Delivarble_7.07.xlsx'

with pd.ExcelWriter(excel_file_name, engine='xlsxwriter') as writer:
    _242526.to_excel(writer, sheet_name='L5_ATR_STAT', index=True)
    _242526_L6.to_excel(writer, sheet_name='L6_ATR_STAT', index=True)




#######################################



#Removed ATR read in for privacy reasons
ATR = ATR[~ATR["Stage Name"].isin(["6 - Closed Lost", "6 - Closed Cancelled"])]


###### ATR L5 ####
L5_ATR = ATR[["Fiscal Year", "SL5", "Annual ATR"]]

L5_ATR_SLED = L5_ATR[L5_ATR['SL5'].str.contains('SLED', na=False)]

L5_ATR_SLED_FY24 = L5_ATR_SLED[L5_ATR_SLED['Fiscal Year'] == 2024]
L5_ATR_SLED_FY24.rename(columns={'Annual ATR' : 'Annual ATR 2024'}, inplace=True)
L5_ATR_SLED_FY24 = L5_ATR_SLED_FY24.groupby("SL5")["Annual ATR 2024"].sum()
L5_ATR_SLED_FY24 = pd.DataFrame(L5_ATR_SLED_FY24)

L5_ATR_SLED_FY25 = L5_ATR_SLED[L5_ATR_SLED['Fiscal Year'] == 2025]
L5_ATR_SLED_FY25.rename(columns={'Annual ATR': 'Annual ATR 2025'}, inplace=True)
L5_ATR_SLED_FY25 = L5_ATR_SLED_FY25.groupby("SL5")["Annual ATR 2025"].sum()
L5_ATR_SLED_FY25 = pd.DataFrame(L5_ATR_SLED_FY25)

L5_ATR_SLED_FY26 = L5_ATR_SLED[L5_ATR_SLED['Fiscal Year'] == 2026]
L5_ATR_SLED_FY26.rename(columns={'Annual ATR': 'Annual ATR 2026'}, inplace=True)
L5_ATR_SLED_FY26 = L5_ATR_SLED_FY26.groupby("SL5")["Annual ATR 2026"].sum()
L5_ATR_SLED_FY26 = pd.DataFrame(L5_ATR_SLED_FY26)

_242526 = pd.merge(
    pd.merge(L5_ATR_SLED_FY24, L5_ATR_SLED_FY25, on='SL5', how='inner', suffixes=('_2024', '_2025')),
    L5_ATR_SLED_FY26,
    on='SL5',
    how='inner')

_242526.rename(columns={'Annual ATR 2024': 'ATR 2024'}, inplace=True)
_242526.rename(columns={'Annual ATR 2025': 'ATR 2025'}, inplace=True)
_242526.rename(columns={'Annual ATR 2026': 'ATR 2026'}, inplace=True)

#Growth Rate Calculation
_242526['YoY Growth 2025 (%)'] = ((_242526['ATR 2025'] - _242526['ATR 2024']) / _242526['ATR 2024']) * 100
_242526['YoY Growth 2026 (%)'] = ((_242526['ATR 2026'] - _242526['ATR 2025']) / _242526['ATR 2025']) * 100

# Replace infinite values with NaN (if they exist)
_242526.replace([np.inf, -np.inf], np.nan, inplace=True)

# Drop rows with NaN values in relevant columns
_242526.dropna(subset=['ATR 2024', 'ATR 2025', 'ATR 2026'], inplace=True)

#Z score outlier detection

for year in ['2024', '2025', '2026']:
    column_name_L5 = f'ATR {year}'

    # Calculate mean and standard deviation for the year
    mean_L5 = _242526[column_name_L5].mean()
    std_L5 = _242526[column_name_L5].std()

    # Compute z-scores
    _242526[f'Z-Score {year}'] = (_242526[column_name_L5] - mean_L5) / std_L5

    # Flag outliers (Z-Score > 3 or < -3)
    _242526[f'Outlier {year}'] = _242526[f'Z-Score {year}'].apply(
        lambda z: 'Outlier' if abs(z) > 3 else 'Normal'
    )


# Performed Ztest instead of Ttest bc this is not a sample, this is the population.




##### ATR L6 #################################
L6_ATR = ATR[["Fiscal Year", "SL5", "SL6", "Annual ATR"]]

L6_ATR_SLED = L6_ATR[L6_ATR['SL5'].str.contains('SLED', na=False)]

L6_ATR_SLED = L6_ATR_SLED.drop(columns=['SL5'])

L6_ATR_SLED_FY24 = L6_ATR_SLED[L6_ATR_SLED['Fiscal Year'] == 2024]
L6_ATR_SLED_FY24.rename(columns={'Annual ATR' : 'Annual ATR 2024'}, inplace=True)
L6_ATR_SLED_FY24 = L6_ATR_SLED_FY24.groupby("SL6")["Annual ATR 2024"].sum()
L6_ATR_SLED_FY24 = pd.DataFrame(L6_ATR_SLED_FY24)


L6_ATR_SLED_FY25 = L6_ATR_SLED[L6_ATR_SLED['Fiscal Year'] == 2025]
L6_ATR_SLED_FY25.rename(columns={"Annual ATR" : "Annual ATR 2025"}, inplace=True)
L6_ATR_SLED_FY25 = L6_ATR_SLED_FY25.groupby("SL6")["Annual ATR 2025"].sum()
L6_ATR_SLED_FY25 = pd.DataFrame(L6_ATR_SLED_FY25)


L6_ATR_SLED_FY26 = L6_ATR_SLED[L6_ATR_SLED['Fiscal Year'] == 2026]
L6_ATR_SLED_FY26.rename(columns={'Annual ATR' : 'Annual ATR 2026'}, inplace=True)
L6_ATR_SLED_FY26 = L6_ATR_SLED_FY26.groupby("SL6")["Annual ATR 2026"].sum()
L6_ATR_SLED_FY26 = pd.DataFrame(L6_ATR_SLED_FY26)

_242526_L6 = pd.merge(
    pd.merge(L6_ATR_SLED_FY24, L6_ATR_SLED_FY25, on='SL6', how='inner', suffixes=('_2024', '_2025')),
    L6_ATR_SLED_FY26,
    on='SL6',
    how='inner')

_242526_L6.rename(columns={'Annual ATR 2024': 'ATR 2024'}, inplace=True)
_242526_L6.rename(columns={'Annual ATR 2025': 'ATR 2025'}, inplace=True)
_242526_L6.rename(columns={'Annual ATR 2026': 'ATR 2026'}, inplace=True)

print(_242526_L6)

#### Growth Rate
_242526_L6['YoY Growth 2025 (%)'] = ((_242526_L6['ATR 2025'] - _242526_L6['ATR 2024']) / _242526_L6['ATR 2024']) * 100
_242526_L6['YoY Growth 2026 (%)'] = ((_242526_L6['ATR 2026'] - _242526_L6['ATR 2025']) / _242526_L6['ATR 2025']) * 100

print(_242526_L6)

#### fixing errors when calculating outliers
_242526_L6.replace([np.inf, -np.inf], np.nan, inplace=True)

# Drop rows with NaN values in relevant columns
_242526_L6.dropna(subset=['ATR 2024', 'ATR 2025', 'ATR 2026'], inplace=True)

for year in ['2024', '2025', '2026']:
    column_name = f'ATR {year}'

    # Calculate mean and standard deviation for the year
    mean = _242526_L6[column_name].mean()
    std = _242526_L6[column_name].std()

    # Compute z-scores
    _242526_L6[f'Z-Score {year}'] = (_242526_L6[column_name] - mean) / std

    # Flag outliers (Z-Score > 3 or < -3)
    _242526_L6[f'Outlier {year}'] = _242526_L6[f'Z-Score {year}'].apply(
        lambda z: 'Outlier' if abs(z) > 3 else 'Normal'
    )

# Display DataFrame with outlier information
print(_242526_L6)

print_outliers = (
    (_242526_L6["Outlier 2024"] != "Normal") |
    (_242526_L6["Outlier 2025"] != "Normal") |
    (_242526_L6["Outlier 2026"] != "Normal")
)

outliers_only = _242526_L6[print_outliers]

# Plot only the SL5 groups with outliers
for sl5_value in outliers_only.index:
    atr_values = outliers_only.loc[sl5_value, ['ATR 2024', 'ATR 2025', 'ATR 2026']]
    plt.figure(figsize=(11, 7))
    plt.plot(['2024', '2025', '2026'], atr_values, marker='o', linestyle='-', color='g', label=f'SL5: {sl5_value}')

    plt.title(f'ATR Trend for {sl5_value} (With Outlier)', fontsize=14)
    plt.xlabel('Fiscal Year', fontsize=12)
    plt.ylabel('Annual ATR', fontsize=12)
    plt.grid(True, which='both', linestyle='--', linewidth=0.5)

    plt.legend()
    plt.tight_layout()
    plt.show()

    filename = f'ATR_Trend_{sl5_value}.png'
    plt.savefig(filename)

    plt.close()



###### What percentage of services makes up L5 and L6 ATR?

############################################################################
#ATR line chart split

L5_ATR_Both = ATR[["Fiscal Year", "Annual ATR", "Selected Dimension 1"]]
L5_ATR_Both_filtered = L5_ATR_Both[L5_ATR_Both["Fiscal Year"].isin([2024, 2025, 2026])]

# Group by 'Fiscal Year' and 'Selected Dimension 1' to get the sum of 'Annual ATR'
# for each category in each year
L5_ATR_Both_SUM = L5_ATR_Both_filtered.groupby(["Fiscal Year", "Selected Dimension 1"])["Annual ATR"].sum().unstack()

# Create the line chart
plt.figure(figsize=(16, 9)) # Adjust figure size for better readability

# Plot each category as a separate line
# Assuming 'Selected Dimension 1' contains 'Product' and 'Services'
# Check if 'Product' and 'Services' columns exist after unstacking
if 'Product' in L5_ATR_Both_SUM.columns:
    plt.plot(L5_ATR_Both_SUM.index, L5_ATR_Both_SUM['Product'], marker='o', label='Product', color='cornflowerblue', linewidth=2)
if 'Services' in L5_ATR_Both_SUM.columns:
    plt.plot(L5_ATR_Both_SUM.index, L5_ATR_Both_SUM['Services'], marker='o', label='Services', color='mediumseagreen', linewidth=2)

# Add title and labels
plt.title("Annual ATR Trend: Product vs. Services (2024-2026)", y=1.05, fontweight='bold', fontsize=20)
plt.xlabel("Fiscal Year", fontsize=15, fontweight='bold')
plt.ylabel("Annual ATR", fontsize=15,  fontweight='bold')

# Set x-axis ticks to be exactly the fiscal years
plt.xticks(L5_ATR_Both_SUM.index)

# Add grid for better readability
plt.grid(True, linestyle='--', alpha=0.7)

# Add legend to distinguish lines
plt.legend(title="Category", loc='best')

# Add data labels to each point
for category in L5_ATR_Both_SUM.columns:
    for i, txt in enumerate(L5_ATR_Both_SUM[category]):
        plt.annotate(f'${txt:,.0f}', # Format as currency
                     (L5_ATR_Both_SUM.index[i], L5_ATR_Both_SUM[category].iloc[i]),
                     textcoords="offset points",
                     xytext=(0,7),
                     ha='center',
                     fontsize=10,        # Increased font size
                     color='black',
                     fontweight='bold')  # Made font bold


# Improve layout
plt.tight_layout()

# Save the figure
filename_Both = "ATR_Product_Services_Line_Chart.png"
plt.savefig(filename_Both)
plt.close()

#NOW DOING SERVICES

L5_ATR_Services = ATR[["Fiscal Year", "Annual ATR", "Selected Dimension 1", "Selected Dimension 2"]]
L5_ATR_Services_filtered = L5_ATR_Services[
    (L5_ATR_Services["Selected Dimension 1"] == "Services") &
    (L5_ATR_Services["Fiscal Year"].isin([2024, 2025, 2026]))]

# Group by 'Fiscal Year' and 'Selected Dimension 2' to get the sum of 'Annual ATR'
# for each service type in each year
L5_ATR_Services_SUM = L5_ATR_Services_filtered.groupby(["Fiscal Year", "Selected Dimension 2"])["Annual ATR"].sum().unstack()

# Create the line chart
plt.figure(figsize=(16, 9)) # Adjust figure size for better readability

# Plot each service type as a separate line
# Check if 'TSS' and 'LCS' columns exist after unstacking
if 'TSS' in L5_ATR_Services_SUM.columns:
    plt.plot(L5_ATR_Services_SUM.index, L5_ATR_Services_SUM['TSS'], marker='o', label='TSS', color='orchid', linewidth=2)
if 'LCS' in L5_ATR_Services_SUM.columns:
    plt.plot(L5_ATR_Services_SUM.index, L5_ATR_Services_SUM['LCS'], marker='o', label='LCS', color='mediumseagreen', linewidth=2)

# Add title and labels
plt.title("Annual ATR Trend: TSS vs. LCS Services (2024-2026)", y=1.05, fontweight='bold', fontsize=20)
plt.xlabel("Fiscal Year", fontsize=15, fontweight='bold')
plt.ylabel("Annual ATR", fontsize=15, fontweight='bold')

# Set x-axis ticks to be exactly the fiscal years
plt.xticks(L5_ATR_Services_SUM.index)

# Add grid for better readability
plt.grid(True, linestyle='--', alpha=0.7)

# Add legend to distinguish lines
plt.legend(title="Service Type", loc='best')

# Add data labels to each point
for service_type in L5_ATR_Services_SUM.columns:
    for i, txt in enumerate(L5_ATR_Services_SUM[service_type]):
        plt.annotate(f'${txt:,.0f}',  # Format as currency
                     (L5_ATR_Services_SUM.index[i], L5_ATR_Services_SUM[service_type].iloc[i]),
                     textcoords="offset points",
                     xytext=(0, -20),
                     ha='center',
                     fontsize=12,  # Increased font size
                     color='black',
                     fontweight='bold')

# Improve layout
plt.tight_layout()

# Save the figure
filename_services = "Services_TSS_LCS_Line_Chart.png"
plt.savefig(filename_services)
plt.close()


### Products
L5_ATR_Products = ATR[["Fiscal Year", "Annual ATR", "Selected Dimension 1", "Selected Dimension 2"]]
L5_ATR_Products_filtered = L5_ATR_Products[
    (L5_ATR_Products["Selected Dimension 1"] == "Product") &
    (L5_ATR_Products["Fiscal Year"].isin([2024, 2025, 2026]))
]

# Group by 'Fiscal Year' and 'Selected Dimension 2' to get the sum of 'Annual ATR'
# for each product type in each year
L5_ATR_Products_SUM = L5_ATR_Products_filtered.groupby(["Fiscal Year", "Selected Dimension 2"])["Annual ATR"].sum().unstack()

# Create the figure
# Increased figure width significantly to accommodate the large legend
fig = plt.figure(figsize=(16, 9))

# Add an axes to the figure for the line chart
# [left, bottom, width, height] in figure coordinates (0 to 1)
# Set width to leave space for the legend on the right (e.g., 0.7 means 70% of figure width for plot)
ax = fig.add_axes([0.05, 0.1, 0.7, 0.8]) # Plotting area starts at 5% from left, 10% from bottom, takes 70% width, 80% height

# Define a color palette for the lines
colors = plt.cm.get_cmap('tab10', len(L5_ATR_Products_SUM.columns))

# Plot each product type as a separate line on the 'ax' object
for i, product_type in enumerate(L5_ATR_Products_SUM.columns):
    ax.plot(L5_ATR_Products_SUM.index, L5_ATR_Products_SUM[product_type],
             marker='o', label=product_type, color=colors(i), linewidth=2)

# Add title and labels to the 'ax' object
ax.set_title("Annual ATR Trend: Product Categories (2024-2026)", y=1.05, fontweight='bold', fontsize=20)
ax.set_xlabel("Fiscal Year", fontsize=15, fontweight='bold')
ax.set_ylabel("Annual ATR", fontsize=15, fontweight='bold')

# Set x-axis ticks to be exactly the fiscal years
ax.set_xticks(L5_ATR_Products_SUM.index)

# Add grid for better readability
ax.grid(True, linestyle='--', alpha=0.7)

# Add data labels to each point with improved staggering
base_vertical_offset = -5 # Initial offset above the point
stagger_step = -5 # Vertical distance between staggered labels

for fiscal_year in L5_ATR_Products_SUM.index:
    year_data_points = []
    for product_type in L5_ATR_Products_SUM.columns:
        value = L5_ATR_Products_SUM.loc[fiscal_year, product_type]
        if pd.notna(value):
            year_data_points.append((value, product_type))

    year_data_points.sort(key=lambda x: x[0])

    num_points = len(year_data_points)
    for i, (value, product_type) in enumerate(year_data_points):
        offset_multiplier = (i - (num_points - 1) / 2)
        vertical_offset = base_vertical_offset + offset_multiplier * stagger_step

        ax.annotate(f'${value:,.0f}',
                     (fiscal_year, value),
                     textcoords="offset points",
                     xytext=(0, vertical_offset),
                     ha='center',
                     fontsize=9,
                     color='dimgray',
                     bbox=dict(boxstyle="round,pad=0.2", fc="white", alpha=0.7, ec="none")
                    )

# Add legend to distinguish lines
# bbox_to_anchor=(1.01, 0.5) places the legend just outside the right edge of 'ax'
# loc='center left' aligns the legend's center-left point to the bbox_to_anchor
ax.legend(title="Product Category", loc='center left', bbox_to_anchor=(1.01, 0.5), borderaxespad=0,
          prop={'size': 14},        # Adjusted font size for legend labels
          title_fontsize='16',      # Adjusted title font size
          borderpad=1.5,
          labelspacing=1.2,
          handlelength=2.5,
          handletextpad=0.8)

# Removed plt.tight_layout() as we've manually set axes position

# Save the figure
filename_products = "Products_Categories_Line_Chart.png"
plt.savefig(filename_products)
plt.close()




#############################################################################

# Define a custom function to format the absolute values (for pie charts)
def func_absolute_value(pct, allvals):
    # Calculate the absolute value from the percentage
    # pct is the percentage (e.g., 25.0 for 25%)
    # allvals is the array of all values (sizes_services in this case)
    absolute_value = int(round(pct/100.*sum(allvals)))
    # Format the value. You can add currency symbols, thousands separators, etc.
    return f'${absolute_value:,.0f}' # Formats as $123,456

# ATR SPLIT SERVICES VS PRODUCTS
L5_ATR_Both = ATR[["Fiscal Year", "Annual ATR","Selected Dimension 1"]]
L5_ATR_Both = L5_ATR_Both[(L5_ATR_Both["Fiscal Year"] == 2026)]

L5_ATR_Both_SUM = L5_ATR_Both.groupby("Selected Dimension 1")["Annual ATR"].sum()

labels_Both = L5_ATR_Both_SUM.index if hasattr(L5_ATR_Both_SUM, 'index') else ["TSS", "LCS"]
sizes_Both = L5_ATR_Both_SUM.values if hasattr(L5_ATR_Both_SUM, 'values') else L5_ATR_Both_SUM

plt.figure(figsize=(16, 9))

patches, texts, autotexts = plt.pie(sizes_Both,
        labels=labels_Both,
        autopct=lambda pct: func_absolute_value(pct, sizes_Both),
        startangle=90,
        colors=("cornflowerblue", "mediumseagreen"),
        wedgeprops={'linewidth': 1, 'edgecolor': 'white'},  # Add borders: 1 unit wide, white color
        shadow=True,
        labeldistance=1.05,
        pctdistance=0.5
        )

for text in texts:
    text.set_fontsize(18) # Adjust font size as needed
    text.set_fontweight('bold')

for autotext in autotexts:
    autotext.set_color('black')       # Set color to black for contrast
    autotext.set_fontsize(22)         # Increase font size [9]
    autotext.set_fontweight('bold')   # Make font bold [8]

plt.title("Annual ATR Distribution Split for 2026", y=1.05, fontweight = 'bold', fontsize=20)
plt.axis('equal')  # Equal aspect ratio ensures the pie chart is circular.
#plt.show()

# Calculate the total sum of ATR
total_atr_sum_both = L5_ATR_Both_SUM.sum()
formatted_total_atr_both = f'Total Annual ATR: ${total_atr_sum_both:,.0f}'

# Add the total sum as a footnote
# x and y coordinates are in figure fraction (0,0 is bottom-left, 1,1 is top-right)
plt.figtext(0.5, 0.02, formatted_total_atr_both, ha="center", fontsize=20, bbox={"facecolor":"white", "alpha":0.5, "pad":5})



filename_Both = "Both.png"
plt.savefig(filename_Both)
plt.close()


plt.figure(figsize=(16, 9))
L5_ATR_Services = ATR[["Fiscal Year", "Annual ATR","Selected Dimension 1", "Selected Dimension 2"]]
L5_ATR_Services = L5_ATR_Services[(L5_ATR_Services["Fiscal Year"] == 2026) & (L5_ATR_Services["Selected Dimension 1"] == "Services")]
L5_ATR_Services = L5_ATR_Services.drop(columns = ["Selected Dimension 1"])

L5_ATR_Services_SUM = L5_ATR_Services.groupby("Selected Dimension 2")["Annual ATR"].sum()

labels_services = L5_ATR_Services_SUM.index if hasattr(L5_ATR_Services_SUM, 'index') else ["TSS", "LCS"]
sizes_services = L5_ATR_Services_SUM.values if hasattr(L5_ATR_Services_SUM, 'values') else L5_ATR_Services_SUM

patches, texts, autotexts = plt.pie(sizes_services,
        labels=labels_services,
        autopct=lambda pct: func_absolute_value(pct, sizes_services),
        startangle=90,
        colors=("lightcoral", "cornflowerblue"),
        wedgeprops={'linewidth': 1, 'edgecolor': 'white'},  # Add borders: 1 unit wide, white color
        shadow=True,
        labeldistance=1.05,
        pctdistance=0.6
        )

for text in texts:
    text.set_fontsize(18) # Adjust font size as needed
    text.set_fontweight('bold')

for autotext in autotexts:
    autotext.set_color('black')       # Set color to black for contrast
    autotext.set_fontsize(22)         # Increase font size [9]
    autotext.set_fontweight('bold')

total_atr_sum_services = L5_ATR_Services_SUM.sum()
formatted_total_atr_services = f'Total Annual ATR: ${total_atr_sum_services:,.0f}'

# Add the total sum as a footnote
# x and y coordinates are in figure fraction (0,0 is bottom-left, 1,1 is top-right)
plt.figtext(0.5, 0.02, formatted_total_atr_services, ha="center", fontsize=20, bbox={"facecolor":"white", "alpha":0.5, "pad":5})

plt.title("Annual ATR Distribution For Services in 2026", fontweight='bold', fontsize=20) # More descriptive title
plt.axis('equal')

filename_services = "Services_Split.png"
plt.savefig(filename_services)
plt.close() # Close the figure for 'Services'


fig = plt.figure(figsize=(16, 9))
# --- CHANGE THIS LINE ---
# Use fig.suptitle() for the overall figure title
fig.suptitle("Annual ATR Distribution For Products in 2026", fontweight='bold', fontsize=20)

# Add an axes to the figure for the pie chart
# [left, bottom, width, height] in figure coordinates (0 to 1)
# Adjust 'left' to move the pie chart to the left (e.g., 0.05 for 5% from left edge)
# Adjust 'width' to control the size of the pie chart within its allocated space
ax = fig.add_axes([0.05, 0.1, 0.45, 0.8]) # Pie chart axes starts at 5% from left, 10% from bottom, takes 45% width, 80% height

L5_ATR_Products = ATR[["Fiscal Year", "Annual ATR","Selected Dimension 1", "Selected Dimension 2"]]
L5_ATR_Products = L5_ATR_Products[(L5_ATR_Products["Fiscal Year"] == 2026) & (L5_ATR_Products["Selected Dimension 1"] == "Product")]
L5_ATR_Products = L5_ATR_Products.drop(columns = ["Selected Dimension 1"])

L5_ATR_Products_SUM = L5_ATR_Products.groupby("Selected Dimension 2")["Annual ATR"].sum()

L5_ATR_Products_SUM_sorted = L5_ATR_Products_SUM.sort_values(ascending=False)

# Use the sorted data for labels and sizes
labels_products = L5_ATR_Products_SUM_sorted.index
sizes_products = L5_ATR_Products_SUM_sorted.values

patches, texts, autotexts = ax.pie(sizes_products, # Plot on the specific axes 'ax'
        labels=None,
        autopct='%.1f%%',
        startangle=90,
        colors=("cornflowerblue", "orchid", "lightcoral", "mediumseagreen", "sandybrown"),
        wedgeprops={'linewidth': 1, 'edgecolor': 'white'},
        shadow=True,
        pctdistance=0.7
        )
threshold_percentage = 5.0
total_sum = sum(sizes_products)

for i, autotext in enumerate(autotexts):
    percentage = (sizes_products[i] / total_sum) * 100

    if percentage < threshold_percentage:
        # For small slices, move the autotext outside and draw a line
        # Get the angle of the wedge
        angle = (patches[i].theta1 + patches[i].theta2) / 2. # Mid-angle of the wedge
        x = np.cos(np.deg2rad(angle))
        y = np.sin(np.deg2rad(angle))

        # Calculate new position for the text
        # Adjust these radii to control how far out the text is and where the line connects
        connection_point_radius = 1 # Point where the line starts from the pie edge
        text_point_radius = 1.1 # Point where the text is placed

        autotext.set_position((text_point_radius * x, text_point_radius * y))
        autotext.set_ha('center') # Center align the text
        autotext.set_va('center')
        autotext.set_color('black') # Set color for visibility
        autotext.set_fontsize(15) # Smaller font for small labels
        autotext.set_fontweight('bold')

        # Draw a line from the edge of the pie to the text
        ax.plot([connection_point_radius * x, text_point_radius * x],
                [connection_point_radius * y, text_point_radius * y],
                color='gray', linestyle='-', linewidth=0.5, zorder=0) # zorder to keep lines behind text
    else:
        # For larger slices, apply general formatting
        autotext.set_color('black')
        autotext.set_fontsize(20)
        autotext.set_fontweight('bold')


# Set the title for the specific axes

ax.axis('equal')  # Equal aspect ratio ensures the pie chart is circular within its axes
ax.set_axis_off()

legend_labels = []
for i, label in enumerate(labels_products):
    dollar_value = sizes_products[i]
    formatted_dollar_value = f'${dollar_value:,.0f}'
    legend_labels.append(f"{label}: {formatted_dollar_value}")

# Place the legend relative to the figure
# The pie chart axes ends at x = 0.05 (left) + 0.45 (width) = 0.5.
# So, set bbox_to_anchor to start the legend just to the right of this, e.g., 0.55.
fig.legend(patches, legend_labels, title="Product ATR", loc="center left", bbox_to_anchor=(0.55, 0.5),
           fontsize='x-large',  # Adjust font size of legend labels
           title_fontsize='xx-large',
           borderpad=1.5,  # Increase padding around the legend entries
           labelspacing=1.2,  # Increase vertical spacing between entries
           handlelength=2.5,  # Increase the length of the legend handles
           handletextpad=0.8)


# Calculate the total sum of ATR
total_atr_sum_products = L5_ATR_Products_SUM.sum()
formatted_total_atr_products = f'Total Annual ATR For Products in 2026: ${total_atr_sum_products:,.0f}'

# Add the total sum as a footnote relative to the figure
plt.figtext(0.5, 0.02, formatted_total_atr_products, ha="center", fontsize=20, bbox={"facecolor":"white", "alpha":0.5, "pad":5})

filename_products = "Products_Split.png"
plt.savefig(filename_products)
plt.close()



#grabbing the why for that one slide with ohio
#check percentage increase of product YOY
#Check percentage increase of services YOY


StateOfOhio = ATR[ATR["SL6"] == "STR_SLED-STATE OF OHIO KEY-L6"]

StateOfOhio = StateOfOhio[["Selected Dimension 1", "Selected Dimension 2", "Fiscal Year", "SL6","Annual ATR"]]

StateOfOhio = StateOfOhio[StateOfOhio["Selected Dimension 1"] == "Services"]

grouped_result_by_year = StateOfOhio.groupby(["Fiscal Year", "Selected Dimension 2"])["Annual ATR"].sum()

rotated_table = grouped_result_by_year.unstack(level= "Fiscal Year")

print(rotated_table)


################# Looking into catalyst

ATR_catalyst = ATR[["Selected Dimension 1", "Selected Dimension 2","Selected Dimension 3", "Fiscal Year","Annual ATR"]]

ATR_catalyst = ATR_catalyst[ATR_catalyst["Selected Dimension 1"] == "Product"]
ATR_catalyst = ATR_catalyst[ATR_catalyst["Selected Dimension 2"] == "Catalyst"]


yearly_catalyst = ATR_catalyst.groupby(["Fiscal Year","Selected Dimension 3"])["Annual ATR"].sum()
yearly_catalyst = yearly_catalyst.unstack(level= "Fiscal Year")

print(yearly_catalyst)