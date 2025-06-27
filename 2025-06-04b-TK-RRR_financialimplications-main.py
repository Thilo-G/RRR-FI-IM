#Press strg I to chat with copilot
# ctrl + enter: gitub copilot suggestion
# ctrl + alt + I: open chat view
#debugging F5
#Autoformation Shift+Alt+F
# """ """ for multiline comments
#==============================================================================
# 1. Importing the data files & Data Exploration
#==============================================================================
# Importing the libraries
import pandas as pd
import openpyxl
import numpy as np

###To DO
# add portfolio descriptions
# add F&F portfolio analysis
# maybe change code sections merging before data analysis, as when merging I discovered that I need to drop a firm -> backwards now

# Importing the data files
# Firm Funamentals
file1_path = "C:\\Users\\thkraft\\eCommerce-Goethe Dropbox\\Thilo Kraft\\Thilo(privat)\\Privat\\Research\\RRR_FinancialImplication\\Data\\2025-0319a-TK-fundamentals-python.xlsx"
# Quarterly Revenue
file2_path = "C:\\Users\\thkraft\\eCommerce-Goethe Dropbox\\Thilo Kraft\\Thilo(privat)\\Privat\\Research\\RRR_FinancialImplication\\Data\\2025-0319a-TK-quarterlyrevenue-collection.xlsx"

try:
    # Read the fundamentals file
    df_fundamentals = pd.read_excel(file1_path, header=None);  # Load with no header
    print("Fundamentals-Daten erfolgreich eingelesen.")
    print("Shape:", df_fundamentals.shape)


    # Read the revenue file
    df_revenue = pd.read_excel(file2_path, header=None);       # Load with no header
    print("\nRevenue-Daten erfolgreich eingelesen.")
    print("Shape:", df_revenue.shape)


except Exception as e:
    #Exception is the base class for all exceptions (FilenotFoundError, ValueError, etc.)
    #As e saves the "exception"
    print(f"Fehler beim Einlesen der Dateien: {e}")


##################################################################
# 1.1 Are there any missing firms in the data files? Compare both#
##################################################################

# Normalize strings to ensure case-insensitive and whitespace-consistent comparison
def normalize_strings(strings):
    return set(s.upper().strip() for s in strings)

# Normalize firm names from the first row of both files and # Extract all unique strings from the first row of both files
normalized_firm_names_fundamentals = normalize_strings(set(df_fundamentals.iloc[0].dropna().astype(str).str.strip()));
normalized_firm_names_revenue = normalize_strings(set(df_revenue.iloc[0].dropna().astype(str).str.strip()));

# Compare normalized firm names
missing_in_revenue = normalized_firm_names_fundamentals - normalized_firm_names_revenue;
missing_in_fundamentals = normalized_firm_names_revenue - normalized_firm_names_fundamentals;

# Print the number of unique firms in each dataset
print(f"Number of Unique Firms in Fundamentals File: {len(normalized_firm_names_fundamentals)}")
print(f"Number of Unique Firms in Revenue File: {len(normalized_firm_names_revenue)}")

# Print results
print("Firm Names Missing in Revenue File:")
print(missing_in_revenue)

print("\nFirm Names Missing in Fundamentals File:")
print(missing_in_fundamentals)

#==============================================================================
# 2. Data Wrangling
#==============================================================================

###################
# 2.1 Revenue Data#
###################

# Extract the first two rows for processing headers
header_rows_rev = df_revenue.iloc[:2];
data_rows_rev = df_revenue.iloc[2:];  # Remaining rows are the actual data

# Combine the first two rows to create a single-level column header
combined_headers = header_rows_rev.apply(lambda x: x.str.strip() if x.dtype == "object" else x);
column_headers = combined_headers.apply(lambda x: '.'.join(x.dropna()), axis=0);
data_rows_rev.columns = column_headers;  # Assign proper headers to the data rows

# Rename the first column to "Date"
data_rows_rev.rename(columns={data_rows_rev.columns[0]: 'Date'}, inplace=True);

# Reset the index for the data rows
df_revenue = data_rows_rev.reset_index(drop=True);

### replace empty cells with NaN
df_revenue = df_revenue.replace(r'^\s*$', pd.NA, regex=True);

#DROP CART US EQUITY columns
df_revenue = df_revenue.loc[
    :, 
    ~df_revenue.columns.str.contains('CART US EQUITY', case=False)
]

#Add totel and RRR and shares
# Remove duplicate columns
df_revenue = df_revenue.loc[:, ~df_revenue.columns.duplicated()];
# Identify columns related to "New Customer" and "Returning Customer"
new_customer_cols = [col for col in df_revenue.columns if '#New_Customers' in col];
returning_customer_cols = [col for col in df_revenue.columns if '#Returning_Customers' in col];

"""
code is slow not optimal-> try to improve
"""

# Create new columns for Total Revenue, share, RRR, revenue growth, and revenue_new_growth
for new_col, ret_col in zip(new_customer_cols, returning_customer_cols):
    # Generate new column names
    total_revenue_col = new_col.replace('#New_Customers', '#Total_Revenue')
    share_ret_revenue_col = ret_col.replace('#Returning_Customers', '#Share_Ret_Revenue')
    rrr_col = ret_col.replace('#Returning_Customers', '#RRR')
    revenue_growth_col = ret_col.replace('#Returning_Customers', '#Revenue_Growth')
    revenue_new_growth_col = new_col.replace('#New_Customers', '#Revenue_New_Growth')

    # Add Total Revenue
    df_revenue[total_revenue_col] = df_revenue[new_col] + df_revenue[ret_col]

    # Calculate Share of Returning Revenue
    #### have some 0 df_revenue[share_ret_revenue_col] = df_revenue[ret_col] / df_revenue[total_revenue_col].replace(0, np.nan)
    df_revenue[share_ret_revenue_col] = df_revenue[ret_col] / df_revenue[total_revenue_col].replace(0, np.nan)

    # Calculate RRR
    df_revenue[rrr_col] = df_revenue[ret_col] / df_revenue[total_revenue_col].shift(1).replace(0, np.nan)

    # Calculate Revenue Growth
    df_revenue[revenue_growth_col] = (
        (df_revenue[total_revenue_col] - df_revenue[total_revenue_col].shift(1)) / df_revenue[total_revenue_col].shift(1)
    )

    # Calculate Revenue New Growth
    df_revenue[revenue_new_growth_col] = df_revenue[new_col] / df_revenue[total_revenue_col].shift(1)
    # Reorder columns to place new columns next to Returning Customers
    cols = list(df_revenue.columns)
    ret_col_index = cols.index(ret_col)
    # Remove and reinsert new columns next to the Returning Customer column
    
    cols.remove(total_revenue_col)
    cols.remove(share_ret_revenue_col)
    cols.remove(rrr_col)
    cols.remove(revenue_growth_col)
    cols.remove(revenue_new_growth_col)
    cols.insert(ret_col_index + 1, total_revenue_col)
    cols.insert(ret_col_index + 2, share_ret_revenue_col)
    cols.insert(ret_col_index + 3, rrr_col)
    cols.insert(ret_col_index + 4, revenue_growth_col)
    cols.insert(ret_col_index + 5, revenue_new_growth_col)
    df_revenue = df_revenue[cols]



#######################
# 2.2 Fundamental Data#
#######################

"""
I have some missing cells or 0 which I need to replace with 0 like for R&D expenses
some variables are named the same as in 2.1
"""
### Rename Headers
### Rename Headers
### Rename Headers
# Step 1: Process Headers
header_rows_fun = df_fundamentals.iloc[:2];  # First two rows as headers
data_rows_fun = df_fundamentals.iloc[2:];  # Remaining rows are data

# Combine headers: "Firm.Variable"
combined_headers = header_rows_fun.apply(lambda x: x.astype(str).str.strip(), axis=0);
column_headers = combined_headers.apply(lambda x: '.'.join(x.dropna()), axis=0);
data_rows_fun.columns = column_headers;  # Assign proper headers to the data rows
df_fundamentals = data_rows_fun.reset_index(drop=True);  # Reset index with proper columns

# Step 2: Handle Missing Data
df_fundamentals.replace('#N/A N/A', np.nan, inplace=True)
# Remove duplicate columns
df_fundamentals = df_fundamentals.loc[:, ~df_fundamentals.columns.duplicated()];

### Check for missing values
### Check for missing values
### Check for missing values
# Extract variable names (after the dot)
variable_names = [col.split('.')[-1] for col in df_fundamentals.columns];

# Compute missing values per variable
missing_counts = df_fundamentals.isnull().sum();
missing_summary = pd.DataFrame({'Variable': variable_names, 'Missing_Count': missing_counts.values});

# Aggregate by variable name
missing_summary_grouped = missing_summary.groupby('Variable')['Missing_Count'].sum().sort_values(ascending=False);

# Display the top variables with the most missing values
print("Top variables with missing values:")
print(missing_summary_grouped.head(10))  # Show top 10 variables with most missing values

columns_to_replace = [col for col in df_fundamentals.columns if "IS_RD_EXPEND" in col or "IS_SGA_EXPENSE" in col]

# Replace NaN values with 0 in the selected columns
df_fundamentals[columns_to_replace] = df_fundamentals[columns_to_replace].fillna(0)


# Extract firm names (before the dot in "Firm.Variable")
firm_names = [col.split('.')[0] for col in df_fundamentals.columns];

# Create a DataFrame with firm names and their missing values
missing_summary_firms = pd.DataFrame({'Firm': firm_names, 'Missing_Count': missing_counts.values});

# Aggregate missing values per firm
missing_summary_firms_grouped = missing_summary_firms.groupby('Firm')['Missing_Count'].sum().sort_values(ascending=False);

# Display the top firms with the most missing values
print("Top firms with missing values:")
print(missing_summary_firms_grouped.head(10))  # Show top 10 firms with most missing values



#==============================================================================
# 3. First Analysis of seperated data
#==============================================================================

###################
# 3.1 Revenue Data#
###################

import seaborn as sns
import matplotlib.pyplot as plt

"""
make it more efficient, create function
"""


# Define the columns of interest
columns_of_interest = ["#RRR", "#Revenue_Growth", "#Revenue_New_Growth"]

# Ensure the Date column is in datetime format for proper sorting
df_revenue['Date'] = pd.to_datetime(df_revenue['Date'])

# Iterate over each metric to generate boxplots and histograms
for metric in columns_of_interest:
    # Extract relevant columns for the current metric
    metric_cols = [col for col in df_revenue.columns if metric in col]

    # Melt the dataset for easier plotting with seaborn
    df_metric_melted = df_revenue.melt(
        id_vars=['Date'], 
        value_vars=metric_cols, 
        var_name='Firm', 
        value_name=metric
    )

    # Remove the top 4 highest values correctly for display purposes (not statistics)
    df_filtered = df_metric_melted.copy()
    df_filtered = df_filtered[df_filtered[metric] < df_filtered[metric].nlargest(4).min()]  # Exclude top 4 values

    # Compute additional statistics
    q25, q75 = df_metric_melted[metric].quantile([0.25, 0.75])

    # Ensure the x-axis (dates) remain sorted
    df_filtered = df_filtered.sort_values(by="Date")

    # Boxplot
    plt.figure(figsize=(12, 6))
    sns.boxplot(x='Date', y=metric, data=df_filtered, color='skyblue')

    # Customize the chart
    plt.xlabel('Date')
    plt.ylabel(metric.replace("#", ""))  # Remove '#' for cleaner labels
    plt.xticks(rotation=45, ha='right')
    plt.grid(axis='y', linestyle='--', alpha=0.6)
    plt.tight_layout()
    plt.show()

    # Print summary statistics including quantiles
    print(f"\nSummary statistics for {metric.replace('#', '')}:")
    print(df_metric_melted[metric].describe().loc[['mean', 'std', 'min', 'max']])
    print(f"25th Percentile (Q1): {q25:.2f}")
    print(f"75th Percentile (Q3): {q75:.2f}")

    # Histogram
    plt.figure(figsize=(10, 6))
    sns.histplot(df_filtered[metric], bins=30, kde=True, color='skyblue')

    # Customize the plot
    plt.xlabel(metric.replace("#", ""))
    plt.ylabel('Frequency')
    plt.grid(axis='y', linestyle='--', alpha=0.6)
    plt.show()

"""
Summary statistics for RRR (Excluding Top 4):
mean    0.799884
std     0.312226
min     0.001062
max     3.889442
Name: #RRR, dtype: float64
25th Percentile (Q1): 0.61
75th Percentile (Q3): 0.95

Summary statistics for Revenue_Growth (Excluding Top 4):
mean    0.080728
std     0.410411
min    -0.998914
max     5.636012
Name: #Revenue_Growth, dtype: float64
25th Percentile (Q1): -0.08
75th Percentile (Q3): 0.17

Summary statistics for Revenue_New_Growth:
mean     0.287643
std      0.411074
min      0.000024
max     16.130858
Name: #Revenue_New_Growth, dtype: float64
25th Percentile (Q1): 0.10
75th Percentile (Q3): 0.37
"""


### Create RRR bins and calculate average growth rate
### Create RRR bins and calculate average growth rate
### Create RRR bins and calculate average growth rate

# Extract company names by identifying unique prefixes before the dot (".")
company_names = set([col.split('.')[0] for col in df_revenue.columns if '.' in col])

# Define bins (quartiles) for RRR within each period
bin_labels = ["Q1 (Lowest 25%)", "Q2", "Q3", "Q4 (Highest 25%)"]

# List to store period-wise bin averages
bin_avg_growth_list = []

# Sort dataset by Date
df_revenue = df_revenue.sort_values(by="Date")

# Identify the first available period (which has no RRR)
first_period = df_revenue["Date"].min()



# Iterate over each period (Date)
for date, group in df_revenue.groupby("Date"):
    if date == first_period:
        print(f"Skipping first period {date} (No RRR exists).")
        continue  # Skip the first period since RRR cannot be computed

    firm_rrr = {}
    firm_growth = {}

    # Extract RRR and Revenue Growth for each company
    for company in company_names:
        rrr_col = f"{company}.#RRR"
        growth_col = f"{company}.#Revenue_Growth"

        # Ensure both columns exist in the dataset
        if rrr_col in group.columns and growth_col in group.columns:
            rrr_value = group[rrr_col].values[0]  # Extract RRR for the period
            growth_value = group[growth_col].values[0]  # Extract Growth Rate

            if not pd.isna(rrr_value) and not pd.isna(growth_value):  # Exclude NaN values
                firm_rrr[company] = rrr_value
                firm_growth[company] = growth_value

    # Convert to DataFrame
    df_firms = pd.DataFrame({"Company": list(firm_rrr.keys()), "RRR": list(firm_rrr.values()), "Growth": list(firm_growth.values())})

    # Skip if there are too few firms for quartiles
    if len(df_firms) < 4:
        print(f"Skipping Date {date}: Not enough firms for quartile binning.")
        continue

    # Create quartile bins based on RRR
    try:
        df_firms["RRR_Quartile"] = pd.qcut(df_firms["RRR"], q=4, labels=bin_labels)
    except ValueError:  # Catch cases where binning is not possible
        print(f"Skipping Date {date}: Unable to compute quartiles due to identical values.")
        continue

    # Calculate the average growth rate for each bin (Fix: observed=True)
    avg_growth_per_bin = df_firms.groupby("RRR_Quartile", observed=True)["Growth"].mean()

    # Store results for this period
    avg_growth_per_bin["Date"] = date
    bin_avg_growth_list.append(avg_growth_per_bin)

# Convert list to DataFrame
df_bin_avg_growth = pd.DataFrame(bin_avg_growth_list)

# Compute the final average growth rate per bin across all periods
final_avg_growth_per_bin = df_bin_avg_growth.drop(columns="Date", errors="ignore").mean()
print(final_avg_growth_per_bin)

"""
RRR_Quartile
Q1 (Lowest 25%)    -0.090065
Q2                 -0.002914
Q3                  0.070511
Q4 (Highest 25%)    0.463775 / 0.323965 ohne outliner
Triggered by one period where its?
"""

"""
filter outliner out?
"""

#==============================================================================
# 4. Merging the Data
#==============================================================================

# Exclude specific columns from df_revenue
excluded_columns = ['Date'];
df_revenue_to_append = df_revenue.drop(columns=excluded_columns, errors='ignore');

df_fundamentals = df_fundamentals.rename(columns={'nan.Dates': 'Date'});
df_combined = pd.concat([df_fundamentals, df_revenue_to_append], axis=1);

# Convert all column headers to uppercase
df_combined.columns = df_combined.columns.str.upper()
# Remove duplicate columns by keeping the first occurrence
df_combined = df_combined.loc[:, ~df_combined.columns.duplicated()]

# Identify all variables except "Date"
all_vars = [col for col in df_combined.columns if col not in ["DATE"]]

# Melt the dataset to create a long format
df_long = df_combined.melt(
    id_vars=["DATE"],  # Keep the Date column as an identifier
    value_vars=all_vars,  # Include all variables (original + calculated)
    var_name="Firm_Variable",  # Combined Firm and Variable column
    value_name="Value"  # Values for all variables
)

# Split "Firm_Variable" into "Firm" and "Variable"
df_long[['FIRM', 'VARIABLE']] = df_long['Firm_Variable'].str.rsplit('.', n=1, expand=True)

# Pivot the dataset to create separate columns for variables
df_long = df_long.pivot_table(
    index=['FIRM', 'DATE'],  # Group by Firm and Date
    columns='VARIABLE',  # Create columns for variables
    values='Value',  # Values to fill
).reset_index()

# Set the MultiIndex for panel regression
df_long = df_long.set_index(['FIRM', 'DATE'])


# Test the revenue data
print(f"Correlation between merged datasets is: {df_long['#TOTAL_REVENUE'].corr(df_long['SALES_REV_TURN']):.2f}")
#Correlation between merged datasets is: 0.94

# Share of #Total_Revenue of SALES_REV_TURN in % 
df_long['TOTAL_REVENUE_PCT_SALES_REV_TURN'] = (df_long['#TOTAL_REVENUE'] / df_long['SALES_REV_TURN'].replace(0, np.nan)) / 1000000 * 100 # /1000000 to account for unit difference * 100 for percentage
# Create the boxplot
plt.figure(figsize=(6, 4))
plt.boxplot(df_long['TOTAL_REVENUE_PCT_SALES_REV_TURN'], vert=True)
plt.title('Boxplot of Share of observed revenue to reported revenue')
plt.ylabel('Percentage')
plt.grid(True)
plt.show()
print(f"Average share of reported revenue in observed revenue: {df_long['TOTAL_REVENUE_PCT_SALES_REV_TURN'].mean():.2f}%")
# Average share of reported revenue in observed revenue: 4.13 % 

'''
I have shares of outliners above 100%-> cant be
# 1) Create the mask for ratios above 100%
mask = df_long['TOTAL_REVENUE_PCT_SALES_REV_TURN'] > 60

# 2) Extract and view those rows
outliers = df_long.loc[mask].reset_index()
print(f"Found {len(outliers)} rows with ratio > 100%:")
print(outliers[['FIRM', 'DATE', 'TOTAL_REVENUE_PCT_SALES_REV_TURN']])

VARIABLE            FIRM       DATE TOTAL_REVENUE_PCT_SALES_REV_TURN
0         CART US EQUITY 2019-03-29                       109.350886
1         CART US EQUITY 2019-06-28                       104.114378
manually check in the data

Data points in ALTD: 44833863.1, 47892613.98
Data points in Fund: 41 46

if i lower the threshold to 60 it seems to only concern the CART US EQUITY
-> exclude CART US Equity
-> exclude it in the alt D data already in case it screws the revenue portfolio analysis

'''


### Add needed columns to the dataset
### Add needed columns to the dataset
### Add needed columns to the dataset

# 1) Estimate new and retained revenue by scaling with the total‐revenue ratio
df_long['NEW_REV_EST']   = df_long['#NEW_CUSTOMERS']      / (df_long['TOTAL_REVENUE_PCT_SALES_REV_TURN'] / 100) / 1000000
df_long['RETAINED_REV_EST']   = df_long['#RETURNING_CUSTOMERS'] / (df_long['TOTAL_REVENUE_PCT_SALES_REV_TURN'] / 100) / 1000000 #to account for uni displayment F Data is displayed in millions

# 2) Sum to get an estimated total revenue
df_long['SALES_REV_TURN_EST'] = df_long['RETAINED_REV_EST'] + df_long['NEW_REV_EST']

#Corr is one by contruction

'''
#Calculate Retained & New Revenue Estimate
df_long['RETAINED_REV_EST'] = df_long['#RRR'] * df_long['SALES_REV_TURN'].shift(1) # in million


df_long['NEW_REV_EST'] = df_long['SALES_REV_TURN'] - df_long['RETAINED_REV_EST'] # in million
# Check if any NEW_REV_EST values are negative
if (df_long['NEW_REV_EST'] < 0).any():
    print("There are rows with NEW_REV_EST < 0.")
else:
    print("No rows with NEW_REV_EST < 0.")


what do you do when est new rev < 0 ? ->nothing? estimation error
# 1) Create a boolean mask for negative new revenue estimates
mask_neg_new_rev = df_long['NEW_REV_EST'] < 0

# 2) View all rows where NEW_REV_EST is negative
negative_new_rev = df_long.loc[mask_neg_new_rev]

# 3) Optionally, display only the key columns (Firm, Date, NEW_REV_EST)
print(negative_new_rev.reset_index()[['FIRM', 'DATE', 'NEW_REV_EST']])
400 rows


df_long['SALES_REV_TURN_EST'] = df_long['RETAINED_REV_EST'] + df_long['NEW_REV_EST'] # in million

print(f"Correlation between SALES REV TURN and the Estimated through RRR and lag is: {df_long['SALES_REV_TURN_EST'].corr(df_long['SALES_REV_TURN']):.2f}")
# Correlation between SALES REV TURN and the Estimated through RRR and lag is: 1.0000 by definition though

print(f"Correlation between #New Customer and the Estimated New Revenue is: {df_long['#NEW_CUSTOMERS'].corr(df_long['NEW_REV_EST']):.2f}")
# Correlation between SALES REV TURN and the Estimated through RRR and lag is: 0.3

BiG Problem

print(f"Correlation between #Returning Customer and the Estimated Returning Revenue is: {df_long['#RETURNING_CUSTOMERS'].corr(df_long['RETAINED_REV_EST']):.2f}")
# Correlation between SALES REV TURN and the Estimated through RRR and lag is: 0.96

'''

#Calculate Rest

df_long['RET_EST_X_SHARE_RET'] = df_long['RETAINED_REV_EST'] * df_long['#SHARE_RET_REVENUE']

# Multiply by 100
df_long['RRR_pct'] = df_long['#RRR'] * 100
# Create a one-period lagged version of RRR_pct for each firm
df_long['RRR_pct_lag1'] = df_long.groupby(level='FIRM')['RRR_pct'].shift(1)

df_long['REVENUE_NEW_GROWTH_PCT'] = df_long['#REVENUE_NEW_GROWTH'] * 100
# Compute IS_SGA_EXPENSE_PCT only when SALES_REV_TURN is valid; otherwise, assign NaN
df_long['IS_SGA_EXPENSE_PCT'] = (df_long['IS_SGA_EXPENSE'] / df_long['SALES_REV_TURN'].replace(0, np.nan)) * 100

#Calculate Operating Income Profit Margin
df_long['PM_OPER_PCT'] = (df_long['IS_OPER_INC'] / df_long['SALES_REV_TURN'].replace(0, np.nan)) * 100

# Calculate Revenue Growth as percentage change of SALES_REV_TURN from previous period for each firm
df_long['SALES_REV_TURN_LAG1'] = (
    df_long
      .groupby(level='FIRM')['SALES_REV_TURN']
      .shift(1)
)
df_long['REV_GROWTH_PCT'] = (np.log(df_long['SALES_REV_TURN'].replace(0, np.nan))- np.log(df_long['SALES_REV_TURN_LAG1'].replace(0, np.nan)))*100 # sales growth as percentage change of sales revenue from previous period for each firm
df_long.drop(columns='SALES_REV_TURN_LAG1', inplace=True)


#Calculate stock returns
#Calculate stock returns
#Calculate stock returns
# 1) Make sure PX_LAST is numeric
df_long['PX_LAST'] = pd.to_numeric(df_long['PX_LAST'], errors='coerce')

# 2) Compute lag‐1 prices per firm
df_long['PX_LAST_LAG1'] = (
    df_long
      .groupby(level='FIRM')['PX_LAST']
      .shift(1)
)

# 3) Raw log-return (decimal)
df_long['RETURN_LOG'] = np.log(df_long['PX_LAST'] / df_long['PX_LAST_LAG1']) # compunded decimal return over period; time series analysis
df_long['RET_ARITH']  = np.exp(df_long['RETURN_LOG']) - 1 # decimal percentage change over period, not additive, actual percentage change
# 4) Drop the helper column
#df_long.drop(columns='PX_LAST_LAG1', inplace=True)


# 2) Calculate risk-free 
### make sure RF is converted into fitting period returns decimal (e.g quarterly, monthly)
### make sure RF is converted into fitting period returns (e.g quarterly, monthly)
### make sure RF is converted into fitting period returns (e.g quarterly, monthly)
rf_ann_pct = df_long.xs('USBMMY3M INDEX', level='FIRM')['PX_LAST']  # annual return in percent
conversion = 4 #for quartlery
rf = ((1 + rf_ann_pct/100)**(1/conversion) - 1) # risk free rate per period return decimal
dates = df_long.index.get_level_values('DATE')
df_long['RF']     = rf.reindex(dates).values


# Calculate market premiums using log-returns
mkt_log   = df_long.xs('SPX INDEX',    level='FIRM')['RETURN_LOG']
mkt_arith = np.exp(mkt_log) - 1
df_long['MKT_RF'] = mkt_arith.reindex(dates).values - df_long['RF']
df_long['MKT_RF'] = pd.to_numeric(df_long['MKT_RF'], errors='coerce')
df_long['MKT_RF_LAG'] = df_long.groupby(level='FIRM')['MKT_RF'].shift(1)

# 3) Lagged firm characteristics: SIZE, BTM and existing RRR_pct_lag1
df_long['HISTORICAL_MARKET_CAP'] = pd.to_numeric(df_long['HISTORICAL_MARKET_CAP'], errors='coerce')
size = np.log(df_long['HISTORICAL_MARKET_CAP'])
df_long['SIZE'] = size
df_long['SIZE_LAG'] = size.groupby(level='FIRM').shift(1)
btm  = (df_long['BS_TOT_ASSET'] - df_long['BS_TOT_LIAB2']) / df_long['HISTORICAL_MARKET_CAP']
df_long['BTM_LAG'] = btm.groupby(level='FIRM').shift(1)
df_long['BTM'] = btm
df_long['RRR_LAG'] = df_long['RRR_pct_lag1'] / 100

# 4) Excess return calculation
df_long['EXCESS_RET'] = df_long['RET_ARITH'] - df_long['RF']


# download the data
# Save the merged dataset to a new Excel file
output_file = "C:\\Users\\thkraft\\eCommerce-Goethe Dropbox\\Thilo Kraft\\Thilo(privat)\\Privat\\Research\\RRR_FinancialImplication\\Data\\df_final-20250315.xlsx"
df_long.to_excel(output_file, index=False)
print(f"Data successfully saved to: {output_file}")
print(f"Number of Firms: {df_long.index.get_level_values('FIRM').nunique()}")

#==============================================================================
# 5. Regression Analysis
#==============================================================================



### function to run regressions and analysis 
### function to run regressions and analysis
### function to run regressions and analysis
import pandas as pd
import matplotlib.pyplot as plt
import statsmodels.api as sm
from linearmodels.panel import PanelOLS
from stargazer.stargazer import Stargazer

def analyze_columns(columns, firm_effect=True, time_effect=True, show_plots=True):
    # Calculate the correlation matrix among the selected columns
    corr_matrix = df_long[columns].corr()
    print("Correlation Matrix:")
    print(corr_matrix)
    
    # Compute summary statistics for each column
    print("\nSummary Statistics of columns:")
    summary_dict = {}
    for col in columns:
        col_data = df_long[col]
        summary_dict[col] = {
            'min': col_data.min(),
            'max': col_data.max(),
            'mean': col_data.mean(),
            'std' : col_data.std(),
            '25% quantile': col_data.quantile(0.25),
            '75% quantile': col_data.quantile(0.75)
        }
    summary_df = pd.DataFrame(summary_dict).T
    print(summary_df)
    
    # Plot histograms (optional)
    if show_plots:
        for col in columns:
            lower = df_long[col].quantile(0.05) #winsorize for better display
            upper = df_long[col].quantile(0.95)
            data_filtered = df_long[col][(df_long[col] >= lower) & (df_long[col] <= upper)]
            plt.figure()
            plt.xlabel(col)
            plt.ylabel('Frequency')
            plt.hist(data_filtered)
            plt.title(f'Histogram of {col}')
            plt.show()
            
    # Run panel regression using PanelOLS
    if len(columns) > 1:
        dep_var = columns[0]
        indep_vars = columns[1:]
        y_reg = df_long[dep_var]
        X_reg = df_long[indep_vars]
        X_reg = sm.add_constant(X_reg)  # add constant
        
        model = PanelOLS(y_reg, X_reg, 
                         entity_effects=firm_effect, 
                         time_effects=time_effect,
                         drop_absorbed=True)
        results = model.fit(cov_type='clustered', cluster_entity=True, cluster_time=True)
        print("\nPanel Regression Results:")
        print(results.summary)
    else:
        print("Not enough columns provided for regression (need at least one independent variable).")




### Regression
### Regression
### Regression

analyze_columns(['IS_SGA_EXPENSE', 'RETAINED_REV_EST', 'NEW_REV_EST'], firm_effect=True, time_effect=True,show_plots=False)

'''
new rev est is not significant check why
'''

analyze_columns(['PM_OPER_PCT', 'RRR_pct', 'RRR_pct_lag1','REVENUE_NEW_GROWTH_PCT', 'IS_SGA_EXPENSE_PCT'], firm_effect=True, time_effect=True,show_plots=False)
analyze_columns(['PM_OPER_PCT', 'RRR_pct', 'REVENUE_NEW_GROWTH_PCT', 'IS_SGA_EXPENSE_PCT'], firm_effect=True, time_effect=True,show_plots=False)

analyze_columns(['IS_OPER_INC', 'RETAINED_REV_EST', 'NEW_REV_EST', 'IS_RD_EXPEND'], firm_effect=True, time_effect=True,show_plots=False)

analyze_columns(['RETAINED_REV_EST','IS_RD_EXPEND'], firm_effect=False, time_effect=False,show_plots=True)
plt.figure(figsize=(7, 5))
plt.scatter(df_long['NEW_REV_EST'], df_long['IS_SGA_EXPENSE'], alpha=0.4)
plt.xlabel('NEW_REV_EST')
plt.ylabel('IS_SGA_EXPENSE')
plt.title('Scatterplot: IS_SGA_EXPENSE vs. NEW_REV_EST')
plt.grid(True, linestyle='--', alpha=0.5)
plt.tight_layout()
plt.show()

plt.figure(figsize=(7, 5))
plt.scatter(df_long['RETAINED_REV_EST'], df_long['IS_SGA_EXPENSE'], alpha=0.4)
plt.xlabel('NEW_REV_EST')
plt.ylabel('IS_SGA_EXPENSE')
plt.title('Scatterplot: IS_SGA_EXPENSE vs. NEW_REV_EST')
plt.grid(True, linestyle='--', alpha=0.5)
plt.tight_layout()
plt.show()


#predict
analyze_columns(['EXCESS_RET', 'MKT_RF_LAG', 'SIZE_LAG', 'BTM_LAG'], firm_effect=False, time_effect=False, show_plots=False)
analyze_columns(['EXCESS_RET', 'MKT_RF_LAG', 'SIZE_LAG', 'BTM_LAG','RRR_LAG'], firm_effect=False, time_effect=False, show_plots=False)
'''
to predict or to explain?
'''
#explain
analyze_columns(['EXCESS_RET', 'MKT_RF', 'SIZE', 'BTM','#RRR'], firm_effect=False, time_effect=False, show_plots=True)
analyze_columns(['EXCESS_RET', 'MKT_RF', 'SIZE', 'BTM','#RRR','RRR_LAG'], firm_effect=False, time_effect=False, show_plots=False)

'''
why do I get those results?
- RRR is not significant, but the lagged RRR is significant
'''


#==============================================================================
# 6. Portfolio Analysis
#==============================================================================


###creating portfolios
###creating portfolios
###creating portfolios
def quartile_cumulative_returns(
    df_long: pd.DataFrame,
    weight_by_mcap: bool = True
):
    """
    1) Assign each firm to one of four RRR quartiles per quarter (using RRR_LAG).
    2) Compute mean log return (RETURN_LOG) for each quartile per quarter,
       optionally weighting by lagged market cap.
    3) Add the SPX INDEX log return.
    4) Build a DataFrame of cumulative returns (start = 1 at a synthetic prior date).
    5) Plot Q1–Q4 and SPX cumulative return curves on one graph, with a “$1” start row.
    """

    # 1) Flatten index to assign quartiles
    df = df_long.reset_index()

    # 1a) Drop rows missing RRR_LAG or RETURN_LOG (and MCAP if weighting)
    required = ['RRR_LAG', 'RETURN_LOG']
    if weight_by_mcap:
        required.append('HISTORICAL_MARKET_CAP')
    df = df.dropna(subset=required)

    # 1b) Convert RRR_LAG to numeric
    df['RRR_LAG'] = pd.to_numeric(df['RRR_LAG'], errors='coerce')

    # 1c) Assign quartile (Q1–Q4) per DATE
    df['quartile'] = df.groupby('DATE')['RRR_LAG'].transform(
        lambda x: pd.qcut(x.astype(float), q=4,
                          labels=['Q1','Q2','Q3','Q4'],
                          duplicates='drop')
    )

    # 2) Compute mean (or weighted) log-return per DATE × quartile
    if not weight_by_mcap:
        # simple mean
        mean_quartile = (
            df
            .groupby(['DATE','quartile'])['RETURN_LOG']
            .mean()
            .unstack('quartile')
            .sort_index()
        )
    else:
        # value-weighted by HISTORICAL_MARKET_CAP at same DATE
        df['MCAP'] = pd.to_numeric(df['HISTORICAL_MARKET_CAP'], errors='coerce')
        # sum of MCAP per (DATE, quartile)
        df['MCAP_SUM'] = df.groupby(['DATE','quartile'])['MCAP'].transform('sum')
        # weight_i = MCAP_i / MCAP_SUM
        df['w'] = df['MCAP'] / df['MCAP_SUM']
        # weighted average of RETURN_LOG
        mean_quartile = (
            df
            .assign(weighted_ret=lambda x: x['w'] * x['RETURN_LOG'])
            .groupby(['DATE','quartile'])['weighted_ret']
            .sum()
            .unstack('quartile')
            .sort_index()
        )

    # Ensure columns Q1..Q4 exist (in case any quartile dropped)
    for q in ['Q1','Q2','Q3','Q4']:
        if q not in mean_quartile.columns:
            mean_quartile[q] = np.nan
    mean_quartile = mean_quartile[['Q1','Q2','Q3','Q4']]

    # 3) SPX log-return, aligned to same index
    spx = df_long.xs('SPX INDEX', level='FIRM')['RETURN_LOG']
    spx = spx.reindex(mean_quartile.index)

    # 4) Combine into mean_returns (log-return)
    mean_returns = mean_quartile.copy()
    mean_returns['SPX'] = spx

    # 5) Cumulative log‐sum then exponentiate → cumulative return factor
    log_cum = mean_returns.cumsum()
    cum_returns = np.exp(log_cum) -1

    # 6) Prepend a “start” row of 1.0 at one quarter before the first real date
    first_date = cum_returns.index.min()
    prior_date = first_date - pd.DateOffset(months=3)
    start_row = pd.DataFrame(
        {col: 0.0 for col in cum_returns.columns},
        index=[prior_date]
    )
    cum_with_start = pd.concat([start_row, cum_returns]).sort_index()

    # 7) Plot all series, including the fake start
    plt.figure(figsize=(10, 6))
    for col in ['Q1','Q2','Q3','Q4','SPX']:
        plt.plot(cum_with_start.index, cum_with_start[col], label=col)
    plt.title('Cumulative Returns: Q1–Q4 Portfolios vs SPX Index')
    plt.xlabel('Date')
    plt.ylabel('Cumulative Return')
    plt.legend(title='Series', loc='upper left')
    plt.grid(True, linestyle='--', alpha=0.5)
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()

    return cum_with_start



# Value‐weighted by market cap:
cum_vw = quartile_cumulative_returns(df_long, weight_by_mcap=True)
cum_vw.head()
# equally weighted:
'''
change to equally weighted S&P index
'''
cum_eq = quartile_cumulative_returns(df_long, weight_by_mcap=False)
cum_eq.head()

###analyzing portfolios with performance metrics
###analyzing portfolios with performance metrics
###analyzing portfolios with performance metrics

import statsmodels.api as sm
from scipy.stats import skew, kurtosis

def portfolio_performance_table(port_ret, benchmark_col='SPX'):
    measures = [
        'Geometric Mean Return', 'Downside Deviation', 'Max Drawdown',
        'Sortino Ratio', 'Skewness', 'Kurtosis', 'Alpha', 'Beta',
        'VaR 5%', 'Hit Ratio'
    ]
    portfolios = list(port_ret.columns)
    perf = pd.DataFrame(index=measures, columns=portfolios)
    
    for col in portfolios:
        returns = port_ret[col].dropna()
        benchmark = port_ret[benchmark_col].reindex(returns.index).dropna()
        returns = returns.loc[benchmark.index]  # align

        # Geometric mean
        geo_mean = np.exp(np.log1p(returns).mean()) - 1 #np.log1p(x) is equivalent to np.log(1 + x) offering improved numerical stability for small values

        # Downside deviation (negative returns only)
        downside = returns[returns < 0]
        dd = downside.std(ddof=0)  # population std for downside

        # Max drawdown
        cum = (1 + returns).cumprod()
        cum_max = cum.cummax() #Return a df or series of the same size containing cumulative maximum 
        drawdown = (cum / cum_max - 1).min()

        # Sortino ratio
        sortino = returns.mean() / dd if dd > 0 else np.nan

        # Skewness and kurtosis
        skewness = skew(returns, nan_policy='omit')
        kurt = kurtosis(returns, nan_policy='omit', fisher=False)  # "normal" = 3

        # Alpha and Beta (CAPM regression vs. SPX)
        X = sm.add_constant(benchmark)
        reg = sm.OLS(returns, X).fit()
        alpha = reg.params['const']
        beta = reg.params[benchmark_col]

        # 5% Value at Risk (VaR, historical)
        var5 = np.percentile(returns, 5)

        # Hit ratio (fraction positive returns)
        hit = (returns > 0).mean()

        perf.at['Geometric Mean Return', col] = geo_mean
        perf.at['Downside Deviation', col] = dd
        perf.at['Max Drawdown', col] = drawdown
        perf.at['Sortino Ratio', col] = sortino
        perf.at['Skewness', col] = skewness
        perf.at['Kurtosis', col] = kurt
        perf.at['Alpha', col] = alpha
        perf.at['Beta', col] = beta
        perf.at['VaR 5%', col] = var5
        perf.at['Hit Ratio', col] = hit

    return perf


port_ret = cum_vw.diff().dropna()  # Each column: Q1, Q2, Q3, Q4, SPX
perf_table = portfolio_performance_table(port_ret)
perf_table = perf_table.map(
    lambda x: f"{x:.4f}" if isinstance(x, (float, np.floating)) else x
)
print(perf_table)

'''
change to equally weighted SPX index
'''
port_ret_eq = cum_eq.diff().dropna()  # Each column: Q1, Q2, Q3, Q4, SPX
perf_table_eq = portfolio_performance_table(port_ret_eq)
perf_table_eq = perf_table_eq.map(  
    lambda x: f"{x:.4f}" if isinstance(x, (float, np.floating)) else x
)   
print(perf_table_eq)


###analyzing portfolios with F&F
###analyzing portfolios with F&F
###analyzing portfolios with F&F


# factors are available monthly or annually
# use monthly factors 
# how? 
# two ways: first (easy way): decompount quarterly returns to monthly returns and then merge it with the monthly factors
# second way: get monthly return data and somehow merge it with the RRR quantile data and the factors

# second way
# align portfolio time series returns to the factor dates
# merge your returns with the factor series on the monthly data 
# run the regression: portfolioreturn -RF = alpha + beta1 * MKT_RF + beta2 * SMB + beta3 * HML



# factors are available at https://mba.tuck.dartmouth.edu/pages/faculty/ken.french/data_library.html
# Download the Fama-French factors data (e.g., 5 factors) and save it as a CSV file

# Header is on line 5 (row 4), data starts on line 6 (row 5)
file3_path = "C:\\Users\\thkraft\\eCommerce-Goethe Dropbox\\Thilo Kraft\\Thilo(privat)\\Privat\\Research\\RRR_FinancialImplication\\Data\\2025-06-27-FF_Factors.csv"
#manually delete annual factors after the monthly factors
try:
    # Read the fundamentals file
    ff_factors = pd.read_csv(file3_path, skiprows=3)
    #ff_factors = pd.read_csv(file3_path, skiprows=4,index_col=0)
    print("Fama and French Fundamentals-Daten erfolgreich eingelesen.")
    print("Shape:", ff_factors.shape)

except Exception as e:
    #Exception is the base class for all exceptions (FilenotFoundError, ValueError, etc.)
    #As e saves the "exception"
    print(f"Fehler beim Einlesen der Dateien: {e}")
ff_factors.head()
# Rename 'Unnamed: 0' to 'Date' for clarity
ff_factors = ff_factors.rename(columns={'Unnamed: 0': 'Date'})

# Convert date to datetime (YYYYMM format)
ff_factors['Date'] = pd.to_datetime(ff_factors['Date'].astype(str), format='%Y%m')
# Set as index and sort
ff_factors = ff_factors.set_index('Date').sort_index()
for col in ['Mkt-RF', 'SMB', 'HML', 'RF']:
    ff_factors[col] = ff_factors[col] / 100
ff_factors.head()



#way two: interpolate quarterly returns to monthly returns
quarter_rets = cum_vw.diff().dropna()
# Let's assume the index of quarter_rets is quarterly, e.g., 2020-03-31, 2020-06-30, etc.
monthly_rets = []
monthly_dates = []

for date, row in quarter_rets.iterrows():
    r_q = row.values
    r_m = (1 + r_q) ** (1/3) - 1  # Decompound
    # Add three months for each quarter
    for i in range(3):
        month_date = (pd.to_datetime(date) - pd.offsets.MonthEnd(0)) + pd.DateOffset(months=i-2)
        monthly_rets.append(r_m)
        monthly_dates.append(month_date)

monthly_df = pd.DataFrame(monthly_rets, columns=quarter_rets.columns, index=pd.to_datetime(monthly_dates))
monthly_df = monthly_df.sort_index()
monthly_df.head()

# Set both to month-end frequency for perfect alignment
monthly_df.index = monthly_df.index.to_period('M').to_timestamp('M')
ff_factors.index = ff_factors.index.to_period('M').to_timestamp('M')
combined = monthly_df.join(ff_factors, how='inner')
combined.head()

portfolios = [col for col in monthly_df.columns if col != 'SPX']

for p in portfolios:
    # Excess returns: subtract risk-free rate (ensure both in percent or decimal!)
    y = combined[p] - combined['RF'] / 100   # if RF is in percent and your returns are decimal
    X = combined[['Mkt-RF', 'SMB', 'HML']] / 100  # if these are in percent
    X = sm.add_constant(X)
    model = sm.OLS(y, X).fit()
    print(f'\nFama-French regression for {p}:')
    print(model.summary())

#==============================================================================