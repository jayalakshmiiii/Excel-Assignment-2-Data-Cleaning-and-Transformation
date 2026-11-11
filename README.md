# Excel Assignment 2 – Data Cleaning and Transformation

## Tools Used

- Microsoft Excel
- Power Query Editor

## Steps Performed

### 1. Data Cleaning

Applied **Trim, Clean, and Capitalize Each Word** to standardize the columns and remove unwanted spaces and characters.

### 2. Corrected Spelling

Used **Replace Values** to correct `Electroni` to `Electronics` in the Category column.

### 3. Removed Duplicates

Selected all columns and used **Remove Rows → Remove Duplicates**. Three duplicate records were removed.

### 4. Handled Missing Prices

Calculated the median of the Price column using the `MEDIAN` function. The overall median was **$130**, which was used to replace the three missing price values.

### 5. Handled Missing Categories Using Merge Queries

Used **Reference** to create a separate Category Lookup query containing Product Name and Category.

Filtered out null categories and removed duplicate product names. Then, used **Merge Queries** with a Left Outer Join, matching Product Name in both queries.

Expanded the merged column to retrieve the matching category. A conditional column was created to retain existing categories or use the looked-up category when the original value was null.

**Power Query Custom Column Expression:**

```powerquery
if [Category_cleaned] <> null
then [Category_cleaned]
else [Category_Lookup]
```

The category for North Face Backpack was marked as `Unknown` because no matching record was available to determine it reliably.

### 6. Split Product ID

Used **Split Column → By Delimiter**, with a hyphen as the delimiter and Right-most delimiter as the split option. Created Manufacturing Date and Country Code columns.

### 7. Formatted Manufacturing Date

Used **2025** as the year, as instructed by the teacher, and formatted the dates as `dd-mm-yyyy`.

### 8. Merged Brand and Product Name

Used **Add Column → Merge Columns** to combine Brand Name and Product Name, with a space as the separator. Named the new column Product Brand.

### 9. Final Formatting

Applied currency formatting to the Price column with the `$` symbol and two decimal places. Renamed the columns appropriately and used **Close & Load** to load the cleaned dataset into Excel.

## Notes

- Overall median used for missing prices: $130.
- Unknown retained for North Face Backpack due to insufficient information.
- Manufacturing year 2025 used as instructed by the teacher.
