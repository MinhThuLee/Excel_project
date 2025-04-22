# Excel_project

## Project Overview
This project aims to analyze customer data to identify factors influencing bicycle purchase behavior. The raw dataset consists of 1,000 customers with various demographic and behavioral attributes. The results are presented through a visual dashboard in Excel to support business and marketing decision-making.

## File Structure
- **Excel Project Dataset.xlsx**: The raw data and the place where data processing is done.
<img width="1055" alt="Image" src="https://github.com/user-attachments/assets/90c01f9c-a61f-45ae-9bc7-550fe029062a" />: The final visual representation of the data analysis results on the dashboard.

## Data Cleaning in Excel
To ensure accuracy in the analysis, the following data cleaning steps were performed:

1. **Remove Duplicates**
   - Used the `Remove Duplicates` feature in the *Data* tab to remove duplicate records based on columns like: ID, Name, or demographic information.
   
2. **Standardize Data**
   - Ensured consistent data types (Text/Number/Date).
   - Harmonized the case of values such as “Male”/“male” or “Yes”/“YES” using functions like `UPPER()` and `PROPER()`.
   
3. **Remove Unnecessary Columns**
   - Removed columns that were not necessary for the analysis (e.g., personal names, detailed addresses, customer IDs, etc.).
   
4. **Handle Nulls**
   - Applied filters to identify blank cells.
   - For important fields such as “Age”, “Income”, and “Commute Distance”, I checked and removed rows with missing data when there were only a few. If many values were missing, I used `AVERAGE()` to fill in missing data with the mean value.

## Pivot Tables
Pivot Tables were used to analyze the data along different dimensions:
- **Gender vs. Bike Purchase**
- **Age vs. Bike Purchase**
- **Average Income by Group**
- **Commute Distance vs. Purchase Behavior**
- **Marital Status, Education Level, and Region**
   
These Pivot Tables served as the basis for creating charts in the dashboard.

## Creating the Dashboard in Excel
The steps to create the dashboard are as follows:
1. **Create Charts from Pivot Tables**
   - Used bar, line, and combo charts from the *Insert* tab.
   - Adjusted labels and colors for better clarity.
   
2. **Add Interactive Filters (Slicer)**
   - Applied slicers to variables such as Gender, Region, Education Level, and Marital Status.
   - Slicers allow for quick filtering across all charts simultaneously.

3. **Design the Dashboard Layout**
   - Placed charts logically for easy viewing.
   - Added annotations on the right side of the dashboard to summarize insights.
   - Used a neutral background color (light gray) to highlight the charts.

## Key Insights from the Dashboard
- ~50% of customers have purchased a bike, indicating a significant market potential.
- Men are more likely to purchase bikes than women, making them a prime target for the sports segment.
- The *Middle-Aged* group has the highest purchase rate — marketing strategies should focus on health and quality of life.
- Customers with a commute distance of 0-1 miles have the highest purchase rate, suggesting marketing bikes as an effective commuting solution.

## Practical Applications
- **Targeted Customer Segmentation:** Identifying the right potential customer segments.
- **Personalized Marketing Campaigns:** Creating targeted marketing campaigns based on customer demographics and behaviors.
- **Data-Driven Decision Making:** Helping businesses make informed decisions based on real customer data.
