# Excess_deaths
Code to demonstrate ONS's new method for estimating the number of expected and excess deaths, as described in the [methodology article](https://www.ons.gov.uk/peoplepopulationandcommunity/healthandsocialcare/causesofdeath/articles/estimatingexcessdeathsintheukmethodologychanges/february2024).

All code was developed using R version 4.1.3.

## 1. Prerequisites
* RStudio installed locally
* The 'openxlsx' package installed in your R library

## 2. Process
1. Create a new local folder (e.g. 'D:\ons_excess_deaths') - this will be your working directory
2. Download 'ons_weekly_ed.R' and 'ons_monthly_ed.R' from this GitHub repo into your working directory
3. Download the dataset [dataset_20240220.xlsx](https://www.ons.gov.uk/peoplepopulationandcommunity/healthandsocialcare/causesofdeath/datasets/estimatingexcessdeathsintheukmethodologychanges) accompanying the methodology article into your working directory
4. Amend the parameters at the top of 'ons_weekly_ed.R' and/or 'ons_monthly_ed.R' (see below)
5. Run 'ons_weekly_ed.R' (for weekly estimates) or 'ons_monthly_ed.R' (for monthly estimates) in RStudio

## 3. Parameters
* `dir`: the path of your working directory (where the input dataset 'dataset_20240220.xlsx' is saved and your outputs will be directed); this should be entered as a character string, is case sensitive, and should be specified using single forward slashes or double back slashes, e.g. "D:/ons_excess_deaths" or "D:\\ons_excess_deaths", not "D:\ons_excess_deaths"
* `ref_period`: the reference week (for 'ons_weekly_ed.R') or month (for 'ons_monthly_ed.R') for which you want to estimate the number of expected and excess deaths; this should be a single numeric value (only one period is allowed per run of the code), entered in the format yyyyww or yyyymm, e.g. 202301 for Week 1 or January 2023

## 4. Outputs
The code will output one CSV file per run, saved in your working directory. The file will be named 'weekly_ed_yyyyww.csv' (from 'ons_weekly_ed.R') or 'monthly_ed_yyyymm' (from 'ons_monthly_ed.R'), where yyyyww or yyyymm corresponds to the `ref_period` paramater you specified before running the code.

The output file contains estimates of the total number of observed, expected and excess deaths in ref_period (as well as 95% confidence limits for expected and excess deaths), and breakdowns of these quantities by:
* Age group and sex (aggregated across all geographies)
* Geography: the four UK countries and the nine English regions (aggregated across all age groups and both sexes)
