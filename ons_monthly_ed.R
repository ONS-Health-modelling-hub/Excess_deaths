################### LOAD PACKAGES ###################

install.packages("openxlsx")
library(openxlsx)


################### SET PARAMETERS ###################

dir = "D:\\ons_excess_deaths"  ## working directory (input dataset needs to be saved here,
                               ## and outputs will also be directed here)

ref_period = 202312            ## reference month for which to calculate ED (yyyymm numeric,
                               ## just one month allowed)


################### DATA PREP - PRE-MODELLING ###################

### read in dataset
deaths <- read.xlsx(paste0(dir, "\\dataset20240220.xlsx"), sheet="Table_2")

### get rid of info rows and assign column names
deaths <- deaths[-(1:4),]
colnames(deaths) <- c("year", "month", "weekdays", "agegrp", "agegrp_coarse",
                      "sex", "geography", "deaths", "population")

### coerse variables to numeric
deaths$year <- as.numeric(deaths$year)
deaths$weekdays <- as.numeric(deaths$weekdays)
deaths$deaths <- as.numeric(deaths$deaths)
deaths$population <- as.numeric(deaths$population)

### set up factor variables
deaths$month <- as.factor(deaths$month)
deaths$month <- factor(deaths$month, levels=levels(deaths$month)[c(5,4,8,1,9,7,6,2,12,11,10,3)])

deaths$agegrp <- as.factor(deaths$agegrp)
deaths$agegrp <- factor(deaths$agegrp, levels=levels(deaths$agegrp)[c(20,1,10,2:9,11:19)])

deaths$agegrp_coarse <- as.factor(deaths$agegrp_coarse)
deaths$agegrp_coarse <- factor(deaths$agegrp_coarse, levels=levels(deaths$agegrp_coarse)[c(7,1:6)])

deaths$sex <- as.factor(deaths$sex)
deaths$sex <- factor(deaths$sex, levels=levels(deaths$sex)[c(1,2)])

deaths$geography <- as.factor(deaths$geography)
deaths$geography <- factor(deaths$geography, levels=levels(deaths$geography)[c(3,8,7,5,6,13,1,12,2,4,9,10,11)])

### create period (year-month) variable
deaths$month_num <- match(deaths$month, month.name)

deaths$period <- ifelse(deaths$month_num %in% 1:9,
                        paste0(deaths$year, "0", deaths$month_num),
                        paste0(deaths$year, deaths$month_num))

deaths$period <- as.numeric(deaths$period)

### derive trend variable
deaths <- deaths[order(deaths$agegrp, deaths$sex, deaths$geography, deaths$period), ]
rownames(deaths) <- NULL

deaths$trend <- ave(deaths$period,
                    deaths$agegrp, deaths$sex, deaths$geography,
                    FUN=seq_along)

### derive flag for fitting England-only model
deaths$eng <- ifelse(deaths$geography %in% c("England and Wales, including non-residents",
                                             "Wales, excluding non-residents",
                                             "Scotland, including non-residents",
                                             "Northern Ireland, including non-residents"), 0, 1)


################### MODEL FITTING AND PREDICTION ###################

### function to fit monthly models
monthly.modelling <- function(df) {

  ### restrict dataset to five-year fitting period
  trend_end <- unique(df$trend[df$period==ref_period]) - 12
  trend_start <- unique(df$trend[df$period==ref_period]) - 12*6 + 1
  df_fitting <- df[df$trend %in% trend_start:trend_end,]
  rownames(df_fitting) <- NULL

  ### remove COVID-affected months from the fitting period
  covid_periods <- c(202004, 202005, 202011, 202012, 202101, 202102)
  df_fitting <- df_fitting[!(df_fitting$period %in% covid_periods),]
  rownames(df_fitting) <- NULL

  ### fit the model - England & Wales (combined), Wales, Scotland or Northern Ireland
  if(sum(df_fitting$eng)==0) {
    df_mod <- glm(deaths ~ agegrp + sex + trend + month + weekdays +
                           agegrp_coarse:sex + agegrp_coarse:trend + agegrp_coarse:month +
                           offset(log(population)),
                  family=quasipoisson,
                  data=df_fitting)
  }

  ### fit the model - England
  if(sum(df_fitting$eng)>0) {
    df_mod <- glm(deaths ~ agegrp + sex + geography + trend + month + weekdays +
                           agegrp_coarse:sex + agegrp_coarse:trend + agegrp_coarse:month +
                           offset(log(population)),
                  family=quasipoisson,
                  data=df_fitting)
  }

  ### restrict dataset to reference month
  df_pred <- df[df$period==ref_period,]
  rownames(df_pred) <- NULL

  ### use the model to predict deaths in the reference month by age group, sex and geography
  df_preds <- as.data.frame(predict(df_mod, newdata=df_pred,
                                    type="response", se.fit=TRUE))[1:2]
  rownames(df_preds) <- NULL
  colnames(df_preds) <- c("d_hat", "se_d_hat")

  ### bind predictions onto the prediction dataset and store in list
  preds_out <- cbind(df_pred, df_preds)

}

### generate and stack outputs for each geography
monthly_ed <- rbind(
  monthly.modelling(deaths[deaths$geography=="England and Wales, including non-residents",]),
  monthly.modelling(deaths[deaths$eng==1,]),
  monthly.modelling(deaths[deaths$geography=="Wales, excluding non-residents",]),
  monthly.modelling(deaths[deaths$geography=="Scotland, including non-residents",]),
  monthly.modelling(deaths[deaths$geography=="Northern Ireland, including non-residents",])
)

### calculate expected and excess deaths plus their variances
monthly_ed$expected_deaths <- monthly_ed$d_hat
monthly_ed$expected_deaths_var <- monthly_ed$se_d_hat^2
monthly_ed$excess_deaths <- monthly_ed$deaths - monthly_ed$expected_deaths
monthly_ed$excess_deaths_var <- monthly_ed$expected_deaths_var


################### AGGREGATION ###################

### aggregate data to publication age groups
monthly_ed$agegrp_publication <- ifelse(
  monthly_ed$agegrp %in% c("1 to 4 years", "5 to 9 years", "10 to 14 years",
                           "15 to 19 years", "20 to 24 years", "25 to 29 years"),
  "1 to 29 years",
  as.character(monthly_ed$agegrp)
)

monthly_ed$agegrp_publication <- as.factor(monthly_ed$agegrp_publication)
monthly_ed$agegrp_publication <- factor(monthly_ed$agegrp_publication,
                                        levels=levels(monthly_ed$agegrp_publication)[c(15,1:14)])

monthly_ed <- aggregate(x=list(deaths=monthly_ed$deaths,
                               expected_deaths=monthly_ed$expected_deaths,
                               expected_deaths_var=monthly_ed$expected_deaths_var,
                               excess_deaths=monthly_ed$excess_deaths,
                               excess_deaths_var=monthly_ed$excess_deaths_var),
                        by=list(year=monthly_ed$year, month=monthly_ed$month,
                                agegrp_publication=monthly_ed$agegrp_publication,
                                sex=monthly_ed$sex, geography=monthly_ed$geography),
                        FUN=sum)

### restrict dataset to geographies that sum to UK and England totals
monthly_ed_uk <- monthly_ed[grep("including non-residents", monthly_ed$geography),]
monthly_ed_eng <- monthly_ed[!(monthly_ed$geography %in% c("England and Wales, including non-residents",
                                                           "Scotland, including non-residents",
                                                           "Northern Ireland, including non-residents",
                                                           "Wales, excluding non-residents")),]

### aggregate across all people in the UK
monthly_agg_all <- aggregate(x=list(deaths=monthly_ed_uk$deaths,
                                    expected_deaths=monthly_ed_uk$expected_deaths,
                                    expected_deaths_var=monthly_ed_uk$expected_deaths_var,
                                    excess_deaths=monthly_ed_uk$excess_deaths,
                                    excess_deaths_var=monthly_ed_uk$excess_deaths_var),
                             by=list(year=monthly_ed_uk$year, month=monthly_ed_uk$month),
                             FUN=sum)

### aggregate across all people in England
monthly_agg_eng <- aggregate(x=list(deaths=monthly_ed_eng$deaths,
                                    expected_deaths=monthly_ed_eng$expected_deaths,
                                    expected_deaths_var=monthly_ed_eng$expected_deaths_var,
                                    excess_deaths=monthly_ed_eng$excess_deaths,
                                    excess_deaths_var=monthly_ed_eng$excess_deaths_var),
                             by=list(year=monthly_ed_eng$year, month=monthly_ed_eng$month),
                             FUN=sum)

### aggregate by age group
monthly_agg_age <- aggregate(x=list(deaths=monthly_ed_uk$deaths,
                                    expected_deaths=monthly_ed_uk$expected_deaths,
                                    expected_deaths_var=monthly_ed_uk$expected_deaths_var,
                                    excess_deaths=monthly_ed_uk$excess_deaths,
                                    excess_deaths_var=monthly_ed_uk$excess_deaths_var),
                             by=list(year=monthly_ed_uk$year, month=monthly_ed_uk$month,
                                     agegrp=monthly_ed_uk$agegrp_publication),
                             FUN=sum)

### aggregate by sex
monthly_agg_sex <- aggregate(x=list(deaths=monthly_ed_uk$deaths,
                                    expected_deaths=monthly_ed_uk$expected_deaths,
                                    expected_deaths_var=monthly_ed_uk$expected_deaths_var,
                                    excess_deaths=monthly_ed_uk$excess_deaths,
                                    excess_deaths_var=monthly_ed_uk$excess_deaths_var),
                             by=list(year=monthly_ed_uk$year, month=monthly_ed_uk$month,
                                     sex=monthly_ed_uk$sex),
                             FUN=sum)

### aggregate by age group and sex
monthly_agg_age_sex <- aggregate(x=list(deaths=monthly_ed_uk$deaths,
                                        expected_deaths=monthly_ed_uk$expected_deaths,
                                        expected_deaths_var=monthly_ed_uk$expected_deaths_var,
                                        excess_deaths=monthly_ed_uk$excess_deaths,
                                        excess_deaths_var=monthly_ed_uk$excess_deaths_var),
                                 by=list(year=monthly_ed_uk$year, month=monthly_ed_uk$month,
                                         agegrp=monthly_ed_uk$agegrp_publication,
                                         sex=monthly_ed_uk$sex),
                                 FUN=sum)

### aggregate by geography (not just those that sum to UK totals)
monthly_agg_geog <- aggregate(x=list(deaths=monthly_ed$deaths,
                                     expected_deaths=monthly_ed$expected_deaths,
                                     expected_deaths_var=monthly_ed$expected_deaths_var,
                                     excess_deaths=monthly_ed$excess_deaths,
                                     excess_deaths_var=monthly_ed$excess_deaths_var),
                              by=list(year=monthly_ed$year, month=monthly_ed$month,
                                      geography=monthly_ed$geography),
                              FUN=sum)

### stack aggregate outputs
monthly_agg_all$agegrp <- NA
monthly_agg_all$sex <- NA
monthly_agg_all$geography <- NA

monthly_agg_eng$agegrp <- NA
monthly_agg_eng$sex <- NA
monthly_agg_eng$geography <- "England, excluding non-residents"

monthly_agg_age$sex <- NA
monthly_agg_age$geography <- NA

monthly_agg_sex$agegrp <- NA
monthly_agg_sex$geography <- NA

monthly_agg_age_sex$geography <- NA

monthly_agg_geog$agegrp <- NA
monthly_agg_geog$sex <- NA

monthly_agg <- rbind(monthly_agg_all, monthly_agg_eng, monthly_agg_age,
                     monthly_agg_sex, monthly_agg_age_sex, monthly_agg_geog)

monthly_agg$agegrp[is.na(monthly_agg$agegrp)] <- "All ages"
monthly_agg$sex[is.na(monthly_agg$sex)] <- "Both sexes"
monthly_agg$geography[is.na(monthly_agg$geography)] <- "UK, including non-residents"

### compute 95% CIs around expected and excess deaths
monthly_agg$expected_deaths_lcl <- monthly_agg$expected_deaths - qnorm(0.975)*sqrt(monthly_agg$expected_deaths_var)
monthly_agg$expected_deaths_ucl <- monthly_agg$expected_deaths + qnorm(0.975)*sqrt(monthly_agg$expected_deaths_var)
monthly_agg$excess_deaths_lcl <- monthly_agg$excess_deaths - qnorm(0.975)*sqrt(monthly_agg$excess_deaths_var)
monthly_agg$excess_deaths_ucl <- monthly_agg$excess_deaths + qnorm(0.975)*sqrt(monthly_agg$excess_deaths_var)

### round estimates to nearest integer
monthly_agg$deaths <- round(monthly_agg$deaths)
monthly_agg$expected_deaths <- round(monthly_agg$expected_deaths)
monthly_agg$expected_deaths_lcl <- round(monthly_agg$expected_deaths_lcl)
monthly_agg$expected_deaths_ucl <- round(monthly_agg$expected_deaths_ucl)
monthly_agg$excess_deaths <- round(monthly_agg$excess_deaths)
monthly_agg$excess_deaths_lcl <- round(monthly_agg$excess_deaths_lcl)
monthly_agg$excess_deaths_ucl <- round(monthly_agg$excess_deaths_ucl)

### keep only columns that are needed in the output
monthly_agg <- monthly_agg[,c("geography", "agegrp", "sex", "year", "month", "deaths",
                              "expected_deaths", "expected_deaths_lcl", "expected_deaths_ucl",
                              "excess_deaths", "excess_deaths_lcl", "excess_deaths_ucl")]

### sort dataset by geography, age group, sex and period
monthly_agg$geography <- as.factor(monthly_agg$geography)
monthly_agg$geography <- factor(monthly_agg$geography, levels=levels(monthly_agg$geography)[c(12,3,9,8,4,6,7,15,1,14,2,5,10,11,13)])

monthly_agg$agegrp <- as.factor(monthly_agg$agegrp)
monthly_agg$agegrp <- factor(monthly_agg$agegrp, levels=levels(monthly_agg$agegrp)[c(15:16,1:14)])

monthly_agg$sex <- as.factor(monthly_agg$sex)
monthly_agg$sex <- factor(monthly_agg$sex, levels=levels(monthly_agg$sex)[c(1:3)])

monthly_agg$month <- as.factor(monthly_agg$month)
monthly_agg$month <- factor(monthly_agg$month, levels=levels(monthly_agg$month)[c(1:12)])

monthly_agg <- monthly_agg[order(monthly_agg$geography, monthly_agg$agegrp, monthly_agg$sex,
                                 monthly_agg$year, monthly_agg$month),]
rownames(monthly_agg) <- NULL

### write out dataset to working directory
write.csv(monthly_agg,
          file=paste0(dir, "\\monthly_ed_", ref_period, ".csv"),
          row.names=FALSE)
