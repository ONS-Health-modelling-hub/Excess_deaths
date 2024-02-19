################### LOAD PACKAGES ###################

install.packages("openxlsx")
library(openxlsx)


################### SET PARAMETERS ###################

dir = "D:\\ons_excess_deaths"  ## working directory (input dataset needs to be saved here,
                               ## and outputs will also be directed here)

ref_period = 202301            ## reference week for which to calculate ED (yyyyww numeric,
                               ## just one week allowed)


################### DATA PREP - PRE-MODELLING ###################

### read in dataset
deaths <- read.xlsx(paste0(dir, "\\dataset_20240220.xlsx"), sheet="Table_1")

### get rid of info rows and assign column names
deaths <- deaths[-(1:4),]
colnames(deaths) <- c("year", "week", "agegrp", "agegrp_coarse",
                      "sex", "geography", "deaths", "population")

### coerse variables to numeric
deaths$year <- as.numeric(deaths$year)
deaths$week <- as.numeric(deaths$week)
deaths$deaths <- as.numeric(deaths$deaths)
deaths$population <- as.numeric(deaths$population)

### re-label Week 53 as Week 52 for the purpose of modelling
deaths$week_model <- ifelse(deaths$week==53, 52, deaths$week)

### set up factor variables
deaths$week_model <- as.factor(deaths$week_model)

deaths$agegrp <- as.factor(deaths$agegrp)
deaths$agegrp <- factor(deaths$agegrp, levels=levels(deaths$agegrp)[c(20,1,10,2:9,11:19)])

deaths$agegrp_coarse <- as.factor(deaths$agegrp_coarse)
deaths$agegrp_coarse <- factor(deaths$agegrp_coarse, levels=levels(deaths$agegrp_coarse)[c(7,1:6)])

deaths$sex <- as.factor(deaths$sex)
deaths$sex <- factor(deaths$sex, levels=levels(deaths$sex)[c(1,2)])

deaths$geography <- as.factor(deaths$geography)
deaths$geography <- factor(deaths$geography, levels=levels(deaths$geography)[c(3,8,7,5,6,13,1,12,2,4,9,10,11)])

### create period (year-week) variable
deaths$period <- ifelse(deaths$week %in% 1:9,
                        paste0(deaths$year, "0", deaths$week),
                        paste0(deaths$year, deaths$week))

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

### function to fit weekly models
weekly.modelling <- function(df) {

  ### restrict dataset to five-year fitting period
  ref_year <- as.numeric(substr(ref_period, 1, 4))
  ref_week <- as.numeric(substr(ref_period, 5, 6))

  trend_end <- max(df$trend[df$year==floor(ref_year-(51/52)) &
                              df$week==ifelse(ref_week==53, 52, ref_week)])

  trend_start <- min(df$trend[df$year==floor(ref_year-5-(51/52)) &
                                df$week==ifelse(ref_week %in% 52:53, 1, ref_week+1)])

  df_fitting <- df[df$trend %in% trend_start:trend_end,]
  rownames(df_fitting) <- NULL

  ### remove COVID-affected weeks from the fitting period
  covid_periods <- c(202014:202022, 202045:202053, 202101:202108)
  df_fitting <- df_fitting[!(df_fitting$period %in% covid_periods),]
  rownames(df_fitting) <- NULL

  ### fit the model - England & Wales (combined), Wales, Scotland or Northern Ireland
  if(sum(df_fitting$eng)==0) {
    df_mod <- glm(deaths ~ agegrp + sex + trend + week_model +
                           agegrp_coarse:sex + agegrp_coarse:trend + agegrp_coarse:week_model +
                           offset(log(population)),
                  family=quasipoisson,
                  data=df_fitting)
  }

  ### fit the model - England
  if(sum(df_fitting$eng)>0) {
    df_mod <- glm(deaths ~ agegrp + sex + geography + trend + week_model +
                           agegrp_coarse:sex + agegrp_coarse:trend + agegrp_coarse:week_model +
                           offset(log(population)),
                  family=quasipoisson,
                  data=df_fitting)
  }

  ### restrict dataset to reference week
  df_pred <- df[df$period==ref_period,]
  rownames(df_pred) <- NULL

  ### use the model to predict deaths in the reference week by age group, sex and geography
  df_preds <- as.data.frame(predict(df_mod, newdata=df_pred,
                                    type="response", se.fit=TRUE))[1:2]
  rownames(df_preds) <- NULL
  colnames(df_preds) <- c("d_hat", "se_d_hat")

  ### bind predictions onto the prediction dataset and store in list
  preds_out <- cbind(df_pred, df_preds)

}

### generate and stack outputs for each geography
weekly_ed <- rbind(
  weekly.modelling(deaths[deaths$geography=="England and Wales, including non-residents",]),
  weekly.modelling(deaths[deaths$eng==1,]),
  weekly.modelling(deaths[deaths$geography=="Wales, excluding non-residents",]),
  weekly.modelling(deaths[deaths$geography=="Scotland, including non-residents",]),
  weekly.modelling(deaths[deaths$geography=="Northern Ireland, including non-residents",])
)

### calculate expected and excess deaths plus their variances
weekly_ed$expected_deaths <- weekly_ed$d_hat
weekly_ed$expected_deaths_var <- weekly_ed$se_d_hat^2
weekly_ed$excess_deaths <- weekly_ed$deaths - weekly_ed$expected_deaths
weekly_ed$excess_deaths_var <- weekly_ed$expected_deaths_var


################### AGGREGATION ###################

### aggregate data to publication age groups
weekly_ed$agegrp_publication <- ifelse(
  weekly_ed$agegrp %in% c("1 to 4 years", "5 to 9 years", "10 to 14 years",
                          "15 to 19 years", "20 to 24 years", "25 to 29 years"),
  "1 to 29 years",
  as.character(weekly_ed$agegrp)
)

weekly_ed$agegrp_publication <- as.factor(weekly_ed$agegrp_publication)
weekly_ed$agegrp_publication <- factor(weekly_ed$agegrp_publication,
                                       levels=levels(weekly_ed$agegrp_publication)[c(15,1:14)])

weekly_ed <- aggregate(x=list(deaths=weekly_ed$deaths,
                              expected_deaths=weekly_ed$expected_deaths,
                              expected_deaths_var=weekly_ed$expected_deaths_var,
                              excess_deaths=weekly_ed$excess_deaths,
                              excess_deaths_var=weekly_ed$excess_deaths_var),
                       by=list(year=weekly_ed$year, week=weekly_ed$week,
                               agegrp_publication=weekly_ed$agegrp_publication,
                               sex=weekly_ed$sex, geography=weekly_ed$geography),
                       FUN=sum)

### restrict dataset to geographies that sum to UK and England totals
weekly_ed_uk <- weekly_ed[grep("including non-residents", weekly_ed$geography),]
weekly_ed_eng <- weekly_ed[!(weekly_ed$geography %in% c("England and Wales, including non-residents",
                                                        "Scotland, including non-residents",
                                                        "Northern Ireland, including non-residents",
                                                        "Wales, excluding non-residents")),]

### aggregate across all people in the UK
weekly_agg_all <- aggregate(x=list(deaths=weekly_ed_uk$deaths,
                                   expected_deaths=weekly_ed_uk$expected_deaths,
                                   expected_deaths_var=weekly_ed_uk$expected_deaths_var,
                                   excess_deaths=weekly_ed_uk$excess_deaths,
                                   excess_deaths_var=weekly_ed_uk$excess_deaths_var),
                            by=list(year=weekly_ed_uk$year, week=weekly_ed_uk$week),
                            FUN=sum)

### aggregate across all people in England
weekly_agg_eng <- aggregate(x=list(deaths=weekly_ed_eng$deaths,
                                   expected_deaths=weekly_ed_eng$expected_deaths,
                                   expected_deaths_var=weekly_ed_eng$expected_deaths_var,
                                   excess_deaths=weekly_ed_eng$excess_deaths,
                                   excess_deaths_var=weekly_ed_eng$excess_deaths_var),
                            by=list(year=weekly_ed_eng$year, week=weekly_ed_eng$week),
                            FUN=sum)

### aggregate by age group
weekly_agg_age <- aggregate(x=list(deaths=weekly_ed_uk$deaths,
                                   expected_deaths=weekly_ed_uk$expected_deaths,
                                   expected_deaths_var=weekly_ed_uk$expected_deaths_var,
                                   excess_deaths=weekly_ed_uk$excess_deaths,
                                   excess_deaths_var=weekly_ed_uk$excess_deaths_var),
                            by=list(year=weekly_ed_uk$year, week=weekly_ed_uk$week,
                                     agegrp=weekly_ed_uk$agegrp_publication),
                            FUN=sum)

### aggregate by sex
weekly_agg_sex <- aggregate(x=list(deaths=weekly_ed_uk$deaths,
                                   expected_deaths=weekly_ed_uk$expected_deaths,
                                   expected_deaths_var=weekly_ed_uk$expected_deaths_var,
                                   excess_deaths=weekly_ed_uk$excess_deaths,
                                   excess_deaths_var=weekly_ed_uk$excess_deaths_var),
                            by=list(year=weekly_ed_uk$year, week=weekly_ed_uk$week,
                                    sex=weekly_ed_uk$sex),
                            FUN=sum)

### aggregate by age group and sex
weekly_agg_age_sex <- aggregate(x=list(deaths=weekly_ed_uk$deaths,
                                       expected_deaths=weekly_ed_uk$expected_deaths,
                                       expected_deaths_var=weekly_ed_uk$expected_deaths_var,
                                       excess_deaths=weekly_ed_uk$excess_deaths,
                                       excess_deaths_var=weekly_ed_uk$excess_deaths_var),
                                by=list(year=weekly_ed_uk$year, week=weekly_ed_uk$week,
                                        agegrp=weekly_ed_uk$agegrp_publication,
                                        sex=weekly_ed_uk$sex),
                                FUN=sum)

### aggregate by geography (not just those that sum to UK totals)
weekly_agg_geog <- aggregate(x=list(deaths=weekly_ed$deaths,
                                    expected_deaths=weekly_ed$expected_deaths,
                                    expected_deaths_var=weekly_ed$expected_deaths_var,
                                    excess_deaths=weekly_ed$excess_deaths,
                                    excess_deaths_var=weekly_ed$excess_deaths_var),
                             by=list(year=weekly_ed$year, week=weekly_ed$week,
                                     geography=weekly_ed$geography),
                             FUN=sum)

### stack aggregate outputs
weekly_agg_all$agegrp <- NA
weekly_agg_all$sex <- NA
weekly_agg_all$geography <- NA

weekly_agg_eng$agegrp <- NA
weekly_agg_eng$sex <- NA
weekly_agg_eng$geography <- "England, excluding non-residents"

weekly_agg_age$sex <- NA
weekly_agg_age$geography <- NA

weekly_agg_sex$agegrp <- NA
weekly_agg_sex$geography <- NA

weekly_agg_age_sex$geography <- NA

weekly_agg_geog$agegrp <- NA
weekly_agg_geog$sex <- NA

weekly_agg <- rbind(weekly_agg_all, weekly_agg_eng, weekly_agg_age,
                    weekly_agg_sex, weekly_agg_age_sex, weekly_agg_geog)

weekly_agg$agegrp[is.na(weekly_agg$agegrp)] <- "All ages"
weekly_agg$sex[is.na(weekly_agg$sex)] <- "Both sexes"
weekly_agg$geography[is.na(weekly_agg$geography)] <- "UK, including non-residents"

### compute 95% CIs around expected and excess deaths
weekly_agg$expected_deaths_lcl <- weekly_agg$expected_deaths - qnorm(0.975)*sqrt(weekly_agg$expected_deaths_var)
weekly_agg$expected_deaths_ucl <- weekly_agg$expected_deaths + qnorm(0.975)*sqrt(weekly_agg$expected_deaths_var)
weekly_agg$excess_deaths_lcl <- weekly_agg$excess_deaths - qnorm(0.975)*sqrt(weekly_agg$excess_deaths_var)
weekly_agg$excess_deaths_ucl <- weekly_agg$excess_deaths + qnorm(0.975)*sqrt(weekly_agg$excess_deaths_var)

### round estimates to nearest integer
weekly_agg$deaths <- round(weekly_agg$deaths)
weekly_agg$expected_deaths <- round(weekly_agg$expected_deaths)
weekly_agg$expected_deaths_lcl <- round(weekly_agg$expected_deaths_lcl)
weekly_agg$expected_deaths_ucl <- round(weekly_agg$expected_deaths_ucl)
weekly_agg$excess_deaths <- round(weekly_agg$excess_deaths)
weekly_agg$excess_deaths_lcl <- round(weekly_agg$excess_deaths_lcl)
weekly_agg$excess_deaths_ucl <- round(weekly_agg$excess_deaths_ucl)

### keep only columns that are needed in the output
weekly_agg <- weekly_agg[,c("geography", "agegrp", "sex", "year", "week", "deaths",
                            "expected_deaths", "expected_deaths_lcl", "expected_deaths_ucl",
                            "excess_deaths", "excess_deaths_lcl", "excess_deaths_ucl")]

### sort dataset by geography, age group, sex and period
weekly_agg$geography <- as.factor(weekly_agg$geography)
weekly_agg$geography <- factor(weekly_agg$geography, levels=levels(weekly_agg$geography)[c(12,3,9,8,4,6,7,15,1,14,2,5,10,11,13)])

weekly_agg$agegrp <- as.factor(weekly_agg$agegrp)
weekly_agg$agegrp <- factor(weekly_agg$agegrp, levels=levels(weekly_agg$agegrp)[c(15:16,1:14)])

weekly_agg$sex <- as.factor(weekly_agg$sex)
weekly_agg$sex <- factor(weekly_agg$sex, levels=levels(weekly_agg$sex)[c(1:3)])

weekly_agg <- weekly_agg[order(weekly_agg$geography, weekly_agg$agegrp, weekly_agg$sex,
                                 weekly_agg$year, weekly_agg$week),]
rownames(weekly_agg) <- NULL

### write out dataset to working directory
write.csv(weekly_agg,
          file=paste0(dir, "\\weekly_ed_", ref_period, ".csv"),
          row.names=FALSE)
