sf <- raw_file


# Load required packages
library(nnet)
library(tidyverse)
library(flextable)
library(officer)
library(rio)
library(dplyr)

export(sf,'sf.xlsx')

# Load your dataset (replace with your file path)

a$Age <- as.factor(a$Age)
a$Status <- as.factor(a$Status)
a$`Place of Residence` <- as.factor(a$`Place of Residence`)
a$`Marital Status` <- as.factor(a$`Marital Status`)
a$`Family income` <- as.factor(a$`Family income`)
a$`Living arrangement` <- as.factor(a$`Living arrangement`)
a$`Daily sleep time` <- as.factor(a$`Daily sleep time`)
#This code creates a new column named **“Knowledge _about_Breast cancer”** and fills it with a **factor version** of the values from **“Signs and symptoms of Breast Cancer (Knowledge)”**.
a$`Knowledge _about_Breast cancer` <- as.factor(a$`Signs and symptoms of Breast Cancer (Knowledge)`)

b <- a %>% 
  mutate(w=
           cut(
             as.numeric(`Signs and symptoms of Breast Cancer (Knowledge)`),
             breaks =c(-Inf, 6, 8, Inf),
             labels = c("Low", "Medium", "High"))) %>% 
  select(Age,Status,`Place of Residence`,`Marital Status`,`Family income`,
         `Living arrangement`,`Daily sleep time`,w)


# Run regression model


model <- multinom(w ~ 
                    Age + Status + `Place of Residence` + `Marital Status` +
                    `Family income` + `Living arrangement` + `Daily sleep time`,
                  data = b)

# Get model summary
model_summary <- summary(model)

# Extract coefficients and standard errors
coef_matrix <- coef(model)
se_matrix <- model_summary$standard.errors
# Calculate statistics

results <- map2(
  .x = list(coef_matrix, se_matrix),
  .y = c("coef", "se"),
  ~ as_tibble(.x, rownames = "comparison") %>%
    pivot_longer(-comparison, names_to = "variable", values_to = .y)
) %>%
  reduce(left_join, by = c("comparison", "variable")) %>%
  mutate(
    odds_ratio = exp(coef),
    p_value = 2 * (1 - pnorm(abs(coef/se))),
    ci_low = exp(coef - 1.96 * se),
    ci_high = exp(coef + 1.96 * se)
  ) %>%
  select(-coef, -se)
#`reduce(left_join, by = c("comparison", "variable"))` means **join all data frames in the list together using left_join, matching rows by the columns `comparison` and `variable`.**

# Create formatted table
final_table <- results %>%
  mutate(across(c(odds_ratio,ci_low, ci_high), ~round(., 3)),
         p_value = round(p_value, 3)) %>%
  select(comparison, variable, odds_ratio,p_value, ci_low, ci_high) %>%
  #The code sorts your table by comparison and variable, then inserts a new row with the text "Signs and symptoms of Breast Cancer (Knowledge)" in both columns at the very top.
  arrange(comparison, variable) %>%
  add_row(comparison = "Signs and symptoms of Breast Cancer (Knowledge)", 
          variable = "Signs and symptoms of Breast Cancer (Knowledge)", .before = 1)  


#Here is the meaning **in one simple sentence**:

#`mutate(across(c(odds_ratio, ci_low, ci_high), ~round(., 3)), p_value = round(p_value, 3))` means **round the columns `odds_ratio`, `ci_low`, `ci_high`, and `p_value` to 3 decimal places.**
#We write `p_value` separately because **`across()` only works on existing columns**, but inside your code:


# Header row

# Add model statistics
model_stats <- tibble(
  `Number of obs` = nrow(a),
  `LR chi2` = model_summary$deviance,
  `Prob > chi2` = pchisq(model_summary$deviance, model_summary$edf, lower.tail = FALSE),
  `Pseudo R2` = 1 - (model_summary$deviance / model_summary$null.deviance)
)


# Create flextable
ft <- final_table %>%
  flextable() %>%
  theme_booktabs() %>%
  set_header_labels(
    comparison = "Signs and symptoms of Breast Cancer (Knowledge)",
    variable = "Variable",
    odds_ratio = "Odds ratio",
    se = "Std. error",
    p_value = "P-value",
    ci_low = "95% CI Low",
    ci_high = "95% CI High"
  ) %>%
  merge_h(part = "header") %>%
  align(align = "center", part = "all") %>%
  fontsize(size = 8, part = "all") %>%
  autofit()

# Add model statistics footer
ft <- ft %>%
  add_footer_lines(values = c(
    "",
    paste("Number of obs =", model_stats$`Number of obs`),
    paste("LR chi2(", length(coef(model))-1, ") =", round(model_stats$`LR chi2`, 2)),
    paste("Prob > chi2 =", round(model_stats$`Prob > chi2`, 4)),
    paste("Pseudo R2 =", round(model_stats$`Pseudo R2`, 4))
  ))

# Save as Word document
doc <- read_docx() %>%
  body_add_flextable(ft) %>%
  body_end_section_landscape() 

# Landscape orientation

print(doc, target = "logistic_Signs_and_symptoms _of_Breast_Cancer (Knowledge).docx")

export(final_table,'logistic_regression_results.csv')
cf <- final_table %>% 
  pivot_wider(names_from = variable,
              values_from = odds_ratio)
export(cf,'logistic_regression.xlsx')

# ------------------------------
# 1️⃣ Load required packages
# ------------------------------
# install.packages(c("officer", "flextable"))
library(officer)
library(flextable)

# ------------------------------
# 2️⃣ Use your dataset
# ------------------------------
b  # replace with your actual dataset

# ------------------------------
# 3️⃣ List of categorical variables
# ------------------------------
cat_vars <- c(
  "Age",
  "Status",
  "Place of Residence",
  "Marital Status",
  "Family income",
  "Living arrangement",
  "Daily sleep time",
  "Signs and symptoms of Breast Cancer (Knowledge)"
)

# ------------------------------
# 4️⃣ Initialize results list
# ------------------------------
results <- list()

# ------------------------------
# 5️⃣ Loop through all pairs
# ------------------------------
for (i in 1:(length(cat_vars) - 1)) {
  for (j in (i + 1):length(cat_vars)) {
    var1 <- cat_vars[i]
    var2 <- cat_vars[j]
    
    # Skip if variables not found
    if (!(var1 %in% names(b)) | !(var2 %in% names(b))) next
    
    v1 <- b[[var1]]
    v2 <- b[[var2]]
    
    # Skip if unequal length
    if (length(v1) != length(v2)) next
    
    # Remove missing values
    complete_idx <- complete.cases(v1, v2)
    v1 <- v1[complete_idx]
    v2 <- v2[complete_idx]
    
    # Skip if no data left
    if (length(v1) == 0) next
    
    # Create contingency table
    tbl <- table(v1, v2)
    
    # Only Chi-square test
    chi <- suppressWarnings(chisq.test(tbl))
    
    results[[paste(var1, "vs", var2)]] <- data.frame(
      Variable1 = var1,
      Variable2 = var2,
      Method = "Chi-square Test",
      Chi_Square = round(as.numeric(chi$statistic), 3),
      DF = as.numeric(chi$parameter),
      P_Value = round(chi$p.value, 4)
    )
  }
}

# ------------------------------
# 6️⃣ Combine results into one dataframe
# ------------------------------
results_df <- do.call(rbind, results)
rownames(results_df) <- NULL

# ------------------------------
# 7️⃣ Export ALL results to Word
# ------------------------------
ft_all <- flextable(results_df)
ft_all <- autofit(ft_all)
ft_all <- bold(ft_all, part = "header")
ft_all <- set_caption(ft_all, caption =
                        "Chi_Results_(Signs and symptoms of Breast Cancer (Knowledge))")

doc_all <- read_docx()
doc_all <- body_add_par(doc_all, "All Chi-square Test Results", style = "heading 1")
doc_all <- body_add_flextable(doc_all, value = ft_all)

#wee also do it by using this code
doc_all <- read_docx() %>%
  body_add_par("All Chi-square Test Results", style = "heading 1") %>%
  body_add_flextable(value = ft_all)

print(doc_all, target = 
        "Chi_Results_(Signs and symptoms of Breast Cancer (Knowledge)Signs and symptoms of Breast Cancer (Knowledge)).docx")

# ------------------------------
# 8️⃣ Export only SIGNIFICANT results (p < 0.05)
# ------------------------------
significant_results <- subset(results_df, P_Value < 0.05)

if (nrow(significant_results) > 0) {
  ft_sig <- flextable(significant_results)
  ft_sig <- autofit(ft_sig)
  ft_sig <- bold(ft_sig, part = "header")
  ft_sig <- set_caption(ft_sig, caption = "Significant Chi-squre Test Results (p < 0.05)")
  
  doc_sig <- read_docx()
  doc_sig <- body_add_par(doc_sig, "Significant Chi-square Test Results", style = "heading 1")
  doc_sig <- body_add_flextable(doc_sig, value = ft_sig)
  print(doc_sig, target = "Significant_chi__Results_Signs and symptoms of Breast Cancer (Knowledge).docx")
}

#frequency table preparation
s <- Breast_Cancer
age=s %>% group_by(s$Age) %>% summarise(Total=n())
stat <- s %>% group_by(s$Status) %>% summarise(Total=n())
place <- s %>% group_by(s$`Place of Residence`) %>% summarise(Total=n())
data_list <- list(age,stat,place)
#find the maximum number of rows among these data frames:
max_row <- max(sapply(data_list,nrow))
print(max_row)
padded_list <- lapply(data_list,function(s){
  if(nrow(s)<max_row){
    s[(nrow(s)+1):max_row,]<-
      NA
  }
  return(s)
})
freq_table <- bind_cols(padded_list)

long_data <- pivot_longer(Book1,
                          cols = -Concentration,
                          names_to = 'Sample',
                          values_to = 'Inhibition')

#**geom plot**#
ggplot(long_data, 
       aes(x = Concentration, y = Inhibition, 
           colour =Sample,group = Sample,shape = Inhibition)) + 
  geom_line(size=1) +
  geom_point(size=2)+
  scale_fill_manual(
    values = c('blueviolet','deepskyblue','aquamarine2',
               'coral','coral3'))+
  
  labs(title = "HRBC membrane stabilization assay", x ="Concentration (µg/ml)" , 
       y = "% protection") + 
  theme_classic()+ 
  theme(plot.title = element_text(family = 'arial',size = 14,
                                  face = 'bold',hjust = .5),
        
        
        axis.title.x = element_text(family = 'arial',size = 12,
                                    face = 'bold',hjust = .5),
        axis.title.y = element_text(family = 'arial',size = 12,
                                    face = 'bold',vjust = .5),
        axis.line.x = element_line(colour = 'black',size = .2,
                                   linetype='solid'))

ggsave("long_data.pdf",plot = my_plot,
       width=10, height=6.5)  

#*geom_BAR*#
ggplot(long_data, 
       aes(x = Inhibition, y = factor(Concentration), fill = Sample)) + 
  geom_bar(stat = "identity",width = .5,
           position = position_dodge(width = .7)) +
  scale_fill_manual(
    values = c('blueviolet','deepskyblue','aquamarine2',
               'coral','coral3'))+
  
  labs(title = "Egg Albumin Assay", x = "% of protection", 
       y = "Concentration (µg/ml)") + 
  theme_classic() +
  theme(plot.title = element_text(family = 'arial',size = 14,
                                  face = 'bold',hjust = .5),
        
        
        axis.title.x = element_text(family = 'arial',size = 12,
                                    face = 'bold',hjust = .5),
        axis.title.y = element_text(family = 'arial',size = 12,
                                    face = 'bold',vjust = .5),
        axis.line.x = element_line(colour = 'black',size = .2,
                                   linetype = 'dashed')
  )

table_1 <- Antimicrobial[2:9,3:7]
table_1 %>% mean(table_1[2:6,6],na.rm = TRUE)
crude_extract <- Antimicrobial %>% 
  rowwise() %>% 
  mutate(
    Mean_absorbance=mean(c(Abs1,Abs2,Ab3),na.rm=TRUE),
    SEM=sd(c(Abs1,Abs2,Ab3),na.rm=TRUE)/sqrt(3)
  ) %>% 
  ungroup()
export(crude_extract,'crude_extract.xlsx')
export(N_Hexane_Fraction,'N Hexane Fraction.xlsx')

long_data <- pivot_longer(Book1,
                          cols = -Concentration,
                          names_to = 'Sample',
                          values_to = 'Inhibition')


ggplot(long_data, 
       aes(x = Inhibition, y = factor(Concentration), fill = Sample)) + 
  geom_bar(stat = "identity",width = .5,
           position = position_dodge(width = .7)) +
  scale_fill_manual(
    values = c('blueviolet','deepskyblue','aquamarine2',
               'coral','coral3'))+
  
  labs(title = "DPPH scavenging activity", x = "% inhibition", 
       y = "Concentration (µg/ml)") + 
  theme_classic() +
  theme(plot.title = element_text(family = 'arial',size = 14,
                                  face = 'bold',hjust = .5),
        
        
        axis.title.x = element_text(family = 'arial',size = 12,
                                    face = 'bold',hjust = .5),
        axis.title.y = element_text(family = 'arial',size = 12,
                                    face = 'bold',vjust = .5),
        axis.line.x = element_line(colour = 'black',size = .2,
                                   linetype = 'dashed')
  )

export(long_data,'long_data.xlsx')

ggplot(long_data, 
       aes(x = Concentration, y = Inhibition, 
           colour =Sample,group = Sample)) + 
  geom_line(size=1) +
  geom_point(size=2)+
  scale_fill_manual(
    values = c('blueviolet','deepskyblue','aquamarine2',
               'coral','coral3'))+
  
  labs(title = "DPPH scavenging activity", x ="Concentration (µg/ml)" , 
       y = "% inhibition") + 
  theme_classic()+ 
  theme(plot.title = element_text(family = 'arial',size = 14,
                                  face = 'bold',hjust = .5),
        
        
        axis.title.x = element_text(family = 'arial',size = 12,
                                    face = 'bold',hjust = .5),
        axis.title.y = element_text(family = 'arial',size = 12,
                                    face = 'bold',vjust = .5),
        axis.line.x = element_line(colour = 'black',size = .2,
                                   linetype='solid'))

