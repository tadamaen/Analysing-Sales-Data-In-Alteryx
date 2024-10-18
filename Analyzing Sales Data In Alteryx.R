# Install The Necessary R Packages 
# Importing Relevant Libraries
library(readxl)
library(tidyverse)
library(lubridate)

# Read Excel For The 4 Files
order_data <- read_excel("Order Details.xlsx", skip = 1)
regions_data <- read_csv("Regions.csv")
customer_data_east <- read_excel("Customer Orders-East.xlsx")
customer_data_west <- read_excel("Customer Orders-West.xlsx")

# Data Preparation and Data Cleaning
customer_data_east <- customer_data_east %>%
                      rename("CustomerName" = "Name") %>%
                      select(`Order ID`, `Order Date`, `CustomerName`, `State`)
customer_data <- bind_rows(customer_data_east, customer_data_west)
customer_data_unique <- customer_data %>%
                        unique()

order_data <- order_data %>%
              mutate(`Sale Price` = as.numeric(`Sale Price`),
                      `Cost per Item` = as.numeric(`Cost per Item`),
                      `COD` = as.numeric(`COD`),
                      `Credit Card` = as.numeric(`Credit Card`),
                      `Debit Card` = as.numeric(`Debit Card`),
                      `EFT` = as.numeric(`EFT`)) %>%
              mutate(`COD` = ifelse(is.na(`COD`), 0, `COD`),
                     `Credit Card` = ifelse(is.na(`Credit Card`), 0, `Credit Card`),
                     `Debit Card` = ifelse(is.na(`Debit Card`), 0, `Debit Card`),
                     `EFT` = ifelse(is.na(`EFT`), 0, `EFT`)) %>%
              mutate("Number Of Transactions" = `COD` + `Credit Card` + `Debit Card` + `EFT`) %>%
              mutate("Profit Per Item" = `Sale Price` - `Cost per Item`) %>%
              mutate("Total Profit" = `Profit Per Item` * `Number Of Transactions`) %>%
              select(`Order ID`, `Category`, `Total Profit`)

# Question 1: What are the top 10 most profitable orders? 

top_ten_profit_orders <- order_data %>%
                         arrange(desc(`Total Profit`)) %>%
                         head(10) %>%
                         left_join(customer_data_unique, "Order ID") %>%
                         left_join(regions_data, "State") %>%
                         select(`CustomerName`, `Region`, `State`, `Order Date`, `Category`, `Total Profit`)

                      
# Question 2: What day of the week is most profitable, and by how much?

most_profitable_days <- order_data %>%
                        group_by(`Order ID`) %>%
                        summarize(`Profit` = sum(`Total Profit`)) %>%
                        left_join(customer_data_unique, "Order ID") %>%
                        select(`Order Date`, `Profit`) %>%
                        mutate(Date = mdy(`Order Date`),  
                               DayOfWeek = wday(Date, label = TRUE, abbr = FALSE)) %>%
                        group_by(`DayOfWeek`) %>%
                        summarize(`Sum Of Profits` = sum(`Profit`)) %>%
                        arrange(desc(`Sum Of Profits`))


# Question 3: How do the profits vary by category and region?

# By Category 

profit_by_category <- order_data %>%
                      select(`Category`, `Total Profit`) %>%
                      mutate(Category = case_when(
                              grepl("Clothing", Category) ~ "Clothing",
                              grepl("Electronics", Category) ~ "Electronics",
                              grepl("Furniture", Category) ~ "Furniture")) %>%
                      group_by(`Category`) %>%
                      summarise(`Profit By Category` = sum(`Total Profit`))
                            
# By Region 

profit_by_region <- order_data %>%
                    left_join(customer_data_unique, "Order ID") %>%
                    left_join(regions_data, "State") %>%
                    select(`Region`, `Total Profit`) %>%
                    group_by(`Region`) %>%
                    summarize(`Profits By Region` = sum(`Total Profit`))




order_categories <- order_data %>%
                    mutate(`Main Category` = case_when(
                           grepl("Clothing", Category) ~ "Clothing",
                           grepl("Electronics", Category) ~ "Electronics",
                           grepl("Furniture", Category) ~ "Furniture")) %>%
                    select(`Main Category`, `Category`, `Total Profit`)

order_data_furniture <- order_categories %>%
                        filter(`Main Category` == "Furniture") %>%
                        group_by(`Category`) %>%
                        summarize(`Sum Of Total Profit` = sum(`Total Profit`)) %>%
                        arrange(desc(`Sum Of Total Profit`)) %>%
                        head(1)

order_data_electronics <- order_categories %>%
                          filter(`Main Category` == "Electronics") %>%
                          group_by(`Category`) %>%
                          summarize(`Sum Of Total Profit` = sum(`Total Profit`)) %>%
                          arrange(desc(`Sum Of Total Profit`)) %>%
                          head(1)

order_data_clothing <- order_categories %>%
                       filter(`Main Category` == "Clothing") %>%
                       group_by(`Category`) %>%
                       summarize(`Sum Of Total Profit` = sum(`Total Profit`)) %>%
                       arrange(desc(`Sum Of Total Profit`)) %>%
                       head(1)
                            


