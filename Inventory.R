####
# Ban thanh pham dung chung
####

library(tidyverse)
library(tidymodels)
library(caret)
library(factoextra)
library(FactoMineR)
library(scales)
library(readxl)
library(dplyr) 
library(writexl)


wb_source <- "/Users/admin/Documents/R Studio/ABC Inventory Analysis/DATA-SX-BTPDC-4.2019-4.2020.xlsx"

wb_sheets <- readxl::excel_sheets(wb_source) 
print(wb_sheets)

# Load everything into the Global Environment
wb_sheets %>%
  purrr::map(function(sheet){ # iterate through each sheet name
  assign(x = sheet,
         value = readxl::read_xlsx(path = wb_source, sheet = sheet),
         envir = .GlobalEnv)
})
binded_df = bind_rows(`ISSUE 1-4 (2020)`,`ISSUE 4-8(2019)`,`ISSUE 9-10(2019)`)
binded_df
binded_df <- rename(binded_df, Month_Issue = 'Tháng Issue')
binded_df <- rename(binded_df, Item_Code = 'Item Code')
df <- binded_df

### Convert Date de ve chart x axis theo ngay
df$Month_Issue <- as.Date(paste0('01-', df$Month_Issue), format="%d-%m-%Y")
library(zoo) #this is a little more forgiving:
df$Month_Issue <- as.yearmon(df$Month_Issue, "%m-%Y")
df$Month_Issue <- as.yearmon(df$Month_Issue, "%b-%Y")
df$Month_Issue <- as.Date(as.yearmon(df$Month_Issue, "%m-%Y"))
table(df$Month_Issue)


#---A-OVERAL KHONG CHIA THEO THANG-GROUP_BY ITEM_CODE, NOT BY MONTH---- 
df2 <- df %>% group_by(Item_Code) %>% summarise(Amount = sum(SL))
#--- 1- BTP nao dang dung nhieu nhat (TOP 100)----
TOP100 <- df2 %>% top_n(100, wt = Amount)  # highest values
BOTTOM100 <- df2 %>% top_n(-100, wt = Amount)  # highest values

df2 <- df2[order(-df2$Amount), ]
# Get the total sales
total_Amount <- sum(df2$Amount)
# Create an empty vector for sales percentages
percentage <- c()
# Get the sales percentage for each item
for (i in df2$Amount){
  percentage <- append(percentage, i/total_Amount * 100)
}

# Import dplyr package
library(dplyr)
# Add the percentage column to inventory data frame
df2 <- mutate(df2, percentage)
# Get the cumulative sum of the products sales' percentages
cumulative <- cumsum(df2$percentage)
# Add the cumulative column to inventory data frame
df2 <- mutate(df2, cumulative)
# Create an empty vector for items classification
classification <- c()


#--- 2-Classify the items by A, B and C----
for (i in df2$cumulative){
  if (i < 70) {
    classification <- append(classification, "A")
  } else if (i >= 70 & i < 90) {
    classification <- append(classification, "B")
  } else {
    classification <- append(classification, "C")
  }
}

# Add the classification column to inventory data frame
df2 <- mutate(df2, classification)
ItemName <- binded_df
ItemName$Month_Issue <- NULL
ItemName$SL <- NULL
ItemName$UOM <- NULL

ItemName <- ItemName[!duplicated(ItemName[ , c("Item_Code")]),]

df3 <- merge(df2,ItemName, by = "Item_Code")
library(writexl)
write_xlsx(x = df3, path = "df3.xlsx",col_names = TRUE)


# Import ggplot2 package
library(ggplot2)

table(df2$classification)
# 3 - Build inventory classification graph----
ggplot(data = df2, aes(x = classification)) +
  geom_bar(aes(fill = classification)) +
  labs(title = "Inventory ABC Distribution",
       subtitle = "",
       x = "ABC Classification",
       y = "Count") + theme_bw() + theme(panel.border = element_blank(), panel.grid.major = element_blank(),panel.grid.minor = element_blank(), axis.line = element_line(colour = "black")) #stack bar

# Print results
summary <- count(df2, classification)
summary$percentage <- summary$n / sum(summary$n) * 100
summary
write_xlsx(x = summary, path = "summary.xlsx",col_names = TRUE)


#---B-OVERAL CHIA THEO THANG- GROUP_BY ITEM_CODE and BY MONTH---- 
df1 <- df %>% group_by(Month_Issue, Item_Code) %>% summarise(Amount = sum(SL))

# Top & Bottom 
for (i in df1$Month_Issue){
  m_TOP10 <- df1 %>% top_n(2, wt = Amount)
}
for (i in df1$Month_Issue){
  m_BOTTOM10 <- df1 %>% top_n(-2, wt = Amount)
}

# Get the total sales
df1 <- df1[order(-df1$Amount), ]

df1 <- df1 %>% group_by(Month_Issue) %>% mutate(m_total_Amount = sum(Amount))
df1 <- df1 %>% group_by(Month_Issue) %>% mutate(percentage = (Amount/m_total_Amount) * 100)
df1 <- df1 %>% group_by(Month_Issue) %>% mutate(cumulative = cumsum(percentage))


# Create an empty vector for items classification
df1 <- df1 %>% group_by(Month_Issue) %>% mutate(classification = ifelse(cumulative < 70,"A", ifelse(cumulative >= 70 & cumulative < 90,"B","C")))

write_xlsx(x = df1, path = "df1.xlsx",col_names = TRUE)

# 3 - Build inventory classification graph----
df1_plot <- df1 %>% group_by(Month_Issue,classification) %>% summarise(N = n())

ggplot(df1_plot, aes(x = Month_Issue, y = N)) +
  geom_bar(
    aes(color = classification, fill = classification),
    stat = "identity", position = position_stack()
    ) +
  scale_color_manual(values = c("#00AFBB", "#E7B800", "#FC4E07"))+
  scale_fill_manual(values = c("#00AFBB", "#E7B800", "#FC4E07")) + theme_bw() + theme(panel.border = element_blank(), panel.grid.major = element_blank(),panel.grid.minor = element_blank(), axis.line = element_line(colour = "black")) #stack bar


# Print results
summary1 <- count(df1, classification)
summary1 <- summary1 %>% group_by(Month_Issue) %>% mutate(percentage = (n/sum(n) * 100))
summary1
# library(writexl)
# write_xlsx(x = summary1, path = "summary1.xlsx",col_names = TRUE)


# TOP BTP 
library(gridExtra)
library(grid)

p1 <- ggplot(data=df1[which(df1$Item_Code == "G1M1000508.000"),], aes(x=Month_Issue, y=Amount)) +
  geom_line() + 
  xlab("") + theme_bw() + theme(panel.border = element_blank(), panel.grid.major = element_blank(),panel.grid.minor = element_blank(), axis.line = element_line(colour = "black")) + labs(title = "G1M1000508.000")

p2 <- ggplot(data=df1[which(df1$Item_Code == "G1M1000529.000"),], aes(x=Month_Issue, y=Amount)) +
  geom_line() + 
  xlab("") + theme_bw() + theme(panel.border = element_blank(), panel.grid.major = element_blank(),panel.grid.minor = element_blank(), axis.line = element_line(colour = "black")) + labs(title = "G1M1000529.000")

p3 <- ggplot(data=df1[which(df1$Item_Code == "G1M1000473.000"),], aes(x=Month_Issue, y=Amount)) +
  geom_line() + 
  xlab("") + theme_bw() + theme(panel.border = element_blank(), panel.grid.major = element_blank(),panel.grid.minor = element_blank(), axis.line = element_line(colour = "black")) + labs(title = "G1M1000473.000")

p4 <- ggplot(data=df1[which(df1$Item_Code == "G1M1000557.000"),], aes(x=Month_Issue, y=Amount)) +
  geom_line() + 
  xlab("") + theme_bw() + theme(panel.border = element_blank(), panel.grid.major = element_blank(),panel.grid.minor = element_blank(), axis.line = element_line(colour = "black")) + labs(title = "G1M1000557.000")

grid.arrange(p1, p2, p3, p4, ncol=2)

# BOTTOM BTP 
library(gridExtra)
library(grid)
p1 <- ggplot(data=df1[which(df1$Item_Code == "G1A1000966.000"),], aes(x=Month_Issue, y=Amount)) +
  geom_line() + 
  xlab("") + theme_bw() + theme(panel.border = element_blank(), panel.grid.major = element_blank(),panel.grid.minor = element_blank(), axis.line = element_line(colour = "red")) + labs(title = "G1A1000966.000")

p2 <- ggplot(data=df1[which(df1$Item_Code == "G1O1000266.000"),], aes(x=Month_Issue, y=Amount)) +
  geom_line() + 
  xlab("") + theme_bw() + theme(panel.border = element_blank(), panel.grid.major = element_blank(),panel.grid.minor = element_blank(), axis.line = element_line(colour = "red")) + labs(title = "G1O1000266.000")


p3 <- ggplot(data=df1[which(df1$Item_Code == "G1A1000312.000"),], aes(x=Month_Issue, y=Amount)) +
  geom_line() + 
  xlab("") + theme_bw() + theme(panel.border = element_blank(), panel.grid.major = element_blank(),panel.grid.minor = element_blank(), axis.line = element_line(colour = "red")) + labs(title = "G1A1000312.000")


p4 <- ggplot(data=df1[which(df1$Item_Code == "G1A1000490.000"),], aes(x=Month_Issue, y=Amount)) +
  geom_line() + 
  xlab("") + theme_bw() + theme(panel.border = element_blank(), panel.grid.major = element_blank(),panel.grid.minor = element_blank(), axis.line = element_line(colour = "red")) + labs(title = "G1A1000490.000")

grid.arrange(p1, p2, p3, p4, ncol=2)


