library(openxlsx)
library(lubridate)
library(dplyr)
library(ggplot2)
library(scales)
library(tidyr)
library(grid)
library(RColorBrewer)
library(gridExtra)

space <- function(x, ...) {format(x, ..., big.mark = " ", scientific = FALSE, trim = TRUE)}

###Data reading and processing
#Read revenue and profit data for plants, select only several columns
marg <- read.xlsx("./data/Копия нитрат кальция new.xlsx", sheet = 1)
marg <- tbl_df(data.frame(productGroup1 = marg[,1], productGroup2 = marg[,2], client = marg[,5],
                          month = marg[,4], volume = marg[,6], priceRub = marg[,7]))

#data processing
#marg$scenario <- toupper(marg$scenario) #all scenario charachters to upper case
#marg$market <- tolower(marg$market)
marg$month <- dmy("01.01.1900") + days(marg$month) - days(2) #date to POSIXct

#marg <- filter(marg, ico != "ВГО")
#marg <- filter(marg, scenario == "ФАКТ") #Take only fact 
#marg$scenario[marg$scenario == "ФАКТ"] <- "Факт" #back to normal spelling
marg$priceRub[is.na(marg$priceRub)] <- 0 #all NA revenue to zeros
marg$volume[is.na(marg$volume)] <- 0 #all NA revenue to zeros
marg <- filter(marg, volume > 0 & priceRub > 0)

#marg <- mutate(marg, priceUsd = priceRub / usdRate)
#marg <- select(marg, -(priceRub:ico))
meanPrice <- group_by(marg, productGroup2, month) %>% summarise(meanPriceRub = sum(volume * priceRub) / sum (volume))
marg <- left_join(marg, meanPrice, by = c("productGroup2", "month"))

#Summarise
#marg <- group_by(marg, scenario, companygroup, product, month) %>% summarise(marginal.profit = sum(marginal.profit))
#To narrow tidy form
#marg <- gather(marg, indicator, value, -(productGroup1:month))

makefigures <- function(product, productN) {
        png(paste0("figures/",productN,".png"), width=3308, height=2339)
        p <- ggplot(filter(marg, productGroup2 == product), aes(x = month, y = priceRub)) + 
                geom_point(aes(size = volume, colour = client), alpha = .7) +
                geom_line(aes(colour = client, group = client)) +
                geom_line(aes(y = meanPriceRub), size = 2) +
                #geom_text(aes(label = client), data = filter(marg, priceRub < meanPriceRub, productGroup2 == product)) + 
                scale_color_discrete(guide = "none") +
                scale_size(range = c(4,22), name  ="Объем по сделке") +
                scale_y_continuous(labels = space) +
                ggtitle(paste0(product,"\nзаключенные сделки по реализации")) +
                theme_bw(base_size = 32) + 
                theme(legend.position="top", plot.margin = unit(c(1, 1, 1, 2), "lines"), 
                      axis.text.y = element_text(angle = 90), panel.grid.major.x = element_line(size = 1),
                      legend.key.size = unit(2, "cm"), panel.grid.major.y = element_line(size = 2)) +
                xlab(NULL) +
                ylab("Цена FCA за тонну, руб.") +
                scale_x_datetime(labels = date_format("%b.%y"), breaks = date_breaks("month"),
                                 limits = c(dmy("01.09.2013"), dmy("01.03.2015"))) +
                geom_dl(aes(label = client), method = "last.bumpup")
        
        v <- ggplot(filter(marg, productGroup2 == product), aes(x = month, y = volume)) + 
                #geom_line(stat = "identity", fun.y="sum") +
                geom_bar(alpha = .6, position = "stack", stat = "summary", fun.y="sum") +
                scale_fill_brewer(palette = "Dark2") +
                theme_bw(base_size = 32) + 
                theme(legend.position="none", plot.margin = unit(c(1, 1, 1, 2), "lines"), 
                      axis.text.y = element_text(angle = 90), axis.text.y = element_text(size=14),
                      panel.grid.major.x = element_line(size = 1), panel.grid.major.y = element_line(size = 2)) +
                scale_y_continuous(labels = space) +
                xlab(NULL) +
                ylab("Объем реализации, т.") +
                scale_x_datetime(labels = date_format("%b.%y"), breaks = date_breaks("month"),
                                 limits = c(dmy("01.09.2013"), dmy("01.03.2015"))) 
        #direct.label(p)
        grid.arrange(p, v, nrow = 2, heights = c(3, 1))                  
        dev.off()      
}

productList <- c("Нитрат кальция безводный марка «Оптимум плюс»", "ИАС, переведенный в нитрат кальция", "Нитрат кальция (композиция солевая «юнисалт»)", "Нитрат кальция безводный марки «Премиум»", "Нитрат кальция концентрированный")
productListNames <- c("Оптимум плюс", "переведенный ИАС", "юнисалт", "Нитрат кальция Премиум", "концентрированный")
lapply(productList, function(product) {makefigures(product, productListNames[which(productList == product)])})      