library(data.table)
library(stringr)
library(dplyr)
library(readxl)
library(tidyr)
library(officer)
library(flextable)
library(mschart)


setwd('your working directory')
#read in data 
superstore <- read_excel('Superstore.xlsx',sheet ='DATA')
all_regions <- unique(superstore$`Region`)

# functions####################################################################################
#categroy summary
cate_summary <- function(regions = c(all_regions),data){
  data <- data %>% filter(`Region`%in%c(regions))
  data_sum <- data %>% group_by(Category)%>%
         summarise(`TY Sales` = sum(`TY Sales`),
         `Sales Growth` = sum(`TY Sales`)/sum(`LY Sales`)-1,
         `TY Profit` = sum(`TY Profit`),
         `Profit Growth` = sum(`TY Profit`)/sum(`LY Profit`)-1,
         `TY Profit Ratio` = sum(`TY Profit`)/ sum(`TY Sales`),
         `LY Profit Ratio` = sum(`LY Profit`)/sum(`LY Sales`),
         .groups = 'drop')
  data_ttl <- data %>% group_by()%>%
    summarise(Category = 'TTL',`TY Sales` = sum(`TY Sales`),
              `Sales Growth` = sum(`TY Sales`)/sum(`LY Sales`)-1,
              `TY Profit` = sum(`TY Profit`),
              `Profit Growth` = sum(`TY Profit`)/sum(`LY Profit`)-1,
              `TY Profit Ratio` = sum(`TY Profit`)/ sum(`TY Sales`),
              `LY Profit Ratio` = sum(`LY Profit`)/sum(`LY Sales`),
              .groups = 'drop')
  data_ttl <- data_ttl %>% select(names(data_sum))
  data_sum <- data_ttl %>% bind_rows(data_sum)
  data_sum <- data_sum %>% mutate_if(is.numeric,~replace(.,is.na(.),0))%>%
    mutate_if(is.numeric,~replace(.,is.infinite(.),0))%>%
    mutate_if(is.numeric,~replace(.,is.nan(.),0))
  
  return(data_sum)
}


#Product summary 
prod_summary <- function(regions = c(all_regions),data){
  data <- data %>% filter(`Region`%in%c(regions))
  
  data_sum <- data %>% group_by(`Product Name`)%>%
    summarise(`TY Sales` = sum(`TY Sales`),
              `Sales Growth` = sum(`TY Sales`)/sum(`LY Sales`)-1,
              `TY Profit` = sum(`TY Profit`),
              `Profit Growth` = sum(`TY Profit`)/sum(`LY Profit`)-1,
              `TY Profit Ratio` = sum(`TY Profit`)/ sum(`TY Sales`),
              `LY Profit Ratio` = sum(`LY Profit`)/sum(`LY Sales`),
              .groups = 'drop')
  
  data_sum <- data_sum %>% arrange(-`TY Sales`)
  print(dim(data_sum)[1])
  select_row <- min(5,dim(data_sum)[1])
  data_sum <- data_sum[c(1:select_row),]
  data_sel <- data %>% filter(`Product Name`%in% c(unique(data_sum$`Product Name`)))
  data_ttl <- data_sel %>% group_by()%>%
               summarise(`Product Name`='TTL',`TY Sales` = sum(`TY Sales`),
              `Sales Growth` = sum(`TY Sales`)/sum(`LY Sales`)-1,
              `TY Profit` = sum(`TY Profit`),
              `Profit Growth` = sum(`TY Profit`)/sum(`LY Profit`)-1,
              `TY Profit Ratio` = sum(`TY Profit`)/ sum(`TY Sales`),
              `LY Profit Ratio` = sum(`LY Profit`)/sum(`LY Sales`),
              .groups = 'drop')
  data_ttl <- data_ttl %>% select(names(data_sum))
  data_sum <- data_ttl %>% bind_rows(data_sum)
  data_sum <- data_sum %>% mutate_if(is.numeric,~replace(.,is.na(.),0))%>%
                           mutate_if(is.numeric,~replace(.,is.infinite(.),0))%>%
                           mutate_if(is.numeric,~replace(.,is.nan(.),0))
  return(data_sum)
}


#flexttable default table, grid line, 10.5 Calibr, bold, red header, no border


#monthly data 
month_summary <-function(regions = c(all_regions),data){
  data <- data %>% filter(`Region`%in%c(regions))
  data_sum <- data %>% group_by(`Month`)%>%
    summarise(`TY Sales` = sum(`TY Sales`),
              `LY Sales` = sum(`LY Sales`),.groups = 'drop')
 
  return(data_sum)
}

get_color<-function(key_mar_table,col_num){
  
  key_man_table_sel <- as.numeric(unlist(key_mar_table[,c(col_num)]))
  #replace is.finite is.nan is. na
  key_man_table_sel[is.infinite(key_man_table_sel)]<-0
  key_man_table_sel[is.nan(key_man_table_sel)] <- 0
  key_man_table_sel[is.na(key_man_table_sel)] <- 0
  
  min_sc <- min(key_man_table_sel)
  max_sc <- max(key_man_table_sel)
  #max_size <- abs((max_sc-min_sc)/8)
  
  col_palette_l <- c("#D73027", "#F46D43", "#FDAE61", "#FEE08B")
  col_palette_h <- c("#D9EF8B", "#A6D96A", "#66BD63", "#1A9850")
  col_palette <- c(col_palette_l,col_palette_h)
  
  if(max_sc<0){
    break_size <- (max_sc - min_sc)/4
    share_breaks <- seq(min_sc,max_sc,break_size)
    mycut_l <- cut(key_man_table_sel, 
                   breaks = share_breaks , 
                   include.lowest = TRUE, label = FALSE)
    
    pat<-col_palette_l[mycut_l]
  }else if(min_sc>=0 & max_sc !=0){
    break_size <- (max_sc - min_sc)/4
    share_breaks <- seq(min_sc,max_sc,break_size)
    mycut_h <- cut(key_man_table_sel, 
                   breaks = share_breaks , 
                   include.lowest = TRUE, label = FALSE)
    
    pat<-col_palette_h[mycut_h]
    
  }else if(min_sc ==0 & max_sc==0){
    pat<-c("#FFFFFF") 
  }
  else if(min_sc<0 & max_sc>=0 ){
    bs_low  <- abs(min_sc)/4
    bs_high <- max_sc/4
    break_l <- seq(min_sc,0,bs_low)
    break_h <- seq(0,max_sc,bs_high)
    share_breaks <- unique(c(break_l,break_h))
    mycut <- cut(key_man_table_sel, 
                 breaks = share_breaks , 
                 include.lowest = TRUE, label = FALSE)
    pat <- col_palette[mycut]
  }
  
  # if = 0 
  pat[key_man_table_sel==0]<-c("#FFFFFF")
  
  return(pat)
}

all_table_format <- function(prod_table){
  row_num <- dim(prod_table)[1]
  ft <- flextable(prod_table)
  ft <- ft %>% border_remove()%>%surround(
    border = fp_border_default(color = "grey", width = 1))
  ft <- ft %>% font(font= "Calibri",part = 'all')
  ft <- ft %>% bg(bg = "#C00000",part = 'header')
  ft <- ft %>% bold(i = 1,part ='header')%>%
    color(i =1,color = 'white',part = 'header')
  ft <- ft %>% set_formatter(`Sales Growth` = function(x) sprintf( "%.1f%%", x*100),
                             `Profit Growth` = function(x) sprintf( "%.1f%%", x*100),
                             `TY Profit Ratio` = function(x) sprintf( "%.1f%%", x*100),
                             `LY Profit Ratio` = function(x) sprintf( "%.1f%%", x*100))
  ft <- ft %>% colformat_num(j = c(2,4), prefix = "$")
  ft <- ft %>% bg(bg = "#D9D9D9",i = 1, part = 'body')%>%bold(i =1,part ='body')
  ft <- ft %>% fontsize(size = 10.5, part = 'header')
  ft <- ft %>% bg(i = c(2:row_num),j = 3,bg = get_color(prod_table,3)[2:row_num],part = 'body')
  ft <- ft %>% bg(i = c(2:row_num) ,j = 5,bg = get_color(prod_table,5)[2:row_num],part = 'body')
  return(ft)
}

score_card_exp <- function(my_pres,regions,data ){
tt_loc <- ph_location(left =1.59, top =0.31 , width =10.16 , height = 0.56)
  
  if(length(regions)==1){
    my_pres <- add_slide(my_pres)
    
    title_name <- toupper(paste(regions,'SALES SCORECARD', sep = " "))
    
    paragraph <- fpar(ftext(title_name,
                            fp_text( font.size = 32,bold = TRUE,color = "black",
                                     font.family = "Century Gothic (Headings)")),fp_p = fp_par(text.align = "center"))
    my_pres <- ph_with(my_pres, value = paragraph, 
                       location = tt_loc)}
  
  t1_loc <- ph_location(left = 0.45, top = 1.2)
  t2_loc <- ph_location(left = 6.34, top = 1.2)
  c3_loc <- ph_location(left = 0.45, top = 3.15,width = 12.46, height = 3.46)
  
  cate_table <- cate_summary(regions,data)
  prod_table <- prod_summary(regions,data)
  month_table <- month_summary(regions,data)
  
  cate_table_ft <- all_table_format(cate_table)
  cate_table_ft <- cate_table_ft %>% fontsize(size = 14, part = 'body')%>%
    width(j = c(1,2,4),width = c(1.3,0.88,0.84))%>%
    width(j = c(3,5,6,7), width = c(0.7))%>%
    height(height = 0.32 ,part = 'body')%>%
    height(height = 0.54, part = 'header')%>%fontsize(size = 13,part = 'body')
  
  prod_table_ft <- all_table_format(prod_table)
  prod_table_ft <- prod_table_ft %>% width(j = c(1), width = c(2.79))%>%
    width(j = c(2:7), width = c(0.62))%>%fontsize(size = 9, part = 'body')%>%
    height(height = 0.21, part = 'body')%>%
    height(height =0.54 , part = 'header')
  
  month_table <- month_table %>%pivot_longer(!Month, names_to = "Metrics", values_to = "Sales")
  month_table <- month_table %>% mutate(Month = as.factor(Month))
  month_table <- month_table %>% mutate(ksales = paste('$',round(Sales/1000,0),sep = ""))
  
  my_barchart <- ms_barchart(data = month_table , 
                             x = "Month", y = "Sales", group = "Metrics",labels = "ksales")  
  fp_text_settings <- list(
    `LY Sales` = fp_text(font.size = 12),
    `TY Sales` = fp_text(font.size = 12)
  )
  
  my_barchart <- chart_labels_text( my_barchart,
                                    values = fp_text_settings )
  my_barchart <- chart_data_labels(my_barchart, position = "outEnd")
  
  my_barchart<- chart_data_fill(my_barchart,
                                values = c(`TY Sales` = "#D9D9D9", `LY Sales` = "#C00000"))
  
  my_barchart <- chart_data_stroke(my_barchart, values = "transparent" )
  my_theme <- mschart_theme(legend_position = 't',grid_major_line = fp_border(color = "#99999999", style = "none"))
  my_barchart <- set_theme(my_barchart, my_theme)
  
  my_barchart<- chart_ax_y(my_barchart,num_fmt = '$#,##0,')
  
  my_pres <- ph_with(my_pres, value = cate_table_ft, 
                     location = t1_loc)
  
  my_pres <- ph_with(my_pres, value = prod_table_ft, 
                     location = t2_loc)
  
  my_pres <- ph_with(my_pres, value = my_barchart,
                     location = c3_loc)
  return(my_pres)
}



###################################Create Slides defalut table################################## 
my_pres <- read_pptx('PerformanceScoreCard1.pptx')

my_pres <- score_card_exp(my_pres,all_regions,data = superstore)
for(i in (1:length(all_regions))){
  my_pres <-score_card_exp(my_pres = my_pres,all_regions[i],data = superstore)
}

print(my_pres, target = "PerformanceScoreCard2.pptx") 


###################################Demo#########################################################
# create 1
my_pres <- read_pptx() 
my_pres <- add_slide(my_pres, layout = "Title and Content", master = "Office Theme")
print(my_pres,'example1.pptx')

my_pres <- read_pptx('PerformanceScoreCard1.pptx') 
my_pres <- add_slide(my_pres, layout = "Title and Content", master = "Office Theme")
print(my_pres,'example2.pptx')
###################################Demo#########################################################

my_pres <- read_pptx('PerformanceScoreCard1.pptx') 
my_pres <- add_slide(my_pres)

title_name <- toupper(paste('WEST','SALES SCORECARD', sep = " "))
tt_loc <- ph_location(left =1.59, top =0.31 , width =10.16 , height = 0.56)
paragraph <- fpar(ftext(title_name,
                        fp_text( font.size = 32,bold = TRUE,color = "black",
                                 font.family = "Century Gothic (Headings)")),fp_p = fp_par(text.align = "center"))
my_pres <- ph_with(my_pres, value = paragraph, 
                   location = tt_loc)


#my_pres <- add_slide(my_pres)
#my_pres <- ph_with(my_pres, value = paragraph, 
#                   location = ph_location_type('body'))

###################################Demo#########################################################

cate_table <- cate_summary(regions,data)
t1_loc <- ph_location(left = 0.45, top = 1.2)
#cate_table_ft <- flextable(cate_table)

cate_table_ft <- all_table_format(cate_table)
cate_table_ft <- cate_table_ft %>% fontsize(size = 14, part = 'body')%>%
  width(j = c(1,2,4),width = c(1.3,0.88,0.84))%>%
  width(j = c(3,5,6,7), width = c(0.7))%>%
  height(height = 0.32 ,part = 'body')%>%
  height(height = 0.54, part = 'header')%>%fontsize(size = 13,part = 'body')

my_pres <- ph_with(my_pres, value = cate_table_ft, 
                   location = t1_loc)

###Adding chart 
# preparing data 
c3_loc <- ph_location(left = 0.45, top = 3.15,width = 12.46, height = 3.46)
month_table <- month_summary(regions,data)
month_table <- month_table %>%pivot_longer(!Month, names_to = "Metrics", values_to = "Sales")
month_table <- month_table %>% mutate(Month = as.factor(Month))
month_table <- month_table %>% mutate(ksales = paste('$',round(Sales/1000,0),sep = ""))

my_barchart <- ms_barchart(data = month_table , 
                           x = "Month", y = "Sales", group = "Metrics",labels = "ksales")  
#change lable setting 
fp_text_settings <- list(
  `LY Sales` = fp_text(font.size = 12),
  `TY Sales` = fp_text(font.size = 12)
)
my_barchart <- chart_labels_text( my_barchart,
                                  values = fp_text_settings )
my_barchart <- chart_data_labels(my_barchart, position = "outEnd")

#change fill color and borderline 
my_barchart<- chart_data_fill(my_barchart,
                              values = c(`TY Sales` = "#D9D9D9", `LY Sales` = "#C00000"))

my_barchart <- chart_data_stroke(my_barchart, values = "transparent" )
#change lengend position and remove border line 
my_theme <- mschart_theme(legend_position = 't',grid_major_line = fp_border(color = "#99999999", style = "none"))
my_barchart <- set_theme(my_barchart, my_theme)
#change y axis format 
my_barchart<- chart_ax_y(my_barchart,num_fmt = '$#,##0,')

#add chart to table
my_pres <- ph_with(my_pres, value = my_barchart,
                   location = c3_loc)



print(my_pres,'example.pptx')

