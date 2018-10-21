Taotao <- new.env()

Taotao$Packages.InstallandLoad <- function(pkg) {
  new.pkg <- pkg[!(pkg %in% installed.packages()[, "Package"])]
  if (length(new.pkg))
    install.packages(new.pkg, dependencies = TRUE)
  sapply(pkg, library, character.only = TRUE)
}


# 00 配置运行环境 ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
packages <- c("tidyverse", "rgdal", "raster", "gstat", "sp", "tmap","openxlsx")
Taotao$Packages.InstallandLoad(packages)
sapply(packages, require, character.only = TRUE)
rm(packages)
#----

Taotao$Preview.CSV <- function(ds) {
  require(stringi)
  require(excel.link)
  sheetName <- stringi::stri_rand_strings(1, 10)
  xl.sheet.add(sheetName)
  xl.sheet.activate(sheetName)
  xlc$a1 <- ds
}

Taotao$Preview.ggplot <- function(G, width = 17.5, height = 20, dpi = 600) {
    
    Taotao$Packages.InstallandLoad("ggplot2")
    tmpfile = sprintf("%s//%s", "99 output", "1.tiff")
    ggsave(plot = G,filename =  tmpfile,units = "cm",width = width,
           height = height, dpi = dpi, compression = "lzw")
    shell.exec(tmpfile)
  }

Taotao$tibble.output <- function(DS){
    dir <- tempdir(TRUE)
    dir <- sprintf("%s/%s/", dir, stringr::str_replace_all(string = date(),pattern = ":","-"))
    dir.create(dir, showWarnings = FALSE)
    
    dataCols <- names(DS) %>% `[`(sapply(DS, is.list))
    DS %>% dplyr::select(one_of(setdiff(names(DS),dataCols))) %>% 
      write_csv(path = sprintf("%s/info.csv",dir))
    
    for (col in dataCols) {
      dir.create(sprintf("%s/%s/", dir,col), showWarnings = FALSE)
      walk2(
        DS$SID,
        DS[[col]],
        .f = function(name, ds) {
          write.csv(x = ds, file = sprintf("%s/%s/%s.csv", dir, col,name))
        }
      )
    }
    shell.exec(dir)
}

Taotao$DataAnalysis.ma <- function(x, n = 5) {
  stats::filter(x, rep(1 / n, n), sides = 1) %>% as.numeric()
}

Taotao$addStationsINFO <- function(x, othercols = NULL) {
  nms <- c("SID", "EnStationName", "latitude", "longitude")
  if (!is.null(othercols)) {
    nms <- append(nms, othercols)
  }
  info <- read_rds(path = "00Dataset/StationsINFO.rds") %>% 
    mutate(SID=as.character(SID)) %>% dplyr::select(one_of(nms))

  info %>% left_join(x, by = "SID") 
}

Taotao$tiff.preview <- function(...,ncol = 1, nrow = 1, width = 10, height = 7, res = 600) {
  require(tmap)
  file <- tempfile(fileext = ".tiff") 
  tiff(
    filename = file,
    width = width,
    height = height,
    units = "cm",
    compression = "lzw",
    res = res
    # family="Arial Unicode MS"
  )
  tmap_arrange(...,ncol = ncol,nrow = nrow, outer.margins = 0) %>% print
  dev.off()
  shell.exec(file)
}

Taotao$regenerating.boundary <- function(){
  crs.geo <- CRS("+proj=longlat +ellps=WGS84 +datum=WGS84")
  boundary_Province <-
    readOGR(dsn = "00 Map", "省级行政区", encoding = "UTF-8")
  boundary_Province <- spTransform(boundary_Province, crs.geo)
  Province <- tm_shape(boundary_Province) +
    tm_polygons(alpha = 0,
                border.col = 1,
                lwd = 1)
  
  boundary_city <- readOGR(dsn = "00 Map", "中国地州界")
  boundary_city <- spTransform(boundary_city, crs.geo)
  
  city <- tm_shape(boundary_city) + tm_polygons(alpha = 0)
  JLboundaries <- list(Province = Province,
                       city = city,
                       OGR_Province = boundary_Province)
  
  write_rds(x = JLboundaries, path = "00 Map/boundaries.rds")
}

Taotao$output.Excel <- function(ds) {
  require(openxlsx)
  wb <- createWorkbook()
  # 找出嵌套数据结构
  dataCols <- names(ds) %>% `[`(sapply(DS, is.list))
  info <- ds %>% dplyr::select(one_of(setdiff(names(ds), dataCols)))
  
  # 将嵌套数据结构用超链接的形式显示
  info[dataCols] <-   map(dataCols,
                          ~ makeHyperlinkString(sprintf("%s_%s", ds$SID,.x),text = "details"))
  
  addWorksheet(wb, "info")
  
  #写入表头
  writeData(wb, sheet = "info", x = info %>% head(0)) 
  
  #添加内容，自动判断文本还是链接
  walk2(
    info,
    seq_along(info),
    .f = function(.x, .y) {
      if (any(str_detect(.x, "=HYPERLI"))) {
        writeFormula(wb,sheet = "info",startCol = .y,startRow = 2,x = .x)
        
      } else {
        writeData(wb,sheet = "info",startCol = .y,startRow = 2,x = .x) 
      }
    }
  )
  # 开始添加嵌套表格
  
  for (col in dataCols) {
    walk2(
      DS$SID,
      DS[[col]],
      .f = function(name, ds) {
        if (is.tibble(ds) || is.data.frame(ds)) {
          name = sprintf("%s_%s", name,col)
          addWorksheet(wb, name)
          writeData(wb, name,startCol = 1,startRow = 2,x =  ds)
          writeFormula(wb,sheet = name,startCol = 1,startRow = 1,
                       x = makeHyperlinkString("info",text = "Back") )
          
        } else if (is.list(ds)) {
          
          name = sprintf("%s_%s",  name,col)
          addWorksheet(wb, name)
          
          xxx <- tibble(nme = names(ds), data = ds) 
          writeData(wb, sheet = name,startRow =2, x = xxx %>% head(0)) 
          
          xxx2 <- map(xxx$nme,
                      ~ makeHyperlinkString(sprintf("%s_%s", name,.x),text = "details"))
          writeData(wb,sheet = name,startCol = 1,startRow = 3,x = xxx$nme)
          writeFormula(wb,sheet = name,startCol = 2,startRow = 3,x = xxx2 %>% map_chr(~.x))
          writeFormula(wb,sheet = name,startCol = 1,startRow = 1,
                       x = makeHyperlinkString("info",text = "Back"))
          
          
          walk2(xxx$data,sprintf("%s_%s", name,xxx$nme),
                .f = function(subds,subname){
                  addWorksheet(wb, subname)
                  writeFormula(wb,sheet = subname,startCol = 2,startRow = 1,
                               x = makeHyperlinkString(name,text = "Back"))
                  
                  writeData(wb, subname, subds,startRow = 2)
                  writeFormula(wb,sheet = subname,startCol = 1,startRow = 1,
                               x = makeHyperlinkString("info",text = "Back"))
                  
                  
                })
        }
      }
    )
  }
  excel.path <- tempfile(fileext = ".xlsx")
  saveWorkbook(wb, file = excel.path, overwrite = TRUE)
  shell.exec(excel.path)
}

Taotao$addUnits <-  function(Response = "N application rate",
           Unit = "kg~m^-3~ha^-1", Preview=FALSE,
           groups = c("(", ")")) {

  out <- sprintf("'%s'~'%s'*%s*'%s'",
                 stri_escape_unicode(str = iconv(x = Response, to = "UTF-8")), 
                 groups[1], Unit, groups[2]) %>%
    parse(text = .,encoding = "UTF-8")
  if(Preview){
    plot.new()
    text(x = 0.5, y = 0.5, label = out)
  }
  return(out)
  }
Taotao$tibble_preview <- function(data_tibble) {
  require(tidyverse)
  require(shiny)
  require(DT)
  require(glue)
  require(excel.link)
  require(stringi)
  title <- substitute(data_tibble)
  ui <- fluidPage(# Application title
    titlePanel(glue("{title}数据集可视化")),
    h5("开发者：陈涛涛"),
    h5("开发日期：2018年9月21日"),
    helpText('本shiny小程序可实现tibble对象可视化，对于字符、数字等数据类型直接窗口中显示；
             对于data.frame、list或者嵌套tibble等复杂数据，可以通过双击窗口中"打开”按钮，用excel打开。'),
    uiOutput("tibble"))
  
  server <- function(input, output) {
    output$tibble <- renderUI({
      fluidRow(column(
        width = 8,
        DT::dataTableOutput("tibble_Window")  ,
        tags$script(
          HTML(
            '$(tibble_Window).on("click", "button", function () {
            Shiny.onInputChange("object_button",this.value);})'
        )
          )
        
        ))
    })
    output$tibble_Window <- renderDataTable({
      expand.grid(
        col = names(data_tibble),
        row = 1:nrow(data_tibble),
        stringsAsFactors = FALSE
      ) %>%
        mutate(button = map2(
          .x = row,
          .y = col,
          ~ ifelse(
            is.list(data_tibble[.x, .y][[1]][[1]]),
            glue(
              '<button type="button" name="object_button" value="{.y}:{.x}" >打开</button>'
            ) %>% as.character(),
            data_tibble[.x, .y][[1]][[1]]
          )
        )) %>%
        spread(key = col, value = button)  %>% dplyr::select(-row) %>%
        dplyr::select(one_of(names(data_tibble)))
    },
    server = FALSE,
    escape = FALSE,
    selection = 'none',
    options = list(ordering = FALSE,
                   columnDefs = list(list(className = 'dt-center',targets= seq_along(data_tibble)))))
    
    HSDData <- observeEvent(input$object_button, {
      temp <- input$object_button %>% stringr::str_split(":") %>% `[[`(1)
      out <- data_tibble[as.integer(temp[2]), temp[1]][[1]][[1]]
      xl.sheet.add(stri_rand_strings(1,length = 10))
      xlc$a1 <- out
    })
  }
  
  
  shinyApp(ui = ui, server = server) %>% print
  }

# Taotao$regenerating.boundary()
write_rds(Taotao,path = "00 SharedFuns/Taotao.rds")
