#' UsagiSan: A package for cleansing dataset and outputting statistical test results with using EXCEL.
#'
#' The package UsagiSan provides you a lot of helps to reduce the time on data-clansing and editting the test results. this package contains four function:
#' excelColor, excelHeadColor, mkDirectories and dataCleanser
#'
#' @section excelColor:
#' The function excelColor helps you to edit the test results with coloring the signigicant variables with specific color.
#'
#' @section excelHeadColor:
#' The function excelHeadColor helps you to add colors on headers of any type of tables including summaty sheets and statistical test tables.
#'
#' @section colorCells_xlsx:
#' The function colorCells_xlsx colors columns with specified condition for rows in a EXCEL-sheet. This enables to color any EXCEL-sheets overwriting previous workbook data. you can freely intuitively edit EXCEL-sheets.
#'
#' @section mkDirectories:
#' The function mkDirectories organizes files in tha working directory.
#'
#' @section dataCleanser:
#' The function dataCleanser helps you to cleans a dataset. This function modifies a dataset  to the form much easier to handle in statistical analysis.
#'
#' @section dataCleansing:
#' The class dataCleansing have several methods to operate dataCleanser. This provides how to classify a vector object into numeric, factor or Date type object.
#'
#' @section mergeRowAndColnamesWithData:
#' The function mergeRowAndColnamesWithData reshapes a dataframe so that it has own rownames and colnames as its components.
#'
#' @section rowBind:
#' The function rowBind is more useful and flexible than rbind. You can merge two objects without adjusting the number of rows or columns.
#'
#' @section colBind:
#' The function colBind is more useful and flexible than cbind. You can merge two objects without adjusting the number of rows or columns.
#'
#' @section getIndex:
#' The function getIndex gives you the indices of an item in a vector or a list.
#'
#' @section getCount:
#' The function getCount gives you the number of an item in a vector or a list.
#'
#' @section append_vecOrList:
#' The function append_vecOrList appends two vector or list type objects.
#'
#' @section extend_vecOrList:
#' The function getCount extend_vecOrList two vector or list type objects to make a extended vector or list.
#'
#' @section insert_vecOrList:
#' The function insert_vecOrList inserts an item at the specified index in a vector or list type object.
#'
#' @section remove_vecOrList:
#' The function remove_vecOrList removes an item from a vector or list type object.
#'
#' @section pop_vecOrList:
#' The function pop_vecOrList pops an item from a vector or list type object.
#'
#' @section isEqualData:
#' The function isEqualData compares two data in txt, csv or xlsx files and returns results whether the two data are the same or not (if not, the row and column indices where components are not equal are printed out).
#'
#' @section getSummaryTable :
#' The function getSummarytable Creates a summary table for numeric and factor data in terms of each level of a factor data.
#'
#' @seealso \code{\link{excelColor}}
#' @seealso \code{\link{excelHeadColor}}
#' @seealso \code{\link{excelHeadColor}}
#' @seealso \code{\link{mkDirectories}}
#' @seealso \code{\link{dataCleanser}}
#' @seealso \code{\linkS4class{dataCleansing}}
#' @seealso \code{\link{mergeRowAndColnamesWithData}}
#' @seealso \code{\link{rowBind}}
#' @seealso \code{\link{colBind}}
#' @seealso \code{\link{getIndex}}
#' @seealso \code{\link{getCount}}
#' @seealso \code{\link{append_vecOrList}}
#' @seealso \code{\link{extend_vecOrList}}
#' @seealso \code{\link{insert_vecOrList}}
#' @seealso \code{\link{remove_vecOrList}}
#' @seealso \code{\link{pop_vecOrList}}
#' @seealso \code{\link{isEqualData}}
#' @seealso \code{\link{getSummaryTable}}
#' @seealso My web site: \url{https://multivariate-statistics.com}
#'
#' @docType package
#' @name <UsagiSan>
NULL
#> NULL
#Discription
#The arguments of this function is as follows:

#dataName          :The name of a csv-file you want to edit with
#                   coloring(*)
#fileName          :The name of a Excel-file you want to save as
#level             :The significance level applied in coloring
#                   significant
#                   variables and p values
#significanceColor :The fore-ground-color of the significant variables
#headerColor       :The fore-ground-color of the headers of
#                   a data-frame(**)
#fontSize          :Font-size
#fontName          :Font-name
#fontColor         :The color of fonts
#intercept         :Allows you to color significant intercept variable
#                   with the fontColor
#adj               :Allows you yo adjust shifted statistical test
#                   tables

#Notes (*):  The data must include the column of which name is
#            "Pr(>|t|)" or "Pr(>|z|)"
#      (**): If the data have more than two results of statistical
#            tests, the process is applied for the each headers.

initialize <- function(data, pValue, header, mode) {
  if (mode != "header") {
    tmp <- apply(data, 1, function(x) {
      bar <- rep(FALSE, length(x))
      for (i in seq_len(length(pValue))) {
        bar <- bar | x == pValue[i]
      }
      x[bar]
    }
    )
    bar <- rep(FALSE, length(tmp))
    for (i in seq_len(length(pValue))) {
      bar <- bar | tmp == pValue[i]
    }

    tmp2 <- data[as.numeric(rownames(data[bar, ])), ]

    tmp_footer <- apply(data, 1, function(x) {
      all(x == "")
    }
    )
    green_header <- as.numeric(rownames(data[bar, ]))
    footer <- as.numeric(rownames(data[tmp_footer, ]))
    }
    else{
    tmp <- apply(data, 1, function(x) {
      x[header == x]
    }
    )
    tmp2 <- NULL
    tmp_footer <- apply(data, 1, function(x) {
      all(x == "")
    }
    )
    green_header <- as.numeric(rownames(data[tmp == header, ]))
    footer <- as.numeric(rownames(data[tmp_footer, ]))
  }

  return(list(tmp = tmp, tmp2 = tmp2, gr_header = green_header, footer = footer))
}

getPotision_sigVar_Intercpt_rLim <- function(green_header, data, pValue, tmp2, col_count, col_intercept, table_rightLim) {
  for (j in seq_len(length(green_header))) {
    for (k in seq_len(ncol(data))) {
      for (l in seq_len(length(pValue))) {
        if (tmp2[j, k] == pValue[l]) {
          col_count[j] <- k
        }
      }
    }
  }
  for (j in seq_len(length(green_header))) {
    for (k in seq_len(length(data[green_header[j], ]))) {
      if (data[green_header[j], k] == "" & data[green_header[j] + 1, k] != "") {
        col_intercept[j] <- k
      }
    }
  }
  for (j in seq_len(length(green_header))) {
    for (k in setdiff(seq_len(length(data[green_header[j], ])), seq_len((col_intercept[j])))) {
      if (data[green_header[j], k] == "") {
        table_rightLim[j] <- (k - 1)
        break
      }
      else if (k == ncol(data)) {
        table_rightLim[j] <- k
      }
    }
  }
  return(list(count = col_count, intercept = col_intercept, rightLim = table_rightLim))
}


removeNA <- function(factor_row, adj) {
  if (adj == TRUE) {
    for (i in seq_len(length(factor_row))) {
      if (!(is.integer(factor_row[[i]]) & length(factor_row[[i]]) == 0L)) {
        options(warn = -1)
        if (all(is.na(factor_row[[i]]))) {
          factor_row[[i]] <- integer(0)
        }
        options(warn = 0)
      }
      if (any(is.na(factor_row[[i]]))) {
        factor_row[[i]] <- as.vector(stats::na.omit(factor_row[[i]]))
      }
    }
  }
  else {
    for (i in seq_len(length(factor_row))) {
      if (!(is.integer(factor_row[i]) & length(factor_row[i]) == 0L)) {
        options(warn = -1)
        if (is.na(factor_row[i])) {
          factor_row[i] <- integer(0)
        }
        options(warn = 0)
      }
    }
  }
  return(factor_row)
}


#'
#' Coloring the signigicant variables and corresponding p values in statistical tests tables on a EXCEL sheet.
#' @encoding UTF-8
#'
#' @param dataName The name of a csv-file you want to edit with coloring.
#' @param fileName The name of a Excel-file you want to save as.
#' @param level The significance level applied in coloring significant variables and p values.
#' @param pValue The character object which indicates the column name in each stitistical test tables.
#' @param significanceColor The fore-ground-color of the significant variables.
#' @param headerColor The fore-ground-color of the headers of a data-frame.
#' @param fontSize Font-size.
#' @param fontName Font-name.
#' @param fontColor The color of fonts.
#' @param intercept Allows you to color significant intercept variable with the fontColor.
#' @param adj Allows you yo adjust shifted statistical test tables.
#' @param fileEncoding File-encoding.
#'
#' @importFrom stats na.omit
#' @importFrom utils read.table
#' @importFrom openxlsx createWorkbook
#' @importFrom openxlsx addWorksheet
#' @importFrom openxlsx createStyle
#' @importFrom openxlsx addStyle
#' @importFrom openxlsx writeData
#' @importFrom openxlsx modifyBaseFont
#' @importFrom openxlsx saveWorkbook
#'
#' @export
#'
excelColor <- function(dataName, fileName, level = 0.05, pValue = c("Pr(>|z|)", "Pr(>|t|)", "p-value"), significanceColor = "#FFFF00", headerColor = "#92D050", fontSize = 11, fontName = "Yu Gothic", fontColor = "#000000",  intercept = FALSE, adj = TRUE, fileEncoding = "CP932") {
  if (!is.character(dataName)) {
    stop("The data-name must be character")
  }
  if (!is.character(fileName)) {
    stop("The file-name must be character")
  }
  data <- utils::read.table(paste0(dataName, ".csv"), fill = TRUE, header = FALSE, sep = ",", blank.lines.skip = FALSE, fileEncoding = fileEncoding)
  data <- replace(data, is.na(data), "")
  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, "Sheet 1")
  st <- openxlsx::createStyle(fontName = fontName, fontSize = fontSize)
  openxlsx::addStyle(wb, "Sheet 1", style = st, cols = 1:2, rows = 1:2)
  openxlsx::writeData(wb, sheet = "Sheet 1", x = data, colNames = F, withFilter = F)
  openxlsx::modifyBaseFont(wb, fontSize = fontSize, fontColour = fontColor, fontName = fontName)

  init <- initialize(data, pValue, mode = "test")
  tmp <- init$tmp
  tmp2 <- init$tmp2
  green_header <- init$gr_header
  footer <- init$footer
  col_count <- NULL
  col_intercept <- NULL
  table_rightLim <- NULL

  options(warn = -1)
  if (length(tmp) == 0) {
    stop("There is no such pvalue' colname")
  }
  options(warn = 0)
  if (adj == TRUE) {
    cols <- getPotision_sigVar_Intercpt_rLim(green_header, data,
                                     pValue, tmp2,  col_count,
                                     col_intercept, table_rightLim)
    col_count <- cols$count
    col_intercept <- cols$intercept
    table_rightLim <- cols$rightLim

    #setStyle for headers
    for (i in seq_len(length(green_header))) {
      st <- openxlsx::createStyle(fontName = fontName, fontSize = fontSize, fgFill = headerColor)
      openxlsx::addStyle(wb, "Sheet 1", style = st, cols = col_intercept[i]:table_rightLim[i], rows = green_header[i])
    }

    factor_list <- list(NULL)
    factor_list <- mkFactorList(green_header, footer, factor_list, 1, data)

    factor_row <- list(NULL)
    for (i in seq_len(length(factor_list))) {
      bar <-  factor_list[[i]]
      factor_row[[i]] <- bar[as.numeric(data[factor_list[[i]], col_count[i]]) < level]
    }

    #removing NA
    factor_row <- removeNA(factor_row, adj)

    #write data
    writeDatas(factor_list, data, wb)

    #setStyle for significant variables
    st <- openxlsx::createStyle(fontName = fontName, fontSize = fontSize, fgFill = significanceColor)
    for (i in seq_len(length(factor_list))) {
      openxlsx::addStyle(wb, "Sheet 1", style = st, cols = col_count[i], rows = factor_row[[i]])
    }

    if (intercept == TRUE) {
      for (i in seq_len(length(factor_list))) {
        openxlsx::addStyle(wb, "Sheet 1", style = st, cols = col_intercept[i], rows = factor_row[[i]])
      }
    }else {
      for (i in seq_len(length(factor_list))) {
        openxlsx::addStyle(wb, "Sheet 1", style = st, cols = col_intercept[i], rows = setdiff(factor_row[[i]], green_header[i] + 1))
      }
    }
  }else {
    st <- openxlsx::createStyle(fontName = fontName, fontSize = fontSize, fgFill = headerColor)
    for (k in seq_len(ncol(data))) {
      for (l in seq_len(length(pValue))) {
        if (tmp2[1, k] ==  pValue[l]) {
          col_count[1] <- k
        }
      }
      openxlsx::addStyle(wb, "Sheet 1", style = st, cols = k, rows = green_header)
    }

    options(warn = -1)
    factor_row <- as.numeric(rownames(data[replace(as.numeric(data[, col_count[1]]) < level, is.na(as.numeric(data[, col_count[1]]) < level), FALSE), ]))
    options(warn = 0)

    #removing NA
    factor_row <- removeNA(factor_row, adj)

    factor_list <- list(NULL)
    factor_list <- mkFactorList(green_header, footer, factor_list, 1, data)

    #write data
    writeDatas(factor_list, data, wb)

    st <- openxlsx::createStyle(fontName = fontName, fontSize = fontSize, fgFill = significanceColor)
    openxlsx::addStyle(wb, "Sheet 1", style = st, cols = col_count[1], rows = factor_row)

    if (intercept == TRUE) {
      openxlsx::addStyle(wb, "Sheet 1", style = st, cols = 1, rows = factor_row)
    }else {
      openxlsx::addStyle(wb, "Sheet 1", style = st, cols = 1, rows = setdiff(factor_row, green_header + 1))
    }
  }

  openxlsx::saveWorkbook(wb, paste0(fileName, ".xlsx"), overwrite = TRUE)
}

getPosition_intercpt_rightLim <- function(green_header, data, col_intercept, table_rightLim) {
  for (j in seq_len(length(green_header))) {
    for (k in seq_len(length(data[green_header[j], ]))) {
      if (data[green_header[j], k] == "" & data[green_header[j] + 1, k] != "") {
        col_intercept[j] <- k
      }else {
        if (k < ncol(data)) {
          if (data[green_header[j], k] == "" & data[green_header[j] + 1, k] == "" & data[green_header[j], k + 1] != "") {
            col_intercept[j] <- k + 1
          }
        }
      }
    }
    if (is.na(col_intercept[j])) {
      col_intercept[j] <- 1
    }
  }
  for (j in seq_len(length(green_header))) {
    for (k in setdiff(seq_len(length(data[green_header[j], ])), seq_len((col_intercept[j])))) {
      if (data[green_header[j], k] == "") {
        table_rightLim[j] <- (k - 1)
        break
      }
      else if (k == ncol(data)) {
        table_rightLim[j] <- k
      }
    }
  }
  return(list(intercept = col_intercept, rightLim = table_rightLim))
}

mkFactorList <- function(green_header, footer, factor_list, count, data) {
  for (i in seq_len(length(green_header))) {
    if (!is.na(footer[count])) {
      while (green_header[i] >= (footer[count] - 1) & count != length(footer)) {
        count <- count + 1
      }
      if (count != length(footer)) {
        factor_list[[i]] <- (green_header[i] + 1):(footer[count] - 1)
        count <- count + 1
      }else {
        if (green_header[i] < footer[count]) {
          factor_list[[i]] <- (green_header[i] + 1):(footer[count] - 1)
        }else {
          factor_list[[i]] <- (green_header[i] + 1):nrow(data)
        }
      }
    }else {
      factor_list[[i]] <- (green_header[i] + 1):nrow(data)
    }
  }
  return(factor_list)
}

writeDatas <- function(factor_list, data, wb) {
  options(warn = -1)
  lapply(factor_list, function(x) {
    bar <- data[x, ]
    for (i in seq_len(ncol(data[x, ]))) {
      bar[bar[, i] == "", i] <- NA
      if (!any(is.na(as.numeric(stats::na.omit(bar[, i]))))) {
        if (all(as.numeric(stats::na.omit(bar[, i])) == as.character(as.numeric(stats::na.omit(bar[, i]))))) {
          bar[, i] <- as.numeric(bar[, i])
        }
      }
    }
    openxlsx::writeData(wb, sheet = "Sheet 1", x = bar, startRow = x[1], startCol = 1, colNames = F, withFilter = F)
  })
  options(warn = 0)
}
#'
#' Coloring headers of tables on a EXCEL sheet
#' @encoding UTF-8
#'
#' @param dataName The name of a csv-file you want to edit with coloring.
#' @param fileName The name of a Excel-file you want to save as.
#' @param header The character object included in each headers of tables.
#' @param headerColor The fore-ground-color of the headers of a data-frame.
#' @param fontSize Font-size.
#' @param fontName Font-name.
#' @param fontColor The color of fonts.
#' @param adj Allows you yo adjust shifted statistical test tables.
#' @param fileEncoding File-encoding.
#'
#' @importFrom stats na.omit
#' @importFrom utils read.table
#' @importFrom openxlsx createWorkbook
#' @importFrom openxlsx addWorksheet
#' @importFrom openxlsx createStyle
#' @importFrom openxlsx addStyle
#' @importFrom openxlsx writeData
#' @importFrom openxlsx modifyBaseFont
#' @importFrom openxlsx saveWorkbook
#'
#' @export
#'
excelHeadColor <- function(dataName, fileName, header, headerColor = "#92D050", fontSize = 11, fontName = "Yu Gothic", fontColor = "#000000", adj = TRUE, fileEncoding = "CP932") {
  if (!is.character(dataName)) {
    stop("The data-name must be character")
  }
  if (!is.character(fileName)) {
    stop("The file-name must be character")
  }
  data <- utils::read.table(paste0(dataName, ".csv"), fill = TRUE, header = FALSE, sep = ",", blank.lines.skip = FALSE, fileEncoding = fileEncoding)
  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, "Sheet 1")
  st <- openxlsx::createStyle(fontName = fontName, fontSize = fontSize)
  openxlsx::addStyle(wb, "Sheet 1", style = st, cols = 1:2, rows = 1:2)
  openxlsx::writeData(wb, sheet = "Sheet 1", x = data, colNames = F, withFilter = F)
  openxlsx::modifyBaseFont(wb, fontSize = fontSize, fontColour = fontColor, fontName = fontName)

  init <- initialize(data, header = header, mode = "header")
  green_header <- init$gr_header
  footer <- init$footer
  col_intercept <- NA
  table_rightLim <- NULL

  if (adj == TRUE) {
    cols <- getPosition_intercpt_rightLim(green_header, data, col_intercept, table_rightLim)
    col_intercept <- cols$intercept
    table_rightLim <- cols$rightLim

    #setStyle for headers
    for (i in seq_len(length(green_header))) {
      st <- openxlsx::createStyle(fontName = fontName, fontSize = fontSize, fgFill = headerColor)
      openxlsx::addStyle(wb, "Sheet 1", style = st, cols = col_intercept[i]:table_rightLim[i], rows = green_header[i])
    }

    factor_list <- list(NULL)
    factor_list <- mkFactorList(green_header, footer, factor_list, 1, data)

    #write data
    writeDatas(factor_list, data, wb)
  }
  else {
    st <- openxlsx::createStyle(fontName = fontName, fontSize = fontSize, fgFill = headerColor)
    for (k in seq_len(ncol(data))) {
      openxlsx::addStyle(wb, "Sheet 1", style = st, cols = k, rows = green_header)
    }
    factor_list <- list(NULL)
    factor_list <- mkFactorList(green_header, footer, factor_list, 1, data)

    #write data
    writeDatas(factor_list, data, wb)
  }
  openxlsx::saveWorkbook(wb, paste0(fileName, ".xlsx"), overwrite = TRUE)
}
#'
#' Coloring the columns specifying rows conditions
#' @encoding UTF-8
#'
#' @param dataName The name of a csv-file you want to edit with coloring.
#' @param fileName The name of a Excel-file you want to save as.
#' @param sheetName The name of a EXCEL-sheet you want to color.
#' @param coloredCols The Columns colored with the specified color.
#' @param coloredCondition The condition for rows. This argument must be a logical vector in which TRUE components indicate row indices colored and neither FALSE components dose.
#' @param cellColor The color you want to color with
#' @param fontSize Font-size.
#' @param fontName Font-name.
#' @param fontColor The color of fonts.
#'
#' @importFrom openxlsx loadWorkbook
#' @importFrom openxlsx createWorkbook
#' @importFrom openxlsx addWorksheet
#' @importFrom openxlsx createStyle
#' @importFrom openxlsx addStyle
#' @importFrom openxlsx writeData
#' @importFrom openxlsx modifyBaseFont
#' @importFrom openxlsx saveWorkbook
#'
#' @export
#'
colorCells_xlsx <- function(dataName, fileName, sheetName, coloredCols, coloredCondition = NULL, cellColor, fontSize = 11, fontName = "Yu Gothic", fontColor = "#000000") {
  if (!is.character(dataName)) {
    stop("The data-name must be character")
  }
  if (!is.character(fileName)) {
    stop("The file-name must be character")
  }
  wb <- openxlsx::loadWorkbook(paste0(dataName, ".xlsx"))
  openxlsx::modifyBaseFont(wb, fontSize = fontSize, fontColour = fontColor, fontName = fontName)
  data <- as.data.frame(openxlsx::read.xlsx(paste0(dataName, ".xlsx"), sheet = sheetName, colNames = FALSE))
  data <- replace(data, is.na(data), "")
  st <- openxlsx::createStyle(fontName = fontName, fontSize = fontSize, fgFill = cellColor)
  coloredRows <- as.numeric(rownames(data[coloredCondition, ]))
  for (i in coloredCols) {
    openxlsx::addStyle(wb, sheetName, style = st, cols = i, rows = coloredRows)
  }
  openxlsx::saveWorkbook(wb, paste0(fileName, ".xlsx"), overwrite = TRUE)
}

mkDirFor_result_data <- function(parentDirName, childDirName, extension, file) {
  if (is.na(strsplit(file, "\\.")[[1]][2]) & any(is.na(extension))) {
    dir.create(paste0(getwd(), "/", parentDirName, "/", childDirName, "/", "No Extension"))
    extension[is.na(extension)] <- ""
  }
  if (!is.na(any(extension == strsplit(file, "\\.")[[1]][length(strsplit(file, "\\.")[[1]])]))) {
    if (any(extension == strsplit(file, "\\.")[[1]][length(strsplit(file, "\\.")[[1]])])) {
      options(warn = -1)
      dir.create(paste0(getwd(), "/", parentDirName, "/", childDirName, "/", strsplit(file, "\\.")[[1]][length(strsplit(file, "\\.")[[1]])]))
      options(warn = 0)
      extension[extension == strsplit(file, "\\.")[[1]][length(strsplit(file, "\\.")[[1]])]] <- ""
    }
  }
  if (is.na(strsplit(file, "\\.")[[1]][2])) {
    file.copy(paste0(getwd(), "/", file), paste0(getwd(), "/", parentDirName, "/", childDirName, "/", "No Extension", "/", file))
  }else {
    file.copy(paste0(getwd(), "/", file), paste0(getwd(), "/", parentDirName, "/", childDirName, "/", strsplit(file, "\\.")[[1]][length(strsplit(file, "\\.")[[1]])], "/", file))
  }
  return(extension)
}

mkDir_noArrange_esult_data <- function(parentDirName, childDirName, file) {
  file.copy(paste0(getwd(), "/", file), paste0(getwd(), "/", parentDirName, "/", childDirName, "/", file))
}
#'
#' Making directories to organize three kinds of datas: data-sets,  script-files and result-files
#' @encoding UTF-8
#'
#' @param parentDirName The name of a parent-directory containing organize datas-files, script-files and result-files.
#' @param dataDirName The name of a directory to organize data-files.
#' @param programmingDirName The name of a directory to organize script-files.
#' @param resultDirName The name of a directory to organize result-files.
#' @param updateTime The time used to divide data-filese into two directories, one is for datas and the other is for results.
#' @param arrange Allows you to organize data-files in the form of file extensions.
#'
#' @export
#'
mkDirectories <- function(parentDirName, dataDirName="data", programmingDirName="program", resultDirName="result", updateTime=1, arrange = TRUE) {
  dir.create(paste0(getwd(), "/", parentDirName))
  dir.create(paste0(getwd(), "/", parentDirName, "/", dataDirName))
  dir.create(paste0(getwd(), "/", parentDirName, "/", programmingDirName))
  dir.create(paste0(getwd(), "/", parentDirName, "/", resultDirName))
  files <- list.files()
  R.files <- grep("\\.R$", files)

  fileExtension <- NULL
  for (i in seq_len(length(strsplit(files[- (R.files)], "\\.")))) {
    if (length(strsplit(files[- (R.files)], "\\.")[[i]]) == 1) {
      fileExtension[length(fileExtension) + 1] <- strsplit(files[- (R.files)], "\\.")[[i]][2]
    }else {
      fileExtension[length(fileExtension) + 1] <- strsplit(files[- (R.files)], "\\.")[[i]][length(strsplit(files[- (R.files)], "\\.")[[i]])]
    }
  }

  fileExtension <- unique(fileExtension)
  resultExtension <- fileExtension
  dataExtension <- fileExtension

  for (i in files[- (R.files)]) {
    if (!is.na(file.info(paste0(getwd(), "/", i))$mtime) &  i != parentDirName) {
      get_resultOrData <- as.numeric(as.POSIXct(as.list(file.info(paste0(getwd(), "/", i)))$mtime, format = "%Y-%m-%d  %H:%M:%S", tz = "Japan") - Sys.time(), units = "mins") > (-1) * updateTime * 60
      if (arrange == TRUE) {
        if (get_resultOrData) {
          resultExtension <- mkDirFor_result_data(parentDirName, resultDirName, resultExtension, i)
        }
        else {
          dataExtension <- mkDirFor_result_data(parentDirName, dataDirName, dataExtension, i)
        }
      }
      else {
        if (get_resultOrData) {
          mkDir_noArrange_esult_data(parentDirName, resultDirName, i)
        }
        else {
          mkDir_noArrange_esult_data(parentDirName, dataDirName, i)
        }
      }
    }
  }
  for (i in files[R.files]) {
    file.copy(paste0(getwd(), "/", i), paste0(getwd(), "/", parentDirName, "/", programmingDirName, "/", i))
  }
}

mkNumericTable <- function(data, index) {
  table <- c(index, rep("", 7))
  table <- rbind(table, c("", "Missing values", "", "Replace the column B with the spesific numbers", "", "Breaks", "", "Labels"))
  options(warn = -1)
  if (length(unique(data[is.na(as.numeric(data[, index])), index])) > 0) {
    numericData <- cbind(rep("", length(unique(data[is.na(as.numeric(data[, index])), index]))),
                         unique(data[is.na(as.numeric(data[, index])), index]),
                         rep("", length(unique(data[is.na(as.numeric(data[, index])), index]))),
                         rep("", length(unique(data[is.na(as.numeric(data[, index])), index]))),
                         rep("", length(unique(data[is.na(as.numeric(data[, index])), index]))),
                         rep("", length(unique(data[is.na(as.numeric(data[, index])), index]))),
                         rep("", length(unique(data[is.na(as.numeric(data[, index])), index]))),
                         rep("", length(unique(data[is.na(as.numeric(data[, index])), index]))))
  }
  else{
    numericData <- NULL
  }
  options(warn = 0)
  table <- rbind(table, numericData)
  table <- rbind(table, rep("", 8))
  return(table)
}

mkFactorTable <- function(data, index) {
  table <- c(index, rep("", 8))
  table <- rbind(table, c("", "", "Levels", "", "Replace the column C with", "", "Pool the column C", "", "The Order of levels"))
  factorData <- cbind(rep("", nlevels(as.factor(data[, index]))),
                      paste0("No.", 1:nlevels(as.factor(data[, index]))), levels(as.factor(data[, index])),
                      rep("", nlevels(as.factor(data[, index]))),
                      rep("", nlevels(as.factor(data[, index]))),
                      rep("", nlevels(as.factor(data[, index]))),
                      rep("", nlevels(as.factor(data[, index]))),
                      rep("", nlevels(as.factor(data[, index]))),
                      rep("", nlevels(as.factor(data[, index]))))

  table <- rbind(table, factorData)
  table <- rbind(table, t(c(rep("", 9))))
  return(table)
}

formatOrder <- function(formatStr) {
  order <- NULL
  for (i in formatStr) {
    if (i == "Y") {
      order <- append(order, 1)
    }
    else if (i == "m") {
      order <- append(order, 2)
    }
    else {
      order <- append(order, 3)
    }
  }
  return(order)
}

dateClassifier <- function(data, index) {
  DateFormats <- list(c("Y", "m", "d"), c("Y", "d", "m"), c("m", "Y", "d"), c("m", "d", "Y"), c("d", "Y", "m"), c("d", "m", "Y"))
  formatAndDate <- NULL
  delimiter <- lapply(strsplit(data[, index], ""), function(x) {
    return(x[grep("[^0-9]", x)])
  })
  indexDelim <- lapply(strsplit(data[, index], ""), function(x) {
    dateList <- list()
    delStr <- c(0, grep("[^0-9]", x))
    for (i in seq_len(length(delStr))) {
      if (i == length(delStr) & length(delStr) < 4) {
        dateList[[i]] <- (delStr[i] + 1) : length(x)
      }
      else if (i < length(delStr)) {
        dateList[[i]] <- (delStr[i] + 1) : (delStr[i + 1] - 1)
      }
    }
    return(list(index = dateList, delim = x[grep("[^0-9]", x)], indexDelim = grep("[^0-9]", x), data = x))
  })
  arrangedIndex <- lapply(indexDelim, function(x) {
    index <- NULL
    for (i in seq_len(2)) {
      index <- append(index, c(x$index[[order(x$delim)[i]]], x$indexDelim[order(x$delim)[i]]))
    }
    if (length(x$delim) == 2) {
      index <- append(index, x$index[[3]])
    }
    else {
      index <- append(index, x$index[[order(x$delim[3])]])
    }
    return(list(index = index, data = x$data))
  })
  data[, index] <- unlist(lapply(arrangedIndex, function(x) {
    return(paste(x$data[x$index], collapse = ""))
  }))

  for (formatStr in DateFormats) {
    format <- paste0("%", formatStr[1], "-%", formatStr[2], "-%", formatStr[3])
    options(warn = -1)

    dateData <- replace(data[, index], data[, index] == "", NA)
    dateData <- gsub("[^0-9]", "-", as.character(dateData))
    dateData <- replace(dateData, nchar(gsub("[^-]", "", dateData)) != 2 | lapply(strsplit(dateData, "-"), length) != 3, NA)
    date_dataFrame <- as.data.frame(strsplit(dateData, "-"), stringsAsFactors = FALSE)
    YmdOrder <- c(match("Y", formatStr), match("m", formatStr), match("d", formatStr))
    charData <- paste0(formatC(as.numeric(date_dataFrame[YmdOrder[1], ]), width = 4, flag = "0"), "-",
                       formatC(as.numeric(date_dataFrame[YmdOrder[2], ]), width = 2, flag = "0"), "-",
                       formatC(as.numeric(date_dataFrame[YmdOrder[3], ]), width = 2, flag = "0"))
    options(warn = 0)
    formatAndDate[[length(formatAndDate) + 1]] <- list(format = format, date = charData)
  }
  return(formatAndDate)
}

changeColName <- function(data, colname, refData, rowNumber) {
  changedName <- NULL
  if (!is.na(refData[rowNumber, 2])) {
    changedName <- refData[rowNumber, 2]
  }
  else {
    changedName <- colname
  }
  return(changedName)
}

replaceMissVal <- function(data, refData, rowNumber) {
  missVal <- unique(data[is.na(as.numeric(data))])
  for (j in seq_len(length(missVal))) {
    if (!is.na(refData[rowNumber + 1 + j, 4])) {
      data <- replace(data, data == missVal[j], refData[rowNumber + 1 + j, 4])
    }
  }
  return(data)
}

cutting <- function(data, refData, rowNumber) {
  if (!is.na(refData[rowNumber + 2,  8])) {
    labels <- strsplit(refData[rowNumber + 2, 8], ",")
    breaks <- strsplit(refData[rowNumber + 2, 6], ",")
    data <- cut(data, breaks = as.numeric(breaks[[1]]), labels = labels[[1]], right = FALSE)
  }
  else {
    breaks <- strsplit(refData[rowNumber + 2, 6], ",")
    data <- cut(data, breaks = as.numeric(breaks[[1]]), right = FALSE)
  }
  return(data)
}

orderer <- function(data, refData, rowNumber) {
  orderVector <- rep(NA, nlevels(as.factor(data)))
  factorLength <- 0
  rowIndex <- 1
  while (!is.na(refData[rowNumber + 1 + rowIndex, 2])) {
    factorLength <- factorLength + 1
    rowIndex <- rowIndex + 1
  }
  pooledFactor <- strsplit(refData[(rowNumber + 2):(rowNumber + 1 + factorLength), 7], "[+]")
  poolLevel <- NULL
  for (i in seq_len(length(pooledFactor))) {
    refDataPool <- NULL
    refDataLevels <- NULL
    if (length(pooledFactor[[i]]) > 1) {
      for (j in seq_len(length(pooledFactor[[i]]))) {
        if (!is.na(refData[rowNumber + 1 + as.numeric(pooledFactor[[i]][j]), 5])) {
          refDataPool <- paste0(refDataPool, "+",  refData[rowNumber + 1 + as.numeric(pooledFactor[[i]][j]), 5])
        }
        else {
          refDataPool <- paste0(refDataPool, "+",  refData[rowNumber + 1 + as.numeric(pooledFactor[[i]][j]), 3])
        }
        refDataLevels <- append(refDataLevels, as.numeric(pooledFactor[[i]][j]))
      }
      refDataPool <- substr(refDataPool, 2, nchar(refDataPool))
      poolLevel[[i]] <- list(pool = refDataPool, levels = refDataLevels)
    }
  }
  for (i in seq_len(length(poolLevel))) {
    refData[rowNumber + 1 + poolLevel[[i]]$levels, 5] <- poolLevel[[i]]$pool
  }
  for (j in seq_len(factorLength)) {
    if (!is.na(refData[rowNumber + 1 + j, 9])) {
      orderNum <- as.numeric(refData[rowNumber + 1 + j, 9])
      if (!is.na(refData[rowNumber + 1 + j, 5])) {
        if (refData[rowNumber + 1 + j, 5] != "N/A") {
          orderVector[orderNum] <- refData[rowNumber + 1 + j, 5]
        }
      }
      else {
        orderVector[orderNum] <- refData[rowNumber + 1 + j, 3]
      }
    }
  }
  remainLevels <- setdiff(levels(as.factor(data)), orderVector)
  count <- 1
  for (i in seq_len(length(orderVector))) {
    if (is.na(orderVector[i])) {
      orderVector[i] <- remainLevels[count]
      count <- count + 1
    }
  }
  return(factor(data, levels = orderVector))
}

pooledName <- function(data, refData, rowNumber) {
  pooledData <- pooler(data, refData, rowNumber)
  pooledFactor <- ""
  for (j in grep("[+]", levels(pooledData))) {
    pooledFactor <- paste0(pooledFactor, ".", levels(pooledData)[j])
  }
  return(substr(pooledFactor, 2, nchar(pooledFactor)))
}

pooler <- function(data, refData, rowNumber) {
  poolList <- strsplit(refData[(rowNumber + 2) : (rowNumber + 1 + nlevels(as.factor(data))), 7], "[+]")
  data <- as.character(data)
  for (j in seq_len(length(poolList))) {
    if (length(poolList[[j]]) > 1) {
      poolLevel <- ""
      for (k in seq_len(length(poolList[[j]]))) {
        if (!is.na(refData[rowNumber + 1 + as.numeric(poolList[[j]][k]), 5])) {
          poolLevel <- paste0(poolLevel, "+",  refData[rowNumber + 1 + as.numeric(poolList[[j]][k]), 5])
        }
        else {
          poolLevel <- paste0(poolLevel, "+",  refData[rowNumber + 1 + as.numeric(poolList[[j]][k]), 3])
        }
      }
      poolLevel <- substr(poolLevel, 2, nchar(poolLevel))
      for (k in seq_len(length(poolList[[j]]))) {
        if (!is.na(refData[rowNumber + 1 + as.numeric(poolList[[j]][k]), 5])) {
          data <- replace(data, data == refData[rowNumber + 1 + as.numeric(poolList[[j]][k]), 5], poolLevel)
        }
        else {
          data <- replace(data, data == refData[rowNumber + 1 + as.numeric(poolList[[j]][k]), 3], poolLevel)
        }
      }
    }
  }
  levels <- levels(as.factor(data))
  return(factor(data, levels = levels))
}

replacer <- function(data, refData, rowNumber) {
  levelsRow <- refData[(rowNumber + 2) : (rowNumber + 1 + nlevels(as.factor(data))), 5] != ""
  factorLength <- 0
  rowIndex <- 1
  while (!is.na(refData[rowNumber + 1 + rowIndex, 2])) {
    factorLength <- factorLength + 1
    rowIndex <- rowIndex + 1
  }
  data <- as.character(data)
  refData[(rowNumber + 2) : (rowNumber + 1 + factorLength), 3] <- replace(refData[(rowNumber + 2) : (rowNumber + 1 + factorLength), 3], is.na(refData[(rowNumber + 2) : (rowNumber + 1 + factorLength), 3]), "")
  for (j in seq_len(nlevels(as.factor(data)))) {
    if (!is.na(levelsRow[j])) {
      if (refData[rowNumber + 1 + j, 5] == "N/A") {
        data <- replace(data, data == refData[rowNumber + 1 + j, 3], NA)
      }
      else {
        data <- replace(data, data == refData[rowNumber + 1 + j, 3], refData[rowNumber + 1 + j, 5])
      }
    }
  }
  levels <- levels(as.factor(data))
  return(factor(data, levels = levels))
}

readData <- function(dataName, fileEncoding) {
  if (fileEncoding == "UTF-8" | fileEncoding == "Latin-1") {
    data <- data.table::fread(paste0(dataName, ".csv"), encoding = fileEncoding)
  }
  else {
    data <- data.table::fread(paste0(dataName, ".csv"), encoding = "unknown")
  }
  return(data)
}

mkTimeTable <- function(data, index, tableTime, leastNumOfDate) {
  asDatedVector <- rep(FALSE, nrow(data))
  haveNumber <- replace(rep(FALSE, nrow(data)), grep("[0-9]", data[, index]), TRUE)
  haveLeastNumStr <- nchar(data[, index]) > 4
  delimiter <- nchar(gsub("[0-9]", "", data[, index])) == 2 | nchar(gsub("[0-9]", "", data[, index])) == 3
  if (length(data[haveNumber & haveLeastNumStr & delimiter, index]) <= leastNumOfDate) {
    return(NULL)
  }
  formCharData <- dateClassifier(data, index)
  for (i in list(c(1, 2), c(3, 5), c(4, 6))) {
    isFirstFormat <- !is.na(isCorrectFormat(formCharData[[i[1]]])) & isCorrectFormat(formCharData[[i[1]]]) & !asDatedVector
    isSecondFormat <- !is.na(isCorrectFormat(formCharData[[i[2]]])) & isCorrectFormat(formCharData[[i[2]]]) & !asDatedVector
    if (length(data[isFirstFormat, index]) > 0 | length(data[isSecondFormat, index]) > 0) {
      if (length(data[isFirstFormat, index]) > length(data[isSecondFormat, index])) {
        asDatedVector[isFirstFormat] <- TRUE
      }
      else {
        asDatedVector[isSecondFormat] <- TRUE
      }
    }
  }
  if (length(asDatedVector[asDatedVector == TRUE]) > leastNumOfDate) {
    tableTime <- rbind(tableTime, c(index, ""))
    return(tableTime)
  }
  return(NULL)
}

writeTablesOnExcel <- function(tableNumeric, tableFactor, tableTime, dataName, filePath = "") {
  sheetNumeric <- list(table = tableNumeric, sheetName = "numeric")
  sheetFactor <- list(table = tableFactor, sheetName = "factor")
  sheetTime <- list(table = tableTime, sheetName = "Date")
  wb <- openxlsx::createWorkbook()
  for (i in list(sheetNumeric, sheetFactor, sheetTime)) {
    openxlsx::addWorksheet(wb, i$sheetName)
  }
  st <- openxlsx::createStyle(fontName = "Yu Gothic", fontSize = 11)
  for (i in list(sheetNumeric, sheetFactor, sheetTime)) {
    openxlsx::addStyle(wb, i$sheetName, style = st, cols = 1:2, rows = 1:2)
    openxlsx::writeData(wb, sheet = i$sheetName, x = i$table, colNames = F, withFilter = F)
  }
  openxlsx::modifyBaseFont(wb, fontSize = 11, fontColour = "#000000", fontName = "Yu Gothic")
  openxlsx::saveWorkbook(wb, paste0(filePath, "dataCleansingForm_", dataName, "_.xlsx"), overwrite = TRUE)
}

mkTableNum_Fac <- function(data, index, numOrFac, tableNumeric, tableFactor) {
  options(warn = -1)
  charEqualNum <-  as.character(as.numeric(data[, index])) == as.numeric(data[, index]) #todo check
  options(warn = 0)
  if (length(na.omit(data[charEqualNum == FALSE, index])) == 0 &  length(na.omit(data[charEqualNum == TRUE, index])) > 0 & nlevels(as.factor(data[, index])) > nrow(data) / numOrFac) {
    numTab <- mkNumericTable(data, index)
    tableNumeric <- rbind(tableNumeric, numTab)
  }
  else {
    facTab <- mkFactorTable(data, index)
    tableFactor <- rbind(tableFactor, facTab)
  }
  return(list(num = tableNumeric, fac = tableFactor))
}

cleansNumeric <- function(data, index, refData, append) {
  if (any(refData[, 1] == index)) {
    options(warn = -1)
    rowNumber <- as.numeric(rownames(refData[refData[, 1] == index & !is.na(refData[, 1]), ]))
    colnames(data)[colnames(data) == index] <- changeColName(data, index, refData, rowNumber)
    index <- changeColName(data, index, refData, rowNumber)
    if (any(!is.na(refData[(rowNumber + 2):(rowNumber + 1 + length(unique(data[is.na(as.numeric(data[, index])), index]))), 4])) & length(unique(data[is.na(as.numeric(data[, index])), index])) > 0) {
      if (append == TRUE) {
        data <- cbind(data, replaceMissVal(data[, index], refData, rowNumber))
        colnames(data)[ncol(data)] <- paste0(index, "_missing Values replaced")
      }
      else {
        data[, index] <- replaceMissVal(data[, index], refData, rowNumber)
      }
    }
    data[, index] <- as.numeric(data[, index])
    options(warn = 0)
    if (any(!is.na(refData[rowNumber + 2, 6]))) {
      if (append == TRUE) {
        data <- cbind(data, cutting(data[, index], refData, rowNumber))
        colnames(data)[ncol(data)] <- paste0(index, "_categorized")
      }
      else {
        data[, index] <- cutting(data[, index], refData, rowNumber)
      }
    }
  }
  return(data)
}

cleansFactor <- function(data, index, refData, append) {
  if (any(refData[, 1] == index)) {
    pooling <- FALSE
    ordering <- FALSE
    options(warn = -1)
    rowNumber <- as.numeric(rownames(refData[refData[, 1] == index & !is.na(refData[, 1]), ]))
    options(warn = 0)
    colnames(data)[colnames(data) == index] <- changeColName(data, index, refData, rowNumber)
    index <- changeColName(data, index, refData, rowNumber)
    nrowLevel <- nlevels(as.factor(data[, index]))
    if (any(!is.na(refData[(rowNumber + 2):(rowNumber + 1 +  nrowLevel), 5]))) { #5 for replace
      data[, index] <- replacer(data[, index], refData, rowNumber)
      if (any(!is.na(refData[(rowNumber + 2):(rowNumber + 1 + nlevels(as.factor(data[, index]))), 7]))) { #7 for pool
        if (append == TRUE) {
          data <- cbind(data, pooler(data[, index], refData, rowNumber))
          colnames(data)[ncol(data)] <- paste0(index, "_", pooledName(data[, index], refData, rowNumber))
          index <- colnames(data)[ncol(data)]
        }
        else {
          data[, index] <- pooler(data[, index], refData, rowNumber)
        }
        pooling <- TRUE
        if (any(!is.na(refData[(rowNumber + 2):(rowNumber + 1 +  nrowLevel), 9]))) { #9 for order
          data[, index] <- orderer(data[, index], refData, rowNumber)
          ordering <- TRUE
        }
      }
    }
    if (any(!is.na(refData[(rowNumber + 2):(rowNumber + 1 + nlevels(as.factor(data[, index]))), 7])) & pooling == FALSE) {
      if (append == TRUE) {
        data <- cbind(data, pooler(data[, index], refData, rowNumber))
        colnames(data)[ncol(data)] <- paste0(index, "_", pooledName(data[, index], refData, rowNumber))
      }
      else {
        data[, index] <- pooler(data[, index], refData, rowNumber)
      }
      pooling <- TRUE
      if (any(!is.na(refData[(rowNumber + 2):(rowNumber + 1 + nlevels(as.factor(data[, index]))), 9])) & ordering == FALSE) {
        data[, index] <- orderer(data[, index], refData, rowNumber)
        ordering <- TRUE
      }
    }
    if (any(!is.na(refData[(rowNumber + 2):(rowNumber + 1 + nlevels(as.factor(data[, index]))), 9])) & ordering == FALSE) {
      data[, index] <- orderer(data[, index], refData, rowNumber)
      ordering <- TRUE
    }
    data[, index] <- as.factor(data[, index])
  }
  return(data)
}

cleansDate <- function(data, index, refData) {
  if (any(refData[, 1] == index)) {
    asDatedVector <- rep(FALSE, nrow(data))
    options(warn = -1)
    rowNumber <- as.numeric(rownames(refData[refData[, 1] == index & !is.na(refData[, 1]), ]))
    colnames(data)[colnames(data) == index] <- changeColName(data, index, refData, rowNumber)
    index <- changeColName(data, index, refData, rowNumber)
    formCharData <- dateClassifier(data, index)
    for (i in list(c(1, 2), c(3, 5), c(4, 6))) {
      isFirstFormat <- !is.na(isCorrectFormat(formCharData[[i[1]]])) & isCorrectFormat(formCharData[[i[1]]]) & !asDatedVector
      isSecondFormat <- !is.na(isCorrectFormat(formCharData[[i[2]]])) & isCorrectFormat(formCharData[[i[2]]]) & !asDatedVector
      if (length(data[isFirstFormat, index]) > 0 | length(data[isSecondFormat, index]) > 0) {
        if (length(data[isFirstFormat, index]) > length(data[isSecondFormat, index])) {
          data[isFirstFormat, index] <- formCharData[[i[1]]]$date[isFirstFormat]
          asDatedVector[isFirstFormat] <- TRUE
        }
        else {
          data[isSecondFormat, index] <- formCharData[[i[2]]]$date[isSecondFormat]
          asDatedVector[isSecondFormat] <- TRUE
        }
      }
    }
    options(warn = 0)
  }
  return(data)
}

isCorrectFormat <- function(dateAsFormat) {
  splitedFCD <- strsplit(dateAsFormat$date, "-")
  options(warn = -1)
  correctDate <- unlist(lapply(splitedFCD, function(x) {
    conditionFormat <- nchar(x[1]) == 4 & nchar(x[2]) == 2 & nchar(x[3]) == 2
    conditionUpper <- as.numeric(x[2]) < 13 & as.numeric(x[3] < 32)
    if (conditionFormat & conditionUpper) {
      return(TRUE)
    }
    else {
      return(FALSE)
    }
  }))
  options(warn = 0)
  return(correctDate)
}

#' Cleansing the dataset on a csv-file to change its form to more arranged one to handle.
#' @encoding UTF-8
#'
#' @param dataName The file-name of a csv file that will be cleansed.
#' @param append Allows you to append the new datas generated from dataCleansingForm__.xlsx.
#' @param numOrFac The criteria for classifying whether the column data is numeric or factor. If the number of levels are greater than the ratio (nrow(data)/numOrFac), then it will be assiged to numeric group.
#' @param leastNumOfDate The criteria for classifying whether the column data is Date of numeric. if the data contains the dateFormat you have chosen and the number of data containing such formats is greater than this value, leastNumOfDate, then the data will be assigned to Date group.
#' @param fileEncoding File-encoding
#'
#' @importFrom data.table fread
#' @export
dataCleanser <- function(dataName, append = FALSE, numOrFac = 10, leastNumOfDate = 10, fileEncoding = "CP932") {
  files <- list.files()
  if (any(files == paste0("dataCleansingForm_", dataName, "_.xlsx")) == FALSE) {
    data <- as.data.frame(readData(dataName, fileEncoding))
    tableTime <- c("ColName", "Change the colName")
    tableNumeric <- c("ColName", "Change the colName", rep("", 6))
    tableFactor <- c("ColName", "Change the colName", rep("", 7))

    for (i in colnames(data)) {
      table_Time <- mkTimeTable(data, i, tableTime, leastNumOfDate)
      if (!is.null(table_Time)) {
        tableTime <- table_Time
        next ()
      }
      tableNum_Fac <- mkTableNum_Fac(data, i, numOrFac, tableNumeric, tableFactor)
      tableNumeric <- tableNum_Fac$num
      tableFactor <- tableNum_Fac$fac
    }
    writeTablesOnExcel(tableNumeric, tableFactor, tableTime, dataName)
  }
  else {
    data <- as.data.frame(readData(dataName, fileEncoding))
    dataList <- NULL
    sheetList <- c("numeric", "factor", "Date")
    for (i in seq_len(length(sheetList))) {
      dataList[[i]] <- openxlsx::read.xlsx(paste0("dataCleansingForm_", dataName, "_.xlsx"), sheet = sheetList[i], colNames = F, skipEmptyRows = FALSE, skipEmptyCols = FALSE, na.strings = c("NA", ""))
    }
    for (i in colnames(data)) {
      if (!is.na(any(dataList[[1]][, 1] == i))) {
        data <- cleansNumeric(data, i, dataList[[1]], append)
      }
      if (!is.na(any(dataList[[2]][, 1] == i))) {
        data <- cleansFactor(data, i, dataList[[2]], append)
      }
      if (!is.na(any(dataList[[3]][, 1] == i))) {
        data <- cleansDate(data, i, dataList[[3]])
      }
    }
    return(data)
  }
}
#'
#' The class dataCleansing have several methods to operate the function dataCleanser.
#' @encoding UTF-8
#' @import methods
#' @field cleansingForm A form for data-cleansing stored into each types and each colnames.
#' @field  dataset A dataset you want to organise.
#'
#' @export
setRefClass(
  Class = "dataCleansing",

  fields = list(
    cleansingForm = "list",
    dataset = "data.frame",
    fileInfo = "list"
  ),

  methods = list(
    initialize = function() {
      cleansingForm <<- list(numeric = NULL, factor = NULL, Date = NULL)
      fileInfo <<- list(name = NULL, data = NULL)
    },

    mkNumericTable = function(data, index) {
      table <- c(index, rep("", 7))
      table <- rbind(table, c("", "Missing values", "", "Replace the column B with the spesific numbers", "", "Breaks", "", "Labels"))
      options(warn = -1)
      if (length(unique(data[is.na(as.numeric(data[, index])), index])) > 0) {
        numericData <- cbind(rep("", length(unique(data[is.na(as.numeric(data[, index])), index]))),
                             unique(data[is.na(as.numeric(data[, index])), index]),
                             rep("", length(unique(data[is.na(as.numeric(data[, index])), index]))),
                             rep("", length(unique(data[is.na(as.numeric(data[, index])), index]))),
                             rep("", length(unique(data[is.na(as.numeric(data[, index])), index]))),
                             rep("", length(unique(data[is.na(as.numeric(data[, index])), index]))),
                             rep("", length(unique(data[is.na(as.numeric(data[, index])), index]))),
                             rep("", length(unique(data[is.na(as.numeric(data[, index])), index]))))
      }
      else{
        numericData <- NULL
      }
      options(warn = 0)
      table <- rbind(table, numericData)
      table <- rbind(table, rep("", 8))
      return(table)
    },

    mkFactorTable = function(data, index) {
      table <- c(index, rep("", 8))
      table <- rbind(table, c("", "", "Levels", "", "Replace the column C with", "", "Pool the column C", "", "The Order of levels"))
      factorData <- cbind(rep("", nlevels(as.factor(data[, index]))),
                          paste0("No.", 1:nlevels(as.factor(data[, index]))), levels(as.factor(data[, index])),
                          rep("", nlevels(as.factor(data[, index]))),
                          rep("", nlevels(as.factor(data[, index]))),
                          rep("", nlevels(as.factor(data[, index]))),
                          rep("", nlevels(as.factor(data[, index]))),
                          rep("", nlevels(as.factor(data[, index]))),
                          rep("", nlevels(as.factor(data[, index]))))

      table <- rbind(table, factorData)
      table <- rbind(table, t(c(rep("", 9))))
      return(table)
    },

    formatOrder = function(formatStr) {
      order <- NULL
      for (i in formatStr) {
        if (i == "Y") {
          order <- append(order, 1)
        }
        else if (i == "m") {
          order <- append(order, 2)
        }
        else {
          order <- append(order, 3)
        }
      }
      return(order)
    },

    dateClassifier = function(data, index) {
      DateFormats <- list(c("Y", "m", "d"), c("Y", "d", "m"), c("m", "Y", "d"), c("m", "d", "Y"), c("d", "Y", "m"), c("d", "m", "Y"))
      formatAndDate <- NULL
      delimiter <- lapply(strsplit(data[, index], ""), function(x) {
        return(x[grep("[^0-9]", x)])
      })
      indexDelim <- lapply(strsplit(data[, index], ""), function(x) {
        dateList <- list()
        delStr <- c(0, grep("[^0-9]", x))
        for (i in seq_len(length(delStr))) {
          if (i == length(delStr) & length(delStr) < 4) {
            dateList[[i]] <- (delStr[i] + 1) : length(x)
          }
          else if (i < length(delStr)) {
            dateList[[i]] <- (delStr[i] + 1) : (delStr[i + 1] - 1)
          }
        }
        return(list(index = dateList, delim = x[grep("[^0-9]", x)], indexDelim = grep("[^0-9]", x), data = x))
      })
      arrangedIndex <- lapply(indexDelim, function(x) {
        index <- NULL
        for (i in seq_len(2)) {
          index <- append(index, c(x$index[[order(x$delim)[i]]], x$indexDelim[order(x$delim)[i]]))
        }
        if (length(x$delim) == 2) {
          index <- append(index, x$index[[3]])
        }
        else {
          index <- append(index, x$index[[order(x$delim[3])]])
        }
        return(list(index = index, data = x$data))
      })
      data[, index] <- unlist(lapply(arrangedIndex, function(x) {
        return(paste(x$data[x$index], collapse = ""))
      }))

      for (formatStr in DateFormats) {
        format <- paste0("%", formatStr[1], "-%", formatStr[2], "-%", formatStr[3])
        options(warn = -1)

        dateData <- replace(data[, index], data[, index] == "", NA)
        dateData <- gsub("[^0-9]", "-", as.character(dateData))
        dateData <- replace(dateData, nchar(gsub("[^-]", "", dateData)) != 2 | lapply(strsplit(dateData, "-"), length) != 3, NA)
        date_dataFrame <- as.data.frame(strsplit(dateData, "-"), stringsAsFactors = FALSE)
        YmdOrder <- c(match("Y", formatStr), match("m", formatStr), match("d", formatStr))
        charData <- paste0(formatC(as.numeric(date_dataFrame[YmdOrder[1], ]), width = 4, flag = "0"), "-",
                           formatC(as.numeric(date_dataFrame[YmdOrder[2], ]), width = 2, flag = "0"), "-",
                           formatC(as.numeric(date_dataFrame[YmdOrder[3], ]), width = 2, flag = "0"))
        options(warn = 0)
        formatAndDate[[length(formatAndDate) + 1]] <- list(format = format, date = charData)
      }
      return(formatAndDate)
    },

    changeColName = function(data, colname, refData, rowNumber) {
      changedName <- NULL
      if (!is.na(refData[rowNumber, 2])) {
        changedName <- refData[rowNumber, 2]
      }
      else {
        changedName <- colname
      }
      return(changedName)
    },

    replaceMissVal = function(data, refData, rowNumber) {
      options(warn = -1)
      missVal <- unique(data[is.na(as.numeric(data))])
      options(warn = 0)
      for (j in seq_len(length(missVal))) {
        if (!is.na(refData[rowNumber + 1 + j, 4])) {
          data <- replace(data, data == missVal[j], refData[rowNumber + 1 + j, 4])
        }
      }
      return(data)
    },

    cutting = function(data, refData, rowNumber) {
      if (!is.na(refData[rowNumber + 2,  8])) {
        labels <- strsplit(refData[rowNumber + 2, 8], ",")
        breaks <- strsplit(refData[rowNumber + 2, 6], ",")
        data <- cut(data, breaks = as.numeric(breaks[[1]]), labels = labels[[1]], right = FALSE)
      }
      else {
        breaks <- strsplit(refData[rowNumber + 2, 6], ",")
        data <- cut(data, breaks = as.numeric(breaks[[1]]), right = FALSE)
      }
      return(data)
    },

    orderer = function(data, refData, rowNumber) {
      orderVector <- rep(NA, nlevels(as.factor(data)))
      factorLength <- 0
      rowIndex <- 1
      while (!is.na(refData[rowNumber + 1 + rowIndex, 2])) {
        factorLength <- factorLength + 1
        rowIndex <- rowIndex + 1
      }
      pooledFactor <- strsplit(refData[(rowNumber + 2):(rowNumber + 1 + factorLength), 7], "[+]")
      poolLevel <- NULL
      for (i in seq_len(length(pooledFactor))) {
        refDataPool <- NULL
        refDataLevels <- NULL
        if (length(pooledFactor[[i]]) > 1) {
          for (j in seq_len(length(pooledFactor[[i]]))) {
            if (!is.na(refData[rowNumber + 1 + as.numeric(pooledFactor[[i]][j]), 5])) {
              refDataPool <- paste0(refDataPool, "+",  refData[rowNumber + 1 + as.numeric(pooledFactor[[i]][j]), 5])
            }
            else {
              refDataPool <- paste0(refDataPool, "+",  refData[rowNumber + 1 + as.numeric(pooledFactor[[i]][j]), 3])
            }
            refDataLevels <- append(refDataLevels, as.numeric(pooledFactor[[i]][j]))
          }
          refDataPool <- substr(refDataPool, 2, nchar(refDataPool))
          poolLevel[[i]] <- list(pool = refDataPool, levels = refDataLevels)
        }
      }
      for (i in seq_len(length(poolLevel))) {
        refData[rowNumber + 1 + poolLevel[[i]]$levels, 5] <- poolLevel[[i]]$pool
      }
      for (j in seq_len(factorLength)) {
        if (!is.na(refData[rowNumber + 1 + j, 9])) {
          orderNum <- as.numeric(refData[rowNumber + 1 + j, 9])
          if (!is.na(refData[rowNumber + 1 + j, 5])) {
            if (refData[rowNumber + 1 + j, 5] != "N/A") {
              orderVector[orderNum] <- refData[rowNumber + 1 + j, 5]
            }
          }
          else {
            orderVector[orderNum] <- refData[rowNumber + 1 + j, 3]
          }
        }
      }
      remainLevels <- setdiff(levels(as.factor(data)), orderVector)
      count <- 1
      for (i in seq_len(length(orderVector))) {
        if (is.na(orderVector[i])) {
          orderVector[i] <- remainLevels[count]
          count <- count + 1
        }
      }
      return(factor(data, levels = orderVector))
    },

    pooledName = function(data, refData, rowNumber) {
      pooledData <- pooler(data, refData, rowNumber)
      pooledFactor <- ""
      for (j in grep("[+]", levels(pooledData))) {
        pooledFactor <- paste0(pooledFactor, ".", levels(pooledData)[j])
      }
      return(substr(pooledFactor, 2, nchar(pooledFactor)))
    },

    pooler = function(data, refData, rowNumber) {
      poolList <- strsplit(refData[(rowNumber + 2) : (rowNumber + 1 + nlevels(as.factor(data))), 7], "[+]")
      data <- as.character(data)
      for (j in seq_len(length(poolList))) {
        if (length(poolList[[j]]) > 1) {
          poolLevel <- ""
          for (k in seq_len(length(poolList[[j]]))) {
            if (!is.na(refData[rowNumber + 1 + as.numeric(poolList[[j]][k]), 5])) {
              poolLevel <- paste0(poolLevel, "+",  refData[rowNumber + 1 + as.numeric(poolList[[j]][k]), 5])
            }
            else {
              poolLevel <- paste0(poolLevel, "+",  refData[rowNumber + 1 + as.numeric(poolList[[j]][k]), 3])
            }
          }
          poolLevel <- substr(poolLevel, 2, nchar(poolLevel))
          for (k in seq_len(length(poolList[[j]]))) {
            if (!is.na(refData[rowNumber + 1 + as.numeric(poolList[[j]][k]), 5])) {
              data <- replace(data, data == refData[rowNumber + 1 + as.numeric(poolList[[j]][k]), 5], poolLevel)
            }
            else {
              data <- replace(data, data == refData[rowNumber + 1 + as.numeric(poolList[[j]][k]), 3], poolLevel)
            }
          }
        }
      }
      levels <- levels(as.factor(data))
      return(factor(data, levels = levels))
    },

    replacer = function(data, refData, rowNumber) {
      levelsRow <- refData[(rowNumber + 2) : (rowNumber + 1 + nlevels(as.factor(data))), 5] != ""
      factorLength <- 0
      rowIndex <- 1
      while (!is.na(refData[rowNumber + 1 + rowIndex, 2])) {
        factorLength <- factorLength + 1
        rowIndex <- rowIndex + 1
      }
      data <- as.character(data)
      refData[(rowNumber + 2) : (rowNumber + 1 + factorLength), 3] <- replace(refData[(rowNumber + 2) : (rowNumber + 1 + factorLength), 3], is.na(refData[(rowNumber + 2) : (rowNumber + 1 + factorLength), 3]), "")
      for (j in seq_len(nlevels(as.factor(data)))) {
        if (!is.na(levelsRow[j])) {
          if (refData[rowNumber + 1 + j, 5] == "N/A") {
            data <- replace(data, data == refData[rowNumber + 1 + j, 3], NA)
          }
          else {
            data <- replace(data, data == refData[rowNumber + 1 + j, 3], refData[rowNumber + 1 + j, 5])
          }
        }
      }
      levels <- levels(as.factor(data))
      return(factor(data, levels = levels))
    },

    readData = function(dataName, fileEncoding) {
      if (fileEncoding == "UTF-8" | fileEncoding == "Latin-1") {
        data <- data.table::fread(paste0(dataName, ".csv"), encoding = fileEncoding)
      }
      else {
        data <- data.table::fread(paste0(dataName, ".csv"), encoding = "unknown")
      }
      return(data)
    },

    mkTimeTable = function(data, index, tableTime, leastNumOfDate) {
      asDatedVector <- rep(FALSE, nrow(data))
      haveNumber <- replace(rep(FALSE, nrow(data)), grep("[0-9]", data[, index]), TRUE)
      haveLeastNumStr <- nchar(data[, index]) > 4
      delimiter <- nchar(gsub("[0-9]", "", data[, index])) == 2 | nchar(gsub("[0-9]", "", data[, index])) == 3
      if (length(data[haveNumber & haveLeastNumStr & delimiter, index]) <= leastNumOfDate) {
        return(NULL)
      }
      formCharData <- dateClassifier(data, index)
      for (i in list(c(1, 2), c(3, 5), c(4, 6))) {
        isFirstFormat <- !is.na(isCorrectFormat(formCharData[[i[1]]])) & isCorrectFormat(formCharData[[i[1]]]) & !asDatedVector
        isSecondFormat <- !is.na(isCorrectFormat(formCharData[[i[2]]])) & isCorrectFormat(formCharData[[i[2]]]) & !asDatedVector
        if (length(data[isFirstFormat, index]) > 0 | length(data[isSecondFormat, index]) > 0) {
          if (length(data[isFirstFormat, index]) > length(data[isSecondFormat, index])) {
            asDatedVector[isFirstFormat] <- TRUE
          }
          else {
            asDatedVector[isSecondFormat] <- TRUE
          }
        }
      }
      if (length(asDatedVector[asDatedVector == TRUE]) > leastNumOfDate) {
        tableTime <- rbind(tableTime, c(index, ""))
        cleansingForm$Date[[length(cleansingForm$Date) + 1]] <<- list(table = as.data.frame(t(c(index, "")), stringsAsFactors = FALSE), colname = index)
        return(tableTime)
      }
      return(NULL)
    },

    writeTablesOnExcel = function(tableNumeric, tableFactor, tableTime, dataName, filePath = "") {
      sheetNumeric <- list(table = tableNumeric, sheetName = "numeric")
      sheetFactor <- list(table = tableFactor, sheetName = "factor")
      sheetTime <- list(table = tableTime, sheetName = "Date")
      wb <- openxlsx::createWorkbook()
      for (i in list(sheetNumeric, sheetFactor, sheetTime)) {
        openxlsx::addWorksheet(wb, i$sheetName)
      }
      st <- openxlsx::createStyle(fontName = "Yu Gothic", fontSize = 11)
      for (i in list(sheetNumeric, sheetFactor, sheetTime)) {
        openxlsx::addStyle(wb, i$sheetName, style = st, cols = 1:2, rows = 1:2)
        openxlsx::writeData(wb, sheet = i$sheetName, x = i$table, colNames = F, withFilter = F)
      }
      openxlsx::modifyBaseFont(wb, fontSize = 11, fontColour = "#000000", fontName = "Yu Gothic")
      openxlsx::saveWorkbook(wb, paste0(filePath, "dataCleansingForm_", dataName, "_.xlsx"), overwrite = TRUE)
    },

    mkTableNum_Fac = function(data, index, numOrFac, tableNumeric, tableFactor) {
      options(warn = -1)
      charEqualNum <-  as.character(as.numeric(data[, index])) == as.numeric(data[, index])
      options(warn = 0)
      if (length(na.omit(data[charEqualNum == FALSE, index])) == 0 &  length(na.omit(data[charEqualNum == TRUE, index])) > 0 & nlevels(as.factor(data[, index])) > nrow(data) / numOrFac) {
        numTab <- mkNumericTable(data, index)
        tableNumeric <- rbind(tableNumeric, numTab)
        if (!is.null(numTab)) {
          cleansingForm$numeric[[length(cleansingForm$numeric) + 1]] <<- list(table = as.data.frame(numTab, stringsAsFactors = FALSE), colname = index)
        }
      }
      else {
        facTab <- mkFactorTable(data, index)
        tableFactor <- rbind(tableFactor, facTab)
        if (!is.null(facTab)) {
          cleansingForm$factor[[length(cleansingForm$factor) + 1]] <<- list(table = as.data.frame(facTab, stringsAsFactors = FALSE), colname = index)
        }
      }
      return(list(num = tableNumeric, fac = tableFactor))
    },

    cleansNumeric = function(data, index, refData, append) {
      if (any(refData[, 1] == index)) {
        options(warn = -1)
        rowNumber <- as.numeric(rownames(refData[refData[, 1] == index & !is.na(refData[, 1]), ]))
        colnames(data)[colnames(data) == index] <- changeColName(data, index, refData, rowNumber)
        index <- changeColName(data, index, refData, rowNumber)
        if (any(!is.na(refData[(rowNumber + 2):(rowNumber + 1 + length(unique(data[is.na(as.numeric(data[, index])), index]))), 4])) & length(unique(data[is.na(as.numeric(data[, index])), index])) > 0) {
          if (append == TRUE) {
            data <- cbind(data, replaceMissVal(data[, index], refData, rowNumber))
            colnames(data)[ncol(data)] <- paste0(index, "_missing Values replaced")
          }
          else {
            data[, index] <- replaceMissVal(data[, index], refData, rowNumber)
          }
        }
        data[, index] <- as.numeric(data[, index])
        options(warn = 0)
        if (any(!is.na(refData[rowNumber + 2, 6]))) {
          if (append == TRUE) {
            data <- cbind(data, cutting(data[, index], refData, rowNumber))
            colnames(data)[ncol(data)] <- paste0(index, "_categorized")
          }
          else {
            data[, index] <- cutting(data[, index], refData, rowNumber)
          }
        }
      }
      return(data)
    },

    cleansFactor = function(data, index, refData, append) {
      if (any(refData[, 1] == index)) {
        pooling <- FALSE
        ordering <- FALSE
        options(warn = -1)
        rowNumber <- as.numeric(rownames(refData[refData[, 1] == index & !is.na(refData[, 1]), ]))
        options(warn = 0)
        colnames(data)[colnames(data) == index] <- changeColName(data, index, refData, rowNumber)
        index <- changeColName(data, index, refData, rowNumber)
        nrowLevel <- nlevels(as.factor(data[, index]))
        if (any(!is.na(refData[(rowNumber + 2):(rowNumber + 1 +  nrowLevel), 5]))) { #5 for replace
          data[, index] <- replacer(data[, index], refData, rowNumber)
          if (any(!is.na(refData[(rowNumber + 2):(rowNumber + 1 + nlevels(as.factor(data[, index]))), 7]))) { #7 for pool
            if (append == TRUE) {
              data <- cbind(data, pooler(data[, index], refData, rowNumber))
              colnames(data)[ncol(data)] <- paste0(index, "_", pooledName(data[, index], refData, rowNumber))
              index <- colnames(data)[ncol(data)]
            }
            else {
              data[, index] <- pooler(data[, index], refData, rowNumber)
            }
            pooling <- TRUE
            if (any(!is.na(refData[(rowNumber + 2):(rowNumber + 1 +  nrowLevel), 9]))) { #9 for order
              data[, index] <- orderer(data[, index], refData, rowNumber)
              ordering <- TRUE
            }
          }
        }
        if (any(!is.na(refData[(rowNumber + 2):(rowNumber + 1 + nlevels(as.factor(data[, index]))), 7])) & pooling == FALSE) {
          if (append == TRUE) {
            data <- cbind(data, pooler(data[, index], refData, rowNumber))
            colnames(data)[ncol(data)] <- paste0(index, "_", pooledName(data[, index], refData, rowNumber))
          }
          else {
            data[, index] <- pooler(data[, index], refData, rowNumber)
          }
          pooling <- TRUE
          if (any(!is.na(refData[(rowNumber + 2):(rowNumber + 1 + nlevels(as.factor(data[, index]))), 9])) & ordering == FALSE) {
            data[, index] <- orderer(data[, index], refData, rowNumber)
            ordering <- TRUE
          }
        }
        if (any(!is.na(refData[(rowNumber + 2):(rowNumber + 1 + nlevels(as.factor(data[, index]))), 9])) & ordering == FALSE) {
          data[, index] <- orderer(data[, index], refData, rowNumber)
          ordering <- TRUE
        }
        data[, index] <- as.factor(data[, index])
      }
      return(data)
    },

    cleansDate = function(data, index, refData) {
      if (any(refData[, 1] == index)) {
        asDatedVector <- rep(FALSE, nrow(data))
        options(warn = -1)
        rowNumber <- as.numeric(rownames(refData[refData[, 1] == index & !is.na(refData[, 1]), ]))
        colnames(data)[colnames(data) == index] <- changeColName(data, index, refData, rowNumber)
        index <- changeColName(data, index, refData, rowNumber)
        formCharData <- dateClassifier(data, index)
        for (i in list(c(1, 2), c(3, 5), c(4, 6))) {
          isFirstFormat <- !is.na(isCorrectFormat(formCharData[[i[1]]])) & isCorrectFormat(formCharData[[i[1]]]) & !asDatedVector
          isSecondFormat <- !is.na(isCorrectFormat(formCharData[[i[2]]])) & isCorrectFormat(formCharData[[i[2]]]) & !asDatedVector
          if (length(data[isFirstFormat, index]) > 0 | length(data[isSecondFormat, index]) > 0) {
            if (length(data[isFirstFormat, index]) > length(data[isSecondFormat, index])) {
              data[isFirstFormat, index] <- formCharData[[i[1]]]$date[isFirstFormat]
              asDatedVector[isFirstFormat] <- TRUE
            }
            else {
              data[isSecondFormat, index] <- formCharData[[i[2]]]$date[isSecondFormat]
              asDatedVector[isSecondFormat] <- TRUE
            }
          }
        }
        options(warn = 0)
      }
      return(data)
    },

    isCorrectFormat = function(dateAsFormat) {
      splitedFCD <- strsplit(dateAsFormat$date, "-")
      options(warn = -1)
      correctDate <- unlist(lapply(splitedFCD, function(x) {
        conditionFormat <- nchar(x[1]) == 4 & nchar(x[2]) == 2 & nchar(x[3]) == 2
        conditionUpper <- as.numeric(x[2]) < 13 & as.numeric(x[3] < 32)
        if (conditionFormat & conditionUpper) {
          return(TRUE)
        }
        else {
          return(FALSE)
        }
      }))
      options(warn = 0)
      return(correctDate)
    },

    mkCleansingForm = function(dataName, dataPath, numOrFac = 10, leastNumOfDate = 10, fileEncoding = "CP932", filePath) {
      data <- as.data.frame(readData(dataPath, fileEncoding), stringsAsFactors = FALSE)
      dataset <<- data
      fileInfo$data <<- data
      if (fileEncoding == "Others") {
        write.csv(data, paste0(filePath, dataName, ".csv"), row.names = FALSE, fileEncoding = "CP932")
      }
      else {
        write.csv(data, paste0(filePath, dataName, ".csv"), row.names = FALSE, fileEncoding = fileEncoding)
      }
      tableTime <- c("ColName", "Change the colName")
      tableNumeric <- c("ColName", "Change the colName", rep("", 6))
      tableFactor <- c("ColName", "Change the colName", rep("", 7))
      for (i in colnames(data)) {
        table_Time <- mkTimeTable(data, i, tableTime, leastNumOfDate)
        if (!is.null(table_Time)) {
          tableTime <- table_Time
          next ()
        }
        tableNum_Fac <- mkTableNum_Fac(data, i, numOrFac, tableNumeric, tableFactor)
        tableNumeric <- tableNum_Fac$num
        tableFactor <- tableNum_Fac$fac
      }
      writeTablesOnExcel(tableNumeric, tableFactor, tableTime, dataName, filePath)
    },

    dataCleanser = function(dataName, append = FALSE, numOrFac = 10, leastNumOfDate = 10, fileEncoding = "CP932", path = "") {
      files <- list.files(path = path)
      if (any(files == paste0("dataCleansingForm_", dataName, "_.xlsx")) == FALSE) {
        data <- as.data.frame(readData(paste0(path, dataName), fileEncoding), stringsAsFactors = FALSE)
        dataset <<- data
        tableTime <- c("ColName", "Change the colName")
        tableNumeric <- c("ColName", "Change the colName", rep("", 6))
        tableFactor <- c("ColName", "Change the colName", rep("", 7))

        for (i in colnames(data)) {
          table_Time <- mkTimeTable(data, i, tableTime, leastNumOfDate)
          if (!is.null(table_Time)) {
            tableTime <- table_Time
            next ()
          }
          tableNum_Fac <- mkTableNum_Fac(data, i, numOrFac, tableNumeric, tableFactor)
          tableNumeric <- tableNum_Fac$num
          tableFactor <- tableNum_Fac$fac
        }
        writeTablesOnExcel(tableNumeric, tableFactor, tableTime, dataName, filePath = path)
      }
      else {
        data <- as.data.frame(readData(paste0(path, dataName), fileEncoding))
        dataList <- NULL
        sheetList <- c("numeric", "factor", "Date")
        for (i in seq_len(length(sheetList))) {
          dataList[[i]] <- openxlsx::read.xlsx(paste0(path, "dataCleansingForm_", dataName, "_.xlsx"), sheet = sheetList[i], colNames = F, skipEmptyRows = FALSE, skipEmptyCols = FALSE, na.strings = c("NA", ""))
        }
        for (i in colnames(data)) {
          if (!is.na(any(dataList[[1]][, 1] == i))) {
            data <- cleansNumeric(data, i, dataList[[1]], append)
          }
          if (!is.na(any(dataList[[2]][, 1] == i))) {
            data <- cleansFactor(data, i, dataList[[2]], append)
          }
          if (!is.na(any(dataList[[3]][, 1] == i))) {
            data <- cleansDate(data, i, dataList[[3]])
          }
        }
        return(data)
      }
    }
  )
)

#' Merging rownames and colnames with a data.frame type object. This function reshapes a dataframe to the form having rownames and colnames as components of a dataframe not as its rownames and colnames.
#'
#' @param x A vector or dataframe type object you want to merge its rownames and colnames with.
#'
#' @export
mergeRowAndColnamesWithData <- function(x) {
  dataFrame <- as.data.frame(x)
  colnames(dataFrame) <- rep("", ncol(dataFrame))
  rownames(dataFrame) <- seq_len(nrow(dataFrame))
  for (i in seq_len(ncol(dataFrame))) {
    if (is.factor(dataFrame[, i])) {
      dataFrame[, i] <- as.character(dataFrame[, i])
    }
  }
  if (!is.null(rownames(x))) {
    rownamesDataFrame <- as.data.frame(rownames(x))
    colnames(rownamesDataFrame) <- rep("", ncol(rownamesDataFrame))
    dataFrame <- cbind(rownamesDataFrame, dataFrame)
  }
  colnames(dataFrame) <- rep("", ncol(dataFrame))
  if (!is.null(colnames(x))) {
    colnamesDataFrame <- as.data.frame(t(colnames(x)))
    if (ncol(x) != ncol(dataFrame)) {
      colnamesDataFrame <- cbind("", colnamesDataFrame)
    }
    colnames(colnamesDataFrame) <- rep("", ncol(colnamesDataFrame))
    dataFrame <- rbind(colnamesDataFrame, dataFrame)
  }
  return(dataFrame)
}

#' Merging two objects (data.frame, vector) in a vertical direction without adjusting the each number of rows or columns.
#' @encoding UTF-8
#'
#' @param x Data.frame type object or vector type object you want to merge with y.
#' @param y Data.frame type object or vector type object merged with x.
#' @param rowNames wheter to include rownames of x and y.
#' @param colNames wheter to include colnames of x and y.
#' @param sep Whether to separate x and y with a empty row.
#'
#' @export
rowBind <- function(x, y, rowNames = TRUE, colNames = TRUE, sep = TRUE) {
  dataFrameX <- mergeRowAndColnamesWithData(x)
  dataFrameY <- mergeRowAndColnamesWithData(y)
  if (!rowNames) {
    dataFrameX <- as.data.frame(dataFrameX[-1, ])
    rownames(dataFrameX) <- seq_len(nrow(dataFrameX))
    dataFrameY <- as.data.frame(dataFrameY[-1, ])
    rownames(dataFrameY) <- seq_len(nrow(dataFrameY))
  }
  if (!colNames) {
    dataFrameX <- as.data.frame(dataFrameX[, -1])
    colnames(dataFrameX) <- rep("", ncol(dataFrameX))
    dataFrameY <- as.data.frame(dataFrameY[, -1])
    colnames(dataFrameY) <- rep("", ncol(dataFrameY))
  }
  diffOfNCol <- ncol(dataFrameX) - ncol(dataFrameY)
  if (diffOfNCol > 0) {
    dataFrameY <- cbind(dataFrameY, matrix(rep("", nrow(dataFrameY) * abs(diffOfNCol)), nrow = nrow(dataFrameY)))
  }
  else {
    dataFrameX <- cbind(dataFrameX, matrix(rep("", nrow(dataFrameX) * abs(diffOfNCol)), nrow = nrow(dataFrameX)))
  }
  colnames(dataFrameX) <- rep("", ncol(dataFrameX))
  colnames(dataFrameY) <- rep("", ncol(dataFrameY))
  bindedDataFrame <- NULL
  if (sep) {
    bindedDataFrame <- rbind(dataFrameX, rep("", ncol(dataFrameX)), dataFrameY)
  }
  else {
    bindedDataFrame <- rbind(dataFrameX, dataFrameY)
  }
  colnames(bindedDataFrame) <- rep("", ncol(bindedDataFrame))
  return(bindedDataFrame)
}
#'
#' Merging two objects (data.frame, vector) in a horizontal direction without adjusting the each number of rows or columns.
#' @encoding UTF-8
#'
#' @param x Data.frame type object or vector type object you want to merge with y.
#' @param y Data.frame type object or vector type object merged with x.
#' @param rowNames wheter to include rownames of x and y.
#' @param colNames wheter to include colnames of x and y.
#' @param sep Whether to separate x and y with a empty column.
#'
#' @export
colBind <- function(x, y, rowNames = TRUE, colNames = TRUE, sep = TRUE) {
  dataFrameX <- mergeRowAndColnamesWithData(x)
  dataFrameY <- mergeRowAndColnamesWithData(y)
  if (!rowNames) {
    dataFrameX <- as.data.frame(dataFrameX[-1, ])
    rownames(dataFrameX) <- seq_len(nrow(dataFrameX))
    dataFrameY <- as.data.frame(dataFrameY[-1, ])
    rownames(dataFrameY) <- seq_len(nrow(dataFrameY))
  }
  if (!colNames) {
    dataFrameX <- as.data.frame(dataFrameX[, -1])
    colnames(dataFrameX) <- rep("", ncol(dataFrameX))
    dataFrameY <- as.data.frame(dataFrameY[, -1])
    colnames(dataFrameY) <- rep("", ncol(dataFrameY))
  }
  diffOfNRow <- nrow(dataFrameX) - nrow(dataFrameY)
  if (diffOfNRow > 0) {
    emptyDataFrame <- as.data.frame(matrix(rep("", ncol(dataFrameY) * abs(diffOfNRow)), ncol = ncol(dataFrameY)))
    colnames(emptyDataFrame) <- rep("", ncol(emptyDataFrame))
    dataFrameY <- rbind(dataFrameY, emptyDataFrame)
  }
  else {
    emptyDataFrame <- as.data.frame(matrix(rep("", ncol(dataFrameX) * abs(diffOfNRow)), ncol = ncol(dataFrameX)))
    colnames(emptyDataFrame) <- rep("", ncol(emptyDataFrame))
    dataFrameX <- rbind(dataFrameX, emptyDataFrame)
  }
  bindedDataFrame <- NULL
  if (sep) {
    bindedDataFrame <- cbind(dataFrameX, rep("", nrow(dataFrameX)), dataFrameY)
  }
  else {
    bindedDataFrame <- cbind(dataFrameX, dataFrameY)
  }
  colnames(bindedDataFrame) <- rep("", ncol(bindedDataFrame))
  return(bindedDataFrame)
}

#'
#' Getting the indices of some item in a vector or a list.
#' @encoding UTF-8
#'
#' @param x A vector or list type object.
#' @param item The item in x indexed.
#'
#' @export
getIndex <- function(x, item) {
  if (!is.vector(x)) {
    stop("The Argument x must be a vector or list type object")
  }
  if (is.list(x)) {
    isItem <- unlist(lapply(x, function(y) {
      return(all(y == item))
    }))
    return(seq_len(length(isItem))[isItem])
  }
  else {
    if (length(item) > 1) {
      stop("The length of the argument item must be 1")
    }
    names(x) <- seq_len(length(x))
    return(as.numeric(names(x[x == item])))
  }
}
#'
#' Getting the number of a component in a vector or a list.
#' @encoding UTF-8
#'
#' @param x A vector or list type object.
#' @param item A component in x.
#'
#' @export
getCount <- function(x, item) {
  return(length(getIndex(x, item)))
}
#'
#' Appending two vector or list type objects.
#' @encoding UTF-8
#'
#' @param x A vector or list type object.
#' @param item A vector or list type object appended to x.
#'
#' @export
append_vecOrList <- function(x, item) {
  if (!is.vector(x)) {
    stop("The Argument x must be a vector or list type object")
  }
  if (is.list(x) & !is.list(item)) {
    item <- list(item)
  }
  return(c(x, item))
}
#'
#' Merging two vector or list type objects.
#' @encoding UTF-8
#'
#' @param x A vector or list type object.
#' @param item A vector or list type object merged with x.
#'
#' @export
extend_vecOrList <- function(x, item) {
  if (!is.vector(x)) {
    stop("The Argument x must be a vector or list type object")
  }
  return(c(x, item))
}
#'
#' Inserting an item into a vector or list type object.
#' @encoding UTF-8
#'
#' @param x A vector or list type object.
#' @param i The index which an item inserted at.
#' @param item A vector or list type object merged with x.
#'
#' @export
insert_vecOrList <- function(x, i, item) {
  if (!is.vector(x)) {
    stop("The Argument x must be a vector or list type object")
  }
  if (is.list(x) & !is.list(item)) {
    item <- list(item)
  }
  return(c(x[seq_len(i - 1)], item, x[i : length(x)]))
}
#'
#' Removing an item in a vector or list type object.
#' @encoding UTF-8
#'
#' @param x A vector or list type object.
#' @param item An item removed from x.
#'
#' @export
remove_vecOrList <- function(x, item) {
  indices <- getIndex(x, item)
  return(x[- indices])
}
#'
#' Popping an item from a vector or list type object.
#' @encoding UTF-8
#'
#' @param x A vector or list type object.
#' @param i The index of a component in x popped.
#'
#' @export
pop_vecOrList <- function(x, i) {
  if (!is.vector(x)) {
    stop("The Argument x must be a vector or list type object")
  }
  if (i < 1 | i > length(x)) {
    stop("The Argument i must greater than 1 and less than the length of x")
  }
  if (i + 1 <= length(x)) {
    poppedX <- c(x[seq_len(i - 1)], x[(i + 1) : length(x)])
  }
  else {
    poppedX <- x[seq_len(i - 1)]
  }
  return(list(poppedX = poppedX, xi = x[i]))
}
#'
#' Compare two data in txt, csv or xlsx files.
#' @encoding UTF-8
#'
#' @param fileName1 The name of file in which one of data you want to compare is saved.
#' @param fileName2 The name of file in which another data you want to compare is saved.
#' @param fileEncoding1 File-encoding for fileName1.
#' @param fileEncoding2 File-encoding for fileName2.
#'
#' @importFrom utils read.csv
#' @importFrom utils read.delim
#'
#' @export
isEqualData <- function(fileName1, fileName2, fileEncoding1 = "CP932", fileEncoding2 = "CP932") {
  splitedName1 <- unlist(strsplit(fileName1, "\\."))
  splitedName2 <- unlist(strsplit(fileName2, "\\."))
  extension1 <- splitedName1[length(splitedName1)]
  extension2 <- splitedName2[length(splitedName2)]
  extensions <- tolower(c(extension1, extension2))
  fileNames <- c(fileName1, fileName2)
  fileEncodings <- c(fileEncoding1, fileEncoding2)
  dataList <- list(data1 = NULL, data2 = NULL)
  for (i in seq_len(length(extensions))) {
    if (extensions[i] == "txt") {
      dataList[[i]] <- read.delim(fileNames[i], fill = TRUE, header = FALSE, sep = ",", blank.lines.skip = FALSE, fileEncoding = fileEncodings[i])
    }
    else if (extensions[i] == "csv") {
      dataList[[i]] <- read.csv(fileNames[i], fill = TRUE, header = FALSE, sep = ",", blank.lines.skip = FALSE, fileEncoding = fileEncodings[i])
    }
    else if (extensions[i] == "xlsx") {
      dataList[[i]] <- openxlsx::read.xlsx(fileNames[i])
    }
  }
  isEqualEachCell <- dataList$data1 == dataList$data2
  isEqualRow <- apply(isEqualEachCell, 1, all)
  isEqualCol <- apply(isEqualEachCell, 2, all)
  names(isEqualRow) <- seq_len(length(isEqualRow))
  names(isEqualCol) <- seq_len(length(isEqualCol))
  lastNotEqualIndex <- c(row = max(as.numeric(names(isEqualRow[!isEqualRow]))), col = max(as.numeric(names(isEqualCol[!isEqualCol]))))
  if (all(isEqualEachCell)) {
    writeLines("TRUE\n================================================================================")
    writeLines(paste0("Data in", fileName1, " is equal to ", "data in ", fileName2))
    writeLines("================================================================================")
  }
  else {
    writeLines("FALSE\n================================================================================")
    for (i in seq_len(nrow(isEqualEachCell))) {
      for (j in seq_len(ncol(isEqualEachCell))) {
        if (!isEqualEachCell[i, j]) {
          if (i != lastNotEqualIndex["row"] | j != lastNotEqualIndex["col"]) {
            writeLines(paste0("The ", i, ", ", j, " components are not the same:\ndata1[", i, ", ", j, "]: ",
                              dataList$data1[i, j], "\ndata2[", i, ", ", j, "]: ", dataList$data2[i, j], "\n"))
          }
          else {
            writeLines(paste0("The ", i, ", ", j, " components are not the same:\ndata1[", i, ", ", j, "]: ",
                              dataList$data1[i, j], "\ndata2[", i, ", ", j, "]: ", dataList$data2[i, j]))
          }
        }
      }
    }
    writeLines("================================================================================")
  }
}

#' Create a summary table for numeric and factor data in terms of each level of a factor data.
#' @encoding UTF-8
#'
#' @param data Dataset you want to summarize.
#' @param namesForRow Column names assigned to row names of a summary table.
#' @param nameForCol Column name assigned to col names of a summary table.
#' @param digits integer indicating the number of decimal places.
#' @param locationPar A Character variable which determine the location parameter.
#' @param sd Whether to include standard deviation for numeric data. and ratio for factor data.
#' @param Qu Whether to include 1-quantile and 2-quantile.
#' @param ratio Whether to include ratio of percentage for each cell.
#'
#' @importFrom stats sd
#' @importFrom stats median
#' @importFrom stats na.omit
#' @importFrom graphics hist
#'
#' @export
getSummaryTable <- function(data, namesForRow, nameForCol, digits = 2, locationPar = "mean", sd = FALSE, Qu = FALSE, ratio = FALSE) {
  if (length(nameForCol) > 1) {
    stop("The length of the argument nameForCol must be one")
  }
  if (is.null(data)) {
    stop("data is null")
  }
  if (!is.factor(data[, nameForCol])) {
    stop("The type of data[, nameForCol] must be factor")
  }
  table <- as.data.frame(matrix(rep(NA, nlevels(data[, nameForCol])), nrow = 1)[numeric(0), ])
  colnames(table) <- rep("", ncol(table))
  table <- rbind(table, t(c("", paste0(nameForCol, " (n=", length(na.omit(data[, nameForCol])), ")"), rep("", nlevels(data[, nameForCol]) - 1))))
  table <- rbind(table, t(c("", paste0(levels(data[, nameForCol]), " (n=", table(na.omit(data[, nameForCol])), ")"))))
  for (nameForRow in namesForRow) {
    if (is.numeric(data[, nameForRow])) {
      table <- rbind(table, getSummaryNumeric(data, nameForRow, nameForCol, digits, locationPar, sd, Qu))
    }
    else if (is.factor(data[, nameForRow])) {
      table <- rbind(table, getSummaryFactor(data, nameForRow, nameForCol, digits, ratio))
    }
  }
  return(table)
}

getSummaryNumeric <- function(data, nameForRow, nameForCol, digits, locationPar, sd, Qu) {
  tableRow <- NULL
  label <- nameForRow
  if (sd) {
    label <- paste0(label, " (sd)")
  }
  if (Qu) {
    label <- paste0(label, " [1st Qu., 3rd Qu.]")
  }
  tableRow <- label
  for (colLevel in levels(data[, nameForCol])) {
    cellValue <- NULL
    if (locationPar == "mean") {
      cellValue <- sprintf(paste0("%.", digits, "f"), mean(na.omit(data[data[, nameForCol] == colLevel, nameForRow])))
    }
    else if (locationPar == "median") {
      cellValue <- sprintf(paste0("%.", digits, "f"), median(na.omit(data[data[, nameForCol] == colLevel, nameForRow])))
    }
    else if (locationPar == "mode") {
      cellValue <- mode(na.omit(data[data[, nameForCol] == colLevel, nameForRow]), digits)
    }
    if (sd) {
      cellValue <- paste0(cellValue, " (", sprintf(paste0("%.", digits, "f"), sd(na.omit(data[data[, nameForCol] == colLevel, nameForRow]))), ")")
    }
    if (Qu) {
      summaryStats <- summary(data[data[, nameForCol] == colLevel, nameForRow])
      cellValue <- paste0(cellValue, " [", sprintf(paste0("%.", digits, "f"), summaryStats["1st Qu."]), ", ", sprintf(paste0("%.", digits, "f"), summaryStats["3rd Qu."]), "]")
    }
    tableRow <- append(tableRow, cellValue)
  }
  return(t(tableRow))
}

getSummaryFactor <- function(data, nameForRow, nameForCol, digits, ratio) {
  table <- as.data.frame(matrix(rep(NA, nlevels(data[, nameForCol])), nrow = 1)[numeric(0), ])
  colnames(table) <- rep("", ncol(table))
  tableRow <- NULL
  if (ratio) {
    tableRow <- append(tableRow, c(paste0(nameForRow, " (%)"), rep("", nlevels(data[, nameForCol]))))
  }
  else {
    tableRow <- append(tableRow, c(nameForRow, rep("", nlevels(data[, nameForCol]))))
  }
  table <- rbind(table, t(tableRow))
  for (rowLevel in levels(data[, nameForRow])) {
    tableRow <- rowLevel
    for (colLevel in levels(data[, nameForCol])) {
      data_eachCell <- na.omit(data[data[, nameForCol] == colLevel & data[, nameForRow] == rowLevel, nameForRow])
      if (ratio) {
        tableRow <- append(tableRow, paste0(length(data_eachCell), " (", sprintf(paste0("%.", digits, "f"), 100 * (length(data_eachCell) / length(data[data[, nameForCol] == colLevel, nameForRow]))), ")"))
      }
      else {
        tableRow <- append(tableRow, length(data_eachCell))
      }
    }
    table <- rbind(table, t(tableRow))
  }
  return(table)
}

mode <- function(x, digits) {
  if (length(x) < 30 | length(table(x)) < 7) {
  tableX <- table(x)
  mode <- as.numeric(names(tableX[tableX == max(tableX)]))
  mode <- sprintf(paste0("%.", digits, "f"), mode)
  }
  else {
    histX <- graphics::hist(x, plot = FALSE)$counts
    names(histX) <- paste0("[", sprintf(paste0("%.", digits, "f"), graphics::hist(x, plot = FALSE)$breaks[seq_len(length(histX))]), ", ",
                           sprintf(paste0("%.", digits, "f"), graphics::hist(x, plot = FALSE)$breaks[seq(2, length(histX) + 1)]), ")")
    mode <- names(histX[histX == max(histX)])
  }
  if (length(mode) > 1){
    mode <- NA
  }
  return(mode)
}
