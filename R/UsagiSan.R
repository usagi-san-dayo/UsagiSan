#' UsagiSan: A package for cleansing dataset and outputting statistical test results with using EXCEL.
#'
#' The package UsagiSan provides you a lot of helps to reduce the time on data-clansing and editting the test results. this package contains four function:
#' excelColor, excelHeadColor, mkDirectries and dataCleanser
#'
#' @section excelColor:
#' The function excelColor helps you to edit the test results with coloring the signigicant variables with specific color.
#'
#' @section excelHeadColor:
#' The function excelHeadColor helps you to add colors on headers of any type of tables including summaty sheets and statistical test tables.
#'
#' @section mkDirectries:
#' The function mkDirectries organizes files in tha working directory.
#'
#' @section dataCleanser:
#' The function dataCleanser helps you to cleans a dataset. This function modifies a dataset  to the form much easier to handle in statistical analysis.
#'
#' @seealso \code{\link{excelColor}}
#' @seealso \code{\link{excelHeadColor}}
#' @seealso \code{\link{mkDirectries}}
#' @seealso \code{\link{dataCleanser}}
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
  options(warn = -1)
  if (is.na(dataName)) {
    stop("There is no data-name")
  }
  options(warn = 0)
  if (is.na(fileName)) {
    stop("There is no file-name")
  }
  if (!(is.character(fileName))) {
    stop("The file-name must be character")
  }
  data <- utils::read.table(paste0(dataName, ".csv"), fill = TRUE, header = FALSE, sep = ",", blank.lines.skip = FALSE, fileEncoding = fileEncoding)
  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, "Sheet 1")
  st <- openxlsx::createStyle(fontName = fontName, fontSize = fontSize)
  openxlsx::addStyle(wb, "Sheet 1", style = st, cols = 1:2, rows = 1:2)
  openxlsx::writeData(wb, sheet = "Sheet 1", x = data, colNames = F, withFilter = F)
  openxlsx::modifyBaseFont(wb, fontSize = fontSize, fontColour = fontColor, fontName = fontName)

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

  col_count <- NULL
  col_intercept <- NULL
  table_rightLim <- NULL
  if (adj == TRUE) {
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
    #setStyle for headers
    for (i in seq_len(length(green_header))) {
      st <- openxlsx::createStyle(fontName = fontName, fontSize = fontSize, fgFill = headerColor)
      openxlsx::addStyle(wb, "Sheet 1", style = st, cols = col_intercept[i]:table_rightLim[i], rows = green_header[i])
    }
    count <- 1
    factor_list <- list(NULL)
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

    factor_row <- list(NULL)
    for (i in seq_len(length(factor_list))) {
      bar <-  factor_list[[i]]
      factor_row[[i]] <- bar[as.numeric(data[factor_list[[i]], col_count[i]]) < level]
    }

    #removing NA
    for (i in seq_len(length(factor_row))) {
      if (!(is.integer(factor_row[[i]]) & length(factor_row[[i]]) == 0L)) {
        options(warn = -1)
        if (all(is.na(factor_row[[i]]))) {
          factor_row[[i]] <- integer(0)
        }
        options(warn = 0)
      }
      if (any(is.na(factor_row[[i]]))) {
        factor_row[[i]] <- stats::na.omit(factor_row[[i]])[1]
      }
    }

    #write data
    options(warn = -1)
    for (i in seq_len(length(factor_list))) {
      bar <- data[factor_list[[i]], ]
      for (j in seq_len(ncol(data[factor_list[[i]], ]))) {
        bar[bar[, j] == "", j] <- NA
        tmp <- data[factor_list[[i]], col_intercept[i] - 1 + j]
        tmp[tmp == ""] <- NA
        if (!any((is.na(as.numeric(stats::na.omit(bar[, j])))))) {
          if (all(as.numeric(stats::na.omit(bar[, j])) == as.character(as.numeric(stats::na.omit(bar[, j]))))) {
            bar[, j] <- as.numeric(bar[, j])
          }
        }
      }
      openxlsx::writeData(wb, sheet = "Sheet 1", x = bar, startRow = factor_list[[i]][1], startCol = 1, colNames = F, withFilter = F)
    }
    options(warn = 0)

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
    for (i in seq_len(length(factor_row))) {
      if (!(is.integer(factor_row[i]) & length(factor_row[i]) == 0L)) {
        options(warn = -1)
        if (is.na(factor_row[i])) {
          factor_row[i] <- integer(0)
        }
        options(warn = 0)
      }
    }


    count <- 1
    factor_list <- list(NULL)
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

    #write data
    options(warn = -1)
    for (i in seq_len(length(factor_list))) {
      bar <- data[factor_list[[i]], ]
      for (j in seq_len(ncol(data[factor_list[[i]], ]))) {
        bar[bar[, j] == "", j] <- NA
        if (!any((is.na(as.numeric(stats::na.omit(bar[, j])))))) {
          if (all(as.numeric(stats::na.omit(bar[, j])) == as.character(as.numeric(stats::na.omit(bar[, j]))))) {
            bar[, j] <- as.numeric(bar[, j])
          }
        }
      }
      openxlsx::writeData(wb, sheet = "Sheet 1", x = bar, startRow = factor_list[[i]][1], startCol = 1, colNames = F, withFilter = F)
    }
    options(warn = 0)

    st <- openxlsx::createStyle(fontName = fontName, fontSize = fontSize, fgFill = significanceColor)
    openxlsx::addStyle(wb, "Sheet 1", style = st, cols = col_count[1], rows = factor_row)

    if (intercept == TRUE) {
      openxlsx::addStyle(wb, "Sheet 1", style = st, cols = 1, rows = factor_row)
    }else {
      openxlsx::addStyle(wb, "Sheet 1", style = st, cols = 1, rows = setdiff(factor_row, green_header[i] + 1))
    }
  }

  openxlsx::saveWorkbook(wb, paste0(fileName, ".xlsx"), overwrite = TRUE)
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
  options(warn = -1)
  if (is.na(dataName)) {
    stop("There is no data-name")
  }
  options(warn = 0)
  if (is.na(fileName)) {
    stop("There is no file-name")
  }
  if (!(is.character(fileName))) {
    stop("The file-name must be character")
  }
  data <- utils::read.table(paste0(dataName, ".csv"), fill = TRUE, header = FALSE, sep = ",", blank.lines.skip = FALSE, fileEncoding = fileEncoding)
  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, "Sheet 1")
  st <- openxlsx::createStyle(fontName = fontName, fontSize = fontSize)
  openxlsx::addStyle(wb, "Sheet 1", style = st, cols = 1:2, rows = 1:2)
  openxlsx::writeData(wb, sheet = "Sheet 1", x = data, colNames = F, withFilter = F)
  openxlsx::modifyBaseFont(wb, fontSize = fontSize, fontColour = fontColor, fontName = fontName)

  tmp <- apply(data, 1, function(x) {
    x[header == x]
    }
  )
  tmp_footer <- apply(data, 1, function(x) {
    all(x == "")
    }
  )
  green_header <- as.numeric(rownames(data[tmp == header, ]))
  footer <- as.numeric(rownames(data[tmp_footer, ]))

  col_intercept <- NA
  table_rightLim <- NULL

  if (adj == TRUE) {
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
    #setStyle for headers
    for (i in seq_len(length(green_header))) {
      st <- openxlsx::createStyle(fontName = fontName, fontSize = fontSize, fgFill = headerColor)
      openxlsx::addStyle(wb, "Sheet 1", style = st, cols = col_intercept[i]:table_rightLim[i], rows = green_header[i])
    }

    count <- 1
    factor_list <- list(NULL)
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

    #write data
    options(warn = -1)
    for (i in seq_len(length(factor_list))) {
      bar <- data[factor_list[[i]], ]
      for (j in seq_len(ncol(data[factor_list[[i]], ]))) {
        bar[bar[, j] == "", j] <- NA
        if (!any((is.na(as.numeric(stats::na.omit(bar[, j])))))) {
          if (all(as.numeric(stats::na.omit(bar[, j])) == as.character(as.numeric(stats::na.omit(bar[, j]))))) {
            bar[, j] <- as.numeric(bar[, j])
          }
        }
      }
      openxlsx::writeData(wb, sheet = "Sheet 1", x = bar, startRow = factor_list[[i]][1], startCol = 1, colNames = F, withFilter = F)
    }
    options(warn = 0)

  }else {
    st <- openxlsx::createStyle(fontName = fontName, fontSize = fontSize, fgFill = headerColor)
    for (k in seq_len(ncol(data))) {
      openxlsx::addStyle(wb, "Sheet 1", style = st, cols = k, rows = green_header)
    }

    count <- 1
    factor_list <- list(NULL)
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

    #write data
    options(warn = -1)
    for (i in seq_len(length(factor_list))) {
      bar <- data[factor_list[[i]], ]
      for (j in seq_len(ncol(data[factor_list[[i]], ]))) {
        bar[bar[, j] == "", j] <- NA
        if (!any((is.na(as.numeric(stats::na.omit(bar[, j])))))) {
          if (all(as.numeric(stats::na.omit(bar[, j])) == as.character(as.numeric(stats::na.omit(bar[, j]))))) {
            bar[, j] <- as.numeric(bar[, j])
          }
        }
      }
      openxlsx::writeData(wb, sheet = "Sheet 1", x = bar, startRow = factor_list[[i]][1], startCol = 1, colNames = F, withFilter = F)
    }
    options(warn = 0)
  }

  openxlsx::saveWorkbook(wb, paste0(fileName, ".xlsx"), overwrite = TRUE)
}
#'
#' Making directries to organize three kinds of datas: data-sets,  script-files and result-files
#' @encoding UTF-8
#'
#' @param parentDirectryName The name of a parent-directry containing organize datas-files, script-files and result-files.
#' @param dataDirectryName The name of a directry to organize data-files.
#' @param programmingDirectryName The name of a directry to organize script-files.
#' @param resultDirectryName The name of a directry to organize result-files.
#' @param updateTime The time used to divide data-filese into two directries, one is for datas and the other is for results.
#' @param arrange Allows you to organize data-files in the form of file extensions.
#'
#' @export
#'
mkDirectries <- function(parentDirectryName, dataDirectryName="data", programmingDirectryName="program", resultDirectryName="result", updateTime=1, arrange = TRUE) {
  dir.create(paste0(getwd(), "/", parentDirectryName))
  dir.create(paste0(getwd(), "/", parentDirectryName, "/", dataDirectryName))
  dir.create(paste0(getwd(), "/", parentDirectryName, "/", programmingDirectryName))
  dir.create(paste0(getwd(), "/", parentDirectryName, "/", resultDirectryName))
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
    if (!is.na(file.info(paste0(getwd(), "/", i))$mtime) &  i != parentDirectryName) {
      if (arrange == TRUE) {
        if (as.numeric(as.POSIXct(as.list(file.info(paste0(getwd(), "/", i)))$mtime, format = "%Y-%m-%d  %H:%M:%S", tz = "Japan") - Sys.time(), units = "mins") > (-1) * updateTime * 60) {
          if (is.na(strsplit(i, "\\.")[[1]][2]) & any(is.na(resultExtension))) {
            dir.create(paste0(getwd(), "/", parentDirectryName, "/", resultDirectryName, "/", "No Extension"))
            resultExtension[is.na(resultExtension)] <- ""
          }
          if (!is.na(any(resultExtension == strsplit(i, "\\.")[[1]][length(strsplit(i, "\\.")[[1]])]))) {
            if (any(resultExtension == strsplit(i, "\\.")[[1]][length(strsplit(i, "\\.")[[1]])])) {
              dir.create(paste0(getwd(), "/", parentDirectryName, "/", resultDirectryName, "/", strsplit(i, "\\.")[[1]][length(strsplit(i, "\\.")[[1]])]))
              resultExtension[resultExtension == strsplit(i, "\\.")[[1]][length(strsplit(i, "\\.")[[1]])]] <- ""
            }
          }
          if (is.na(strsplit(i, "\\.")[[1]][2])) {
            file.copy(paste0(getwd(), "/", i), paste0(getwd(), "/", parentDirectryName, "/", resultDirectryName, "/", "No Extension", "/", i))
          }else {
            file.copy(paste0(getwd(), "/", i), paste0(getwd(), "/", parentDirectryName, "/", resultDirectryName, "/", strsplit(i, "\\.")[[1]][length(strsplit(i, "\\.")[[1]])], "/", i))
          }
        }
        else {
          if (is.na(strsplit(i, "\\.")[[1]][2]) & any(is.na(dataExtension))) {
            dir.create(paste0(getwd(), "/", parentDirectryName, "/", dataDirectryName, "/", "No Extension"))
            dataExtension[is.na(dataExtension)] <- ""
          }
          if (!is.na(any(dataExtension == strsplit(i, "\\.")[[1]][length(strsplit(i, "\\.")[[1]])]))) {
            if (any(dataExtension == strsplit(i, "\\.")[[1]][length(strsplit(i, "\\.")[[1]])])) {
              dir.create(paste0(getwd(), "/", parentDirectryName, "/", dataDirectryName, "/", strsplit(i, "\\.")[[1]][length(strsplit(i, "\\.")[[1]])]))
              dataExtension[dataExtension == strsplit(i, "\\.")[[1]][length(strsplit(i, "\\.")[[1]])]] <- ""
            }
          }
          if (is.na(strsplit(i, "\\.")[[1]][2])) {
            file.copy(paste0(getwd(), "/", i), paste0(getwd(), "/", parentDirectryName, "/", dataDirectryName, "/", "No Extension", "/", i))
          }else {
            file.copy(paste0(getwd(), "/", i), paste0(getwd(), "/", parentDirectryName, "/", dataDirectryName, "/", strsplit(i, "\\.")[[1]][length(strsplit(i, "\\.")[[1]])], "/", i))
          }
        }
      }else {
        if (as.numeric(as.POSIXct(as.list(file.info(paste0(getwd(), "/", i)))$mtime, format = "%Y-%m-%d  %H:%M:%S", tz = "Japan") - Sys.time(), units = "mins") > (-1) * updateTime * 60) {
          file.copy(paste0(getwd(), "/", i), paste0(getwd(), "/", parentDirectryName, "/", resultDirectryName, "/", i))
        }
        else {
          file.copy(paste0(getwd(), "/", i), paste0(getwd(), "/", parentDirectryName, "/", dataDirectryName, "/", i))
        }
      }
    }
  }

  for (i in files[R.files]) {
    file.copy(paste0(getwd(), "/", i), paste0(getwd(), "/", parentDirectryName, "/", programmingDirectryName, "/", i))
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
    numericData <- t(rep("", 8))
  }
  options(warn = 0)
  table <- rbind(table, numericData)
  table <- rbind(table, rep("", 8))
  return(table)
}

mkFactorTable <- function(data, index) {
  table <- c(index, rep("", 8))
  table <- rbind(table, c("", "", "Levels", "", "Replace the column B with", "", "Pool the column B", "", "The Order of levels"))
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

dataClassifier <- function(data, index, dateIndex, dateList) {
  if (length(dateList[[dateIndex]]) > 1) {
    form1 <- paste0("%m", dateList[[dateIndex]][2], "%d", dateList[[dateIndex]][3], "%Y", dateList[[dateIndex]][1])
    form2 <- paste0("%Y", dateList[[dateIndex]][1], "%m", dateList[[dateIndex]][2], "%d", dateList[[dateIndex]][3])

    charData1 <- gsub(dateList[[dateIndex]][2], dateList[[dateIndex]][1],  data[, index])
    charData1 <- gsub(dateList[[dateIndex]][3], dateList[[dateIndex]][1],  data[, index])
    options(warn = -1)
    charData1 <- as.vector(apply(as.data.frame(strsplit(charData1, dateList[[dateIndex]][1])), 2, function(x) {
      paste0(formatC(as.numeric(x[3]), width = 4, flag = "0"), "-",
             formatC(as.numeric(x[1]), width = 2, flag = "0"), "-",
             formatC(as.numeric(x[2]), width = 2, flag = "0"))
    }
    )
    )
    options(warn = 0)
    charData2 <- gsub(dateList[[dateIndex]][2], dateList[[dateIndex]][1], data[, index])
    charData2 <- gsub(dateList[[dateIndex]][3], dateList[[dateIndex]][1], charData2)
    options(warn = -1)
    charData2 <- as.vector(apply(as.data.frame(strsplit(charData2, dateList[[dateIndex]][1])), 2, function(x) {
      paste0(formatC(as.numeric(x[1]), width = 4, flag = "0"), "-",
             formatC(as.numeric(x[2]), width = 2, flag = "0"), "-",
             formatC(as.numeric(x[3]), width = 2, flag = "0"))
    }
    )
    )
    options(warn = 0)
  }
  else {
    form1 <- paste0("%m", dateList[[dateIndex]][1], "%d", dateList[[dateIndex]][1], "%Y")
    form2 <- paste0("%Y", dateList[[dateIndex]][1], "%m", dateList[[dateIndex]][1], "%d")
    options(warn = -1)
    charData1 <- as.vector(apply(as.data.frame(strsplit(as.character(data[, index]), dateList[[dateIndex]][1])), 2, function(x) {
      paste0(formatC(as.numeric(x[3]), width = 4, flag = "0"), "-",
             formatC(as.numeric(x[1]), width = 2, flag = "0"), "-",
             formatC(as.numeric(x[2]), width = 2, flag = "0"))
    }
    )
    )
    charData2 <- as.vector(apply(as.data.frame(strsplit(as.character(data[, index]), dateList[[dateIndex]][1])), 2, function(x) {
      paste0(formatC(as.numeric(x[1]), width = 4, flag = "0"), "-",
             formatC(as.numeric(x[2]), width = 2, flag = "0"), "-",
             formatC(as.numeric(x[3]), width = 2, flag = "0"))
    }
    )
    )
    options(warn = 0)
  }
  return(list(list(form = form1, date = charData1), list(form = form2, date = charData2)))
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
  lengthMissVal <- length(unique(data[is.na(as.numeric(data))]))
  tmp <- data
  for (j in 1:lengthMissVal) {
    if (!is.na(refData[rowNumber + 1 + j, 4])) {
      data <- replace(data, data == unique(tmp[is.na(as.numeric(tmp))])[j], refData[rowNumber + 1 + j, 4])
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
  poolLevel <- list(NULL)
  for (i in seq_len(length(pooledFactor))) {
    if (length(pooledFactor[[i]]) > 1) {
      for (j in seq_len(length(pooledFactor[[i]]))) {
        if (!is.na(refData[rowNumber + 1 + as.numeric(pooledFactor[[i]][j]), 5])) {
          poolLevel[[i]]$pool <- paste0(poolLevel[[i]]$pool, "+",  refData[rowNumber + 1 + as.numeric(pooledFactor[[i]][j]), 5])
        }
        else {
          poolLevel[[i]]$pool <- paste0(poolLevel[[i]]$pool, "+",  refData[rowNumber + 1 + as.numeric(pooledFactor[[i]][j]), 3])
        }
        poolLevel[[i]]$levels <- append(poolLevel[[i]]$levels, as.numeric(pooledFactor[[i]][j]))
      }
      poolLevel[[i]]$pool <- substr(poolLevel[[i]]$pool, 2, nchar(poolLevel[[i]]$pool))
    }
  }

  for (i in seq_len(length(poolLevel))) {
    refData[rowNumber + 1 + poolLevel[[i]]$levels, 5] <- poolLevel[[i]]$pool
  }

  for (j in seq_len(factorLength)) {
    if (!is.na(refData[rowNumber + 1 + j, 9])) {
      orderNum <- as.numeric(refData[rowNumber + 1 + j, 9])
      if (!is.na(refData[rowNumber + 1 + j, 5])) {
        if (refData[rowNumber + 1 + j, 5] != "\"NA\"") {
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
  refData[(rowNumber + 2) : (rowNumber + 1 + nlevels(as.factor(data))), 5] != ""
  levelsRow <- refData[(rowNumber + 2) : (rowNumber + 1 + nlevels(as.factor(data))), 5] != ""
  factorLength <- 0
  rowIndex <- 1

  while (!is.na(refData[rowNumber + 1 + rowIndex, 2])) {
    factorLength <- factorLength + 1
    rowIndex <- rowIndex + 1
  }

  data <- as.character(data)

  refData[(rowNumber + 2) : (rowNumber + 1 + factorLength), 3] <- replace(refData[(rowNumber + 2) : (rowNumber + 1 + factorLength), 3], is.na(refData[(rowNumber + 2) : (rowNumber + 1 + factorLength), 3]), "")
  for (j in 1:nlevels(as.factor(data))) {
    if (!is.na(levelsRow[j])) {
      if (refData[rowNumber + 1 + j, 5] == "\"NA\"") {
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



#' Cleansing the dataset on a csv-file to change its form to more arranged one to handle.
#' @encoding UTF-8
#'
#' @param dataName The file-name of a csv file that will be cleansed.
#' @param dateFormat The format assigned to Date group.
#' @param append Allows you to append the new datas generated from dataCleansingForm__.xlsx.
#' @param NumOrFac The criteria for classifying whether the column data is numeric or factor. If the number of levels are greater than the ratio (nrow(data)/NumOrFac), then it will be assiged to numeric group.
#' @param leastNumOfDate The criteria for classifying whether the column data is Date of numeric. if the data contains the dateFormat you have chosen and the number of data containing such formats is greater than this value, leastNumOfDate, then the data will be assigned to Date group.
#' @param fileEncoding File-encoding
#'
#' @importFrom data.table fread
#' @export
dataCleanser <- function(dataName, dateFormat = list("/", "-"), append = FALSE, NumOrFac = 10, leastNumOfDate = 10, fileEncoding = "CP932") {
  files <- list.files()
  if (any(files == paste0("dataCleansingForm_", dataName, "_.xlsx")) == FALSE) {
    if (fileEncoding == "UTF-8" | fileEncoding == "Latin-1") {
      data <- data.table::fread(paste0(dataName, ".csv"), encoding = fileEncoding)
    }
    else {
      data <- data.table::fread(paste0(dataName, ".csv"), encoding = "unknown")
    }
    data <- as.data.frame(data)
    tableTime <- c("ColName", "Change the colName")
    tableNumeric <- c("ColName", "Change the colName", rep("", 6))
    tableFactor <- c("ColName", "Change the colName", rep("", 7))
    for (i in colnames(data)) {
      asDatedVector <- rep(FALSE, nrow(data))
      dateList <- dateFormat
      for (j in seq_len(length(dateList))) {
        tmp <- NULL
        for (k in strsplit(as.character(data[, i]), dateList[[j]][1])) {
          tmp <- append(tmp, length(k))
        }
        if ((all(tmp[grep(dateList[[j]][1], data[, i])] == 3) & length(tmp[grep(dateList[[j]][1], data[, i])] == 3) > 0) | (length(dateList[[j]]) > 1 & all(tmp[grep(dateList[[j]][1], data[, i])] == 2) & length(tmp[grep(dateList[[j]][1], data[, i])] == 2))) {
          formCharData <- dataClassifier(data, i, j, dateList)
          options(warn = -1)
          rowData <- as.numeric(rownames(data[as.character(as.Date(data[, i], form = formCharData[[1]]$form, origin = "1970-01-01")) == formCharData[[1]]$date & !is.na(as.character(as.Date(data[, i], form = formCharData[[1]]$form, origin = "1970-01-01")) == formCharData[[1]]$date), ]))
          if (all(asDatedVector[rowData] == FALSE)) {
            asDatedVector[rowData] <- TRUE
          }

          rowData <- as.numeric(rownames(data[as.character(as.Date(data[, i], form = formCharData[[2]]$form, origin = "1970-01-01")) == formCharData[[2]]$date & !is.na(as.character(as.Date(data[, i], form = formCharData[[2]]$form, origin = "1970-01-01")) == formCharData[[2]]$date), ]))
          options(warn = 0)
          if (all(asDatedVector[rowData] == FALSE)) {
            asDatedVector[rowData] <- TRUE
          }
        }
      }
      tmp <- NULL
      for (j in seq_len(length(dateList))) {
        tmp <- append(tmp, grep(dateList[[j]][1], data[, i]))
      }
      tmp <- setdiff(seq_len(nrow(data)), tmp)
      if (length(asDatedVector[asDatedVector == TRUE]) > leastNumOfDate) {
        asDatedVector[tmp] <- TRUE
      }
      if (all(asDatedVector == TRUE)) {
        tableTime <- rbind(tableTime, i)
        next ()
      }
      options(warn = -1)
      charEqualNum <- (as.numeric(data[, i]) == data[, i])
      options(warn = 0)
      if (length(na.omit(data[charEqualNum == FALSE, i])) == 0 &  length(na.omit(data[charEqualNum == TRUE, i])) > 0 & nlevels(as.factor(data[, i])) > nrow(data) / NumOrFac) {
        tableNumeric <- rbind(tableNumeric, mkNumericTable(data, i))
      }
      else {
        tableFactor <- rbind(tableFactor, mkFactorTable(data, i))
      }
    }
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
    openxlsx::saveWorkbook(wb, paste0("dataCleansingForm_", dataName, "_.xlsx"), overwrite = TRUE)
  }
  else {
    if (fileEncoding == "UTF-8" | fileEncoding == "Latin-1") {
      data <- data.table::fread(paste0(dataName, ".csv"), encoding = fileEncoding)
    }
    else {
      data <- data.table::fread(paste0(dataName, ".csv"), encoding = "unknown")
    }
    data <- as.data.frame(data)
    dataList <- NULL
    sheetList <- c("numeric", "factor", "Date")
    for (i in seq_len(length(sheetList))) {
      dataList[[i]] <- openxlsx::read.xlsx(paste0("dataCleansingForm_", dataName, "_.xlsx"), sheet = sheetList[i], colNames = F, skipEmptyRows = FALSE, skipEmptyCols = FALSE, na.strings = c("NA", ""))
    }
    for (i in colnames(data)) {
      if (!is.na(any(dataList[[1]][, 1] == i))) {
        if (any(dataList[[1]][, 1] == i)) {
          options(warn = -1)
          rowNumber <- as.numeric(rownames(dataList[[1]][dataList[[1]][, 1] == i & !is.na(dataList[[1]][, 1]), ]))
          colnames(data)[colnames(data) == i] <- changeColName(data, i, dataList[[1]], rowNumber)
          i <- changeColName(data, i, dataList[[1]], rowNumber)
          if (any(!is.na(dataList[[1]][(rowNumber + 2):(rowNumber + 1 + length(unique(data[is.na(as.numeric(data[, i])), i]))), 4])) & length(unique(data[is.na(as.numeric(data[, i])), i])) > 0) {
            if (append == TRUE) {
              data <- cbind(data, replaceMissVal(data[, i], dataList[[1]], rowNumber))
              colnames(data)[ncol(data)] <- paste0(i, "_missing Values replaced")
            }
            else {
              data[, i] <- replaceMissVal(data[, i], dataList[[1]], rowNumber)
            }
          }
          data[, i] <- as.numeric(data[, i])
          options(warn = 0)
          if (any(!is.na(dataList[[1]][rowNumber + 2, 6]))) {
            if (append == TRUE) {
              data <- cbind(data, cutting(data[, i], dataList[[1]], rowNumber))
              colnames(data)[ncol(data)] <- paste0(i, "_categorized")
            }
            else {
              data[, i] <- cutting(data[, i], dataList[[1]], rowNumber)
            }
          }
        }
      }
      if (!is.na(any(dataList[[2]][, 1] == i))) {
        if (any(dataList[[2]][, 1] == i)) {
          pooling <- FALSE
          ordering <- FALSE
          options(warn = -1)
          rowNumber <- as.numeric(rownames(dataList[[2]][dataList[[2]][, 1] == i & !is.na(dataList[[2]][, 1]), ]))
          options(warn = 0)
          colnames(data)[colnames(data) == i] <- changeColName(data, i, dataList[[2]], rowNumber)
          i <- changeColName(data, i, dataList[[2]], rowNumber)
          nrowLevel <- nlevels(as.factor(data[, i]))
          if (any(!is.na(dataList[[2]][(rowNumber + 2):(rowNumber + 1 +  nrowLevel), 5]))) {
            data[, i] <- replacer(data[, i], dataList[[2]], rowNumber)
            if (any(!is.na(dataList[[2]][(rowNumber + 2):(rowNumber + 1 + nlevels(as.factor(data[, i]))), 7]))) {
              if (append == TRUE) {
                data <- cbind(data, pooler(data[, i], dataList[[2]], rowNumber))
                colnames(data)[ncol(data)] <- paste0(i, "_", pooledName(data[, i], dataList[[2]], rowNumber))
                i <- colnames(data)[ncol(data)]
              }
              else {
                data[, i] <- pooler(data[, i], dataList[[2]], rowNumber)
              }
              pooling <- TRUE
              if (any(!is.na(dataList[[2]][(rowNumber + 2):(rowNumber + 1 +  nrowLevel), 9]))) {
                data[, i] <- orderer(data[, i], dataList[[2]], rowNumber)
                ordering <- TRUE
              }
            }
          }
          if (any(!is.na(dataList[[2]][(rowNumber + 2):(rowNumber + 1 + nlevels(as.factor(data[, i]))), 7])) & pooling == FALSE) {
            if (append == TRUE) {
              data <- cbind(data, pooler(data[, i], dataList[[2]], rowNumber))
              colnames(data)[ncol(data)] <- paste0(i, "_", pooledName(data[, i], dataList[[2]], rowNumber))
            }
            else {
              data[, i] <- pooler(data[, i], dataList[[2]], rowNumber)
            }
            pooling <- TRUE
            if (any(!is.na(dataList[[2]][(rowNumber + 2):(rowNumber + 1 + nlevels(as.factor(data[, i]))), 9])) & ordering == FALSE) {
              data[, i] <- orderer(data[, i], dataList[[2]], rowNumber)
              ordering <- TRUE
            }
          }
          if (any(!is.na(dataList[[2]][(rowNumber + 2):(rowNumber + 1 + nlevels(as.factor(data[, i]))), 9])) & ordering == FALSE) {
            data[, i] <- orderer(data[, i], dataList[[2]], rowNumber)
            ordering <- TRUE
          }
          data[, i] <- as.factor(data[, i])
        }
      }
      if (!is.na(any(dataList[[3]][, 1] == i))) {
        if (any(dataList[[3]][, 1] == i)) {
          asDatedVector <- rep(FALSE, nrow(data))
          dateList <- dateFormat
          options(warn = -1)
          rowNumber <- as.numeric(rownames(dataList[[3]][dataList[[3]][, 1] == i & !is.na(dataList[[3]][, 1]), ]))
          colnames(data)[colnames(data) == i] <- changeColName(data, i, dataList[[3]], rowNumber)
          i <- changeColName(data, i, dataList[[3]], rowNumber)
          for (j in seq_len(length(dateList))) {
            tmp <- NULL
            for (k in strsplit(as.character(data[, i]), dateList[[j]][1])) {
              tmp <- append(tmp, length(k))
            }
            if (all(tmp[grep(dateList[[j]][1], data[, i])] == 3) | (length(dateList[[j]]) > 1 & all(tmp[grep(dateList[[j]][1], data[, i])] == 2))) {
              formCharData <- dataClassifier(data, i, j, dateList)
              rowData <- as.numeric(rownames(data[as.character(as.Date(data[, i], form = formCharData[[1]]$form, origin = "1970-01-01")) == formCharData[[1]]$date & !is.na(as.character(as.Date(data[, i], form = formCharData[[1]]$form, origin = "1970-01-01")) == formCharData[[1]]$date), ]))
              if (length(rowData) > 0) {
                if (all(asDatedVector[rowData] == FALSE)) {
                  data[rowData, i] <- formCharData[[1]]$date[rowData]
                  asDatedVector[rowData] <- TRUE
                }
              }
              rowData <- as.numeric(rownames(data[as.character(as.Date(data[, i], form = formCharData[[2]]$form, origin = "1970-01-01")) == formCharData[[2]]$date & !is.na(as.character(as.Date(data[, i], form = formCharData[[2]]$form, origin = "1970-01-01")) == formCharData[[2]]$date), ]))
              if (length(rowData) > 0) {
                if (all(asDatedVector[rowData] == FALSE)) {
                  data[rowData, i] <- formCharData[[2]]$date[rowData]
                  asDatedVector[rowData] <- TRUE
                }
              }
            }
          }
          options(warn = 0)
        }
      }
    }
    return(data)
  }
}

