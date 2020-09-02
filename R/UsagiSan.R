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

#' Coloring the signigicant variables and corresponding p values in statistical tests tables on a EXCEL sheet
#' @encoding UTF-8
#'
#' @param dataName The name of a csv-file you want to edit with coloring
#' @param fileName The name of a Excel-file you want to save as
#' @param level The significance level applied in coloring significant variables and p values
#' @param pValue The character object which indicates the column name in each stitistical test tables
#' @param significanceColor The fore-ground-color of the significant variables
#' @param headerColor The fore-ground-color of the headers of a data-frame
#' @param fontSize Font-size
#' @param fontName Font-name
#' @param fontColor The color of fonts
#' @param intercept Allows you to color significant intercept variable with the fontColor
#' @param adj Allows you yo adjust shifted statistical test tables
#' @param fileEncoding File-encoding
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
excelColor <- function(dataName, fileName, level = 0.05, pValue = c("Pr(>|z|)", "Pr(>|t|)", "p-value"), significanceColor = "#FFFF00", headerColor = "#92D050", fontSize = 11, fontName = enc2native("\u6e38\u30b4\u30b7\u30c3\u30af"), fontColor = "#000000",  intercept = FALSE, adj = TRUE, fileEncoding = "CP932") {
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
#' @param dataName The name of a csv-file you want to edit with coloring
#' @param fileName The name of a Excel-file you want to save as
#' @param header The character object included in each headers of tables
#' @param headerColor The fore-ground-color of the headers of a data-frame
#' @param fontSize Font-size
#' @param fontName Font-name
#' @param fontColor The color of fonts
#' @param adj Allows you yo adjust shifted statistical test tables
#' @param fileEncoding File-encoding
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
excelHeadColor <- function(dataName, fileName, header, headerColor = "#92D050", fontSize = 11, fontName = enc2native("\u6e38\u30b4\u30b7\u30c3\u30af"), fontColor = "#000000", adj = TRUE, fileEncoding = "CP932") {
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
#' @param parentDirectryName The name of a parent-directry containing organize datas-files, script-files and result-files
#' @param dataDirectryName The name of a directry to organize data-files
#' @param programmingDirectryName The name of a directry to organize script-files
#' @param resultDirectryName The name of a directry to organize result-files
#' @param updateTime The time used to divide data-filese into two directries, one is for datas and the other is for results
#' @param arrange Allows you to organize data-files in the form of file extensions
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