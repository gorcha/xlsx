
# Create a streaming workbook
createStreamWorkbook <- function()
{
  wb <- .jnew("org/apache/poi/xssf/streaming/SXSSFWorkbook", as.integer(-1))
  return(wb)
}

loadStreamWorkbook <- function(file, password=NULL)
{
  wb <- loadWorkbook(file, password)
  wb <- .jnew("org/apache/poi/xssf/streaming/SXSSFWorkbook", wb, as.integer(-1))
  return(wb)
}

saveStreamWorkbook <- function(wb, file, password=NULL)
{
  saveWorkbook(wb, file, password)

  # Run dispose on SXSSF to remove temporary files
  wb$dispose()

  invisible()
}

# Add a data.frame to a sheet
addStreamDataFrame <- function(x, sheet, col.names=TRUE, row.names=TRUE,
    startRow=1, startColumn=1, colStyle=NULL, colnamesStyle=NULL,
    rownamesStyle=NULL, showNA=FALSE, characterNA="", byrow=FALSE,
    writeRows=100)
{
  if (!is.data.frame(x))
    x <- data.frame(x)    # just because the error message is too ugly

  if (row.names) {        # add rownames to data x as the first column
    x <- cbind(rownames=rownames(x), x)
    if (!is.null(colStyle))
      names(colStyle) <- as.numeric(names(colStyle)) + 1
  }
  
  wb <- sheet$getWorkbook()
  classes <- unlist(sapply(x, class))
  if ("Date" %in% classes) 
    csDate <- CellStyle(wb) + DataFormat(getOption("xlsx.date.format"))
  if ("POSIXct" %in% classes) 
    csDateTime <- CellStyle(wb) + DataFormat(getOption("xlsx.datetime.format"))

  # offset required to give space for column names
  # (either excel columns if byrow=TRUE or rows if byrow=FALSE)
  iOffset <- if (col.names) 1L else 0L
  # offset required to give space for row names
  # (either excel rows if byrow=TRUE or columns if byrow=FALSE)
  jOffset <- if (row.names) 1L else 0L

  if ( byrow ) {
      # write data.frame columns data row-wise
      setDataMethod   <- "setRowData"
      setHeaderMethod <- "setColData"
      blockRows <- ncol(x)
      blockCols <- nrow(x) + iOffset # row-wise data + column names
  } else {
      # write data.frame columns data column-wise, DEFAULT
      setDataMethod   <- "setColData"
      setHeaderMethod <- "setRowData"
      blockRows <- nrow(x) + iOffset # column-wise data + column names
      blockCols <- ncol(x)
  }
  
  # insert colnames
  if (col.names) {
    if (!(ncol(x) == 1 && colnames(x)=="rownames"))  {
      # create a CellBlock, not sure why the usual .jnew doesn't work
      cellBlock <- CellBlock( sheet,
          as.integer(startRow), as.integer(startColumn),
          1L, as.integer(blockCols),
          TRUE)
  
      .jcall( cellBlock$ref, "V", setHeaderMethod, 0L, jOffset,
         .jarray(colnames(x)[(1+jOffset):ncol(x)]), showNA,
         if ( !is.null(colnamesStyle) ) colnamesStyle$ref else
             .jnull('org/apache/poi/ss/usermodel/CellStyle') )
      sheet$flushRows
    }
  }
  
  # write one data.frame column at a time, and style it if it has style
  # Dates and POSIXct columns get styled if not overridden. 
  for (i in seq(1, nrow(x), writeRows)) {
    # number of records for CellBlock (to avoid blank columns at end of sheet)
    if ((nrow(x) - i + 1) < writeRows) {
      cbRows <- nrow(x) - i + 1
    } else {
      cbRows <- writeRows
    }
      
    # create a CellBlock, not sure why the usual .jnew doesn't work
    cellBlock <- CellBlock( sheet,
        as.integer(startRow + i - 1L + as.logical(col.names)), as.integer(startColumn),
        as.integer(cbRows), as.integer(blockCols),
        TRUE)

    for (j in 1:ncol(x)) {
    thisColStyle <-
      if ((j==1) && (row.names) && (!is.null(rownamesStyle))) {
        rownamesStyle
      } else if (as.character(j) %in% names(colStyle)) {
        colStyle[[as.character(j)]]
      } else if ("Date" %in% class(x[i:(i + cbRows - 1),j])) {
        csDate
      } else if ("POSIXt" %in% class(x[i:(i + cbRows - 1),j])) {
        csDateTime
      } else {
        NULL
      }
      
    xj <- x[i:(i + cbRows - 1),j]
    if ("integer" %in% class(xj)) {
      aux <- xj
    } else if (any(c("numeric", "Date", "POSIXt") %in% class(xj))) {
      aux <- if ("Date" %in% class(xj)) {
          as.numeric(xj)+25569
        } else if ("POSIXt" %in% class(x[,j])) {
          as.numeric(xj)/86400 + 25569
        } else {
          xj
        }
      haveNA <- is.na(aux)
      if (any(haveNA))
        aux[haveNA] <- NaN          # encode the numeric NAs as NaN for java
      } else {
      aux <- as.character(x[i:(i + cbRows - 1),j])
      haveNA <- is.na(aux)
      if (any(haveNA))
        aux[haveNA] <- characterNA

      # Excel max cell size limit 
      if (max(nchar(aux)) > .EXCEL_LIMIT_MAX_CHARS_IN_CELL) {
          warning(sprintf("Some cells exceed Excel's limit of %d characters and they will be truncated", 
                          .EXCEL_LIMIT_MAX_CHARS_IN_CELL))
          aux <- strtrim(aux, .EXCEL_LIMIT_MAX_CHARS_IN_CELL)   
    }
    }

    # Write data to cell block
   .jcall( cellBlock$ref, "V", setDataMethod,
      as.integer(j-1L),   #  -1L for Java index 
      0L,                 # no offset needed for SXSSF
      .jarray(aux), showNA, 
      if ( !is.null(thisColStyle) ) thisColStyle$ref else
        .jnull('org/apache/poi/ss/usermodel/CellStyle') )
  }
    # Flush data to disk
    sheet$flushRows()
  }
  
  # return the cellBlock occupied by the generated data frame
  invisible(cellBlock)
}

# Write a data.frame to a new xlsx file using streaming XSSF. 
write.xlsx.stream <- function(x, file, sheetName="Sheet1",
  col.names=TRUE, row.names=TRUE, append=FALSE,
  password=NULL, writeRows=100, ...)
{
  if (append && file.exists(file)){
    wb <- loadStreamWorkbook(file)
  } else {
    wb  <- createStreamWorkbook()
  }  
  sheet <- createSheet(wb, sheetName)
  
  addStreamDataFrame(x, sheet, col.names=col.names, row.names=row.names,
    startRow=1, startColumn=1, colStyle=NULL, colnamesStyle=NULL,
    rownamesStyle=NULL, writeRows=writeRows)
  
  saveStreamWorkbook(wb, file, password=password)  
  
  invisible()
}
