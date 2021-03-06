\name{Workbook}
\alias{Workbook}
\alias{createWorkbook}
\alias{loadWorkbook}
\alias{saveWorkbook}
\title{Functions to manipulate Excel 2007 workbooks.}
\description{
Functions to manipulate Excel 2007 workbooks. 
}
\usage{

createWorkbook(type="xlsx")

loadWorkbook(file, password=NULL)

saveWorkbook(wb, file, password=NULL)

}
\arguments{
  \item{type}{a String, either \code{xlsx} for Excel 2007 OOXML format,
	or \code{xls} for Excel 95 binary format.}
  \item{file}{the path to the file you intend to read or write.
	Can be an xls or xlsx format.}
  \item{wb}{a workbook object as returned by \code{createWorkbook} or
    \code{loadWorkbook}.}
  \item{password}{a String with the password.}
}

\details{
  \code{createWorkbook} creates an empty workbook object.

  \code{loadWorkbook} loads a workbook from a file.

  \code{saveWorkbook} saves an existing workook to an Excel 2007 file.

  Reading or writing of password protected workbooks is supported for
  Excel 2007 OOXML format only.  Note that in Linux, LibreOffice is not
  able to read password protected spreadsheets. 
}
\value{
  \code{createWorkbook} returns a java object reference pointing to an 
  empty workbook object.
  
  \code{loadWorkbook} creates a java object reference corresponding to
  the file to load.

}
\author{Adrian Dragulescu}
\seealso{\code{\link{write.xlsx}} for writing a \code{data.frame} to an
  \code{xlsx} file.  \code{\link{read.xlsx}} for reading the content of
  a \code{xlsx} worksheet into a \code{data.frame}.  To extract
  worksheets and manipulate them, see \code{\link{Worksheet}}.}
\examples{

wb <- createWorkbook()

# see all the available java methods that you can call
#.jmethods(wb)

# for example
wb$getNumberOfSheets()   # no sheet yet!

# loadWorkbook("C:/Temp/myFile.xls")

}

