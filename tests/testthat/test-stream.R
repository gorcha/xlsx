context('write.xlsx.stream')

test_that('stream export', {
  ##Testing high level export ...  
  x <- data.frame(mon=month.abb[1:10], day=1:10, year=2000:2009,
                  date=seq(as.Date("2009-01-01"), by="1 month", length.out=10),
                  bool=ifelse(1:10 %% 2, TRUE, FALSE))
  
  file <- test_tmp("test_stream_export.xlsx")
  ## write an xlsx file with char, int, double, date, bool columns ...
  write.xlsx.stream(x, file, sheetName="writexlsx")
  
  ## test the append argument by adding another sheet ...
  write.xlsx.stream(USArrests, file, sheetName="usarrests", append=TRUE)
  
  ## test writing/reading data.frames with NA values ...
  file <- test_tmp("test_writeread_NA.xlsx")
  
  x <- data.frame(matrix(c(1.0, 2.0, 3.0, NA), 2, 2))
  write.xlsx.stream(x, file, row.names=FALSE)
  xx <- read.xlsx(file, 1)
  expect_identical(x,xx)
})

test_that('write password protected stream workbook succeeds', {
  ## issue #49
  
  x <- data.frame(values=c(1,2,3),stringsAsFactors=FALSE)
  filename <- test_tmp('issue49.xlsx')
  
  ## write
  write.xlsx.stream(x, filename, password='test', row.names=FALSE)
  
  ## read
  r <- read.xlsx2(filename, sheetIndex = 1, password='test'
                  , stringsAsFactors=FALSE
                  , colClasses = 'numeric')
  
  expect_identical(x, r)
})
