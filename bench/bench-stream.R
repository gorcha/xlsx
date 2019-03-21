options(java.parameters = "-Xmx1024m", scipen = 999)
library(xlsx)
library(rJava)

test_tmp <- function(file) {
  tmp_folder <- rprojroot::is_r_package$find_file('bench/tmp')
  if (!file.exists(tmp_folder))
    dir.create(tmp_folder)
  rprojroot::is_r_package$find_file(paste0('bench/tmp/',file))
}

remove_tmp <- function() {
  tmp_folder <- rprojroot::is_r_package$find_file('bench/tmp')
  if (file.exists(tmp_folder))
    unlink(tmp_folder, recursive=TRUE)
}


bench::press(
  rows = c(100, 1000, 10000, 100000),
  {
    javagc()
    file <- test_tmp(paste0("bench-", rows, ".xlsx"))
    x <- data.frame(a = 1:rows,
                    b = sample(letters, rows, replace = TRUE),
                    c = sample(letters, rows, replace = TRUE),
                    d = seq(as.Date("2009-01-01"), by="1 month", length.out=rows))
    bench::mark(write.xlsx2(x, file),
                write.xlsx.stream(x, file),
                write.xlsx.stream(x, file, writeRows = 1000),
                min_iterations = 2,
                check = FALSE)
  }
)

remove_tmp()
