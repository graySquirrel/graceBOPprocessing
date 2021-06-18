library(openxlsx)

# go to the folder.  assume below this folder is 'newmice' with new bop files, and 'masterfile' where the big xls is.
setwd("/Users/fritzebner/Documents/R/Grace/temp/graceBOPprocessing/newmice/")

# make a backup copy of master file
nowstamp <- format(Sys.time(), "%Y-%m-%d_%H_%M_%s")
masterfilename <- "../masterfiles/BlastAnimalOutcomes.xlsx"
file.copy(masterfilename, paste0(sub('\\.xlsx$','', masterfilename),nowstamp,".xlsx"))

# only take files > 210616
newfiles <- list.files(pattern="\\BOP000.csv$")
newdates <- as.Date(newfiles, format = "%y%m%d")
newfiles <- newfiles[newdates > as.Date("2021-06-16")]
for (f in newfiles) 
{
  print(f)
  # get date from filename
  thedate <- as.Date(f, format = "%y%m%d")
  # read in the bop file
  df <- read.csv(f, header = FALSE, stringsAsFactors = FALSE, fileEncoding="latin1")
  # get psi
  psi <- df[df$V1=="BOP","V2"]
  psi <- trimws(psi)
  psi <- substr(psi,start = 1, stop = (nchar(psi) - 4))
  # get mouse number
  mousenum <- trimws(df$V1[dim(df)[1]])
  len <- nchar(mousenum)
  mousenum <- substr(mousenum, start = len - 3, stop = len)
  # read in whole workbook
  wb <- loadWorkbook(masterfilename)
  # get master, append new data
  master <- read.xlsx(masterfilename, 1, detectDates = TRUE)
  newrownum <- dim(master)[1] + 1
  master[newrownum,"Date"] <- thedate
  master[newrownum,"Subject"] <- mousenum
  master[newrownum,"BOP.(psi)"] <- psi
  writeData(wb, sheet = "Master", master, colNames = T)
  saveWorkbook(wb, masterfilename, overwrite = T)
  # change name of file, so its not processed twice
  file.rename(f, paste0(sub('\\.csv$','', f),"processed.csv"))
}