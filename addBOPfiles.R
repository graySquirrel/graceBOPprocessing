library(openxlsx)

# To run, go to the folder where the R script is and run 'Rscript addBOPfiles.R'

# assume you are running this script at the level where subdirectories masterfiles and "BOP outputs" exists
bopfilefolder <- "BOP outputs"

# only take files > 210616
newfiles <- list.files(path=bopfilefolder, pattern="\\BOP000.csv$")
newdates <- as.Date(newfiles, format = "%y%m%d")
newfiles <- newfiles[newdates > as.Date("2021-06-16")]
if (length(newfiles) > 0) 
{
  # make a backup copy of master file
  nowstamp <- format(Sys.time(), "%Y-%m-%d_%H_%M_%s")
  masterfilename <- "masterfiles/BlastAnimalOutcomes.xlsx"
  newmasterfilename <- paste0(sub('\\.xlsx$','', masterfilename),nowstamp,".xlsx")
  print(paste("copying",masterfilename,"to",newmasterfilename))
  file.copy(masterfilename, newmasterfilename)
  for (f in paste0(bopfilefolder,"/",newfiles))
  {
    print(paste("processing",f))
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
    newf <- paste0(sub('\\.csv$','', f),"processed.csv")
    print(paste("renaming",f,"to",newf))
    file.rename(f, newf)
  }
}