library(openxlsx)

# To run, go to Blaster - General/BOP outputs/graceBOPprocessing and run 'Rscript addBOPfiles.R'

setwd("..") # First, go up a dir.  Execution will happen in "BOP outputs"

# assume you are running this script at the level where subdirectories masterfiles and "BOP outputs" exists
bopfilefolder <- "."
#masterfilename <- "masterfiles/BlastAnimalOutcomes.xlsx"
masterfilefolder <- "masterfiles"

# only take files > 210616
newfiles <- list.files(path=bopfilefolder, pattern="\\BOP000.csv$")
newdates <- as.Date(newfiles, format = "%y%m%d")
newfiles <- newfiles[newdates > as.Date("2021-06-16")]

if (length(newfiles) > 0) 
{
  # make a backup copy of master file
  nowstamp <- format(Sys.time(), "%Y-%m-%d_%H_%M_%s")
  #newmasterfilename <- paste0(sub('\\.xlsx$','', masterfilename),nowstamp,".xlsx")
  #print(paste("copying",masterfilename,"to",newmasterfilename))
  #file.copy(masterfilename, newmasterfilename)
  # load workbook up
  #wb <- loadWorkbook(masterfilename)
  #master <- read.xlsx(masterfilename, 1, detectDates = TRUE)
  
  master <- data.frame(matrix(ncol = 13, nrow=0))
  columnNames <- c("Date","Subject","BOP Input (psi)", "BOP (psi)", 
                   "Outcome", "Measures (Experiments)","48h NS", "7 day NS",
                   "Actual Driver Pressure", "Still have", "","Filename", "Notes")
  colnames(master) <- columnNames
  
  for (f in newfiles)
  {
    bopfilepath <- paste0(bopfilefolder,"/",f)
    print(paste("processing",f))
    # get date from filename
    thedate <- as.Date(f, format = "%y%m%d")
    # read in the bop file
    df <- read.csv(bopfilepath, header = FALSE, stringsAsFactors = FALSE, fileEncoding="latin1")
    # get input pressure from 'Driver Set'
    driverset <- df[df$V1=="Driver Set","V2"]
    driverset <- trimws(driverset)
    driverset <- substr(driverset, start=1,stop=(nchar(driverset) - 4))
    # get psi
    psi <- df[df$V1=="BOP","V2"]
    psi <- trimws(psi)
    psi <- substr(psi,start = 1, stop = (nchar(psi) - 4))
    # get mouse number
    notes <- trimws(df$V1[dim(df)[1]])
    len <- nchar(notes)
    mousenum <- as.integer(substr(notes, start = len - 3, stop = len))
    newrownum <- dim(master)[1] + 1
    master[newrownum,"Date"] <- difftime(thedate, "1899-12-30", units = "days") + 1
    master[newrownum,"Subject"] <- mousenum
    master[newrownum,"BOP Input (psi)"] <- driverset
    master[newrownum,"BOP (psi)"] <- psi
    master[newrownum, "Notes"] <- notes
    # change name of file, so its not processed twice
    newf <- paste0(sub('\\.csv$','', bopfilepath),"processed.csv")
    print(paste("renaming",bopfilepath,"to",newf))
    file.rename(bopfilepath, newf)
  }
  # update xls and save it back out
  #writeData(wb, sheet = "Master", master, colNames = T)
  #saveWorkbook(wb, masterfilename, overwrite = T)
  write.xlsx(master, file = paste0(masterfilefolder,"/","newBops",nowstamp,".xlsx"))
  #write.csv(master, file = paste0(masterfilefolder,"/","newBops",nowstamp,".csv"),
  #          row.names = FALSE, na="")
}
