
# -----------------------------------------------------------------------------
# --- Clear your desk first!
rm(list=ls())

# -----------------------------------------------------------------------------
# --- Make sure we are were we need to be
setwd("~/mycloud/Work/AAA_Projects/AAA 2017/AAA Meetup/src")
t1 <- Sys.time()

# -----------------------------------------------------------------------------
# --- Init all libraries needed
source("01_LoadLibraries.R")

# -----------------------------------------------------------------------------
# --- Data fethers and code here
td1 <- Sys.time()
source("02_LoadData.R")
td2 <- Sys.time()

# -----------------------------------------------------------------------------
# --- Post processing and plot code here
source("03_ProcessPlotData.R")

# -----------------------------------------------------------------------------
# --- Create the slides
source("04_CreateSlides.R")

# -----------------------------------------------------------------------------
# --- How much time did we save so we can go home earlier, 
# --- still crank out stuff with quality, consistency and look smart?
t2 <- Sys.time()
paste("The process completed in ",
      format(t2-t1, digits = 4),
      paste(". Aqcuiring data took ", format(td2-td1, digits = 4), sep = " "),
      sep = "")

# -----------------------------------------------------------------------------
# --- As good citizens we should cleanup memory and not leave a mess behind
#rm(list=ls())