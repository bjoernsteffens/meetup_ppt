# --- Install all these packages of this is a new install of RStudio
# install.packages("rJava")
# install.packages("scales")
# install.packages("dtplyr")
# install.packages("dplyr")
# install.packages("data.table")
# install.packages("formattable")
# install.packages("tidyr")
# install.packages("ggthemes")
# install.packages("RColorBrewer")
# install.packages("gridExtra")
# install.packages("grid")
# install.packages("xlsx")
# install.packages("markdown")
# install.packages("rmarkdown")
# install.packages("ReporteRs")
# install.packages("ReporteRsjars")
# install.packages("ggplot2")
# install.packages("RJDBC")
# install.packages("dtplyr")
# install.packages("extrafont")
# install.packages("lazyeval")


# --- Load what you need
library(lazyeval)
library(extrafont)
library(markdown)
library(rmarkdown)
library(xlsx, quietly = TRUE)
library(RColorBrewer, quietly = TRUE)
require(rJava, quietly = TRUE)
.jinit()
require(ReporteRsjars, quietly = TRUE)
require(ReporteRs, quietly = TRUE)
require(ggplot2, quietly = TRUE)
require(scales, quietly = TRUE)
require(RJDBC, quietly = TRUE)
require(formattable, quietly = TRUE)
require(tidyr, quietly = TRUE)
require(dplyr, quietly = TRUE)
library(data.table, quietly = TRUE)
#require(dtplyr, quietly = TRUE)
require(ggthemes, quietly = TRUE)
require(stringr, quietly = TRUE)
library(gridExtra, quietly = TRUE)
library(grid, quietly = TRUE)


# -----------------------------------------------------------------------------------------
# Save a plot to disk
# -----------------------------------------------------------------------------------------
savePlotToDisk <- function (plotName, plotImage) {
    
    tmpName <- paste("../03_images/",plotName,".png",sep = "")
    png(tmpName, width = 1500, height = 1000, res = 144)
    plot(plotImage)
    dev.off()
    
}

# -----------------------------------------------------------------------------------------
# Apply thousand separators to numbers
# -----------------------------------------------------------------------------------------
fmtVals <- function (x) {
    ifelse(x != 0, format(x, digits = 3, justify = "left", big.mark = ","), x)
}

# -----------------------------------------------------------------------------------------
# Drop thousand separators to numbers
# -----------------------------------------------------------------------------------------
# --- Feature used to annotate the final dataframe
fmtValsRemove <- function (x) {
    
    ifelse(x != 0, as.numeric(gsub("'", '', x)), x)
    
}

# -----------------------------------------------------------------------------------------
# Remove Zero value
# -----------------------------------------------------------------------------------------
dropZero <- function (x) {
    ifelse(x == 0, "", x)
}


getColRowTotals <- function (x, sortOrder = "sortDesc", format = "formatYes", indicate = "indicateNo") {
    
    #dfa <- df_ProductData
    dfa <- x
    dfb <- spread(dfa, MONTH,  REVENUE, fill = 0)
    
    # --- Row Totals
    r <- rowSums(dfb[,-1], na.rm = TRUE)
    dfr <- data.frame(t(r))
    dfr <- t(dfr)
    colnames(dfr) <- c("Total")
    dfa <- as.data.frame(cbind(dfb,dfr))
    
    # --- Sort the table before adding the row totals
    if ( sortOrder == "sortDesc" ) { dfa <- arrange(dfa, desc(Total)) }
    if ( sortOrder == "sortAsc")   { dfa <- arrange(dfa, Total) }
    
    # --- Column Totals
    c <- colSums(dfa[,-1], na.rm = TRUE)
    dfb <- data.frame(t(c))
    dfb <- cbind(Brand = "Total", dfb)
    colnames(dfb) <- colnames(dfa)
    dfa <- as.data.frame(rbind(dfa,dfb))
    
    if ( format == "formatYes") {
        # --- Format the numbers and drop zeros
        dfa[,2:length(dfa)] <- sapply(dfa[,2:length(dfa)], fmtVals)
        dfa[,2:length(dfa)] <- sapply(dfa[,2:length(dfa)], dropZero)
    }
    
    # --- Add indicators to the text
    # --- loop across the frame and compare n to n-1. If positive ↑ if Negative ↓
    if ( indicate == 'indicateYes') {
        for ( i in 1:nrow(dfa) ) {
            #i = 12
            
            dfx <- dfa[i,2:(length(dfa)-1)]
            cmpWith <- as.numeric(gsub(",", "", dfx[1,1]))
            
            
            for ( k in 1:(length(dfx)-1) ) {
                #k = 7
                #cmpCurr <- trimws(as.character(dfx[1,c(k+1)]))
                
                # --- Catch the case when it starts with a value being NA
                if ( is.na(cmpWith) )  {
                    cmpWith <- as.numeric(gsub(",", "", dfx[1,c(k+1)]))
                    next
                }
                
                cmpCurr <- as.numeric(gsub(",", "", dfx[1,c(k+1)]))
                # --- or there is a NA in the middle
                if ( is.na(cmpCurr) ) {
                    cmpWith <- cmpCurr
                    next
                }
                
                if ( cmpCurr > cmpWith ) {
                    cmpWith <- as.numeric(gsub(",", "", dfx[1,c(k+1)]))
                    dfa[i,c(k+2)] <- paste(dfa[i,c(k+2)], '↑', sep = "")
                } else if (cmpCurr < cmpWith) {
                    cmpWith <- as.numeric(gsub(",", "", dfx[1,c(k+1)]))
                    dfa[i,c(k+2)] <- paste(dfa[i,c(k+2)], '↓', sep = "")
                } else if (cmpCurr == cmpWith) {
                    cmpWith <- as.numeric(gsub(",", "", dfx[1,c(k+1)]))
                }
            }
        }
    }
    
    dfa
}
