mydoc <- pptx(template = "../template.pptx")
docFooterText <- paste("All data on this slide was extracted on", as.character(t1), sep = " ")
margin <- theme(plot.margin = unit(c(0,0,0.1,0), "cm"))

# -----------------------------------------------------------------------------------------
# --- Title Slides
# -----------------------------------------------------------------------------------------
mydoc = addSlide( mydoc, slide.layout = "Title Slide", bookmark = 1 )
mydoc = addTitle( mydoc, "Leveraging R to Explore Digital Analytics Data" )
v_SubTitle <- as.character((paste(paste("Meetup@Amazee ", format((seq(Sys.Date(), length = 2, by = "-1 months")[2]), '%B %Y'), sep = ""), 
                                  paste("Last updated ", format(Sys.Date(), '%m/%d/%Y'), sep = ""), 
                                  paste("", "", sep = "\n"),
                                  paste("bjoern.steffens1@gmail.com", "+41 78 759 70 80", sep = " "),
                                  sep = "\n")))
mydoc = addSubtitle( mydoc, v_SubTitle)

# -----------------------------------------------------------------------------------------
# --- Links to Github Slide
# -----------------------------------------------------------------------------------------
pot1 = pot("Sample code on Github", textProperties( color = "blue", font.size = 24, font.family = "Calibri (Body)" ), hyperlink = "https://github.com/bjoernsteffens/meetup_ppt")
mydoc = addSlide( mydoc, slide.layout = "Hyperlinks" )
mydoc = addTitle( mydoc, "Source code is on Github" )
mydoc = addParagraph( mydoc, set_of_paragraphs(pot1) )
mydoc = addFooter( mydoc, docFooterText)
mydoc = addPageNumber( mydoc )

# -----------------------------------------------------------------------------------------
# --- Product Overview - A
# -----------------------------------------------------------------------------------------
mydoc = addSlide( mydoc, slide.layout = "Hyperlinks" )
mydoc = addTitle( mydoc, "Product Revenue Overview" )
mydoc = addPlot(mydoc, function() print(plotProductRevevueLegendLeft), vector.graphic=TRUE, editable = TRUE)
mydoc = addFooter( mydoc, docFooterText)
mydoc = addPageNumber( mydoc )


# -----------------------------------------------------------------------------------------
# --- Product Overview - B
# -----------------------------------------------------------------------------------------
pot1 = pot(paste(trimws(format(df_ProductDataXtab[1, 14], digits = 3, justify = "left", big.mark = ",")), 
                 "-", 
                 df_ProductDataXtab[1, 1], 
                 sep = " "),
           textProperties( font.size = 18, font.family = "Calibri Light" ))

pot2 = pot(paste(trimws(format(df_ProductDataXtab[2, 14], digits = 3, justify = "left", big.mark = ",")), 
                 "-", 
                 df_ProductDataXtab[2, 1], 
                 sep = " "),
           textProperties( font.size = 18, font.family = "Calibri Light" ))

pot3 = pot(paste(trimws(format(df_ProductDataXtab[3, 14], digits = 3, justify = "left", big.mark = ",")), 
                 "-", 
                 df_ProductDataXtab[3, 1], 
                 sep = " "),
           textProperties( font.size = 18, font.family = "Calibri Light" ))

pot4 = pot(paste(trimws(format(df_ProductDataXtab[4, 14], digits = 3, justify = "left", big.mark = ",")), 
                 "-", 
                 df_ProductDataXtab[4, 1], 
                 sep = " "),
           textProperties( font.size = 18, font.family = "Calibri Light" ))

pot5 = pot(paste(trimws(format(df_ProductDataXtab[5, 14], digits = 3, justify = "left", big.mark = ",")), 
                 "-", 
                 df_ProductDataXtab[5, 1], 
                 sep = " "),
           textProperties( font.size = 18, font.family = "Calibri Light" ))

pot6 = pot(paste(trimws(format(df_ProductDataXtab[6, 14], digits = 3, justify = "left", big.mark = ",")), 
                 "-", 
                 df_ProductDataXtab[6, 1], 
                 sep = " "),
           textProperties( font.size = 18, font.family = "Calibri Light" ))

mydoc = addSlide( mydoc, slide.layout = "Two Content" )
mydoc = addTitle( mydoc, "Product Revenue Overview" )
mydoc = addPlot(mydoc, function() print(plotProductRevevueLegendBottom), vector.graphic=TRUE, editable = TRUE)
mydoc = addParagraph( mydoc, set_of_paragraphs(pot1, pot2, pot3, pot4, pot5, pot6) )
mydoc = addFooter( mydoc, docFooterText)
mydoc = addPageNumber( mydoc )

# -----------------------------------------------------------------------------------------
# --- Product Details
# -----------------------------------------------------------------------------------------
flexTable <- vanilla.table(df_ProductDataXtab)
flexTable <- setFlexTableWidths(flexTable, widths = c(2.5,rep(0.75, each = length(df_ProductDataXtab)-1)))
flexTable <- setZebraStyle(flexTable, odd = "#FAFAFA", even = "#D3D3D3")

mydoc = addSlide( mydoc, slide.layout = "Title and Content" )
mydoc = addTitle( mydoc, "Product Revenue Details" )
mydoc = addFlexTable( mydoc, flexTable)
mydoc = addFooter( mydoc, docFooterText)
mydoc = addPageNumber( mydoc )

# -----------------------------------------------------------------------------------------
# --- Save to disk
# -----------------------------------------------------------------------------------------
mydoc = addSlide( mydoc, slide.layout = "Title Slide")
mydoc = addTitle( mydoc, "Backup Charts" )
mydoc = addSubtitle( mydoc, paste("Last updated ", format(Sys.Date(), '%m/%d/%Y'), sep = ""))
writeDoc( mydoc, file = "../Meetup PPT Demo.pptx" )

