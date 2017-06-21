# ------------------------------------------------------------
# Product Revenue Plot
# ------------------------------------------------------------
# --- Create the history overview
plotProductRevevueLegendLeft <- ggplot(df_ProductData, aes(x = MONTH, y = REVENUE, group = PRODUCT, color = PRODUCT)) +
    geom_line(stat = "identity", size = 1) +
    labs(title = "Revenue for XYZ",
         subtitle = "",
         x = "Month",
         y = "Revenue") +
    theme_classic() +
    theme(legend.title=element_blank()) +
    theme(axis.text.x = element_text(angle = 90, hjust = 1))

plotProductRevevueLegendBottom <- ggplot(df_ProductData, aes(x = MONTH, y = REVENUE, group = PRODUCT, color = PRODUCT)) +
    geom_line(stat = "identity", size = 1) +
    labs(title = "Revenue for XYZ",
         subtitle = "",
         x = "Month",
         y = "Revenue") +
    theme_classic() +
    theme(legend.title=element_blank(),
          axis.text.x = element_text(angle = 90, hjust = 1),
          legend.position = "bottom")
      

# --- Save to disk in case someone wants to use it or if you want to
# --- compile a PDF afterwards using Markdown => Simply pick up the image
savePlotToDisk("plotProductRevevueLegendLeft", plotProductRevevueLegendLeft)
savePlotToDisk("plotProductRevevueLegendBottom", plotProductRevevueLegendBottom)

# --- Go and get the crosstab version of this data for PPT presentation.
df_ProductDataXtab <- getColRowTotals(df_ProductData, "sortDesc", "formatYes", "indicateYes")

# --- Store any result sets of processed data in xlsx format
v_FileName = "../02_data_processed/product_revenue_crosstab.xlsx"
write.xlsx2(x = df_ProductDataXtab, file = v_FileName, sheetName = "All Products - Monthly View", col.names = T, append = F)
