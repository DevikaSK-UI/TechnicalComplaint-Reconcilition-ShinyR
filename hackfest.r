library(shiny)
library(readxl)
library(openxlsx)
library(DT)
library(dplyr)
library(shinyjs)  # For showing/hiding elements

# Define UI
ui <- fluidPage(
  includeCSS("styles1.css"), 
  useShinyjs(),  # Enable JavaScript functionalities
  titlePanel("Technical Complaint Reconciliation Tool"),
  
  sidebarLayout(
    sidebarPanel(
      h2("File Selection", class = "file-selection"),
      fileInput("ccc_file", "Select CCC Excel File", accept = c(".xlsx")),
      fileInput("edc_file", "Select EDC Excel File", accept = c(".xlsx")),
      actionButton("compare_btn", "Compare Files", class = "btn-compare"),
      br(),
      downloadButton("downloadData", "Download Result"),
      br(),
      actionButton("show_results_btn", "Show Results", style = "display: none;", class = "btn-show-results"),
      textOutput("output_file")
    ),
    
    mainPanel(
      h2(class = "comparison-results-heading", "Comparison Results"),
      dataTableOutput("comparison_table"),
      textOutput("output_file")
    )
  )
)

# Define server logic
server <- function(input, output, session) {
  # Store the comparison result as a reactive value
  comparison_result <- reactiveVal(NULL)
  
  # Perform comparison when the button is clicked
  observeEvent(input$compare_btn, {
    req(input$ccc_file, input$edc_file)
    
    ccc_data <- read_excel(input$ccc_file$datapath)
    edc_data <- read_excel(input$edc_file$datapath)
    
    # Select and rename columns to match
    ccc_columns <- c("Subject/Patient ID", "Technical Complaint No.", "AE related", "DUN Number", "Trial/Study Number")
    edc_columns <- c("Subject", "Seq No", "AE related", "Dispense Unit Number ID", "Trial/Study Number")
    
    ccc_data <- ccc_data %>% select(all_of(ccc_columns))
    edc_data <- edc_data %>% select(all_of(edc_columns))
    
    colnames(ccc_data) <- c("Subject", "Seq No", "AE related", "Dispense Unit Number ID", "Trial/Study Number")
    
    # Initialize result columns in EDC data
    edc_data$Status <- "Not Present"
    edc_data$Mismatch_Details <- ""
    
    # Comparison logic
    for (i in 1:nrow(edc_data)) {
      matching_row <- ccc_data %>% filter(Subject == edc_data$Subject[i])
      
      if (nrow(matching_row) > 0) {
        mismatches <- list()
        if (edc_data$`Seq No`[i] != matching_row$`Seq No`) mismatches <- c(mismatches, paste("Seq No mismatch"))
        if (edc_data$`AE related`[i] != matching_row$`AE related`) mismatches <- c(mismatches, paste("AE mismatch"))
        if (edc_data$`Dispense Unit Number ID`[i] != matching_row$`Dispense Unit Number ID`) mismatches <- c(mismatches, paste("DUN mismatch"))
        if (edc_data$`Trial/Study Number`[i] != matching_row$`Trial/Study Number`) mismatches <- c(mismatches, paste("Trial mismatch"))
        
        if (length(mismatches) == 0) {
          edc_data$Status[i] <- "Match"
        } else {
          edc_data$Status[i] <- "Mismatch"
          edc_data$Mismatch_Details[i] <- paste(mismatches, collapse = "; ")
        }
      } else {
        edc_data$Mismatch_Details[i] <- "Not present in CCC"
      }
    }
    
    # Store the result for further use
    comparison_result(edc_data)  # Update reactive value
    
    # Enable result display button and download button
    shinyjs::show("show_results_btn")
    shinyjs::show("downloadData")
  })
  
  # Show comparison result in a table
  observeEvent(input$show_results_btn, {
    output$comparison_table <- renderDataTable({
      datatable(comparison_result(), options = list(page_length=10)) %>%
        formatStyle(
          'Status',
          backgroundColor = styleEqual(c('Match', 'Mismatch', 'Not Present'),
                                        c('green', 'red', 'orange'))
        )
    })
  })
  
  # Download comparison result as Excel
  output$downloadData <- downloadHandler(
    filename = function() {
      paste0("comparison_result_", Sys.Date(), ".xlsx")
    },
    content = function(file) {
      data_to_write <- comparison_result()
      if (is.null(data_to_write)) {
        showNotification("No data available to download", type = "error")
        return(NULL)
      }
      
      wb <- createWorkbook()
      addWorksheet(wb, "comparison_result")
      
      # Write data to workbook
      writeData(wb, "comparison_result", data_to_write)
      
      # Define styles
      headerStyle <- createStyle(fontColour = "white", fgFill = "#003366", halign = "center", valign = "center", textDecoration = "bold", border = "Bottom", fontSize = 12)
      matchStyle <- createStyle(fgFill = "green", halign = "center")
      mismatchStyle <- createStyle(fgFill = "red", halign = "center")
      notPresentStyle <- createStyle(fgFill = "orange", halign = "center")
      defaultStyle <- createStyle(halign = "center", valign = "center")
      
      # Apply styles to the worksheet
      addStyle(wb, sheet = 1, headerStyle, rows = 1, cols = 1:ncol(data_to_write), gridExpand = TRUE)
      addStyle(wb, sheet = 1, defaultStyle, rows = 2:(nrow(data_to_write)+1), cols = 1:ncol(data_to_write), gridExpand = TRUE)
      
      # Apply conditional formatting for 'Status' column
      status_col <- which(colnames(data_to_write) == "Status")
      status_rows <- 2:(nrow(data_to_write) + 1)
      
      for (i in seq_along(status_rows)) {
        if (data_to_write$Status[i] == "Match") {
          addStyle(wb, sheet = 1, matchStyle, rows = status_rows[i], cols = status_col, gridExpand = TRUE)
        } else if (data_to_write$Status[i] == "Mismatch") {
          addStyle(wb, sheet = 1, mismatchStyle, rows = status_rows[i], cols = status_col, gridExpand = TRUE)
        } else if (data_to_write$Status[i] == "Not Present") {
          addStyle(wb, sheet = 1, notPresentStyle, rows = status_rows[i], cols = status_col, gridExpand = TRUE)
        }
      }
      
      # Set column width for better readability
      setColWidths(wb, sheet = 1, cols = 1:ncol(data_to_write), widths = "auto")
      
      # Save workbook to the specified file
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
}

# Run the application 
shinyApp(ui = ui, server = server)

