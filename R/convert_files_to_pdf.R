#' Convert Word & Excel Files to PDFs using 'VBScript'
#' @param folder_path The directory containing Word and Excel files
#' @return No return value, called for side effects.
#' @export
convert_files_to_pdf <- function(folder_path) {
  excel_script <- system.file("excel.vbs", package = "LJexm")
  word_script <- system.file("word_new.vbs", package = "LJexm")

  run_vbs <- function(script_path, folder) {
    if (file.exists(script_path)) {
      system2("wscript", args = c(shQuote(script_path), shQuote(folder)), stdout = FALSE, stderr = FALSE)
    } else {
      stop(paste("Error:", basename(script_path), "not found at", script_path))
    }
  }

  # Run Excel and Word conversion scripts
  run_vbs(excel_script, folder_path)
  run_vbs(word_script, folder_path)

  # Read and print the log file in R
  log_file <- file.path(folder_path, "conversion_log.txt")
  if (file.exists(log_file)) {
    log_content <- readLines(log_file)
    message(log_content, appendLF = TRUE)

    # Delete the log file after printing
    file.remove(log_file)
    message("Log file deleted after execution.")
  }
}
