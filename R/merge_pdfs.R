#' Merge PDFs in Alphabetical Order
#' @param folder_path The directory containing sub-folders with PDF files
#' @return Character vector with the path of the merged PDF file.
#' @export
merge_pdfs <- function(folder_path) {
  subfolders <- fs::dir_ls(folder_path, type = "directory")

  for (subfolder in subfolders) {
    pdf_files <- fs::dir_ls(subfolder)
    pdf_files <- pdf_files[grepl("\\.pdf$", pdf_files, ignore.case = TRUE)]
    pdf_files <- sort(pdf_files)

    output_pdf <- file.path(subfolder, paste0(basename(subfolder), "_merged.pdf"))

    if (length(pdf_files) >= 2) {
      tryCatch({
        pdftools::pdf_combine(pdf_files, output_pdf)
        message("Merged PDF: ", output_pdf)
      }, error = function(e) {
        warning("Error merging PDFs in ", subfolder, " - ", e$message)
      })
    } else {
      message("Not enough PDFs in ", subfolder, " to merge.")
    }
  }
}
