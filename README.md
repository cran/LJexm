# LJexm

**LJexm** is an 'R' package designed to automate the process of extracting zip files, converting 'Word' and 'Excel' files to 'pdf' files using 'VBScript', and merging 'pdf' files in a structured order.

## Installation

To install 'LJexm' from source, run:

```r
# Install from source
install.packages("LJexm", repos = NULL, type = "source")
```

## Usage

Load the package and run the application:

```r
library(LJexm)
run_app()
```

### **Error Handling**
- Errors and status messages are **logged in `conversion_log.txt`** during execution.
- After processing, **messages are printed using `message()`** in R.

#### **Manually Checking Errors (If Needed)**
If you need to check errors manually before running the script, you can view the log file:

```r
log_file <- file.path("path/to/your/folder", "conversion_log.txt")
if (file.exists(log_file)) {
  log_content <- readLines(log_file)
  message(log_content, appendLF = TRUE)
}
```

## Features
- Extracts 'zip' files in a given directory.
- Converts `.docx` and `.xlsx` files to 'pdf' using 'VBScript'.
- Merges 'pdf' files in alphabetical order, ensuring correct sequencing.
- **Automatically logs errors and prints them in R using `message()`.**
- Works in both 'RStudio' and Base R.

## Dependencies
This package requires the following dependencies:
- `fs`
- `pdftools`
- `rstudioapi`
- `utils`

These dependencies will be automatically installed when you install 'LJexm'.

## License

This package is licensed under GPL-3.

