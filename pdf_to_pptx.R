# Utility to check and install R packages
install_if_missing <- function(pkg) {
  if (!requireNamespace(pkg, quietly = TRUE)) {
    install.packages(pkg, dependencies = TRUE)
  }
  suppressPackageStartupMessages(library(pkg, character.only = TRUE))
}

# Ensure required packages
install_if_missing("magick")
install_if_missing("officer")
install_if_missing("pdftools")

# Check if Poppler is available
check_poppler <- function() {
  res <- tryCatch(
    system2("pdftoppm", "--version", stdout = TRUE, stderr = TRUE),
    warning = function(w) w,
    error = function(e) e
  )
  
  # If command is missing
  if (inherits(res, "error") ||
      (inherits(res, "warning") && grepl("not found", res$message, ignore.case = TRUE))) {
    stop(
      "Poppler is required for PDF reading with magick.\n",
      "Please install Poppler:\n",
      "- macOS: brew install poppler\n",
      "- Windows: download from https://blog.alivate.com.au/poppler-windows/ and add bin/ to PATH\n",
      "- Linux: sudo apt install poppler-utils"
    )
  }
  invisible(TRUE)
}

# Main function: Convert PDF slides to PPTX
pdf_to_pptx <- function(pdf_file,
                        pptx_file = NULL,
                        template_pptx,
                        dpi = 300) {
  
  check_poppler()
  
  if (is.null(pptx_file)) {
    pptx_file <- sub("\\.pdf$", ".pptx", pdf_file, ignore.case = TRUE)
  }
  
  if (missing(template_pptx) || !file.exists(template_pptx)) {
    stop("You must supply an existing PPTX template via template_pptx")
  }
  
  # Stop if template is old .ppt format
  if (tolower(tools::file_ext(template_pptx)) == "ppt") {
    stop(
      "Template file is a .ppt (old PowerPoint format). Only .pptx templates are supported.\n",
      "Please convert the template to .pptx before using this function."
    )
  }
  
  # Read PDF pages as high-resolution images
  img_list <- magick::image_read_pdf(pdf_file, density = dpi)
  n_pages <- length(img_list)
  message("PDF has ", n_pages, " page(s). Starting conversion...")
  
  # Get PDF first page aspect ratio
  pdf_info <- magick::image_info(img_list[1])
  pdf_ratio <- pdf_info$width / pdf_info$height
  
  # Read template
  pptx <- officer::read_pptx(template_pptx)
  dims <- officer::slide_size(pptx)
  slide_w <- dims[["width"]]
  slide_h <- dims[["height"]]
  pptx_ratio <- slide_w / slide_h
  
  # Warn if PDF vs template aspect ratio differs
  if (abs(pdf_ratio - pptx_ratio) > 0.01) {
    warning(
      sprintf(
        "Aspect ratio mismatch: PDF %.3f vs template %.3f. Slides may be scaled or cropped.",
        pdf_ratio, pptx_ratio
      )
    )
  }
  
  # Add each PDF page as a new slide
  for (i in seq_along(img_list)) {
    tmp_img <- tempfile(fileext = ".png")
    magick::image_write(img_list[i], path = tmp_img, format = "png")
    
    pptx <- officer::add_slide(pptx, layout = "Blank", master = "Office Theme")
    pptx <- officer::ph_with(
      pptx,
      officer::external_img(tmp_img, width = slide_w, height = slide_h),
      location = officer::ph_location(left = 0, top = 0, width = slide_w, height = slide_h)
    )
    
    message("Processed page ", i, " / ", n_pages)
  }
  
  # Save PPTX
  print(pptx, target = pptx_file)
  message("PDF successfully converted to PPTX: ", pptx_file)
}
