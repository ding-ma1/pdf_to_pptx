# pdf_to_pptx
A (naïve) way to convert PDF slides to PPTX slides using R

## Motivation
Some conferences and presentations require slides in PowerPoint format, but you prefer to create them in PDF format (e.g., using LaTeX or markdown).

## Description
This R script converts each page of a PDF file into a slide in a PowerPoint (PPTX) file. Each slide is represented as an image, preserving the original layout and design of
the PDF slides. The script uses a template PPTX file to maintain consistent slide dimensions and formatting.

## Requirements
- R packages: `officer`, `magick`, `pdftools`
- poppler-utils (for PDF rendering)
- ImageMagick (for image processing, probably required by `magick` package on Windows systems, can be downloaded from [here](https://imagemagick.org/script/download.php))
- A PDF file containing the slides you want to convert
- A PPTX template file with the desired slide dimensions (optional, if the provided templates with 16:9 or 4:3 aspect ratios do not suit your need)

## Usage
```R
source("pdf_to_pptx.R")
pdf_to_pptx("your_slides.pdf", "template_for_aspect_ratio.pptx")
```
Then open the generated PPTX file (in the same directory of `your_slides.pdf` by default) in PowerPoint and double check, delete the pages inherited from `template_for_aspect_ratio.pptx`.

Done.
