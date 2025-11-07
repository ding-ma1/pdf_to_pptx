source("pdf_to_pptx.R")
pdf_to_pptx(pdf_file = "example/metropolis_beamer_example_16_9.pdf",
            template_pptx = "aspect_ratio/template_16_9.pptx", # use the provided templates template_16_9.pptx for 16:9 or template_4_3.pptx for 4:3 aspect ratios, or use your own template
            # pptx_file = "example.pptx", # output file name, NULL to auto-generate
            dpi = 300) # change the DPI (resolution) if needed, 300 is usually good

# Please double check the output PPTX file in PowerPoint after running the code,
# and remove the page(s) inherited from the template that you do not need.

# Warning message about aspect ratio mismatch (if appears):
# 1.333 refers to 4:3 aspect ratio, 1.778 refers to 16:9 aspect ratio