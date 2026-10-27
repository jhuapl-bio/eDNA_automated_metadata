# Printing DYMO Labels for eDNA Samples

#### NOTE: These instructions allow printing of 2x1 cryo labels on a DYMO LabelWriter 500. Additional optimization may be required for other printers.

#### This document describes the process for using the DYMO label printer to print sample labels directly from google sheets.

1. Ensure the following metadata fields are filled out correctly:

    * dive_or_cast_number
    * asset  
    * collection_method 	
    * bottle_number

2. Highlight the rows you would like to print labels for, and select the DYMO Labels extension at the top of the sheet

3. Click, “Generate Label PDF from Selected Rows

4. A sidebar will pop-up to the right of the screen, click “Generate PDF” when ready

5. A PDF window will pop-up with your labels

6. You may now print your labels. Select the print icon
   
7. You will need to select the correct dimensions of the labels to print properly, click “Print using system dialog” 

8. Under “Paper Size” you will need to select “Manage Custom Sizes, and make a custom label size (I called ours Dymo_2x1) and input the following dimensions:

    * Width: 2 in
    * Height: 1 in
    * Margin Top: 0.25 in
    * Margin Bottom: 0.56 in
    * Margin Left: 0.25 in
    * Margin Right: 0.25 in

10. Click “OK” and now you should be able to select the custom size and print your labels.

You can also download the DYMO Connect Software to use the label maker for any additional custom labels.
