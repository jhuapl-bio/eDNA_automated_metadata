# eDNA_automated_metadata

The code in this repo are the scripts which generate and execute the entensions' functions for the eDNA metadata tracking sheet.
The sheet is based on the (FAIRe template)[https://fair-edna.github.io/index.html] and has been adapted for systems and needs aboard the OceanXplorer.

# Set Up and Usage
In order to use the metadata sheet, you do not need to pull anything from this repo. The code in this repo exists for those who which to understand the functions and/or adapt the scripts to for their own purposes.

To use the metadata simply follow the step below to make your own copy, link it to your data files, and start processing you data!
1) Make a copy of (this sheet)[https://docs.google.com/spreadsheets/d/12Q50EyFOjgzUrDFeDTH9WpJVlWXpVbs3SB3SPf1Glhk/edit?usp=sharing] in your Google Drive. Make folders CTD, ROV, and Flowthrough in the same directory as your sheet copy (or just copy this whole directory).																									
2) Copy your data files to their respective directories. For CTD you need the .btl files. For ROV you need the RBR _data.txt files and the _sealogExport.csv files. For Flowthrough you need the TSG SBE45 and the Fluorometer_Wetlab .txt files.																									
3) Go to the projectMetadata sheet and add the cruise ID to Unique Cruise ID under value1. Add the names of any relevant personnel after recordedBy. Add your insitutuon name after institution. 																									
     Add any other relevant information to the appropriate columns.																									
4) Go to the samplesheet_v1 sheet. This is where you will enter all your data.																									
5) For CTD, add the cast number under dive_or_cast_number, select CTD under asset, and select Niskin under collection_method. Finally, for each of your samples, select the relevant bottle number under bottle_number.																									
     You may select more than one number for samples that pooled water from multiple bottles.																									
6) Once you have added the above information for each of your CTD samples, select the rows for all of these samples, then click the Import Data button at the top, then select "Import CTD data from BTL" from the dropdown menu.																									
     This should populate each of your selected samples with the appropriate metadata.																									
7) To import metadata for ROV samples, add the dive number under dive_or_cast_number, select ROV under asset, select Niskin under collection_method, and select the appropraite bottle numbers under bottle_number for each of your samples.																									
     You may select more than one number for samples that pooled water from multiple bottles.																									
8) Once you have added the above information for each of your ROV samples, select the rows for all of these samples, then click the Import Data button at the top, then select "Import ROV data from CSV/TXT" from the dropdown menu.																									
     This should populate each of your selected samples with the appropriate metadata.																									
9) To import metadata for Flowthrough samples, add the sampling date under eventDate (YYYY-MM-DD), add the Z-time of the sample under z-time in the format YYYY-MM-DDTHH:MM:SSZ, and select Flowthrough under both asset and collection_method.																									
     You may select more than one number for samples that pooled water from multiple bottles.																									
10) Once you have added the above information for each of your Flowthrough samples, select the rows for all of these samples, then click the Import Data button at the top, then select "Import Flowthrough data from TXTs" from the dropdown menu.																									
     This should populate each of your selected samples with the appropriate metadata.																									
11) See instructions for how to print sample labels using the DYMO Labels button																									
12) To display your samples on a map, select the rows of the samples you would like to display and click the Sample Map Button at the top. You can choose to display the plotted samples in a sidebar or in a new tab. You can toggle the map layer between ESRI bathymetry or Google Maps.																									
