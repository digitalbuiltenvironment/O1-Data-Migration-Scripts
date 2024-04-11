# PWDM Powershell Script

## Table of Contents

- [About](#about)
- [Getting Started](#getting_started)
- [Usage](#usage)

## About <a name = "about"></a>

Contains 3 different powershell script for exporting of PWDM data.

1) PWDM_Single_Export.ps1 - Exports a single project's PWDM Data
2) PWDM_Multi_Export.ps1 - Exports multiple project's PWDM Data
3) PWDM_Multi_Export_Metadata.ps1 - Exports multiple project's PWDM Metadata only

## Getting Started <a name = "getting_started"></a>

1) Edit Line 11 in the `PWDM_Multi_Export.ps1 | PWDM_Multi_Export_Metadata.ps1` or Line 8 in the `PWDM_Single_Export.ps1` script to include the full path of a folder that which you want to save the exported data to.

2) FOR `PWDM_Multi_Export.ps1 | PWDM_Multi_Export_Metadata.ps1`:
   - Edit the 'project_ids.txt' file to include the Project IDs you want to export:
   
   ![image](https://github.com/digitalbuiltenvironment/O1-Data-Migration-Scripts/assets/21101460/569d3a5a-0317-47df-8c87-565dd19e5dbb)

   - Project ID can be found in the URL by going to connect.bentley.com and entering the particular project:
   
   ![image](https://github.com/digitalbuiltenvironment/O1-Data-Migration-Scripts/assets/21101460/3b2437cc-67df-48ef-9bea-29e17e0897ba)

  
4) FOR `PWDM_Single_Export.ps1`:
   - Edit Line 6 in the script to include the Project ID of the single project you want to export.

## Usage <a name = "usage"></a>

Open Powershell in Administrator mode, then `cd` to the folder containing the script and run the script you need or you can provide the full path to the script in the commandline.
