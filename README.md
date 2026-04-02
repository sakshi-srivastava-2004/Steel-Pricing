#  Steel-Pricing Analytics System

<p align="center">
  <b>Automated Excel Data Processing & Weekly Price Analysis</b><br>
  Transform raw steel pricing data into structured insights 
</p>



##  Overview

This project is designed to process multiple Excel files containing steel pricing data and generate a **consolidated weekly analytics report**.

It intelligently extracts, cleans, and transforms raw Excel data into meaningful insights by handling:
-Highlighted headers 
- Merged cells
- Missing and inconsistent values
- Special cases like `'H'` values



## ✨ Features

-  **Automated Excel Processing**  
  Reads and processes multiple `.xlsx` files from a directory  

-  **Header Detection Using Color**  
  Detects table structure based on yellow-highlighted cells  

-  **Smart Weekly Aggregation**  
  Calculates weekly average prices using Date/Week logic  

  **Intelligent Data Cleaning**  
  - Handles missing values  
  - Replaces `'H'` values with nearest valid values  
  - Removes invalid rows  

  **Multi-Dimensional Analysis**  
  Works across:
  - Materials  
  - Regions  
  - Weekly timelines  

  **Consolidated Output**  
  Generates a single structured Excel file with clean data  



**Tech Stack**

- 🐍 Python  
- 📊 Pandas  
- 📄 OpenPyXL  
- 🔢 NumPy  




