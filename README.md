# VBA-Challenge: Stock Metrics Calculator

## Table of Contents

- [Overview](#overview)
- [Prerequisites](#prerequisites)
- [Methodology Applied](#usage)
- [Sample Data](#sample-data)
- [Screenshots](#screenshots)
- [VBA files](#VBAfiles)
- [License](#license)

## Overview

The **Stock Metrics Calculator** is a collection of VBA scripts designed to efficiently analyze stock data, calculate essential metrics, and provide valuable insights. This repository showcases separate VBA scripts for each sheet within an Excel workbook.

## Prerequisites

- Microsoft Excel (for running the VBA scripts)
- Basic understanding of VBA scripting

## Methodology Applied

1. Open the provided `alphabetical_testing.xlsx` file, which contains a subset of stock data for demonstration purposes

2. Access the VBA editor in Excel:
   - Press `ALT` + `F11` to open the Visual Basic for Applications editor.
   - For each sheet in your workbook, insert a new module (`Insert` > `Module`).

3. Copy and paste the relevant VBA script from the corresponding file in the `VBA_Scripts` directory into the newly created module for each sheet.

4. **Looping through Sheets**: Each sheet should have its own VBA script. To apply the script to all sheets:
   - Create a button (Form Control or ActiveX Control) on each sheet.
   - Assign the corresponding VBA script to the button.
   - The script will analyze the data on that specific sheet.

5. **Applying to Other Data**: To apply the script to different data:
   - Copy and paste the script into a new module in the VBA editor.
   - Modify the script to reference the correct sheet name and data range.

6. After inserting the scripts and buttons, return to the workbook.
  
7. Run the VBA script for each sheet by clicking the respective execution button. This will calculate and display the metrics on each sheet.


## Sample Data

Alphabetical testing data has been used to showcase the functionality of the scripts while keeping the dataset size manageable.

Feel free to replace the sample data with your own dataset for a comprehensive analysis.

## Screenshots

The [`Screenshots`](Screenshots/) directory contains screenshots displaying the results obtained after running each script on its respective sheet.

## VBA files

The [`VBA Files`](VBAFiles/) directory contains the VBA codescript corresponding to each sheet of the alphabetical_testing workbook.

## License

This project is licensed under the [MIT License](LICENSE).
