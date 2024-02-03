# Project "Excel tables comparator"

## Description

The "Excel tables comparator" project is an application for comparing data in Excel tables. The project is organized according to the MVC (Model-View-Controller) architecture and includes the following components:

1. **Analysis Table Module (`analysTable.js`):** This module implements functions for creating an Excel table with data analysis, setting color labels, and processing comparative data between the original and test tables. And creating second sheet (information sheet) , where explaining terms and colors from first sheet.

2. **Excel Tables Model (`excel.model.js`):** The model is responsible for obtaining data from Excel tables, processing information about clients, their data, and comparing this data between the original and test tables.

3. **Excel Tables Controller (`excel.controller.js`):** The controller combines the functionalities of the model and the analysis table module, providing methods to initiate the data analysis process and create results in an Excel table.

4. **Main Application File (`app.js`):** In this file, an instance of the controller is created, and the process of analyzing data is initiated.

## Installation and Running

1. Clone the repository: `git clone <repository URL>`
2. Install dependencies: `npm install`
3. Run the project: `npm start`

## Features

- Creating an Excel table with data analysis for clients.
- Comparing data between the original and test tables.
- Setting color labels to highlight differences and full match of data.

## Project Structure

The project is organized as follows:

```plaintext
/project-root
│
├── server/
│   ├── src/
│   │   ├── controller/
│   │   │   ├── excel.controller.js
│   │   ├── model/
│   │   │   ├── analysTable.model.js
│   │   │   ├── excel.model.js
│   │   ├── app.js
│   └── data/
│       ├── Analys_Table.xlsx
│       ├── originalTable.xlsx
├── package.json
└── README.md
