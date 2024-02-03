"use strict"
import ExcelJS from 'exceljs';
const { Workbook } = ExcelJS;
class AnalysTableModel {
  setHeadersCellColor(worksheet, rowNum) {
    let row = worksheet.getRow(rowNum);

    row.font = { bold: true };

    const actualCellCount = row.actualCellCount;

    for (let colNumber = 1; colNumber <= actualCellCount; colNumber++) {
      const cell = row.getCell(colNumber);
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'ffbdb9b9' }
      };
    }

    for (let colIndex = 1; colIndex <= 9; colIndex++) {
      const cell = worksheet.getCell(1, colIndex);
      cell.border = { top: { style: 'medium' }, left: { style: 'medium' }, bottom: { style: 'medium' }, right: { style: 'medium' } };
    }
    for (let colIndex = 1; colIndex <= 20; colIndex++) {
      const cell = worksheet.getCell(3, colIndex);
      cell.border = { top: { style: 'medium' }, left: { style: 'medium' }, bottom: { style: 'medium' }, right: { style: 'medium' } };
    }

  }


  setInfoHeadersCellColor(worksheet, rowNum = 1) {
    let row = worksheet.getRow(rowNum);

    row.font = { bold: true };

    const actualCellCount = row.actualCellCount;

    for (let colNumber = 1; colNumber <= actualCellCount; colNumber++) {
      const cell = row.getCell(colNumber);
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'ffbdb9b9' }
      };
    }

    for (let colIndex = 1; colIndex <= 4; colIndex++) {
      const cell = worksheet.getCell(1, colIndex);
      cell.border = { top: { style: 'medium' }, left: { style: 'medium' }, bottom: { style: 'medium' }, right: { style: 'medium' } };
    }
  }


  async setColorsForInfoCellColA(worksheetInfo, objColors) {
    worksheetInfo.getColumn(1).eachCell({ includeEmpty: false }, (cell, rowNumber) => {
      if (rowNumber > 1 && rowNumber < 9) {
        let color = objColors[cell.value];
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: color }
        };
      }
    });
  }


  async createAnalysTable(pathNewTable = 'server/data/Analys_Table.xlsx') {
    try {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Sheet1');
      const worksheetInfo = workbook.addWorksheet('Info');

      const headerRow = ["CountClientIdInOriginalTable", "CountClientIdInTestTable", "CountMatchingClientIds", "CountMissingCLientIdsInTheOriginalTable", "ClientsIdsMissingInTheOriginalTable", "CountOfDifferenceMoreInOriginalTable", "CountOfDifferenceLessInOriginTable", "CountOfAllDifferenceBetweenTables", "CountOfClientDataFullMatch"];
      worksheet.addRow(headerRow);

      const mainHeaderRow = ["ClientId", "AllCountDifferents", "CountDifferentsMore", "CountDifferentsLess", "RowCountMore", "RowCountLess", "PreliminaryIssueMore", "PreliminaryIssueLess", "PreliminaryIssueStatusMore", "PreliminaryIssueStatusLess", "PrefileConsultTableMore", "PrefileConsultTableLess", "DataPageAddressMore", "DataPageAddressLess", "DataPageSequenceValueMore", "DataPageSequenceValueLess", "RequiredServiceMore", "RequiredServiceLess", "PrimaryConditionMore", "PrimaryConditionLess"];
      worksheet.spliceRows(3, 0, mainHeaderRow);

      const headersInfo = ['colors', 'explanationOfColors', 'terms', 'explanationOfTerms'];
      worksheetInfo.addRow(headersInfo);

      const colorNames = {
        'Blue': 'FF554fd5',
        'light red': 'FFE37272',
        'Lime green': 'FF00FF00',
        'Yellow-green': 'FFabf452',
        'Bright red': 'FFFF0000',
        'Very light pink': 'FFFFCCCC',
        'Pink': 'FFFF00FF'
      };


      const dataInfo = [
        [
          'Blue',
          'count of unique client ids in the table',
          'More/Less',
          'the count of values in the original table is greater/less than in the test table'
        ],
        [
          'light red',
          'client id wich data is not full matching',
          'CountClientIdInOriginalTable',
          'the count of client id in original table'
        ],
        [
          'Lime green',
          'when all values matching',
          'CountClientIdInTestTable',
          'the count of client id in test table'
        ],
        [
          'Yellow-green',
          'when all values of this column in both tables match',
          'CountMatchingClientIds',
          'the count of client id wich matched in original and test table'
        ],
        [
          'Bright red',
          'when there are some differents in the row',
          'CountMissingCLientIdsInTheOriginalTable',
          'the count of client id wich missing in original table'
        ],
        [
          'Very light pink',
          'values and there count wich less in original table',
          'ClientsIdsMissingInTheOriginalTable',
          'the name of client id wich missing in original table'
        ],
        [
          'Pink',
          'values and there count wich more in original table',
          'CountOfDifferenceMoreInOriginalTable',
          'the count of all values missing in test table'
        ],
        [
          '',
          '',
          'CountOfDifferenceLessInOriginTable',
          'the count of all values missing in original table'
        ],
        [
          '',
          '',
          'CountOfAllDifferenceBetweenTables',
          'the count of all differents between original and test tables'
        ],
        [
          '',
          '',
          'CountOfClientDataFullMatch',
          'the count of clients whos all values matching in original and test tables'
        ]
      ];

      for (let j = 0; j < dataInfo.length; j++) {
        worksheetInfo.addRow(dataInfo[j]);
      }


      this.setHeadersCellColor(worksheet, 1);
      this.setHeadersCellColor(worksheet, 3);
      this.setInfoHeadersCellColor(worksheetInfo);
      await this.setColorsForInfoCellColA(worksheetInfo, colorNames);


      for (let colIndex = 1; colIndex <= mainHeaderRow.length; colIndex++) {
        const column = worksheet.getColumn(colIndex);
        const maxLength = headerRow[colIndex - 1];
        if (maxLength == undefined) {
          const maxLengthMain = mainHeaderRow[colIndex - 1]
          column.width = maxLengthMain.length + 1;
        } else {
          column.width = maxLength.length + 1;
        }
      }

      const widthArr = [15, 45, 36, 63]

      for (let i = 1; i <= headersInfo.length; i++) {
        const colum = worksheetInfo.getColumn(i)
        colum.width = widthArr[i - 1];
      }

      await workbook.xlsx.writeFile(pathNewTable);
      console.log('Excel file created successfully');
    } catch (error) {
      console.error("ERROR", error);
      throw error
    }
  }

}

export default AnalysTableModel;
