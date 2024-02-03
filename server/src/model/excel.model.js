"use strict"
import ExcelJS from "exceljs";

class ExcelTablesModel {

  async getAllClientsId(worksheetTest) { // we get all unique client ids from table
    try {

      const allClientsId = {};

      worksheetTest.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        if (rowNumber > 4) {
          const cellValue = row.getCell(4).value;
          allClientsId[cellValue] = true;
        }
      });

      const clientsIdArray = Object.keys(allClientsId);

      return clientsIdArray; // fn return array of client ids 

    } catch (err) {
      console.error("ERROR ", err);
      throw err
    }
  }

  async getClientRowsArray(clientId, worksheet) { // we get an array , and push there every row numbers, where this client has data
    try {
      const clientRows = [];

      worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        const cellValue = row.getCell(4).value; // get value from column 4 for each row

        if (cellValue === clientId) {
          clientRows.push(rowNumber);
        }
      });

      return clientRows;

    } catch (err) {
      console.error('Error', err);
      throw err;
    }
  }

  async getClientData(worksheet, clientRows) { // get all data from column 5 to column 11 from given rows of array clientRows

    try {
      // we create variables, in each of which we will store the entire value of a certain column, where there will be a value and the number of duplicates

      const rowCount = clientRows.length;


      let colE = {};
      let colF = {};
      let colG = {};
      let colH = {};
      let colI = {};
      let colJ = {};
      let colK = {};

      for (let i = 0; i < clientRows.length; i++) {
        let rowNum = clientRows[i];
        // in each circle we take the entire value of columns E, F, G, H, I, J, K of the given row
        let cellE = worksheet.getCell(rowNum, 5).value;
        let cellF = worksheet.getCell(rowNum, 6).value;
        let cellG = worksheet.getCell(rowNum, 7).value;
        let cellH = worksheet.getCell(rowNum, 8).value;
        let cellI = worksheet.getCell(rowNum, 9).value;
        let cellJ = worksheet.getCell(rowNum, 10).value;
        let cellK = worksheet.getCell(rowNum, 11).value;

        // if cell value === null, rename as 'empty cell' , else don't change anything
        cellE = cellE === null ? 'empty cell' : cellE;
        cellF = cellF === null ? 'empty cell' : cellF;
        cellG = cellG === null ? 'empty cell' : cellG;
        cellH = cellH === null ? 'empty cell' : cellH;
        cellI = cellI === null ? 'empty cell' : cellI;
        cellJ = cellJ === null ? 'empty cell' : cellJ;
        cellK = cellK === null ? 'empty cell' : cellK;


        // if this field already exists, add + 1, else add this field with value = 1
        colE[cellE] = cellE in colE ? colE[cellE] += 1 : 1;
        colF[cellF] = cellF in colF ? colF[cellF] += 1 : 1;
        colG[cellG] = cellG in colG ? colG[cellG] += 1 : 1;
        colH[cellH] = cellH in colH ? colH[cellH] += 1 : 1;
        colI[cellI] = cellI in colI ? colI[cellI] += 1 : 1;
        colJ[cellJ] = cellJ in colJ ? colJ[cellJ] += 1 : 1;
        colK[cellK] = cellK in colK ? colK[cellK] += 1 : 1;

      }

      const clientData = { rowCount: rowCount, PreliminaryIssue: colE, PreliminaryIssueStatus: colF, PrefileConsultTable: colG, DataPageAddress: colH, DataPageSequenceValue: colI, RequiredService: colJ, PrimaryCondition: colK };

      return clientData

    } catch (err) {
      console.error("ERROR ", err);
      throw err
    }
  }

  async isAllIdsInOrigSheet(clientIdTestArr, clientIdOrigArr) { // we check are all clients from test table in original table and return array with client ids wich are not in origin table
    try {
      const foreignIds = [];

      for (let i = 0; i < clientIdTestArr.length; i++) {
        let clientIdTest = clientIdTestArr[i];
        let res = clientIdOrigArr.some(clientIdOrig => clientIdOrig === clientIdTest);

        if (res === false) {
          foreignIds.push(clientIdTest);
        };
      };
      let result = foreignIds.length > 0 ? false : true;
      return { check: result, arrExceptions: foreignIds }

    } catch (err) {
      console.error("ERROR ", err);
      throw err
    }
  }

  async getEachTableClientsAllData(worksheet, clientIds) { // we get all data about all clientIds from given table
    const allClientsAllData = {};


    for (let i = 0; i < clientIds.length; i++) {
      let rowsArr = await this.getClientRowsArray(clientIds[i], worksheet);
      let clientData = await this.getClientData(worksheet, rowsArr);
      allClientsAllData[clientIds[i]] = clientData; // {clientId : { rowCount: rowCount, PreliminaryIssue: colE, PreliminaryIssueStatus: colF, PrefileConsultTable: colG, DataPageAddress: colH, DataPageSequenceValue: colI, RequiredService: colJ, PrimaryCondition: colK }}
    }
    console.log("*** getEachTableClientsAllData has done ***");
    return allClientsAllData;
  }

  async getAllClientsAllData(pathOriginTable = 'server/data/originalTable.xlsx', pathTestTable = 'server/data/testTable.xlsx') {// ***
    try {

      const workbookOrigin = new ExcelJS.Workbook();
      await workbookOrigin.xlsx.readFile(pathOriginTable);
      const worksheetOrigin = workbookOrigin.getWorksheet(1);

      const workbookTest = new ExcelJS.Workbook();
      await workbookTest.xlsx.readFile(pathTestTable);
      const worksheetTest = workbookTest.getWorksheet(1);

      const arrOrigAllIds = await this.getAllClientsId(worksheetOrigin); // get all ids from original table
      const arrTestAllIds = await this.getAllClientsId(worksheetTest); // get all ids from test table
      const result = await this.isAllIdsInOrigSheet(arrTestAllIds, arrOrigAllIds);
      const arrayExceptions = result.arrExceptions;
      const countClientIdInOriginalTable = arrOrigAllIds.length;
      const countClientIdInTestTable = arrTestAllIds.length;
      const countMissingCLientIdsInTheOriginalTable = arrayExceptions.length;
      const countMatchingClientIds = countClientIdInTestTable - countMissingCLientIdsInTheOriginalTable;
      let resExcept = arrayExceptions.length === 0 ? 0 : arrayExceptions;

      const valuesForSheet2 = [countClientIdInOriginalTable, countClientIdInTestTable, countMatchingClientIds, countMissingCLientIdsInTheOriginalTable, resExcept];

      console.log(result);
      console.log(result.check === false ? 'foreignIds is/are ' + result.arrExceptions : result.check);


      const allOriginClientData = await this.getEachTableClientsAllData(worksheetOrigin, arrTestAllIds); // { }
      console.log("allOriginClientData  received");

      const allTestClientData = await this.getEachTableClientsAllData(worksheetTest, arrTestAllIds);
      console.log("allTestClientData  received");



      return { originTable: allOriginClientData, testTable: allTestClientData, exceptions: arrayExceptions, generalResults: valuesForSheet2 }
    } catch (error) {
      console.error("ERROR", error.message);
      throw error;
    }
  }

  //-----------------------------------------------------------------------------------------------------------

  compareSameClientData(clientTestAllData, clientOrigAllData) { // this function compares the all values in same columns of two clients of the same name

    const objDifferentsLess = {};
    const objDifferentsMore = {};
    const arrKeys = ["PreliminaryIssue", "PreliminaryIssueStatus", "PrefileConsultTable", "DataPageAddress", "DataPageSequenceValue", "RequiredService", "PrimaryCondition"];         //['colE', 'colF', 'colG', 'colH', 'colI', 'colJ', 'colK'];


    const rowCountRes = clientOrigAllData.rowCount - clientTestAllData.rowCount;

    if (rowCountRes > 0) {  // if rowCountRes true value create property rowCount in objDifferents and assigned value rowCountRes
      objDifferentsMore.rowCount = rowCountRes;
    } else if (rowCountRes < 0) {
      objDifferentsLess.rowCount = rowCountRes;
    }

    for (let i = 0; i < arrKeys.length; i++) { // ['colE', 'colF', 'colG', 'colH', 'colI', 'colJ', 'colK']
      let colName = arrKeys[i];                //       first loop let colName === colE
      let origCol = clientOrigAllData[colName]; // we get property clientOrigAllData[colE]  PreliminaryIssue: {ke1:5,key:48484}
      let testCol = clientTestAllData[colName]; // we get property clientTestAllData[colE]


      let objColKeys = { ...origCol, ...testCol }; // made an object with unique keys from two objects

      for (const colKey in objColKeys) {
        let valOrig = origCol[colKey] ? origCol[colKey] : 0;
        let valTest = testCol[colKey] ? testCol[colKey] : 0;

        let res = valOrig - valTest;
        if (res > 0) {
          objDifferentsMore[colName] = { [colKey]: res };
        } else if (res < 0) {
          objDifferentsLess[colName] = { [colKey]: res };
        }
      }
    }
    return { more: objDifferentsMore, less: objDifferentsLess };
  }


  async compareAllClientsData() {// *** this function compare all data for same client from original and test tables
    const data = await this.getAllClientsAllData();
    const objResultDifference = {};

    const originData = data.originTable;
    const testData = data.testTable;
    const exceptions = data.exceptions;
    const generalResults = data.generalResults;


    if (exceptions.length === 0) { // if there are no exceptions
      for (let clientId in testData) {
        const origClientData = originData[clientId];
        const testClientData = testData[clientId];

        let res = this.compareSameClientData(testClientData, origClientData);
        objResultDifference[clientId] = res;
      }
    } else {
      for (let clientId in testData) {
        let blockedId = exceptions.some(val => val === clientId);
        if (!blockedId) {
          const origClientData = originData[clientId];
          const testClientData = testData[clientId];

          let res = this.compareSameClientData(testClientData, origClientData);
          objResultDifference[clientId] = res; // clientId :{ more :{ colName1 :{ properties : values},colName2 :{ properties : values} .... }, less :{ colName1 :{ properties : values},colName2 :{ properties : values} .... }}
        }
      }
    }



    console.log('***--------------------compareAllClientsData received--------------------***');

    return [objResultDifference, generalResults]
  }


  //------------------------------------------------------------------------------------------------------------


  setRowColor(worksheet, rowIndex, color = 'FF00FF00') { // set color for certain row
    const row = worksheet.getRow(rowIndex);
    const actualCellCount = row.actualCellCount;

    for (let colNumber = 1; colNumber <= actualCellCount; colNumber++) {
      const cell = row.getCell(colNumber);
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: color }
      };
      cell.border = { bottom: { style: 'thin', color: { argb: '000000' } }, right: { style: 'thin', color: { argb: '000000' } } };
    }
  }

  setCellsColor(worksheet, rowNum) { // set color for certain cells 

    const cellColA = worksheet.getCell('A' + rowNum);

    cellColA.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFE37272' }
    };

    cellColA.border = { bottom: { style: 'thin', color: { argb: '000000' } }, right: { style: 'thin', color: { argb: '000000' } } };

    const cellColB = worksheet.getCell('B' + rowNum);

    cellColB.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFF0000' }
    };

    cellColB.border = { bottom: { style: 'thin', color: { argb: '000000' } }, right: { style: 'thin', color: { argb: '000000' } } };

    const targetRow = worksheet.getRow(rowNum);


    for (let colNumber = 3; colNumber <= targetRow.actualCellCount; colNumber += 2) {
      const cell = targetRow.getCell(colNumber);
      const val = cell.value;

      cell.border = { bottom: { style: 'thin', color: { argb: '000000' } }, right: { style: 'thin', color: { argb: '000000' } } };

      if (val === 0) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'ffabf452' }
        }
      } else {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFF00FF' }
        }
      }
    }

    for (let colNumber = 4; colNumber <= targetRow.actualCellCount; colNumber += 2) {
      const cell = targetRow.getCell(colNumber);
      const val = cell.value;

      cell.border = { bottom: { style: 'thin', color: { argb: '000000' } }, right: { style: 'thin', color: { argb: '000000' } } };

      if (val === 0) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'ffabf452' }
        }
      } else {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFCCCC' }
        }
      }
    }

  }

  setCellsColorForGeneralResults(worksheet, rowNum = 2) { // set color for certain cells for general results 
    const targetRow = worksheet.getRow(rowNum);
    const lastColIndex = targetRow.actualCellCount;
    const cellColA = worksheet.getCell('A' + rowNum);
    const cellColB = worksheet.getCell('B' + rowNum);
    const cellColC = worksheet.getCell('C' + rowNum);
    const lastCell = worksheet.getCell('I' + rowNum);

    cellColA.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'ff554fd5' }
    };

    cellColA.border = { bottom: { style: 'medium', color: { argb: '000000' } }, right: { style: 'thin', color: { argb: '000000' } } };


    cellColB.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'ff554fd5' }
    };

    cellColB.border = { bottom: { style: 'medium', color: { argb: '000000' } }, right: { style: 'thin', color: { argb: '000000' } } };

    if (cellColB.value === cellColC.value) {
      cellColC.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'ffabf452' }
      };
    } else {
      cellColC.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFF0000' }
      };
    }

    cellColC.border = { bottom: { style: 'medium', color: { argb: '000000' } }, right: { style: 'thin', color: { argb: '000000' } } };



    lastCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'ffabf452' }
    };

    lastCell.border = { bottom: { style: 'medium', color: { argb: '000000' } }, right: { style: 'thin', color: { argb: '000000' } } };


    for (let colNumber = 4; colNumber < lastColIndex; colNumber++) {
      const cell = targetRow.getCell(colNumber);
      const val = cell.value;

      cell.border = { bottom: { style: 'medium', color: { argb: '000000' } }, right: { style: 'thin', color: { argb: '000000' } } };

      if (val === 0) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'ffabf452' }
        }
      } else if (val > 0) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFF00FF' }
        }
      } else {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFCCCC' }
        }
      }
    }




  }

  getCountDifferentsForEachClient(arr) { // get the number of differences between the test and original tables for each client

    let countDifferentsMore = arr[0] //rowCountMore {val1:number,val2:number};
    let countDifferentsLess = arr[1] //rowCountLess {val1:number,val2:number};

    for (let i = 2; i < arr.length; i += 2) {
      const element = arr[i];

      let values = Object.values(element);
      let sum = values.reduce((accumulator, currentValue) => {
        return accumulator + currentValue;
      }, 0);

      countDifferentsMore += sum;

    }

    for (let j = 3; j < arr.length; j += 2) {
      const element = arr[j];

      let values = Object.values(element);
      let sum = values.reduce((accumulator, currentValue) => {
        return accumulator + currentValue;
      }, 0);

      countDifferentsLess += sum;
    }

    let allCountDifferents = countDifferentsMore + -countDifferentsLess;
    return [allCountDifferents, countDifferentsMore, countDifferentsLess]
  }


  async analyzeDataAndAddToAnalysTable(pathAnalysTable = 'server/data/Analys_Table.xlsx') {// ***
    try {

      // create AnalysTable before use this function

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(pathAnalysTable);
      const worksheet = workbook.getWorksheet('Sheet1');

      const data = await this.compareAllClientsData();
      const comparedData = data[0];  // {clientId : { more:{rowCount:number,key2:object},less:{rowCount:number,key2:object} }, clientId2 : { more:{},less:{} }}
      const generalResults = data[1];

      const colNames = ["rowCount", "PreliminaryIssue", "PreliminaryIssueStatus", "PrefileConsultTable", "DataPageAddress", "DataPageSequenceValue", "RequiredService", "PrimaryCondition"];
      let countOfClientDataFullMatch = 0;
      let countOfAllDiffMoreInOrigTable = 0;
      let countOfAllDiffLessInOrigTable = 0;
      let countOfAllDifferenceBetweenTables = 0;

      for (const clientId in comparedData) {
        let resArr = []; // [rowCountMore, rowCountLess, preliminaryIssueMore, preliminaryIssueLess.....]

        for (let i = 0; i < colNames.length; i++) {
          let colName = colNames[i];
          let valMore = comparedData[clientId].more?.[colName] || 0;
          let valLess = comparedData[clientId].less?.[colName] || 0;

          resArr.push(valMore, valLess);

        }

        let countDifferents = this.getCountDifferentsForEachClient(resArr);

        countOfAllDifferenceBetweenTables += countDifferents[0];
        countOfAllDiffMoreInOrigTable += countDifferents[1];
        countOfAllDiffLessInOrigTable += countDifferents[2];

        let result = [clientId, ...countDifferents, ...resArr];
        worksheet.addRow(result);

        let checkAreEqual = result[1] === 0;
        if (checkAreEqual) {
          // set row color green if condition is true
          countOfClientDataFullMatch++;
          this.setRowColor(worksheet, worksheet.actualRowCount + 1);
        } else {
          this.setCellsColor(worksheet, worksheet.actualRowCount + 1);
        }
      }

      console.log('countOfClientDataFullMatch  = ', countOfClientDataFullMatch);

      const sheet2AllData = [...generalResults, countOfAllDiffMoreInOrigTable, countOfAllDiffLessInOrigTable, countOfAllDifferenceBetweenTables, countOfClientDataFullMatch];

      worksheet.spliceRows(2, 0, sheet2AllData);
      this.setCellsColorForGeneralResults(worksheet);

      await workbook.xlsx.writeFile(pathAnalysTable);

    } catch (error) {
      console.error("ERROR", error);
      throw error;
    }
  }
}


export default ExcelTablesModel;