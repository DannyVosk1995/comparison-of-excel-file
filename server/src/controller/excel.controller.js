"use strict"
import ExcelTablesModel from "../model/excel.model.js";
const excelTablemodels = new ExcelTablesModel();
import AnalysTableModel from "../model/analysTable.model.js";
const analysTableModel = new AnalysTableModel();

class ExcelTablesController {

  async projectResult() {
    await analysTableModel.createAnalysTable();
    await excelTablemodels.analyzeDataAndAddToAnalysTable();
  }
}

export default ExcelTablesController;