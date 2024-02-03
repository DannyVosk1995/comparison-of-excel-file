"use strict"
import ExcelTablesController from "./controller/excel.controller.js";
const controller = new ExcelTablesController();

async function runProject() {
  await controller.projectResult();
}

runProject();
