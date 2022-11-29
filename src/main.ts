import * as excel from 'exceljs';
import * as path from 'path';

import { Constant } from './constant';
import {
  duplicateDependents,
  duplicateEmployees,
  getDependents,
  getEmployees,
  writeDependents,
  writeDependentsCsv,
  writeEmployees,
  writeId
} from './helper';

async function processExcel(filename) {
  let workBook = new excel.Workbook();
  await workBook.xlsx.readFile(filename);

  let sheet = workBook.getWorksheet(2);

  const phase1lowerBound = 6;
  const phase1upperBound = 16;

  const phase2lowerBound = 19;
  const phase2upperBound = 27;

  const phase3lowerBound = 29;
  const phase3upperBound = 35;

  const phase4lowerBound = 37;
  const phase4upperBound = 43;

  await processData(Constant.PHASE_1, sheet, phase1lowerBound, phase1upperBound);
  //await processAndDuplicateData(Constant.PHASE_1, sheet, phase1lowerBound, phase1upperBound, 2);
  await processData(Constant.PHASE_2, sheet, phase2lowerBound, phase2upperBound);
  await processData(Constant.PHASE_3, sheet, phase3lowerBound, phase3upperBound);
  await processData(Constant.PHASE_4, sheet, phase4lowerBound, phase4upperBound);

  await writeId();
}

async function processData(phase, sheet, lowerBound, upperBound) {
  const employees = await getEmployees(sheet, lowerBound, upperBound);
  const dependents = await getDependents(sheet, lowerBound, upperBound);
  await writeEmployees(employees, phase);

  await writeDependentsCsv(dependents, phase);
}

async function processAndDuplicateData(phase, sheet, lowerBound, upperBound, multiplier) {
  const employees = await getEmployees(sheet, lowerBound, upperBound);
  const dependents = await getDependents(sheet, lowerBound, upperBound);

  const duplicatedEmployees = await duplicateEmployees(employees, multiplier);
  const duplicatedDependents = await duplicateDependents(dependents, multiplier);

  await writeEmployees(duplicatedEmployees, phase);
  await writeDependents(duplicatedDependents, phase);
}

// @ts-ignore
processExcel(path.join(__dirname, '/test.xlsx'));
