import path from 'path';
import fs from 'fs';

export function createCsvFields(objArray) {
  let str = '';

  for (const index in objArray[0]) {
    str += ','
    str += index;
  }

  return str.slice(1);
}

export function convertToCsv(objArray) {
  let str = '';

  for (let i = 0; i < objArray.length; i++) {
    let line = '';
    for (const index in objArray[i]) {
      if (line != '') line += ','

      if (objArray[i][index] != undefined) {
        line += objArray[i][index];
      }
    }

    str += line + '\r\n';
  }

  return str;
}

export async function getEmployees(sheet, lowerBound, upperBound) {
  const id = await getId();
  const employee = [];
  for (let i = lowerBound; i <= upperBound; i++) {
    const employeeCell = {
      full_name: sheet.getRow(3).getCell(i).value?.toString(),
      gender: sheet.getRow(5).getCell(i).value?.toString(),
      place_of_birth: sheet.getRow(6).getCell(i).value?.toString(),
      date_of_birth: sheet.getRow(7).getCell(i).value?.toString(),
      marital_status: sheet.getRow(8).getCell(i).value?.toString(),
      religion: sheet.getRow(9).getCell(i).value?.toString(),
      foreign_labor: sheet.getRow(10).getCell(i).value?.toString(),
      nationality: sheet.getRow(11).getCell(i).value?.toString(),
      'kartu_tanda_penduduk_(ktp)': sheet.getRow(14).getCell(i).value?.toString(),
      ktp_expiry_date: sheet.getRow(15).getCell(i).value?.toString(),
      passport_number: sheet.getRow(16).getCell(i).value?.toString(),
      passport_expiry_date: sheet.getRow(17).getCell(i).value?.toString(),
      sim_a: sheet.getRow(18).getCell(i).value?.toString(),
      sim_a_expiry_date: sheet.getRow(19).getCell(i).value?.toString(),
      sim_c: sheet.getRow(20).getCell(i).value?.toString(),
      sim_c_expiry_date: sheet.getRow(21).getCell(i).value?.toString(),
      employee_id: sheet.getRow(2).getCell(i).value?.toString() + `-id${id}`,
      date_of_joining: sheet.getRow(28).getCell(i).value?.toString(),
      employment_status_number: sheet.getRow(31).getCell(i).value ? sheet.getRow(31).getCell(i).value.toString() + `/id${id}` : '',
      employment_status_type: sheet.getRow(32).getCell(i).value?.toString(),
      employee_type: sheet.getRow(33).getCell(i).value?.toString(),
      effective_date: sheet.getRow(34).getCell(i).value?.toString(),
      contract_start_date: sheet.getRow(36).getCell(i).value?.toString(),
      contract_end_date: sheet.getRow(37).getCell(i).value?.toString(),
      date_of_confirmation: sheet.getRow(38).getCell(i).value?.toString(),
      department: sheet.getRow(39).getCell(i).value?.toString(),
      designations: sheet.getRow(40).getCell(i).value?.toString(),
      job_level: sheet.getRow(41).getCell(i).value?.toString(),
      office_location: sheet.getRow(42).getCell(i).value?.toString(),
      cost_center: sheet.getRow(43).getCell(i).value?.toString(),
      direct_manager_employee_id: sheet.getRow(44).getCell(i).value ? sheet.getRow(44).getCell(i).value.toString() + `-id${id}` : '',
      bpjs_ketenagakerjaan_branch_office: sheet.getRow(45).getCell(i).value?.toString(),
      bpjs_ketenagakerjaan_number: sheet.getRow(46).getCell(i).value?.toString(),
      bpjs_ketenagakerjaan_start_date: sheet.getRow(47).getCell(i).value?.toString(),
      bpjs_ketenagakerjaan_end_date: sheet.getRow(48).getCell(i).value?.toString(),
      bpjs_ketenagakerjaan_template: sheet.getRow(49).getCell(i).value?.toString(),
      bpjs_kesehatan_branch_office: sheet.getRow(50).getCell(i).value?.toString(),
      bpjs_kesehatan_number: sheet.getRow(51).getCell(i).value?.toString(),
      bpjs_kesehatan_start_date: sheet.getRow(52).getCell(i).value?.toString(),
      bpjs_kesehatan_end_date: sheet.getRow(53).getCell(i).value?.toString(),
      bpjs_kesehatan_template: sheet.getRow(54).getCell(i).value?.toString(),
      npwp_effective_period: sheet.getRow(55).getCell(i).value?.toString(),
      employee_tax_object: sheet.getRow(56).getCell(i).value?.toString(),
      tax_calculation_method: sheet.getRow(57).getCell(i).value?.toString(),
      foreign_tax_subject: sheet.getRow(58).getCell(i).value?.toString(),
      more_than_one_employer: sheet.getRow(59).getCell(i).value?.toString(),
      'nomor_pokok_wajib_pajak_(npwp)': sheet.getRow(60).getCell(i).value?.toString(),
      address_in_npwp: sheet.getRow(61).getCell(i).value?.toString(),
      ptkp_status: sheet.getRow(62).getCell(i).value?.toString(),
      kpp_code: sheet.getRow(63).getCell(i).value?.toString(),
      previous_net_income: sheet.getRow(64).getCell(i).value?.toString(),
      previous_paid_pph: sheet.getRow(65).getCell(i).value?.toString(),
      payment_method: sheet.getRow(66).getCell(i).value?.toString(),
      'percentage/amount': sheet.getRow(73).getCell(i).value?.toString(),
      distribution_method: sheet.getRow(67).getCell(i).value?.toString(),
      'bank_name_(indonesia)': sheet.getRow(68).getCell(i).value?.toString(),
      bank_branch: sheet.getRow(69).getCell(i).value?.toString(),
      bank_account_name: sheet.getRow(70).getCell(i).value?.toString(),
      bank_account: sheet.getRow(71).getCell(i).value?.toString(),
      bank_priority: sheet.getRow(72).getCell(i).value?.toString(),
      company_email_id: sheet.getRow(4).getCell(i).value?.toString(),
      employee_separation_reason: sheet.getRow(79).getCell(i).value?.toString(),
      admin_deactivate_reason: sheet.getRow(79).getCell(i).value?.toString(),
      date_of_resignation: sheet.getRow(75).getCell(i).value?.toString(),
      date_of_exit: sheet.getRow(76).getCell(i).value?.toString(),
      experience_in_current_role: sheet.getRow(78).getCell(i).value?.toString(),
      payroll_method: sheet.getRow(80).getCell(i).value?.toString(),
      npwp_end_date: sheet.getRow(81).getCell(i).value?.toString(),
      latest_modified: sheet.getRow(82).getCell(i).value?.toString(),
      latest_modified_any_attribute: sheet.getRow(82).getCell(i).value?.toString(),
      latest_modified_timestamp: sheet.getRow(82).getCell(i).value?.toString()
    };
    employee.push(employeeCell);
  }
  return employee;
}

export async function writeEmployees(employee, phase) {
  const writePath = path.join(__dirname, `/output-${phase}.csv`);
  await fs.writeFile(writePath, `${createCsvFields(employee)}\n${convertToCsv(employee)}`, function (err) {
    if (err) {
      console.log('An error occured while writing JSON Object to File.');
      return console.log(err);
    }

    console.log('CSV file has been saved.');
  });
}

export async function duplicateEmployees(employees, multiplier) {
  let duplicatedEmployees = [];
  for (let i = 0; i < multiplier; i++) {
    for (const employee of employees) {
      const modifiedEmployees = {
        ...employee,
        employee_id: employee.employee_id + `-${i}`,
        direct_manager_employee_id: employee.direct_manager_employee_id ? employee.direct_manager_employee_id + `-${i}` : null,
        employment_status_number: employee.employee_status_number + `-${i}`
      };
      duplicatedEmployees.push(modifiedEmployees);
    }
  }
  return duplicatedEmployees;
}

export async function getDependents(sheet, lowerBound, upperBound) {
  const id = await getId();
  const dependents = [];
  for (let i = lowerBound; i <= upperBound; i++) {
    const fullName = sheet.getRow(22).getCell(i).value?.toString();
    const { firstName, middleName, lastName } = getName(fullName);

    const dependentCell = {
      employee_id: sheet.getRow(2).getCell(i).value?.toString() + `-id${id}`,
      first_name: firstName,
      middle_name: middleName,
      last_name: lastName,
      relation_type: sheet.getRow(23).getCell(i).value?.toString(),
      gender: sheet.getRow(26).getCell(i).value?.toString(),
      date_of_death: sheet.getRow(27).getCell(i).value?.toString()
    };
    dependents.push(dependentCell);
  }
  return dependents;
}

export async function duplicateDependents(dependents, multiplier) {
  let duplicatedEmployees = [];
  for (let i = 0; i < multiplier; i++) {
    for (const dependent of dependents) {
      const modifiedDependents = {
        ...dependent,
        employee_id: dependent.employee_id + `-${i}`
      };
      duplicatedEmployees.push(modifiedDependents);
    }
  }
  return duplicatedEmployees;
}

function getName(fullName: string) {
  let firstName = '';
  let middleName = '';
  let lastName = '';
  console.log(fullName);
  if (!!fullName) {
    const splittedFullName = fullName.split(' ');
    const splittedFullNameLength = splittedFullName.length;
    if (splittedFullNameLength == 1) {
      firstName = splittedFullName[0];
    } else if (splittedFullNameLength == 2) {
      firstName = splittedFullName[0];
      lastName = splittedFullName[1];
    } else if (splittedFullNameLength > 2) {
      firstName = splittedFullName[0];
      lastName = splittedFullName[splittedFullNameLength - 1];
      middleName = splittedFullName.splice(1, splittedFullNameLength - 2).join(' ');
    }
  }

  return {
    firstName,
    middleName,
    lastName
  };
}

export async function writeDependents(dependents, phase) {
  for (const dependentData of dependents) {
    const writePath = path.join(__dirname, `/dependents-data/${phase}/dependents-${dependentData['employee_id']}.json`);
    await fs.writeFile(writePath, JSON.stringify(dependentData['first_name'] ? [dependentData] : []), function (err) {
      if (err) {
        console.log('An error occured while writing JSON Object to File.');
        return console.log(err);
      }

      console.log('JSON file has been saved.');
    });
  }
}

async function getId() {
  const config = require('./config');
  return config['id'];
}

export async function writeId() {
  const id = await getId();
  const writePath = path.join(__dirname, '/config.json');
  await fs.writeFile(writePath, JSON.stringify({ id: id + 1 }), function (err) {
    if (err) {
      console.log('An error occured while writing JSON Object to File.');
      return console.log(err);
    }

    console.log('JSON file has been saved.');
  });
}
