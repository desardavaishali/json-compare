const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

function compareJSON(obj1, obj2, filteredFields) {
  const differences = getDifferences(obj1, obj2, '', filteredFields);
  return differences;
}

function getDifferences(obj1, obj2, path, filteredFields) {
  const differences = [];

  const keys = new Set([...Object.keys(obj1), ...Object.keys(obj2)]);
  for (const key of keys) {
    const currentPath = path ? `${path}.${key}` : key;
    const value1 = obj1[key];
    const value2 = obj2[key];

    if (!compareFilteredFields(currentPath, filteredFields)) {
      if (isObject(value1) && isObject(value2)) {
        const nestedDifferences = getDifferences(value1, value2, currentPath, filteredFields);
        differences.push(...nestedDifferences);
      } else if (isArray(value1) && isArray(value2)) {
        const nestedDifferences = getArrayDifferences(value1, value2, currentPath, filteredFields);
        differences.push(...nestedDifferences);
      } else if (!isEqual(value1, value2)) {
        differences.push({
          Field: currentPath,
          Value: value2,
          ComparedField: currentPath,
          ComparedValue: value1,
        });
      }
    }
  }

  return differences;
}

function getArrayDifferences(arr1, arr2, path, filteredFields) {
  const maxLength = Math.max(arr1.length, arr2.length);
  const differences = [];

  for (let i = 0; i < maxLength; i++) {
    const currentPath = `${path}[${i}]`;
    const value1 = arr1[i];
    const value2 = arr2[i];

    if (!compareFilteredFields(currentPath, filteredFields)) {
      if (isObject(value1) && isObject(value2)) {
        const nestedDifferences = getDifferences(value1, value2, currentPath, filteredFields);
        differences.push(...nestedDifferences);
      } else if (isArray(value1) && isArray(value2)) {
        const nestedDifferences = getArrayDifferences(value1, value2, currentPath, filteredFields);
        differences.push(...nestedDifferences);
      } else if (!isEqual(value1, value2)) {
        differences.push({
          Field: currentPath,
          Value: value2,
          ComparedField: currentPath,
          ComparedValue: value1,
        });
      }
    }
  }

  return differences;
}

function isObject(value) {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function isArray(value) {
  return Array.isArray(value);
}

function isEqual(value1, value2) {
  if (isObject(value1) || isArray(value1)) {
    return JSON.stringify(value1) === JSON.stringify(value2);
  }
  return value1 === value2;
}

function compareFilteredFields(field, filteredFields) {
  for (const filteredField of filteredFields) {
    if (field.includes(filteredField)) {
      return true;
    }
  }
  return false;
}

function generateExcel(differences) {
  const worksheetData = differences.map(diff => [
    diff.Field,
    JSON.stringify(diff.Value),
    diff.ComparedField,
    JSON.stringify(diff.ComparedValue),
  ]);

  const headers = ['Field', 'Value', 'ComparedField', 'ComparedValue'];
  const data = [headers, ...worksheetData];

  const worksheet = XLSX.utils.aoa_to_sheet(data);

  // Cell Formatting
  const cellStyles = {
    header: {
      fill: {
        fgColor: { rgb: 'EFEFEF' },
      },
      font: {
        bold: true,
      },
    },
    value: {
      fill: {
        fgColor: { rgb: 'FFFFFF' },
      },
    },
  };

  // Apply formatting to header cells
  const headerRange = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: 0, c: headers.length - 1 } });
  for (let i = 0; i < headers.length; i++) {
    const cellAddress = XLSX.utils.encode_cell({ r: 0, c: i });
    const cell = worksheet[cellAddress] || {};
    cell.s = cellStyles.header;
    worksheet[cellAddress] = cell;
  }

  // Apply formatting to value cells
  const dataRange = XLSX.utils.encode_range({ s: { r: 1, c: 0 }, e: { r: data.length - 1, c: headers.length - 1 } });
  for (let row = 1; row < data.length; row++) {
    for (let col = 0; col < headers.length; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
      const cell = worksheet[cellAddress] || {};
      cell.s = cellStyles.value;
      worksheet[cellAddress] = cell;
    }
  }

  const workbook = XLSX.utils.book_new();
  const sheetName = 'Differences';
  XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);

  return workbook;
}

function writeExcelFile(workbook, filePath) {
  const fileExtension = path.extname(filePath);
  const fileName = path.basename(filePath, fileExtension);
  const excelFilePath = `${fileName}.xlsx`;
  XLSX.writeFile(workbook, excelFilePath);
  console.log(`Excel file generated: ${excelFilePath}`);
}

function readJSONFile(filePath) {
  try {
    const fileContent = fs.readFileSync(filePath, 'utf8');
    return JSON.parse(fileContent);
  } catch (error) {
    console.error(`Error reading JSON file: ${filePath}`);
    console.error(error.message);
    process.exit(1);
  }
}

function getFilteredFields(filePath) {
  try {
    const filteredFieldsText = fs.readFileSync(filePath, 'utf8');
    return filteredFieldsText.split('\n').map(field => field.trim());
  } catch (error) {
    console.error(`Error reading filtered fields file: ${filePath}`);
    console.error(error.message);
    process.exit(1);
  }
}

function validateFilePath(filePath) {
  if (!fs.existsSync(filePath)) {
    console.error(`File not found: ${filePath}`);
    process.exit(1);
  }
}

function validateArguments() {
  if (process.argv.length < 4 || process.argv.length > 5) {
    console.error('Invalid number of arguments.');
    console.error('Usage: node json-compare-excel.js <file1> <file2> [filteredFieldsFile]');
    process.exit(1);
  }
}

function run() {
  validateArguments();

  const file1 = process.argv[2];
  const file2 = process.argv[3];
  const filteredFieldsFile = process.argv[4];

  validateFilePath(file1);
  validateFilePath(file2);

  const obj1 = readJSONFile(file1);
  const obj2 = readJSONFile(file2);
  const filteredFields = filteredFieldsFile ? getFilteredFields(filteredFieldsFile) : [];

  const diff = compareJSON(obj1, obj2, filteredFields);
  const workbook = generateExcel(diff);
  writeExcelFile(workbook, file2);
}

run();
//Final cmd line execution