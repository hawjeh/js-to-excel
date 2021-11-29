import { ParseTableToExcel, ParseJsonToExcel } from './worker.js';
import Blob from 'cross-blob';
import saveAs from "file-saver";
import ExcelJS from 'exceljs';

const defaultOptions = {
    name: "export.xlsx",
    saveLocal: false,
    sheet: {
        name: "Sheet 1"
    }
};

const initWorkbook = () => {
    let wb = new ExcelJS.Workbook();
    return wb;
}

const initWorkSheet = (wb, sheetName) => {
    let ws = wb.addWorksheet(sheetName);
    return ws;
}

const buildSheet = (isTable, wb, obj, options) => {
    let ws = initWorkSheet(wb, options.sheet.name);
    if (isTable) {
        ws = ParseTableToExcel(ws, obj);
    } else {
        const parsedJson = typeof obj === 'string' ? JSON.parse(obj) : obj;
        parsedJson.options = { ...defaultOptions, ...parsedJson.options };
        ws = ParseJsonToExcel(ws, parsedJson);
    }
}

const save = (wb, fileName, saveLocal) => {
    if (saveLocal) {
        wb.xlsx.writeFile(fileName);
    } else {
        wb.xlsx.writeBuffer().then(function (buffer) {
            saveAs(
                new Blob([buffer], { type: "application/octet-stream" }),
                fileName
            );
        });
    }
}

const exportExcel = (isTable, obj, newOptions = {}) => {
    const options = { ...defaultOptions, ...newOptions };
    let wb = initWorkbook();
    buildSheet(isTable, wb, obj, options);
    save(wb, options.name, options.saveLocal);
}

const exportHtmlToExcel = (table, newOptions = {}) => {
    exportExcel(true, table, newOptions);
}

const exportJsonToExcel = (json, newOptions = {}) => {
    exportExcel(false, json, newOptions);
}

export {
    exportHtmlToExcel,
    exportJsonToExcel
};

// Testing Local
// exportJsonToExcel('{"options":{"showHeader":true,"globalStyles":{},"headerStyles":{},"rowStyles":[]},"keys": ["Header1","Header2","Header3"],"rows":[{"Header1":"Val1","Header2":"Val2","Header3":"Val3"},{"Header1":"Val1","Header2":"Val2","Header3":"Val3"},{"Header1":"Val1","Header2":"Val2","Header3":"Val3"},{"Header1":"Val1","Header2":"Val2","Header3":"Val3"},{"Header1":"Val1","Header2":"Val2","Header3":"Val3"},{"Header1":"Val1","Header2":"Val2","Header3":"Val3"}]}', {
//     name: 'C:\\export.xlsx',
//     saveLocal: true
// })