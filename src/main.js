import { ParseTableToExcel, ParseJsonToExcel } from './worker.js';
import Blob from 'cross-blob';
import saveAs from "file-saver";
import ExcelJS from 'exceljs';

const defaultOptions = {
    name: "export.xlsx",
    saveLocal: false,
    autoStyle: false,
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

const htmlToSheet = (wb, table, options) => {
    let ws = initWorkSheet(wb, options.sheet.name);
    ParseTableToExcel(ws, table);
    return wb;
}

const jsonToSheet = (wb, json, options) => {
    let ws = initWorkSheet(wb, options.sheet.name);
    const parsedJson = typeof json === 'string' ? JSON.parse(json) : json;
    parsedJson.options = { ...defaultOptions, ...parsedJson.options };
    ParseJsonToExcel(ws, parsedJson);
    return wb;
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

const exportHtmlToExcel = (table, newOptions = {}) => {
    const options = { ...defaultOptions, ...newOptions };
    let wb = initWorkbook();
    wb = htmlToSheet(wb, table, options);
    save(wb, options.name, options.saveLocal);
}

const exportJsonToExcel = (json, newOptions = {}) => {
    const options = { ...defaultOptions, ...newOptions };
    let wb = initWorkbook();
    wb = jsonToSheet(wb, json, options);
    save(wb, options.name, options.saveLocal);
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