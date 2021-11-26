import { defaultStyle, defaultOptions } from './default.js';
import { buildStyle, getHeaders, getValue } from './helper.js';

const ParseTableToExcel = (ws, table) => {

    const rows = [...table.rows];
    const options = buildStyle(table, defaultOptions);

    let json = {
        options: options,
        keys: [],
        rows: []
    }

    // Columns & Rows
    for (let i = 0; i < rows.length; i++) {

        // columns
        const currentRow = rows[i];
        const tds = [...currentRow.children];

        for (let j = 0; j < tds.length; j++) {
            const td = tds[j];
            if (td.getAttribute('data-ex') && td.getAttribute('data-ex') === 'true') {
                continue;
            }

            if (typeof json.rows[i] === 'undefined') {
                json.rows.push({});
            }

            if (i === 0) {
                // header
                let headerName = 'header' + j;
                json.keys.push(headerName);
            }

            let tempObject = {};
            tempObject[json.keys[j]] = getValue(td.innerText, options.rowStyles.find(x => x.row === i && x.column === j));
            json.rows[i] = { ...tempObject, ...json.rows[i] };
        }
    }

    json.rows = json.rows.filter(x => Object.keys(x).length > 0)
    return ParseJsonToExcel(ws, json);
}

const ParseJsonToExcel = (ws, json) => {
    try {
        // styling
        const options = { ...json.options };
        const globalFont = options.globalStyles.font || {};
        const globalBorderedTable = options.globalStyles.borderedTable || false;
        let style = { ...defaultStyle };
        if (globalFont !== {}) {
            style.font = globalFont;
        }
        if (globalBorderedTable) {
            style.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };

            options.headerStyles.border = style.border;
            options.rowStyles = options.rowStyles.map(x => {
                x.style.border = style.border;
                return x;
            });
        }

        const showHeader = options.showHeader || false;
        const headers = getHeaders(json.rows[0], json.keys);
        const rows = json.rows.splice(1) || [];

        // Header
        if (headers && headers.length > 0 && headers !== []) {
            ws.columns = headers;

            // styling
            let headerRow = ws.getRow(1);
            let length = headers.length;
            let counter = 1;
            while (counter <= length) {
                headerRow.getCell(counter).style = options.headerStyles;
                headerRow.hidden = !showHeader;
                counter++;
            }

            headerRow.commit();
        }

        // Rows
        if (rows && rows.length > 0 && rows !== []) {
            for (let i in rows) {
                ws.addRow(rows[i]);
            }

            const min = 2;
            const max = rows.length + min;
            for (let i = min; i < max; i++) {
                let bodyRow = ws.getRow(i);
                let counter = 1;
                const length = Object.keys(rows[i - min]).length;
                while (counter <= length) {
                    const rowStyle = options.rowStyles.find(x => x.row === (i - 1) && x.column === (counter - 1));
                    bodyRow.getCell(counter).style = rowStyle && rowStyle.style ? rowStyle.style : style;
                    bodyRow.getCell(counter).numFmt = rowStyle && rowStyle.numFmt ? rowStyle.numFmt : '';
                    counter++;
                }

                bodyRow.commit();
            }
        }

    } catch {
        console.log('Invalid JSON');
    }

    return ws;
}

export {
    ParseTableToExcel,
    ParseJsonToExcel
};