const buildBorderStyle = (hasBorder = false, border, position = 'data-b-side-all', style = 'thin', color = "FF00000") => {
    if (hasBorder) {
        switch (position) {
            case 'data-b-side-top':
                border.top = { style: style, color: { argb: color } };
                break;
            case 'data-b-side-bottom':
                border.bottom = { style: style, color: { argb: color } };
                break;
            case 'data-b-side-right':
                border.right = { style: style, color: { argb: color } };
                break;
            case 'data-b-side-left':
                border.left = { style: style, color: { argb: color } };
                break;
            default:
                border.top = { style: style, color: { argb: color } };
                border.left = { style: style, color: { argb: color } };
                border.bottom = { style: style, color: { argb: color } };
                border.right = { style: style, color: { argb: color } };
        }
    }

    return border;
}

const buildStyle = (table, defaultOptions) => {
    let options = JSON.parse(JSON.stringify(defaultOptions));

    // Global Styles
    options.showHeader = table.getAttribute('data-header') && table.getAttribute('data-header') === 'true' || false;
    options.globalStyles.borderedTable = table.getAttribute('data-bordered') && table.getAttribute('data-bordered') === 'true' || false;
    options.globalStyles.font.name = table.getAttribute('data-f-name') ? table.getAttribute('data-f-name') : options.globalStyles.font.name;
    options.globalStyles.font.size = table.getAttribute('data-f-size') ? table.getAttribute('data-f-size') : options.globalStyles.font.size;
    options.globalStyles.font.color.argb = table.getAttribute('data-f-color') ? table.getAttribute('data-f-color') : options.globalStyles.font.color.argb;
    options.globalStyles.font.bold = table.getAttribute('data-f-bold') && table.getAttribute('data-f-bold') === 'true' || false;
    options.globalStyles.font.italic = table.getAttribute('data-f-italic') && table.getAttribute('data-f-italic') === 'true' || false;
    options.globalStyles.font.strike = table.getAttribute('data-f-strike') && table.getAttribute('data-f-strike') === 'true' || false;
    options.globalStyles.font.underline = table.getAttribute('data-f-underline') && table.getAttribute('data-f-underline') === 'true' || false;

    // Header
    const header = table.tHead;
    let tempHeader = JSON.parse(JSON.stringify(options.globalStyles));
    tempHeader.font.name = header.getAttribute('data-f-name') ? header.getAttribute('data-f-name') : tempHeader.font.name;
    tempHeader.font.size = header.getAttribute('data-f-size') ? header.getAttribute('data-f-size') : tempHeader.font.size;
    tempHeader.font.color.argb = header.getAttribute('data-f-color') ? header.getAttribute('data-f-color') : tempHeader.font.color.argb;
    tempHeader.font.bold = header.getAttribute('data-f-bold') && header.getAttribute('data-f-bold') === 'true' || tempHeader.font.bold;
    tempHeader.font.italic = header.getAttribute('data-f-italic') && header.getAttribute('data-f-italic') === 'true' || tempHeader.font.italic;
    tempHeader.font.strike = header.getAttribute('data-f-strike') && header.getAttribute('data-f-strike') === 'true' || tempHeader.font.strike;
    tempHeader.font.underline = header.getAttribute('data-f-underline') && header.getAttribute('data-f-underline') === 'true' || tempHeader.font.underline;
    tempHeader.alignment.horizontal = header.getAttribute('data-a-hori') ? header.getAttribute('data-a-hori') : tempHeader.alignment.horizontal;
    tempHeader.alignment.vertical = header.getAttribute('data-a-vert') ? header.getAttribute('data-a-vert') : tempHeader.alignment.vertical;
    tempHeader.alignment.wrapText = header.getAttribute('data-a-wrap') && header.getAttribute('data-a-wrap') === 'true' || tempHeader.alignment.wrapText;
    tempHeader.alignment.shrinkToFit = header.getAttribute('data-a-stf') && header.getAttribute('data-a-stf') === 'true' || tempHeader.alignment.shrinkToFit;
    tempHeader.alignment.indent = header.getAttribute('data-a-indt') ? header.getAttribute('data-a-indt') : tempHeader.alignment.indent;
    tempHeader.alignment.readingOrder = header.getAttribute('data-a-rtl') && header.getAttribute('data-a-rtl') === 'true' ? 'rtl' : tempHeader.alignment.readingOrder;
    tempHeader.alignment.textRotation = header.getAttribute('data-a-rota') ? header.getAttribute('data-a-rota') : tempHeader.alignment.textRotation;

    let borderColor = header.getAttribute('data-b-style-color') ? header.getAttribute('data-b-style-color') : 'FF000000';
    let borderStyle = header.getAttribute('data-b-style') ? header.getAttribute('data-b-style') : 'thin';
    let border = { top: {}, left: {}, bottom: {}, right: {} };
    [...header.attributes].filter(x => x.nodeName.startsWith('data-b-side')).forEach((x) => {
        if (x.nodeName) {
            border = buildBorderStyle(true, border, x.nodeName, borderStyle, borderColor);
        }
    });
    tempHeader.border = border;
    
    options.headerStyles = tempHeader;

    // Rows
    const rows = [...table.rows];
    options.rowStyles = [];
    for (let i = 0; i < rows.length; i++) {

        if (typeof options.rowStyles[i] === 'undefined') {
            options.rowStyles.push({});
        }

        // columns
        const currentRow = rows[i];
        const tds = [...currentRow.children];
        for (let j = 0; j < tds.length; j++) {
            const td = tds[j];

            if (td.getAttribute('data-ex') && td.getAttribute('data-ex') === 'true') continue;

            let temp = { ...options.rowStyles[i] };
            temp.row = i;
            temp.column = j;
            temp.type = td.getAttribute('data-type') || 's';
            temp.numFmt = td.getAttribute('data-fmt') || '';

            let tempStyle = JSON.parse(JSON.stringify(options.globalStyles));
            tempStyle.font.name = td.getAttribute('data-f-name') ? td.getAttribute('data-f-name') : tempStyle.font.name;
            tempStyle.font.size = td.getAttribute('data-f-size') ? td.getAttribute('data-f-size') : tempStyle.font.size;
            tempStyle.font.color.argb = td.getAttribute('data-f-color') ? td.getAttribute('data-f-color') : tempStyle.font.color.argb;
            tempStyle.font.bold = td.getAttribute('data-f-bold') && td.getAttribute('data-f-bold') === 'true' || tempStyle.font.bold;
            tempStyle.font.italic = td.getAttribute('data-f-italic') && td.getAttribute('data-f-italic') === 'true' || tempStyle.font.italic;
            tempStyle.font.strike = td.getAttribute('data-f-strike') && td.getAttribute('data-f-strike') === 'true' || tempStyle.font.strike;
            tempStyle.font.underline = td.getAttribute('data-f-underline') && td.getAttribute('data-f-underline') === 'true' || tempStyle.font.underline;
            tempStyle.alignment.horizontal = td.getAttribute('data-a-hori') ? td.getAttribute('data-a-hori') : tempStyle.alignment.horizontal;
            tempStyle.alignment.vertical = td.getAttribute('data-a-vert') ? td.getAttribute('data-a-vert') : tempStyle.alignment.vertical;
            tempStyle.alignment.wrapText = td.getAttribute('data-a-wrap') && td.getAttribute('data-a-wrap') === 'true' || tempStyle.alignment.wrapText;
            tempStyle.alignment.shrinkToFit = td.getAttribute('data-a-stf') && td.getAttribute('data-a-stf') === 'true' || tempStyle.alignment.shrinkToFit;
            tempStyle.alignment.indent = td.getAttribute('data-a-indt') ? td.getAttribute('data-a-indt') : tempStyle.alignment.indent;
            tempStyle.alignment.readingOrder = td.getAttribute('data-a-rtl') && td.getAttribute('data-a-rtl') === 'true' ? 'rtl' : tempStyle.alignment.readingOrder;
            tempStyle.alignment.textRotation = td.getAttribute('data-a-rota') ? td.getAttribute('data-a-rota') : tempStyle.alignment.textRotation;

            let borderColor = td.getAttribute('data-b-style-color') ? td.getAttribute('data-b-style-color') : 'FF000000';
            let borderStyle = td.getAttribute('data-b-style') ? td.getAttribute('data-b-style') : 'thin';
            let border = { top: {}, left: {}, bottom: {}, right: {} };
            [...td.attributes].filter(x => x.nodeName.startsWith('data-b-side')).forEach((x) => {
                if (x.nodeName) {
                    border = buildBorderStyle(true, border, x.nodeName, borderStyle, borderColor);
                }
            });
            tempStyle.border = border;

            temp.style = tempStyle;
            options.rowStyles.push(temp);
        }
    }

    options.rowStyles = options.rowStyles.filter(x => Object.keys(x).length > 0);
    return options;
}

const getHeaders = (row, keys) => {
    let headers = [];
    for (let i = 0; i < keys.length; i++) {
        headers.push({
            key: keys[i],
            header: row[keys[i]]
        });
    }
    return headers;
}

const getValue = (text, setting) => {
    let val;
    const valueType = setting.type;
    switch (valueType) {
        case 's':
            val = String(text);
            break;
        case 'n':
            val = Number(text);
            break;
        case 'd':
            let date = new Date(text);
            val = new Date(
                Date.UTC(
                    date.getFullYear(),
                    date.getMonth(),
                    date.getDate(),
                    date.getHours(),
                    date.getMinutes(),
                    date.getSeconds()
                )
            );
            break;
        case 'b':
            val = String(text).toLowerCase() === "true"
                ? true : String(text).toLowerCase() === "false"
                    ? false : Boolean(parseInt(String(text)));
            break;
        default:
            val = { text: text, hyperlink: valueType, tooltip: valueType };
            break;
    }

    return val;
}

export { buildStyle, getValue, getHeaders };