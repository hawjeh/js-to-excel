// styles: refers to https://github.com/exceljs/exceljs#styles

const defaultStyle = {
    font: {
        name: "Calibri",
        size: 11,
        color: { argb: 'FF000000' },
        bold: false,
        italic: false,
        strike: false,
        underline: false
    },
    alignment: {
        horizontal: 'left',
        vertical: 'top',
        wrapText: false,
        shrinkToFit: false,
        indent: 0,
        readingOrder: 'ltr',
        textRotation: ''
    },
    border: {
        top: {},
        left: {},
        bottom: {},
        right: {}
    }
};

const defaultOptions = {
    showHeader: true,
    globalStyles: { ...defaultStyle },
    headerStyles: {},
    rowStyles: [
        // {
        //     row: 1,
        //     column: 1,
        //     type: 's'.
        //     numFmt: '',
        //     style: { ...defaultStyle }
        // }
    ]
}

export { defaultStyle, defaultOptions };