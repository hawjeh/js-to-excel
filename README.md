# JS to Table

[![hawjeh/js-to-excel Build](https://github.com/hawjeh/js-to-excel/actions/workflows/webpack.yml/badge.svg)](https://github.com/hawjeh/js-to-excel/actions/workflows/webpack.yml)

A javascript tool to export HTML table / JSON to valid excel file effortlessly.
This library uses [exceljs/exceljs](https://github.com/exceljs/exceljs) under the hood to create the excel.

## Sample Table

[Visit here](https://hawjeh.com/js-to-excel)

or

Checkout the sample project in the repo

## Installation / Usage

### Browser

Just add a script tag:

```html
<script type="text/javascript" src="./dist/JsToExcel.js"></script>
<script>
    const { exportHtmlToExcel } = window.JsToExcel;
    function init() {
        JsToExcel.exportHtmlToExcel(document.getElementById('table1'));
    }
    
    function init_with_name() {
        JsToExcel.exportHtmlToExcel(document.getElementById('table1'), {
            name: 'new_excel.xlsx',
            saveLocal: false,
            sheet: {
                name: "My Sheet 1"
            }
        });
    }
</script>
```

### Node

```bash
npm install @hawjeh/js-to-excel --save

import "@hawjeh/js-to-excel";
const { exportHtmlToExcel, exportJsonToExcel } = window.JsToExcel;

const onExportHtmlClick = () => {
    const table = window.document.getElementById('entry-table');
    exportHtmlToExcel(table);
}

const onExportJsonClick = () => {
    exportJsonToExcel('{"options":{"showHeader":true,"globalStyles":{},"headerStyles":{},"rowStyles":[]},"keys": ["Header1"],"rows":[{"Header1":"Val1"},{"Header1":"Val1"},{"Header1":"Val1"},{"Header1":"Val1"]}', {
    name: 'C:\\export.xlsx',
    saveLocal: true
    })
}
```

or

Run the sample project:
```bash
npm install @hawjeh/js-to-excel --save
cd ./node_modules/@hawjeh/js-to-excel
npm install && npm run start
```
Then open browser and go to [http://localhost:8080](http://localhost:8080)

### Example HTML:

Check the file here: [src/test/index.html](https://github.com/hawjeh/js-to-excel/blob/main/src/test/index.html)

<br>

## Cell

Cell can be set using the following data attributes:

| Attribute     | Description                        | Values                                              |
| --------------| ---------------------------------- | --------------------------------------------------- |
| `data-type`   | To specify the data type of a cell | `s` : String (Default) <br> `n` : Number <br> `b` : Boolean <br> `d` :  Date <br> `https://`: Hyperlink  |
| `numFmt`      | Number Format                      | "0", "0.00%", "0.0%"                                |
| `data-exclude`| Exclude Cell                       |                                                     |

Example:

```html
<!-- For setting a cell type as number -->
<td data-type="n">2500</td>

<!-- For setting a cell type as number with format -->
<td data-type="n" data-fmt="0.00">12345</td>

<!-- For setting a cell type as date -->
<td data-type="d">05-23-2018</td>

<!-- For setting a cell type as boolean -->
<td data-t="b">true</td>

<!-- For setting a cell type as hyperlink -->
<td data-type="https://google.com">Google</td>

<!-- For excluding a cell -->
<td data-exclude="true">Exclude me</td>

```

### Cell Styling

Cell styling can be set using the following data attributes:

#### Cell Font Styling:

| Attribute         | Description                   | Values                                              |
| ----------------- | ----------------------------- | --------------------------------------------------- |
| `data-f-name`     | To specify font-name          | "Calibri", "Arial", "Times"                         |
| `data-f-size`     | To specify font-size          | "11" // font size in points                         |
| `data-f-color`    | To specify font-color         | A hex ARGB value. Eg: FFFFOOOO for opaque red       |
| `data-f-bold`     | To bold font                  | `False` (Default) or `True`                         |
| `data-f-italic`   | To italic font                | `False` (Default) or `True`                         |
| `data-f-strike`   | To strike through font        | `False` (Default) or `True`                         |
| `data-f-underline`| To underline font             | `False` (Default) or `True`                         |

Example:

```html
<!-- For setting a cell font name, size and color -->
<td data-f-name="Arial" data-f-size="14" data-f-color="FFFF0000" >Arial, Size 14, Red Text</td>

<!-- For setting a cell bold, italic, strike and underline -->
<td data-f-bold="true">Bold</td>
<td data-f-italic="true">Italic</td>
<td data-f-strike="true">Strike Through</td>
<td data-f-underline="true">Underline</td>

```

##### Cell Alignment:

| Attribute         | Description                   | Values                                              |
| ----------------- | ----------------------------- | --------------------------------------------------- |
| `data-a-hori`     | To set horizontal alignment   | `left` (Default), `center`, `right`, `fill`, `justify`, `centerContinuous`, `distributed`                   |
| `data-a-vert`     | To set vertical alignment     | `top` (Default), `middle`, `bottom`, `distributed`, `justify`           |
| `data-a-wrap`     | To determine wrap text        | `False` (Default) or `True`                         |
| `data-a-indt`     | To determine text indent      | "0" // integer                                      |
| `data-a-rtl`      | To determine text direction   | `False` (Default) - left to right or `True` - right to left                |
| `data-a-stf`      | To determine shrink to fit    | `False` (Default) or `True`                         |
| `data-a-rota`     | To determine text rotation    | "" (Default), "0 to 90", "-1 to -90", "vertical"    |

Example:

```html
<!-- For setting a cell align center horizontal -->
<td data-a-hori="center">Horizontal Center</td>

<!-- For setting a cell align bottom vertical -->
<td data-a-vert="bottom">Vertical Bottom</td>

<!-- For setting a cell wrap text -->
<td data-a-wrap="true">Wrap Text</td>

<!-- For setting a cell to indent by 10 -->
<td data-a-indt="10">Indent 10</td>

<!-- For setting a cell direction from right to left -->
<td data-a-rtl="true">Right to left</td>

<!-- For setting a cell shrink to fit -->
<td data-a-stf="true">Shrink to fit</td>

<!-- For setting a cell rotation -->
<td data-a-rota="vertical">Vertical Rotate</td>

```

##### Cell Border:

| Attribute             | Description                   | Values                                        |
| --------------------- | ----------------------------- | --------------------------------------------- |
| `data-b-style`        | To set border style           | Check `BORDER_STYLES`                         |
| `data-b-style-color`  | To set border color           | A hex ARGB value. Eg: FFFFOOOO for opaque red |
| `data-b-side-all`     | To set border for all sides   |                                               |
| `data-b-side-top`     | To set border for top only    |                                               |
| `data-b-side-bottom`  | To set border for bottom only |                                               |
| `data-b-side-left`    | To set border for left only   |                                               |
| `data-b-side-right`   | To set border for right only  |                                               |

**`BORDER_STYLES:`** `thin`, `dotted`, `dashDot`, `hair`, `dashDotDot`, `slantDashDot`, `mediumDashed`, `mediumDashDotDot`, `mediumDashDot`, `medium`, `double`, `thick`

Example:

```html
<!-- For setting a cell bordered -->
<td data-b-side-all="">Bordered</td>

<!-- For setting a cell bordered left and right with dotted style -->
<td data-b-side-left="" data-b-side-right="" data-b-style="dotted">Dotted left and right</td>

<!-- For setting a cell bordered top only with color red -->
<td data-b-side-top="" data-b-style-color="FFFF0000">Top Red Only</td>

<!-- For setting a cell bordered bottom only with double line style -->
<td data-b-side-bottom="" data-b-style="double">Bottom Double</td>

```


## Inspired by:

- [ExcelJS](https://github.com/exceljs/exceljs)
- [@linways/table-to-excel](https://github.com/linways/table-to-excel)

