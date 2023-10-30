import { read } from 'xlsx';
import { set_cptable, utils } from 'xlsx';
import * as cptable from 'xlsx/dist/cpexcel.full.mjs';
const ExcelJS = require('exceljs');
set_cptable(cptable);

export const getBook = async files => {
    if (files && files[0]) {
        const selectedFile = files[0];
        const workbook = new ExcelJS.Workbook();
        const data = await (await fetch(window.URL.createObjectURL(selectedFile))).arrayBuffer();
        await workbook.xlsx.load(data);

        return workbook;
    }
};

export const getRows = async files => {
    if (files && files[0]) {
        const selectedFile = files[0];

        const data = await (await fetch(window.URL.createObjectURL(selectedFile))).arrayBuffer();
        const book = read(data, { cellStyles: true });
        const mainPage = book.Sheets[book.SheetNames[0]];
        const rows = utils.sheet_to_json(mainPage, { header: 1 });

        return rows;
    }
};
