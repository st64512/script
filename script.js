const xlsx = require('xlsx');
const fs = require('fs');
const inputFile = process.argv[2];
const outputFile = process.argv[3];

function prepareObjectPudniBlok(row) {
    row.nazev = row["č.vz."];
    delete row['č.vz.'];
    return {...row};
}

function completeObjectPudniBlok(pudniBlok, metaInfo, vzorky) {
    if (pudniBlok) {
        pudniBlok.metaInfo = {...metaInfo};
        pudniBlok.vzorky = [...vzorky];
    }
}

function getFirstCellOfRow(row) {
    let rowCells = Object.entries(row);
    return rowCells[0];
}

function isSummarizedRow(row) {
    const [key, value] = getFirstCellOfRow(row);
    return key === "č.vz." && typeof value === "string";
}

function loadXlsDataAsJson(fileName) {
    if (!fs.existsSync(fileName)) {
        console.error('Soubor ' + fileName + ' neexistuje!');
        process.exit(1);
    }
    const workbook = xlsx.readFile(fileName);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    return xlsx.utils.sheet_to_json(sheet);
}

function writeJsonResultIntoOutputFile(fileName, jsonData) {
    try {
        fs.writeFileSync(fileName, JSON.stringify(jsonData));
        console.log('Skript proběhl v pořádku. Data se zapsala do ' + fileName + ' souboru.');
    } catch (err) {
        console.error('Při zápisu dat do výstupního souboru nastala chyba!');
        console.error(err);
        process.exit(1);
    }
}

/* -------- Začátek skriptu -------- */

const jsonData = loadXlsDataAsJson(inputFile);
let isMetaRow = false;
let pudniBlok = {};
let pudniBlokyStructurizedJson = [];
let vzorky = [];

jsonData?.forEach((row) => {
    if (isMetaRow) {
        completeObjectPudniBlok(pudniBlok, row, vzorky);
        pudniBlokyStructurizedJson.push(pudniBlok);
        pudniBlok = {};
        vzorky = [];
        isMetaRow = false;
    } else if (isSummarizedRow(row)) {
        pudniBlok = prepareObjectPudniBlok(row);
        isMetaRow = true;
    } else {
        vzorky.push(row);
    }
})

writeJsonResultIntoOutputFile(outputFile, pudniBlokyStructurizedJson);