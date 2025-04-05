const ExcelJS = require('exceljs');
const fs = require('fs');
const readline = require('readline');

/**
 * Cria um objeto workbook a partir de um arquivo Excel.
 *
 * @param {string} fileExcel - O caminho do arquivo Excel.
 * @returns {Object} - O objeto workbook do ExcelJS.
 * @throws {Error} - Lança um erro se o arquivo Excel não puder ser lido.
 */
async function getExcelWorkbook(fileExcel) {
    if (!fileExcel || typeof fileExcel !== 'string') {
        throw new Error('Caminho do arquivo Excel inválido fornecido.');
    }

    const workbook = new ExcelJS.Workbook();
    try {
        await workbook.xlsx.readFile(fileExcel);
        return workbook;
    } catch (error) {
        console.error('Erro ao ler o arquivo Excel:', error);
        throw error;
    }
}

/**
 * Retorna uma planilha do Excel a partir do índice fornecido.
 *
 * @param {Object} workbook - O objeto workbook do ExcelJS.
 * @param {string|number} [sheetIndex=1] - O índice da planilha a ser retornada (padrão é 1).
 * @returns {Object} - O objeto worksheet do ExcelJS.
 * @throws {Error} - Lança um erro se a planilha não for encontrada.
 */
function getWorksheet(workbook, sheetIndex = 1) {
    if (!workbook || typeof workbook.getWorksheet !== 'function') {
        throw new Error('Workbook inválido fornecido.');
    }

    let worksheet = workbook.getWorksheet(sheetIndex);

    if (!worksheet) {
        const truncatedSheetName = sheetIndex.slice(0, 31);
        worksheet = workbook.getWorksheet(truncatedSheetName);
    }

    if (!worksheet) {
        throw new Error(`A planilha ${sheetIndex} não foi encontrada no workbook.`);
    }

    return worksheet;
}

/**
 * Retorna a linha de cabeçalho de uma planilha do Excel.
 *
 * @param {Object} worksheet - O objeto worksheet do ExcelJS.
 * @param {number} [header=1] - O número da linha de cabeçalho a ser retornada (padrão é 1).
 * @returns {Object} - O objeto headerRow do ExcelJS.
 * @throws {Error} - Lança um erro se a linha de cabeçalho não for encontrada.
 */
function getHeaderRow(worksheet, header = 1) {
    if (!worksheet || typeof worksheet.getRow !== 'function') {
        throw new Error('Planilha inválida fornecida.');
    }
    if (typeof header !== 'number' || header < 1) {
        console.log(header);
        throw new Error('Número da linha de cabeçalho inválido fornecido. Deve ser um número maior ou igual a 1.');
    }

    const headerRow = worksheet.getRow(header);
    if (!headerRow) {
        throw new Error(`A linha de cabeçalho ${header} não foi encontrada na planilha.`);
    }

    return headerRow;
}

/**
 * Converte um arquivo CSV em um arquivo XLSX.
 *
 * @param {string} csvFilePath - Caminho para o arquivo CSV de entrada.
 * @param {string} xlsxFilePath - Caminho para salvar o arquivo XLSX gerado.
 * @param {string} [aba='Planilha1'] - Nome da aba da planilha dentro do arquivo XLSX (padrão: 'Planilha1').
 * @returns {Promise<void>} - Retorna uma Promise que é resolvida quando o arquivo XLSX é salvo.
 */
async function csvToXlsx(csvFilePath, xlsxFilePath, aba = 'Planilha1') {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(aba);

    const fileStream = fs.createReadStream(csvFilePath, { encoding: 'utf-8' });
    const rl = readline.createInterface({
        input: fileStream,
        crlfDelay: Infinity,
    });

    for await (const line of rl) {
        const cleanedLine = line.trim(); // Remove espaços em branco extras
        const row = cleanedLine.split(/;|,/); // Divide usando ';' ou ','
        worksheet.addRow(row);
    }

    await workbook.xlsx.writeFile(xlsxFilePath);
    console.log(`Arquivo XLSX salvo em: ${xlsxFilePath}`);
}


module.exports = {
    getExcelWorkbook,
    getWorksheet,
    getHeaderRow,
    csvToXlsx,
};