const { getExcelWorkbook, getWorksheet } = require('./reader');

/**
 * Oculta colunas específicas de uma planilha em um arquivo Excel.
 *
 * @param {string} fileExcel - Caminho completo do arquivo Excel (.xlsx).
 * @param {string} sheetName - Nome da planilha onde as colunas devem ser ocultadas.
 * @param {number} startColumn - Índice da primeira coluna a ser ocultada (1 = coluna A).
 * @param {number} numColumns - Número de colunas consecutivas a serem ocultadas.
 * @returns {Promise<void>} - Retorna uma Promise que é resolvida após salvar o arquivo.
 *
 */
async function hiddenColumns(fileExcel, sheetName, startColumn, numColumns) {

    let workbook = await getExcelWorkbook(fileExcel);
    let worksheet = await getWorksheet(workbook, sheetName);
    // Ocultar as colunas especificadas
    for (let i = 0; i < numColumns; i++) {
        worksheet.getColumn(startColumn + i).hidden = true;
    }
    // Salvar o workbook modificado
    await workbook.xlsx.writeFile(fileExcel);

}

module.exports = { hiddenColumns };
