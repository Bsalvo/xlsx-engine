const { formatReplacement } = require('./formatter');
const { getSheetColumnMap } = require('./transformer');
const { getWorksheet, getExcelWorkbook, getHeaderRow } = require('./reader');
const { formatTextToIdentifier } = require('./transformer');

/**
 * Converte uma planilha do Excel em um array de objetos JSON formatados.
 * 
 * @param {string} fileExcel - O caminho do arquivo Excel.
 * @param {number} initRow - O número da linha inicial para começar a conversão.
 * @param {string} sheetIndex - O índice da planilha a ser retornada (padrão é 1).
 * @param {number} [headerIndex=1] - O cabeçalho da planilha (padrão é 1).
 * @param {Array} necessaryColumns - Colunas obrigatórias para a construção essencial dos dados da planilha
 * @param {Object} config - Objeto de configuração para retorno de dados
 * @returns {Array<Object>} - Um array de objetos JSON formatados.
 * @throws {Error} - Lança um erro se os parâmetros fornecidos não forem válidos.
 */
async function excelToJson(fileExcel, initRow, sheetIndex = 1, headerIndex = 1, necessaryColumns, config = null) {

    const { worksheet, columnMap, headerRow } = await getExcelData(fileExcel, necessaryColumns, sheetIndex, headerIndex);

    if (typeof initRow !== 'number' || initRow < 1) {
        throw new Error('Número de linha inicial inválido fornecido. Deve ser um número maior ou igual a 1.');
    }

    let formattedReplacements = [];
    try {
        for (let rowNumber = initRow; rowNumber <= worksheet.rowCount; rowNumber++) {
            const row = worksheet.getRow(rowNumber);

            let data = setObjectReplacements(row, headerRow, worksheet);

            const replacements = data.replacements
            //  console.log(data.styles);
            const formattedReplacement = formatReplacement(replacements);
            if (Object.keys(formattedReplacement).length > 0) {
                formattedReplacements.push(formattedReplacement);
            }
        }
        if (config && config.header) {
            return {
                data: formattedReplacements,
                header: columnMap,
            }
        }
        return formattedReplacements;
    } catch (error) {
        console.log(error);
    }
}

/**
 * Lê um arquivo Excel e retorna a planilha específica, mapeamento de colunas e linha de cabeçalho.
 *
 * @param {string} fileExcel - O caminho do arquivo Excel.
 * @param {Array} necessaryColumns - Colunas obrigatórias para a construção essencial dos dados da planilha.
 * @param {number} [sheetIndex=1] - O índice da planilha a ser retornada (padrão é 1).
 * @returns {Promise<Object>} - Um objeto contendo a planilha, mapeamento de colunas e linha de cabeçalho.
 * @throws {Error} - Lança um erro se o arquivo Excel não puder ser lido ou se houver problemas com a planilha.
 */
async function getExcelData(fileExcel, necessaryColumns, sheetIndex = 1, headerIndex = 1) {
    if (!fileExcel || typeof fileExcel !== 'string') {
        throw new Error('Caminho do arquivo Excel inválido fornecido.');
    }

    try {
        const workbook = await getExcelWorkbook(fileExcel);
        const worksheet = getWorksheet(workbook, sheetIndex);
        const columnMap = getSheetColumnMap(worksheet, headerIndex, necessaryColumns);
        const headerRow = getHeaderRow(worksheet, headerIndex);

        return { worksheet, columnMap, headerRow };
    } catch (error) {
        console.error('Erro ao obter dados do arquivo Excel:', error);
    }
}

/**
 * Cria um objeto de substituições a partir de uma linha e uma linha de cabeçalho de uma planilha do Excel.
 *
 * @param {Object} row - A linha de dados do ExcelJS.
 * @param {Number} rowNumber - A linha de cabeçalho do ExcelJS.
 * @param {Object} headerRow - A linha de cabeçalho do ExcelJS.
 * @returns {Object} - Um objeto com os identificadores formatados como chaves e os valores das células correspondentes.
 * @throws {Error} - Lança um erro se os parâmetros fornecidos não forem válidos.
 */
function setObjectReplacements(row, headerRow, worksheet) {
    if (!headerRow || typeof headerRow.getCell !== 'function') {
        throw new Error('Linha de cabeçalho inválida fornecida.');
    }

    let replacements = {};
    let styles = {};

    row.eachCell((cell, colNumber) => {
        const columnCell = headerRow.getCell(colNumber);
        const column = worksheet.getColumn(colNumber);
        if (columnCell && columnCell.value && !column.hidden) {
            let columnName = columnCell.value;
            const upString = formatTextToIdentifier(columnName);

            if (upString) {
                let newKey = upString;
                let counter = 1;

                // Verifica se a chave já existe e incrementa o sufixo
                while (replacements.hasOwnProperty(newKey)) {
                    counter++;
                    newKey = `${upString}_${counter}`;
                }

                // Define o valor no objeto com a chave única
                //console.log(cell.style.border);
                replacements[newKey] = cell.value;
                styles[newKey] = cell.style;
            }
        }
    });

    return { replacements, styles };
}

module.exports = {
    excelToJson,
    getExcelData,
    setObjectReplacements
};
