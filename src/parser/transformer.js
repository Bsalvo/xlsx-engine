const unidecode = require('unidecode');
const { getHeaderRow } = require('./reader');

/**
 * Formata o texto de cabeçalho para identificador, convertendo para minúsculas, removendo acentos e substituindo espaços por underscores.
 * Esses identificadores devem estar presentes no documento template. Dessa forma será possível substituir identificadores pelo texto necessário.
 *
 * @param {string} texto - O texto a ser formatado.
 * @returns {string} - O texto formatado como identificador.
 * @throws {Error} - Lança um erro se o texto fornecido não for uma string.
 */
function formatTextToIdentifier(texto) {
    if (typeof texto !== 'string') {
        return null;
        // throw new Error('Texto inválido fornecido. Deve ser uma string.');
    }
    let formattedText = texto.trim();
    formattedText = texto.toLowerCase();
    formattedText = unidecode(formattedText);
    formattedText = formattedText.trim();
    formattedText = formattedText.replace(/ /g, "_");

    return formattedText;
}

/**
 * Verifica se as colunas necessárias estão presentes no mapeamento de colunas.
 *
 * @param {Object} columnMap - O mapeamento de colunas do Excel.
 * @param {Array<string>} necessaryColumns - As colunas necessárias a serem verificadas.
 * @throws {Error} - Lança um erro se alguma coluna necessária estiver faltando.
 */
function verifyNecessaryColumns(columnMap, necessaryColumns) {
    if (typeof columnMap !== 'object' || columnMap === null) {
        throw new Error('Mapeamento de colunas inválido fornecido.');
    }
    if (!Array.isArray(necessaryColumns)) {
        throw new Error('Lista de colunas necessária inválida fornecida.');
    }

    const missingColumns = necessaryColumns.filter(column => !columnMap[column]);
    if (missingColumns.length > 0) {
        throw new Error(`Colunas "${missingColumns.join(', ')}" são necessárias para processar os dados da planilha corretamente.`);
    }
}

/**
 * Retorna um mapeamento de colunas de uma planilha do Excel, onde as chaves são os nomes das colunas (em maiúsculas) e os valores são os números das colunas.
 *
 * @param {Object} worksheet - O objeto worksheet do ExcelJS.
 * @param {number} headerIndex - Qual linha o cabeçalho está.
 * @param {Array} necessaryColumns - Colunas obrigatórias para a construção essencial dos dados da planilha
 * @returns {Object} - Um mapa de colunas com nomes de colunas como chaves e números de colunas como valores.
 * @throws {Error} - Lança um erro se a linha de cabeçalho não existir.
 */
function getSheetColumnMap(worksheet, headerIndex, necessaryColumns) {
    if (!worksheet || typeof worksheet.getRow !== 'function') {
        throw new Error('Planilha inválida fornecida.');
    }

    /** Cabeçalho da planilha analiasda */
    const headerRow = getHeaderRow(worksheet, headerIndex);
    if (!headerRow) {
        throw new Error('A linha de cabeçalho não foi encontrada na planilha.');
    }

    const columnMap = {};
    headerRow.eachCell((cell, colNumber) => {
        if (cell.value && typeof cell.value != 'object') {
            columnMap[cell.value.toUpperCase()] = colNumber;
        }
    });

    try {
        if (necessaryColumns) {
            verifyNecessaryColumns(columnMap, necessaryColumns);
        }
        return columnMap;
    } catch (error) {
        throw new Error(error.message);
    }

}

module.exports = {
    formatTextToIdentifier,
    getSheetColumnMap,
    verifyNecessaryColumns,
};