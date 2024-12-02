const ExcelJS = require('exceljs');
const unidecode = require('unidecode');
const { formatFullDate, getExtendedDate, getScheduleDate } = require('../utils/controller');

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
            
            const replacements = setObjectReplacements(row, headerRow, worksheet);
            const formattedReplacement = formatReplacement(replacements);
            if(Object.keys(formattedReplacement).length > 0){
                formattedReplacements.push(formattedReplacement);
            }
        }
        if(config && config.header){
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
        if(necessaryColumns){
            verifyNecessaryColumns(columnMap, necessaryColumns);
        }
        return columnMap;
    } catch (error) {
        throw new Error(error.message);
    }

    
}

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
                replacements[newKey] = cell.value;
            }
        }
    });

    return replacements;
}


/**
 * Formata os valores de um objeto de substituições. Como datas; trim() no valores..
 *
 * @param {Object} obj - O objeto de substituições a ser formatado.
 * @returns {Object} - O objeto de substituições com os valores formatados.
 * @throws {Error} - Lança um erro se o parâmetro fornecido não for um objeto.
 */
function formatReplacement(obj) {
    if (typeof obj !== 'object' || obj === null) {
        throw new Error('Parâmetro inválido fornecido. Deve ser um objeto.');
    }

    const formattedObj = {};
    for (const key in obj) {
        if (obj.hasOwnProperty(key)) {
            const value = obj[key];
            formattedObj[key] = formatData(key, value); // Formata o valor
        }
    }

    return formattedObj;
}

/**
 * Formata o valor vindo do excel com base no tipo e na chave fornecida.
 *
 * @param {string} key - A chave que pode influenciar a formatação.
 * @param {any} value - O valor a ser formatado.
 * @returns {string} - O valor formatado.
 * @throws {Error} - Lança um erro se a chave não for uma string.
 */
function formatData(key, value) {
    if (typeof key !== 'string') {
        throw new Error('A chave fornecida deve ser uma string.');
    }

    if (value instanceof Date) {
        return formatDate(key, value);
    }

    if (typeof value === 'string') {
        return value.trim();
    }

    if (typeof value === 'object' && value !== null) {
        return value.result || '';
    }

    return String(value);
}

/**
 * Formata uma data ou horário com base na regra fornecida.
 *
 * @param {string} rule - A regra que indica o tipo de formatação (data ou horário, ou data extensa).
 * @param {string|Date} dateValue - O valor da data a ser formatado.
 * @returns {string} - A data ou horário formatado.
 * @throws {Error} - Lança um erro se a data fornecida não for válida.
 */
function formatDate(rule = '', dateValue) {
    const date = new Date(dateValue);
    if (isNaN(date.getTime())) { // Verifica se a data é inválida
        throw new Error('Valor de data inválido fornecido.');
    }

    // Corrige o fuso horário
    date.setMinutes(date.getMinutes() + date.getTimezoneOffset());

    if (rule.includes('horario')) { // Verifica se o campo é de horário e retorna apenas o horário específico
        return getScheduleDate(date);
    } else if(rule.includes('extenso')){
        return getExtendedDate(date);
    } else { // Caso contrário, formata como data dia/mês/ano
        return formatFullDate(date);
    }
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

async function hiddenColumns(fileExcel, sheetName, startColumn, numColumns){

    let workbook = await getExcelWorkbook(fileExcel);
    let worksheet = await  getWorksheet(workbook, sheetName);

    // Ocultar as colunas especificadas
    for (let i = 0; i < numColumns; i++) {
        worksheet.getColumn(startColumn + i).hidden = true;
    }

    // Salvar o workbook modificado
    await workbook.xlsx.writeFile(fileExcel);

}

module.exports = {
    getExcelWorkbook,
    getSheetColumnMap,
    getWorksheet,
    getHeaderRow,
    setObjectReplacements,
    formatReplacement,
    getExcelData,
    excelToJson,
    formatTextToIdentifier,
    hiddenColumns
}