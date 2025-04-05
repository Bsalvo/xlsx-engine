const fs = require('fs');
const util = require('util');
const writeFile = util.promisify(fs.writeFile);

/**
 * Cria um arquivo CSV diretamente a partir de colunas e linhas fornecidas.
 *
 * @param {string} outputFilePath - Caminho onde o arquivo CSV será salvo.
 * @param {Array} columns - Lista de colunas (strings) que definem o cabeçalho.
 * @param {Array<Object>} rows - Dados em formato de array de objetos.
 */
async function createExcelCsv(outputFilePath, columns, rows) {
    try {
        const csvContent = [
            columns.map((col) => `"${col}"`).join(';'),
            ...rows.map((row) => {
                return columns
                    .map((col) => {
                        const value = row[col] || '';
                        return `"${value.toString().replace(/"/g, '""')}"`;
                    })
                    .join(';');
            }),
        ].join('\n');

        await writeFile(outputFilePath, '\uFEFF' + csvContent, 'utf8');
        console.log(`Arquivo CSV criado com sucesso: ${outputFilePath}`);
    } catch (error) {
        throw new Error(`Erro ao criar o arquivo CSV: ${error.message}`);
    }
}

module.exports = { createExcelCsv };