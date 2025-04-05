const { formatFullDate, getExtendedDate, getScheduleDate } = require('../utils/dateUtils');

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
    } else if (rule.includes('extenso')) {
        return getExtendedDate(date);
    } else { // Caso contrário, formata como data dia/mês/ano
        return formatFullDate(date);
    }
}

module.exports = {

    formatReplacement,
    formatData,
    formatDate
}