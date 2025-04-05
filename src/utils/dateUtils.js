/**
 * Retorna uma data formatada por extenso (ex: "01 de janeiro de 2024").
 *
 * @param {Date|string} dateInput - O objeto Date ou string de data a ser formatado.
 * @returns {string} - A data formatada por extenso.
 * @throws {Error} - Lança um erro se o parâmetro fornecido não for um objeto Date válido ou uma string de data válida.
 */
function getExtendedDate(dateInput) {
    let date;

    if (dateInput instanceof Date) {
        date = dateInput;
    } else if (typeof dateInput === 'string') {
        const parts = dateInput.split('/');
        if (parts.length === 3) {
            const day = parseInt(parts[0], 10);
            const month = parseInt(parts[1], 10) - 1;
            const year = parseInt(parts[2], 10);
            date = new Date(year, month, day);
        } else {
            throw new Error('Formato de data inválido. Esperado: DD/MM/YYYY');
        }
    } else {
        throw new Error('Parâmetro inválido fornecido. Deve ser um objeto Date ou uma string de data.');
    }

    if (isNaN(date.getTime())) {
        throw new Error('Data fornecida é inválida.');
    }

    const months = [
        "janeiro", "fevereiro", "março", "abril", "maio", "junho",
        "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
    ];

    const day_a = String(date.getDate()).padStart(2, '0');
    const month_a = months[date.getMonth()];
    const year_a = date.getFullYear();
    
    return `${day_a} de ${month_a} de ${year_a}`;
}

/**
 * Formata uma data no formato DD/MM/YYYY.
 *
 * @param {Date} date - O objeto Date a ser formatado.
 * @returns {string} - A data formatada no formato DD/MM/YYYY.
 * @throws {Error} - Lança um erro se o parâmetro fornecido não for um objeto Date válido.
 */
function formatFullDate(date) {
    if (!(date instanceof Date) || isNaN(date.getTime())) {
        throw new Error('Parâmetro inválido fornecido. Deve ser um objeto Date válido.');
    }

    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0'); // Mês começa do zero
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
}

/**
 * Retorna uma data formatada como horário (HH:MM).
 *
 * @param {Date} date - O objeto Date a ser formatado.
 * @returns {string} - A data formatada como horário.
 * @throws {Error} - Lança um erro se o parâmetro fornecido não for um objeto Date válido.
 */
function getScheduleDate(date) {
    if (!(date instanceof Date) || isNaN(date.getTime())) {
        throw new Error('Parâmetro inválido fornecido. Deve ser um objeto Date válido.');
    }

    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    return `${hours}:${minutes}`;
}

module.exports = { 
    getExtendedDate,
    formatFullDate,
    getScheduleDate,
}