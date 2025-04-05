const path = require('path');
const fs = require('fs');

/**
 * Define o diretório a partir da pasta do projeto se necessário
 *
 * @param {bool} isReadingMode Se o diretório está sendo acessado para leitura ou escrita
 */
function setDirectory(directory, pastaProjeto, isReadingMode = true) {
    if (directory && !path.isAbsolute(directory)) {
        directory = path.join(pastaProjeto, directory);
        console.log('aqui', directory);
    } else if (!directory) {
        if (!isReadingMode) {
            const timestamp = Date.now();
            directory = path.join(pastaProjeto, `temp_${timestamp}.xlsx`);
        }
    }
    // Cria o diretório se não existir
    const dirPath = path.dirname(directory);
    if (!fs.existsSync(dirPath)) {
        fs.mkdirSync(dirPath, { recursive: true });
    }
    return directory;
}

module.exports = {
    setDirectory,
};