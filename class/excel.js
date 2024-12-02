const os = require('os');
const path = require('path');
const fs = require('fs');
const { excelToJson } = require('../controller');
const { createExcelXlsx, csvToXlsx } = require('../creator');

class Excel {
  /** Indica qual o diretório onde os arquivos manipulados irão ser inseridos ou acessados */
  pastaProjeto;

  constructor(pastaProjeto) {
    this.pastaProjeto = pastaProjeto ?? null;
    if (this.pastaProjeto) {
      this.pastaProjeto = path.isAbsolute(this.pastaProjeto)
        ? this.pastaProjeto
        : path.join(os.homedir(), this.pastaProjeto);
    }
  }

  /** Pré-definições de estilos de cabeçalhos para a criação de planilha */
  headerPredefinitions = {
    default: {
      header: { fixed: true, style: { font: 'bold', alignment: 'center' } },
    },
  };

  /**
   * Converte uma planilha do Excel em um array de objetos JSON formatados.
   *
   * @param {string} fileExcel - O caminho do arquivo Excel.
   * @param {string} sheetName - O nome da aba a ser acessada e retornada.
   * @param {number} [headerIndex=1] - O cabeçalho da planilha (padrão é 1).
   * @param {number} initRow - O número da linha inicial para começar a conversão.
   * @param {Array} necessaryColumns - Colunas obrigatórias para a construção essencial dos dados da planilha
   * @returns {Array<Object>} - Um array de objetos JSON formatados.
   * @throws {Error} - Lança um erro se os parâmetros fornecidos não forem válidos.
   */
  async toJson(
    fileExcel,
    sheetName,
    headerIndex = 1,
    initRow = 2,
    necessaryColumns = []
  ) {
    let json = await excelToJson(
      this.setDirectory(fileExcel),
      initRow,
      sheetName,
      headerIndex,
      necessaryColumns,
      { header: true }
    );

    return {
      header: [...Object.keys(json.header)],
      data: json.data,
    };
  }

  /**
   * Função assíncrona para criar um arquivo Excel (.xlsx) com uma ou múltiplas abas.
   *
   * @param {string | Array} sheetConfigOrName - Nome da aba (string) para uma única aba ou array de configurações de abas para múltiplas abas.
   * Cada configuração de aba deve conter os campos `sheetName`, `columns` e `rows`.
   * @param {Array} columns - Definição das colunas para a aba única ou o diretório de saída, caso `sheetConfigOrName` seja um array.
   * @param {Array} rows - Linhas de dados para a aba única ou configurações adicionais, caso `sheetConfigOrName` seja um array.
   * @param {string} directory - Diretório para salvar o arquivo .xlsx.
   * @param {Object} config - Objeto de configurações opcionais, ou uma string com um estilo pré-definido.
   *
   * @returns {Promise<void>} Retorna uma promessa que cria e salva o arquivo Excel.
   */
  async create(
    sheetConfigOrName,
    columns,
    rows,
    directory = null,
    config = 'default'
  ) {

    if (typeof config === 'string') {
      config = this.headerPredefinitions[config] ?? {};
    }

    directory = this.setDirectory(directory, false);
    if (Array.isArray(sheetConfigOrName)) {
      await createExcelXlsx(sheetConfigOrName, directory, config);
    } else {
      await createExcelXlsx(
        sheetConfigOrName,
        columns,
        rows,
        directory,
        config
      );
    }
  }

  /**
   * Converte um arquivo CSV em um arquivo XLSX.
   *
   * @param {string} csvFilePath - Caminho para o arquivo CSV de entrada.
   * @param {string} xlsxFilePath - Caminho para salvar o arquivo XLSX gerado.
   * @param {string} aba - Nome da aba da planilha dentro do arquivo XLSX (padrão: 'Planilha1').
   * @returns {Promise<void>} - Retorna uma Promise que é resolvida quando o arquivo XLSX é salvo.
   */
  async toXlsx(csvFilePath, xlsxFilePath = null, aba) {

    try {
      if(!csvFilePath){
        throw new Error('Caminho do arquivo Excel (.csv) inválido fornecido.');
      }

      if(!xlsxFilePath){
        xlsxFilePath = `${path.parse(csvFilePath).name}.xlsx`;
      }

      await csvToXlsx(this.setDirectory(csvFilePath), this.setDirectory(xlsxFilePath), aba);

    } catch (error) {
      throw error;
    }

  }

  /**
   * Define o diretório a partir da pasta do projeto se necessário
   *
   * @param {bool} isReadingMode Se o diretório está sendo acessado para leitura ou escrita
   */
  setDirectory(directory, isReadingMode = true) {
    if (directory && !path.isAbsolute(directory)) {
      directory = path.join(this.pastaProjeto, directory);
    } else if (!directory) {
      if (!isReadingMode) {
        const timestamp = Date.now();
        directory = path.join(this.pastaProjeto, `temp_${timestamp}.xlsx`);
      }
    }

    // Cria o diretório se não existir
    const dirPath = path.dirname(directory);
    if (!fs.existsSync(dirPath)) {
      fs.mkdirSync(dirPath, { recursive: true });
    }
    return directory;
  }
}

module.exports = Excel;
