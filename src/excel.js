const os = require('os');
const path = require('path');
const { createExcelXlsx } = require('./creator/creator');
const { createExcelCsv } = require('./creator/csvCreator');
const { excelToJson } = require('./parser/extractor');
const { csvToXlsx } = require('./parser/reader');
const { formatTextToIdentifier } = require('./parser/transformer');
const { setDirectory } = require('./utils/pathUtils');

class Excel {
  /** Indica qual o diretório onde os arquivos manipulados irão ser inseridos ou acessados */
  pastaProjeto;

  constructor(pastaProjeto, config = {}) {
    this.pastaProjeto = pastaProjeto ?? null;
    if (this.pastaProjeto) {
      this.pastaProjeto = path.isAbsolute(this.pastaProjeto)
        ? this.pastaProjeto
        : path.join(os.homedir(), this.pastaProjeto);
    }

    this.document_font = config?.document?.font ?? 'Aptos Narrow';
    this.document_size = config?.document?.size ?? 9;

    /** Pré-definições de estilos de cabeçalhos para a criação de planilha */
    this.documentPredefinitions = {
      default: {
        header: { fixed: true, style: { font: { name: this.document_font, size: this.document_size, bold: true }, alignment: 'center' } },
        global: { style: { font: { name: this.document_font, size: this.document_size } } },
      },
    };
  }



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
      setDirectory(fileExcel, this.pastaProjeto),
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
      config = this.documentPredefinitions[config] ?? {};
    }

    directory = setDirectory(directory, this.pastaProjeto, false);
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
      if (!csvFilePath) {
        throw new Error('Caminho do arquivo Excel (.csv) inválido fornecido.');
      }

      if (!xlsxFilePath) {
        xlsxFilePath = `${path.parse(csvFilePath).name}.xlsx`;
      }

      await csvToXlsx(setDirectory(csvFilePath), setDirectory(xlsxFilePath), aba);

    } catch (error) {
      throw error;
    }

  }

  toIdentifier(value) {
    return formatTextToIdentifier(value);
  }

  async createExcelCsv(filePath, columns, rows) {
    createExcelCsv(filePath, columns, rows);
  }

}

module.exports = Excel;
