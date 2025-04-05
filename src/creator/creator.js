const ExcelJS = require('exceljs');

const { formatTextToIdentifier } = require('../parser/transformer');
const { applyCellStyle } = require('./styles');

/**
 * Cria e retorna uma nova instância de um workbook Excel.
 *
 * @returns {Promise<Object>} - Uma instância de `ExcelJS.Workbook`.
 *
 */
async function setExcelWorkbook() {
  const workbook = new ExcelJS.Workbook();
  return workbook;
}

/**
 * Cria e adiciona uma nova planilha ao workbook com o nome especificado.
 *
 * @param {Object} workbook - Instância de `ExcelJS.Workbook`.
 * @param {string} sheetName - Nome da nova planilha a ser criada.
 *
 * @returns {Promise<Object>} - Uma instância de `ExcelJS.Worksheet` referente à planilha criada.
 *
 * @throws {Error} Caso o nome da planilha não seja fornecido.
 */
async function setWorksheet(workbook, sheetName) {
  if (!sheetName) {
    throw new Error('Necessário informar o nome da planilha.');
  }
  const worksheet = workbook.addWorksheet(sheetName);
  return worksheet;
}

/**
 * Configura a linha de cabeçalho de uma planilha Excel.
 *
 * @param {Object} worksheet - Instância de `ExcelJS.Worksheet` onde o cabeçalho será configurado.
 * @param {Array} columns - Lista de colunas, podendo ser strings simples ou objetos com configurações específicas.
 * @param {Object} [config={}] - Configurações opcionais para o cabeçalho:
 *   - `fixed`: Booleano para fixar o cabeçalho.
 *   - `row`: Número da linha onde o cabeçalho será fixado (padrão: 1).
 *   - `style`: Estilos globais aplicáveis ao cabeçalho (fonte, alinhamento, rotação de texto).
 *   - `headerRow`: Número da linha de cabeçalho (padrão: 1).
 *
 * @throws {Error} Caso as colunas não sejam fornecidas ou estejam no formato errado.
 */
function setHeaderRow(worksheet, columns, config = {}) {
  if (!columns || !Array.isArray(columns)) {
    throw new Error('As colunas devem ser fornecidas como um array.');
  }

  const preparedColumns = columns.map((column) => {
    if (typeof column === 'string') {
      return { header: column, key: formatTextToIdentifier(column) };
    }
    if (typeof column === 'object' && column.value) {
      return {
        header: column.value,
        key: column.key || formatTextToIdentifier(column.value),
        width: column.width || null, // Respeita larguras manuais
        style: column.style || null,
        rotation: column.rotation || null,
      };
    }
    throw new Error(
      'Cada coluna deve ser uma string ou um objeto com o campo "value".'
    );
  });

  // Configuração inicial do worksheet.columns, incluindo larguras
  worksheet.columns = preparedColumns.map(({ header, key, width }) => ({
    header,
    key,
    width, // Respeita as larguras manuais na configuração inicial
  }));

  if (config.fixed) {
    const headerRow = config.row || 1;
    worksheet.views = [{ state: 'frozen', ySplit: headerRow }];
  }

  // Aplicar estilos (globais e específicos) ao cabeçalho
  preparedColumns.forEach((column, index) => {
    const cell = worksheet.getRow(config.headerRow || 1).getCell(index + 1);

    // Estilo global
    if (config.style) {
      if (config.style.font) {
        cell.font = { ...cell.font, bold: config.style.font === 'bold' };
      }
      if (config.style.alignment) {
        cell.alignment = {
          ...cell.alignment,
          horizontal: config.style.alignment,
        };
      }
      if (config.style.textRotation) {
        cell.alignment = {
          ...cell.alignment,
          textRotation: config.style.textRotation,
        };
      }
    }

    // Estilo específico da coluna
    if (column.style) {
      if (column.style.font === 'bold') {
        cell.font = { ...cell.font, bold: true };
      }
      if (column.style.alignment) {
        cell.alignment = {
          ...cell.alignment,
          horizontal: column.style.alignment,
        };
      }
    }

    // Rotação específica da coluna
    if (column.rotation) {
      cell.alignment = { ...cell.alignment, textRotation: column.rotation };
    }
  });

  if (typeof config.filter == 'undefined' || config.filter == true) {
    enableColumnFilters(worksheet, preparedColumns, config?.row);
  }
}

/**
 * Salva um workbook Excel no diretório especificado.
 *
 * @param {Object} workbook - Instância de `ExcelJS.Workbook` a ser salva.
 * @param {string} directory - Caminho completo onde o arquivo será salvo.
 *
 * @throws {Error} Caso ocorra algum problema ao salvar o arquivo.
 */
async function saveXlsxFile(workbook, directory) {
  try {
    await workbook.xlsx.writeFile(directory);
  } catch (error) {
    throw new Error(
      'Não foi possível salvar a planilha no diretório informado: ' +
        error.message
    );
  }
}

/**
 * Função assíncrona para criar um arquivo Excel (.xlsx) com uma ou múltiplas abas.
 *
 * @param {string | Array} sheetConfigOrName - Nome da aba (string) para uma única aba ou array de configurações de abas para múltiplas abas.
 * Cada configuração de aba deve conter os campos `sheetName`, `columns` e `rows`.
 * @param {Array} columns - Definição das colunas para a aba única ou o diretório de saída, caso `sheetConfigOrName` seja um array.
 * @param {Array} rows - Linhas de dados para a aba única ou configurações adicionais, caso `sheetConfigOrName` seja um array.
 * @param {string} directory - Diretório para salvar o arquivo .xlsx.
 * @param {Object} config - Configurações opcionais de estilo ou propriedades para cada aba (default: {}).
 *
 * @returns {Promise<void>} Retorna uma promessa que cria e salva o arquivo Excel.
 */
async function createExcelXlsx(
  sheetConfigOrName,
  columns,
  rows,
  directory,
  config = {}
) {
  try {
    const workbook = await setExcelWorkbook();

    // Verifica se é um array, o que indica múltiplas abas
    if (Array.isArray(sheetConfigOrName)) {
      // Modo de múltiplas abas
      const sheetConfigs = sheetConfigOrName;
      directory = columns; // Define `directory` a partir do segundo parâmetro
      config = rows || {}; // Define `config` a partir do terceiro parâmetro, caso exista

      sheetConfigs.forEach(({ sheetName, columns, rows }) => {
        const worksheet = workbook.addWorksheet(sheetName);
        configureSheet(worksheet, columns, rows, config);
      });
    } else {
      // Modo de uma única aba
      const worksheet = await setWorksheet(workbook, sheetConfigOrName);
      configureSheet(worksheet, columns, rows, config);
    }

    await saveXlsxFile(workbook, directory);

    const dataAtual = new Date();
    const dia = dataAtual.getDate();
    const mes = dataAtual.getMonth() + 1; // Mês começa em 0
    const ano = dataAtual.getFullYear();
    const horas = dataAtual.getHours();
    const minutos = dataAtual.getMinutes();
    const segundos = dataAtual.getSeconds();

    console.log(
      `Planilha criada com sucesso! - ${dia}/${mes}/${ano} às ${horas}:${minutos}:${segundos}`
    );
  } catch (error) {
    console.error('Impossível criar planilha: ', error);
  }
}

/**
 * Configura uma planilha Excel com cabeçalho, preenchimento de dados e ajuste de colunas.
 *
 * @param {Object} worksheet - Objeto da planilha fornecido pela biblioteca ExcelJS.
 * @param {Array} columns - Configuração das colunas, podendo ser strings ou objetos com propriedades como header, key e width.
 * @param {Array} rows - Dados das linhas a serem inseridos na planilha. Cada objeto no array deve usar as chaves correspondentes às colunas.
 * @param {Object} [config] - Configurações adicionais para personalização da planilha:
 *   - `header`: Estilos ou configurações específicas para o cabeçalho.
 *   - `ajustColumn`: Define se as larguras das colunas devem ser ajustadas automaticamente (padrão: true).
 */
function configureSheet(worksheet, columns, rows, config) {
  const preparedColumns = columns.map((column) =>
    typeof column === 'string'
      ? { header: column, key: formatTextToIdentifier(column) }
      : {
          header: column.value,
          key: column.key || formatTextToIdentifier(column.value),
          width: column.width,
          style: Array.isArray(column.style)
            ? column.style
            : typeof column.style === 'string'
            ? [column.style]
            : [],
        }
  );

  const columnKeys = preparedColumns.map((col) => col.key);

  setHeaderRow(worksheet, columns, config?.header || {});

  // Preenchimento de dados
  if (rows && rows.length) {
    rows.forEach((rowData) => {
      const row = worksheet.addRow();
      columnKeys.forEach((colKey, index) => {
        const cell = row.getCell(index + 1);
        const matchedValue = rowData[colKey];
        const columnStyle = preparedColumns[index].style || []; // Obter estilo da coluna

        if (matchedValue && typeof matchedValue === 'object') {
          if (matchedValue.style) {
            const cellStyle = Array.isArray(matchedValue.style)
              ? matchedValue.style
              : [matchedValue.style];
            const combinedStyle = [...columnStyle, ...cellStyle];
            applyCellStyle(cell, combinedStyle);
          } else if (columnStyle.length) {
            applyCellStyle(cell, columnStyle); // Aplica apenas estilos da coluna
          }
          if (matchedValue.note) {
            cell.note = { texts: [{ text: matchedValue.note }] };
          }
          cell.value = matchedValue.value;
        } else {
          if (columnStyle.length) {
            applyCellStyle(cell, columnStyle); // Aplica estilos da coluna se não houver configuração específica
          }
          cell.value = matchedValue || null;
        }
      });
    });
  }

  // Ajusta larguras, passando as colunas originais como referência
  if (!config || config.ajustColumn !== false) {
    adjustColumnWidths(worksheet, columns);
  }
}

/**
 * Ajusta automaticamente a largura das colunas em uma planilha Excel.
 *
 * @param {Object} worksheet - Objeto da planilha fornecido pela biblioteca ExcelJS.
 * @param {Array} columns - Configurações de colunas definidas manualmente, contendo informações como largura.
 *
 * Observação:
 * - Colunas com largura manualmente definida nas configurações não serão ajustadas automaticamente.
 * - A largura de cada coluna é ajustada com base no conteúdo das células e no tamanho do cabeçalho.
 */
function adjustColumnWidths(worksheet, columns) {
  worksheet.columns.forEach((column, index) => {
    const originalConfig = columns[index];

    if (
      originalConfig &&
      typeof originalConfig === 'object' &&
      originalConfig.width
    ) {
      // Respeita larguras definidas manualmente nas configurações originais
      return;
    }

    // Ajusta automaticamente apenas colunas sem largura definida
    let maxLength = column.header ? column.header.length : 10; // Começa com o tamanho do cabeçalho
    column.eachCell({ includeEmpty: true }, (cell) => {
      if (cell.value) {
        const cellLength = cell.value.toString().length;
        if (cellLength > maxLength) {
          maxLength = cellLength;
        }
      }
    });

    column.width = maxLength + 5; // Define a largura ideal com margem
  });
}

/**
 * Habilita filtros automáticos nas colunas de uma planilha Excel.
 *
 * @param {Object} worksheet - Instância de `ExcelJS.Worksheet` onde os filtros serão aplicados.
 * @param {Array} columns - Array contendo as colunas que serão filtráveis.
 * @param {number} [row=1] - Número da linha onde os filtros serão aplicados (padrão: 1).
 */
function enableColumnFilters(worksheet, columns, row = 1) {
  const totalColumns = columns.length;
  if (totalColumns > 0) {
    // Define a área do filtro automático na linha especificada
    worksheet.autoFilter = {
      from: { row: row, column: 1 }, // Início na primeira coluna
      to: { row: row, column: totalColumns }, // Fim na última coluna
    };
  }
}


module.exports = {
  setExcelWorkbook,
  setHeaderRow,
  setWorksheet,
  saveXlsxFile,
  createExcelXlsx,
};
