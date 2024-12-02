/** Estilos pré definidos para as células */
const PRE_CELL_STYLES = {
  Bom: {
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'C6EFCE' } },
    font: { color: { argb: '006100' } },
  },
  Ruim: {
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC7CE' } },
    font: { color: { argb: '9C0006' } },
  },
  Neutro: {
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEB9C' } },
    font: { color: { argb: '9C6500' } },
  },
  Centralizado: {
    alignment: { horizontal: 'center' },
  },
};

/**
 * Aplica estilos a uma célula em uma planilha Excel.
 *
 * @param {Object} cell - Objeto da célula fornecido pela biblioteca ExcelJS.
 * @param {string|Array|string[]|Object|Object[]} styleType - Tipo(s) de estilo a ser(em) aplicado(s):
 *   - Pode ser um nome de estilo predefinido (e.g., "Bom", "Ruim").
 *   - Pode ser um array contendo múltiplos estilos predefinidos.
 *   - Pode ser um objeto com propriedades de estilo personalizadas.
 *
 * @example
 * // Aplicar um único estilo predefinido:
 * applyCellStyle(cell, 'Bom');
 *
 * @example
 * // Aplicar múltiplos estilos predefinidos:
 * applyCellStyle(cell, ['Atenção', 'Centralizado']);
 *
 * @example
 * // Aplicar um estilo personalizado:
 * applyCellStyle(cell, { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0000' } } });
 *
 * @throws {Error} Caso um estilo não seja encontrado ou seja inválido.
 */
function applyCellStyle(cell, styleType) {
  const styleArray = Array.isArray(styleType) ? styleType : [styleType];
  styleArray.forEach((style) => {
    if (typeof style === 'string' && PRE_CELL_STYLES[style]) {
      cell.fill = PRE_CELL_STYLES[style]?.fill ?? cell.fill;
      cell.font = PRE_CELL_STYLES[style]?.font ?? cell.font;
      cell.border = PRE_CELL_STYLES[style]?.border ?? cell.border;
      cell.alignment = PRE_CELL_STYLES[style]?.alignment ?? cell.alignment;
    } else if (typeof style === 'object') {
      cell.fill = style.fill ?? cell.fill;
      cell.font = style.font ?? cell.font;
      cell.border = style.border ?? cell.border;
      cell.alignment = style.alignment ?? cell.alignment;
    } else {
      throw new Error(`Estilo inválido ou não encontrado: ${style}`);
    }
  });
}

module.exports = {
  applyCellStyle,
};
