# 📊 Xlsx Engine (Node.js)

Um módulo poderoso e modular para **leitura, criação, formatação e conversão de planilhas** Excel (`.xlsx`) e CSV com Node.js, usando [ExcelJS](https://github.com/exceljs/exceljs).

---

## 📦 Funcionalidades

- ✅ Leitura de arquivos `.xlsx`
- ✅ Conversão de `.xlsx` para JSON
- ✅ Conversão de `.csv` para `.xlsx`
- ✅ Criação de planilhas com múltiplas abas
- ✅ Estilos personalizados em células
- ✅ Aplicação de filtros, larguras e estilos automáticos
- ✅ Organização modular (responsabilidade separada por arquivo)

---

## 🚀 Instalação

Requisitos:
- Node.js v14+
- Biblioteca `exceljs`
- Biblioteca `unidecode` (para identificadores)

Instale as dependências:

```bash
npm install exceljs unidecode
```

---

## 🧠 Estrutura dos arquivos

```
src/
├── parser/           # Leitura, extração e transformação de planilhas
│   ├── reader.js
│   ├── extractor.js
│   ├── formatter.js
│   ├── transformer.js
│   └── modifier.js
│
├── creator/          # Criação de planilhas Excel e CSV
│   ├── creator.js
│   ├── csvCreator.js
│   └── styles.js
│
├── utils/            # Utilitários auxiliares
│   ├── pathUtils.js
│   └── dateUtils.js
│
└── Excel.js          # Classe principal para interface externa
```

---

## ✍️ Exemplo de uso

```js
const Excel = require('./src/Excel');
const E = new Excel('planilhas');

(async () => {
  // Converter planilha para JSON
  const resultado = await E.toJson('exemplo.xlsx', 'Dados');
  console.log(resultado.data);

  // Criar nova planilha
  await E.create('Relatório', ['Nome', 'Idade'], [
    { Nome: 'João', Idade: 30 },
    { Nome: 'Maria', Idade: 25 }
  ]);

  // Converter CSV para XLSX
  await excel.toXlsx('usuarios.csv');
})();
```

---

## ⚙️ Configurações extras

- Estilos de cabeçalho predefinidos no objeto `headerPredefinitions`
- Identificadores limpos com `toIdentifier()`
- Criação automática de diretórios e arquivos temporários

---

## 📄 Licença

Este projeto é livre para uso e modificação interna. Adapte conforme necessário para o seu contexto.

---

