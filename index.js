const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const axios = require('axios');
const fs = require('fs');
const path = require('path');

const app = express();
const upload = multer({ dest: 'uploads/' });
app.use(express.json());

function isRowRed(row, maxCol) {
  for (let c = 1; c <= maxCol; c++) {
    const cell = row.getCell(c);
    if (!cell || !cell.font || !cell.font.color) continue;
    const color = cell.font.color;
    if (color.index && String(color.index).toUpperCase().includes('FF0000')) return true;
    if (color.argb) {
      const rgb = color.argb.slice(2).toUpperCase();
      const r = parseInt(rgb.slice(0,2), 16);
      const g = parseInt(rgb.slice(2,4), 16);
      const b = parseInt(rgb.slice(4,6), 16);
      if (r > 150 && g < 80 && b < 80) return true;
    }
  }
  return false;
}

// Копируем стиль ячейки
function copyCellStyle(src, dst) {
  if (src.font) dst.font = JSON.parse(JSON.stringify(src.font));
  if (src.fill) dst.fill = JSON.parse(JSON.stringify(src.fill));
  if (src.border) dst.border = JSON.parse(JSON.stringify(src.border));
  if (src.alignment) dst.alignment = JSON.parse(JSON.stringify(src.alignment));
  if (src.numFmt) dst.numFmt = src.numFmt;
}

app.post('/process', upload.single('file'), async (req, res) => {
  try {
    const userRequest = req.body.request || 'Проанализируй данные';
    const originalFilename = req.body.filename || 'result.xlsx';
    const filePath = req.file.path;

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.worksheets[0];

    // Собираем заголовки
    const headerRow = worksheet.getRow(1);
    const headers = [];
    headerRow.eachCell({ includeEmpty: true }, (cell, colNum) => {
      headers[colNum] = cell.value;
    });
    const maxCol = headers.length - 1;

    // Считаем цвета и собираем данные
    let redCount = 0;
    let blackCount = 0;
    const cleanData = [];

    worksheet.eachRow((row, rowNum) => {
      if (rowNum === 1) return;
      const firstCell = row.getCell(1);
      if (!firstCell || firstCell.value === null || firstCell.value === undefined) return;

      if (isRowRed(row, maxCol)) redCount++; else blackCount++;

      const rowData = {};
      for (let c = 1; c <= maxCol; c++) {
        let val = row.getCell(c).value;
        if (val && typeof val === 'object' && val.text) val = val.text;
        if (val && typeof val === 'object' && val.result !== undefined) val = val.result;
        rowData[headers[c]] = val;
      }
      cleanData.push(rowData);
    });

    const requestLower = userRequest.toLowerCase();
    const isColorTask = requestLower.includes('красн') || requestLower.includes('черн') ||
                        requestLower.includes('цвет') || requestLower.includes('шрифт') ||
                        requestLower.includes('запретн') || requestLower.includes('аварийн');

    // Спрашиваем Claude: что добавить и куда вставить
    const claudeResponse = await axios.post(
      'https://api.anthropic.com/v1/messages',
      {
        model: 'claude-sonnet-4-6',
        max_tokens: 2048,
        messages: [{
          role: 'user',
          content: `Файл Excel содержит ${cleanData.length} строк данных (не считая заголовок).
Заголовки: ${JSON.stringify(headers.slice(1))}
${isColorTask ? `Строк с красным шрифтом: ${redCount}. Строк с чёрным шрифтом: ${blackCount}.` : `Данные: ${JSON.stringify(cleanData, null, 2)}`}

Запрос пользователя: "${userRequest}"

Верни ТОЛЬКО JSON объект такого формата (никакого текста вокруг):
{
  "insertPosition": "end",
  "emptyRowsBefore": 0,
  "rows": [
    { "values": ["текст col1", "значение col2", null, ...], "fontColor": "FF0000", "bold": true }
  ]
}

Правила:
- insertPosition: "start" (после заголовка), "end" (в конец), или число (номер строки данных после которой вставить, 1 = после первой строки данных)
- emptyRowsBefore: количество пустых строк перед вставкой (0-5)
- rows: массив строк для вставки
- values: массив значений для каждой колонки (по порядку заголовков), null для пустых
- fontColor: цвет шрифта в формате RRGGBB (без FF впереди), например "FF0000" для красного, "0000FF" для синего, "000000" для чёрного
- bold: true или false`
        }]
      },
      {
        headers: {
          'x-api-key': process.env.ANTHROPIC_API_KEY,
          'anthropic-version': '2023-06-01',
          'Content-Type': 'application/json'
        }
      }
    );

    const claudeText = claudeResponse.data.content[0].text;
    let instruction;
    try {
      const clean = claudeText.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
      instruction = JSON.parse(clean.substring(clean.indexOf('{'), clean.lastIndexOf('}') + 1));
    } catch(e) {
      instruction = { insertPosition: 'end', emptyRowsBefore: 0, rows: [] };
    }

    // Читаем все строки оригинала в память (с форматированием)
    const originalRows = [];
    worksheet.eachRow({ includeEmpty: false }, (row, rowNum) => {
      const cells = [];
      for (let c = 1; c <= maxCol; c++) {
        const cell = row.getCell(c);
        cells.push({
          value: cell.value,
          font: cell.font ? JSON.parse(JSON.stringify(cell.font)) : null,
          fill: cell.fill ? JSON.parse(JSON.stringify(cell.fill)) : null,
          border: cell.border ? JSON.parse(JSON.stringify(cell.border)) : null,
          alignment: cell.alignment ? JSON.parse(JSON.stringify(cell.alignment)) : null,
          numFmt: cell.numFmt || null,
        });
      }
      originalRows.push({ rowNum, cells });
    });

    // Строим новые строки для вставки
    const newRows = [];
    for (let i = 0; i < (instruction.emptyRowsBefore || 0); i++) {
      newRows.push({ isEmpty: true });
    }
    for (const r of (instruction.rows || [])) {
      newRows.push({
        isEmpty: false,
        values: r.values || [],
        fontColor: r.fontColor || '000000',
        bold: r.bold || false
      });
    }

    // Определяем позицию вставки (индекс в originalRows после заголовка)
    let insertAfterIndex; // индекс в originalRows
    const pos = instruction.insertPosition;
    if (pos === 'start') {
      insertAfterIndex = 0; // после заголовка (индекс 0)
    } else if (pos === 'end' || pos === undefined) {
      insertAfterIndex = originalRows.length; // в самый конец
    } else {
      // число = после N-й строки данных (1-based), заголовок = индекс 0
      insertAfterIndex = Math.min(parseInt(pos) + 1, originalRows.length);
    }

    // Пересобираем worksheet
    // Сначала очищаем все строки
    const totalNewRows = originalRows.length + newRows.length;
    for (let i = 1; i <= totalNewRows + 5; i++) {
      const row = worksheet.getRow(i);
      for (let c = 1; c <= maxCol; c++) row.getCell(c).value = null;
    }

    // Записываем оригинальные строки + вставляем новые
    let writeIndex = 1;
    for (let i = 0; i < originalRows.length; i++) {
      // Если это позиция вставки — вставляем новые строки
      if (i === insertAfterIndex) {
        for (const nr of newRows) {
          const row = worksheet.getRow(writeIndex++);
          if (!nr.isEmpty) {
            for (let c = 0; c < Math.min(nr.values.length, maxCol); c++) {
              const cell = row.getCell(c + 1);
              cell.value = nr.values[c] !== undefined ? nr.values[c] : null;
              cell.font = { bold: nr.bold, color: { argb: 'FF' + nr.fontColor }, size: 12 };
            }
          }
        }
      }
      // Пишем оригинальную строку
      const origRow = originalRows[i];
      const row = worksheet.getRow(writeIndex++);
      for (let c = 0; c < origRow.cells.length; c++) {
        const srcCell = origRow.cells[c];
        const dstCell = row.getCell(c + 1);
        dstCell.value = srcCell.value;
        if (srcCell.font) dstCell.font = srcCell.font;
        if (srcCell.fill) dstCell.fill = srcCell.fill;
        if (srcCell.border) dstCell.border = srcCell.border;
        if (srcCell.alignment) dstCell.alignment = srcCell.alignment;
        if (srcCell.numFmt) dstCell.numFmt = srcCell.numFmt;
      }
    }

    // Если вставка в конец
    if (insertAfterIndex >= originalRows.length) {
      for (const nr of newRows) {
        const row = worksheet.getRow(writeIndex++);
        if (!nr.isEmpty) {
          for (let c = 0; c < Math.min(nr.values.length, maxCol); c++) {
            const cell = row.getCell(c + 1);
            cell.value = nr.values[c] !== undefined ? nr.values[c] : null;
            cell.font = { bold: nr.bold, color: { argb: 'FF' + nr.fontColor }, size: 12 };
          }
        }
      }
    }

    // Авторазмер первых 3 колонок
    [1, 2, 3].forEach(colNum => {
      let maxLen = 10;
      worksheet.getColumn(colNum).eachCell({ includeEmpty: false }, cell => {
        const len = String(cell.value || '').length;
        if (len > maxLen) maxLen = len;
      });
      worksheet.getColumn(colNum).width = Math.min(maxLen + 2, 50);
    });

    const outputPath = path.join('uploads', `result_${Date.now()}.xlsx`);
    await workbook.xlsx.writeFile(outputPath);

    const resultName = 'result_' + originalFilename;
    res.download(outputPath, resultName, () => {
      fs.unlinkSync(filePath);
      fs.unlinkSync(outputPath);
    });

  } catch (error) {
    console.error(error);
    res.status(500).json({ error: error.message });
  }
});

app.get('/', (req, res) => {
  res.json({ status: 'Excel Processor is running' });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
