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

app.post('/process', upload.single('file'), async (req, res) => {
  try {
    const userRequest = req.body.request || 'Проанализируй данные';
    const originalFilename = req.body.filename || 'result.xlsx';
    const filePath = req.file.path;

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.worksheets[0];

    const headerRow = worksheet.getRow(1);
    const headers = [];
    headerRow.eachCell({ includeEmpty: true }, (cell, colNum) => {
      headers[colNum] = cell.value;
    });
    const maxCol = headers.length - 1;

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

    let instruction;

    if (isColorTask) {
      // Для задач про цвета — строим инструкцию сами, без Claude
      instruction = {
        insertPosition: 'end',
        emptyRowsBefore: 3,
        rows: [
          { values: ['Запретные ТС', redCount], fontColor: 'FF0000', bold: true },
          { values: ['Аварийные ТС', blackCount], fontColor: '0000FF', bold: true }
        ]
      };
    } else {
      // Для других задач — спрашиваем Claude
      const claudeResponse = await axios.post(
        'https://api.anthropic.com/v1/messages',
        {
          model: 'claude-sonnet-4-6',
          max_tokens: 2048,
          messages: [{
            role: 'user',
            content: `Файл Excel содержит ${cleanData.length} строк данных (не считая заголовок).
Заголовки: ${JSON.stringify(headers.slice(1))}
Данные: ${JSON.stringify(cleanData, null, 2)}

Запрос пользователя: "${userRequest}"

Верни ТОЛЬКО JSON объект (никакого текста вокруг):
{
  "insertPosition": "end",
  "emptyRowsBefore": 0,
  "rows": [
    { "values": ["текст col1", "значение col2", null], "fontColor": "FF0000", "bold": true }
  ]
}

Правила:
- insertPosition: "start" (сразу после заголовка), "end" (в самый конец), или число N (после N-й строки данных)
- emptyRowsBefore: пустых строк перед вставкой (0-5)
- rows: строки для вставки, values по порядку заголовков
- fontColor: RRGGBB без альфа — "FF0000" красный, "0000FF" синий, "000000" чёрный
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
      try {
        const clean = claudeText.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
        instruction = JSON.parse(clean.substring(clean.indexOf('{'), clean.lastIndexOf('}') + 1));
      } catch(e) {
        instruction = { insertPosition: 'end', emptyRowsBefore: 0, rows: [] };
      }
    }

    // Читаем все строки в память
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
      originalRows.push(cells);
    });

    const emptyBefore = instruction.emptyRowsBefore || 0;
    const newRows = instruction.rows || [];

    let insertAfterIndex;
    const pos = instruction.insertPosition;
    if (pos === 'start') {
      insertAfterIndex = 1;
    } else if (pos === 'end' || pos === undefined) {
      insertAfterIndex = originalRows.length;
    } else {
      insertAfterIndex = Math.min(parseInt(pos) + 1, originalRows.length);
    }

    // Очищаем worksheet
    const maxPossibleRow = originalRows.length + emptyBefore + newRows.length + 5;
    for (let i = 1; i <= maxPossibleRow; i++) {
      const row = worksheet.getRow(i);
      for (let c = 1; c <= maxCol; c++) {
        const cell = row.getCell(c);
        cell.value = null;
        cell.font = null;
        cell.fill = null;
        cell.border = null;
        cell.alignment = null;
        cell.numFmt = null;
      }
    }

    let writeIndex = 1;

    for (let i = 0; i < originalRows.length; i++) {
      if (i === insertAfterIndex) {
        // Пустые строки
        for (let e = 0; e < emptyBefore; e++) {
          const emptyRow = worksheet.getRow(writeIndex++);
          for (let c = 1; c <= maxCol; c++) {
            emptyRow.getCell(c).value = null;
            emptyRow.getCell(c).font = null;
          }
        }
        // Новые строки
        for (const nr of newRows) {
          const row = worksheet.getRow(writeIndex++);
          const vals = nr.values || [];
          const argb = 'FF' + nr.fontColor.toUpperCase();
          for (let c = 0; c < Math.min(vals.length, maxCol); c++) {
            const cell = row.getCell(c + 1);
            cell.value = vals[c] !== undefined ? vals[c] : null;
            cell.style = { font: { bold: nr.bold || false, color: { argb: argb }, size: 12 } };
          }
        }
      }

      // Оригинальная строка
      const origCells = originalRows[i];
      const row = worksheet.getRow(writeIndex++);
      for (let c = 0; c < origCells.length; c++) {
        const src = origCells[c];
        const dst = row.getCell(c + 1);
        dst.value = src.value;
        if (src.font) dst.font = src.font;
        if (src.fill) dst.fill = src.fill;
        if (src.border) dst.border = src.border;
        if (src.alignment) dst.alignment = src.alignment;
        if (src.numFmt) dst.numFmt = src.numFmt;
      }
    }

    // Если вставка в конец
    if (insertAfterIndex >= originalRows.length) {
      for (let e = 0; e < emptyBefore; e++) {
        const emptyRow = worksheet.getRow(writeIndex++);
        for (let c = 1; c <= maxCol; c++) {
          emptyRow.getCell(c).value = null;
          emptyRow.getCell(c).font = null;
        }
      }
      for (const nr of newRows) {
        const row = worksheet.getRow(writeIndex++);
        const vals = nr.values || [];
        const argb = 'FF' + nr.fontColor.toUpperCase();
        for (let c = 0; c < Math.min(vals.length, maxCol); c++) {
          const cell = row.getCell(c + 1);
          cell.value = vals[c] !== undefined ? vals[c] : null;
          cell.style = { font: { bold: nr.bold || false, color: { argb: argb }, size: 12 } };
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
