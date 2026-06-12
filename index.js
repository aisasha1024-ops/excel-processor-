const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const axios = require('axios');
const fs = require('fs');
const path = require('path');

const app = express();
const upload = multer({ dest: 'uploads/' });
app.use(express.json());

// Проверяем красный ли шрифт в строке
function isRowRed(row) {
  for (const cell of row) {
    if (!cell || !cell.font || !cell.font.color) continue;
    const color = cell.font.color;
    // indexed color: FFFF0000
    if (color.index && String(color.index).toUpperCase().includes('FF0000')) return true;
    // argb color
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

    // Читаем файл через ExcelJS (сохраняет цвета)
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

    // Собираем данные и считаем цвета
    const data = [];
    let redCount = 0;
    let blackCount = 0;

    worksheet.eachRow((row, rowNum) => {
      if (rowNum === 1) return;
      const cells = [];
      row.eachCell({ includeEmpty: true }, (cell) => cells.push(cell));
      if (!cells[0] || cells[0].value === null) return;

      const rowData = {};
      for (let c = 1; c <= maxCol; c++) {
        const cell = row.getCell(c);
        let val = cell.value;
        if (val && typeof val === 'object' && val.text) val = val.text;
        if (val && typeof val === 'object' && val.result !== undefined) val = val.result;
        rowData[headers[c]] = val;
      }

      if (isRowRed(cells)) {
        redCount++;
        rowData['__color'] = 'red';
      } else {
        blackCount++;
        rowData['__color'] = 'black';
      }
      data.push(rowData);
    });

    const requestLower = userRequest.toLowerCase();
    const isColorTask = requestLower.includes('красн') || requestLower.includes('черн') ||
                        requestLower.includes('цвет') || requestLower.includes('шрифт') ||
                        requestLower.includes('запретн') || requestLower.includes('аварийн');

    // Создаём результирующий файл через ExcelJS (чтобы сохранить форматирование)
    const resultWorkbook = new ExcelJS.Workbook();
    const resultSheet = resultWorkbook.addWorksheet('Result');

    // Добавляем заголовки
    const cleanHeaders = [];
    for (let c = 1; c <= maxCol; c++) {
      if (headers[c] !== '__color') cleanHeaders.push(headers[c]);
    }
    resultSheet.addRow(cleanHeaders);

    // Добавляем данные
    for (const row of data) {
      const rowArr = cleanHeaders.map(h => row[h] ?? null);
      resultSheet.addRow(rowArr);
    }

    if (isColorTask) {
      // Отступ 3 строки
      resultSheet.addRow([]);
      resultSheet.addRow([]);
      resultSheet.addRow([]);

      // Строка "Запретные ТС" — красный жирный
      const redRow = resultSheet.addRow([`Запретные ТС`, redCount]);
      redRow.eachCell((cell) => {
        cell.font = { bold: true, color: { argb: 'FFFF0000' }, size: 12 };
      });

      // Строка "Аварийные ТС" — синий жирный
      const blackRow = resultSheet.addRow([`Аварийные ТС`, blackCount]);
      blackRow.eachCell((cell) => {
        cell.font = { bold: true, color: { argb: 'FF0000FF' }, size: 12 };
      });

    } else {
      // Другие задачи — отправляем в Claude
      const cleanData = data.map(({ __color, ...rest }) => rest);
      const claudeResponse = await axios.post(
        'https://api.anthropic.com/v1/messages',
        {
          model: 'claude-sonnet-4-6',
          max_tokens: 4096,
          messages: [{
            role: 'user',
            content: `Вот данные из Excel файла в формате JSON:\n${JSON.stringify(cleanData, null, 2)}\n\nЗапрос пользователя: ${userRequest}\n\nВерни ТОЛЬКО один JSON объект (не массив!) для добавления итоговой строки. Объект должен содержать только нужные ключи из данных. Никакого текста, только { ... }.`
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
      let summaryRow;
      try {
        const clean = claudeText.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
        summaryRow = JSON.parse(clean.substring(clean.indexOf('{'), clean.lastIndexOf('}') + 1));
      } catch(e) {
        summaryRow = { [cleanHeaders[0]]: 'ИТОГО' };
      }
      resultSheet.addRow(cleanHeaders.map(h => summaryRow[h] ?? null));
    }

    // Авторазмер первых 3 колонок
    [1, 2, 3].forEach(colNum => {
      let maxLen = 10;
      resultSheet.getColumn(colNum).eachCell({ includeEmpty: false }, cell => {
        const len = String(cell.value || '').length;
        if (len > maxLen) maxLen = len;
      });
      resultSheet.getColumn(colNum).width = Math.min(maxLen + 2, 50);
    });

    // Сохраняем
    const outputPath = path.join('uploads', `result_${Date.now()}.xlsx`);
    await resultWorkbook.xlsx.writeFile(outputPath);

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
