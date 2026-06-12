const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const axios = require('axios');
const fs = require('fs');
const path = require('path');

const app = express();
const upload = multer({ dest: 'uploads/' });

app.use(express.json());

// Определяем цвет шрифта строки
function getRowFontColor(row) {
  for (const cell of row.values) {
    if (!cell || typeof cell !== 'object') continue;
    const font = cell.font;
    if (font && font.color) {
      const argb = font.color.argb || '';
      const rgb = argb.length === 8 ? argb.slice(2).toUpperCase() : argb.toUpperCase();
      return rgb;
    }
  }
  return null;
}

function isRed(rgb) {
  if (!rgb) return false;
  // Красный: FF0000 или близкие оттенки
  const r = parseInt(rgb.slice(0, 2), 16);
  const g = parseInt(rgb.slice(2, 4), 16);
  const b = parseInt(rgb.slice(4, 6), 16);
  return r > 150 && g < 100 && b < 100;
}

app.post('/process', upload.single('file'), async (req, res) => {
  try {
    const userRequest = req.body.request || 'Проанализируй данные';
    const originalFilename = req.body.filename || 'result.xlsx';
    const filePath = req.file.path;

    // Читаем Excel через ExcelJS (сохраняет форматирование)
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.worksheets[0];

    // Собираем данные с цветами
    const data = [];
    let headers = [];
    let redCount = 0;
    let blackCount = 0;

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) {
        headers = row.values.slice(1); // убираем первый пустой элемент
        return;
      }
      const rowData = {};
      row.values.slice(1).forEach((val, i) => {
        rowData[headers[i]] = (val && typeof val === 'object' && val.text) ? val.text : val;
      });

      // Определяем цвет первой непустой ячейки строки
      let rowColor = null;
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (!rowColor && cell.font && cell.font.color && cell.font.color.argb) {
          rowColor = cell.font.color.argb.slice(2).toUpperCase();
        }
      });

      if (isRed(rowColor)) {
        redCount++;
        rowData['_цвет'] = 'red';
      } else {
        blackCount++;
        rowData['_цвет'] = 'black';
      }

      data.push(rowData);
    });

    // Проверяем тип задачи
    const requestLower = userRequest.toLowerCase();
    const isColorTask = requestLower.includes('красн') || requestLower.includes('черн') || 
                        requestLower.includes('цвет') || requestLower.includes('шрифт');

    let resultData;

    if (isColorTask) {
      // Задача про цвета — считаем сами, без Claude
      const cleanData = data.map(({ _цвет, ...rest }) => rest); // убираем служебное поле
      resultData = cleanData;

      // Добавляем итоговые строки
      const summaryRed = {};
      const summaryBlack = {};
      headers.forEach((h, i) => {
        if (i === 0) {
          summaryRed[h] = 'Запретные ТС';
          summaryBlack[h] = 'Аварийные ТС';
        } else if (i === 1) {
          summaryRed[h] = redCount;
          summaryBlack[h] = blackCount;
        } else {
          summaryRed[h] = null;
          summaryBlack[h] = null;
        }
      });
      resultData.push(summaryRed);
      resultData.push(summaryBlack);

    } else {
      // Остальные задачи — отправляем в Claude
      const cleanData = data.map(({ _цвет, ...rest }) => rest);

      const claudeResponse = await axios.post(
        'https://api.anthropic.com/v1/messages',
        {
          model: 'claude-sonnet-4-6',
          max_tokens: 4096,
          messages: [
            {
              role: 'user',
              content: `Вот данные из Excel файла в формате JSON:\n${JSON.stringify(cleanData, null, 2)}\n\nЗапрос пользователя: ${userRequest}\n\nВерни ТОЛЬКО один JSON объект (не массив!) для добавления итоговой строки. Объект должен содержать только нужные ключи из данных, остальные оставь без упоминания. Никакого текста, только { ... }.`
            }
          ]
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
        const start = clean.indexOf('{');
        const end = clean.lastIndexOf('}');
        summaryRow = JSON.parse(clean.substring(start, end + 1));
      } catch(e) {
        summaryRow = { [headers[0]]: 'ИТОГО', '_info': claudeText.substring(0, 200) };
      }

      resultData = cleanData;
      resultData.push(summaryRow);
    }

    // Создаём новый Excel через XLSX
    const newWorkbook = XLSX.utils.book_new();
    const newSheet = XLSX.utils.json_to_sheet(resultData);
    XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'Result');

    const outputPath = path.join('uploads', `result_${Date.now()}.xlsx`);
    XLSX.writeFile(newWorkbook, outputPath);

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
