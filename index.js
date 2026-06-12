const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const axios = require('axios');
const fs = require('fs');
const path = require('path');

const app = express();
const upload = multer({ dest: 'uploads/' });

app.use(express.json());

app.post('/process', upload.single('file'), async (req, res) => {
  try {
    const userRequest = req.body.request || 'Проанализируй данные';
    const originalFilename = req.body.filename || 'result.xlsx';
    const filePath = req.file.path;

    // Читаем Excel
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet);

    // Отправляем в Claude
    const claudeResponse = await axios.post(
      'https://api.anthropic.com/v1/messages',
      {
        model: 'claude-sonnet-4-6',
        max_tokens: 4096,
        messages: [
          {
            role: 'user',
            content: `Вот данные из Excel файла в формате JSON:\n${JSON.stringify(data, null, 2)}\n\nЗапрос пользователя: ${userRequest}\n\nВАЖНО: Верни ТОЛЬКО валидный JSON массив. Никакого текста до или после. Никаких markdown блоков. Только [ ... ]. Все объекты должны иметь те же ключи что и входные данные. Если нужно добавить итоговую строку — добавь объект с теми же ключами, где ненужные поля оставь null.`
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

    // Парсим ответ Claude
    const claudeText = claudeResponse.data.content[0].text;
    
    let resultData;
    try {
      const clean = claudeText.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
      const start = clean.indexOf('[');
      const end = clean.lastIndexOf(']');
      const jsonStr = clean.substring(start, end + 1);
      resultData = JSON.parse(jsonStr);
    } catch(e) {
      resultData = data;
      resultData.push({ '_Ошибка': 'Claude не смог обработать: ' + e.message });
    }

    // Создаём новый Excel
    const newWorkbook = XLSX.utils.book_new();
    const newSheet = XLSX.utils.json_to_sheet(resultData);
    XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'Result');

    const outputPath = path.join('uploads', `result_${Date.now()}.xlsx`);
    XLSX.writeFile(newWorkbook, outputPath);

    // Отправляем файл
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
