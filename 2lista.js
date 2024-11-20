document.getElementById('search-button').addEventListener('click', async () => {
    const fileUrl = '/SheetJS-git/file/Jobs_kalendar.xlsx'; // Укажите URL-адрес Excel файла
    const jsonFileUrl = '/SheetJS-git/json/jobs.json'; // Укажите URL-адрес json файла
    const searchDateInput = document.getElementById('search-date').value;

    if (!searchDateInput) {
        alert('Укажите дату для поиска!');
        return;
    }

    try {
        // Преобразуем дату в числовой формат Excel
        const searchDate = dateToExcelDate(new Date(searchDateInput));

        // Загружаем файл по URL
        const response = await fetch(fileUrl);
        if (!response.ok) {
            throw new Error('Не удалось загрузить файл');
        }

        const fileData = await response.arrayBuffer();
        const workbook = XLSX.read(fileData, { type: 'array' });


        // Загружаем JSON файл
        const jsonResponse = await fetch(jsonFileUrl);
        if (!jsonResponse.ok) {throw new Error('Не удалось загрузить JSON файл');}
        const jsonData = await jsonResponse.json();

        const results = [];

        // Обрабатываем каждый лист
        for (const sheetName of workbook.SheetNames) {
            const sheet = workbook.Sheets[sheetName];
            const sheetData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
            console.log(`Обрабатываем лист: ${sheetName}`); // Лог текущего листа
            if (sheetData.length === 0) {
                results.push(`${sheetName}: Лист пустой`);                
                continue;
            }

            let columnIndex = -1;

            // Ищем дату в первой строке (заголовке)
            const targetRow = sheetData[0];
            targetRow.forEach((cell, index) => {
                if (cell === searchDate) {
                    columnIndex = index;
                }
            });
            

            if (columnIndex === -1) {
                results.push(`${sheetName}: Дата не найдена`);
                continue;
            }

            // Проверяем, есть ли значение "1" в найденном столбце
            const colData = sheetData
                .slice(2) // Пропускаем заголовок
                .map(row => row[columnIndex]); // Значения из найденного столбца
            const hasOne = colData.some(value => value === 1);
            
            console.log(`Количество найденных точек на листе "${sheetName}":`, colData[0]);
            if (hasOne) {
                // Выводим значения из столбца B
                const columnBData = sheetData
                    .slice(2) // Пропускаем заголовок
                    .filter((row, index) => colData[index] === 1) // Берем только строки, где в найденном столбце "1"
                    .map(row => row[1]); // Значения столбца B
                    // Получаем данные из JSON только по существующим ключам
                    const filteredJsonResults = columnBData
                    .filter(key => jsonData[sheetName][key] !== undefined) // Оставляем только те ключи, которые есть в JSON
                    .map(key => `${key}: position: ${jsonData[sheetName][key].position}, vycka: ${jsonData[sheetName][key].vycka}, date: ${jsonData[sheetName][key].date}, JTSK: ${jsonData[sheetName][key].systemCoordinates}, positionType: ${jsonData[sheetName][key].positionType}`);
                    if (filteredJsonResults.length > 0) {
                        results.push(`${sheetName}:\n` + filteredJsonResults.join('\n'));
                    }
      
            } else {
                results.push(`${sheetName}: В столбце нет значения "1".`);
            }
        }

        // Форматируем вывод и отображаем на странице
        const outputElement = document.getElementById('output'); 
        outputElement.innerText = results.join('\n\n\n');      

    } catch (error) {
        console.error('Ошибка при обработке файла:', error);
        alert('Ошибка при обработке файла. Проверьте файл и повторите попытку.');
    }
});

// Функция преобразования даты в формат Excel
function dateToExcelDate(date) {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30)); // Excel "нулевой день"
    const dayInMilliseconds = 24 * 60 * 60 * 1000;
    return Math.floor((date - excelEpoch) / dayInMilliseconds);
}