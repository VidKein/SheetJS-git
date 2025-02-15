// URL файла Excel на сервере
const FILE_URL = '/SheetJS-git/file/Jobs_kalendar.xlsx'; // Укажите путь к файлу

// Функция преобразования даты в формат Excel
const dateToExcelNumber = (date) => {
        const jsDate = new Date(date);
        return Math.floor((jsDate - new Date(1899, 11, 30)) / (1000 * 60 * 60 * 24));
    };

    // Обработчик поиска
    const searchHandler = () => {
        const searchDateInput = document.getElementById('search-date').value; // Получаем дату из ввода
        if (!searchDateInput) {
            document.getElementById('output').textContent = 'Пожалуйста, введите дату.';
            return;
        }

        const excelDate = dateToExcelNumber(searchDateInput); // Преобразуем дату в Excel-формат
        fetch(FILE_URL)
            .then(response => response.arrayBuffer())
            .then(data => {
                const workbook = XLSX.read(data, { type: 'array' });

                const sheetName = workbook.SheetNames[0];
                const sheet = workbook.Sheets[sheetName];

                const output = document.getElementById('output');

                // Поиск даты в строке 1
                const range = XLSX.utils.decode_range(sheet['!ref']);
                const startCol = range.s.c;
                const endCol = range.e.c;

                let columnIndex = -1; // Идентификатор столбца с датой
                for (let col = startCol; col <= endCol; col++) {
                    const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col }); // Строка 1, номер колонки
                    const cell = sheet[cellAddress];
                    if (cell && cell.v === excelDate) {
                        columnIndex = col;
                        break;
                    }
                }

                if (columnIndex === -1) {
                    output.textContent = 'Дата не найдена.';
                } else {
                    output.textContent = `Дата найдена в столбце ${XLSX.utils.encode_col(columnIndex)}\n`;
                    // Проверяем, есть ли число 1 в найденном столбце
                    let foundOne = false;
                    let rowCount = 0; // Счетчик строк с данными из столбца B

                    for (let row = range.s.r + 1; row <= range.e.r; row++) { // Пропускаем строку 1
                        const cellAddress = XLSX.utils.encode_cell({ r: row, c: columnIndex });
                        const cell = sheet[cellAddress];
                        if (cell && cell.v === 1) {
                            foundOne = true;
                            //output.textContent += `Найдено число 1 в строке ${row + 1}.\n`;

                            // Выводим содержимое столбца B для этой строки
                            const colBAddress = XLSX.utils.encode_cell({ r: row, c: 1 }); // Столбец B
                            const colBCell = sheet[colBAddress];
                            console.log(colBCell);
                            if (colBCell !== undefined) {
                              output.textContent += `Строка ${row + 1}: ${colBCell ? colBCell.v : 'Пусто'}\n`;
                              rowCount++; // Увеличиваем счетчик строк  
                            }
                        }
                    }

                    if (!foundOne) {
                        output.textContent += `Число 1 не найдено в столбце ${XLSX.utils.encode_col(columnIndex)}.\n`;
                    }else {
                        output.textContent += `\nОбщее количество строк, выведенных из столбца В: ${rowCount}\n`;
                    }
                }
            })
            .catch(error => {
                console.error('Ошибка при загрузке файла:', error);
                document.getElementById('output').textContent = 'Не удалось загрузить файл.';
            });
    };

    // Привязываем обработчик к кнопке
    document.getElementById('search-button').addEventListener('click', searchHandler);