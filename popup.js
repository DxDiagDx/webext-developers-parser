let developersData = [];

document.getElementById('parseBtn').addEventListener('click', startParsing);
document.getElementById('exportBtn').addEventListener('click', exportToExcel);

async function startParsing() {
  const statusDiv = document.getElementById('status');
  const progressDiv = document.getElementById('progress');
  const progressBar = document.getElementById('progressBar');
  const parseBtn = document.getElementById('parseBtn');
  const exportBtn = document.getElementById('exportBtn');
  
  parseBtn.disabled = true;
  statusDiv.style.display = 'block';
  progressDiv.style.display = 'block';
  statusDiv.className = 'loading';
  statusDiv.textContent = 'Начинаем парсинг данных...';
  
  developersData = [];
  let offset = 0;
  const limit = 1000;
  let hasMoreData = true;
  let totalProcessed = 0;
  
  try {
    // Первый запрос для получения общего количества
    const firstUrl = `https://наш.дом.рф/сервисы/api/erz/main/filter?offset=0&limit=1&sortField=devShortNm&sortType=asc&objStatus=0`;
    const firstResponse = await fetch(firstUrl);
    const firstData = await firstResponse.json();
    const totalCount = firstData.data?.count || 0;
    
    if (totalCount === 0) {
      throw new Error('Не найдено данных для обработки');
    }
    
    while (hasMoreData) {
      const currentOffset = offset;
      statusDiv.textContent = `Загружаем данные: ${Math.min(currentOffset + limit, totalCount)} из ${totalCount}...`;
      progressBar.style.width = `${(currentOffset / totalCount) * 100}%`;
      
      const url = `https://наш.дом.рф/сервисы/api/erz/main/filter?offset=${offset}&limit=${limit}&sortField=devShortNm&sortType=asc&objStatus=0`;
      
      const response = await fetch(url, {
        method: 'GET',
        headers: {
          'Accept': 'application/json',
          'Content-Type': 'application/json'
        }
      });
      
      if (!response.ok) {
        throw new Error(`HTTP ошибка! status: ${response.status}`);
      }
      
      const data = await response.json();
      
      if (data.errcode !== "0") {
        throw new Error(`API ошибка: ${data.errcode}`);
      }
      
      if (data.data && data.data.developers && data.data.developers.length > 0) {
        // Извлекаем нужные данные из каждого разработчика
        data.data.developers.forEach(dev => {
          developersData.push({
            'Название компании': dev.devShortNm || '',
            'Полное название': dev.devFullCleanNm || '',
            'Регион': dev.regRegionDesc || '',
            'Телефон': dev.devPhoneNum || '',
            'Сайт': dev.devSite || '',
            'Email': dev.devEmail || '',
            'Контактное лицо': dev.devEmplMainFullNm || '',
            'ИНН': dev.devInn || '',
            'ОГРН': dev.devOgrn || '',
            'КПП': dev.devKpp || ''
          });
        });
        
        totalProcessed += data.data.developers.length;
        
        // Проверяем, есть ли еще данные
        if (data.data.developers.length < limit || totalProcessed >= totalCount) {
          hasMoreData = false;
        } else {
          offset += limit;
        }
      } else {
        hasMoreData = false;
      }
      
      // Добавляем небольшую задержку между запросами
      await new Promise(resolve => setTimeout(resolve, 200));
    }
    
    progressBar.style.width = '100%';
    statusDiv.className = 'success';
    statusDiv.textContent = `Парсинг завершен! Обработано ${developersData.length} записей.`;
    
    exportBtn.disabled = false;
    
  } catch (error) {
    statusDiv.className = 'error';
    statusDiv.textContent = `Ошибка: ${error.message}`;
    parseBtn.disabled = false;
    progressDiv.style.display = 'none';
  }
}

function exportToExcel() {
  if (developersData.length === 0) {
    alert('Нет данных для экспорта!');
    return;
  }
  
  try {
    // Создаем новую книгу Excel
    const wb = XLSX.utils.book_new();
    
    // Преобразуем данные в рабочий лист
    const ws = XLSX.utils.json_to_sheet(developersData);
    
    // Настраиваем ширину колонок
    const colWidths = [
      { wch: 25 }, // Название компании
      { wch: 40 }, // Полное название
      { wch: 20 }, // Регион
      { wch: 15 }, // Телефон
      { wch: 20 }, // Сайт
      { wch: 25 }, // Email
      { wch: 25 }, // Контактное лицо
      { wch: 12 }, // ИНН
      { wch: 15 }, // ОГРН
      { wch: 10 }  // КПП
    ];
    ws['!cols'] = colWidths;
    
    // Добавляем рабочий лист в книгу
    XLSX.utils.book_append_sheet(wb, ws, 'Застройщики');
    
    // Генерируем файл
    const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    
    // Создаем Blob и ссылку для скачивания
    const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    
    const link = document.createElement('a');
    link.href = url;
    link.download = `застройщики_${new Date().toISOString().split('T')[0]}.xlsx`;
    document.body.appendChild(link);
    
    // Запускаем скачивание
    link.click();
    
    // Очищаем
    setTimeout(() => {
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    }, 100);
    
    const statusDiv = document.getElementById('status');
    statusDiv.className = 'success';
    statusDiv.textContent = 'Файл Excel успешно экспортирован!';
    
  } catch (error) {
    const statusDiv = document.getElementById('status');
    statusDiv.className = 'error';
    statusDiv.textContent = `Ошибка при экспорте: ${error.message}`;
  }
}