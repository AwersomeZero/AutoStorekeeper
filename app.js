/**
 * Автоскладер (AutoStorekeeper) Web Engine
 * Клиентский конвертер списков материалов Litematica (.txt и .json) в Excel для GitHub Pages
 */

(function () {
  'use strict';

  // Константы
  const DOUBLE_CHEST_SLOTS = 54;
  const SHULKER_BOX_SLOTS = 27;
  const DEFAULT_STACK_SIZE = 64;
  const UNKNOWN_LOCAL_ID = 999999;

  // Состояние приложения
  let sortingDb = null;
  let currentFile = null;
  let convertedRows = null;
  let generatedWorkbook = null;
  let outputFileName = '';

  // Элементы DOM
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('fileInput');
  const openFileBtn = document.getElementById('openFileBtn');
  const fileInfo = document.getElementById('fileInfo');
  const fileNameDisplay = document.getElementById('fileNameDisplay');
  const fileSizeDisplay = document.getElementById('fileSizeDisplay');

  const downloadBtn = document.getElementById('downloadBtn');
  const downloadFilename = document.getElementById('downloadFilename');
  const statRows = document.getElementById('statRows');
  const statChests = document.getElementById('statChests');
  const statShulkers = document.getElementById('statShulkers');

  const consoleBody = document.getElementById('consoleBody');
  const clearLogsBtn = document.getElementById('clearLogsBtn');
  const copyLogsBtn = document.getElementById('copyLogsBtn');

  const tablePlaceholder = document.getElementById('tablePlaceholder');
  const dataTable = document.getElementById('dataTable');
  const tableBody = document.getElementById('tableBody');
  const groupBadgeBig = document.getElementById('groupBadgeBig');
  const groupBadgeMid = document.getElementById('groupBadgeMid');
  const groupBadgeMin = document.getElementById('groupBadgeMin');

  // ==========================================================================
  // Логирование в консоль отладки
  // ==========================================================================
  function log(message, type = 'info') {
    const now = new Date();
    const timeStr = now.toTimeString().split(' ')[0] + '.' + String(now.getMilliseconds()).padStart(3, '0');

    const entry = document.createElement('div');
    entry.className = `log-entry log-${type}`;

    const timeSpan = document.createElement('span');
    timeSpan.className = 'log-time';
    timeSpan.textContent = `[${timeStr}]`;

    entry.appendChild(timeSpan);
    entry.appendChild(document.createTextNode(` ${message}`));

    consoleBody.appendChild(entry);
    consoleBody.scrollTop = consoleBody.scrollHeight;
  }

  // ==========================================================================
  // Инициализация базы данных предметов (sorting_list)
  // ==========================================================================
  async function initDatabase() {
    log('Инициализация базы сортировки предметов...', 'info');

    // 1. Проверяем предварительно загруженные данные через conf/sorting_list.js (без CORS)
    if (window.SORTING_LIST_DATA && Array.isArray(window.SORTING_LIST_DATA)) {
      buildDatabase(window.SORTING_LIST_DATA);
      log(`База данных успешно загружена из встроенного модуля (${window.SORTING_LIST_DATA.length} предметов).`, 'success');
      return;
    }

    // 2. Фоллбэк: загрузка через fetch(conf/sorting_list.json)
    try {
      const resp = await fetch('conf/sorting_list.json');
      if (resp.ok) {
        const json = await resp.json();
        buildDatabase(json);
        log(`База данных загружена из conf/sorting_list.json (${json.length} предметов).`, 'success');
        return;
      }
    } catch (e) {
      // Игнорируем и пробуем следующий способ
    }

    // 3. Фоллбэк: загрузка conf/sorting_list.xlsx через SheetJS
    if (window.XLSX) {
      try {
        const resp = await fetch('conf/sorting_list.xlsx');
        if (resp.ok) {
          const buffer = await resp.arrayBuffer();
          const wb = XLSX.read(buffer, { type: 'array' });
          const firstSheet = wb.Sheets[wb.SheetNames[0]];
          const rawRows = XLSX.utils.sheet_to_json(firstSheet);
          const formatted = rawRows.map((r, idx) => ({
            i: idx,
            ru: r.RU_Name || '',
            en: r.EN_name || '',
            id: r.ID || '',
            size: r.Size || 64
          }));
          buildDatabase(formatted);
          log(`База данных загружена из conf/sorting_list.xlsx (${formatted.length} предметов).`, 'success');
          return;
        }
      } catch (e) {
        // Ошибка загрузки
      }
    }

    log('[!] ВНИМАНИЕ: Не удалось автоматически загрузить conf/sorting_list. "Автоскладер" будет использовать дефолтный размер стака (64).', 'warn');
    buildDatabase([]);
  }

  function buildDatabase(items) {
    const ruMap = new Map();
    const enMap = new Map();
    const idMap = new Map();

    const ruLower = new Map();
    const enLower = new Map();
    const idLower = new Map();
    const shortIdMap = new Map();

    const warnedItems = new Set();

    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      const idx = item.i !== undefined ? item.i : i;
      const ru = String(item.ru || item.RU_Name || '').trim();
      const en = String(item.en || item.EN_name || '').trim();
      const id = String(item.id || item.ID || '').trim();
      const size = parseInt(item.size || item.Size, 10) || DEFAULT_STACK_SIZE;

      const record = { localId: idx, stackSize: size > 0 ? size : DEFAULT_STACK_SIZE, id, ru, en };

      if (ru && !ruMap.has(ru)) {
        ruMap.set(ru, record);
        ruLower.set(ru.toLowerCase(), record);
      }
      if (en && !enMap.has(en)) {
        enMap.set(en, record);
        enLower.set(en.toLowerCase(), record);
      }
      if (id && !idMap.has(id)) {
        idMap.set(id, record);
        idLower.set(id.toLowerCase(), record);
        if (id.includes(':')) {
          const short = id.split(':')[1].toLowerCase();
          if (!shortIdMap.has(short)) {
            shortIdMap.set(short, record);
          }
        }
      }
    }

    sortingDb = {
      getItemInfo(rawName) {
        const name = (rawName || '').trim();
        const lower = name.toLowerCase();

        let info = null;

        // 1. Поиск по ID предмета
        if (idMap.has(name)) {
          info = idMap.get(name);
        } else if (idLower.has(lower)) {
          info = idLower.get(lower);
        } else if (shortIdMap.has(lower)) {
          info = shortIdMap.get(lower);
        } else if (idLower.has('minecraft:' + lower)) {
          info = idLower.get('minecraft:' + lower);
        }
        // 2. Поиск по русскому названию
        else if (ruMap.has(name)) {
          info = ruMap.get(name);
        } else if (ruLower.has(lower)) {
          info = ruLower.get(lower);
        }
        // 3. Поиск по английскому названию
        else if (enMap.has(name)) {
          info = enMap.get(name);
        } else if (enLower.has(lower)) {
          info = enLower.get(lower);
        }

        if (info) {
          const display = info.ru ? info.ru : name;
          return {
            localId: info.localId,
            stackSize: info.stackSize,
            ruName: display
          };
        }

        if (!warnedItems.has(name)) {
          warnedItems.add(name);
          log(`Предмет/ID "${name}" не найден в базе сортировки (назначен стак: ${DEFAULT_STACK_SIZE}, ID: ${UNKNOWN_LOCAL_ID}).`, 'warn');
        }

        return { localId: UNKNOWN_LOCAL_ID, stackSize: DEFAULT_STACK_SIZE, ruName: name };
      }
    };
  }

  // ==========================================================================
  // Парсинг текстового формата MaterialList из Litematica (.txt)
  // ==========================================================================
  function parseMaterialListText(text) {
    const lines = text.split(/\r?\n/);
    const materials = [];

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      if (!line.startsWith('|') || line.startsWith('|+')) continue;

      const cols = line.replace(/^\||\|$/g, '').split('|').map(s => s.trim());
      if (cols.length < 2) continue;

      const itemName = cols[0];
      const totalStr = cols[1];

      // Пропуск заголовков
      const lowerName = itemName.toLowerCase();
      const lowerTotal = totalStr.toLowerCase();
      if (['item', 'предмет', 'название', 'наименование'].includes(lowerName) ||
          ['total', 'всего', 'кол-во', 'количество'].includes(lowerTotal)) {
        continue;
      }

      const cleanTotal = totalStr.replace(/[\s,\u00A0]/g, '');
      if (!/^\d+$/.test(cleanTotal)) continue;

      const total = parseInt(cleanTotal, 10);
      if (itemName) {
        materials.push({ name: itemName, total: total });
      }
    }

    return materials;
  }

  // ==========================================================================
  // Парсинг JSON формата MaterialList из Litematica (.json)
  // ==========================================================================
  function parseMaterialListJson(jsonText) {
    const data = JSON.parse(jsonText);
    const rawItems = data.Materials || data.materials || data.Items || data.items || data;
    const materials = [];

    if (Array.isArray(rawItems)) {
      for (const entry of rawItems) {
        if (!entry || typeof entry !== 'object') continue;
        const itemId = entry.Item || entry.item || entry.ID || entry.id || entry.Name || entry.name;
        const totalVal = entry.Total !== undefined ? entry.Total : (entry.total !== undefined ? entry.total : entry.count);
        if (itemId !== undefined && totalVal !== undefined) {
          const cleanTotal = parseInt(String(totalVal).replace(/[\s,\u00A0]/g, ''), 10);
          if (!isNaN(cleanTotal)) {
            materials.push({ name: String(itemId).trim(), total: cleanTotal });
          }
        }
      }
    } else if (typeof rawItems === 'object') {
      for (const [k, v] of Object.entries(rawItems)) {
        const cleanTotal = parseInt(String(v).replace(/[\s,\u00A0]/g, ''), 10);
        if (!isNaN(cleanTotal)) {
          materials.push({ name: String(k).trim(), total: cleanTotal });
        }
      }
    }

    return materials;
  }

  // ==========================================================================
  // Конвертация материалов и группировка
  // ==========================================================================
  function processMaterials(materials) {
    if (!sortingDb) {
      log('База данных сортировки ещё не готова!', 'error');
      return [];
    }

    let group1Count = 0;
    let group2Count = 0;
    let group3Count = 0;

    const rows = materials.map(item => {
      const info = sortingDb.getItemInfo(item.name);
      const localId = info.localId;
      const stackSize = info.stackSize;
      const displayName = info.ruName || item.name;
      const total = item.total;

      const stacks = Math.ceil(total / stackSize);
      const chests = Math.ceil(stacks / DOUBLE_CHEST_SLOTS);
      const shulkerboxes = Math.ceil(stacks / SHULKER_BOX_SLOTS);

      // Группы сортировки:
      // 1. Более 27 стаков (stacks > 27)
      // 2. От 1 до 27 стаков (total >= stackSize и stacks <= 27)
      // 3. Меньше 1 стака (total < stackSize)
      let groupPriority;
      if (stacks > 27) {
        groupPriority = 1;
        group1Count++;
      } else if (total >= stackSize && stacks <= 27) {
        groupPriority = 2;
        group2Count++;
      } else {
        groupPriority = 3;
        group3Count++;
      }

      return {
        Local_ID: localId,
        Item: displayName,
        Total: total,
        'Стаки': stacks,
        'Даблчесты': chests,
        'Шалкербоксы': shulkerboxes,
        _group: groupPriority
      };
    });

    // Сортировка: сначала по группе (1 -> 2 -> 3),
    // затем внутри каждой группы строго по Local_ID,
    // при равенстве — по убыванию Total
    rows.sort((a, b) => {
      if (a._group !== b._group) return a._group - b._group;
      if (a.Local_ID !== b.Local_ID) return a.Local_ID - b.Local_ID;
      return b.Total - a.Total;
    });

    log(`Распределение по группам:`, 'info');
    log(`  • Более 27 стаков: ${group1Count} поз.`, 'info');
    log(`  • От 1 до 27 стаков: ${group2Count} поз.`, 'info');
    log(`  • Меньше 1 стака: ${group3Count} поз.`, 'info');

    groupBadgeBig.textContent = `Более 27 стаков: ${group1Count}`;
    groupBadgeMid.textContent = `От 1 до 27 стаков: ${group2Count}`;
    groupBadgeMin.textContent = `Меньше 1 стака: ${group3Count}`;

    return rows;
  }

  // ==========================================================================
  // Генерация Excel (.xlsx) через SheetJS
  // ==========================================================================
  function generateExcelWorkbook(rows) {
    if (!window.XLSX) {
      log('Библиотека SheetJS (XLSX) не загружена!', 'error');
      return null;
    }

    const cleanRows = rows.map(({ _group, ...rest }) => rest);

    const worksheet = XLSX.utils.json_to_sheet(cleanRows, {
      header: ['Local_ID', 'Item', 'Total', 'Стаки', 'Даблчесты', 'Шалкербоксы']
    });

    // Автоподбор ширины столбцов
    const headers = ['Local_ID', 'Item', 'Total', 'Стаки', 'Даблчесты', 'Шалкербоксы'];
    const colWidths = headers.map(header => {
      let maxLen = header.length;
      for (let i = 0; i < cleanRows.length; i++) {
        const val = cleanRows[i][header];
        const str = String(val !== undefined && val !== null ? val : '');
        if (str.length > maxLen) maxLen = str.length;
      }
      return { wch: Math.max(maxLen + 3, 10) };
    });
    worksheet['!cols'] = colWidths;

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    return workbook;
  }

  // ==========================================================================
  // Отображение таблицы в секции предпросмотра
  // ==========================================================================
  function renderPreviewTable(rows) {
    tableBody.innerHTML = '';

    if (!rows || rows.length === 0) {
      tablePlaceholder.style.display = 'flex';
      dataTable.style.display = 'none';
      return;
    }

    tablePlaceholder.style.display = 'none';
    dataTable.style.display = 'table';

    // Рендерим строки
    const fragment = document.createDocumentFragment();
    rows.forEach(row => {
      const tr = document.createElement('tr');
      tr.className = `row-group-${row._group}`;

      tr.innerHTML = `
        <td>${row.Local_ID}</td>
        <td><strong>${escapeHtml(row.Item)}</strong></td>
        <td>${row.Total.toLocaleString()}</td>
        <td>${row['Стаки']}</td>
        <td>${row['Даблчесты']}</td>
        <td>${row['Шалкербоксы']}</td>
      `;
      fragment.appendChild(tr);
    });

    tableBody.appendChild(fragment);
  }

  function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  // ==========================================================================
  // Обработка выбранного файла
  // ==========================================================================
  function handleFile(file) {
    if (!file) return;

    currentFile = file;
    fileNameDisplay.textContent = file.name;
    fileSizeDisplay.textContent = `${(file.size / 1024).toFixed(1)} КБ`;
    fileInfo.classList.add('active');

    outputFileName = file.name.replace(/\.[^/.]+$/, '') + '.xlsx';
    downloadFilename.textContent = outputFileName;

    log(`Выбран файл: "${file.name}" (${file.size} байт)`, 'info');

    const reader = new FileReader();
    reader.onload = function (e) {
      try {
        const text = e.target.result;
        log(`Чтение файла завершено. Анализ содержимого...`, 'info');

        const isJson = file.name.toLowerCase().endsWith('.json');
        const materials = isJson ? parseMaterialListJson(text) : parseMaterialListText(text);

        if (materials.length === 0) {
          log(`В файле "${file.name}" не найдено материалов из Litematica! Проверьте формат.`, 'error');
          downloadBtn.disabled = true;
          renderPreviewTable([]);
          return;
        }

        log(`Успешно извлечено позиций: ${materials.length}`, 'success');

        convertedRows = processMaterials(materials);

        // Расчёт суммарной статистики
        let totalChests = 0;
        let totalShulkers = 0;
        convertedRows.forEach(r => {
          totalChests += r['Даблчесты'];
          totalShulkers += r['Шалкербоксы'];
        });

        statRows.textContent = convertedRows.length;
        statChests.textContent = totalChests;
        statShulkers.textContent = totalShulkers;

        // Генерация Excel
        generatedWorkbook = generateExcelWorkbook(convertedRows);
        if (generatedWorkbook) {
          downloadBtn.disabled = false;
          log(`Таблица Excel успешно сгенерирована и готова к скачиванию!`, 'success');
        }

        // Рендеринг таблицы внизу
        renderPreviewTable(convertedRows);

      } catch (err) {
        log(`Критическая ошибка обработки файла: ${err.message}`, 'error');
        console.error(err);
      }
    };

    reader.onerror = function () {
      log(`Ошибка чтения файла "${file.name}".`, 'error');
    };

    reader.readAsText(file, 'utf-8');
  }

  // ==========================================================================
  // Слушатели событий интерфейса
  // ==========================================================================

  // Кнопка "Открыть"
  openFileBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    fileInput.click();
  });

  dropzone.addEventListener('click', () => {
    fileInput.click();
  });

  fileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files[0]) {
      handleFile(e.target.files[0]);
    }
  });

  // Drag and Drop
  ['dragenter', 'dragover'].forEach(eventName => {
    dropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.add('dragover');
    });
  });

  ['dragleave', 'drop'].forEach(eventName => {
    dropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.remove('dragover');
    });
  });

  dropzone.addEventListener('drop', (e) => {
    const dt = e.dataTransfer;
    if (dt && dt.files && dt.files[0]) {
      handleFile(dt.files[0]);
    }
  });

  // Кнопка скачивания готового Excel
  downloadBtn.addEventListener('click', () => {
    if (!generatedWorkbook || !outputFileName) return;
    try {
      XLSX.writeFile(generatedWorkbook, outputFileName);
      log(`Файл "${outputFileName}" успешно сохранён на устройство.`, 'success');
    } catch (e) {
      log(`Ошибка при сохранении файла: ${e.message}`, 'error');
    }
  });

  // Управление логами
  clearLogsBtn.addEventListener('click', () => {
    consoleBody.innerHTML = '';
    log('Логи очищены.', 'info');
  });

  copyLogsBtn.addEventListener('click', () => {
    const text = Array.from(consoleBody.children).map(c => c.textContent).join('\n');
    navigator.clipboard.writeText(text).then(() => {
      log('Логи скопированы в буфер обмена.', 'success');
    }).catch(() => {
      log('Не удалось скопировать логи в буфер обмена.', 'warn');
    });
  });

  // Запуск
  window.addEventListener('DOMContentLoaded', () => {
    initDatabase();
  });

})();
