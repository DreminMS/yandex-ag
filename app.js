// Yandex Task Calendar - PWA Application Logic (Multi-Sheet Matrix Parser)
(function() {
  'use strict';

  // --- Constants & Config ---
  const RU_MONTHS = ['январ', 'феврал', 'март', 'апрел', 'май', 'июн', 'июл', 'август', 'сентябр', 'октябр', 'ноябр', 'декабр'];

  // --- State ---
  let allTasks = [];
  let activeAssignees = ['Миша', 'Таня', 'Илья', 'Дима'];
  try {
    const saved = localStorage.getItem('active_assignees');
    if (saved) {
      const parsed = JSON.parse(saved);
      if (Array.isArray(parsed) && parsed.length > 0) activeAssignees = parsed;
    }
  } catch (e) {}

  let selectedAssignee = localStorage.getItem('selectedAssignee') || 'Все';
  let hideCompleted = localStorage.getItem('hideCompleted') === 'true';
  let sortOrder = localStorage.getItem('agendaSortOrder') || 'dateAsc';
  let sourceUrl = (localStorage.getItem('sourceUrl') || '').trim();
  let activeTab = 'calendar';
  let selectedDate = new Date();
  let currentMonth = new Date(selectedDate.getFullYear(), selectedDate.getMonth(), 1);
  let searchText = '';
  let collapsedSections = new Set(['completed_past']);
  let activeModalTask = null;

  // --- Date Formatter Helpers ---
  function toDateString(d) {
    if (!d || !(d instanceof Date) || isNaN(d)) return '';
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }

  function getTodayString() {
    return toDateString(new Date());
  }

  function parseExcelOrDate(val) {
    if (val === undefined || val === null || val === '') return null;
    if (val instanceof Date && !isNaN(val)) {
      return toDateString(val);
    }

    const num = parseFloat(val);
    if (!isNaN(num) && num >= 40000 && num <= 55000) {
      // Excel epoch 1899-12-30: 25569 days between 1899-12-30 and 1970-01-01
      const utcDays = Math.floor(num - 25569);
      const d = new Date(utcDays * 86400 * 1000);
      const y = d.getUTCFullYear();
      const m = String(d.getUTCMonth() + 1).padStart(2, '0');
      const day = String(d.getUTCDate()).padStart(2, '0');
      return `${y}-${m}-${day}`;
    }

    const str = String(val).trim();
    const m = str.match(/(\d{1,2})[./](\d{1,2})[./](\d{2,4})/);
    if (m) {
      const day = String(parseInt(m[1], 10)).padStart(2, '0');
      const month = String(parseInt(m[2], 10)).padStart(2, '0');
      let year = parseInt(m[3], 10);
      if (year < 100) year += 2000;
      return `${year}-${month}-${day}`;
    }

    return null;
  }

  // --- Status Parser (matching Android TaskStatus.kt) ---
  function parseStatus(raw) {
    if (!raw) return { name: 'Запланировано', color: '#007aff', isCompleted: false, priority: 3 };
    const lower = String(raw).toLowerCase().trim();

    if (lower.includes('сдан') || lower.includes('готов') || lower.includes('завершен') || lower.includes('выполнен') || lower === '+') {
      return { name: 'Сдано', color: '#34c759', isCompleted: true, priority: 5 };
    }
    if (lower.includes('отмен') || lower.includes('перенос') || lower.includes('просроч')) {
      return { name: 'Отменено', color: '#8e8e93', isCompleted: true, priority: 6 };
    }
    if (lower.includes('в работе') || lower.includes('процесс') || lower.includes('монтаж') || lower.includes('съемк')) {
      return { name: 'В работе', color: '#ff9500', isCompleted: false, priority: 1 };
    }
    if (lower.includes('отгул') || lower.includes('отпуск')) {
      return { name: 'Отгул', color: '#8e8e93', isCompleted: false, priority: 4 };
    }
    if (lower.includes('деж')) {
      return { name: 'Дежурство', color: '#af52de', isCompleted: false, priority: 2 };
    }
    return { name: 'Запланировано', color: '#007aff', isCompleted: false, priority: 3 };
  }

  // --- Clean Assignee Name ---
  function cleanAssigneeName(raw) {
    if (!raw) return '';
    const firstLine = String(raw).split(/[\r\n]+/)[0].trim();
    const low = firstLine.toLowerCase();

    if (['=---', 'обозначен', 'столбец', 'шаблон', 'планирование', 'дата', 'фрилансер', 'подрядчик', 'видеограф 2', '9999'].some(k => low.includes(k))) {
      return '';
    }

    if (low.includes('миш') || low.includes('михаил')) return 'Миша';
    if (low.includes('тан') || low.includes('татьян')) return 'Таня';
    if (low.includes('иль')) return 'Илья';
    if (low.includes('владимир')) return 'Владимир';
    if (low.includes('дима') || low.includes('дмит') || low.startsWith('дим') || low.includes(' дим')) return 'Дима';
    if (low.includes('наст')) return 'Настя';
    if (low.includes('саш') || low.includes('александр')) return 'Саша';
    if (low.includes('маш') || low.includes('мари')) return 'Маша';
    if (low.includes('серг')) return 'Сергей';
    if (low.includes('артем') || low.includes('артём')) return 'Артём';
    if (low.includes('алексей')) return 'Алексей';

    let cleaned = firstLine.replace(/^(?:фотограф|видеограф|оператор|монтажер|монтажёр)\s+/i, '');
    cleaned = cleaned.replace(/(\+?\d[\d\s-]{8,})/g, '').trim();
    const parts = cleaned.split(/\s+/);
    if (parts.length > 0 && parts[0].length >= 2) {
      return parts[0];
    }
    return cleaned || firstLine;
  }

  // --- Duty Filter: only show duties for the nearest month ---
  function isDutyTask(task) {
    if (!task) return false;
    const stName = (task.status && task.status.name) ? task.status.name : '';
    if (stName === 'Дежурство') return true;
    const title = (task.title || '').toLowerCase();
    const raw = (task.rawText || '').toLowerCase();
    return title.includes('деж') || raw.includes('деж');
  }

  function isDutyInNearestMonth(taskDateStr) {
    if (!taskDateStr) return false;
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const todayStr = toDateString(today);

    // End of next month (e.g. in Oct -> through end of Nov)
    const maxDate = new Date(today.getFullYear(), today.getMonth() + 2, 0);
    const maxDateStr = toDateString(maxDate);

    return taskDateStr >= todayStr && taskDateStr <= maxDateStr;
  }

  function filterTasks(tasks) {
    if (!Array.isArray(tasks)) return [];
    return tasks.filter(t => {
      if (!isDutyTask(t)) return true;
      return isDutyInNearestMonth(t.date);
    });
  }

  // --- Detail Extractor (matching Android TaskDetailExtractor.kt) ---
  function extractDetails(rawText, sheetName = '', assignee = '') {
    const text = String(rawText || '').trim();
    if (!text) {
      return { cleanTitle: 'Без названия', timeRange: null, location: null, tzUrl: null, bitrixUrl: null, deadline: null, formattedDescription: '' };
    }

    // 1. URLs
    const urlRegex = /https?:\/\/[^\s)\]},;]+/gi;
    const urls = text.match(urlRegex) || [];
    let tzUrl = null;
    let bitrixUrl = null;

    for (const u of urls) {
      const low = u.toLowerCase();
      if (low.includes('bitrix')) {
        bitrixUrl = u;
      } else if (low.includes('disk.yandex') || low.includes('yadi.sk') || low.includes('docs.google') || low.includes('drive.google')) {
        tzUrl = u;
      } else if (!tzUrl) {
        tzUrl = u;
      }
    }

    // 2. Location
    let location = null;
    const locMatch = text.match(/(?:локаци[яи]|адрес|место)\s*:\s*([^\n\r]+)/i);
    if (locMatch) {
      location = locMatch[1].trim();
    } else if (text.toLowerCase().includes('тз:')) {
      const tzLocMatch = text.match(/тз\s*:\s*([^:\n\r]+?)(?=\s*тз|\s*https?:|$)/i);
      if (tzLocMatch && !tzLocMatch[1].toLowerCase().startsWith('http') && tzLocMatch[1].trim().length > 5) {
        location = tzLocMatch[1].trim();
      }
    }

    if (location) {
      location = location.replace(/(?:тз\s*:|https?:\/\/[^\s]+).*/i, '').replace(/^[3.]+\s*/, '').trim().replace(/^[-–—,.\s]+|[-–—,.\s]+$/g, '');
      if (location.length === 0) location = null;
    }

    // 3. Time Range
    let timeRange = null;
    const timeMatch = text.match(/\b(\d{1,2}[:.]\d{2}\s*[-–—]\s*(?:\d{1,2}[:.]\d{2}|по ситуации))\b/i);
    if (timeMatch) {
      timeRange = timeMatch[1].replace(/\./g, ':').trim();
    } else {
      const prelimMatch = text.match(/\(([^)]*половина[^)]*)\)/i);
      if (prelimMatch) {
        timeRange = prelimMatch[0].trim();
      }
    }

    // 4. Deadline
    let deadline = null;
    const dlMatch = text.match(/((?:сдать(?:\s+фото)?\s+до|сдача|до\s+\d{1,2}[:.]\d{2}|обработка\s+до)[^\n\r,.]+)/i);
    if (dlMatch) {
      deadline = dlMatch[1].trim();
    }

    // 5. Clean Title
    const lines = text.split(/[\r\n]+/).map(l => l.trim()).filter(l => l.length > 0);
    const firstLine = lines[0] || '';
    let cleanTitle = firstLine;

    if (timeRange && cleanTitle.includes(timeRange)) {
      cleanTitle = cleanTitle.replace(timeRange, '');
    }
    cleanTitle = cleanTitle.replace(/\b\d{1,2}[:.]\d{2}\s*[-–—]\s*(?:\d{1,2}[:.]\d{2}|по ситуации)\b/gi, '');
    cleanTitle = cleanTitle.replace(/\([^)]*половина[^)]*\)/gi, '');
    cleanTitle = cleanTitle.replace(/https?:\/\/[^\s]+/gi, '');
    cleanTitle = cleanTitle.replace(/(?:локаци[яи]|адрес|место|тз)\s*:.*/gi, '');
    cleanTitle = cleanTitle.trim().replace(/^[-–—:,. \s]+|[-–—:,. \s]+$/g, '');

    if (!cleanTitle && lines.length > 1) {
      cleanTitle = lines[1].replace(/https?:\/\/[^\s]+/gi, '').trim().replace(/^[-–—:,. \s]+|[-–—:,. \s]+$/g, '');
    }
    if (!cleanTitle) {
      cleanTitle = firstLine.slice(0, 60).trim();
    }

    // 6. Extra Details Snippet (details from TZ and cell lines)
    let detailsSnippet = '';
    const detailLines = [];
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      if (line === firstLine) continue;
      const withoutUrls = line.replace(/https?:\/\/[^\s]+/gi, '')
                              .replace(/^(?:тз|ссылка|документ|файл|титры|описание)\s*:?\s*/gi, '')
                              .trim()
                              .replace(/^[-–—:,. \s]+|[-–—:,. \s]+$/g, '');
      if (location && line.includes(location)) continue;
      if (withoutUrls.length > 2) {
        detailLines.push(withoutUrls);
      }
    }
    detailsSnippet = detailLines.join(' • ');

    // 7. Formatted Description
    let descParts = [];
    if (location) descParts.push(`📍 Локация: ${location}`);
    if (timeRange) descParts.push(`⏰ Время: ${timeRange}`);
    if (deadline) descParts.push(`⏳ Срок сдачи: ${deadline}`);
    if (tzUrl) descParts.push(`🔗 ТЗ: ${tzUrl}`);
    if (bitrixUrl) descParts.push(`📌 Битрикс24: ${bitrixUrl}`);
    if (assignee) descParts.push(`👤 Ответственный: ${assignee}`);
    if (sheetName) descParts.push(`📄 Лист: ${sheetName}`);
    descParts.push(`\n📝 Подробности ТЗ:\n${text}`);

    return {
      cleanTitle,
      timeRange,
      location,
      tzUrl,
      bitrixUrl,
      deadline,
      detailsSnippet,
      formattedDescription: descParts.join('\n')
    };
  }

  // --- Helper: Domain for TZ links ---
  function getDomain(url) {
    if (!url) return '';
    try {
      const u = new URL(url);
      const host = u.hostname.toLowerCase();
      if (host.includes('yandex') || host.includes('yadi.sk')) return 'Яндекс.Диск';
      if (host.includes('google')) return 'Google Docs';
      if (host.includes('bitrix')) return 'Битрикс24';
      return host.replace(/^www\./, '');
    } catch (e) {
      return 'Файл';
    }
  }

  // --- Helper: Make URLs clickable in text ---
  function linkify(text) {
    if (!text) return '';
    const escaped = escapeHtml(text);
    return escaped.replace(/(https?:\/\/[^\s<]+)/g, (url) => {
      let cleanUrl = url.replace(/[.,;)]+$/, '');
      return `<a href="${cleanUrl}" target="_blank" rel="noopener noreferrer" class="tz-inline-link">${cleanUrl} ↗</a>`;
    });
  }

  // --- Assignee Matching ---
  function matchesAssignee(task, filter) {
    if (!filter || filter === 'Все') return true;
    const candidate = (task.assignee || '').toLowerCase();
    const f = filter.toLowerCase();
    if (f === 'миша') return candidate.includes('миш') || candidate.includes('михаил');
    if (f === 'таня') return candidate.includes('тан') || candidate.includes('татьян');
    if (f === 'илья') return candidate.includes('иль');
    if (f === 'дима') return candidate.includes('дим') || candidate.includes('дмит');
    if (f === 'владимир') return candidate.includes('владимир');
    return candidate.includes(f) || f.includes(candidate);
  }

  // --- Dynamic Assignee Extraction & Filter Bar ---
  function getBadgeId(name) {
    if (name === 'Все') return 'badge-all';
    if (name === 'Миша') return 'badge-misha';
    if (name === 'Таня') return 'badge-tanya';
    if (name === 'Илья') return 'badge-ilya';
    if (name === 'Дима') return 'badge-dima';
    return 'badge-' + encodeURIComponent(name).replace(/%/g, '_');
  }

  function renderAssigneeFilters() {
    const filterBar = document.getElementById('filterBar');
    const desktopList = document.getElementById('desktopAssigneesList');

    if (selectedAssignee !== 'Все' && !activeAssignees.includes(selectedAssignee)) {
      selectedAssignee = 'Все';
      localStorage.setItem('selectedAssignee', 'Все');
    }

    const items = ['Все', ...activeAssignees];

    // 1. Mobile Filter Bar
    if (filterBar) {
      filterBar.innerHTML = items.map(name => {
        const isActive = name === selectedAssignee ? 'active' : '';
        const badgeId = getBadgeId(name);
        return `<button class="filter-pill ${isActive}" data-assignee="${escapeHtml(name)}">${escapeHtml(name)} <span class="pill-badge" id="${badgeId}">0</span></button>`;
      }).join('');

      filterBar.querySelectorAll('.filter-pill').forEach(pill => {
        pill.addEventListener('click', () => {
          selectAssignee(pill.dataset.assignee);
        });
      });
    }

    // 2. Desktop Sidebar List
    if (desktopList) {
      desktopList.innerHTML = items.map(name => {
        const isActive = name === selectedAssignee ? 'active' : '';
        const badgeId = getBadgeId(name);
        const dotColor = name === 'Все' ? 'var(--ios-blue)' :
                         name === 'Миша' ? '#007aff' :
                         name === 'Таня' ? '#34c759' :
                         name === 'Илья' ? '#ff9500' : '#af52de';
        return `
          <button class="sidebar-assignee-btn ${isActive}" data-assignee="${escapeHtml(name)}">
            <div class="sidebar-assignee-left">
              <span class="sidebar-assignee-dot" style="background:${dotColor}"></span>
              <span>${escapeHtml(name)}</span>
            </div>
            <span class="sidebar-assignee-badge" id="d-${badgeId}">0</span>
          </button>
        `;
      }).join('');

      desktopList.querySelectorAll('.sidebar-assignee-btn').forEach(btn => {
        btn.addEventListener('click', () => {
          selectAssignee(btn.dataset.assignee);
        });
      });
    }

    updateBadges();
  }

  function selectAssignee(name) {
    selectedAssignee = name;
    localStorage.setItem('selectedAssignee', selectedAssignee);

    // Sync mobile filter pills
    document.querySelectorAll('.filter-pill').forEach(p => {
      p.classList.toggle('active', p.dataset.assignee === selectedAssignee);
    });

    // Sync desktop sidebar buttons
    document.querySelectorAll('.sidebar-assignee-btn').forEach(b => {
      b.classList.toggle('active', b.dataset.assignee === selectedAssignee);
    });

    renderAll();
  }

  function extractActiveAssignees(workbook) {
    const today = new Date();
    const currYear = String(today.getFullYear());
    const relevantMonthRoots = [];
    for (let m = today.getMonth(); m < 12; m++) {
      relevantMonthRoots.push(RU_MONTHS[m]);
    }

    let activeSheets = workbook.SheetNames.filter(name => 
      name.includes(currYear) && relevantMonthRoots.some(m => name.toLowerCase().includes(m))
    );
    if (activeSheets.length === 0) {
      activeSheets = workbook.SheetNames.filter(name => name.includes(currYear));
    }
    if (activeSheets.length === 0) {
      activeSheets = workbook.SheetNames.slice(-3);
    }

    const found = [];
    for (const sheetName of activeSheets) {
      const sheet = workbook.Sheets[sheetName];
      if (!sheet) continue;
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: '' });
      if (!rows || rows.length === 0) continue;
      const row0 = rows[0] || [];
      for (let c = 0; c < row0.length; c++) {
        const text = String(row0[c] || '').trim();
        if (text) {
          const clean = cleanAssigneeName(text);
          if (clean && !found.includes(clean)) {
            found.push(clean);
          }
        }
      }
    }

    return found.length > 0 ? found : ['Миша', 'Таня', 'Илья', 'Дима'];
  }

  // --- Multi-Sheet XLSX Parser ---
  function parseWorkbook(workbook) {
    const detectedAssignees = extractActiveAssignees(workbook);
    if (detectedAssignees && detectedAssignees.length > 0) {
      activeAssignees = detectedAssignees;
      try {
        localStorage.setItem('active_assignees', JSON.stringify(activeAssignees));
      } catch (e) {}
      renderAssigneeFilters();
    }

    const currentYear = new Date().getFullYear();
    const targetYears = [(currentYear - 1).toString(), currentYear.toString(), (currentYear + 1).toString()]; // ["2025", "2026", "2027"]

    // Select sheets containing current, previous, or next year
    let targetSheets = workbook.SheetNames.filter(name => 
      targetYears.some(y => name.includes(y))
    );
    if (targetSheets.length === 0) {
      targetSheets = workbook.SheetNames;
    }

    const tasks = [];
    let idCounter = 1;

    for (const sheetName of targetSheets) {
      const sheet = workbook.Sheets[sheetName];
      if (!sheet) continue;

      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: '' });
      if (!rows || rows.length < 2) continue;

      // 1. Read Row 1 to find assignees by column (only ACTIVE assignees)
      const assigneesByCol = {};
      const row0 = rows[0] || [];
      for (let c = 0; c < row0.length; c++) {
        const cellText = String(row0[c] || '').trim();
        if (cellText && !cellText.toLowerCase().startsWith('обозначен')) {
          const clean = cleanAssigneeName(cellText);
          if (clean && activeAssignees.includes(clean)) {
            assigneesByCol[c] = clean;
          }
        }
      }

      let currentDateStr = null;

      // 2. Iterate remaining rows
      for (let r = 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row || row.length === 0) continue;

        // Check if row sets a date
        for (let c = 0; c < row.length; c++) {
          const val = row[c];
          if (val === undefined || val === null || val === '') continue;
          const dStr = parseExcelOrDate(val);
          if (dStr) {
            currentDateStr = dStr;
            break;
          }
        }

        // If date is active and assignees exist, extract tasks
        if (currentDateStr && Object.keys(assigneesByCol).length > 0) {
          for (const colStr in assigneesByCol) {
            const colIdx = parseInt(colStr, 10);
            const val = row[colIdx];
            if (val === undefined || val === null) continue;
            const text = String(val).trim();
            const cleanDigits = text.replace(/[\s.-]/g, '');

            if (text.length > 1 && !/^\d+$/.test(cleanDigits) && !text.toLowerCase().startsWith('обозначен')) {
              const assignee = assigneesByCol[colIdx];
              const details = extractDetails(text, sheetName, assignee);
              const status = parseStatus(text);

              if (isDutyTask({ status, title: details.cleanTitle, rawText: text }) && !isDutyInNearestMonth(currentDateStr)) {
                continue;
              }

              tasks.push({
                id: `task_${idCounter++}`,
                title: details.cleanTitle,
                date: currentDateStr,
                timeRange: details.timeRange,
                assignee: assignee,
                status: status,
                location: details.location,
                tzUrl: details.tzUrl,
                bitrixUrl: details.bitrixUrl,
                deadline: details.deadline,
                detailsSnippet: details.detailsSnippet,
                description: details.formattedDescription,
                rawText: text,
                sheetName: sheetName,
                sourceRow: r + 1
              });
            }
          }
        }
      }
    }

    return tasks;
  }

  // --- Online Cloud Sync (events.json generated by GitHub Actions) ---
  async function loadOnlineEvents() {
    try {
      const resp = await fetch('events.json?v=' + Date.now(), { cache: 'no-store' });
      if (!resp.ok) return false;
      const data = await resp.json();
      if (!data || !data.all_events || data.all_events.length === 0) return false;

      let idCounter = 1;
      const tasks = data.all_events.map(ev => {
        const cleanName = cleanAssigneeName(ev.person_name || ev.person_key);
        const status = parseStatus(ev.full_text);
        const timeRange = ev.start_time ? (ev.end_time ? `${ev.start_time} - ${ev.end_time}` : ev.start_time) : '';
        const tzUrl = (ev.urls && ev.urls.length > 0) ? ev.urls[0] : '';
        const bitrixUrl = (ev.urls && ev.urls.find(u => u.includes('bitrix'))) || '';

        return {
          id: ev.uid || `task_${idCounter++}`,
          title: ev.title || 'Событие',
          date: ev.date,
          timeRange: timeRange,
          assignee: cleanName,
          status: status,
          location: ev.location || '',
          tzUrl: tzUrl,
          bitrixUrl: bitrixUrl,
          deadline: '',
          detailsSnippet: '',
          description: ev.description || ev.full_text,
          rawText: ev.full_text,
          sheetName: ev.sheet || '',
          sourceRow: ev.row || ''
        };
      });

      allTasks = filterTasks(tasks);
      localStorage.setItem('cached_tasks', JSON.stringify(allTasks));
      if (data.updated_at) {
        localStorage.setItem('last_sync_date', data.updated_at);
      } else {
        localStorage.setItem('last_sync_date', new Date().toISOString());
      }
      return true;
    } catch (e) {
      console.warn('Не удалось загрузить онлайн-расписание (events.json):', e);
      return false;
    }
  }

  // --- Network Sync (Direct Yandex Disk Cloud API + Online Fallback) ---
  async function syncTasks(force = false) {
    const syncBtn = document.getElementById('syncBtn');
    const manualBtn = document.getElementById('manualSyncBtn');
    syncBtn.classList.add('spinning');
    manualBtn.textContent = 'Обновление данных...';

    try {
      // 1. First try instant online events.json
      const onlineOk = await loadOnlineEvents();
      if (onlineOk) {
        renderAll();
        if (!sourceUrl || !force) {
          syncBtn.classList.remove('spinning');
          manualBtn.textContent = 'Синхронизировать сейчас';
          return;
        }
      }

      if (!sourceUrl) {
        if (!onlineOk && allTasks.length === 0) {
          alert('🔒 Расписание не настроено.\nДля загрузки задач укажите секретную ссылку на таблицу в Настройках.');
          const settingsTabBtn = document.querySelector('[data-tab="settings"]');
          if (settingsTabBtn) settingsTabBtn.click();
        }
        return;
      }

      // 2. Download from Yandex Disk if sourceUrl is configured
      manualBtn.textContent = 'Подключение к Яндекс.Диску...';
      const apiUrl = `https://cloud-api.yandex.net/v1/disk/public/resources/download?public_key=${encodeURIComponent(sourceUrl)}`;
      const apiResp = await fetch(apiUrl, { referrerPolicy: 'no-referrer' });
      if (!apiResp.ok) {
        throw new Error(`Ошибка Яндекс.Диска (${apiResp.status}). Проверьте правильность ссылки.`);
      }
      const apiData = await apiResp.json();
      if (!apiData.href) {
        throw new Error('Яндекс.Диск не вернул ссылку для скачивания файла.');
      }

      manualBtn.textContent = 'Скачивание таблицы...';
      const fileResp = await fetch(apiData.href, { referrerPolicy: 'no-referrer' });
      if (!fileResp.ok) {
        throw new Error(`Ошибка скачивания файла таблицы (${fileResp.status})`);
      }
      const buffer = await fileResp.arrayBuffer();
      if (!buffer || buffer.byteLength === 0) {
        throw new Error('Файл таблицы пуст');
      }

      manualBtn.textContent = 'Обработка данных...';
      const workbook = XLSX.read(buffer, { type: 'array' });
      let tasks = parseWorkbook(workbook);

      if (!tasks || tasks.length === 0) {
        throw new Error('Задачи в таблице не найдены. Убедитесь, что ссылка ведёт на правильный файл расписания.');
      }

      tasks = filterTasks(tasks);
      allTasks = tasks;
      localStorage.setItem('cached_tasks', JSON.stringify(tasks));
      const now = new Date();
      localStorage.setItem('last_sync_date', now.toISOString());

      renderAll();
      alert(`Синхронизация завершена!\nЗагружено задач: ${tasks.length}`);
    } catch (err) {
      console.error(err);
      if (allTasks.length === 0) {
        alert('Ошибка при синхронизации:\n' + err.message);
      }
    } finally {
      syncBtn.classList.remove('spinning');
      manualBtn.textContent = 'Синхронизировать сейчас';
    }
  }

  // --- Render Functions ---

  function updateBadges() {
    const items = ['Все', ...activeAssignees];
    items.forEach(w => {
      const count = allTasks.filter(t => {
        if (!matchesAssignee(t, w)) return false;
        if (hideCompleted && t.status.isCompleted) return false;
        if (isDutyTask(t) && !isDutyInNearestMonth(t.date)) return false;
        return true;
      }).length;
      const badgeId = getBadgeId(w);
      const el = document.getElementById(badgeId);
      if (el) el.textContent = count;
      const del = document.getElementById('d-' + badgeId);
      if (del) del.textContent = count;
    });

    const totalActive = allTasks.filter(t => !isDutyTask(t) || isDutyInNearestMonth(t.date)).length;
    const totalEl = document.getElementById('totalTasksLabel');
    if (totalEl) totalEl.textContent = totalActive;
    const dTotalEl = document.getElementById('desktopTotalTasksLabel');
    if (dTotalEl) dTotalEl.textContent = totalActive;

    const lastSync = localStorage.getItem('last_sync_date');
    let lastSyncFormatted = 'Никогда';
    if (lastSync) {
      const d = new Date(lastSync);
      lastSyncFormatted = d.toLocaleString('ru-RU', { dateStyle: 'short', timeStyle: 'short' });
    }
    const lastSyncEl = document.getElementById('lastSyncLabel');
    if (lastSyncEl) lastSyncEl.textContent = lastSyncFormatted;
    const dLastSyncEl = document.getElementById('desktopLastSyncLabel');
    if (dLastSyncEl) dLastSyncEl.textContent = lastSyncFormatted;
  }

  function getChipBg(color) {
    if (color === '#34c759') return 'rgba(52, 199, 89, 0.12)';
    if (color === '#ff9500') return 'rgba(255, 149, 0, 0.12)';
    if (color === '#af52de') return 'rgba(175, 82, 222, 0.12)';
    if (color === '#8e8e93') return 'rgba(142, 142, 147, 0.12)';
    return 'rgba(0, 122, 255, 0.12)';
  }

  function renderCalendar() {
    const grid = document.getElementById('calendarGrid');
    grid.innerHTML = '';

    const year = currentMonth.getFullYear();
    const month = currentMonth.getMonth();

    const monthNames = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"];
    const monthTitleText = `${monthNames[month]} ${year}`;
    const mt = document.getElementById('monthTitle');
    if (mt) mt.textContent = monthTitleText;
    const mtd = document.getElementById('monthTitleDesktop');
    if (mtd) mtd.textContent = monthTitleText;

    // First day of month (Monday = 0, Sunday = 6)
    const firstDay = new Date(year, month, 1);
    let startDayOfWeek = firstDay.getDay() - 1;
    if (startDayOfWeek === -1) startDayOfWeek = 6;

    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const daysInPrevMonth = new Date(year, month, 0).getDate();

    // Prev month padding
    for (let p = startDayOfWeek - 1; p >= 0; p--) {
      const d = daysInPrevMonth - p;
      const cellDate = new Date(year, month - 1, d);
      grid.appendChild(createDayCell(cellDate, d, true));
    }

    // Current month days
    for (let d = 1; d <= daysInMonth; d++) {
      const cellDate = new Date(year, month, d);
      grid.appendChild(createDayCell(cellDate, d, false));
    }

    // Next month padding to fill multiple of 7
    const totalCells = grid.children.length;
    const remaining = (7 - (totalCells % 7)) % 7;
    for (let n = 1; n <= remaining; n++) {
      const cellDate = new Date(year, month + 1, n);
      grid.appendChild(createDayCell(cellDate, n, true));
    }

    renderSelectedDayTasks();
  }

  function createDayCell(cellDate, dayNumber, isOtherMonth) {
    const div = document.createElement('div');
    div.className = 'cal-day';
    if (isOtherMonth) div.classList.add('other-month');

    const cellDateStr = toDateString(cellDate);
    const selectedDateStr = toDateString(selectedDate);
    const todayStr = getTodayString();

    if (cellDateStr === todayStr) div.classList.add('today');
    if (cellDateStr === selectedDateStr) div.classList.add('selected');

    const dayTasks = allTasks.filter(t => {
      if (t.date !== cellDateStr) return false;
      if (!matchesAssignee(t, selectedAssignee)) return false;
      if (hideCompleted && t.status.isCompleted) return false;
      if (isDutyTask(t) && !isDutyInNearestMonth(t.date)) return false;
      return true;
    }).sort((a, b) => (a.timeRange || '').localeCompare(b.timeRange || ''));

    const hasTasks = dayTasks.length > 0;

    div.innerHTML = `
      <div class="cal-day-top">
        <span class="cal-day-num">${dayNumber}</span>
        ${hasTasks ? '<span class="cal-dot"></span>' : ''}
        ${hasTasks ? `<span class="cal-day-count-badge">${dayTasks.length}</span>` : ''}
      </div>
      <div class="cal-day-events"></div>
    `;

    const eventsContainer = div.querySelector('.cal-day-events');
    if (eventsContainer && dayTasks.length > 0) {
      dayTasks.slice(0, 3).forEach(task => {
        const chip = document.createElement('div');
        chip.className = 'desktop-event-chip';
        if (activeModalTask && activeModalTask.id === task.id) {
          chip.classList.add('active-chip');
        }
        chip.style.borderLeftColor = task.status.color;
        chip.style.backgroundColor = getChipBg(task.status.color);
        chip.title = `${task.timeRange ? task.timeRange + ' ' : ''}${task.assignee}: ${task.title}`;
        chip.innerHTML = `
          ${task.timeRange ? `<span class="chip-time">${escapeHtml(task.timeRange.split(/[-–—]/)[0].trim())}</span>` : ''}
          <span class="chip-assignee">${escapeHtml(task.assignee)}</span>
          <span class="chip-title">${escapeHtml(task.title)}</span>
        `;
        chip.addEventListener('click', (e) => {
          e.stopPropagation();
          selectedDate = cellDate;
          if (activeModalTask && activeModalTask.id === task.id) {
            activeModalTask = null;
            renderCalendar();
            renderSelectedDayTasks();
          } else {
            activeModalTask = task;
            renderCalendar();
            openTaskModal(task);
          }
        });
        eventsContainer.appendChild(chip);
      });

      if (dayTasks.length > 3) {
        const more = document.createElement('div');
        more.className = 'desktop-more-chip';
        more.textContent = `+ ещё ${dayTasks.length - 3}`;
        more.addEventListener('click', (e) => {
          e.stopPropagation();
          selectedDate = cellDate;
          activeModalTask = null;
          renderCalendar();
          renderSelectedDayTasks();
        });
        eventsContainer.appendChild(more);
      }
    }

    div.addEventListener('click', () => {
      selectedDate = cellDate;
      activeModalTask = null;
      renderCalendar();
      renderSelectedDayTasks();
    });

    return div;
  }

  function renderDesktopInspectorDay(date, tasks) {
    const body = document.getElementById('inspectorBody');
    const headerTitle = document.getElementById('inspectorHeaderTitle');
    if (!body) return;

    if (!sourceUrl && allTasks.length === 0) {
      if (headerTitle) headerTitle.textContent = 'Доступ закрыт';
      body.innerHTML = `
        <div class="inspector-placeholder">
          <div style="font-size:36px;margin-bottom:10px">🔒</div>
          <div style="font-weight:600;font-size:15px;margin-bottom:6px">Расписание заблокировано</div>
          <div style="font-size:13px;line-height:1.4;margin-bottom:16px">Укажите секретную ссылку на таблицу в Настройках.</div>
          <button class="modal-btn primary" onclick="document.querySelector('[data-tab=settings]').click()">Перейти в Настройки</button>
        </div>
      `;
      return;
    }

    const dayStr = date.toLocaleDateString('ru-RU', { day: 'numeric', month: 'long', weekday: 'long' });
    if (headerTitle) headerTitle.textContent = `Съёмки: ${dayStr}`;

    if (!tasks || tasks.length === 0) {
      body.innerHTML = `
        <div class="inspector-placeholder">
          <div style="font-size:32px;margin-bottom:10px">☕️</div>
          <div style="font-weight:600;font-size:15px;margin-bottom:4px">Нет съёмок на этот день</div>
          <div style="font-size:13px;color:var(--ios-text-secondary)">На ${dayStr} съёмок не запланировано.</div>
        </div>
      `;
      return;
    }

    body.innerHTML = '';
    const intro = document.createElement('div');
    intro.style.fontSize = '12px';
    intro.style.color = 'var(--ios-text-secondary)';
    intro.style.fontWeight = '600';
    intro.style.marginBottom = '4px';
    intro.textContent = `${tasks.length} ${pluralize(tasks.length)} • Нажмите на съёмку для деталей:`;
    body.appendChild(intro);

    tasks.forEach(task => {
      const card = createTaskCard(task, false);
      body.appendChild(card);
    });
  }

  function renderDesktopInspectorTask(task) {
    const body = document.getElementById('inspectorBody');
    const headerTitle = document.getElementById('inspectorHeaderTitle');
    if (!body || !task) return;
    if (headerTitle) headerTitle.textContent = 'Детали задачи';

    const domainName = getDomain(task.tzUrl);
    let dtStr = 'Дата не указана';
    if (task.date) {
      const parts = task.date.split('-');
      if (parts.length === 3) {
        const d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
        dtStr = d.toLocaleDateString('ru-RU', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
      }
    }
    if (task.timeRange) dtStr += ` • ${task.timeRange}`;

    body.innerHTML = `
      <div style="display:flex;justify-content:space-between;align-items:center;gap:8px">
        <span class="status-badge" style="background:${task.status.color}20;color:${task.status.color}">${escapeHtml(task.status.name)}</span>
        <span class="assignee-chip" style="font-size:13px;padding:3px 10px;background:var(--ios-pill-bg);border-radius:12px">👤 ${escapeHtml(task.assignee)}</span>
      </div>

      <h2 style="font-size:18px;font-weight:700;line-height:1.3;margin:4px 0">${escapeHtml(task.title)}</h2>

      <div class="modal-detail-card" style="margin:0">
        <div style="font-size:12px;color:var(--ios-text-secondary);margin-bottom:2px">Дата и время</div>
        <div style="font-size:15px;font-weight:600">${escapeHtml(dtStr)}</div>
        ${task.deadline ? `<div style="font-size:13px;color:var(--ios-red);margin-top:4px">⏳ ${escapeHtml(task.deadline)}</div>` : ''}
      </div>

      ${task.location ? `
      <div class="modal-detail-card" style="margin:0">
        <div style="font-size:12px;color:var(--ios-text-secondary);margin-bottom:4px">Локация / Адрес</div>
        <div style="font-size:14px;margin-bottom:8px">${escapeHtml(task.location)}</div>
        <div class="modal-btn-row">
          <a class="modal-btn maps-apple" href="https://maps.apple.com/?q=${encodeURIComponent(task.location)}" target="_blank">Apple Карты</a>
          <a class="modal-btn maps-yandex" href="https://yandex.ru/maps/?text=${encodeURIComponent(task.location)}" target="_blank">Яндекс Карты</a>
        </div>
      </div>
      ` : ''}

      ${task.tzUrl ? `
      <div class="modal-detail-card tz-card-highlight" style="margin:0">
        <div style="font-size:12px;font-weight:700;color:var(--ios-blue);margin-bottom:6px;display:flex;align-items:center;gap:6px">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>
          <span>Файл ТЗ (${escapeHtml(domainName)})</span>
        </div>
        <div style="margin-bottom:10px;word-break:break-all">
          <a class="tz-active-url" href="${escapeHtml(task.tzUrl)}" target="_blank" rel="noopener noreferrer">${escapeHtml(task.tzUrl)}</a>
        </div>
        <div class="modal-btn-row">
          <a class="modal-btn primary" href="${escapeHtml(task.tzUrl)}" target="_blank" rel="noopener noreferrer" style="gap:6px">
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
            Открыть файл ТЗ ↗
          </a>
          <button class="modal-btn secondary" id="desktopCopyTzBtn">Копировать ссылку</button>
        </div>
      </div>
      ` : ''}

      ${task.bitrixUrl ? `
      <div class="modal-detail-card" style="margin:0">
        <div style="font-size:12px;color:var(--ios-text-secondary);margin-bottom:4px">Битрикс24</div>
        <a class="modal-btn secondary" href="${escapeHtml(task.bitrixUrl)}" target="_blank" rel="noopener noreferrer" style="color:var(--ios-orange);display:flex">Открыть в Битрикс24 ↗</a>
      </div>
      ` : ''}

      ${(task.rawText || task.description) ? `
      <div class="modal-detail-card" style="margin:0">
        <div style="font-size:12px;font-weight:600;color:var(--ios-text-secondary);margin-bottom:6px">Детали из ТЗ и таблицы</div>
        <div class="modal-desc-text" style="max-height:220px;overflow-y:auto">${linkify(task.rawText || task.description)}</div>
      </div>
      ` : ''}

      <div class="modal-btn-row" style="margin-top:auto">
        <button class="modal-btn secondary" id="desktopCopyBtn">Скопировать детали</button>
      </div>
    `;

    const copyTzBtn = document.getElementById('desktopCopyTzBtn');
    if (copyTzBtn && task.tzUrl) {
      copyTzBtn.onclick = () => {
        navigator.clipboard.writeText(task.tzUrl).then(() => {
          copyTzBtn.textContent = '✓ Скопировано!';
          setTimeout(() => { copyTzBtn.textContent = 'Копировать ссылку'; }, 1800);
        });
      };
    }

    const copyBtn = document.getElementById('desktopCopyBtn');
    if (copyBtn) {
      copyBtn.onclick = () => {
        let text = `Задача: ${task.title}\n`;
        if (task.date) text += `Дата: ${task.date}\n`;
        if (task.timeRange) text += `Время: ${task.timeRange}\n`;
        if (task.assignee) text += `Исполнитель: ${task.assignee}\n`;
        if (task.location) text += `Адрес: ${task.location}\n`;
        if (task.tzUrl) text += `ТЗ: ${task.tzUrl}\n`;
        navigator.clipboard.writeText(text).then(() => {
          copyBtn.textContent = '✓ Скопировано!';
          setTimeout(() => { copyBtn.textContent = 'Скопировать детали'; }, 1800);
        });
      };
    }
  }

  function renderSelectedDayTasks() {
    const list = document.getElementById('dayTasksList');
    list.innerHTML = '';

    if (!sourceUrl && allTasks.length === 0) {
      document.getElementById('dayTasksTitle').textContent = 'Доступ закрыт';
      document.getElementById('dayTasksCount').textContent = '';
      list.innerHTML = `
        <div style="text-align:center;padding:36px 20px;color:var(--ios-text-secondary)">
          <div style="font-size:42px;margin-bottom:12px">🔒</div>
          <div style="font-weight:600;font-size:16px;color:var(--ios-text-primary);margin-bottom:6px">Расписание не подключено</div>
          <div style="font-size:14px;line-height:1.4;margin-bottom:16px">Укажите секретную ссылку на таблицу в Настройках для доступа к задачам.</div>
          <button class="modal-btn primary" onclick="document.querySelector('[data-tab=settings]').click()" style="max-width:220px;margin:0 auto">Перейти в Настройки</button>
        </div>
      `;
      renderDesktopInspectorDay(selectedDate, []);
      return;
    }

    const selectedDateStr = toDateString(selectedDate);
    const dayStr = selectedDate.toLocaleDateString('ru-RU', { day: 'numeric', month: 'long', weekday: 'long' });
    document.getElementById('dayTasksTitle').textContent = `Задачи на ${dayStr}`;

    const tasks = allTasks.filter(t => {
      if (t.date !== selectedDateStr) return false;
      if (!matchesAssignee(t, selectedAssignee)) return false;
      if (hideCompleted && t.status.isCompleted) return false;
      if (isDutyTask(t) && !isDutyInNearestMonth(t.date)) return false;
      return true;
    }).sort((a, b) => (a.timeRange || '').localeCompare(b.timeRange || ''));

    document.getElementById('dayTasksCount').textContent = `${tasks.length} ${pluralize(tasks.length)}`;

    if (!activeModalTask) {
      renderDesktopInspectorDay(selectedDate, tasks);
    } else {
      renderDesktopInspectorTask(activeModalTask);
    }

    if (tasks.length === 0) {
      list.innerHTML = `<div style="text-align:center;color:var(--ios-text-secondary);padding:30px">На этот день задач нет</div>`;
      return;
    }

    tasks.forEach(t => list.appendChild(createTaskCard(t, false)));
  }

  function createTaskCard(task, showDate) {
    const card = document.createElement('div');
    card.className = 'task-card';

    let dateBadge = '';
    if (showDate && task.date) {
      const parts = task.date.split('-');
      if (parts.length === 3) {
        const d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
        dateBadge = `<span class="time-badge">${d.toLocaleDateString('ru-RU', { day: 'numeric', month: 'short' })}</span>`;
      }
    }

    const domainName = task.tzUrl ? getDomain(task.tzUrl) : '';

    card.innerHTML = `
      <div class="card-header">
        <span class="status-badge" style="background:${task.status.color}20;color:${task.status.color}">
          <span style="width:6px;height:6px;border-radius:50%;background:${task.status.color}"></span>
          ${task.status.name}
        </span>
        <div style="display:flex;gap:6px;align-items:center">
          ${task.timeRange ? `<span class="time-badge">🕒 ${task.timeRange}</span>` : ''}
          ${dateBadge}
        </div>
      </div>
      <div class="card-title">${escapeHtml(task.title)}</div>
      ${task.detailsSnippet ? `<div class="card-details-snippet">📝 ${escapeHtml(task.detailsSnippet)}</div>` : ''}
      ${task.location ? `<div class="card-location">📍 ${escapeHtml(task.location)}</div>` : ''}
      ${task.tzUrl ? `
        <div class="card-tz-row">
          <a class="card-tz-btn" href="${escapeHtml(task.tzUrl)}" target="_blank" rel="noopener noreferrer" onclick="event.stopPropagation()">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>
            <span>Открыть файл ТЗ ↗</span>
          </a>
          <span class="card-tz-domain">${escapeHtml(domainName)}</span>
        </div>
      ` : ''}
      <div class="card-footer">
        <span class="assignee-chip">👤 ${escapeHtml(task.assignee)}</span>
        <div class="card-actions">
          ${task.bitrixUrl ? `<a class="action-pill b24" href="${escapeHtml(task.bitrixUrl)}" target="_blank" onclick="event.stopPropagation()">Битрикс24 ↗</a>` : ''}
        </div>
      </div>
    `;

    card.addEventListener('click', () => openTaskModal(task));
    return card;
  }

  function renderAgenda() {
    const container = document.getElementById('agendaSectionsList');
    container.innerHTML = '';

    if (!sourceUrl && allTasks.length === 0) {
      container.innerHTML = `
        <div style="text-align:center;padding:50px 20px;color:var(--ios-text-secondary)">
          <div style="font-size:44px;margin-bottom:14px">🔒</div>
          <div style="font-weight:600;font-size:17px;color:var(--ios-text-primary);margin-bottom:8px">Расписание заблокировано</div>
          <div style="font-size:14px;line-height:1.4;margin-bottom:18px">Без секретной ссылки расписание недоступно.<br>Укажите ссылку на файл в Настройках.</div>
          <button class="modal-btn primary" onclick="document.querySelector('[data-tab=settings]').click()" style="max-width:220px;margin:0 auto">Перейти в Настройки</button>
        </div>
      `;
      return;
    }

    const todayStr = getTodayString();
    const tomorrow = new Date(); tomorrow.setDate(tomorrow.getDate() + 1);
    const tomorrowStr = toDateString(tomorrow);
    const endOfWeek = new Date(); endOfWeek.setDate(endOfWeek.getDate() + 7);
    const endOfWeekStr = toDateString(endOfWeek);

    let filtered = allTasks.filter(t => {
      if (!matchesAssignee(t, selectedAssignee)) return false;
      if (hideCompleted && t.status.isCompleted) return false;
      if (isDutyTask(t) && !isDutyInNearestMonth(t.date)) return false;
      if (searchText) {
        const q = searchText.toLowerCase();
        return t.title.toLowerCase().includes(q) || 
               (t.location || '').toLowerCase().includes(q) ||
               (t.description || '').toLowerCase().includes(q) ||
               (t.assignee || '').toLowerCase().includes(q);
      }
      return true;
    });

    const todayTasks = [];
    const tomorrowTasks = [];
    const thisWeekTasks = [];
    const upcomingTasks = [];
    const noDateTasks = [];
    const completedPastTasks = [];

    filtered.forEach(t => {
      if (t.status.isCompleted) {
        completedPastTasks.push(t);
        return;
      }
      if (!t.date) {
        noDateTasks.push(t);
        return;
      }
      if (t.date < todayStr) {
        completedPastTasks.push(t);
      } else if (t.date === todayStr) {
        todayTasks.push(t);
      } else if (t.date === tomorrowStr) {
        tomorrowTasks.push(t);
      } else if (t.date <= endOfWeekStr) {
        thisWeekTasks.push(t);
      } else {
        upcomingTasks.push(t);
      }
    });

    const sortFn = getSortComparator();
    todayTasks.sort(sortFn);
    tomorrowTasks.sort(sortFn);
    thisWeekTasks.sort(sortFn);
    upcomingTasks.sort(sortFn);
    noDateTasks.sort(sortFn);
    completedPastTasks.sort(sortFn);

    const sections = [
      { id: 'today', title: '🕒 Сегодня', tasks: todayTasks },
      { id: 'tomorrow', title: '📅 Завтра', tasks: tomorrowTasks },
      { id: 'this_week', title: '🗓 На этой неделе', tasks: thisWeekTasks },
      { id: 'upcoming', title: '➡️ Предстоящие', tasks: upcomingTasks },
      { id: 'no_date', title: '❓ Без даты', tasks: noDateTasks },
      { id: 'completed_past', title: '✅ Завершённые и прошедшие', tasks: completedPastTasks, isPast: true }
    ];

    let renderedAny = false;
    sections.forEach(sec => {
      if (sec.tasks.length === 0) return;
      renderedAny = true;

      const isCollapsed = collapsedSections.has(sec.id);
      const secDiv = document.createElement('div');
      secDiv.className = `agenda-section ${isCollapsed ? 'collapsed' : ''}`;

      secDiv.innerHTML = `
        <div class="section-header">
          <div class="section-title-wrap">
            <span class="section-title">${sec.title}</span>
            <span class="section-count">${sec.tasks.length}</span>
          </div>
          <span class="section-chevron">▼</span>
        </div>
        <div class="section-content"></div>
      `;

      secDiv.querySelector('.section-header').addEventListener('click', () => {
        if (collapsedSections.has(sec.id)) {
          collapsedSections.delete(sec.id);
          secDiv.classList.remove('collapsed');
        } else {
          collapsedSections.add(sec.id);
          secDiv.classList.add('collapsed');
        }
      });

      const content = secDiv.querySelector('.section-content');
      sec.tasks.forEach(t => content.appendChild(createTaskCard(t, true)));
      container.appendChild(secDiv);
    });

    if (!renderedAny) {
      container.innerHTML = `<div style="text-align:center;color:var(--ios-text-secondary);padding:40px">Задач не найдено</div>`;
    }
  }

  function getSortComparator() {
    switch (sortOrder) {
      case 'dateDesc':
        return (a, b) => (b.date || '').localeCompare(a.date || '');
      case 'byAssignee':
        return (a, b) => (a.assignee || 'я').localeCompare(b.assignee || 'я');
      case 'byStatus':
        return (a, b) => a.status.priority - b.status.priority;
      case 'dateAsc':
      default:
        return (a, b) => (a.date || '9999').localeCompare(b.date || '9999');
    }
  }

  // --- Task Detail Modal ---
  function openTaskModal(task) {
    activeModalTask = task;
    const badge = document.getElementById('modalStatusBadge');
    badge.textContent = task.status.name;
    badge.style.background = task.status.color + '20';
    badge.style.color = task.status.color;

    document.getElementById('modalTitle').textContent = task.title;

    let dtStr = 'Дата не указана';
    if (task.date) {
      const parts = task.date.split('-');
      if (parts.length === 3) {
        const d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
        dtStr = d.toLocaleDateString('ru-RU', { dateStyle: 'full' });
      }
    }
    if (task.timeRange) dtStr += ` • ${task.timeRange}`;
    document.getElementById('modalDateTime').textContent = dtStr;

    const dlEl = document.getElementById('modalDeadline');
    if (task.deadline) {
      dlEl.textContent = `Дедлайн: ${task.deadline}`;
      dlEl.style.display = 'block';
    } else {
      dlEl.style.display = 'none';
    }

    const locCard = document.getElementById('modalLocationCard');
    if (task.location) {
      locCard.style.display = 'block';
      document.getElementById('modalLocationText').textContent = task.location;
      document.getElementById('modalAppleMapsBtn').onclick = () => {
        window.open(`maps://?q=${encodeURIComponent(task.location)}`, '_system');
      };
      document.getElementById('modalYandexMapsBtn').onclick = () => {
        window.open(`https://yandex.ru/maps/?text=${encodeURIComponent(task.location)}`, '_blank');
      };
    } else {
      locCard.style.display = 'none';
    }

    // ТЗ Section
    const tzSection = document.getElementById('modalTzSection');
    if (task.tzUrl) {
      tzSection.style.display = 'block';
      const tzUrlLink = document.getElementById('modalTzUrlLink');
      tzUrlLink.href = task.tzUrl;
      tzUrlLink.textContent = task.tzUrl;

      const tzLink = document.getElementById('modalTzLink');
      tzLink.href = task.tzUrl;

      document.getElementById('modalCopyTzBtn').onclick = (e) => {
        e.preventDefault();
        navigator.clipboard.writeText(task.tzUrl).then(() => {
          const btn = document.getElementById('modalCopyTzBtn');
          const orig = btn.textContent;
          btn.textContent = '✓ Скопировано!';
          setTimeout(() => { btn.textContent = orig; }, 1800);
        }).catch(() => {
          prompt('Ссылка на ТЗ:', task.tzUrl);
        });
      };
    } else {
      tzSection.style.display = 'none';
    }

    // Bitrix24 Section
    const b24Section = document.getElementById('modalBitrixSection');
    if (task.bitrixUrl) {
      b24Section.style.display = 'block';
      document.getElementById('modalB24Link').href = task.bitrixUrl;
    } else {
      b24Section.style.display = 'none';
    }

    // Full Details / Description with active clickable links
    const descCard = document.getElementById('modalDescCard');
    const fullText = task.rawText || task.description || '';
    if (fullText) {
      descCard.style.display = 'block';
      document.getElementById('modalDescText').innerHTML = linkify(fullText);
    } else {
      descCard.style.display = 'none';
    }

    document.getElementById('taskModal').classList.add('active');
    renderDesktopInspectorTask(task);
  }

  function closeTaskModal() {
    document.getElementById('taskModal').classList.remove('active');
    activeModalTask = null;
    renderSelectedDayTasks();
  }

  function renderAll() {
    updateBadges();
    if (activeTab === 'calendar') {
      renderCalendar();
    } else if (activeTab === 'agenda') {
      renderAgenda();
    }
  }

  function escapeHtml(str) {
    if (!str) return '';
    return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function pluralize(n) {
    const mod10 = n % 10;
    const mod100 = n % 100;
    if (mod100 >= 11 && mod100 <= 19) return 'задач';
    if (mod10 === 1) return 'задача';
    if (mod10 >= 2 && mod10 <= 4) return 'задачи';
    return 'задач';
  }

  // --- Event Listeners Setup ---
  function setupEvents() {
    // Unified Tab switching for Mobile & Desktop
    function switchTab(tab) {
      activeTab = tab;
      document.querySelectorAll('.tab-btn').forEach(b => {
        b.classList.toggle('active', b.dataset.tab === tab);
      });
      document.querySelectorAll('.d-view-tab').forEach(b => {
        b.classList.toggle('active', b.dataset.tab === tab);
      });

      document.querySelectorAll('.tab-view').forEach(v => v.classList.remove('active'));
      const activeEl = document.getElementById(`tab-${tab}`);
      if (activeEl) activeEl.classList.add('active');

      const titleMap = { calendar: 'Календарь', agenda: 'Расписание', settings: 'Настройки' };
      const pageTitle = document.getElementById('pageTitle');
      if (pageTitle) pageTitle.textContent = titleMap[tab] || 'Задачи АГ';

      const fb = document.getElementById('filterBar');
      if (fb) fb.style.display = (tab === 'settings') ? 'none' : '';

      renderAll();
    }

    document.querySelectorAll('.tab-btn').forEach(btn => {
      btn.addEventListener('click', () => switchTab(btn.dataset.tab));
    });

    document.querySelectorAll('.d-view-tab').forEach(btn => {
      btn.addEventListener('click', () => switchTab(btn.dataset.tab));
    });

    // Month Navigation (Mobile & Desktop)
    const handlePrevMonth = () => {
      currentMonth.setMonth(currentMonth.getMonth() - 1);
      renderCalendar();
    };
    const handleNextMonth = () => {
      currentMonth.setMonth(currentMonth.getMonth() + 1);
      renderCalendar();
    };
    const handleToday = () => {
      selectedDate = new Date();
      currentMonth = new Date(selectedDate.getFullYear(), selectedDate.getMonth(), 1);
      renderCalendar();
    };

    document.getElementById('prevMonthBtn').addEventListener('click', handlePrevMonth);
    document.getElementById('nextMonthBtn').addEventListener('click', handleNextMonth);
    document.getElementById('todayBtn').addEventListener('click', handleToday);

    const prevDesktop = document.getElementById('prevMonthBtnDesktop');
    if (prevDesktop) prevDesktop.addEventListener('click', handlePrevMonth);
    const nextDesktop = document.getElementById('nextMonthBtnDesktop');
    if (nextDesktop) nextDesktop.addEventListener('click', handleNextMonth);
    const todayDesktop = document.getElementById('todayBtnDesktop');
    if (todayDesktop) todayDesktop.addEventListener('click', handleToday);

    // Search synchronization (Mobile & Desktop)
    const desktopSearch = document.getElementById('desktopSearchInput');
    const mobileSearch = document.getElementById('searchInput');

    if (desktopSearch) {
      desktopSearch.addEventListener('input', (e) => {
        searchText = e.target.value.trim();
        if (mobileSearch) mobileSearch.value = e.target.value;
        if (activeTab !== 'agenda' && searchText.length > 0) {
          switchTab('agenda');
        } else {
          renderAgenda();
        }
      });
    }

    if (mobileSearch) {
      mobileSearch.addEventListener('input', (e) => {
        searchText = e.target.value.trim();
        if (desktopSearch) desktopSearch.value = e.target.value;
        renderAgenda();
      });
    }

    const sortSel = document.getElementById('sortSelect');
    if (sortSel) {
      sortSel.value = sortOrder;
      sortSel.addEventListener('change', (e) => {
        sortOrder = e.target.value;
        localStorage.setItem('agendaSortOrder', sortOrder);
        renderAgenda();
      });
    }

    // Hide completed toggle (Mobile & Desktop)
    function toggleHideCompleted() {
      hideCompleted = !hideCompleted;
      localStorage.setItem('hideCompleted', hideCompleted ? 'true' : 'false');

      const hideBtn = document.getElementById('hideCompletedBtn');
      if (hideBtn) hideBtn.classList.toggle('active', hideCompleted);
      const dot = document.getElementById('hideCompletedDot');
      if (dot) dot.textContent = hideCompleted ? '●' : '○';

      const dBtn = document.getElementById('desktopHideCompletedBtn');
      if (dBtn) dBtn.classList.toggle('active', hideCompleted);
      const dDot = document.getElementById('desktopHideCompletedDot');
      if (dDot) dDot.textContent = hideCompleted ? '●' : '○';

      renderAll();
    }

    const hideBtn = document.getElementById('hideCompletedBtn');
    if (hideBtn) {
      if (hideCompleted) hideBtn.classList.add('active');
      hideBtn.addEventListener('click', toggleHideCompleted);
    }

    const dHideBtn = document.getElementById('desktopHideCompletedBtn');
    if (dHideBtn) {
      if (hideCompleted) dHideBtn.classList.add('active');
      dHideBtn.addEventListener('click', toggleHideCompleted);
    }

    // Desktop Inspector close button
    const inspectorClose = document.getElementById('inspectorCloseBtn');
    if (inspectorClose) {
      inspectorClose.addEventListener('click', () => {
        activeModalTask = null;
        renderCalendar();
        renderSelectedDayTasks();
      });
    }

    // Sync buttons
    document.getElementById('syncBtn').addEventListener('click', () => syncTasks(true, false));
    document.getElementById('manualSyncBtn').addEventListener('click', () => syncTasks(true, false));

    // Settings URL
    const urlInput = document.getElementById('sourceUrlInput');
    urlInput.value = sourceUrl || '';
    document.getElementById('saveUrlBtn').addEventListener('click', async () => {
      const val = urlInput.value.trim();
      if (!val) {
        alert('Пожалуйста, вставьте ссылку на файл Яндекс.Диска.');
        return;
      }
      sourceUrl = val;
      localStorage.setItem('sourceUrl', sourceUrl);
      alert('Ссылка сохранена! Загружаем расписание...');
      await syncTasks(true);
      const calTabBtn = document.querySelector('[data-tab="calendar"]');
      if (calTabBtn) calTabBtn.click();
    });
    document.getElementById('resetUrlBtn').addEventListener('click', () => {
      if (confirm('Очистить ссылку и удалить сохранённое расписание с этого устройства?')) {
        sourceUrl = '';
        localStorage.removeItem('sourceUrl');
        localStorage.removeItem('cached_tasks');
        localStorage.removeItem('last_sync_date');
        allTasks = [];
        urlInput.value = '';
        renderAll();
        alert('Ссылка и расписание удалены с этого устройства. Доступ заблокирован.');
      }
    });

    // Modal Sheet close
    document.getElementById('modalCloseBtn').addEventListener('click', closeTaskModal);
    document.getElementById('taskModal').addEventListener('click', (e) => {
      if (e.target.id === 'taskModal') closeTaskModal();
    });

    // Modal Copy and Share
    document.getElementById('modalCopyBtn').addEventListener('click', () => {
      if (!activeModalTask) return;
      let text = `Задача: ${activeModalTask.title}\n`;
      if (activeModalTask.date) text += `Дата: ${activeModalTask.date}\n`;
      if (activeModalTask.timeRange) text += `Время: ${activeModalTask.timeRange}\n`;
      if (activeModalTask.assignee) text += `Исполнитель: ${activeModalTask.assignee}\n`;
      if (activeModalTask.location) text += `Адрес: ${activeModalTask.location}\n`;
      if (activeModalTask.tzUrl) text += `ТЗ: ${activeModalTask.tzUrl}\n`;
      navigator.clipboard.writeText(text).then(() => {
        const btn = document.getElementById('modalCopyBtn');
        const orig = btn.textContent;
        btn.textContent = '✓ Скопировано!';
        setTimeout(() => { btn.textContent = orig; }, 1800);
      }).catch(() => {
        prompt('Скопируйте детали задачи:', text);
      });
    });

    document.getElementById('modalShareBtn').addEventListener('click', () => {
      if (!activeModalTask) return;
      const shareData = {
        title: activeModalTask.title,
        text: `${activeModalTask.title}\n${activeModalTask.timeRange || ''}\n${activeModalTask.location || ''}\n${activeModalTask.tzUrl || ''}`
      };
      if (navigator.share) {
        navigator.share(shareData).catch(() => {});
      } else {
        alert('Поделиться:\n' + shareData.text);
      }
    });

    document.querySelectorAll('.filter-pill').forEach(p => {
      p.classList.toggle('active', p.dataset.assignee === selectedAssignee);
    });
  }

  // --- Init ---
  function init() {
    renderAssigneeFilters();
    setupEvents();

    if ('serviceWorker' in navigator) {
      navigator.serviceWorker.getRegistrations().then(regs => {
        for (let r of regs) r.unregister();
      }).catch(err => console.log('SW unregister error:', err));
    }

    // Check if secret URL was passed via parameter or hash
    const urlParams = new URLSearchParams(window.location.search);
    let passedUrl = urlParams.get('url');
    if (!passedUrl && window.location.hash.startsWith('#http')) {
      passedUrl = window.location.hash.slice(1);
    }
    if (passedUrl) {
      sourceUrl = passedUrl.trim();
      localStorage.setItem('sourceUrl', sourceUrl);
      const urlInput = document.getElementById('sourceUrlInput');
      if (urlInput) urlInput.value = sourceUrl;
      try {
        window.history.replaceState({}, document.title, window.location.pathname);
      } catch (e) {}
    }

    const cached = localStorage.getItem('cached_tasks');
    if (cached) {
      try {
        allTasks = filterTasks(JSON.parse(cached));
      } catch (e) {
        allTasks = [];
      }
    } else {
      allTasks = [];
    }

    renderAll();

    // Silently fetch fresh online events from events.json in cloud
    loadOnlineEvents().then(success => {
      if (success) {
        renderAll();
      } else if (sourceUrl && allTasks.length === 0) {
        syncTasks();
      }
    });

    // Auto-update when app tab is reopened or screen unlocked
    document.addEventListener('visibilitychange', () => {
      if (document.visibilityState === 'visible') {
        loadOnlineEvents().then(ok => { if (ok) renderAll(); });
      }
    });
  }

  window.addEventListener('DOMContentLoaded', init);
})();
