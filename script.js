(() => {
  'use strict';

  const STATUS = {
    WEEKLY: '週週聚會',
    REGULAR: '常聚會',
    OCCASIONAL: '偶聚會',
    SPORADIC: '零星聚會',
    INACTIVE: '近期未聚會'
  };

  const STATUS_ORDER = [STATUS.WEEKLY, STATUS.REGULAR, STATUS.OCCASIONAL, STATUS.SPORADIC, STATUS.INACTIVE];
  const STATUS_CLASS = {
    [STATUS.WEEKLY]: 'weekly',
    [STATUS.REGULAR]: 'regular',
    [STATUS.OCCASIONAL]: 'occasional',
    [STATUS.SPORADIC]: 'sporadic',
    [STATUS.INACTIVE]: 'inactive'
  };

  const MS_WEEK = 7 * 24 * 60 * 60 * 1000;

  const state = {
    fileName: '',
    rows: [],
    dateColumns: [],
    people: [],
    windowWeeks: 12,
    indexes: {},
    recentCols: [],
    selectedGroups: new Set(),
    selectedStatuses: new Set()
  };

  const $ = (id) => document.getElementById(id);

  const els = {
    fileInput: $('fileInput'),
    fileName: $('fileName'),
    error: $('errorMessage'),
    dashboard: $('dashboard'),
    clearBtn: $('clearBtn'),
    infoFile: $('infoFile'),
    infoPeriod: $('infoPeriod'),
    infoWeeks: $('infoWeeks'),
    infoPeople: $('infoPeople'),
    windowSelect: $('windowSelect'),
    ruleText: $('ruleText'),
    summaryCards: $('summaryCards'),
    overallBar: $('overallBar'),
    overallLegend: $('overallLegend'),
    districtBody: document.querySelector('#districtTable tbody'),
    searchInput: $('searchInput'),
    districtFilter: $('districtFilter'),
    smallDistrictFilter: $('smallDistrictFilter'),
    newBelieverFilter: $('newBelieverFilter'),

    groupMultiSelect: $('groupMultiSelect'),
    groupFilterSummary: $('groupFilterSummary'),
    groupFilterOptions: $('groupFilterOptions'),
    clearGroupFilter: $('clearGroupFilter'),
    closeGroupFilter: $('closeGroupFilter'),

    statusMultiSelect: $('statusMultiSelect'),
    statusFilterSummary: $('statusFilterSummary'),
    statusFilterOptions: $('statusFilterOptions'),
    clearStatusFilter: $('clearStatusFilter'),
    closeStatusFilter: $('closeStatusFilter'),

    resultCount: $('resultCount'),
    peopleBody: document.querySelector('#peopleTable tbody')
  };

  function excelDateToDate(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) return startOfDay(value);

    if (typeof value === 'number') {
      const parts = XLSX.SSF.parse_date_code(value);
      if (!parts) return null;
      return new Date(parts.y, parts.m - 1, parts.d);
    }

    if (typeof value === 'string') {
      const s = value.trim();
      if (!s) return null;

      let m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})$/);
      if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));

      m = s.match(/^(\d{4})[-\/.](\d{1,2})$/);
      if (m) return new Date(Number(m[1]), Number(m[2]) - 1, 1);
    }

    return null;
  }

  function startOfDay(d) {
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }

  function isAttendance(value) {
    if (typeof value === 'number') return value > 0;
    if (typeof value === 'boolean') return value;

    if (typeof value === 'string') {
      const s = value.trim().toLowerCase();
      return ['1', 'v', '✓', '✔', 'true', '有', '出席'].includes(s);
    }

    return false;
  }

  function fmtDate(d) {
    return d ? `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}` : '—';
  }

  function analyzeWorkbook(arrayBuffer, fileName) {
    const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: false, raw: true });
    if (!workbook.SheetNames.length) throw new Error('找不到工作表。');

    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: '' });
    if (rows.length < 2) throw new Error('工作表沒有足夠資料。');

    const header = rows[0].map(v => String(v ?? '').trim());
    const required = ['大區', '小區', '姓名', '羣組'];
    const idx = Object.fromEntries(required.map(name => [name, header.indexOf(name)]));
    const missing = required.filter(name => idx[name] < 0);

    if (missing.length) throw new Error(`缺少必要欄位：${missing.join('、')}`);

    const baptismNames = ['受浸日期', '受浸日', '受浸'];
    const baptismIndex = baptismNames.map(n => header.indexOf(n)).find(i => i >= 0) ?? -1;

    const today = new Date();
    today.setHours(23, 59, 59, 999);

    const dateColumns = [];
    rows[0].forEach((value, col) => {
      const d = excelDateToDate(value);
      if (d && d <= today) dateColumns.push({ col, date: d });
    });

    dateColumns.sort((a, b) => a.date - b.date);
    if (!dateColumns.length) throw new Error('找不到可辨識的聚會日期欄位。');

    state.fileName = fileName;
    state.rows = rows;
    state.dateColumns = dateColumns;
    state.indexes = { ...idx, baptism: baptismIndex };

    recalculate();

    if (baptismIndex < 0) {
      els.error.textContent = '提醒：這份 Excel 找不到「受浸日期」欄位，因此目前無法判斷初信；其他聚會分析不受影響。';
    }
  }

  function classify({ attended, totalWeeks, weekly12Attendance, weekly12Weeks, weeksSinceLastAttendance }) {
    if (!totalWeeks || weeksSinceLastAttendance === null || weeksSinceLastAttendance > 8) {
      return STATUS.INACTIVE;
    }

    if (weekly12Weeks >= 12 && weekly12Attendance === weekly12Weeks) {
      return STATUS.WEEKLY;
    }

    const regularNeeded = Math.ceil(totalWeeks / 3);
    if (attended >= regularNeeded && weeksSinceLastAttendance <= 3) {
      return STATUS.REGULAR;
    }

    const occasionalNeeded = Math.ceil(totalWeeks / 4);
    if (attended >= occasionalNeeded && weeksSinceLastAttendance <= 4) {
      return STATUS.OCCASIONAL;
    }

    return STATUS.SPORADIC;
  }

  function recalculate() {
    if (!state.rows.length) return;

    state.windowWeeks = Number(els.windowSelect.value || 12);
    const recentCols = state.dateColumns.slice(-state.windowWeeks);
    state.recentCols = recentCols;

    const referenceDate = recentCols.length ? recentCols[recentCols.length - 1].date : null;
    const { indexes } = state;

    state.people = state.rows.slice(1).map((row, i) => {
      const name = String(row[indexes['姓名']] ?? '').trim();
      if (!name) return null;

      const recentAttendance = recentCols.reduce(
        (sum, x) => sum + (isAttendance(row[x.col]) ? 1 : 0),
        0
      );

      const weekly12Cols = state.dateColumns.slice(-12);
      const weekly12Attendance = weekly12Cols.reduce(
        (sum, x) => sum + (isAttendance(row[x.col]) ? 1 : 0),
        0
      );

      const attendedDates = state.dateColumns
        .filter(x => isAttendance(row[x.col]))
        .map(x => x.date);

      const lastAttendanceDate = attendedDates.length
        ? attendedDates[attendedDates.length - 1]
        : null;

      const weeksSinceLastAttendance = lastAttendanceDate && referenceDate
        ? Math.floor((referenceDate - lastAttendanceDate) / MS_WEEK)
        : null;

      const totalAttendance = attendedDates.length;

      const baptismDate = indexes.baptism >= 0
        ? excelDateToDate(row[indexes.baptism])
        : null;

      const weeksSinceBaptism = baptismDate && referenceDate
        ? Math.floor((referenceDate - baptismDate) / MS_WEEK)
        : null;

      let newBelieverStatus = 'unknown';
      if (weeksSinceBaptism !== null) {
        newBelieverStatus =
          (weeksSinceBaptism >= 0 && weeksSinceBaptism <= 104) ? 'yes' : 'no';
      }

      return {
        rowNumber: i + 2,
        district: String(row[indexes['大區']] ?? '').trim(),
        smallDistrict: String(row[indexes['小區']] ?? '').trim(),
        name,
        group: String(row[indexes['羣組']] ?? '').trim(),
        baptismDate,
        newBelieverStatus,
        isNewBeliever: newBelieverStatus === 'yes',
        recentAttendance,
        recentWeeks: recentCols.length,
        recentRate: recentCols.length ? recentAttendance / recentCols.length : 0,
        totalAttendance,
        lastAttendanceDate,
        weeksSinceLastAttendance,
        weekly12Attendance,
        weekly12Weeks: weekly12Cols.length,
        status: classify({
          attended: recentAttendance,
          totalWeeks: recentCols.length,
          weekly12Attendance,
          weekly12Weeks: weekly12Cols.length,
          weeksSinceLastAttendance
        })
      };
    }).filter(Boolean);

    renderAll();
  }

  function countStatuses(people) {
    const counts = Object.fromEntries(STATUS_ORDER.map(s => [s, 0]));
    people.forEach(p => counts[p.status]++);
    return counts;
  }

  function renderAll() {
    renderInfo();
    renderRuleText();
    renderSummary();
    renderOverall();
    renderDistricts();
    buildFilters();
    renderPeople();

    els.dashboard.classList.remove('hidden');
    els.clearBtn.disabled = false;
  }

  function renderInfo() {
    const allDates = state.dateColumns;
    const recent = state.recentCols;
    const latestDate = allDates[allDates.length - 1].date;

    els.infoFile.textContent = state.fileName;
    els.infoPeriod.textContent = `${fmtDate(allDates[0].date)} – ${fmtDate(latestDate)}`;
    els.infoWeeks.textContent = `${allDates.length} 週｜目前分析：截至 ${fmtDate(latestDate)} 最近 ${recent.length} 週`;
    els.infoPeople.textContent = `${state.people.length} 人`;
  }

  function renderRuleText() {
    const actual = state.recentCols.length;
    const latestDate = state.dateColumns.length
      ? state.dateColumns[state.dateColumns.length - 1].date
      : null;

    const earliestRecent = actual ? state.recentCols[0].date : null;

    const periodText = actual < state.windowWeeks
      ? `檔案目前只有 ${actual} 個有效週次`
      : `${fmtDate(earliestRecent)} – ${fmtDate(latestDate)}`;

    els.ruleText.textContent =
      `以匯入資料中最新主日 ${fmtDate(latestDate)} 為「目前」，累計最近 ${state.windowWeeks} 週（${periodText}）。初信＝受浸日期在目前截止日前 104 週內。`;
  }

  function renderSummary() {
    const counts = countStatuses(state.people);
    const total = state.people.length || 1;
    const newBelievers = state.people.filter(p => p.newBelieverStatus === 'yes').length;
    const unknownBaptism = state.people.filter(p => p.newBelieverStatus === 'unknown').length;

    const cards = [
      ['總人數', state.people.length, '匯入名單'],
      ['初信', newBelievers, `${(newBelievers / total * 100).toFixed(1)}%`],
      ['日期不明', unknownBaptism, `${(unknownBaptism / total * 100).toFixed(1)}%`],
      ...STATUS_ORDER.map(s => [s, counts[s], `${(counts[s] / total * 100).toFixed(1)}%`])
    ];

    els.summaryCards.innerHTML = cards.map(([label, value, note]) => `
      <div class="card">
        <span class="label">${escapeHtml(label)}</span>
        <strong class="value">${value}</strong>
        <span class="note">${escapeHtml(note)}</span>
      </div>
    `).join('');
  }

  function renderOverall() {
    const counts = countStatuses(state.people);
    const total = state.people.length || 1;

    els.overallBar.innerHTML = STATUS_ORDER.map(s => {
      const pct = counts[s] / total * 100;
      return `<div class="seg-${STATUS_CLASS[s]}" style="width:${pct}%" title="${s} ${counts[s]} 人 (${pct.toFixed(1)}%)"></div>`;
    }).join('');

    els.overallLegend.innerHTML = STATUS_ORDER.map(s => {
      const pct = counts[s] / total * 100;
      return `
        <div class="legend-item">
          <span><i class="dot ${STATUS_CLASS[s]}"></i>${s}</span>
          <strong>${counts[s]} <small>${pct.toFixed(1)}%</small></strong>
        </div>`;
    }).join('');
  }

  function renderDistricts() {
    const groups = groupBy(state.people, p => p.district || '未分類');

    els.districtBody.innerHTML = [...groups.entries()]
      .sort((a, b) => a[0].localeCompare(b[0], 'zh-Hant'))
      .map(([district, people]) => {
        const c = countStatuses(people);
        const stableRate = people.length
          ? (c[STATUS.WEEKLY] + c[STATUS.REGULAR]) / people.length * 100
          : 0;

        const nb = people.filter(p => p.newBelieverStatus === 'yes').length;
        const unknown = people.filter(p => p.newBelieverStatus === 'unknown').length;

        return `
          <tr>
            <td><strong>${escapeHtml(district)}</strong></td>
            <td>${people.length}</td>
            <td>${nb}</td>
            <td>${unknown}</td>
            <td>${c[STATUS.WEEKLY]}</td>
            <td>${c[STATUS.REGULAR]}</td>
            <td>${c[STATUS.OCCASIONAL]}</td>
            <td>${c[STATUS.SPORADIC]}</td>
            <td>${c[STATUS.INACTIVE]}</td>
            <td>${stableRate.toFixed(1)}%</td>
          </tr>`;
      }).join('');
  }

  function buildFilters() {
    preserveOptions(els.districtFilter, unique(state.people.map(p => p.district)));
    updateSmallDistrictOptions();
    buildGroupOptions();
    buildStatusOptions();
  }

  function availableGroups() {
    return unique(state.people.map(p => p.group).filter(Boolean))
      .sort((a, b) => a.localeCompare(b, 'zh-Hant'));
  }

  function buildGroupOptions() {
    const groups = availableGroups();

    state.selectedGroups = new Set(
      [...state.selectedGroups].filter(g => groups.includes(g))
    );

    els.groupFilterOptions.innerHTML = groups.map(g => `
      <label class="group-option">
        <input type="checkbox" value="${escapeAttr(g)}" ${state.selectedGroups.has(g) ? 'checked' : ''}>
        <span>${escapeHtml(g)}</span>
      </label>
    `).join('');

    updateGroupSummary();
  }

  function updateGroupSummary() {
    const selected = [...state.selectedGroups];

    if (!selected.length) {
      els.groupFilterSummary.textContent = '全部';
    } else if (selected.length <= 2) {
      els.groupFilterSummary.textContent = selected.join('＋');
    } else {
      els.groupFilterSummary.textContent = `已選 ${selected.length} 組`;
    }
  }

  function setSelectedGroups(groups) {
    const available = availableGroups();
    const aliases = {
      college: ['大專', '大學'],
      youth: ['青職'],
      junior: ['國中'],
      senior: ['高中'],
      child: ['國小']
    };

    const resolve = key => aliases[key].filter(x => available.includes(x));

    let wanted = [];
    if (groups === 'all') wanted = [];
    if (groups === 'college-youth') wanted = [...resolve('college'), ...resolve('youth')];
    if (groups === 'teen') wanted = [...resolve('junior'), ...resolve('senior')];
    if (groups === 'student') wanted = [...resolve('college'), ...resolve('junior'), ...resolve('senior')];
    if (groups === 'child') wanted = [...resolve('child')];

    state.selectedGroups = new Set(wanted);
    buildGroupOptions();
    renderPeople();
  }

  function buildStatusOptions() {
    state.selectedStatuses = new Set(
      [...state.selectedStatuses].filter(s => STATUS_ORDER.includes(s))
    );

    els.statusFilterOptions.innerHTML = STATUS_ORDER.map(status => `
      <label class="group-option">
        <input type="checkbox" value="${escapeAttr(status)}" ${state.selectedStatuses.has(status) ? 'checked' : ''}>
        <span><i class="dot ${STATUS_CLASS[status]}"></i>${escapeHtml(status)}</span>
      </label>
    `).join('');

    updateStatusSummary();
  }

  function updateStatusSummary() {
    const selected = STATUS_ORDER.filter(s => state.selectedStatuses.has(s));

    if (!selected.length) {
      els.statusFilterSummary.textContent = '全部';
    } else if (selected.length === 1) {
      els.statusFilterSummary.textContent = selected[0];
    } else {
      els.statusFilterSummary.textContent = `已選 ${selected.length} 項`;
    }
  }

  function setSelectedStatuses(statuses) {
    state.selectedStatuses = new Set(statuses);
    buildStatusOptions();
    renderPeople();
  }

  function updateSmallDistrictOptions() {
    const selectedDistrict = els.districtFilter.value;

    const peopleInDistrict = selectedDistrict
      ? state.people.filter(p => p.district === selectedDistrict)
      : state.people;

    preserveOptions(
      els.smallDistrictFilter,
      unique(peopleInDistrict.map(p => p.smallDistrict))
    );
  }

  function preserveOptions(select, values) {
    const previous = select.value;

    select.innerHTML =
      '<option value="">全部</option>' +
      values
        .filter(Boolean)
        .sort((a, b) => a.localeCompare(b, 'zh-Hant'))
        .map(v => `<option value="${escapeAttr(v)}">${escapeHtml(v)}</option>`)
        .join('');

    if ([...select.options].some(o => o.value === previous)) {
      select.value = previous;
    }
  }

  function renderPeople() {
    const q = els.searchInput.value.trim().toLowerCase();
    const district = els.districtFilter.value;
    const small = els.smallDistrictFilter.value;
    const nb = els.newBelieverFilter.value;
    const selectedGroups = state.selectedGroups;
    const selectedStatuses = state.selectedStatuses;

    const filtered = state.people.filter(p =>
      (!q || p.name.toLowerCase().includes(q)) &&
      (!district || p.district === district) &&
      (!small || p.smallDistrict === small) &&
      (!selectedGroups.size || selectedGroups.has(p.group)) &&
      (!selectedStatuses.size || selectedStatuses.has(p.status)) &&
      (!nb || p.newBelieverStatus === nb)
    );

    els.resultCount.textContent = `顯示 ${filtered.length} / ${state.people.length} 人`;

    els.peopleBody.innerHTML = filtered.map(p => `
      <tr>
        <td>${escapeHtml(p.district)}</td>
        <td>${escapeHtml(p.smallDistrict)}</td>
        <td><strong>${escapeHtml(p.name)}</strong></td>
        <td>${escapeHtml(p.group)}</td>
        <td>${p.baptismDate ? fmtDate(p.baptismDate) : '—'}</td>
        <td>${
          p.newBelieverStatus === 'yes'
            ? '<span class="new-believer">初信</span>'
            : p.newBelieverStatus === 'no'
              ? '非初信'
              : '日期不明'
        }</td>
        <td><span class="status ${STATUS_CLASS[p.status]}">${p.status}</span></td>
        <td>${p.lastAttendanceDate ? fmtDate(p.lastAttendanceDate) : '—'}</td>
        <td>${p.recentAttendance} / ${p.recentWeeks}</td>
        <td>${(p.recentRate * 100).toFixed(1)}%</td>
        <td>${p.totalAttendance}</td>
      </tr>
    `).join('');
  }

  function closeMultiSelects(except = null) {
    [els.groupMultiSelect, els.statusMultiSelect].forEach(details => {
      if (details && details !== except) details.open = false;
    });
  }

  function groupBy(items, fn) {
    const m = new Map();
    items.forEach(x => {
      const k = fn(x);
      if (!m.has(k)) m.set(k, []);
      m.get(k).push(x);
    });
    return m;
  }

  function unique(arr) {
    return [...new Set(arr)];
  }

  function escapeHtml(v) {
    return String(v ?? '').replace(
      /[&<>'"]/g,
      c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[c])
    );
  }

  function escapeAttr(v) {
    return escapeHtml(v);
  }

  async function onFile(file) {
    els.error.textContent = '';
    els.fileName.textContent = file.name;

    try {
      if (!globalThis.XLSX) {
        throw new Error('Excel 解析元件尚未載入。請確認網路連線，或依 README 將 SheetJS 改為本機 vendor。');
      }

      const buffer = await file.arrayBuffer();
      analyzeWorkbook(buffer, file.name);
    } catch (err) {
      console.error(err);
      els.error.textContent = err?.message || '無法讀取這份 Excel。';
      els.dashboard.classList.add('hidden');
    }
  }

  function clearData() {
    state.fileName = '';
    state.rows = [];
    state.dateColumns = [];
    state.people = [];
    state.recentCols = [];
    state.windowWeeks = 12;
    state.selectedGroups = new Set();
    state.selectedStatuses = new Set();

    els.fileInput.value = '';
    els.fileName.textContent = '尚未選擇檔案';
    els.error.textContent = '';
    els.dashboard.classList.add('hidden');
    els.clearBtn.disabled = true;
    els.searchInput.value = '';

    closeMultiSelects();
  }

  els.fileInput.addEventListener('change', e => {
    const file = e.target.files?.[0];
    if (file) onFile(file);
  });

  els.windowSelect.addEventListener('change', () => {
    closeMultiSelects();
    recalculate();
  });

  els.districtFilter.addEventListener('change', () => {
    closeMultiSelects();
    updateSmallDistrictOptions();
    renderPeople();
  });

  [els.searchInput, els.smallDistrictFilter, els.newBelieverFilter].forEach(el => {
    el.addEventListener('input', () => {
      closeMultiSelects();
      renderPeople();
    });

    el.addEventListener('focus', () => closeMultiSelects());
  });

  els.groupFilterOptions.addEventListener('change', e => {
    if (!e.target.matches('input[type="checkbox"]')) return;

    if (e.target.checked) {
      state.selectedGroups.add(e.target.value);
    } else {
      state.selectedGroups.delete(e.target.value);
    }

    updateGroupSummary();
    renderPeople();
  });

  document.querySelectorAll('[data-group-preset]').forEach(btn => {
    btn.addEventListener('click', () => setSelectedGroups(btn.dataset.groupPreset));
  });

  els.clearGroupFilter.addEventListener('click', () => setSelectedGroups('all'));
  els.closeGroupFilter.addEventListener('click', () => {
    els.groupMultiSelect.open = false;
  });

  els.statusFilterOptions.addEventListener('change', e => {
    if (!e.target.matches('input[type="checkbox"]')) return;

    if (e.target.checked) {
      state.selectedStatuses.add(e.target.value);
    } else {
      state.selectedStatuses.delete(e.target.value);
    }

    updateStatusSummary();
    renderPeople();
  });

  els.clearStatusFilter.addEventListener('click', () => setSelectedStatuses([]));
  els.closeStatusFilter.addEventListener('click', () => {
    els.statusMultiSelect.open = false;
  });

  els.groupMultiSelect.addEventListener('toggle', () => {
    if (els.groupMultiSelect.open) closeMultiSelects(els.groupMultiSelect);
  });

  els.statusMultiSelect.addEventListener('toggle', () => {
    if (els.statusMultiSelect.open) closeMultiSelects(els.statusMultiSelect);
  });

  /* 點擊任何其他下拉式選單時，關閉羣組/聚會情況複選 */
  document.querySelectorAll('select').forEach(select => {
    select.addEventListener('pointerdown', () => closeMultiSelects());
    select.addEventListener('focus', () => closeMultiSelects());
  });

  /* 點擊複選區以外的地方，也自動收合 */
  document.addEventListener('pointerdown', e => {
    const insideGroup = els.groupMultiSelect.contains(e.target);
    const insideStatus = els.statusMultiSelect.contains(e.target);

    if (!insideGroup && !insideStatus) {
      closeMultiSelects();
    }
  });

  els.clearBtn.addEventListener('click', clearData);
})();
