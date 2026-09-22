const CONFIG = {
  recordsSheet: 'Tap_Records',
  auditSheet: 'Audit_Log',
  settingsSheet: 'Settings',
  timezone: Session.getScriptTimeZone() || 'Asia/Manila',
  taskTypes: ['Compression', 'RT', 'Free-Up'],
  operators: ['1 Operator', '2 Operators', '3 Operators']
};

const RECORD_HEADERS = [
  'Record ID',
  'Created At',
  'Work Date',
  'Operator',
  'SKU',
  'SKU Count',
  'Pallet Location',
  'Task Type',
  'Start Time',
  'End Time',
  'Active Duration Seconds',
  'Pause Duration Seconds',
  'Total Elapsed Seconds',
  'Status'
];

const AUDIT_HEADERS = [
  'Audit ID',
  'Timestamp',
  'Action',
  'Record ID',
  'Details'
];

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Stockman Productivity')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function setupWorkbook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const records = getOrCreateSheet_(ss, CONFIG.recordsSheet);
  ensureHeaders_(records, RECORD_HEADERS);

  const audit = getOrCreateSheet_(ss, CONFIG.auditSheet);
  ensureHeaders_(audit, AUDIT_HEADERS);

  const settings = getOrCreateSheet_(ss, CONFIG.settingsSheet);
  ensureHeaders_(settings, ['Setting', 'Value']);

  if (settings.getLastRow() === 1) {
    settings.getRange(2, 1, 3, 2).setValues([
      ['Timezone', CONFIG.timezone],
      ['Task Types', CONFIG.taskTypes.join(', ')],
      ['Operators', CONFIG.operators.join(', ')]
    ]);
  }

  records.setFrozenRows(1);
  audit.setFrozenRows(1);
  settings.setFrozenRows(1);

  records.getRange(1, 1, 1, RECORD_HEADERS.length)
    .setFontWeight('bold');

  audit.getRange(1, 1, 1, AUDIT_HEADERS.length)
    .setFontWeight('bold');

  settings.getRange(1, 1, 1, 2)
    .setFontWeight('bold');

  return {
    success: true,
    message: 'Workbook setup completed.'
  };
}

function getAppConfig() {
  return {
    success: true,
    taskTypes: getTaskTypes_(),
    operators: CONFIG.operators
  };
}

function saveTapTask(record) {
  validateRecord_(record);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(ss, CONFIG.recordsSheet);

  ensureHeaders_(sheet, RECORD_HEADERS);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const recordId = record.recordId || createId_('TASK');
    const now = new Date();

    const startTime = parseDate_(record.startTime);
    const endTime = parseDate_(record.endTime);

    const activeSeconds = Math.max(
      0,
      Number(record.activeDurationSeconds || 0)
    );

    const pauseSeconds = Math.max(
      0,
      Number(record.pauseDurationSeconds || 0)
    );

    const totalElapsedSeconds = Math.max(
      activeSeconds + pauseSeconds,
      Number(record.totalElapsedSeconds || 0)
    );

    const row = [
      recordId,
      now,
      record.workDate,
      String(record.operator).trim(),
      String(record.sku).trim(),
      Number(record.skuCount),
      String(record.palletLocation).trim(),
      String(record.taskType).trim(),
      startTime,
      endTime,
      activeSeconds,
      pauseSeconds,
      totalElapsedSeconds,
      'Completed'
    ];

    sheet.appendRow(row);

    writeAudit_(
      'CREATE_TASK',
      recordId,
      JSON.stringify({
        operator: record.operator,
        sku: record.sku,
        skuCount: record.skuCount,
        palletLocation: record.palletLocation,
        taskType: record.taskType
      })
    );

    return {
      success: true,
      recordId: recordId,
      message: 'Task saved successfully.'
    };
  } finally {
    lock.releaseLock();
  }
}

function getDashboardData(filters) {
  filters = filters || {};

  const records = readRecords_();

  const filtered = records.filter(record => {
    const dateMatch = !filters.workDate ||
      record.workDate === filters.workDate;

    const operatorMatch = !filters.operator ||
      record.operator === filters.operator;

    const taskMatch = !filters.taskType ||
      record.taskType === filters.taskType;

    return dateMatch && operatorMatch && taskMatch;
  });

  const completed = filtered.filter(
    record => record.status === 'Completed'
  );

  const totalActive = sum_(completed, 'activeDurationSeconds');
  const totalPause = sum_(completed, 'pauseDurationSeconds');
  const totalElapsed = sum_(completed, 'totalElapsedSeconds');
  const totalSkuCount = sum_(completed, 'skuCount');

  return {
    success: true,
    records: completed.slice().reverse(),
    summary: {
      taskCount: completed.length,
      totalSkuCount: totalSkuCount,
      totalActiveSeconds: totalActive,
      totalPauseSeconds: totalPause,
      totalElapsedSeconds: totalElapsed,
      averageActiveSeconds: completed.length
        ? Math.round(totalActive / completed.length)
        : 0
    },
    operatorSummary: groupSummary_(completed, 'operator'),
    taskSummary: groupSummary_(completed, 'taskType'),
    operators: unique_(records.map(record => record.operator))
      .filter(Boolean)
      .sort()
  };
}

function getAuditLog() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.auditSheet);

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  return sheet.getDataRange()
    .getValues()
    .slice(1)
    .reverse()
    .map(row => ({
      auditId: row[0],
      timestamp: formatDateTime_(row[1]),
      action: row[2],
      recordId: row[3],
      details: row[4]
    }));
}

function exportRecordsCsv(filters) {
  const data = getDashboardData(filters || {}).records;

  const headers = [
    'Record ID',
    'Created At',
    'Work Date',
    'Operator',
    'SKU',
    'SKU Count',
    'Pallet Location',
    'Task Type',
    'Start Time',
    'End Time',
    'Active Duration',
    'Pause Duration',
    'Total Elapsed',
    'Status'
  ];

  const rows = data.map(record => [
    record.recordId,
    record.createdAt,
    record.workDate,
    record.operator,
    record.sku,
    record.skuCount,
    record.palletLocation,
    record.taskType,
    record.startTime,
    record.endTime,
    secondsToText_(record.activeDurationSeconds),
    secondsToText_(record.pauseDurationSeconds),
    secondsToText_(record.totalElapsedSeconds),
    record.status
  ]);

  return [headers, ...rows]
    .map(row => row.map(csvEscape_).join(','))
    .join('\n');
}

function readRecords_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.recordsSheet);

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  return sheet.getDataRange()
    .getValues()
    .slice(1)
    .map(row => ({
      recordId: row[0],
      createdAt: formatDateTime_(row[1]),
      workDate: formatDate_(row[2]),
      operator: String(row[3] || ''),
      sku: String(row[4] || ''),
      skuCount: Number(row[5] || 0),
      palletLocation: String(row[6] || ''),
      taskType: String(row[7] || ''),
      startTime: formatDateTime_(row[8]),
      endTime: formatDateTime_(row[9]),
      activeDurationSeconds: Number(row[10] || 0),
      pauseDurationSeconds: Number(row[11] || 0),
      totalElapsedSeconds: Number(row[12] || 0),
      status: String(row[13] || '')
    }));
}

function getTaskTypes_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.settingsSheet);

  if (!sheet || sheet.getLastRow() < 2) {
    return CONFIG.taskTypes.slice();
  }

  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const key = String(rows[i][0] || '').trim();

    if (key === 'Task Types') {
      const raw = String(rows[i][1] || '').trim();

      if (!raw) {
        break;
      }

      const types = raw
        .split(',')
        .map(type => type.trim())
        .filter(Boolean);

      if (types.length) {
        return unique_(types);
      }
    }
  }

  return CONFIG.taskTypes.slice();
}

function groupSummary_(records, key) {
  const groups = {};

  records.forEach(record => {
    const group = record[key] || 'Unspecified';

    if (!groups[group]) {
      groups[group] = {
        name: group,
        taskCount: 0,
        skuCount: 0,
        activeSeconds: 0,
        pauseSeconds: 0,
        elapsedSeconds: 0
      };
    }

    groups[group].taskCount++;
    groups[group].skuCount += Number(record.skuCount || 0);
    groups[group].activeSeconds +=
      Number(record.activeDurationSeconds || 0);
    groups[group].pauseSeconds +=
      Number(record.pauseDurationSeconds || 0);
    groups[group].elapsedSeconds +=
      Number(record.totalElapsedSeconds || 0);
  });

  return Object.values(groups)
    .map(group => ({
      ...group,
      averageActiveSeconds: group.taskCount
        ? Math.round(group.activeSeconds / group.taskCount)
        : 0
    }))
    .sort((a, b) => b.activeSeconds - a.activeSeconds);
}

function validateRecord_(record) {
  if (!record) {
    throw new Error('Record is required.');
  }

  const required = [
    'workDate',
    'operator',
    'sku',
    'palletLocation',
    'taskType',
    'startTime',
    'endTime'
  ];

  required.forEach(field => {
    if (
      record[field] === undefined ||
      record[field] === null ||
      String(record[field]).trim() === ''
    ) {
      throw new Error('Missing required field: ' + field);
    }
  });

  if (!CONFIG.operators.includes(String(record.operator).trim())) {
    throw new Error('Invalid operator selection.');
  }

  const skuCount = Number(record.skuCount);

  if (!Number.isInteger(skuCount) || skuCount < 1) {
    throw new Error('SKU Count must be a whole number greater than zero.');
  }

  const taskTypes = getTaskTypes_();

  if (!taskTypes.includes(String(record.taskType).trim())) {
    throw new Error('Invalid task type.');
  }

  const startTime = parseDate_(record.startTime);
  const endTime = parseDate_(record.endTime);

  if (endTime < startTime) {
    throw new Error('End time cannot be earlier than start time.');
  }
}

function writeAudit_(action, recordId, details) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(ss, CONFIG.auditSheet);

  ensureHeaders_(sheet, AUDIT_HEADERS);

  sheet.appendRow([
    createId_('AUDIT'),
    new Date(),
    action,
    recordId,
    details || ''
  ]);
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function ensureHeaders_(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers]);
  }
}

function createId_(prefix) {
  return prefix + '-' +
    Utilities.getUuid()
      .replace(/-/g, '')
      .slice(0, 12)
      .toUpperCase();
}

function parseDate_(value) {
  if (value instanceof Date) {
    return value;
  }

  const date = new Date(value);

  if (isNaN(date.getTime())) {
    throw new Error('Invalid date: ' + value);
  }

  return date;
}

function formatDate_(value) {
  if (!value) {
    return '';
  }

  return Utilities.formatDate(
    parseDate_(value),
    CONFIG.timezone,
    'yyyy-MM-dd'
  );
}

function formatDateTime_(value) {
  if (!value) {
    return '';
  }

  return Utilities.formatDate(
    parseDate_(value),
    CONFIG.timezone,
    'yyyy-MM-dd HH:mm:ss'
  );
}

function secondsToText_(seconds) {
  seconds = Math.max(0, Math.round(Number(seconds || 0)));

  const hours = Math.floor(seconds / 3600);
  const minutes = Math.floor((seconds % 3600) / 60);
  const remaining = seconds % 60;

  return [
    String(hours).padStart(2, '0'),
    String(minutes).padStart(2, '0'),
    String(remaining).padStart(2, '0')
  ].join(':');
}

function sum_(records, field) {
  return records.reduce(
    (total, record) => total + Number(record[field] || 0),
    0
  );
}

function unique_(array) {
  return [...new Set(array)];
}

function csvEscape_(value) {
  const text = String(value ?? '');
  return '"' + text.replace(/"/g, '""') + '"';
}
