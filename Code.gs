const APP_NAME = 'TMA Cabin Crew Portal';
const SHEET_ID = '17oMCgoCVjPVekuZxErQmfxW_kiXCIuaFOBJSFlhv0nM';
const SESSION_PREFIX = 'TMA_PORTAL_SESSION_';
const SESSION_TTL = 60 * 60 * 6;


function doGet(e) {

  // If browser requests the PWA manifest
  if (e && e.parameter && e.parameter.pwa === 'manifest') {
    return ContentService
      .createTextOutput(JSON.stringify(getPwaManifest()))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Normal portal load
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function setupPortal() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  ensureSheet_(ss, 'Settings', ['Key', 'Value'], [
    ['PortalTitle', 'TMA Cabin Crew Portal'],
    ['Subtitle', 'Reports,Important Docx, Policies & Quick Links'],
    ['ContactText', 'Crew Support Office'],
    ['LogoURL', ''],
    ['ThemeColor', '#C8102E'],
    ['LastUpdated', new Date()]
  ]);

  ensureSheet_(ss, 'Announcements', ['ID', 'Title', 'Message', 'Active', 'SortOrder'], [
    ['ANN-001', 'Important Notice', 'Crew updates, schedules, reports, and quick links in one place.', true, 1]
  ]);

  ensureSheet_(ss, 'Marquee', ['ID', 'Message', 'Active', 'SortOrder'], [
    ['MQ-001', 'Welcome to TMA Cabin Crew Portal', true, 1],
    ['MQ-002', 'Please check latest reports and schedules', true, 2]
  ]);

  ensureSheet_(ss, 'TopCrew', ['Rank', 'Name', 'PhotoURL', 'Caption', 'Active'], [
    [1, 'Crew Member 01', '', 'Outstanding service', true],
    [2, 'Crew Member 02', '', 'Excellent teamwork', true],
    [3, 'Crew Member 03', '', 'Strong punctuality', true]
  ]);

  ensureSheet_(ss, 'Menu', ['MainMenu', 'SubMenu', 'LinkType', 'URL', 'SortOrder', 'Active'], [
    ['Monthly Reports 2026', 'JAN', 'report', '', 1, true],
    ['Monthly Reports 2026', 'FEB', 'report', '', 2, true],
    ['Monthly Reports 2026', 'MAR', 'report', '', 3, true],
    ['Monthly Reports 2026', 'APR', 'report', '', 4, true],
    ['Monthly Reports 2026', 'MAY', 'report', '', 5, true],
    ['Monthly Reports 2026', 'JUN', 'report', '', 6, true],
    ['Monthly Reports 2026', 'JUL', 'report', '', 7, true],
    ['Monthly Reports 2026', 'AUG', 'report', '', 8, true],
    ['Monthly Reports 2026', 'SEP', 'report', '', 9, true],
    ['Monthly Reports 2026', 'OCT', 'report', '', 10, true],
    ['Monthly Reports 2026', 'NOV', 'report', '', 11, true],
    ['Monthly Reports 2026', 'DEC', 'report', '', 12, true],
    ['Feedback Policy for Epaulet Upgrade and Downgrade', 'Upgrade Policy', 'url', '', 13, true],
    ['Feedback Policy for Epaulet Upgrade and Downgrade', 'Downgrade Policy', 'url', '', 14, true],
    ['Feedback Policy for Epaulet Upgrade and Downgrade', 'Review Period', 'url', '', 15, true],
    ['Feedback Policy for Epaulet Upgrade and Downgrade', 'Standards / Criteria', 'url', '', 16, true],
    ['Bus Schedule', 'Normal', 'bus', '', 17, true],
    ['Bus Schedule', 'Ramadhan', 'bus', '', 18, true],
    ['Important Links', 'Open Important Links', 'internal', 'important-links', 19, true],
    ['Phone Book', 'Open Phone Book', 'internal', 'phone-book', 20, true],
    ['CC of the Month', 'Open Month History', 'internal', 'cc-month-history', 21, true]
  ]);

  ensureSheet_(ss, 'Reports', ['Month', 'Title', 'URL', 'Status', 'SortOrder', 'Active'], [
    ['JAN', 'January Report', '', 'missing', 1, true],
    ['FEB', 'February Report', '', 'missing', 2, true],
    ['MAR', 'March Report', '', 'missing', 3, true],
    ['APR', 'April Report', '', 'missing', 4, true],
    ['MAY', 'May Report', '', 'missing', 5, true],
    ['JUN', 'June Report', '', 'missing', 6, true],
    ['JUL', 'July Report', '', 'missing', 7, true],
    ['AUG', 'August Report', '', 'missing', 8, true],
    ['SEP', 'September Report', '', 'missing', 9, true],
    ['OCT', 'October Report', '', 'missing', 10, true],
    ['NOV', 'November Report', '', 'missing', 11, true],
    ['DEC', 'December Report', '', 'missing', 12, true]
  ]);

  ensureSheet_(ss, 'BusSchedule', ['Type', 'Title', 'URL', 'Active'], [
    ['Normal', 'Normal Bus Schedule', '', true],
    ['Ramadhan', 'Ramadhan Bus Schedule', '', true]
  ]);

  ensureSheet_(ss, 'ImportantLinks', ['ID', 'Title', 'URL', 'SortOrder', 'Active', 'Icon'], [
    ['LINK-001', 'Crew Notice Board', '', 1, true, 'link'],
    ['LINK-002', 'Operations Manual', '', 2, true, 'book']
  ]);

  ensurePhoneSheetStructure_(ss);

  ensureSheet_(ss, 'CCMonth', ['Rank', 'Month', 'CrewCode', 'Caption', 'Active'], [
    [1, 'JAN', 'AFKA', 'Outstanding leadership', true],
    [2, 'JAN', 'HAMS', 'Great teamwork', true],
    [3, 'JAN', 'FARU', 'Excellent punctuality', true],
    [4, 'JAN', 'IBRA', 'Guest care excellence', true],
    [5, 'JAN', 'ZAIN', 'Safety focus', true]
  ]);

  ensureSheet_(ss, 'CCYear', ['Rank', 'CrewCode', 'Caption', 'Active'], [
    [1, 'AFKA', 'Crew excellence', true],
    [2, 'HAMS', 'Outstanding leadership', true],
    [3, 'FARU', 'Strong performance', true]
  ]);

  ensureSheet_(ss, 'TopFlightTime', ['Rank', 'CrewCode', 'Caption', 'Active'], [
    [1, 'FARU', 'Highest flight hours', true],
    [2, 'IBRA', 'Strong operational contribution', true],
    [3, 'AFKA', 'Excellent monthly total', true]
  ]);

  ensureSheet_(ss, 'Admin', ['Username', 'Password', 'Role', 'Active'], [
    ['admin', 'ChangeMe123!', 'superadmin', true]
  ]);

  formatAllHeaders_(ss);
  updateLastUpdated_(ss);

  return json_({
    ok: true,
    message: 'Setup complete',
    sheetId: SHEET_ID
  });
}

function getPortalBootstrap() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);

    ensurePortalSheets_(ss);

    const settings = keyValueSheetToObject_(ss.getSheetByName('Settings'));
    const announcements = rowsToObjects_(ss.getSheetByName('Announcements'));
    const menuRows = rowsToObjects_(ss.getSheetByName('Menu'));
    const reports = rowsToObjects_(ss.getSheetByName('Reports'));
    const busSchedule = rowsToObjects_(ss.getSheetByName('BusSchedule'));
    const importantLinks = rowsToObjects_(ss.getSheetByName('ImportantLinks'));
    const marqueee = rowsToObjects_(ss.getSheetByName('Marqueee'));
    const phoneMap = buildPhoneMap_(rowsToObjects_(ss.getSheetByName('phone')));

    const ccMonthRaw = rowsToObjects_(ss.getSheetByName('CCMonth'));
    const ccYearRaw = rowsToObjects_(ss.getSheetByName('CCYear'));
    const topFlightRaw = rowsToObjects_(ss.getSheetByName('TopFlightTime'));

    const ccMonthCurrentMonth = detectLatestMonthWithData_(ccMonthRaw) || 'JAN';

    const sliderSets = {
      ccMonth: buildSliderSetFromMonth_(ccMonthRaw, ccMonthCurrentMonth, phoneMap),
      ccYear: buildSliderSetSimple_(ccYearRaw, phoneMap),
      topFlight: buildSliderSetSimple_(topFlightRaw, phoneMap)
    };

    const monthHistory = buildMonthHistory_(ccMonthRaw, phoneMap);
    const groupedMenu = buildPublicMenu_(menuRows, reports, busSchedule, importantLinks, monthHistory);

    const portal = {
      settings: {
        portalTitle: String(settings.PortalTitle || APP_NAME),
        subtitle: String(settings.Subtitle || 'Reports, Schedules, Internal Docx, Policies & Quick Links'),
        contactText: String(settings.ContactText || 'Crew Support Office'),
        logoUrl: String(settings.LogoURL || ''),
        themeColor: String(settings.ThemeColor || '#C8102E'),
        lastUpdated: settings.LastUpdated ? String(settings.LastUpdated) : ''
      },
     activeAnnouncement: announcements.find(r => isTrue_(r.Active)) || null,
      marqueee: marqueee
        .filter(r => isTrue_(r.Active))
        .sort((a, b) => Number(a.SortOrder || 0) - Number(b.SortOrder || 0))
        .map(r => String(r.Message || '').trim())
        .filter(Boolean),
      sliderSets: sliderSets,
      monthHistory: monthHistory,
      currentMonthLabel: ccMonthCurrentMonth,
      menu: groupedMenu
    };

    return json_({
      ok: true,
      portal: portal
    });
  } catch (e) {
    return json_({
      ok: false,
      message: String(e)
    });
  }
}

function getPhoneDirectory() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    ensurePortalSheets_(ss);

    const sheet = ss.getSheetByName('phone');
    const rows = sheet ? rowsToObjects_(sheet) : [];

    const contacts = rows
      .filter(r => r.CrewCode || r.Name || r.Number)
      .sort((a, b) => String(a.CrewCode || '').localeCompare(String(b.CrewCode || '')))
      .map(r => ({
        crewCode: String(r.CrewCode || ''),
        name: String(r.Name || ''),
        number: String(r.Number || ''),
        photoUrl: String(normalizeImageUrl_(extractImageFormulaUrl_(r.PhotoURL || '')) || '')
      }));

    return json_({
      ok: true,
      contacts: contacts
    });
  } catch (e) {
    return json_({
      ok: false,
      contacts: [],
      message: String(e)
    });
  }
}

function loginAdmin(username, password) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    ensurePortalSheets_(ss);

    const rows = rowsToObjects_(ss.getSheetByName('Admin'));
    const user = rows.find(r =>
      String(r.Username || '').toLowerCase() === String(username || '').trim().toLowerCase() &&
      String(r.Password || '') === String(password || '') &&
      isTrue_(r.Active)
    );

    if (!user) {
      return json_({ ok: false, message: 'Invalid username or password.' });
    }

    const token = Utilities.getUuid();
    CacheService.getScriptCache().put(
      SESSION_PREFIX + token,
      JSON.stringify({
        username: String(user.Username || ''),
        role: String(user.Role || 'admin')
      }),
      SESSION_TTL
    );

    return json_({
      ok: true,
      token: token,
      user: {
        username: String(user.Username || ''),
        role: String(user.Role || 'admin')
      }
    });
  } catch (e) {
    return json_({ ok: false, message: String(e) });
  }
}

function logoutAdmin(token) {
  try {
    if (token) CacheService.getScriptCache().remove(SESSION_PREFIX + token);
    return json_({ ok: true });
  } catch (e) {
    return json_({ ok: false, message: String(e) });
  }
}

function validateAdminSession(token) {
  try {
    const session = getSession_(token);
    if (!session) return json_({ ok: false });
    return json_({ ok: true, user: session });
  } catch (e) {
    return json_({ ok: false, message: String(e) });
  }
}

function getAdminDashboard(token) {
  try {
    const session = getSession_(token);
    if (!session) return json_({ ok: false, message: 'Unauthorized' });

    const ss = SpreadsheetApp.openById(SHEET_ID);
    ensurePortalSheets_(ss);

    const data = {
      settings: keyValueSheetToObject_(ss.getSheetByName('Settings')),
      announcements: rowsToObjects_(ss.getSheetByName('Announcements')),
      marqueee: rowsToObjects_(ss.getSheetByName('Marqueee')),
      menu: rowsToObjects_(ss.getSheetByName('Menu')),
      reports: rowsToObjects_(ss.getSheetByName('Reports')),
      busSchedule: rowsToObjects_(ss.getSheetByName('BusSchedule')),
      importantLinks: rowsToObjects_(ss.getSheetByName('ImportantLinks')),
      phone: rowsToObjects_(ss.getSheetByName('phone')),
      ccMonth: rowsToObjects_(ss.getSheetByName('CCMonth')),
      ccYear: rowsToObjects_(ss.getSheetByName('CCYear')),
      topFlight: rowsToObjects_(ss.getSheetByName('TopFlightTime')),
      user: session
    };

    return json_({ ok: true, data: data });
  } catch (e) {
    return json_({ ok: false, message: String(e) });
  }
}

function saveAdminDashboard(token, payload) {
  try {
    const session = getSession_(token);
    if (!session) {
      return json_({ ok: false, message: 'Session expired. Please login again.' });
    }

    const parsed = typeof payload === 'string' ? JSON.parse(payload) : payload;
    const ss = SpreadsheetApp.openById(SHEET_ID);
    ensurePortalSheets_(ss);

    if (parsed.settings) {
      writeKeyValueSheet_(ss.getSheetByName('Settings'), parsed.settings, [
        'PortalTitle',
        'Subtitle',
        'ContactText',
        'LogoURL',
        'ThemeColor',
        'LastUpdated'
      ]);
    }

    if (parsed.announcements) {
      writeObjectsToSheet_(
        ss.getSheetByName('Announcements'),
        removeEmptyItems_(parsed.announcements, ['Title', 'Message']),
        ['ID', 'Title', 'Message', 'Active', 'SortOrder']
      );
    }

    if (parsed.marqueee) {
      writeObjectsToSheet_(
        ss.getSheetByName('Marqueee'),
        removeEmptyItems_(parsed.marqueee, ['Message']),
        ['ID', 'Message', 'Active', 'SortOrder']
      );
    }

    if (parsed.menu) {
      writeObjectsToSheet_(
        ss.getSheetByName('Menu'),
        removeEmptyItems_(parsed.menu, ['MainMenu', 'SubMenu', 'URL']),
        ['MainMenu', 'SubMenu', 'LinkType', 'URL', 'SortOrder', 'Active']
      );
    }

    if (parsed.reports) {
      writeObjectsToSheet_(
        ss.getSheetByName('Reports'),
        removeEmptyItems_(parsed.reports, ['Month', 'Title', 'URL']),
        ['Month', 'Title', 'URL', 'Status', 'SortOrder', 'Active']
      );
    }

    if (parsed.busSchedule) {
      writeObjectsToSheet_(
        ss.getSheetByName('BusSchedule'),
        removeEmptyItems_(parsed.busSchedule, ['Type', 'Title', 'URL']),
        ['Type', 'Title', 'URL', 'Active']
      );
    }

    if (parsed.importantLinks) {
      writeObjectsToSheet_(
        ss.getSheetByName('ImportantLinks'),
        removeEmptyItems_(parsed.importantLinks, ['Title', 'URL']),
        ['ID', 'Title', 'URL', 'SortOrder', 'Active', 'Icon']
      );
    }

    if (parsed.phone) {
      writeObjectsToSheet_(
        ss.getSheetByName('phone'),
        removeEmptyItems_(parsed.phone, ['CrewCode', 'Name', 'Number', 'PhotoURL']),
        ['CrewCode', 'Name', 'Number', 'PhotoURL']
      );
    }

    if (parsed.ccMonth) {
      writeObjectsToSheet_(
        ss.getSheetByName('CCMonth'),
        removeEmptyItems_(parsed.ccMonth, ['Month', 'CrewCode', 'Caption']),
        ['Rank', 'Month', 'CrewCode', 'Caption', 'Active']
      );
    }

    if (parsed.ccYear) {
      writeObjectsToSheet_(
        ss.getSheetByName('CCYear'),
        removeEmptyItems_(parsed.ccYear, ['CrewCode', 'Caption']),
        ['Rank', 'CrewCode', 'Caption', 'Active']
      );
    }

    if (parsed.topFlight) {
      writeObjectsToSheet_(
        ss.getSheetByName('TopFlightTime'),
        removeEmptyItems_(parsed.topFlight, ['CrewCode', 'Caption']),
        ['Rank', 'CrewCode', 'Caption', 'Active']
      );
    }

    formatAllHeaders_(ss);
    updateLastUpdated_(ss);

    return json_({
      ok: true,
      message: 'Changes saved successfully.'
    });
  } catch (e) {
    return json_({
      ok: false,
      message: 'Save failed: ' + String(e)
    });
  }
}

function ensurePortalSheets_(ss) {
  ensureSheet_(ss, 'Settings', ['Key', 'Value'], [
    ['PortalTitle', 'TMA Cabin Crew Portal'],
    ['Subtitle', 'Reports,Internal Docx, Policies & Quick Links'],
    ['ContactText', 'Crew Support Office'],
    ['LogoURL', ''],
    ['ThemeColor', '#C8102E'],
    ['LastUpdated', new Date()]
  ]);

  ensureSheet_(ss, 'Announcements', ['ID', 'Title', 'Message', 'Active', 'SortOrder'], []);
  ensureSheet_(ss, 'Marqueee', ['ID', 'Message', 'Active', 'SortOrder'], []);
  ensureSheet_(ss, 'TopCrew', ['Rank', 'Name', 'PhotoURL', 'Caption', 'Active'], []);
  ensureSheet_(ss, 'Menu', ['MainMenu', 'SubMenu', 'LinkType', 'URL', 'SortOrder', 'Active'], []);
  ensureSheet_(ss, 'Reports', ['Month', 'Title', 'URL', 'Status', 'SortOrder', 'Active'], []);
  ensureSheet_(ss, 'BusSchedule', ['Type', 'Title', 'URL', 'Active'], []);
  ensureSheet_(ss, 'ImportantLinks', ['ID', 'Title', 'URL', 'SortOrder', 'Active', 'Icon'], []);
  ensurePhoneSheetStructure_(ss);
  ensureSheet_(ss, 'CCMonth', ['Rank', 'Month', 'CrewCode', 'Caption', 'Active'], []);
  ensureSheet_(ss, 'CCYear', ['Rank', 'CrewCode', 'Caption', 'Active'], []);
  ensureSheet_(ss, 'TopFlightTime', ['Rank', 'CrewCode', 'Caption', 'Active'], []);
  ensureSheet_(ss, 'Admin', ['Username', 'Password', 'Role', 'Active'], [
    ['admin', 'ChangeMe123!', 'superadmin', true]
  ]);
}

function ensureSheet_(ss, name, headers, sampleRows) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    if (sampleRows && sampleRows.length) {
      sh.getRange(2, 1, sampleRows.length, headers.length).setValues(sampleRows);
    }
  } else {
    const currentHeaders = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), headers.length)).getValues()[0];
    let rewriteHeaders = false;
    for (let i = 0; i < headers.length; i++) {
      if (String(currentHeaders[i] || '') !== String(headers[i])) {
        rewriteHeaders = true;
        break;
      }
    }
    if (rewriteHeaders) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    sh.setFrozenRows(1);
  }
}

function ensurePhoneSheetStructure_(ss) {
  let sh = ss.getSheetByName('phone');
  const headers = ['CrewCode', 'Name', 'Number', 'PhotoURL'];

  if (!sh) sh = ss.insertSheet('phone');

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(2, 1, 3, headers.length).setValues([
      ['AFKA', 'Afraasheem Kamal', '+9607000001', ''],
      ['HAMS', 'Hamsa', '+9607000002', ''],
      ['FARU', 'Faru', '+9607000003', '']
    ]);
  } else {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
}

function rowsToObjects_(sheet) {
  if (!sheet) return [];

  const range = sheet.getDataRange();
  const values = range.getValues();
  const formulas = range.getFormulas();

  if (values.length < 2) return [];

  const headers = values[0];

  return values.slice(1).map(function(row, rowIndex) {
    const obj = {};
    headers.forEach(function(h, colIndex) {
      let value = row[colIndex];
      const formula = formulas[rowIndex + 1][colIndex];

      if (String(h) === 'PhotoURL' && formula) {
        value = formula;
      }

      obj[h] = value;
    });
    return obj;
  }).filter(function(obj) {
    return Object.keys(obj).some(function(key) {
      return String(obj[key] || '').trim() !== '';
    });
  });
}

function keyValueSheetToObject_(sheet) {
  const rows = rowsToObjects_(sheet);
  const out = {};
  rows.forEach(r => out[String(r.Key)] = r.Value);
  return out;
}

function writeKeyValueSheet_(sheet, obj, order) {
  const keys = order && order.length ? order : Object.keys(obj || {});
  const rows = keys.map(k => [k, k === 'LastUpdated' ? new Date() : (obj[k] !== undefined ? obj[k] : '')]);

  sheet.clearContents();
  sheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }
}

function writeObjectsToSheet_(sheet, items, headers) {
  const cleanItems = (items || []).map(function(item) {
    const out = {};
    headers.forEach(function(h) {
      out[h] = item[h] !== undefined ? item[h] : '';
    });
    return out;
  });

  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  if (cleanItems.length) {
    const rows = cleanItems.map(function(item) {
      return headers.map(function(h) {
        return normalizeCellValue_(item[h]);
      });
    });
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
}

function removeEmptyItems_(items, importantKeys) {
  importantKeys = importantKeys || [];
  return (items || []).filter(function(item) {
    return importantKeys.some(function(key) {
      return String(item[key] || '').trim() !== '';
    });
  });
}

function normalizeCellValue_(v) {
  if (v === undefined || v === null) return '';
  return v;
}

function buildPhoneMap_(rows) {
  const map = {};
  (rows || []).forEach(function(r) {
    const code = String(r.CrewCode || '').trim().toUpperCase();
    if (!code) return;
    map[code] = {
      crewCode: String(r.CrewCode || ''),
      name: String(r.Name || ''),
      number: String(r.Number || ''),
      photoUrl: String(normalizeImageUrl_(extractImageFormulaUrl_(r.PhotoURL || '')) || '')
    };
  });
  return map;
}

function buildSliderSetFromMonth_(rows, month, phoneMap) {
  return (rows || [])
    .filter(r => isTrue_(r.Active) && String(r.Month || '').toUpperCase() === String(month || '').toUpperCase())
    .sort((a, b) => Number(a.Rank || 0) - Number(b.Rank || 0))
    .map(function(r) {
      const code = String(r.CrewCode || '').trim().toUpperCase();
      const phone = phoneMap[code] || {};
      return {
        rank: Number(r.Rank || 0),
        crewCode: String(r.CrewCode || ''),
        name: phone.name || String(r.CrewCode || ''),
        caption: String(r.Caption || ''),
        photoUrl: String(phone.photoUrl || '')
      };
    });
}

function buildSliderSetSimple_(rows, phoneMap) {
  return (rows || [])
    .filter(r => isTrue_(r.Active))
    .sort((a, b) => Number(a.Rank || 0) - Number(b.Rank || 0))
    .map(function(r) {
      const code = String(r.CrewCode || '').trim().toUpperCase();
      const phone = phoneMap[code] || {};
      return {
        rank: Number(r.Rank || 0),
        crewCode: String(r.CrewCode || ''),
        name: phone.name || String(r.CrewCode || ''),
        caption: String(r.Caption || ''),
        photoUrl: String(phone.photoUrl || '')
      };
    });
}

function detectLatestMonthWithData_(rows) {
  const order = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
  const available = {};
  (rows || []).forEach(function(r) {
    if (isTrue_(r.Active) && r.Month) available[String(r.Month).toUpperCase()] = true;
  });
  for (let i = order.length - 1; i >= 0; i--) {
    if (available[order[i]]) return order[i];
  }
  return '';
}

function buildMonthHistory_(rows, phoneMap) {
  const order = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
  const grouped = {};

  (rows || []).forEach(function(r) {
    const month = String(r.Month || '').toUpperCase();
    if (!month) return;
    if (!grouped[month]) grouped[month] = [];
    grouped[month].push(r);
  });

  return order.map(function(month) {
    const items = (grouped[month] || [])
      .filter(r => isTrue_(r.Active))
      .sort((a, b) => Number(a.Rank || 0) - Number(b.Rank || 0))
      .map(function(r) {
        const code = String(r.CrewCode || '').trim().toUpperCase();
        const phone = phoneMap[code] || {};
        return {
          rank: Number(r.Rank || 0),
          crewCode: String(r.CrewCode || ''),
          name: phone.name || String(r.CrewCode || ''),
          photoUrl: String(phone.photoUrl || '')
        };
      });

    return {
      month: month,
      items: items
    };
  }).filter(function(m) {
    return m.items.length > 0;
  });
}

function buildPublicMenu_(menuRows, reports, busSchedule, importantLinks, monthHistory) {
  const grouped = {};

  menuRows
    .filter(r => isTrue_(r.Active))
    .sort((a, b) => Number(a.SortOrder || 0) - Number(b.SortOrder || 0))
    .forEach(function(r) {
      const mainMenu = String(r.MainMenu || '').trim();
      if (!mainMenu) return;
      if (!grouped[mainMenu]) grouped[mainMenu] = [];

      const linkType = String(r.LinkType || 'url');
      let finalUrl = String(r.URL || '');
      let badge = 'LINK';

      if (linkType === 'internal') badge = 'APP';
      if (linkType === 'report') {
        const match = reports.find(x => String(x.Month || '') === String(r.SubMenu || ''));
        finalUrl = match ? String(match.URL || '') : '';
        badge = 'PDF';
      }
      if (linkType === 'bus') {
        const match = busSchedule.find(x => String(x.Type || '') === String(r.SubMenu || ''));
        finalUrl = match ? String(match.URL || '') : '';
        badge = 'LINK';
      }

      if (linkType === 'internal' && finalUrl === 'important-links') {
        importantLinks
          .filter(x => isTrue_(x.Active))
          .sort((a, b) => Number(a.SortOrder || 0) - Number(b.SortOrder || 0))
          .forEach(function(link) {
            grouped[mainMenu].push({
              title: String(link.Title || ''),
              linkType: 'url',
              url: String(link.URL || ''),
              badge: 'LINK'
            });
          });
        return;
      }

      if (linkType === 'internal' && finalUrl === 'cc-month-history') {
        monthHistory.forEach(function(m) {
          grouped[mainMenu].push({
            title: m.month,
            linkType: 'month-history',
            url: m.month,
            badge: 'LIST',
            historyItems: m.items
          });
        });
        return;
      }

      grouped[mainMenu].push({
        title: String(r.SubMenu || ''),
        linkType: linkType,
        url: finalUrl,
        badge: badge
      });
    });

  return Object.keys(grouped).map(function(title) {
    return {
      title: title,
      items: grouped[title]
    };
  });
}

function updateLastUpdated_(ss) {
  const sheet = ss.getSheetByName('Settings');
  if (!sheet) return;

  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === 'LastUpdated') {
      sheet.getRange(i + 1, 2).setValue(new Date());
      return;
    }
  }

  const lastRow = sheet.getLastRow() + 1;
  sheet.getRange(lastRow, 1, 1, 2).setValues([['LastUpdated', new Date()]]);
}

function formatAllHeaders_(ss) {
  ss.getSheets().forEach(function(sh) {
    const lastCol = sh.getLastColumn();
    if (lastCol > 0) {
      sh.getRange(1, 1, 1, lastCol)
        .setFontWeight('bold')
        .setBackground('#C8102E')
        .setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }
  });
}

function extractImageFormulaUrl_(value) {
  value = String(value || '').trim();
  if (!value) return '';

  const match = value.match(/^=IMAGE\("([^"]+)"/i);
  if (match && match[1]) return match[1];

  return value;
}

function normalizeImageUrl_(url) {
  url = String(url || '').trim();
  if (!url) return '';

  let match;

  if (url.indexOf('thumbnail?id=') !== -1) return url;

  match = url.match(/[?&]id=([-\w]{25,})/);
  if (match && match[1]) return 'https://drive.google.com/thumbnail?id=' + match[1] + '&sz=w400';

  match = url.match(/\/d\/([-\w]{25,})/);
  if (match && match[1]) return 'https://drive.google.com/thumbnail?id=' + match[1] + '&sz=w400';

  match = url.match(/[-\w]{25,}/);
  if (url.indexOf('drive.google.com') !== -1 && match) {
    return 'https://drive.google.com/thumbnail?id=' + match[0] + '&sz=w400';
  }

  return url;
}

function isTrue_(v) {
  return v === true || String(v).toLowerCase() === 'true' || String(v) === '1' || String(v).toLowerCase() === 'yes';
}

function getSession_(token) {
  if (!token) return null;
  const raw = CacheService.getScriptCache().get(SESSION_PREFIX + token);
  if (!raw) return null;
  return JSON.parse(raw);
}

function json_(obj) {
  return JSON.stringify(obj);
}