const SHEET_NAMES = {
  USERS: 'Users',
  LEADS: 'Leads',
  PROJECTS: 'Projects',
  TRANSACTIONS: 'Transactions',
  ACTIVITY_LOGS: 'Activity_Logs',
  IT_TICKETS: 'IT_Tickets'
};

const LEAD_HEADERS = {
  LEAD_ID: 'LeadID',
  AGENT_USERNAME: 'AgentUsername',
  CLIENT_NAME: 'ClientName',
  PHONE: 'Phone',
  COUNTRY: 'Country',
  CLASSIFICATION: 'Classification (OFW, Locally Employed, Self-Employed)',
  SOURCE: 'Source (TikTok Ads, FB Ads, Organic, Referral, KKK)',
  TEMPERATURE: 'Temperature (Hot, Warm, Cold)',
  PROJECT: 'Project',
  TOTAL_SELLING_PRICE: 'TotalSellingPrice',
  STATUS: 'Status',
  REMARKS: 'Remarks',
  DATE_ADDED: 'DateAdded'
};

const USER_HEADERS = {
  ID: 'ID',
  FULL_NAME: 'FullName',
  USERNAME: 'Username',
  PASSWORD: 'Password',
  TEAM: 'Team',
  ROLE: 'Role'
};

const PROJECT_HEADERS = {
  PROJECT_NAME: 'ProjectName',
  COMMISSION: 'Commission',
  DRIVE_LINK: 'DriveLink'
};

const TRANSACTION_HEADERS = {
  TRANSACTION_ID: 'TransactionID',
  LEAD_ID: 'LeadID',
  PAYMENT_TYPE: 'PaymentType (Spot / Installment)',
  TERM_MONTHS: 'TermMonths',
  CURRENT_MONTH_PAID: 'CurrentMonthPaid',
  RECEIPT_DRIVE_URL: 'ReceiptDriveURL',
  RESERVATION_DATE: 'ReservationDate'
};

const LOG_HEADERS = {
  LOG_ID: 'LogID',
  LEAD_ID: 'LeadID',
  AGENT_USERNAME: 'AgentUsername',
  TIMESTAMP: 'Timestamp',
  ACTION_TEXT: 'ActionText'
};

const TICKET_HEADERS = {
  TICKET_ID: 'TicketID',
  USERNAME: 'Username',
  ISSUE_TYPE: 'IssueType',
  PRIORITY_LEVEL: 'PriorityLevel',
  DESCRIPTION: 'Description',
  STATUS: 'Status'
};

const PIPELINE_STATUSES = [
  'Inquiry',
  'Site Tour',
  'Down Payment',
  'Approval',
  'Takeout',
  'Turnover',
  'Closed/Lost'
];

const CLASSIFICATION_OPTIONS = ['OFW', 'Locally Employed', 'Self-Employed', 'Unknown'];
const SOURCE_OPTIONS = ['TikTok Ads', 'FB Ads', 'Organic', 'Referral', 'KKK'];
const TEMPERATURE_OPTIONS = ['Hot', 'Warm', 'Cold'];
const PAYMENT_TYPE_OPTIONS = ['Spot', 'Installment'];

const SESSION_CACHE_PREFIX = 'lms_session:';
const SESSION_TTL_SECONDS = 6 * 60 * 60;
const VOICE_RATE_PREFIX = 'voice_rate:';

const REQUIRED_HEADERS_BY_SHEET = {};
REQUIRED_HEADERS_BY_SHEET[SHEET_NAMES.USERS] = [
  USER_HEADERS.ID,
  USER_HEADERS.FULL_NAME,
  USER_HEADERS.USERNAME,
  USER_HEADERS.PASSWORD,
  USER_HEADERS.TEAM,
  USER_HEADERS.ROLE
];
REQUIRED_HEADERS_BY_SHEET[SHEET_NAMES.LEADS] = [
  LEAD_HEADERS.LEAD_ID,
  LEAD_HEADERS.AGENT_USERNAME,
  LEAD_HEADERS.CLIENT_NAME,
  LEAD_HEADERS.PHONE,
  LEAD_HEADERS.COUNTRY,
  LEAD_HEADERS.CLASSIFICATION,
  LEAD_HEADERS.SOURCE,
  LEAD_HEADERS.TEMPERATURE,
  LEAD_HEADERS.PROJECT,
  LEAD_HEADERS.TOTAL_SELLING_PRICE,
  LEAD_HEADERS.STATUS,
  LEAD_HEADERS.REMARKS,
  LEAD_HEADERS.DATE_ADDED
];
REQUIRED_HEADERS_BY_SHEET[SHEET_NAMES.PROJECTS] = [
  PROJECT_HEADERS.PROJECT_NAME,
  PROJECT_HEADERS.COMMISSION,
  PROJECT_HEADERS.DRIVE_LINK
];
REQUIRED_HEADERS_BY_SHEET[SHEET_NAMES.TRANSACTIONS] = [
  TRANSACTION_HEADERS.TRANSACTION_ID,
  TRANSACTION_HEADERS.LEAD_ID,
  TRANSACTION_HEADERS.PAYMENT_TYPE,
  TRANSACTION_HEADERS.TERM_MONTHS,
  TRANSACTION_HEADERS.CURRENT_MONTH_PAID,
  TRANSACTION_HEADERS.RECEIPT_DRIVE_URL,
  TRANSACTION_HEADERS.RESERVATION_DATE
];
REQUIRED_HEADERS_BY_SHEET[SHEET_NAMES.ACTIVITY_LOGS] = [
  LOG_HEADERS.LOG_ID,
  LOG_HEADERS.LEAD_ID,
  LOG_HEADERS.AGENT_USERNAME,
  LOG_HEADERS.TIMESTAMP,
  LOG_HEADERS.ACTION_TEXT
];
REQUIRED_HEADERS_BY_SHEET[SHEET_NAMES.IT_TICKETS] = [
  TICKET_HEADERS.TICKET_ID,
  TICKET_HEADERS.USERNAME,
  TICKET_HEADERS.ISSUE_TYPE,
  TICKET_HEADERS.PRIORITY_LEVEL,
  TICKET_HEADERS.DESCRIPTION,
  TICKET_HEADERS.STATUS
];

function doGet(e) {
  if (e && e.parameter && e.parameter.health === '1') {
    return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
  }

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Inner SPARC Lead Management')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function initialDatabaseSetup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNames = Object.keys(REQUIRED_HEADERS_BY_SHEET);
  var report = {
    ok: true,
    createdSheets: [],
    preparedSheets: []
  };

  for (var i = 0; i < sheetNames.length; i++) {
    var sheetName = sheetNames[i];
    var headers = REQUIRED_HEADERS_BY_SHEET[sheetName];
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      report.createdSheets.push(sheetName);
    }

    if (headers.length) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    sheet.setFrozenRows(1);
    report.preparedSheets.push(sheetName);
  }

  return report;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function authenticateUser(username, password) {
  username = sanitizeText_(username, 60);
  password = sanitizeText_(password, 120);

  if (!username || !password) {
    throw new Error('Username and password are required.');
  }

  var usersData = getSheetRows_(SHEET_NAMES.USERS);
  var userRow = null;

  for (var i = 0; i < usersData.rows.length; i++) {
    var row = usersData.rows[i];
    if (
      normalizeUsername_(row[USER_HEADERS.USERNAME]) === normalizeUsername_(username) &&
      String(row[USER_HEADERS.PASSWORD] || '') === password
    ) {
      userRow = row;
      break;
    }
  }

  if (!userRow) {
    throw new Error('Invalid credentials.');
  }

  var session = {
    username: sanitizeText_(userRow[USER_HEADERS.USERNAME], 60),
    fullName: sanitizeText_(userRow[USER_HEADERS.FULL_NAME], 80),
    role: sanitizeText_(userRow[USER_HEADERS.ROLE], 50),
    team: sanitizeText_(userRow[USER_HEADERS.TEAM], 50),
    userId: sanitizeText_(userRow[USER_HEADERS.ID], 50),
    issuedAt: Date.now()
  };

  var token = Utilities.getUuid();
  CacheService.getScriptCache().put(SESSION_CACHE_PREFIX + token, JSON.stringify(session), SESSION_TTL_SECONDS);

  return {
    token: token,
    user: {
      username: session.username,
      fullName: session.fullName,
      role: session.role,
      team: session.team
    }
  };
}

function logout(token) {
  token = sanitizeText_(token, 120);
  if (token) {
    CacheService.getScriptCache().remove(SESSION_CACHE_PREFIX + token);
  }
  return { ok: true };
}

function getAgentWorkspace(token) {
  var session = getSessionOrThrow_(token);
  var leads = getLeadsForAgent_(session.username);

  return {
    user: {
      username: session.username,
      fullName: session.fullName,
      role: session.role,
      team: session.team
    },
    statuses: PIPELINE_STATUSES.slice(),
    options: {
      classifications: CLASSIFICATION_OPTIONS.slice(),
      sources: SOURCE_OPTIONS.slice(),
      temperatures: TEMPERATURE_OPTIONS.slice(),
      paymentTypes: PAYMENT_TYPE_OPTIONS.slice()
    },
    projects: getProjectList_(),
    leadsByStatus: groupLeadsByStatus_(leads)
  };
}

function getLeadDetails(token, leadId) {
  var session = getSessionOrThrow_(token);
  leadId = sanitizeText_(leadId, 80);

  if (!leadId) {
    throw new Error('Lead ID is required.');
  }

  var leadRow = getLeadRowForAgent_(session.username, leadId);
  if (!leadRow) {
    throw new Error('Lead not found.');
  }

  return {
    lead: mapLeadRow_(leadRow),
    activities: getActivitiesForLead_(leadId),
    transaction: getLatestTransactionByLead_(leadId)
  };
}

function saveLead(token, payload) {
  var session = getSessionOrThrow_(token);
  payload = payload || {};

  var clientName = sanitizeText_(payload.clientName || payload.client_name, 80);
  var phone = sanitizePhone_(payload.phone || payload.phone_number);
  var country = sanitizeText_(payload.country, 60) || 'Philippines';
  var classification = normalizeClassification_(payload.classification);
  var source = normalizeSource_(payload.source);
  var temperature = normalizeTemperature_(payload.temperature);
  var project = sanitizeText_(payload.project || payload.project_interest, 120);
  var totalSellingPrice = sanitizeText_(payload.totalSellingPrice || payload.total_selling_price, 40);
  var remarks = sanitizeText_(payload.remarks, 1000);
  var status = normalizeStatus_(payload.status || 'Inquiry');
  var entrySource = sanitizeText_(payload.entrySource, 30) || 'Manual';

  if (!clientName) {
    throw new Error('Client name is required.');
  }

  var leadId = generateId_('LEAD');
  getSheetOrThrow_(SHEET_NAMES.LEADS).appendRow([
    leadId,
    session.username,
    clientName,
    phone,
    country,
    classification === 'Unknown' ? 'Locally Employed' : classification,
    source,
    temperature,
    project,
    totalSellingPrice,
    status,
    remarks,
    new Date()
  ]);

  appendActivity_(leadId, session.username, 'Lead created via ' + entrySource + ' entry.');

  return {
    ok: true,
    leadId: leadId
  };
}

function updateLeadStatus(token, leadId, nextStatus) {
  var session = getSessionOrThrow_(token);
  leadId = sanitizeText_(leadId, 80);
  nextStatus = normalizeStatus_(nextStatus);

  var leadData = getSheetRows_(SHEET_NAMES.LEADS);
  var rowInfo = findLeadRowInfo_(leadData, session.username, leadId);

  if (!rowInfo) {
    throw new Error('Lead not found for this account.');
  }

  leadData.sheet.getRange(rowInfo.rowNumber, 11).setValue(nextStatus);

  appendActivity_(leadId, session.username, 'Lead status updated to "' + nextStatus + '".');

  return { ok: true };
}

function addActivityLog(token, leadId, actionText) {
  var session = getSessionOrThrow_(token);
  leadId = sanitizeText_(leadId, 80);
  actionText = sanitizeText_(actionText, 500);

  if (!leadId || !actionText) {
    throw new Error('Lead ID and activity text are required.');
  }

  var leadRow = getLeadRowForAgent_(session.username, leadId);
  if (!leadRow) {
    throw new Error('Lead not found for this account.');
  }

  appendActivity_(leadId, session.username, actionText);
  return {
    ok: true,
    activities: getActivitiesForLead_(leadId)
  };
}

function getDPTransaction(token, leadId) {
  var session = getSessionOrThrow_(token);
  leadId = sanitizeText_(leadId, 80);

  if (!getLeadRowForAgent_(session.username, leadId)) {
    throw new Error('Lead not found for this account.');
  }

  return {
    transaction: getLatestTransactionByLead_(leadId)
  };
}

function saveDPTransaction(token, payload) {
  var session = getSessionOrThrow_(token);
  payload = payload || {};

  var leadId = sanitizeText_(payload.leadId, 80);
  if (!leadId) {
    throw new Error('Lead ID is required.');
  }

  if (!getLeadRowForAgent_(session.username, leadId)) {
    throw new Error('Lead not found for this account.');
  }

  var paymentType = normalizePaymentType_(payload.paymentType);
  var termMonths = sanitizeInt_(payload.termMonths, 1, 360, paymentType === 'Spot' ? 1 : 12);
  var currentMonthPaid = sanitizeInt_(payload.currentMonthPaid, 0, termMonths, paymentType === 'Spot' ? 1 : 0);
  var reservationDate = parseDateOrToday_(payload.reservationDate);
  var receiptDriveURL = sanitizeUrl_(payload.receiptDriveURL, 500);

  var txData = getSheetRows_(SHEET_NAMES.TRANSACTIONS);
  var txRow = findLatestTransactionInfo_(txData.rows, leadId);

  if (txRow) {
    txData.sheet.getRange(txRow.rowNumber, 3).setValue(paymentType);
    txData.sheet.getRange(txRow.rowNumber, 4).setValue(termMonths);
    txData.sheet.getRange(txRow.rowNumber, 5).setValue(currentMonthPaid);
    txData.sheet.getRange(txRow.rowNumber, 6).setValue(
      receiptDriveURL || txRow.row[TRANSACTION_HEADERS.RECEIPT_DRIVE_URL]
    );
    txData.sheet.getRange(txRow.rowNumber, 7).setValue(reservationDate);
  } else {
    getSheetOrThrow_(SHEET_NAMES.TRANSACTIONS).appendRow([
      generateId_('TX'),
      leadId,
      paymentType,
      termMonths,
      currentMonthPaid,
      receiptDriveURL,
      reservationDate
    ]);
  }

  appendActivity_(
    leadId,
    session.username,
    'Down payment tracker updated (' + currentMonthPaid + '/' + termMonths + ' months paid, ' + paymentType + ').'
  );

  return {
    ok: true,
    transaction: getLatestTransactionByLead_(leadId)
  };
}

function uploadReceipt(base64Data, fileName, mimeType, leadId, token) {
  var session = getSessionOrThrow_(token);
  leadId = sanitizeText_(leadId, 80);

  if (!leadId) {
    throw new Error('Lead ID is required.');
  }

  if (!getLeadRowForAgent_(session.username, leadId)) {
    throw new Error('Lead not found for this account.');
  }

  var safeFileName = sanitizeText_(fileName, 120) || ('receipt-' + Date.now() + '.jpg');
  var safeMimeType = sanitizeText_(mimeType, 80) || 'application/octet-stream';
  var cleanBase64 = sanitizeBase64_(base64Data);

  if (!cleanBase64) {
    throw new Error('Receipt file payload is empty.');
  }

  var bytes = Utilities.base64Decode(cleanBase64);
  var blob = Utilities.newBlob(bytes, safeMimeType, safeFileName);

  var folderId = PropertiesService.getScriptProperties().getProperty('RECEIPT_FOLDER_ID');
  var folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  var file = folder.createFile(blob);

  var url = file.getUrl();
  var txData = getSheetRows_(SHEET_NAMES.TRANSACTIONS);
  var txRow = findLatestTransactionInfo_(txData.rows, leadId);

  if (txRow) {
    txData.sheet.getRange(txRow.rowNumber, 6).setValue(url);
  } else {
    getSheetOrThrow_(SHEET_NAMES.TRANSACTIONS).appendRow([
      generateId_('TX'),
      leadId,
      'Spot',
      1,
      0,
      url,
      new Date()
    ]);
  }

  appendActivity_(leadId, session.username, 'Uploaded down payment receipt: ' + safeFileName + '.');

  return {
    ok: true,
    url: url,
    fileName: safeFileName,
    fileId: file.getId()
  };
}

function processVoiceTranscript(token, transcript) {
  var session = getSessionOrThrow_(token);
  transcript = sanitizeText_(transcript, 6000);

  if (!transcript) {
    throw new Error('Transcript is empty.');
  }

  enforceVoiceRateLimit_(session.username);

  var endpoint = PropertiesService.getScriptProperties().getProperty('GEMMA_JSON_ENDPOINT');
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMMA_API_KEY');

  var parsed;
  if (endpoint) {
    parsed = callGemmaJsonEndpoint_(endpoint, apiKey, transcript);
  } else {
    parsed = fallbackVoiceExtraction_(transcript);
  }

  return normalizeVoicePayload_(parsed, transcript);
}

function getProfile(token) {
  var session = getSessionOrThrow_(token);
  var usersData = getSheetRows_(SHEET_NAMES.USERS);

  for (var i = 0; i < usersData.rows.length; i++) {
    var row = usersData.rows[i];
    if (normalizeUsername_(row[USER_HEADERS.USERNAME]) === normalizeUsername_(session.username)) {
      return {
        fullName: sanitizeText_(row[USER_HEADERS.FULL_NAME], 80),
        username: sanitizeText_(row[USER_HEADERS.USERNAME], 60),
        team: sanitizeText_(row[USER_HEADERS.TEAM], 50),
        role: sanitizeText_(row[USER_HEADERS.ROLE], 50)
      };
    }
  }

  throw new Error('User profile not found.');
}

function changePassword(token, currentPassword, newPassword) {
  var session = getSessionOrThrow_(token);
  currentPassword = sanitizeText_(currentPassword, 120);
  newPassword = sanitizeText_(newPassword, 120);

  if (!currentPassword || !newPassword) {
    throw new Error('Current and new passwords are required.');
  }

  if (newPassword.length < 6) {
    throw new Error('New password must be at least 6 characters.');
  }

  var usersData = getSheetRows_(SHEET_NAMES.USERS);

  for (var i = 0; i < usersData.rows.length; i++) {
    var row = usersData.rows[i];
    if (normalizeUsername_(row[USER_HEADERS.USERNAME]) !== normalizeUsername_(session.username)) {
      continue;
    }

    if (String(row[USER_HEADERS.PASSWORD] || '') !== currentPassword) {
      throw new Error('Current password is incorrect.');
    }

    usersData.sheet.getRange(row.__rowNumber, 4).setValue(newPassword);
    return { ok: true };
  }

  throw new Error('User not found.');
}

function submitTicket(token, issueType, priorityLevel, description) {
  var session = getSessionOrThrow_(token);
  issueType = sanitizeText_(issueType, 60);
  priorityLevel = sanitizeText_(priorityLevel, 30);
  description = sanitizeText_(description, 1200);

  if (!issueType || !priorityLevel || !description) {
    throw new Error('Issue type, priority, and description are required.');
  }

  getSheetOrThrow_(SHEET_NAMES.IT_TICKETS).appendRow([
    generateId_('TKT'),
    session.username,
    issueType,
    priorityLevel,
    description,
    'Open'
  ]);

  return { ok: true };
}

function getSessionOrThrow_(token) {
  token = sanitizeText_(token, 120);
  if (!token) {
    throw new Error('Unauthorized request.');
  }

  var raw = CacheService.getScriptCache().get(SESSION_CACHE_PREFIX + token);
  if (!raw) {
    throw new Error('Session expired. Please sign in again.');
  }

  var session = JSON.parse(raw);
  if (!session || !session.username) {
    throw new Error('Invalid session.');
  }

  return session;
}

function getLeadsForAgent_(username) {
  var leadsData = getSheetRows_(SHEET_NAMES.LEADS);
  var out = [];

  for (var i = 0; i < leadsData.rows.length; i++) {
    var row = leadsData.rows[i];
    if (normalizeUsername_(row[LEAD_HEADERS.AGENT_USERNAME]) !== normalizeUsername_(username)) {
      continue;
    }

    out.push(mapLeadRow_(row));
  }

  out.sort(function(a, b) {
    return new Date(b.dateAdded).getTime() - new Date(a.dateAdded).getTime();
  });

  return out;
}

function groupLeadsByStatus_(leads) {
  var grouped = {};

  for (var i = 0; i < PIPELINE_STATUSES.length; i++) {
    grouped[PIPELINE_STATUSES[i]] = [];
  }

  for (var j = 0; j < leads.length; j++) {
    var lead = leads[j];
    var status = normalizeStatus_(lead.status || 'Inquiry');
    if (!grouped[status]) {
      grouped[status] = [];
    }
    grouped[status].push(lead);
  }

  return grouped;
}

function getLeadRowForAgent_(username, leadId) {
  var leadsData = getSheetRows_(SHEET_NAMES.LEADS);
  var info = findLeadRowInfo_(leadsData, username, leadId);
  return info ? info.row : null;
}

function findLeadRowInfo_(leadsData, username, leadId) {
  for (var i = 0; i < leadsData.rows.length; i++) {
    var row = leadsData.rows[i];
    if (String(row[LEAD_HEADERS.LEAD_ID] || '') !== leadId) {
      continue;
    }

    if (normalizeUsername_(row[LEAD_HEADERS.AGENT_USERNAME]) !== normalizeUsername_(username)) {
      return null;
    }

    return {
      row: row,
      rowNumber: row.__rowNumber
    };
  }

  return null;
}

function mapLeadRow_(row) {
  return {
    leadId: sanitizeText_(row[LEAD_HEADERS.LEAD_ID], 80),
    agentUsername: sanitizeText_(row[LEAD_HEADERS.AGENT_USERNAME], 60),
    clientName: sanitizeText_(row[LEAD_HEADERS.CLIENT_NAME], 80),
    phone: sanitizeText_(row[LEAD_HEADERS.PHONE], 40),
    country: sanitizeText_(row[LEAD_HEADERS.COUNTRY], 60),
    classification: sanitizeText_(row[LEAD_HEADERS.CLASSIFICATION], 40),
    source: sanitizeText_(row[LEAD_HEADERS.SOURCE], 60),
    temperature: normalizeTemperature_(row[LEAD_HEADERS.TEMPERATURE]),
    project: sanitizeText_(row[LEAD_HEADERS.PROJECT], 120),
    totalSellingPrice: sanitizeText_(row[LEAD_HEADERS.TOTAL_SELLING_PRICE], 40),
    status: normalizeStatus_(row[LEAD_HEADERS.STATUS] || 'Inquiry'),
    remarks: sanitizeText_(row[LEAD_HEADERS.REMARKS], 1000),
    dateAdded: toIsoDateString_(row[LEAD_HEADERS.DATE_ADDED])
  };
}

function appendActivity_(leadId, username, actionText) {
  getSheetOrThrow_(SHEET_NAMES.ACTIVITY_LOGS).appendRow([
    generateId_('LOG'),
    leadId,
    username,
    new Date(),
    sanitizeText_(actionText, 500)
  ]);
}

function getActivitiesForLead_(leadId) {
  var activityData = getSheetRows_(SHEET_NAMES.ACTIVITY_LOGS);
  var out = [];

  for (var i = 0; i < activityData.rows.length; i++) {
    var row = activityData.rows[i];
    if (String(row[LOG_HEADERS.LEAD_ID] || '') !== String(leadId)) {
      continue;
    }

    out.push({
      logId: sanitizeText_(row[LOG_HEADERS.LOG_ID], 80),
      leadId: sanitizeText_(row[LOG_HEADERS.LEAD_ID], 80),
      agentUsername: sanitizeText_(row[LOG_HEADERS.AGENT_USERNAME], 60),
      timestamp: toIsoDateString_(row[LOG_HEADERS.TIMESTAMP]),
      actionText: sanitizeText_(row[LOG_HEADERS.ACTION_TEXT], 500)
    });
  }

  out.sort(function(a, b) {
    return new Date(a.timestamp).getTime() - new Date(b.timestamp).getTime();
  });

  return out;
}

function getLatestTransactionByLead_(leadId) {
  var txData = getSheetRows_(SHEET_NAMES.TRANSACTIONS);
  var info = findLatestTransactionInfo_(txData.rows, leadId);

  if (!info) {
    return null;
  }

  var row = info.row;
  var termMonths = sanitizeInt_(row[TRANSACTION_HEADERS.TERM_MONTHS], 1, 360, 1);
  var currentMonthPaid = sanitizeInt_(row[TRANSACTION_HEADERS.CURRENT_MONTH_PAID], 0, termMonths, 0);
  var progressPercent = termMonths > 0 ? Math.round((currentMonthPaid / termMonths) * 100) : 0;

  return {
    transactionId: sanitizeText_(row[TRANSACTION_HEADERS.TRANSACTION_ID], 80),
    leadId: sanitizeText_(row[TRANSACTION_HEADERS.LEAD_ID], 80),
    paymentType: normalizePaymentType_(row[TRANSACTION_HEADERS.PAYMENT_TYPE]),
    termMonths: termMonths,
    currentMonthPaid: currentMonthPaid,
    receiptDriveURL: sanitizeUrl_(row[TRANSACTION_HEADERS.RECEIPT_DRIVE_URL], 500),
    reservationDate: toIsoDateString_(row[TRANSACTION_HEADERS.RESERVATION_DATE]),
    progressPercent: progressPercent
  };
}

function findLatestTransactionInfo_(rows, leadId) {
  var latest = null;

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (String(row[TRANSACTION_HEADERS.LEAD_ID] || '') !== String(leadId)) {
      continue;
    }

    var stamp = new Date(row[TRANSACTION_HEADERS.RESERVATION_DATE] || 0).getTime();
    if (!latest || stamp >= latest.stamp) {
      latest = {
        row: row,
        rowNumber: row.__rowNumber,
        stamp: stamp
      };
    }
  }

  return latest;
}

function getProjectList_() {
  var projectsData = getSheetRows_(SHEET_NAMES.PROJECTS);
  var out = [];

  for (var i = 0; i < projectsData.rows.length; i++) {
    var row = projectsData.rows[i];
    out.push({
      projectName: sanitizeText_(row[PROJECT_HEADERS.PROJECT_NAME], 120),
      commission: sanitizeText_(row[PROJECT_HEADERS.COMMISSION], 60),
      driveLink: sanitizeUrl_(row[PROJECT_HEADERS.DRIVE_LINK], 500)
    });
  }

  return out;
}

function enforceVoiceRateLimit_(username) {
  var key = VOICE_RATE_PREFIX + normalizeUsername_(username);
  var cache = CacheService.getScriptCache();
  var now = Date.now();
  var history = [];

  var raw = cache.get(key);
  if (raw) {
    try {
      history = JSON.parse(raw);
    } catch (err) {
      history = [];
    }
  }

  var valid = [];
  for (var i = 0; i < history.length; i++) {
    if (now - Number(history[i]) <= 60000) {
      valid.push(Number(history[i]));
    }
  }

  if (valid.length >= 5) {
    throw new Error('Voice request limit reached. Please wait a minute and try again.');
  }

  valid.push(now);
  cache.put(key, JSON.stringify(valid), 120);
}

function callGemmaJsonEndpoint_(endpoint, apiKey, transcript) {
  var projectCatalog = getProjectCatalogForPrompt_();
  var systemPrompt = [
    'You are a real estate data extractor. Parse the agent transcript.',
    'Default country to Philippines if unstated.',
    'Map project to known catalog: ' + projectCatalog + '.',
    'Infer temperature when unstated (Hot=ready, Cold=just asking).',
    'Output strictly as JSON with the requested fields.'
  ].join(' ');

  var schema = {
    type: 'object',
    properties: {
      client_name: { type: 'string' },
      phone_number: { type: 'string' },
      country: { type: 'string' },
      classification: {
        type: 'string',
        enum: ['OFW', 'Locally Employed', 'Self-Employed', 'Unknown']
      },
      source: { type: 'string' },
      temperature: {
        type: 'string',
        enum: ['Hot', 'Warm', 'Cold']
      },
      project_interest: { type: 'string' },
      total_selling_price: { type: 'string' },
      remarks: { type: 'string' }
    },
    required: ['client_name', 'temperature']
  };

  var payload = {
    model: 'gemma-4',
    prompt: {
      system: systemPrompt,
      user: transcript
    },
    response_schema: schema
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    payload: JSON.stringify(payload),
    headers: {}
  };

  if (apiKey) {
    options.headers.Authorization = 'Bearer ' + apiKey;
  }

  var response = UrlFetchApp.fetch(endpoint, options);
  var code = response.getResponseCode();
  var text = response.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error('AI endpoint error (' + code + ').');
  }

  var parsed = JSON.parse(text || '{}');
  var candidate = parsed;

  if (parsed.output) {
    candidate = parsed.output;
  }
  if (parsed.data) {
    candidate = parsed.data;
  }
  if (parsed.json) {
    candidate = parsed.json;
  }

  if (typeof candidate === 'string') {
    candidate = JSON.parse(candidate);
  }

  return candidate;
}

function fallbackVoiceExtraction_(transcript) {
  var text = String(transcript || '');
  var lower = text.toLowerCase();
  var clientName = extractClientNameFromTranscript_(text) || 'Unknown Client';
  var phone = extractPhoneFromTranscript_(text);
  var project = extractProjectFromTranscript_(text);

  var temp = 'Warm';
  if (lower.indexOf('ready') >= 0 || lower.indexOf('urgent') >= 0 || lower.indexOf('now') >= 0) {
    temp = 'Hot';
  } else if (lower.indexOf('ask') >= 0 || lower.indexOf('inquire') >= 0 || lower.indexOf('just checking') >= 0) {
    temp = 'Cold';
  }

  return {
    client_name: clientName,
    phone_number: phone,
    country: 'Philippines',
    classification: 'Unknown',
    source: 'Organic',
    temperature: temp,
    project_interest: project,
    total_selling_price: '',
    remarks: text
  };
}

function getProjectCatalogForPrompt_() {
  var projects = getProjectList_();
  var names = [];

  for (var i = 0; i < projects.length; i++) {
    if (projects[i] && projects[i].projectName) {
      names.push(projects[i].projectName);
    }
  }

  if (!names.length) {
    return 'Unknown';
  }

  return names.join(', ');
}

function extractClientNameFromTranscript_(text) {
  var directMatch = text.match(/(?:client|name)\s*(?:is|:)\s*([a-zA-Z][a-zA-Z\s\.-]{1,50})/i);
  if (directMatch && directMatch[1]) {
    return sanitizeText_(directMatch[1], 80);
  }

  var possible = text.match(/\b([A-Z][a-z]+\s[A-Z][a-z]+)\b/);
  if (possible && possible[1]) {
    return sanitizeText_(possible[1], 80);
  }

  return '';
}

function extractPhoneFromTranscript_(text) {
  var match = text.match(/(\+?\d[\d\s-]{7,})/);
  if (!match || !match[1]) {
    return '';
  }
  return sanitizePhone_(match[1]);
}

function extractProjectFromTranscript_(text) {
  var projects = getProjectList_();
  var lowerText = String(text || '').toLowerCase();

  for (var i = 0; i < projects.length; i++) {
    var name = String(projects[i].projectName || '').trim();
    if (!name) {
      continue;
    }

    if (lowerText.indexOf(name.toLowerCase()) >= 0) {
      return name;
    }
  }

  return '';
}

function normalizeVoicePayload_(payload, transcript) {
  payload = payload || {};

  var normalized = {
    client_name: sanitizeText_(payload.client_name || payload.clientName, 80),
    phone_number: sanitizePhone_(payload.phone_number || payload.phone),
    country: sanitizeText_(payload.country, 60) || 'Philippines',
    classification: normalizeClassification_(payload.classification),
    source: normalizeSource_(payload.source),
    temperature: normalizeTemperature_(payload.temperature),
    project_interest: sanitizeText_(payload.project_interest || payload.project, 120),
    total_selling_price: sanitizeText_(payload.total_selling_price, 40),
    remarks: sanitizeText_(payload.remarks || transcript, 1200)
  };

  if (!normalized.client_name) {
    throw new Error('AI output is missing client_name. Please retry with a clearer voice note.');
  }

  return normalized;
}

function getSheetOrThrow_(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Sheet "' + sheetName + '" not found. Run initialDatabaseSetup() first.');
  }
  return sheet;
}

function getSheetRows_(sheetName) {
  var sheet = getSheetOrThrow_(sheetName);

  var values = sheet.getDataRange().getValues();
  if (!values || values.length === 0) {
    return {
      sheet: sheet,
      headers: [],
      rows: []
    };
  }

  var headers = [];
  for (var h = 0; h < values[0].length; h++) {
    headers.push(String(values[0][h] || '').trim());
  }

  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var rowValues = values[i];
    if (isBlankRow_(rowValues)) {
      continue;
    }

    var rowObj = { __rowNumber: i + 1 };
    for (var c = 0; c < headers.length; c++) {
      rowObj[headers[c]] = rowValues[c];
    }
    rows.push(rowObj);
  }

  return {
    sheet: sheet,
    headers: headers,
    rows: rows
  };
}

function sanitizeText_(value, maxLen) {
  if (value === null || value === undefined) {
    return '';
  }

  var out = String(value)
    .replace(/[<>]/g, '')
    .replace(/[\u0000-\u001f\u007f]/g, ' ')
    .trim();

  if (maxLen && out.length > maxLen) {
    out = out.substring(0, maxLen);
  }

  return out;
}

function sanitizePhone_(phone) {
  var clean = sanitizeText_(phone, 40);
  if (!clean) {
    return '';
  }
  return clean.replace(/[^0-9+]/g, '');
}

function sanitizeUrl_(url, maxLen) {
  var clean = sanitizeText_(url, maxLen || 500);
  if (!clean) {
    return '';
  }
  if (!/^https?:\/\//i.test(clean)) {
    return '';
  }
  return clean;
}

function sanitizeInt_(value, min, max, fallback) {
  var num = Number(value);
  if (!isFinite(num)) {
    return fallback;
  }
  num = Math.round(num);
  if (num < min) {
    num = min;
  }
  if (num > max) {
    num = max;
  }
  return num;
}

function parseDateOrToday_(value) {
  var date = value ? new Date(value) : new Date();
  if (isNaN(date.getTime())) {
    return new Date();
  }
  return date;
}

function toIsoDateString_(value) {
  if (!value) {
    return '';
  }

  var date = new Date(value);
  if (isNaN(date.getTime())) {
    return '';
  }

  return date.toISOString();
}

function sanitizeBase64_(value) {
  var text = String(value || '').trim();
  if (!text) {
    return '';
  }

  var commaIndex = text.indexOf(',');
  if (commaIndex >= 0) {
    text = text.substring(commaIndex + 1);
  }

  return text.replace(/\s/g, '');
}

function normalizeUsername_(username) {
  return sanitizeText_(username, 60).toLowerCase();
}

function normalizeStatus_(status) {
  var clean = sanitizeText_(status, 30);
  for (var i = 0; i < PIPELINE_STATUSES.length; i++) {
    if (PIPELINE_STATUSES[i].toLowerCase() === clean.toLowerCase()) {
      return PIPELINE_STATUSES[i];
    }
  }
  return 'Inquiry';
}

function normalizeTemperature_(temperature) {
  var clean = sanitizeText_(temperature, 20);
  for (var i = 0; i < TEMPERATURE_OPTIONS.length; i++) {
    if (TEMPERATURE_OPTIONS[i].toLowerCase() === clean.toLowerCase()) {
      return TEMPERATURE_OPTIONS[i];
    }
  }
  return 'Warm';
}

function normalizeClassification_(classification) {
  var clean = sanitizeText_(classification, 40);
  for (var i = 0; i < CLASSIFICATION_OPTIONS.length; i++) {
    if (CLASSIFICATION_OPTIONS[i].toLowerCase() === clean.toLowerCase()) {
      return CLASSIFICATION_OPTIONS[i];
    }
  }
  return 'Unknown';
}

function normalizeSource_(source) {
  var clean = sanitizeText_(source, 60);
  for (var i = 0; i < SOURCE_OPTIONS.length; i++) {
    if (SOURCE_OPTIONS[i].toLowerCase() === clean.toLowerCase()) {
      return SOURCE_OPTIONS[i];
    }
  }
  return 'Organic';
}

function normalizePaymentType_(paymentType) {
  var clean = sanitizeText_(paymentType, 20);
  if (clean.toLowerCase() === 'installment') {
    return 'Installment';
  }
  return 'Spot';
}

function isBlankRow_(rowValues) {
  for (var i = 0; i < rowValues.length; i++) {
    if (String(rowValues[i] || '').trim() !== '') {
      return false;
    }
  }
  return true;
}

function generateId_(prefix) {
  return prefix + '-' + Utilities.getUuid().split('-')[0].toUpperCase();
}
