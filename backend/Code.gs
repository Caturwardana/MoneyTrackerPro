
let SS;
let SH_TRX;
let SH_WAL;
let SH_USR;
let SH_CAT;
let APP_SETUP_RUNNING = false;
const TZ = Session.getScriptTimeZone();
const CACHE = CacheService.getScriptCache();
const PROP = PropertiesService.getScriptProperties();
const WALLET_CACHE_KEY = "wallet_balances_v2";
const DASH_CACHE_PREFIX = "dash_cache_v2";
const USER_CACHE_KEY = "users_cache_v2";
const CATEGORY_CACHE_KEY = "categories_cache_v2";
const MASTER_CACHE_PREFIX = "master_cache_v2";
const MIGRATION_DONE_KEY = "migration_done_v2";
const DEFAULT_USER_PIN = "1234";
const DEFAULT_USER_NAME = "Saya";
const CACHE_TTL_SECONDS = 120;

function ensureAppSetup() {
  if (APP_SETUP_RUNNING) return;
  // Fast-path: if setup ran recently, skip heavy operations
  try {
    if (CACHE.get('SETUP_DONE') === 'true') {
      // If setup already ran recently, ensure quick references exist.
      if (!SS) SS = SpreadsheetApp.getActiveSpreadsheet();
      SH_TRX = SH_TRX || SS.getSheetByName('Transaksi');
      SH_WAL = SH_WAL || SS.getSheetByName('Wallets');
      SH_USR = SH_USR || SS.getSheetByName('Users');
      SH_CAT = SH_CAT || SS.getSheetByName('Kategori');
      // If all sheets are present, skip heavy setup. Otherwise continue to full setup.
      if (SH_TRX && SH_WAL && SH_USR && SH_CAT) return;
    }
  } catch (e) {}
  APP_SETUP_RUNNING = true;
  try {
    if (!SS) {
      SS = SpreadsheetApp.getActiveSpreadsheet();
    }
  SH_TRX = SS.getSheetByName('Transaksi') || SS.insertSheet('Transaksi');
  SH_WAL = SS.getSheetByName('Wallets') || SS.insertSheet('Wallets');
  SH_USR = SS.getSheetByName('Users') || SS.insertSheet('Users');
  SH_CAT = SS.getSheetByName('Kategori') || SS.insertSheet('Kategori');

    ensureHeaders(SH_TRX, ['tgl', 'tipe', 'kategori', 'wallet', 'nominal', 'catatan', 'trxId', 'userId']);
    ensureHeaders(SH_WAL, ['name', 'balance', 'userId']);
    ensureHeaders(SH_USR, ['id', 'name', 'pin', 'createdAt', 'isAdmin']);
    ensureHeaders(SH_CAT, ['name', 'type', 'userId']);

    if (SH_USR.getLastRow() <= 1) {
      SH_USR.appendRow([generateId(), DEFAULT_USER_NAME, DEFAULT_USER_PIN, new Date(), true]);
    }
    if (SH_WAL.getLastRow() <= 1) {
      const uid = getDefaultUserId();
      SH_WAL.appendRow(['Cash', 0, uid]);
    }
    if (SH_CAT.getLastRow() <= 1) {
      const uid = getDefaultUserId();
      SH_CAT.appendRow(['Makan & Minum', 'Pengeluaran', uid]);
      SH_CAT.appendRow(['Transportasi', 'Pengeluaran', uid]);
      SH_CAT.appendRow(['Belanja', 'Pengeluaran', uid]);
      SH_CAT.appendRow(['Lainnya', 'Pengeluaran', uid]);
      SH_CAT.appendRow(['Gaji', 'Pemasukan', uid]);
    }

    if (!PROP.getProperty(MIGRATION_DONE_KEY)) {
      migrateLegacyData();
      PROP.setProperty(MIGRATION_DONE_KEY, 'true');
    }
    // mark setup as done for a short period to avoid repeated sheet introspection
    try { CACHE.put('SETUP_DONE', 'true', 3600); } catch (e) {}
  } finally {
    APP_SETUP_RUNNING = false;
  }
}

function ensureHeaders(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    return;
  }
  const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const hasHeaders = firstRow.every((value, index) => String(value || '').toLowerCase() === headers[index].toLowerCase());
  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Money Tracker Pro')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function generateId() { return Utilities.getUuid(); }

function getDefaultUserId() {
  ensureAppSetup();
  const users = getUsers();
  return users[0] ? users[0].id : generateId();
}

function getUserById(userId) {
  ensureAppSetup();
  const users = getUsers();
  return users.find(u => u.id === userId) || users[0] || null;
}

function getSheetValues(sheet, startRow, numCols) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= startRow - 1) return [];
  return sheet.getRange(startRow, 1, lastRow - startRow + 1, numCols).getValues();
}

function getUsers() {
  ensureAppSetup();
  const cached = CACHE.get(USER_CACHE_KEY);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }
  const values = getSheetValues(SH_USR, 2, 5);
  const result = values.filter(r => r[0]).map(r => ({
    id: r[0],
    name: String(r[1] || '').trim(),
    pin: String(r[2] || '').trim(),
    createdAt: r[3],
    isAdmin: String(r[4] || '').toLowerCase() === 'true' || String(r[4] || '').toLowerCase() === '1' || (String(r[1] || '').trim() === DEFAULT_USER_NAME && String(r[2] || '').trim() === DEFAULT_USER_PIN)
  }));
  CACHE.put(USER_CACHE_KEY, JSON.stringify(result), CACHE_TTL_SECONDS);
  return result;
}

function verifyPin(pin) {
  ensureAppSetup();
  const normalized = String(pin || '').trim();
  if (!normalized) throw new Error('PIN wajib diisi');
  const users = getUsers();
  const found = users.find(u => String(u.pin || '').trim() === normalized);
  if (!found) throw new Error('PIN tidak cocok');
  return { id: found.id, name: found.name, isAdmin: found.isAdmin };
}

function loginInit(pin) {
  ensureAppSetup();
  const normalized = String(pin || '').trim();
  if (!normalized) throw new Error('PIN wajib diisi');
  const users = getUsers();
  const found = users.find(u => String(u.pin || '').trim() === normalized);
  if (!found) throw new Error('PIN tidak cocok');
  const uid = found.id;
  const masterData = getMasterData(uid);
  const start = new Date();
  start.setDate(start.getDate() - 30);
  const end = new Date();
  end.setHours(23, 59, 59);
  return {
    user: { id: found.id, name: found.name, isAdmin: found.isAdmin },
    users: masterData.users,
    masterData: {
      categories: masterData.categories,
      wallets: masterData.wallets
    },
    transactions: getTransactions({ userId: uid, tipe: 'Semua', start: start.toISOString(), end: end.toISOString() }),
    dashboard: getDashboardData({ userId: uid, tipe: 'Pengeluaran', start: start.toISOString(), end: end.toISOString() })
  };
}

const SESSION_CACHE_PREFIX = 'sess_';
const SESSION_TTL_SECONDS = 86400; // 24 jam

function loginSession(pin) {
  const result = loginInit(pin);
  const token = generateId();
  CACHE.put(SESSION_CACHE_PREFIX + token, result.user.id, SESSION_TTL_SECONDS);
  result.sessionToken = token;
  return result;
}

function verifyAutoLogin(token) {
  ensureAppSetup();
  if (!token) return null;
  const userId = CACHE.get(SESSION_CACHE_PREFIX + token);
  if (!userId) return null;
  // Refresh session TTL
  CACHE.put(SESSION_CACHE_PREFIX + token, userId, SESSION_TTL_SECONDS);
  const user = getUserById(userId);
  if (!user) return null;
  const masterData = getMasterData(userId);
  const start = new Date();
  start.setDate(start.getDate() - 30);
  const end = new Date();
  end.setHours(23, 59, 59);
  return {
    user: { id: user.id, name: user.name, isAdmin: user.isAdmin },
    users: masterData.users,
    masterData: {
      categories: masterData.categories,
      wallets: masterData.wallets
    },
    transactions: getTransactions({ userId, tipe: 'Semua', start: start.toISOString(), end: end.toISOString() }),
    dashboard: getDashboardData({ userId, tipe: 'Pengeluaran', start: start.toISOString(), end: end.toISOString() })
  };
}

function destroySession(token) {
  if (token) CACHE.remove(SESSION_CACHE_PREFIX + token);
}

function createUser(payload) {
  ensureAppSetup();
  const name = String(payload && payload.name ? payload.name : '').trim();
  const pin = String(payload && payload.pin ? payload.pin : '').trim();
  const requesterId = payload && (payload.userId || payload.adminUserId || payload.requesterId);
  const requester = requesterId ? getUserById(requesterId) : null;
  if (!name || !pin) throw new Error('Nama dan PIN wajib diisi');
  if (!requester || !requester.isAdmin) throw new Error('Hanya akun admin yang bisa menambah akun');
  if (getUsers().some(u => u.pin === pin)) throw new Error('PIN sudah dipakai');
  const id = generateId();
  SH_USR.appendRow([id, name, pin, new Date(), false]);
  ensureUserSeedData(id);
  clearAppCache(id);
  return { id, name, pin, isAdmin: false };
}

function resolveUserId(userId) {
  ensureAppSetup();
  if (userId && typeof userId === 'object') {
    userId = userId.userId || userId.id || null;
  }
  return userId || getDefaultUserId();
}

function ensureUserSeedData(userId) {
  ensureAppSetup();
  const uid = resolveUserId(userId);
  const walletRows = getSheetValues(SH_WAL, 2, 3);
  const hasWalletForUser = walletRows.some(r => String(r[2] || '').trim() === uid && String(r[0] || '').trim());
  if (!hasWalletForUser) {
    SH_WAL.appendRow(['Cash', 0, uid]);
  }

  const categoryRows = getSheetValues(SH_CAT, 2, 3);
  const hasCategoryForUser = categoryRows.some(r => String(r[2] || '').trim() === uid && String(r[0] || '').trim());
  if (!hasCategoryForUser) {
    SH_CAT.appendRow(['Makan & Minum', 'Pengeluaran', uid]);
    SH_CAT.appendRow(['Transportasi', 'Pengeluaran', uid]);
    SH_CAT.appendRow(['Belanja', 'Pengeluaran', uid]);
    SH_CAT.appendRow(['Lainnya', 'Pengeluaran', uid]);
    SH_CAT.appendRow(['Gaji', 'Pemasukan', uid]);
  }

  clearAppCache(uid);
}

function migrateLegacyData() {
  const targetUserId = getDefaultUserId();
  let needsUpdate = false;

  const walletValues = SH_WAL.getDataRange().getValues();
  for (let i = 1; i < walletValues.length; i++) {
    if (walletValues[i] && walletValues[i][0]) {
      const rowUser = String(walletValues[i][2] || '').trim();
      if (!rowUser) {
        SH_WAL.getRange(i + 1, 3).setValue(targetUserId);
        needsUpdate = true;
      }
    }
  }

  const categoryValues = SH_CAT.getDataRange().getValues();
  for (let i = 1; i < categoryValues.length; i++) {
    if (categoryValues[i] && categoryValues[i][0]) {
      const rowUser = String(categoryValues[i][2] || '').trim();
      if (!rowUser) {
        SH_CAT.getRange(i + 1, 3).setValue(targetUserId);
        needsUpdate = true;
      }
    }
  }

  const transactionValues = SH_TRX.getDataRange().getValues();
  for (let i = 1; i < transactionValues.length; i++) {
    if (transactionValues[i] && transactionValues[i][0]) {
      const rowUser = String(transactionValues[i][7] || '').trim();
      if (!rowUser) {
        SH_TRX.getRange(i + 1, 8).setValue(targetUserId);
        needsUpdate = true;
      }
    }
  }

  if (needsUpdate) {
    clearDashboardCache();
  }
}

function clearDashboardCache() {
  CACHE.remove(WALLET_CACHE_KEY);
  const newVersion = new Date().getTime().toString();
  CACHE.put('DASH_VER', newVersion, 600);
}

function clearAppCache(userId) {
  CACHE.remove(USER_CACHE_KEY);
  CACHE.remove(CATEGORY_CACHE_KEY);
  CACHE.remove(WALLET_CACHE_KEY);
  if (userId) {
    CACHE.remove(MASTER_CACHE_PREFIX + '_' + userId);
    CACHE.remove('wallets_' + userId);
    CACHE.remove('categories_' + userId);
  }
  clearDashboardCache();
}

function validateTransaksi(p) {
  ensureAppSetup();
  if (!p) throw new Error('Payload kosong');
  if (!p.wallet) throw new Error('Wallet wajib dipilih');
  if (!p.tgl) throw new Error('Tanggal wajib diisi');
  const tglDate = new Date(p.tgl);
  if (isNaN(tglDate.getTime())) throw new Error('Format tanggal tidak valid');
  const nominal = Number(p.jumlah);
  if (!nominal || nominal <= 0) throw new Error('Nominal harus lebih dari 0');
  if (nominal > 999999999999) throw new Error('Nominal terlalu besar');
  if (p.tipe === 'Transfer') {
    if (!p.walletTujuan) throw new Error('Wallet tujuan wajib dipilih untuk Transfer');
    if (p.wallet === p.walletTujuan) throw new Error('Wallet sumber dan tujuan tidak boleh sama');
  }
}

function getWalletBalances(userId) {
  ensureAppSetup();
  const uid = resolveUserId(userId);
  const cacheKey = 'wallets_' + uid;
  const cached = CACHE.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }
  ensureUserSeedData(uid);
  const values = getSheetValues(SH_WAL, 2, 3);
  const res = values
    .filter(r => r[0] && String(r[2] || '').trim() === uid)
    .map(r => ({ name: r[0], balance: Number(r[1] || 0), userId: r[2] || uid }));
  const sorted = res.sort((a, b) => a.name.localeCompare(b.name));
  CACHE.put(cacheKey, JSON.stringify(sorted), CACHE_TTL_SECONDS);
  return sorted;
}

function saveWallet(payload) {
  ensureAppSetup();
  const uid = resolveUserId(payload && payload.userId ? payload.userId : null);
  const name = String(payload && payload.name ? payload.name : '').trim();
  if (!name) throw new Error('Nama sumber dana wajib diisi');
  const existing = getWalletBalances(uid).find(w => w.name.toLowerCase() === name.toLowerCase());
  if (existing) {
    return getWalletBalances(uid);
  }
  SH_WAL.appendRow([name, Number(payload && payload.initialBalance ? payload.initialBalance : 0), uid]);
  clearAppCache(uid);
  return getWalletBalances(uid);
}

function updateBalance(name, amt, userId) {
  ensureAppSetup();
  const uid = resolveUserId(userId);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const values = getSheetValues(SH_WAL, 2, 3);
    let found = false;
    for (let i = 0; i < values.length; i++) {
      const rowUser = values[i][2] || uid;
      if (values[i][0] === name && String(rowUser) === uid) {
        const newBalance = Number(values[i][1] || 0) + amt;
        SH_WAL.getRange(i + 2, 2).setValue(newBalance);
        found = true;
        break;
      }
    }
    if (!found) throw new Error('Wallet not found: ' + name);
    clearAppCache(uid);
    return true;
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function getCategories(userId) {
  ensureAppSetup();
  const uid = resolveUserId(userId);
  const cacheKey = 'categories_' + uid;
  const cached = CACHE.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }
  ensureUserSeedData(uid);
  const values = getSheetValues(SH_CAT, 2, 3);
  const result = values
    .filter(r => r[0] && String(r[2] || '').trim() === uid)
    .map(r => ({ name: r[0], type: r[1] || 'Pengeluaran' }))
    .sort((a, b) => a.name.localeCompare(b.name));
  CACHE.put(cacheKey, JSON.stringify(result), CACHE_TTL_SECONDS);
  return result;
}

function saveCategory(payload) {
  ensureAppSetup();
  const uid = resolveUserId(payload && payload.userId ? payload.userId : null);
  const name = String(payload && payload.name ? payload.name : '').trim();
  const type = String(payload && payload.type ? payload.type : 'Pengeluaran').trim();
  if (!name) throw new Error('Nama kategori wajib diisi');
  const existing = getCategories(uid).find(c => c.name.toLowerCase() === name.toLowerCase() && c.type === type);
  if (existing) {
    return getCategories(uid);
  }
  SH_CAT.appendRow([name, type, uid]);
  clearAppCache(uid);
  return getCategories(uid);
}

function editWallet(payload) {
  ensureAppSetup();
  const uid = resolveUserId(payload && payload.userId ? payload.userId : null);
  const oldName = String(payload && payload.oldName ? payload.oldName : '').trim();
  const newName = String(payload && payload.newName ? payload.newName : '').trim();
  if (!oldName || !newName) throw new Error('Nama lama dan baru wajib diisi');
  const values = getSheetValues(SH_WAL, 2, 3);
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === oldName && String(values[i][2] || '').trim() === uid) {
      SH_WAL.getRange(i + 2, 1).setValue(newName);
      // Cascade-update existing transaksi records
      const trxValues = getSheetValues(SH_TRX, 2, 8);
      for (let j = 0; j < trxValues.length; j++) {
        if (trxValues[j][3] === oldName && String(trxValues[j][7] || '').trim() === uid) {
          SH_TRX.getRange(j + 2, 4).setValue(newName);
        }
      }
      clearAppCache(uid);
      return getWalletBalances(uid);
    }
  }
  throw new Error('Wallet tidak ditemukan: ' + oldName);
}

function deleteWallet(payload) {
  ensureAppSetup();
  const uid = resolveUserId(payload && payload.userId ? payload.userId : null);
  const name = String(payload && payload.name ? payload.name : '').trim();
  if (!name) throw new Error('Nama wallet wajib diisi');
  const values = getSheetValues(SH_WAL, 2, 3);
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === name && String(values[i][2] || '').trim() === uid) {
      SH_WAL.deleteRow(i + 2);
      clearAppCache(uid);
      return getWalletBalances(uid);
    }
  }
  throw new Error('Wallet tidak ditemukan: ' + name);
}

function editCategory(payload) {
  ensureAppSetup();
  const uid = resolveUserId(payload && payload.userId ? payload.userId : null);
  const oldName = String(payload && payload.oldName ? payload.oldName : '').trim();
  const newName = String(payload && payload.newName ? payload.newName : '').trim();
  const type = String(payload && payload.type ? payload.type : 'Pengeluaran').trim();
  if (!oldName || !newName) throw new Error('Nama lama dan baru wajib diisi');
  const values = getSheetValues(SH_CAT, 2, 3);
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === oldName && String(values[i][2] || '').trim() === uid) {
      SH_CAT.getRange(i + 2, 1).setValue(newName);
      SH_CAT.getRange(i + 2, 2).setValue(type);
      // Cascade-update existing transaksi records
      const trxValues = getSheetValues(SH_TRX, 2, 8);
      for (let j = 0; j < trxValues.length; j++) {
        if (trxValues[j][2] === oldName && String(trxValues[j][7] || '').trim() === uid) {
          SH_TRX.getRange(j + 2, 3).setValue(newName);
        }
      }
      clearAppCache(uid);
      return getCategories(uid);
    }
  }
  throw new Error('Kategori tidak ditemukan: ' + oldName);
}

function deleteCategory(payload) {
  ensureAppSetup();
  const uid = resolveUserId(payload && payload.userId ? payload.userId : null);
  const name = String(payload && payload.name ? payload.name : '').trim();
  if (!name) throw new Error('Nama kategori wajib diisi');
  const values = getSheetValues(SH_CAT, 2, 3);
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === name && String(values[i][2] || '').trim() === uid) {
      SH_CAT.deleteRow(i + 2);
      clearAppCache(uid);
      return getCategories(uid);
    }
  }
  throw new Error('Kategori tidak ditemukan: ' + name);
}

function getMasterData(userId) {
  ensureAppSetup();
  const uid = resolveUserId(userId);
  const cacheKey = MASTER_CACHE_PREFIX + '_' + uid;
  const cached = CACHE.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }
  ensureUserSeedData(uid);
  // Batch-read all sheets in one go to avoid N+1
  const catValues = getSheetValues(SH_CAT, 2, 3);
  const walValues = getSheetValues(SH_WAL, 2, 3);
  const categories = catValues
    .filter(r => r[0] && String(r[2] || '').trim() === uid)
    .map(r => ({ name: r[0], type: r[1] || 'Pengeluaran' }))
    .sort((a, b) => a.name.localeCompare(b.name));
  const wallets = walValues
    .filter(r => r[0] && String(r[2] || '').trim() === uid)
    .map(r => ({ name: r[0], balance: Number(r[1] || 0), userId: r[2] || uid }))
    .sort((a, b) => a.name.localeCompare(b.name));
  const result = {
    user: getUserById(uid),
    users: getUsers(),
    categories: categories,
    wallets: wallets
  };
  CACHE.put(cacheKey, JSON.stringify(result), CACHE_TTL_SECONDS);
  return result;
}

function simpanTransaksi(p) {
  validateTransaksi(p);
  const uid = resolveUserId(p && p.userId ? p.userId : null);
  const tgl = new Date(p.tgl);
  const nominal = Number(p.jumlah);
  const kategoriFinal = (p.kategori === 'Lainnya' && p.kategoriKustom) ? p.kategoriKustom : p.kategori;
  const trxId = p.trxId || generateId();
  if (p.rowId) {
    updateExistingTransaction(p, uid);
    return true;
  }
  if (p.tipe === 'Transfer') {
    try {
      SH_TRX.appendRow([tgl, 'Transfer Out', 'Sistem', p.wallet, nominal, `Ke: ${p.walletTujuan} | ${p.catatan}`, trxId, uid]);
      SH_TRX.appendRow([tgl, 'Transfer In', 'Sistem', p.walletTujuan, nominal, `Dari: ${p.wallet} | ${p.catatan}`, trxId, uid]);
      updateBalance(p.wallet, -nominal, uid);
      updateBalance(p.walletTujuan, nominal, uid);
    } catch (e) {
      deleteByTransactionId(trxId, uid);
      throw new Error('Transfer gagal: ' + e.message);
    }
  } else {
    SH_TRX.appendRow([tgl, p.tipe, kategoriFinal, p.wallet, nominal, p.catatan, trxId, uid]);
    updateBalance(p.wallet, p.tipe === 'Pemasukan' ? nominal : -nominal, uid);
  }
  clearAppCache(uid);
  return true;
}

function updateExistingTransaction(p, userId) {
  ensureAppSetup();
  const row = Number(p.rowId);
  if (row <= 1) throw new Error('Row invalid');
  const old = SH_TRX.getRange(row, 1, 1, 8).getValues()[0];
  const trxId = old[6];
  if (trxId) deleteByTransactionId(trxId, userId || old[7]);
  else { revertBalance(old, userId || old[7]); SH_TRX.deleteRow(row); }
  simpanTransaksi({ ...p, kategori: p.kategori || old[2], trxId: p.trxId || old[6], rowId: null, userId: userId || old[7] });
}

function revertBalance(row, userId) {
  const tipe = row[1];
  const wallet = row[3];
  const nominal = Number(row[4]);
  if (tipe === 'Pemasukan' || tipe === 'Transfer In') updateBalance(wallet, -nominal, userId);
  else if (tipe === 'Pengeluaran' || tipe === 'Transfer Out') updateBalance(wallet, nominal, userId);
}

function hapusTransaksi(rowId, userId) {
  ensureAppSetup();
  const row = Number(rowId);
  if (row <= 1) return false;
  const data = SH_TRX.getRange(row, 1, 1, 8).getValues()[0];
  const trxId = data[6];
  const uid = resolveUserId(userId || data[7]);
  if (trxId) deleteByTransactionId(trxId, uid);
  else { revertBalance(data, uid); SH_TRX.deleteRow(row); }
  clearAppCache(uid);
  return true;
}

function deleteByTransactionId(trxId, userId) {
  ensureAppSetup();
  const values = getSheetValues(SH_TRX, 2, 8);
  const uid = resolveUserId(userId);
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][6] === trxId && String(values[i][7] || uid) === uid) {
      revertBalance(values[i], uid);
      SH_TRX.deleteRow(i + 2);
    }
  }
  clearAppCache(uid);
}

function getTransactions(f) {
  ensureAppSetup();
  const uid = resolveUserId(f && f.userId ? f.userId : null);
  const values = getSheetValues(SH_TRX, 2, 8);
  const start = new Date(f && f.start ? f.start : new Date());
  const end = new Date(f && f.end ? f.end : new Date());
  end.setHours(23, 59, 59);
  const result = [];
  values.forEach((r, index) => {
    const rowUser = String(r[7] || '').trim();
    if (!rowUser || rowUser !== uid || !r[0]) return;
    const d = new Date(r[0]);
    const matchTipe = f && f.tipe === 'Semua' ? true : r[1] === f.tipe || (f.tipe === 'Pemasukan' && r[1] === 'Transfer In') || (f.tipe === 'Pengeluaran' && r[1] === 'Transfer Out');
    if (d >= start && d <= end && matchTipe) {
      result.push({
        tgl: (function(dt){ const d=new Date(dt); const dd=('0'+d.getDate()).slice(-2); const mm=('0'+(d.getMonth()+1)).slice(-2); return dd + '/' + mm; })(r[0]),
        tglRaw: (function(dt){ const d=new Date(dt); const yyyy=d.getFullYear(); const mm=('0'+(d.getMonth()+1)).slice(-2); const dd=('0'+d.getDate()).slice(-2); const hh=('0'+d.getHours()).slice(-2); const min=('0'+d.getMinutes()).slice(-2); return `${yyyy}-${mm}-${dd}T${hh}:${min}`; })(r[0]),
        tipe: r[1],
        kat: r[2],
        wallet: r[3],
        nominal: Number(r[4] || 0),
        note: r[5],
        row: index + 2,
        trxId: r[6]
      });
    }
  });
  const sorted = result.sort((a, b) => new Date(b.tglRaw) - new Date(a.tglRaw));
  const page = Number(f && f.page ? f.page : 0);
  const limit = Number(f && f.limit ? f.limit : 0);
  if (page > 0 && limit > 0) {
    const startIdx = (page - 1) * limit;
    return { data: sorted.slice(startIdx, startIdx + limit), total: sorted.length, page, limit };
  }
  return sorted;
}

function getDashboardData(f) {
  ensureAppSetup();
  const uid = resolveUserId(f && f.userId ? f.userId : null);
  const ver = CACHE.get('DASH_VER') || '0';
  const cacheKey = DASH_CACHE_PREFIX + ver + '_' + Utilities.base64Encode(JSON.stringify({ ...f, userId: uid }));
  const cached = CACHE.get(cacheKey);
  if (cached) return JSON.parse(cached);

  const trx = getSheetValues(SH_TRX, 2, 8);
  const wal = getSheetValues(SH_WAL, 2, 3);
  const start = new Date(f && f.start ? f.start : new Date());
  const end = new Date(f && f.end ? f.end : new Date());
  end.setHours(23, 59, 59);

  let totalSemua = 0;
  const walletDetails = wal
    .filter(r => r[0] && String(r[2] || '').trim() === uid)
    .map(r => {
      totalSemua += Number(r[1] || 0);
      return { name: r[0], balance: Number(r[1] || 0) };
    });

  const catStats = {};
  let totalFiltered = 0;
  trx.forEach(r => {
    const rowUser = String(r[7] || '').trim();
    if (!rowUser || rowUser !== uid) return;
    const d = new Date(r[0]);
    const selectedType = (f && f.tipe) || 'Pengeluaran';
    const matchType = selectedType === 'Semua'
      ? true
      : r[1] === selectedType || (selectedType === 'Pemasukan' && r[1] === 'Transfer In') || (selectedType === 'Pengeluaran' && r[1] === 'Transfer Out');
    if (d >= start && d <= end && matchType) {
      const nom = Number(r[4] || 0);
      catStats[r[2]] = (catStats[r[2]] || 0) + nom;
      totalFiltered += nom;
    }
  });

  const finalResult = { totalSemua, walletDetails, catStats, totalFiltered };
  CACHE.put(cacheKey, JSON.stringify(finalResult), 600);
  return finalResult;
}

function getExportData(params) {
  ensureAppSetup();
  const uid = resolveUserId(Array.isArray(params) ? null : params && params.userId ? params.userId : null);
  const startStr = Array.isArray(params) ? params[0] : params && params.start;
  const endStr = Array.isArray(params) ? params[1] : params && params.end;
  if (!startStr || !endStr) throw new Error('Invalid export date range');

  const start = new Date(startStr);
  const end = new Date(endStr);
  end.setHours(23, 59, 59);

  const data = SH_TRX.getDataRange().getValues().slice(1).filter(r => String(r[7] || '').trim() === uid);
  data.sort((a, b) => new Date(a[0]) - new Date(b[0]));

  let runningTotal = 0;
  const exportData = [];
  const yearSuffix = new Date().getFullYear();
  data.forEach((r, i) => {
    if (!r[0]) return;
    const tglTrx = new Date(r[0]);
    if (isNaN(tglTrx)) return;
    const nominal = Number(r[4]) || 0;
    const isDebit = r[1] === 'Pemasukan' || r[1] === 'Transfer In';
    runningTotal += isDebit ? nominal : -nominal;
    if (tglTrx >= start && tglTrx <= end) {
      exportData.push({
        kode: `PP${yearSuffix}-${String(i + 1).padStart(3, '0')}`,
        tgl: Utilities.formatDate(tglTrx, TZ, 'dd/MM/yyyy'),
        wallet: r[3] || '-',
        keterangan: r[2] || '-',
        catatan: r[5] || '-',
        debit: isDebit ? nominal : 0,
        kredit: !isDebit ? nominal : 0,
        total: Number(runningTotal.toFixed(2))
      });
    }
  });
  return exportData;
}

function getMonthlyRecap(f) {
  ensureAppSetup();
  const uid = resolveUserId(f && f.userId ? f.userId : null);
  const year = Number(f && f.year ? f.year : new Date().getFullYear());
  const trx = getSheetValues(SH_TRX, 2, 8);
  const months = [];
  for (let m = 0; m < 12; m++) {
    let income = 0, expense = 0, count = 0;
    trx.forEach(r => {
      if (String(r[7] || '').trim() !== uid || !r[0]) return;
      const d = new Date(r[0]);
      if (d.getFullYear() === year && d.getMonth() === m) {
        const tipe = r[1];
        const nominal = Number(r[4] || 0);
        if (tipe === 'Pemasukan' || tipe === 'Transfer In') { income += nominal; } else if (tipe === 'Pengeluaran' || tipe === 'Transfer Out') { expense += nominal; }
        count++;
      }
    });
    months.push({ month: m + 1, income, expense, count });
  }
  return { year, months };
}

function escapeCsvField(str) {
  if (!str) return '';
  str = String(str);
  if (str.includes(',') || str.includes('"') || str.includes('\n') || str.includes('\r')) {
    return '"' + str.replace(/"/g, '""') + '"';
  }
  return str;
}