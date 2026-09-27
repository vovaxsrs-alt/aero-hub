/**
 * Бекенд сайту (Google Apps Script), опублікований як веб-додаток.
 *
 *  doGet  — віддає JSON з товарами, новинами, медіа, партнерами й
 *           налаштуваннями (?action=data).
 *  doPost — приймає заявку з форми замовлення (записує в лист "Замовлення"
 *           і надсилає email) або запит від адмінки (action:"admin_update",
 *           див. handleAdminUpdate).
 */

// Email для сповіщень про нові заявки. Порожній рядок — брати email з листа "Налаштування".
const NOTIFY_EMAIL_OVERRIDE = "";

// Пароль адмінки (admin.html). Після зміни: Розгорнути -> Керування розгортаннями -> Нова версія.
const ADMIN_PASSWORD = "aerohub-admin-2026";

const SHEET_PRODUCTS = "Товари";
const SHEET_SETTINGS = "Налаштування";
const SHEET_ORDERS = "Замовлення";
const SHEET_NEWS = "Новини";
const SHEET_MEDIA = "Медіа";
const SHEET_PARTNERS = "Партнери";

/** Обробка GET-запитів від сайту (?action=data) */
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const data = {
    settings: readSettings(ss),
    products: readProducts(ss),
    news: readNews(ss),
    media: readMedia(ss),
    partners: readPartners(ss)
  };

  return jsonOut(data);
}

/** Обробка POST-запитів: нова заявка з форми або оновлення з адмінки */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Запит від адмінки (admin.html) на оновлення сайту
    if (body.action === "admin_update") {
      return handleAdminUpdate(ss, body);
    }

    // --- звичайна заявка з форми замовлення на сайті ---
    let sheet = ss.getSheetByName(SHEET_ORDERS);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_ORDERS);
      sheet.appendRow(["Дата", "Ім'я", "Телефон", "Товар", "Коментар"]);
    }

    sheet.appendRow([
      new Date(),
      body.name || "",
      body.phone || "",
      body.product || "",
      body.comment || ""
    ]);

    notifyOwner(ss, body);

    return jsonOut({ ok: true });

  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

/**
 * Обробляє запит з адмінки (admin.html).
 * Очікує body = { action:"admin_update", password:"...",
 *                  settings:{...},
 *                  products:[...], deleteProductIds:[...],
 *                  news:[...],     deleteNewsIds:[...],
 *                  media:[...],    deleteMediaIds:[...],
 *                  partners:[...], deletePartnersIds:[...] }.
 * Усі поля крім password необов'язкові — можна прислати тільки те, що
 * реально змінилось (або взагалі нічого, щоб просто перевірити пароль).
 * Спочатку завжди перевіряє пароль і нічого не змінює, якщо він невірний.
 */
function handleAdminUpdate(ss, body) {
  if (String(body.password || "") !== ADMIN_PASSWORD) {
    return jsonOut({ ok: false, error: "Невірний пароль" });
  }

  if (body.settings && Object.keys(body.settings).length) {
    const err = updateSettings(ss, body.settings);
    if (err) return jsonOut({ ok: false, error: err });
  }

  const collections = [
    { rows: body.products, deleteIds: body.deleteProductIds, sheet: SHEET_PRODUCTS },
    { rows: body.news, deleteIds: body.deleteNewsIds, sheet: SHEET_NEWS },
    { rows: body.media, deleteIds: body.deleteMediaIds, sheet: SHEET_MEDIA },
    { rows: body.partners, deleteIds: body.deletePartnersIds, sheet: SHEET_PARTNERS }
  ];

  for (let i = 0; i < collections.length; i++) {
    const c = collections[i];
    if (c.rows && c.rows.length) {
      const err = updateGenericRows(ss, c.sheet, c.rows);
      if (err) return jsonOut({ ok: false, error: err });
    }
    if (c.deleteIds && c.deleteIds.length) {
      const err = deleteGenericRows(ss, c.sheet, c.deleteIds);
      if (err) return jsonOut({ ok: false, error: err });
    }
  }

  return jsonOut({ ok: true });
}

/** Оновлює лист "Налаштування". Повертає текст помилки або null, якщо все ок. */
function updateSettings(ss, settings) {
  const sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet) return "Лист «Налаштування» не знайдено";

  const values = sheet.getDataRange().getValues();
  const keyRow = {}; // ключ -> номер рядка (1-based, для getRange)
  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || "").trim();
    if (key) keyRow[key] = i + 1;
  }

  Object.keys(settings).forEach(key => {
    const value = settings[key];
    if (keyRow[key]) {
      sheet.getRange(keyRow[key], 2).setValue(value);
    } else {
      // такого ключа ще нема в таблиці — додаємо новий рядок
      sheet.appendRow([key, value]);
    }
  });

  return null;
}

/**
 * Створює або оновлює рядки в будь-якому листі, де перша колонка (за
 * заголовком) називається "id" (Товари, Новини, Медіа, Партнери — усі
 * мають однакову форму: id + довільний набір текстових колонок).
 * Рядок без id або з id, якого нема в таблиці, додається як новий з
 * автоматично призначеним id. Повертає текст помилки або null.
 */
function updateGenericRows(ss, sheetName, rows) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return "Лист «" + sheetName + "» не знайдено";

  const values = sheet.getDataRange().getValues();
  if (values.length < 1) return "Лист «" + sheetName + "» порожній (нема заголовків)";

  const headers = values[0].map(h => String(h).trim());
  const idCol = headers.indexOf("id");
  if (idCol === -1) return "У листі «" + sheetName + "» нема колонки id";

  const idRow = {}; // id (рядком) -> номер рядка (1-based)
  let maxId = 0;
  for (let i = 1; i < values.length; i++) {
    const rawId = values[i][idCol];
    const idStr = String(rawId || "").trim();
    if (idStr) idRow[idStr] = i + 1;
    const num = Number(rawId);
    if (!isNaN(num) && num > maxId) maxId = num;
  }

  rows.forEach(row => {
    const idStr = (row.id !== undefined && row.id !== null) ? String(row.id).trim() : "";

    if (idStr && idRow[idStr]) {
      // оновлюємо існуючий рядок
      const rowNum = idRow[idStr];
      headers.forEach((h, idx) => {
        if (h === "id") return; // id не змінюємо
        if (Object.prototype.hasOwnProperty.call(row, h)) {
          sheet.getRange(rowNum, idx + 1).setValue(row[h]);
        }
      });
    } else {
      // новий рядок — призначаємо наступний вільний id
      maxId += 1;
      const newId = maxId;
      const newRow = headers.map(h => {
        if (h === "id") return newId;
        return Object.prototype.hasOwnProperty.call(row, h) ? row[h] : "";
      });
      sheet.appendRow(newRow);
      idRow[String(newId)] = sheet.getLastRow();
    }
  });

  return null;
}

/**
 * Видаляє рядки за списком id з будь-якого листа з колонкою id.
 * Повертає текст помилки або null.
 */
function deleteGenericRows(ss, sheetName, ids) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return "Лист «" + sheetName + "» не знайдено";

  const values = sheet.getDataRange().getValues();
  if (values.length < 1) return null;

  const headers = values[0].map(h => String(h).trim());
  const idCol = headers.indexOf("id");
  if (idCol === -1) return "У листі «" + sheetName + "» нема колонки id";

  const idSet = {};
  ids.forEach(id => { idSet[String(id).trim()] = true; });

  // Видаляємо знизу вгору, щоб номери рядків не зʼїжджали під час видалення
  for (let i = values.length - 1; i >= 1; i--) {
    const idStr = String(values[i][idCol] || "").trim();
    if (idStr && idSet[idStr]) {
      sheet.deleteRow(i + 1);
    }
  }

  return null;
}

/** Читає лист "Налаштування" (2 колонки: ключ | значення) у звичайний об'єкт */
function readSettings(ss) {
  const sheet = ss.getSheetByName(SHEET_SETTINGS);
  const settings = {};
  if (!sheet) return settings;

  const values = sheet.getDataRange().getValues();
  // Пропускаємо перший рядок (заголовки "key" / "value")
  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || "").trim();
    const value = values[i][1];
    if (key) settings[key] = value;
  }
  return settings;
}

/** Читає будь-який лист (Товари, Новини, Медіа, Партнери) у масив об'єктів */
function readGenericSheet(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(h => String(h).trim());
  const rows = values.slice(1);

  return rows
    .filter(row => row.some(cell => cell !== "" && cell !== null))
    .map(row => {
      const obj = {};
      headers.forEach((h, idx) => { obj[h] = row[idx]; });
      return obj;
    });
}

function readProducts(ss) { return readGenericSheet(ss, SHEET_PRODUCTS); }
function readNews(ss) { return readGenericSheet(ss, SHEET_NEWS); }
function readMedia(ss) { return readGenericSheet(ss, SHEET_MEDIA); }
function readPartners(ss) { return readGenericSheet(ss, SHEET_PARTNERS); }

/** Надсилає власнику email-сповіщення про нову заявку */
function notifyOwner(ss, order) {
  const settings = readSettings(ss);
  const to = NOTIFY_EMAIL_OVERRIDE || settings.email;
  if (!to) return; // нема куди слати — просто пропускаємо

  const subject = "Нова заявка з сайту: " + (order.name || "без імені");
  const body =
    "Нова заявка з сайту-візитки:\n\n" +
    "Ім'я: " + (order.name || "-") + "\n" +
    "Телефон: " + (order.phone || "-") + "\n" +
    "Товар: " + (order.product || "-") + "\n" +
    "Коментар: " + (order.comment || "-") + "\n";

  MailApp.sendEmail(to, subject, body);
}

/** Допоміжна функція: віддає обʼєкт як JSON-відповідь */
function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

