/**
 * ==========================================================================
 *  БЕКЕНД САЙТУ-ВІЗИТКИ (Google Apps Script)
 * ==========================================================================
 *  Що робить цей файл:
 *   1) doGet   -> віддає сайту JSON з товарами та налаштуваннями
 *                 (сайт робить GET-запит на твій опублікований URL)
 *   2) doPost  -> приймає заявку з форми замовлення, записує її у лист
 *                 "Замовлення" і надсилає тобі лист на пошту
 *
 *  Як підключити — див. README.md, розділ "Крок 2".
 * ==========================================================================
 */

// Якщо хочеш, щоб сповіщення про замовлення завжди йшли на конкретну
// адресу (а не на ту, що вказана у листі "Налаштування"), впиши її тут.
// Залиш порожнім (""), щоб бекенд сам узяв email з листа "Налаштування".
const NOTIFY_EMAIL_OVERRIDE = "";

const SHEET_PRODUCTS = "Товари";
const SHEET_SETTINGS = "Налаштування";
const SHEET_ORDERS = "Замовлення";

/** Обробка GET-запитів від сайту (?action=data) */
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const data = {
    settings: readSettings(ss),
    products: readProducts(ss)
  };

  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Обробка POST-запитів — нова заявка з форми замовлення */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

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

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
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

/** Читає лист "Товари" у масив об'єктів */
function readProducts(ss) {
  const sheet = ss.getSheetByName(SHEET_PRODUCTS);
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
