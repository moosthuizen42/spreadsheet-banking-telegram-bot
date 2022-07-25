/* eslint-disable @typescript-eslint/no-unused-vars */
/* global console setInterval, clearInterval */

/**
 * Add two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

/**
 * Displays the current time once a second
 * @customfunction
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 */
function clock(invocation) {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time
 * @returns {string} String with the current time formatted for the current locale.
 */
function currentTime() {
  return new Date().toLocaleTimeString();
  // return new Date().toISOString();
}

/**
 * Increments a value once a second.
 * @customfunction
 * @param {number} incrementBy Amount to increment
 * @param {CustomFunctions.StreamingInvocation<number>} invocation
 */
function increment(incrementBy, invocation) {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param {string} message String to write.
 * @returns String to write.
 */
function logMessage(message) {
  console.log(`${new Date().toLocaleTimeString()} | ${message}`);

  return message;
}

/**
 * @customfunction
 * @param {string} address The address of the cell from which to retrieve the value.
 * @returns The value of the cell at the input address.
 **/
 async function getRangeValue(address) {
  // Retrieve the context object.
  console.log("Retrieving the context object...");
  const context = new Excel.RequestContext();
  console.log("Success.");



  // const range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
  const range = context.workbook.worksheets.getItem(TELEGRAM_BOT_TOKEN_SHEET).getRange(address);
  range.load("values");
  await context.sync();

  // Return the value of the cell at the input address.
  return range.values[0][0];
}

/**
 * @customfunction
 * @param {string} tableName The name of the table to load
 * @returns A value
 **/
 async function loadTableValues(tableName) {

  console.log("Retrieving the context object...");
  // Retrieve the context object.
  const context = new Excel.RequestContext();
  console.log("Success.");

  const sheet = context.workbook.worksheets.getItem(TELEGRAM_BOT_TOKEN_SHEET);
  const table = sheet.tables.getItem(TELEGRAM_BOT_TOKEN_TABLE).getRange().load();

  // const range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
  await context.sync();

  return table.cellCount;
}


const TELEGRAM_MESSAGE_HISTORY_SHEET = `💬 Telegram log`;
const TELEGRAM_MESSAGE_HISTORY_TABLE = `tblMessageHistory`;

const TELEGRAM_CONFIGURATOR_SHEET = `💬 Telegram configurator`;
const TELEGRAM_CONFIGURATOR_TABLE = `tblConfigurator`;

const TELEGRAM_BOT_SETUP_SHEET = `💬 Telegram setup`;
const TELEGRAM_BOT_USERNAME_RANGE = `B16`;
const TELEGRAM_BOT_TOKEN_RANGE = `B23`;


/**
 * @customfunction
 * @param {string} message Name of the sheet where the table can be found.
 * @param {string} message Name of the table (set under "Table Design").
 * @param {string[]} message Array representing the data within a single row.
 * @returns {number} The number of data rows in the table after new data has been inserted.
 **/
async function insertTableRow (sheetName, tableName, data) {
  try {
    console.log("Retrieving the context object...");
    // Retrieve the context object.
    const context = new Excel.RequestContext();
    console.log("OK");

    console.log("Retrieving the worksheet...");
    const sheet = context.workbook.worksheets.getItem(sheetName);
    console.log("OK");

    console.log("Retrieving the table...");
    const table = sheet.tables.getItem(tableName);
    console.log("OK");

    console.log("Inserting data into table...");
    table.rows.add(null, [data], true);
    console.log("OK");

    console.log("Synchronizing context...");
    const loadedTable = table.getDataBodyRange().load();
    await context.sync();
    console.log("OK");

    console.log(`New data row count: ${loadedTable.rowCount}`);
    return loadedTable.rowCount;
  }
  catch (error) {
    console.error(error);
  }
}

const TELEGRAM_UPDATE_LOOP_INTERVAL = 5000; // ms

/**
 * @customfunction
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 */
 async function telegramCheckUpdatesLoop (setting, invocation) {
  const timer = setInterval(async () => {
    if (setting != "On") {
      invocation.setResult("Telegram bot disabled.");
    }
    else {
      await telegramCheckUpdates();
      invocation.setResult("Telegram bot online! Last updated " + new Date().toLocaleString());
    }
  }, TELEGRAM_UPDATE_LOOP_INTERVAL);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * @customfunction
 * @returns {string} A timestamp of when the function completed execution.
 **/
 async function telegramCheckUpdates() {

  const botUsername = await loadCellValue(TELEGRAM_BOT_SETUP_SHEET, TELEGRAM_BOT_USERNAME_RANGE);
  const botToken = await loadCellValue(TELEGRAM_BOT_SETUP_SHEET, TELEGRAM_BOT_TOKEN_RANGE);

  const offset = (await telegramGetLastUpdateId()) + 1;
  console.log(`update_id offset: ${offset}`);

  await loadConversationTree();

  // const url = `https://api.telegram.org/bot5390338534:AAFXsi3Pjf14PPxyh15nDZBksutzkbuuuF8/getUpdates?offset=${offset}`;
  const url = `https://api.telegram.org/bot${botToken}/getUpdates?offset=${offset}`;
  const response = await fetch(url);

  const updates = (await response.json())?.result;

  if (!updates || !Array.isArray(updates) || !updates.length) {
    return `Checked at ${new Date().toLocaleTimeString()}, no updates.`;
  }

  console.log(updates);
  updates.forEach((update => telegramProcessUpdate(update, botToken, botUsername)));

  return `Checked at ${new Date().toLocaleTimeString()}, updates received: ${updates.length}`;

}

async function telegramProcessUpdate (update, botToken, botUsername) {
  const updateId = update?.update_id;
  const chatId = update?.message?.chat?.id;
  const fromId = update?.message?.from?.id;
  const fromName = update?.message?.from?.first_name + " " + update?.message?.from?.last_name;
  const receivedTime = new Date(update?.message?.date * 1000).toLocaleString();
  const receivedMessage = update?.message?.text;

  if (!updateId || !fromId || !receivedMessage) return;


  await Promise.all([
    await insertTableRow(TELEGRAM_MESSAGE_HISTORY_SHEET, TELEGRAM_MESSAGE_HISTORY_TABLE, [updateId, fromId, fromName, receivedTime, receivedMessage]),
    await telegramDetermineAction(chatId, fromId, receivedMessage, botToken, botUsername),
  ]);

}

function button (text) {
  return { text };
}

async function telegramDetermineAction (chatId, fromId, receivedText, botToken, botUsername) {

  const replyId = undefined;
  const options = await loadConversationTree();

  // If the message if prefixed with "Back to " remove it.
  if (receivedText.match(/^Back to .*$/)) receivedText = receivedText.substring(8);

  // Alias "/start" to whatever the first option is (presumably the "Main menu").
  if (receivedText.match(/^\/start$/)) receivedText = options[0].optionName;

  // Check if the received message matches one of the options.
  let matchingOption = options.find(i => i.optionName === receivedText);
  console.log("First pass option matching");
  console.log(matchingOption);

  if (!matchingOption) {
    // If the received message DOES NOT match one of the options, find the most recent matching message.
    matchingOption = await telegramGetMostRecentValidOption(fromId, options);
  }

  console.log("Second pass option matching");
  console.log(matchingOption);

  if (!matchingOption) {
    await telegramSendMessage(chatId, replyId, `No option "${receivedText}" has been programmed.`, undefined, botToken, botUsername);
    return;
  }

  const responseText = matchingOption.response || `You selected ${matchingOption.optionName}`; // Use the specified response, otherwise just echo.

  const keyboardOptions = [];
  const childOptions = options
    .filter(i => i.parentOption === receivedText)
    .map(i => [ button(i.optionName) ]);

  console.log("Performing action...");
  console.log(matchingOption.action);
  switch (matchingOption.action) {
    case "Show child options": {
      keyboardOptions.push(...childOptions);
      if (matchingOption.parentOption) keyboardOptions.unshift([ button(`Back to ${matchingOption.parentOption}`) ]);
      telegramSendMessage(chatId, replyId, responseText, { keyboard: keyboardOptions }, botToken, botUsername);
      break;
    }
    case "Return a value": {
      telegramSendMessage(chatId, replyId, responseText, undefined, botToken, botUsername); // undefined = Don't change the keyboard
      const value = await loadCellValue(matchingOption.inputSheet, matchingOption.inputRange);
      telegramSendMessage(chatId, replyId, value, undefined, botToken, botUsername); // undefined = Don't change the keyboard
      break;
    }
    case "Enumerate a list (limit to top 5 rows)": {
      telegramSendMessage(chatId, replyId, responseText, undefined, botToken, botUsername); // undefined = Don't change the keyboard
      const list = await loadListFromTable(matchingOption.inputSheet, matchingOption.inputRange, 0, 5);
      list.forEach(i => { telegramSendMessage(chatId, replyId, i, undefined, botToken, botUsername) });
      break;
    }
    case "Enumerate a list (limit to bottom 5 rows)": {
      telegramSendMessage(chatId, replyId, responseText, undefined, botToken, botUsername); // undefined = Don't change the keyboard
      const list = await loadListFromTable(matchingOption.inputSheet, matchingOption.inputRange, -5, undefined);
      list.forEach(i => { telegramSendMessage(chatId, replyId, i, undefined, botToken, botUsername) });
      break;
    }
    case "Ask the AI a question": {
      const waitingForQuestion = receivedText == matchingOption.optionName;
      if (waitingForQuestion) {
        telegramSendMessage(chatId, replyId, responseText, undefined, botToken, botUsername); // undefined = Don't change the keyboard
      }
      else {
        telegramSendMessage(chatId, replyId, `Let me think on this...`, undefined, botToken, botUsername); // undefined = Don't change the keyboard
        const data = await loadCellValue(matchingOption.inputSheet, matchingOption.inputRange);
        const response = await askGPT(`${data}\n${receivedText}`);
        telegramSendMessage(chatId, replyId, response, undefined, botToken, botUsername);
      }
      break;
    }
    default: {
      telegramSendMessage(chatId, replyId, `The wizard has been hampered by a missing page in the tome of incantations! (No action has been specified for the "${matchingOption.optionName}" option. Please check your spreadsheet.)`, undefined, botToken, botUsername); // undefined = Don't change the keyboard
      break;
    }
  }




}

async function telegramSendMessage (chatId, replyId, messageText, replyMarkup, botToken, botUsername) {

  const botUserId = botToken.split(":")[0];

  const messageUrl = `https://api.telegram.org/bot${botToken}/sendMessage`;
  const messageBody = {
    chat_id: chatId,
    text: messageText,
    reply_to_message_id: replyId,
    reply_markup: replyMarkup,
  };

  return Promise.all([
    insertTableRow(TELEGRAM_MESSAGE_HISTORY_SHEET, TELEGRAM_MESSAGE_HISTORY_TABLE, [0, botUserId, botUsername, new Date().toLocaleString(), messageText]),
    fetch(messageUrl, { method: "POST", headers: { "Content-Type": "application/json"}, body: JSON.stringify(messageBody) })
  ]);

}

async function loadTableHeaders(sheetName, tableName) {
  const context = new Excel.RequestContext();

  const sheet = context.workbook.worksheets.getItem(sheetName);
  const table = sheet.tables.getItem(tableName);

  const loadedRange = table.getHeaderRowRange().load();
  await context.sync();

  const values = loadedRange.values;
  // console.log(values);

  return values;
}

async function loadTableBody(sheetName, tableName) {
  const context = new Excel.RequestContext();

  const sheet = context.workbook.worksheets.getItem(sheetName);
  const table = sheet.tables.getItem(tableName);

  const loadedRange = table.getDataBodyRange().load();
  await context.sync();

  const values = loadedRange.values;
  // console.log(values);

  return values;
}

async function loadCellValue (sheetName, range) {

  const context = new Excel.RequestContext();

  const sheet = context.workbook.worksheets.getItem(sheetName);
  const rangeValues = sheet.getRange(range)
  rangeValues.load("values");
  await context.sync();

  console.log(rangeValues);
  console.log(rangeValues.values[0][0]);
  return rangeValues.values[0][0];
}


async function loadListFromTable(sheetName, tableName, start, end) {

  const [header, rows] = await Promise.all([
    loadTableHeaders(sheetName, tableName),
    loadTableBody(sheetName, tableName),
  ]);

  console.log(rows);
  const nonEmpty = rows
    .filter(row => !row.every(col => !col))
    .slice(start, end);
  console.log(nonEmpty);
  const list = nonEmpty.map(row => header[0].map((value,index)=>`${value}: ${row[index] || "<no data>"}`).join("\n"));
  console.log(list);

  return list;

}


/**
 * @customfunction
 * @returns {string} The largest update ID found in the message history.
 **/
 async function telegramGetLastUpdateId() {

  const values = await loadTableBody(TELEGRAM_MESSAGE_HISTORY_SHEET, TELEGRAM_MESSAGE_HISTORY_TABLE);
  const updateIds = values.map(v => v[0])

  updateIds.sort((a,b) => a < b ? 1 : -1);

  return updateIds.length && updateIds[0] || 0;

}

async function telegramGetMostRecentValidOption(userId, options) {

  const values = await loadTableBody(TELEGRAM_MESSAGE_HISTORY_SHEET, TELEGRAM_MESSAGE_HISTORY_TABLE);
  const filtered = values
    .filter(v => v[1] == userId)
    .map(v => { return {
      date: v[3],
      text: v[4]
    }})

  filtered.sort((a,b) => a.date < b.date ? 1 : -1);

  for (const f of filtered) {
      // If the message if prefixed with "Back to " remove it.
    if (f.text.match(/^Back to .*$/)) f.text = f.text.substring(8);

    // Alias "/start" to whatever the first option is (presumably the "Main menu").
    if (f.text.match(/^\/start$/)) f.text = options[0].optionName;

    const option = options.find(i => i.optionName === f.text);
    if (option) return option;
  }

  return null;
}

async function loadConversationTree() {
  const values = await loadTableBody(TELEGRAM_CONFIGURATOR_SHEET, TELEGRAM_CONFIGURATOR_TABLE);
  const options = values
    .map(v => { return {
      parentOption: v[0],
      optionName: v[1],
      response: v[2],
      action: v[3],
      inputSheet: v[4],
      inputRange: v[5]
    }})
    .filter(v => v.optionName);
  console.log(options);
  return options;
}

/**
 * @customfunction
 * @returns {string[]} An array of valid parent option selections.
 **/
 async function getParentOptions() {

  const values = await loadTableBody(TELEGRAM_CONFIGURATOR_SHEET, TELEGRAM_CONFIGURATOR_TABLE);
  const options = values.map(v => v[1]);
  return options;

}

/**
 * @customfunction
 * @returns {string[]} An array of valid sheet names.
 **/
 async function getSheetNames() {

  const context = new Excel.RequestContext();

  console.log(`Loading...`);
  const sheets = context.workbook.worksheets.forEach(s => s.load());
  await context.sync();

  console.log(sheets);

}

/**
 * @customfunction
 * @returns {string} A message describing whether validation was passed.
 **/
 async function validateBotName(botUsername) {
    if (!botUsername) {
      return "Enter your bot's username above.";
    }

    if (!botUsername.toLowerCase().match(/^[^@]+bot$/)) {
      return "That bot username seems invalid.";
    }

    return "Success! An excellent username my liege!";
}

/**
 * @customfunction
 * @returns {string} A message describing whether validation was passed.
 **/
 async function validateBotToken(botToken) {
  if (!botToken) {
    return `Enter the bot token above.`;
  }

  try {
    const url = `https://api.telegram.org/bot${botToken}/getWebhookInfo`;
    const response = await fetch(url);
    if (!(await response.json()).ok) {
      throw new Error(`"ok" field was not truthy.`);
    }
    return `Success! You are now ready to dispatch your messenger on its communicative errands!`;
  }
  catch (error) {
    console.log(`An error was thrown, likely because of an invalid bot token`);
    console.warn(error);
    return `Something went wrong. Please double-check that the token was correctly copied.`;
  }

}


const GPT_API_SECRET = `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`;

/**
 * Taken from Spreadsheet Banking Google Sheets example (not sure who to attribute credit to).
 * @customfunction
 * @param {string} prompt A string containing the input to GPT.
 * @returns {string} A string containing GPT output.
 **/
async function askGPT(prompt) {
  try {
    const data = {
      prompt: prompt,
      max_tokens: 100
    };
    const options = {
      method: 'POST',
      headers: { "Authorization": `Bearer ${GPT_API_SECRET}`, "Content-Type": "application/json"  },
      body: JSON.stringify(data)
    };
    const response = await fetch("https://api.openai.com/v1/engines/text-davinci-002/completions", options);
    const completion = (await response.json()).choices[0].text;
    return completion;
  }
  catch (error) {
    console.error(error);
    return `Who knows? (an error occured)`;
  }

}

