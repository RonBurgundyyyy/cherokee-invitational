function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Cherokee Invitational')
    .setFaviconUrl('favicon-32.png');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

const CHANNELS = [
  "Practice Round",
  "Tournament",
  "Sportsbook",
  "Casino",
  "Ruth's Chris",
  "Hotel",
  "Driving",
  "Shit Talk"
];

const ENTRY_HEADERS = [
  "Timestamp",
  "Name",
  "Email",
  "Handicap",
  "Practice",
  "Tournament",
  "Sportsbook"
];

const CHAT_HEADERS = [
  "Timestamp",
  "User",
  "Channel",
  "Message",
  "MessageID",
  "ParentID"
];

function saveEntry(name, email, handicap, practice, tournament, sportsbook) {
  const sheet = getSheet_("Entries", ENTRY_HEADERS);
  const entryName = cleanText_(name, 80);
  const entryEmail = cleanText_(email, 120);

  if (!entryName || !entryEmail) {
    throw new Error("Name and email are required.");
  }

  sheet.appendRow([
    new Date(),
    entryName,
    entryEmail,
    cleanText_(handicap, 20),
    yesNo_(practice),
    yesNo_(tournament),
    yesNo_(sportsbook)
  ]);
}

function getEntryNames() {
  const sheet = getSheet_("Entries", ENTRY_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  const unique = new Set(values.flat().map(n => cleanText_(n, 80)).filter(Boolean));
  return Array.from(unique).sort();
}

// Chat sheet columns:
// A: Timestamp, B: User, C: Channel, D: Message, E: MessageID, F: ParentID
function saveChatMessage(user, channel, text, parentId) {
  const sheet = getSheet_("Chat", CHAT_HEADERS);
  const entryNames = getEntryNames();
  const cleanUser = cleanText_(user, 80);
  const cleanChannel = validateChannel_(channel);
  const cleanMessage = cleanText_(text, 500);
  const cleanParentId = cleanText_(parentId, 80);

  if (!cleanUser || !entryNames.includes(cleanUser)) {
    throw new Error("Select a registered name before chatting.");
  }
  if (!cleanMessage) {
    throw new Error("Message cannot be empty.");
  }

  const id = Utilities.getUuid();

  sheet.appendRow([
    new Date(),
    cleanUser,
    cleanChannel,
    cleanMessage,
    id,
    cleanParentId || ""
  ]);
}

function getChatMessages(channel) {
  const cleanChannel = validateChannel_(channel);
  return getChatRows_()
    .filter(row => row.channel === cleanChannel)
    .map(row => row.message);
}

function getChatSnapshot(channel, lastSeen) {
  const cleanChannel = validateChannel_(channel);
  const seen = lastSeen || {};
  const unread = {};
  const rows = getChatRows_();

  CHANNELS.forEach(ch => {
    const last = seen[ch] ? new Date(seen[ch]) : new Date(0);
    unread[ch] = rows.filter(row => row.channel === ch && new Date(row.message.time) > last).length;
  });

  return {
    messages: rows
      .filter(row => row.channel === cleanChannel)
      .map(row => row.message),
    unread: unread
  };
}

function getChatData(channel) {
  const cleanChannel = validateChannel_(channel);

  return {
    channel: cleanChannel,
    channels: CHANNELS,
    names: getEntryNames(),
    messages: getChatRows_()
      .filter(row => row.channel === cleanChannel)
      .map(row => row.message)
  };
}

function getChatRows_() {
  const sheet = getSheet_("Chat", CHAT_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  return sheet.getRange(2, 1, lastRow - 1, CHAT_HEADERS.length)
    .getValues()
    .map(r => ({
      channel: cleanText_(r[2], 80),
      message: {
        time: formatTimeValue_(r[0]),
        user: cleanText_(r[1], 80),
        channel: cleanText_(r[2], 80),
        text: cleanText_(r[3], 500),
        id: cleanText_(r[4], 80),
        parent: cleanText_(r[5], 80)
      }
    }))
    .filter(row => row.channel && row.message.id);
}

function getSheet_(name, headers) {
  const ss = SpreadsheetApp.getActive();
  if (!ss) {
    throw new Error("This script must be bound to a Google Sheet.");
  }

  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }

  return sheet;
}

function validateChannel_(channel) {
  const cleanChannel = cleanText_(channel, 80);
  if (!CHANNELS.includes(cleanChannel)) {
    throw new Error("Unknown chat channel.");
  }
  return cleanChannel;
}

function cleanText_(value, maxLength) {
  return String(value || "").trim().slice(0, maxLength);
}

function yesNo_(value) {
  return value === "Yes" ? "Yes" : "No";
}

function formatTimeValue_(value) {
  if (value instanceof Date) {
    return value.toISOString();
  }
  return cleanText_(value, 80);
}
