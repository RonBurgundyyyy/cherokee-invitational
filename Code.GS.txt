function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Cherokee Invitational')
    .setFaviconUrl('favicon-32.png');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function saveEntry(name, email, handicap, practice, tournament, sportsbook) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("Entries");

  sheet.appendRow([
    new Date(),
    name,
    email,
    handicap,
    practice,
    tournament,
    sportsbook
  ]);
}

function getEntryNames() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Entries");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  return values.flat().filter(n => n);
}

// Chat sheet columns:
// A: Timestamp, B: User, C: Channel, D: Message, E: ThreadID, F: ParentID
function saveChatMessage(user, channel, text, parentId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Chat");
  const id = Utilities.getUuid();
  const threadId = parentId || id;

  sheet.appendRow([
    new Date(),
    user,
    channel,
    text,
    threadId,
    parentId || ""
  ]);
}

function getChatMessages(channel) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Chat");
  const rows = sheet.getDataRange().getValues().slice(1);

  return rows
    .filter(r => r[2] === channel)
    .map(r => ({
      time: r[0],
      user: r[1],
      channel: r[2],
      text: r[3],
      id: r[4],
      parent: r[5]
    }));
}
