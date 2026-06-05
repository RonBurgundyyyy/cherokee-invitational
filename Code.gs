function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  template.initialTeamNumber = cleanTeamNumber_(e && e.parameter ? e.parameter.team : "");
  return template
    .evaluate()
    .setTitle('Cherokee Invitational')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
  "Sportsbook",
  "Hotel 2 Nights Minimum",
  "Own Room",
  "Team #",
  "Sequoyah Tee Time",
  "Tourney Tee Time"
];

const CHAT_HEADERS = [
  "Timestamp",
  "User",
  "Channel",
  "Message",
  "MessageID",
  "ParentID"
];

const ITINERARY_HEADERS = [
  "date",
  "time",
  "title",
  "location",
  "notes",
  "type",
  "image"
];

const LEADERBOARD_HEADERS = [
  "Team #",
  "Total",
  "Thru",
  "Strokes",
  "Base Score",
  "Badges Owned"
];

const BONUS_BADGES = [
  { key: "fireball3", label: "Fireball 3", hole: 3 },
  { key: "fireball5", label: "Fireball 5", hole: 5 },
  { key: "fireball7", label: "Fireball 7", hole: 7 },
  { key: "ctp9", label: "9 CTP", hole: 9, detailHeader: "9 CTP Yardage" },
  { key: "fireball10", label: "Fireball 10", hole: 10 },
  { key: "fireball12", label: "Fireball 12", hole: 12 },
  { key: "longDrive14", label: "14 Long Drive", hole: 14, detailHeader: "14 Long Drive Yardage" },
  { key: "fireball16", label: "Fireball 16", hole: 16 },
  { key: "happy18", label: "18 Happy", hole: 18, detailHeader: "18 Happy Yardage" }
];

const TEAM_SCORE_HEADERS = [
  "Timestamp",
  "Team #",
  "Updated By",
  "Hole 1",
  "Hole 2",
  "Hole 3",
  "Hole 4",
  "Hole 5",
  "Hole 6",
  "Hole 7",
  "Hole 8",
  "Hole 9",
  "Hole 10",
  "Hole 11",
  "Hole 12",
  "Hole 13",
  "Hole 14",
  "Hole 15",
  "Hole 16",
  "Hole 17",
  "Hole 18",
  "Thru",
  "Strokes",
  ...BONUS_BADGES.map(badge => badge.label),
  ...BONUS_BADGES.filter(badge => badge.detailHeader).map(badge => badge.detailHeader)
];

const BOOKING_HEADERS = [
  "ITEM",
  "LATEST UPDATE",
  "PRICE/PP",
  "Made Contact",
  "Confirmed Availability",
  "Secured",
  "Contract?",
  "Deposit?",
  "CANCEL BY DATE?",
  "TOTAL LIAB",
  "Contract"
];

const GALLERY_HEADERS = [
  "Timestamp",
  "Year",
  "Uploaded By",
  "Caption",
  "Image URL",
  "File ID",
  "Source"
];

const GALLERY_FOLDER_PROPERTY = "GALLERY_FOLDER_ID";
const GALLERY_FOLDER_NAME = "Cherokee Invitational Gallery Uploads";
const MAX_GALLERY_UPLOAD_BYTES = 8 * 1024 * 1024;

function saveEntry(name, email, handicap, practice, tournament, sportsbook, hotel, ownRoom) {
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
    yesNo_(sportsbook),
    yesNo_(hotel),
    yesNo_(ownRoom)
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

function getGolfersRows() {
  const sheet = getSheet_("Entries", ENTRY_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  return sheet.getRange(2, 1, lastRow - 1, ENTRY_HEADERS.length)
    .getDisplayValues()
    .map(row => {
      const commitments = [];
      if (isYes_(row[4])) commitments.push("Practice Round");
      if (isYes_(row[5])) commitments.push("Tournament");
      if (isYes_(row[6])) commitments.push("SportsBook");
      if (isYes_(row[7])) commitments.push("Hotel");

      return {
        name: cleanText_(row[1], 80),
        handicap: cleanText_(row[3], 20),
        commitments: commitments
      };
    })
    .filter(row => row.name);
}

function getItineraryEvents() {
  const sheet = getSheet_("Itinerary", ITINERARY_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  return sheet.getRange(2, 1, lastRow - 1, ITINERARY_HEADERS.length)
    .getValues()
    .map(row => ({
      date: formatDateValue_(row[0]),
      time: cleanText_(row[1], 40),
      title: cleanText_(row[2], 120),
      location: cleanText_(row[3], 180),
      notes: cleanText_(row[4], 260),
      type: cleanText_(row[5], 60),
      image: cleanText_(row[6], 500)
    }))
    .filter(event => event.title);
}

function getLeaderboardRows() {
  const sheet = getSheet_("Leaderboard", LEADERBOARD_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  return sheet.getRange(2, 1, lastRow - 1, LEADERBOARD_HEADERS.length)
    .getValues()
    .map(row => ({
      team: cleanText_(row[0], 80),
      total: cleanText_(row[1], 20),
      thru: cleanText_(row[2], 20),
      strokes: cleanText_(row[3], 20),
      baseScore: cleanText_(row[4], 20),
      badgesOwned: cleanText_(row[5], 200)
    }))
    .filter(row => row.team || row.total || row.thru || row.strokes || row.baseScore || row.badgesOwned);
}

function getTeamScore(teamNumber) {
  const cleanTeamNumber = cleanTeamNumber_(teamNumber);
  if (!cleanTeamNumber) {
    throw new Error("Unknown team number.");
  }

  const sheet = getSheet_("Team Scores", TEAM_SCORE_HEADERS);
  const rowIndex = findTeamScoreRow_(sheet, cleanTeamNumber);
  if (!rowIndex) {
    return {
      teamNumber: cleanTeamNumber,
      updatedBy: "",
      scores: Array(18).fill(""),
      badges: {},
      thru: "0",
      strokes: "0",
      updatedAt: ""
    };
  }

  const row = sheet.getRange(rowIndex, 1, 1, TEAM_SCORE_HEADERS.length).getDisplayValues()[0];
  return {
    teamNumber: cleanTeamNumber,
    updatedBy: cleanText_(row[2], 80),
    scores: row.slice(3, 21).map(value => cleanText_(value, 4)),
    badges: readBadgeValues_(row),
    thru: cleanText_(row[21], 10) || "0",
    strokes: cleanText_(row[22], 10) || "0",
    updatedAt: cleanText_(row[0], 80)
  };
}

function saveTeamScore(teamNumber, scores, updatedBy, badges) {
  const cleanTeamNumber = cleanTeamNumber_(teamNumber);
  if (!cleanTeamNumber) {
    throw new Error("Unknown team number.");
  }

  const cleanUpdatedBy = cleanText_(updatedBy, 80);
  const cleanScores = normalizeScores_(scores);
  const cleanBadges = normalizeBadges_(badges);
  const badgesOwned = badgeLabels_(cleanBadges).join(", ");
  const thru = cleanScores.filter(score => score !== "").length;
  const strokes = cleanScores.reduce((sum, score) => sum + (score === "" ? 0 : Number(score)), 0);
  const lock = LockService.getDocumentLock();

  lock.waitLock(10000);
  try {
    const scoreSheet = getSheet_("Team Scores", TEAM_SCORE_HEADERS);
    const scoreRow = [
      new Date(),
      cleanTeamNumber,
      cleanUpdatedBy,
      ...cleanScores,
      thru,
      strokes,
      ...BONUS_BADGES.map(badge => cleanBadges[badge.key].earned ? "Yes" : ""),
      ...BONUS_BADGES.filter(badge => badge.detailHeader).map(badge => cleanBadges[badge.key].yardage)
    ];
    const rowIndex = findTeamScoreRow_(scoreSheet, cleanTeamNumber);

    if (rowIndex) {
      scoreSheet.getRange(rowIndex, 1, 1, TEAM_SCORE_HEADERS.length).setValues([scoreRow]);
    } else {
      scoreSheet.appendRow(scoreRow);
    }

    updateLeaderboardScore_(cleanTeamNumber, thru, strokes, badgesOwned);
  } finally {
    lock.releaseLock();
  }

  return {
    teamNumber: cleanTeamNumber,
    updatedBy: cleanUpdatedBy,
    scores: cleanScores,
    badges: cleanBadges,
    thru: String(thru),
    strokes: String(strokes)
  };
}

function getBookingRows() {
  const sheet = getSheet_("BOOKING", BOOKING_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { rows: [], summary: {} };

  const rowCount = lastRow - 1;
  const rawValues = sheet.getRange(2, 1, rowCount, BOOKING_HEADERS.length).getValues();
  const displayValues = sheet.getRange(2, 1, rowCount, BOOKING_HEADERS.length).getDisplayValues();
  const contractLinks = sheet.getRange(2, 11, rowCount, 1).getRichTextValues();
  let summary = {};

  const rows = displayValues
    .map((displayRow, index) => {
      const rawRow = rawValues[index];
      const contractRichText = contractLinks[index][0];
      const item = cleanText_(displayRow[0], 120);
      const row = {
        item: item,
        latestUpdate: cleanText_(displayRow[1], 500),
        pricePerPerson: cleanText_(displayRow[2], 40),
        madeContact: isChecked_(rawRow[3]),
        confirmedAvailability: isChecked_(rawRow[4]),
        secured: isChecked_(rawRow[5]),
        contractSigned: isChecked_(rawRow[6]),
        depositPaid: isChecked_(rawRow[7]),
        cancelBy: cleanText_(displayRow[8], 120),
        totalLiability: cleanText_(displayRow[9], 60),
        contractUrl: cleanText_(contractRichText.getLinkUrl() || displayRow[10], 500)
      };

      if (item.toUpperCase() === "TOTAL") {
        summary = {
          pricePerPerson: row.pricePerPerson,
          madeContact: cleanText_(displayRow[3], 20),
          confirmedAvailability: cleanText_(displayRow[4], 20),
          secured: cleanText_(displayRow[5], 20),
          contractSigned: cleanText_(displayRow[6], 20),
          depositPaid: cleanText_(displayRow[7], 20),
          totalLiability: row.totalLiability
        };
        return null;
      }

      return row;
    })
    .filter(row => row && row.item);

  return {
    rows: rows,
    summary: summary
  };
}

function getGalleryPhotos() {
  const sheet = getSheet_("Gallery", GALLERY_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  return sheet.getRange(2, 1, lastRow - 1, GALLERY_HEADERS.length)
    .getDisplayValues()
    .map(row => {
      const fileId = cleanText_(row[5], 120);
      return {
        timestamp: cleanText_(row[0], 80),
        year: cleanText_(row[1], 20),
        uploadedBy: cleanText_(row[2], 80),
        caption: cleanText_(row[3], 160),
        imageUrl: cleanText_(row[4], 500) || galleryImageUrl_(fileId),
        fileId: fileId,
        source: cleanText_(row[6], 40)
      };
    })
    .filter(photo => photo.imageUrl)
    .reverse();
}

function uploadGalleryPhoto(payload) {
  const data = payload || {};
  const uploadedBy = cleanText_(data.uploadedBy, 80);
  const year = cleanText_(data.year, 20) || "2026";
  const caption = cleanText_(data.caption, 160);
  const dataUrl = String(data.dataUrl || "");
  const match = dataUrl.match(/^data:(image\/(?:jpeg|jpg|png|gif|webp));base64,(.+)$/i);

  if (!uploadedBy) {
    throw new Error("Add your name before uploading a photo.");
  }
  if (!match) {
    throw new Error("Upload a JPG, PNG, GIF, or WebP image.");
  }

  const mimeType = match[1].toLowerCase() === "image/jpg" ? "image/jpeg" : match[1].toLowerCase();
  const base64 = match[2];
  const estimatedBytes = Math.ceil(base64.length * 0.75);
  if (estimatedBytes > MAX_GALLERY_UPLOAD_BYTES) {
    throw new Error("Photo is too large. Please upload an image under 8 MB.");
  }

  const extension = imageExtension_(mimeType);
  const bytes = Utilities.base64Decode(base64);
  const filename = [
    "cherokee",
    year.replace(/[^0-9a-z-]/gi, ""),
    uploadedBy.replace(/[^0-9a-z-]/gi, "-").replace(/-+/g, "-"),
    Utilities.getUuid().slice(0, 8)
  ].filter(Boolean).join("-") + "." + extension;

  const folder = getGalleryFolder_();
  const file = folder.createFile(Utilities.newBlob(bytes, mimeType, filename));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const fileId = file.getId();
  const imageUrl = galleryImageUrl_(fileId);
  const sheet = getSheet_("Gallery", GALLERY_HEADERS);
  sheet.appendRow([
    new Date(),
    year,
    uploadedBy,
    caption,
    imageUrl,
    fileId,
    "Upload"
  ]);

  return {
    year: year,
    uploadedBy: uploadedBy,
    caption: caption,
    imageUrl: imageUrl,
    fileId: fileId,
    source: "Upload"
  };
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
  } else {
    const existingHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    headers.forEach((header, index) => {
      if (!existingHeaders[index]) {
        sheet.getRange(1, index + 1).setValue(header);
      }
    });
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

function isChecked_(value) {
  return value === true || String(value).toUpperCase() === "TRUE" || String(value).trim() === "1";
}

function isYes_(value) {
  return String(value || "").trim().toUpperCase() === "YES";
}

function cleanTeamNumber_(value) {
  const text = String(value || "").trim();
  const match = text.match(/^\d{1,2}$/);
  if (!match) return "";
  const numeric = Number(text);
  return numeric >= 1 && numeric <= 99 ? String(numeric) : "";
}

function normalizeScores_(scores) {
  const source = Array.isArray(scores) ? scores : [];
  return Array.from({ length: 18 }, (_, index) => {
    const value = String(source[index] || "").trim();
    if (!value) return "";

    const numeric = Number(value);
    if (!Number.isInteger(numeric) || numeric < 1 || numeric > 20) {
      throw new Error("Hole " + (index + 1) + " must be a whole number from 1 to 20.");
    }

    return numeric;
  });
}

function cleanYardage_(value) {
  const text = String(value || "").trim().replace(/yards?|yds?\.?/i, "").trim();
  if (!text) return "";

  const numeric = Number(text);
  if (!Number.isFinite(numeric) || numeric < 0 || numeric > 700) {
    throw new Error("Badge yardage must be a number from 0 to 700.");
  }

  return String(Math.round(numeric * 10) / 10);
}

function normalizeBadges_(badges) {
  const source = badges && typeof badges === "object" ? badges : {};
  return BONUS_BADGES.reduce((result, badge) => {
    const value = source[badge.key];
    const valueObject = value && typeof value === "object" ? value : {};
    const earned = value === true || value === "true" || value === "Yes" || valueObject.earned === true || valueObject.earned === "true";
    result[badge.key] = {
      earned: earned,
      yardage: badge.detailHeader && earned ? cleanYardage_(valueObject.yardage) : ""
    };
    return result;
  }, {});
}

function readBadgeValues_(row) {
  const detailBadges = BONUS_BADGES.filter(badge => badge.detailHeader);
  return BONUS_BADGES.reduce((result, badge, index) => {
    const detailIndex = detailBadges.findIndex(detailBadge => detailBadge.key === badge.key);
    result[badge.key] = {
      earned: isYes_(row[23 + index]),
      yardage: detailIndex === -1 ? "" : cleanText_(row[23 + BONUS_BADGES.length + detailIndex], 20)
    };
    return result;
  }, {});
}

function badgeLabels_(badges) {
  return BONUS_BADGES
    .filter(badge => badges[badge.key] && badges[badge.key].earned)
    .map(badge => {
      const yardage = badges[badge.key].yardage;
      return yardage ? badge.label + " (" + yardage + " yds)" : badge.label;
    });
}

function findTeamScoreRow_(sheet, teamNumber) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const values = sheet.getRange(2, 2, lastRow - 1, 1).getDisplayValues();
  const index = values.findIndex(row => cleanTeamNumber_(row[0]) === teamNumber);
  return index === -1 ? 0 : index + 2;
}

function updateLeaderboardScore_(teamNumber, thru, strokes, badgesOwned) {
  const sheet = getSheet_("Leaderboard", LEADERBOARD_HEADERS);
  const lastRow = sheet.getLastRow();
  let targetRow = 0;

  if (lastRow >= 2) {
    const teams = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
    const index = teams.findIndex(row => cleanLeaderboardTeamNumber_(row[0]) === teamNumber);
    targetRow = index === -1 ? 0 : index + 2;
  }

  if (!targetRow) {
    sheet.appendRow(["Team " + teamNumber, "", thru, strokes, strokes, badgesOwned]);
    return;
  }

  sheet.getRange(targetRow, 3).setValue(thru);
  sheet.getRange(targetRow, 4).setValue(strokes);
  sheet.getRange(targetRow, 5).setValue(strokes);
  sheet.getRange(targetRow, 6).setValue(badgesOwned);
}

function cleanLeaderboardTeamNumber_(value) {
  const text = String(value || "").trim();
  const match = text.match(/\d{1,2}/);
  return match ? cleanTeamNumber_(match[0]) : "";
}

function getGalleryFolder_() {
  const properties = PropertiesService.getScriptProperties();
  const existingId = properties.getProperty(GALLERY_FOLDER_PROPERTY);
  if (existingId) {
    try {
      return DriveApp.getFolderById(existingId);
    } catch (e) {
      properties.deleteProperty(GALLERY_FOLDER_PROPERTY);
    }
  }

  const folder = DriveApp.createFolder(GALLERY_FOLDER_NAME);
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  properties.setProperty(GALLERY_FOLDER_PROPERTY, folder.getId());
  return folder;
}

function galleryImageUrl_(fileId) {
  const cleanFileId = cleanText_(fileId, 120);
  return cleanFileId ? "https://drive.google.com/thumbnail?id=" + encodeURIComponent(cleanFileId) + "&sz=w1400" : "";
}

function imageExtension_(mimeType) {
  const cleanMimeType = String(mimeType || "").toLowerCase();
  if (cleanMimeType === "image/png") return "png";
  if (cleanMimeType === "image/gif") return "gif";
  if (cleanMimeType === "image/webp") return "webp";
  return "jpg";
}

function formatTimeValue_(value) {
  if (value instanceof Date) {
    return value.toISOString();
  }
  return cleanText_(value, 80);
}

function formatDateValue_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "M/d/yyyy");
  }
  return cleanText_(value, 40);
}
