// ══════════════════════════════════════════════════════
// Renova Quiz — Google Apps Script
// 1. צור Google Sheet חדש
// 2. העתק את ה-ID של הגיליון מה-URL
//    (החלק בין /d/ ל-/edit)
// 3. הדבק אותו ב-SHEET_ID למטה
// 4. Deploy → New deployment → Web app
//    Execute as: Me | Who has access: Anyone
// 5. העתק את ה-URL שקיבלת ל-index.html שורה 1
// ══════════════════════════════════════════════════════

const SHEET_ID = 'YOUR_GOOGLE_SHEET_ID_HERE'; // ← החלף

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('לידים');

    // צור גיליון אם לא קיים
    if (!sheet) {
      sheet = ss.insertSheet('לידים');
      const headers = [
        'תאריך ושעה','שם','טלפון','מייל',
        'ציפורניים','רגל','הפרעה 1-10','משך הפטרת',
        'מטרות','טיפולים שנוסו',
        'ידע שציפורן קשה','ידע על פצירה',
        'אורח חיים','מחויבות 1-10',
        'חשיבות קיץ','מוכן לנסות עם ערבות',
        'ניקוד','חבילה מומלצת','לחץ לרכישה'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setBackground('#0a6e6e').setFontColor('#ffffff').setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    // תרגומים לעברית
    const legMap   = { right:'ימין', left:'שמאל', both:'שתיים' };
    const durMap   = { '<1':'פחות משנה','1-3':'1-3 שנים','3-5':'3-5 שנים','5+':'5+ שנים' };
    const lifeMap  = {
      closed_shoes:'נעליים סגורות', gym:'חדר כושר', pool:'בריכה/ים',
      diabetic:'סוכרתי', military:'מילואים'
    };
    const goalsMap = {
      beach:'ים וחוף', sandals:'סנדלים', pedicure:'פדיקור',
      socks:'ללא גרביים', partner:'בן/בת זוג'
    };
    const triedMap = {
      nail_polish:'לק טיפולי', tea_tree:'שמן עץ תה', pills:'כדורים',
      laser:'לייזר', medical_pedicure:'פדיקור רפואי', nothing:'לא ניסה'
    };

    function translateList(str, map) {
      return (str||'').split(', ').filter(Boolean).map(v => map[v]||v).join(', ');
    }

    const row = [
      new Date(data.timestamp || Date.now()).toLocaleString('he-IL'),
      data.name  || '',
      data.phone || '',
      data.email || '',
      data.q1_nails || '',
      legMap[data.q2_leg] || data.q2_leg || '',
      data.q3_bothers || '',
      durMap[data.q4_duration] || data.q4_duration || '',
      translateList(data.q5_goals, goalsMap),
      translateList(data.q7_tried, triedMap),
      data.q8_knew_hard   === 'no' ? 'לא ידע' : 'ידע',
      data.q9_knew_filing === 'no' ? 'לא ידע' : 'ידע',
      lifeMap[data.q10_lifestyle] || data.q10_lifestyle || '',
      data.q11_commitment || '',
      data.q12_summer    === 'yes' ? 'כן' : 'לא',
      data.q13_guarantee === 'yes' ? 'כן' : 'לא',
      data.score || 0,
      data.recommended || '',
      data.clicked_purchase ? 'כן' : 'לא'
    ];

    sheet.appendRow(row);

    // עצב שורה לפי חבילה מומלצת
    const lastRow = sheet.getLastRow();
    const colors = { '3 חודשים':'#e8f5e9', '6 חודשים':'#e3f2fd', '9 חודשים':'#fff3e0' };
    const bg = colors[data.recommended] || '#ffffff';
    sheet.getRange(lastRow, 1, 1, row.length).setBackground(bg);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', row: lastRow }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// GET — לבדיקה שה-script עובד
function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'Renova Quiz Script פעיל' }))
    .setMimeType(ContentService.MimeType.JSON);
}
