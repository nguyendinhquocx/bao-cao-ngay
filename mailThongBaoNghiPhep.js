/**
 * MAIL THONG BAO NGHI PHEP
 *
 * Tu dong gui email thong bao nhan vien dang ky nghi phep/cong tac
 *
 * SHEET STRUCTURE:
 * Column A: ma nhan vien
 * Column B: ten nhan vien
 * Column C: date
 * Column D: thu
 * Column E: check
 * Column F: nghi phep or cong tac (dropdown: "Nghi phep sang", "Nghi phep chieu", "Nghi phep ca ngay", "Cong tac sang", "Cong tac chieu", "Cong tac ca ngay")
 * Column G: note (ly do)
 * Column H: x (danh dau da gui email - tu dong mark sau khi gui)
 *
 * FLOW:
 * 1. Scan 10 spreadsheets qua IMPORTRANGE
 * 2. Tim records co column F khong rong (da dang ky nghi phep/cong tac)
 * 3. Tim tat ca records co column F khong rong NHUNG column H = rong (chua gui email)
 * 4. Gui email tong hop
 * 5. Danh dau 'x' vao column H sau khi gui thanh cong
 *
 * FEATURES:
 * - Multi-spreadsheet aggregation (10 sheets)
 * - Auto sync column H based on column F
 * - Phan biet "Nghi phep" vs "Cong tac"
 * - HTML email template dep
 * - Daily summary
 *
 * @version 1.0
 * @author Nguyen Dinh Quoc
 * @updated 2025-01-21
 *
 * USAGE:
 * sendLeaveNotification() - Gui thong bao nghi phep cho ngay hien tai
 */
function sendLeaveNotification() {
  const CONFIG = {
    // 10 Spreadsheet IDs
    spreadsheetIds: [
      '1tcMDkaxhHIvoBAi1yzcnZddD46nm56mfFB7A7aLt5ek',
      '1L1wIc5gVEh4hz1mfVorGF48vfEGQnZ6VBXGcndxRlwA',
      '10CJ8qC028-CbX1UjmI-HcbPkQxRoiWFKPXlDT8f5uSY',
      '15ZO4BGyOkHMyPIttE4BBVvaZhJbmqTTJ5jKWsHLfQfg',
      '13ACyogQDDBRw9QEZc0YQ42Qis5Z_yVNOu7WmS67TWVw',
      '1s2KXuk9Kph_6vFUUrutBUZMamikyyR00-C05bQxXW6k',
      '1n6iq0G2aC6rzIJ7Ir7P6Z4LmBqXiA6jKhfh7R5axxKE',
      '1mij4KC6yZ8joioMpcj1-rGXkhL6gBYByHryjRh0edy4',
      '19kqn8JcKp3TzdrwoUEiJUmyUInmk2Jv9fhMY-D01ktc',
      '1ZZ47Rf5aAV5ixHpAWrbGzBfhFXSjTQQTsRXiwWe3ydI',
      '1D1bPi44OL8skQJW0mqhqTcfCAWM4ZvXOUqa-OW-2viw'
    ],

    sheetName: 'tick',
    emailTo: ['quoc.nguyen3@hoanmy.com'],
    // emailTo: ['quoc.nguyen3@hoanmy.com', 'luan.tran@hoanmy.com'],

    // Leave types (CO DAU - phai khop voi dropdown trong sheet)
    leaveTypes: {
      nghiPhep: ['Nghỉ phép sáng', 'Nghỉ phép chiều', 'Nghỉ phép cả ngày'],
      congTac: ['Công tác sáng', 'Công tác chiều', 'Công tác cả ngày']
    },

    debugMode: false
  };

  try {
    // Step 1: Load all leave data from 10 spreadsheets
    const allLeaveData = loadAllLeaveData(CONFIG);

    if (CONFIG.debugMode) {
      Logger.log(`Total leave records: ${allLeaveData.length}`);
    }

    // Step 2: Get unsent leave registrations (column F has value, column H = empty)
    const unsentLeave = getUnsentLeaveRegistrations(allLeaveData, CONFIG);

    if (CONFIG.debugMode) {
      Logger.log(`Future leave count: ${unsentLeave.nghiPhep.length + unsentLeave.congTac.length}`);
      Logger.log(`Past leave count: ${unsentLeave.pastLeave.length}`);
    }

    // Step 3: Auto-mark past leave (tu dong danh 'x' cho ngay qua khu, khong gui email)
    if (unsentLeave.pastLeave.length > 0) {
      markAsSent({ nghiPhep: unsentLeave.pastLeave, congTac: [] }, CONFIG);
      Logger.log(`Da tu dong danh dau ${unsentLeave.pastLeave.length} ngay qua khu`);
    }

    // Step 4: Send email for future/today leave
    if (unsentLeave.nghiPhep.length > 0 || unsentLeave.congTac.length > 0) {
      sendLeaveEmail(unsentLeave, CONFIG);

      // Step 5: Mark as sent (column H = 'x')
      markAsSent(unsentLeave, CONFIG);

      Logger.log(`Email thong bao nghi phep da duoc gui`);
    } else {
      Logger.log(`Khong co dang ky nghi phep moi`);
    }

  } catch (error) {
    Logger.log(`Loi khi gui thong bao nghi phep: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
  }
}

/**
 * Load all leave data from 10 spreadsheets
 */
function loadAllLeaveData(CONFIG) {
  const allData = [];

  CONFIG.spreadsheetIds.forEach((ssId, index) => {
    try {
      const ss = SpreadsheetApp.openById(ssId);
      const sheet = ss.getSheetByName(CONFIG.sheetName);

      if (!sheet) {
        Logger.log(`Sheet '${CONFIG.sheetName}' khong ton tai trong spreadsheet ${index + 1}`);
        return;
      }

      const dataRange = sheet.getDataRange();
      const values = dataRange.getValues();

      if (values.length <= 1) {
        return; // Skip neu chi co header
      }

      const headers = values[0];

      // Parse data
      for (let i = 1; i < values.length; i++) {
        const row = values[i];
        const record = {
          spreadsheetId: ssId,
          sheetName: CONFIG.sheetName,
          rowIndex: i + 1, // 1-indexed for sheet
          employeeId: row[0],
          employeeName: row[1],
          date: row[2],
          dayOfWeek: row[3],
          check: row[4],
          leaveType: row[5], // Column F
          note: row[6],      // Column G
          mailSent: row[7]   // Column H
        };

        // Only include records with leave type
        if (record.leaveType) {
          allData.push(record);
        }
      }

      if (CONFIG.debugMode) {
        Logger.log(`Loaded ${allData.length} records from spreadsheet ${index + 1}`);
      }

    } catch (error) {
      Logger.log(`Loi khi load spreadsheet ${index + 1}: ${error.message}`);
    }
  });

  return allData;
}

/**
 * Get unsent leave registrations (column F has value, column H = empty)
 * Return:
 * - nghiPhep/congTac: ngay hom nay va sau -> gui email
 * - pastLeave: ngay truoc hom nay -> tu dong danh 'x', khong gui email
 */
function getUnsentLeaveRegistrations(leaveData, CONFIG) {
  const futureNghiPhep = [];
  const futureCongTac = [];
  const pastLeave = [];

  // Get today at 00:00:00 for date comparison
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  leaveData.forEach(record => {
    const hasLeaveType = record.leaveType && record.leaveType.trim() !== '';
    const isMailSent = record.mailSent === 'x' || record.mailSent === 'X';

    // Only process records with leaveType AND mailSent = empty
    if (hasLeaveType && !isMailSent) {
      // Parse record date
      const recordDate = record.date instanceof Date ? new Date(record.date) : new Date(record.date);
      recordDate.setHours(0, 0, 0, 0);

      const leaveInfo = {
        spreadsheetId: record.spreadsheetId,
        sheetName: record.sheetName,
        rowIndex: record.rowIndex,
        employeeName: record.employeeName,
        employeeId: record.employeeId,
        date: record.date,
        leaveType: record.leaveType,
        note: record.note || 'Khong co ghi chu'
      };

      if (recordDate >= today) {
        // Future/today: gui email
        if (CONFIG.leaveTypes.nghiPhep.includes(record.leaveType)) {
          futureNghiPhep.push(leaveInfo);
        } else if (CONFIG.leaveTypes.congTac.includes(record.leaveType)) {
          futureCongTac.push(leaveInfo);
        }
      } else {
        // Past: tu dong danh 'x', khong gui email
        pastLeave.push(leaveInfo);
      }
    }
  });

  // Sort by date
  const sortByDate = (a, b) => {
    const dateA = a.date instanceof Date ? a.date : new Date(a.date);
    const dateB = b.date instanceof Date ? b.date : new Date(b.date);
    return dateA - dateB;
  };

  futureNghiPhep.sort(sortByDate);
  futureCongTac.sort(sortByDate);
  pastLeave.sort(sortByDate);

  return {
    nghiPhep: futureNghiPhep,
    congTac: futureCongTac,
    pastLeave: pastLeave
  };
}

/**
 * Mark records as sent (column H = 'x')
 */
function markAsSent(leaveData, CONFIG) {
  const allRecords = [...leaveData.nghiPhep, ...leaveData.congTac];

  allRecords.forEach(record => {
    try {
      const ss = SpreadsheetApp.openById(record.spreadsheetId);
      const sheet = ss.getSheetByName(record.sheetName);
      const cell = sheet.getRange(record.rowIndex, 8); // Column H

      cell.setValue('x');

      if (CONFIG.debugMode) {
        Logger.log(`Marked mail sent for row ${record.rowIndex}`);
      }
    } catch (error) {
      Logger.log(`Loi khi mark column H: ${error.message}`);
    }
  });

  if (CONFIG.debugMode) {
    Logger.log(`Total marked as sent: ${allRecords.length}`);
  }
}

/**
 * Send leave notification email
 */
function sendLeaveEmail(todayLeave, CONFIG) {
  const today = new Date();
  const dayNames = ['Chu nhat', 'Thu hai', 'Thu ba', 'Thu tu', 'Thu nam', 'Thu sau', 'Thu bay'];
  const dayOfWeek = dayNames[today.getDay()];
  const detailedDate = `${dayOfWeek}, ngay ${today.getDate()} thang ${today.getMonth() + 1} nam ${today.getFullYear()}`;

  const subject = `HMSG | P.KD - THONG BAO NGHI PHEP & CONG TAC`;

  // Group by date first
  const allLeave = [...todayLeave.nghiPhep, ...todayLeave.congTac];
  const dateMap = {};

  allLeave.forEach(leave => {
    const dateObj = leave.date instanceof Date ? leave.date : new Date(leave.date);
    const dateKey = Utilities.formatDate(dateObj, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyy-MM-dd");

    if (!dateMap[dateKey]) {
      dateMap[dateKey] = {
        date: dateObj,
        leaves: []
      };
    }
    dateMap[dateKey].leaves.push(leave);
  });

  // Group dates by month
  const monthMap = {};

  Object.keys(dateMap).forEach(dateKey => {
    const dateGroup = dateMap[dateKey];
    const monthKey = Utilities.formatDate(dateGroup.date, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyy-MM");

    if (!monthMap[monthKey]) {
      monthMap[monthKey] = {
        month: dateGroup.date,
        dates: []
      };
    }

    monthMap[monthKey].dates.push({
      dateKey: dateKey,
      date: dateGroup.date,
      leaves: dateGroup.leaves
    });
  });

  // Sort months ascending (oldest first - xa den gan)
  const sortedMonths = Object.keys(monthMap).sort((a, b) => {
    return a.localeCompare(b); // Ascending yyyy-MM
  });

  // Build HTML grouped by month
  let dateHtml = '';

  sortedMonths.forEach(monthKey => {
    const monthGroup = monthMap[monthKey];
    const monthStr = Utilities.formatDate(monthGroup.month, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "MM/yyyy");

    // Sort dates within month ascending (oldest first - xa den gan)
    monthGroup.dates.sort((a, b) => a.date - b.date);

    // Build date sections for this month
    let monthDatesHtml = '';

    monthGroup.dates.forEach(dateGroup => {
      const dateStr = Utilities.formatDate(dateGroup.date, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd/MM/yyyy");
      const leaveCount = dateGroup.leaves.length;

      // Sort employees within each date alphabetically
      dateGroup.leaves.sort((a, b) => a.employeeName.localeCompare(b.employeeName));

      // Build leave items for this date
      const leaveItems = dateGroup.leaves.map(leave => {
        return `
          <div style="padding: 8px 0; border-bottom: 1px solid #f5f5f5;">
            <div style="font-size: 14px; font-weight: 500; color: #1a1a1a; margin-bottom: 4px;">
              ${leave.employeeName}
            </div>
            <div style="font-size: 13px; color: #8e8e93; margin-bottom: 4px;">
              ${leave.leaveType}
            </div>
            <div style="font-size: 13px; color: #495057; font-style: italic;">
              ${leave.note}
            </div>
          </div>
        `;
      }).join('');

      monthDatesHtml += `
        <div style="margin-bottom: 20px;">
          <div style="font-size: 15px; font-weight: 500; color: #1a1a1a; margin-bottom: 12px;">
            ${dateStr} (${leaveCount})
          </div>
          <div style="padding-left: 16px;">
            ${leaveItems}
          </div>
        </div>
      `;
    });

    // Add month header and dates
    dateHtml += `
      <div style="margin-bottom: 32px;">
        <div style="font-size: 16px; font-weight: 600; color: #1a1a1a; margin-bottom: 16px; padding-bottom: 8px; border-bottom: 2px solid #e0e0e0;">
          Thang ${monthStr}
        </div>
        <div style="padding-left: 12px;">
          ${monthDatesHtml}
        </div>
      </div>
    `;
  });

  // HTML Email Template
  const htmlBody = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Thong bao nghi phep va cong tac</title>
    </head>
    <body style="margin: 0; padding: 0; background-color: #ffffff; font-family: Arial, sans-serif;">

      <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color: #ffffff;">
        <tr>
          <td style="padding: 20px;">
            <table width="600" cellpadding="0" cellspacing="0" border="0" style="max-width: 600px; margin: 0 auto;" class="container">
              <tr>
                <td style="padding: 20px;">

        <!-- Date Sections -->
        <div style="margin-bottom: 40px;">
          ${dateHtml}
        </div>

        <!-- Link to Dashboard -->
        <div style="margin-bottom: 24px;">
          <a href="https://docs.google.com/spreadsheets/d/15eMfEvqNvy1qBNG1NXwr7eSBsYZA6KqlBB3lTyzTfhM/edit?gid=1077073904#gid=1077073904"
             style="font-size: 13px; color: #1a1a1a; text-decoration: underline;">
            Xem tong quat
          </a>
        </div>

        <!-- Footer -->
        <div style="text-align: left; padding-top: 8px;">
          <p style="margin: 0; font-size: 12px; font-weight: 400; color: #1a1a1a;">
            Tran trong
          </p>
        </div>

                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>

    </body>
    </html>
  `;

  // Send email
  try {
    const recipients = Array.isArray(CONFIG.emailTo) ? CONFIG.emailTo.join(',') : CONFIG.emailTo;

    GmailApp.sendEmail(
      recipients,
      subject,
      '', // Plain text body
      {
        htmlBody: htmlBody,
        name: "Thong Bao Nghi Phep"
      }
    );
    Logger.log(`Email sent successfully to: ${recipients}`);
  } catch (error) {
    Logger.log(`Email send failed: ${error.message}`);
  }
}

/**
 * TEST FUNCTION - Test voi debug mode
 */
function testLeaveNotification() {
  Logger.log('TESTING LEAVE NOTIFICATION');

  // Enable debug mode temporarily
  const CONFIG = {
    spreadsheetIds: [
      '1tcMDkaxhHIvoBAi1yzcnZddD46nm56mfFB7A7aLt5ek'
      // Test voi 1 spreadsheet truoc
    ],
    sheetName: 'tick',
    emailTo: ['quoc.nguyen3@hoanmy.com'],
    leaveTypes: {
      nghiPhep: ['Nghỉ phép sáng', 'Nghỉ phép chiều', 'Nghỉ phép cả ngày'],
      congTac: ['Công tác sáng', 'Công tác chiều', 'Công tác cả ngày']
    },
    debugMode: true
  };

  const allLeaveData = loadAllLeaveData(CONFIG);
  Logger.log(`Sample leave data:`, allLeaveData.slice(0, 5));

  Logger.log('Test completed');
}
