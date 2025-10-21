/**
 * RAW DATA VERSION: Gửi email báo cáo tổng hợp từ raw data trong sheet 'tick'
 *
 * SHEET STRUCTURE:
 * Column A: ma nhan vien
 * Column B: ten nhan vien
 * Column C: date
 * Column D: thu
 * Column E: check
 *
 * EXCLUDED EMPLOYEES:
 * - Cấu hình trong CONFIG.excludedEmployees bằng mã nhân viên
 * - Nhân viên trong list sẽ KHÔNG xuất hiện trong email báo cáo
 * - Data vẫn giữ nguyên trong sheet, chỉ filter khi tạo báo cáo
 *
 * FEATURES:
 * ✅ Use raw transactional data từ sheet 'tick' thay vì processed data 'check bc'
 * ✅ Flexible data querying với date ranges
 * ✅ Weekly stars calculation từ raw data
 * ✅ Custom date support
 * ✅ Mobile responsive email template
 * ✅ Medal system với HTML entities
 * ✅ Excluded employees filter by employee ID
 *
 * @version 3.2 Excluded Employees Support
 * @author Nguyễn Đình Quốc
 * @updated 2025-01-21
 *
 * @param {string|Date} customDate - Ngày tuỳ chọn (format: 'YYYY-MM-DD' hoặc Date object). Nếu không truyền thì dùng ngày hiện tại
 *
 * USAGE:
 * sendDailyReportSummaryRaw() - Gửi báo cáo ngày hiện tại
 * sendDailyReportSummaryRaw('2025-07-15') - Gửi báo cáo ngày 15/7/2025
 * sendDailyReportSummaryRaw(new Date('2025-07-15')) - Gửi báo cáo ngày 15/7/2025
 *
 * EXCLUDE EMPLOYEES:
 * 1. Mở file baoCaoNgayRaw.js
 * 2. Tìm CONFIG.excludedEmployees (dòng ~40)
 * 3. Thêm mã nhân viên vào array: ['004620', '005123', ...]
 */
function sendDailyReportSummaryRaw(customDate = null) {
  const CONFIG = {
    sheetName: 'tick', // Changed to raw data sheet

    // Uncomment khi deploy production
    // emailTo: 'luan.tran@hoanmy.com, khanh.tran@hoanmy.com, hong.le@hoanmy.com, quynh.bui@hoanmy.com, thuy.pham@hoanmy.com, anh.ngo@hoanmy.com, truc.nguyen3@hoanmy.com, trang.nguyen9@hoanmy.com, tram.mai@hoanmy.com, vuong.duong@hoanmy.com, hoang.vo4@hoanmy.com, phi.tran@hoanmy.com, quoc.nguyen3@hoanmy.com',
    emailTo: 'quoc.nguyen3@hoanmy.com',

    // EXCLUDED EMPLOYEES - Không tính vào báo cáo
    // Thêm mã nhân viên vào array này để loại bỏ khỏi email (vẫn giữ data trong sheet)
    excludedEmployees: [
      '004620'  // Trần Thị Phương Phi - không tính vào báo cáo
      // Thêm các mã nhân viên khác nếu cần, VD: '005123', '006456'
    ],

    // Raw data column mapping - Simplified structure
    columns: {
      employeeId: 'ma nhan vien',        // Column A
      employeeName: 'ten nhan vien',     // Column B
      date: 'date',                      // Column C
      dayOfWeek: 'thu',                  // Column D
      check: 'check'                     // Column E
    },

    // ICON mặc định (đen/xám)
    starIconDefault: 'https://cdn-icons-png.flaticon.com/128/2956/2956792.png',
    calendarIconDefault: 'https://cdn-icons-png.flaticon.com/128/3239/3239948.png',
    completedIconDefault: 'https://cdn-icons-png.flaticon.com/128/7046/7046053.png',
    pendingIconDefault: 'https://cdn-icons-png.flaticon.com/128/17694/17694317.png',

    // ICON màu xanh khi perfect day
    starIconPerfect: 'https://cdn-icons-png.flaticon.com/128/18245/18245248.png',
    calendarIconPerfect: 'https://cdn-icons-png.flaticon.com/128/15881/15881446.png',
    completedIconPerfect: 'https://cdn-icons-png.flaticon.com/128/10995/10995390.png',
    pendingIconPerfect: 'https://cdn-icons-png.flaticon.com/128/17694/17694222.png',

    // Achievement icons
    celebrationIcon: 'https://cdn-icons-png.flaticon.com/128/9422/9422222.png',

    // DEBUG MODE
    debugMode: false // Set true để troubleshoot
  };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.sheetName);

    if (!sheet) {
      Logger.log(`❌ Sheet '${CONFIG.sheetName}' không tồn tại`);
      return;
    }

    // Parse custom date or use current date
    const targetDate = parseTargetDate(customDate);
    const today = new Date(); // Keep for reference
    const targetDateStr = Utilities.formatDate(targetDate, ss.getSpreadsheetTimeZone(), "M/d/yyyy");
    const isWeekend = targetDate.getDay() === 0; // Chủ nhật
    const isCustomDate = customDate !== null;

    // Định dạng ngày chi tiết
    const dayNames = ['Chủ nhật', 'Thứ hai', 'Thứ ba', 'Thứ tư', 'Thứ năm', 'Thứ sáu', 'Thứ bảy'];
    const dayOfWeek = dayNames[targetDate.getDay()];
    const detailedDate = `${dayOfWeek}, ngày ${targetDate.getDate()} tháng ${targetDate.getMonth() + 1} năm ${targetDate.getFullYear()}`;

    if (CONFIG.debugMode) {
      Logger.log(`🎯 Target date: ${targetDateStr} (${isCustomDate ? 'Custom' : 'Current'})`);
      Logger.log(`📅 Detailed date: ${detailedDate}`);
    }

    // Load va parse raw data tu sheet
    const rawData = loadRawDataFromSheet(sheet, CONFIG);

    // Build date index once for performance optimization
    const dateIndex = buildDateIndexRaw(rawData, ss);

    // Get employees who reported on target date
    const targetReports = getEmployeeReportsForDate(rawData, targetDate, ss);
    const reported = targetReports.reported;
    const notReported = targetReports.notReported;

    // Kiểm tra perfect day và tính totals
    const totalEmployees = reported.length + notReported.length;
    const isPerfectDay = notReported.length === 0 && reported.length > 0;
    const subject = isWeekend ?
      `HMSG | P.KD - THỐNG KÊ TUẦN` :
      `HMSG | P.KD - TỔNG HỢP BÁO CÁO NGÀY ${targetDateStr}${isCustomDate ? ' ⭐' : ''}`;

    // Chọn icons theo trạng thái
    const calendarIcon = isPerfectDay ? CONFIG.calendarIconPerfect : CONFIG.calendarIconDefault;
    const completedIcon = isPerfectDay ? CONFIG.completedIconPerfect : CONFIG.completedIconDefault;
    const pendingIcon = isPerfectDay ? CONFIG.pendingIconPerfect : CONFIG.pendingIconDefault;

    // Color scheme
    const colors = isPerfectDay ? {
      border: '#22c55e',
      headerTitle: '#22c55e',
      headerSubtitle: '#22c55e',
      dateText: '#22c55e',
      sectionTitle: '#22c55e',
      namesList: '#22c55e',
      footerName: '#22c55e',
      footerLabel: '#22c55e', // Xanh khi perfect day
      disclaimerColor: '#22c55e'
    } : {
      border: '#000000',
      headerTitle: '#1a1a1a',
      headerSubtitle: '#8e8e93',
      dateText: '#495057',
      sectionTitle: '#1a1a1a',
      pendingTitle: '#dc3545',
      namesList: '#1a1a1a',
      footerName: '#8e8e93',
      footerLabel: '#1a1a1a', // Đen khi không perfect
      disclaimerColor: '#8e8e93'
    };

    // Nếu là Chủ nhật, tạo Weekly Performance Dashboard
    let weeklyDashboard = '';
    if (isWeekend) {
      weeklyDashboard = buildWeeklyDashboardRaw(rawData, CONFIG, colors, targetDate, ss, dateIndex);
    }

    // Smart Badge Function
    const getPerformanceBadgeStyle = (completed, total) => {
      const rate = completed / total;
      if (rate === 1) return 'background: linear-gradient(135deg, #22c55e, #16a34a); color: white;';
      if (rate >= 0.8) return 'background: linear-gradient(135deg, #84cc16, #65a30d); color: white;';
      if (rate >= 0.6) return 'background: linear-gradient(135deg, #eab308, #ca8a04); color: white;';
      return 'background: linear-gradient(135deg, #ef4444, #dc2626); color: white;';
    };

    // Build employee lists (chỉ hiển thị nếu không phải weekly dashboard)
    let reportedHtml = '', notReportedHtml = '';

    if (!isWeekend) {
      // Danh sách đã báo cáo với star calculation chính xác
      if (reported.length > 0) {
        const reportedWithStars = reported.map(name => ({
          name,
          stars: getWeeklyStarsRaw(rawData, name, CONFIG, targetDate, ss, dateIndex)
        }));
        reportedWithStars.sort((a, b) => b.stars - a.stars);

        reportedHtml = reportedWithStars.map(person => {
          const starColor = getStarColor(person.stars);
          const starsDisplay = person.stars > 0
            ? `<span style="color: ${starColor}; font-size: 16px;">★</span>`.repeat(person.stars)
            : '';

          return `
            <table width="100%" cellpadding="0" cellspacing="0" border="0" style="padding: 16px 0;">
              <tr>
                <td style="font-size: 15px; font-weight: 400; color: ${colors.namesList}; vertical-align: middle;">
                  ${person.name}
                </td>
                <td style="text-align: right; vertical-align: middle;">
                  ${person.stars > 0 ? `<span style="white-space: nowrap;">${starsDisplay}</span>` : ''}
                </td>
              </tr>
            </table>
          `;
        }).join('');
      } else {
        reportedHtml = `<div style="padding: 16px 0; font-size: 15px; color: #8e8e93; font-style: italic;">Chưa có báo cáo nào</div>`;
      }

      // Danh sách chưa báo cáo với star calculation chính xác
      if (notReported.length > 0) {
        const notReportedWithStars = notReported.map(name => ({
          name,
          stars: getWeeklyStarsRaw(rawData, name, CONFIG, targetDate, ss, dateIndex)
        }));
        notReportedWithStars.sort((a, b) => b.stars - a.stars);

        notReportedHtml = notReportedWithStars.map(person => {
          const starColor = getStarColor(person.stars);
          const starsDisplay = person.stars > 0
            ? `<span style="color: ${starColor}; font-size: 16px;">★</span>`.repeat(person.stars)
            : '';

          return `
            <table width="100%" cellpadding="0" cellspacing="0" border="0" style="padding: 16px 0;">
              <tr>
                <td style="font-size: 15px; font-weight: 400; color: ${colors.namesList}; vertical-align: middle;">
                  ${person.name}
                </td>
                <td style="text-align: right; vertical-align: middle;">
                  ${person.stars > 0 ? `<span style="white-space: nowrap;">${starsDisplay}</span>` : ''}
                </td>
              </tr>
            </table>
          `;
        }).join('');
      } else {
        notReportedHtml = ``; // Bỏ trống khi perfect day
      }
    }

    // Daily sections for non-weekend days
    const dailySections = !isWeekend ? `
      <!-- Completed Section -->
      <div style="margin-bottom: 32px; background-color: #ffffff; border-radius: 6px; overflow: hidden;">
        <div style="padding: 20px 24px 16px;">
          <table width="100%" cellpadding="0" cellspacing="0" border="0">
            <tr>
              <td style="vertical-align: middle;">
                <h2 style="margin: 0; font-size: 18px; font-weight: 500; color: ${colors.sectionTitle};">
                  ${isPerfectDay ? 'Tất cả đã báo cáo' : 'Đã báo cáo'}
                </h2>
              </td>
              <td style="vertical-align: middle; text-align: right;">
                <span style="${getPerformanceBadgeStyle(reported.length, totalEmployees)} padding: 6px 12px; border-radius: 4px; font-weight: 600; font-size: 13px; min-width: 60px; text-align: center; display: inline-block;">
                  ${reported.length}/${totalEmployees}
                </span>
              </td>
            </tr>
          </table>
        </div>
        <div style="padding: 0 24px 8px;">
          ${reportedHtml}
        </div>
      </div>

      <!-- Pending Section -->
      ${!isPerfectDay ? `<div style="margin-bottom: 40px; background-color: #ffffff; border-radius: 6px; overflow: hidden;">
        <div style="padding: 20px 24px 16px;">
          <table width="100%" cellpadding="0" cellspacing="0" border="0">
            <tr>
              <td style="vertical-align: middle;">
                <h2 style="margin: 0; font-size: 18px; font-weight: 500; color: ${colors.pendingTitle};">
                  Chưa báo cáo
                </h2>
              </td>
              <td style="vertical-align: middle; text-align: right;">
                <span style="${getPerformanceBadgeStyle(totalEmployees - notReported.length, totalEmployees)} padding: 6px 12px; border-radius: 4px; font-weight: 600; font-size: 13px; min-width: 60px; text-align: center; display: inline-block;">
                  ${notReported.length}/${totalEmployees}
                </span>
              </td>
            </tr>
          </table>
        </div>
        <div style="padding: 0 24px 8px;">
          ${notReportedHtml}
        </div>
      </div>` : ''}
    ` : '';

    // HTML Email Template
    const htmlBody = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>${isWeekend ? 'Thống kê tuần' : 'Báo cáo ngày'} ${targetDateStr}${isCustomDate ? ' ⭐' : ''}</title>
        <!--[if mso]>
        <style type="text/css">
          table { border-collapse: collapse; }
          .container { width: 600px !important; }
        </style>
        <![endif]-->
      </head>
      <body style="margin: 0; padding: 0; background-color: #ffffff; font-family: Arial, sans-serif;">

        <!-- Outer Container for Outlook Desktop -->
        <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color: #ffffff;">
          <tr>
            <td style="padding: 20px;">
              <!-- Inner Container -->
              <table width="600" cellpadding="0" cellspacing="0" border="0" style="max-width: 600px; margin: 0 auto;" class="container">
                <tr>
                  <td style="padding: 20px;">

          <!-- Header -->
          <div style="text-align: center; margin-bottom: 48px;">
            <h1 style="margin: 0; font-size: 28px; font-weight: 300; color: ${colors.headerTitle}; letter-spacing: -0.5px;">
              ${isWeekend ? 'Thống kê tuần' : `Báo cáo tổng hợp ${isPerfectDay ? '⭐' : ''}`}
            </h1>
            <p style="margin: 8px 0 0; font-size: 16px; font-weight: 400; color: ${colors.headerSubtitle};">
              Phòng Kinh Doanh
            </p>
          </div>

          <!-- Date -->
          <div style="margin-bottom: 32px;">
            <span style="font-size: 14px; font-weight: 500; color: ${colors.dateText};">
              ${detailedDate}
            </span>
          </div>

          <!-- Weekly Dashboard (chỉ Chủ nhật) -->
          ${weeklyDashboard}

          <!-- Daily Sections (Thu 2-7) -->
          ${dailySections}

          <!-- Footer -->
          <div style="text-align: center; padding-top: 32px;">
            <p style="margin: 0; font-size: 12px; font-weight: 400; color: ${colors.footerLabel};">
              Trân trọng
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

    // DEBUG: Log email content instead of sending when debugging
    if (CONFIG.debugMode) {
      Logger.log(`📧 DEBUG - Email subject: ${subject}`);
      Logger.log(`📧 DEBUG - Email to: ${CONFIG.emailTo}`);
      Logger.log(`📧 DEBUG - Total employees: ${totalEmployees}`);
      Logger.log(`📧 DEBUG - Reported: ${reported.length}, Not reported: ${notReported.length}`);
      Logger.log(`📧 DEBUG - HTML Body length: ${htmlBody.length} characters`);

      // Log first part of HTML to check content
      Logger.log(`📧 DEBUG - HTML Preview (first 500 chars): ${htmlBody.substring(0, 500)}`);

      // Log employee data
      Logger.log(`📧 DEBUG - Reported employees:`, reported);
      Logger.log(`📧 DEBUG - Not reported employees:`, notReported);
    }

    // Gửi email với retry mechanism (skip when debugging to avoid quota)
    if (!CONFIG.debugMode) {
      sendEmailWithRetry({
        to: CONFIG.emailTo,
        subject: subject,
        htmlBody: htmlBody
      });
    } else {
      Logger.log(`📧 DEBUG - Email sending skipped (debug mode active)`);
    }

    Logger.log(`✅ Email báo cáo ${isWeekend ? 'tuần' : 'ngày'} đã được gửi thành công (Raw Data Version)`);

  } catch (error) {
    Logger.log(`❌ Lỗi khi gửi email báo cáo: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
  }
}

/**
 * Load raw data từ sheet 'tick' và parse thành array objects
 * Tự động filter excluded employees based on CONFIG.excludedEmployees
 */
function loadRawDataFromSheet(sheet, CONFIG) {
  try {
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();

    if (values.length === 0) {
      Logger.log(`❌ Sheet '${CONFIG.sheetName}' trống`);
      return [];
    }

    // First row is headers
    const headers = values[0];
    const data = [];

    // Parse each row into object và filter excluded employees
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const record = {};

      headers.forEach((header, index) => {
        record[header] = row[index];
      });

      // Filter excluded employees by employee ID
      const employeeId = String(record['ma nhan vien'] || '').trim();
      const isExcluded = CONFIG.excludedEmployees && CONFIG.excludedEmployees.includes(employeeId);

      if (!isExcluded) {
        data.push(record);
      }
    }

    if (CONFIG.debugMode) {
      Logger.log(`📊 Loaded ${data.length} records from raw data (after filtering excluded employees)`);
      if (CONFIG.excludedEmployees && CONFIG.excludedEmployees.length > 0) {
        Logger.log(`🚫 Excluded employee IDs: ${CONFIG.excludedEmployees.join(', ')}`);
      }
      Logger.log(`📋 Sample record:`, JSON.stringify(data[0], null, 2));
    }

    return data;
  } catch (error) {
    Logger.log(`❌ Lỗi khi load raw data: ${error.message}`);
    return [];
  }
}

/**
 * Get employee reports for specific date from raw data
 */
function getEmployeeReportsForDate(rawData, targetDate, ss) {
  try {
    const targetDateStr = Utilities.formatDate(targetDate, ss.getSpreadsheetTimeZone(), "M/d/yyyy");

    // Get all unique employees
    const allEmployees = [...new Set(rawData.map(record => record['ten nhan vien']))].filter(Boolean);

    // Find who reported on target date - OPTIMIZED VERSION
    const reportedEmployees = rawData
      .filter(record => {
        const recordDate = record['date'];
        const recordCheck = record['check'];
        const recordName = record['ten nhan vien'];

        if (!recordName) return false;

        let recordDateStr = '';
        if (recordDate instanceof Date) {
          recordDateStr = Utilities.formatDate(recordDate, ss.getSpreadsheetTimeZone(), "M/d/yyyy");
        } else if (typeof recordDate === 'string') {
          // Try to parse string date
          const parsedDate = new Date(recordDate);
          if (!isNaN(parsedDate.getTime())) {
            recordDateStr = Utilities.formatDate(parsedDate, ss.getSpreadsheetTimeZone(), "M/d/yyyy");
          }
        }

        const dateMatches = recordDateStr === targetDateStr;
        const hasReport = (recordCheck === 'TRUE' || recordCheck === true || recordCheck === 'X' || recordCheck === 1);

        return dateMatches && hasReport;
      })
      .map(record => record['ten nhan vien'])
      .filter(Boolean);

    const reported = [...new Set(reportedEmployees)];
    const notReported = allEmployees.filter(name => !reported.includes(name));

    return { reported, notReported };
  } catch (error) {
    Logger.log(`❌ Lỗi khi get employee reports: ${error.message}`);
    return { reported: [], notReported: [] };
  }
}

/**
 * OPTIMIZED: Calculate weekly stars from raw data with caching
 */
function getWeeklyStarsRaw(rawData, employeeName, CONFIG, currentDate, ss, dateIndex = null) {
  try {
    // Use cached date index if provided
    if (!dateIndex) {
      dateIndex = buildDateIndexRaw(rawData, ss);
    }

    const currentDayOfWeek = currentDate.getDay(); // 0=CN, 1=T2, 2=T3, 3=T4, 4=T5, 5=T6, 6=T7

    // FIXED: Tìm thứ 2 của tuần hiện tại
    let mondayOffset;
    if (currentDayOfWeek === 0) {
      // Nếu hôm nay là Chủ nhật -> lấy thứ 2 tuần trước (6 ngày trước)
      mondayOffset = -6;
    } else {
      // Nếu là T2-T7 -> lấy thứ 2 tuần này
      mondayOffset = -(currentDayOfWeek - 1);
    }

    const mondayThisWeek = new Date(currentDate);
    mondayThisWeek.setDate(currentDate.getDate() + mondayOffset);

    let stars = 0;

    // FIXED: Tính số ngày từ thứ 2 tuần này đến hôm nay (bao gồm hôm nay)
    let daysToCheck;
    if (currentDayOfWeek === 0) {
      // Chủ nhật: check 6 ngày (T2->T7 tuần trước)
      daysToCheck = 6;
    } else {
      // T2->T7: check từ T2 tuần này đến hôm nay
      daysToCheck = currentDayOfWeek;
    }

    // Duyệt từng ngày từ thứ 2 tuần này đến hôm nay - OPTIMIZED WITH INDEX
    for (let dayOffset = 0; dayOffset < daysToCheck; dayOffset++) {
      const checkDate = new Date(mondayThisWeek);
      checkDate.setDate(mondayThisWeek.getDate() + dayOffset);
      const checkDateStr = Utilities.formatDate(checkDate, ss.getSpreadsheetTimeZone(), "M/d/yyyy");

      // Use index for fast lookup
      const dateRecords = dateIndex[checkDateStr] || [];
      const hasReport = dateRecords.some(record =>
        record['ten nhan vien'] === employeeName &&
        (record['check'] === 'TRUE' || record['check'] === true || record['check'] === 'X')
      );

      if (hasReport) {
        stars++;
      }
    }

    return stars;
  } catch (error) {
    Logger.log(`❌ Lỗi khi lay weekly stars raw cho ${employeeName}: ${error.message}`);
    return 0;
  }
}

/**
 * Build date index for fast lookups - O(n) preprocessing instead of O(n²) searches
 */
function buildDateIndexRaw(rawData, ss) {
  const dateIndex = {};

  rawData.forEach(record => {
    const recordDate = record['date'];
    if (!recordDate) return;

    let dateStr = '';
    if (recordDate instanceof Date) {
      dateStr = Utilities.formatDate(recordDate, ss.getSpreadsheetTimeZone(), "M/d/yyyy");
    } else if (typeof recordDate === 'string') {
      const parsedDate = new Date(recordDate);
      if (!isNaN(parsedDate.getTime())) {
        dateStr = Utilities.formatDate(parsedDate, ss.getSpreadsheetTimeZone(), "M/d/yyyy");
      }
    }

    if (dateStr) {
      if (!dateIndex[dateStr]) {
        dateIndex[dateStr] = [];
      }
      dateIndex[dateStr].push(record);
    }
  });

  return dateIndex;
}

/**
 * Build Weekly Dashboard từ raw data
 */
function buildWeeklyDashboardRaw(rawData, CONFIG, colors, targetDate, ss, dateIndex = null) {
  try {
    // FIXED: Proper Monday calculation for weekly dashboard
    const monday = new Date(targetDate);
    const currentDayOfWeek = targetDate.getDay(); // 0=CN, 1=T2, 2=T3, 3=T4, 4=T5, 5=T6, 6=T7

    let mondayOffset;
    if (currentDayOfWeek === 0) {
      // Nếu hôm nay là Chủ nhật -> lấy thứ 2 tuần trước (6 ngày trước)
      mondayOffset = -6;
    } else {
      // Nếu là T2-T7 -> lấy thứ 2 tuần này
      mondayOffset = -(currentDayOfWeek - 1);
    }

    monday.setDate(targetDate.getDate() + mondayOffset);

    if (CONFIG.debugMode) {
      Logger.log(`📅 RAW Weekly Dashboard - Target date: ${targetDate.toDateString()}`);
      Logger.log(`📅 RAW Calculated Monday: ${monday.toDateString()}`);
      Logger.log(`📊 RAW Day of week: ${currentDayOfWeek} (0=CN)`);
    }

    // Get all employees performance data
    const allEmployees = getAllEmployeesWeeklyDataRaw(rawData, CONFIG, monday, ss, dateIndex);

    // Daily Performance Heatmap
    const heatmap = buildMobileResponsiveHeatmapRaw(allEmployees, monday, ss, CONFIG);

    // Individual Performance Dashboard
    const leaderboard = buildSimplifiedLeaderboardRaw(allEmployees, CONFIG);

    return `
      ${heatmap}
      <div style="border-top: 1px solid #22c55e; margin: 20px 0;"></div>
      ${leaderboard}
    `;
  } catch (error) {
    Logger.log(`❌ Lỗi khi tao Weekly Dashboard Raw: ${error.message}`);
    return `<div style="color: #dc3545; text-align: center; padding: 20px;">Khong the tai thong ke tuan</div>`;
  }
}

/**
 * Get all employees weekly data from raw data
 */
function getAllEmployeesWeeklyDataRaw(rawData, CONFIG, monday, ss, dateIndex = null) {
  const employees = [];

  try {
    // Get all unique employees
    const allEmployeeNames = [...new Set(rawData.map(record => record['ten nhan vien']))].filter(Boolean);

    allEmployeeNames.forEach(employeeName => {
      const weeklyData = getEmployeeWeeklyPerformanceRaw(rawData, employeeName, CONFIG, monday, ss, dateIndex);
      employees.push({
        name: employeeName,
        id: '', // Raw data may not have consistent IDs
        dailyReports: weeklyData.dailyReports,
        totalReports: weeklyData.totalReports,
        completionRate: weeklyData.completionRate,
        streak: weeklyData.streak,
        trend: weeklyData.trend
      });
    });
  } catch (error) {
    Logger.log(`❌ Lỗi khi lay du lieu nhan vien raw: ${error.message}`);
  }

  return employees;
}

/**
 * Get employee weekly performance from raw data
 */
function getEmployeeWeeklyPerformanceRaw(rawData, employeeName, CONFIG, monday, ss, dateIndex = null) {
  const dailyReports = [];
  let totalReports = 0;

  try {
    for (let dayOffset = 0; dayOffset < 6; dayOffset++) {
      const checkDate = new Date(monday);
      checkDate.setDate(monday.getDate() + dayOffset);
      const checkDateStr = Utilities.formatDate(checkDate, ss.getSpreadsheetTimeZone(), "M/d/yyyy");

      // Use index for fast lookup if available
      let hasReport = false;
      if (dateIndex) {
        const dateRecords = dateIndex[checkDateStr] || [];
        hasReport = dateRecords.some(record =>
          record['ten nhan vien'] === employeeName &&
          (record['check'] === 'TRUE' || record['check'] === true || record['check'] === 'X')
        );
      } else {
        // Fallback to original method
        hasReport = rawData.some(record => {
          const recordName = record['ten nhan vien'];
          const recordDate = record['date'];
          const recordCheck = record['check'];

          let recordDateStr = '';
          if (recordDate instanceof Date) {
            recordDateStr = Utilities.formatDate(recordDate, ss.getSpreadsheetTimeZone(), "M/d/yyyy");
          } else if (typeof recordDate === 'string') {
            const parsedDate = new Date(recordDate);
            if (!isNaN(parsedDate.getTime())) {
              recordDateStr = Utilities.formatDate(parsedDate, ss.getSpreadsheetTimeZone(), "M/d/yyyy");
            }
          }

          return recordName === employeeName &&
                 recordDateStr === checkDateStr &&
                 (recordCheck === 'TRUE' || recordCheck === true || recordCheck === 'X');
        });
      }

      dailyReports.push(hasReport);
      if (hasReport) {
        totalReports++;
      }
    }
  } catch (error) {
    Logger.log(`❌ Lỗi khi lay performance raw cua ${employeeName}: ${error.message}`);
  }

  const completionRate = totalReports / 6;
  const streak = calculateStreak(dailyReports);
  const trend = calculateTrend(dailyReports);

  return {
    dailyReports,
    totalReports,
    completionRate,
    streak,
    trend
  };
}

/**
 * Build heatmap from raw data
 */
function buildMobileResponsiveHeatmapRaw(employees, monday, ss, CONFIG) {
  const dayNames = ['T2', 'T3', 'T4', 'T5', 'T6', 'T7'];
  let heatmapHtml = '';

  // Tính tỷ lệ cho từng ngày
  const dayRates = [];
  for (let day = 0; day < 6; day++) {
    const dayReports = employees.filter(emp => emp.dailyReports[day]).length;
    const totalEmployees = employees.length;
    const dayRate = dayReports / totalEmployees;
    dayRates.push(dayRate);
  }

  // Tìm tỷ lệ thấp nhất
  const minRate = Math.min(...dayRates);

  for (let day = 0; day < 6; day++) {
    const dayRate = dayRates[day];
    const percentage = Math.round(dayRate * 100);

    let boxStyle = '';
    let textColor = '#1a1a1a';
    let displayText = '';

    if (dayRate === 0) {
      // Ngay nghi (0%) -> hien thi 'x'
      boxStyle = 'background-color: #ffffff; color: #1a1a1a;';
      textColor = '#1a1a1a';
      displayText = 'x';
    } else if (dayRate === 1) {
      // Perfect day (100%) -> màu xanh
      boxStyle = 'background-color: #ffffff; color: #22c55e;';
      textColor = '#22c55e';
      displayText = '100';
    } else {
      // Ngày thường (dưới 100%) -> màu đen
      boxStyle = 'background-color: #ffffff; color: #1a1a1a;';
      textColor = '#1a1a1a';
      displayText = `${percentage}`;
    }

    heatmapHtml += `
      <td style="text-align: center; width: 16.66%;">
        <div style="${boxStyle} padding: 12px 4px; border-radius: 6px; margin: 0 2px;">
          <div style="font-size: 10px; font-weight: 600; margin-bottom: 6px; color: ${textColor};">${dayNames[day]}</div>
          <div style="font-size: 14px; font-weight: 700; color: ${textColor};">${displayText}</div>
        </div>
      </td>
    `;
  }

  return `
    <div style="margin-bottom: 32px; background-color: #ffffff; border-radius: 6px; padding: 20px;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr>
          ${heatmapHtml}
        </tr>
      </table>
    </div>
  `;
}

/**
 * Build leaderboard from raw data
 */
function buildSimplifiedLeaderboardRaw(employees, CONFIG) {
  // Remove duplicates by name
  const uniqueEmployees = [];
  const employeeMap = new Map();

  employees.forEach(emp => {
    const existing = employeeMap.get(emp.name);
    if (!existing || emp.totalReports > existing.totalReports) {
      employeeMap.set(emp.name, emp);
    }
  });

  employeeMap.forEach(emp => uniqueEmployees.push(emp));

  // Group by star count
  const starGroups = {};
  uniqueEmployees.forEach(emp => {
    const stars = emp.totalReports;
    if (!starGroups[stars]) {
      starGroups[stars] = [];
    }
    starGroups[stars].push(emp);
  });

  // Sort star levels descending
  const sortedStarLevels = Object.keys(starGroups)
    .map(Number)
    .sort((a, b) => b - a);

  // Medal system with proper fallback for email clients
  const medalMap = {
    0: '&#x1F947;', // 🥇 - HTML entity for gold medal
    1: '&#x1F948;', // 🥈 - HTML entity for silver medal
    2: '&#x1F949;'  // 🥉 - HTML entity for bronze medal
  };

  let leaderboardHtml = '';
  let currentRank = 1;

  sortedStarLevels.forEach((starLevel, groupIndex) => {
    const employeesInGroup = starGroups[starLevel];

    // Use HTML entity medals for top 3 groups, then numbers
    const medal = groupIndex < 3 ? medalMap[groupIndex] : '';

    employeesInGroup.sort((a, b) => a.name.localeCompare(b.name));

    employeesInGroup.forEach(emp => {
      const starColor = getStarColor(emp.totalReports);
      const starsDisplay = emp.totalReports > 0
        ? `<span style="color: ${starColor}; font-size: 16px;">★</span>`.repeat(emp.totalReports)
        : '<span style="color: #94a3b8; font-size: 14px;">Chua bao cao</span>';

      // Display medal or rank number
      const rankDisplay = medal || currentRank;

      leaderboardHtml += `
        <table width="100%" cellpadding="0" cellspacing="0" border="0" style="padding: 12px 0;">
          <tr>
            <td style="width: 40px; text-align: center; font-size: 16px; vertical-align: middle;">
              ${rankDisplay}
            </td>
            <td style="padding-left: 12px; vertical-align: middle;">
              <div style="font-size: 14px; font-weight: 400; color: #22c55e;">${emp.name}</div>
            </td>
            <td style="text-align: right; vertical-align: middle;">
              <span style="white-space: nowrap;">
                ${starsDisplay}
              </span>
            </td>
          </tr>
        </table>
      `;
      currentRank++;
    });
  });

  return `
    <div style="margin-bottom: 16px; background-color: #ffffff; border-radius: 6px; padding: 16px;">
      ${leaderboardHtml}
    </div>
  `;
}

// =====================
// REUSE UTILITY FUNCTIONS FROM ORIGINAL VERSION
// =====================

/**
 * Parse target date tu input cua user
 * @param {string|Date|null} customDate - Ngay tuy chon
 * @returns {Date} - Date object da duoc parse
 */
function parseTargetDate(customDate) {
  if (!customDate) {
    return new Date(); // Ngay hien tai
  }

  try {
    if (customDate instanceof Date) {
      return new Date(customDate);
    }

    if (typeof customDate === 'string') {
      // Support cac format: 'YYYY-MM-DD', 'MM/DD/YYYY', 'DD/MM/YYYY'
      let parsedDate;

      if (customDate.includes('-')) {
        // Format: YYYY-MM-DD
        parsedDate = new Date(customDate);
      } else if (customDate.includes('/')) {
        // Format: MM/DD/YYYY hoac DD/MM/YYYY
        parsedDate = new Date(customDate);
      } else {
        throw new Error('Invalid date format');
      }

      if (isNaN(parsedDate.getTime())) {
        throw new Error('Invalid date');
      }

      return parsedDate;
    }

    throw new Error('Unsupported date type');
  } catch (error) {
    Logger.log(`⚠️ Loi parse custom date '${customDate}': ${error.message}. Su dung ngay hien tai.`);
    return new Date();
  }
}

/**
 * FIXED: Gui email voi retry mechanism
 */
function sendEmailWithRetry(emailConfig, maxRetries = 3) {
  for (let i = 0; i < maxRetries; i++) {
    try {
      GmailApp.sendEmail(
        emailConfig.to,
        emailConfig.subject,
        '', // body text rong, vi dung htmlBody
        {
          htmlBody: emailConfig.htmlBody,
          name: "Báo Cáo Ngày" // Dat ten ngau vao day
        }
      );
      Logger.log(`✅ Email sent successfully on attempt ${i + 1}`);
      return true;
    } catch (error) {
      Logger.log(`❌ Email attempt ${i + 1} failed: ${error.message}`);
      if (i === maxRetries - 1) throw error;
      Utilities.sleep(1000 * (i + 1)); // Exponential backoff
    }
  }
  return false;
}

/**
 * SIMPLIFIED: Star Color Function - Chỉ dựa vào số sao tuyệt đối
 */
function getStarColor(starCount) {
  // Sử dụng thang màu đơn giản theo số sao
  if (starCount >= 6) return '#22c55e';       // 6 sao - Xanh dam hoan hao
  if (starCount >= 5) return '#84cc16';       // 5 sao - Xanh lime xuat sac
  if (starCount >= 4) return '#22c55e';       // 4 sao - Xanh tot
  if (starCount >= 3) return '#eab308';       // 3 sao - Vang kha
  if (starCount >= 2) return '#f97316';       // 2 sao - Cam trung binh
  if (starCount >= 1) return '#94a3b8';       // 1 sao - Xam nhat can cai thien
  return '#d1d5db';                           // 0 sao - Xam nhat chua bat dau
}

/**
 * Utility functions
 */
function calculateStreak(dailyReports) {
  let streak = 0;
  for (let i = dailyReports.length - 1; i >= 0; i--) {
    if (dailyReports[i]) {
      streak++;
    } else {
      break;
    }
  }
  return streak;
}

function calculateTrend(dailyReports) {
  const firstHalf = dailyReports.slice(0, 3).filter(Boolean).length;
  const secondHalf = dailyReports.slice(3, 6).filter(Boolean).length;

  if (secondHalf > firstHalf) return 'up';
  if (secondHalf < firstHalf) return 'down';
  return 'stable';
}

// =====================
// HELPER FUNCTIONS FOR RAW VERSION
// =====================

/**
 * HELPER FUNCTION: Gui bao cao cho ngay cu the (raw version)
 * @param {string} dateString - Ngay theo format 'YYYY-MM-DD' (VD: '2025-07-15')
 *
 * USAGE:
 * sendReportForDateRaw('2025-07-15') - Gui bao cao ngay 15/7/2025
 * sendReportForDateRaw('2025-06-30') - Gui bao cao ngay 30/6/2025
 */
function sendReportForDateRaw(dateString) {
  Logger.log(`🎯 Gui bao cao RAW cho ngay: ${dateString}`);
  sendDailyReportSummaryRaw(dateString);
}

/**
 * HELPER FUNCTION: Gui bao cao cho ngay hom qua (raw version)
 */
function sendReportForYesterdayRaw() {
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const yesterdayStr = Utilities.formatDate(yesterday, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyy-MM-dd");
  Logger.log(`📅 Gui bao cao RAW cho ngay hom qua: ${yesterdayStr}`);
  sendDailyReportSummaryRaw(yesterday);
}

/**
 * HELPER FUNCTION: Gui bao cao cho tuan truoc (Chu nhat) (raw version)
 */
function sendReportForLastSundayRaw() {
  const today = new Date();
  const lastSunday = new Date(today);
  const daysToLastSunday = today.getDay() === 0 ? 7 : today.getDay();
  lastSunday.setDate(today.getDate() - daysToLastSunday);

  const lastSundayStr = Utilities.formatDate(lastSunday, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyy-MM-dd");
  Logger.log(`📊 Gui bao cao tuan RAW cho Chu nhat truoc: ${lastSundayStr}`);
  sendDailyReportSummaryRaw(lastSunday);
}

/**
 * TEST FUNCTION - Test raw data version
 */
function testRawDataVersion() {
  Logger.log('🧪 TESTING RAW DATA VERSION');

  // Test loading raw data
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('tick');

  if (!sheet) {
    Logger.log('❌ Sheet "tick" khong ton tai');
    return;
  }

  const CONFIG = { debugMode: true };
  const rawData = loadRawDataFromSheet(sheet, CONFIG);

  Logger.log(`📊 Raw data sample:`, rawData.slice(0, 3));

  // Test date querying
  const testDate = new Date('2025-01-01');
  const reports = getEmployeeReportsForDate(rawData, testDate, ss);

  Logger.log(`📅 Reports for ${testDate.toDateString()}:`);
  Logger.log(`✅ Reported (${reports.reported.length}):`, reports.reported);
  Logger.log(`❌ Not Reported (${reports.notReported.length}):`, reports.notReported);

  Logger.log('✅ Raw data version test completed');
}

/**
 * TEST EMAIL CONTENT cho ngày hôm qua (không gửi thật)
 */
function testEmailContentYesterday() {
  Logger.log('🧪 TESTING EMAIL CONTENT FOR YESTERDAY');

  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);

  // Call the main function with debug mode already active
  sendDailyReportSummaryRaw(yesterday);
}

/**
 * SIMPLE DEBUG TEST - Kiểm tra data matching logic
 */
function debugDataMatching() {
  Logger.log('🧪 DEBUGGING DATA MATCHING LOGIC');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('tick');

  if (!sheet) {
    Logger.log('❌ Sheet "tick" không tồn tại');
    return;
  }

  // Load raw data
  const CONFIG = { debugMode: true, sheetName: 'tick' };
  const rawData = loadRawDataFromSheet(sheet, CONFIG);

  Logger.log(`📊 Loaded ${rawData.length} records`);

  if (rawData.length > 0) {
    Logger.log(`📋 Sample record:`, rawData[0]);

    // Test with today
    const today = new Date();
    const targetDateStr = Utilities.formatDate(today, ss.getSpreadsheetTimeZone(), "M/d/yyyy");
    Logger.log(`🎯 Target date string: ${targetDateStr}`);

    // Check last 10 records for date formats
    Logger.log(`📅 Checking last 10 records for date formats:`);
    const lastRecords = rawData.slice(-10);

    lastRecords.forEach((record, index) => {
      const recordDate = record['date'];
      const recordCheck = record['check'];
      const recordName = record['ten nhan vien'];

      let recordDateStr = '';
      if (recordDate instanceof Date) {
        recordDateStr = Utilities.formatDate(recordDate, ss.getSpreadsheetTimeZone(), "M/d/yyyy");
      } else {
        recordDateStr = recordDate ? recordDate.toString() : 'NO_DATE';
      }

      Logger.log(`   ${index}: ${recordName || 'NO_NAME'} | ${recordDateStr} | ${recordCheck || 'NO_CHECK'}`);
    });

    // Test the actual function
    const reports = getEmployeeReportsForDate(rawData, today, ss);
    Logger.log(`🎭 FINAL RESULT: Reported: ${reports.reported.length}, Not reported: ${reports.notReported.length}`);
  }

  Logger.log('✅ Debug data matching completed');
}

/**
 * QUICK FIX - Temporary function to debug thật sự
 */
function quickDebugRaw() {
  Logger.log('🔧 QUICK DEBUG RAW');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('tick');

  const range = sheet.getDataRange();
  const values = range.getValues();

  Logger.log(`📊 Sheet có ${values.length} rows, ${values[0]?.length || 0} columns`);
  Logger.log(`📋 Headers:`, values[0]);

  // Check 3 sample data rows
  for (let i = 1; i <= Math.min(3, values.length - 1); i++) {
    Logger.log(`📋 Row ${i}:`, values[i]);
  }

  // Check today's data specifically
  const today = new Date();
  const todayStr = Utilities.formatDate(today, ss.getSpreadsheetTimeZone(), "M/d/yyyy");
  Logger.log(`🎯 Looking for today: ${todayStr}`);

  // Find headers
  const headers = values[0];
  const nameCol = headers.indexOf('ten nhan vien');
  const dateCol = headers.indexOf('date');
  const checkCol = headers.indexOf('check');

  Logger.log(`🔍 Column indexes: name=${nameCol}, date=${dateCol}, check=${checkCol}`);

  // Find today's entries
  let todayCount = 0;
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowDate = row[dateCol];
    let rowDateStr = '';

    if (rowDate instanceof Date) {
      rowDateStr = Utilities.formatDate(rowDate, ss.getSpreadsheetTimeZone(), "M/d/yyyy");
    }

    if (rowDateStr === todayStr) {
      todayCount++;
      Logger.log(`📅 TODAY MATCH: ${row[nameCol]} | ${rowDateStr} | ${row[checkCol]}`);
    }
  }

  Logger.log(`🎭 Found ${todayCount} entries for today`);
}

/**
 * TEST EXCLUSION - Kiểm tra excluded employees có bị loại bỏ không
 */
function testExcludedEmployees() {
  Logger.log('🧪 TESTING EXCLUDED EMPLOYEES FILTER');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('tick');

  if (!sheet) {
    Logger.log('❌ Sheet "tick" không tồn tại');
    return;
  }

  const CONFIG = {
    sheetName: 'tick',
    excludedEmployees: ['004620'], // Test với Trần Thị Phương Phi
    debugMode: true
  };

  // Load raw data with exclusion
  const rawData = loadRawDataFromSheet(sheet, CONFIG);

  // Check if excluded employee still exists
  const excludedEmployee = rawData.find(record => record['ma nhan vien'] === '004620');

  if (excludedEmployee) {
    Logger.log('❌ FAIL - Employee 004620 vẫn còn trong data');
  } else {
    Logger.log('✅ SUCCESS - Employee 004620 đã bị loại bỏ khỏi data');
  }

  // Show unique employees
  const uniqueEmployees = [...new Set(rawData.map(r => `${r['ma nhan vien']} - ${r['ten nhan vien']}`))];
  Logger.log(`📊 Total unique employees (after exclusion): ${uniqueEmployees.length}`);
  Logger.log(`📋 Employee list:`, uniqueEmployees.slice(0, 10)); // Show first 10
}

/**
 * TEST SHEET STRUCTURE - Kiểm tra cấu trúc sheet 'tick'
 */
function testSheetStructure() {
  Logger.log('🧪 TESTING SHEET STRUCTURE');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('tick');

  if (!sheet) {
    Logger.log('❌ Sheet "tick" không tồn tại');
    return;
  }

  const range = sheet.getDataRange();
  const values = range.getValues();

  Logger.log(`📊 Sheet có ${values.length} dòng và ${values[0]?.length || 0} cột`);

  if (values.length > 0) {
    Logger.log(`📋 Headers (dòng 1):`, values[0]);

    if (values.length > 1) {
      Logger.log(`📋 Sample data (dòng 2):`, values[1]);
      Logger.log(`📋 Sample data (dòng 3):`, values[2] || 'Không có dòng 3');
      Logger.log(`📋 Sample data (dòng 4):`, values[3] || 'Không có dòng 4');
      Logger.log(`📋 Sample data (dòng 5):`, values[4] || 'Không có dòng 5');
    }

    // Test column mapping
    const headers = values[0];
    Logger.log(`🔍 Tìm cột 'ten nhan vien':`, headers.indexOf('ten nhan vien'));
    Logger.log(`🔍 Tìm cột 'date':`, headers.indexOf('date'));
    Logger.log(`🔍 Tìm cột 'check':`, headers.indexOf('check'));

    // Detailed analysis
    Logger.log(`🔍 All headers:`, headers);

    // Check recent date entries
    const today = new Date();
    const yesterday = new Date(today);
    yesterday.setDate(today.getDate() - 1);

    const todayStr = Utilities.formatDate(today, ss.getSpreadsheetTimeZone(), "M/d/yyyy");
    const yesterdayStr = Utilities.formatDate(yesterday, ss.getSpreadsheetTimeZone(), "M/d/yyyy");

    Logger.log(`🎯 Looking for today: ${todayStr}`);
    Logger.log(`🎯 Looking for yesterday: ${yesterdayStr}`);

    // Check date formats in recent entries
    for (let i = Math.max(1, values.length - 20); i < Math.min(values.length, values.length); i++) {
      const row = values[i];
      const dateValue = row[headers.indexOf('date')];
      const checkValue = row[headers.indexOf('check')];
      const nameValue = row[headers.indexOf('ten nhan vien')];

      if (dateValue) {
        let dateStr = '';
        if (dateValue instanceof Date) {
          dateStr = Utilities.formatDate(dateValue, ss.getSpreadsheetTimeZone(), "M/d/yyyy");
        } else {
          dateStr = dateValue.toString();
        }

        Logger.log(`📅 Row ${i}: ${nameValue} | ${dateStr} | ${checkValue}`);
      }
    }
  }

  Logger.log('✅ Sheet structure test completed');
}

/**
 * TEST FUNCTION - Test raw data version for a specific date with known data
 */
function testRawWithKnownDate() {
  Logger.log('🧪 TESTING RAW DATA VERSION WITH KNOWN DATE');

  // Test with a date we know has data: 2/8/2025 (from sample data)
  const testDate = '2025-02-08'; // February 8, 2025 - has data in sample
  Logger.log(`🎯 Testing with known date: ${testDate}`);

  sendDailyReportSummaryRaw(testDate);
}