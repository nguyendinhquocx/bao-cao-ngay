/**
 * Gửi email báo cáo tổng hợp ngày với enhanced UI/UX.
 * Perfect Day: header đơn giản với emoji, loại bỏ ratio display
 */
function sendDailyReportSummary() {
  const CONFIG = {
    sheetName: 'check bc',
    emailTo: 'quoc.nguyen3@hoanmy.com',
    dateHeaderRanges: ['e3:n3', 'e17:n17', 'e30:o30'],
    dataRanges: ['B4:n12', 'B18:n26', 'B31:o39'],
    
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
    
    // Enhanced celebration icons
    celebrationIcon: 'https://cdn-icons-png.flaticon.com/128/9422/9422222.png'
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheetName);
  const today = new Date();
  const todayStr = Utilities.formatDate(today, ss.getSpreadsheetTimeZone(), "M/d/yyyy");

  // Định dạng ngày chi tiết với thứ
  const dayNames = ['Chủ nhật', 'Thứ hai', 'Thứ ba', 'Thứ tư', 'Thứ năm', 'Thứ sáu', 'Thứ bảy'];
  const dayOfWeek = dayNames[today.getDay()];
  const detailedDate = `${dayOfWeek}, ngày ${today.getDate()} tháng ${today.getMonth() + 1} năm ${today.getFullYear()}`;

  // Tìm vị trí cột ngày hôm nay trong các vùng header
  let dateColumnIndex = null, dataRange = null, values = null;
  for (let i = 0; i < CONFIG.dateHeaderRanges.length; i++) {
    const headerRange = sheet.getRange(CONFIG.dateHeaderRanges[i]);
    const headerValues = headerRange.getValues()[0];
    for (let j = 0; j < headerValues.length; j++) {
      const cell = headerValues[j];
      if (cell instanceof Date) {
        const dateStr = Utilities.formatDate(cell, ss.getSpreadsheetTimeZone(), "M/d/yyyy");
        if (dateStr === todayStr) {
          dateColumnIndex = headerRange.getColumn() + j;
          dataRange = sheet.getRange(CONFIG.dataRanges[i]);
          values = dataRange.getValues();
          break;
        }
      }
    }
    if (dateColumnIndex !== null) break;
  }
  
  if (!values) {
    Logger.log("Không tìm thấy cột ngày hôm nay trong bất kỳ vùng tiêu đề nào.");
    return;
  }

  // Lấy danh sách đã báo cáo và chưa báo cáo
  let reported = [], notReported = [];
  for (let row of values) {
    const maNV = row[0];
    const tenNV = row[2];
    const reportMark = row[dateColumnIndex - dataRange.getColumn()];
    if (maNV && tenNV) {
      if (reportMark === 'X') {
        reported.push(tenNV);
      } else {
        notReported.push(tenNV);
      }
    }
  }

  // Kiểm tra perfect day và tính totals
  const totalEmployees = reported.length + notReported.length;
  const isPerfectDay = notReported.length === 0 && reported.length > 0;
  const subject = `HMSG | P.KD - TỔNG HỢP BÁO CÁO NGÀY ${todayStr}`;

  // Chọn icons theo trạng thái
  const calendarIcon = isPerfectDay ? CONFIG.calendarIconPerfect : CONFIG.calendarIconDefault;
  const completedIcon = isPerfectDay ? CONFIG.completedIconPerfect : CONFIG.completedIconDefault;
  const pendingIcon = isPerfectDay ? CONFIG.pendingIconPerfect : CONFIG.pendingIconDefault;

  // Enhanced color scheme
  const colors = isPerfectDay ? {
    border: '#22c55e',
    headerTitle: '#22c55e',
    headerSubtitle: '#16a34a',
    dateText: '#16a34a',
    sectionTitle: '#22c55e',
    namesList: '#15803d',
    footerName: '#16a34a',
    footerLabel: '#22c55e',
    disclaimerColor: '#16a34a'
  } : {
    border: '#000000',
    headerTitle: '#1a1a1a',
    headerSubtitle: '#8e8e93',
    dateText: '#495057',
    sectionTitle: '#1a1a1a',
    pendingTitle: '#dc3545',
    namesList: '#1a1a1a',
    footerName: '#8e8e93',
    footerLabel: '#1a1a1a',
    disclaimerColor: '#8e8e93'
  };

  // Smart Badge Function
  const getPerformanceBadgeStyle = (completed, total) => {
    const rate = completed / total;
    if (rate === 1) return 'background: linear-gradient(135deg, #22c55e, #16a34a); color: white;';
    if (rate >= 0.8) return 'background: linear-gradient(135deg, #84cc16, #65a30d); color: white;';
    if (rate >= 0.6) return 'background: linear-gradient(135deg, #eab308, #ca8a04); color: white;';
    return 'background: linear-gradient(135deg, #ef4444, #dc2626); color: white;';
  };

  // Progressive Star Color Function
  const getStarColor = (starCount, totalPossible) => {
    const ratio = starCount / totalPossible;
    if (ratio >= 0.9) return '#22c55e'; // Xanh đậm - Xuất sắc
    if (ratio >= 0.7) return '#84cc16'; // Xanh lime - Tốt  
    if (ratio >= 0.5) return '#eab308'; // Vàng - Trung bình
    return '#94a3b8'; // Xám - Cần cải thiện
  };

  // Tính số ngày làm việc tuần này cho progressive stars
  const today_day = new Date().getDay();
  const workDaysThisWeek = today_day === 0 ? 6 : today_day;

  // Danh sách đã báo cáo với progressive stars (không có ratio)
  let reportedHtml = '';
  if (reported.length > 0) {
    const reportedWithStars = reported.map(name => ({ 
      name, 
      stars: getWeeklyStars(sheet, name, ss, CONFIG) 
    }));
    reportedWithStars.sort((a, b) => b.stars - a.stars);
    
    reportedHtml = reportedWithStars.map(person => {
      const starColor = getStarColor(person.stars, workDaysThisWeek);
      const starsDisplay = person.stars > 0
        ? `<span style="color: ${starColor}; font-size: 16px;">★</span>`.repeat(person.stars)
        : '';
      
      return `
        <div style="padding: 16px 0; 
                    font-size: 15px; 
                    font-weight: 400; 
                    color: ${colors.namesList}; 
                    border-bottom: 1px solid #f5f5f5; 
                    display: flex; 
                    justify-content: space-between; 
                    align-items: center;
                    min-height: 48px;">
          <span style="flex: 1;">${person.name}</span>
          ${person.stars > 0 ? `
            <span style="display: flex; gap: 2px;">${starsDisplay}</span>
          ` : ''}
        </div>
      `;
    }).join('');
  } else {
    reportedHtml = `<div style="padding: 16px 0; font-size: 15px; color: #8e8e93; font-style: italic;">Chưa có báo cáo nào</div>`;
  }

  // Danh sách chưa báo cáo với progressive stars (không có ratio)
  let notReportedHtml = '';
  if (notReported.length > 0) {
    const notReportedWithStars = notReported.map(name => ({ 
      name, 
      stars: getWeeklyStars(sheet, name, ss, CONFIG) 
    }));
    notReportedWithStars.sort((a, b) => b.stars - a.stars);
    
    notReportedHtml = notReportedWithStars.map(person => {
      const starColor = getStarColor(person.stars, workDaysThisWeek);
      const starsDisplay = person.stars > 0
        ? `<span style="color: ${starColor}; font-size: 16px;">★</span>`.repeat(person.stars)
        : '';
      
      return `
        <div style="padding: 16px 0; 
                    font-size: 15px; 
                    font-weight: 400; 
                    color: ${colors.namesList}; 
                    border-bottom: 1px solid #f5f5f5; 
                    display: flex; 
                    justify-content: space-between; 
                    align-items: center;
                    min-height: 48px;">
          <span style="flex: 1;">${person.name}</span>
          ${person.stars > 0 ? `
            <span style="display: flex; gap: 2px;">${starsDisplay}</span>
          ` : ''}
        </div>
      `;
    }).join('');
  } else {
    notReportedHtml = `<div style="padding: 16px 0; font-size: 15px; color: ${colors.namesList}; font-style: italic;">
      Tất cả đã báo cáo
      <img src="${CONFIG.celebrationIcon}" width="20" height="20" style="margin-left: 8px;" alt="Celebration">
    </div>`;
  }

  // HTML Email Template
  const htmlBody = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Báo cáo ngày ${todayStr}</title>
    </head>
    <body style="margin: 0; padding: 0; background-color: #ffffff; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;">
      
      <!-- Main Container -->
      <div style="max-width: 600px; margin: 40px auto; padding: 40px; border: 1px solid ${colors.border}; border-radius: 12px;">
        
        <!-- Header (đơn giản, có emoji khi perfect day) -->
        <div style="text-align: center; margin-bottom: 48px;">
          <h1 style="margin: 0; font-size: 28px; font-weight: 300; color: ${colors.headerTitle}; letter-spacing: -0.5px;">
            Báo cáo tổng hợp ${isPerfectDay ? '🎉' : ''}
          </h1>
          <p style="margin: 8px 0 0; font-size: 16px; font-weight: 400; color: ${colors.headerSubtitle};">
            Phòng Kinh Doanh HMSG
          </p>
        </div>

        <!-- Date -->
        <div style="margin-bottom: 32px;">
          <span style="font-size: 14px; font-weight: 500; color: ${colors.dateText};">
            <img src="${calendarIcon}" width="16" height="16" style="vertical-align: middle; margin-right: 8px;" alt="Calendar">
            ${detailedDate}
          </span>
        </div>

        <!-- Completed Section -->
        <div style="margin-bottom: 32px; background-color: #ffffff; border: 1px solid #e9ecef; border-radius: 12px; overflow: hidden;">
          <div style="padding: 20px 24px 16px; border-bottom: 1px solid #f5f5f5;">
            <div style="display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 12px;">
              <h2 style="margin: 0; font-size: 18px; font-weight: 500; color: ${colors.sectionTitle}; display: flex; align-items: center;">
                <img src="${completedIcon}" width="20" height="20" style="margin-right: 12px;" alt="Completed">
                Đã báo cáo
              </h2>
              <span style="${getPerformanceBadgeStyle(reported.length, totalEmployees)} 
                           padding: 6px 12px; 
                           border-radius: 12px; 
                           font-weight: 600; 
                           font-size: 13px;
                           min-width: 60px;
                           text-align: center;">
                ${reported.length}/${totalEmployees} ${reported.length === totalEmployees ? '🏆' : ''}
              </span>
            </div>
          </div>
          <div style="padding: 0 24px 8px;">
            ${reportedHtml}
          </div>
        </div>

        <!-- Pending Section -->
        <div style="margin-bottom: 40px; background-color: #ffffff; border: 1px solid #e9ecef; border-radius: 12px; overflow: hidden;">
          <div style="padding: 20px 24px 16px; border-bottom: 1px solid #f5f5f5;">
            <div style="display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 12px;">
              <h2 style="margin: 0; font-size: 18px; font-weight: 500; color: ${isPerfectDay ? colors.sectionTitle : colors.pendingTitle}; display: flex; align-items: center;">
                <img src="${pendingIcon}" width="20" height="20" style="margin-right: 12px;" alt="Pending">
                Chưa báo cáo
              </h2>
              <span style="${getPerformanceBadgeStyle(totalEmployees - notReported.length, totalEmployees)} 
                           padding: 6px 12px; 
                           border-radius: 12px; 
                           font-weight: 600; 
                           font-size: 13px;
                           min-width: 60px;
                           text-align: center;">
                ${notReported.length}/${totalEmployees} ${notReported.length === 0 ? '🎯' : ''}
              </span>
            </div>
          </div>
          <div style="padding: 0 24px 8px;">
            ${notReportedHtml}
          </div>
        </div>

        <!-- Footer -->
        <div style="text-align: center; padding-top: 32px; border-top: 1px solid #f5f5f5;">
          <p style="margin: 0 0 6px; font-size: 12px; font-weight: 400; color: ${colors.footerLabel};">
            Trân trọng
          </p>
          <p style="margin: 0; font-size: 12px; font-weight: 500; color: ${colors.footerName};">
            Nguyen Dinh Quoc
          </p>
        </div>

        <!-- Disclaimer -->
        <div style="margin-top: 40px; text-align: center;">
          <p style="margin: 0; font-size: 12px; color: ${colors.disclaimerColor}; line-height: 1.4; font-style: italic;">
            Đây là báo cáo tự động. Vui lòng không trả lời email này.<br>
            Liên hệ: quoc.nguyen3@hoanmy.com
          </p>
        </div>

      </div>
      
    </body>
    </html>
  `;

  // Gửi email
  MailApp.sendEmail({
    to: CONFIG.emailTo,
    subject: subject,
    htmlBody: htmlBody
  });
}

/**
 * Tính số sao theo tuần cho một nhân viên.
 */
function getWeeklyStars(sheet, employeeName, ss, CONFIG) {
  const today = new Date();
  const currentDay = today.getDay(); // 0 = CN, 1 = T2, ..., 6 = T7
  const mondayOffset = currentDay === 0 ? -6 : -(currentDay - 1);
  const monday = new Date(today);
  monday.setDate(today.getDate() + mondayOffset);
  
  let stars = 0;
  const daysToCheck = currentDay === 0 ? 6 : currentDay; // Từ thứ 2 đến hôm nay
  
  for (let dayOffset = 0; dayOffset < daysToCheck; dayOffset++) {
    const checkDate = new Date(monday);
    checkDate.setDate(monday.getDate() + dayOffset);
    const checkDateStr = Utilities.formatDate(checkDate, ss.getSpreadsheetTimeZone(), "M/d/yyyy");
    
    // Tìm cột cho ngày này
    for (let i = 0; i < CONFIG.dateHeaderRanges.length; i++) {
      const headerRange = sheet.getRange(CONFIG.dateHeaderRanges[i]);
      const headerValues = headerRange.getValues()[0];
      
      for (let j = 0; j < headerValues.length; j++) {
        const cell = headerValues[j];
        if (cell instanceof Date) {
          const dateStr = Utilities.formatDate(cell, ss.getSpreadsheetTimeZone(), "M/d/yyyy");
          if (dateStr === checkDateStr) {
            const dateColumnIndex = headerRange.getColumn() + j;
            const dataRange = sheet.getRange(CONFIG.dataRanges[i]);
            const values = dataRange.getValues();
            
            // Tìm nhân viên trong data range này
            for (let row of values) {
              const tenNV = row[2]; // Cột tên nhân viên
              const reportMark = row[dateColumnIndex - dataRange.getColumn()];
              
              if (tenNV === employeeName && reportMark === 'X') {
                stars++;
                break;
              }
            }
            break;
          }
        }
      }
    }
  }
  
  return stars;
}
