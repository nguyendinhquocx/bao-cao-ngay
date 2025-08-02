# Hướng dẫn sử dụng tính năng Custom Date

## Tổng quan
Tính năng mới cho phép gửi báo cáo cho bất kỳ ngày nào trong quá khứ, không chỉ ngày hiện tại.

## Cách sử dụng

### 1. Gửi báo cáo ngày hiện tại (như trước)
```javascript
sendDailyReportSummary()
```

### 2. Gửi báo cáo cho ngày cụ thể
```javascript
// Format YYYY-MM-DD (khuyến nghị)
sendDailyReportSummary('2025-07-15')

// Format MM/DD/YYYY
sendDailyReportSummary('07/15/2025')

// Sử dụng Date object
sendDailyReportSummary(new Date('2025-07-15'))
```

### 3. Các helper functions tiện lợi

#### Gửi báo cáo cho ngày cụ thể (dễ sử dụng)
```javascript
sendReportForDate('2025-07-15')
```

#### Gửi báo cáo cho ngày hôm qua
```javascript
sendReportForYesterday()
```

#### Gửi báo cáo tuần cho Chủ nhật trước
```javascript
sendReportForLastSunday()
```

## Tính năng

### ✅ Hỗ trợ multiple date formats
- `'2025-07-15'` (YYYY-MM-DD)
- `'07/15/2025'` (MM/DD/YYYY)
- `new Date('2025-07-15')` (Date object)

### ✅ Error handling
- Nếu date format không hợp lệ → tự động fallback về ngày hiện tại
- Logging chi tiết để debug

### ✅ Email customization
- Subject line có thêm "" khi gửi custom date
- Disclaimer thông báo rõ đây là báo cáo tùy chỉnh
- Title trong HTML cũng được cập nhật

### ✅ Weekly dashboard support
- Nếu custom date là Chủ nhật → vẫn hiển thị weekly dashboard
- Weekly stars calculation được tính theo tuần của custom date

## Testing

### Test parsing functions
```javascript
testCustomDateFeature()
```

### Test weekly stars logic
```javascript
testWeeklyStarsLogic()
```

## Ví dụ thực tế

### Gửi lại báo cáo ngày 15/7/2025
```javascript
sendReportForDate('2025-07-15')
```

### Gửi báo cáo cho tuần trước
```javascript
sendReportForLastSunday()
```

### Gửi báo cáo cho ngày làm việc cuối tuần trước
```javascript
sendReportForDate('2025-07-26') // Thứ 6 tuần trước
```

## Lưu ý
- Chỉ có thể gửi báo cáo cho các ngày đã có dữ liệu trong sheet
- Nếu không tìm thấy cột ngày trong sheet → function sẽ dừng và log error
- Custom date không ảnh hưởng đến logic scheduling tự động