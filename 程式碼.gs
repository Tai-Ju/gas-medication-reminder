

// @ts-nocheck
// 全局變量
const SPREADSHEET_ID = '1XQ1OR8QdmEAT74B5CCs6siCmfQ9Ueb4K8wIJinc5sAo';
const SHEET_NAME = '假卡申請';
const OVERTIME_SHEET_NAME = '加班費申請';


// 定義津貼金額常數
const ALLOWANCE_RATES = {
  '夜診18:00-21:00': 400,
  '夜診18:00-21:30': 400,
  '夜診18:00-22:00': 400,
  '小夜半': 300,
  '小小班': 300,
  '小夜全': 600,
  '大夜(8+4)': 700,
  '大夜(10+2)': 800
};


function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('員工打卡記錄查詢系統')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}


function getEmployeeData(empId) {
  try {
    if (!empId) throw new Error('請輸入員工編號');
    
    // 標準化員工編號
    const searchEmpId = empId.toString().trim().toUpperCase();
    console.log('開始搜尋員工:', searchEmpId);
    
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('找不到工作表');
    
    const data = sheet.getDataRange().getValues();
    const records = [];
    let found = false;
    let startIndex = -1;
    
    // 遍歷數據
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const currentEmpId = row[2] ? row[2].toString().trim().toUpperCase() : '';
      
      // 如果已經找到匹配項，且當前員工編號不同，則停止搜尋
      if (found && currentEmpId !== searchEmpId) {
        break;
      }
      
      if (currentEmpId === searchEmpId) {
        if (!found) {
          found = true;
          startIndex = i;
          console.log('找到第一個匹配記錄:', {
            行號: i + 1,
            員工編號: currentEmpId,
            員工姓名: row[3]
          });
        }
        
        // 只針對 timeCode, startTime, endTime, hours 做 fallback
        const timeCode = row[27] ? row[27] : (row[17] || '');
        const startTime = row[28] ? row[28] : (row[18] || '');
        const endTime = row[29] ? row[29] : (row[19] || '');
        const hours = row[30] ? row[30] : (row[20] || '');

        // 確保日期格式正確
        const formattedDate = row[4] ? formatDate(row[4]) : '';
        
        records.push({
          ud_code: row[1] || '',
          emp_id: row[2] || '',
          emp_name: row[3] || '',
          date: formattedDate,
          weekday: row[5] || '',
          punch3: formatTime(row[8]),
          punch4: formatTime(row[9]),
          punch5: formatTime(row[10]),
          punch6: formatTime(row[11]),
          workType: row[12] || '',
          status: row[13] || '',
          overtimeHours: row[14] || '',
          compensationType: row[15] || '',
          leaveHours: row[16] || '',
          timeCode: timeCode,
          startTime: startTime ? formatTime(startTime) : '',
          endTime: endTime ? formatTime(endTime) : '',
          hours: hours,
          originalTimeCode: row[17] || '',
          originalStartTime: row[18] ? formatTime(row[18]) : '',
          originalEndTime: row[19] ? formatTime(row[19]) : '',
          originalHours: row[20] || ''
        });
      }
    }

    if (!found) {
      console.log('未找到匹配的員工記錄');
      throw new Error(`查無此員工 (搜尋編號: ${searchEmpId})`);
    }
    
    const overtimeData = getOvertimeData(empId);
    console.log('getEmployeeData 最終回傳:', { 
      attendance: records, 
      overtime: overtimeData,
      找到的記錄數: records.length,
      起始行號: startIndex + 1
    });
    return { attendance: records, overtime: overtimeData };
  } catch (error) {
    console.error('查詢失敗:', error);
    throw new Error(`查詢失敗: ${error.message}`);
  }
}

function getOvertimeData(empId) {
  try {
    const overtimeSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(OVERTIME_SHEET_NAME);
    if (!overtimeSheet) {
      throw new Error('找不到加班費申請單工作表');
    }

    const data = overtimeSheet.getDataRange().getValues();
    empId = empId.toString().toUpperCase().trim();
    
    console.log('開始查找加班費申請單數據:', {
      empId: empId,
      totalRows: data.length
    });
    
    // 查找現有記錄
    const existingRecord = data.find((row, index) => {
      if (index === 0) return false;
      const rowEmpId = row[2] ? row[2].toString().trim().toUpperCase() : '';
      const isMatch = rowEmpId === empId;
      
      if (isMatch) {
        console.log('找到匹配記錄:', {
          rowIndex: index + 1,
          empId: rowEmpId,
          empName: row[1]
        });
      }
      
      return isMatch;
    });

    // 初始化返回數據結構
    const result = {
      shiftDates: {},
      employeeName: '',
      saturdayData: Array(5).fill(null).map(() => ({ date: '', time: '' })),
      timeData: {},
      columnData: {}
    };

    if (existingRecord) {
      console.log('處理現有記錄數據');
      
      // 設置員工姓名
      result.employeeName = existingRecord[1] || '';  // B欄：員工姓名

      // 讀取夜診資料（I-K欄）
      const nightShiftTypes = ['夜診18:00-21:00', '夜診18:00-21:30', '夜診18:00-22:00'];
      nightShiftTypes.forEach((type, index) => {
        const dates = existingRecord[8 + index];  // I, J, K欄
        if (dates) {
          result.shiftDates[type] = dates.toString().split(',')
            .map(d => d.trim())
            .filter(d => d);
        }
      });

      // 讀取週六資料（L-U欄）
      for (let i = 0; i < 5; i++) {
        const dateIndex = 11 + (i * 2);  // L, N, P, R, T欄
        const timeIndex = dateIndex + 1;  // M, O, Q, S, U欄
        result.saturdayData[i] = {
          date: existingRecord[dateIndex] ? String(existingRecord[dateIndex]).padStart(2, '0') : '',
          time: existingRecord[timeIndex] || ''
        };
      }

      // 讀取小夜和大夜班資料
      const shiftColumns = {
        '小夜半': { dates: 22, time: 23 },     // W, X欄
        '小小班': { dates: 25 },               // Z欄
        '小夜全': { dates: 27 },               // AB欄
        '大夜(8+4)': { dates: 29, time: 30 },  // AD, AE欄
        '大夜(10+2)': { dates: 32, time: 33 }  // AG, AH欄
      };

      Object.entries(shiftColumns).forEach(([type, { dates, time }]) => {
        const datesValue = existingRecord[dates];
        if (datesValue) {
          result.shiftDates[type] = datesValue.toString().split(',')
            .map(d => d.trim())
            .filter(d => d);
        }
        if (time !== undefined) {
          result.timeData[type] = existingRecord[time] || '';
        }
      });

      // 讀取 columnData
      result.columnData = {
        AA: existingRecord[26] || '',  // AA欄：小夜全天數
        AB: existingRecord[27] || '',  // AB欄：小夜全日期
        AC: existingRecord[28] || '',  // AC欄：大夜(8+4)天數
        AD: existingRecord[29] || '',  // AD欄：大夜(8+4)日期
        AE: existingRecord[30] || '',  // AE欄：大夜(8+4)時間
        AF: existingRecord[31] || '',  // AF欄：大夜(10+2)天數
        AG: existingRecord[32] || '',  // AG欄：大夜(10+2)日期
        AH: existingRecord[33] || ''   // AH欄：大夜(10+2)時間
      };
    } else {
      console.log('未找到現有記錄，返回空數據結構');
    }

    console.log('返回的加班費申請單資料:', result);
    return result;

  } catch (error) {
    console.error('獲取加班費申請資料失敗:', error);
    throw new Error(`獲取加班費申請資料失敗: ${error.message}`);
  }
}


function saveEmployeeData(empId, formData) {
  try {
    // 驗證輸入資料
    if (!empId || !formData) {
      throw new Error('缺少必要的資料');
    }

    // 確保 formData 有正確的結構
    formData.attendance = formData.attendance || [];
    formData.overtime = formData.overtime || {};
    formData.activeTab = formData.activeTab || '';

    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) {
      throw new Error('無法找到工作表');
    }

    const data = sheet.getDataRange().getValues();
    let errors = [];
    let successCount = 0;

    console.log('開始處理資料:', {
      empId,
      recordCount: formData.attendance.length
    });

    // 更新打卡記錄
    formData.attendance.forEach((entry, index) => {
      // 確保 entry 物件有所有必要的屬性
      entry = {
        date: entry.date || '',
        workType: entry.workType || '',
        status: entry.status || '',
        overtimeHours: entry.overtimeHours || '',
        compensationType: entry.compensationType || '',
        leaveHours: entry.leaveHours || '',
        actualTimeCode: entry.actualTimeCode || '',
        actualStartTime: entry.actualStartTime || '',
        actualEndTime: entry.actualEndTime || '',
        actualHours: entry.actualHours || ''
      };

      // 跳過空白記錄
      if (!entry.date) return;
      
      console.log(`處理第 ${index + 1} 筆記錄:`, entry);
      
      const rowIndex = findRowIndex(data, empId, entry.date);
      console.log('找到的行索引:', rowIndex);
      
      if (rowIndex > 0) {
        try {
          // 更新 M欄(班別/工作說明)和 N欄(異常狀態)
          sheet.getRange(rowIndex + 1, 13, 1, 2).setValues([[
            entry.workType,
            entry.status
          ]]);
          
          // 更新其他欄位 (O-Q欄)
          sheet.getRange(rowIndex + 1, 15, 1, 3).setValues([[
            entry.overtimeHours,
            entry.compensationType,
            entry.leaveHours
          ]]);

          // 更新實際值欄位 (AB-AE欄)
          sheet.getRange(rowIndex + 1, 28, 1, 4).setValues([[
            entry.actualTimeCode,
            entry.actualStartTime,
            entry.actualEndTime,
            entry.actualHours
          ]]);
          
          successCount++;
        } catch (error) {
          const errorMsg = `更新第 ${rowIndex + 1} 行時發生錯誤: ${error.message}`;
          console.error(errorMsg);
          errors.push(errorMsg);
        }
      }
    });

    // 自動產生加班費申請單數據
    const shiftDates = {};
    const saturdayData = Array(5).fill(null).map(() => ({ date: '', time: '' }));
    let saturdayDates = [];
    
    // 從打卡記錄中收集加班費申請單數據
    formData.attendance.forEach(record => {
      const date = record.date;
      const workType = record.workType;
      const compensationType = record.compensationType;
      
      if (workType === '週六(休假日)' && compensationType === '加班費') {
        const day = date.split('-')[2];
        saturdayDates.push(day);
      }
      
      if (ALLOWANCE_RATES[workType] && compensationType === '加班費' && workType !== '週六(休假日)') {
        if (!shiftDates[workType]) {
          shiftDates[workType] = [];
        }
        const day = date.split('-')[2];
        if (!shiftDates[workType].includes(day)) {
          shiftDates[workType].push(day);
        }
      }
    });
    
    // 排序週六數據
    saturdayDates.sort((a, b) => parseInt(a) - parseInt(b));
    saturdayDates.forEach((date, index) => {
      if (index < 5) {
        saturdayData[index] = { date: date, time: '' };
      }
    });

    // 更新加班費申請單
    const overtimeSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(OVERTIME_SHEET_NAME);
    if (!overtimeSheet) {
      throw new Error('無法找到加班費申請單工作表');
    }
    
    let overtimeRowIndex = findOvertimeRow(overtimeSheet, empId);
    console.log('加班費申請單行索引:', overtimeRowIndex);
    
    // 如果找不到現有記錄，在最後添加新行
    if (overtimeRowIndex === -1) {
      overtimeRowIndex = overtimeSheet.getLastRow() + 1;
    }

    // 計算各類津貼和天數
    let nightShiftAllowance = 0;
    let shiftAllowance = 0;
    let totalNightShiftDays = 0;
    
    // 夜診組數據
    const nightShiftTypes = ['夜診18:00-21:00', '夜診18:00-21:30', '夜診18:00-22:00'];
    const nightShiftData = nightShiftTypes.map(shiftType => {
      const dates = shiftDates[shiftType] || [];
      const rate = ALLOWANCE_RATES[shiftType] || 0;
      nightShiftAllowance += dates.length * rate;
      totalNightShiftDays += dates.length;
      return {
        count: dates.length || '',
        dates: dates.sort((a, b) => parseInt(a) - parseInt(b)).join(',') || ''
      };
    });

    // 小夜和大夜組數據
    const shiftTypes = ['小夜半', '小小班', '小夜全', '大夜(8+4)', '大夜(10+2)'];
    const shiftData = shiftTypes.map(shiftType => {
      const dates = shiftDates[shiftType] || [];
      const rate = ALLOWANCE_RATES[shiftType] || 0;
      shiftAllowance += dates.length * rate;
      return {
        count: dates.length || '',
        dates: dates.sort((a, b) => parseInt(a) - parseInt(b)).join(',') || '',
        time: formData.overtime.timeData?.[shiftType] || ''
      };
    });

    // 準備要寫入的數據
    const rowData = new Array(42).fill('');

    // 基本資料
    rowData[0] = overtimeRowIndex;                // A欄：序號
    rowData[1] = formData.overtime.employeeName || getEmployeeName(empId);  // B欄：員工姓名
    rowData[2] = empId;                          // C欄：員工編號
    rowData[3] = totalNightShiftDays;            // D欄：夜診總天數
    rowData[4] = shiftAllowance;                 // E欄：大小夜津貼
    rowData[5] = nightShiftAllowance;            // F欄：夜診津貼
    rowData[6] = shiftAllowance + nightShiftAllowance; // G欄：津貼合計
    rowData[7] = totalNightShiftDays;            // H欄：夜診天數

    // 夜診組數據
    rowData[8] = nightShiftData[0].dates;        // I欄：夜診18:00-21:00日期
    rowData[9] = nightShiftData[1].dates;        // J欄：夜診18:00-21:30日期
    rowData[10] = nightShiftData[2].dates;       // K欄：夜診18:00-22:00日期

    // 週六組數據
    saturdayData.forEach((data, index) => {
      const baseIndex = 11 + (index * 2);
      rowData[baseIndex] = data.date || '';      // L,N,P,R,T欄：日期
      rowData[baseIndex + 1] = data.time || '';  // M,O,Q,S,U欄：時間
    });

    // 小夜半
    rowData[21] = shiftData[0].count;           // V欄：小夜半天數
    rowData[22] = shiftData[0].dates;           // W欄：小夜半日期
    rowData[23] = shiftData[0].time;            // X欄：小夜半時間

    // 小小班
    rowData[24] = shiftData[1].count;           // Y欄：小小班天數
    rowData[25] = shiftData[1].dates;           // Z欄：小小班日期

    // 小夜全
    rowData[26] = shiftData[2].count;           // AA欄：小夜全天數
    rowData[27] = shiftData[2].dates;           // AB欄：小夜全日期

    // 大夜(8+4)
    rowData[28] = shiftData[3].count;           // AC欄：大夜(8+4)天數
    rowData[29] = shiftData[3].dates;           // AD欄：大夜(8+4)日期
    rowData[30] = shiftData[3].time;            // AE欄：大夜(8+4)時間

    // 大夜(10+2)
    rowData[31] = shiftData[4].count;           // AF欄：大夜(10+2)天數
    rowData[32] = shiftData[4].dates;           // AG欄：大夜(10+2)日期
    rowData[33] = shiftData[4].time;            // AH欄：大夜(10+2)時間

    console.log('準備寫入加班費申請單數據:', rowData);

    // 寫入加班費申請單數據
    overtimeSheet.getRange(overtimeRowIndex, 1, 1, rowData.length).setValues([rowData]);

    // 強制重新載入數據
    SpreadsheetApp.flush();
    
    // 返回更新後的數據
    const updatedData = getEmployeeData(empId);
    if (!updatedData || !updatedData.attendance || updatedData.attendance.length === 0) {
      throw new Error('更新後無法獲取資料');
    }
    
    return updatedData;

  } catch (error) {
    console.error('保存失敗:', error);
    throw new Error(`保存失敗: ${error.message}`);
  }
}


function calculateAllowance(formData) {
  const allowance = {
    nightShiftAllowance: 0,
    shiftAllowance: 0,
    totalAllowance: 0
  };


  const shiftCounts = {};
  formData.forEach(entry => {
    if (entry.workType) {
      shiftCounts[entry.workType] = (shiftCounts[entry.workType] || 0) + 1;
    }
  });


  ['小夜半', '小小班', '小夜全', '大夜(8+4)', '大夜(10+2)'].forEach(shiftType => {
    if (shiftCounts[shiftType]) {
      allowance.shiftAllowance += shiftCounts[shiftType] * ALLOWANCE_RATES[shiftType];
    }
  });


  if (shiftCounts['夜診']) {
    allowance.nightShiftAllowance = shiftCounts['夜診'] * ALLOWANCE_RATES['夜診'];
  }


  allowance.totalAllowance = allowance.shiftAllowance + allowance.nightShiftAllowance;


  return allowance;
}


function updateOvertimeSheet(empId, allowance, overtimeData) {
  const overtimeSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(OVERTIME_SHEET_NAME);
  if (!overtimeSheet) return;


  let rowIndex = findEmployeeRow(overtimeSheet, empId);
  if (rowIndex === -1) {
    rowIndex = overtimeSheet.getLastRow() + 1;
    overtimeSheet.getRange(rowIndex, 1, 1, 3).setValues([[
      rowIndex - 1,
      getEmployeeName(empId),
      empId
    ]]);
  }


  // 更新津貼金額
  overtimeSheet.getRange(rowIndex, 3, 1, 3).setValues([[
    allowance.nightShiftAllowance,
    allowance.shiftAllowance,
    allowance.totalAllowance
  ]]);


  // 更新其他欄位
  if (overtimeData) {
    // 這裡可以添加更新其他欄位的邏輯
  }
}


function findRowIndex(data, empId, date) {
  // 標準化輸入
  empId = empId.toString().trim().toUpperCase();
  
  // 添加日誌
  console.log('findRowIndex 開始搜尋:', {
    empId: empId,
    date: date,
    totalRows: data.length
  });
  
  // 遍歷數據
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowEmpId = row[2] ? row[2].toString().trim().toUpperCase() : '';
    const rowDate = row[4] ? formatDate(row[4]) : '';
    
    // 添加詳細日誌
    console.log(`檢查第 ${i + 1} 行:`, {
      rowEmpId: rowEmpId,
      rowDate: rowDate,
      targetDate: date,
      isEmpMatch: rowEmpId === empId,
      isDateMatch: rowDate === date
    });
    
    if (rowEmpId === empId && rowDate === date) {
      console.log(`找到匹配記錄: 第 ${i + 1} 行`);
      return i;
    }
  }
  
  console.log('未找到匹配記錄');
  return -1;
}


function findEmployeeRow(sheet, empId) {
  const data = sheet.getDataRange().getValues();
  return data.findIndex((row, index) =>
    index > 0 && String(row[2]).trim() === String(empId).trim()
  ) + 1;
}


function getEmployeeName(empId) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const employeeRow = data.find(row => String(row[2]).trim() === String(empId).trim());
  return employeeRow ? employeeRow[3] : '';
}


function formatDate(date) {
  if (!date) return '';
  if (typeof date === 'string') return date;
  return Utilities.formatDate(new Date(date), 'GMT+8', 'yyyy-MM-dd');
}


function formatTime(time) {
  if (!time) return '';
  const timeStr = time.toString().trim();
  
  // 如果是数字（Excel时间格式），转换为标准时间格式
  if (typeof time === 'number') {
    const totalMinutes = Math.round(time * 24 * 60);
    const hours = Math.floor(totalMinutes / 60) % 24;
    const minutes = totalMinutes % 60;
    return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
  }
  
  // 如果是标准时间格式（HH:MM:SS），只取HH:MM部分
  const timeMatch = timeStr.match(/(\d{1,2}):(\d{2})/);
  if (timeMatch) {
    const [_, hours, minutes] = timeMatch;
    return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
  }
  
  // 如果是GMT时间格式
  if (timeStr.includes('GMT')) {
    const date = new Date(timeStr);
    return date.toTimeString().split(' ')[0].substring(0, 5);
  }
  
  return timeStr;
}


function formatTimeOnly(timeStr) {
  if (!timeStr) return '';
  try {
    if (typeof timeStr === 'number') {
      const totalMinutes = Math.round(timeStr * 24 * 60);
      // 減去23分鐘
      //const adjustedMinutes = totalMinutes - 23;
      const adjustedMinutes = totalMinutes - 23;
      const hours = Math.floor(adjustedMinutes / 60) % 24;
      const minutes = adjustedMinutes % 60;
      return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
    }


    const timeMatch = timeStr.toString().match(/(\d{1,2}):(\d{2})/);
    if (timeMatch) {
      const [_, hoursStr, minutesStr] = timeMatch;
      // 將時間轉換為分鐘並減去23分鐘
      let totalMinutes = parseInt(hoursStr) * 60 + parseInt(minutesStr) - 23;
      // 處理跨日的情況
      if (totalMinutes < 0) {
        totalMinutes += 24 * 60;
      }
      const hours = Math.floor(totalMinutes / 60) % 24;
      const minutes = totalMinutes % 60;
      return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
    }


    return '';
  } catch (error) {
    console.error('時間格式化錯誤:', error, timeStr);
    return '';
  }
}


function getScheduleData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('說明');
    if (!sheet) {
      throw new Error('找不到"說明"工作表');
    }

    // 直接獲取範圍的值
    const range = sheet.getRange('B1:G14');
    // 獲取格式化的文字值
    const values = range.getDisplayValues();
    
    // 檢查是否有數據
    if (!values || values.length === 0) {
      throw new Error('範圍內沒有數據');
    }
    
    // 檢查並清理數據
    const cleanedValues = values.map(row => 
      row.map(cell => 
        // 保留原始值，包括空白
        cell ? cell.toString() : ''
      )
    );

    return cleanedValues;
  } catch (error) {
    console.error('獲取排班表數據失敗:', error);
    throw new Error(`獲取數據失敗: ${error.message}`);
  }
}


// 尋找員工在加班費申請單中的列位置
function findOvertimeRow(sheet, empId) {
  const data = sheet.getDataRange().getValues();
  console.log('開始查找員工:', empId);
  console.log('工作表數據行數:', data.length);
  
  for (let i = 1; i < data.length; i++) {
    const currentEmpId = data[i][2] ? data[i][2].toString().trim().toUpperCase() : '';
    console.log('比較員工編號:', currentEmpId, '與', empId.toString().trim().toUpperCase());
    
    if (currentEmpId === empId.toString().trim().toUpperCase()) {
      console.log('找到匹配的行:', i + 1);
      return i + 1;
    }
  }
  console.log('未找到匹配的行');
  return -1;
}


// 将时间转换为分钟数的辅助函数
function timeToMinutes(time) {
  if (!time) return NaN;
  
  // 如果是数字（Excel时间格式），转换为分钟
  if (typeof time === 'number') {
    return Math.round(time * 24 * 60);
  }
  
  // 如果是字符串，解析时间格式
  const timeStr = time.toString();
  const match = timeStr.match(/(\d{1,2}):(\d{2})/);
  if (match) {
    const [_, hours, minutes] = match;
    return parseInt(hours) * 60 + parseInt(minutes);
  }
  
  return NaN;
}





