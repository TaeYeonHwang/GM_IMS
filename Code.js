// 구글시트 설정 값 가져오기.
// Code.gs
function getConfig() {
  const scriptProperties = PropertiesService.getScriptProperties();
  return {
    SPREADSHEET_ID: scriptProperties.getProperty('SPREADSHEET_ID'),
    GUIDE_IMAGE_ID: scriptProperties.getProperty('GUIDE_IMAGE_ID')
  };
}

const config = getConfig();
const SPREADSHEET_ID = config.SPREADSHEET_ID;
const SHEET_NAME = 'ItemInfo';
const LOG_SHEET_NAME = 'AccessLog';
const PURCHASE_ORDER_SHEET_NAME = 'PurchaseOrder';

// HTML 파일 include 함수.
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// GUIDE_IMAGE_ID 가져오기 함수
function getGuideImageId() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const imageId = scriptProperties.getProperty('GUIDE_IMAGE_ID');
  
  if (!imageId) {
    Logger.log('GUIDE_IMAGE_ID가 설정되지 않았습니다.');
    return null;
  }
  
  return imageId;
}

// ✅ 웹 앱 진입점 수정 (.evaluate() 추가)
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('바코드 스캔 재고 관리 시스템')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 접근 로그 기록.
function logAccess(codeNum, userIP) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let logSheet = ss.getSheetByName(LOG_SHEET_NAME);
    
    if (!logSheet) {
      logSheet = ss.insertSheet(LOG_SHEET_NAME);
      logSheet.appendRow(['Access_IP', 'Time', 'Scaned_CodeNum']);
    }
    
    const currentTime = new Date();
    logSheet.appendRow([userIP, currentTime, codeNum]);
    
    return { success: true };
  } catch (error) {
    Logger.log('로그 기록 오류: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// 코드번호로 아이템 검색
function searchByCodeNum(codeNum) {
  try {
    // 입력값 검증
    if (!codeNum || codeNum.toString().trim() === '') {
      return {
        success: false,
        message: '코드번호를 입력해주세요.'
      };
    }
    
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      return {
        success: false,
        message: 'ItemInfo 시트를 찾을 수 없습니다.'
      };
    }
    
    // ✅ 마지막 행만 확인하여 필요한 범위만 읽기
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      return {
        success: false,
        message: '등록된 품목이 없습니다.'
      };
    }
    
    // 헤더(1행) 제외, 2행부터 lastRow까지, 1열부터 7열(StockNum)까지만 읽기
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 7);
    const data = dataRange.getValues();
    
    const searchCode = codeNum.toString().trim();
    
    // ✅ 더 효율적인 검색 (Array.find 사용)
    const foundIndex = data.findIndex(row => 
      row[3] && row[3].toString().trim() === searchCode
    );
    
    if (foundIndex !== -1) {
      const row = data[foundIndex];
      return {
        success: true,
        item: {
          serialNum: row[0],
          name: row[1],
          description: row[2],
          codeNum: row[3],
          costB2B: row[4],
          costB2C: row[5],
          stockNum: row[6]
        },
        rowNumber: foundIndex + 2  // 실제 시트의 행 번호
      };
    }
    
    return {
      success: false,
      message: '해당 코드번호를 찾을 수 없습니다: ' + searchCode
    };
    
  } catch (error) {
    Logger.log('searchByCodeNum 오류: ' + error.toString());
    return {
      success: false,
      message: '검색 중 오류가 발생했습니다: ' + error.message
    };
  }
}

// 전체 아이템 목록 가져오기
function getAllItems() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    const items = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        const item = {
          serialNum: data[i][0],
          name: data[i][1],
          description: data[i][2],
          codeNum: data[i][3],
          costB2B: data[i][4],
          costB2C: data[i][5],
          stockNum: data[i][6]
        };
        items.push(item);
      }
    }
    
    return {
      success: true,
      items: items
    };
  } catch (error) {
    return {
      success: false,
      message: '오류가 발생했습니다: ' + error.toString()
    };
  }
}

// Revision 정보 가져오기
function getLatestRevision() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('RevisionHistory');
    
    if (!sheet) {
      return {
        success: false,
        message: 'RevisionHistory 시트를 찾을 수 없습니다.'
      };
    }
    
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return {
        success: false,
        message: 'Revision 데이터가 없습니다.'
      };
    }
    
    const headers = data[0];
    const revisionCol = headers.indexOf('Revision');
    const dateCol = headers.indexOf('Date');
    
    if (revisionCol === -1 || dateCol === -1) {
      return {
        success: false,
        message: 'Revision 또는 Date 열을 찾을 수 없습니다.'
      };
    }
    
    let maxRevision = -1;
    let maxRevisionDate = '';
    
    for (let i = 1; i < data.length; i++) {
      const revision = parseFloat(data[i][revisionCol]);
      if (!isNaN(revision) && revision > maxRevision) {
        maxRevision = revision;
        maxRevisionDate = data[i][dateCol];
      }
    }
    
    if (maxRevision === -1) {
      return {
        success: false,
        message: '유효한 Revision 값을 찾을 수 없습니다.'
      };
    }
    
    if (maxRevisionDate instanceof Date) {
      maxRevisionDate = Utilities.formatDate(maxRevisionDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    
    return {
      success: true,
      revision: maxRevision,
      date: maxRevisionDate
    };
  } catch (error) {
    Logger.log('Revision 조회 오류: ' + error.toString());
    return {
      success: false,
      message: '오류가 발생했습니다: ' + error.toString()
    };
  }
}

// 전체 Revision History 가져오기
function getRevisionHistory() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('RevisionHistory');
    
    if (!sheet) {
      return {
        success: false,
        message: 'RevisionHistory 시트를 찾을 수 없습니다.'
      };
    }
    
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return {
        success: false,
        message: 'Revision 데이터가 없습니다.'
      };
    }
    
    const headers = data[0];
    const targetColumns = ['Revision', 'Author', 'Date', 'Description'];
    const columnIndices = {};
    
    targetColumns.forEach(col => {
      const index = headers.indexOf(col);
      if (index !== -1) {
        columnIndices[col] = index;
      }
    });
    
    const missingColumns = targetColumns.filter(col => columnIndices[col] === undefined);
    if (missingColumns.length > 0) {
      return {
        success: false,
        message: '다음 열을 찾을 수 없습니다: ' + missingColumns.join(', ')
      };
    }
    
    const revisions = [];
    for (let i = 1; i < data.length; i++) {
      const row = {};
      targetColumns.forEach(col => {
        let value = data[i][columnIndices[col]];
        if (value instanceof Date) {
          value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        row[col] = value;
      });
      revisions.push(row);
    }
    
    revisions.sort((a, b) => {
      const revA = parseFloat(a.Revision) || 0;
      const revB = parseFloat(b.Revision) || 0;
      return revB - revA;
    });
    
    return {
      success: true,
      headers: targetColumns,
      revisions: revisions
    };
  } catch (error) {
    Logger.log('Revision History 조회 오류: ' + error.toString());
    return {
      success: false,
      message: '오류가 발생했습니다: ' + error.toString()
    };
  }
}

// 신규 주문서 번호 가져오기
function getNewOrderIndex(orderDate) {
  try {
    // ✅ 입력값 검증 추가
    if (!orderDate) {
      return {
        success: false,
        message: '주문 날짜가 제공되지 않았습니다.'
      };
    }
    
    // 날짜 형식 검증 (YYYYMMDD, 8자리 숫자)
    const dateStr = orderDate.toString();
    if (dateStr.length !== 8 || isNaN(dateStr)) {
      return {
        success: false,
        message: '올바르지 않은 날짜 형식입니다. YYYYMMDD 형식이어야 합니다. (입력값: ' + dateStr + ')'
      };
    }
    
    // 날짜 유효성 검증
    const year = parseInt(dateStr.substring(0, 4));
    const month = parseInt(dateStr.substring(4, 6));
    const day = parseInt(dateStr.substring(6, 8));
    
    if (year < 2000 || year > 2100) {
      return {
        success: false,
        message: '유효하지 않은 연도입니다: ' + year
      };
    }
    
    if (month < 1 || month > 12) {
      return {
        success: false,
        message: '유효하지 않은 월입니다: ' + month
      };
    }
    
    if (day < 1 || day > 31) {
      return {
        success: false,
        message: '유효하지 않은 일입니다: ' + day
      };
    }
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // ✅ 스프레드시트 접근 검증
    if (!ss) {
      return {
        success: false,
        message: '스프레드시트에 접근할 수 없습니다. ID를 확인해주세요.'
      };
    }
    
    let sheet = ss.getSheetByName(PURCHASE_ORDER_SHEET_NAME);
    
    if (!sheet) {
      Logger.log('PurchaseOrder sheet does not exist, attempting to create...');
      
      // ✅ Updated: Added PayType and IsCanceled columns
      try {
        sheet = ss.insertSheet(PURCHASE_ORDER_SHEET_NAME);
        sheet.appendRow([
          'Order_SerialNumber', 
          'Order_Date', 
          'Order_Index', 
          'Order_CodeNum',
          'Order_Name', 
          'Order_Description', 
          'Order_CostB2B', 
          'Order_CostB2C', 
          'Order_IsB2B', 
          'Order_Cnt',
          'PayType',           // ✅ NEW
          'Order_TotalCost',
          'IsCanceled'         // ✅ NEW
        ]);
        Logger.log('PurchaseOrder sheet created successfully.');
        
        return {
          success: true,
          orderIndex: 1,
          message: 'New sheet created.'
        };
      } catch (createError) {
        return {
          success: false,
          message: 'Failed to find or create PurchaseOrder sheet: ' + createError.toString()
        };
      }
    }
    
    const data = sheet.getDataRange().getValues();
    
    // ✅ 데이터 없음 = 첫 주문
    if (data.length <= 1) {
      return {
        success: true,
        orderIndex: 1,
        message: '첫 번째 주문서입니다.'
      };
    }
    
    const headers = data[0];
    const dateColIndex = headers.indexOf('Order_Date');
    const indexColIndex = headers.indexOf('Order_Index');
    
    // ✅ 컬럼 존재 검증
    if (dateColIndex === -1) {
      return {
        success: false,
        message: 'Order_Date 열을 찾을 수 없습니다. 시트 구조를 확인해주세요.'
      };
    }
    
    if (indexColIndex === -1) {
      return {
        success: false,
        message: 'Order_Index 열을 찾을 수 없습니다. 시트 구조를 확인해주세요.'
      };
    }
    
    let maxIndex = 0;
    let sameDataCount = 0;
    
    // 같은 날짜의 최대 인덱스 찾기
    for (let i = 1; i < data.length; i++) {
      const rowDate = data[i][dateColIndex];
      const rowIndex = parseInt(data[i][indexColIndex]) || 0;
      
      if (rowDate && rowDate.toString() === orderDate.toString()) {
        sameDataCount++;
        if (rowIndex > maxIndex) {
          maxIndex = rowIndex;
        }
      }
    }
    
    const newIndex = maxIndex + 1;
    
    // ✅ 인덱스 범위 검증 (9999까지만)
    if (newIndex > 9999) {
      return {
        success: false,
        message: '하루 최대 주문서 개수(9999)를 초과했습니다.'
      };
    }
    
    // ✅ 기존 주문서 개수 = 마지막 주문서 번호
    const existingOrderCount = maxIndex;

    Logger.log(`날짜 ${orderDate}의 주문서: 기존 ${existingOrderCount}개, 새 인덱스: ${newIndex}`);
    
    return {
      success: true,
      orderIndex: newIndex,
      existingOrders: existingOrderCount
    };
    
  } catch (error) {
    // ✅ 자세한 에러 로깅
    const errorDetails = {
      message: error.message,
      stack: error.stack,
      orderDate: orderDate,
      timestamp: new Date().toISOString()
    };
    
    Logger.log('신규 주문서 번호 조회 오류: ' + JSON.stringify(errorDetails, null, 2));
    
    return {
      success: false,
      message: '오류가 발생했습니다: ' + error.message
    };
  }
}

// 주문서 조회
function getOrderData(orderDate, orderIndex) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(PURCHASE_ORDER_SHEET_NAME);
    
    if (!sheet) {
      return {
        success: false,
        message: 'PurchaseOrder sheet not found.'
      };
    }
    
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return {
        success: false,
        message: 'No order data found.'
      };
    }
    
    const headers = data[0];
    const colIndices = {};
    const requiredCols = [
      'Order_SerialNumber', 'Order_Date', 'Order_Time', 'Order_Index', 
      'Order_CodeNum', 'Order_Name', 'Order_Description', 
      'Order_CostB2B', 'Order_CostB2C', 'Order_IsB2B', 'Order_Cnt', 
      'PayType', 'Order_TotalCost', 'IsCanceled'
    ];
    
    requiredCols.forEach(col => {
      const index = headers.indexOf(col);
      if (index !== -1) {
        colIndices[col] = index;
      }
    });
    
    const orderIndexStr = orderIndex.toString().padStart(4, '0');
    const targetSerialNumber = orderDate.toString() + orderIndexStr;
    const orders = [];
    
    for (let i = 1; i < data.length; i++) {
      const rowSerialNumber = data[i][colIndices['Order_SerialNumber']];
      
      if (rowSerialNumber && rowSerialNumber.toString() === targetSerialNumber) {
        // ✅ Order_Time 데이터 처리
        let orderTime = data[i][colIndices['Order_Time']];
        if (orderTime instanceof Date) {
          orderTime = Utilities.formatDate(orderTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
        } else if (orderTime) {
          orderTime = orderTime.toString();
        } else {
          orderTime = '-';  // 값이 없으면 '-'
        }

        // ✅ Changed: Check if IsCanceled is '취소' text
        const canceledValue = data[i][colIndices['IsCanceled']];
        const isCanceled = canceledValue === '취소';

        orders.push({
          serialNumber: data[i][colIndices['Order_SerialNumber']],
          date: data[i][colIndices['Order_Date']],
          time: orderTime,
          index: data[i][colIndices['Order_Index']],
          codeNum: data[i][colIndices['Order_CodeNum']],
          name: data[i][colIndices['Order_Name']],
          description: data[i][colIndices['Order_Description']],
          costB2B: data[i][colIndices['Order_CostB2B']],
          costB2C: data[i][colIndices['Order_CostB2C']],
          isB2B: data[i][colIndices['Order_IsB2B']],
          cnt: data[i][colIndices['Order_Cnt']],
          payType: data[i][colIndices['PayType']] || '-',              // ✅ NEW
          totalCost: data[i][colIndices['Order_TotalCost']],
          isCanceled: isCanceled           // ✅ NEW
        });
      }
    }
    
    if (orders.length === 0) {
      return {
        success: false,
        message: 'No orders found for the specified date and order number.'
      };
    }
    
    return {
      success: true,
      orders: orders
    };
    
  } catch (error) {
    Logger.log('Error retrieving order: ' + error.toString());
    return {
      success: false,
      message: 'An error occurred: ' + error.toString()
    };
  }
}

// 주문 저장
function saveOrder(orderData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const purchaseSheet = ss.getSheetByName(PURCHASE_ORDER_SHEET_NAME);
    const itemSheet = ss.getSheetByName(SHEET_NAME);
    
    if (!purchaseSheet || !itemSheet) {
      return { success: false, message: 'Required sheets not found.' };
    }
    
    const itemData = itemSheet.getDataRange().getValues();
    const itemHeaders = itemData[0];
    const codeNumColIndex = itemHeaders.indexOf('CodeNum');
    const stockNumColIndex = itemHeaders.indexOf('StockNum');
    
    if (codeNumColIndex === -1 || stockNumColIndex === -1) {
      return { success: false, message: 'Required columns not found in ItemInfo sheet.' };
    }
    
    // Step 1: Validate stock for all items
    const stockValidation = [];
    for (let item of orderData.items) {
      let found = false;
      for (let i = 1; i < itemData.length; i++) {
        if (itemData[i][codeNumColIndex] && 
            itemData[i][codeNumColIndex].toString() === item.codeNum.toString()) {
          const currentStock = itemData[i][stockNumColIndex] || 0;
          
          // Validate stock availability
          if (currentStock < item.cnt) {
            return {
              success: false,
              message: `Insufficient stock: ${item.name} (Requested: ${item.cnt}, Available: ${currentStock})`
            };
          }
          
          stockValidation.push({
            rowIndex: i,
            codeNum: item.codeNum,
            name: item.name,
            currentStock: currentStock,
            orderCnt: item.cnt,
            newStock: currentStock - item.cnt
          });
          found = true;
          break;
        }
      }
      
      if (!found) {
        return {
          success: false,
          message: `Item not found: ${item.name} (Code: ${item.codeNum})`
        };
      }
    }
    
    // Step 2: Save order and update stock
    const orderDate = orderData.date;
    const orderIndex = orderData.index.toString().padStart(4, '0');
    const orderSerialNumber = orderDate.toString() + orderIndex;
    const payType = orderData.payType || '카드';  // ✅ NEW: Default to card if not specified

    // 트랜잭션처럼 처리 (모두 성공하거나 모두 실패)
    try {
      // Add order items
      orderData.items.forEach(item => {
        const cost = item.isB2B ? (item.costB2B || 0) : (item.costB2C || 0);
        const totalCost = cost * item.cnt;
        // Get current time for Order_Time
        const currentTime = new Date();

        purchaseSheet.appendRow([
          orderSerialNumber,
          parseInt(orderDate),
          currentTime,
          orderIndex,
          item.codeNum,
          item.name,
          item.description,
          item.costB2B || 0,
          item.costB2C || 0,
          item.isB2B ? 1 : 0,
          item.cnt,
          payType,           // ✅ NEW
          totalCost,
          ''                 // ✅ NEW: IsCanceled (empty = not canceled)
        ]);
      });
      
      // Update stock
      stockValidation.forEach(stock => {
        itemSheet.getRange(stock.rowIndex + 1, stockNumColIndex + 1).setValue(stock.newStock);
      });
      
      // ✅ 3단계: 로그 기록 (선택사항)
      Logger.log(`Order completed - Number: ${orderSerialNumber}, Items: ${orderData.items.length}`);
      
      return {
        success: true,
        message: 'Order saved successfully.',
        orderSerialNumber: orderSerialNumber,
        stockUpdates: stockValidation
      };
      
    } catch (saveError) {
      // 롤백은 어려우므로 에러 로그만 남김
      Logger.log('Error during order save (partial save may have occurred): ' + saveError.toString());
      throw saveError;
    }
    
  } catch (error) {
    Logger.log('Error saving order: ' + error.toString());
    return {
      success: false,
      message: 'An error occurred while saving order: ' + error.toString()
    };
  }
}

// ===== NEW FUNCTION: Cancel Order =====
function cancelOrder(orderSerialNumber) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const purchaseSheet = ss.getSheetByName(PURCHASE_ORDER_SHEET_NAME);
    const itemSheet = ss.getSheetByName(SHEET_NAME);
    
    if (!purchaseSheet || !itemSheet) {
      return { success: false, message: 'Required sheets not found.' };
    }
    
    const data = purchaseSheet.getDataRange().getValues();
    const headers = data[0];
    
    const serialColIndex = headers.indexOf('Order_SerialNumber');
    const canceledColIndex = headers.indexOf('IsCanceled');
    const codeNumColIndex = headers.indexOf('Order_CodeNum');
    const cntColIndex = headers.indexOf('Order_Cnt');
    
    if (serialColIndex === -1 || canceledColIndex === -1) {
      return { success: false, message: 'Required columns not found.' };
    }
    
    // Find all rows with matching serial number
    const rowsToCancel = [];
    const itemsToRestore = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][serialColIndex] && 
          data[i][serialColIndex].toString() === orderSerialNumber.toString()) {
        
        // Check if already canceled
        if (data[i][canceledColIndex] === '취소') {
          return { success: false, message: 'Order is already canceled.' };
        }
        
        rowsToCancel.push(i + 1); // +1 for 1-indexed sheet rows
        itemsToRestore.push({
          codeNum: data[i][codeNumColIndex],
          cnt: data[i][cntColIndex]
        });
      }
    }
    
    if (rowsToCancel.length === 0) {
      return { success: false, message: 'Order not found: ' + orderSerialNumber };
    }
    
    // Mark rows as canceled and apply red color
    rowsToCancel.forEach(rowNum => {
      purchaseSheet.getRange(rowNum, canceledColIndex + 1).setValue('취소');
      
      // ✅ Apply red color to entire row
      const lastCol = purchaseSheet.getLastColumn();
      purchaseSheet.getRange(rowNum, 1, 1, lastCol).setFontColor('#dc2626');
    });
    
    // Restore stock
    const itemData = itemSheet.getDataRange().getValues();
    const itemHeaders = itemData[0];
    const itemCodeColIndex = itemHeaders.indexOf('CodeNum');
    const stockColIndex = itemHeaders.indexOf('StockNum');
    
    itemsToRestore.forEach(item => {
      for (let i = 1; i < itemData.length; i++) {
        if (itemData[i][itemCodeColIndex] && 
            itemData[i][itemCodeColIndex].toString() === item.codeNum.toString()) {
          const currentStock = itemData[i][stockColIndex] || 0;
          const newStock = currentStock + item.cnt;
          itemSheet.getRange(i + 1, stockColIndex + 1).setValue(newStock);
          break;
        }
      }
    });
    
    Logger.log(`Order canceled - Serial: ${orderSerialNumber}, Rows: ${rowsToCancel.length}`);
    
    return {
      success: true,
      message: 'Order canceled successfully.',
      canceledRows: rowsToCancel.length,
      restoredItems: itemsToRestore
    };
    
  } catch (error) {
    Logger.log('Error canceling order: ' + error.toString());
    return {
      success: false,
      message: 'An error occurred while canceling order: ' + error.toString()
    };
  }
}

// 특정 날짜의 모든 주문서 목록 가져오기
function getOrderListByDate(orderDate) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(PURCHASE_ORDER_SHEET_NAME);
    
    if (!sheet) {
      return {
        success: false,
        message: 'PurchaseOrder 시트를 찾을 수 없습니다.'
      };
    }
    
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return {
        success: false,
        message: '주문 데이터가 없습니다.'
      };
    }
    
    const headers = data[0];
    const dateColIndex = headers.indexOf('Order_Date');
    const indexColIndex = headers.indexOf('Order_Index');
    
    if (dateColIndex === -1 || indexColIndex === -1) {
      return {
        success: false,
        message: '필요한 열을 찾을 수 없습니다.'
      };
    }
    
    // 해당 날짜의 주문서 인덱스 목록 추출
    const orderIndexSet = new Set();
    
    for (let i = 1; i < data.length; i++) {
      const rowDate = data[i][dateColIndex];
      const rowIndex = data[i][indexColIndex];
      
      if (rowDate && rowDate.toString() === orderDate.toString()) {
        orderIndexSet.add(parseInt(rowIndex));
      }
    }
    
    if (orderIndexSet.size === 0) {
      return {
        success: false,
        message: '해당 날짜의 주문서가 없습니다.'
      };
    }
    
    // Set을 배열로 변환하고 정렬
    const orderList = Array.from(orderIndexSet).sort((a, b) => a - b);
    
    return {
      success: true,
      orderList: orderList,
      count: orderList.length
    };
    
  } catch (error) {
    Logger.log('주문서 목록 조회 오류: ' + error.toString());
    return {
      success: false,
      message: '오류가 발생했습니다: ' + error.toString()
    };
  }
}