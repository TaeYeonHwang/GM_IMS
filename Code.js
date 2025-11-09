// 구글시트 설정 값 가져오기.
// Code.gs

const RECEIPT_TEMPLATE = 'ReceiptTemplate_2'; // 'ReceiptTemplate' 또는 'ReceiptTemplate_2'

function getConfig() {
  const scriptProperties = PropertiesService.getScriptProperties();
  return {
    SPREADSHEET_ID: scriptProperties.getProperty('SPREADSHEET_ID'),
    GUIDE_IMAGE_ID: scriptProperties.getProperty('GUIDE_IMAGE_ID'),
    FOLDER_ID: scriptProperties.getProperty('FOLDER_ID')
  };
}

const config = getConfig();
const SPREADSHEET_ID = config.SPREADSHEET_ID;
const SHEET_NAME = 'ItemInfo';
const LOG_SHEET_NAME = 'AccessLog';
const PURCHASE_ORDER_SHEET_NAME = 'PurchaseOrder';

// ===== Cache Service Utility Functions =====
const CACHE_KEYS = {
  DASHBOARD_INFO: 'dashboard_info',
  INVENTORY_STATUS: 'inventory_status',
  REVISION_INFO: 'revision_info',
  LATEST_ORDER: 'latest_order_info'
};
const CACHE_DURATION = {
  DASHBOARD: 600,      // 10 minutes
  INVENTORY: 600,      // 10 minutes
  REVISION: 3600,      // 1 hour
  LATEST_ORDER: 3600   // 1 hour
};

// Get latest order created today
function getLatestTodayOrder() {
  try {
    const today = new Date();
    const year = today.getFullYear();
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const day = String(today.getDate()).padStart(2, '0');
    const todayDate = parseInt(year + month + day);
    Logger.log('Today date: ' + todayDate);

    // ✅ Step 1: Check cache first
    const cached = getCachedData(CACHE_KEYS.LATEST_ORDER);
    if (cached) {
      Logger.log('✓ Latest order from cache');
      return cached;
    }
    
    // ✅ Step 2: Cache miss - fetch from DB
    Logger.log('✗ Cache miss - fetching latest order from DB');

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const orderSheet = ss.getSheetByName(PURCHASE_ORDER_SHEET_NAME);
    
    if (!orderSheet) {
      return {
        success: false,
        message: 'PurchaseOrder sheet not found.'
      };
    }
    
    const data = orderSheet.getDataRange().getValues();
    
    if (data.length < 1) {
      return {
        success: true,
        hasOrder: false,
        message: '오늘 생성된 주문서가 없습니다.'
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
    
    // Find max order index for today
    let maxIndex = -1;
    for (let i = 1; i < data.length; i++) {
      const rowDate = data[i][colIndices['Order_Date']];
      const rowIndex = parseInt(data[i][colIndices['Order_Index']]) || 0;
      
      if (rowDate && parseInt(rowDate.toString()) === todayDate) {
        if (rowIndex > maxIndex) {
          maxIndex = rowIndex;
        }
      }
    }
    
    // No orders today
    if (maxIndex === -1) {
      const result = {
        success: true,
        hasOrder: false,
        message: '오늘 생성된 주문서가 없습니다.'
      };
      
      // Cache for 1 hour
      setCachedData(CACHE_KEYS.LATEST_ORDER, result, CACHE_DURATION.LATEST_ORDER);
      return result;
    }
    
    // Collect all items for the latest order
    const orderIndexStr = maxIndex.toString().padStart(4, '0');
    const targetSerialNumber = todayDate.toString() + orderIndexStr;
    const orders = [];
    
    for (let i = 1; i < data.length; i++) {
      const rowSerialNumber = data[i][colIndices['Order_SerialNumber']];
      
      if (rowSerialNumber && rowSerialNumber.toString() === targetSerialNumber) {
        let orderTime = data[i][colIndices['Order_Time']];
        if (orderTime instanceof Date) {
          orderTime = Utilities.formatDate(orderTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
        } else if (orderTime) {
          orderTime = orderTime.toString();
        } else {
          orderTime = '-';
        }
        
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
          payType: data[i][colIndices['PayType']] || '-',
          totalCost: data[i][colIndices['Order_TotalCost']],
          isCanceled: isCanceled
        });
      }
    }
    
    if (orders.length === 0) {
      const result = {
        success: true,
        hasOrder: false,
        message: '오늘 생성된 주문서가 없습니다.'
      };
      
      setCachedData(CACHE_KEYS.LATEST_ORDER, result, CACHE_DURATION.LATEST_ORDER);
      return result;
    }
    
    // Calculate summary
    let totalAmount = 0;
    let totalQty = 0;
    orders.forEach(order => {
      totalAmount += order.totalCost || 0;
      totalQty += order.cnt || 0;
    });
    
    const result = {
      success: true,
      hasOrder: true,
      orderSerialNumber: targetSerialNumber,
      orderIndex: maxIndex,
      orderTime: orders[0].time,
      payType: orders[0].payType,
      isCanceled: orders[0].isCanceled,
      itemCount: orders.length,
      totalQty: totalQty,
      totalAmount: totalAmount,
      items: orders.slice(0, 3)  // First 3 items for preview
    };
    
    // ✅ Step 3: Cache for 1 hour
    setCachedData(CACHE_KEYS.LATEST_ORDER, result, CACHE_DURATION.LATEST_ORDER);
    Logger.log('✓ Latest order cached for 1 hour');
    
    return result;
    
  } catch (error) {
    Logger.log('Error getting latest order: ' + error.toString());
    return {
      success: false,
      message: 'An error occurred: ' + error.toString()
    };
  }
}

/**
Get data from cache, return null if not found or error
*/
function getCachedData(key) {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(key);
    if (cached) {
      Logger.log('Cache HIT: ' + key);
      return JSON.parse(cached);
    }
    Logger.log('Cache MISS: ' + key);
    return null;
  } catch (error) {
    Logger.log('Cache read error for key ' + key + ': ' + error.toString());
    return null;
  }
}

/**
Save data to cache with TTL
*/
function setCachedData(key, data, ttlSeconds) {
  try {
    const cache = CacheService.getScriptCache();
    cache.put(key, JSON.stringify(data), ttlSeconds);
    Logger.log('Cache SET: ' + key + ' (TTL: ' + ttlSeconds + 's)');
    return true;
  } catch (error) {
    Logger.log('Cache write error for key ' + key + ': ' + error.toString());
    return false;
  }
}

/**
Invalidate specific cache key
*/
function invalidateCache(key) {
  try {
    const cache = CacheService.getUserCache();
    const fullKey = CACHE_KEYS[cacheKey];
    
    if (fullKey) {
      cache.remove(fullKey);
      Logger.log(`✓ Cache invalidated: ${fullKey}`);
      return { success: true };
    }
    
    return { success: false, message: 'Invalid cache key' };
  } catch (error) {
    Logger.log('Error invalidating cache: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
Invalidate multiple cache keys
*/
function invalidateCaches(keys) {
  keys.forEach(key => invalidateCache(key));
}
// ===== END: Cache Service Utility Functions =====

// HTML file include function.
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Get GUIDE_IMAGE_ID function.
function getGuideImageId() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const imageId = scriptProperties.getProperty('GUIDE_IMAGE_ID');
  
  if (!imageId) {
    Logger.log('GUIDE_IMAGE_ID is not set.');
    return null;
  }
  
  return imageId;
}

// Web app entry point
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('GM - Inventory Management System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Log access
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
    Logger.log('Log access error: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// Search item by code number
function searchByCodeNum(codeNum) {
  try {
    // Input validation
    if (!codeNum || codeNum.toString().trim() === '') {
      return {
        success: false,
        message: 'Please enter a code number.'
      };
    }
    
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      return {
        success: false,
        message: 'ItemInfo sheet not found.'
      };
    }
    
    // ✅ 마지막 행만 확인하여 필요한 범위만 읽기
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      return {
        success: false,
        message: 'No items registered.'
      };
    }
    
    // 헤더(1행) 제외, 2행부터 lastRow까지, 1열부터 9열(IsShortage)까지만 읽기
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 9);
    const data = dataRange.getValues();
    
    const searchCode = codeNum.toString().trim();
    
    // Efficient search using Array.find
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
          stockNum: row[6],
          shortageNum: row[7],
          isShortage: row[8]
        },
        rowNumber: foundIndex + 2  // 실제 시트의 행 번호
      };
    }
    
    return {
      success: false,
      message: 'Code number not found: ' + searchCode
    };
    
  } catch (error) {
    Logger.log('searchByCodeNum error: ' + error.toString());
    return {
      success: false,
      message: 'An error occurred during search: ' + error.message
    };
  }
}

// Get all items
function getAllItems() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    // ✅ 개선점 1: 마지막 행 확인
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      // No Data case. (only exist header)
      return {
        success: true,
        items: []
      };
    }

    // ✅ 개선점 2: A열~I열(9개 컬럼)만 읽기
    // getRange(startRow, startCol, numRows, numCols)
    const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

    Logger.log('Total rows: ' + data.length);

    const items = [];
    for (let i = 0; i < data.length; i++) {  // ✅ 인덱스 0부터 시작 (헤더 제외했으므로)
      if (data[i][0]) { // SerialNum이 있는 행만 처리
        const item = {
          serialNum: data[i][0],  // A열
          name: data[i][1],       // B열
          description: data[i][2],// C열
          codeNum: data[i][3],    // D열
          costB2B: data[i][4],    // E열
          costB2C: data[i][5],    // F열
          stockNum: data[i][6],   // G열
          shortageNum: data[i][7],// H열
          isShortage: data[i][8]  // I열
        };
        items.push(item);
      }
    }
    
    return {
      success: true,
      items: items
    };
  } catch (error) {
    Logger.log('getAllItems error: ' + error.toString());
    return {
      success: false,
      message: 'An error occurred:  ' + error.toString()
    };
  }
}

// Get latest revision
function getLatestRevision() {
  try {

    // ✅ Step 1: Check Cache (1 hour valid)
    const cached = getCachedData(CACHE_KEYS.REVISION_INFO);
    if (cached) {
      Logger.log('✓ Revision info from cache');
      return cached;
    }

    // ✅ Step 2: Cache Miss - Search in Sheet
    Logger.log('✗ Cache miss - fetching Revision from DB');

    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('RevisionHistory');
    
    if (!sheet) {
      return {
        success: false,
        message: 'RevisionHistory sheet not found.'
      };
    }
    
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return {
        success: false,
        message: 'No revision data found.'
      };
    }
    
    const headers = data[0];
    const revisionCol = headers.indexOf('Revision');
    const dateCol = headers.indexOf('Date');
    
    if (revisionCol === -1 || dateCol === -1) {
      return {
        success: false,
        message: 'Revision or Date column not found.'
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
        message: 'No valid revision value found.'
      };
    }
    
    if (maxRevisionDate instanceof Date) {
      maxRevisionDate = Utilities.formatDate(maxRevisionDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    
     const result = {
      success: true,
      revision: maxRevision,
      date: maxRevisionDate
    };
    
    // ✅ Step 3: Save Result to Cache (1 Hour)
    setCachedData(CACHE_KEYS.REVISION_INFO, result, CACHE_DURATION.REVISION);
    Logger.log('✓ Revision info cached for 1 hour');
    
    return result;

  } catch (error) {
    Logger.log('Revision query error: ' + error.toString());
    return {
      success: false,
      message: 'An error occurred: ' + error.toString()
    };
  }
}

// Get revision history
function getRevisionHistory() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('RevisionHistory');
    
    if (!sheet) {
      return {
        success: false,
        message: 'RevisionHistory sheet not found.'
      };
    }
    
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return {
        success: false,
        message: 'No revision data found.'
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
        message: 'Cannot find columns: ' + missingColumns.join(', ')
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
    Logger.log('Revision History query error: ' + error.toString());
    return {
      success: false,
      message: 'An error occurred:  ' + error.toString()
    };
  }
}

// Get new order index
function getNewOrderIndex(orderDate) {
  try {
    //  Input validation
    if (!orderDate) {
      return {
        success: false,
        message: 'Order date not provided.'
      };
    }
    
    // Date format validation (YYYYMMDD, 8 digits)
    const dateStr = orderDate.toString();
    if (dateStr.length !== 8 || isNaN(dateStr)) {
      return {
        success: false,
        message: 'Invalid date format. Must be YYYYMMDD format. (Input: ' + dateStr + ')'
      };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const dashboardSheet = ss.getSheetByName('Dashboard');
    
    if (!ss) {
      return {
        success: false,
        message: 'Cannot access spreadsheet. Please check ID.'
      };
    }
    
    if (!dashboardSheet) {
      return {
        success: false,
        message: 'Dashboard sheet not found.'
      };
    }
    
    // ✅ Dashboard B5 값 읽기 (오늘 주문서 개수)
    const existingOrderCount = dashboardSheet.getRange('B5').getValue() || 0;
    const newIndex = existingOrderCount + 1;
    
    // Validate index range (max 9999)
    if (newIndex > 9999) {
      return {
        success: false,
        message: 'Exceeded maximum orders per day (9999).'
      };
    }

    Logger.log(`날짜 ${orderDate}: 기존 ${existingOrderCount}개, 새 인덱스: ${newIndex}`);
    
    return {
      success: true,
      orderIndex: newIndex,
      existingOrders: existingOrderCount
    };
    
  } catch (error) {
    Logger.log('Error getting new order index: ' + error.toString());
    return {
      success: false,
      message: 'An error occurred: ' + error.message
    };
  }
}

// Get order data
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
          payType: data[i][colIndices['PayType']] || '-',
          totalCost: data[i][colIndices['Order_TotalCost']],
          isCanceled: isCanceled
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

// Save order
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
          payType,
          totalCost,
          ''                 // ✅ NEW: IsCanceled (empty = not canceled)
        ]);
      });
      
      // Update stock
      stockValidation.forEach(stock => {
        itemSheet.getRange(stock.rowIndex + 1, stockNumColIndex + 1).setValue(stock.newStock);
      });
      
      Logger.log(`Order completed - Number: ${orderSerialNumber}, Items: ${orderData.items.length}`);
      
      // ✅ NEW: Invalidate caches after order creation
      invalidateCaches([
        CACHE_KEYS.DASHBOARD_INFO,
        CACHE_KEYS.INVENTORY_STATUS,
        CACHE_KEYS.LATEST_ORDER
      ]);

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

// Cancel order
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
    
    // ✅ NEW: Invalidate caches after order cancellation
    invalidateCaches([
      CACHE_KEYS.DASHBOARD_INFO,
      CACHE_KEYS.INVENTORY_STATUS,
      CACHE_KEYS.LATEST_ORDER
    ]);

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

//  Get order list by date
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

// ===== Memo Functions =====
const MEMO_SHEET_NAME = 'Memo';

// Get latest 10 memos
function getLatestMemos() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(MEMO_SHEET_NAME);
    
    if (!sheet) {
      // Create sheet if not exists
      sheet = ss.insertSheet(MEMO_SHEET_NAME);
      sheet.appendRow(['Date', 'Index', 'Content']);
      Logger.log('Memo sheet created.');
      
      return {
        success: true,
        memos: []
      };
    }
    
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return {
        success: true,
        memos: []
      };
    }
    
    const headers = data[0];
    const dateColIndex = headers.indexOf('Date');
    const indexColIndex = headers.indexOf('Index');
    const contentColIndex = headers.indexOf('Content');
    
    if (dateColIndex === -1 || indexColIndex === -1 || contentColIndex === -1) {
      return {
        success: false,
        message: 'Required columns not found in Memo sheet.'
      };
    }
    
    // Collect all memos with row numbers
    const memos = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][dateColIndex] && data[i][contentColIndex]) {
        memos.push({
          rowNumber: i + 1, // Actual sheet row number
          date: data[i][dateColIndex],
          index: data[i][indexColIndex],
          content: data[i][contentColIndex]
        });
      }
    }
    
    // Sort by date descending (latest first), then by index descending
    memos.sort((a, b) => {
      const dateA = parseInt(a.date.toString());
      const dateB = parseInt(b.date.toString());
      
      if (dateB !== dateA) {
        return dateB - dateA;
      }
      return b.index - a.index;
    });
    
    // Return top 10
    const latest10 = memos.slice(0, 10);
    
    return {
      success: true,
      memos: latest10
    };
    
  } catch (error) {
    Logger.log('Error getting memos: ' + error.toString());
    return {
      success: false,
      message: 'An error occurred: ' + error.toString()
    };
  }
}

// Add new memo
function addMemo(content) {
  try {
    if (!content || content.trim() === '') {
      return {
        success: false,
        message: 'Memo content cannot be empty.'
      };
    }
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(MEMO_SHEET_NAME);
    
    if (!sheet) {
      sheet = ss.insertSheet(MEMO_SHEET_NAME);
      sheet.appendRow(['Date', 'Index', 'Content']);
      Logger.log('Memo sheet created.');
    }
    
    // Get today's date in YYYYMMDD format
    const today = new Date();
    const year = today.getFullYear();
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const day = String(today.getDate()).padStart(2, '0');
    const todayDate = parseInt(year + month + day);
    
    // Find the next index for today
    const data = sheet.getDataRange().getValues();
    let maxIndex = 0;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && parseInt(data[i][0].toString()) === todayDate) {
        const currentIndex = parseInt(data[i][1]) || 0;
        if (currentIndex > maxIndex) {
          maxIndex = currentIndex;
        }
      }
    }
    
    const newIndex = maxIndex + 1;
    
    // Add new memo
    sheet.appendRow([todayDate, newIndex, content.trim()]);
    
    Logger.log(`Memo added - Date: ${todayDate}, Index: ${newIndex}`);
    
    return {
      success: true,
      message: 'Memo saved successfully.',
      date: todayDate,
      index: newIndex
    };
    
  } catch (error) {
    Logger.log('Error adding memo: ' + error.toString());
    return {
      success: false,
      message: 'An error occurred while saving memo: ' + error.toString()
    };
  }
}

// Update memo
function updateMemo(rowNumber, newContent) {
  try {
    if (!newContent || newContent.trim() === '') {
      return {
        success: false,
        message: 'Memo content cannot be empty.'
      };
    }
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(MEMO_SHEET_NAME);
    
    if (!sheet) {
      return {
        success: false,
        message: 'Memo sheet not found.'
      };
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const contentColIndex = headers.indexOf('Content');
    
    if (contentColIndex === -1) {
      return {
        success: false,
        message: 'Content column not found.'
      };
    }
    
    // Update the content
    sheet.getRange(rowNumber, contentColIndex + 1).setValue(newContent.trim());
    
    Logger.log(`Memo updated - Row: ${rowNumber}`);
    
    return {
      success: true,
      message: 'Memo updated successfully.'
    };
    
  } catch (error) {
    Logger.log('Error updating memo: ' + error.toString());
    return {
      success: false,
      message: 'An error occurred while updating memo: ' + error.toString()
    };
  }
}

// Delete memo
function deleteMemo(rowNumber) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(MEMO_SHEET_NAME);
    
    if (!sheet) {
      return {
        success: false,
        message: 'Memo sheet not found.'
      };
    }
    
    // Delete the row
    sheet.deleteRow(rowNumber);
    
    Logger.log(`Memo deleted - Row: ${rowNumber}`);
    
    return {
      success: true,
      message: 'Memo deleted successfully.'
    };
    
  } catch (error) {
    Logger.log('Error deleting memo: ' + error.toString());
    return {
      success: false,
      message: 'An error occurred while deleting memo: ' + error.toString()
    };
  }
}

// ===== Dashboard Info Functions =====

// Get today's dashboard info (order count and stock status)
function getDashboardInfo() {
  try {
    // ✅ NEW: Try cache first
    const cached = getCachedData(CACHE_KEYS.DASHBOARD_INFO);
    if (cached) {
      return cached;
    }

    // ✅ Cache miss - fetch from DB
    const today = new Date();
    const year = today.getFullYear();
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const day = String(today.getDate()).padStart(2, '0');
    const todayDate = parseInt(year + month + day);
    
    // Get today's order count and stock status
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const dashboardSheet = ss.getSheetByName('Dashboard');

    if (!dashboardSheet) {
      return {
        success: false,
        message: 'Dashboard sheet not found. Please create it with formulas.'
      };
    }

    // ✅ 개선점 1: Dashboard B5에서 오늘 주문 수를 바로 읽음 (PurchaseOrder 전체 스캔 X)
    const orderCount = dashboardSheet.getRange('B5').getValue() || 0;
    
    // ✅ 개선점 2: Dashboard B2~B4에서 재고 상태를 바로 읽음 (ItemInfo 전체 스캔 X)
    const outCount = dashboardSheet.getRange('B2').getValue() || 0;
    const lowCount = dashboardSheet.getRange('B3').getValue() || 0;
    
    let stockStatus = 2; // Default: 양호
    if (outCount > 0) {
      stockStatus = 0; // 경고 (품절 있음)
    } else if (lowCount > 0) {
      stockStatus = 1; // 관심 (부족 있음)
    }
    
    const result = {
      success: true,
      todayDate: `${year}년 ${month}월 ${day}일`,
      orderCount: orderCount,
      stockStatus: stockStatus
    };
    
    // ✅ NEW: Save to cache (10 minutes)
    setCachedData(CACHE_KEYS.DASHBOARD_INFO, result, CACHE_DURATION.DASHBOARD);
    
    return result;
    
  } catch (error) {
    Logger.log('Error getting dashboard info: ' + error.toString());
    return {
      success: false,
      message: 'An error occurred: ' + error.toString()
    };
  }
}

function getOrdersByDateRange(startDate, endDate) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(PURCHASE_ORDER_SHEET_NAME);
    if (!sheet) {
      return {
        success: false,
        message: 'PurchaseOrder sheet not found.'
      };
    }

    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return {
        success: true,
        orders: []
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

    const startDateInt = parseInt(startDate);
    const endDateInt = parseInt(endDate);

    // Collect all orders in date range
    const ordersMap = {}; // Group by serial number

    for (let i = 1; i < data.length; i++) {
      const rowDate = parseInt(data[i][colIndices['Order_Date']]);
      
      if (rowDate >= startDateInt && rowDate <= endDateInt) {
        const serialNumber = data[i][colIndices['Order_SerialNumber']].toString();
        
        // Format order time
        let orderTime = data[i][colIndices['Order_Time']];
        if (orderTime instanceof Date) {
          orderTime = Utilities.formatDate(orderTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
        } else if (orderTime) {
          orderTime = orderTime.toString();
        } else {
          orderTime = '-';
        }
        
        // Check if canceled
        const canceledValue = data[i][colIndices['IsCanceled']];
        const isCanceled = canceledValue === '취소';
        
        const order = {
          serialNumber: serialNumber,
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
          payType: data[i][colIndices['PayType']] || '-',
          totalCost: data[i][colIndices['Order_TotalCost']],
          isCanceled: isCanceled
        };
        
        // Group by serial number
        if (!ordersMap[serialNumber]) {
          ordersMap[serialNumber] = [];
        }
        ordersMap[serialNumber].push(order);
      }
    }

    // Convert map to array of order groups
    const orderGroups = [];
    for (const serialNumber in ordersMap) {
      const orders = ordersMap[serialNumber];
      if (orders.length > 0) {
        orderGroups.push({
          date: orders[0].date.toString(),
          index: orders[0].index,
          orders: orders
        });
      }
    }

    // Sort by date and index
    orderGroups.sort((a, b) => {
      if (a.date !== b.date) {
        return parseInt(a.date) - parseInt(b.date);
      }
      return a.index - b.index;
    });

    Logger.log(`Found ${orderGroups.length} orders in date range ${startDate}-${endDate}`);

    return {
      success: true,
      orderGroups: orderGroups
    };
  } catch (error) {
    Logger.log('Error getting orders by date range: ' + error.toString());
    return {
      success: false,
      message: 'An error occurred: ' + error.toString()
    };
  }
}


function getInventoryStatusCountsFromSheet() {
  try {
    // ✅ NEW: Try cache first
    const cached = getCachedData(CACHE_KEYS.INVENTORY_STATUS);
    if (cached) {
      return cached;
    }
    
    // ✅ Cache miss - fetch from sheet
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let dashboardSheet = ss.getSheetByName('Dashboard');

    if (!dashboardSheet) {
      return {
        success: false,
        message: 'Dashboard sheet not found.'
      };
    }

    // 값 읽기 (B2(품절 개수), B3(품절임박 개수), B4(정상 개수))
    const outCount = dashboardSheet.getRange('B2').getValue() || 0;
    const lowCount = dashboardSheet.getRange('B3').getValue() || 0;
    const normalCount = dashboardSheet.getRange('B4').getValue() || 0;

    Logger.log(`Inventory counts from sheet - Out: ${outCount}, Low: ${lowCount}, Normal: ${normalCount}`);

    const result = {
      success: true,
      outCount: outCount,
      lowCount: lowCount,
      normalCount: normalCount
    };
    
    // ✅ NEW: Save to cache (10 minutes)
    setCachedData(CACHE_KEYS.INVENTORY_STATUS, result, CACHE_DURATION.INVENTORY);
    
    return result;

  } catch (error) {
    Logger.log('Error getting inventory counts from sheet: ' + error.toString());
    return {
      success: false,
      message: 'An error occurred: ' + error.toString()
    };
  }
}

// Generate receipt PDF(s) from order data
function generateReceiptPDF(orderSerialNumber) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const templateSheet = ss.getSheetByName(RECEIPT_TEMPLATE);
    
    if (!templateSheet) {
      return { success: false, message: `${RECEIPT_TEMPLATE} sheet not found.`};
    }
    
    // Get order data
    const orderSheet = ss.getSheetByName(PURCHASE_ORDER_SHEET_NAME);
    if (!orderSheet) {
      return { success: false, message: 'PurchaseOrder sheet not found.' };
    }
    
    const data = orderSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find order items
    const orders = [];
    for (let i = 1; i < data.length; i++) {
      const rowSerialNumber = data[i][headers.indexOf('Order_SerialNumber')];
      if (rowSerialNumber && rowSerialNumber.toString() === orderSerialNumber.toString()) {
        orders.push({
          date: data[i][headers.indexOf('Order_Date')],
          name: data[i][headers.indexOf('Order_Name')],
          description: data[i][headers.indexOf('Order_Description')],
          cnt: data[i][headers.indexOf('Order_Cnt')],
          costB2B: data[i][headers.indexOf('Order_CostB2B')],
          costB2C: data[i][headers.indexOf('Order_CostB2C')],
          isB2B: data[i][headers.indexOf('Order_IsB2B')],
          totalCost: data[i][headers.indexOf('Order_TotalCost')]
        });
      }
    }
    
    if (orders.length === 0) {
      return { success: false, message: 'No orders found for this serial number.' };
    }
    
    // Choose template-specific settings
    let itemsPerPage, startRow, endRow, exportRange, portrait;
    
    if (RECEIPT_TEMPLATE === 'ReceiptTemplate_2') {
      itemsPerPage = 11;
      startRow = 9;
      endRow = 19;
      exportRange = 'A1:AJ44';
      portrait = true;
    } else {
      itemsPerPage = 14;
      startRow = 14;
      endRow = 27;
      exportRange = 'A1:Y30';
      portrait = false;
    }

    // Split into pages
    const pages = [];
    for (let i = 0; i < orders.length; i += itemsPerPage) {
      pages.push(orders.slice(i, i + itemsPerPage));
    }
    
    const pdfFiles = [];
    const folderId = config.FOLDER_ID;
    if (!folderId) {
      return { success: false, message: 'FOLDER_ID is not set in script properties.' };
    }
    const folder = DriveApp.getFolderById(folderId);
    
    // Format date
    const dateStr = orders[0].date.toString();
    const year = dateStr.substring(0, 4);
    const month = dateStr.substring(4, 6);
    const day = dateStr.substring(6, 8);
    const formattedDateFull = `${year}.${month}.${day}`;
    const formattedDateShort = `${month}.${day}`;
    
    // Generate PDF for each page
    for (let pageNum = 0; pageNum < pages.length; pageNum++) {
      const pageItems = pages[pageNum];
      
      // Create temp sheet
      const tempSheet = templateSheet.copyTo(ss);
      const tempSheetId = tempSheet.getSheetId();
      tempSheet.setName('Temp_' + Date.now() + '_' + pageNum);
      
      if (RECEIPT_TEMPLATE === 'ReceiptTemplate_2') {
        // ===== ReceiptTemplate_2 Logic =====
        
        // Q2: Full date (YYYY.MM.DD)
        tempSheet.getRange('Q2').setValue(formattedDateFull);
        
        // AH2: Page number (1/2, 2/2, etc.)
        tempSheet.getRange('AH2').setValue(`${pageNum + 1}/${pages.length}`);
        
        // Calculate total for this page
        let pageTotal = 0;
        pageItems.forEach(item => {
          pageTotal += item.totalCost;
        });
        
        // Fill items (rows 9-19)
        let row = startRow;
        pageItems.forEach(item => {
          // B: Date (MM.DD)
          tempSheet.getRange(row, 2).setValue(formattedDateShort);
          
          // D: Item name + description
          const itemText = item.description 
            ? `${item.name} (${item.description})` 
            : item.name;
          tempSheet.getRange(row, 4).setValue(itemText);
          
          // Q: Quantity
          tempSheet.getRange(row, 17).setValue(item.cnt);
          
          // T: Unit price (excluding VAT - 90% of original price)
          const originalPrice = item.isB2B === 1 ? item.costB2B : item.costB2C;
          const unitPriceExVAT = Math.round(originalPrice * 0.9);
          tempSheet.getRange(row, 20).setValue(unitPriceExVAT);
          
          // AB: VAT (10% of original price)
          const vat = Math.round(originalPrice * 0.1);
          tempSheet.getRange(row, 28).setValue(vat);
          
          // X: Total (original price * quantity)
          const totalPrice = originalPrice * item.cnt;
          tempSheet.getRange(row, 24).setValue(totalPrice);
          
          row++;
        });
        
        // D20: Total amount (sum of all items)
        tempSheet.getRange('D20').setValue(pageTotal);
        
      } else {
        // ===== ReceiptTemplate Logic (기존 코드) =====
        
        // B11: Full date
        tempSheet.getRange('B11').setValue(formattedDateFull);
        
        // Calculate total for this page
        let pageTotal = 0;
        pageItems.forEach(item => {
          pageTotal += item.totalCost;
        });
        tempSheet.getRange('F11').setValue(pageTotal);
        
        // Fill items (rows 14-27)
        let row = startRow;
        pageItems.forEach(item => {
          // Date (B column)
          tempSheet.getRange(row, 2).setValue(formattedDateShort);
          
          // Item name + description (D column)
          const itemText = item.description 
            ? `${item.name} (${item.description})` 
            : item.name;
          tempSheet.getRange(row, 4).setValue(itemText);
          
          // Quantity (G column)
          tempSheet.getRange(row, 7).setValue(item.cnt);
          
          // Unit price (I column)
          const unitPrice = item.isB2B === 1 ? item.costB2B : item.costB2C;
          tempSheet.getRange(row, 9).setValue(unitPrice);
          
          // Total price (K column)
          tempSheet.getRange(row, 11).setValue(item.totalCost);
          
          row++;
        });
        
        // Total at bottom (K28)
        tempSheet.getRange('K28').setValue(pageTotal);
      }
      
      SpreadsheetApp.flush();
      
      // Export to PDF
      const url = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/export?format=pdf&gid=${tempSheetId}&portrait=${portrait}&fitw=true&range=${exportRange}&gridlines=false&printtitle=false&sheetnames=false&pagenum=UNDEFINED&attachment=false&top_margin=0.2&bottom_margin=0.2&left_margin=0.2&right_margin=0.2`;
      
      const token = ScriptApp.getOAuthToken();
      const response = UrlFetchApp.fetch(url, {
        headers: { 'Authorization': 'Bearer ' + token }
      });
      
      // File name
      const filePrefix = RECEIPT_TEMPLATE === 'ReceiptTemplate_2' ? '거래명세표' : '영수증';
      const pageSuffix = pages.length > 1 ? `_${pageNum + 1}` : '';
      const fileName = `${filePrefix}_${orderSerialNumber}${pageSuffix}.pdf`;
      const pdfBlob = response.getBlob().setName(fileName);
      
      // Save to Drive
      const file = folder.createFile(pdfBlob);
      pdfFiles.push({
        name: fileName,
        url: file.getUrl(),
        id: file.getId()
      });
      
      // Delete temp sheet
      ss.deleteSheet(tempSheet);
      
      // Small delay between pages
      if (pageNum < pages.length - 1) {
        Utilities.sleep(500);
      }
    }
    
    Logger.log(`Generated ${pdfFiles.length} PDF(s) for order ${orderSerialNumber} using ${RECEIPT_TEMPLATE}`);
    
    return {
      success: true,
      files: pdfFiles,
      count: pdfFiles.length
    };
    
  } catch (error) {
    Logger.log('Error generating receipt: ' + error.toString());
    return {
      success: false,
      message: error.toString()
    };
  }
}

function getTodayOrderCount() {
    try {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      const dashboardSheet = ss.getSheetByName('Dashboard');
      
      if (!dashboardSheet) {
        return {
          success: false,
          message: 'Dashboard sheet not found.'
        };
      }
      
      const orderCount = dashboardSheet.getRange('B5').getValue() || 0;
      
      return {
        success: true,
        orderCount: orderCount
      };
      
    } catch (error) {
      Logger.log('Error getting today order count: ' + error.toString());
      return {
        success: false,
        message: 'An error occurred: ' + error.toString()
      };
    }
  }