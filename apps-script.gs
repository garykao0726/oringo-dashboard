// ==========================================
// 林果鞋業｜週報自動發送腳本 v3.4 FINAL
// 每週一 9:00 發送至 Google Chat
//
// 修正記錄：
//   - 月累計改為加總每日 P 欄，解決台中等結構不同門市的問題
//   - Shopline 金額改用 total.dollars（TWD 不用 ÷100）
//
// 欄位對應（row index 從 0 起算，資料列從 row6 = index 5 開始）：
//   A(row[0])  = 日期
//   C(row[2])  = 日目標（元）
//   I(row[8])  = 新客來客數
//   J(row[9])  = 舊客來客數
//   K(row[10]) = 維修鞋來客數
//   L(row[11]) = 新客成交數
//   M(row[12]) = 回購成交數
//   N(row[13]) = VIP成交數
//   O(row[14]) = 惜履券成交數
//   P(row[15]) = 總成交金額（元）
//
// 作者：Gary × Claude
// ==========================================

const CONFIG = {
  STORES: [
    { name: '中山店', id: '1xEKIVjJL5aWZsrpqrefyELCgXfYKoowMpbN256T87uo' },
    { name: '松菸店', id: '18UdW4--SeI7Wd1IiBentKW0keaCYE3NYdYYcCGgCyfg' },
    { name: '台中店', id: '1Jd6MIpv6FZCHtW0M59GJrvpTcYfsaLAVNwF9grOBviU' },
    { name: '東門店', id: '1au2yqNCUXNhzStjgMWevfySUFdZbYpnYarNVSrH7lwY' },
  ],
  WEBHOOK: 'https://chat.googleapis.com/v1/spaces/AAAAeOLlxg0/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=mHB_-6pbU6dIUNyoLsJEuAy7iIpVHKAC-CmGfovU_j0',
  DASHBOARD_URL: 'https://garykao0726.github.io/oringo-dashboard/',

  SHOPLINE_TOKEN:  'e7d6778e15c50da41da4184207f6a84140c80b00ef4124887dc65a3868577977',
  SHOPLINE_HANDLE: 'oringoshoes',

  GA4_PROPERTY_ID: '259936795',

  EC_TARGETS: {
    SESSIONS: 850000,
    CONV_RATE: 0.0035,
    AVG_TICKET: 5000,
  },

  DASHBOARD_SHEET_ID:    '17hTgCpF0mBTHf3RcnDmRFrJMeLBn3nRdTMuG8l58QjY',
  DASHBOARD_TAB_WEEK:    '週報數據',
  DASHBOARD_TAB_MONTHLY: '月度數據',

  ANNUAL_TARGETS: {
    '官網':   [125, 115, 115, 100, 100,  90,  90,  90, 100, 115, 115, 125],
    '中山店': [180, 170, 170, 150, 140, 125, 125, 125, 140, 170, 170, 180],
    '松菸店': [160, 150, 150, 140, 120, 110, 110, 110, 120, 140, 150, 160],
    '台中店': [130, 120, 120, 120, 110,  95,  95,  95, 110, 120, 120, 130],
    '東門店': [140, 130, 130, 120, 110, 100, 100, 100, 120, 130, 130, 140],
  },
  BENCHMARKS: {
    AVG_TICKET: 6800,
    CONV_RATE:  0.58,
  },
  CLAUDE_API_KEY: '',
};

// ==========================================
// 主函式：每週一自動執行
// ==========================================
function sendWeeklyReport() {
  var today = new Date();
  var currentMonth = today.getMonth() + 1;
  var ranges = getWeekRanges();
  var lastWeek = ranges.lastWeek;
  var thisWeek = ranges.thisWeek;

  Logger.log('▶ 上週：' + formatDate(lastWeek.start) + ' ~ ' + formatDate(lastWeek.end));
  Logger.log('▶ 本週：' + formatDate(thisWeek.start) + ' ~ ' + formatDate(thisWeek.end));

  var storesData = CONFIG.STORES.map(function(store) {
    return getStoreData(store, lastWeek, currentMonth);
  });

  var ecData = getShoplineData(lastWeek, currentMonth);
  storesData.push(ecData);

  var totalMTD = 0, totalMonthTarget = 0;
  storesData.forEach(function(s) {
    if (!s.error) {
      totalMTD         += (s.mtdRevenue  || 0);
      totalMonthTarget += (s.monthTarget || 0);
    }
  });
  var companyMPct = totalMonthTarget > 0 ? totalMTD / totalMonthTarget : 0;

  var weekTotalTarget = 0;
  CONFIG.STORES.forEach(function(store) {
    var t = (CONFIG.ANNUAL_TARGETS[store.name] || [])[currentMonth - 1] || 0;
    weekTotalTarget += Math.round(t / 4.3 * 10000);
  });

  var encouragement = CONFIG.CLAUDE_API_KEY
    ? generateEncouragementWithClaude(companyMPct, storesData)
    : getCompanyEncouragement(companyMPct);

  var message = buildMessage(storesData, thisWeek, lastWeek, currentMonth, companyMPct, weekTotalTarget, encouragement);
  Logger.log('> 訊息：' + message);
  sendToChat(message);

  pushDashboardData(storesData, lastWeek, currentMonth);
}

// ==========================================
// 讀取門市資料
// 月累計：加總當月每日 P 欄（不依賴合計列位置）
// ==========================================
function getStoreData(storeConfig, weekRange, currentMonth) {
  var monthTarget = ((CONFIG.ANNUAL_TARGETS[storeConfig.name] || [])[currentMonth - 1] || 0) * 10000;
  try {
    var ss   = SpreadsheetApp.openById(storeConfig.id);
    var year = new Date().getFullYear();

    // ── 月累計（MTD）：直接加總每日 P 欄 ──
    var tabName = year + '.' + String(currentMonth).padStart(2, '0');
    var sheet   = ss.getSheetByName(tabName);
    var mtdRevenue = 0;
    if (sheet) {
      var daysInMonth = new Date(year, currentMonth, 0).getDate(); // 當月天數
      var pValues     = sheet.getRange(6, 16, daysInMonth, 1).getValues(); // P6~P(5+天數)
      pValues.forEach(function(row) {
        mtdRevenue += toNumber(row[0]);
      });
    } else {
      Logger.log(storeConfig.name + ' 找不到分頁：' + tabName);
    }

    // ── 上週每日資料 ──
    var rows = getRowsInRange(ss, weekRange.start, weekRange.end);
    var weekRevenue        = 0;
    var weekTarget         = 0;
    var weekVisitors       = 0;
    var weekNewVisitors    = 0;
    var weekOldVisitors    = 0;
    var weekRepairVisitors = 0;
    var weekNewTxn         = 0;
    var weekOldTxn         = 0;
    var weekRepairConv     = 0;
    var weekAvgTicketSum   = 0;
    var weekAvgTicketDays  = 0;

    rows.forEach(function(row) {
      var dailyTarget  = toNumber(row[2]);
      var dailyRevenue = toNumber(row[15]);

      weekTarget          += dailyTarget;
      weekRevenue         += dailyRevenue;
      weekNewVisitors     += toNumber(row[8]);
      weekOldVisitors     += toNumber(row[9]);
      weekRepairVisitors  += toNumber(row[10]);
      weekNewTxn          += toNumber(row[11]);
      weekOldTxn          += toNumber(row[12]);
      weekRepairConv      += toNumber(row[13]);
      weekVisitors        += toNumber(row[8]) + toNumber(row[9]) + toNumber(row[10]);

      var dailyTxn = toNumber(row[11]) + toNumber(row[12]) + toNumber(row[13]) + toNumber(row[14]);
      if (dailyTxn > 0 && dailyRevenue > 0) {
        weekAvgTicketSum  += dailyRevenue / dailyTxn;
        weekAvgTicketDays += 1;
      }
    });

    var weekTotalTxn     = weekNewTxn + weekOldTxn + weekRepairConv;
    var weekAvgTicket    = weekAvgTicketDays > 0 ? Math.round(weekAvgTicketSum / weekAvgTicketDays) : 0;
    var weekConvRate     = weekVisitors > 0 ? weekTotalTxn / weekVisitors : 0;
    var weekAchievement  = weekTarget  > 0 ? weekRevenue  / weekTarget   : 0;
    var monthAchievement = monthTarget > 0 ? mtdRevenue   / monthTarget  : 0;

    return {
      name:                storeConfig.name,
      weekAchievement:     weekAchievement,
      weekRevenue:         Math.round(weekRevenue),
      weekTarget:          weekTarget,
      weekVisitors:        weekVisitors,
      weekNewVisitors:     weekNewVisitors,
      weekOldVisitors:     weekOldVisitors,
      weekRepairVisitors:  weekRepairVisitors,
      weekNewTxn:          weekNewTxn,
      weekOldTxn:          weekOldTxn,
      weekRepairConv:      weekRepairConv,
      weekTotalTxn:        weekTotalTxn,
      weekAvgTicket:       weekAvgTicket,
      weekConvRate:        weekConvRate,
      monthAchievement:    monthAchievement,
      mtdRevenue:          Math.round(mtdRevenue),
      monthTarget:         monthTarget,
    };
  } catch(e) {
    Logger.log('ERROR ' + storeConfig.name + ': ' + e.toString());
    return { name: storeConfig.name, error: e.toString(), monthTarget: monthTarget };
  }
}

// ==========================================
// Shopline API
// ==========================================
function getShoplineData(weekRange, currentMonth) {
  var monthTarget = ((CONFIG.ANNUAL_TARGETS['官網'] || [])[currentMonth - 1] || 0) * 10000;
  try {
    var weekStartMs = new Date(weekRange.start).getTime();
    var weekEndMs   = new Date(weekRange.end).getTime();
    var today       = new Date();
    var monthStart  = new Date(today.getFullYear(), currentMonth - 1, 1);

    // 一次拉取整月訂單，同時算週業績與 MTD（減少 API 呼叫次數）
    var allMtd = fetchShoplineOrders_(monthStart, today);

    // 本週前（月初到週前一天）已下過單的 customer_id = 舊客
    var priorCustIds = {};
    allMtd.forEach(function(o) {
      var t = new Date(o.created_at).getTime();
      if (t < weekStartMs && o.customer_id) priorCustIds[o.customer_id] = true;
    });

    var weekRevenue = 0, weekNewTxn = 0, weekOldTxn = 0, mtdRevenue = 0;
    var weekCustIds = {}; // 本週已出現的 customer_id
    var weekOrders  = 0;

    allMtd.forEach(function(o) {
      var dollars = (o.total && o.total.dollars != null) ? o.total.dollars : 0;
      mtdRevenue += dollars;
      var t = new Date(o.created_at).getTime();
      if (t >= weekStartMs && t <= weekEndMs) {
        weekRevenue += dollars;
        weekOrders++;
        var cid = o.customer_id || '';
        // 舊客：月初到上週已買過，或本週重複購買
        if (cid && (priorCustIds[cid] || weekCustIds[cid])) {
          weekOldTxn++;
        } else {
          weekNewTxn++;
        }
        if (cid) weekCustIds[cid] = true;
      }
    });

    var weekTotalTxn  = weekNewTxn + weekOldTxn;
    var weekAvgTicket = weekOrders > 0 ? Math.round(weekRevenue / weekOrders) : 0;
    var weekAchievement  = monthTarget > 0 ? weekRevenue / (monthTarget / 4.3) : 0;
    var monthAchievement = monthTarget > 0 ? mtdRevenue  / monthTarget          : 0;
    var monthOrders     = allMtd.length;
    var monthAvgTicket  = monthOrders > 0 ? Math.round(mtdRevenue / monthOrders) : 0;

    // GA4 流量資料（若服務未啟用會回傳 0）
    var gaWeek  = fetchGA4_(weekRange.start, weekRange.end);
    var gaMonth = fetchGA4_(monthStart, today);
    var weekGaConv  = gaWeek.sessions  > 0 ? weekOrders  / gaWeek.sessions  : 0;
    var monthGaConv = gaMonth.sessions > 0 ? monthOrders / gaMonth.sessions : 0;

    Logger.log('Shopline 上週營收：' + weekRevenue + '，月累計：' + mtdRevenue +
      '，週訂單：' + weekOrders + '，新客：' + weekNewTxn + '，舊客：' + weekOldTxn +
      '，GA 週 Sessions：' + gaWeek.sessions + '，月 Sessions：' + gaMonth.sessions);

    return {
      name:              '官網',
      weekAchievement:   weekAchievement,
      weekRevenue:       Math.round(weekRevenue),
      weekOrders:        weekOrders,
      weekAvgTicket:     weekAvgTicket,
      weekNewTxn:        weekNewTxn,
      weekOldTxn:        weekOldTxn,
      weekTotalTxn:      weekTotalTxn,
      weekConvRate:      0,
      weekVisitors:      0,
      weekNewVisitors:   0,
      weekOldVisitors:   0,
      weekRepairVisitors:0,
      weekRepairConv:    0,
      weekNewMembers:    weekNewTxn,
      monthAchievement:  monthAchievement,
      mtdRevenue:        Math.round(mtdRevenue),
      monthTarget:       monthTarget,
      monthOrders:       monthOrders,
      monthAvgTicket:    monthAvgTicket,
      weekSessions:      gaWeek.sessions,
      weekUsers:         gaWeek.users,
      weekGaConv:        weekGaConv,
      monthSessions:     gaMonth.sessions,
      monthUsers:        gaMonth.users,
      monthGaConv:       monthGaConv,
    };
  } catch(e) {
    Logger.log('Shopline ERROR: ' + e.toString());
    return { name: '官網', error: e.toString(), monthTarget: monthTarget };
  }
}

// ==========================================
// GA4 Data API（需在 Apps Script「服務」啟用 AnalyticsData）
// 使用前：Apps Script 編輯器 → 服務(+) → 新增「Google Analytics Data API」
// 執行帳號需對 GA4 Property 259936795 擁有至少檢視權限
// ==========================================
function fetchGA4_(startDate, endDate) {
  try {
    var request = {
      dateRanges: [{
        startDate: Utilities.formatDate(startDate, 'Asia/Taipei', 'yyyy-MM-dd'),
        endDate:   Utilities.formatDate(endDate,   'Asia/Taipei', 'yyyy-MM-dd'),
      }],
      metrics: [
        { name: 'sessions' },
        { name: 'totalUsers' },
      ],
    };
    var report = AnalyticsData.Properties.runReport(request, 'properties/' + CONFIG.GA4_PROPERTY_ID);
    var mv = (report.rows && report.rows[0]) ? report.rows[0].metricValues : [];
    var sessions = Number((mv[0] || {}).value || 0);
    var users    = Number((mv[1] || {}).value || 0);
    return { sessions: sessions, users: users };
  } catch (e) {
    Logger.log('GA4 ERROR: ' + e.toString());
    return { sessions: 0, users: 0, error: e.toString() };
  }
}

// created_from==='shop' 是線上訂單；admin_openapi 是實體門市 POS，排除
// 嚴格雙邊日期過濾：startMs <= t <= endMs，不依賴 API 排序
function fetchShoplineOrders_(startDate, endDate) {
  var startMs  = startDate.getTime();
  var endMs    = endDate.getTime();
  // ISO8601 for Shopline date filter params
  var minStr   = startDate.toISOString();
  var maxStr   = endDate.toISOString();
  var matched  = [], page = 1;
  while (page <= 50) {
    var url = 'https://open.shopline.io/v1/orders'
      + '?handle='           + CONFIG.SHOPLINE_HANDLE
      + '&per_page=100&page=' + page
      + '&created_at_min='   + encodeURIComponent(minStr)
      + '&created_at_max='   + encodeURIComponent(maxStr);
    var resp = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { 'Authorization': 'Bearer ' + CONFIG.SHOPLINE_TOKEN, 'Content-Type': 'application/json' },
      muteHttpExceptions: true,
    });
    if (resp.getResponseCode() !== 200) { Logger.log('Shopline HTTP ' + resp.getResponseCode() + ' p' + page); break; }
    var data   = JSON.parse(resp.getContentText());
    var orders = data.items || [];
    if (orders.length === 0) break;
    var allBeforeStart = true;
    for (var i = 0; i < orders.length; i++) {
      var t = new Date(orders[i].created_at).getTime();
      if (t >= startMs) allBeforeStart = false;
      if (orders[i].created_from !== 'shop') continue;
      if (t >= startMs && t <= endMs && orders[i].status !== 'cancelled') matched.push(orders[i]);
    }
    // 若整頁訂單都早於查詢起始日，代表已超出範圍，提早結束
    if (allBeforeStart) { Logger.log('Shopline early-exit at p' + page); break; }
    var perPage = (data.pagination || {}).per_page || 100;
    if (orders.length < perPage) break;
    page++;
  }
  Logger.log('Shopline matched=' + matched.length + ' pages=' + (page - 1));
  return matched;
}

function fetchShoplineRevenue(startDate, endDate) {
  var total = 0;
  fetchShoplineOrders_(startDate, endDate).forEach(function(order) {
    total += (order.total && order.total.dollars != null) ? order.total.dollars : 0;
  });
  return total;
}

function fetchShoplineOrderCount(startDate, endDate) {
  return fetchShoplineOrders_(startDate, endDate).length;
}

// ==========================================
// 推送儀表板數據
// ==========================================
function pushDashboardData(storesData, lastWeek, currentMonth) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.DASHBOARD_SHEET_ID);
    pushWeeklyTab(ss, storesData, lastWeek);
    pushMonthlyTab(ss, storesData, currentMonth);
    Logger.log('✅ 儀表板數據已推送');
  } catch(e) {
    Logger.log('pushDashboardData ERROR: ' + e.toString());
  }
}

function pushWeeklyTab(ss, storesData, lastWeek) {
  var tab = ss.getSheetByName(CONFIG.DASHBOARD_TAB_WEEK);
  if (!tab) tab = ss.insertSheet(CONFIG.DASHBOARD_TAB_WEEK);
  tab.clearContents();
  var weekLabel = formatDate(lastWeek.start) + ' ~ ' + formatDate(lastWeek.end);
  var now = formatDate(new Date()) + ' ' + new Date().toTimeString().slice(0, 5);
  tab.appendRow(['更新時間','週期','門市','週營收(元)','週目標(元)','週達成率',
    '總來客','新客來客','舊客來客','維修來客',
    '新客成交','回購成交','總成交','客單價','成交率',
    '月累計(元)','月目標(元)','月達成率','週訂單數(電商)','維修成交','新會員數',
    '週Sessions','週Users','週GA轉換率','月Sessions','月Users','月GA轉換率',
    '月訂單數','月客單價']);
  storesData.forEach(function(s) {
    if (s.error) { tab.appendRow([now, weekLabel, s.name, 'ERROR: ' + s.error]); return; }
    var weekTgt = s.monthTarget > 0 ? Math.round(s.monthTarget / 4.3) : 0;
    tab.appendRow([now, weekLabel, s.name,
      Math.round(s.weekRevenue  || 0), weekTgt, s.weekAchievement  || 0,
      s.weekVisitors || 0, s.weekNewVisitors || 0, s.weekOldVisitors || 0, s.weekRepairVisitors || 0,
      s.weekNewTxn   || 0, s.weekOldTxn     || 0, s.weekTotalTxn    || 0,
      s.weekAvgTicket|| 0, s.weekConvRate    || 0,
      Math.round(s.mtdRevenue || 0), Math.round(s.monthTarget || 0), s.monthAchievement || 0,
      s.weekOrders   || 0, s.weekRepairConv  || 0, s.weekNewMembers  || 0,
      s.weekSessions || 0, s.weekUsers       || 0, s.weekGaConv     || 0,
      s.monthSessions|| 0, s.monthUsers      || 0, s.monthGaConv    || 0,
      s.monthOrders  || 0, s.monthAvgTicket  || 0,
    ]);
  });
  Logger.log('週報數據分頁已更新');
}

function pushMonthlyTab(ss, storesData, currentMonth) {
  var tab = ss.getSheetByName(CONFIG.DASHBOARD_TAB_MONTHLY);
  if (!tab) {
    tab = ss.insertSheet(CONFIG.DASHBOARD_TAB_MONTHLY);
    tab.appendRow(['年份','月份','門市','月營收(元)','月目標(元)','月達成率','更新時間']);
  }
  var year = new Date().getFullYear();
  var now  = formatDate(new Date());
  storesData.forEach(function(s) {
    if (s.error) return;
    var data = tab.getDataRange().getValues();
    var foundRow = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === year && data[i][1] === currentMonth && data[i][2] === s.name) {
        foundRow = i + 1; break;
      }
    }
    var rowData = [year, currentMonth, s.name,
      Math.round(s.mtdRevenue || 0), Math.round(s.monthTarget || 0),
      s.monthAchievement || 0, now];
    if (foundRow > 0) {
      tab.getRange(foundRow, 1, 1, rowData.length).setValues([rowData]);
    } else {
      tab.appendRow(rowData);
    }
  });
  Logger.log('月度數據分頁已更新');
}

// ==========================================
// 組合 Google Chat 訊息
// ==========================================
function buildMessage(storesData, thisWeek, lastWeek, currentMonth, companyMPct, weekTotalTarget, encouragement) {
  var lws = (lastWeek.start.getMonth() + 1) + '/' + lastWeek.start.getDate();
  var lwe = (lastWeek.end.getMonth()   + 1) + '/' + lastWeek.end.getDate();
  var tws = (thisWeek.start.getMonth() + 1) + '/' + thisWeek.start.getDate();
  var twe = (thisWeek.end.getMonth()   + 1) + '/' + thisWeek.end.getDate();

  var msg = '';
  msg += '*林果營業額週報 | ' + formatDate(new Date()) + '（週一）*\n\n';
  msg += '*📊 上週（' + lws + '-' + lwe + '）業績 + ' + currentMonth + '月達成率*\n';

  var physicalStores = storesData.filter(function(s) {
    return !s.error && s.name !== '官網' && s.weekAchievement > 0;
  });
  var bestStore = physicalStores.slice().sort(function(a, b) {
    return b.weekAchievement - a.weekAchievement;
  })[0];

  storesData.forEach(function(store) {
    if (store.error) { msg += '⚠️ *' + store.name + '*　資料讀取異常\n'; return; }
    var wPct = Math.round((store.weekAchievement  || 0) * 100);
    var mPct = Math.round((store.monthAchievement || 0) * 100);
    if (store.name === '官網') {
      var emoji = wPct >= 100 ? '🏆' : wPct >= 80 ? '✅' : wPct >= 60 ? '💪' : '🌐';
      var ecExtra = '';
      if (store.weekOrders > 0) {
        ecExtra = '\n     📦 訂單 *' + store.weekOrders + '* 筆';
        if (store.weekAvgTicket > 0) ecExtra += '　客單 *NT$' + formatNumber(store.weekAvgTicket) + '*';
      }
      msg += emoji + ' *官網*　上週 *' + wPct + '%*（NT$' + formatNumber(store.weekRevenue) + '）／ ' + currentMonth + '月 *' + mPct + '%*' + ecExtra + '\n';
    } else {
      var emoji = wPct >= 100 ? '🏆' : wPct >= 80 ? '✅' : wPct >= 60 ? '💪' : '📈';
      var visitorsText = '';
      if (store.weekVisitors > 0) {
        visitorsText = '\n     👥 來客 *' + formatNumber(store.weekVisitors) + '* 人';
        if (store.weekAvgTicket > 0) visitorsText += '　客單 *NT$' + formatNumber(store.weekAvgTicket) + '*';
      }
      msg += emoji + ' *' + store.name + '*　上週 *' + wPct + '%*（NT$' + formatNumber(store.weekRevenue) + '）／ ' + currentMonth + '月 *' + mPct + '%*' + visitorsText + '\n';
    }
  });

  var totalVisitors = 0, totalRevenue = 0, totalNewTxn = 0, totalOldTxn = 0, totalTxn = 0;
  storesData.forEach(function(s) {
    if (!s.error && s.name !== '官網') {
      totalVisitors += (s.weekVisitors  || 0);
      totalRevenue  += (s.weekRevenue   || 0);
      totalNewTxn   += (s.weekNewTxn    || 0);
      totalOldTxn   += (s.weekOldTxn    || 0);
      totalTxn      += (s.weekTotalTxn  || 0);
    }
  });

  var totalAvgTicket = totalTxn > 0 ? Math.round(totalRevenue / totalTxn) : 0;
  var totalTicketPct = totalAvgTicket > 0 ? Math.round(totalAvgTicket / CONFIG.BENCHMARKS.AVG_TICKET * 100) : 0;
  var totalConvRate  = totalVisitors > 0 ? Math.round(totalTxn / totalVisitors * 100) : 0;
  var convRateDiff   = totalConvRate - Math.round(CONFIG.BENCHMARKS.CONV_RATE * 100);
  var convDiffText   = convRateDiff >= 0 ? '↑ +' + convRateDiff + '%' : '↓ ' + convRateDiff + '%';

  var newOldText = '';
  if (totalTxn > 0) {
    var newRatio = Math.round(totalNewTxn / totalTxn * 10);
    var oldRatio = 10 - newRatio;
    newOldText = '（新舊客 ' + newRatio + ':' + oldRatio + '）';
  }

  var companyPct = Math.round(companyMPct * 100);
  msg += '━━━━━━━━━━━━━━━━━━━━\n';
  if (totalVisitors > 0) msg += '👥 *四店來客：' + formatNumber(totalVisitors) + ' 人' + newOldText + '*\n';
  if (totalAvgTicket > 0) msg += '🧾 *四店客單：NT$' + formatNumber(totalAvgTicket) + '（達成 ' + totalTicketPct + '%）*\n';
  if (totalConvRate > 0) msg += '🎯 *四店成交率：' + totalConvRate + '%（基準 ' + Math.round(CONFIG.BENCHMARKS.CONV_RATE * 100) + '%，' + convDiffText + '）*\n';
  msg += '📌 *全公司 ' + currentMonth + '月營業額達成率：' + companyPct + '%*\n\n';
  if (weekTotalTarget > 0) msg += '🎯 *本週（' + tws + '-' + twe + '）四店目標：NT$ ' + formatNumber(weekTotalTarget) + '*\n\n';
  if (bestStore) msg += '🥇 上週 *' + bestStore.name + '* 達成率最高，替大家拍拍手！\n\n';
  msg += '💬 ' + encouragement + '\n';
  msg += '加油！林果全體夥伴 💪👟\n\n';
  msg += '📊 *完整儀表板：* ' + CONFIG.DASHBOARD_URL;

  return msg;
}

// ==========================================
// 日期範圍工具
// ==========================================
function getWeekRanges() {
  var today = new Date();
  var dow   = today.getDay();

  var thisMonday = new Date(today);
  thisMonday.setDate(today.getDate() - (dow === 0 ? 6 : dow - 1));
  thisMonday.setHours(0, 0, 0, 0);

  var thisSunday = new Date(thisMonday);
  thisSunday.setDate(thisMonday.getDate() + 6);
  thisSunday.setHours(23, 59, 59, 0);

  var lastMonday = new Date(thisMonday);
  lastMonday.setDate(thisMonday.getDate() - 7);

  var lastSunday = new Date(thisMonday);
  lastSunday.setDate(thisMonday.getDate() - 1);
  lastSunday.setHours(23, 59, 59, 0);

  return {
    lastWeek: { start: lastMonday, end: lastSunday },
    thisWeek: { start: thisMonday, end: thisSunday },
  };
}

function getRowsInRange(ss, startDate, endDate) {
  var results = [];
  var months  = getMonthsInRange(startDate, endDate);
  months.forEach(function(m) {
    var tabName = m.year + '.' + String(m.month).padStart(2, '0');
    var sheet   = ss.getSheetByName(tabName);
    if (!sheet) { Logger.log('找不到分頁：' + tabName); return; }
    var data = sheet.getDataRange().getValues();
    for (var i = 5; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      var cellDate = row[0] instanceof Date ? row[0] : new Date(row[0]);
      if (isNaN(cellDate.getTime())) continue;
      var d = ymd(cellDate), s = ymd(startDate), e = ymd(endDate);
      if (d >= s && d <= e) results.push(row);
    }
  });
  return results;
}

function getMonthsInRange(startDate, endDate) {
  var months = [];
  var cur    = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
  var end    = new Date(endDate.getFullYear(),   endDate.getMonth(),   1);
  while (cur <= end) {
    months.push({ year: cur.getFullYear(), month: cur.getMonth() + 1 });
    cur.setMonth(cur.getMonth() + 1);
  }
  return months;
}

function ymd(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

// ==========================================
// 鼓勵語
// ==========================================
function getCompanyEncouragement(companyMPct) {
  var pct        = Math.round(companyMPct * 100);
  var weekNumber = Math.floor(new Date().getTime() / (7 * 24 * 60 * 60 * 1000));
  var msgs;
  if (pct >= 100) {
    msgs = ['全公司達標！謝謝每位夥伴的努力，繼續保持，加油！','漂亮過關！成果是用心換來的，繼續衝，加油！','全員達成！每筆銷售都算數，繼續前進，加油！'];
  } else if (pct >= 85) {
    msgs = ['差一步就到了，本週再拼一把，加油！','快到了，把每位客人服務好，數字自然到位，加油！','基礎打穩了，後面繼續衝，加油！'];
  } else if (pct >= 70) {
    msgs = ['進度還有空間，每天一步一步來，加油！','方向對了，把節奏抓回來，加油！','每位客人都是機會，一起把數字拉上來，加油！'];
  } else if (pct >= 50) {
    msgs = ['調整節奏再出發，每天的累積都算數，加油！','重新出發，把每位客人服務好，成果會跟上，加油！'];
  } else {
    msgs = ['先把每位客人服務好，一步一步往目標走，加油！','每天重新出發，大家一起一步步往前，加油！'];
  }
  return msgs[weekNumber % msgs.length];
}

function generateEncouragementWithClaude(companyMPct, storesData) {
  try {
    var pct = Math.round(companyMPct * 100);
    var storeSummary = storesData.filter(function(s) { return !s.error; })
      .map(function(s) { return s.name + '：月達成 ' + Math.round((s.monthAchievement || 0) * 100) + '%'; }).join('；');
    var prompt = '你是林果鞋業的營運夥伴，用溫暖簡短的語氣給全體員工每週一句話。全公司本月達成率：' + pct + '%，各店況：' + storeSummary + '。請根據達成率調整語氣（低則給力加油，高則讚美鼓勵），約 25-35 字，繁體中文，結尾一定要有「加油」，直接有溫度，不要廢話。只輸出激勵語本身。';
    var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      headers: { 'x-api-key': CONFIG.CLAUDE_API_KEY, 'anthropic-version': '2023-06-01', 'content-type': 'application/json' },
      payload: JSON.stringify({ model: 'claude-haiku-4-5-20251001', max_tokens: 200, messages: [{ role: 'user', content: prompt }] }),
    });
    return JSON.parse(response.getContentText()).content[0].text.trim();
  } catch(e) {
    Logger.log('Claude API failed: ' + e.toString());
    return getCompanyEncouragement(companyMPct);
  }
}

// ==========================================
// 發送至 Google Chat
// ==========================================
function sendToChat(message) {
  var options = { method: 'post', contentType: 'application/json', payload: JSON.stringify({ text: message }) };
  try {
    var resp = UrlFetchApp.fetch(CONFIG.WEBHOOK, options);
    Logger.log('Chat 回應：' + resp.getResponseCode());
  } catch(e) {
    Logger.log('Chat 發送失敗：' + e.toString());
    throw e;
  }
}

// ==========================================
// 工具函式
// ==========================================
function toNumber(value) {
  if (typeof value === 'number') return isNaN(value) ? 0 : value;
  if (typeof value === 'string') {
    var isPercent = value.indexOf('%') >= 0;
    var num = parseFloat(value.replace(/[%,]/g, '').trim());
    if (isNaN(num)) return 0;
    return isPercent ? num / 100 : num;
  }
  return 0;
}

function formatNumber(num) {
  if (!num || isNaN(num)) return '0';
  return Math.round(num).toLocaleString('zh-TW');
}

function formatDate(date) {
  return date.getFullYear() + '/' + (date.getMonth() + 1) + '/' + date.getDate();
}

// ==========================================
// 測試用函式
// ==========================================
function test() {
  sendWeeklyReport();
}

function fixShoplineJan() {
  var year  = 2026, month = 1;
  var start = new Date(year, 0, 1);
  var end   = new Date(year, 1, 0, 23, 59, 59, 999); // Jan 31
  var rev   = fetchShoplineRevenue(start, end);
  var tgt   = ((CONFIG.ANNUAL_TARGETS['官網'] || [])[0] || 0) * 10000;
  Logger.log('官網 2026/1 營收：' + rev + '，目標：' + tgt);

  var ss  = SpreadsheetApp.openById(CONFIG.DASHBOARD_SHEET_ID);
  var tab = ss.getSheetByName(CONFIG.DASHBOARD_TAB_MONTHLY);
  var data = tab.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === year && data[i][1] === month && data[i][2] === '官網') {
      tab.getRange(i + 1, 4).setValue(Math.round(rev));
      tab.getRange(i + 1, 6).setValue(tgt > 0 ? rev / tgt : 0);
      Logger.log('✅ 已更新第 ' + (i+1) + ' 列');
      return;
    }
  }
  Logger.log('找不到對應列，改用 append');
  tab.appendRow([year, month, '官網', Math.round(rev), tgt, tgt > 0 ? rev / tgt : 0, '手動修正']);
}

function diagZhongshan() {
  var store = CONFIG.STORES[0]; // 中山店
  var ranges = getWeekRanges();
  var ss = SpreadsheetApp.openById(store.id);

  // 確認上週分頁存在
  var months = getMonthsInRange(ranges.lastWeek.start, ranges.lastWeek.end);
  months.forEach(function(m) {
    var tabName = m.year + '.' + String(m.month).padStart(2, '0');
    var sheet = ss.getSheetByName(tabName);
    if (!sheet) { Logger.log('❌ 找不到分頁：' + tabName); return; }
    Logger.log('✅ 找到分頁：' + tabName + '，共 ' + sheet.getLastRow() + ' 列，' + sheet.getLastColumn() + ' 欄');

    // 印出上週每日的 A欄(日期) 和 P欄(row[15])
    var data = sheet.getDataRange().getValues();
    var s = ymd(ranges.lastWeek.start), e = ymd(ranges.lastWeek.end);
    for (var i = 5; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      var cellDate = row[0] instanceof Date ? row[0] : new Date(row[0]);
      if (isNaN(cellDate.getTime())) continue;
      var d = ymd(cellDate);
      if (d >= s && d <= e) {
        Logger.log('  ' + Utilities.formatDate(cellDate,'Asia/Taipei','M/d') +
          ' | C(目標)=' + row[2] + ' | P(金額)=' + row[15] +
          ' | 總欄數=' + row.length);
      }
    }
  });

  // MTD（P欄加總）
  var tabName = new Date().getFullYear() + '.' + String(new Date().getMonth()+1).padStart(2,'0');
  var sheet = ss.getSheetByName(tabName);
  if (sheet) {
    var days = new Date().getDate();
    var pVals = sheet.getRange(6, 16, days, 1).getValues();
    var mtd = 0;
    pVals.forEach(function(r){ mtd += toNumber(r[0]); });
    Logger.log('本月 MTD (P欄加總 ' + days + '天) = ' + mtd);
  }
}

function diagShoplineRaw() {
  var url = 'https://open.shopline.io/v1/orders'
    + '?handle=' + CONFIG.SHOPLINE_HANDLE
    + '&per_page=5&page=1';
  var resp = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'Authorization': 'Bearer ' + CONFIG.SHOPLINE_TOKEN, 'Content-Type': 'application/json' },
    muteHttpExceptions: true,
  });
  Logger.log('HTTP: ' + resp.getResponseCode());
  var data = JSON.parse(resp.getContentText());
  Logger.log('pagination: ' + JSON.stringify(data.pagination || {}));
  var orders = data.items || [];
  orders.forEach(function(o, i) {
    var t = new Date(o.created_at).getTime();
    Logger.log('Order[' + i + '] created_at=' + o.created_at + ' | t=' + t + ' | created_from=' + o.created_from + ' | status=' + o.status + ' | total.dollars=' + (o.total ? o.total.dollars : '?'));
  });
}

// 診斷 Shopline 客戶欄位（確認 orders_count 是否有回傳）
function diagShoplineCustomers() {
  var url = 'https://open.shopline.io/v1/orders'
    + '?handle=' + CONFIG.SHOPLINE_HANDLE
    + '&per_page=10&page=1';
  var resp = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'Authorization': 'Bearer ' + CONFIG.SHOPLINE_TOKEN, 'Content-Type': 'application/json' },
    muteHttpExceptions: true,
  });
  var orders = (JSON.parse(resp.getContentText()).items || []);
  var shopOrders = orders.filter(function(o){ return o.created_from === 'shop'; });
  shopOrders.forEach(function(o, i) {
    var c = o.customer || {};
    Logger.log('Order[' + i + '] id=' + o.id
      + ' | customer.id='         + c.id
      + ' | orders_count='        + c.orders_count
      + ' | total_spent='         + c.total_spent
      + ' | accepts_marketing='   + c.accepts_marketing
      + ' | state='               + c.state
      + ' | tags='                + c.tags);
  });
  Logger.log('共 ' + shopOrders.length + ' 筆網店訂單（前10頁）');
}

function testShopline() {
  var ranges = getWeekRanges();
  var data = getShoplineData(ranges.lastWeek, new Date().getMonth() + 1);
  Logger.log(JSON.stringify(data));
}

function testPushDashboard() {
  var ranges       = getWeekRanges();
  var currentMonth = new Date().getMonth() + 1;
  var storesData   = CONFIG.STORES.map(function(store) {
    return getStoreData(store, ranges.lastWeek, currentMonth);
  });
  var ecData = getShoplineData(ranges.lastWeek, currentMonth);
  storesData.push(ecData);
  pushDashboardData(storesData, ranges.lastWeek, currentMonth);
}

// ==========================================
// 歷史資料回填：補齊 2026/01 ~ 本月的月度數據
// 在 Apps Script 編輯器手動執行此函式一次即可
// ==========================================
function backfillHistoricalData() {
  var ss           = SpreadsheetApp.openById(CONFIG.DASHBOARD_SHEET_ID);
  var today        = new Date();
  var currentYear  = today.getFullYear();
  var currentMonth = today.getMonth() + 1;

  for (var month = 1; month <= currentMonth; month++) {
    Logger.log('▶ 回填 ' + currentYear + '/' + month + ' ...');

    var monthStart = new Date(currentYear, month - 1, 1);
    var isCurrentMonth = (month === currentMonth);
    var monthEnd = isCurrentMonth
      ? new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59, 59, 999)
      : new Date(currentYear, month, 0, 23, 59, 59, 999);

    // 實體門市
    var storesData = CONFIG.STORES.map(function(storeConfig) {
      var monthTarget = ((CONFIG.ANNUAL_TARGETS[storeConfig.name] || [])[month - 1] || 0) * 10000;
      try {
        var storeSs  = SpreadsheetApp.openById(storeConfig.id);
        var tabName  = currentYear + '.' + String(month).padStart(2, '0');
        var sheet    = storeSs.getSheetByName(tabName);
        var mtdRevenue = 0;
        if (sheet) {
          var daysToCount = isCurrentMonth ? today.getDate() : new Date(currentYear, month, 0).getDate();
          var pValues = sheet.getRange(6, 16, daysToCount, 1).getValues();
          pValues.forEach(function(row) { mtdRevenue += toNumber(row[0]); });
        } else {
          Logger.log('  找不到分頁：' + tabName);
        }
        return {
          name:             storeConfig.name,
          mtdRevenue:       Math.round(mtdRevenue),
          monthTarget:      monthTarget,
          monthAchievement: monthTarget > 0 ? mtdRevenue / monthTarget : 0,
        };
      } catch(e) {
        Logger.log('  ERROR ' + storeConfig.name + ': ' + e);
        return { name: storeConfig.name, error: e.toString(), monthTarget: monthTarget };
      }
    });

    // 官網（Shopline）
    var shoplineTgt = ((CONFIG.ANNUAL_TARGETS['官網'] || [])[month - 1] || 0) * 10000;
    try {
      var mtdRevenue = fetchShoplineRevenue(monthStart, monthEnd);
      storesData.push({
        name:             '官網',
        mtdRevenue:       Math.round(mtdRevenue),
        monthTarget:      shoplineTgt,
        monthAchievement: shoplineTgt > 0 ? mtdRevenue / shoplineTgt : 0,
      });
    } catch(e) {
      Logger.log('  Shopline ERROR: ' + e);
      storesData.push({ name: '官網', error: e.toString(), monthTarget: shoplineTgt });
    }

    // 寫入月度數據分頁
    var tab = ss.getSheetByName(CONFIG.DASHBOARD_TAB_MONTHLY);
    if (!tab) {
      tab = ss.insertSheet(CONFIG.DASHBOARD_TAB_MONTHLY);
      tab.appendRow(['年份','月份','門市','月營收(元)','月目標(元)','月達成率','更新時間']);
    }
    var now = formatDate(new Date());
    storesData.forEach(function(s) {
      if (s.error) return;
      var data = tab.getDataRange().getValues();
      var foundRow = -1;
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] === currentYear && data[i][1] === month && data[i][2] === s.name) {
          foundRow = i + 1; break;
        }
      }
      var rowData = [currentYear, month, s.name,
        Math.round(s.mtdRevenue || 0), Math.round(s.monthTarget || 0),
        s.monthAchievement || 0, now];
      if (foundRow > 0) {
        tab.getRange(foundRow, 1, 1, rowData.length).setValues([rowData]);
      } else {
        tab.appendRow(rowData);
      }
    });
    Logger.log('  ✅ ' + currentYear + '/' + month + ' 寫入完成');
    Utilities.sleep(3000); // 避免 API 頻率限制
  }
  Logger.log('🎉 全部回填完成');
}
