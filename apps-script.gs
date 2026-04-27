// ==========================================
// 林果鞋業｜週報自動發送腳本 v3.8
// 每週一 9:00 發送至 Google Chat
//
// v3.8 變更：
//   - 週目標換算改用「該月實際天數 ÷ 7」當週數，取代固定 ÷4.3
//     例：4月 30 天 → 4.286 週；2月 28 天 → 4 週；3月 31 天 → 4.43 週
//   - 影響：pushWeeklyTab 的 weekTgt、buildMessage 的 weekTotalTarget、
//     getShoplineData 的 weekAchievement、getStoreData 的 weekVisitorTarget
// v3.7 變更：
//   - fetchShoplineOrders_ 加入 CacheService 快取（chunked），避免 UrlFetchApp 頻寬配額用盡
//     當月 TTL 30 分鐘、歷史月份 6 小時；快取命中時不發外部請求
//   - 新增 fetchWithBackoff_ 包裝器：429/5xx 自動指數退避重試（1s→3s→10s→30s）
//   - Shopline 分頁間 sleep 300ms，降低 Shopline 限速風險
//   - 新增 clearShoplineCache() 工具函式：強制清除 Shopline 快取
// v3.6 變更：
//   - 新增 CONFIG.STORE_BENCHMARKS：各門市轉換率/客單價目標
//   - 維修成交數改用 O 欄（惜履券，row[14]），原本錯用 N 欄（VIP）
//   - VIP 成交數獨立追蹤，併入 weekTotalTxn 維持總成交完整
//   - buildMessage 改用各店個別基準判斷成交率/客單達成
// v3.5：整合 GA4 Sessions + 官網月訂單/月客單
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
//   O(row[14]) = 惜履券成交數（= 維修成交）
//   P(row[15]) = 總成交金額（元）
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

  // 各門市轉換率/客單價目標（用於來客數目標反推 + 達成判斷）
  STORE_BENCHMARKS: {
    '中山店': { CONV_RATE: 0.58, AVG_TICKET: 6500 },
    '松菸店': { CONV_RATE: 0.58, AVG_TICKET: 6800 },
    '台中店': { CONV_RATE: 0.56, AVG_TICKET: 6700 },
    '東門店': { CONV_RATE: 0.55, AVG_TICKET: 6500 },
  },

  // 全公司基準（向後相容用，已不再做為門市判斷依據）
  BENCHMARKS: {
    AVG_TICKET: 6625,  // 四店平均
    CONV_RATE:  0.5675, // 四店平均
  },

  EC_TARGETS: {
    SESSIONS:   850000,
    CONV_RATE:  0.0035,
    AVG_TICKET: 5000,
  },

  CLAUDE_API_KEY: '',
};

// ==========================================
// 每日執行：只刷新儀表板資料
// ==========================================
function refreshDashboard() {
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

  pushDashboardData(storesData, lastWeek, currentMonth);

  Logger.log('✅ 儀表板資料已更新（' + formatDate(today) + '）');

  return {
    storesData: storesData,
    thisWeek: thisWeek,
    lastWeek: lastWeek,
    currentMonth: currentMonth
  };
}

// ==========================================
// 每週一執行：刷新資料 + 發送 Google Chat 週報
// ==========================================
function sendWeeklyReport() {
  var result = refreshDashboard();
  var storesData   = result.storesData;
  var thisWeek     = result.thisWeek;
  var lastWeek     = result.lastWeek;
  var currentMonth = result.currentMonth;

  var totalMTD = 0, totalMonthTarget = 0;
  storesData.forEach(function(s) {
    if (!s.error) {
      totalMTD         += (s.mtdRevenue  || 0);
      totalMonthTarget += (s.monthTarget || 0);
    }
  });
  var companyMPct = totalMonthTarget > 0 ? totalMTD / totalMonthTarget : 0;

  var weekTotalTarget = 0;
  var _today        = new Date();
  var _daysInMonth  = new Date(_today.getFullYear(), currentMonth, 0).getDate();
  var _weeksInMonth = _daysInMonth / 7;
  CONFIG.STORES.forEach(function(store) {
    var t = (CONFIG.ANNUAL_TARGETS[store.name] || [])[currentMonth - 1] || 0;
    weekTotalTarget += Math.round(t / _weeksInMonth * 10000);
  });

  var encouragement = CONFIG.CLAUDE_API_KEY
    ? generateEncouragementWithClaude(companyMPct, storesData)
    : getCompanyEncouragement(companyMPct);

  var message = buildMessage(storesData, thisWeek, lastWeek, currentMonth, companyMPct, weekTotalTarget, encouragement);
  Logger.log('> 訊息：' + message);
  sendToChat(message);
}

// ==========================================
// 讀取門市資料
// ==========================================
function getStoreData(storeConfig, weekRange, currentMonth) {
  var monthTarget = ((CONFIG.ANNUAL_TARGETS[storeConfig.name] || [])[currentMonth - 1] || 0) * 10000;
  var benchmark   = CONFIG.STORE_BENCHMARKS[storeConfig.name] || CONFIG.BENCHMARKS;

  try {
    var ss   = SpreadsheetApp.openById(storeConfig.id);
    var year = new Date().getFullYear();

    var tabName = year + '.' + String(currentMonth).padStart(2, '0');
    var sheet   = ss.getSheetByName(tabName);
    var mtdRevenue = 0;
    if (sheet) {
      var daysInMonth = new Date(year, currentMonth, 0).getDate();
      var pValues     = sheet.getRange(6, 16, daysInMonth, 1).getValues();
      pValues.forEach(function(row) {
        mtdRevenue += toNumber(row[0]);
      });
    } else {
      Logger.log(storeConfig.name + ' 找不到分頁：' + tabName);
    }

    var rows = getRowsInRange(ss, weekRange.start, weekRange.end);
    var weekRevenue=0, weekTarget=0, weekVisitors=0;
    var weekNewVisitors=0, weekOldVisitors=0, weekRepairVisitors=0;
    var weekNewTxn=0, weekOldTxn=0, weekVipTxn=0, weekRepairConv=0;
    var weekAvgTicketSum=0, weekAvgTicketDays=0;

    rows.forEach(function(row) {
      var dailyTarget  = toNumber(row[2]);
      var dailyRevenue = toNumber(row[15]);
      weekTarget         += dailyTarget;
      weekRevenue        += dailyRevenue;
      weekNewVisitors    += toNumber(row[8]);
      weekOldVisitors    += toNumber(row[9]);
      weekRepairVisitors += toNumber(row[10]);
      weekNewTxn         += toNumber(row[11]);
      weekOldTxn         += toNumber(row[12]);
      weekVipTxn         += toNumber(row[13]);   // VIP 成交（獨立追蹤）
      weekRepairConv     += toNumber(row[14]);   // 維修成交 = 惜履券（v3.6 修正）
      weekVisitors       += toNumber(row[8]) + toNumber(row[9]) + toNumber(row[10]);
      var dailyTxn = toNumber(row[11]) + toNumber(row[12]) + toNumber(row[13]) + toNumber(row[14]);
      if (dailyTxn > 0 && dailyRevenue > 0) {
        weekAvgTicketSum  += dailyRevenue / dailyTxn;
        weekAvgTicketDays += 1;
      }
    });

    var weekTotalTxn     = weekNewTxn + weekOldTxn + weekVipTxn + weekRepairConv;
    var weekAvgTicket    = weekAvgTicketDays > 0 ? Math.round(weekAvgTicketSum / weekAvgTicketDays) : 0;
    var weekConvRate     = weekVisitors > 0 ? weekTotalTxn / weekVisitors : 0;
    var weekAchievement  = weekTarget  > 0 ? weekRevenue  / weekTarget   : 0;
    var monthAchievement = monthTarget > 0 ? mtdRevenue   / monthTarget  : 0;

    // 來客數目標（基於各店基準反推）：月目標 ÷ 客單 ÷ 轉換率 ÷ (該月天數/7)
    var daysInMonth  = new Date(year, currentMonth, 0).getDate();
    var weeksInMonth = daysInMonth / 7;
    var weekVisitorTarget = (monthTarget > 0 && benchmark.AVG_TICKET > 0 && benchmark.CONV_RATE > 0)
      ? Math.round(monthTarget / benchmark.AVG_TICKET / benchmark.CONV_RATE / weeksInMonth)
      : 0;

    return {
      name:                storeConfig.name,
      weekAchievement:     weekAchievement,
      weekRevenue:         Math.round(weekRevenue),
      weekTarget:          weekTarget,
      weekVisitors:        weekVisitors,
      weekVisitorTarget:   weekVisitorTarget,
      weekNewVisitors:     weekNewVisitors,
      weekOldVisitors:     weekOldVisitors,
      weekRepairVisitors:  weekRepairVisitors,
      weekNewTxn:          weekNewTxn,
      weekOldTxn:          weekOldTxn,
      weekVipTxn:          weekVipTxn,
      weekRepairConv:      weekRepairConv,
      weekTotalTxn:        weekTotalTxn,
      weekAvgTicket:       weekAvgTicket,
      weekConvRate:        weekConvRate,
      monthAchievement:    monthAchievement,
      mtdRevenue:          Math.round(mtdRevenue),
      monthTarget:         monthTarget,
      benchmarkConv:       benchmark.CONV_RATE,
      benchmarkTicket:     benchmark.AVG_TICKET,
    };
  } catch(e) {
    Logger.log('ERROR ' + storeConfig.name + ': ' + e.toString());
    return { name: storeConfig.name, error: e.toString(), monthTarget: monthTarget };
  }
}

// ==========================================
// 官網 Shopline + GA4
// ==========================================
function getShoplineData(weekRange, currentMonth) {
  var monthTarget = ((CONFIG.ANNUAL_TARGETS['官網'] || [])[currentMonth - 1] || 0) * 10000;
  try {
    var weekStartMs = new Date(weekRange.start).getTime();
    var weekEndMs   = new Date(weekRange.end).getTime();
    var today       = new Date();
    var monthStart  = new Date(today.getFullYear(), currentMonth - 1, 1);

    var allMtd = fetchShoplineOrders_(monthStart, today);

    var priorCustIds = {};
    allMtd.forEach(function(o) {
      var t = new Date(o.created_at).getTime();
      if (t < weekStartMs && o.customer_id) priorCustIds[o.customer_id] = true;
    });

    var weekRevenue = 0, weekNewTxn = 0, weekOldTxn = 0, mtdRevenue = 0;
    var weekCustIds = {};
    var weekOrders  = 0;

    allMtd.forEach(function(o) {
      var dollars = (o.total && o.total.dollars != null) ? o.total.dollars : 0;
      mtdRevenue += dollars;
      var t = new Date(o.created_at).getTime();
      if (t >= weekStartMs && t <= weekEndMs) {
        weekRevenue += dollars;
        weekOrders++;
        var cid = o.customer_id || '';
        if (cid && (priorCustIds[cid] || weekCustIds[cid])) {
          weekOldTxn++;
        } else {
          weekNewTxn++;
        }
        if (cid) weekCustIds[cid] = true;
      }
    });

    var weekTotalTxn    = weekNewTxn + weekOldTxn;
    var weekAvgTicket   = weekOrders > 0 ? Math.round(weekRevenue / weekOrders) : 0;
    var _daysInMo       = new Date(today.getFullYear(), currentMonth, 0).getDate();
    var weekAchievement = monthTarget > 0 ? weekRevenue / (monthTarget * 7 / _daysInMo) : 0;
    var monthAchievement= monthTarget > 0 ? mtdRevenue  / monthTarget          : 0;
    var monthOrders     = allMtd.length;
    var monthAvgTicket  = monthOrders > 0 ? Math.round(mtdRevenue / monthOrders) : 0;

    var gaWeek  = fetchGA4_(weekRange.start, weekRange.end);
    var gaMonth = fetchGA4_(monthStart, today);
    var lyWeekStart  = new Date(weekRange.start.getTime()); lyWeekStart.setFullYear(lyWeekStart.getFullYear()-1);
    var lyWeekEnd    = new Date(weekRange.end.getTime());   lyWeekEnd.setFullYear(lyWeekEnd.getFullYear()-1);
    var lyMonthStart = new Date(monthStart.getTime());      lyMonthStart.setFullYear(lyMonthStart.getFullYear()-1);
    var lyToday      = new Date(today.getTime());           lyToday.setFullYear(lyToday.getFullYear()-1);
    var gaWeekLY  = fetchGA4_(lyWeekStart, lyWeekEnd);
    var gaMonthLY = fetchGA4_(lyMonthStart, lyToday);
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
      weekVisitorTarget: 0,
      weekNewVisitors:   0,
      weekOldVisitors:   0,
      weekRepairVisitors:0,
      weekRepairConv:    0,
      weekVipTxn:        0,
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
      weekSessionsLY:    gaWeekLY.sessions,
      weekUsersLY:       gaWeekLY.users,
      monthSessionsLY:   gaMonthLY.sessions,
      monthUsersLY:      gaMonthLY.users,
    };
  } catch(e) {
    Logger.log('Shopline ERROR: ' + e.toString());
    return { name: '官網', error: e.toString(), monthTarget: monthTarget };
  }
}

// UrlFetchApp + 指數退避重試（用於 Shopline 限速）
// 遇到 429/5xx/exception 自動退避：1s → 3s → 10s → 30s
function fetchWithBackoff_(url, options) {
  var delays = [1000, 3000, 10000, 30000];
  var lastErr = null;
  for (var attempt = 0; attempt <= delays.length; attempt++) {
    try {
      var resp = UrlFetchApp.fetch(url, options);
      var code = resp.getResponseCode();
      if (code === 200) return resp;
      // 429 限速 / 5xx 暫時錯誤 → 重試
      if (code === 429 || (code >= 500 && code < 600)) {
        var bodySnippet = (resp.getContentText() || '').slice(0, 200);
        Logger.log('fetchWithBackoff HTTP ' + code + ' (attempt ' + (attempt+1) + '): ' + bodySnippet);
        if (attempt < delays.length) { Utilities.sleep(delays[attempt]); continue; }
      }
      // 其他狀態碼直接回傳，由呼叫端處理
      return resp;
    } catch (e) {
      lastErr = e;
      var msg = e.toString();
      Logger.log('fetchWithBackoff exception (attempt ' + (attempt+1) + '): ' + msg);
      // 頻寬/限速類錯誤才重試；明顯非暫時性的就直接拋
      if (/bandwidth|quota|rate|timeout|temporarily|reset|unavailable/i.test(msg) && attempt < delays.length) {
        Utilities.sleep(delays[attempt]); continue;
      }
      throw e;
    }
  }
  throw lastErr || new Error('fetchWithBackoff: unknown failure');
}

// 加上 CacheService 快取，避免每日 UrlFetchApp 頻寬配額用盡
// 只快取下游真正需要的欄位：created_at / customer_id / total.dollars / status / created_from
// TTL：歷史月份 6 小時；當月 30 分鐘
function fetchShoplineOrders_(startDate, endDate) {
  var startMs  = startDate.getTime();
  var endMs    = endDate.getTime();
  var cache    = CacheService.getScriptCache();
  var cacheKey = 'shopline_v1_' + startMs + '_' + endMs;

  // 快取命中：拼接所有 chunk
  var cachedMeta = cache.get(cacheKey);
  if (cachedMeta) {
    try {
      var meta = JSON.parse(cachedMeta);
      var keys = [];
      for (var k = 0; k < meta.parts; k++) keys.push(cacheKey + '_p' + k);
      var chunks = cache.getAll(keys);
      var allOk = true, joined = '';
      for (var i = 0; i < keys.length; i++) {
        if (!chunks[keys[i]]) { allOk = false; break; }
        joined += chunks[keys[i]];
      }
      if (allOk) {
        Logger.log('Shopline cache HIT: ' + cacheKey + ' (' + meta.count + ' orders, ' + meta.parts + ' parts)');
        return JSON.parse(joined);
      }
    } catch (e) { Logger.log('Shopline cache parse error: ' + e); }
  }

  var minStr   = startDate.toISOString();
  var maxStr   = endDate.toISOString();
  var matched  = [], page = 1;
  while (page <= 50) {
    var url = 'https://open.shopline.io/v1/orders'
      + '?handle='           + CONFIG.SHOPLINE_HANDLE
      + '&per_page=100&page=' + page
      + '&created_at_min='   + encodeURIComponent(minStr)
      + '&created_at_max='   + encodeURIComponent(maxStr);
    var resp = fetchWithBackoff_(url, {
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
      if (t >= startMs && t <= endMs && orders[i].status !== 'cancelled') {
        matched.push({
          created_at:   orders[i].created_at,
          customer_id:  orders[i].customer_id,
          status:       orders[i].status,
          created_from: orders[i].created_from,
          total:        { dollars: (orders[i].total && orders[i].total.dollars != null) ? orders[i].total.dollars : 0 },
        });
      }
    }
    if (allBeforeStart) { Logger.log('Shopline early-exit at p' + page); break; }
    var perPage = (data.pagination || {}).per_page || 100;
    if (orders.length < perPage) break;
    page++;
    Utilities.sleep(300);  // 分頁間禮貌 sleep，降低 Shopline 限速風險
  }
  Logger.log('Shopline matched=' + matched.length + ' pages=' + (page - 1));

  // 寫入 chunked 快取（單一 key 上限約 100KB，分塊存放）
  try {
    var json = JSON.stringify(matched);
    var CHUNK = 90000; // 90KB safety margin
    var parts = [];
    for (var off = 0; off < json.length; off += CHUNK) parts.push(json.substr(off, CHUNK));
    if (parts.length === 0) parts.push('[]');
    var today    = new Date();
    var isCurMo  = (endDate.getFullYear() === today.getFullYear() && endDate.getMonth() === today.getMonth());
    var ttl      = isCurMo ? 1800 : 21600; // 30 min vs 6 hr
    var payload  = { meta: JSON.stringify({ count: matched.length, parts: parts.length }) };
    for (var p = 0; p < parts.length; p++) payload[cacheKey + '_p' + p] = parts[p];
    // 主 key 與 chunks 一起寫入（putAll 一次完成）
    var allMap = {};
    allMap[cacheKey] = payload.meta;
    for (var pp = 0; pp < parts.length; pp++) allMap[cacheKey + '_p' + pp] = parts[pp];
    cache.putAll(allMap, ttl);
    Logger.log('Shopline cache PUT: ' + cacheKey + ' (' + parts.length + ' parts, ttl=' + ttl + 's)');
  } catch (e) { Logger.log('Shopline cache put error: ' + e); }

  return matched;
}

// 手動清除 Shopline 快取（在編輯器執行可強制下次重新拉取）
function clearShoplineCache() {
  var today = new Date();
  var keysToRemove = [];
  for (var m = 0; m < 12; m++) {
    var monthStart = new Date(today.getFullYear(), m, 1).getTime();
    var monthEnd   = new Date(today.getFullYear(), m + 1, 0, 23, 59, 59, 999).getTime();
    var base = 'shopline_v1_' + monthStart + '_' + monthEnd;
    keysToRemove.push(base);
    for (var p = 0; p < 20; p++) keysToRemove.push(base + '_p' + p);
  }
  CacheService.getScriptCache().removeAll(keysToRemove);
  Logger.log('已清除 Shopline 月度快取（' + keysToRemove.length + ' 個 key）');
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
// GA4 Data API
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

function testGA4() {
  var today = new Date();
  var monthStart = new Date(today.getFullYear(), today.getMonth(), 1);
  var r = fetchGA4_(monthStart, today);
  Logger.log(JSON.stringify(r));
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
  // 週目標 = 月目標 × 7 ÷ 該月天數（取代固定 ÷4.3）
  var today        = new Date();
  var daysInMonth  = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();
  var weeksInMonth = daysInMonth / 7;
  tab.appendRow(['更新時間','週期','門市','週營收(元)','週目標(元)','週達成率',
    '總來客','新客來客','舊客來客','維修來客',
    '新客成交','回購成交','總成交','客單價','成交率',
    '月累計(元)','月目標(元)','月達成率','週訂單數(電商)','維修成交','新會員數',
    '週Sessions','週Users','週GA轉換率','月Sessions','月Users','月GA轉換率',
    '月訂單數','月客單價',
    '週Sessions去年','週Users去年','月Sessions去年','月Users去年',
    '來客目標','客單目標','轉換率目標','VIP成交']);
  storesData.forEach(function(s) {
    if (s.error) { tab.appendRow([now, weekLabel, s.name, 'ERROR: ' + s.error]); return; }
    var weekTgt = s.monthTarget > 0 ? Math.round(s.monthTarget / weeksInMonth) : 0;
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
      s.weekSessionsLY || 0, s.weekUsersLY   || 0,
      s.monthSessionsLY|| 0, s.monthUsersLY  || 0,
      s.weekVisitorTarget || 0, s.benchmarkTicket || 0, s.benchmarkConv || 0, s.weekVipTxn || 0,
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
        var visTgtTxt = store.weekVisitorTarget > 0 ? '/' + formatNumber(store.weekVisitorTarget) : '';
        visitorsText = '\n     👥 來客 *' + formatNumber(store.weekVisitors) + visTgtTxt + '* 人';
        if (store.weekAvgTicket > 0) {
          var ticketMark = (store.benchmarkTicket && store.weekAvgTicket < store.benchmarkTicket) ? ' ⚠️' : '';
          visitorsText += '　客單 *NT$' + formatNumber(store.weekAvgTicket) + '*' + ticketMark;
        }
        if (store.weekConvRate > 0 && store.benchmarkConv) {
          var convMark = store.weekConvRate < store.benchmarkConv ? ' ⚠️' : '';
          visitorsText += '　成交 *' + Math.round(store.weekConvRate * 100) + '%*' + convMark;
        }
      }
      msg += emoji + ' *' + store.name + '*　上週 *' + wPct + '%*（NT$' + formatNumber(store.weekRevenue) + '）／ ' + currentMonth + '月 *' + mPct + '%*' + visitorsText + '\n';
    }
  });

  // 四店合計：客單/成交率與「各店個別目標」比較
  var totalVisitors = 0, totalRevenue = 0, totalNewTxn = 0, totalOldTxn = 0, totalTxn = 0;
  var ticketAchSum = 0, convAchSum = 0, benchCount = 0;
  storesData.forEach(function(s) {
    if (!s.error && s.name !== '官網') {
      totalVisitors += (s.weekVisitors  || 0);
      totalRevenue  += (s.weekRevenue   || 0);
      totalNewTxn   += (s.weekNewTxn    || 0);
      totalOldTxn   += (s.weekOldTxn    || 0);
      totalTxn      += (s.weekTotalTxn  || 0);
      if (s.benchmarkTicket && s.weekAvgTicket) { ticketAchSum += s.weekAvgTicket / s.benchmarkTicket; benchCount++; }
      if (s.benchmarkConv   && s.weekConvRate)  { convAchSum   += s.weekConvRate   / s.benchmarkConv;            }
    }
  });

  var totalAvgTicket = totalTxn > 0 ? Math.round(totalRevenue / totalTxn) : 0;
  var totalTicketAchPct = benchCount > 0 ? Math.round(ticketAchSum / benchCount * 100) : 0;
  var totalConvRate  = totalVisitors > 0 ? Math.round(totalTxn / totalVisitors * 100) : 0;
  var totalConvAchPct = benchCount > 0 ? Math.round(convAchSum / benchCount * 100) : 0;

  var newOldText = '';
  if (totalTxn > 0) {
    var newRatio = Math.round(totalNewTxn / totalTxn * 10);
    var oldRatio = 10 - newRatio;
    newOldText = '（新舊客 ' + newRatio + ':' + oldRatio + '）';
  }

  var companyPct = Math.round(companyMPct * 100);
  msg += '━━━━━━━━━━━━━━━━━━━━\n';
  if (totalVisitors > 0) msg += '👥 *四店來客：' + formatNumber(totalVisitors) + ' 人' + newOldText + '*\n';
  if (totalAvgTicket > 0) msg += '🧾 *四店平均客單：NT$' + formatNumber(totalAvgTicket) + '（達成各店目標 ' + totalTicketAchPct + '%）*\n';
  if (totalConvRate > 0)  msg += '🎯 *四店平均成交率：' + totalConvRate + '%（達成各店目標 ' + totalConvAchPct + '%）*\n';
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

function fixShoplineJan() {
  var year  = 2026, month = 1;
  var start = new Date(year, 0, 1);
  var end   = new Date(year, 1, 0, 23, 59, 59, 999);
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
  tab.appendRow([year, month, '官網', Math.round(rev), tgt, tgt > 0 ? rev / tgt : 0, '手動修正']);
}

// ==========================================
// 歷史資料回填（官網月度營收，不含 GA4）
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
    Utilities.sleep(3000);
  }
  Logger.log('🎉 全部回填完成');
}
