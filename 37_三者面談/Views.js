/**
 * 表示（全体ビュー・クラス別予約表）の更新を、予約の確定から切り離す。
 *
 * これまでは予約が1件入るたびに全クラスの表を作り直していた。
 * 1学年4クラスなら許容できたが、全校17クラスでは1件あたり70秒を超え、
 * 排他ロックのぶん、そのまま保護者の待ち行列になってしまう。
 *
 * そこで予約時は「表が古くなった」印だけを付け、
 * 実際の作り直しは数分おきのタイマーか、先生が見るときに行う。
 * 保護者の予約可否は「枠マスタ」だけで決まるので、表示が遅れても取り違えは起きない。
 */

var VIEW_STALE_KEY = 'views_stale_count';
var VIEW_UPDATED_KEY = 'views_updated_at';
var VIEW_TRIGGER_FN = 'refreshViewsIfStale';

/** 表が古くなったことを記録する（予約・変更・取消のたびに呼ぶ） */
function markViewsStale_() {
  try {
    var props = PropertiesService.getScriptProperties();
    var n = Number(props.getProperty(VIEW_STALE_KEY) || 0) + 1;
    props.setProperty(VIEW_STALE_KEY, String(n));
  } catch (e) {
    console.warn('表示の更新予約に失敗:', e);
  }
}

/** 未反映の件数 */
function pendingViewUpdates_() {
  try {
    return Number(PropertiesService.getScriptProperties().getProperty(VIEW_STALE_KEY) || 0);
  } catch (e) {
    return 0;
  }
}

/** 最後に表を作り直した時刻。まだ無ければ null。 */
function viewsUpdatedAt_() {
  try {
    var v = PropertiesService.getScriptProperties().getProperty(VIEW_UPDATED_KEY);
    return v ? new Date(Number(v)) : null;
  } catch (e) {
    return null;
  }
}

/**
 * 表を作り直す。
 * @param {boolean} force 未反映が無くても作り直すか
 * @return {{refreshed:boolean, pending:number}}
 */
function refreshViews(force) {
  var pending = pendingViewUpdates_();
  if (!force && !pending) {
    return { refreshed: false, pending: 0 };
  }

  // 作り直している間に入った予約を取りこぼさないよう、先に印を消す
  try {
    PropertiesService.getScriptProperties().deleteProperty(VIEW_STALE_KEY);
  } catch (e) { /* 無視 */ }

  rebuildOverview();
  rebuildClassSheets();
  hideInternalSheets();

  try {
    PropertiesService.getScriptProperties()
      .setProperty(VIEW_UPDATED_KEY, String(new Date().getTime()));
  } catch (e) { /* 無視 */ }

  return { refreshed: true, pending: pending };
}

/** タイマーから呼ばれる。溜まっていれば作り直す。 */
function refreshViewsIfStale() {
  try {
    refreshViews(false);
  } catch (err) {
    console.error('表示の自動更新に失敗:', err);
  }
}

/* ---------------- 自動更新のタイマー ---------------- */

function viewTriggerMinutes_() {
  var t = ScriptApp.getProjectTriggers();
  for (var i = 0; i < t.length; i++) {
    if (t[i].getHandlerFunction() === VIEW_TRIGGER_FN) return true;
  }
  return false;
}

function autoRefreshEnabled_() {
  var on = countTriggers_(VIEW_TRIGGER_FN) > 0;
  syncAutoFlag_(VIEW_AUTO_KEY, on);
  return on;
}

/**
 * 表の自動更新を有効にする。
 * @param {number} minutes 1・5・10・15・30 のいずれか
 */
function enableAutoRefresh(minutes) {
  disableAutoRefresh();
  var m = [1, 5, 10, 15, 30].indexOf(Number(minutes)) >= 0 ? Number(minutes) : 5;
  ScriptApp.newTrigger(VIEW_TRIGGER_FN).timeBased().everyMinutes(m).create();
  syncAutoFlag_(VIEW_AUTO_KEY, true);
  return m;
}

function disableAutoRefresh() {
  var n = deleteTriggers_(VIEW_TRIGGER_FN);
  syncAutoFlag_(VIEW_AUTO_KEY, false);
  return n;
}
