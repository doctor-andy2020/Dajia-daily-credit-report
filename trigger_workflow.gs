// ============================================================
// 大家资产每日舆情日报 — 兜底触发器
// 部署：Google Apps Script (script.google.com)
// 北京时间 09:00-10:00 触发，作为 GitHub cron(07:30) 的兜底
// 先查询当天是否已有成功 run，有则跳过，避免重复推送
// GitHub Actions workflow 内也有去重作为第二道防线
// GAS 控制台：https://script.google.com/
// ============================================================

function triggerWorkflow() {
  // ⚠️ Token 存储在 GAS Script Properties 中（文件 > 项目属性 > 脚本属性）
  // 添加属性：GITHUB_TOKEN = <your-github-personal-access-token>
  // 切勿将 token 硬编码在代码中
  var token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) {
    Logger.log('❌ GITHUB_TOKEN not found in Script Properties. Set it via File > Project Properties > Script Properties.');
    return 'no_token';
  }
  var headers = {
    'Authorization': 'Bearer ' + token,
    'Accept': 'application/vnd.github+json',
    'Content-Type': 'application/json',
  };

  // Step 1: 检查今天是否已有成功 run
  var runsUrl = 'https://api.github.com/repos/doctor-andy2020/Dajia-daily-credit-report/actions/runs?branch=master&status=success&per_page=10';
  var runsResp = UrlFetchApp.fetch(runsUrl, { headers: headers, muteHttpExceptions: true });

  if (runsResp.getResponseCode() === 200) {
    var data = JSON.parse(runsResp.getContentText());
    var runs = data.workflow_runs || [];
    var today = Utilities.formatDate(new Date(), 'Asia/Shanghai', 'yyyy-MM-dd');
    for (var i = 0; i < runs.length; i++) {
      var createdAt = runs[i].created_at;
      if (createdAt && createdAt.indexOf(today) === 0) {
        Logger.log('⏭  Skipped: today already has a successful run at ' + createdAt);
        return 'skipped';
      }
    }
  } else {
    Logger.log('⚠ Check runs failed (status ' + runsResp.getResponseCode() + '), dispatching anyway');
  }

  // Step 2: 调度 workflow
  var dispatchUrl = 'https://api.github.com/repos/doctor-andy2020/Dajia-daily-credit-report/actions/workflows/292559747/dispatches';
  var options = {
    'method': 'POST',
    'headers': headers,
    'payload': JSON.stringify({ ref: 'master' }),
    'muteHttpExceptions': true,
  };

  var response = UrlFetchApp.fetch(dispatchUrl, options);
  var status = response.getResponseCode();

  Logger.log('Trigger time: ' + new Date().toLocaleString('zh-CN', { timeZone: 'Asia/Shanghai' }));
  Logger.log('HTTP Status: ' + status);

  if (status === 204) {
    Logger.log('✅ Workflow triggered successfully');
  } else {
    Logger.log('❌ Failed: ' + response.getContentText());
  }

  return status;
}
