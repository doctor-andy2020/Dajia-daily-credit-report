// ============================================================
// 大家资产每日舆情报告 — 兜底触发器（双工作流）
// 部署：Google Apps Script (script.google.com)
//
// 触发器时间设置（GAS 编辑器 → 触发器 → 添加触发器）：
//   1. triggerWorkflow    — 时间驱动 → 日定时器 → 上午 8-9点  → 舆情日报兜底
//   2. triggerDMWorkflow  — 时间驱动 → 日定时器 → 上午 8-9点  → DM早报兜底
//
// 工作原理：
//   先查询当天是否已有成功 run，有则跳过，避免重复推送
//   GitHub Actions workflow 内也有去重作为第二道防线
//
// GAS 控制台：https://script.google.com/
// ============================================================

// Token 读取优先级：
//   1. GAS Script Properties（推荐）→ 项目设置 ⚙️ → 脚本属性 → GITHUB_TOKEN
//   2. 全局变量（兜底）→ 新建 config.gs 文件，内容：var GITHUB_TOKEN = 'ghp_xxx';
// config.gs 不要提交到 GitHub（已在 .gitignore 中忽略）

// ── 舆情日报 Workflow ID（daily-report.yml）──
var DAILY_REPORT_WORKFLOW_ID = '292559747';

// ── DM早报 Workflow ID（dm-morning-report.yml）──
// ⚠️ 部署步骤：push dm-morning-report.yml 到 GitHub 后，
//    运行以下命令获取新 workflow ID，替换下面的值：
//    gh api repos/doctor-andy2020/Dajia-daily-credit-report/actions/workflows --jq '.workflows[] | select(.name=="DM Credit Morning Report") | .id'
var DM_MORNING_WORKFLOW_ID = '301380840';


function _getToken() {
  var token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) {
    try { token = GITHUB_TOKEN; } catch(e) {}
  }
  return token;
}

function _checkExistingRun(token, workflowId) {
  // 检查今天是否已有成功/进行中/排队的 run
  //   - success: 已跑完，无需再触发
  //   - in_progress / queued: 正在跑或等待中，再触发会造成竞态重复
  var statuses = ['success', 'in_progress', 'queued'];
  var today = Utilities.formatDate(new Date(), 'Asia/Shanghai', 'yyyy-MM-dd');
  var headers = {
    'Authorization': 'Bearer ' + token,
    'Accept': 'application/vnd.github+json',
    'Content-Type': 'application/json',
  };

  for (var s = 0; s < statuses.length; s++) {
    var runsUrl = 'https://api.github.com/repos/doctor-andy2020/Dajia-daily-credit-report/actions/runs'
      + '?branch=master&status=' + statuses[s] + '&per_page=10'
      + '&workflow_id=' + workflowId;

    var runsResp = UrlFetchApp.fetch(runsUrl, { headers: headers, muteHttpExceptions: true });
    if (runsResp.getResponseCode() === 200) {
      var data = JSON.parse(runsResp.getContentText());
      var runs = data.workflow_runs || [];
      for (var i = 0; i < runs.length; i++) {
        var createdAt = runs[i].created_at;
        if (createdAt && createdAt.indexOf(today) === 0) {
          return { skip: true, createdAt: createdAt, status: statuses[s] };
        }
      }
    }
  }
  return { skip: false };
}

function _dispatchWorkflow(token, workflowId) {
  var dispatchUrl = 'https://api.github.com/repos/doctor-andy2020/Dajia-daily-credit-report/actions/workflows/'
    + workflowId + '/dispatches';
  var headers = {
    'Authorization': 'Bearer ' + token,
    'Accept': 'application/vnd.github+json',
    'Content-Type': 'application/json',
  };
  var options = {
    'method': 'POST',
    'headers': headers,
    'payload': JSON.stringify({ ref: 'master' }),
    'muteHttpExceptions': true,
  };
  var response = UrlFetchApp.fetch(dispatchUrl, options);
  return response.getResponseCode();
}


// ============================================================
// 舆情日报触发器（北京时间 8:00-9:00）
// ============================================================
function triggerWorkflow() {
  var token = _getToken();
  if (!token) {
    Logger.log('❌ GITHUB_TOKEN not found.');
    return 'no_token';
  }

  // 检查今天是否已有成功 run
  var check = _checkExistingRun(token, DAILY_REPORT_WORKFLOW_ID);
  if (check.skip) {
    Logger.log('⏭  [舆情日报] Skipped: today already has a ' + (check.status || 'success') + ' run at ' + check.createdAt);
    return 'skipped';
  }

  // 调度 workflow
  var status = _dispatchWorkflow(token, DAILY_REPORT_WORKFLOW_ID);

  Logger.log('[舆情日报] Trigger time: ' + new Date().toLocaleString('zh-CN', { timeZone: 'Asia/Shanghai' }));
  Logger.log('[舆情日报] HTTP Status: ' + status);

  if (status === 204) {
    Logger.log('✅ [舆情日报] Workflow triggered successfully');
  } else {
    Logger.log('❌ [舆情日报] Failed (status: ' + status + ')');
  }

  return status;
}


// ============================================================
// DM早报触发器（北京时间 8:00-9:00）
// ============================================================
function triggerDMWorkflow() {
  var token = _getToken();
  if (!token) {
    Logger.log('❌ GITHUB_TOKEN not found.');
    return 'no_token';
  }

  // 检查 workflow ID 是否已配置
  if (DM_MORNING_WORKFLOW_ID === 'PLACEHOLDER_GET_FROM_GITHUB') {
    Logger.log('⚠  [DM早报] Workflow ID not configured yet. Run the gh CLI command to get it.');
    return 'not_configured';
  }

  // 去重：DM早报实际由 daily-report.yml 的 dm-morning job 承担，
  // 所以检查 daily-report.yml 当天是否已成功（成功=舆情+DM两个job都过=DM已发）。
  // 若 daily-report.yml 失败（含 DM job 失败），这里查不到成功 → 正确触发兜底补发。
  var checkPrimary = _checkExistingRun(token, DAILY_REPORT_WORKFLOW_ID);
  if (checkPrimary.skip) {
    Logger.log('⏭  [DM早报] Skipped: daily-report.yml 今日已成功（DM已随主力发送）at ' + checkPrimary.createdAt);
    return 'skipped';
  }
  // 二次去重：避免本兜底 workflow 自身重复触发
  var checkSelf = _checkExistingRun(token, DM_MORNING_WORKFLOW_ID);
  if (checkSelf.skip) {
    Logger.log('⏭  [DM早报] Skipped: 兜底 workflow 今日已成功 at ' + checkSelf.createdAt);
    return 'skipped';
  }

  // 调度 workflow
  var status = _dispatchWorkflow(token, DM_MORNING_WORKFLOW_ID);

  Logger.log('[DM早报] Trigger time: ' + new Date().toLocaleString('zh-CN', { timeZone: 'Asia/Shanghai' }));
  Logger.log('[DM早报] HTTP Status: ' + status);

  if (status === 204) {
    Logger.log('✅ [DM早报] Workflow triggered successfully');
  } else {
    Logger.log('❌ [DM早报] Failed (status: ' + status + ')');
  }

  return status;
}
