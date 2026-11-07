const APP_CONFIG = {
  CHURCH_CODE: 'GG',
  CHURCH_NAME: 'Garden Grove',
  CHURCH_NAME_ZH: '園林教會',
  CONTACT_EMAIL: 'gg.info@tjc.org',
  CC_EMAIL: 'garden_grove@tjc.org',
  ADMIN_ALERT_EMAIL: 'gg.info@tjc.org',
  TIMEZONE: 'America/Los_Angeles',
  HOLY_WORK_SCHEDULES_FOLDER_ID: '1sH-ELkxwRLB78ts2rJeTpPmYtDY8KfzE',
  LEGACY_2026_FOLDER_ID: '1A2B8lmTT0qBWHT4TFzXG4RRGIWRhzagP',
  DEFAULT_SERMON_TOPICS_ID: '17qRYYuZTV3PWf92CnMmqwARaC2bk4BUPxz8DQ_9f3hU',
  DEFAULT_UNAVAILABILITY_ID: '1Ge3iKjPUIkJCuAdlRDqhgEWt4DlO7dL3mFyl0TVLuaY',
  ESSC_SCHEDULE_ID: '1YupXj1wjmjTefQCPArSAJKNy1aF95CIxMxUV6JeFi3w',
  TJC_LOGO_FILE_ID: '15-OcvnePRkYVBg5yz63ILCcHxvVYLQIr',
  THEME: {
    PRIMARY_COLOR: '#1F4E79',     // Main header blue
    ACCENT_SUCCESS: '#d9ead3',    // Light green highlight
    ACCENT_WARNING: '#fff2cc',    // Light yellow highlight
    ACCENT_DANGER:  '#f4cccc'     // Light red alert
  },
  DEFAULT_TEAMS: ['Audio/Visual', 'Ushers', 'Interpreters', 'Piano', 'Cleaning'],
  AUTHORIZED_EDITORS: [
    "david.liou@tjc.org",
    "anthony.tsai@tjc.org",
    "gg.info@tjc.org",
    "jessica.chen@tjc.org",
    "anthony.chang@tjc.org",
    "jessica.jeh.chen@gmail.com"
  ],
  BIBLE_PLAN_URL: "https://docs.google.com/spreadsheets/d/1Omlqzklhiw7I1Wyp-5tgxuVZQdUlszZEufPh5O5zCpM/edit"
};

/**
 * Structured logger compatible with Google Cloud Logging and Apps Script Execution Log.
 */
var AppLogger = {
  info: function(fnName, message, meta) {
    var payload = { severity: 'INFO', function: fnName, message: message, meta: meta || {} };
    if (typeof console !== 'undefined' && console.info) console.info(JSON.stringify(payload));
    else Logger.log('[INFO][' + fnName + '] ' + message);
  },
  warn: function(fnName, message, meta) {
    var payload = { severity: 'WARNING', function: fnName, message: message, meta: meta || {} };
    if (typeof console !== 'undefined' && console.warn) console.warn(JSON.stringify(payload));
    else Logger.log('[WARN][' + fnName + '] ' + message);
  },
  error: function(fnName, err, meta) {
    var errMsg = (err && err.message) ? err.message : String(err);
    var payload = {
      severity: 'ERROR',
      function: fnName,
      error: errMsg,
      stack: (err && err.stack) ? err.stack : null,
      meta: meta || {}
    };
    if (typeof console !== 'undefined' && console.error) console.error(JSON.stringify(payload));
    else Logger.log('[ERROR][' + fnName + '] ' + errMsg);
  }
};

/**
 * Parses an A1 cell reference (e.g., "E12", "AY6") into 1-based { row, col } integers in memory
 * without invoking any SpreadsheetApp range RPC calls.
 *
 * @param {string} a1 Cell A1 notation
 * @return {{row: number, col: number}|null}
 */
function parseA1Coords_(a1) {
  var m = /^([A-Z]+)(\d+)$/i.exec(String(a1 || '').trim());
  if (!m) return null;
  var letters = m[1].toUpperCase();
  var col = 0;
  for (var i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return { row: parseInt(m[2], 10), col: col };
}

/**
 * Concurrency helper: executes an operation with LockService, exponential backoff for transient
 * Google API errors, and guaranteed lock release in a finally block.
 *
 * @param {string} actionName Descriptive name of the action for logging
 * @param {Function} fn Function to execute while lock is held
 * @param {Object} [options]
 * @param {string} [options.lockType='script'] 'script', 'document', or 'user'
 * @param {number} [options.timeoutMs=30000] Milliseconds to wait to acquire lock
 * @param {number} [options.retries=2] Number of retries on lock timeout or transient error
 * @param {number} [options.retryDelayMs=2000] Base delay between retries
 * @param {boolean} [options.flushBeforeRelease=true] Flush pending SpreadsheetApp writes before releasing lock
 * @return {*} Result of fn()
 */
function executeWithLockAndRetry_(actionName, fn, options) {
  options = options || {};
  var lockType = options.lockType || 'script';
  var timeoutMs = options.timeoutMs || 30000;
  var retries = options.retries !== undefined ? options.retries : 2;
  var retryDelayMs = options.retryDelayMs || 2000;
  var flushBeforeRelease = options.flushBeforeRelease !== false;

  var lock;
  if (lockType === 'document') {
    lock = LockService.getDocumentLock();
  } else if (lockType === 'user') {
    lock = LockService.getUserLock();
  } else {
    lock = LockService.getScriptLock();
  }

  var attempt = 0;
  while (attempt <= retries) {
    var hasLock = false;
    try {
      hasLock = lock.tryLock(timeoutMs);
    } catch (e) {
      AppLogger.warn(actionName, 'tryLock threw: ' + e);
    }

    if (hasLock) {
      try {
        var apiAttempt = 0;
        while (true) {
          try {
            var res = fn();
            if (flushBeforeRelease) {
              try { SpreadsheetApp.flush(); } catch (flushErr) {}
            }
            return res;
          } catch (apiErr) {
            var msg = String(apiErr && apiErr.message ? apiErr.message : apiErr);
            var isTransient = /Service invoked too many times|Timed out|Internal error|try again|Simultaneous invocations/i.test(msg);
            if (!isTransient || apiAttempt >= retries) {
              AppLogger.error(actionName, apiErr, { apiAttempt: apiAttempt });
              throw apiErr;
            }
            apiAttempt++;
            var backoff = Math.pow(2, apiAttempt) * 500 + Math.floor(Math.random() * 250);
            AppLogger.warn(actionName, 'Transient API error; retrying in ' + backoff + 'ms', { apiAttempt: apiAttempt, error: msg });
            Utilities.sleep(backoff);
          }
        }
      } finally {
        try {
          lock.releaseLock();
        } catch (releaseErr) {
          AppLogger.warn(actionName, 'Error releasing lock: ' + releaseErr);
        }
      }
    }

    attempt++;
    if (attempt <= retries) {
      AppLogger.warn(actionName, 'Lock contention, attempt ' + attempt + ' failed. Waiting ' + retryDelayMs + 'ms...');
      Utilities.sleep(retryDelayMs);
    }
  }

  var timeoutErr = new Error('Lock timeout: Unable to acquire ' + lockType + ' lock for "' + actionName + '" after ' + (retries + 1) + ' attempts.');
  AppLogger.error(actionName, timeoutErr);
  throw timeoutErr;
}


