const APP_CONFIG = {
  CHURCH_CODE: 'GG',
  CHURCH_NAME: 'Garden Grove',
  CHURCH_NAME_ZH: '園林教會',
  CONTACT_EMAIL: 'gg.info@tjc.org',
  CC_EMAIL: 'garden_grove@tjc.org',
  ADMIN_ALERT_EMAIL: 'gg.info@tjc.org',
  TIMEZONE: 'America/Los_Angeles',
  HOLY_WORK_SCHEDULES_FOLDER_ID: '1A2B8lmTT0qBWHT4TFzXG4RRGIWRhzagP',
  DEFAULT_SERMON_TOPICS_ID: '1zrMaPubQ7uaUvgBjYZ3lWIBRUclg0GsGliHMqpvOark',
  DEFAULT_UNAVAILABILITY_ID: '',
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
 * Concurrency helper: executes an operation with LockService and optional retries.
 * Ensures locks are always released in a finally block.
 *
 * @param {string} actionName Descriptive name of the action for logging
 * @param {Function} fn Function to execute while lock is held
 * @param {Object} [options]
 * @param {string} [options.lockType='script'] 'script', 'document', or 'user'
 * @param {number} [options.timeoutMs=30000] Milliseconds to wait to acquire lock
 * @param {number} [options.retries=2] Number of retries on lock timeout
 * @param {number} [options.retryDelayMs=2000] Delay between retries
 * @return {*} Result of fn()
 */
function executeWithLockAndRetry_(actionName, fn, options) {
  options = options || {};
  var lockType = options.lockType || 'script';
  var timeoutMs = options.timeoutMs || 30000;
  var retries = options.retries !== undefined ? options.retries : 2;
  var retryDelayMs = options.retryDelayMs || 2000;

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
      Logger.log('[executeWithLockAndRetry_] tryLock threw: ' + e);
    }

    if (hasLock) {
      try {
        return fn();
      } finally {
        try {
          lock.releaseLock();
        } catch (releaseErr) {
          Logger.log('[executeWithLockAndRetry_] Error releasing lock: ' + releaseErr);
        }
      }
    }

    attempt++;
    if (attempt <= retries) {
      Logger.log('[executeWithLockAndRetry_] Lock contention for "' + actionName + '", attempt ' + attempt + ' failed. Waiting ' + retryDelayMs + 'ms...');
      Utilities.sleep(retryDelayMs);
    }
  }

  throw new Error('Lock timeout: Unable to acquire ' + lockType + ' lock for "' + actionName + '" after ' + (retries + 1) + ' attempts.');
}

