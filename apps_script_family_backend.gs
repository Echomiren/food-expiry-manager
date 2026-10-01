const SHEETS = {
  USERS: 'Users',
  FAMILIES: 'Families',
  MEMBERS: 'FamilyMembers',
  FOODS: 'Foods',
  SESSIONS: 'Sessions',
  RESETS: 'PasswordResets',
  ACTIVITY: 'ActivityLog',
  REQUESTS: 'RequestLog',
  NOTIFY_USER: 'NotificationUserSettings',
  NOTIFY_GROUP: 'NotificationGroupSettings',
  NOTIFY_LOG: 'NotificationLog',
  FOOD_IMAGES: 'FoodImages',
  IMAGE_BACKUPS: 'ImageBackupLog'
};

const SESSION_DAYS = 30;
const SESSION_CACHE_SECONDS = 21600;
const HASH_ROUNDS = 1500; // 舊帳號相容用；新帳號不再使用這個慢迴圈
const PASSWORD_HASH_VERSION = 'hmac-v2';
const RESET_CODE_MINUTES = 10;
const RESET_CODE_COOLDOWN_SECONDS = 60;
const RESET_MAX_ATTEMPTS = 5;

// v4.4 Drive 圖片 / 備份與清理
const BACKUP_RETENTION_DAYS = 93;       // 約 3 個月
const SESSION_CLEANUP_GRACE_DAYS = 14; // Session 到期後再保留 14 天
const RESET_CLEANUP_DAYS = 7;
const REQUEST_LOG_RETENTION_DAYS = 30;
const NOTIFY_LOG_RETENTION_DAYS = 90;
const ACTIVITY_LOG_RETENTION_DAYS = 90;
const PENDING_IMAGE_RETENTION_DAYS = 14;
const DELETED_IMAGE_META_RETENTION_DAYS = 35;
const IMAGE_BACKUP_RETENTION_DAYS = 93;
const MAX_PRODUCT_IMAGE_BYTES = 450000; // 前端正常約 50~150 KB，後端再設硬上限

// 單次 Apps Script 執行上限通常為 6 分鐘。圖片備份主動在 4 分鐘內收尾，
// 並限制每批數量，保留時間給狀態寫入與建立下一次續跑 Trigger。
const IMAGE_BACKUP_BATCH_TIME_BUDGET_MS = 4*60*1000;
const IMAGE_BACKUP_MAX_ITEMS_PER_BATCH = 50;
const IMAGE_BACKUP_CONTINUATION_DELAY_MS = 2*60*1000;
const IMAGE_BACKUP_MAX_FAILURE_RETRIES = 3;
const IMAGE_BACKUP_RUN_STATE_KEY = 'IMAGE_BACKUP_RUN_STATE_V1';
const IMAGE_BACKUP_CONTINUATION_HANDLER =
  'continueWeeklyBackupAndMaintenance';

const NOTIFICATION_LOG_HEADERS = [
  'userId','familyId','foodId','expiry','daysBefore','sentAt','foodName','familyName'
];

const IMAGE_BACKUP_LOG_HEADERS = [
  'imageId','familyId','foodId','sourceDriveFileId','backupFileId',
  'mimeType','sizeBytes','sourceStatus','backedUpAt','updatedAt',
  'deletedAt','lastError'
];

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureSheet_(ss, SHEETS.USERS,
    ['userId','email','displayName','passwordSalt','passwordHash','status','createdAt','hashVersion']);

  ensureHeaders_(ss.getSheetByName(SHEETS.USERS),
    ['userId','email','displayName','passwordSalt','passwordHash','status','createdAt','hashVersion']);

  ensureSheet_(ss, SHEETS.FAMILIES,
    ['familyId','familyName','ownerUserId','inviteCode','createdAt']);

  ensureSheet_(ss, SHEETS.MEMBERS,
    ['familyId','userId','role','status','joinedAt']);

  ensureSheet_(ss, SHEETS.FOODS,
    ['familyId','id','name','qty','location','expiry','note','createdBy','createAt','updatedAt','notifyMode','notifyDaysBefore','imageId','status','consumedAt']);

  ensureHeaders_(ss.getSheetByName(SHEETS.FOODS),
    ['familyId','id','name','qty','location','expiry','note','createdBy','createAt','updatedAt','notifyMode','notifyDaysBefore','imageId','status','consumedAt']);

  ensureSheet_(ss, SHEETS.SESSIONS,
    ['sessionToken','userId','expiresAt','createdAt']);

  ensureSheet_(ss, SHEETS.RESETS,
    ['resetId','userId','email','codeHash','expiresAt','attempts','usedAt','createdAt']);

  ensureSheet_(ss, SHEETS.ACTIVITY,
    ['familyId','userId','action','targetId','detail','createdAt']);

  ensureSheet_(ss, SHEETS.REQUESTS,
    ['requestId','userId','action','status','resultJson','createdAt','updatedAt']);

  ensureSheet_(ss, SHEETS.NOTIFY_USER,
    ['userId','emailEnabled','sendHour','updatedAt']);

  ensureSheet_(ss, SHEETS.NOTIFY_GROUP,
    ['userId','familyId','enabled','defaultDaysBefore','updatedAt']);

  ensureSheet_(ss, SHEETS.NOTIFY_LOG,
    NOTIFICATION_LOG_HEADERS);

  ensureHeaders_(
    ss.getSheetByName(SHEETS.NOTIFY_LOG),
    NOTIFICATION_LOG_HEADERS
  );

  ensureSheet_(ss, SHEETS.FOOD_IMAGES,
    ['imageId','familyId','foodId','driveFileId','mimeType','sizeBytes','width','height','status','uploadedBy','createdAt','updatedAt','deletedAt']);

  ensureSheet_(ss, SHEETS.IMAGE_BACKUPS,
    IMAGE_BACKUP_LOG_HEADERS);

  ensureHeaders_(
    ss.getSheetByName(SHEETS.IMAGE_BACKUPS),
    IMAGE_BACKUP_LOG_HEADERS
  );

  const props = PropertiesService.getScriptProperties();

  // 供背景 Trigger 正確找到這份資料庫試算表。
  props.setProperty('DATABASE_SPREADSHEET_ID', ss.getId());

  if (!props.getProperty('AUTH_PEPPER')) {
    props.setProperty('AUTH_PEPPER', Utilities.getUuid() + Utilities.getUuid());
  }

  if (!props.getProperty('RESET_PEPPER')) {
    props.setProperty('RESET_PEPPER', Utilities.getUuid() + Utilities.getUuid());
  }
}

function doGet() {
  return json_({ ok:true, service:'Food Expiry Family API v4.4' });
}

function doPost(e) {
  let action='';

  try {
    const p = e.parameter || {};
    action = p.action || '';

    // 只記錄 API 動作名稱，不記 sessionToken、密碼或其他敏感資料。
    console.log('API action: ' + action);

    if (action === 'register') return json_(register_(p));
    if (action === 'login') return json_(login_(p));
    if (action === 'requestPasswordReset') return json_(requestPasswordReset_(p));
    if (action === 'resetPassword') return json_(resetPassword_(p));

    const user = requireSession_(p.sessionToken);

    if (action === 'logout') return json_(logout_(p.sessionToken));
    if (action === 'me') return json_(me_(user));
    if (action === 'getNotificationSettings') return json_(getNotificationSettings_(user));
    if (action === 'saveNotificationSettings') return json_(saveNotificationSettings_(user,p));

    if (action === 'adminSummary') return json_(adminSummary_(user,p));
    if (action === 'adminLogs') return json_(adminLogs_(user,p));
    if (action === 'adminListAdmins') return json_(adminListAdmins_(user,p));
    if (action === 'adminAddAdmin') return json_(adminAddAdmin_(user,p));
    if (action === 'adminRemoveAdmin') return json_(adminRemoveAdmin_(user,p));

    if (action === 'uploadFoodImage') return json_(uploadFoodImage_(user,p));
    if (action === 'discardFoodImage') return json_(discardFoodImage_(user,p));
    if (action === 'getFoodImage') return json_(getFoodImage_(user,p));

    if (action === 'createFamily') return json_(createFamily_(user,p));
    if (action === 'joinFamily') return json_(joinFamily_(user,p));
    if (action === 'checkRequest') return json_(checkRequest_(user,p));
    if (action === 'familyDetails') return json_(familyDetails_(user,p));
    if (action === 'renameFamily') return json_(renameFamily_(user,p));
    if (action === 'regenerateInviteCode') return json_(regenerateInviteCode_(user,p));
    if (action === 'setMemberRole') return json_(setMemberRole_(user,p));
    if (action === 'removeMember') return json_(removeMember_(user,p));
    if (action === 'leaveFamily') return json_(leaveFamily_(user,p));
    if (action === 'deleteFamily') return json_(deleteFamily_(user,p));

    if (action === 'listFoods') return json_(listFoods_(user,p));
    if (action === 'syncChanges') return json_(syncChanges_(user,p));

    throw new Error('unknown action');
  } catch (err) {
    console.error(
      'API error (' + (action || 'unknown') + '): ' +
      String(err.message || err)
    );
    return json_({ ok:false, error:String(err.message || err) });
  }
}

// ---------------- AUTH ----------------

function register_(p) {
  const email = String(p.email || '').trim().toLowerCase();
  const displayName = String(p.displayName || '').trim();
  const password = String(p.password || '');

  if (!email || !email.includes('@')) throw new Error('Email 格式不正確');
  if (!displayName) throw new Error('請輸入顯示名稱');
  if (password.length < 8) throw new Error('密碼至少 8 碼');

  const users = rows_(SHEETS.USERS);
  if (users.some(x => String(x.email).toLowerCase() === email)) {
    throw new Error('這個 Email 已經註冊');
  }

  const userId = Utilities.getUuid();
  const salt = Utilities.getUuid();
  const hash = passwordHashFast_(password,salt);

  append_(SHEETS.USERS,[
    userId,email,displayName,salt,hash,'active',isoNow_(),PASSWORD_HASH_VERSION
  ]);
  return { ok:true };
}

function login_(p) {
  const email = String(p.email || '').trim().toLowerCase();
  const password = String(p.password || '');

  const user = rows_(SHEETS.USERS).find(x =>
    String(x.email).toLowerCase() === email && x.status === 'active'
  );

  if (!user) {
    throw new Error('Email 或密碼錯誤');
  }

  const isFast = String(user.hashVersion || '') === PASSWORD_HASH_VERSION;
  let valid = false;

  if (isFast) {
    valid = passwordHashFast_(password,user.passwordSalt) === user.passwordHash;
  } else {
    // 舊帳號只在第一次登入時跑舊的 1500 次 SHA-256。
    valid = passwordHashLegacy_(password,user.passwordSalt) === user.passwordHash;

    if (valid) {
      // 驗證成功立即升級，之後登入就走快速版本。
      const newSalt = Utilities.getUuid();
      const newHash = passwordHashFast_(password,newSalt);
      updateUserPasswordVersion_(user.userId,newSalt,newHash,PASSWORD_HASH_VERSION);

      user.passwordSalt = newSalt;
      user.passwordHash = newHash;
      user.hashVersion = PASSWORD_HASH_VERSION;
    }
  }

  if (!valid) {
    throw new Error('Email 或密碼錯誤');
  }

  const token = Utilities.getUuid() + Utilities.getUuid();
  const expires = new Date(Date.now() + SESSION_DAYS * 86400000);

  append_(SHEETS.SESSIONS,[token,user.userId,expires.toISOString(),isoNow_()]);
  cacheSession_(token,user,expires.getTime());

  return {
    ok:true,
    sessionToken:token,
    user:{...publicUser_(user),isAdmin:isAdmin_(user)},
    families:userFamilies_(user.userId)
  };
}

function requestPasswordReset_(p) {
  const email = String(p.email || '').trim().toLowerCase();
  const user = rows_(SHEETS.USERS).find(x =>
    String(x.email).toLowerCase() === email && x.status === 'active'
  );

  if (!user) return { ok:true };

  const resets = rows_(SHEETS.RESETS)
    .filter(x => x.userId === user.userId && !x.usedAt)
    .sort((a,b)=>new Date(b.createdAt)-new Date(a.createdAt));

  if (resets.length) {
    const latest = new Date(resets[0].createdAt).getTime();
    if (Date.now() - latest < RESET_CODE_COOLDOWN_SECONDS * 1000) {
      throw new Error('驗證碼剛剛已寄出，請稍候再試。');
    }
  }

  const code = String(Math.floor(100000 + Math.random()*900000));
  const resetId = Utilities.getUuid();
  const expires = new Date(Date.now() + RESET_CODE_MINUTES*60000);

  append_(SHEETS.RESETS,[
    resetId,user.userId,user.email,resetCodeHash_(resetId,code),
    expires.toISOString(),0,'',isoNow_()
  ]);

  MailApp.sendEmail({
    to:user.email,
    subject:'食品過期管理 - 密碼重設驗證碼',
    body:'你的食品過期管理密碼重設驗證碼是：' + code +
      '\n\n驗證碼將在 ' + RESET_CODE_MINUTES + ' 分鐘後失效。' +
      '\n如果不是你本人操作，請忽略這封信。'
  });

  return { ok:true };
}

function resetPassword_(p) {
  const email = String(p.email || '').trim().toLowerCase();
  const code = String(p.code || '').trim();
  const newPassword = String(p.newPassword || '');

  if (!/^\d{6}$/.test(code)) throw new Error('驗證碼必須是 6 位數');
  if (newPassword.length < 8) throw new Error('新密碼至少 8 碼');

  const user = rows_(SHEETS.USERS).find(x =>
    String(x.email).toLowerCase() === email && x.status === 'active'
  );
  if (!user) throw new Error('驗證碼無效或已過期');

  const sh = sheet_(SHEETS.RESETS);
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(String);

  const resetIdI=headers.indexOf('resetId');
  const userIdI=headers.indexOf('userId');
  const codeHashI=headers.indexOf('codeHash');
  const expiresI=headers.indexOf('expiresAt');
  const attemptsI=headers.indexOf('attempts');
  const usedAtI=headers.indexOf('usedAt');
  const createdAtI=headers.indexOf('createdAt');

  let target=null;

  for(let r=1;r<values.length;r++) {
    if(String(values[r][userIdI])===user.userId && !values[r][usedAtI]) {
      const item={
        row:r+1,
        resetId:String(values[r][resetIdI]),
        codeHash:String(values[r][codeHashI]),
        expiresAt:values[r][expiresI],
        attempts:Number(values[r][attemptsI]||0),
        createdAt:values[r][createdAtI]
      };

      if(!target || new Date(item.createdAt)>new Date(target.createdAt)) {
        target=item;
      }
    }
  }

  if(!target) throw new Error('驗證碼無效或已過期');
  if(new Date(target.expiresAt).getTime()<=Date.now()) throw new Error('驗證碼已過期，請重新申請');
  if(target.attempts>=RESET_MAX_ATTEMPTS) throw new Error('驗證錯誤次數過多，請重新申請驗證碼');

  if(resetCodeHash_(target.resetId,code)!==target.codeHash) {
    sh.getRange(target.row,attemptsI+1).setValue(target.attempts+1);
    throw new Error('驗證碼錯誤');
  }

  const newSalt=Utilities.getUuid();
  const newHash=passwordHashFast_(newPassword,newSalt);
  updateUserPasswordVersion_(
    user.userId,newSalt,newHash,PASSWORD_HASH_VERSION
  );
  sh.getRange(target.row,usedAtI+1).setValue(isoNow_());
  invalidateUserSessions_(user.userId);

  return { ok:true };
}

function logout_(token) {
  token=String(token||'');
  CacheService.getScriptCache().remove('session:'+token);

  const sh=sheet_(SHEETS.SESSIONS);
  const values=sh.getDataRange().getValues();

  for(let r=values.length-1;r>=1;r--) {
    if(String(values[r][0])===token) {
      sh.deleteRow(r+1);
      break;
    }
  }
  return { ok:true };
}

function me_(user) {
  return {
    ok:true,
    user:{...publicUser_(user),isAdmin:isAdmin_(user)},
    families:userFamilies_(user.userId)
  };
}

function publicUser_(user) {
  return { userId:user.userId,email:user.email,displayName:user.displayName };
}

// ---------------- ADMIN ----------------

function getAdminUserIds_() {
  const raw=String(
    PropertiesService.getScriptProperties().getProperty('ADMIN_USER_IDS') || ''
  );

  return new Set(
    raw
      .split(',')
      .map(x=>String(x||'').trim())
      .filter(Boolean)
  );
}

function saveAdminUserIds_(ids) {
  const values=Array.from(ids||[])
    .map(x=>String(x||'').trim())
    .filter(Boolean)
    .sort();

  PropertiesService
    .getScriptProperties()
    .setProperty('ADMIN_USER_IDS',values.join(','));

  return values;
}

function getPrimaryAdminUserId_() {
  return String(
    PropertiesService.getScriptProperties().getProperty('PRIMARY_ADMIN_USER_ID') || ''
  ).trim();
}

function setPrimaryAdminUserId_(userId) {
  const id=String(userId||'').trim();
  if(!id) throw new Error('主要管理者 User ID 不可空白');

  const props=PropertiesService.getScriptProperties();
  const existing=String(props.getProperty('PRIMARY_ADMIN_USER_ID')||'').trim();

  if(existing && existing!==id) {
    throw new Error('主要管理者已設定，不能由程式自動覆寫');
  }

  props.setProperty('PRIMARY_ADMIN_USER_ID',id);
  return id;
}

function isAdmin_(user) {
  if(!user || !user.userId) return false;
  return getAdminUserIds_().has(String(user.userId));
}

function requireAdmin_(user) {
  if(!isAdmin_(user)) throw new Error('沒有系統管理權限');
  return user;
}

// 第一次建立管理者時可直接從 Apps Script 編輯器執行。
// 會以「目前執行 Apps Script 的 Google 帳號 Email」尋找本系統 Users。
// 第一次成功執行時，該帳號也會被鎖定為 PRIMARY_ADMIN_USER_ID。
// PRIMARY_ADMIN_USER_ID 不提供網站介面修改或刪除；需要更換時只能由 Script Properties 手動處理。
// 若你的食品管理登入 Email 與 Apps Script 擁有者相同，執行一次即可。
function setupCurrentGoogleAccountAsAdmin() {
  const email=String(Session.getEffectiveUser().getEmail()||'')
    .trim()
    .toLowerCase();

  if(!email) {
    throw new Error('無法取得目前 Google 帳號 Email，請改用 Script Properties 手動設定 ADMIN_USER_IDS');
  }

  const user=rows_(SHEETS.USERS).find(x=>
    String(x.email||'').trim().toLowerCase()===email &&
    String(x.status||'')==='active'
  );

  if(!user) {
    throw new Error(
      '找不到相同 Email 的食品管理帳號：'+email+
      '。請先用此 Email 建立食品管理帳號並登入一次，或手動設定 ADMIN_USER_IDS。'
    );
  }

  const ids=getAdminUserIds_();
  ids.add(String(user.userId));
  const saved=saveAdminUserIds_(ids);

  const existingPrimary=getPrimaryAdminUserId_();
  const primaryUserId=existingPrimary || setPrimaryAdminUserId_(user.userId);

  Logger.log(
    'Admin enabled: '+user.displayName+' <'+user.email+'> '+user.userId+
    ' | primaryAdmin='+primaryUserId
  );

  return {
    ok:true,
    userId:user.userId,
    email:user.email,
    displayName:user.displayName,
    adminCount:saved.length,
    primaryAdminUserId:primaryUserId,
    isPrimary:String(primaryUserId)===String(user.userId)
  };
}

function adminUserMaps_() {
  const users=rows_(SHEETS.USERS);
  const families=rows_(SHEETS.FAMILIES);
  const foods=rows_(SHEETS.FOODS);
  const activities=rows_(SHEETS.ACTIVITY)
    .slice()
    .sort((a,b)=>
      (toMillis_(a.createdAt)||0)-(toMillis_(b.createdAt)||0)
    );

  const userMap={};
  users.forEach(u=>{
    userMap[String(u.userId)]={
      userId:String(u.userId),
      displayName:String(u.displayName||''),
      email:String(u.email||''),
      status:String(u.status||'')
    };
  });

  // 先保留 ActivityLog 中曾經出現過的名稱，讓原始資料被刪除後仍能查 Log。
  const historicalFamilyMap={};
  const historicalFoodMap={};

  activities.forEach(row=>{
    const action=String(row.action||'');
    const targetId=String(row.targetId||'');
    const familyId=String(row.familyId||'');
    const detail=String(row.detail||'').trim();

    if(
      detail &&
      targetId &&
      action.indexOf('food_')===0
    ) {
      // 依時間排序後覆寫，保留最後一次已知食品名稱。
      historicalFoodMap[targetId]=detail;
    }

    if(
      detail &&
      familyId &&
      ['family_create','family_rename','family_delete'].includes(action)
    ) {
      historicalFamilyMap[familyId]=detail;
    }
  });

  const familyMap={...historicalFamilyMap};
  families.forEach(f=>{
    // 目前仍存在的群組資料優先。
    familyMap[String(f.familyId)]=String(f.familyName||'');
  });

  const foodMap={...historicalFoodMap};
  foods.forEach(f=>{
    // 目前仍存在的食品資料優先。
    foodMap[String(f.id)]=String(f.name||'');
  });

  return {
    users,
    families,
    foods,
    activities,
    userMap,
    familyMap,
    foodMap,
    historicalFamilyMap,
    historicalFoodMap
  };
}

function shortId_(value) {
  const s=String(value||'');
  if(s.length<=16) return s;
  return s.slice(0,8)+'…'+s.slice(-4);
}

function adminActivityTargetLabel_(row,maps) {
  const action=String(row.action||'');
  const targetId=String(row.targetId||'');
  const detail=String(row.detail||'');

  if(action.indexOf('food_')===0) {
    return detail || maps.foodMap[targetId] || shortId_(targetId);
  }

  if(
    action.indexOf('member_')===0 ||
    action==='family_join' ||
    action==='family_leave'
  ) {
    const u=maps.userMap[targetId];
    return u ? (u.displayName || u.email) : (detail || shortId_(targetId));
  }

  if(action.indexOf('family_')===0 || action==='invite_regenerate') {
    return maps.familyMap[targetId] || detail || shortId_(targetId);
  }

  return detail || shortId_(targetId);
}

function adminSummary_(user,p) {
  requireAdmin_(user);

  const maps=adminUserMaps_();
  const now=Date.now();
  const sessions=rows_(SHEETS.SESSIONS);
  const members=rows_(SHEETS.MEMBERS);

  const activeFoods=maps.foods.filter(f=>
    normalizeFoodStatus_(f.status)==='active'
  ).length;

  const consumedFoods=maps.foods.filter(f=>
    normalizeFoodStatus_(f.status)==='consumed'
  ).length;

  const activeSessions=sessions.filter(s=>{
    const expires=toMillis_(s.expiresAt);
    return expires && expires>now;
  }).length;

  return {
    ok:true,
    summary:{
      users:maps.users.filter(u=>String(u.status||'')==='active').length,
      families:maps.families.length,
      activeMembers:members.filter(m=>String(m.status||'')==='active').length,
      activeFoods,
      consumedFoods,
      activeSessions,
      activityLog:rows_(SHEETS.ACTIVITY).length,
      requestLog:rows_(SHEETS.REQUESTS).length,
      notificationLog:rows_(SHEETS.NOTIFY_LOG).length,
      adminCount:getAdminUserIds_().size,
      retentionDays:{
        activity:ACTIVITY_LOG_RETENTION_DAYS,
        request:REQUEST_LOG_RETENTION_DAYS,
        notification:NOTIFY_LOG_RETENTION_DAYS
      },
      serverTime:isoNow_()
    }
  };
}

function adminLogs_(user,p) {
  requireAdmin_(user);

  const type=String(p.type||'activity');
  if(!['activity','notification','request'].includes(type)) {
    throw new Error('不支援的 Log 類型');
  }

  const page=Math.max(1,Math.floor(Number(p.page||1)));
  const pageSize=Math.max(
    10,
    Math.min(100,Math.floor(Number(p.pageSize||20)))
  );
  const query=String(p.query||'').trim().toLowerCase();
  const maps=adminUserMaps_();

  let items=[];

  if(type==='activity') {
    items=rows_(SHEETS.ACTIVITY).map(row=>{
      const u=maps.userMap[String(row.userId)]||{};
      return {
        time:String(row.createdAt||''),
        familyId:String(row.familyId||''),
        familyName:maps.familyMap[String(row.familyId)]||'',
        userId:String(row.userId||''),
        userName:String(u.displayName||''),
        userEmail:String(u.email||''),
        action:String(row.action||''),
        targetId:String(row.targetId||''),
        targetLabel:adminActivityTargetLabel_(row,maps),
        detail:String(row.detail||'')
      };
    });
  } else if(type==='notification') {
    items=rows_(SHEETS.NOTIFY_LOG).map(row=>{
      const u=maps.userMap[String(row.userId)]||{};
      return {
        time:String(row.sentAt||''),
        familyId:String(row.familyId||''),
        familyName:String(
          row.familyName ||
          maps.familyMap[String(row.familyId)] ||
          ''
        ),
        userId:String(row.userId||''),
        userName:String(u.displayName||''),
        userEmail:String(u.email||''),
        foodId:String(row.foodId||''),
        foodName:String(
          row.foodName ||
          maps.foodMap[String(row.foodId)] ||
          ''
        ),
        expiry:formatDate_(row.expiry),
        daysBefore:Number(row.daysBefore||0)
      };
    });
  } else {
    items=rows_(SHEETS.REQUESTS).map(row=>{
      const u=maps.userMap[String(row.userId)]||{};
      return {
        time:String(row.updatedAt||row.createdAt||''),
        createdAt:String(row.createdAt||''),
        updatedAt:String(row.updatedAt||''),
        requestId:String(row.requestId||''),
        requestIdShort:shortId_(row.requestId),
        userId:String(row.userId||''),
        userName:String(u.displayName||''),
        userEmail:String(u.email||''),
        action:String(row.action||''),
        status:String(row.status||'')
      };
    });
  }

  if(query) {
    items=items.filter(item=>
      Object.values(item).some(value=>
        String(value??'').toLowerCase().includes(query)
      )
    );
  }

  items.sort((a,b)=>
    (toMillis_(b.time)||0)-(toMillis_(a.time)||0)
  );

  const total=items.length;
  const totalPages=Math.max(1,Math.ceil(total/pageSize));
  const safePage=Math.min(page,totalPages);
  const start=(safePage-1)*pageSize;
  const pageItems=items.slice(start,start+pageSize);

  return {
    ok:true,
    type,
    page:safePage,
    pageSize,
    total,
    totalPages,
    items:pageItems
  };
}

function adminListAdmins_(user,p) {
  requireAdmin_(user);

  const ids=getAdminUserIds_();
  const primaryUserId=getPrimaryAdminUserId_();
  const users=rows_(SHEETS.USERS);
  const admins=Array.from(ids).map(userId=>{
    const u=users.find(x=>String(x.userId)===String(userId));
    return {
      userId:String(userId),
      displayName:u ? String(u.displayName||'') : '',
      email:u ? String(u.email||'') : '',
      status:u ? String(u.status||'') : 'missing',
      isSelf:String(userId)===String(user.userId),
      isPrimary:Boolean(primaryUserId) && String(userId)===String(primaryUserId)
    };
  });

  admins.sort((a,b)=>{
    if(a.isPrimary!==b.isPrimary) return a.isPrimary ? -1 : 1;
    return String(a.displayName||a.email||a.userId)
      .localeCompare(String(b.displayName||b.email||b.userId));
  });

  return {
    ok:true,
    admins,
    primaryConfigured:Boolean(primaryUserId)
  };
}

function adminAddAdmin_(user,p) {
  requireAdmin_(user);

  const email=String(p.email||'').trim().toLowerCase();
  if(!email) throw new Error('請輸入要設為管理者的 Email');

  const target=rows_(SHEETS.USERS).find(x=>
    String(x.email||'').trim().toLowerCase()===email &&
    String(x.status||'')==='active'
  );

  if(!target) {
    throw new Error('找不到這個已啟用的食品管理帳號');
  }

  const ids=getAdminUserIds_();
  ids.add(String(target.userId));
  saveAdminUserIds_(ids);

  console.log(
    'Admin added by '+user.userId+': '+target.userId
  );

  return {
    ok:true,
    added:{
      userId:String(target.userId),
      displayName:String(target.displayName||''),
      email:String(target.email||'')
    }
  };
}

function adminRemoveAdmin_(user,p) {
  requireAdmin_(user);

  const targetUserId=String(p.userId||'').trim();
  if(!targetUserId) throw new Error('缺少管理者 User ID');

  const ids=getAdminUserIds_();
  if(!ids.has(targetUserId)) return {ok:true,removed:false};

  const primaryUserId=getPrimaryAdminUserId_();
  if(!primaryUserId) {
    throw new Error('尚未設定主要管理者，請先由 Apps Script 執行 setupCurrentGoogleAccountAsAdmin()');
  }

  if(targetUserId===primaryUserId) {
    throw new Error('主要管理者不可移除');
  }

  if(ids.size<=1) {
    throw new Error('至少要保留一位系統管理者');
  }

  ids.delete(targetUserId);
  saveAdminUserIds_(ids);

  console.log(
    'Admin removed by '+user.userId+': '+targetUserId
  );

  return {
    ok:true,
    removed:true,
    selfRemoved:targetUserId===String(user.userId)
  };
}

// ---------------- NOTIFICATIONS ----------------

function getNotificationSettings_(user) {
  const userRow=rows_(SHEETS.NOTIFY_USER).find(x=>x.userId===user.userId);
  const groupRows=rows_(SHEETS.NOTIFY_GROUP).filter(x=>x.userId===user.userId);
  const families=userFamilies_(user.userId);

  return {
    ok:true,
    settings:{
      emailEnabled:userRow ? toBool_(userRow.emailEnabled) : false,
      sendHour:userRow ? Math.max(0,Math.min(23,Number(userRow.sendHour||8))) : 8,
      groups:families.map(f=>{
        const s=groupRows.find(x=>x.familyId===f.familyId);
        return {
          familyId:f.familyId,
          familyName:f.familyName,
          enabled:s ? toBool_(s.enabled) : true,
          defaultDaysBefore:s ? Math.max(0,Math.min(365,Number(s.defaultDaysBefore||3))) : 3
        };
      })
    }
  };
}

function saveNotificationSettings_(user,p) {
  const emailEnabled=toBool_(p.emailEnabled);
  const sendHour=Math.max(0,Math.min(23,Number(p.sendHour||0)));

  let groups=[];
  try { groups=JSON.parse(String(p.groups||'[]')); }
  catch { throw new Error('群組提醒設定格式錯誤'); }

  if(!Array.isArray(groups)) throw new Error('群組提醒設定格式錯誤');

  upsertNotificationUser_(user.userId,emailEnabled,sendHour);

  for(const g of groups) {
    const familyId=String(g.familyId||'');
    requireMembership_(user.userId,familyId);

    upsertNotificationGroup_(
      user.userId,
      familyId,
      Boolean(g.enabled),
      Math.max(0,Math.min(365,Number(g.defaultDaysBefore||0)))
    );
  }

  return getNotificationSettings_(user);
}

function setupNotificationTrigger() {
  const handler='checkExpiryNotifications';

  ScriptApp.getProjectTriggers().forEach(t=>{
    if(t.getHandlerFunction()===handler) ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger(handler)
    .timeBased()
    .everyHours(1)
    .create();
}

function checkExpiryNotifications() {
  // 自動補齊 NotificationLog 的快照欄位，避免 Trigger 在升級後遇到舊 schema。
  ensureNotificationLogSchema_();

  const tz=Session.getScriptTimeZone();
  const now=new Date();
  const currentHour=Number(Utilities.formatDate(now,tz,'H'));
  let sentUsers=0;
  let sentItems=0;
  let failedUsers=0;

  const users=rows_(SHEETS.USERS).filter(x=>x.status==='active');
  const userSettings=rows_(SHEETS.NOTIFY_USER);
  const groupSettings=rows_(SHEETS.NOTIFY_GROUP);
  const memberships=rows_(SHEETS.MEMBERS).filter(x=>x.status==='active');
  const families=rows_(SHEETS.FAMILIES);
  const foods=rows_(SHEETS.FOODS);
  const logs=rows_(SHEETS.NOTIFY_LOG);

  for(const u of users) {
    const us=userSettings.find(x=>x.userId===u.userId);
    if(!us || !toBool_(us.emailEnabled)) continue;

    // Trigger 不保證精準在整點執行。設定時間之前不寄；設定時間之後若
    // 當天尚未成功寄送，後續每小時執行時都可以補寄。
    const sendHour=Math.max(0,Math.min(23,Number(us.sendHour||0)));
    if(currentHour<sendHour) continue;

    const memberRows=memberships.filter(x=>x.userId===u.userId);
    if(!memberRows.length) continue;

    const items=[];
    const queuedItemKeys=new Set();

    for(const m of memberRows) {
      const fam=families.find(x=>x.familyId===m.familyId);
      if(!fam) continue;

      const gset=groupSettings.find(x=>x.userId===u.userId && x.familyId===m.familyId);
      const enabled=gset ? toBool_(gset.enabled) : true;
      const defaultDays=gset ? Math.max(0,Math.min(365,Number(gset.defaultDaysBefore||3))) : 3;
      if(!enabled) continue;

      const groupFoods=foods.filter(x=>
        x.familyId===m.familyId &&
        normalizeFoodStatus_(x.status)==='active'
      );

      for(const f of groupFoods) {
        const mode=String(f.notifyMode||'inherit');
        if(mode==='off') continue;

        const daysBefore=mode==='custom'
          ? Math.max(0,Math.min(365,Number(f.notifyDaysBefore||0)))
          : defaultDays;

        const days=daysUntilDate_(f.expiry,tz);

        // 已進入提醒區間、但尚未過期的食品都可補寄。
        // 例如設定提前 3 天，Trigger 在第 3 天失敗，第 2 天仍會補寄一次。
        if(!Number.isFinite(days) || days<0 || days>daysBefore) continue;

        const already=logs.some(l=>
          l.userId===u.userId &&
          l.familyId===m.familyId &&
          l.foodId===f.id &&
          formatDate_(l.expiry)===formatDate_(f.expiry) &&
          Number(l.daysBefore)===daysBefore
        );

        if(already) continue;

        const itemKey=[
          u.userId,
          m.familyId,
          f.id,
          formatDate_(f.expiry),
          daysBefore
        ].join('|');

        if(queuedItemKeys.has(itemKey)) continue;
        queuedItemKeys.add(itemKey);

        items.push({
          familyId:m.familyId,
          familyName:fam.familyName,
          foodId:f.id,
          name:f.name,
          qty:Number(f.qty||1),
          expiry:formatDate_(f.expiry),
          daysBefore,
          daysLeft:days
        });
      }
    }

    if(!items.length) continue;

    const grouped={};
    items.forEach(i=>{
      if(!grouped[i.familyName]) grouped[i.familyName]=[];
      grouped[i.familyName].push(i);
    });

    let body='食品過期管理提醒\n\n';

    Object.keys(grouped).forEach(groupName=>{
      body+='【'+groupName+'】\n';
      grouped[groupName].forEach(i=>{
        const when=i.daysLeft===0 ? '今天到期' : '剩 '+i.daysLeft+' 天到期';
        body+='- '+i.name+' ×'+i.qty+'：'+when+'（'+i.expiry+'）\n';
      });
      body+='\n';
    });

    try {
      MailApp.sendEmail({
        to:u.email,
        subject:'食品過期提醒：'+items.length+' 項食品需要注意',
        body
      });

      items.forEach(i=>{
        append_(SHEETS.NOTIFY_LOG,[
          u.userId,
          i.familyId,
          i.foodId,
          i.expiry,
          i.daysBefore,
          isoNow_(),
          i.name,
          i.familyName
        ]);

        // 保持本次執行的記憶體快照與剛寫入的 Log 一致。
        logs.push({
          userId:u.userId,
          familyId:i.familyId,
          foodId:i.foodId,
          expiry:i.expiry,
          daysBefore:i.daysBefore
        });
      });

      sentUsers++;
      sentItems+=items.length;
    } catch(err) {
      failedUsers++;
      console.error(
        'Expiry notification failed for user '+u.userId+': '+
        String(err.message||err)
      );
    }
  }

  const result={
    ok:failedUsers===0,
    sentUsers,
    sentItems,
    failedUsers,
    checkedAt:isoNow_()
  };

  console.log('Expiry notification result: '+JSON.stringify(result));
  return result;
}

function ensureNotificationLogSchema_() {
  const sh=sheet_(SHEETS.NOTIFY_LOG);
  ensureHeaders_(sh,NOTIFICATION_LOG_HEADERS);
  return sh;
}

// 升級後可手動執行一次：
// 盡量替舊 NotificationLog 補上食品/群組名稱。
// 能從目前資料或 ActivityLog 歷史推回的會自動補；
// 已經完全沒有任何名稱線索的舊資料仍會保留空白並由管理頁顯示 ID。
function backfillNotificationLogLabels() {
  const sh=ensureNotificationLogSchema_();
  const values=sh.getDataRange().getValues();

  if(values.length<=1) {
    return {
      ok:true,
      rows:0,
      foodNamesFilled:0,
      familyNamesFilled:0
    };
  }

  const headers=values[0].map(String);
  const familyIdI=headers.indexOf('familyId');
  const foodIdI=headers.indexOf('foodId');
  const foodNameI=headers.indexOf('foodName');
  const familyNameI=headers.indexOf('familyName');

  if(
    familyIdI<0 ||
    foodIdI<0 ||
    foodNameI<0 ||
    familyNameI<0
  ) {
    throw new Error('NotificationLog 欄位不完整');
  }

  const maps=adminUserMaps_();
  let foodNamesFilled=0;
  let familyNamesFilled=0;

  const foodNameValues=[];
  const familyNameValues=[];

  for(let r=1;r<values.length;r++) {
    const row=values[r];

    let foodName=String(row[foodNameI]||'').trim();
    let familyName=String(row[familyNameI]||'').trim();

    if(!foodName) {
      const recovered=String(
        maps.foodMap[String(row[foodIdI]||'')] || ''
      ).trim();

      if(recovered) {
        foodName=recovered;
        foodNamesFilled++;
      }
    }

    if(!familyName) {
      const recovered=String(
        maps.familyMap[String(row[familyIdI]||'')] || ''
      ).trim();

      if(recovered) {
        familyName=recovered;
        familyNamesFilled++;
      }
    }

    foodNameValues.push([foodName]);
    familyNameValues.push([familyName]);
  }

  sh.getRange(
    2,
    foodNameI+1,
    foodNameValues.length,
    1
  ).setValues(foodNameValues);

  sh.getRange(
    2,
    familyNameI+1,
    familyNameValues.length,
    1
  ).setValues(familyNameValues);

  const result={
    ok:true,
    rows:values.length-1,
    foodNamesFilled,
    familyNamesFilled
  };

  Logger.log(JSON.stringify(result));
  return result;
}

function upsertNotificationUser_(userId,emailEnabled,sendHour) {
  const sh=sheet_(SHEETS.NOTIFY_USER);
  const values=sh.getDataRange().getValues();
  const headers=values[0].map(String);
  const ui=headers.indexOf('userId');

  for(let r=1;r<values.length;r++) {
    if(String(values[r][ui])===userId) {
      sh.getRange(r+1,headers.indexOf('emailEnabled')+1).setValue(emailEnabled);
      sh.getRange(r+1,headers.indexOf('sendHour')+1).setValue(sendHour);
      sh.getRange(r+1,headers.indexOf('updatedAt')+1).setValue(isoNow_());
      return;
    }
  }

  append_(SHEETS.NOTIFY_USER,[userId,emailEnabled,sendHour,isoNow_()]);
}

function upsertNotificationGroup_(userId,familyId,enabled,days) {
  const sh=sheet_(SHEETS.NOTIFY_GROUP);
  const values=sh.getDataRange().getValues();
  const headers=values[0].map(String);
  const ui=headers.indexOf('userId');
  const fi=headers.indexOf('familyId');

  for(let r=1;r<values.length;r++) {
    if(String(values[r][ui])===userId && String(values[r][fi])===familyId) {
      sh.getRange(r+1,headers.indexOf('enabled')+1).setValue(enabled);
      sh.getRange(r+1,headers.indexOf('defaultDaysBefore')+1).setValue(days);
      sh.getRange(r+1,headers.indexOf('updatedAt')+1).setValue(isoNow_());
      return;
    }
  }

  append_(SHEETS.NOTIFY_GROUP,[userId,familyId,enabled,days,isoNow_()]);
}

function daysUntilDate_(value,tz) {
  const s=formatDate_(value);
  if(!/^\d{4}-\d{2}-\d{2}$/.test(s)) return NaN;

  const todayStr=Utilities.formatDate(new Date(),tz,'yyyy-MM-dd');
  const a=new Date(todayStr+'T00:00:00');
  const b=new Date(s+'T00:00:00');
  return Math.round((b.getTime()-a.getTime())/86400000);
}

function toBool_(v) {
  if(v===true) return true;
  const s=String(v||'').toLowerCase();
  return s==='true' || s==='1' || s==='yes' || s==='on';
}

// ---------------- FAMILY ----------------

function createFamily_(user,p) {
  const familyName=String(p.familyName||'').trim();
  const requestId=String(p.requestId||'').trim() ||
    legacyRequestId_(user.userId,'createFamily',familyName);

  if(!familyName) throw new Error('請輸入群組名稱');

  const lock=LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const existing=findRequest_(requestId,user.userId,'createFamily');

    if(existing) {
      const result=parseRequestResult_(existing.resultJson);

      if(existing.status==='done' && result) {
        return {ok:true,...result,replayed:true};
      }

      // 如果前一次已經寫入 Families，但來不及把 RequestLog 標成 done，
      // 用預先記錄的 familyId 確認並補完成狀態。
      if(result && result.family && result.family.familyId) {
        const familyExists=rows_(SHEETS.FAMILIES)
          .some(x=>x.familyId===result.family.familyId);

        if(familyExists) {
          ensureOwnerMembership_(
            result.family.familyId,
            user.userId
          );
          markRequestDone_(requestId,user.userId,'createFamily',result);
          return {ok:true,...result,replayed:true};
        }
      }
    }

    const familyId = existing
      ? parseRequestResult_(existing.resultJson)?.family?.familyId
      : Utilities.getUuid();

    const inviteCode = existing
      ? parseRequestResult_(existing.resultJson)?.family?.inviteCode
      : uniqueInviteCode_();

    const result = {
      family:{
        familyId,
        familyName,
        role:'owner',
        inviteCode
      }
    };

    if(!existing) {
      append_(SHEETS.REQUESTS,[
        requestId,
        user.userId,
        'createFamily',
        'processing',
        JSON.stringify(result),
        isoNow_(),
        isoNow_()
      ]);
    }

    const families=rows_(SHEETS.FAMILIES);

    if(!families.some(x=>x.familyId===familyId)) {
      append_(SHEETS.FAMILIES,[
        familyId,familyName,user.userId,inviteCode,isoNow_()
      ]);
    }

    ensureOwnerMembership_(familyId,user.userId);

    markRequestDone_(
      requestId,
      user.userId,
      'createFamily',
      result
    );

    logActivity_(
      familyId,user.userId,'family_create',familyId,familyName
    );

    return {ok:true,...result};
  } finally {
    lock.releaseLock();
  }
}

function joinFamily_(user,p) {
  const code=String(p.inviteCode||'').trim().toUpperCase();
  const requestId=String(p.requestId||'').trim() ||
    legacyRequestId_(user.userId,'joinFamily',code);

  if(!code) throw new Error('請輸入群組邀請碼');

  const lock=LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const existing=findRequest_(requestId,user.userId,'joinFamily');

    if(existing && existing.status==='done') {
      const result=parseRequestResult_(existing.resultJson);
      if(result) return {ok:true,...result,replayed:true};
    }

    const family=rows_(SHEETS.FAMILIES).find(x =>
      String(x.inviteCode).toUpperCase()===code
    );
    if(!family) throw new Error('邀請碼不存在');

    const result={
      family:{
        familyId:family.familyId,
        familyName:family.familyName,
        role:'member',
        inviteCode:''
      }
    };

    if(!existing) {
      append_(SHEETS.REQUESTS,[
        requestId,
        user.userId,
        'joinFamily',
        'processing',
        JSON.stringify(result),
        isoNow_(),
        isoNow_()
      ]);
    }

    const members=rows_(SHEETS.MEMBERS);
    const member=members.find(x =>
      x.familyId===family.familyId &&
      x.userId===user.userId
    );

    if(member && member.status==='active') {
      // 同一 requestId 的安全重播不算錯誤。
      if(existing) {
        markRequestDone_(
          requestId,user.userId,'joinFamily',result
        );
        return {ok:true,...result,replayed:true};
      }

      throw new Error('你已經是這個群組的成員');
    }

    if(member) {
      updateMemberStatusRole_(
        family.familyId,user.userId,'active','member'
      );
    } else {
      append_(SHEETS.MEMBERS,[
        family.familyId,user.userId,'member','active',isoNow_()
      ]);
    }

    markRequestDone_(
      requestId,user.userId,'joinFamily',result
    );

    logActivity_(
      family.familyId,user.userId,'family_join',user.userId,''
    );

    return {ok:true,...result};
  } finally {
    lock.releaseLock();
  }
}

function checkRequest_(user,p) {
  const requestId=String(p.requestId||'').trim();
  const requestAction=String(p.requestAction||'').trim();

  if(!requestId || !requestAction) {
    throw new Error('缺少 requestId');
  }

  const item=findRequest_(
    requestId,user.userId,requestAction
  );

  if(!item) {
    return {
      ok:true,
      found:false,
      status:'not_found'
    };
  }

  return {
    ok:true,
    found:true,
    status:item.status,
    result:parseRequestResult_(item.resultJson)
  };
}

function findRequest_(requestId,userId,action) {
  return rows_(SHEETS.REQUESTS).find(x =>
    x.requestId===requestId &&
    x.userId===userId &&
    x.action===action
  ) || null;
}

function parseRequestResult_(s) {
  try {
    return JSON.parse(String(s||''));
  } catch {
    return null;
  }
}

function markRequestDone_(requestId,userId,action,result) {
  const sh=sheet_(SHEETS.REQUESTS);
  const values=sh.getDataRange().getValues();
  const headers=values[0].map(String);

  const ri=headers.indexOf('requestId');
  const ui=headers.indexOf('userId');
  const ai=headers.indexOf('action');
  const si=headers.indexOf('status');
  const ji=headers.indexOf('resultJson');
  const ti=headers.indexOf('updatedAt');

  for(let r=1;r<values.length;r++) {
    if(
      String(values[r][ri])===requestId &&
      String(values[r][ui])===userId &&
      String(values[r][ai])===action
    ) {
      sh.getRange(r+1,si+1).setValue('done');
      sh.getRange(r+1,ji+1).setValue(JSON.stringify(result));
      sh.getRange(r+1,ti+1).setValue(isoNow_());
      return;
    }
  }

  append_(SHEETS.REQUESTS,[
    requestId,userId,action,'done',
    JSON.stringify(result),isoNow_(),isoNow_()
  ]);
}

function ensureOwnerMembership_(familyId,userId) {
  const members=rows_(SHEETS.MEMBERS);

  const existing=members.find(x =>
    x.familyId===familyId &&
    x.userId===userId
  );

  if(existing) {
    if(
      existing.status!=='active' ||
      existing.role!=='owner'
    ) {
      updateMemberStatusRole_(
        familyId,userId,'active','owner'
      );
    }
    return;
  }

  append_(SHEETS.MEMBERS,[
    familyId,userId,'owner','active',isoNow_()
  ]);
}

function familyDetails_(user,p) {
  const familyId=String(p.familyId||'');
  const membership=requireMembership_(user.userId,familyId);

  const family=rows_(SHEETS.FAMILIES).find(x=>x.familyId===familyId);
  if(!family) throw new Error('家庭不存在');

  const users=rows_(SHEETS.USERS);
  const members=rows_(SHEETS.MEMBERS)
    .filter(x=>x.familyId===familyId && x.status==='active')
    .map(m=>{
      const u=users.find(x=>x.userId===m.userId);
      return {
        userId:m.userId,
        displayName:u?u.displayName:'未知使用者',
        email:u?u.email:'',
        role:m.role
      };
    });

  return {
    ok:true,
    family:{
      familyId,
      familyName:family.familyName,
      ownerUserId:family.ownerUserId,
      inviteCode:membership.role==='owner'?family.inviteCode:'',
      myRole:membership.role,
      members
    }
  };
}

function renameFamily_(user,p) {
  const familyId=String(p.familyId||'');
  requireOwner_(user.userId,familyId);

  const familyName=String(p.familyName||'').trim();
  if(!familyName) throw new Error('家庭名稱不可空白');

  updateFamilyCell_(familyId,'familyName',familyName);
  logActivity_(familyId,user.userId,'family_rename',familyId,familyName);

  return { ok:true,familyName };
}

function regenerateInviteCode_(user,p) {
  const familyId=String(p.familyId||'');
  requireOwner_(user.userId,familyId);

  const inviteCode=uniqueInviteCode_();
  updateFamilyCell_(familyId,'inviteCode',inviteCode);
  logActivity_(familyId,user.userId,'invite_regenerate',familyId,'');

  return { ok:true,inviteCode };
}

function setMemberRole_(user,p) {
  const familyId=String(p.familyId||'');
  const owner=requireOwner_(user.userId,familyId);

  const targetUserId=String(p.targetUserId||'');
  const role=String(p.role||'');

  if(!['member','viewer'].includes(role)) throw new Error('角色不正確');
  if(targetUserId===owner.userId) throw new Error('不能修改 Owner 的角色');

  const target=requireMembership_(targetUserId,familyId);
  if(target.role==='owner') throw new Error('不能修改 Owner 的角色');

  updateMemberRole_(familyId,targetUserId,role);
  logActivity_(familyId,user.userId,'member_role',targetUserId,role);

  return { ok:true };
}

function removeMember_(user,p) {
  const familyId=String(p.familyId||'');
  const owner=requireOwner_(user.userId,familyId);
  const targetUserId=String(p.targetUserId||'');

  if(targetUserId===owner.userId) throw new Error('不能移除 Owner');

  const target=requireMembership_(targetUserId,familyId);
  if(target.role==='owner') throw new Error('不能移除 Owner');

  updateMemberStatusRole_(familyId,targetUserId,'removed',target.role);
  logActivity_(familyId,user.userId,'member_remove',targetUserId,'');

  return { ok:true };
}

function leaveFamily_(user,p) {
  const familyId=String(p.familyId||'');
  const membership=requireMembership_(user.userId,familyId);

  if(membership.role==='owner') {
    throw new Error('Owner 目前不能直接離開家庭');
  }

  updateMemberStatusRole_(familyId,user.userId,'left',membership.role);
  logActivity_(familyId,user.userId,'family_leave',user.userId,'');

  return { ok:true };
}

function deleteFamily_(user,p) {
  const familyId=String(p.familyId||'');
  requireOwner_(user.userId,familyId);

  if(String(p.confirm||'')!=='DELETE') throw new Error('確認文字不正確');

  const family=rows_(SHEETS.FAMILIES).find(x=>
    String(x.familyId)===familyId
  );
  const familyName=family ? String(family.familyName||'') : '';

  markFamilyImagesPendingDelete_(familyId);

  deleteRowsByValue_(SHEETS.FOODS,'familyId',familyId);
  deleteRowsByValue_(SHEETS.MEMBERS,'familyId',familyId);

  // ActivityLog 是 90 天稽核紀錄，不跟著群組立即刪除。
  // 群組本體刪除後仍保留原有操作紀錄，並新增一筆刪除群組紀錄。
  deleteRowsByValue_(SHEETS.FAMILIES,'familyId',familyId);

  logActivity_(
    familyId,
    user.userId,
    'family_delete',
    familyId,
    familyName
  );

  return { ok:true };
}

// ---------------- FOOD / BATCH SYNC ----------------

function listFoods_(user,p) {
  const familyId=String(p.familyId||'');
  const membership=requireMembership_(user.userId,familyId);

  // 前端需要同時顯示「目前食品」與「已用完」歷史，
  // 所以這裡回傳群組內全部食品；前端再依 status 分頁顯示。
  const foods=rows_(SHEETS.FOODS)
    .filter(x=>x.familyId===familyId)
    .map(foodForClient_);

  return {
    ok:true,
    foods,
    myRole:membership.role
  };
}

function syncChanges_(user,p) {
  const familyId=String(p.familyId||'');
  requireEditableMembership_(user.userId,familyId);

  let changes=[];
  try {
    changes=JSON.parse(String(p.changes||'[]'));
  } catch {
    throw new Error('變更資料格式錯誤');
  }

  if(!Array.isArray(changes)) throw new Error('變更資料格式錯誤');
  if(changes.length>100) throw new Error('單次同步最多 100 筆變更');

  const lock=LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const sh=sheet_(SHEETS.FOODS);
    let values=sh.getDataRange().getValues();
    const headers=values[0].map(String);

    const idx={};
    headers.forEach((h,i)=>idx[h]=i);

    if(idx.imageId==null || idx.status==null || idx.consumedAt==null) {
      throw new Error('Foods 欄位尚未更新，請先執行 setupDatabase()');
    }

    // 每一筆 clientKey 都會寫入 RequestLog。
    // 同一個 change 即使因網路逾時被前端重送，也只會作用在同一筆食品上。
    const reqSh=sheet_(SHEETS.REQUESTS);
    let reqValues=reqSh.getDataRange().getValues();
    const reqHeaders=reqValues[0].map(String);
    const reqIdx={};
    reqHeaders.forEach((h,i)=>reqIdx[h]=i);

    const requiredRequestHeaders=[
      'requestId','userId','action','status','resultJson','createdAt','updatedAt'
    ];
    if(requiredRequestHeaders.some(h=>reqIdx[h]==null)) {
      throw new Error('RequestLog 欄位不完整，請先執行 setupDatabase()');
    }

    const requestMap=new Map();
    const requestMapKey_=(requestId,action)=>requestId+'\u0001'+action;

    for(let r=1;r<reqValues.length;r++) {
      if(String(reqValues[r][reqIdx.userId])!==user.userId) continue;
      requestMap.set(
        requestMapKey_(
          String(reqValues[r][reqIdx.requestId]),
          String(reqValues[r][reqIdx.action])
        ),
        r
      );
    }

    const getRequest_=(requestId,action)=>{
      const rowIndex=requestMap.get(requestMapKey_(requestId,action));
      if(rowIndex==null) return null;

      return {
        rowIndex,
        status:String(reqValues[rowIndex][reqIdx.status]||''),
        result:parseRequestResult_(reqValues[rowIndex][reqIdx.resultJson])
      };
    };

    const saveRequest_=(requestId,action,status,result)=>{
      const now=isoNow_();
      const key=requestMapKey_(requestId,action);
      let rowIndex=requestMap.get(key);

      if(rowIndex==null) {
        const row=[
          requestId,
          user.userId,
          action,
          status,
          JSON.stringify(result),
          now,
          now
        ];

        reqSh.appendRow(row);
        reqValues.push(row);
        rowIndex=reqValues.length-1;
        requestMap.set(key,rowIndex);
      } else {
        reqSh.getRange(rowIndex+1,reqIdx.status+1).setValue(status);
        reqSh.getRange(rowIndex+1,reqIdx.resultJson+1)
          .setValue(JSON.stringify(result));
        reqSh.getRange(rowIndex+1,reqIdx.updatedAt+1).setValue(now);

        reqValues[rowIndex][reqIdx.status]=status;
        reqValues[rowIndex][reqIdx.resultJson]=JSON.stringify(result);
        reqValues[rowIndex][reqIdx.updatedAt]=now;
      }

      return {
        rowIndex,
        status,
        result
      };
    };

    const requestIdFor_=(clientKey)=>
      'food-sync:'+familyId+':'+clientKey;

    const requestActionFor_=(type)=>
      'syncFood:'+type;

    const applied=[];
    const items=[];
    const deletedIds=[];
    const conflicts=[];

    for(const ch of changes) {
      const type=String(ch.type||'');
      const clientKey=String(ch.clientKey||'').trim();

      // 新版前端一定會帶 clientKey；舊版若沒帶，仍允許執行但無法提供跨重試去重。
      const effectiveClientKey=clientKey || Utilities.getUuid();
      const requestId=requestIdFor_(effectiveClientKey);
      const requestAction=requestActionFor_(type);

      if(type==='add') {
        const f=validateFoodPayload_(ch.food||{});
        const incomingTempId=String(ch.tempId||'');

        let req=getRequest_(requestId,requestAction);
        let state=req && req.result ? req.result : null;

        if(state && state.familyId && state.familyId!==familyId) {
          conflicts.push({
            clientKey:effectiveClientKey,
            id:'',
            reason:'request_family_mismatch'
          });
          continue;
        }

        if(!state || !state.food || !state.food.id) {
          const id=Utilities.getUuid();
          const now=isoNow_();

          state={
            familyId,
            tempId:incomingTempId,
            expectedUpdatedAt:'',
            oldImageId:'',
            food:savedFoodFromPayload_(
              id,
              f,
              now.slice(0,10),
              now
            )
          };

          req=saveRequest_(
            requestId,
            requestAction,
            'processing',
            state
          );
        }

        const id=String(state.food.id||'');
        let rowIndex=findFoodRow_(values,idx,familyId,id);

        if(rowIndex<1) {
          // 已標記 done 卻找不到原本新增的食品，代表之後可能已被其他裝置刪除。
          // 不可在 retry 時把它重新建立回來。
          if(req && req.status==='done') {
            conflicts.push({
              clientKey:effectiveClientKey,
              id,
              reason:'not_found'
            });
            continue;
          }

          validateImageReferenceForChange_(
            user,
            familyId,
            f.imageId,
            ''
          );

          // processing 期間若使用者又編輯了 tmp 食品，保留同一個正式 id，
          // 但把尚未真正落盤的內容更新為最新版本。
          state.tempId=state.tempId || incomingTempId;
          state.food=savedFoodFromPayload_(
            id,
            f,
            state.food.createAt || isoNow_().slice(0,10),
            state.food.updatedAt || isoNow_()
          );
          state.oldImageId='';

          saveRequest_(
            requestId,
            requestAction,
            'processing',
            state
          );

          const saved=state.food;

          sh.appendRow([
            familyId,id,saved.name,saved.qty,saved.location,saved.expiry,saved.note,
            user.userId,saved.createAt,saved.updatedAt,
            saved.notifyMode,saved.notifyDaysBefore,saved.imageId,
            saved.status,saved.consumedAt
          ]);

          values.push([
            familyId,id,saved.name,saved.qty,saved.location,saved.expiry,saved.note,
            user.userId,saved.createAt,saved.updatedAt,
            saved.notifyMode,saved.notifyDaysBefore,saved.imageId,
            saved.status,saved.consumedAt
          ]);
          rowIndex=values.length-1;

          if(saved.imageId) {
            attachImageToFood_(
              saved.imageId,
              familyId,
              id,
              user.userId
            );
          }

          logActivity_(
            familyId,
            user.userId,
            'food_add',
            id,
            saved.name
          );
        } else {
          let current=foodFromValuesRow_(values,rowIndex,idx);

          // 如果前一次其實已成功，只是前端沒有收到回應，直接重播結果。
          // 若使用者在這段期間又修改了 tmp 食品，且伺服器那筆尚未被別人改過，
          // 則把「同一個尚未確認完成的新增」更新成最新內容，而不是再 append 一列。
          if(!foodPayloadMatches_(current,f)) {
            const lastKnown=String(state.food.updatedAt||'');

            if(
              current.updatedAt &&
              lastKnown &&
              current.updatedAt!==lastKnown
            ) {
              conflicts.push({
                clientKey:effectiveClientKey,
                id,
                reason:'modified_elsewhere'
              });
              continue;
            }

            validateImageReferenceForChange_(
              user,
              familyId,
              f.imageId,
              id
            );

            const oldImageId=current.imageId||'';
            const now=isoNow_();
            const saved=savedFoodFromPayload_(
              id,
              f,
              current.createAt,
              now
            );

            state.expectedUpdatedAt=current.updatedAt||'';
            state.oldImageId=oldImageId;
            state.food=saved;

            saveRequest_(
              requestId,
              requestAction,
              'processing',
              state
            );

            writeFoodRow_(
              sh,
              values,
              rowIndex,
              idx,
              f,
              now
            );

            if(f.imageId!==oldImageId) {
              if(f.imageId) {
                attachImageToFood_(
                  f.imageId,
                  familyId,
                  id,
                  user.userId
                );
              }

              if(oldImageId) {
                markImagePendingDelete_(
                  oldImageId,
                  'food_image_replaced'
                );
              }
            }

            current=saved;

            logActivity_(
              familyId,
              user.userId,
              'food_update',
              id,
              f.name
            );
          } else {
            // RequestLog 可能仍停在 processing，但食品列已經寫入。
            // 補做可安全重入的圖片綁定後即可完成這筆 request。
            if(current.imageId) {
              attachImageToFood_(
                current.imageId,
                familyId,
                id,
                user.userId
              );
            }
          }

          state.food=current;
        }

        state.tempId=state.tempId || incomingTempId;
        state.oldImageId='';
        saveRequest_(
          requestId,
          requestAction,
          'done',
          state
        );

        applied.push({clientKey:effectiveClientKey});
        items.push({
          clientKey:effectiveClientKey,
          tempId:incomingTempId || state.tempId || '',
          food:state.food
        });

        continue;
      }

      const id=String(ch.id||'');

      if(type==='update') {
        const f=validateFoodPayload_(ch.food||{});
        let req=getRequest_(requestId,requestAction);
        let state=req && req.result ? req.result : null;
        let rowIndex=findFoodRow_(values,idx,familyId,id);

        if(rowIndex<1) {
          conflicts.push({
            clientKey:effectiveClientKey,
            id,
            reason:'not_found'
          });
          continue;
        }

        let current=foodFromValuesRow_(values,rowIndex,idx);

        if(state && state.familyId && state.familyId!==familyId) {
          conflicts.push({
            clientKey:effectiveClientKey,
            id,
            reason:'request_family_mismatch'
          });
          continue;
        }

        if(state && state.food && state.food.id===id) {
          const lastApplied=String(state.food.updatedAt||'');

          if(current.updatedAt===lastApplied) {
            if(!foodPayloadMatches_(current,f)) {
              validateImageReferenceForChange_(
                user,
                familyId,
                f.imageId,
                id
              );

              const oldImageId=current.imageId||'';
              const now=isoNow_();
              const saved=savedFoodFromPayload_(
                id,
                f,
                current.createAt,
                now
              );

              state.expectedUpdatedAt=current.updatedAt||'';
              state.oldImageId=oldImageId;
              state.food=saved;

              saveRequest_(
                requestId,
                requestAction,
                'processing',
                state
              );

              writeFoodRow_(
                sh,
                values,
                rowIndex,
                idx,
                f,
                now
              );

              if(f.imageId!==oldImageId) {
                if(f.imageId) {
                  attachImageToFood_(
                    f.imageId,
                    familyId,
                    id,
                    user.userId
                  );
                }

                if(oldImageId) {
                  markImagePendingDelete_(
                    oldImageId,
                    'food_image_replaced'
                  );
                }
              }

              current=saved;

              logActivity_(
                familyId,
                user.userId,
                'food_update',
                id,
                f.name
              );
            } else {
              // 同一個 update 的安全重播。
              if(current.imageId) {
                attachImageToFood_(
                  current.imageId,
                  familyId,
                  id,
                  user.userId
                );
              }

              if(
                state.oldImageId &&
                state.oldImageId!==current.imageId
              ) {
                markImagePendingDelete_(
                  state.oldImageId,
                  'food_image_replaced'
                );
              }
            }

            state.food=current;
            state.oldImageId='';
            saveRequest_(
              requestId,
              requestAction,
              'done',
              state
            );

            applied.push({clientKey:effectiveClientKey});
            items.push({clientKey:effectiveClientKey,food:current});
            continue;
          }

          // processing 還沒真正寫入 Foods 時，current 仍會是 expectedUpdatedAt。
          if(
            req &&
            req.status==='processing' &&
            String(state.expectedUpdatedAt||'')===current.updatedAt
          ) {
            // 繼續走到下方，用目前前端最新 payload 完成同一筆 update。
          } else {
            conflicts.push({
              clientKey:effectiveClientKey,
              id,
              reason:'modified_elsewhere'
            });
            continue;
          }
        } else {
          const expected=String(ch.expectedUpdatedAt||'');

          if(
            expected &&
            current.updatedAt &&
            expected!==current.updatedAt
          ) {
            conflicts.push({
              clientKey:effectiveClientKey,
              id,
              reason:'modified_elsewhere'
            });
            continue;
          }
        }

        validateImageReferenceForChange_(
          user,
          familyId,
          f.imageId,
          id
        );

        const oldImageId=current.imageId||'';
        const now=isoNow_();
        const saved=savedFoodFromPayload_(
          id,
          f,
          current.createAt,
          now
        );

        state={
          familyId,
          id,
          expectedUpdatedAt:current.updatedAt||'',
          oldImageId,
          food:saved
        };

        saveRequest_(
          requestId,
          requestAction,
          'processing',
          state
        );

        writeFoodRow_(
          sh,
          values,
          rowIndex,
          idx,
          f,
          now
        );

        if(f.imageId!==oldImageId) {
          if(f.imageId) {
            attachImageToFood_(
              f.imageId,
              familyId,
              id,
              user.userId
            );
          }

          if(oldImageId) {
            markImagePendingDelete_(
              oldImageId,
              'food_image_replaced'
            );
          }
        }

        state.oldImageId='';
        saveRequest_(
          requestId,
          requestAction,
          'done',
          state
        );

        applied.push({clientKey:effectiveClientKey});
        items.push({clientKey:effectiveClientKey,food:saved});

        logActivity_(
          familyId,
          user.userId,
          'food_update',
          id,
          f.name
        );

        continue;
      }

      if(type==='delete') {
        let req=getRequest_(requestId,requestAction);
        let state=req && req.result ? req.result : null;
        let rowIndex=findFoodRow_(values,idx,familyId,id);

        if(state && state.familyId && state.familyId!==familyId) {
          conflicts.push({
            clientKey:effectiveClientKey,
            id,
            reason:'request_family_mismatch'
          });
          continue;
        }

        if(req && req.status==='done') {
          // delete 成功後重送：即使 Foods 已找不到，也視為同一操作的成功重播。
          applied.push({clientKey:effectiveClientKey});
          deletedIds.push(id);
          continue;
        }

        if(rowIndex<1) {
          if(req && req.status==='processing' && state) {
            if(state.oldImageId) {
              markImagePendingDelete_(
                state.oldImageId,
                'food_deleted'
              );
            }

            saveRequest_(
              requestId,
              requestAction,
              'done',
              state
            );

            applied.push({clientKey:effectiveClientKey});
            deletedIds.push(id);
            continue;
          }

          conflicts.push({
            clientKey:effectiveClientKey,
            id,
            reason:'not_found'
          });
          continue;
        }

        const current=foodFromValuesRow_(values,rowIndex,idx);
        const expected=state
          ? String(state.expectedUpdatedAt||'')
          : String(ch.expectedUpdatedAt||'');

        if(
          expected &&
          current.updatedAt &&
          expected!==current.updatedAt
        ) {
          conflicts.push({
            clientKey:effectiveClientKey,
            id,
            reason:'modified_elsewhere'
          });
          continue;
        }

        state={
          familyId,
          id,
          expectedUpdatedAt:current.updatedAt||'',
          oldImageId:current.imageId||''
        };

        saveRequest_(
          requestId,
          requestAction,
          'processing',
          state
        );

        if(state.oldImageId) {
          markImagePendingDelete_(
            state.oldImageId,
            'food_deleted'
          );
        }

        sh.deleteRow(rowIndex+1);
        values.splice(rowIndex,1);

        saveRequest_(
          requestId,
          requestAction,
          'done',
          state
        );

        applied.push({clientKey:effectiveClientKey});
        deletedIds.push(id);

        logActivity_(
          familyId,
          user.userId,
          'food_delete',
          id,
          current.name
        );

        continue;
      }

      conflicts.push({
        clientKey:effectiveClientKey,
        id,
        reason:'unknown_type'
      });
    }

    return {
      ok:true,
      applied,
      items,
      deletedIds,
      conflicts
    };
  } finally {
    lock.releaseLock();
  }
}

function normalizeFoodStatus_(value) {
  return String(value||'').toLowerCase()==='consumed'
    ? 'consumed'
    : 'active';
}

function validateFoodPayload_(f) {
  const name=String(f.name||'').trim();
  const expiry=String(f.expiry||'').trim();
  const status=normalizeFoodStatus_(f.status);
  const rawQty=Number(f.qty);
  const qty=status==='consumed'
    ? Math.max(0,Number.isFinite(rawQty) ? rawQty : 0)
    : Math.max(1,Number.isFinite(rawQty) ? rawQty : 1);
  const consumedAt=status==='consumed'
    ? String(f.consumedAt||'').trim()
    : '';

  if(!name) throw new Error('食品名稱不可空白');
  if(!/^\d{4}-\d{2}-\d{2}$/.test(expiry)) throw new Error('到期日格式不正確');
  if(status==='consumed' && !consumedAt) throw new Error('已用完食品缺少 consumedAt');

  return {
    name,
    qty,
    location:String(f.location||''),
    expiry,
    note:String(f.note||''),
    notifyMode:['inherit','custom','off'].includes(String(f.notifyMode||'inherit')) ? String(f.notifyMode||'inherit') : 'inherit',
    notifyDaysBefore:String(f.notifyMode||'inherit')==='custom' ? Math.max(0,Math.min(365,Number(f.notifyDaysBefore||0))) : '',
    imageId:String(f.imageId||'').trim(),
    status,
    consumedAt
  };
}

function findFoodRow_(values,idx,familyId,id) {
  for(let r=1;r<values.length;r++) {
    if(String(values[r][idx.familyId])===familyId &&
       String(values[r][idx.id])===id) return r;
  }
  return -1;
}

function foodFromValuesRow_(values,rowIndex,idx) {
  return {
    id:String(values[rowIndex][idx.id]||''),
    name:String(values[rowIndex][idx.name]||''),
    qty:Number(values[rowIndex][idx.qty]||1),
    location:String(values[rowIndex][idx.location]||''),
    expiry:normalizeCell_(values[rowIndex][idx.expiry],'expiry'),
    note:String(values[rowIndex][idx.note]||''),
    createAt:normalizeCell_(values[rowIndex][idx.createAt],'createAt'),
    updatedAt:normalizeCell_(values[rowIndex][idx.updatedAt],'updatedAt'),
    notifyMode:String(values[rowIndex][idx.notifyMode]||'inherit'),
    notifyDaysBefore:String(values[rowIndex][idx.notifyMode]||'inherit')==='custom'
      ? Number(values[rowIndex][idx.notifyDaysBefore]||0)
      : '',
    imageId:String(values[rowIndex][idx.imageId]||''),
    status:normalizeFoodStatus_(values[rowIndex][idx.status]),
    consumedAt:String(values[rowIndex][idx.consumedAt]||'')
  };
}

function savedFoodFromPayload_(id,f,createAt,updatedAt) {
  return {
    id:String(id||''),
    name:f.name,
    qty:f.qty,
    location:f.location,
    expiry:f.expiry,
    note:f.note,
    createAt:String(createAt||''),
    updatedAt:String(updatedAt||''),
    notifyMode:f.notifyMode,
    notifyDaysBefore:f.notifyDaysBefore,
    imageId:f.imageId,
    status:f.status,
    consumedAt:f.consumedAt
  };
}

function foodPayloadMatches_(food,f) {
  return (
    String(food.name||'')===String(f.name||'') &&
    Number(food.qty||1)===Number(f.qty||1) &&
    String(food.location||'')===String(f.location||'') &&
    String(food.expiry||'')===String(f.expiry||'') &&
    String(food.note||'')===String(f.note||'') &&
    String(food.notifyMode||'inherit')===String(f.notifyMode||'inherit') &&
    String(food.notifyDaysBefore??'')===String(f.notifyDaysBefore??'') &&
    String(food.imageId||'')===String(f.imageId||'') &&
    normalizeFoodStatus_(food.status)===normalizeFoodStatus_(f.status) &&
    String(food.consumedAt||'')===String(f.consumedAt||'')
  );
}

function writeFoodRow_(sh,values,rowIndex,idx,f,updatedAt) {
  sh.getRange(rowIndex+1,idx.name+1).setValue(f.name);
  sh.getRange(rowIndex+1,idx.qty+1).setValue(f.qty);
  sh.getRange(rowIndex+1,idx.location+1).setValue(f.location);
  sh.getRange(rowIndex+1,idx.expiry+1).setValue(f.expiry);
  sh.getRange(rowIndex+1,idx.note+1).setValue(f.note);
  sh.getRange(rowIndex+1,idx.notifyMode+1).setValue(f.notifyMode);
  sh.getRange(rowIndex+1,idx.notifyDaysBefore+1)
    .setValue(f.notifyDaysBefore);
  sh.getRange(rowIndex+1,idx.imageId+1).setValue(f.imageId);
  sh.getRange(rowIndex+1,idx.status+1).setValue(f.status);
  sh.getRange(rowIndex+1,idx.consumedAt+1).setValue(f.consumedAt);
  sh.getRange(rowIndex+1,idx.updatedAt+1).setValue(updatedAt);

  values[rowIndex][idx.name]=f.name;
  values[rowIndex][idx.qty]=f.qty;
  values[rowIndex][idx.location]=f.location;
  values[rowIndex][idx.expiry]=f.expiry;
  values[rowIndex][idx.note]=f.note;
  values[rowIndex][idx.notifyMode]=f.notifyMode;
  values[rowIndex][idx.notifyDaysBefore]=f.notifyDaysBefore;
  values[rowIndex][idx.imageId]=f.imageId;
  values[rowIndex][idx.status]=f.status;
  values[rowIndex][idx.consumedAt]=f.consumedAt;
  values[rowIndex][idx.updatedAt]=updatedAt;
}

function foodForClient_(x) {
  return {
    id:String(x.id),
    name:String(x.name),
    qty:Number(x.qty||1),
    location:String(x.location||''),
    expiry:formatDate_(x.expiry),
    note:String(x.note||''),
    createAt:formatDate_(x.createAt),
    updatedAt:String(x.updatedAt||''),
    notifyMode:String(x.notifyMode||'inherit'),
    notifyDaysBefore:String(x.notifyMode||'inherit')==='custom' ? Number(x.notifyDaysBefore||0) : '',
    imageId:String(x.imageId||''),
    status:normalizeFoodStatus_(x.status),
    consumedAt:String(x.consumedAt||'')
  };
}


// ---------------- DRIVE IMAGE STORAGE / BACKUP ----------------

function setupStorageAndMaintenance() {
  setupDatabase();

  const folders=getOrCreateStorageFolders_();

  // 避免重複建立同一個每週 Trigger，並清除舊的續跑狀態。
  ScriptApp.getProjectTriggers()
    .filter(t =>
      [
        'weeklyBackupAndMaintenance',
        IMAGE_BACKUP_CONTINUATION_HANDLER
      ].includes(t.getHandlerFunction())
    )
    .forEach(t => ScriptApp.deleteTrigger(t));

  PropertiesService.getScriptProperties()
    .deleteProperty(IMAGE_BACKUP_RUN_STATE_KEY);

  ScriptApp.newTrigger('weeklyBackupAndMaintenance')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(3)
    .create();

  const result={
    ok:true,
    rootFolderId:folders.root.getId(),
    imageFolderId:folders.images.getId(),
    imageBackupFolderId:folders.imageBackups.getId(),
    backupFolderId:folders.backups.getId(),
    message:'已建立 Drive 儲存資料夾、增量圖片備份與每週備份 / 清理 Trigger'
  };

  Logger.log(JSON.stringify(result));
  return result;
}

function weeklyBackupAndMaintenance() {
  return runWeeklyBackupAndMaintenance_();
}

// 由一次性 Trigger 呼叫。圖片尚未全部完成時會自動建立下一個續跑 Trigger。
function continueWeeklyBackupAndMaintenance() {
  return runWeeklyBackupAndMaintenance_();
}

function newImageBackupRunState_() {
  return {
    version:1,
    runId:Utilities.getUuid(),
    startedAt:isoNow_(),
    phase:'images',
    nextIndex:0,
    processed:0,
    created:0,
    alreadyBackedUp:0,
    errors:0,
    totalCandidates:0,
    skippedUnattachedPending:0,
    failureRetries:0,
    lastError:'',
    imageBackup:null,
    results:{}
  };
}

function loadImageBackupRunState_() {
  const raw=PropertiesService.getScriptProperties()
    .getProperty(IMAGE_BACKUP_RUN_STATE_KEY);

  if(!raw) return newImageBackupRunState_();

  try {
    const state=JSON.parse(raw);
    if(
      state &&
      Number(state.version)===1 &&
      ['images','maintenance'].includes(String(state.phase||''))
    ) {
      state.results=state.results||{};
      return state;
    }
  } catch(err) {
    console.warn('Invalid image backup state: '+String(err.message||err));
  }

  return newImageBackupRunState_();
}

function saveImageBackupRunState_(state) {
  PropertiesService.getScriptProperties().setProperty(
    IMAGE_BACKUP_RUN_STATE_KEY,
    JSON.stringify(state)
  );
}

function deleteImageBackupContinuationTriggers_() {
  ScriptApp.getProjectTriggers()
    .filter(t =>
      t.getHandlerFunction()===IMAGE_BACKUP_CONTINUATION_HANDLER
    )
    .forEach(t => ScriptApp.deleteTrigger(t));
}

function scheduleImageBackupContinuation_() {
  // 每次只保留一個一次性續跑 Trigger，避免手動執行與排程重疊時累積。
  deleteImageBackupContinuationTriggers_();
  ScriptApp.newTrigger(IMAGE_BACKUP_CONTINUATION_HANDLER)
    .timeBased()
    .after(IMAGE_BACKUP_CONTINUATION_DELAY_MS)
    .create();
}

function runWeeklyBackupAndMaintenance_() {
  const executionStartedAt=Date.now();
  const lock=LockService.getScriptLock();

  if(!lock.tryLock(5000)) {
    scheduleImageBackupContinuation_();
    const busy={
      ok:true,
      complete:false,
      busy:true,
      continuationScheduled:true,
      message:'另一個備份程序仍在執行，已安排稍後續跑'
    };
    Logger.log(JSON.stringify(busy));
    return busy;
  }

  // 若本次正是一次性 Trigger，先刪除已觸發的 Trigger；未完成時會再建一個。
  deleteImageBackupContinuationTriggers_();
  let state=loadImageBackupRunState_();

  try {
    if(state.phase==='images') {
      const batch=backupFoodImagesBatch_(state,executionStartedAt);
      state=batch.state;
      saveImageBackupRunState_(state);

      if(!batch.complete) {
        scheduleImageBackupContinuation_();
        const pending={
          ok:state.errors===0,
          complete:false,
          runId:state.runId,
          imageBackup:batch.summary,
          continuationScheduled:true,
          message:'圖片尚未全部完成，已保存進度並安排續跑'
        };
        Logger.log(JSON.stringify(pending));
        return pending;
      }

      state.phase='maintenance';
      state.imageBackup=batch.summary;
      saveImageBackupRunState_(state);
    }

    // 每完成一個維護階段就保存結果。若服務暫時錯誤，下次續跑不會
    // 重複建立已經成功的資料庫備份。
    if(!state.results.backup) {
      state.results.backup=createDatabaseBackup_();
      saveImageBackupRunState_(state);
    }

    if(!state.results.cleanup) {
      state.results.cleanup=cleanupTransientData_();
      saveImageBackupRunState_(state);
    }

    if(!state.results.imageCleanup) {
      state.results.imageCleanup=state.imageBackup.errors===0
        ? cleanupFoodImageMetadata_()
        : {
            skipped:true,
            reason:'圖片備份有錯誤，為避免刪除未備份圖片，本次略過圖片清理'
          };
      saveImageBackupRunState_(state);
    }

    if(!state.results.imageBackupCleanup) {
      state.results.imageBackupCleanup=cleanupExpiredImageBackups_();
      saveImageBackupRunState_(state);
    }

    if(!state.results.oldBackups) {
      state.results.oldBackups=cleanupOldBackups_();
      saveImageBackupRunState_(state);
    }

    const result={
      ok:state.imageBackup.errors===0,
      complete:true,
      runId:state.runId,
      imageBackup:state.imageBackup,
      backup:state.results.backup,
      cleanup:state.results.cleanup,
      imageCleanup:state.results.imageCleanup,
      imageBackupCleanup:state.results.imageBackupCleanup,
      oldBackups:state.results.oldBackups
    };

    PropertiesService.getScriptProperties()
      .deleteProperty(IMAGE_BACKUP_RUN_STATE_KEY);
    deleteImageBackupContinuationTriggers_();
    Logger.log(JSON.stringify(result));
    return result;
  } catch(err) {
    // 服務偶發錯誤時保留目前進度並自動重試。即使這裡建立 Trigger
    // 失敗，下一次每週排程仍會讀取相同狀態繼續。
    state.failureRetries=Number(state.failureRetries||0)+1;
    state.lastError=String(err.message||err);
    state.lastFailedAt=isoNow_();
    saveImageBackupRunState_(state);

    // 避免權限或配額等永久性錯誤造成無限 Trigger 迴圈；下週排程
    // 仍會保留狀態並再次嘗試，也可在修正問題後手動執行主程序。
    if(state.failureRetries<=IMAGE_BACKUP_MAX_FAILURE_RETRIES) {
      try {
        scheduleImageBackupContinuation_();
      } catch(triggerErr) {
        console.error(
          'Unable to schedule image backup continuation: '+
          String(triggerErr.message||triggerErr)
        );
      }
    }
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function createDatabaseBackup_() {
  const props=PropertiesService.getScriptProperties();

  const ssId=
    props.getProperty('DATABASE_SPREADSHEET_ID') ||
    SpreadsheetApp.getActiveSpreadsheet().getId();

  props.setProperty('DATABASE_SPREADSHEET_ID',ssId);

  const ss=SpreadsheetApp.openById(ssId);
  const source=DriveApp.getFileById(ssId);
  const folders=getOrCreateStorageFolders_();

  const tz=
    ss.getSpreadsheetTimeZone() ||
    Session.getScriptTimeZone() ||
    'Asia/Taipei';

  const stamp=Utilities.formatDate(
    new Date(),
    tz,
    'yyyyMMdd_HHmmss'
  );

  const copy=source.makeCopy(
    `FoodExpiryBackup_${stamp}`,
    folders.backups
  );

  return {
    fileId:copy.getId(),
    name:copy.getName()
  };
}

function cleanupOldBackups_() {
  const folders=getOrCreateStorageFolders_();
  const cutoff=
    Date.now() -
    BACKUP_RETENTION_DAYS*86400000;

  const files=folders.backups.getFiles();
  let trashed=0;

  while(files.hasNext()) {
    const file=files.next();

    if(file.getDateCreated().getTime()<cutoff) {
      file.setTrashed(true);
      trashed++;
    }
  }

  return {
    trashed,
    retentionDays:BACKUP_RETENTION_DAYS
  };
}

function imageBackupExtension_(mimeType) {
  const mime=String(mimeType||'').toLowerCase();
  if(mime==='image/webp') return 'webp';
  if(mime==='image/png') return 'png';
  return 'jpg';
}

function getAvailableDriveFile_(fileId) {
  const id=String(fileId||'');
  if(!id) return null;

  try {
    const file=DriveApp.getFileById(id);
    return file.isTrashed() ? null : file;
  } catch {
    return null;
  }
}

function createImageBackupContext_() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();

  ensureSheet_(
    ss,
    SHEETS.IMAGE_BACKUPS,
    IMAGE_BACKUP_LOG_HEADERS
  );

  const sh=ss.getSheetByName(SHEETS.IMAGE_BACKUPS);
  ensureHeaders_(sh,IMAGE_BACKUP_LOG_HEADERS);

  const values=sh.getDataRange().getValues();
  const headers=values[0].map(String);
  const idx={};
  headers.forEach((h,i)=>idx[h]=i);

  const entries={};
  for(let r=1;r<values.length;r++) {
    const imageId=String(values[r][idx.imageId]||'');
    if(!imageId) continue;

    entries[imageId]={
      rowNumber:r+1,
      values:values[r].slice()
    };
  }

  return {
    sh,
    headers,
    idx,
    entries,
    folders:getOrCreateStorageFolders_()
  };
}

function writeImageBackupLog_(ctx,meta,details) {
  const imageId=String(meta.imageId||'');
  if(!imageId) throw new Error('圖片備份缺少 imageId');

  const existing=ctx.entries[imageId]||null;
  const row=existing
    ? existing.values.slice()
    : new Array(ctx.headers.length).fill('');

  const set=(name,value)=>{
    const i=ctx.idx[name];
    if(i!=null && i>=0) row[i]=value;
  };

  set('imageId',imageId);
  set('familyId',String(meta.familyId||''));
  set('foodId',String(meta.foodId||''));
  set('sourceDriveFileId',String(meta.driveFileId||''));
  set('backupFileId',String(details.backupFileId||''));
  set('mimeType',String(meta.mimeType||''));
  set('sizeBytes',Math.max(0,Number(meta.sizeBytes||0)));
  set('sourceStatus',String(meta.status||''));
  set('backedUpAt',String(details.backedUpAt||''));
  set('updatedAt',isoNow_());
  set(
    'deletedAt',
    String(meta.status||'')==='pending_delete'
      ? String(meta.deletedAt||details.deletedAt||'')
      : ''
  );
  set('lastError',String(details.lastError||''));

  if(existing) {
    ctx.sh.getRange(
      existing.rowNumber,
      1,
      1,
      row.length
    ).setValues([row]);

    existing.values=row;
  } else {
    ctx.sh.appendRow(row);
    ctx.entries[imageId]={
      rowNumber:ctx.sh.getLastRow(),
      values:row
    };
  }
}

function imageBackupLogIsCurrent_(ctx,existing,meta,details) {
  if(!existing) return false;

  const value=name=>String(existing.values[ctx.idx[name]]||'');
  const expectedDeletedAt=String(meta.status||'')==='pending_delete'
    ? String(meta.deletedAt||details.deletedAt||'')
    : '';

  return (
    value('imageId')===String(meta.imageId||'') &&
    value('familyId')===String(meta.familyId||'') &&
    value('foodId')===String(meta.foodId||'') &&
    value('sourceDriveFileId')===String(meta.driveFileId||'') &&
    value('backupFileId')===String(details.backupFileId||'') &&
    value('mimeType')===String(meta.mimeType||'') &&
    Number(value('sizeBytes')||0)===Math.max(0,Number(meta.sizeBytes||0)) &&
    value('sourceStatus')===String(meta.status||'') &&
    Boolean(value('backedUpAt')) &&
    value('deletedAt')===expectedDeletedAt &&
    !value('lastError')
  );
}

function ensureFoodImageBackup_(meta,context) {
  const ctx=context||createImageBackupContext_();
  const imageId=String(meta.imageId||'');
  const sourceFileId=String(meta.driveFileId||'');

  if(!imageId || !sourceFileId) {
    throw new Error('圖片備份資料不完整');
  }

  const existing=ctx.entries[imageId]||null;
  const existingBackupId=existing
    ? String(existing.values[ctx.idx.backupFileId]||'')
    : '';
  const existingBackedUpAt=existing
    ? String(existing.values[ctx.idx.backedUpAt]||'')
    : '';

  let backupFile=getAvailableDriveFile_(existingBackupId);
  let created=false;

  try {
    if(!backupFile) {
      // 原圖即使已在垃圾桶，只要尚未永久刪除，仍嘗試讀出並建立備份。
      const sourceFile=DriveApp.getFileById(sourceFileId);
      const blob=sourceFile.getBlob();

      backupFile=ctx.folders.imageBackups.createFile(blob);
      backupFile.setName(
        `backup_${imageId}.${imageBackupExtension_(
          meta.mimeType||blob.getContentType()
        )}`
      );
      created=true;
    }

    const backedUpAt=created || !existingBackedUpAt
      ? isoNow_()
      : existingBackedUpAt;

    const details={
      backupFileId:backupFile.getId(),
      backedUpAt,
      deletedAt:meta.deletedAt||'',
      lastError:''
    };
    const logUpdated=
      !imageBackupLogIsCurrent_(ctx,existing,meta,details);

    // 已存在且 metadata 沒變時不重寫試算表，降低大量圖片的服務呼叫數。
    if(logUpdated) writeImageBackupLog_(ctx,meta,details);

    return {
      imageId,
      backupFileId:backupFile.getId(),
      created,
      logUpdated
    };
  } catch(err) {
    writeImageBackupLog_(ctx,meta,{
      backupFileId:backupFile ? backupFile.getId() : '',
      backedUpAt:existingBackedUpAt,
      deletedAt:meta.deletedAt||'',
      lastError:String(err.message||err)
    });
    throw err;
  }
}

function backupFoodImagesBatch_(state,executionStartedAt) {
  const ctx=createImageBackupContext_();
  const images=rows_(SHEETS.FOOD_IMAGES);
  const candidates=images.filter(image=>
    String(image.status||'')!=='pending' ||
    Boolean(String(image.foodId||''))
  );
  const startedAt=Number(executionStartedAt||Date.now());
  let nextIndex=Math.max(0,Number(state.nextIndex||0));
  nextIndex=Math.min(nextIndex,candidates.length);
  let batchProcessed=0;

  state.totalCandidates=candidates.length;
  state.skippedUnattachedPending=images.length-candidates.length;

  while(nextIndex<candidates.length) {
    if(
      batchProcessed>=IMAGE_BACKUP_MAX_ITEMS_PER_BATCH ||
      (
        batchProcessed>0 &&
        Date.now()-startedAt>=IMAGE_BACKUP_BATCH_TIME_BUDGET_MS
      )
    ) {
      break;
    }

    const image=candidates[nextIndex];

    if(image.imageId && image.driveFileId) {
      try {
        const result=ensureFoodImageBackup_(image,ctx);
        if(result.created) state.created++;
        else state.alreadyBackedUp++;
      } catch(err) {
        state.errors++;
        console.error(
          'Image backup failed '+String(image.imageId)+': '+
          String(err.message||err)
        );
      }
    }

    nextIndex++;
    batchProcessed++;
    state.processed++;
    state.nextIndex=nextIndex;

    // 週期性存檔可防範 Apps Script 或 Drive 服務意外中止；重跑仍具冪等性。
    if(batchProcessed%10===0) saveImageBackupRunState_(state);
  }

  state.nextIndex=nextIndex;
  const complete=nextIndex>=candidates.length;
  const summary={
    ok:state.errors===0,
    complete,
    checked:state.processed,
    totalCandidates:candidates.length,
    remaining:Math.max(0,candidates.length-nextIndex),
    batchProcessed,
    skippedUnattachedPending:state.skippedUnattachedPending,
    created:state.created,
    alreadyBackedUp:state.alreadyBackedUp,
    errors:state.errors,
    folderId:ctx.folders.imageBackups.getId(),
    startedAt:state.startedAt
  };

  return {state,summary,complete};
}

// 相容既有的手動測試入口；正式排程請由 weeklyBackupAndMaintenance 啟動，
// 才會在未完成時自動建立續跑 Trigger。
function backupFoodImages_() {
  return backupFoodImagesBatch_(
    newImageBackupRunState_(),
    Date.now()
  ).summary;
}

function cleanupExpiredImageBackups_() {
  const ctx=createImageBackupContext_();
  const values=ctx.sh.getDataRange().getValues();
  const cutoff=
    Date.now()-IMAGE_BACKUP_RETENTION_DAYS*86400000;
  let trashed=0;
  let rowsRemoved=0;

  for(let r=values.length-1;r>=1;r--) {
    const status=String(
      values[r][ctx.idx.sourceStatus]||''
    );
    const deletedAt=toMillis_(
      values[r][ctx.idx.deletedAt]
    );

    if(
      status!=='pending_delete' ||
      !deletedAt ||
      deletedAt>=cutoff
    ) {
      continue;
    }

    const backupFileId=String(
      values[r][ctx.idx.backupFileId]||''
    );

    if(backupFileId) {
      try {
        const file=DriveApp.getFileById(backupFileId);
        if(!file.isTrashed()) {
          file.setTrashed(true);
          trashed++;
        }
      } catch(err) {
        // 檔案若已不存在，仍移除過期 Log；其他錯誤會留在執行記錄。
        console.warn(
          'Expired image backup cleanup '+backupFileId+': '+
          String(err.message||err)
        );
      }
    }

    ctx.sh.deleteRow(r+1);
    rowsRemoved++;
  }

  return {
    trashed,
    rowsRemoved,
    retentionDays:IMAGE_BACKUP_RETENTION_DAYS
  };
}

// 災難復原工具：傳入 imageId 可還原單張圖片。
// 會保留備份檔，並在 FoodImages 建立新的正式 Drive 檔案參照。
function restoreFoodImageFromBackup(imageId) {
  imageId=String(imageId||'').trim();
  if(!imageId) throw new Error('請提供 imageId');

  const ctx=createImageBackupContext_();
  const entry=ctx.entries[imageId];
  if(!entry) throw new Error('找不到這張圖片的備份記錄');

  const backupFileId=String(
    entry.values[ctx.idx.backupFileId]||''
  );
  const backupFile=getAvailableDriveFile_(backupFileId);
  if(!backupFile) throw new Error('圖片備份檔不存在或已刪除');

  const sh=sheet_(SHEETS.FOOD_IMAGES);
  const values=sh.getDataRange().getValues();
  const headers=values[0].map(String);
  const idx={};
  headers.forEach((h,i)=>idx[h]=i);

  for(let r=1;r<values.length;r++) {
    if(String(values[r][idx.imageId]||'')!==imageId) continue;

    const currentFileId=String(
      values[r][idx.driveFileId]||''
    );
    const currentFile=getAvailableDriveFile_(currentFileId);

    if(currentFile) {
      return {
        ok:true,
        restored:false,
        reason:'原始圖片仍可使用',
        imageId,
        driveFileId:currentFile.getId()
      };
    }

    const mimeType=String(
      values[r][idx.mimeType] ||
      entry.values[ctx.idx.mimeType] ||
      backupFile.getBlob().getContentType() ||
      'image/jpeg'
    );
    const familyId=String(values[r][idx.familyId]||'');
    const restoredFile=ctx.folders.images.createFile(
      backupFile.getBlob()
    );

    restoredFile.setName(
      `food_${familyId}_${imageId}.${imageBackupExtension_(mimeType)}`
    );

    const referenced=rows_(SHEETS.FOODS).some(f=>
      String(f.imageId||'')===imageId
    );
    const status=referenced ? 'active' : 'pending';
    const now=isoNow_();

    sh.getRange(r+1,idx.driveFileId+1)
      .setValue(restoredFile.getId());
    sh.getRange(r+1,idx.status+1)
      .setValue(status);
    sh.getRange(r+1,idx.updatedAt+1)
      .setValue(now);
    sh.getRange(r+1,idx.deletedAt+1)
      .setValue('');

    const meta={};
    headers.forEach((h,i)=>{
      meta[h]=normalizeCell_(values[r][i],h);
    });
    meta.driveFileId=restoredFile.getId();
    meta.status=status;
    meta.deletedAt='';

    writeImageBackupLog_(ctx,meta,{
      backupFileId,
      backedUpAt:String(
        entry.values[ctx.idx.backedUpAt]||now
      ),
      deletedAt:'',
      lastError:''
    });

    return {
      ok:true,
      restored:true,
      imageId,
      driveFileId:restoredFile.getId(),
      status
    };
  }

  throw new Error('FoodImages 找不到這張圖片的 metadata');
}

// 可直接從 Apps Script 編輯器手動執行；只修復狀態為 active、
// 但正式 Drive 檔案已遺失或在垃圾桶的圖片。
function restoreMissingFoodImagesFromBackups() {
  const missing=rows_(SHEETS.FOOD_IMAGES).filter(image=>
    String(image.status||'')==='active' &&
    !getAvailableDriveFile_(image.driveFileId)
  );
  const results=[];
  let restored=0;
  let errors=0;

  for(const image of missing) {
    try {
      const result=restoreFoodImageFromBackup(image.imageId);
      results.push(result);
      if(result.restored) restored++;
    } catch(err) {
      errors++;
      results.push({
        ok:false,
        imageId:String(image.imageId||''),
        error:String(err.message||err)
      });
    }
  }

  const summary={
    ok:errors===0,
    missing:missing.length,
    restored,
    errors,
    results
  };

  Logger.log(JSON.stringify(summary));
  return summary;
}

function cleanupTransientData_() {
  const now=Date.now();

  const result={};

  result.sessions=deleteRowsWhere_(
    SHEETS.SESSIONS,
    row => {
      const expires=toMillis_(row.expiresAt);
      return expires &&
        expires <
          now -
          SESSION_CLEANUP_GRACE_DAYS*86400000;
    }
  );

  result.passwordResets=deleteRowsWhere_(
    SHEETS.RESETS,
    row => {
      const expires=toMillis_(row.expiresAt);
      const used=toMillis_(row.usedAt);

      if(
        used &&
        used <
          now -
          RESET_CLEANUP_DAYS*86400000
      ) {
        return true;
      }

      return expires &&
        expires <
          now -
          RESET_CLEANUP_DAYS*86400000;
    }
  );

  result.requestLog=deleteRowsWhere_(
    SHEETS.REQUESTS,
    row => {
      const time=
        toMillis_(row.updatedAt) ||
        toMillis_(row.createdAt);

      return time &&
        time <
          now -
          REQUEST_LOG_RETENTION_DAYS*86400000;
    }
  );

  result.notificationLog=deleteRowsWhere_(
    SHEETS.NOTIFY_LOG,
    row => {
      const time=toMillis_(row.sentAt);

      return time &&
        time <
          now -
          NOTIFY_LOG_RETENTION_DAYS*86400000;
    }
  );

  // ActivityLog 與 NotificationLog 同樣保留 90 天，
  // 每週維護會先建立完整備份，再刪除超過保留期的資料。
  result.activityLog=deleteRowsWhere_(
    SHEETS.ACTIVITY,
    row => {
      const time=toMillis_(row.createdAt);

      return time &&
        time <
          now -
          ACTIVITY_LOG_RETENTION_DAYS*86400000;
    }
  );

  return result;
}

function cleanupFoodImageMetadata_() {
  const sh=sheet_(SHEETS.FOOD_IMAGES);
  const values=sh.getDataRange().getValues();

  if(values.length<2) {
    return {
      orphanImagesTrashed:0,
      metadataRowsRemoved:0
    };
  }

  const headers=values[0].map(String);
  const idx={};
  headers.forEach((h,i)=>idx[h]=i);

  const now=Date.now();
  let orphanImagesTrashed=0;
  let metadataRowsRemoved=0;

  for(let r=values.length-1;r>=1;r--) {
    const status=String(
      values[r][idx.status]||''
    );

    const created=toMillis_(
      values[r][idx.createdAt]
    );

    const deleted=toMillis_(
      values[r][idx.deletedAt]
    );

    if(
      status==='pending' &&
      created &&
      created <
        now -
        PENDING_IMAGE_RETENTION_DAYS*86400000
    ) {
      const imageId=String(
        values[r][idx.imageId]||''
      );

      markImagePendingDelete_(
        imageId,
        'orphan_pending_cleanup'
      );

      orphanImagesTrashed++;
      continue;
    }

    if(
      status==='pending_delete' &&
      deleted &&
      deleted <
        now -
        DELETED_IMAGE_META_RETENTION_DAYS*86400000
    ) {
      sh.deleteRow(r+1);
      metadataRowsRemoved++;
    }
  }

  return {
    orphanImagesTrashed,
    metadataRowsRemoved
  };
}

function uploadFoodImage_(user,p) {
  const familyId=String(p.familyId||'');
  requireEditableMembership_(user.userId,familyId);

  const mimeType=String(
    p.mimeType||'image/jpeg'
  ).toLowerCase();

  if(
    ![
      'image/jpeg',
      'image/jpg',
      'image/webp',
      'image/png'
    ].includes(mimeType)
  ) {
    throw new Error('不支援的圖片格式');
  }

  const base64=String(p.imageBase64||'');

  if(!base64) {
    throw new Error('圖片資料不可空白');
  }

  let bytes;

  try {
    bytes=Utilities.base64Decode(base64);
  } catch {
    throw new Error('圖片資料格式錯誤');
  }

  if(!bytes.length) {
    throw new Error('圖片資料不可空白');
  }

  if(bytes.length>MAX_PRODUCT_IMAGE_BYTES) {
    throw new Error(
      '圖片壓縮後仍太大，請重新選擇圖片'
    );
  }

  const imageId=Utilities.getUuid();
  const folders=getOrCreateStorageFolders_();

  const ext=
    mimeType==='image/webp'
      ? 'webp'
      : mimeType==='image/png'
        ? 'png'
        : 'jpg';

  const blob=Utilities.newBlob(
    bytes,
    mimeType,
    `${imageId}.${ext}`
  );

  const file=folders.images.createFile(blob);
  file.setName(
    `food_${familyId}_${imageId}.${ext}`
  );

  const now=isoNow_();

  append_(SHEETS.FOOD_IMAGES,[
    imageId,
    familyId,
    '',
    file.getId(),
    mimeType,
    bytes.length,
    Math.max(0,Number(p.width||0)),
    Math.max(0,Number(p.height||0)),
    'pending',
    user.userId,
    now,
    now,
    ''
  ]);

  return {
    ok:true,
    imageId,
    sizeBytes:bytes.length
  };
}

function getFoodImage_(user,p) {
  const familyId=String(p.familyId||'');
  requireMembership_(user.userId,familyId);

  const imageId=String(p.imageId||'');
  const meta=findImageMeta_(imageId);

  if(
    !meta ||
    meta.familyId!==familyId ||
    meta.status!=='active'
  ) {
    throw new Error('圖片不存在或已刪除');
  }

  const file=DriveApp.getFileById(
    meta.driveFileId
  );

  if(file.isTrashed()) {
    throw new Error('圖片已刪除');
  }

  const blob=file.getBlob();
  const mimeType=
    String(meta.mimeType||blob.getContentType()||'image/jpeg');

  return {
    ok:true,
    imageId,
    mimeType,
    dataUrl:
      `data:${mimeType};base64,`+
      Utilities.base64Encode(blob.getBytes())
  };
}

function discardFoodImage_(user,p) {
  const familyId=String(p.familyId||'');
  requireEditableMembership_(user.userId,familyId);

  const imageId=String(p.imageId||'');
  if(!imageId) return {ok:true};

  const meta=findImageMeta_(imageId);

  if(!meta) return {ok:true};

  if(meta.familyId!==familyId) {
    throw new Error('圖片不屬於目前群組');
  }

  // 只有尚未綁定食品的 pending 圖片可以從前端直接放棄。
  if(
    meta.status!=='pending' ||
    meta.uploadedBy!==user.userId ||
    meta.foodId
  ) {
    throw new Error('這張圖片不能直接放棄');
  }

  markImagePendingDelete_(
    imageId,
    'client_discard'
  );

  return {ok:true};
}

function validateImageReferenceForChange_(
  user,
  familyId,
  imageId,
  foodId
) {
  imageId=String(imageId||'');
  if(!imageId) return;

  const meta=findImageMeta_(imageId);

  if(!meta) {
    throw new Error('找不到指定的食品圖片');
  }

  if(meta.familyId!==familyId) {
    throw new Error('食品圖片不屬於目前群組');
  }

  if(meta.status==='pending') {
    if(meta.uploadedBy!==user.userId) {
      throw new Error('這張待上傳圖片不屬於目前使用者');
    }

    if(meta.foodId) {
      throw new Error('圖片已經綁定其他食品');
    }

    return;
  }

  if(
    meta.status==='active' &&
    foodId &&
    meta.foodId===foodId
  ) {
    return;
  }

  throw new Error('食品圖片狀態不正確');
}

function attachImageToFood_(
  imageId,
  familyId,
  foodId,
  userId
) {
  if(!imageId) return;

  const sh=sheet_(SHEETS.FOOD_IMAGES);
  const values=sh.getDataRange().getValues();
  const headers=values[0].map(String);

  const idx={};
  headers.forEach((h,i)=>idx[h]=i);

  for(let r=1;r<values.length;r++) {
    if(
      String(values[r][idx.imageId])===imageId
    ) {
      const rowFamily=String(
        values[r][idx.familyId]||''
      );

      if(rowFamily!==familyId) {
        throw new Error('圖片群組不符');
      }

      const status=String(
        values[r][idx.status]||''
      );

      const existingFoodId=String(
        values[r][idx.foodId]||''
      );

      if(
        status==='active' &&
        existingFoodId===foodId
      ) {
        return;
      }

      if(status!=='pending') {
        throw new Error('圖片無法綁定食品');
      }

      const uploader=String(
        values[r][idx.uploadedBy]||''
      );

      if(uploader!==userId) {
        throw new Error('圖片上傳者不符');
      }

      const now=isoNow_();

      sh.getRange(r+1,idx.foodId+1)
        .setValue(foodId);

      sh.getRange(r+1,idx.status+1)
        .setValue('active');

      sh.getRange(r+1,idx.updatedAt+1)
        .setValue(now);

      sh.getRange(r+1,idx.deletedAt+1)
        .setValue('');

      return;
    }
  }

  throw new Error('找不到指定的食品圖片');
}

function markImagePendingDelete_(
  imageId,
  reason
) {
  if(!imageId) return;

  const sh=sheet_(SHEETS.FOOD_IMAGES);
  const values=sh.getDataRange().getValues();

  if(values.length<2) return;

  const headers=values[0].map(String);
  const idx={};
  headers.forEach((h,i)=>idx[h]=i);

  for(let r=1;r<values.length;r++) {
    if(
      String(values[r][idx.imageId])!==imageId
    ) {
      continue;
    }

    const currentStatus=String(
      values[r][idx.status]||''
    );

    if(currentStatus==='pending_delete') {
      return;
    }

    const fileId=String(
      values[r][idx.driveFileId]||''
    );

    const now=isoNow_();
    const backupMeta={};
    headers.forEach((h,i)=>{
      backupMeta[h]=normalizeCell_(values[r][i],h);
    });
    backupMeta.status='pending_delete';
    backupMeta.deletedAt=now;

    // 正式使用過的圖片在進入垃圾桶前，至少要有一份可用備份。
    // 從未綁定食品、被使用者取消的暫存圖片不占用備份空間。
    const shouldBackup=
      currentStatus==='active' ||
      Boolean(String(values[r][idx.foodId]||''));

    if(shouldBackup) {
      // 若備份失敗就中止刪除，避免只留下無法還原的 metadata。
      ensureFoodImageBackup_(backupMeta);
    }

    if(fileId) {
      try {
        DriveApp.getFileById(fileId)
          .setTrashed(true);
      } catch(err) {
        // 即使檔案已不存在，也要更新 metadata，
        // 避免之後一直重試同一筆。
        Logger.log(
          `markImagePendingDelete_ ${imageId}: ${err}`
        );
      }
    }

    sh.getRange(r+1,idx.status+1)
      .setValue('pending_delete');

    sh.getRange(r+1,idx.updatedAt+1)
      .setValue(now);

    sh.getRange(r+1,idx.deletedAt+1)
      .setValue(now);

    if(reason) {
      Logger.log(
        `image ${imageId} -> pending_delete (${reason})`
      );
    }

    return;
  }
}

function markFamilyImagesPendingDelete_(familyId) {
  rows_(SHEETS.FOOD_IMAGES)
    .filter(x =>
      x.familyId===familyId &&
      (
        x.status==='active' ||
        x.status==='pending'
      )
    )
    .forEach(x =>
      markImagePendingDelete_(
        x.imageId,
        'family_deleted'
      )
    );
}

function findImageMeta_(imageId) {
  imageId=String(imageId||'');

  return rows_(SHEETS.FOOD_IMAGES)
    .find(x => x.imageId===imageId) ||
    null;
}

function getOrCreateStorageFolders_() {
  const props=PropertiesService.getScriptProperties();

  let root=getFolderByProperty_(
    'APP_STORAGE_ROOT_ID'
  );

  if(!root) {
    root=DriveApp.createFolder(
      'FoodExpiryManager'
    );

    props.setProperty(
      'APP_STORAGE_ROOT_ID',
      root.getId()
    );
  }

  let images=getFolderByProperty_(
    'IMAGE_FOLDER_ID'
  );

  if(!images) {
    images=root.createFolder(
      'FoodImages'
    );

    props.setProperty(
      'IMAGE_FOLDER_ID',
      images.getId()
    );
  }

  let imageBackups=getFolderByProperty_(
    'IMAGE_BACKUP_FOLDER_ID'
  );

  if(!imageBackups) {
    imageBackups=root.createFolder(
      'ImageBackups'
    );

    props.setProperty(
      'IMAGE_BACKUP_FOLDER_ID',
      imageBackups.getId()
    );
  }

  let backups=getFolderByProperty_(
    'BACKUP_FOLDER_ID'
  );

  if(!backups) {
    backups=root.createFolder(
      'Backups'
    );

    props.setProperty(
      'BACKUP_FOLDER_ID',
      backups.getId()
    );
  }

  return {
    root,
    images,
    imageBackups,
    backups
  };
}

function getFolderByProperty_(key) {
  const id=PropertiesService
    .getScriptProperties()
    .getProperty(key);

  if(!id) return null;

  try {
    const folder=DriveApp.getFolderById(id);

    if(folder.isTrashed()) {
      return null;
    }

    return folder;
  } catch {
    return null;
  }
}

function deleteRowsWhere_(
  sheetName,
  predicate
) {
  const sh=sheet_(sheetName);
  const values=sh.getDataRange().getValues();

  if(values.length<2) return 0;

  const headers=values[0].map(String);
  let deleted=0;

  for(let r=values.length-1;r>=1;r--) {
    const obj={};

    headers.forEach((h,i)=>{
      obj[h]=normalizeCell_(
        values[r][i],
        h
      );
    });

    if(predicate(obj)) {
      sh.deleteRow(r+1);
      deleted++;
    }
  }

  return deleted;
}

function toMillis_(value) {
  if(!value) return 0;

  if(
    Object.prototype.toString.call(value)==='[object Date]'
  ) {
    return value.getTime();
  }

  const t=new Date(value).getTime();
  return Number.isFinite(t) ? t : 0;
}


// ---------------- MEMBERSHIP / DATA HELPERS ----------------

function userFamilies_(userId) {
  const families=rows_(SHEETS.FAMILIES);
  const memberships=rows_(SHEETS.MEMBERS)
    .filter(x=>x.userId===userId && x.status==='active');

  return memberships.map(m=>{
    const f=families.find(x=>x.familyId===m.familyId);
    if(!f) return null;

    return {
      familyId:f.familyId,
      familyName:f.familyName,
      role:m.role,
      inviteCode:m.role==='owner'?f.inviteCode:''
    };
  }).filter(Boolean);
}

function requireSession_(token) {
  token=String(token||'');
  if(!token) throw new Error('尚未登入');

  const cache=CacheService.getScriptCache();
  const cacheKey='session:'+token;
  const cached=cache.get(cacheKey);

  if(cached) {
    const item=JSON.parse(cached);
    if(Number(item.expiresAt)>Date.now()) return item.user;
    cache.remove(cacheKey);
  }

  const session=rows_(SHEETS.SESSIONS).find(x =>
    x.sessionToken===token && new Date(x.expiresAt).getTime()>Date.now()
  );
  if(!session) throw new Error('登入已過期');

  const user=rows_(SHEETS.USERS).find(x =>
    x.userId===session.userId && x.status==='active'
  );
  if(!user) throw new Error('使用者不存在');

  cacheSession_(token,user,new Date(session.expiresAt).getTime());
  return user;
}

function cacheSession_(token,user,expiresAt) {
  const ttl=Math.max(1,Math.min(
    SESSION_CACHE_SECONDS,
    Math.floor((expiresAt-Date.now())/1000)
  ));

  CacheService.getScriptCache().put(
    'session:'+token,
    JSON.stringify({user:publicUser_(user),expiresAt}),
    ttl
  );
}

function requireMembership_(userId,familyId) {
  const m=rows_(SHEETS.MEMBERS).find(x =>
    x.userId===userId && x.familyId===familyId && x.status==='active'
  );
  if(!m) throw new Error('你不是這個家庭的成員');
  return m;
}

function requireEditableMembership_(userId,familyId) {
  const m=requireMembership_(userId,familyId);
  if(!['owner','member'].includes(m.role)) throw new Error('你沒有編輯權限');
  return m;
}

function requireOwner_(userId,familyId) {
  const m=requireMembership_(userId,familyId);
  if(m.role!=='owner') throw new Error('只有 Owner 可以執行這個操作');
  return m;
}

function updateFamilyCell_(familyId,column,value) {
  const sh=sheet_(SHEETS.FAMILIES);
  const values=sh.getDataRange().getValues();
  const headers=values[0].map(String);

  const idI=headers.indexOf('familyId');
  const colI=headers.indexOf(column);

  for(let r=1;r<values.length;r++) {
    if(String(values[r][idI])===familyId) {
      sh.getRange(r+1,colI+1).setValue(value);
      return;
    }
  }
  throw new Error('家庭不存在');
}

function updateMemberRole_(familyId,userId,role) {
  updateMemberStatusRole_(familyId,userId,'active',role);
}

function updateMemberStatusRole_(familyId,userId,status,role) {
  const sh=sheet_(SHEETS.MEMBERS);
  const values=sh.getDataRange().getValues();
  const headers=values[0].map(String);

  const fi=headers.indexOf('familyId');
  const ui=headers.indexOf('userId');
  const ri=headers.indexOf('role');
  const si=headers.indexOf('status');

  for(let r=1;r<values.length;r++) {
    if(String(values[r][fi])===familyId && String(values[r][ui])===userId) {
      sh.getRange(r+1,ri+1).setValue(role);
      sh.getRange(r+1,si+1).setValue(status);
      return;
    }
  }
  throw new Error('找不到家庭成員');
}

function deleteRowsByValue_(sheetName,column,value) {
  const sh=sheet_(sheetName);
  const values=sh.getDataRange().getValues();
  if(values.length<=1) return;

  const headers=values[0].map(String);
  const ci=headers.indexOf(column);
  if(ci<0) return;

  for(let r=values.length-1;r>=1;r--) {
    if(String(values[r][ci])===String(value)) sh.deleteRow(r+1);
  }
}

function logActivity_(familyId,userId,action,targetId,detail) {
  append_(SHEETS.ACTIVITY,[
    familyId,userId,action,targetId,String(detail||''),isoNow_()
  ]);
}

// ---------------- GENERIC HELPERS ----------------

function updateUserPassword_(userId,salt,hash) {
  const sh=sheet_(SHEETS.USERS);
  const values=sh.getDataRange().getValues();
  const headers=values[0].map(String);

  const idI=headers.indexOf('userId');
  const saltI=headers.indexOf('passwordSalt');
  const hashI=headers.indexOf('passwordHash');

  for(let r=1;r<values.length;r++) {
    if(String(values[r][idI])===userId) {
      sh.getRange(r+1,saltI+1).setValue(salt);
      sh.getRange(r+1,hashI+1).setValue(hash);
      return;
    }
  }
  throw new Error('使用者不存在');
}

function updateUserPasswordVersion_(userId,salt,hash,version) {
  const sh=sheet_(SHEETS.USERS);
  const values=sh.getDataRange().getValues();
  const headers=values[0].map(String);

  const idI=headers.indexOf('userId');
  const saltI=headers.indexOf('passwordSalt');
  const hashI=headers.indexOf('passwordHash');
  const versionI=headers.indexOf('hashVersion');

  if(versionI < 0) {
    throw new Error('Users 缺少 hashVersion 欄位，請先執行 setupDatabase()');
  }

  for(let r=1;r<values.length;r++) {
    if(String(values[r][idI])===userId) {
      sh.getRange(r+1,saltI+1).setValue(salt);
      sh.getRange(r+1,hashI+1).setValue(hash);
      sh.getRange(r+1,versionI+1).setValue(version);
      return;
    }
  }

  throw new Error('使用者不存在');
}

function invalidateUserSessions_(userId) {
  const sh=sheet_(SHEETS.SESSIONS);
  const values=sh.getDataRange().getValues();
  if(values.length<=1) return;

  const headers=values[0].map(String);
  const tokenI=headers.indexOf('sessionToken');
  const userI=headers.indexOf('userId');
  const cache=CacheService.getScriptCache();

  for(let r=values.length-1;r>=1;r--) {
    if(String(values[r][userI])===userId) {
      cache.remove('session:'+String(values[r][tokenI]));
      sh.deleteRow(r+1);
    }
  }
}

function passwordHashFast_(password,salt) {
  // Apps Script 上大量 computeDigest 迴圈非常慢。
  // 新版使用 server-side secret pepper + per-user salt 的 HMAC-SHA256。
  // Sheet 單獨外洩時，沒有 Script Properties 裡的 pepper 無法直接驗證密碼。
  const pepper =
    PropertiesService.getScriptProperties().getProperty('AUTH_PEPPER') || '';

  const bytes = Utilities.computeHmacSha256Signature(
    salt + '|' + password,
    pepper
  );

  return bytes.map(b =>
    ('0' + ((b < 0 ? b + 256 : b).toString(16))).slice(-2)
  ).join('');
}

function passwordHashLegacy_(password,salt) {
  const pepper=PropertiesService.getScriptProperties().getProperty('AUTH_PEPPER')||'';
  let s=pepper+'|'+salt+'|'+password;

  for(let i=0;i<HASH_ROUNDS;i++) {
    const bytes=Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      s,
      Utilities.Charset.UTF_8
    );

    s=bytes.map(b=>('0'+((b<0?b+256:b).toString(16))).slice(-2)).join('');
  }
  return s;
}

function resetCodeHash_(resetId,code) {
  const pepper=PropertiesService.getScriptProperties().getProperty('RESET_PEPPER')||'';
  const bytes=Utilities.computeHmacSha256Signature(resetId+'|'+code,pepper);
  return bytes.map(b=>('0'+((b<0?b+256:b).toString(16))).slice(-2)).join('');
}

function legacyRequestId_(userId,action,payload) {
  // 向下相容舊版 HTML：如果前端沒有 requestId，
  // 以「使用者 + 操作 + 內容 + 10 分鐘時間窗」生成穩定 ID。
  // 同一操作在短時間重試不會重複建立。
  const bucket=Math.floor(Date.now()/(10*60*1000));
  const raw=[
    String(userId||''),
    String(action||''),
    String(payload||''),
    String(bucket)
  ].join('|');

  const bytes=Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    raw,
    Utilities.Charset.UTF_8
  );

  const hex=bytes.map(b=>
    ('0'+((b<0?b+256:b).toString(16))).slice(-2)
  ).join('');

  return 'legacy-'+hex.slice(0,32);
}

function uniqueInviteCode_() {
  const existing=new Set(rows_(SHEETS.FAMILIES).map(x=>String(x.inviteCode)));
  let code='';
  do {
    const s=Utilities.getUuid().replace(/-/g,'').toUpperCase();
    code=s.slice(0,4)+'-'+s.slice(4,8);
  } while(existing.has(code));
  return code;
}

function ensureSheet_(ss,name,headers) {
  let sh=ss.getSheetByName(name);
  if(!sh) sh=ss.insertSheet(name);
  if(sh.getLastRow()===0) sh.appendRow(headers);
}

function ensureHeaders_(sh,headers) {
  if(!sh) return;

  const lastColumn=Math.max(1,sh.getLastColumn());
  const current=sh.getRange(1,1,1,lastColumn).getValues()[0].map(String);

  for(const h of headers) {
    if(!current.includes(h)) {
      const newCol=sh.getLastColumn()+1;
      sh.getRange(1,newCol).setValue(h);
      current.push(h);
    }
  }
}

function sheet_(name) {
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if(!sh) throw new Error('缺少工作表：'+name+'，請先執行 setupDatabase()');
  return sh;
}

function rows_(name) {
  const values=sheet_(name).getDataRange().getValues();
  if(values.length<=1) return [];

  const headers=values[0].map(String);

  return values.slice(1)
    .filter(r=>r.some(v=>v!==''))
    .map(r=>{
      const o={};
      headers.forEach((h,i)=>o[h]=normalizeCell_(r[i],h));
      return o;
    });
}

function normalizeCell_(v,h) {
  if(v instanceof Date) {
    if(['expiry','createAt'].includes(h)) {
      return Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd');
    }

    return Utilities.formatDate(
      v,Session.getScriptTimeZone(),"yyyy-MM-dd'T'HH:mm:ssXXX"
    );
  }
  return String(v);
}

function formatDate_(v) {
  if(!v) return '';
  if(v instanceof Date) {
    return Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd');
  }
  return String(v).slice(0,10);
}

function append_(sheetName,row) {
  sheet_(sheetName).appendRow(row);
}

function isoNow_() {
  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd'T'HH:mm:ssXXX"
  );
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function resetExpiryNotificationTrigger() {
  // 刪除既有的 checkExpiryNotifications Trigger
  ScriptApp.getProjectTriggers()
    .filter(t =>
      t.getHandlerFunction() === 'checkExpiryNotifications'
    )
    .forEach(t => ScriptApp.deleteTrigger(t));

  // 重新建立：每小時一次，盡量靠近整點
  ScriptApp.newTrigger('checkExpiryNotifications')
    .timeBased()
    .nearMinute(0)
    .everyHours(1)
    .create();

  Logger.log('checkExpiryNotifications Trigger 已重新建立，設定為每小時接近整點執行');
}
