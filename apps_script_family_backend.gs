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
  FOOD_IMAGES: 'FoodImages'
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
const PENDING_IMAGE_RETENTION_DAYS = 2;
const DELETED_IMAGE_META_RETENTION_DAYS = 35;
const MAX_PRODUCT_IMAGE_BYTES = 450000; // 前端正常約 50~150 KB，後端再設硬上限

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
    ['familyId','id','name','qty','location','expiry','note','createdBy','createAt','updatedAt','notifyMode','notifyDaysBefore','imageId']);

  ensureHeaders_(ss.getSheetByName(SHEETS.FOODS),
    ['familyId','id','name','qty','location','expiry','note','createdBy','createAt','updatedAt','notifyMode','notifyDaysBefore','imageId']);

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
    ['userId','familyId','foodId','expiry','daysBefore','sentAt']);

  ensureSheet_(ss, SHEETS.FOOD_IMAGES,
    ['imageId','familyId','foodId','driveFileId','mimeType','sizeBytes','width','height','status','uploadedBy','createdAt','updatedAt','deletedAt']);

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
  try {
    const p = e.parameter || {};
    const action = p.action || '';

    if (action === 'register') return json_(register_(p));
    if (action === 'login') return json_(login_(p));
    if (action === 'requestPasswordReset') return json_(requestPasswordReset_(p));
    if (action === 'resetPassword') return json_(resetPassword_(p));

    const user = requireSession_(p.sessionToken);

    if (action === 'logout') return json_(logout_(p.sessionToken));
    if (action === 'me') return json_(me_(user));
    if (action === 'getNotificationSettings') return json_(getNotificationSettings_(user));
    if (action === 'saveNotificationSettings') return json_(saveNotificationSettings_(user,p));

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
    user:publicUser_(user),
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
  return { ok:true, user:publicUser_(user), families:userFamilies_(user.userId) };
}

function publicUser_(user) {
  return { userId:user.userId,email:user.email,displayName:user.displayName };
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
  const tz=Session.getScriptTimeZone();
  const now=new Date();
  const currentHour=Number(Utilities.formatDate(now,tz,'H'));

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
    if(Number(us.sendHour)!==currentHour) continue;

    const memberRows=memberships.filter(x=>x.userId===u.userId);
    if(!memberRows.length) continue;

    const items=[];

    for(const m of memberRows) {
      const fam=families.find(x=>x.familyId===m.familyId);
      if(!fam) continue;

      const gset=groupSettings.find(x=>x.userId===u.userId && x.familyId===m.familyId);
      const enabled=gset ? toBool_(gset.enabled) : true;
      const defaultDays=gset ? Math.max(0,Math.min(365,Number(gset.defaultDaysBefore||3))) : 3;
      if(!enabled) continue;

      const groupFoods=foods.filter(x=>x.familyId===m.familyId);

      for(const f of groupFoods) {
        const mode=String(f.notifyMode||'inherit');
        if(mode==='off') continue;

        const daysBefore=mode==='custom'
          ? Math.max(0,Math.min(365,Number(f.notifyDaysBefore||0)))
          : defaultDays;

        const days=daysUntilDate_(f.expiry,tz);
        if(days!==daysBefore) continue;

        const already=logs.some(l=>
          l.userId===u.userId &&
          l.familyId===m.familyId &&
          l.foodId===f.id &&
          formatDate_(l.expiry)===formatDate_(f.expiry) &&
          Number(l.daysBefore)===daysBefore
        );

        if(already) continue;

        items.push({
          familyId:m.familyId,
          familyName:fam.familyName,
          foodId:f.id,
          name:f.name,
          qty:Number(f.qty||1),
          expiry:formatDate_(f.expiry),
          daysBefore
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
        const when=i.daysBefore===0 ? '今天到期' : '剩 '+i.daysBefore+' 天到期';
        body+='- '+i.name+' ×'+i.qty+'：'+when+'（'+i.expiry+'）\n';
      });
      body+='\n';
    });

    MailApp.sendEmail({
      to:u.email,
      subject:'食品過期提醒：'+items.length+' 項食品需要注意',
      body
    });

    items.forEach(i=>{
      append_(SHEETS.NOTIFY_LOG,[
        u.userId,i.familyId,i.foodId,i.expiry,i.daysBefore,isoNow_()
      ]);
    });
  }
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

  markFamilyImagesPendingDelete_(familyId);

  deleteRowsByValue_(SHEETS.FOODS,'familyId',familyId);
  deleteRowsByValue_(SHEETS.MEMBERS,'familyId',familyId);
  deleteRowsByValue_(SHEETS.ACTIVITY,'familyId',familyId);
  deleteRowsByValue_(SHEETS.FAMILIES,'familyId',familyId);

  return { ok:true };
}

// ---------------- FOOD / BATCH SYNC ----------------

function listFoods_(user,p) {
  const familyId=String(p.familyId||'');
  const membership=requireMembership_(user.userId,familyId);

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

    if(idx.imageId==null) {
      throw new Error('Foods 尚未建立 imageId 欄位，請先執行 setupDatabase()');
    }

    const applied=[];
    const items=[];
    const deletedIds=[];
    const conflicts=[];

    for(const ch of changes) {
      const type=String(ch.type||'');
      const clientKey=String(ch.clientKey||Utilities.getUuid());

      if(type==='add') {
        const f=validateFoodPayload_(ch.food||{});
        validateImageReferenceForChange_(
          user,
          familyId,
          f.imageId,
          ''
        );

        const id=Utilities.getUuid();
        const now=isoNow_();

        sh.appendRow([
          familyId,id,f.name,f.qty,f.location,f.expiry,f.note,
          user.userId,now.slice(0,10),now,
          f.notifyMode,f.notifyDaysBefore,f.imageId
        ]);

        if(f.imageId) {
          attachImageToFood_(
            f.imageId,
            familyId,
            id,
            user.userId
          );
        }

        const saved={
          id,
          name:f.name,
          qty:f.qty,
          location:f.location,
          expiry:f.expiry,
          note:f.note,
          createAt:now.slice(0,10),
          updatedAt:now,
          notifyMode:f.notifyMode,
          notifyDaysBefore:f.notifyDaysBefore,
          imageId:f.imageId
        };

        applied.push({clientKey});
        items.push({
          tempId:String(ch.tempId||''),
          food:saved
        });

        logActivity_(
          familyId,
          user.userId,
          'food_add',
          id,
          f.name
        );

        values.push([
          familyId,id,f.name,f.qty,f.location,f.expiry,f.note,
          user.userId,now.slice(0,10),now,
          f.notifyMode,f.notifyDaysBefore,f.imageId
        ]);

        continue;
      }

      const id=String(ch.id||'');
      const rowIndex=findFoodRow_(
        values,
        idx,
        familyId,
        id
      );

      if(rowIndex<1) {
        conflicts.push({
          clientKey,
          id,
          reason:'not_found'
        });
        continue;
      }

      const currentUpdatedAt=normalizeCell_(
        values[rowIndex][idx.updatedAt],
        'updatedAt'
      );

      const expected=String(
        ch.expectedUpdatedAt||''
      );

      if(
        expected &&
        currentUpdatedAt &&
        expected!==currentUpdatedAt
      ) {
        conflicts.push({
          clientKey,
          id,
          reason:'modified_elsewhere'
        });
        continue;
      }

      const oldImageId=String(
        values[rowIndex][idx.imageId]||''
      );

      if(type==='update') {
        const f=validateFoodPayload_(ch.food||{});

        validateImageReferenceForChange_(
          user,
          familyId,
          f.imageId,
          id
        );

        const now=isoNow_();

        sh.getRange(rowIndex+1,idx.name+1)
          .setValue(f.name);
        sh.getRange(rowIndex+1,idx.qty+1)
          .setValue(f.qty);
        sh.getRange(rowIndex+1,idx.location+1)
          .setValue(f.location);
        sh.getRange(rowIndex+1,idx.expiry+1)
          .setValue(f.expiry);
        sh.getRange(rowIndex+1,idx.note+1)
          .setValue(f.note);
        sh.getRange(rowIndex+1,idx.notifyMode+1)
          .setValue(f.notifyMode);
        sh.getRange(rowIndex+1,idx.notifyDaysBefore+1)
          .setValue(f.notifyDaysBefore);
        sh.getRange(rowIndex+1,idx.imageId+1)
          .setValue(f.imageId);
        sh.getRange(rowIndex+1,idx.updatedAt+1)
          .setValue(now);

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

        values[rowIndex][idx.name]=f.name;
        values[rowIndex][idx.qty]=f.qty;
        values[rowIndex][idx.location]=f.location;
        values[rowIndex][idx.expiry]=f.expiry;
        values[rowIndex][idx.note]=f.note;
        values[rowIndex][idx.notifyMode]=f.notifyMode;
        values[rowIndex][idx.notifyDaysBefore]=f.notifyDaysBefore;
        values[rowIndex][idx.imageId]=f.imageId;
        values[rowIndex][idx.updatedAt]=now;

        applied.push({clientKey});

        items.push({
          food:{
            id,
            name:f.name,
            qty:f.qty,
            location:f.location,
            expiry:f.expiry,
            note:f.note,
            createAt:normalizeCell_(
              values[rowIndex][idx.createAt],
              'createAt'
            ),
            updatedAt:now,
            notifyMode:f.notifyMode,
            notifyDaysBefore:f.notifyDaysBefore,
            imageId:f.imageId
          }
        });

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
        if(oldImageId) {
          markImagePendingDelete_(
            oldImageId,
            'food_deleted'
          );
        }

        sh.deleteRow(rowIndex+1);
        values.splice(rowIndex,1);

        applied.push({clientKey});
        deletedIds.push(id);

        logActivity_(
          familyId,
          user.userId,
          'food_delete',
          id,
          ''
        );

        continue;
      }

      conflicts.push({
        clientKey,
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

function validateFoodPayload_(f) {
  const name=String(f.name||'').trim();
  const expiry=String(f.expiry||'').trim();
  const qty=Math.max(1,Number(f.qty||1));

  if(!name) throw new Error('食品名稱不可空白');
  if(!/^\d{4}-\d{2}-\d{2}$/.test(expiry)) throw new Error('到期日格式不正確');

  return {
    name,
    qty,
    location:String(f.location||''),
    expiry,
    note:String(f.note||''),
    notifyMode:['inherit','custom','off'].includes(String(f.notifyMode||'inherit')) ? String(f.notifyMode||'inherit') : 'inherit',
    notifyDaysBefore:String(f.notifyMode||'inherit')==='custom' ? Math.max(0,Math.min(365,Number(f.notifyDaysBefore||0))) : '',
    imageId:String(f.imageId||'').trim()
  };
}

function findFoodRow_(values,idx,familyId,id) {
  for(let r=1;r<values.length;r++) {
    if(String(values[r][idx.familyId])===familyId &&
       String(values[r][idx.id])===id) return r;
  }
  return -1;
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
    imageId:String(x.imageId||'')
  };
}


// ---------------- DRIVE IMAGE STORAGE / BACKUP ----------------

function setupStorageAndMaintenance() {
  setupDatabase();

  const folders=getOrCreateStorageFolders_();

  // 避免重複建立同一個每週 Trigger。
  ScriptApp.getProjectTriggers()
    .filter(t =>
      t.getHandlerFunction()==='weeklyBackupAndMaintenance'
    )
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('weeklyBackupAndMaintenance')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(3)
    .create();

  const result={
    ok:true,
    rootFolderId:folders.root.getId(),
    imageFolderId:folders.images.getId(),
    backupFolderId:folders.backups.getId(),
    message:'已建立 Drive 儲存資料夾與每週備份 / 清理 Trigger'
  };

  Logger.log(JSON.stringify(result));
  return result;
}

function weeklyBackupAndMaintenance() {
  // 順序刻意是「先備份，再清理」。
  const backup=createDatabaseBackup_();
  const cleanup=cleanupTransientData_();
  const imageCleanup=cleanupFoodImageMetadata_();
  const oldBackups=cleanupOldBackups_();

  const result={
    ok:true,
    backup,
    cleanup,
    imageCleanup,
    oldBackups
  };

  Logger.log(JSON.stringify(result));
  return result;
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

    const now=isoNow_();

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
