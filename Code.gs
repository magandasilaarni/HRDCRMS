const CFG = {
  REQUESTS: 'Requests',
  APPROVALS: 'Approvals',
  USERS: 'Users',
  AUDIT: 'AuditTrail',
  SETTINGS: 'Settings',
  TZ: 'Asia/Manila',
  PREFIX: 'HRD-CERT',
  SESSION_SECONDS: 21600
};

const RH = [
  'RequestID','TrackingHash','ControlNumber','SubmittedAt','UpdatedAt',
  'ProponentEmail','ProponentName','OfficeSchool','ActivityTitle','ActivityType',
  'CertificateTypes','TotalQuantity','ActivityStart','ActivityEnd','Notes','UrgencyStatus','Status','HRDAssignee',
  'DraftURL','FinalURL','VerificationToken','ReleasedAt','ReturnReason'
];

const AH = [
  'ApprovalID','RequestID','Step','ApproverEmail','ApproverName',
  'Decision','Comments','ActionAt'
];

const UH = [
  'Email','FullName','Role','Active','Office','ApprovalStep','Salt','PasswordHash'
];

const TH = [
  'Timestamp','RequestID','ActorEmail','Role','Action','Details'
];

const SETTINGS_HEADERS = ['Key','Value'];

function setupSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Create this project from Extensions > Apps Script in your Google Sheet.');

  ensureSheet_(ss, CFG.REQUESTS, RH);
  ensureSheet_(ss, CFG.APPROVALS, AH);
  ensureSheet_(ss, CFG.USERS, UH);
  ensureSheet_(ss, CFG.AUDIT, TH);
  ensureSheet_(ss, CFG.SETTINGS, SETTINGS_HEADERS);

  const settings = ss.getSheetByName(CFG.SETTINGS);
  if (settings.getLastRow() < 2) {
    settings.getRange(2, 1, 3, 2).setValues([
      ['ControlNumberPrefix', CFG.PREFIX],
      ['OrganizationName', 'HRD Laguna'],
      ['PublicPortalEnabled', 'TRUE']
    ]);
  }

  [CFG.REQUESTS, CFG.APPROVALS, CFG.USERS, CFG.AUDIT, CFG.SETTINGS]
    .forEach(n => ss.getSheetByName(n).setFrozenRows(1));

  return 'Setup complete.';
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('HRD Certificate Request & Approval System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function apiPublicInfo() {
  return {
    organization: setting_('OrganizationName') || 'HRD Laguna',
    enabled: String(setting_('PublicPortalEnabled')).toUpperCase() !== 'FALSE'
  };
}

function calculateUrgency_(activityStartDateStr, submittedDate) {
  if (!activityStartDateStr) return 'Normal (No Start Date)';
  const startDate = new Date(activityStartDateStr);
  if (isNaN(startDate.getTime())) return 'Normal';
  
  const startMs = startDate.setHours(0,0,0,0);
  const subMs = new Date(submittedDate).setHours(0,0,0,0);
  const diffDays = Math.ceil((startMs - subMs) / (1000 * 60 * 60 * 24));

  if (diffDays < 5) {
    return `URGENT (${diffDays < 0 ? 'Past Event' : diffDays + ' day(s) lead time'})`;
  } else {
    return `TIMELY (${diffDays} day(s) lead time)`;
  }
}

function apiSubmitPublic(p) {
  if (String(setting_('PublicPortalEnabled')).toUpperCase() === 'FALSE') {
    throw new Error('The public request portal is temporarily closed.');
  }

  p = p || {};
  if (!p.proponentName || !p.proponentEmail || !p.officeSchool || !p.activityTitle || !p.activityType || !p.activityStart) {
    throw new Error('Complete all required fields including Activity Start Date.');
  }

  let certTypesSummary = [];
  let totalQty = 0;

  if (p.certDetails && typeof p.certDetails === 'object') {
    Object.keys(p.certDetails).forEach(type => {
      const qty = Number(p.certDetails[type] || 0);
      if (qty > 0) {
        certTypesSummary.push(`${type} (${qty})`);
        totalQty += qty;
      }
    });
  }

  if (certTypesSummary.length === 0) {
    throw new Error('Enable and specify a quantity for at least one Certificate Type requested.');
  }

  if (!p.emailAttachmentConfirmed) {
    throw new Error('Please confirm that you have sent your supporting document attachments via email.');
  }

  const proponentEmail = email_(p.proponentEmail);
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(proponentEmail)) {
    throw new Error('Enter a valid email address.');
  }

  const approvers = configuredApprovers_();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const requestId = 'REQ-' + Utilities.getUuid().split('-')[0].toUpperCase();
    const trackingKey = randomSecret_();
    const now = new Date();
    const certTypesStr = certTypesSummary.join(', ');
    const urgency = calculateUrgency_(p.activityStart, now);

    const sh = SpreadsheetApp.getActive().getSheetByName(CFG.REQUESTS);
    sh.appendRow([
      requestId, secretHash_(trackingKey), '', now, now,
      proponentEmail, String(p.proponentName).trim(),
      String(p.officeSchool).trim(), String(p.activityTitle).trim(),
      String(p.activityType), certTypesStr, totalQty, p.activityStart || '', p.activityEnd || '',
      String(p.notes || ''), urgency, 'Submitted', '', '', '', '', '', ''
    ]);

    const rows = approvers.map(a => [
      'APR-' + Utilities.getUuid().split('-')[0].toUpperCase(),
      requestId, a.step, a.email, a.name, 'Pending', '', ''
    ]);

    const as = SpreadsheetApp.getActive().getSheetByName(CFG.APPROVALS);
    if (rows.length > 0) {
      as.getRange(as.getLastRow() + 1, 1, rows.length, AH.length).setValues(rows);
    }

    audit_(requestId, { email: proponentEmail, role: 'PROPONENT' }, 'SUBMITTED', `Request submitted. Urgency: ${urgency}. Certificates requested: ${certTypesStr}. Total quantity: ${totalQty}.`);

    const confirmationEmailSent = sendSubmissionConfirmationEmail_(proponentEmail, String(p.proponentName).trim(), requestId, String(p.activityTitle).trim(), certTypesStr, totalQty, urgency, p.activityStart, p.activityEnd, now);
    const hrdEmailSent = notifyHrdNewRequest_(requestId, String(p.proponentName).trim(), proponentEmail, String(p.officeSchool).trim(), String(p.activityTitle).trim(), String(p.activityType), certTypesStr, totalQty, p.activityStart, p.activityEnd, urgency, now);
    if (!confirmationEmailSent) audit_(requestId, {email:proponentEmail,role:'SYSTEM'}, 'EMAIL_FAILED', 'Could not send submission confirmation email to proponent.');
    if (!hrdEmailSent) audit_(requestId, {email:'hrds.laguna@deped.gov.ph',role:'SYSTEM'}, 'EMAIL_FAILED', 'Could not send new-request notification to HRD. Check MailApp authorization/quota.');

    return { ok: true, requestId: requestId, trackingKey: trackingKey, confirmationEmailSent: confirmationEmailSent, hrdEmailSent: hrdEmailSent };
  } finally {
    lock.releaseLock();
  }
}

function apiResubmitPublic(requestId) {
  requestId = String(requestId || '').trim();
  if (!requestId) throw new Error('Missing Request ID.');

  const r = requestById_(requestId);
  if (!r) throw new Error('Request not found.');

  if (!['Returned to Proponent', 'Returned for Revision'].includes(r.Status)) {
    throw new Error('Only requests returned for revision can be resubmitted.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    // Preserve the distinction between a first-time submission and a resubmission.
    // HRD can process this status through the same Prepare Draft workflow.
    updateRequest_(requestId, {
      Status: 'Resubmitted',
      ReturnReason: ''
    });

    resetApprovals_(requestId);
    audit_(requestId, { email: r.ProponentEmail, role: 'PROPONENT' }, 'RESUBMITTED', `Updated request resubmitted. Certificates requested: ${r.CertificateTypes || 'Not recorded'}. Total quantity: ${r.TotalQuantity || 0}.`);

    const confirmationEmailSent = sendSubmissionConfirmationEmail_(r.ProponentEmail, r.ProponentName, requestId, r.ActivityTitle, r.CertificateTypes, r.TotalQuantity, r.UrgencyStatus, r.ActivityStart, r.ActivityEnd, new Date());
    const hrdEmailSent = notifyHrdNewRequest_(requestId, r.ProponentName, r.ProponentEmail, r.OfficeSchool, r.ActivityTitle, r.ActivityType, r.CertificateTypes, r.TotalQuantity, r.ActivityStart, r.ActivityEnd, r.UrgencyStatus, new Date());
    if (!confirmationEmailSent) audit_(requestId, {email:r.ProponentEmail,role:'SYSTEM'}, 'EMAIL_FAILED', 'Resubmission confirmation email failed.');
    if (!hrdEmailSent) audit_(requestId, {email:'hrds.laguna@deped.gov.ph',role:'SYSTEM'}, 'EMAIL_FAILED', 'HRD resubmission notification failed.');

    return { ok: true, requestId: requestId };
  } finally {
    lock.releaseLock();
  }
}

function sendSubmissionConfirmationEmail_(toEmail, name, requestId, title, certTypes, totalQty, urgency, activityStart, activityEnd, submittedAt) {
  const subject = `[HRD Certificate Request] Request Recorded - ${requestId}`;
  const body = `Dear ${name},

` +
    `Your certificate request (${requestId}) has been recorded successfully.

` +
    `Request Details:
` +
    `- Request ID: ${requestId}
` +
    `- Date/Time Submitted: ${formatDateTime_(submittedAt)}
` +
    `- Activity Title: ${title}
` +
    `- Activity Date: ${formatDateOnly_(activityStart)}${activityEnd ? ' to ' + formatDateOnly_(activityEnd) : ''}
` +
    `- Urgency Classification: ${urgency}
` +
    `- Total Certificates Requested: ${totalQty}
` +
    `- Certificate Breakdown: ${certTypes}

` +
    `Please keep your Request ID for tracking.
` +
    `Ensure supporting files are emailed to hrds.laguna@deped.gov.ph with Request ID (${requestId}) in the subject line.

` +
    `Thank you.
HRD Office - Schools Division of Laguna`;
  return sendSystemEmail_(toEmail, subject, body);
}

function notifyHrdNewRequest_(requestId, name, email, office, title, type, certTypes, totalQty, startDate, endDate, urgency, submittedAt) {
  const subject = `[HRD Certificate Request] New request ${requestId}`;
  const body = `A new certificate request has been submitted.

` +
    `Request ID: ${requestId}
Date/Time Submitted: ${formatDateTime_(submittedAt)}
` +
    `Proponent: ${name}
Email: ${email}
Office/School: ${office}
` +
    `Activity: ${title}
Activity Type: ${type}
` +
    `Activity Date: ${formatDateOnly_(startDate)}${endDate ? ' to ' + formatDateOnly_(endDate) : ''}
` +
    `Urgency: ${urgency}
Certificate Types / Quantities: ${certTypes}
Total Quantity: ${totalQty}

` +
    `Please sign in to the HRD Certificate Request & Approval System to process this request.`;
  return sendSystemEmail_(setting_('HRDEmail') || 'hrds.laguna@deped.gov.ph', subject, body);
}

function sendSystemEmail_(toEmail, subject, body) {
  if (!toEmail) throw new Error('Email recipient is missing.');
  try {
    MailApp.sendEmail({to:String(toEmail), subject:String(subject), body:String(body), name:'HRD Certificate Request System'});
    return true;
  } catch (e) {
    Logger.log('Email failed to ' + toEmail + ': ' + e.message);
    return false;
  }
}

function formatDateOnly_(value) {
  if (!value) return 'Not specified';
  const d = value instanceof Date ? value : new Date(value);
  return isNaN(d.getTime()) ? String(value) : Utilities.formatDate(d, CFG.TZ, 'MMMM d, yyyy');
}
function formatDateTime_(value) {
  if (!value) return Utilities.formatDate(new Date(), CFG.TZ, 'MMMM d, yyyy h:mm a');
  const d = value instanceof Date ? value : new Date(value);
  return isNaN(d.getTime()) ? String(value) : Utilities.formatDate(d, CFG.TZ, 'MMMM d, yyyy h:mm a');
}

function apiTrackPublic(query) {
  query = String(query || '').trim();
  if (!query) throw new Error('Please enter a Request ID or Control Number.');

  const requests = objects_(CFG.REQUESTS);
  const r = requests.find(x => 
    String(x.RequestID).toLowerCase() === query.toLowerCase() || 
    (x.ControlNumber && String(x.ControlNumber).toLowerCase().includes(query.toLowerCase()))
  );

  if (!r) throw new Error('No request found matching the provided search string.');

  const approvals = objects_(CFG.APPROVALS)
    .filter(a => a.RequestID === r.RequestID)
    .sort((a,b) => Number(a.Step) - Number(b.Step));

  const auditTrail = objects_(CFG.AUDIT)
    .filter(a => a.RequestID === r.RequestID && !['EMAIL_FAILED','STAFF_LOGIN','STAFF_LOGOUT','TRANSACTION_BACKUP_EXPORTED'].includes(String(a.Action || '').toUpperCase()))
    .sort((a,b) => new Date(a.Timestamp) - new Date(b.Timestamp));

  let returnReason = r.ReturnReason || '';
  if (!returnReason) {
    const rejected = approvals.find(a => a.Decision === 'Disapproved' || a.Decision === 'Returned for Revision');
    if (rejected && rejected.Comments) returnReason = rejected.Comments;
  }

  return {
    requestId: r.RequestID,
    proponentName: r.ProponentName,
    proponentEmail: r.ProponentEmail,
    activityTitle: r.ActivityTitle,
    activityType: r.ActivityType,
    certificateTypes: r.CertificateTypes || '',
    certificateBreakdown: buildCertificateBreakdown_(r.CertificateTypes, r.ControlNumber, r.Status),
    totalQuantity: r.TotalQuantity || 0,
    officeSchool: r.OfficeSchool,
    activityStart: r.ActivityStart || '',
    activityEnd: r.ActivityEnd || '',
    notes: r.Notes || '',
    submittedAt: r.SubmittedAt,
    updatedAt: r.UpdatedAt,
    urgencyStatus: r.UrgencyStatus || 'Normal',
    status: r.Status,
    returnReason: returnReason,
    controlNumber: r.ControlNumber || '',
    finalURL: r.Status === 'Released' ? (r.FinalURL || '') : '',
    approvalProgress: approvals.map(a => ({
      step: Number(a.Step),
      name: a.ApproverName || ('Approver Step ' + a.Step),
      decision: a.Decision,
      comments: a.Comments || ''
    })),
    auditTrail: auditTrail.map(t => ({
      timestamp: t.Timestamp,
      action: friendlyAuditAction_(t.Action),
      details: publicAuditDetails_(t.Action, t.Details)
    }))
  };
}

function apiStaffLogin(email, password) {
  email = email_(email);
  password = String(password || '');

  const cache = CacheService.getScriptCache();
  const throttleKey = 'login_' + secretHash_(email);
  const attempts = Number(cache.get(throttleKey) || 0);

  if (attempts >= 10) throw new Error('Too many login attempts. Try again in a few minutes.');

  const users = staffRows_();
  const user = users.find(u => email_(u.Email) === email && isTrue_(u.Active));

  if (!user || String(user.PasswordHash) !== passwordHash_(password, String(user.Salt))) {
    cache.put(throttleKey, String(attempts + 1), 300);
    throw new Error('Email or password is incorrect.');
  }

  cache.remove(throttleKey);
  const token = randomSecret_() + randomSecret_();
  const session = {
    email: email_(user.Email),
    fullName: String(user.FullName || user.Email),
    role: String(user.Role).toUpperCase(),
    office: String(user.Office || ''),
    step: Number(user.ApprovalStep || 0)
  };

  cache.put('session_' + secretHash_(token), JSON.stringify(session), CFG.SESSION_SECONDS);
  audit_('SYSTEM', session, 'STAFF_LOGIN', 'Staff login succeeded.');

  return { token: token, user: session };
}

function apiStaffLogout(token) {
  const session = staffSession_(token);
  CacheService.getScriptCache().remove('session_' + secretHash_(token));
  audit_('SYSTEM', session, 'STAFF_LOGOUT', 'Staff logged out.');
  return {ok:true};
}

function apiAdminCreateStaffAccount(token, payload) {
  const u = staffSession_(token);
  requireRole_(u, ['ADMIN']);
  const msg = createStaffAccount_(payload.email, payload.fullName, payload.role, payload.password, payload.office, payload.approvalStep);
  return { ok: true, message: msg };
}

function staffSession_(token) {
  token = String(token || '');
  if (token.length < 30) throw new Error('Your staff session has expired. Please sign in again.');

  const cacheKey = 'session_' + secretHash_(token);
  const raw = CacheService.getScriptCache().get(cacheKey);

  if (!raw) throw new Error('Your staff session has expired. Please sign in again.');

  const session = JSON.parse(raw);
  const user = staffRows_().find(u => email_(u.Email) === session.email && isTrue_(u.Active));

  if (!user || String(user.Role).toUpperCase() !== session.role) {
    CacheService.getScriptCache().remove(cacheKey);
    throw new Error('Your account is inactive or permissions changed.');
  }

  return session;
}

function apiStaffBootstrap(token) {
  const u = staffSession_(token);
  const requests = objects_(CFG.REQUESTS);
  const approvals = objects_(CFG.APPROVALS);

  let visible = requests;
  if (u.role === 'APPROVER') {
    const assignedIds = new Set(approvals
      .filter(a => email_(a.ApproverEmail) === u.email || Number(a.Step) === u.step)
      .map(a => a.RequestID));
    visible = requests.filter(r => assignedIds.has(r.RequestID));
  }

  return {
    user: u,
    requests: visible.map(r => {
      const apprs = approvals.filter(a => a.RequestID === r.RequestID)
        .sort((a,b) => Number(a.Step) - Number(b.Step));

      let returnReason = r.ReturnReason || '';
      if (!returnReason) {
        const rejected = apprs.find(a => a.Decision === 'Disapproved' || a.Decision === 'Returned for Revision');
        if (rejected && rejected.Comments) returnReason = rejected.Comments;
      }

      return {
        requestId: r.RequestID,
        controlNumber: r.ControlNumber || '',
        submittedAt: r.SubmittedAt,
        updatedAt: r.UpdatedAt,
        proponentName: r.ProponentName,
        proponentEmail: r.ProponentEmail,
        officeSchool: r.OfficeSchool,
        activityTitle: r.ActivityTitle,
        activityType: r.ActivityType,
        certificateTypes: r.CertificateTypes || '',
        certificateBreakdown: buildCertificateBreakdown_(r.CertificateTypes, r.ControlNumber, r.Status),
        totalQuantity: r.TotalQuantity || 0,
        activityStart: r.ActivityStart,
        activityEnd: r.ActivityEnd,
        notes: r.Notes,
        urgencyStatus: r.UrgencyStatus || 'Normal',
        status: r.Status,
        returnReason: returnReason,
        hrdAssignee: r.HRDAssignee,
        draftURL: r.DraftURL,
        finalURL: r.FinalURL,
        approvals: apprs.map(a => ({
          step:Number(a.Step), email:a.ApproverEmail,
          name:a.ApproverName, decision:a.Decision,
          comments:a.Comments, actionAt:a.ActionAt
        }))
      };
    })
  };
}

function apiHRDAction(token, p) {
  const u = staffSession_(token);
  requireRole_(u, ['ADMIN','HRD']);

  const r = requestById_(p.requestId);
  if (!r) throw new Error('Request not found.');

  const action = String(p.action || '');
  const updates = {};

  if (action === 'Prepare Draft') {
    if (p.draftURL) checkHttps_(p.draftURL);
    updates.Status = 'Draft in Preparation';
    updates.HRDAssignee = p.assignee || u.email;
    updates.DraftURL = p.draftURL || r.DraftURL || '';
  } else if (action === 'Send for Approval') {
    if (!p.draftURL && !r.DraftURL) throw new Error('Provide the draft certificate link first.');
    if (p.draftURL) checkHttps_(p.draftURL);

    updates.Status = 'Awaiting Approvals';
    updates.DraftURL = p.draftURL || r.DraftURL;
    updates.HRDAssignee = p.assignee || u.email;
    resetApprovals_(r.RequestID);
    configuredApprovers_().forEach(a => {
      try {
        sendSystemEmail_(a.email, `[HRD Approval Required] ${r.RequestID}`,
          `A certificate request is ready for review.\nRequest ID: ${r.RequestID}\nActivity: ${r.ActivityTitle}\nProponent: ${r.ProponentName}\nActivity Date: ${formatDateOnly_(r.ActivityStart)}${r.ActivityEnd ? ' to ' + formatDateOnly_(r.ActivityEnd) : ''}\nDraft: ${p.draftURL || r.DraftURL}\nPlease sign in to record your decision.`);
      } catch (mailErr) {
        audit_(r.RequestID, u, 'EMAIL_FAILED', 'Could not notify approver '+a.email+': '+mailErr.message);
      }
    });
  } else if (action === 'Return to Proponent') {
    if (!p.comments) throw new Error('Enter revision instructions.');
    updates.Status = 'Returned to Proponent';
    updates.ReturnReason = p.comments;
    sendRevisionEmail_(r.ProponentEmail, r.ProponentName, r.RequestID, r.ActivityTitle, p.comments);
  } else if (action === 'Disapprove Request') {
    if (!p.comments) throw new Error('Enter reason for disapproval.');
    updates.Status = 'Disapproved';
    updates.ReturnReason = p.comments;
    sendDisapprovalEmail_(r.ProponentEmail, r.ProponentName, r.RequestID, r.ActivityTitle, p.comments);
  } else {
    throw new Error('Unknown HRD action.');
  }

  updateRequest_(r.RequestID, updates);
  audit_(r.RequestID, u, action.toUpperCase(), 'Recorded at '+formatDateTime_(new Date())+'. '+(p.comments || p.draftURL || ''));
  return {ok:true};
}

function apiApproverDecision(token, p) {
  const u = staffSession_(token);
  requireRole_(u, ['APPROVER', 'ADMIN']);

  const r = requestById_(p.requestId);
  if (!r || r.Status !== 'Awaiting Approvals') {
    throw new Error('This request is not currently awaiting approvals.');
  }

  const decision = String(p.decision || '');
  if (!['Approved','Disapproved','Returned for Revision'].includes(decision)) {
    throw new Error('Choose a valid decision.');
  }
  if (decision !== 'Approved' && !p.comments) {
    throw new Error('Comments/Reasons are required for return or disapproval.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(CFG.APPROVALS);
    const data = sh.getDataRange().getValues();
    const ix = headerIndex_(data[0]);
    let row = -1;

    for (let i = 1; i < data.length; i++) {
      const matchEmail = email_(data[i][ix.ApproverEmail]) === u.email || u.role === 'ADMIN';
      const matchStep = Number(data[i][ix.Step]) === u.step || u.role === 'ADMIN';

      if (String(data[i][ix.RequestID]) === r.RequestID && matchEmail && matchStep && data[i][ix.Decision] === 'Pending') {
        row = i + 1;
        break;
      }
    }

    if (row < 0) {
      throw new Error('No pending approval step is assigned to your account for this request.');
    }

    sh.getRange(row, ix.Decision + 1).setValue(decision);
    sh.getRange(row, ix.Comments + 1).setValue(p.comments || '');
    sh.getRange(row, ix.ActionAt + 1).setValue(new Date());

    if (decision === 'Disapproved') {
      updateRequest_(r.RequestID, {Status:'Disapproved', ReturnReason: p.comments || ''});
      sendDisapprovalEmail_(r.ProponentEmail, r.ProponentName, r.RequestID, r.ActivityTitle, p.comments);
    } else if (decision === 'Returned for Revision') {
      updateRequest_(r.RequestID, {Status:'Returned for Revision', ReturnReason: p.comments || ''});
      sendRevisionEmail_(r.ProponentEmail, r.ProponentName, r.RequestID, r.ActivityTitle, p.comments);
    } else {
      const all = objects_(CFG.APPROVALS).filter(a => a.RequestID === r.RequestID);
      if (all.length > 0 && all.every(a => a.Decision === 'Approved')) {
        const perCertificateControlNumbers = generateIndividualCertificateControlNumbers_(r.RequestID, r.CertificateTypes);

        updateRequest_(r.RequestID, {
          Status:'Approved - For Release',
          ControlNumber: perCertificateControlNumbers,
          VerificationToken: randomSecret_()
        });
        audit_(r.RequestID, u, 'ALL_APPROVALS_COMPLETE', 'All required approvals recorded.');
      }
    }

    audit_(r.RequestID, u, decision.toUpperCase(), 'Recorded at '+formatDateTime_(new Date())+'. '+(p.comments || ''));
    try {
      sendSystemEmail_('hrds.laguna@deped.gov.ph', `[HRD Approval Update] ${r.RequestID}`,
        `Approval decision recorded.\\nDate/Time: ${formatDateTime_(new Date())}\\nRequest ID: ${r.RequestID}\\nActivity: ${r.ActivityTitle}\\nApprover: ${u.fullName} (${u.email})\\nDecision: ${decision}\\nComments: ${p.comments || 'None'}\\nCurrent status: ${(requestById_(r.RequestID)||{}).Status || 'Updated'}`);
    } catch (mailErr) {
      audit_(r.RequestID, u, 'EMAIL_FAILED', 'HRD approval update notification failed: '+mailErr.message);
    }
    return {ok:true};
  } finally {
    lock.releaseLock();
  }
}

function generateIndividualCertificateControlNumbers_(requestId, certTypesStr) {
  const prefix = String(setting_('ControlNumberPrefix') || CFG.PREFIX);
  const year = Utilities.formatDate(new Date(), CFG.TZ, 'yyyy');
  const existing = objects_(CFG.REQUESTS).map(r => String(r.ControlNumber || ''));
  const count = existing.filter(c => c.includes(prefix + '-' + year + '-')).length + 1;
  const baseNum = `${prefix}-${year}-${String(count).padStart(5, '0')}`;
  const types = parseCertificateTypes_(certTypesStr);
  if (!types.length) return `Other (1): ${baseNum}-CERT-1`;

  return types.map(item => {
    const code = certificateTypeCode_(item.type);
    const nums = [];
    for (let i = 1; i <= item.quantity; i++) nums.push(`${baseNum}-${code}-${i}`);
    return `${item.type} (${item.quantity}): ${nums.join(', ')}`;
  }).join(' | ');
}

function parseCertificateTypes_(value) {
  return String(value || '').split(',').map(part => {
    const match = part.trim().match(/^(.*?)\s*\((\d+)\)$/);
    return match ? {type: match[1].trim(), quantity: Math.max(0, Number(match[2]) || 0)} : null;
  }).filter(item => item && item.type && item.quantity > 0);
}

function certificateTypeCode_(typeName) {
  const code = String(typeName || '').replace(/Certificate of /i, '').replace(/[^A-Za-z]/g, '').toUpperCase().substring(0, 4);
  return code || 'CERT';
}

function buildCertificateBreakdown_(typesString, controlNumberString, status) {
  const types = parseCertificateTypes_(typesString);
  const raw = String(controlNumberString || '');
  const result = [];
  let legacyNumbers = [];
  // Older requests stored all control numbers as one pipe-separated list.
  if (raw && raw.indexOf(':') < 0) legacyNumbers = raw.split('|').map(x => x.trim()).filter(Boolean);
  let legacyOffset = 0;
  types.forEach(item => {
    let numbers = [];
    const group = raw.split(' | ').find(part => part.indexOf(item.type + ' (') === 0);
    if (group && group.indexOf(':') >= 0) {
      numbers = group.substring(group.indexOf(':') + 1).split(',').map(x => x.trim()).filter(Boolean);
    } else if (legacyNumbers.length) {
      numbers = legacyNumbers.slice(legacyOffset, legacyOffset + item.quantity);
      legacyOffset += item.quantity;
    }
    result.push({
      type: item.type,
      quantity: item.quantity,
      status: certificateStatusForRequest_(status),
      controlNumbers: numbers
    });
  });
  return result;
}

function certificateStatusForRequest_(status) {
  const s = String(status || 'Submitted');
  if (s === 'Released') return 'Released';
  if (s === 'Approved - For Release') return 'Approved - For Release';
  if (['Returned to Proponent','Returned for Revision'].includes(s)) return 'Returned for Revision';
  if (s === 'Disapproved') return 'Disapproved';
  if (s === 'Awaiting Approvals') return 'Awaiting Approval';
  if (s === 'Draft in Preparation') return 'Draft in Preparation';
  if (s === 'Resubmitted') return 'Resubmitted';
  return s;
}

function apiReleaseCertificate(token, p) {
  const u = staffSession_(token);
  requireRole_(u, ['ADMIN','HRD']);

  const r = requestById_(p.requestId);
  if (!r || r.Status !== 'Approved - For Release') {
    throw new Error('All required approvals must be completed before release.');
  }

  checkHttps_(p.finalURL);

  updateRequest_(r.RequestID, {
    Status:'Released',
    FinalURL:p.finalURL,
    ReleasedAt:new Date()
  });

  sendCertificateReleaseEmail_(r.ProponentEmail, r.ProponentName, r.RequestID, r.ActivityTitle, r.ControlNumber, p.finalURL);
  audit_(r.RequestID, u, 'CERTIFICATE_RELEASED', 'Released at '+formatDateTime_(new Date())+'. Certificates and quantities: '+r.CertificateTypes+'. Per-type control numbers: '+r.ControlNumber+'. Final link recorded and release email attempted.');
  return {ok:true};
}

function sendRevisionEmail_(toEmail, name, requestId, title, comments) {
  const subject = `[HRD Action Required] Request Returned - ${requestId}`;
  const body = `Dear ${name},\n\n` +
    `Your request for "${title}" (${requestId}) has been returned for revision.\n` +
    `Date/Time: ${formatDateTime_(new Date())}\n\n` +
    `Reason / Feedback:\n${comments}\n\n` +
    `Please address the feedback, email your updated files to hrds.laguna@deped.gov.ph, and click "Resubmit Request" on the tracking portal.\n\n` +
    `Thank you.\nHRD Office - Schools Division of Laguna`;
  return sendSystemEmail_(toEmail, subject, body);
}

function sendDisapprovalEmail_(toEmail, name, requestId, title, comments) {
  const subject = `[HRD Certificate Notice] Request Disapproved - ${requestId}`;
  const body = `Dear ${name},\n\n` +
    `Your request for "${title}" (${requestId}) has been disapproved.\n` +
    `Date/Time: ${formatDateTime_(new Date())}\n\n` +
    `Reason for Disapproval:\n${comments}\n\n` +
    `Thank you.\nHRD Office - Schools Division of Laguna`;
  return sendSystemEmail_(toEmail, subject, body);
}

function sendCertificateReleaseEmail_(toEmail, name, requestId, title, controlNumbers, finalUrl) {
  const subject = `[HRD Certificate Released] Official Link for ${requestId}`;
  const body = `Dear ${name},\n\n` +
    `Your certificate request for "${title}" (${requestId}) has been approved and officially released.\n\n` +
    `Date/Time Released: ${formatDateTime_(new Date())}\n` +
    `Assigned Per-Certificate Control Numbers (one number per requested certificate quantity):\n${controlNumbers}\n\n` +
    `Download Link:\n${finalUrl}\n\n` +
    `Thank you.\nHRD Office - Schools Division of Laguna`;
  return sendSystemEmail_(toEmail, subject, body);
}

function apiGetAudit(token, requestId) {
  const u = staffSession_(token);
  requireRole_(u, ['ADMIN','HRD','APPROVER']);
  return objects_(CFG.AUDIT).filter(a => a.RequestID === requestId && !['EMAIL_FAILED','STAFF_LOGIN','STAFF_LOGOUT','TRANSACTION_BACKUP_EXPORTED'].includes(String(a.Action || '').toUpperCase()))
    .sort((a,b) => new Date(a.Timestamp) - new Date(b.Timestamp))
    .map(a => ({Timestamp:a.Timestamp, Action:friendlyAuditAction_(a.Action), Details:publicAuditDetails_(a.Action,a.Details)}));
}

function friendlyAuditAction_(action) {
  const labels = {
    SUBMITTED:'Request submitted', RESUBMITTED:'Request resubmitted', 'PREPARE DRAFT':'Certificate draft prepared',
    'SEND FOR APPROVAL':'Draft sent for approval', 'RETURN TO PROPONENT':'Returned to requester for revision',
    'DISAPPROVE REQUEST':'Request disapproved', APPROVED:'Approval recorded', REJECTED:'Revision requested by approver',
    'RETURNED FOR REVISION':'Returned for revision', 'ALL_APPROVALS_COMPLETE':'All approvals completed',
    'CERTIFICATE_RELEASED':'Certificate released', 'APPROVE':'Approved', 'REJECT':'Returned for revision'
  };
  const key = String(action || '').toUpperCase();
  return labels[key] || String(action || '').toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
}

function publicAuditDetails_(action, details) {
  const key = String(action || '').toUpperCase();
  let value = String(details || '');
  value = value.replace(/^Recorded at [^.]+\.\s*/i, '').trim();
  if (['SUBMITTED','RESUBMITTED','PREPARE DRAFT','SEND FOR APPROVAL','ALL_APPROVALS_COMPLETE','APPROVED','CERTIFICATE_RELEASED'].includes(key)) {
    if (key === 'SUBMITTED') return 'Your request has been received. ' + certificateAuditSummary_(value);
    if (key === 'RESUBMITTED') return 'Your updated request has been received. ' + certificateAuditSummary_(value);
    if (key === 'PREPARE DRAFT') return value || 'The certificate draft is being prepared.';
    if (key === 'SEND FOR APPROVAL') return 'The certificate draft is now under approval.';
    if (key === 'ALL_APPROVALS_COMPLETE') return 'All required approvals have been completed.';
    if (key === 'APPROVED') return 'An approval has been recorded.';
    if (key === 'CERTIFICATE_RELEASED') return 'The certificate has been released.';
  }
  return value;
}

function certificateAuditSummary_(value) {
  const text = String(value || '');
  const match = text.match(/Certificates requested:\s*(.*?)(?:\.\s*Total quantity:|$)/i);
  const total = text.match(/Total quantity:\s*(\d+)/i);
  if (!match) return '';
  return 'Requested: ' + match[1].trim() + (total ? '. Total certificates: ' + total[1] + '.' : '');
}

function apiExportSummaryReport(token) {
  const u = staffSession_(token);
  requireRole_(u, ['ADMIN','HRD']);

  const requests = objects_(CFG.REQUESTS);
  let csv = 'Request ID,Date Requested,Proponent Name,Email,Office/School,Activity Title,Activity Type,Certificates Requested,Total Qty,Start Date,End Date,Urgency Status,Current Status,Control Numbers\n';

  requests.forEach(r => {
    csv += `"${r.RequestID}","${r.SubmittedAt}","${r.ProponentName}","${r.ProponentEmail}","${r.OfficeSchool}","${r.ActivityTitle}","${r.ActivityType}","${r.CertificateTypes}","${r.TotalQuantity}","${r.ActivityStart}","${r.ActivityEnd}","${r.UrgencyStatus}","${r.Status}","${r.ControlNumber || ''}"\n`;
  });

  return csv;
}

function apiGetAnalytics(token) {
  const u = staffSession_(token);
  requireRole_(u, ['ADMIN','HRD']);
  const requests = objects_(CFG.REQUESTS);
  const byStatusMap = {};
  const certMap = {};
  requests.forEach(r => {
    const status = String(r.Status || 'Unknown');
    byStatusMap[status] = (byStatusMap[status] || 0) + 1;
    String(r.CertificateTypes || '').split(',').forEach(part => {
      const m = part.trim().match(/^(.*?)\s*\((\d+)\)$/);
      if (m) certMap[m[1].trim()] = (certMap[m[1].trim()] || 0) + Number(m[2]);
    });
  });
  return {
    generatedAt: formatDateTime_(new Date()),
    totalRequests: requests.length,
    urgentRequests: requests.filter(r => String(r.UrgencyStatus || '').toUpperCase().includes('URGENT')).length,
    releasedRequests: requests.filter(r => r.Status === 'Released').length,
    byStatus: Object.keys(byStatusMap).sort().map(status => ({status, count:byStatusMap[status]})),
    byCertificateType: Object.keys(certMap).sort().map(type => ({type, quantity:certMap[type]}))
  };
}

// Transaction backup intentionally excludes Users/password hashes.
function apiExportTransactionBackup(token) {
  const u = staffSession_(token);
  requireRole_(u, ['ADMIN','HRD']);
  const payload = {
    backupType: 'HRD_CERTIFICATE_TRANSACTION_BACKUP',
    schemaVersion: 1,
    exportedAt: formatDateTime_(new Date()),
    exportedBy: u.email,
    timezone: CFG.TZ,
    transactions: {
      requests: objects_(CFG.REQUESTS),
      approvals: objects_(CFG.APPROVALS),
      auditTrail: objects_(CFG.AUDIT)
    }
  };
  audit_('SYSTEM', u, 'TRANSACTION_BACKUP_EXPORTED', 'Transaction backup downloaded at '+formatDateTime_(new Date()));
  return JSON.stringify(payload, null, 2);
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,headers.length).setValues([headers]);
  }
}

function objects_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const data = sh.getDataRange().getValues();
  const headers = data.shift();

  return data.filter(row => row.some(v => v !== '' && v !== null))
    .map(row => {
      const o = {};
      headers.forEach((h,i) => o[h] = serialize_(row[i]));
      return o;
    });
}

function staffRows_() { return objects_(CFG.USERS); }

function createStaffAccount_(email, fullName, role, password, office, approvalStep) {
  email = email_(email);
  role = String(role || '').toUpperCase();
  password = String(password || '');

  if (!email || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Enter a valid staff email.');
  if (!['ADMIN', 'HRD', 'APPROVER'].includes(role)) throw new Error('Role must be ADMIN, HRD, or APPROVER.');
  if (password.length < 12) throw new Error('Password must be at least 12 characters.');

  const step = Number(approvalStep || 0);
  if (role === 'APPROVER' && ![1,2,3,4].includes(step)) throw new Error('Approver step must be 1, 2, 3, or 4.');

  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.USERS);
  if (staffRows_().some(u => email_(u.Email) === email)) throw new Error('Email is already registered.');

  const salt = Utilities.getUuid() + Utilities.getUuid();
  sh.appendRow([email, String(fullName || email), role, true, String(office || ''), step || '', salt, passwordHash_(password, salt)]);

  return 'Staff account created for ' + email;
}

function configuredApprovers_() {
  const users = staffRows_().filter(u => String(u.Role).toUpperCase() === 'APPROVER' && isTrue_(u.Active));
  const out = [];
  for (let step = 1; step <= 4; step++) {
    const matches = users.filter(u => Number(u.ApprovalStep) === step);
    if (matches.length > 0) {
      out.push({ step: step, email: email_(matches[0].Email), name: String(matches[0].FullName || '') });
    }
  }
  return out;
}

function resetApprovals_(requestId) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.APPROVALS);
  const data = sh.getDataRange().getValues();
  const ix = headerIndex_(data[0]);

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][ix.RequestID]) === requestId) {
      sh.getRange(i + 1, ix.Decision + 1).setValue('Pending');
      sh.getRange(i + 1, ix.Comments + 1).clearContent();
      sh.getRange(i + 1, ix.ActionAt + 1).clearContent();
    }
  }
}

function requestById_(id) { return objects_(CFG.REQUESTS).find(r => r.RequestID === String(id)); }

function updateRequest_(id, updates) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.REQUESTS);
  const data = sh.getDataRange().getValues();
  const ix = headerIndex_(data[0]);

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][ix.RequestID]) === String(id)) {
      Object.keys(updates).forEach(k => {
        if (ix[k] !== undefined) sh.getRange(i + 1, ix[k] + 1).setValue(updates[k]);
      });
      sh.getRange(i + 1, ix.UpdatedAt + 1).setValue(new Date());
      return;
    }
  }
  throw new Error('Request not found.');
}

function audit_(id, u, action, details) {
  SpreadsheetApp.getActive().getSheetByName(CFG.AUDIT).appendRow([new Date(), id, u.email, u.role, action, details || '']);
}

function requireRole_(u, roles) { if (!roles.includes(u.role)) throw new Error('Permission denied.'); }

function setting_(key) {
  const item = objects_(CFG.SETTINGS).find(x => x.Key === key);
  return item ? item.Value : '';
}

function headerIndex_(headers) {
  const out = {};
  headers.forEach((h,i) => out[h] = i);
  return out;
}

function serialize_(v) {
  return v instanceof Date ? Utilities.formatDate(v, CFG.TZ, "yyyy-MM-dd'T'HH:mm:ss") : v;
}

function email_(v) { return String(v || '').trim().toLowerCase(); }
function isTrue_(v) { return v === true || ['true','yes','1'].includes(String(v).toLowerCase()); }

function checkHttps_(url) {
  if (!/^https:\/\/\S+$/i.test(String(url || ''))) throw new Error('Provide a valid HTTPS link.');
}

function passwordHash_(password, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(salt) + ':' + String(password), Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(bytes);
}

function secretHash_(value) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(value), Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(bytes);
}

function randomSecret_() { return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, ''); }
// Run once manually from the Apps Script editor to authorize MailApp.
function authorizeEmailNotifications() {
  const remaining = MailApp.getRemainingDailyQuota();
  return 'Email service authorization checked. Remaining recipient quota today: ' + remaining;
}
