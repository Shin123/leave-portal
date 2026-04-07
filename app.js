// =========================================================================
// LEAVE REQUEST PORTAL — app.js
// Kết nối SharePoint Online qua REST API + MSAL.js (Microsoft Identity)
// =========================================================================

// ─── CONFIGURATION ───────────────────────────────────────────────────────────
// ⚠️  BẠN CẦN CẬP NHẬT 3 GIÁ TRỊ NÀY THEO TENANT CỦA BẠN
const CONFIG = {
  // 1) Azure AD App Registration → Application (client) ID
  clientId: 'b6de2590-c2aa-44f5-9ed7-0c335eb7d7e8',

  // 2) Azure AD → Tenant ID
  tenantId: '922f6579-71c0-482f-881b-25ddba79b524',

  // 3) SharePoint site URL (không có trailing slash)
  sharePointSiteUrl: 'https://masterdx1.sharepoint.com/sites/MDX-OrienMDMteam',

  // Tên SharePoint List (đúng như trên SharePoint)
  leaveListName: 'Leave List',
  leaveBalanceName: 'LeaveBalance',

  // Redirect URI — phải trùng với cấu hình trong Azure AD App Registration
  redirectUri: window.location.origin + window.location.pathname,
};

// ─── MSAL CONFIG ─────────────────────────────────────────────────────────────
const msalConfig = {
  auth: {
    clientId: CONFIG.clientId,
    authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
    redirectUri: CONFIG.redirectUri,
    navigateToLoginRequestUrl: true, // Quay lại đúng trang sau khi login
  },
  cache: {
    cacheLocation: 'localStorage',        // Giữ session qua tab/refresh
    storeAuthStateInCookie: true,          // Fix lỗi IE/Edge + cross-tab
  },
  system: {
    allowRedirectInIframe: false,
    loggerOptions: {
      logLevel: 0, // Error only
    },
  },
};

// Login scopes - chỉ cần cơ bản
const loginRequest = {
  scopes: ['openid', 'profile', 'User.Read'],
};

// SharePoint API scopes - yêu cầu khi gọi REST API
const sharePointScopes = [
  'https://masterdx1.sharepoint.com/AllSites.Read',
  'https://masterdx1.sharepoint.com/AllSites.Write',
];

// ─── APP STATE ───────────────────────────────────────────────────────────────
let msalInstance = null;
let currentAccount = null;
let accessToken = null;

// Data caches
let leaveRequests = [];
let leaveBalance = null;
let currentFilter = 'all';
let autoRefreshTimer = null;
const AUTO_REFRESH_INTERVAL = 900000; // 15 phút

// ─── DEMO MODE ───────────────────────────────────────────────────────────────
// Khi chưa cấu hình Azure AD, app chạy với dữ liệu demo
const IS_DEMO = CONFIG.clientId === 'YOUR_CLIENT_ID_HERE';

const DEMO_USER = {
  name: 'Nguyễn Văn Việt',
  username: 'viet.nguyenvan@masterdx.com',
};

const DEMO_BALANCE = {
  EmployeeEmail: 'viet.nguyenvan@masterdx.com',
  JoinDate: '2026-03-20',
  AnnualQuota: 49,
  DaysTaken: 10,
  RemainingDays: 39,
  Status: 'Confirmed',
};

const DEMO_REQUESTS = [
  {
    ID: 1, Title: 'viet.nguyenvan@masterdx.com', Typeofleave: 'Nghỉ phép năm',
    FromDate: '2026-04-05', ToDate: '2026-04-06', DayLeave: 1,
    Reason: 'Nghỉ phép đi du lịch cùng gia đình', Status: 'Approved',
    ManagementEmail: 'manager@masterdx.com', ApprovalTime: '2026-04-02',
    Note: '', RemainingVacationDays: 48, NumberofVacationDaysUsed: 1,
  },
  {
    ID: 2, Title: 'viet.nguyenvan@masterdx.com', Typeofleave: 'Nghỉ ốm',
    FromDate: '2026-04-05', ToDate: '2026-04-14', DayLeave: 9,
    Reason: 'Bị cảm cúm, cần nghỉ điều trị', Status: 'Approved',
    ManagementEmail: 'manager@masterdx.com', ApprovalTime: '2026-04-02',
    Note: '', RemainingVacationDays: 40, NumberofVacationDaysUsed: 9,
  },
  {
    ID: 3, Title: 'viet.nguyenvan@masterdx.com', Typeofleave: 'Nghỉ ốm',
    FromDate: '2026-04-13', ToDate: '2026-04-14', DayLeave: 1,
    Reason: 'Khám sức khỏe định kỳ', Status: 'Approved',
    ManagementEmail: 'manager@masterdx.com', ApprovalTime: '2026-04-02',
    Note: '', RemainingVacationDays: 48, NumberofVacationDaysUsed: 1,
  },
  {
    ID: 4, Title: 'viet.nguyenvan@masterdx.com', Typeofleave: 'Nghỉ không phép',
    FromDate: '2026-04-03', ToDate: '2026-04-05', DayLeave: 2,
    Reason: 'Việc gia đình đột xuất', Status: 'Approved',
    ManagementEmail: 'manager@masterdx.com', ApprovalTime: '2026-04-02',
    Note: '', RemainingVacationDays: 46, NumberofVacationDaysUsed: 3,
  },
  {
    ID: 5, Title: 'viet.nguyenvan@masterdx.com', Typeofleave: 'Nghỉ ốm',
    FromDate: '2026-04-15', ToDate: '2026-04-17', DayLeave: 2,
    Reason: '', Status: 'Waiting',
    ManagementEmail: 'manager@masterdx.com', ApprovalTime: null,
    Note: '', RemainingVacationDays: 44, NumberofVacationDaysUsed: 5,
  },
  {
    ID: 6, Title: 'viet.nguyenvan@masterdx.com', Typeofleave: 'Nghỉ phép năm',
    FromDate: '2026-04-12', ToDate: '2026-04-13', DayLeave: 1,
    Reason: '', Status: 'Rejected',
    ManagementEmail: 'manager@masterdx.com', ApprovalTime: '2026-04-02',
    Note: '', RemainingVacationDays: 39, NumberofVacationDaysUsed: 10,
  },
  {
    ID: 7, Title: 'staff1@masterdx.com', Typeofleave: 'Nghỉ phép năm',
    FromDate: '2026-04-20', ToDate: '2026-04-22', DayLeave: 2,
    Reason: 'Đi họp phụ huynh', Status: 'Waiting',
    ManagementEmail: 'viet.nguyenvan@masterdx.com', ApprovalTime: null,
    Note: '',
  },
  {
    ID: 8, Title: 'staff2@masterdx.com', Typeofleave: 'Nghỉ ốm',
    FromDate: '2026-04-18', ToDate: '2026-04-19', DayLeave: 1,
    Reason: 'Đi khám răng', Status: 'Waiting',
    ManagementEmail: 'viet.nguyenvan@masterdx.com', ApprovalTime: null,
    Note: '',
  },
];

// ─── INIT ────────────────────────────────────────────────────────────────────
let msalReady = false;
let loginInProgress = false;

document.addEventListener('DOMContentLoaded', async () => {
  // Nếu đang chạy trong popup → bỏ qua
  if (window.opener && window.opener !== window) return;

  updateHeaderDate();

  if (IS_DEMO) {
    console.log('🟡 Running in DEMO mode');
    msalReady = true;
    return;
  }

  try {
    msalInstance = new msal.PublicClientApplication(msalConfig);

    // 1. Xử lý redirect callback TRƯỚC TIÊN (bắt buộc)
    let redirectResp = null;
    try {
      redirectResp = await msalInstance.handleRedirectPromise();
    } catch (redirectErr) {
      console.warn('Redirect handling error:', redirectErr);
      // Xóa state kẹt nếu có
      clearMsalState();
    }

    msalReady = true;

    // 2. Nếu vừa redirect về từ login
    if (redirectResp && redirectResp.account) {
      currentAccount = redirectResp.account;
      onLoginSuccess();
      return;
    }

    // 3. Kiểm tra session cũ
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      currentAccount = accounts[0];
      onLoginSuccess();
    }
  } catch (e) {
    console.error('MSAL init error:', e);
    clearMsalState();
    msalReady = true;
  }
});

// Xóa MSAL state bị kẹt trong storage
function clearMsalState() {
  ['sessionStorage', 'localStorage'].forEach(storageType => {
    const storage = window[storageType];
    Object.keys(storage).forEach(key => {
      if (key.includes('msal.interaction')) {
        storage.removeItem(key);
      }
    });
  });
}

// ─── DATE HEADER ─────────────────────────────────────────────────────────────
function updateHeaderDate() {
  const el = document.getElementById('headerDate');
  const now = new Date();
  const opts = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
  el.textContent = now.toLocaleDateString('vi-VN', opts);
}

// ─── AUTH ────────────────────────────────────────────────────────────────────
async function handleLogin() {
  if (IS_DEMO) {
    currentAccount = DEMO_USER;
    onLoginSuccess();
    return;
  }

  if (!msalReady || loginInProgress) {
    showToast('⏳ Đang xử lý, vui lòng đợi...', 'info');
    return;
  }

  loginInProgress = true;
  const btn = document.getElementById('btnLogin');
  if (btn) btn.disabled = true;

  // Xóa state cũ trước khi login (phòng lỗi interaction_in_progress)
  clearMsalState();

  // Redirect — ổn định nhất, không popup
  try {
    await msalInstance.loginRedirect(loginRequest);
  } catch (err) {
    console.error('Login redirect error:', err);
    loginInProgress = false;
    if (btn) btn.disabled = false;
    showToast('Lỗi đăng nhập. Vui lòng thử lại.', 'error');
  }
}

function handleLogout() {
  if (IS_DEMO) {
    document.getElementById('appLayout').classList.remove('active');
    document.getElementById('loginScreen').style.display = 'flex';
    currentAccount = null;
    return;
  }
  msalInstance.logoutRedirect({ postLogoutRedirectUri: CONFIG.redirectUri });
}

async function getAccessToken() {
  if (IS_DEMO) return 'DEMO_TOKEN';

  const tokenRequest = {
    scopes: sharePointScopes,
    account: currentAccount,
  };

  try {
    // 1. Thử lấy token im lặng (từ cache)
    const resp = await msalInstance.acquireTokenSilent(tokenRequest);
    return resp.accessToken;
  } catch (silentErr) {
    console.warn('Silent token failed, trying redirect...');
    try {
      // 2. Thử lại với forceRefresh
      const resp = await msalInstance.acquireTokenSilent({ ...tokenRequest, forceRefresh: true });
      return resp.accessToken;
    } catch (retryErr) {
      // 3. Redirect để lấy token mới
      clearMsalState();
      msalInstance.acquireTokenRedirect({ scopes: sharePointScopes });
      return null; // Trang sẽ redirect, không tiếp tục
    }
  }
}

function onLoginSuccess() {
  document.getElementById('loginScreen').style.display = 'none';
  document.getElementById('appLayout').classList.add('active');

  // Update user info
  const displayName = currentAccount.name || currentAccount.username || DEMO_USER.name;
  const email = currentAccount.username || DEMO_USER.username;

  document.getElementById('userName').textContent = displayName;
  document.getElementById('userEmail').textContent = email;
  document.getElementById('userAvatar').textContent = getInitials(displayName);
  document.getElementById('formEmployee').textContent = email;

  // Load data
  loadDashboard();
}

function getInitials(name) {
  return name
    .split(' ')
    .map(w => w[0])
    .join('')
    .toUpperCase()
    .slice(0, 2);
}

// ─── SHAREPOINT REST API HELPERS ─────────────────────────────────────────────
async function spGet(listName, filter, select, orderby, top) {
  if (IS_DEMO) return getDemoData(listName, filter);

  const token = await getAccessToken();
  if (!token) return []; // Đang redirect để lấy token, trả về rỗng

  let url = `${CONFIG.sharePointSiteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items?`;
  const params = [];
  if (select) params.push(`$select=${select}`);
  if (filter) params.push(`$filter=${filter}`);
  if (orderby) params.push(`$orderby=${orderby}`);
  if (top) params.push(`$top=${top}`);
  url += params.join('&');

  const resp = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/json;odata=nometadata',
    },
  });

  if (resp.status === 401) {
    // Token hết hạn → redirect lấy token mới
    clearMsalState();
    msalInstance.acquireTokenRedirect({ scopes: sharePointScopes });
    return [];
  }

  if (!resp.ok) throw new Error(`SharePoint API error: ${resp.status}`);
  const data = await resp.json();
  return data.value;
}

async function spCreate(listName, itemData) {
  if (IS_DEMO) return createDemoItem(listName, itemData);

  const token = await getAccessToken();
  if (!token) throw new Error('Đang chuyển hướng để xác thực...');

  const url = `${CONFIG.sharePointSiteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items`;

  const resp = await fetch(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=nometadata',
    },
    body: JSON.stringify(itemData),
  });

  if (!resp.ok) {
    const errText = await resp.text();
    throw new Error(`Create failed: ${resp.status} - ${errText}`);
  }
  return resp.json();
}

async function spUpdate(listName, itemId, itemData) {
  if (IS_DEMO) return updateDemoItem(listName, itemId, itemData);

  const token = await getAccessToken();
  const url = `${CONFIG.sharePointSiteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items(${itemId})`;

  const resp = await fetch(url, {
    method: 'PATCH',
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=nometadata',
      'IF-MATCH': '*',
      'X-HTTP-Method': 'MERGE',
    },
    body: JSON.stringify(itemData),
  });

  if (!resp.ok) {
    const errText = await resp.text();
    throw new Error(`Update failed: ${resp.status} - ${errText}`);
  }
}

// ─── DEMO DATA HELPERS ──────────────────────────────────────────────────────
function getDemoData(listName, filter) {
  if (listName === CONFIG.leaveBalanceName) {
    return [DEMO_BALANCE];
  }
  return [...DEMO_REQUESTS];
}

function createDemoItem(listName, data) {
  const newItem = {
    ID: DEMO_REQUESTS.length + 1,
    ...data,
    Status: 'Waiting',
  };
  DEMO_REQUESTS.unshift(newItem);
  return newItem;
}

function updateDemoItem(listName, itemId, data) {
  const idx = DEMO_REQUESTS.findIndex(r => r.ID === itemId);
  if (idx !== -1) {
    Object.assign(DEMO_REQUESTS[idx], data);
  }
}

// ─── NAVIGATION ──────────────────────────────────────────────────────────────
const PAGE_TITLES = {
  dashboard: 'Dashboard',
  'new-request': 'Tạo đơn nghỉ phép',
  'my-requests': 'Đơn nghỉ của tôi',
};

let currentPage = 'dashboard';

function navigateTo(page) {
  currentPage = page;

  // Update nav
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  const navBtn = document.querySelector(`.nav-item[data-page="${page}"]`);
  if (navBtn) navBtn.classList.add('active');

  // Update page views
  document.querySelectorAll('.page-view').forEach(v => v.classList.remove('active'));
  document.getElementById(`page-${page}`).classList.add('active');

  // Update title
  document.getElementById('pageTitle').textContent = PAGE_TITLES[page] || 'Dashboard';

  // Close sidebar on mobile
  document.getElementById('sidebar').classList.remove('open');
  document.getElementById('sidebarOverlay').classList.remove('active');

  // Load data for the page
  if (page === 'dashboard') loadDashboard();
  if (page === 'my-requests') loadMyRequests();
  if (page === 'new-request') loadNewRequestForm();

  // Auto-refresh cho dashboard và my-requests
  startAutoRefresh(page);
}

function startAutoRefresh(page) {
  // Xóa timer cũ
  if (autoRefreshTimer) {
    clearInterval(autoRefreshTimer);
    autoRefreshTimer = null;
  }

  // Chỉ auto-refresh cho dashboard và my-requests
  if (page === 'dashboard' || page === 'my-requests') {
    autoRefreshTimer = setInterval(() => {
      console.log('🔄 Auto-refreshing data...');
      if (currentPage === 'dashboard') loadDashboard();
      else if (currentPage === 'my-requests') loadMyRequests();
    }, AUTO_REFRESH_INTERVAL);
  }
}

// Refresh thủ công
function refreshData() {
  if (currentPage === 'dashboard') loadDashboard();
  else if (currentPage === 'my-requests') loadMyRequests();
  showToast('🔄 Dữ liệu đã được cập nhật!', 'success');
}

function toggleSidebar() {
  document.getElementById('sidebar').classList.toggle('open');
  document.getElementById('sidebarOverlay').classList.toggle('active');
}

// ─── LOAD DASHBOARD ──────────────────────────────────────────────────────────
async function loadDashboard() {
  try {
    // Load balance
    const balanceData = await spGet(CONFIG.leaveBalanceName);
    const userEmail = currentAccount.username || DEMO_USER.username;
    
    console.log('🔍 User email:', userEmail);

    // LeaveBalance dùng cột Title để lưu EmployeeEmail
    leaveBalance = balanceData.find(b => {
      const email = (b.Title || '').toLowerCase().trim();
      return email === userEmail.toLowerCase().trim();
    });

    if (!leaveBalance) {
      console.warn('⚠️ No LeaveBalance found for', userEmail);
      leaveBalance = { AnnualQuota: 0, DaysTaken: 0, RemainingDays: 0 };
    } else {
      console.log('✅ Found balance:', leaveBalance);
    }

    // Load requests
    leaveRequests = await spGet(CONFIG.leaveListName);
    const myRequests = leaveRequests.filter(
      r => r.Title?.toLowerCase() === userEmail.toLowerCase()
    );

    // Stats - không làm tròn để giữ 0.5
    const quota = leaveBalance?.AnnualQuota || 0;
    const remaining = leaveBalance?.RemainingDays || 0;
    const taken = leaveBalance?.DaysTaken || 0;
    const pending = myRequests.filter(r => r.Status === 'Waiting').length;

    animateNumber('statQuota', quota);
    animateNumber('statRemaining', remaining);
    animateNumber('statTaken', taken);
    animateNumber('statPending', pending);

    // Balance overview
    document.getElementById('balanceNumber').textContent = formatNum(remaining);
    document.getElementById('barUsed').textContent = formatNum(taken);
    document.getElementById('barTotal').textContent = formatNum(quota);

    const pct = quota > 0 ? ((remaining / quota) * 100) : 0;
    setTimeout(() => {
      document.getElementById('balanceBar').style.width = pct + '%';
    }, 300);

    // Detail breakdown
    const countByType = (type) => myRequests
      .filter(r => r.Typeofleave === type && r.Status === 'Approved')
      .reduce((sum, r) => sum + (r.DayLeave || 0), 0);

    document.getElementById('detailAnnual').textContent = countByType('Nghỉ phép năm');
    document.getElementById('detailSick').textContent = countByType('Nghỉ ốm');
    document.getElementById('detailUnpaid').textContent = countByType('Nghỉ không phép');
    document.getElementById('detailOther').textContent = countByType('Khác');

    // WFH count (dùng DayWFH thay vì DayLeave)
    const wfhDays = myRequests
      .filter(r => r.Typeofleave === 'WFH' && r.Status === 'Approved')
      .reduce((sum, r) => sum + (r.DayWFH || r.DayLeave || 0), 0);
    document.getElementById('detailWfh').textContent = wfhDays;

    // Recent requests (last 5)
    const sorted = [...myRequests].sort((a, b) => new Date(b.FromDate) - new Date(a.FromDate));
    renderRecentRequests(sorted.slice(0, 5));

  } catch (err) {
    console.error('Load dashboard error:', err);
    showToast('Lỗi tải dữ liệu: ' + err.message, 'error');
  }
}

// Format số: 2.5 -> "2.5", 3.0 -> "3"
function formatNum(n) {
  return Number.isInteger(n) ? n : parseFloat(n.toFixed(1));
}

function animateNumber(elementId, target) {
  const el = document.getElementById(elementId);
  const duration = 800;
  const start = performance.now();
  const initial = parseFloat(el.textContent) || 0;
  const isDecimal = !Number.isInteger(target);

  function step(now) {
    const elapsed = now - start;
    const progress = Math.min(elapsed / duration, 1);
    const eased = 1 - Math.pow(1 - progress, 3);
    const current = initial + (target - initial) * eased;
    el.textContent = progress >= 1 ? formatNum(target) : (isDecimal ? current.toFixed(1) : Math.round(current));
    if (progress < 1) requestAnimationFrame(step);
  }
  requestAnimationFrame(step);
}

// ─── RENDER TABLES ───────────────────────────────────────────────────────────
function formatDate(dateStr) {
  if (!dateStr) return '—';
  const d = new Date(dateStr);
  return d.toLocaleDateString('vi-VN', { day: '2-digit', month: '2-digit', year: 'numeric' });
}

function getLeaveTypeClass(type) {
  if (type === 'Nghỉ phép năm') return 'annual';
  if (type === 'Nghỉ ốm') return 'sick';
  if (type === 'Nghỉ không phép') return 'unpaid';
  if (type === 'WFH') return 'wfh';
  return 'other';
}

function getStatusBadge(status) {
  const cls = status === 'Approved' ? 'badge-approved'
    : status === 'Rejected' ? 'badge-rejected'
    : 'badge-waiting';
  const label = status === 'Approved' ? 'Đã duyệt'
    : status === 'Rejected' ? 'Từ chối'
    : 'Chờ duyệt';
  return `<span class="badge ${cls}"><span class="badge-dot"></span>${label}</span>`;
}

function renderRecentRequests(items) {
  const tbody = document.getElementById('recentRequestsBody');
  if (items.length === 0) {
    tbody.innerHTML = `<tr><td colspan="6">
      <div class="empty-state">
        <div class="empty-icon">📭</div>
        <div class="empty-title">Chưa có đơn nghỉ nào</div>
        <div class="empty-subtitle">Tạo đơn nghỉ phép đầu tiên của bạn</div>
      </div>
    </td></tr>`;
    return;
  }

  tbody.innerHTML = items.map(r => `
    <tr onclick="showDetail(${r.ID})" style="cursor:pointer">
      <td><span class="leave-type ${getLeaveTypeClass(r.Typeofleave)}">${r.Typeofleave}</span></td>
      <td>${formatDate(r.FromDate)}</td>
      <td>${formatDate(r.ToDate)}</td>
      <td><strong>${r.DayLeave || 0}</strong></td>
      <td><strong style="color:var(--accent-cyan)">${r.DayWFH || 0}</strong></td>
      <td>${getStatusBadge(r.Status)}</td>
    </tr>
  `).join('');
}

// ─── MY REQUESTS ─────────────────────────────────────────────────────────────
async function loadMyRequests() {
  try {
    leaveRequests = await spGet(CONFIG.leaveListName);
    const userEmail = currentAccount.username || DEMO_USER.username;
    const myItems = leaveRequests.filter(
      r => r.Title?.toLowerCase() === userEmail.toLowerCase()
    );
    renderMyRequests(myItems);
  } catch (err) {
    showToast('Lỗi tải danh sách: ' + err.message, 'error');
  }
}

function renderMyRequests(items) {
  const filtered = currentFilter === 'all'
    ? items
    : items.filter(r => r.Status === currentFilter);

  const sorted = [...filtered].sort((a, b) => new Date(b.FromDate) - new Date(a.FromDate));

  const tbody = document.getElementById('myRequestsBody');
  if (sorted.length === 0) {
    tbody.innerHTML = `<tr><td colspan="9">
      <div class="empty-state">
        <div class="empty-icon">📭</div>
        <div class="empty-title">Không có đơn nào</div>
        <div class="empty-subtitle">Không tìm thấy đơn nghỉ phù hợp</div>
      </div>
    </td></tr>`;
    return;
  }

  tbody.innerHTML = sorted.map(r => `
    <tr>
      <td><span class="leave-type ${getLeaveTypeClass(r.Typeofleave)}">${r.Typeofleave}</span></td>
      <td>${formatDate(r.FromDate)}</td>
      <td>${formatDate(r.ToDate)}</td>
      <td><strong>${r.DayLeave || 0}</strong></td>
      <td><strong style="color:var(--accent-cyan)">${r.DayWFH || 0}</strong></td>
      <td style="max-width:150px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis">${r.Reason || '—'}</td>
      <td>${getStatusBadge(r.Status)}</td>
      <td style="max-width:150px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; font-size:12px; color:var(--text-muted)">${r.Note || '—'}</td>
      <td>
        <button class="btn btn-ghost btn-sm btn-icon" onclick="showDetail(${r.ID})" title="Xem chi tiết">👁️</button>
      </td>
    </tr>
  `).join('');
}

function filterRequests(status, btn) {
  currentFilter = status;
  document.querySelectorAll('#page-my-requests .filter-chip').forEach(c => c.classList.remove('active'));
  btn.classList.add('active');
  loadMyRequests();
}

// ─── NEW REQUEST FORM ────────────────────────────────────────────────────────
function loadNewRequestForm() {
  const userEmail = currentAccount.username || DEMO_USER.username;
  document.getElementById('formEmployee').textContent = userEmail;

  // Balance stats
  const quota = Math.round(leaveBalance?.AnnualQuota || 0);
  const remaining = Math.round(leaveBalance?.RemainingDays || 0);
  const taken = Math.round(leaveBalance?.DaysTaken || 0);

  // Đếm đơn chờ duyệt
  const pending = leaveRequests.filter(
    r => r.Title?.toLowerCase() === userEmail.toLowerCase() && r.Status === 'Waiting'
  ).length;

  document.getElementById('formQuota').textContent = quota;
  document.getElementById('formRemaining').textContent = remaining;
  document.getElementById('formTaken').textContent = taken;
  document.getElementById('formPending').textContent = pending;

  // Set default dates
  const today = new Date();
  const tomorrow = new Date(today);
  tomorrow.setDate(tomorrow.getDate() + 1);

  document.getElementById('fromDate').value = formatDateInput(tomorrow);
  document.getElementById('toDate').value = formatDateInput(tomorrow);
  calculateDays();
}

function formatDateInput(d) {
  return d.toISOString().split('T')[0];
}

// Khi đổi loại nghỉ
function onLeaveTypeChange() {
  const type = document.getElementById('leaveType').value;
  const hint = document.getElementById('wfhHint');
  const daysOffGroup = document.getElementById('daysOffGroup');
  const daysWfhGroup = document.getElementById('daysWfhGroup');

  if (type === 'WFH') {
    if (hint) hint.style.display = 'block';
    if (daysOffGroup) daysOffGroup.style.display = 'none';
    if (daysWfhGroup) daysWfhGroup.style.display = 'block';
  } else {
    if (hint) hint.style.display = 'none';
    if (daysOffGroup) daysOffGroup.style.display = 'block';
    if (daysWfhGroup) daysWfhGroup.style.display = 'none';
  }
  calculateDays();
}

function calculateDays() {
  const fromVal = document.getElementById('fromDate').value;
  const toVal = document.getElementById('toDate').value;
  const isHalfDay = document.getElementById('halfDayToggle')?.checked;

  if (!fromVal || !toVal) {
    document.getElementById('formDaysOff').textContent = '0 ngày';
    return 0;
  }

  // Parse dates without timezone issues
  const [fy, fm, fd] = fromVal.split('-').map(Number);
  const [ty, tm, td] = toVal.split('-').map(Number);
  const from = new Date(fy, fm - 1, fd);
  const to = new Date(ty, tm - 1, td);

  if (to < from) {
    document.getElementById('formDaysOff').textContent = '0 ngày';
    return 0;
  }

  // Tính tất cả ngày (bao gồm cả T7/CN)
  const diffTime = to.getTime() - from.getTime();
  let count = Math.round(diffTime / (1000 * 60 * 60 * 24)) + 1;

  // Nghỉ nửa ngày
  if (isHalfDay) {
    count = 0.5;
    document.getElementById('toDate').value = fromVal;
  }

  const leaveType = document.getElementById('leaveType')?.value;
  if (leaveType === 'WFH') {
    document.getElementById('formDaysWfh').textContent = count + ' ngày';
  } else {
    document.getElementById('formDaysOff').textContent = count + ' ngày';
  }
  return count;
}

async function handleSubmitLeave(e) {
  e.preventDefault();

  const fromDate = document.getElementById('fromDate').value;
  const toDate = document.getElementById('toDate').value;
  const leaveType = document.getElementById('leaveType').value;
  const reason = document.getElementById('reason').value;
  const days = calculateDays();

  // Luôn kiểm tra ngày hợp lệ (kể cả WFH)
  if (fromDate > toDate) {
    showToast('Ngày kết thúc phải sau hoặc bằng ngày bắt đầu', 'error');
    return;
  }

  if (days <= 0) {
    showToast('Vui lòng chọn ngày hợp lệ', 'error');
    return;
  }

  const userEmail = currentAccount.username || DEMO_USER.username;

  const submitBtn = document.getElementById('btnSubmitLeave');
  submitBtn.disabled = true;
  submitBtn.innerHTML = '<div class="loading-spinner"></div> Đang gửi...';

  try {
    // WFH → DayWFH, Nghỉ phép → DayLeave
    const submitData = {
      Title: userEmail,
      Typeofleave: leaveType,
      FromDate: fromDate + 'T00:00:00',
      ToDate: toDate + 'T00:00:00',
      Reason: reason,
      Status: 'Waiting',
    };

    if (leaveType === 'WFH') {
      submitData.DayWFH = days;
      submitData.DayLeave = 0;
    } else {
      submitData.DayLeave = days;
      submitData.DayWFH = 0;
    }

    await spCreate(CONFIG.leaveListName, submitData);

    showToast('🎉 Gửi đơn thành công! Đơn đang chờ duyệt.', 'success');
    document.getElementById('leaveForm').reset();
    loadNewRequestForm();

    // Navigate to my requests after a short delay
    setTimeout(() => navigateTo('my-requests'), 1500);

  } catch (err) {
    showToast('Lỗi gửi đơn: ' + err.message, 'error');
  } finally {
    submitBtn.disabled = false;
    submitBtn.innerHTML = '📤 Gửi đơn';
  }
}

function resetForm() {
  document.getElementById('leaveForm').reset();
  loadNewRequestForm();
}

// ─── DETAIL MODAL ────────────────────────────────────────────────────────────
function showDetail(itemId) {
  const item = leaveRequests.find(r => r.ID === itemId) || DEMO_REQUESTS.find(r => r.ID === itemId);
  if (!item) return;

  const content = document.getElementById('detailContent');
  content.innerHTML = `
    <div class="detail-item">
      <div class="detail-label">Nhân viên</div>
      <div class="detail-value">${item.Title}</div>
    </div>
    <div class="detail-item">
      <div class="detail-label">Loại nghỉ</div>
      <div class="detail-value"><span class="leave-type ${getLeaveTypeClass(item.Typeofleave)}">${item.Typeofleave}</span></div>
    </div>
    <div class="detail-item">
      <div class="detail-label">Từ ngày</div>
      <div class="detail-value">${formatDate(item.FromDate)}</div>
    </div>
    <div class="detail-item">
      <div class="detail-label">Đến ngày</div>
      <div class="detail-value">${formatDate(item.ToDate)}</div>
    </div>
    <div class="detail-item">
      <div class="detail-label">Số ngày nghỉ</div>
      <div class="detail-value" style="font-size:20px; font-weight:700; color:var(--accent-blue)">${item.DayLeave || 0}</div>
    </div>
    <div class="detail-item">
      <div class="detail-label">Số ngày WFH</div>
      <div class="detail-value" style="font-size:20px; font-weight:700; color:var(--accent-cyan)">${item.DayWFH || 0}</div>
    </div>
    <div class="detail-item">
      <div class="detail-label">Trạng thái</div>
      <div class="detail-value">${getStatusBadge(item.Status)}</div>
    </div>
    <div class="detail-item full">
      <div class="detail-label">Lý do</div>
      <div class="detail-value">${item.Reason || '—'}</div>
    </div>
    <div class="detail-item">
      <div class="detail-label">Người duyệt</div>
      <div class="detail-value">${item.ManagementEmail || '—'}</div>
    </div>
    <div class="detail-item">
      <div class="detail-label">Thời gian duyệt</div>
      <div class="detail-value">${formatDate(item.ApprovalTime)}</div>
    </div>
    ${item.Note ? `
    <div class="detail-item full">
      <div class="detail-label">Ghi chú</div>
      <div class="detail-value">${item.Note}</div>
    </div>` : ''}
  `;

  document.getElementById('detailModal').classList.add('active');
}

function closeModal() {
  document.getElementById('detailModal').classList.remove('active');
}

// Close modal on Escape
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') closeModal();
});

// ─── TOAST NOTIFICATIONS ────────────────────────────────────────────────────
function showToast(message, type = 'info') {
  const container = document.getElementById('toastContainer');
  const iconMap = { success: '✅', error: '❌', info: 'ℹ️' };

  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  toast.innerHTML = `
    <span class="toast-icon">${iconMap[type]}</span>
    <span class="toast-message">${message}</span>
    <button class="toast-close" onclick="this.parentElement.remove()">&times;</button>
  `;

  container.appendChild(toast);

  // Auto remove after 5s
  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateX(40px)';
    toast.style.transition = 'all 0.3s ease';
    setTimeout(() => toast.remove(), 300);
  }, 5000);
}
