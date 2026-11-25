
import './style.css';
import { createClient } from '@supabase/supabase-js';
import * as XLSX from 'xlsx';

const app = document.querySelector('#app');
const url = import.meta.env.VITE_SUPABASE_URL;
const key = import.meta.env.VITE_SUPABASE_PUBLISHABLE_KEY;

if (!url || !key) {
  app.textContent = 'Supabase settings are missing. Check .env.local.';
  throw new Error('Missing Supabase configuration');
}

const supabase = createClient(url, key);


const fields = [
  ['document_no', 'Document No'],
  ['revision', 'Revision'],
  ['title', 'Title'],
  ['type', 'Type'],
  ['status', 'Status'],
  ['review_status', 'Review Status'],
  ['discipline', 'Discipline'],
  ['stage', 'Stage'],
  ['location_area', 'Location/Area'],
  ['date_modified', 'Date Modified'],
  ['volume_system', 'Volume/System'],
  ['review_source', 'Review Source']
];

let projects = [];
let documents = [];
let selectedProjectId = '';
let searchText = '';
let activeRegister = 'DASHBOARD';

// Document register pagination
const PAGE_SIZE = 100;
let currentPage = 1;
let totalDocuments = 0;
let totalPages = 0;
let dashboardCountsProjectId = '';
let dashboardRows = [];
let dashboardCounts = {
  total: 0,
  accepted: 0,
  comments: 0,
  revise: 0,
  rejected: 0,
  terminated: 0,
  pending: 0
};
function safe(value) {
  return String(value ?? '').replace(/[&<>"']/g, ch => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  })[ch]);
}

function notice(text, error = false) {
  const el = document.querySelector('#notice');
  if (el) {
    el.textContent = text;
    el.className = error ? 'message error' : 'message';
  }
}

function showLogin() {
  app.innerHTML = `
    <main class="login-card">
      <div class="login-brand">
        <h1>Project Master Log</h1>
        <p>Document Control System</p>
      </div>
      <h2>Sign in</h2>
      <p>Use your existing Supabase user.</p>
      <form id="loginForm">
        <label for="email">Email address</label>
        <input id="email" type="email" required autocomplete="username">
        <label for="password">Password</label>
        <input id="password" type="password" required
          autocomplete="current-password">
        <p><a href="#" id="forgotPassword">Forgot password?</a></p>
        <button id="loginButton" type="submit">Sign in</button>
        <p id="loginNotice" class="message"></p>
      </form>
    </main>
  `;

  document.querySelector('#loginForm').addEventListener('submit', async e => {
    e.preventDefault();
    const button = document.querySelector('#loginButton');
    const feedback = document.querySelector('#loginNotice');

    button.disabled = true;
    feedback.textContent = 'Signing in...';

    const { error } = await supabase.auth.signInWithPassword({
      email: document.querySelector('#email').value.trim(),
      password: document.querySelector('#password').value
    });

    if (error) {
      feedback.textContent = error.message;
      button.disabled = false;
    }
  });

  document.querySelector('#forgotPassword').addEventListener('click', async e => {
    e.preventDefault();

    const email = document.querySelector('#email').value.trim();
    const feedback = document.querySelector('#loginNotice');

    if (!email) {
      feedback.textContent = 'Enter your email address first.';
      document.querySelector('#email').focus();
      return;
    }

    feedback.textContent = 'Sending password reset email...';

    const { error } = await supabase.auth.resetPasswordForEmail(email, {
      redirectTo: 'https://project-master-log-system.netlify.app/',
    });

    feedback.textContent = error
      ? error.message
      : 'Password reset email sent. Check your inbox.';
  });
}

function showUpdatePassword() {
  app.innerHTML = `
    <main class="login-card">
      <h2>Reset your password</h2>
      <p>Enter a new password for your account.</p>
      <form id="updatePasswordForm">
        <label for="newPassword">New password</label>
        <input id="newPassword" type="password" required minlength="8"
          autocomplete="new-password">
        <label for="confirmPassword">Confirm new password</label>
        <input id="confirmPassword" type="password" required minlength="8"
          autocomplete="new-password">
        <button id="updatePasswordButton" type="submit">
          Update password
        </button>
        <p id="passwordNotice" class="message"></p>
      </form>
    </main>
  `;

  document.querySelector('#updatePasswordForm').addEventListener('submit', async e => {
    e.preventDefault();

    const password = document.querySelector('#newPassword').value;
    const confirmPassword = document.querySelector('#confirmPassword').value;
    const feedback = document.querySelector('#passwordNotice');
    const button = document.querySelector('#updatePasswordButton');

    if (password.length < 8) {
      feedback.textContent = 'Use at least 8 characters.';
      return;
    }

    if (password !== confirmPassword) {
      feedback.textContent = 'The passwords do not match.';
      return;
    }

    button.disabled = true;
    feedback.textContent = 'Updating password...';

    const { error } = await supabase.auth.updateUser({ password });

    if (error) {
      feedback.textContent = error.message;
      button.disabled = false;
      return;
    }

    feedback.textContent = 'Password updated successfully.';
    setTimeout(() => {
      window.history.replaceState(null, '', window.location.pathname);
      start();
    }, 1000);
  });
}

/* Dashboard: display all six projects as clickable cards. */
function showProjectsDashboard() {
  searchText = '';

  app.innerHTML = `
    <header class="topbar dashboard-topbar">
      <div>
        <h1>Project Master Log</h1>
        <p>Document Control System</p>
      </div>
      <button id="logoutButton" class="secondary">Sign out</button>
    </header>

    <main class="page project-home">
      <section class="welcome-panel">
        <p class="eyebrow">PROJECT PORTFOLIO</p>
        <h2>Select a project</h2>
        <p>
          Open a project to manage its document register, revisions,
          review status, and Excel records.
        </p>
      </section>

      <section class="project-grid" aria-label="Projects">
        ${projects.map((project, index) => `
          <button type="button"
            class="project-card"
            data-project-id="${safe(project.id)}">
            <span class="project-card-number">
              ${String(index + 1).padStart(2, '0')}
            </span>
            <span class="project-card-icon" aria-hidden="true">
              ${['🏨', '🌊', '🏢', '🏙️', '🏥', '⛳'][index % 6]}
            </span>
            <span class="project-card-name">
              ${safe(project.project_name)}
            </span>
            <span class="project-card-code">
              Project Code: ${safe(project.project_code)}
            </span>
            <span class="project-card-open">
              Open project <span aria-hidden="true">→</span>
            </span>
          </button>
        `).join('')}
      </section>

      <p class="dashboard-footnote">
        Choose a project card to open its document register.
      </p>
    </main>
  `;

  document.querySelector('#logoutButton').onclick = async () => {
    const { error } = await supabase.auth.signOut();
    if (error) {
      app.insertAdjacentHTML(
        'afterbegin',
        `<p class="message error">${safe(error.message)}</p>`
      );
    }
  };


  document.querySelectorAll('.project-card').forEach(card => {
    card.addEventListener('click', async () => {
      selectedProjectId = card.dataset.projectId;

      sessionStorage.setItem('selectedProjectId', selectedProjectId);
      sessionStorage.setItem('activeRegister', 'DASHBOARD');

      activeRegister = 'DASHBOARD';
      searchText = '';

      await loadDocuments();
    });
  });

}

/* Excel workbook tabs represented in the website. */
let registers = [
  { key: 'DASHBOARD', label: 'Dashboard' }
];

function registerKey(name) {
  return normalize(name);
}

// Compare register names in singular form so dashboard labels such as
// 'Shop Drawings' match workbook tabs stored as 'Shop Drawing'.
function registerMatchKey(name) {
  const key = registerKey(name);
  if (key.endsWith('ies') && key.length > 3) return key.slice(0, -3) + 'y';
  if (key.endsWith('s') && key.length > 1) return key.slice(0, -1);
  return key;
}

function registerForDocument(doc) {
  const type = registerKey(doc.type);
  if (!type) return null;

  // Match against the visible worksheet name, not a possibly old or
  // inconsistent register_key stored in the database.
  return registers.find(register =>
    register.key !== 'DASHBOARD' &&
    registerMatchKey(register.label) === registerMatchKey(doc.type)
  ) || null;
}

/*
 * Load this project's register tabs from the dedicated register list.
 * This includes worksheets with zero document rows.
 */
async function loadProjectRegisters() {
  const { data, error } = await supabase
    .from('master_log_registers')
    .select('register_name, register_key')
    .eq('project_id', selectedProjectId)
    .order('register_name', { ascending: true });

  if (error) throw error;

  const seen = new Set(['DASHBOARD']);
  const dynamic = (data ?? [])
    .map(row => {
      const label = String(row.register_name ?? '').trim();
      // Build a consistent runtime key from the worksheet name. This keeps
      // navigation working even if an older stored register_key differs.
      return { key: registerMatchKey(label), label };
    })
    .filter(register => {
      if (!register.label || !register.key ||
        register.key === 'DASHBOARD' ||
        register.key === 'NOTSPECIFIED' ||
        seen.has(register.key)) return false;
      seen.add(register.key);
      return true;
    });

  registers = [
    { key: 'DASHBOARD', label: 'Dashboard' },
    ...dynamic
  ];

  if (activeRegister !== 'DASHBOARD' &&
    !registers.some(register => register.key === activeRegister)) {
    activeRegister = 'DASHBOARD';
    sessionStorage.setItem('activeRegister', activeRegister);
  }
}
function statusCounts(list) {
  const counts = {
    total: list.length,
    accepted: 0,
    comments: 0,
    revise: 0,
    rejected: 0,
    terminated: 0,
    pending: 0
  };

  for (const doc of list) {
    const status = normalize(
      `${doc.status ?? ''} ${doc.review_status ?? ''}`
    );

    if (
      status.includes('acceptedwithcomments') ||
      status.includes('acceptedwithcomment')
    ) {
      counts.comments++;
    } else if (
      status.includes('accepted') ||
      status.includes('approved')
    ) {
      counts.accepted++;
    } else if (
      status.includes('revise') ||
      status.includes('resubmit')
    ) {
      counts.revise++;
    } else if (status.includes('reject')) {
      counts.rejected++;
    } else if (status.includes('terminat')) {
      counts.terminated++;
    } else if (
      status.includes('pending') ||
      status.includes('inreview') ||
      status.includes('underreview') ||
      status.includes('awaiting') ||
      status.includes('submitted')
    ) {
      counts.pending++;
    }
  }

  return counts;
}

async function loadDashboardCounts() {
  if (String(dashboardCountsProjectId) === String(selectedProjectId)) {
    return;
  }
  const allStatuses = [];
  const batchSize = 100;
  let from = 0;

  while (true) {
    const { data, error } = await supabase
      .from('master_log_documents')
      .select('status, review_status, type, discipline')
      .eq('project_id', selectedProjectId)
      .order('id', { ascending: true })
      .range(from, from + batchSize - 1);

    if (error) {
      throw error;
    }

    allStatuses.push(...(data ?? []));

    if (!data || data.length < batchSize) {
      break;
    }

    from += batchSize;
  }

  dashboardRows = allStatuses;
  dashboardCounts = statusCounts(allStatuses);
  dashboardCountsProjectId = String(selectedProjectId);
  console.log('Full dashboard counts:', dashboardCounts);
}

function dashboardGroupSummary(field) {
  const map = new Map();
  for (const row of dashboardRows) {
    const label = String(row[field] ?? '').trim() || 'Not specified';
    map.set(label, (map.get(label) || 0) + 1);
  }
  return [...map.entries()]
    .map(([label, count]) => ({ label, count, pct: dashboardCounts.total ? count / dashboardCounts.total * 100 : 0 }))
    .sort((a, b) => b.count - a.count || a.label.localeCompare(b.label));
}

function dashboardStatusRows() {
  const items = [
    ['A', 'Accepted', dashboardCounts.accepted],
    ['B', 'Accepted with Comments', dashboardCounts.comments],
    ['C', 'Revise & Resubmit', dashboardCounts.revise],
    ['D', 'Rejected', dashboardCounts.rejected],
    ['E', 'Terminated', dashboardCounts.terminated],
    ['F', 'Pending / Under Review', dashboardCounts.pending]
  ];
  return items.map(([code, label, count]) => ({
    code, label, count,
    pct: dashboardCounts.total ? count / dashboardCounts.total * 100 : 0
  }));
}

function dashboardTable(title, heading, rows, emptyText, clickable = false) {
  return `
    <section class="dash-summary-card">
      <h3>${title}</h3>
      <div class="dash-table-scroll">
        <table class="dash-summary-table">
          <thead><tr><th>${heading}</th><th>Documents</th><th>% of Total</th></tr></thead>
          <tbody>${rows.length ? (clickable ? rows : rows.slice(0, 12)).map(row => {
    const rowKey = registerKey(row.label);
    const targetRegister = clickable && rowKey !== 'NOTSPECIFIED'
      ? registers.find(register =>
        register.key !== 'DASHBOARD' &&
        (register.key === rowKey ||
          registerMatchKey(register.label) === registerMatchKey(row.label))
      )
      : null;
    return `
              <tr ${targetRegister ? `class="dash-register-row" data-register="${safe(targetRegister.key)}" tabindex="0" role="button" title="Open ${safe(targetRegister.label)} register"` : ''}>
                <td>${safe(row.label)}</td><td>${row.count.toLocaleString()}</td><td>${row.pct.toFixed(1)}%</td>
              </tr>`;
  }).join('') : `<tr><td colspan="3">${emptyText}</td></tr>`}</tbody>
        </table>
      </div>
    </section>`;
}

function renderMasterDashboard() {
  const statusRows = dashboardStatusRows();
  const registerRows = dashboardGroupSummary('type');
  const disciplineRows = dashboardGroupSummary('discipline');
  const colors = ['#4f8b36', '#83b93e', '#ef7b2d', '#cf1717', '#858585', '#3973c9'];
  const total = Math.max(1, dashboardCounts.total);
  let cursor = 0;
  const segments = statusRows.map((row, i) => {
    const start = cursor;
    cursor += row.count / total * 100;
    return `${colors[i]} ${start}% ${cursor}%`;
  });
  const donutStyle = `background:conic-gradient(${segments.join(',')})`;
  const bars = registerRows.slice(0, 10).map(row => `
    <div class="dash-bar-row"><span title="${safe(row.label)}">${safe(row.label)}</span>
      <div class="dash-bar-track"><i style="width:${Math.max(row.pct, row.count ? 1 : 0)}%"></i></div>
      <b>${row.count.toLocaleString()}</b></div>`).join('');
  const disciplineBars = disciplineRows.slice(0, 10).map(row => `
    <div class="dash-bar-row"><span title="${safe(row.label)}">${safe(row.label)}</span>
      <div class="dash-bar-track"><i style="width:${Math.max(row.pct, row.count ? 1 : 0)}%"></i></div>
      <b>${row.count.toLocaleString()}</b></div>`).join('');
  return `
    <section class="master-dashboard">
      <div class="master-dash-heading">
        <h2>${safe(projects.find(p => String(p.id) === String(selectedProjectId))?.project_name || 'PROJECT')} | MASTER DOCUMENT CONTROL DASHBOARD</h2>
        <p>Project Code: ${safe(projects.find(p => String(p.id) === String(selectedProjectId))?.project_code || '')} | Consolidated from available register worksheets</p>
      </div>
      <div class="dash-kpi-grid">
        <div class="dash-kpi total"><span>Total Documents</span><strong>${dashboardCounts.total.toLocaleString()}</strong></div>
        <div class="dash-kpi accepted"><span>Code A | Accepted</span><strong>${dashboardCounts.accepted.toLocaleString()}</strong></div>
        <div class="dash-kpi comments"><span>Code B | Accepted with Comments</span><strong>${dashboardCounts.comments.toLocaleString()}</strong></div>
        <div class="dash-kpi revise"><span>Code C | Revise & Resubmit</span><strong>${dashboardCounts.revise.toLocaleString()}</strong></div>
        <div class="dash-kpi rejected"><span>Code D | Rejected</span><strong>${dashboardCounts.rejected.toLocaleString()}</strong></div>
        <div class="dash-kpi terminated"><span>Code E | Terminated</span><strong>${dashboardCounts.terminated.toLocaleString()}</strong></div>
        <div class="dash-kpi pending"><span>Code F | Pending / Under Review</span><strong>${dashboardCounts.pending.toLocaleString()}</strong></div>
      </div>
      <div class="dash-summary-grid">
        ${dashboardTable('DOCUMENT STATUS SUMMARY', 'Code / Review Status', statusRows, 'No status records found')}
        ${dashboardTable('REGISTER-WISE DOCUMENT SUMMARY', 'Register', registerRows, 'No register data found', true)}
        ${dashboardTable('DISCIPLINE-WISE SUMMARY', 'Discipline', disciplineRows, 'No discipline data found')}
      </div>
      <div class="dash-chart-grid">
        <section class="dash-chart-card"><h3>DOCUMENT REVIEW STATUS</h3>
          <div class="dash-donut-layout"><div class="dash-donut" style="${donutStyle}"><div><b>${dashboardCounts.total.toLocaleString()}</b><small>Documents</small></div></div>
            <ul class="dash-legend">${statusRows.map((row, i) => `<li><i style="background:${colors[i]}"></i><span>${row.code} · ${safe(row.label)}</span><b>${row.count.toLocaleString()}</b></li>`).join('')}</ul>
          </div>
        </section>
        <section class="dash-chart-card"><h3>DOCUMENTS BY REGISTER</h3><div class="dash-bars">${bars || '<p>No register data available.</p>'}</div></section>
        <section class="dash-chart-card"><h3>DOCUMENTS BY DISCIPLINE</h3><div class="dash-bars">${disciplineBars || '<p>No discipline data available.</p>'}</div></section>
      </div>
      <p class="dash-note">Dashboard summaries are calculated from all document records currently stored for this project. Register and discipline names are taken from imported workbook data.</p>
    </section>`;
}

/* Open the selected project's document register. */
function renderTable() {
  const project = projects.find(
    p => String(p.id) === String(selectedProjectId)
  );

  const registerDocs = activeRegister === 'DASHBOARD'
    ? documents
    : documents.filter(doc => registerForDocument(doc)?.key === activeRegister);

  const shown = registerDocs.filter(doc =>
    fields.some(([key]) =>
      String(doc[key] ?? '').toLowerCase().includes(searchText.toLowerCase())
    )
  );
  const counts = dashboardCounts;
  const currentRegister = registers.find(r => r.key === activeRegister) || registers[0];

  app.innerHTML = `
    <header class="topbar register-topbar">
      <div>
        <h1>Project Master Log</h1>
        <p>
          ${safe(project?.project_name ?? 'Project Document Register')}
          · ${safe(project?.project_code ?? '')}
        </p>
      </div>
      <div class="topbar-actions">
        <button id="projectsButton" class="secondary">← All projects</button>
        <button id="logoutButton" class="secondary">Sign out</button>
      </div>
    </header>

    <main class="page">
      <section class="panel">
        <div class="toolbar">
          <div class="project-picker">
            <label for="projectSelect">Select Project</label>
            <select id="projectSelect">
              ${projects.map(p => `
                <option value="${safe(p.id)}"
                  ${String(p.id) === String(selectedProjectId) ? 'selected' : ''}>
                  ${safe(p.project_name)} (${safe(p.project_code)})
                </option>
              `).join('')}
            </select>
          </div>

          <div class="actions">
            <button id="importButton">Import Excel</button>
            <button id="exportButton" class="secondary">Export Excel</button>
            <input id="excelFile" type="file"
              accept=".xlsx,.xls,.csv" hidden>
          </div>
        </div>

        <nav class="register-tabs" aria-label="Document registers">
          ${registers.map(r => `
            <button type="button" class="register-tab ${r.key === activeRegister ? 'active' : ''}"
              data-register="${safe(r.key)}">${safe(r.key === 'DASHBOARD' ? 'Dashboard' : r.label)}</button>
          `).join('')}
        </nav>

        ${activeRegister === 'DASHBOARD' ? `
  ${renderMasterDashboard()}
  <p class="help">Select a register tab to view its documents.</p>
` : `
  <div class="summary">
    <span>
      ${safe(currentRegister.label)} ·
      ${shown.length.toLocaleString()} displayed ·
      ${totalDocuments.toLocaleString()} total
    </span>
    <span>${safe(project?.project_code ?? '')}</span>
  </div>

  <label for="searchInput">Search documents</label>
  <input id="searchInput" type="search"
    placeholder="Search number, title, revision, status..."
    value="${safe(searchText)}">

  <p id="notice" class="message"></p>

  <div class="table-wrap">
    <table>
      <thead>
        <tr>
          <th>S.No</th>
          ${fields.map(([, label]) => `<th>${label}</th>`).join('')}
        </tr>
      </thead>
      <tbody>
        ${shown.length ? shown.map((doc, i) => `
          <tr>
            <td>${(currentPage - 1) * PAGE_SIZE + i + 1}</td>
            ${fields.map(([key]) => `<td>${safe(doc[key] ?? '')}</td>`).join('')}
          </tr>
        `).join('') : `
          <tr>
            <td colspan="${fields.length + 1}">
              No documents found in this register.
            </td>
          </tr>
        `}
      </tbody>
    </table>
  </div>

  <div class="pagination">
    <button id="previousPage" ${currentPage <= 1 ? 'disabled' : ''}>
      Previous
    </button>

    <span>
      Page ${currentPage.toLocaleString()} of ${totalPages.toLocaleString()}
    </span>

    <button id="nextPage" ${currentPage >= totalPages ? 'disabled' : ''}>
      Next
    </button>
  </div>

  <p class="help">
    Import one project's Excel file at a time. Back up your original
    Excel file before importing.
  </p>
`}
        ${activeRegister === 'DASHBOARD' ? '' : ''}

        <p class="help">
          Import one project's Excel file at a time. Back up your original
          Excel file before importing.
        </p>
      </section>
    </main>
  `;

  document.querySelector('#projectsButton').onclick =
    () => showProjectsDashboard();

  document.querySelector('#logoutButton').onclick = async () => {
    const { error } = await supabase.auth.signOut();
    if (error) notice(error.message, true);
  };


  document.querySelector('#projectSelect').onchange = async e => {
    selectedProjectId = e.target.value;
    dashboardCountsProjectId = '';
    dashboardRows = [];
    currentPage = 1;
    totalDocuments = 0;
    totalPages = 0;
    activeRegister = 'DASHBOARD';
    searchText = '';

    await loadDocuments();
  };

  // Register tab buttons
  document.querySelectorAll('.register-tab').forEach(button => {
    button.onclick = async () => {
      activeRegister = button.dataset.register;
      sessionStorage.setItem('activeRegister', activeRegister);

      searchText = '';
      currentPage = 1;

      await loadDocuments();
    };
  });

  // Clicking a register-wise dashboard row opens that register directly.
  document.querySelectorAll('.dash-register-row').forEach(row => {
    const openRegister = async () => {
      activeRegister = row.dataset.register;
      sessionStorage.setItem('activeRegister', activeRegister);
      searchText = '';
      currentPage = 1;
      await loadDocuments();
    };
    row.addEventListener('click', openRegister);
    row.addEventListener('keydown', event => {
      if (event.key === 'Enter' || event.key === ' ') {
        event.preventDefault();
        openRegister();
      }
    });
  });

  // Previous page button
  document.querySelector('#previousPage')?.addEventListener('click', async () => {
    if (currentPage > 1) {
      currentPage--;
      await loadDocuments();
    }
  });

  // Next page button
  document.querySelector('#nextPage')?.addEventListener('click', async () => {
    if (currentPage < totalPages) {
      currentPage++;
      await loadDocuments();
    }
  });

  // Search documents
  const searchInput = document.querySelector('#searchInput');

  if (searchInput) {
    searchInput.onkeydown = async e => {
      if (e.key === 'Enter') {
        e.preventDefault();
        searchText = e.currentTarget.value.trim();
        currentPage = 1;
        await loadDocuments();
      }
    };
  }

  // Excel Import button
  const importButton = document.querySelector('#importButton');
  const excelFile = document.querySelector('#excelFile');

  if (importButton && excelFile) {
    importButton.onclick = () => {
      excelFile.click();
    };

    excelFile.onchange = importExcel;
  }

  // Excel Export button
  const exportButton = document.querySelector('#exportButton');

  if (exportButton) {
    exportButton.onclick = () => {
      exportExcel();
    };
  }

}
// End of renderTable()

document.querySelector('#statusFilter')?.addEventListener('change', () => {
  renderTable();
});

document.querySelector('#disciplineFilter')?.addEventListener('change', () => {
  renderTable();
});

document.querySelector('#clearFiltersButton')?.addEventListener('click', () => {
  searchText = '';
  renderTable();
});


function normalize(value) {
  return String(value ?? '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function formatDate(value) {
  if (value === null || value === undefined || value === '') {
    return null;
  }

  // Excel may already provide a JavaScript Date object.
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10);
  }

  const text = String(value).trim();

  // Convert DD/MM/YYYY to YYYY-MM-DD.
  const match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);

  if (match) {
    const [, day, month, year] = match;
    return `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`;
  }

  return text;
}


const aliases = {
  document_no: ['documentno', 'documentnumber', 'docno', 'documentid'],
  revision: ['revision', 'rev', 'revisionno', 'revno'],
  title: ['title', 'documenttitle', 'documentname', 'description'],
  type: ['type', 'documenttype', 'doctype'],
  status: ['status', 'documentstatus'],
  review_status: ['reviewstatus', 'reviewresult', 'reviewoutcome'],
  discipline: ['discipline', 'documentdiscipline'],
  stage: ['stage', 'projectstage'],
  location_area: ['locationarea', 'location', 'area', 'locationorarea'],
  date_modified: ['datemodified', 'modifieddate', 'lastmodified', 'date'],
  volume_system: ['volumesystem', 'volume', 'system'],
  review_source: ['reviewsource', 'source', 'reviewer', 'reviewedby']
};

function importColumn(headers, field) {
  return headers.find(header =>
    aliases[field]?.includes(normalize(header))
  );
}

async function importExcel(event) {
  const file = event.target.files?.[0];
  event.target.value = '';
  if (!file || !selectedProjectId) return;

  if (!window.confirm(
    `Import document registers from ${file.name} into the selected project? ` +
    'Matching document and revision records may be updated.'
  )) return;

  try {
    notice('Reading Excel workbook...');

    const workbook = XLSX.read(await file.arrayBuffer(), {
      type: 'array',
      cellDates: true
    });

    const records = [];
    const skippedSheets = [];
    const detectedRegisters = [];

    for (const sheetName of workbook.SheetNames) {
      const cleanSheetName = String(sheetName ?? '').trim();
      if (!cleanSheetName) continue;

      // Register names are captured before checking document headers.
      // This preserves empty worksheets as visible register tabs.
      detectedRegisters.push({
        project_id: selectedProjectId,
        register_name: cleanSheetName,
        register_key: registerKey(cleanSheetName)
      });

      const sheet = workbook.Sheets[sheetName];
      if (!sheet) continue;

      const matrix = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: '',
        raw: false
      });

      const headerIndex = matrix.findIndex(row =>
        row.some(cell => {
          const h = normalize(cell);
          return ['documentno', 'documentnumber', 'docno', 'mailno'].includes(h);
        })
      );

      // Keep the worksheet as a register even if it has no document-number
      // heading. It will simply have no document rows imported.
      if (headerIndex === -1) {
        skippedSheets.push(sheetName);
        continue;
      }

      const headers = matrix[headerIndex].map(cell => String(cell ?? '').trim());
      const headerMap = {};
      headers.forEach(h => { headerMap[normalize(h)] = h; });

      const findHeader = names => {
        for (const name of names) {
          if (headerMap[name]) return headerMap[name];
        }
        return undefined;
      };

      const documentHeader = findHeader(['documentno', 'documentnumber', 'docno', 'mailno']);
      const titleHeader = findHeader(['title', 'documenttitle', 'documentname', 'description', 'subject']);
      const dateHeader = findHeader(['datemodified', 'modifieddate', 'lastmodified', 'date']);
      const statusHeader = findHeader(['status', 'documentstatus', 'reviewstatus', 'reviewresult', 'reviewoutcome']);
      const revisionHeader = findHeader(['revision', 'rev', 'revisionno', 'revno']);
      // The worksheet name is the register name for this imported sheet.
      const sheetRegisterName = String(sheetName).trim();

      const rows = matrix.slice(headerIndex + 1)
        .filter(row => row.some(cell => String(cell ?? '').trim() !== ''));

      for (const row of rows) {
        const get = header => header === undefined ? '' : row[headers.indexOf(header)] ?? '';
        const documentNo = String(get(documentHeader) ?? '').trim();
        if (!documentNo) continue;

        const record = { project_id: selectedProjectId };
        for (const [field] of fields) {
          const header = importColumn(headers, field);
          let value = header ? get(header) : '';

          // RFI registers commonly use Mail No and Subject rather than
          // Document No and Title.
          if (field === 'document_no') value = documentNo;
          if (field === 'title' && !value && titleHeader) value = get(titleHeader);
          if (field === 'revision' && !value && revisionHeader) value = get(revisionHeader);
          if (field === 'status' && !value && statusHeader) value = get(statusHeader);
          if (field === 'date_modified' && !value && dateHeader) value = get(dateHeader);

          value = value === null || value === undefined ? '' : String(value).trim();

          if (field === 'document_no') {
            record[field] = value;
          } else if (field === 'date_modified') {
            record[field] = formatDate(value);
          } else if (field === 'type') {
            // Keep the exact Excel worksheet name so it appears as a tab.
            record[field] = sheetRegisterName || value || null;
          } else {
            record[field] = value || null;
          }
        }

        records.push(record);
      }
    }

    if (!detectedRegisters.length) {
      throw new Error('No worksheets were found in this workbook.');
    }

    // Save every worksheet name for this project, including empty registers.
    const registerBatchSize = 500;
    for (let i = 0; i < detectedRegisters.length; i += registerBatchSize) {
      const { error } = await supabase
        .from('master_log_registers')
        .upsert(detectedRegisters.slice(i, i + registerBatchSize), {
          onConflict: 'project_id,register_key'
        });

      if (error) throw error;
    }

    // Import document rows only when worksheets contain usable document IDs.
    // Existing document records are retained; matching keys are updated.
    let imported = 0;
    const batchSize = 500;
    for (let i = 0; i < records.length; i += batchSize) {
      const batch = records.slice(i, i + batchSize);
      const { error } = await supabase
        .from('master_log_documents')
        .upsert(batch, {
          onConflict: 'project_id,document_no,revision'
        });

      if (error) throw error;
      imported += batch.length;
      notice(`Importing registers... ${imported.toLocaleString()} of ${records.length.toLocaleString()} rows`);
    }

    const skippedMessage = skippedSheets.length
      ? ` These worksheets were registered but had no Document No/Mail No heading, so no document rows were imported: ${skippedSheets.join(', ')}.`
      : '';
    notice(`Successfully detected ${detectedRegisters.length.toLocaleString()} worksheets and imported/updated ${imported.toLocaleString()} document rows.${skippedMessage}`);
    dashboardCountsProjectId = '';
    dashboardRows = [];
    activeRegister = 'DASHBOARD';
    await loadDocuments();
  } catch (error) {
    notice(`Import failed: ${error.message}`, true);
  }
}


async function exportExcel() {
  const project = projects.find(
    p => String(p.id) === String(selectedProjectId)
  );

  if (!project) {
    notice('Select a project before exporting.', true);
    return;
  }

  try {
    notice('Preparing Excel export...');

    const exportDocuments = [];
    const batchSize = 1000;
    let from = 0;

    while (true) {
      let query = supabase
        .from('master_log_documents')
        .select(`
          id,
          project_id,
          document_no,
          revision,
          title,
          type,
          status,
          review_status,
          discipline,
          stage,
          location_area,
          date_modified,
          volume_system,
          review_source
        `)
        .eq('project_id', selectedProjectId);

      if (activeRegister !== 'DASHBOARD') {
        const selectedRegister = registers.find(r => r.key === activeRegister);
        query = query.eq('type', selectedRegister?.label ?? '__NO_MATCHING_REGISTER__');
      }

      const { data, error } = await query
        .order('document_no', { ascending: true })
        .order('id', { ascending: true })
        .range(from, from + batchSize - 1);

      if (error) throw error;

      exportDocuments.push(...(data ?? []));

      notice(
        `Preparing export... ${exportDocuments.length.toLocaleString()} records`
      );

      if (!data || data.length < batchSize) break;

      from += batchSize;
    }

    const rows = exportDocuments.map((doc, i) => {
      const row = { 'S.No': i + 1 };

      for (const [field, label] of fields) {
        row[label] = doc[field] ?? '';
      }

      return row;
    });

    const workbook = XLSX.utils.book_new();
    const sheet = XLSX.utils.json_to_sheet(rows);

    XLSX.utils.book_append_sheet(workbook, sheet, 'Master Log');

    const filename = project.project_name.replace(/[\/:*?"<>|]/g, '_');

    XLSX.writeFile(
      workbook,
      `${filename}_${activeRegister}_Master_Log.xlsx`
    );

    notice(
      `Export successful: ${exportDocuments.length.toLocaleString()} records.`
    );

  } catch (error) {
    notice(`Export failed: ${error.message}`, true);
  }
}

async function loadDocuments() {
  app.innerHTML = `
    <main class="page">
      <section class="custom-loading-screen">
        <div class="loading-top">
          <span>PROJECT MASTER LOG</span>
          <p>Loading project documents...</p>
        </div>
        <div class="loading-center">
          <div class="loading-spinner"></div>
          <h1>Loading Documents</h1>
          <p>Please wait while we fetch your documents...</p>
          <div class="loading-progress">
            <div class="loading-progress-bar"></div>
          </div>
        </div>
        <div class="loading-footer">
          DOCUMENT MANAGEMENT SYSTEM
        </div>
      </section>
    </main>
  `;

  try {
    activeRegister =
      sessionStorage.getItem('activeRegister') || 'DASHBOARD';

    await loadProjectRegisters();

    const from = (currentPage - 1) * PAGE_SIZE;
    const to = from + PAGE_SIZE - 1;

    let query = supabase
      .from('master_log_documents')
      .select(`
        id,
        project_id,
        document_no,
        revision,
        title,
        type,
        status,
        review_status,
        discipline,
        stage,
        location_area,
        date_modified,
        volume_system,
        review_source
      `, { count: 'exact' })
      .eq('project_id', selectedProjectId);

    // Filter by the selected register before pagination.
    if (activeRegister !== 'DASHBOARD') {
      const selectedRegister = registers.find(r => r.key === activeRegister);
      query = query.eq('type', selectedRegister?.label ?? '__NO_MATCHING_REGISTER__');
    }
    // Search across all matching records in the selected register.
    const term = searchText.trim();
    console.log('Search text:', term);
    console.log('Selected project:', selectedProjectId);
    console.log('Selected register:', activeRegister);

    if (term) {
      query = query.or(
        `document_no.ilike.%${term}%,` +
        `title.ilike.%${term}%,` +
        `revision.ilike.%${term}%,` +
        `type.ilike.%${term}%,` +
        `status.ilike.%${term}%,` +
        `review_status.ilike.%${term}%,` +
        `discipline.ilike.%${term}%,` +
        `stage.ilike.%${term}%,` +
        `location_area.ilike.%${term}%,` +
        `volume_system.ilike.%${term}%,` +
        `review_source.ilike.%${term}%`
      );
    }
    const { data, error, count } = await query
      .order('document_no', { ascending: true })
      .order('id', { ascending: true })
      .range(from, to);

    if (error) {
      throw error;
    }

    documents = data ?? [];
    totalDocuments = count ?? 0;
    totalPages = Math.ceil(totalDocuments / PAGE_SIZE);

    await loadDashboardCounts();

    console.log('Active register:', activeRegister);
    console.log('Loaded page:', currentPage);
    console.log('Documents on page:', documents.length);
    console.log('Total matching documents:', totalDocuments);
    console.log('Total pages:', totalPages);

    renderTable();

  } catch (error) {
    documents = [];
    totalDocuments = 0;
    totalPages = 0;
    renderTable();
    notice(`Could not load documents: ${error.message}`, true);
  }
}

async function loadProjects() {
  const { data, error } = await supabase
    .from('master_log_projects')
    .select('id, project_code, project_name')
    .order('project_name');

  if (error) throw error;

  projects = data ?? [];

  if (!projects.length) {
    throw new Error('No projects found.');
  }

  // Restore the saved project, but do not load its documents here.
  selectedProjectId = sessionStorage.getItem('selectedProjectId') || '';

  if (
    selectedProjectId &&
    !projects.some(p => String(p.id) === String(selectedProjectId))
  ) {
    selectedProjectId = '';
    sessionStorage.removeItem('selectedProjectId');
  }
}




async function start() {
  const { data, error } = await supabase.auth.getSession();

  if (error) {
    app.textContent = error.message;
    return;
  }

  if (!data.session) {
    showLogin();
    return;
  }

  app.innerHTML = '<main class="page"><p>Loading projects...</p></main>';

  try {
    // Load projects once
    await loadProjects();

    // Restore the last selected register
    activeRegister =
      sessionStorage.getItem('activeRegister') || 'DASHBOARD';

    // Load documents only when a project is selected
    if (selectedProjectId) {
      await loadDocuments();
    } else {
      showProjectsDashboard();
    }

  } catch (error) {
    app.innerHTML = `
      <main class="page">
        <p class="message error">${safe(error.message)}</p>
      </main>
    `;
  }
}

let startPromise = null;

supabase.auth.onAuthStateChange((event) => {
  if (event === 'PASSWORD_RECOVERY') {
    showUpdatePassword();
  }

  if (event === 'SIGNED_IN' && !startPromise) {
    startPromise = start().finally(() => {
      startPromise = null;
    });
  }

  if (event === 'SIGNED_OUT') {
    showLogin();
  }
});

startPromise = start().finally(() => {
  startPromise = null;
});

