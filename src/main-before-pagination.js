
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
      redirectTo: 'https://peaceful-khapse-85b5d7.netlify.app/',
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
const registers = [
  { key: 'DASHBOARD', label: 'Dashboard', aliases: [] },
  { key: 'WIR', label: 'Work Inspection Request', aliases: ['workinspectionrequest', 'wir'] },
  { key: 'MIR', label: 'Material Inspection Request', aliases: ['materialinspectionrequest', 'mir'] },
  { key: 'MAT', label: 'Material Submittal', aliases: ['materialsubmittal', 'mat'] },
  { key: 'MTS', label: 'Method Statement', aliases: ['methodstatement', 'mts'] },
  { key: 'SDW', label: 'Shop Drawings', aliases: ['shopdrawings', 'sdw', 'shopdrawing'] },
  { key: 'ABD', label: 'As Built Drawings', aliases: ['asbuiltdrawings', 'abd', 'asbuiltdrawing'] },
  { key: 'PQ', label: 'Pre-Qualification', aliases: ['prequalification', 'pq', 'pqr'] },
  { key: 'LTR_AMAALA', label: 'LTR AMAALA', aliases: ['ltramaala'] },
  { key: 'LTR_STE', label: 'LTR STE', aliases: ['ltrste'] },
  { key: 'RFI', label: 'RFI', aliases: ['rfi', 'requestforinformation'] },
  { key: 'ITP', label: 'Inspection & Test Plan', aliases: ['inspectiontestplan', 'itp'] },
  { key: 'ENV', label: 'Environmental Report', aliases: ['environmentalreport', 'env'] },
  { key: 'SUS', label: 'Sustainability Report', aliases: ['sustainabilityreport', 'sus'] }
];

function registerForSheetName(name) {
  const n = normalize(name);
  return registers.find(r => r.key !== 'DASHBOARD' &&
    (r.aliases.includes(n) || normalize(r.label) === n)) || null;
}

function registerForDocument(doc) {
  const type = normalize(doc.type);
  const no = normalize(doc.document_no);
  return registers.find(r => r.key !== 'DASHBOARD' &&
    ((type && r.aliases.some(alias => type.includes(alias) || alias.includes(type))) ||
      (r.key === 'RFI' && (type === 'rfi' || no.includes('rfi'))) ||
      (r.key === 'WIR' && no.includes('wir')) ||
      (r.key === 'MIR' && no.includes('mir')) ||
      (r.key === 'MAT' && no.includes('mat')) ||
      (r.key === 'MTS' && no.includes('mts')) ||
      (r.key === 'SDW' && no.includes('sdw')) ||
      (r.key === 'ABD' && no.includes('abd')))) || null;
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
  const counts = statusCounts(documents);
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
          <section class="status-dashboard">
            <div class="status-card total"><span>Total documents</span><strong>${counts.total.toLocaleString()}</strong></div>
            <div class="status-card accepted"><span>Accepted / Approved</span><strong>${counts.accepted.toLocaleString()}</strong></div>
            <div class="status-card comments"><span>Accepted with Comments</span><strong>${counts.comments.toLocaleString()}</strong></div>
            <div class="status-card revise"><span>Revise & Resubmit</span><strong>${counts.revise.toLocaleString()}</strong></div>
            <div class="status-card rejected"><span>Rejected</span><strong>${counts.rejected.toLocaleString()}</strong></div>
            <div class="status-card terminated"><span>Terminated</span><strong>${counts.terminated.toLocaleString()}</strong></div>
            <div class="status-card pending"><span>Pending / Under Review</span><strong>${counts.pending.toLocaleString()}</strong></div>
          </section>
          <p class="help">Select a register tab to view its documents. Dashboard totals are calculated from records currently loaded for this project.</p>
        ` : `
          <div class="summary">
            <span>${safe(currentRegister.label)} · ${shown.length.toLocaleString()} documents displayed</span>
            <span>${safe(project?.project_code ?? '')}</span>
          </div>
          <label for="searchInput">Search documents</label>
          <input id="searchInput" type="search"
            placeholder="Search number, title, revision, status..."
            value="${safe(searchText)}">
          <p id="notice" class="message"></p>
          <div class="table-wrap">
            <table>
              <thead><tr><th>S.No</th>${fields.map(([, label]) => `<th>${label}</th>`).join('')}</tr></thead>
              <tbody>
                ${shown.length ? shown.map((doc, i) => `<tr><td>${i + 1}</td>${fields.map(([key]) => `<td>${safe(doc[key] ?? '')}</td>`).join('')}</tr>`).join('') : `
                  <tr><td colspan="${fields.length + 1}">No documents found in this register. Import the workbook or check the document Type column.</td></tr>
                `}
              </tbody>
            </table>
          </div>
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
    activeRegister = 'DASHBOARD';
    searchText = '';
    await loadDocuments();

  };

  document.querySelectorAll('.register-tab').forEach(button => {
    button.onclick = () => {
      activeRegister = button.dataset.register;

      sessionStorage.setItem('activeRegister', activeRegister);

      searchText = '';
      renderTable();
    };
  });
}
const searchInput = document.querySelector('#searchInput');

if (searchInput) {
  searchInput.oninput = e => {
    searchText = e.target.value;
    renderTable();

    const input = document.querySelector('#searchInput');
    if (input) {
      input.focus();
      input.setSelectionRange(
        searchText.length,
        searchText.length
      );
    }
  };
}

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

    for (const sheetName of workbook.SheetNames) {
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

      // Dashboard and summary sheets are expected not to have document IDs.
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
      const sheetRegister = registerForSheetName(sheetName);

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
          } else if (field === 'type' && sheetRegister) {
            record[field] = sheetRegister.label;
          } else {
            record[field] = value || null;
          }
        }

        records.push(record);
      }
    }

    if (!records.length) {
      throw new Error(
        'No document records were found. Check that your workbook has a Document No or Mail No heading.'
      );
    }

    // Send manageable batches and keep the database's existing unique key.
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
      ? ` Summary/non-register sheets skipped: ${skippedSheets.join(', ')}.`
      : '';
    notice(`Successfully imported/updated ${imported.toLocaleString()} rows across workbook registers.${skippedMessage}`);
    activeRegister = 'DASHBOARD';
    await loadDocuments();
  } catch (error) {
    notice(`Import failed: ${error.message}`, true);
  }
}

function exportExcel() {
  const project = projects.find(
    p => String(p.id) === String(selectedProjectId)
  );

  if (!project) {
    notice('Select a project before exporting.', true);
    return;
  }

  const exportDocuments = activeRegister === 'DASHBOARD'
    ? documents
    : documents.filter(doc => registerForDocument(doc)?.key === activeRegister);

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

  const filename = project.project_name.replace(/[\\/:*?"<>|]/g, '_');
  XLSX.writeFile(workbook, `${filename}_Master_Log.xlsx`);
  notice('Excel file exported.');
}


async function loadDocuments() {

  app.innerHTML = `
  <main class="page">
    <section class="custom-loading-screen">
      <div class="loading-top">
        <span>PROJECT MASTER LOG</span>
        <p>Loading all project documents...</p>
      </div>

      <div class="loading-center">
        <div class="loading-spinner"></div>
        <h1>Loading Projects</h1>
        <p>Please wait while we fetch all project documents...</p>
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


  const allDocuments = [];
  const pageSize = 500;
  let from = 0;

  try {
    while (true) {
      const { data, error } = await supabase
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
        .eq('project_id', selectedProjectId)
        .order('document_no', { ascending: true })
        .order('id', { ascending: true })
        .range(from, from + pageSize - 1);

      if (error) {
        throw error;
      }

      allDocuments.push(...(data ?? []));
      console.log("Loaded batch:", data?.length, "Starting at:", from);

      if (!data || data.length < pageSize) {
        break;
      }

      from += pageSize;
    }

    documents = allDocuments;
    console.log("Finished loading documents:", documents.length);
    activeRegister = sessionStorage.getItem('activeRegister') || 'DASHBOARD';
    renderTable();

  } catch (error) {
    documents = [];
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

