// ==========================================================================
// Frontend Logic & Auth Controller
// ==========================================================================

document.addEventListener('DOMContentLoaded', () => {
  // Authentication Elements
  const modalAuth = document.getElementById('modal-auth');
  const modalStatus = document.getElementById('modal-status');
  const formLogin = document.getElementById('form-login');
  const formRegister = document.getElementById('form-register');
  const linkShowRegister = document.getElementById('link-show-register');
  const linkShowLogin = document.getElementById('link-show-login');
  const loginErrorMsg = document.getElementById('login-error-msg');
  const regErrorMsg = document.getElementById('reg-error-msg');
  const btnLogout = document.getElementById('btn-logout');
  const btnStatusLogout = document.getElementById('btn-status-logout');
  const headerUsername = document.getElementById('header-username');
  const mainAppContainer = document.getElementById('main-app-container');

  // Navigation Tabs & Panels
  const mainTabs = document.getElementById('main-tabs');
  const tabBtnUsers = document.getElementById('tab-btn-users');
  const tabContentRegistry = document.getElementById('tab-content-registry');
  const tabContentDrivers = document.getElementById('tab-content-drivers');
  const tabContentUsers = document.getElementById('tab-content-users');
  
  // Users elements
  const usersTableBody = document.getElementById('users-table-body');
  const usersCount = document.getElementById('users-count');

  // Add Vehicle Drawer Elements
  const btnAddVehicle = document.getElementById('btn-add-vehicle');
  const modalAddVehicle = document.getElementById('modal-add-vehicle');
  const btnCloseAddModal = document.getElementById('btn-close-add-modal');
  const btnCancelAdd = document.getElementById('btn-cancel-add');
  const formAddVehicle = document.getElementById('form-add-vehicle');
  const vehicleTableBody = document.getElementById('vehicle-table-body');
  
  // Vehicle File inputs
  const dlInput = document.getElementById('dlPhoto');
  const dlPreview = document.getElementById('dl-file-preview');
  const pcInput = document.getElementById('punchCard');
  const pcPreview = document.getElementById('pc-file-preview');

  // Drivers Tab Elements
  const btnAddDriver = document.getElementById('btn-add-driver');
  const modalAddDriver = document.getElementById('modal-add-driver');
  const btnCloseDriverModal = document.getElementById('btn-close-driver-modal');
  const btnCancelAddDriver = document.getElementById('btn-cancel-add-driver');
  const formAddDriver = document.getElementById('form-add-driver');
  const driversTableBody = document.getElementById('drivers-table-body');
  const driversCount = document.getElementById('drivers-count');
  const driverSearchInput = document.getElementById('driver-search-input');
  const driverStatusFilter = document.getElementById('driver-status-filter');

  // Driver File Inputs & Previews
  const dlInputD = document.getElementById('dlPhotoD');
  const dlPreviewD = document.getElementById('dl-file-preview-d');
  const pcInputD = document.getElementById('punchCardD');
  const pcPreviewD = document.getElementById('pc-file-preview-d');
  const panInputD = document.getElementById('panPhotoD');
  const panPreviewD = document.getElementById('pan-file-preview-d');
  const passbookInputD = document.getElementById('passbookPhotoD');
  const passbookPreviewD = document.getElementById('pb-file-preview-d');
  const baseDocInputD = document.getElementById('baseDocumentD');
  const baseDocPreviewD = document.getElementById('bd-file-preview-d');

  // Stats elements
  const statTotal = document.getElementById('stat-total');
  const statExpired = document.getElementById('stat-expired');
  const statWarning = document.getElementById('stat-warning');
  const statActive = document.getElementById('stat-active');
  const alertsSection = document.getElementById('alerts-section');
  const alertsList = document.getElementById('alerts-list');

  // Search & Filters (Vehicles)
  const searchInput = document.getElementById('search-input');
  const statusFilter = document.getElementById('status-filter');
  const registryCount = document.getElementById('registry-count');

  // Lightbox Modal
  const modalLightbox = document.getElementById('modal-lightbox');
  const btnCloseLightbox = document.getElementById('btn-close-lightbox');
  const lightboxOverlay = document.getElementById('lightbox-overlay');
  const lightboxBody = document.getElementById('lightbox-body');
  const lightboxTitle = document.getElementById('lightbox-title');
  const btnDownloadDoc = document.getElementById('btn-download-doc');

  // Global State
  let allVehicles = [];
  let allDrivers = [];
  let currentUser = null;
  let currentFilter = 'all';
  let currentDriverFilter = 'all';
  let activeTab = 'registry';
  let isEditMode = false;
  let editVehicleNo = null;

  // Initialize Icons
  lucide.createIcons();

  // 1. Authenticate check on load
  checkAuth();

  async function checkAuth() {
    try {
      const response = await fetch('/api/auth/me');
      const data = await response.json();

      if (response.ok && data.authenticated) {
        currentUser = data;
        
        if (data.status === 'pending') {
          showStatusOverlay('pending', 'Access Request Pending', 'Your account registration is currently pending administrator review. Please check back later.');
        } else if (data.status === 'rejected') {
          showStatusOverlay('rejected', 'Access Request Rejected', 'Your registration approval has been denied. Contact your registry supervisor.');
        } else if (data.status === 'approved') {
          // Hide login and status overlays
          modalAuth.classList.remove('active');
          modalStatus.classList.remove('active');
          mainAppContainer.classList.remove('hidden');
          
          // Setup dashboard access levels
          configureUserPermissions();
        }
      } else {
        // Not logged in, show auth form
        currentUser = null;
        mainAppContainer.classList.add('hidden');
        modalStatus.classList.remove('active');
        modalAuth.classList.add('active');
        showLoginForm();
      }
    } catch (e) {
      console.error("Auth check failed:", e);
    }
  }

  function showStatusOverlay(statusType, title, message) {
    const iconContainer = document.getElementById('status-icon-container');
    const iconEl = document.getElementById('status-icon');
    const titleEl = document.getElementById('status-title');
    const msgEl = document.getElementById('status-message');

    titleEl.textContent = title;
    msgEl.textContent = message;

    if (statusType === 'pending') {
      iconContainer.className = 'status-icon-box bg-amber';
      iconEl.setAttribute('data-lucide', 'clock');
    } else {
      iconContainer.className = 'status-icon-box bg-rose';
      iconEl.setAttribute('data-lucide', 'alert-octagon');
    }

    modalAuth.classList.remove('active');
    mainAppContainer.classList.add('hidden');
    modalStatus.classList.add('active');
    lucide.createIcons();
  }

  function configureUserPermissions() {
    headerUsername.textContent = currentUser.username;
    
    const adminOnlyCells = document.querySelectorAll('.admin-only-cell');
    const btnExportExcel = document.getElementById('btn-export-excel');

    const canEdit = currentUser.role === 'admin' || currentUser.role === 'editor';
    const canManageUsers = currentUser.role === 'admin';

    // Show navigation tabs
    mainTabs.classList.remove('hidden');

    // Show/Hide Add Vehicle button based on edit permissions
    if (canEdit) {
      btnAddVehicle.classList.remove('hidden');
      btnAddDriver.classList.remove('hidden');
      btnExportExcel.classList.remove('hidden');
      adminOnlyCells.forEach(el => el.classList.remove('hidden'));
    } else {
      btnAddVehicle.classList.add('hidden');
      btnAddDriver.classList.add('hidden');
      btnExportExcel.classList.add('hidden');
      adminOnlyCells.forEach(el => el.classList.add('hidden'));
    }

    // Show/Hide User Access tab link based on admin role
    if (canManageUsers) {
      tabBtnUsers.classList.remove('hidden');
    } else {
      tabBtnUsers.classList.add('hidden');
      
      // If we are viewer/editor and somehow on users tab, switch back
      if (activeTab === 'users') {
        activeTab = 'registry';
        tabContentRegistry.classList.add('active');
        tabContentRegistry.classList.remove('hidden');
        tabContentUsers.classList.add('hidden');
        tabContentUsers.classList.remove('active');
        document.querySelectorAll('.tab-btn').forEach(b => {
          b.classList.remove('active');
          if (b.getAttribute('data-tab') === 'registry') b.classList.add('active');
        });
      }
    }

    // Load initial tab content
    if (activeTab === 'registry') {
      fetchVehicles();
    } else if (activeTab === 'drivers') {
      fetchDrivers();
    } else if (activeTab === 'users') {
      fetchUsers();
    }
  }

  // Switch between Login and Register Forms
  linkShowRegister.addEventListener('click', (e) => {
    e.preventDefault();
    formLogin.classList.add('hidden');
    formRegister.classList.remove('hidden');
    loginErrorMsg.classList.add('hidden');
  });

  linkShowLogin.addEventListener('click', (e) => {
    e.preventDefault();
    formRegister.classList.add('hidden');
    formLogin.classList.remove('hidden');
    regErrorMsg.classList.add('hidden');
  });

  function showLoginForm() {
    formRegister.classList.add('hidden');
    formLogin.classList.remove('hidden');
    loginErrorMsg.classList.add('hidden');
    regErrorMsg.classList.add('hidden');
  }

  // Handle Login submission
  formLogin.addEventListener('submit', async (e) => {
    e.preventDefault();
    loginErrorMsg.classList.add('hidden');
    
    const username = document.getElementById('login-username').value;
    const password = document.getElementById('login-password').value;
    
    const submitBtn = formLogin.querySelector('button[type="submit"]');
    setLoadingState(submitBtn, true);

    try {
      const response = await fetch('/api/auth/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username, password })
      });
      const data = await response.json();

      if (response.ok) {
        formLogin.reset();
        await checkAuth();
      } else {
        loginErrorMsg.textContent = data.error || "Login failed.";
        loginErrorMsg.classList.remove('hidden');
      }
    } catch (err) {
      console.error(err);
      loginErrorMsg.textContent = "Server error. Please try again.";
      loginErrorMsg.classList.remove('hidden');
    } finally {
      setLoadingState(submitBtn, false);
    }
  });

  // Handle Registration submission
  formRegister.addEventListener('submit', async (e) => {
    e.preventDefault();
    regErrorMsg.classList.add('hidden');

    const username = document.getElementById('reg-username').value;
    const password = document.getElementById('reg-password').value;

    const submitBtn = formRegister.querySelector('button[type="submit"]');
    setLoadingState(submitBtn, true);

    try {
      const response = await fetch('/api/auth/register', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username, password })
      });
      const data = await response.json();

      if (response.ok) {
        formRegister.reset();
        alert(data.message || "Registration request submitted. Awaiting admin approval.");
        // Switch back to login form
        formRegister.classList.add('hidden');
        formLogin.classList.remove('hidden');
      } else {
        regErrorMsg.textContent = data.error || "Registration failed.";
        regErrorMsg.classList.remove('hidden');
      }
    } catch (err) {
      console.error(err);
      regErrorMsg.textContent = "Server error. Please try again.";
      regErrorMsg.classList.remove('hidden');
    } finally {
      setLoadingState(submitBtn, false);
    }
  });

  // Handle Logout
  const triggerLogout = async () => {
    try {
      const response = await fetch('/api/auth/logout', { method: 'POST' });
      if (response.ok) {
        allVehicles = [];
        allDrivers = [];
        currentUser = null;
        await checkAuth();
      }
    } catch (err) {
      console.error("Logout failed:", err);
    }
  };

  btnLogout.addEventListener('click', triggerLogout);
  btnStatusLogout.addEventListener('click', triggerLogout);

  // Tab switching for Admin Dashboard
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const tabName = btn.getAttribute('data-tab');
      activeTab = tabName;

      // Update active tab buttons styling
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');

      // Hide all pane content
      tabContentRegistry.classList.add('hidden');
      tabContentRegistry.classList.remove('active');
      tabContentDrivers.classList.add('hidden');
      tabContentDrivers.classList.remove('active');
      tabContentUsers.classList.add('hidden');
      tabContentUsers.classList.remove('active');

      // Update active pane visibility
      if (tabName === 'registry') {
        tabContentRegistry.classList.add('active');
        tabContentRegistry.classList.remove('hidden');
        fetchVehicles();
      } else if (tabName === 'drivers') {
        tabContentDrivers.classList.add('active');
        tabContentDrivers.classList.remove('hidden');
        fetchDrivers();
      } else if (tabName === 'users') {
        tabContentUsers.classList.add('active');
        tabContentUsers.classList.remove('hidden');
        fetchUsers();
      }
    });
  });

  // Admin user approvals
  async function fetchUsers() {
    if (!currentUser || currentUser.role !== 'admin') return;

    try {
      const response = await fetch('/api/admin/users');
      if (!response.ok) throw new Error("Unauthorized");
      const users = await response.json();

      usersCount.textContent = `Showing ${users.length} users`;
      usersTableBody.innerHTML = '';

      if (users.length === 0) {
        usersTableBody.innerHTML = `
          <tr>
            <td colspan="5" style="text-align:center; padding: 2rem; color: var(--text-muted);">No users found.</td>
          </tr>
        `;
        return;
      }

      users.forEach(u => {
        const row = document.createElement('tr');
        
        let statusBadge = '';
        if (u.status === 'approved') {
          statusBadge = `<span class="badge badge-approved">Approved</span>`;
        } else if (u.status === 'rejected') {
          statusBadge = `<span class="badge badge-rejected">Rejected</span>`;
        } else {
          statusBadge = `<span class="badge badge-pending">Pending</span>`;
        }

        const isMainAdmin = u.username.toLowerCase() === 'admin';
        
        let roleSelector = '';
        if (!isMainAdmin) {
          roleSelector = `
            <select class="user-role-select" data-user="${u.username}" style="background: rgba(0, 0, 0, 0.35); border: 1px solid var(--glass-border); color: white; border-radius: 6px; padding: 0.25rem 0.5rem; font-size: 0.85rem; outline: none; cursor: pointer;">
              <option value="viewer" ${u.role === 'viewer' ? 'selected' : ''}>Viewer (Read-Only)</option>
              <option value="editor" ${u.role === 'editor' ? 'selected' : ''}>Editor (Registry Modifier)</option>
              <option value="admin" ${u.role === 'admin' ? 'selected' : ''}>Admin (Full Control)</option>
            </select>
          `;
        } else {
          roleSelector = `<span class="badge badge-role">${u.role}</span>`;
        }

        let actionButtons = '';
        if (!isMainAdmin) {
          actionButtons = `
            <div class="docs-cell">
              ${u.status !== 'approved' ? `
                <button class="btn btn-sm btn-primary btn-approve-user" data-user="${u.username}" style="background:#10b981; color:white; border:none; box-shadow:none;">
                  Approve
                </button>
              ` : ''}
              ${u.status !== 'rejected' ? `
                <button class="btn btn-sm btn-secondary btn-reject-user" data-user="${u.username}">
                  Reject
                </button>
              ` : ''}
              <button class="btn btn-sm btn-danger btn-delete-user" data-user="${u.username}">
                Delete
              </button>
            </div>
          `;
        } else {
          actionButtons = `<span style="font-size:0.8rem; font-style:italic; color:var(--text-muted);">Locked Admin Account</span>`;
        }

        row.innerHTML = `
          <td style="font-weight:600; color:white;">${u.username}</td>
          <td>${roleSelector}</td>
          <td>${statusBadge}</td>
          <td style="font-size:0.85rem; color:var(--text-muted);">${u.createdAt || '-'}</td>
          <td>${actionButtons}</td>
        `;

        usersTableBody.appendChild(row);
      });

      // Bind actions
      usersTableBody.querySelectorAll('.btn-approve-user').forEach(btn => {
        btn.addEventListener('click', () => updateUserStatus(btn.getAttribute('data-user'), 'approved'));
      });
      usersTableBody.querySelectorAll('.btn-reject-user').forEach(btn => {
        btn.addEventListener('click', () => updateUserStatus(btn.getAttribute('data-user'), 'rejected'));
      });
      usersTableBody.querySelectorAll('.btn-delete-user').forEach(btn => {
        btn.addEventListener('click', () => deleteUserAccount(btn.getAttribute('data-user')));
      });

      // Bind role change selectors
      usersTableBody.querySelectorAll('.user-role-select').forEach(select => {
        select.addEventListener('change', (e) => {
          const targetUser = select.getAttribute('data-user');
          const targetRole = e.target.value;
          updateUserRole(targetUser, targetRole);
        });
      });

    } catch (e) {
      console.error(e);
      usersTableBody.innerHTML = `
        <tr>
          <td colspan="5" style="text-align:center; padding: 2rem; color: var(--danger);">Failed to retrieve users.</td>
        </tr>
      `;
    }
  }

  async function updateUserStatus(username, newStatus) {
    try {
      const response = await fetch(`/api/admin/users/${encodeURIComponent(username)}/status`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ status: newStatus })
      });
      const data = await response.json();
      if (response.ok) {
        alert(data.message);
        await fetchUsers();
      } else {
        alert(data.error);
      }
    } catch (err) {
      console.error(err);
      alert("Error updating user status.");
    }
  }

  async function updateUserRole(username, newRole) {
    try {
      const response = await fetch(`/api/admin/users/${encodeURIComponent(username)}/role`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ role: newRole })
      });
      const data = await response.json();
      if (response.ok) {
        alert(data.message);
        await fetchUsers();
      } else {
        alert(data.error);
      }
    } catch (err) {
      console.error(err);
      alert("Error updating user permission level.");
    }
  }

  async function deleteUserAccount(username) {
    if (confirm(`Are you sure you want to permanently delete user account '${username}'?`)) {
      try {
        const response = await fetch(`/api/admin/users/${encodeURIComponent(username)}`, {
          method: 'DELETE'
        });
        const data = await response.json();
        if (response.ok) {
          alert(data.message);
          await fetchUsers();
        } else {
          alert(data.error);
        }
      } catch (err) {
        console.error(err);
        alert("Error deleting user.");
      }
    }
  }

  // --- Vehicle Registry Controller ---

  // Drawer Toggle
  btnAddVehicle.addEventListener('click', () => {
    modalAddVehicle.classList.add('active');
  });

  const closeAddModal = () => {
    modalAddVehicle.classList.remove('active');
    formAddVehicle.reset();
    dlPreview.textContent = 'No file chosen';
    pcPreview.textContent = 'No file chosen';

    // Reset edit state variables
    isEditMode = false;
    editVehicleNo = null;

    const modalTitle = document.querySelector('#modal-add-vehicle .modal-header h2');
    const submitBtnText = document.querySelector('#btn-submit-form .submit-text');
    
    if (modalTitle) modalTitle.textContent = "Register New Vehicle";
    if (submitBtnText) submitBtnText.textContent = "Register Vehicle";

    const vehicleNoInput = document.getElementById('vehicleNo');
    if (vehicleNoInput) {
      vehicleNoInput.readOnly = false;
      vehicleNoInput.style.opacity = '1.0';
    }
  };

  btnCloseAddModal.addEventListener('click', closeAddModal);
  btnCancelAdd.addEventListener('click', closeAddModal);

  // File pre-render listeners
  dlInput.addEventListener('change', (e) => {
    dlPreview.textContent = e.target.files.length > 0 ? e.target.files[0].name : 'No file chosen';
  });

  pcInput.addEventListener('change', (e) => {
    pcPreview.textContent = e.target.files.length > 0 ? e.target.files[0].name : 'No file chosen';
  });

  // Search & Filter listeners
  searchInput.addEventListener('input', filterAndRender);
  statusFilter.addEventListener('change', (e) => {
    currentFilter = e.target.value;
    filterAndRender();
  });

  // Stats Card quick filters
  document.querySelectorAll('.stat-card.cursor-pointer').forEach(card => {
    card.addEventListener('click', () => {
      const filterVal = card.getAttribute('data-filter');
      statusFilter.value = filterVal;
      currentFilter = filterVal;
      filterAndRender();
      document.querySelector('.registry-panel').scrollIntoView({ behavior: 'smooth' });
    });
  });

  // Form submission: Add or Edit Vehicle
  formAddVehicle.addEventListener('submit', async (e) => {
    e.preventDefault();
    
    const submitBtn = document.getElementById('btn-submit-form');
    setLoadingState(submitBtn, true);

    const formData = new FormData(formAddVehicle);
    const url = isEditMode ? `/api/vehicles/${encodeURIComponent(editVehicleNo)}/update` : '/api/vehicles';

    try {
      const response = await fetch(url, {
        method: 'POST',
        body: formData
      });

      const result = await response.json();

      if (response.ok) {
        alert(result.message || 'Vehicle details saved successfully!');
        closeAddModal();
        await fetchVehicles();
      } else {
        alert(result.error || 'Failed to save vehicle.');
      }
    } catch (error) {
      console.error(error);
      alert('An error occurred. Check server logs.');
    } finally {
      setLoadingState(submitBtn, false);
    }
  });

  // Fetch Vehicles API
  async function fetchVehicles() {
    if (!currentUser) return;
    
    try {
      const response = await fetch('/api/vehicles');
      if (!response.ok) throw new Error("Unauthorized access.");
      allVehicles = await response.json();
      
      updateStats();
      renderAlerts();
      filterAndRender();
    } catch (error) {
      console.error(error);
      vehicleTableBody.innerHTML = `
        <tr class="error-row">
          <td colspan="5" style="text-align: center; color: #ef4444; padding: 2rem;">
            Failed to load registry data from server.
          </td>
        </tr>
      `;
    }
  }

  // Stats Summary
  function updateStats() {
    const total = allVehicles.length;
    const expired = allVehicles.filter(v => v.status === 'Expired').length;
    const warning = allVehicles.filter(v => v.status === 'Warning').length;
    const active = allVehicles.filter(v => v.status === 'Active').length;

    statTotal.textContent = total;
    statExpired.textContent = expired;
    statWarning.textContent = warning;
    statActive.textContent = active;
  }

  // Render Expired Alerts
  function renderAlerts() {
    const expiredVehicles = allVehicles.filter(v => v.status === 'Expired');
    
    if (expiredVehicles.length === 0) {
      alertsSection.classList.add('hidden');
      return;
    }

    alertsSection.classList.remove('hidden');
    alertsList.innerHTML = '';

    const canEdit = currentUser.role === 'admin' || currentUser.role === 'editor';

    expiredVehicles.forEach(v => {
      const docsList = v.expiredDocs.map(d => `<span class="alert-doc-tag">${d.name} (${d.date})</span>`).join(', ');
      
      const alertItem = document.createElement('div');
      alertItem.className = 'alert-item';
      alertItem.innerHTML = `
        <div class="alert-info">
          <div class="alert-headline">
            <span class="vehicle-tag">${v.vehicleNo}</span>
            <span>- Expired: ${docsList}</span>
          </div>
          <div class="alert-subtext">
            Driver: <strong>${v.driverName || 'N/A'}</strong> ${v.punchNumber ? `| Punch No: <strong>${v.punchNumber}</strong>` : ''}
          </div>
          ${v.remark ? `<div class="alert-remark">Remark: "${v.remark}"</div>` : ''}
        </div>
        ${canEdit ? `
          <div class="alert-actions">
            <button class="btn btn-sm btn-danger btn-delete-vehicle" data-vehicle="${v.vehicleNo}">
              <i data-lucide="trash-2"></i> Remove
            </button>
          </div>
        ` : ''}
      `;
      alertsList.appendChild(alertItem);
    });

    alertsList.querySelectorAll('.btn-delete-vehicle').forEach(btn => {
      btn.addEventListener('click', () => {
        confirmDeleteVehicle(btn.getAttribute('data-vehicle'));
      });
    });

    lucide.createIcons();
  }

  // Registry Table Rendering
  function filterAndRender() {
    const searchQuery = searchInput.value.toLowerCase().trim();
    
    const filtered = allVehicles.filter(v => {
      const matchesSearch = v.vehicleNo.toLowerCase().includes(searchQuery) ||
                            v.driverName.toLowerCase().includes(searchQuery) ||
                            (v.punchNumber && v.punchNumber.toLowerCase().includes(searchQuery)) ||
                            (v.stageInspection && v.stageInspection.toLowerCase().includes(searchQuery)) ||
                            (v.stageSafety && v.stageSafety.toLowerCase().includes(searchQuery)) ||
                            (v.stageDept && v.stageDept.toLowerCase().includes(searchQuery)) ||
                            (v.status && v.status.toLowerCase().includes(searchQuery));
                            
      if (currentFilter === 'all') return matchesSearch;
      return matchesSearch && v.status.toLowerCase() === currentFilter.toLowerCase();
    });

    registryCount.textContent = `Showing ${filtered.length} of ${allVehicles.length} vehicles`;
    
    // Fixed 5 columns layout
    const colspanCount = 5;
    const tableColspanCell = document.getElementById('table-colspan-cell');
    if (tableColspanCell) tableColspanCell.colSpan = colspanCount;

    if (filtered.length === 0) {
      vehicleTableBody.innerHTML = `
        <tr class="empty-row">
          <td colspan="${colspanCount}" style="text-align: center; color: var(--text-muted); padding: 3rem;">
            No vehicles found matching filters.
          </td>
        </tr>
      `;
      return;
    }

    vehicleTableBody.innerHTML = '';
    
    const canEdit = currentUser.role === 'admin' || currentUser.role === 'editor';

    filtered.forEach(v => {
      const row = document.createElement('tr');
      
      let statusBadge = '';
      if (v.status === 'Expired') {
        statusBadge = `<span class="badge badge-expired"><i data-lucide="alert-octagon" style="width:12px;height:12px"></i> Expired</span>`;
      } else if (v.status === 'Warning') {
        statusBadge = `<span class="badge badge-warning"><i data-lucide="alert-triangle" style="width:12px;height:12px"></i> Warning</span>`;
      } else {
        statusBadge = `<span class="badge badge-active"><i data-lucide="check-circle-2" style="width:12px;height:12px"></i> Active</span>`;
      }

      // Clearance Stage Column with 3 Toggles
      const getStagePillHTML = (val, label, stageKey, vehicleNo, editable) => {
        const isDone = val === 'Done';
        const bg = isDone ? 'rgba(16, 185, 129, 0.15)' : 'rgba(255, 255, 255, 0.03)';
        const color = isDone ? '#10b981' : '#94a3b8';
        const border = isDone ? '1px solid rgba(16, 185, 129, 0.3)' : '1px solid var(--glass-border)';
        const cursor = editable ? 'cursor:pointer;' : '';
        const hoverClass = editable ? 'stage-pill-toggle' : '';
        
        return `
          <span class="${hoverClass}" data-vehicle="${vehicleNo}" data-stage="${stageKey}" style="display:inline-flex; align-items:center; gap:0.25rem; padding:0.2rem 0.5rem; border-radius:6px; font-size:0.75rem; font-weight:600; ${border}; background:${bg}; color:${color}; ${cursor}">
            <i data-lucide="${isDone ? 'check' : 'minus'}" style="width:10px; height:10px;"></i>
            ${label}
          </span>
        `;
      };

      const isAllDone = v.stageInspection === 'Done' && v.stageSafety === 'Done' && v.stageDept === 'Done';
      const summaryBadge = isAllDone
        ? `<span class="badge badge-approved" style="margin-top:0.25rem; font-size:0.65rem; font-weight:700;"><i data-lucide="check-circle-2" style="width:10px;height:10px;"></i> Ready: Gatepass Office</span>`
        : `<span class="badge badge-pending" style="margin-top:0.25rem; font-size:0.65rem; font-weight:700;"><i data-lucide="clock" style="width:10px;height:10px;"></i> Pending Clearances</span>`;

      const stageCell = `
        <td>
          <div style="display:flex; flex-direction:column; gap:0.25rem; min-width:140px;">
            <div style="display:flex; flex-wrap:wrap; gap:0.25rem;">
              ${getStagePillHTML(v.stageInspection, 'Insp', 'Inspection', v.vehicleNo, canEdit)}
              ${getStagePillHTML(v.stageSafety, 'Safety', 'Safety', v.vehicleNo, canEdit)}
              ${getStagePillHTML(v.stageDept, 'Dept', 'Department', v.vehicleNo, canEdit)}
            </div>
            ${summaryBadge}
          </div>
        </td>
      `;

      // Compliance Status Column
      let complianceCell = '';
      if (v.status === 'Expired') {
        const count = v.expiredDocs.length;
        complianceCell = `<td><span class="badge badge-expired"><i data-lucide="alert-octagon" style="width:12px;height:12px"></i> Expired (${count} Doc${count > 1 ? 's' : ''})</span></td>`;
      } else if (v.status === 'Warning') {
        const count = v.expiringSoonDocs.length;
        complianceCell = `<td><span class="badge badge-warning"><i data-lucide="alert-triangle" style="width:12px;height:12px"></i> Warning (${count} Doc${count > 1 ? 's' : ''})</span></td>`;
      } else {
        complianceCell = `<td><span class="badge badge-active"><i data-lucide="check-circle-2" style="width:12px;height:12px"></i> Compliant</span></td>`;
      }

      const dlBtnClass = v.dlPhoto ? '' : 'disabled';
      const pcBtnClass = v.punchCard ? '' : 'disabled';

      row.innerHTML = `
        <td>
          <div class="vehicle-cell">
            <div class="vehicle-number-wrapper">
              <span class="vehicle-no" title="Click to view details" style="cursor:pointer; font-weight:700; text-decoration:underline;">${v.vehicleNo}</span>
              ${v.status === 'Expired' ? '<span class="badge badge-expired" style="font-size:0.65rem; padding:0.1rem 0.3rem;">Expired</span>' : ''}
            </div>
            <span class="driver-name">${v.driverName || 'No Driver Assigned'}</span>
            ${v.punchNumber ? `<span class="punch-number-text" style="font-size:0.75rem; color:var(--primary); font-weight:600; display:flex; align-items:center; gap:2px; margin-top:2px;"><i data-lucide="hash" style="width:10px;height:10px;"></i>Punch: ${v.punchNumber}</span>` : ''}
          </div>
        </td>
        ${stageCell}
        ${complianceCell}
        <td>
          <div class="docs-cell">
            <button class="btn-doc ${dlBtnClass}" data-file="${v.dlPhoto}" data-title="DL: ${v.vehicleNo}" title="View Driver's License">
              <i data-lucide="file-image"></i> DL
            </button>
            <button class="btn-doc ${pcBtnClass}" data-file="${v.punchCard}" data-title="Punch Card: ${v.vehicleNo}" title="View Punch Card">
              <i data-lucide="file-text"></i> PC
            </button>
          </div>
        </td>
        <td>
          <div class="docs-cell">
            <button class="btn btn-sm btn-primary btn-view-vehicle" data-vehicle="${v.vehicleNo}" style="background:var(--primary); border:none; box-shadow:none;" title="View Details">
              <i data-lucide="eye"></i> View
            </button>
            ${canEdit ? `
              <button class="btn btn-sm btn-danger btn-delete-vehicle" data-vehicle="${v.vehicleNo}" title="Remove Vehicle">
                <i data-lucide="trash-2"></i>
              </button>
            ` : ''}
          </div>
        </td>
      `;
      
      vehicleTableBody.appendChild(row);
    });

    // Bind table actions - View Details
    vehicleTableBody.querySelectorAll('.btn-view-vehicle').forEach(btn => {
      btn.addEventListener('click', () => {
        const vehicleNo = btn.getAttribute('data-vehicle');
        const v = allVehicles.find(x => x.vehicleNo === vehicleNo);
        if (v) openViewModal(v);
      });
    });

    // Bind clicking on Vehicle Number
    vehicleTableBody.querySelectorAll('.vehicle-no').forEach(el => {
      el.addEventListener('click', () => {
        const row = el.closest('tr');
        const viewBtn = row.querySelector('.btn-view-vehicle');
        if (viewBtn) viewBtn.click();
      });
    });

    // Bind delete action
    vehicleTableBody.querySelectorAll('.btn-delete-vehicle').forEach(btn => {
      btn.addEventListener('click', () => {
        confirmDeleteVehicle(btn.getAttribute('data-vehicle'));
      });
    });

    // Bind file lightboxes
    vehicleTableBody.querySelectorAll('.btn-doc:not(.disabled)').forEach(btn => {
      btn.addEventListener('click', () => {
        const file = btn.getAttribute('data-file');
        const title = btn.getAttribute('data-title');
        openLightbox(file, title);
      });
    });

    // Bind stage toggles
    vehicleTableBody.querySelectorAll('.stage-pill-toggle').forEach(pill => {
      pill.addEventListener('click', async (e) => {
        e.stopPropagation();
        const vehicleNo = pill.getAttribute('data-vehicle');
        const stageKey = pill.getAttribute('data-stage');
        
        try {
          const response = await fetch(`/api/vehicles/${encodeURIComponent(vehicleNo)}/toggle-stage`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ stage: stageKey })
          });
          const result = await response.json();
          if (response.ok) {
            await fetchVehicles(); // Reload to refresh table
          } else {
            alert(result.error || 'Failed to toggle stage.');
          }
        } catch (error) {
          console.error(error);
          alert('An error occurred during stage update.');
        }
      });
    });

    lucide.createIcons();
  }

  // View Vehicle Modal controllers
  function openViewModal(v) {
    const viewTitle = document.getElementById('view-title');
    const viewBody = document.getElementById('view-modal-body');
    const viewFooter = document.getElementById('view-modal-footer');

    viewTitle.textContent = `Details: ${v.vehicleNo}`;

    const getDocStatusHTML = (fieldVal, fieldName) => {
      if (!fieldVal) return `<span class="date-val text-muted">- (Not Provided)</span>`;
      const isExp = v.expiredDocs.some(d => d.name === fieldName);
      const isWarn = v.expiringSoonDocs.some(d => d.name === fieldName);
      
      let badgeClass = 'badge-approved';
      let statusText = 'Compliant';
      let dateCls = '';
      
      if (isExp) {
        badgeClass = 'badge-expired';
        statusText = 'Expired';
        dateCls = 'expired';
      } else if (isWarn) {
        badgeClass = 'badge-warning';
        const details = v.expiringSoonDocs.find(d => d.name === fieldName);
        statusText = `Expires in ${details.daysLeft} days`;
        dateCls = 'warning';
      }
      
      return `
        <div class="detail-doc-row" style="display:flex; justify-content:space-between; align-items:center; padding:0.6rem 0; border-bottom:1px solid rgba(255,255,255,0.06)">
          <span style="font-weight:600; color:var(--text-muted); font-size:0.9rem;">${fieldName} Expiry</span>
          <div style="display:flex; align-items:center; gap:0.75rem;">
            <span class="date-val ${dateCls}" style="font-family:var(--font-sans); font-size:0.9rem; font-weight:600;">${fieldVal}</span>
            <span class="badge ${badgeClass}" style="text-transform:none; font-size:0.7rem; font-weight:600; padding:0.15rem 0.4rem;">${statusText}</span>
          </div>
        </div>
      `;
    };

    const dlThumbnail = v.dlPhoto 
      ? `<div class="attachment-thumb" data-file="${v.dlPhoto}" data-title="DL: ${v.vehicleNo}" style="cursor:pointer; display:flex; align-items:center; gap:0.5rem; background:rgba(255,255,255,0.04); border:1px solid var(--glass-border); padding:0.5rem 1rem; border-radius:10px; margin-top:0.25rem;">
           <i data-lucide="file-image" style="color:var(--primary)"></i>
           <span style="font-size:0.85rem; color:white; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; max-width:180px;">${v.dlPhoto}</span>
         </div>`
      : `<span style="font-size:0.85rem; color:var(--text-muted);">No DL document uploaded.</span>`;

    const pcThumbnail = v.punchCard 
      ? `<div class="attachment-thumb" data-file="${v.punchCard}" data-title="Punch Card: ${v.vehicleNo}" style="cursor:pointer; display:flex; align-items:center; gap:0.5rem; background:rgba(255,255,255,0.04); border:1px solid var(--glass-border); padding:0.5rem 1rem; border-radius:10px; margin-top:0.25rem;">
           <i data-lucide="file-text" style="color:var(--primary)"></i>
           <span style="font-size:0.85rem; color:white; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; max-width:180px;">${v.punchCard}</span>
         </div>`
      : `<span style="font-size:0.85rem; color:var(--text-muted);">No Punch Card document uploaded.</span>`;

    const isAllDone = v.stageInspection === 'Done' && v.stageSafety === 'Done' && v.stageDept === 'Done';

    viewBody.innerHTML = `
      <!-- Info Section -->
      <div class="form-section">
        <h3>Vehicle & Driver Information</h3>
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:1.25rem; margin-top:0.75rem;">
          <div>
            <span style="font-size:0.8rem; color:var(--text-muted); display:block; margin-bottom:0.2rem;">Vehicle Number</span>
            <strong style="font-size:1.15rem; color:white; font-family:var(--font-display); letter-spacing:0.5px;">${v.vehicleNo}</strong>
          </div>
          <div>
            <span style="font-size:0.8rem; color:var(--text-muted); display:block; margin-bottom:0.2rem;">Driver Name</span>
            <strong style="font-size:1.15rem; color:white;">${v.driverName || 'N/A'}</strong>
          </div>
          <div>
            <span style="font-size:0.8rem; color:var(--text-muted); display:block; margin-bottom:0.2rem;">Punch Card Number</span>
            <strong style="font-size:1rem; color:white;">${v.punchNumber || 'N/A'}</strong>
          </div>
          <div>
            <span style="font-size:0.8rem; color:var(--text-muted); display:block; margin-bottom:0.2rem;">Clearance Summary</span>
            <strong style="font-size:1rem; display:block; margin-top:0.2rem;">
              ${isAllDone 
                ? `<span class="badge badge-approved" style="font-size:0.7rem; font-weight:700;"><i data-lucide="check-circle-2" style="width:10px;height:10px;"></i> Ready: Gatepass Office</span>`
                : `<span class="badge badge-pending" style="font-size:0.7rem; font-weight:700;"><i data-lucide="clock" style="width:10px;height:10px;"></i> Pending Signs</span>`}
            </strong>
          </div>
        </div>
        
        <div style="display:grid; grid-template-columns:repeat(3, 1fr); gap:0.5rem; margin-top:1rem; padding-top:0.75rem; border-top:1px solid rgba(255,255,255,0.05)">
          <div>
            <span style="font-size:0.75rem; color:var(--text-muted); display:block;">Inspection Stage</span>
            <span class="badge ${v.stageInspection === 'Done' ? 'badge-approved' : 'badge-role'}" style="margin-top:0.25rem; font-size:0.75rem;">${v.stageInspection}</span>
          </div>
          <div>
            <span style="font-size:0.75rem; color:var(--text-muted); display:block;">Safety Sign</span>
            <span class="badge ${v.stageSafety === 'Done' ? 'badge-approved' : 'badge-role'}" style="margin-top:0.25rem; font-size:0.75rem;">${v.stageSafety}</span>
          </div>
          <div>
            <span style="font-size:0.75rem; color:var(--text-muted); display:block;">Department Sign</span>
            <span class="badge ${v.stageDept === 'Done' ? 'badge-approved' : 'badge-role'}" style="margin-top:0.25rem; font-size:0.75rem;">${v.stageDept}</span>
          </div>
        </div>
      </div>

      <!-- Compliance Documents -->
      <div class="form-section">
        <h3>Compliance Documents Status</h3>
        <div style="display:flex; flex-direction:column; margin-top:0.5rem; background:rgba(0,0,0,0.15); border-radius:12px; padding:0.5rem 1rem; border:1px solid var(--glass-border);">
          ${getDocStatusHTML(v.gatepassExpiry, 'Gatepass')}
          ${getDocStatusHTML(v.insuranceExpiry, 'Insurance')}
          ${getDocStatusHTML(v.pucExpiry, 'PUC')}
          ${getDocStatusHTML(v.fitnessExpiry, 'Fitness')}
          ${getDocStatusHTML(v.taxExpiry, 'Tax')}
          ${getDocStatusHTML(v.permitExpiry, 'Permit')}
        </div>
      </div>

      <!-- Remarks -->
      <div class="form-section">
        <h3>Remarks & Notes</h3>
        <div style="background:rgba(0,0,0,0.25); border:1px solid var(--glass-border); padding:0.75rem 1rem; border-radius:10px; font-size:0.9rem; line-height:1.5; color:white; min-height:60px; word-break:break-word;">
          ${v.remark || '<em style="color:var(--text-muted)">No remarks added for this vehicle.</em>'}
        </div>
      </div>

      <!-- Attachments -->
      <div class="form-section">
        <h3>Document Attachments</h3>
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:1rem; margin-top:0.5rem;">
          <div>
            <span style="font-size:0.8rem; color:var(--text-muted); display:block; margin-bottom:0.25rem;">Driver's License (DL)</span>
            ${dlThumbnail}
          </div>
          <div>
            <span style="font-size:0.8rem; color:var(--text-muted); display:block; margin-bottom:0.25rem;">Punch Card</span>
            ${pcThumbnail}
          </div>
        </div>
      </div>

      <!-- Footer Info -->
      <div style="font-size:0.75rem; color:var(--text-muted); text-align:right; margin-top:1.5rem;">
        Registered on: ${v.createdAt || '-'}
      </div>
    `;

    // Footer actions
    const canEdit = currentUser.role === 'admin' || currentUser.role === 'editor';
    viewFooter.innerHTML = `
      <div style="display:flex; gap:0.5rem; width:100%;">
        <button class="btn btn-secondary" id="btn-close-view-drawer" style="flex:1;">Close</button>
        ${canEdit ? `
          <button class="btn btn-primary" id="btn-edit-view-drawer" style="flex:1; background:#3b82f6; border:none; box-shadow:none; display:inline-flex; align-items:center; justify-content:center; gap:4px;">
            <i data-lucide="edit"></i> Edit Details
          </button>
          <button class="btn btn-danger" id="btn-delete-view-drawer" style="flex:1; display:inline-flex; align-items:center; justify-content:center; gap:4px;">
            <i data-lucide="trash-2"></i> Remove
          </button>
        ` : ''}
      </div>
    `;

    // Open Modal
    document.getElementById('modal-view-vehicle').classList.add('active');

    // Bind inner actions
    document.getElementById('btn-close-view-drawer').addEventListener('click', closeViewModal);
    
    if (canEdit) {
      document.getElementById('btn-edit-view-drawer').addEventListener('click', () => {
        closeViewModal();
        openEditModal(v);
      });
      document.getElementById('btn-delete-view-drawer').addEventListener('click', () => {
        closeViewModal();
        confirmDeleteVehicle(v.vehicleNo);
      });
    }

    viewBody.querySelectorAll('.attachment-thumb').forEach(thumb => {
      thumb.addEventListener('click', () => {
        const file = thumb.getAttribute('data-file');
        const title = thumb.getAttribute('data-title');
        openLightbox(file, title);
      });
    });

    lucide.createIcons();
  }

  const closeViewModal = () => {
    document.getElementById('modal-view-vehicle').classList.remove('active');
  };

  document.getElementById('btn-close-view-modal').addEventListener('click', closeViewModal);
  document.getElementById('view-modal-overlay').addEventListener('click', closeViewModal);

  function openEditModal(v) {
    isEditMode = true;
    editVehicleNo = v.vehicleNo;

    // Update modal labels
    const modalTitle = document.querySelector('#modal-add-vehicle .modal-header h2');
    const submitBtnText = document.querySelector('#btn-submit-form .submit-text');
    
    if (modalTitle) modalTitle.textContent = "Edit Vehicle Details";
    if (submitBtnText) submitBtnText.textContent = "Save Changes";

    // Fill inputs
    const vehicleNoInput = document.getElementById('vehicleNo');
    if (vehicleNoInput) {
      vehicleNoInput.value = v.vehicleNo;
      vehicleNoInput.readOnly = true;
      vehicleNoInput.style.opacity = '0.6';
    }

    document.getElementById('driverName').value = v.driverName || '';
    document.getElementById('punchNumber').value = v.punchNumber || '';
    document.getElementById('gatepassExpiry').value = v.gatepassExpiry || '';
    document.getElementById('insuranceExpiry').value = v.insuranceExpiry || '';
    document.getElementById('pucExpiry').value = v.pucExpiry || '';
    document.getElementById('fitnessExpiry').value = v.fitnessExpiry || '';
    document.getElementById('taxExpiry').value = v.taxExpiry || '';
    document.getElementById('permitExpiry').value = v.permitExpiry || '';
    
    document.getElementById('stageInspection').value = v.stageInspection || 'Pending';
    document.getElementById('stageSafety').value = v.stageSafety || 'Pending';
    document.getElementById('stageDept').value = v.stageDept || 'Pending';
    
    document.getElementById('remark').value = v.remark || '';

    dlPreview.textContent = v.dlPhoto ? v.dlPhoto : 'No file chosen';
    pcPreview.textContent = v.punchCard ? v.punchCard : 'No file chosen';

    // Show modal
    modalAddVehicle.classList.add('active');
  }

  // Deletion logic
  async function confirmDeleteVehicle(vehicleNo) {
    const canEdit = currentUser.role === 'admin' || currentUser.role === 'editor';
    if (!canEdit) return;

    if (confirm(`Are you sure you want to remove vehicle ${vehicleNo} and all its uploaded documents?`)) {
      try {
        const response = await fetch(`/api/vehicles/${encodeURIComponent(vehicleNo)}`, {
          method: 'DELETE'
        });
        const result = await response.json();
        if (response.ok) {
          alert(result.message || 'Vehicle removed successfully.');
          await fetchVehicles();
        } else {
          alert(result.error || 'Failed to remove vehicle.');
        }
      } catch (error) {
        console.error(error);
        alert('An error occurred during vehicle deletion.');
      }
    }
  }


  // --- Drivers / Punch Cards Controller ---

  // Drawer Toggle
  btnAddDriver.addEventListener('click', () => {
    modalAddDriver.classList.add('active');
  });

  const closeDriverModal = () => {
    modalAddDriver.classList.remove('active');
    formAddDriver.reset();
    dlPreviewD.textContent = 'No file chosen';
    pcPreviewD.textContent = 'No file chosen';
    panPreviewD.textContent = 'No file chosen';
    passbookPreviewD.textContent = 'No file chosen';
    baseDocPreviewD.textContent = 'No file chosen';
  };

  btnCloseDriverModal.addEventListener('click', closeDriverModal);
  btnCancelAddDriver.addEventListener('click', closeDriverModal);
  document.getElementById('add-driver-overlay').addEventListener('click', closeDriverModal);

  // File pre-render listeners for Driver Form
  dlInputD.addEventListener('change', (e) => {
    dlPreviewD.textContent = e.target.files.length > 0 ? e.target.files[0].name : 'No file chosen';
  });
  pcInputD.addEventListener('change', (e) => {
    pcPreviewD.textContent = e.target.files.length > 0 ? e.target.files[0].name : 'No file chosen';
  });
  panInputD.addEventListener('change', (e) => {
    panPreviewD.textContent = e.target.files.length > 0 ? e.target.files[0].name : 'No file chosen';
  });
  passbookInputD.addEventListener('change', (e) => {
    passbookPreviewD.textContent = e.target.files.length > 0 ? e.target.files[0].name : 'No file chosen';
  });
  baseDocInputD.addEventListener('change', (e) => {
    baseDocPreviewD.textContent = e.target.files.length > 0 ? e.target.files[0].name : 'No file chosen';
  });

  // Search & Filter listeners for Drivers
  driverSearchInput.addEventListener('input', filterAndRenderDrivers);
  driverStatusFilter.addEventListener('change', (e) => {
    currentDriverFilter = e.target.value;
    filterAndRenderDrivers();
  });

  // Fetch Drivers API
  async function fetchDrivers() {
    if (!currentUser) return;

    try {
      const response = await fetch('/api/drivers');
      if (!response.ok) throw new Error("Unauthorized access.");
      allDrivers = await response.json();
      filterAndRenderDrivers();
    } catch (err) {
      console.error(err);
      driversTableBody.innerHTML = `
        <tr class="error-row">
          <td colspan="6" style="text-align: center; color: #ef4444; padding: 2rem;">
            Failed to load driver punch card registry from server.
          </td>
        </tr>
      `;
    }
  }

  // Filter & Render Drivers Table
  function filterAndRenderDrivers() {
    const searchQuery = driverSearchInput.value.toLowerCase().trim();

    const filtered = allDrivers.filter(d => {
      const matchesSearch = d.driverName.toLowerCase().includes(searchQuery) ||
                            d.punchNumber.toLowerCase().includes(searchQuery) ||
                            (d.dlNumber && d.dlNumber.toLowerCase().includes(searchQuery)) ||
                            (d.panNumber && d.panNumber.toLowerCase().includes(searchQuery));

      if (currentDriverFilter === 'all') return matchesSearch;
      return matchesSearch && d.status.toLowerCase() === currentDriverFilter.toLowerCase();
    });

    driversCount.textContent = `Showing ${filtered.length} of ${allDrivers.length} drivers`;

    if (filtered.length === 0) {
      driversTableBody.innerHTML = `
        <tr class="empty-row">
          <td colspan="6" style="text-align: center; color: var(--text-muted); padding: 3rem;">
            No driver profiles found matching filters.
          </td>
        </tr>
      `;
      return;
    }

    driversTableBody.innerHTML = '';
    const canEdit = currentUser.role === 'admin' || currentUser.role === 'editor';

    filtered.forEach(d => {
      const row = document.createElement('tr');

      // Attachment buttons classes
      const dlClass = d.dlPhoto ? '' : 'disabled';
      const pcClass = d.punchCard ? '' : 'disabled';
      const panClass = d.panPhoto ? '' : 'disabled';
      const pbClass = d.passbookPhoto ? '' : 'disabled';
      const bdClass = d.baseDocument ? '' : 'disabled';

      // Status Toggle Badge
      const isActive = d.status === 'Active';
      const statusBadgeClass = isActive ? 'badge-approved' : 'badge-expired';
      const cursorStyle = canEdit ? 'cursor:pointer;' : '';
      const hoverClass = canEdit ? 'driver-status-toggle' : '';

      row.innerHTML = `
        <td>
          <div style="display:flex; flex-direction:column;">
            <strong style="color:white; font-size:0.95rem;">${d.driverName}</strong>
            <span style="font-size:0.75rem; color:var(--text-muted); margin-top:2px;">Reg on: ${d.createdAt.split(' ')[0]}</span>
          </div>
        </td>
        <td>
          <span class="badge badge-role" style="font-family:var(--font-sans); text-transform:none; font-size:0.8rem; font-weight:600;">
            <i data-lucide="hash" style="width:10px;height:10px;display:inline;"></i> ${d.punchNumber}
          </span>
        </td>
        <td>
          <div style="display:flex; flex-direction:column; gap:0.2rem; font-size:0.8rem; color:var(--text-muted);">
            <div>DL No: <strong style="color:white; font-family:var(--font-sans);">${d.dlNumber || 'N/A'}</strong></div>
            <div>PAN: <strong style="color:white; font-family:var(--font-sans);">${d.panNumber || 'N/A'}</strong></div>
            <div>Bank Acc: <strong style="color:white; font-family:var(--font-sans);">${d.accountNumber || 'N/A'}</strong></div>
          </div>
        </td>
        <td>
          <div class="docs-cell">
            <button class="btn-doc ${dlClass}" data-file="${d.dlPhoto}" data-title="DL: ${d.driverName}" title="Driver's License">
              <i data-lucide="file-image"></i> DL
            </button>
            <button class="btn-doc ${pcClass}" data-file="${d.punchCard}" data-title="Punch Card: ${d.driverName}" title="Punch Card">
              <i data-lucide="file-text"></i> PC
            </button>
            <button class="btn-doc ${panClass}" data-file="${d.panPhoto}" data-title="PAN: ${d.driverName}" title="PAN Card Document">
              <i data-lucide="credit-card"></i> PAN
            </button>
            <button class="btn-doc ${pbClass}" data-file="${d.passbookPhoto}" data-title="Passbook: ${d.driverName}" title="Bank Passbook Document">
              <i data-lucide="book-open"></i> Bank
            </button>
            <button class="btn-doc ${bdClass}" data-file="${d.baseDocument}" data-title="Base Doc: ${d.driverName}" title="Punch Card Request / Verification Document">
              <i data-lucide="file-check"></i> Base Doc
            </button>
          </div>
        </td>
        <td>
          <span class="badge ${statusBadgeClass} ${hoverClass}" data-punch="${d.punchNumber}" style="${cursorStyle} font-size:0.75rem; font-weight:700; text-transform:uppercase;">
            <i data-lucide="${isActive ? 'check' : 'clock'}" style="width:10px;height:10px;"></i>
            ${d.status}
          </span>
        </td>
        <td>
          <div class="docs-cell">
            ${canEdit ? `
              <button class="btn btn-sm btn-danger btn-delete-driver" data-punch="${d.punchNumber}" title="Delete Driver Profile">
                <i data-lucide="trash-2"></i> Remove
              </button>
            ` : `<span style="font-size:0.8rem; color:var(--text-muted); font-style:italic;">Read-Only</span>`}
          </div>
        </td>
      `;

      driversTableBody.appendChild(row);
    });

    // Bind document lightboxes
    driversTableBody.querySelectorAll('.btn-doc:not(.disabled)').forEach(btn => {
      btn.addEventListener('click', () => {
        const file = btn.getAttribute('data-file');
        const title = btn.getAttribute('data-title');
        openLightbox(file, title);
      });
    });

    // Bind toggling statuses
    if (canEdit) {
      driversTableBody.querySelectorAll('.driver-status-toggle').forEach(el => {
        el.addEventListener('click', async () => {
          const punchNo = el.getAttribute('data-punch');
          try {
            const response = await fetch(`/api/drivers/${encodeURIComponent(punchNo)}/toggle-status`, {
              method: 'POST'
            });
            const result = await response.json();
            if (response.ok) {
              await fetchDrivers();
            } else {
              alert(result.error || 'Failed to toggle status.');
            }
          } catch (err) {
            console.error(err);
            alert('An error occurred updating driver status.');
          }
        });
      });

      // Bind deleting driver
      driversTableBody.querySelectorAll('.btn-delete-driver').forEach(btn => {
        btn.addEventListener('click', async () => {
          const punchNo = btn.getAttribute('data-punch');
          if (confirm(`Are you sure you want to permanently delete driver profile with punch card '${punchNo}' and all associated document uploads?`)) {
            try {
              const response = await fetch(`/api/drivers/${encodeURIComponent(punchNo)}`, {
                method: 'DELETE'
              });
              const result = await response.json();
              if (response.ok) {
                alert(result.message || 'Driver removed successfully.');
                await fetchDrivers();
              } else {
                alert(result.error || 'Failed to delete driver.');
              }
            } catch (err) {
              console.error(err);
              alert('An error occurred during deletion.');
            }
          }
        });
      });
    }

    lucide.createIcons();
  }

  // Handle Driver Registration submission
  formAddDriver.addEventListener('submit', async (e) => {
    e.preventDefault();

    const submitBtn = document.getElementById('btn-submit-driver-form');
    setLoadingState(submitBtn, true);

    const formData = new FormData(formAddDriver);

    try {
      const response = await fetch('/api/drivers', {
        method: 'POST',
        body: formData
      });

      const result = await response.json();

      if (response.ok) {
        alert(result.message || 'Driver profile registered successfully!');
        closeDriverModal();
        await fetchDrivers();
      } else {
        alert(result.error || 'Failed to save driver profile.');
      }
    } catch (err) {
      console.error(err);
      alert('An error occurred. Check server logs.');
    } finally {
      setLoadingState(submitBtn, false);
    }
  });


  // Lightbox Viewer
  function openLightbox(filename, title) {
    lightboxTitle.textContent = title;
    const fileUrl = `/uploads/${filename}`;
    btnDownloadDoc.href = fileUrl;
    
    // Hide download button for view-only users
    const canEdit = currentUser.role === 'admin' || currentUser.role === 'editor';
    if (!canEdit) {
      btnDownloadDoc.classList.add('hidden');
    } else {
      btnDownloadDoc.classList.remove('hidden');
    }

    const ext = filename.split('.').pop().toLowerCase();
    
    if (ext === 'pdf') {
      lightboxBody.innerHTML = `<embed src="${fileUrl}" type="application/pdf">`;
    } else {
      lightboxBody.innerHTML = `<img src="${fileUrl}" alt="${title}">`;
    }
    
    modalLightbox.classList.add('active');
  }

  const closeLightbox = () => {
    modalLightbox.classList.remove('active');
    setTimeout(() => { lightboxBody.innerHTML = ''; }, 200);
  };

  btnCloseLightbox.addEventListener('click', closeLightbox);
  lightboxOverlay.addEventListener('click', closeLightbox);

  // Helper function to toggle button loading spinners
  function setLoadingState(buttonEl, isLoading) {
    const textEl = buttonEl.querySelector('.btn-text') || buttonEl.querySelector('.submit-text');
    const spinnerEl = buttonEl.querySelector('.btn-spinner');
    
    if (isLoading) {
      buttonEl.disabled = true;
      if (textEl) textEl.classList.add('hidden');
      if (spinnerEl) spinnerEl.classList.remove('hidden');
    } else {
      buttonEl.disabled = false;
      if (textEl) textEl.classList.remove('hidden');
      if (spinnerEl) spinnerEl.classList.add('hidden');
    }
  }
});
