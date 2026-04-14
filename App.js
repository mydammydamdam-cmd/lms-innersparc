(() => {
  'use strict';

  const state = {
    token: '',
    user: null,
    statuses: [
      'Inquiry',
      'Site Tour',
      'Down Payment',
      'Approval',
      'Takeout',
      'Turnover',
      'Closed/Lost'
    ],
    options: {
      classifications: ['OFW', 'Locally Employed', 'Self-Employed', 'Unknown'],
      sources: ['TikTok Ads', 'FB Ads', 'Organic', 'Referral', 'KKK'],
      temperatures: ['Hot', 'Warm', 'Cold'],
      paymentTypes: ['Spot', 'Installment']
    },
    projects: [],
    leadsByStatus: {},
    activeLead: null,
    activeTransaction: null,
    voiceMode: 'manual',
    transcript: '',
    recognition: null,
    isRecording: false
  };

  const tempBadgeClass = {
    Hot: {
      dot: 'bg-red-500',
      text: 'text-red-600'
    },
    Warm: {
      dot: 'bg-orange-500',
      text: 'text-orange-600'
    },
    Cold: {
      dot: 'bg-blue-500',
      text: 'text-blue-600'
    }
  };

  const views = {
    login: null,
    dashboard: null
  };

  const els = {};

  document.addEventListener('DOMContentLoaded', init);

  function init() {
    cacheDom();
    bindEvents();
    showView('login');
    setLoader(false);
    initSelectDefaults();
    updateDPProgress();
  }

  function cacheDom() {
    views.login = byId('viewLogin');
    views.dashboard = byId('viewDashboard');

    els.toast = byId('toast');
    els.globalLoader = byId('globalLoader');
    els.globalLoaderText = byId('globalLoaderText');

    els.loginForm = byId('loginForm');
    els.loginUsername = byId('loginUsername');
    els.loginPassword = byId('loginPassword');
    els.loginSubmitBtn = byId('loginSubmitBtn');

    els.welcomeText = byId('welcomeText');
    els.openMenuBtn = byId('openMenuBtn');
    els.menuPanel = byId('menuPanel');
    els.manualAddBtn = byId('manualAddBtn');
    els.refreshBtn = byId('refreshBtn');
    els.pipelineBoard = byId('pipelineBoard');
    els.voiceFab = byId('voiceFab');

    els.leadModal = byId('leadModal');
    els.leadName = byId('leadName');
    els.leadMeta = byId('leadMeta');
    els.callBtn = byId('callBtn');
    els.whatsAppBtn = byId('whatsAppBtn');
    els.viberBtn = byId('viberBtn');
    els.leadStatusSelect = byId('leadStatusSelect');
    els.saveStatusBtn = byId('saveStatusBtn');
    els.manageDpBtn = byId('manageDpBtn');
    els.activityTimeline = byId('activityTimeline');
    els.activityInput = byId('activityInput');
    els.addActivityBtn = byId('addActivityBtn');

    els.voiceModal = byId('voiceModal');
    els.voiceForm = byId('voiceForm');
    els.vfClientName = byId('vfClientName');
    els.vfPhone = byId('vfPhone');
    els.vfCountry = byId('vfCountry');
    els.vfClassification = byId('vfClassification');
    els.vfSource = byId('vfSource');
    els.vfTemperature = byId('vfTemperature');
    els.vfProject = byId('vfProject');
    els.vfPrice = byId('vfPrice');
    els.vfRemarks = byId('vfRemarks');
    els.saveLeadBtn = byId('saveLeadBtn');

    els.dpModal = byId('dpModal');
    els.dpForm = byId('dpForm');
    els.dpTermMonths = byId('dpTermMonths');
    els.dpCurrentMonth = byId('dpCurrentMonth');
    els.dpReservationDate = byId('dpReservationDate');
    els.receiptFileInput = byId('receiptFileInput');
    els.uploadSpinner = byId('uploadSpinner');
    els.receiptLabel = byId('receiptLabel');
    els.dpProgressText = byId('dpProgressText');
    els.dpProgressBar = byId('dpProgressBar');

    els.profileName = byId('profileName');
    els.profileUser = byId('profileUser');
    els.profileTeamRole = byId('profileTeamRole');
    els.changePasswordForm = byId('changePasswordForm');
    els.cpCurrent = byId('cpCurrent');
    els.cpNew = byId('cpNew');
    els.ticketForm = byId('ticketForm');
    els.ticketIssueType = byId('ticketIssueType');
    els.ticketPriority = byId('ticketPriority');
    els.ticketDescription = byId('ticketDescription');
    els.logoutBtn = byId('logoutBtn');

    document.querySelectorAll('[data-close]').forEach((el) => {
      el.addEventListener('click', () => hideModal(el.getAttribute('data-close')));
    });
  }

  function bindEvents() {
    els.loginForm.addEventListener('submit', onLoginSubmit);
    els.refreshBtn.addEventListener('click', refreshWorkspace);
    els.manualAddBtn.addEventListener('click', onManualAddLead);
    els.voiceFab.addEventListener('click', onVoiceCapture);

    els.openMenuBtn.addEventListener('click', () => showModal('menuPanel'));
    els.logoutBtn.addEventListener('click', onLogout);

    els.saveStatusBtn.addEventListener('click', onSaveLeadStatus);
    els.addActivityBtn.addEventListener('click', onAddActivity);
    els.manageDpBtn.addEventListener('click', onOpenDPModal);

    els.voiceForm.addEventListener('submit', onSaveVerifiedLead);

    els.dpForm.addEventListener('submit', onSaveDP);
    els.dpTermMonths.addEventListener('input', updateDPProgress);
    els.dpCurrentMonth.addEventListener('input', updateDPProgress);
    els.receiptFileInput.addEventListener('change', onReceiptSelected);

    els.changePasswordForm.addEventListener('submit', onChangePassword);
    els.ticketForm.addEventListener('submit', onSubmitTicket);
  }

  async function onLoginSubmit(evt) {
    evt.preventDefault();

    const username = els.loginUsername.value.trim();
    const password = els.loginPassword.value.trim();

    if (!username || !password) {
      toast('Username and password are required.', 'error');
      return;
    }

    try {
      setButtonLoading(els.loginSubmitBtn, true, 'Signing in...');
      const result = await gas('authenticateUser', username, password);

      state.token = result.token;
      state.user = result.user;

      await refreshWorkspace();
      await refreshProfile();

      showView('dashboard');
      toast('Signed in successfully.', 'success');
      els.loginPassword.value = '';
    } catch (err) {
      toast(getErrorMessage(err), 'error');
    } finally {
      setButtonLoading(els.loginSubmitBtn, false, 'Sign In');
    }
  }

  async function refreshWorkspace() {
    if (!state.token) {
      return;
    }

    try {
      setLoader(true, 'Syncing dashboard...');
      const payload = await gas('getAgentWorkspace', state.token);

      state.user = payload.user || state.user;
      state.statuses = Array.isArray(payload.statuses) && payload.statuses.length ? payload.statuses : state.statuses;
      state.options = payload.options || state.options;
      state.projects = Array.isArray(payload.projects) ? payload.projects : [];
      state.leadsByStatus = payload.leadsByStatus || {};

      els.welcomeText.textContent = 'Welcome, ' + (state.user && state.user.fullName ? state.user.fullName : 'Agent');

      renderStaticOptions();
      renderPipelineBoard();
    } catch (err) {
      toast(getErrorMessage(err), 'error');
    } finally {
      setLoader(false);
    }
  }

  function renderStaticOptions() {
    renderSelectOptions(els.vfClassification, state.options.classifications || [], '', true);
    renderSelectOptions(els.vfSource, state.options.sources || [], '', true);
    renderSelectOptions(els.vfTemperature, state.options.temperatures || [], '', true);

    const projectNames = state.projects.map((p) => p.projectName).filter(Boolean);
    renderSelectOptions(els.vfProject, projectNames, '', true);

    renderSelectOptions(els.leadStatusSelect, state.statuses, 'Inquiry', false);
  }

  function renderPipelineBoard() {
    const fragment = document.createDocumentFragment();

    state.statuses.forEach((status) => {
      const leads = Array.isArray(state.leadsByStatus[status]) ? state.leadsByStatus[status] : [];

      const col = document.createElement('section');
      col.className = 'min-w-[260px] max-w-[260px] rounded-2xl border border-slate-200 bg-white p-3 shadow-sm';

      const header = document.createElement('div');
      header.className = 'mb-3 flex items-center justify-between';
      header.innerHTML =
        '<h4 class="text-sm font-extrabold text-slate-800">' +
        escapeHtml(status) +
        '</h4><span class="inline-flex h-6 min-w-6 items-center justify-center rounded-full bg-slate-100 px-2 text-xs font-bold text-slate-700">' +
        leads.length +
        '</span>';

      col.appendChild(header);

      const body = document.createElement('div');
      body.className = 'space-y-2';

      if (!leads.length) {
        const empty = document.createElement('p');
        empty.className = 'rounded-xl border border-dashed border-slate-200 px-3 py-4 text-center text-xs text-slate-400';
        empty.textContent = 'No leads yet';
        body.appendChild(empty);
      } else {
        leads.forEach((lead) => body.appendChild(renderLeadCard(lead)));
      }

      col.appendChild(body);
      fragment.appendChild(col);
    });

    els.pipelineBoard.innerHTML = '';
    els.pipelineBoard.appendChild(fragment);
  }

  function renderLeadCard(lead) {
    const badge = tempBadgeClass[lead.temperature] || tempBadgeClass.Warm;

    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'w-full rounded-xl border border-slate-200 bg-slate-50 p-3 text-left transition hover:bg-slate-100 active:scale-[0.99]';

    btn.innerHTML =
      '<div class="flex items-start justify-between gap-2">' +
      '<h5 class="line-clamp-1 text-sm font-bold text-slate-800">' +
      escapeHtml(lead.clientName || '-') +
      '</h5>' +
      '<span class="inline-flex items-center gap-1 text-xs font-bold ' +
      badge.text +
      '"><span class="h-2.5 w-2.5 rounded-full ' +
      badge.dot +
      '"></span>' +
      escapeHtml(lead.temperature || 'Warm') +
      '</span>' +
      '</div>' +
      '<p class="mt-1 text-xs text-slate-500">' +
      escapeHtml(lead.phone || 'No phone') +
      '</p>' +
      '<p class="mt-1 line-clamp-1 text-xs text-slate-500">' +
      escapeHtml(lead.project || 'No project') +
      '</p>';

    btn.addEventListener('click', () => openLead(lead.leadId));
    return btn;
  }

  async function openLead(leadId) {
    try {
      setLoader(true, 'Loading lead details...');
      const payload = await gas('getLeadDetails', state.token, leadId);

      state.activeLead = payload.lead || null;
      state.activeTransaction = payload.transaction || null;

      if (!state.activeLead) {
        toast('Lead not found.', 'error');
        return;
      }

      fillLeadModal(payload.lead, payload.activities || [], payload.transaction || null);
      showModal('leadModal');
    } catch (err) {
      toast(getErrorMessage(err), 'error');
    } finally {
      setLoader(false);
    }
  }

  function fillLeadModal(lead, activities, transaction) {
    els.leadName.textContent = lead.clientName || 'Unnamed Lead';
    els.leadMeta.textContent = [lead.project || 'No project', lead.country || 'No country'].filter(Boolean).join(' | ');

    const normalizedPhone = normalizePhoneForLinks(lead.phone);
    els.callBtn.href = normalizedPhone ? 'tel:' + normalizedPhone : '#';
    els.whatsAppBtn.href = normalizedPhone ? 'https://wa.me/' + normalizedPhone.replace(/\+/g, '') : '#';
    els.viberBtn.href = normalizedPhone ? 'viber://chat?number=' + encodeURIComponent(normalizedPhone) : '#';

    renderSelectOptions(els.leadStatusSelect, state.statuses, lead.status || 'Inquiry', false);

    if ((lead.status || '') === 'Down Payment') {
      els.manageDpBtn.classList.remove('hidden');
    } else {
      els.manageDpBtn.classList.add('hidden');
    }

    if (transaction && transaction.progressPercent >= 0) {
      els.manageDpBtn.textContent = 'Manage DP (' + transaction.progressPercent + '% complete)';
    } else {
      els.manageDpBtn.textContent = 'Manage DP';
    }

    renderTimeline(activities);
  }

  function renderTimeline(items) {
    els.activityTimeline.innerHTML = '';

    if (!items.length) {
      const empty = document.createElement('p');
      empty.className = 'rounded-xl border border-dashed border-slate-200 px-3 py-4 text-center text-xs text-slate-400';
      empty.textContent = 'No activity yet';
      els.activityTimeline.appendChild(empty);
      return;
    }

    const fragment = document.createDocumentFragment();
    items.forEach((item) => {
      const row = document.createElement('div');
      row.className = 'rounded-xl border border-slate-200 p-3';
      row.innerHTML =
        '<p class="text-sm font-semibold text-slate-700">' +
        escapeHtml(item.actionText || '') +
        '</p>' +
        '<p class="mt-1 text-[11px] text-slate-500">' +
        escapeHtml(formatDate(item.timestamp)) +
        ' | ' +
        escapeHtml(item.agentUsername || '') +
        '</p>';
      fragment.appendChild(row);
    });

    els.activityTimeline.appendChild(fragment);
  }

  async function onSaveLeadStatus() {
    if (!state.activeLead) {
      return;
    }

    const status = els.leadStatusSelect.value;

    try {
      setButtonLoading(els.saveStatusBtn, true, 'Saving...');
      await gas('updateLeadStatus', state.token, state.activeLead.leadId, status);

      state.activeLead.status = status;
      if (status === 'Down Payment') {
        els.manageDpBtn.classList.remove('hidden');
      } else {
        els.manageDpBtn.classList.add('hidden');
      }

      toast('Lead status updated.', 'success');
      await refreshWorkspace();
      await openLead(state.activeLead.leadId);
    } catch (err) {
      toast(getErrorMessage(err), 'error');
    } finally {
      setButtonLoading(els.saveStatusBtn, false, 'Save');
    }
  }

  async function onAddActivity() {
    if (!state.activeLead) {
      return;
    }

    const text = els.activityInput.value.trim();
    if (!text) {
      toast('Activity note is empty.', 'error');
      return;
    }

    try {
      setButtonLoading(els.addActivityBtn, true, 'Adding...');
      const res = await gas('addActivityLog', state.token, state.activeLead.leadId, text);
      els.activityInput.value = '';
      renderTimeline(res.activities || []);
      toast('Activity added.', 'success');
    } catch (err) {
      toast(getErrorMessage(err), 'error');
    } finally {
      setButtonLoading(els.addActivityBtn, false, 'Add');
    }
  }

  function onManualAddLead() {
    state.voiceMode = 'manual';
    state.transcript = '';
    fillVoiceForm({
      client_name: '',
      phone_number: '',
      country: 'Philippines',
      classification: 'Unknown',
      source: 'Organic',
      temperature: 'Warm',
      project_interest: '',
      total_selling_price: '',
      remarks: ''
    });
    showModal('voiceModal');
  }

  async function onVoiceCapture() {
    if (!window.webkitSpeechRecognition) {
      toast('Voice recognition is not supported on this browser.', 'error');
      return;
    }

    if (state.isRecording) {
      try {
        state.recognition.stop();
      } catch (err) {
        console.warn(err);
      }
      return;
    }

    state.voiceMode = 'voice';
    state.transcript = '';

    const recognition = new window.webkitSpeechRecognition();
    recognition.lang = 'en-PH';
    recognition.interimResults = false;
    recognition.maxAlternatives = 1;

    state.recognition = recognition;
    state.isRecording = true;
    els.voiceFab.classList.add('bg-red-500');

    toast('Listening... speak now.', 'info');

    recognition.onresult = async (event) => {
      const transcript =
        event && event.results && event.results[0] && event.results[0][0]
          ? String(event.results[0][0].transcript || '').trim()
          : '';

      if (!transcript) {
        toast('No transcript captured. Please try again.', 'error');
        return;
      }

      state.transcript = transcript;

      try {
        setLoader(true, 'Analyzing voice note...');
        const aiData = await gas('processVoiceTranscript', state.token, transcript);
        fillVoiceForm(aiData);
        showModal('voiceModal');
        toast('Voice note parsed. Verify details before saving.', 'success');
      } catch (err) {
        toast(getErrorMessage(err), 'error');
      } finally {
        setLoader(false);
      }
    };

    recognition.onerror = () => {
      toast('Voice recognition failed. Please try again.', 'error');
    };

    recognition.onend = () => {
      state.isRecording = false;
      els.voiceFab.classList.remove('bg-red-500');
    };

    recognition.start();
  }

  function fillVoiceForm(data) {
    els.vfClientName.value = data.client_name || '';
    els.vfPhone.value = data.phone_number || '';
    els.vfCountry.value = data.country || 'Philippines';
    setSelectValue(els.vfClassification, data.classification || 'Unknown');
    setSelectValue(els.vfSource, data.source || 'Organic');
    setSelectValue(els.vfTemperature, data.temperature || 'Warm');
    setSelectValue(els.vfProject, data.project_interest || '');
    els.vfPrice.value = data.total_selling_price || '';
    els.vfRemarks.value = data.remarks || state.transcript || '';
  }

  async function onSaveVerifiedLead(evt) {
    evt.preventDefault();

    const payload = {
      clientName: els.vfClientName.value.trim(),
      phone: els.vfPhone.value.trim(),
      country: els.vfCountry.value.trim() || 'Philippines',
      classification: els.vfClassification.value,
      source: els.vfSource.value,
      temperature: els.vfTemperature.value,
      project: els.vfProject.value,
      totalSellingPrice: els.vfPrice.value.trim(),
      remarks: els.vfRemarks.value.trim() || state.transcript,
      status: 'Inquiry',
      entrySource: state.voiceMode === 'voice' ? 'Voice' : 'Manual'
    };

    if (!payload.clientName) {
      toast('Client name is required.', 'error');
      return;
    }

    try {
      setButtonLoading(els.saveLeadBtn, true, 'Saving lead...');
      await gas('saveLead', state.token, payload);
      hideModal('voiceModal');
      await refreshWorkspace();
      toast('Lead saved successfully.', 'success');
    } catch (err) {
      toast(getErrorMessage(err), 'error');
    } finally {
      setButtonLoading(els.saveLeadBtn, false, 'Verify & Save Lead');
    }
  }

  async function onOpenDPModal() {
    if (!state.activeLead) {
      return;
    }

    try {
      setLoader(true, 'Loading transaction details...');
      const payload = await gas('getDPTransaction', state.token, state.activeLead.leadId);
      const tx = payload && payload.transaction ? payload.transaction : null;
      state.activeTransaction = tx;

      applyDPTransactionToForm(tx);
      showModal('dpModal');
    } catch (err) {
      toast(getErrorMessage(err), 'error');
    } finally {
      setLoader(false);
    }
  }

  function applyDPTransactionToForm(tx) {
    const paymentType = tx && tx.paymentType ? tx.paymentType : 'Spot';
    const radios = els.dpForm.querySelectorAll('input[name="paymentType"]');
    radios.forEach((r) => {
      r.checked = r.value === paymentType;
    });

    els.dpTermMonths.value = tx && tx.termMonths ? tx.termMonths : paymentType === 'Spot' ? 1 : 12;
    els.dpCurrentMonth.value = tx && tx.currentMonthPaid >= 0 ? tx.currentMonthPaid : paymentType === 'Spot' ? 1 : 0;
    els.dpReservationDate.value = tx && tx.reservationDate ? tx.reservationDate.slice(0, 10) : todayISO();

    els.receiptFileInput.value = '';
    els.receiptLabel.textContent = tx && tx.receiptDriveURL ? 'Latest: Receipt already uploaded' : 'No file selected';

    updateDPProgress();
  }

  function onReceiptSelected() {
    const file = els.receiptFileInput.files && els.receiptFileInput.files[0];
    els.receiptLabel.textContent = file ? file.name : 'No file selected';
  }

  async function onSaveDP(evt) {
    evt.preventDefault();

    if (!state.activeLead) {
      toast('No active lead selected.', 'error');
      return;
    }

    const selectedRadio = els.dpForm.querySelector('input[name="paymentType"]:checked');
    const paymentType = selectedRadio ? selectedRadio.value : 'Spot';

    let receiptDriveURL = state.activeTransaction && state.activeTransaction.receiptDriveURL ? state.activeTransaction.receiptDriveURL : '';

    try {
      setLoader(true, 'Saving down payment details...');

      const file = els.receiptFileInput.files && els.receiptFileInput.files[0];
      if (file) {
        toggleUploadSpinner(true);
        const base64 = await fileToDataURL(file);
        const uploadRes = await gas('uploadReceipt', base64, file.name, file.type, state.activeLead.leadId, state.token);
        receiptDriveURL = uploadRes.url || '';
        toggleUploadSpinner(false);
      }

      const payload = {
        leadId: state.activeLead.leadId,
        paymentType,
        termMonths: Number(els.dpTermMonths.value || 1),
        currentMonthPaid: Number(els.dpCurrentMonth.value || 0),
        reservationDate: els.dpReservationDate.value || todayISO(),
        receiptDriveURL
      };

      const res = await gas('saveDPTransaction', state.token, payload);
      state.activeTransaction = res.transaction || null;

      hideModal('dpModal');
      toast('Down payment details saved.', 'success');

      await refreshWorkspace();
      await openLead(state.activeLead.leadId);
    } catch (err) {
      toast(getErrorMessage(err), 'error');
    } finally {
      toggleUploadSpinner(false);
      setLoader(false);
    }
  }

  function updateDPProgress() {
    const term = Math.max(1, Number(els.dpTermMonths.value || 1));
    const paid = Math.max(0, Math.min(term, Number(els.dpCurrentMonth.value || 0)));
    const percent = Math.round((paid / term) * 100);

    els.dpProgressText.textContent = percent + '%';
    els.dpProgressBar.style.width = percent + '%';
  }

  async function refreshProfile() {
    if (!state.token) {
      return;
    }

    try {
      const profile = await gas('getProfile', state.token);
      els.profileName.textContent = profile.fullName || '-';
      els.profileUser.textContent = 'Username: ' + (profile.username || '-');
      els.profileTeamRole.textContent = 'Team: ' + (profile.team || '-') + ' | Role: ' + (profile.role || '-');
    } catch (err) {
      toast(getErrorMessage(err), 'error');
    }
  }

  async function onChangePassword(evt) {
    evt.preventDefault();

    const currentPassword = els.cpCurrent.value;
    const newPassword = els.cpNew.value;

    if (!currentPassword || !newPassword) {
      toast('Both password fields are required.', 'error');
      return;
    }

    try {
      const btn = els.changePasswordForm.querySelector('button[type="submit"]');
      setButtonLoading(btn, true, 'Updating...');
      await gas('changePassword', state.token, currentPassword, newPassword);
      els.changePasswordForm.reset();
      toast('Password updated.', 'success');
    } catch (err) {
      toast(getErrorMessage(err), 'error');
    } finally {
      const btn = els.changePasswordForm.querySelector('button[type="submit"]');
      setButtonLoading(btn, false, 'Update Password');
    }
  }

  async function onSubmitTicket(evt) {
    evt.preventDefault();

    const issueType = els.ticketIssueType.value;
    const priority = els.ticketPriority.value;
    const description = els.ticketDescription.value.trim();

    if (!issueType || !priority || !description) {
      toast('Issue type, priority, and description are required.', 'error');
      return;
    }

    try {
      const btn = els.ticketForm.querySelector('button[type="submit"]');
      setButtonLoading(btn, true, 'Submitting...');
      await gas('submitTicket', state.token, issueType, priority, description);
      els.ticketForm.reset();
      toast('Ticket submitted to support.', 'success');
    } catch (err) {
      toast(getErrorMessage(err), 'error');
    } finally {
      const btn = els.ticketForm.querySelector('button[type="submit"]');
      setButtonLoading(btn, false, 'Submit Ticket');
    }
  }

  async function onLogout() {
    try {
      if (state.token) {
        await gas('logout', state.token);
      }
    } catch (err) {
      console.warn(err);
    }

    state.token = '';
    state.user = null;
    state.activeLead = null;
    state.activeTransaction = null;
    state.leadsByStatus = {};

    hideModal('menuPanel');
    hideModal('leadModal');
    hideModal('voiceModal');
    hideModal('dpModal');

    els.loginForm.reset();
    showView('login');
    toast('Signed out.', 'info');
  }

  function showView(name) {
    if (name === 'dashboard') {
      views.login.classList.add('hidden');
      views.dashboard.classList.remove('hidden');
      return;
    }

    views.dashboard.classList.add('hidden');
    views.login.classList.remove('hidden');
  }

  function showModal(id) {
    const node = byId(id);
    if (!node) {
      return;
    }
    node.classList.remove('hidden');
    document.body.classList.add('overflow-hidden');
  }

  function hideModal(id) {
    const node = byId(id);
    if (!node) {
      return;
    }
    node.classList.add('hidden');

    const open = ['leadModal', 'voiceModal', 'dpModal', 'menuPanel'].some((modalId) => {
      const modal = byId(modalId);
      return modal && !modal.classList.contains('hidden');
    });

    if (!open) {
      document.body.classList.remove('overflow-hidden');
    }
  }

  function setLoader(show, text) {
    if (!els.globalLoader) {
      return;
    }

    if (show) {
      els.globalLoader.classList.remove('hidden');
      els.globalLoader.classList.add('flex');
      els.globalLoaderText.textContent = text || 'Working...';
    } else {
      els.globalLoader.classList.add('hidden');
      els.globalLoader.classList.remove('flex');
      els.globalLoaderText.textContent = 'Working...';
    }
  }

  function toggleUploadSpinner(show) {
    if (show) {
      els.uploadSpinner.classList.remove('hidden');
      els.uploadSpinner.classList.add('flex');
      return;
    }

    els.uploadSpinner.classList.add('hidden');
    els.uploadSpinner.classList.remove('flex');
  }

  function toast(message, type) {
    const kind = type || 'info';
    const classes = {
      success: 'bg-emerald-600',
      error: 'bg-red-600',
      info: 'bg-brandBlue'
    };

    els.toast.className =
      'pointer-events-none fixed left-1/2 top-4 z-[100] w-[92%] max-w-sm -translate-x-1/2 rounded-xl px-4 py-3 text-sm font-semibold text-white shadow-lg ' +
      (classes[kind] || classes.info);
    els.toast.textContent = message;
    els.toast.classList.remove('hidden');

    window.clearTimeout(els.toast._timer);
    els.toast._timer = window.setTimeout(() => {
      els.toast.classList.add('hidden');
    }, 2200);
  }

  function setButtonLoading(button, loading, loadingLabel) {
    if (!button) {
      return;
    }

    if (loading) {
      button.dataset.originalText = button.textContent;
      button.textContent = loadingLabel || 'Working...';
      button.disabled = true;
      button.classList.add('opacity-70');
      return;
    }

    button.textContent = button.dataset.originalText || button.textContent;
    button.disabled = false;
    button.classList.remove('opacity-70');
  }

  function renderSelectOptions(select, values, selectedValue, includeBlank) {
    if (!select) {
      return;
    }

    const html = [];
    if (includeBlank) {
      html.push('<option value="">Select</option>');
    }

    values.forEach((value) => {
      const selected = selectedValue && String(value) === String(selectedValue) ? ' selected' : '';
      html.push('<option value="' + escapeAttr(value) + '"' + selected + '>' + escapeHtml(value) + '</option>');
    });

    select.innerHTML = html.join('');

    if (selectedValue) {
      setSelectValue(select, selectedValue);
    }
  }

  function setSelectValue(select, value) {
    if (!select) {
      return;
    }

    const val = String(value || '');
    const exists = Array.from(select.options).some((opt) => opt.value === val);

    if (exists) {
      select.value = val;
      return;
    }

    select.value = '';
  }

  function initSelectDefaults() {
    renderStaticOptions();
  }

  function byId(id) {
    return document.getElementById(id);
  }

  function formatDate(iso) {
    if (!iso) {
      return '-';
    }
    const date = new Date(iso);
    if (Number.isNaN(date.getTime())) {
      return '-';
    }

    return date.toLocaleString('en-PH', {
      year: 'numeric',
      month: 'short',
      day: 'numeric',
      hour: '2-digit',
      minute: '2-digit'
    });
  }

  function todayISO() {
    const now = new Date();
    const y = now.getFullYear();
    const m = String(now.getMonth() + 1).padStart(2, '0');
    const d = String(now.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + d;
  }

  function normalizePhoneForLinks(phone) {
    return String(phone || '').replace(/[^0-9+]/g, '');
  }

  function fileToDataURL(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result || '');
      reader.onerror = () => reject(new Error('Failed to read selected file.'));
      reader.readAsDataURL(file);
    });
  }

  function getErrorMessage(err) {
    if (!err) {
      return 'Something went wrong.';
    }

    if (typeof err === 'string') {
      return err;
    }

    if (err.message) {
      return err.message;
    }

    if (err.details) {
      return String(err.details);
    }

    return 'Something went wrong.';
  }

  function gas(functionName) {
    const args = Array.prototype.slice.call(arguments, 1);

    return new Promise((resolve, reject) => {
      if (!window.google || !google.script || !google.script.run) {
        reject(new Error('google.script.run is unavailable. Open this app from the deployed Apps Script web app URL.'));
        return;
      }

      let runner = google.script.run
        .withSuccessHandler((data) => resolve(data))
        .withFailureHandler((err) => reject(err));

      if (typeof runner[functionName] !== 'function') {
        reject(new Error('Backend function not found: ' + functionName));
        return;
      }

      runner[functionName].apply(runner, args);
    });
  }

  function escapeHtml(input) {
    return String(input || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function escapeAttr(input) {
    return escapeHtml(input).replace(/`/g, '&#96;');
  }
})();
