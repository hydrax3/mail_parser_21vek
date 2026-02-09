/**
 * Email Notifications - Renderer Process
 * Handles UI logic, data fetching, filtering, and notifications
 */

// State
let emails = [];
let filteredEmails = [];
let currentFilter = 'all';
let searchQuery = '';
let previousOverdueIds = new Set();
let isFirstLoad = true;

// Pagination State
let currentPage = 1;
const itemsPerPage = 10;

// Constants
const REFRESH_INTERVAL = 30000; // 30 seconds
const OVERDUE_HOURS = 6;
const REMINDER_HOURS = 3; // Hours after external reply to trigger reminder

// Reminder system state
let operators = [];
let notifiedReminderIds = new Set();

// DOM Elements
const elements = {
    emailTableBody: document.getElementById('emailTableBody'),
    loadingState: document.getElementById('loadingState'),
    emptyState: document.getElementById('emptyState'),
    errorState: document.getElementById('errorState'),
    errorMessage: document.getElementById('errorMessage'),
    searchInput: document.getElementById('searchInput'),
    refreshBtn: document.getElementById('refreshBtn'),
    retryBtn: document.getElementById('retryBtn'),
    totalCount: document.getElementById('totalCount'),
    overdueCount: document.getElementById('overdueCount'),
    answeredCount: document.getElementById('answeredCount'),
    lastUpdate: document.getElementById('lastUpdate'),
    statusIndicator: document.getElementById('statusIndicator'),
    filterButtons: document.querySelectorAll('.filter-btn'),

    // Modal elements
    subjectModal: document.getElementById('subjectModal'),
    modalCloseBtn: document.getElementById('modalCloseBtn'),
    modalSubjectText: document.getElementById('modalSubjectText'),

    // Pagination elements
    paginationContainer: document.getElementById('paginationContainer')
};

// ========================================
// Initialization
// ========================================
document.addEventListener('DOMContentLoaded', () => {
    initTitlebarButtons();
    initEventListeners();
    loadOperators(); // Load operators for reminder system
    loadEmails();
    startAutoRefresh();
});

function initTitlebarButtons() {
    // Window control buttons
    document.getElementById('minimizeBtn')?.addEventListener('click', () => {
        window.api.windowMinimize();
    });

    document.getElementById('maximizeBtn')?.addEventListener('click', () => {
        window.api.windowMaximize();
    });

    document.getElementById('closeBtn')?.addEventListener('click', () => {
        window.api.windowClose();
    });
}

function initEventListeners() {
    // Search input
    elements.searchInput.addEventListener('input', (e) => {
        searchQuery = e.target.value.toLowerCase();
        currentPage = 1; // Reset to first page on search
        applyFilters();
        renderTable();
    });

    // Refresh button
    elements.refreshBtn.addEventListener('click', () => {
        loadEmails();
    });

    // Retry button
    elements.retryBtn.addEventListener('click', () => {
        loadEmails();
    });

    // Filter buttons
    elements.filterButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            elements.filterButtons.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            currentFilter = btn.dataset.filter;
            currentPage = 1; // Reset to first page on filter change
            applyFilters();
            renderTable();
        });
    });

    // Modal close button
    elements.modalCloseBtn?.addEventListener('click', () => {
        closeSubjectModal();
    });

    // Close modal when clicking outside
    window.addEventListener('click', (event) => {
        if (event.target === elements.subjectModal) {
            closeSubjectModal();
        }
    });

    initAdminPanel();
    initDelegation();
}

function initDelegation() {
    // Pagination delegation
    if (elements.paginationContainer) {
        elements.paginationContainer.addEventListener('click', (e) => {
            const btn = e.target.closest('.pagination-btn');
            if (btn && !btn.disabled) {
                const page = parseInt(btn.dataset.page);
                if (!isNaN(page)) {
                    changePage(page);
                }
            }
        });
    }

    // Subject modal delegation
    if (elements.emailTableBody) {
        elements.emailTableBody.addEventListener('click', (e) => {
            const subjectEl = e.target.closest('.subject-text');
            if (subjectEl) {
                const subject = subjectEl.dataset.subject;
                if (subject) {
                    openSubjectModal(subject);
                }
            }

            // Close reminder button delegation
            const closeBtn = e.target.closest('.btn-close-reminder');
            if (closeBtn) {
                const rowIndex = parseInt(closeBtn.dataset.rowIndex);
                if (!isNaN(rowIndex)) {
                    closeReminder(rowIndex);
                }
            }
        });
    }
}

// ========================================
// Data Loading
// ========================================
async function loadEmails() {
    showLoading(true);
    setRefreshButtonLoading(true);

    try {
        const result = await window.api.getEmails();

        if (result.success) {
            emails = result.data.map(processEmail);
            applyFilters();
            renderTable();
            updateStats();
            updateLastRefreshTime();
            checkForNewOverdueEmails();
            checkForAwaitingReplyEmails();
            setConnectionStatus(true);
            isFirstLoad = false;
        } else {
            throw new Error(result.error || 'Unknown error');
        }
    } catch (error) {
        console.error('Failed to load emails:', error);
        showError(error.message);
        setConnectionStatus(false);
    } finally {
        showLoading(false);
        setRefreshButtonLoading(false);
    }
}

function processEmail(email) {
    // Determine effective date for calculation
    // Priority: reminderStatus (Col H) > time (Col D)
    // This allows correctly calculating time since LAST activity (e.g. last client reply)
    let effectiveDateStr = email.time;

    if (email.reminderStatus && email.reminderStatus.toLowerCase() !== 'закрыт') {
        const d = parseDate(email.reminderStatus);
        if (d && !isNaN(d.getTime())) {
            effectiveDateStr = email.reminderStatus;
        }
    }

    // Calculate overdue status
    const emailDate = parseDate(effectiveDateStr);
    const now = new Date();
    const hoursDiff = emailDate ? (now - emailDate) / (1000 * 60 * 60) : 0;

    // Use backend logic if availble, or fallback
    // Note: notificationCheck is provided by main.js
    const isOverdue = email.notificationCheck ? email.notificationCheck.notify : false;
    const reason = email.notificationCheck ? email.notificationCheck.reason : '';

    return {
        ...email,
        parsedDate: emailDate, // This is now effective date
        hoursSinceReceived: hoursDiff,
        isOverdue: isOverdue,
        overdueReason: reason,
        hoursOverdue: Math.max(0, hoursDiff - OVERDUE_HOURS)
    };
}

function parseDate(dateStr) {
    if (!dateStr) return null;
    try {
        const d = new Date(dateStr);
        if (isNaN(d.getTime())) return null;
        return d;
    } catch (e) {
        return null;
    }
}

// ========================================
// Filtering
// ========================================
function applyFilters() {
    filteredEmails = emails.filter(email => {
        // Search filter
        if (searchQuery) {
            const searchFields = [
                email.subject,
                email.sender,
                email.status
            ].map(f => (f || '').toLowerCase());

            if (!searchFields.some(f => f.includes(searchQuery))) {
                return false;
            }
        }

        // Status filter
        switch (currentFilter) {
            case 'overdue':
                return email.isOverdue;
            case 'pending':
                // Pending means monitored statuses but not yet overdue
                const monitored = ['ответа нет', 'ответ не от оператора'];
                return monitored.includes((email.status || '').toLowerCase()) && !email.isOverdue;
            case 'answered':
                return email.status === 'отвечено' || email.status === 'получен ответ';
            case 'awaiting':
                return isAwaitingReply(email);
            default:
                return true;
        }
    });
}

// ========================================
// Rendering
// ========================================
function renderTable() {
    if (filteredEmails.length === 0) {
        elements.emailTableBody.innerHTML = '';
        elements.paginationContainer.innerHTML = '';
        showEmpty(true);
        return;
    }

    showEmpty(false);

    // Sort by date (newest first)
    const sortedEmails = [...filteredEmails].sort((a, b) => {
        if (!a.parsedDate) return 1;
        if (!b.parsedDate) return -1;
        return b.parsedDate - a.parsedDate;
    });

    // Pagination
    const totalPages = Math.ceil(sortedEmails.length / itemsPerPage);
    if (currentPage > totalPages) currentPage = totalPages || 1;

    const startIndex = (currentPage - 1) * itemsPerPage;
    const paginatedEmails = sortedEmails.slice(startIndex, startIndex + itemsPerPage);

    elements.emailTableBody.innerHTML = paginatedEmails.map(email => createEmailRow(email)).join('');
    renderPagination(totalPages);
}

function renderPagination(totalPages) {
    if (totalPages <= 1) {
        elements.paginationContainer.innerHTML = '';
        return;
    }

    let html = `
        <div class="pagination">
            <button class="pagination-btn" ${currentPage === 1 ? 'disabled' : ''} data-page="${currentPage - 1}">
                &laquo;
            </button>
    `;

    for (let i = 1; i <= totalPages; i++) {
        html += `
            <button class="pagination-btn ${i === currentPage ? 'active' : ''}" data-page="${i}">
                ${i}
            </button>
        `;
    }

    html += `
            <button class="pagination-btn" ${currentPage === totalPages ? 'disabled' : ''} data-page="${currentPage + 1}">
                &raquo;
            </button>
        </div>
    `;

    elements.paginationContainer.innerHTML = html;
}

window.changePage = function (page) {
    if (page < 1) return;
    currentPage = page;
    renderTable();
};

function createEmailRow(email) {
    const statusBadge = getStatusBadge(email);
    const overdueIndicator = getOverdueIndicator(email);
    const rowClass = email.isOverdue ? 'row-overdue' : (isAwaitingReply(email) ? 'row-awaiting' : '');
    const closeButton = isAwaitingReply(email)
        ? `<button class="btn-close-reminder" data-row-index="${email.rowIndex}" title="Закрыть напоминание">✓ Закрыть</button>`
        : '';

    return `
        <tr class="${rowClass}">
            <td>${statusBadge}</td>
            <td><span class="sender-name" title="${escapeHtml(email.sender)}">${escapeHtml(extractName(email.sender))}</span></td>
            <td>
                <span class="subject-text" 
                      data-subject="${escapeHtml(email.subject)}"
                      title="Кликните для просмотра полной темы">
                    ${escapeHtml(email.subject)}
                </span>
            </td>
            <td><span class="time-text">${formatDate(email.parsedDate)}</span></td>
            <td>${overdueIndicator}</td>
            <td>${closeButton}</td>
        </tr>
    `;
}

function openSubjectModal(subject) {
    if (!elements.subjectModal) return;
    elements.modalSubjectText.textContent = subject;
    elements.subjectModal.style.display = 'flex';
}

function closeSubjectModal() {
    if (!elements.subjectModal) return;
    elements.subjectModal.style.display = 'none';
}

function escapeJs(str) {
    if (!str) return '';
    return str
        .replace(/\\/g, '\\\\')
        .replace(/'/g, "\\'")
        .replace(/"/g, '\\"')
        .replace(/`/g, '\\`')
        .replace(/\n/g, '\\n')
        .replace(/\r/g, '\\r');
}

function getStatusBadge(email) {
    if (email.isOverdue) {
        return '<span class="status-badge overdue">⚠️ Просрочено</span>';
    }

    switch (email.status) {
        case 'ответа нет':
            return '<span class="status-badge pending">⏳ Ожидает</span>';
        case 'отвечено':
            return '<span class="status-badge answered">✅ Отвечено</span>';
        case 'получен ответ':
            return '<span class="status-badge answered">📩 Ответ</span>';
        default:
            return `<span class="status-badge">${escapeHtml(email.status)}</span>`;
    }
}

function getOverdueIndicator(email) {
    const status = (email.status || '').toLowerCase();
    const monitoredStatuses = ['ответа нет', 'ответ не от оператора'];

    if (!monitoredStatuses.includes(status)) {
        return '<span class="overdue-indicator ok">—</span>';
    }

    if (email.isOverdue) {
        const hours = Math.floor(email.hoursOverdue);
        const days = Math.floor(hours / 24);
        const text = days > 0 ? `${days}д ${hours % 24}ч` : `${hours}ч`;
        return `<span class="overdue-indicator danger">+${text}</span>`;
    }

    // Not overdue yet
    const hoursLeft = Math.floor(OVERDUE_HOURS - email.hoursSinceReceived);
    if (hoursLeft < 2) { // Yellow zone nearing 6h
        return `<span class="overdue-indicator warning">${hoursLeft}ч осталось</span>`;
    }

    return `<span class="overdue-indicator ok">${hoursLeft}ч</span>`;
}

// ========================================
// Notifications
// ========================================
function checkForNewOverdueEmails() {
    const currentOverdue = emails.filter(e => e.isOverdue);
    const currentOverdueIds = new Set(currentOverdue.map(e => e.id));

    // Find new overdue emails
    const newOverdue = currentOverdue.filter(e => !previousOverdueIds.has(e.id));

    // Show notification for new overdue (skip first load)
    if (!isFirstLoad && newOverdue.length > 0) {
        showOverdueNotification(newOverdue);
    }

    previousOverdueIds = currentOverdueIds;
}

function showOverdueNotification(overdueEmails) {
    const count = overdueEmails.length;
    let title, body;

    if (count === 1) {
        const email = overdueEmails[0];
        const reason = email.overdueReason;

        if (reason === 'ответ не от оператора') {
            title = '⚠️ Ответ не от оператора';
        } else if (reason === 'ответа нет') {
            title = '⚠️ Нет ответа (6ч+)';
        } else {
            title = '⚠️ Требует внимания';
        }

        body = `${extractName(email.sender)}: ${email.subject.substring(0, 50)}`;
    } else {
        title = `⚠️ ${count} писем требуют внимания`;
        body = overdueEmails.slice(0, 3).map(e => extractName(e.sender)).join(', ');
    }

    window.api.showNotification(title, body);
}

// ========================================
// Reminder System (3-hour external reply)
// ========================================

/**
 * Load list of operators from Google Sheets
 */
async function loadOperators() {
    try {
        const result = await window.api.getOperators();
        if (result.success) {
            operators = result.data.map(e => e.toLowerCase());
            console.log(`Loaded ${operators.length} operators`);
        } else {
            console.error('Failed to load operators:', result.error);
        }
    } catch (error) {
        console.error('Error loading operators:', error);
    }
}

/**
 * Check if email is awaiting reply (3h after external response)
 */
function isAwaitingReply(email) {
    // Only check emails with "получен ответ" status
    if (email.status !== 'получен ответ') return false;

    // Skip if already closed
    if (email.reminderStatus === 'закрыт') return false;

    // Check if last replyer is external (not an operator)
    const lastReplyer = (email.lastReplyer || '').toLowerCase();
    const isExternalReply = !operators.some(op => lastReplyer.includes(op));
    if (!isExternalReply) return false;

    // Check if 3+ hours have passed
    return email.hoursSinceReceived >= REMINDER_HOURS;
}

/**
 * Check for awaiting reply emails and show notifications
 */
function checkForAwaitingReplyEmails() {
    const awaiting = emails.filter(e => isAwaitingReply(e));
    const newAwaiting = awaiting.filter(e => !notifiedReminderIds.has(e.id));

    if (newAwaiting.length > 0) {
        showAwaitingReplyNotification(newAwaiting);
        newAwaiting.forEach(e => notifiedReminderIds.add(e.id));
    }
}

/**
 * Show notification for awaiting reply emails
 */
function showAwaitingReplyNotification(awaitingEmails) {
    const count = awaitingEmails.length;
    let title, body;

    if (count === 1) {
        const email = awaitingEmails[0];
        title = '📩 Ожидает ответа (3ч+)';
        body = `${extractName(email.sender)}: ${email.subject.substring(0, 50)}`;
    } else {
        title = `📩 ${count} писем ожидают ответа`;
        body = awaitingEmails.slice(0, 3).map(e => extractName(e.sender)).join(', ');
    }

    window.api.showNotification(title, body);
}

/**
 * Close reminder for an email
 */
async function closeReminder(rowIndex) {
    try {
        const result = await window.api.closeReminder(rowIndex);
        if (result.success) {
            // Update local state
            const email = emails.find(e => e.rowIndex === rowIndex);
            if (email) {
                email.reminderStatus = 'закрыт';
                notifiedReminderIds.delete(email.id);
            }
            applyFilters();
            renderTable();
        } else {
            console.error('Failed to close reminder:', result.error);
        }
    } catch (error) {
        console.error('Error closing reminder:', error);
    }
}

// ========================================
// Stats & UI Updates
// ========================================
function updateStats() {
    const total = emails.length;
    const overdue = emails.filter(e => e.isOverdue).length;
    const answered = emails.filter(e =>
        e.status === 'отвечено' || e.status === 'получен ответ'
    ).length;

    animateCounter(elements.totalCount, total);
    animateCounter(elements.overdueCount, overdue);
    animateCounter(elements.answeredCount, answered);
}

function animateCounter(element, target) {
    const current = parseInt(element.textContent) || 0;
    if (current === target) return;

    const duration = 300;
    const steps = 20;
    const increment = (target - current) / steps;
    let step = 0;

    const timer = setInterval(() => {
        step++;
        if (step >= steps) {
            element.textContent = target;
            clearInterval(timer);
        } else {
            element.textContent = Math.round(current + increment * step);
        }
    }, duration / steps);
}

function updateLastRefreshTime() {
    const now = new Date();
    const time = now.toLocaleTimeString('ru-RU', { hour: '2-digit', minute: '2-digit' });
    elements.lastUpdate.textContent = time;
}

// ========================================
// UI State Management
// ========================================
function showLoading(show) {
    elements.loadingState.style.display = show ? 'flex' : 'none';
    elements.errorState.style.display = 'none';
    if (show) {
        elements.emptyState.style.display = 'none';
    }
}

function showEmpty(show) {
    elements.emptyState.style.display = show ? 'flex' : 'none';
}

function showError(message) {
    elements.loadingState.style.display = 'none';
    elements.emptyState.style.display = 'none';
    elements.errorState.style.display = 'flex';
    elements.errorMessage.textContent = message;
}

function setRefreshButtonLoading(loading) {
    if (loading) {
        elements.refreshBtn.classList.add('loading');
    } else {
        elements.refreshBtn.classList.remove('loading');
    }
}

function setConnectionStatus(connected) {
    const dot = elements.statusIndicator.querySelector('.status-dot');
    const text = elements.statusIndicator.querySelector('.status-text');

    if (connected) {
        elements.statusIndicator.style.borderColor = 'rgba(34, 197, 94, 0.2)';
        elements.statusIndicator.style.background = 'rgba(34, 197, 94, 0.1)';
        dot.style.background = '#22c55e';
        text.style.color = '#22c55e';
        text.textContent = 'Подключено';
    } else {
        elements.statusIndicator.style.borderColor = 'rgba(239, 68, 68, 0.2)';
        elements.statusIndicator.style.background = 'rgba(239, 68, 68, 0.1)';
        dot.style.background = '#ef4444';
        text.style.color = '#ef4444';
        text.textContent = 'Ошибка';
    }
}

// ========================================
// Auto Refresh
// ========================================
function startAutoRefresh() {
    setInterval(() => {
        loadEmails();
    }, REFRESH_INTERVAL);
}

// ========================================
// Utilities
// ========================================
function extractName(senderStr) {
    if (!senderStr) return 'Неизвестно';

    // Extract name from "Name <email>" format
    const match = senderStr.match(/^([^<]+)/);
    if (match) {
        return match[1].trim() || senderStr;
    }
    return senderStr;
}

function formatDate(date) {
    if (!date) return '—';

    const now = new Date();
    const isToday = date.toDateString() === now.toDateString();
    const time = date.toLocaleTimeString('ru-RU', { hour: '2-digit', minute: '2-digit' });

    if (isToday) {
        return `<div class="time-cell"><span class="date-part">Сегодня</span><span class="time-part">${time}</span></div>`;
    }

    const dateStr = date.toLocaleDateString('ru-RU', { day: '2-digit', month: '2-digit' });
    return `<div class="time-cell"><span class="date-part">${dateStr}</span><span class="time-part">${time}</span></div>`;
}

function escapeHtml(str) {
    if (!str) return '';
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
}

// ========================================
// Admin Panel
// ========================================
let adminPanel = null;
let testEmails = [];

function initAdminPanel() {
    adminPanel = document.getElementById('adminPanel');

    // Hotkey Ctrl+Shift+D (works on any keyboard layout)
    document.addEventListener('keydown', (e) => {
        if (e.ctrlKey && e.shiftKey && e.code === 'KeyD') {
            e.preventDefault();
            toggleAdminPanel();
        }
    });

    // Close button
    document.getElementById('adminCloseBtn')?.addEventListener('click', () => {
        closeAdminPanel();
    });

    // Add test email button
    document.getElementById('addTestEmailBtn')?.addEventListener('click', async () => {
        const subject = document.getElementById('testSubject').value;
        const sender = document.getElementById('testSender').value;
        const overdueMinutes = parseInt(document.getElementById('testOverdue').value) || 0;
        const status = document.getElementById('testStatus').value;

        const result = await window.api.addTestEmail({
            subject,
            sender,
            overdueMinutes,
            status
        });

        if (result.success) {
            testEmails.push(result.email);
            mergeTestEmailsAndRender();
        }
    });

    // Trigger notification button with delay
    document.getElementById('triggerNotificationBtn')?.addEventListener('click', () => {
        const delay = parseInt(document.getElementById('notificationDelay')?.value) || 0;
        const delayMs = delay * 1000;

        if (delayMs > 0) {
            setTimeout(() => {
                window.api.showNotification('🔔 Тестовое уведомление', `Уведомление после ${delay} сек задержки`);
            }, delayMs);
        } else {
            window.api.showNotification('🔔 Тестовое уведомление', 'Это тестовое уведомление из админ-панели');
        }
    });

    // Clear test emails button
    document.getElementById('clearTestEmailsBtn')?.addEventListener('click', async () => {
        await window.api.clearTestEmails();
        testEmails = [];
        mergeTestEmailsAndRender();
    });
}

function toggleAdminPanel() {
    if (adminPanel) {
        adminPanel.classList.toggle('open');
    }
}

function closeAdminPanel() {
    if (adminPanel) {
        adminPanel.classList.remove('open');
    }
}

function mergeTestEmailsAndRender() {
    // Re-filter and render with test emails included
    applyFilters();
    renderTable();
    updateStats();
}

// Override loadEmails to include test emails
const originalLoadEmails = loadEmails;
loadEmails = async function () {
    showLoading(true);
    setRefreshButtonLoading(true);

    try {
        const result = await window.api.getEmails();
        const testResult = await window.api.getTestEmails();

        if (result.success) {
            // Merge real emails with test emails
            const realEmails = result.data.map(processEmail);
            const testEmailsProcessed = (testResult.data || []).map(processEmail);

            emails = [...realEmails, ...testEmailsProcessed];
            testEmails = testResult.data || [];

            applyFilters();
            renderTable();
            updateStats();
            updateLastRefreshTime();
            checkForNewOverdueEmails();
            setConnectionStatus(true);
            isFirstLoad = false;
        } else {
            throw new Error(result.error || 'Unknown error');
        }
    } catch (error) {
        console.error('Failed to load emails:', error);
        showError(error.message);
        setConnectionStatus(false);
    } finally {
        showLoading(false);
        setRefreshButtonLoading(false);
    }
};

// Initialize admin panel via initEventListeners (called on DOMContentLoaded)
// No need for a separate listener here to avoid double initialization
