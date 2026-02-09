const { contextBridge, ipcRenderer } = require('electron');

// Expose protected methods to the renderer process
contextBridge.exposeInMainWorld('api', {
    // Window controls
    windowMinimize: () => ipcRenderer.invoke('window-minimize'),
    windowMaximize: () => ipcRenderer.invoke('window-maximize'),
    windowClose: () => ipcRenderer.invoke('window-close'),

    // Get emails from Google Sheets
    getEmails: () => ipcRenderer.invoke('get-emails'),

    // Show Windows notification
    showNotification: (title, body) => ipcRenderer.invoke('show-notification', { title, body }),

    // Admin panel API
    addTestEmail: (options) => ipcRenderer.invoke('add-test-email', options),
    getTestEmails: () => ipcRenderer.invoke('get-test-emails'),
    clearTestEmails: () => ipcRenderer.invoke('clear-test-emails'),

    // Reminder system API
    getOperators: () => ipcRenderer.invoke('get-operators'),
    closeReminder: (rowIndex) => ipcRenderer.invoke('close-reminder', rowIndex),


    // Listen for new overdue emails (from main process)
    onNewOverdueEmails: (callback) => {
        ipcRenderer.on('new-overdue-emails', (event, emails) => callback(emails));
    },

    // Remove listener
    removeOverdueListener: () => {
        ipcRenderer.removeAllListeners('new-overdue-emails');
    }
});
