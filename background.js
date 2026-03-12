// background.js — MV3 service worker
// Always open as popup. Sidebar is opt-in per session via the toggle button in the popup.
// Mode is NOT persisted — reopening always starts as popup.
chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: false }).catch(e =>
  console.warn("[D365SpeedUp]", e.message)
);

// When switching tabs, close the panel for non-sidebar tabs and try to restore it for the sidebar tab.
chrome.tabs.onActivated.addListener(async ({ tabId, windowId }) => {
  const { sidebarTabId } = await chrome.storage.session.get("sidebarTabId");
  if (tabId === sidebarTabId) {
    chrome.sidePanel.setOptions({ tabId, enabled: true }).catch(() => {});
    chrome.sidePanel.open({ windowId }).catch(() => {});
  } else {
    chrome.sidePanel.setOptions({ tabId, enabled: false }).catch(() => {});
  }
});

// Fires when popup is "" for a tab (i.e. sidebar mode is active for that tab).
// Re-opens / focuses the side panel instead of spawning a new popup.
chrome.action.onClicked.addListener((tab) => {
  if (tab.windowId) {
    chrome.sidePanel.open({ windowId: tab.windowId }).catch(() => {});
  }
});
