
interface SnippetInputOption { label?: string; value?: string; }
interface SnippetDataSource {
  entity: string; select: string[]; orderby?: string; top?: number;
  filterTemplate?: string; minChars?: number; displayFormat?: string;
  valueField?: string; popupateAdditionalFields?: { schemaName: string; fieldName: string }[];
}
interface SnippetInput {
  id: string; label: string; type?: string; placeholder?: string;
  required?: boolean; areaRows?: number; autopopulate?: boolean;
  populateFrom?: string; options?: SnippetInputOption[]; dataSource?: SnippetDataSource;
  defaultValue?: string; showWhen?: { inputId: string; value: string }; appendSuffix?: string; stripSpaces?: boolean;
}
interface Snippet {
  id: string; title: string; description: string; script: string;
  inputs?: SnippetInput[]; outputType?: string; outputSubType?: string;
  copyButtonRequired?: boolean; note?: string; noteShowWhen?: { inputId: string; value: string }; runMode?: string; inputNote?: string;
}
interface Category { categoryName: string; categoryIcon?: string; snippets: Snippet[]; description?: string; }
interface PageContext {
  clientUrl: string; environmentId: string; logicalName: string; displayName: string;
  recordId: string; appId: string; pageType: string; userId: string; userName: string;
  userEmail: string; userRoles: string[]; businessUnitId: string; userLanguage: string;
  orgId: string; orgName: string; orgBaseCurrencyId: string; clientType: string;
  formType: number; formTypeName: string;
}
interface GridOptions {
  enableSearch?: boolean; enableFilters?: boolean; enableSorting?: boolean;
  enableResizing?: boolean; showRenderTime?: boolean; allowHtml?: boolean;
  minSearchChars?: number; collapsed?: boolean; columnOrder?: string[] | null;
}

// ============================================================================
// D365SPEEDUP OBJECT
// ============================================================================
const D365Speedup = {
  Constants: {
    CONFIG_PATH: "config.json",
    DEFAULT_ICON: "📁",
    PAGE_WORLD: "MAIN" as const,
  },
  State: {
    configData: [] as Category[],
    selectedSnippet: null as Snippet | null,
  },
  DOM: {
    sidebar: document.getElementById("sidebar") as HTMLElement,
    burgerMenu: document.getElementById("burgerMenu") as HTMLElement,
    overlay: document.getElementById("overlay") as HTMLElement,
    navContainer: document.getElementById("navContainer") as HTMLElement,
    contentTitle: document.getElementById("contentTitle") as HTMLElement,
    contentArea: document.getElementById("contentArea") as HTMLElement,
    mainTabs: document.getElementById("mainTabs") as HTMLElement,
    panelToggleBtn: document.getElementById("panelToggleBtn") as HTMLButtonElement,

  },
  Enums: {} as Record<string, unknown>,
  Helpers: {} as any,
  Handlers: {} as any,
  Storage: {} as any,
  Core: {} as any,
};

// ============================================================================
// HELPERS
// ============================================================================
D365Speedup.Helpers = {

  isSidebarContext: function (): boolean {
    return new URLSearchParams(window.location.search).get("mode") === "sidebar";
  },

  getEntityDisplayName: async function (logicalName: string): Promise<any> {
    try {
      if (!logicalName) return "";

      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
      if (!tab || !tab.id) return "";

      const [{ result }] = await chrome.scripting.executeScript({
        target: { tabId: tab.id! },
        world: D365Speedup.Constants.PAGE_WORLD,
        func: async (logicalName: string) => {
          try {
            const XrmContext = (window as any).Xrm || (window as any).parent?.Xrm || (window as any).top?.Xrm;
            if (!XrmContext) return "";

            // page cache
            (window as any).__d365_entityDisplayCache ||= {};
            if ((window as any).__d365_entityDisplayCache[logicalName]) {
              return (window as any).__d365_entityDisplayCache[logicalName];
            }

            const globalCtx = XrmContext.Utility.getGlobalContext();
            const clientUrl = globalCtx.getClientUrl?.() || "";
            if (!clientUrl) return "";

            const url =
              `${clientUrl}/api/data/v9.2/EntityDefinitions(LogicalName='${encodeURIComponent(logicalName)}')` +
              `?$select=LogicalName,DisplayName,DisplayCollectionName`;

            const res = await fetch(url, {
              headers: {
                "OData-MaxVersion": "4.0",
                "OData-Version": "4.0",
                Accept: "application/json"
              },
              credentials: "include"
            });

            if (!res.ok) return "";

            const data = await res.json();

            const singular =
              data?.DisplayName?.UserLocalizedLabel?.Label ||
              data?.DisplayName?.LocalizedLabels?.[0]?.Label ||
              "";

            const plural =
              data?.DisplayCollectionName?.UserLocalizedLabel?.Label ||
              data?.DisplayCollectionName?.LocalizedLabels?.[0]?.Label ||
              "";

            const payload = { singular, plural };

            (window as any).__d365_entityDisplayCache[logicalName] = payload;
            return payload;
          } catch (e) {
            return "";
          }
        },
        args: [logicalName]
      });

      // result is either "" or {singular, plural}
      return result || "";
    } catch (err) {
      console.warn("getEntityDisplayName failed:", err);
      return "";
    }
  },

  getCurrentUserEmail: async function (): Promise<string> {
    try {
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
      if (!tab || !tab.id) return "";

      const [{ result }] = await chrome.scripting.executeScript({
        target: { tabId: tab.id! },
        world: D365Speedup.Constants.PAGE_WORLD,
        func: async () => {
          try {
            const XrmContext = (window as any).Xrm || (window as any).parent?.Xrm || (window as any).top?.Xrm;
            if (!XrmContext) return "";

            // page cache (email rarely changes)
            (window as any).__d365_currentUserEmail ||= null;
            if ((window as any).__d365_currentUserEmail) return (window as any).__d365_currentUserEmail;

            const globalCtx = XrmContext.Utility.getGlobalContext();
            const clientUrl = globalCtx.getClientUrl?.() || "";
            if (!clientUrl) return "";

            const userCtx = globalCtx.userSettings || {};
            const userId = (userCtx.userId || "").replace(/[{}]/g, "");
            if (!userId) return "";

            const url =
              `${clientUrl}/api/data/v9.2/systemusers(${encodeURIComponent(userId)})` +
              `?$select=internalemailaddress,domainname`;

            const res = await fetch(url, {
              headers: {
                "OData-MaxVersion": "4.0",
                "OData-Version": "4.0",
                Accept: "application/json",
              },
              credentials: "include",
            });

            if (!res.ok) return "";

            const data = await res.json();
            const email = data?.internalemailaddress || data?.domainname || "";

            (window as any).__d365_currentUserEmail = email;
            return email;
          } catch (e) {
            return "";
          }
        },
        args: []
      });

      return (result as string) || "";
    } catch (err) {
      console.warn("getCurrentUserEmail failed:", err);
      return "";
    }
  },

  getContextInfo: async function (): Promise<Partial<PageContext>> {
    try {
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
      if (!tab || !tab.id) {
        console.warn("No active tab found.");
        return {};
      }

      const [{ result }] = await chrome.scripting.executeScript({
        target: { tabId: tab.id! },
        world: D365Speedup.Constants.PAGE_WORLD,
        func: () => {
          try {
            const XrmContext = (window as any).Xrm || (window as any).parent?.Xrm || (window as any).top?.Xrm;
            if (!XrmContext) return {};

            const globalCtx = XrmContext.Utility.getGlobalContext();
            const pageCtx = XrmContext.Utility?.getPageContext?.();
            const page = XrmContext.Page;

            const environmentId = globalCtx.organizationSettings?.bapEnvironmentId || "";

            const clientUrl = globalCtx.getClientUrl?.() || "";

            const logicalName =
              pageCtx?.input?.entityName ||
              page?.data?.entity?.getEntityName?.() || "";

            const displayName =
              pageCtx?.input?.entityDisplayName ||
              (logicalName
                ? logicalName.charAt(0).toUpperCase() + logicalName.slice(1)
                : "");

            const recordId = (page?.data?.entity?.getId?.() || "").toString().trim().replace(/[{}]/g, "");

            const appId = globalCtx?.organizationSettings?.organizationId || "";
            const pageType = pageCtx?.input?.pageType || "";

            const userCtx = globalCtx.userSettings || {};
            const userId = userCtx.userId?.replace(/[{}]/g, "") || "";
            const userName = userCtx.userName || "";

            const userRoles = userCtx.securityRoles || [];
            const businessUnitId = userCtx.businessUnitId?.replace(/[{}]/g, "") || "";
            const userLanguage = userCtx.languageId || "";

            const orgId = globalCtx.organizationSettings?.organizationId || "";
            const orgName = globalCtx.organizationSettings?.uniqueName || "";
            const orgBaseCurrencyId =
              globalCtx.organizationSettings?.baseCurrencyId || "";

            const clientType = globalCtx.client?.getClient?.() || "";
            const formType = page?.ui?.getFormType?.() || 0;
            const formTypeName =
              ({
                0: "Undefined",
                1: "Create",
                2: "Update",
                3: "Read Only",
                4: "Disabled",
                6: "Bulk Edit",
              } as Record<number, string>)[formType] || "Unknown";

            return {
              clientUrl,
              environmentId,
              logicalName,
              displayName,
              recordId,
              appId,
              pageType,
              userId,
              userName,
              userRoles,
              businessUnitId,
              userLanguage,
              orgId,
              orgName,
              orgBaseCurrencyId,
              clientType,
              formType,
              formTypeName,
            };
          } catch (innerErr) {
            console.warn("Error inside Xrm context:", innerErr);
            return {};
          }
        },
      });

      // ---- Fallback object (unchanged) ----
      const ctx: any =
        result || {
          clientUrl: "",
          environmentId: "",
          logicalName: "",
          displayName: "",
          recordId: "",
          appId: "",
          pageType: "",
          userId: "",
          userName: "",
          userEmail: "",
          userRoles: [],
          businessUnitId: "",
          userLanguage: "",
          orgId: "",
          orgName: "",
          orgBaseCurrencyId: "",
          clientType: "",
          formType: 0,
          formTypeName: "",
        };

      // ---- Upgrade displayName using sibling helper ----
      if (ctx.logicalName) {
        const naive =
          ctx.logicalName
            ? ctx.logicalName.charAt(0).toUpperCase() + ctx.logicalName.slice(1)
            : "";

        const shouldUpgrade = !ctx.displayName || ctx.displayName === naive;

        if (shouldUpgrade && D365Speedup?.Helpers?.getEntityDisplayName) {
          try {
            const meta = await D365Speedup.Helpers.getEntityDisplayName(ctx.logicalName);

            const betterName = (meta && typeof meta === "object")
              ? (meta.singular || "")
              : (meta || "");

            if (betterName) ctx.displayName = betterName;
          } catch (e: any) {
            console.warn("DisplayName metadata lookup failed:", e?.message || e);
          }
        }
      }

      ctx.userEmail = await D365Speedup.Helpers.getCurrentUserEmail();
      return ctx;
    } catch (err) {
      console.warn("Could not get Xrm context:", err);
      return {};
    }
  },

  fetchODataAutoComplete: async function (dataSource: SnippetDataSource, query: string): Promise<any[]> {
    try {
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
      if (!tab || !tab.id) {
        console.warn("No active D365 tab found.");
        return [];
      }

      const [{ result }] = await chrome.scripting.executeScript({
        target: { tabId: tab.id! },
        world: D365Speedup.Constants.PAGE_WORLD,
        func: async (dataSource: any, query: string) => {
          try {
            const XrmContext = (window as any).Xrm || (window as any).parent?.Xrm || (window as any).top?.Xrm;
            if (!XrmContext) {
              console.warn("Xrm not available in page context.");
              return [];
            }

            const baseUrl = XrmContext.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2";
            const safeQuery = query.replace(/'/g, "''");

            const { entity, select, orderby, top, filterTemplate } = dataSource;
            if (!entity) throw new Error("Missing entity in dataSource.");

            const toPlural = (name: string) => {
              if (!name) return "";
              if (name.toLowerCase().includes("definition")) return name;
              if (name.endsWith("y") && !/[aeiou]y$/i.test(name)) return name.slice(0, -1) + "ies";
              if (/(s|sh|ch|x|z)$/i.test(name)) return name + "es";
              return name + "s";
            };

            const isMetadata = /definition/i.test(entity);

            if (isMetadata) {
              const url =
                `${baseUrl}/${entity}` +
                `?$select=${select.join(",")}` +
                `&$filter=${filterTemplate}`;

              const response = await fetch(url, { headers: { Accept: "application/json" } });
              const json = await response.json();

              if (!response.ok) throw new Error(json.error?.message || response.statusText);

              let results = json.value || [];
              const q = query.toLowerCase();

              results = results.filter((r: any) => {
                const ln = (r.LogicalName || "").toLowerCase();
                const dn = (r.DisplayName?.UserLocalizedLabel?.Label || "").toLowerCase();
                return ln.includes(q) || dn.includes(q);
              });

              return results.slice(0, top || 10);
            }

            const filter = (filterTemplate || "")
              .replace(/{query}/g, safeQuery)
              .replace(/\s+/g, " ")
              .trim();

            const encodedFilter = encodeURI(filter);
            const pluralEntity = toPlural(entity);

            const url =
              `${baseUrl}/${pluralEntity}` +
              `?$select=${select.join(",")}` +
              `${orderby ? `&$orderby=${orderby}` : ""}` +
              `${top ? `&$top=${top}` : ""}` +
              (filter ? `&$filter=${encodedFilter}` : "");

            const response = await fetch(url, { headers: { Accept: "application/json" } });
            const json = await response.json();

            if (!response.ok) throw new Error(json.error?.message || response.statusText);

            return json.value || [];
          } catch (err: any) {
            console.warn("AutoComplete error:", err.message);
            return [];
          }
        },
        args: [dataSource, query],
      });

      return (result as any[]) || [];
    } catch (err: any) {
      console.warn("AutoComplete fetch failed:", err.message);
      return [];
    }
  },

  attachAutoCompleteInput: function (inputEl: HTMLInputElement, dataSource: SnippetDataSource): void {
    const container = document.createElement("div");
    container.className = "autocomplete-container";
    (inputEl.parentNode as HTMLElement).insertBefore(container, inputEl);

    const wrapper = document.createElement("div");
    wrapper.className = "input-wrapper";
    container.appendChild(wrapper);
    wrapper.appendChild(inputEl);

    const clearBtn = document.createElement("span");
    clearBtn.className = "clear-btn";
    clearBtn.title = "Clear";
    clearBtn.textContent = "✕";
    wrapper.appendChild(clearBtn);

    const spinner = document.createElement("div");
    spinner.className = "autocomplete-spinner";
    wrapper.appendChild(spinner);

    const list = document.createElement("div");
    list.className = "autocomplete-list";
    container.appendChild(list);

    const highlightMatch = (text: string, query: string): string => {
      const regex = new RegExp(`(${query})`, "ig");
      return text.replace(regex, "<mark>$1</mark>");
    };

    const minChars = dataSource.minChars ?? 2;
    const displayTemplate = dataSource.displayFormat || "{name}";
    const valueField = dataSource.valueField || "name";
    let debounceTimer: ReturnType<typeof setTimeout> | null = null;

    clearBtn.addEventListener("click", () => {
      inputEl.value = "";
      list.innerHTML = "";
      clearBtn.style.display = "none";
    });

    const toggleClearBtn = () => {
      clearBtn.style.display = inputEl.value.trim() ? "block" : "none";
    };

    inputEl.addEventListener("input", async () => {
      toggleClearBtn();
      const query = inputEl.value.trim();
      if (debounceTimer !== null) clearTimeout(debounceTimer);

      if (query.length < minChars) {
        list.innerHTML = "";
        wrapper.classList.remove("loading");
        return;
      }

      debounceTimer = setTimeout(async () => {
        try {
          wrapper.classList.add("loading");
          const results = await D365Speedup.Helpers.fetchODataAutoComplete(dataSource, query);
          wrapper.classList.remove("loading");
          toggleClearBtn();

          list.innerHTML = "";

          if (!results.length) {
            const noRes = document.createElement("div");
            noRes.className = "autocomplete-item disabled";
            noRes.textContent = "No results found";
            list.appendChild(noRes);
            return;
          }

          results.forEach((r: any) => {
            const displayText = D365Speedup.Helpers.resolveTemplate(r, displayTemplate);
            const item = document.createElement("div") as any;
            item.className = "autocomplete-item";
            item.innerHTML = highlightMatch(displayText, query);
            item.data = r;

            item.onclick = function (this: any) {
              inputEl.value = (this as any).data[valueField];

              if (dataSource.popupateAdditionalFields && Array.isArray(dataSource.popupateAdditionalFields)) {
                dataSource.popupateAdditionalFields.forEach((extra: { schemaName: string; fieldName: string }) => {
                  const { schemaName, fieldName } = extra;
                  if (!schemaName || !fieldName) return;

                  const resolvedValue = D365Speedup.Helpers.resolveTemplate((this as any).data, `{${schemaName}}`);
                  const targetEl = document.getElementById(fieldName) as HTMLInputElement | null;
                  if (targetEl) {
                    let val = resolvedValue || "";
                    if (targetEl.dataset.stripSpaces) val = val.replace(/\s+/g, "");
                    const suffix = targetEl.dataset.appendSuffix;
                    if (suffix && !val.endsWith(suffix)) val += suffix;
                    targetEl.value = val;

                    const w = targetEl.closest(".input-wrapper-dv") || targetEl.closest(".input-wrapper");
                    const c = w?.querySelector(".clear-btn") as HTMLElement | null;
                    if (c) c.style.display = targetEl.value ? "block" : "none";
                  }
                });
              }

              list.innerHTML = "";
              toggleClearBtn();
            };

            list.appendChild(item);
          });
        } catch (err: any) {
          wrapper.classList.remove("loading");
          console.warn("Autocomplete fetch error:", err.message);
        }
      }, 300);
    });

    document.addEventListener("click", (e: MouseEvent) => {
      if (!container.contains(e.target as Node)) list.innerHTML = "";
    });
  },

  showToast: function (msg: string): void {
    const toast = document.createElement("div");
    toast.className = "toast";
    toast.textContent = msg;
    document.body.appendChild(toast);
    setTimeout(() => toast.classList.add("show"), 10);
    setTimeout(() => toast.remove(), 2000);
  },

  showError: function (message: string): void {
    D365Speedup.DOM.contentArea.innerHTML = `<div class="error-box">${message}</div>`;
  },

  generateTable: function (data: any): string {
    if (!Array.isArray(data)) return `<div>${data}</div>`;
    if (data.length === 0) return `<div>No results</div>`;

    const headers = Object.keys(data[0]);
    const rows = data
      .map((r: any) => `<tr>${headers.map((h) => `<td>${r[h] ?? ""}</td>`).join("")}</tr>`)
      .join("");

    return `
      <table class="output-table">
        <thead><tr>${headers.map((h) => `<th>${h}</th>`).join("")}</tr></thead>
        <tbody>${rows}</tbody>
      </table>
    `;
  },

  copyToClipboard: function (text: string): void {
    navigator.clipboard.writeText(text || "");
    D365Speedup.Helpers.showToast("Copied to clipboard!");
  },

  escapeHTML: function (str: any): string {
    return String(str ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  },

  resolveTemplate: function (obj: any, template: string): string {
    return template.replace(/\{([\w./]+)\}/g, (_: string, path: string) => {
      try {
        return path.split('/').reduce((acc: any, key: string) => acc?.[key], obj) ?? '';
      } catch {
        return '';
      }
    });
  },

  // =====================================================================
  // Generic: Collapsible Sections + Advanced Tables (reusable by snippets)
  // =====================================================================
  renderCollapsibleTables: function (sections: any[], options: any = {}): string {
    const {
      openFirst = true,
      emptyText = "No records",
      showCounts = true,
    } = options;

    const esc = (s: any): string =>
      String(s ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");

    const renderCell = (val: any): string => {
      if (val == null) return "";
      const s = String(val);
      if (s.includes("<a ") || s.includes("</a>") || s.includes("<span") || s.includes("<div")) return s;
      return esc(s);
    };

    if (!Array.isArray(sections) || !sections.length) {
      return `<div class="no-data">${esc(emptyText)}</div>`;
    }

    const htmlSections = sections.map((sec: any, idx: number) => {
      const title = sec?.title || `Section ${idx + 1}`;
      const open = typeof sec?.open === "boolean" ? sec.open : (openFirst && idx === 0);
      const key = (sec?.key || `sec${idx + 1}`).replace(/[^\w-]/g, "");

      const payload = sec?.rows;

      // interactiveTable payload
      if (payload && typeof payload === "object" && payload.__type === "interactiveTable") {
        const gridId = `grid_${key}_${Math.random().toString(36).slice(2, 8)}`;
        const rowsArr = Array.isArray(payload.rows) ? payload.rows : [];
        const countText = showCounts ? ` (${rowsArr.length})` : "";

        return `
          <details ${open ? "open" : ""} data-grid="1" data-gridid="${gridId}">
            <summary>${esc(title)}${countText}</summary>
            <div id="${gridId}"></div>
          </details>
        `;
      }

      // normal HTML-table section (existing behavior)
      const rows = Array.isArray(sec?.rows) ? sec.rows : [];
      const countText = showCounts ? ` (${rows.length})` : "";

      if (!rows.length) {
        return `
          <details ${open ? "open" : ""}>
            <summary>${esc(title)}${countText}</summary>
            <div class="no-data" style="padding:10px;">${esc(emptyText)}</div>
          </details>
        `;
      }

      const headers = Object.keys(rows[0] || {});
      const thead = headers.map((h: string) => `<th class="sortable" data-col="${esc(h)}">${esc(h)}</th>`).join("");
      const tbody = rows.map((r: any) =>
        `<tr>${headers.map((h: string) => `<td>${renderCell(r?.[h])}</td>`).join("")}</tr>`
      ).join("");

      return `
        <details ${open ? "open" : ""}>
          <summary>${esc(title)}${countText}</summary>
          <div class="table-wrapper">
            <table class="advanced-table sortable">
              <thead><tr>${thead}</tr></thead>
              <tbody>${tbody}</tbody>
            </table>
          </div>
        </details>
      `;
    }).join("");

    // After HTML is placed in DOM, bind any interactive grids
    requestAnimationFrame(() => {
      const root = document.querySelector(".collapsible-tables");
      if (!root) return;

      root.querySelectorAll('details[data-grid="1"]').forEach((detailsEl: Element) => {
        const gridId = detailsEl.getAttribute("data-gridid");
        if (!gridId) return;

        // find the matching section again by id pattern
        const secKey = gridId.split("_")[1]; // grid_<key>_<rand>
        const sec = sections.find((s: any) => (String(s?.key || "").replace(/[^\w-]/g, "") === secKey));
        const payload = sec?.rows;

        if (!payload || payload.__type !== "interactiveTable") return;

        const container = document.getElementById(gridId);
        if (!container) return;

        D365Speedup.Helpers.bindInteractiveGrid(
          container,
          payload.rows || [],
          payload.datasetName || (sec?.title || "Result"),
          payload.gridOptions || {}
        );
      });
    });

    return `
      <div class="collapsible-tables">
        ${htmlSections}
      </div>
    `;
  },

  // =====================================================================
  // Interactive Data Grid (Search + Sort + Column Filters + Resize)
  // =====================================================================
  bindInteractiveGrid: function (containerEl: HTMLElement, data: any[], datasetName: string, options: GridOptions = {}): void {
    if (!containerEl) return;
    containerEl.innerHTML = "";

    // Defaults (backwards compatible)
    const defaultOptions: GridOptions = {
      enableSearch: true,
      enableFilters: true,
      enableSorting: true,
      enableResizing: true,
      showRenderTime: false,
      allowHtml: false,
      minSearchChars: 2,
      collapsed: false,
      columnOrder: null
    };

    const mergedOptions: GridOptions = { ...defaultOptions, ...(options || {}) };

    if (!D365Speedup.Helpers.InteractiveTable) {
      D365Speedup.Helpers.InteractiveTable = class InteractiveTable {
        container: HTMLElement;
        datasetName: string;
        rawData: any[];
        options: GridOptions;
        columns: string[];
        globalSearchTerm: string;
        columnFilters: Record<string, any>;
        sortState: { columnKey: string | null; direction: string | null };
        distinctValues: Record<string, string[]>;
        columnMeta: Record<string, { filterType: string }>;
        globalSearchInput: HTMLInputElement | null;
        refreshButton: HTMLButtonElement | null;
        titleLeft: HTMLElement | null;
        chevron: HTMLElement | null;
        colgroup: HTMLElement | null;
        headerRow: HTMLElement | null;
        tbody: HTMLElement | null;
        footer: HTMLElement | null;
        filterPopup: HTMLElement | null;
        activeFilterColumn: string | null;
        _applyCollapsedUi: (() => void) | null;

        constructor(config: { containerId: string; datasetName?: string; data: any[]; options?: GridOptions }) {
          this.container = document.getElementById(config.containerId) as HTMLElement;
          this.datasetName = config.datasetName || "Dataset";
          this.rawData = Array.isArray(config.data) ? config.data.slice() : [];
          this.options = { ...defaultOptions, ...(config.options || {}) };

          this.columns = this.inferColumns();
          this.globalSearchTerm = "";
          this.columnFilters = {};
          this.sortState = { columnKey: null, direction: null };
          this.distinctValues = this.computeDistinctValues();
          this.columnMeta = this.buildColumnMeta();

          this.globalSearchInput = null;
          this.refreshButton = null;
          this.titleLeft = null;
          this.chevron = null;
          this.colgroup = null;
          this.headerRow = null;
          this.tbody = null;
          this.footer = null;
          this.filterPopup = null;
          this.activeFilterColumn = null;
          this._applyCollapsedUi = null;

          this.buildStructure();
          this.buildHeaderCells();
          this.attachGlobalEvents();

          // apply initial collapsed state AFTER DOM exists
          if (this.options.collapsed) {
            this.container.classList.add("collapsed");
          } else {
            this.container.classList.remove("collapsed");
          }

          this.updateTable();
        }

        inferColumns(): string[] {
          const cols: string[] = [];
          const colSet = new Set<string>();

          // discover columns from data
          this.rawData.forEach((row: any) => {
            Object.keys(row || {}).forEach((k) => {
              if (!colSet.has(k)) {
                colSet.add(k);
                cols.push(k);
              }
            });
          });

          // apply preferred order if provided
          const preferred = Array.isArray(this.options?.columnOrder) ? this.options.columnOrder : null;
          if (preferred && preferred.length) {
            const ordered: string[] = [];
            preferred.forEach((k) => {
              if (colSet.has(k) && !ordered.includes(k)) ordered.push(k);
            });
            cols.forEach((k) => {
              if (!ordered.includes(k)) ordered.push(k);
            });
            return ordered;
          }

          return cols;
        }

        computeDistinctValues(): Record<string, string[]> {
          const map: Record<string, string[]> = {};
          this.columns.forEach((key) => {
            const set = new Set<string>();
            this.rawData.forEach((row: any) => {
              const val = row?.[key];
              if (val !== undefined && val !== null && String(val) !== "") set.add(String(val));
            });
            map[key] = Array.from(set);
          });
          return map;
        }

        buildColumnMeta(): Record<string, { filterType: string }> {
          const meta: Record<string, { filterType: string }> = {};
          this.columns.forEach((key) => {
            const values = this.distinctValues[key] || [];
            let longest = 0;
            values.forEach((v) => { longest = Math.max(longest, String(v ?? "").length); });
            const filterType = values.length > 0 && values.length <= 15 && longest <= 25 ? "list" : "text";
            meta[key] = { filterType };
          });
          return meta;
        }

        buildStructure(): void {
          this.container.classList.add("data-table-container");

          const titleBar = document.createElement("div");
          titleBar.className = "table-title-bar";

          const titleLeft = document.createElement("div");
          titleLeft.className = "table-title-left";

          const chevron = document.createElement("span");
          chevron.className = "chevron-icon";
          chevron.innerHTML = "&#9660;";

          const datasetNameEl = document.createElement("span");
          datasetNameEl.className = "dataset-name";
          datasetNameEl.textContent = this.datasetName;

          titleLeft.appendChild(chevron);
          titleLeft.appendChild(datasetNameEl);

          this.chevron = chevron;
          this.titleLeft = titleLeft;

          const titleRight = document.createElement("div");
          titleRight.className = "table-title-right";

          const globalSearch = document.createElement("input");
          globalSearch.type = "text";
          globalSearch.placeholder = "Search all columns...";
          globalSearch.className = "global-search-input";
          this.globalSearchInput = globalSearch;

          const refreshButton = document.createElement("button");
          refreshButton.type = "button";
          refreshButton.className = "refresh-button";
          refreshButton.title = "Refresh grid (reset filters/search/sort)";
          refreshButton.innerHTML = "&#x21bb;";
          this.refreshButton = refreshButton;

          titleRight.appendChild(globalSearch);
          titleRight.appendChild(refreshButton);

          // feature toggle: search
          if (!this.options.enableSearch) globalSearch.style.display = "none";

          titleBar.appendChild(titleLeft);
          titleBar.appendChild(titleRight);
          this.container.appendChild(titleBar);

          const inner = document.createElement("div");
          inner.className = "table-inner";

          const scrollWrapper = document.createElement("div");
          scrollWrapper.className = "table-scroll-wrapper";

          const table = document.createElement("table");
          table.className = "data-table";

          const colgroup = document.createElement("colgroup");
          table.appendChild(colgroup);
          this.colgroup = colgroup;

          const thead = document.createElement("thead");
          const headerRow = document.createElement("tr");
          thead.appendChild(headerRow);
          this.headerRow = headerRow;

          const tbody = document.createElement("tbody");
          this.tbody = tbody;

          table.appendChild(thead);
          table.appendChild(tbody);

          scrollWrapper.appendChild(table);
          inner.appendChild(scrollWrapper);

          const footer = document.createElement("div");
          footer.className = "table-footer-info";
          this.footer = footer;
          inner.appendChild(footer);

          this.container.appendChild(inner);

          const filterPopup = document.createElement("div");
          filterPopup.className = "column-filter-popup hidden";
          this.filterPopup = filterPopup;
          document.body.appendChild(filterPopup);

          // collapse-aware UI (hide search + refresh when collapsed)
          const applyCollapsedUi = () => {
            const isCollapsed = this.container.classList.contains("collapsed");

            // Search box: only if feature enabled
            if (this.options.enableSearch) {
              globalSearch.style.display = isCollapsed ? "none" : "";
            }

            // Refresh: hide when collapsed
            refreshButton.style.display = isCollapsed ? "none" : "";
          };

          // store for chevron click handler to reuse
          this._applyCollapsedUi = applyCollapsedUi;

          // initial state (supports options.collapsed)
          if (this.options.collapsed) {
            this.container.classList.add("collapsed");
          }

          // apply initial visibility
          applyCollapsedUi();
        }

        buildHeaderCells(): void {
          (this.headerRow as HTMLElement).innerHTML = "";
          (this.colgroup as HTMLElement).innerHTML = "";

          this.columns.forEach((key, index) => {
            const col = document.createElement("col");
            col.style.width = "160px";
            (this.colgroup as HTMLElement).appendChild(col);

            const th = document.createElement("th");
            th.dataset.key = key;
            th.dataset.index = String(index);

            const content = document.createElement("div");
            content.className = "th-content";

            const label = document.createElement("span");
            label.className = "header-label";
            label.textContent = key;

            const sortIcons = document.createElement("span");
            sortIcons.className = "sort-icons";
            sortIcons.innerHTML = '<span class="sort-up">▲</span><span class="sort-down">▼</span>';

            const filterIcon = document.createElement("span");
            filterIcon.className = "filter-icon";
            filterIcon.title = "Filter column";
            filterIcon.innerHTML =
              '<svg viewBox="0 0 16 16" aria-hidden="true">' +
              '<path d="M2 2h12L10 7v5.5a1 1 0 0 1-.553.894l-2 1A1 1 0 0 1 6 13.5V7L2 2z"></path>' +
              "</svg>";

            const resizer = document.createElement("span");
            resizer.className = "col-resizer";

            content.appendChild(label);

            if (this.options.enableSorting) content.appendChild(sortIcons);
            if (this.options.enableFilters) content.appendChild(filterIcon);
            if (this.options.enableResizing) content.appendChild(resizer);

            th.appendChild(content);
            (this.headerRow as HTMLElement).appendChild(th);

            th.addEventListener("click", (e: MouseEvent) => {
              if (!this.options.enableSorting) return;
              if ((e.target as Element).closest(".filter-icon") || (e.target as Element).closest(".col-resizer")) return;
              this.handleSortClick(key);
            });

            if (this.options.enableFilters) {
              filterIcon.addEventListener("click", (e: MouseEvent) => {
                e.stopPropagation();
                this.openFilterPopup(key, th);
              });
            }

            if (this.options.enableResizing) {
              this.attachResizerEvents(resizer, index);
            }
          });

          this.updateFilterIconClasses();
        }

        attachGlobalEvents(): void {
          if (this.options.enableSearch && this.globalSearchInput) {
            this.globalSearchInput.addEventListener("input", (e: Event) => {
              this.globalSearchTerm = (e.target as HTMLInputElement).value || "";
              this.updateTable();
            });
          }

          if (this.refreshButton) {
            this.refreshButton.addEventListener("click", (e: MouseEvent) => {
              e.stopPropagation();
              this.resetState();
            });
          }

          // collapse toggle (always available)
          if (this.chevron) {
            this.chevron.addEventListener("click", (e: MouseEvent) => {
              if ((e.target as Element).closest(".global-search-input") || (e.target as Element).closest(".refresh-button")) return;

              this.container.classList.toggle("collapsed");

              // keep search + refresh hidden when collapsed, visible when expanded
              if (typeof (this as any)._applyCollapsedUi === "function") {
                (this as any)._applyCollapsedUi();
              }
            });
          }

          document.addEventListener("click", (e: MouseEvent) => {
            if ((this.filterPopup as HTMLElement).classList.contains("hidden")) return;
            if ((this.filterPopup as HTMLElement).contains(e.target as Node)) return;
            if ((e.target as Element).closest(".filter-icon")) return;
            this.closeFilterPopup();
          });

          document.addEventListener("keydown", (e: KeyboardEvent) => {
            if (e.key === "Escape") this.closeFilterPopup();
          });
        }

        resetState(): void {
          this.globalSearchTerm = "";
          if (this.globalSearchInput) this.globalSearchInput.value = "";
          this.columnFilters = {};
          this.sortState = { columnKey: null, direction: null };
          this.updateSortClasses();
          this.updateFilterIconClasses();
          this.updateTable();
        }

        handleSortClick(columnKey: string): void {
          const state = this.sortState;
          if (state.columnKey === columnKey) {
            if (state.direction === "asc") state.direction = "desc";
            else if (state.direction === "desc") { state.columnKey = null; state.direction = null; }
            else state.direction = "asc";
          } else {
            state.columnKey = columnKey;
            state.direction = "asc";
          }
          this.updateSortClasses();
          this.updateTable();
        }

        updateSortClasses(): void {
          const ths = (this.headerRow as HTMLElement).querySelectorAll("th");
          ths.forEach((th) => {
            th.classList.remove("sorted-asc", "sorted-desc");
            const key = (th as HTMLElement).dataset.key;
            if (this.sortState.columnKey === key) {
              if (this.sortState.direction === "asc") th.classList.add("sorted-asc");
              else if (this.sortState.direction === "desc") th.classList.add("sorted-desc");
            }
          });
        }

        updateFilterIconClasses(): void {
          if (!this.options.enableFilters) return;
          const ths = (this.headerRow as HTMLElement).querySelectorAll("th");
          ths.forEach((th) => {
            const key = (th as HTMLElement).dataset.key!;
            const icon = th.querySelector(".filter-icon");
            if (!icon) return;
            if (this.columnFilters[key]) icon.classList.add("active");
            else icon.classList.remove("active");
          });
        }

        attachResizerEvents(resizer: HTMLElement, colIndex: number): void {
          let startX = 0;
          let startWidth = 0;
          const colElement = (this.colgroup as HTMLElement).children[colIndex] as HTMLElement;
          const minWidth = 60;

          const onMouseMove = (e: MouseEvent) => {
            const delta = e.clientX - startX;
            const newWidth = Math.max(minWidth, startWidth + delta);
            colElement.style.width = newWidth + "px";
          };

          const onMouseUp = () => {
            document.removeEventListener("mousemove", onMouseMove);
            document.removeEventListener("mouseup", onMouseUp);
            resizer.classList.remove("resizing");
          };

          resizer.addEventListener("mousedown", (e: MouseEvent) => {
            e.preventDefault();
            e.stopPropagation();
            startX = e.clientX;
            startWidth = colElement.getBoundingClientRect().width;
            resizer.classList.add("resizing");
            document.addEventListener("mousemove", onMouseMove);
            document.addEventListener("mouseup", onMouseUp);
          });
        }

        openFilterPopup(columnKey: string, headerCell: HTMLElement): void {
          if (!this.options.enableFilters) return;

          this.activeFilterColumn = columnKey;
          const popup = this.filterPopup as HTMLElement;
          popup.innerHTML = "";

          const meta = this.columnMeta[columnKey] || { filterType: "text" };
          const filterType = meta.filterType;
          const currentFilter = this.columnFilters[columnKey];
          const distinct = this.distinctValues[columnKey] || [];

          const header = document.createElement("div");
          header.className = "filter-header";
          header.textContent = `Filter: ${columnKey}`;
          popup.appendChild(header);

          const section = document.createElement("div");
          section.className = "filter-section";

          if (filterType === "list") {
            const selectedSet =
              currentFilter && currentFilter.type === "list" && currentFilter.values
                ? new Set(currentFilter.values)
                : null;
            const allSelected = !selectedSet || selectedSet.size === 0;

            distinct.forEach((val) => {
              const row = document.createElement("div");
              row.className = "filter-row";
              const input = document.createElement("input");
              input.type = "checkbox";
              input.value = val;
              input.checked = allSelected || (selectedSet !== null && selectedSet.has(val));
              const labelEl = document.createElement("label");
              labelEl.textContent = val;
              row.appendChild(input);
              row.appendChild(labelEl);
              section.appendChild(row);
            });
          } else {
            const textInput = document.createElement("input");
            textInput.className = "filter-text-input";
            textInput.type = "text";
            textInput.placeholder = "Contains text...";
            textInput.value =
              currentFilter && currentFilter.type === "text" ? currentFilter.text || "" : "";
            section.appendChild(textInput);
          }

          popup.appendChild(section);

          const actions = document.createElement("div");
          actions.className = "filter-actions";
          const actionsLeft = document.createElement("div");
          actionsLeft.className = "filter-actions-left";
          const actionsRight = document.createElement("div");
          actionsRight.className = "filter-actions-right";

          const btnSelectAll = document.createElement("button");
          btnSelectAll.className = "btn";
          btnSelectAll.textContent = "Select All";

          const btnClear = document.createElement("button");
          btnClear.className = "btn";
          btnClear.textContent = "Clear";

          const btnApply = document.createElement("button");
          btnApply.className = "btn btn-primary";
          btnApply.textContent = "Apply";

          actionsLeft.appendChild(btnSelectAll);
          actionsLeft.appendChild(btnClear);
          actionsRight.appendChild(btnApply);
          actions.appendChild(actionsLeft);
          actions.appendChild(actionsRight);
          popup.appendChild(actions);

          if (filterType === "list") {
            const checkboxes = Array.from(section.querySelectorAll('input[type="checkbox"]')) as HTMLInputElement[];
            btnSelectAll.addEventListener("click", () => checkboxes.forEach((cb) => (cb.checked = true)));
            btnClear.addEventListener("click", () => checkboxes.forEach((cb) => (cb.checked = false)));
            btnApply.addEventListener("click", () => {
              const selected = checkboxes.filter((cb) => cb.checked).map((cb) => cb.value);
              if (selected.length === 0 || selected.length === distinct.length) delete this.columnFilters[columnKey];
              else this.columnFilters[columnKey] = { type: "list", values: new Set(selected) };
              this.closeFilterPopup();
              this.updateTable();
            });
          } else {
            const textInput = section.querySelector("input.filter-text-input") as HTMLInputElement;
            btnSelectAll.style.display = "none";
            btnClear.addEventListener("click", () => { textInput.value = ""; });
            btnApply.addEventListener("click", () => {
              const value = (textInput.value || "").trim();
              if (value) this.columnFilters[columnKey] = { type: "text", text: value };
              else delete this.columnFilters[columnKey];
              this.closeFilterPopup();
              this.updateTable();
            });
          }

          const rect = headerCell.getBoundingClientRect();
          popup.style.left = rect.left + window.scrollX + "px";
          popup.style.top = rect.bottom + window.scrollY + "px";
          popup.classList.remove("hidden");
        }

        closeFilterPopup(): void {
          (this.filterPopup as HTMLElement).classList.add("hidden");
          this.activeFilterColumn = null;
        }

        rowPassesColumnFilters(row: any): boolean {
          for (const key in this.columnFilters) {
            const cfg = this.columnFilters[key];
            const rawVal = row?.[key];
            const val = rawVal == null ? "" : String(rawVal);

            if (cfg.type === "list") {
              const set: Set<string> = cfg.values;
              if (set && set.size > 0 && !set.has(val)) return false;
            } else if (cfg.type === "text") {
              const text = (cfg.text || "").toLowerCase();
              if (text && !val.toLowerCase().includes(text)) return false;
            }
          }
          return true;
        }

        computeFilteredSortedData(): any[] {
          const termRaw = this.globalSearchTerm || "";
          const minChars = this.options.minSearchChars ?? 2;
          const globalTerm = termRaw.length >= minChars ? termRaw.toLowerCase() : "";

          let rows = this.rawData.filter((row: any) => {
            if (!this.rowPassesColumnFilters(row)) return false;
            if (globalTerm) {
              return this.columns.some((key) => {
                const value = row?.[key];
                return value != null && String(value).toLowerCase().includes(globalTerm);
              });
            }
            return true;
          });

          if (this.options.enableSorting) {
            const state = this.sortState;
            if (state.columnKey && state.direction) {
              const key = state.columnKey;
              const direction = state.direction === "asc" ? 1 : -1;

              const toNum = (v: any): number | null => {
                if (v === null || v === undefined || v === "") return null;
                const n = Number(v);
                return Number.isFinite(n) ? n : null;
              };

              const toDateMs = (v: any): number | null => {
                if (v === null || v === undefined || v === "") return null;
                const ms = Date.parse(String(v));
                return Number.isFinite(ms) ? ms : null;
              };

              const sample = rows
                .slice(0, 50)
                .map((r: any) => r?.[key])
                .filter((v: any) => v !== null && v !== undefined && v !== "");

              const numericHits = sample.filter((v: any) => toNum(v) !== null).length;
              const dateHits = sample.filter((v: any) => toDateMs(v) !== null).length;

              const sortMode =
                sample.length === 0
                  ? "string"
                  : dateHits / sample.length >= 0.7
                    ? "date"
                    : numericHits / sample.length >= 0.9
                      ? "number"
                      : "string";

              rows.sort((a: any, b: any) => {
                const av = a?.[key];
                const bv = b?.[key];

                if (sortMode === "date") {
                  const da = toDateMs(av);
                  const db = toDateMs(bv);
                  if (da === null && db === null) return 0;
                  if (da === null) return 1;
                  if (db === null) return -1;
                  return (da < db ? -1 : da > db ? 1 : 0) * direction;
                }

                if (sortMode === "number") {
                  const na = toNum(av);
                  const nb = toNum(bv);
                  if (na === null && nb === null) return 0;
                  if (na === null) return 1;
                  if (nb === null) return -1;
                  return (na < nb ? -1 : na > nb ? 1 : 0) * direction;
                }

                const sa = String(av ?? "").toLowerCase();
                const sb = String(bv ?? "").toLowerCase();
                if (sa < sb) return -1 * direction;
                if (sa > sb) return 1 * direction;
                return 0;
              });
            }
          }

          return rows;
        }

        escapeHTML(text: any): string {
          return String(text)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;");
        }

        getHighlightedHTML(text: string, term: string): string {
          if (!term || term.length < 2) return this.escapeHTML(text);
          const escTerm = term.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
          const regex = new RegExp(escTerm, "ig");
          const escapedText = this.escapeHTML(text);
          return escapedText.replace(regex, (match) => `<span class="cell-highlight">${match}</span>`);
        }

        renderBodyRows(rows: any[], highlightTerm: string): void {
          (this.tbody as HTMLElement).innerHTML = "";
          const frag = document.createDocumentFragment();

          const isHtmlCell = (v: any): boolean => {
            const s = String(v ?? "");
            return s.includes("<a ") || s.includes("</a>") || s.includes("<span") || s.includes("<div");
          };

          rows.forEach((row: any) => {
            const tr = document.createElement("tr");

            this.columns.forEach((key) => {
              const td = document.createElement("td");
              const raw = row?.[key];
              const value = raw == null ? "" : String(raw);

              const allowHtml = !!this.options.allowHtml;

              // Tooltip: show full value on hover
              const tooltipText = value
                .replace(/<[^>]*>/g, "")
                .trim();

              if (tooltipText) td.title = tooltipText;

              if (allowHtml && isHtmlCell(value)) {
                td.innerHTML = value;
              } else if (
                this.options.enableSearch &&
                highlightTerm &&
                highlightTerm.length >= (this.options.minSearchChars ?? 2) &&
                value.toLowerCase().includes(highlightTerm.toLowerCase())
              ) {
                td.innerHTML = this.getHighlightedHTML(value, highlightTerm);
              } else {
                td.textContent = value;
              }

              tr.appendChild(td);
            });

            frag.appendChild(tr);
          });

          (this.tbody as HTMLElement).appendChild(frag);
        }

        updateFooterInfo(info: { visible: number; total: number; ms?: number }): void {
          const visible = info.visible;
          const total = info.total;
          const cols = this.columns.length;

          const rt = this.options.showRenderTime
            ? `<span>Render: <strong>${(info.ms || 0).toFixed(1)} ms</strong></span>`
            : "";

          (this.footer as HTMLElement).innerHTML = `
            <span><strong>${visible}</strong> of <strong>${total}</strong> records shown</span>
            <span>Columns: <strong>${cols}</strong></span>
            ${rt}
          `;
        }

        updateTable(): void {
          const t0 = performance.now();
          const rows = this.computeFilteredSortedData();
          const t1 = performance.now();

          const highlightTerm = this.globalSearchTerm || "";
          this.renderBodyRows(rows, highlightTerm);

          this.updateFooterInfo({ visible: rows.length, total: this.rawData.length, ms: (t1 - t0) });
          this.updateFilterIconClasses();
        }
      };
    }

    const id = containerEl.id || ("grid_" + Math.random().toString(36).slice(2, 9));
    containerEl.id = id;

    new D365Speedup.Helpers.InteractiveTable({
      containerId: id,
      data: Array.isArray(data) ? data : [],
      datasetName: datasetName || "Dataset",
      options: mergedOptions
    });
  }
};

// ============================================================================
// HANDLERS (UI and Event Handlers)
// ============================================================================
D365Speedup.Handlers = {

  setupMainTabs: function (): void {
    const tabs = document.querySelectorAll(".main-tab");

    const activateTab = (tabName: string) => {
      tabs.forEach((t) => t.classList.remove("active"));
      document
        .querySelector(`.main-tab[data-tab="${tabName}"]`)
        ?.classList.add("active");
    };

    tabs.forEach((tab) => {
      tab.addEventListener("click", async () => {
        const selected = (tab as HTMLElement).dataset.tab;

        activateTab(selected!);
        D365Speedup.DOM.contentTitle.innerHTML = "";

        if (selected === "favorites") {
          await D365Speedup.Handlers.renderFavoritesInMain();
          return;
        }

        if (selected === "quick") {
          await D365Speedup.Handlers.renderQuickAccess();
          return;
        }

        if (selected === "info") {
          const { contentArea, contentTitle } = D365Speedup.DOM;
          const manifest = chrome.runtime.getManifest();
          contentTitle.innerHTML = "";
          contentArea.innerHTML = `
            <div class="info-card">
              <div class="info-card-header">
                <img src="assets/icon128.png" class="info-card-icon-img" />
                <div>
                  <div class="info-card-title">${manifest.name}</div>
                  <div class="info-card-version">Version ${manifest.version}</div>
                </div>
              </div>
              <p class="info-card-tagline">Accelerate Dynamics 365 customization and development with smart tools and generators.</p>
              <div class="info-card-divider"></div>
              <div class="info-card-row"><span class="info-label">GitHub</span><a href="https://github.com/CS4OGGY/D365Speedup_ChromeExtension" target="_blank" class="info-card-link">github.com/CS4OGGY/D365Speedup_ChromeExtension</a></div>
              <div class="info-card-row"><span class="info-label">Feedback</span><a href="https://github.com/CS4OGGY/D365Speedup_ChromeExtension/issues" target="_blank" class="info-card-link">Report an issue on GitHub</a></div>
              <div class="info-card-divider"></div>
              <div class="info-card-section-title">What's New</div>
              <p class="info-card-whats-new">v1.0 - Initial release</p>
              <div class="info-card-divider"></div>
              <p class="info-card-privacy">Privacy: No data is collected or stored externally. All operations run directly in the active browser tab.</p>
            </div>
          `;
          return;
        }
      });
    });
  },

  renderPlaceholder: function (icon: string, text: string): void {
    const { contentArea, contentTitle } = D365Speedup.DOM;

    contentTitle.innerHTML = "";
    contentArea.innerHTML = `
      <div class="empty-state">
        <div class="empty-state-icon">${icon}</div>
        <p>${text}</p>
      </div>
    `;
  },

  setupSidebarToggle: function (): void {
    const { sidebar, burgerMenu, overlay } = D365Speedup.DOM;

    burgerMenu.addEventListener("click", (e: MouseEvent) => {
      e.stopPropagation();
      D365Speedup.Handlers.toggleSidebar(!sidebar.classList.contains("open"));
    });

    overlay.addEventListener("click", () => D365Speedup.Handlers.toggleSidebar(false));

    document.addEventListener("click", (e: MouseEvent) => {
      if (!sidebar.contains(e.target as Node) && !burgerMenu.contains(e.target as Node)) {
        D365Speedup.Handlers.toggleSidebar(false);
      }
    });
  },

  toggleSidebar: function (show: boolean): void {
    const { sidebar, overlay, burgerMenu } = D365Speedup.DOM;
    if (show) {
      sidebar.classList.add("open");
      overlay.classList.add("visible");
      burgerMenu.classList.add("hidden");
    } else {
      sidebar.classList.remove("open");
      overlay.classList.remove("visible");
      burgerMenu.classList.remove("hidden");
    }
  },

  loadConfig: async function (): Promise<void> {
    const url = chrome.runtime.getURL(D365Speedup.Constants.CONFIG_PATH);
    try {
      const res = await fetch(url);
      if (!res.ok) throw new Error(`Failed to load config.json (${res.status})`);
      D365Speedup.State.configData = await res.json();
    } catch (err: any) {
      D365Speedup.Helpers.showError(`Unable to load configuration: ${err.message}`);
    }
  },

  renderFavoritesInMain: async function (): Promise<void> {
    const { configData } = D365Speedup.State;
    const { contentArea, contentTitle } = D365Speedup.DOM;
    const favorites: string[] = await D365Speedup.Storage.getFavorites();

    contentTitle.innerHTML = ``;

    if (!favorites.length) {
      contentArea.innerHTML = `
        <div class="empty-state">
          <div class="empty-state-icon">⚡</div>
          <p>No favorite snippets yet.<br/>Use ★ to pin your most used tools!</p>
        </div>`;
      return;
    }

    const allSnippets = configData.flatMap((c: Category) => c.snippets);

    const cardsHTML = favorites.map((favId: string) => {
      const snippet = allSnippets.find((s: Snippet) => s.id === favId);
      if (!snippet) return "";

      return `
        <div class="fav-card" data-id="${snippet.id}">
          <div class="fav-header">
            <div class="fav-title-wrap">
              <span class="fav-title">${snippet.title}</span>
            </div>
            <span class="fav-remove" title="Remove from favorites">★</span>
          </div>
        </div>
      `;
    }).join("");

    contentArea.innerHTML = `<div class="favorites-grid">${cardsHTML}</div>`;

    document.querySelectorAll(".fav-card").forEach((card) => {
      const id = (card as HTMLElement).dataset.id!;
      card.addEventListener("click", () => {
        const snippet = allSnippets.find((s: Snippet) => s.id === id);
        if (snippet) D365Speedup.Handlers.openSnippet(snippet, card);
      });
    });

    document.querySelectorAll(".fav-remove").forEach((btn) => {
      btn.addEventListener("click", async (e: Event) => {
        e.stopPropagation();
        const id = (btn as HTMLElement).parentElement!.parentElement!.dataset.id!;
        await D365Speedup.Storage.toggleFavorite(id);
        D365Speedup.Helpers.showToast("Removed from favorites");
        await D365Speedup.Handlers.renderFavoritesInMain();
      });
    });
  },

  renderSidebar: async function (): Promise<void> {
    const { configData } = D365Speedup.State;
    const { navContainer } = D365Speedup.DOM;

    navContainer.innerHTML = "";

    const mainSection = document.createElement("div");
    mainSection.className = "nav-section";

    const mainTab = document.createElement("div");
    mainTab.className = "nav-tab";
    mainTab.innerHTML = `
      <span class="main-menu-icon">🏠</span>
      <span class="main-menu-text">Main Menu</span>
    `;
    mainTab.addEventListener("click", async () => {
      document.querySelectorAll(".nav-tab").forEach((t) => t.classList.remove("expanded", "active"));
      mainTab.classList.add("active");
      await D365Speedup.Handlers.renderFavoritesInMain();
      D365Speedup.Handlers.toggleSidebar(false);
    });
    mainSection.appendChild(mainTab);
    navContainer.appendChild(mainSection);

    const searchContainer = document.createElement("div");
    searchContainer.className = "search-container";
    searchContainer.innerHTML = `<input id="snippetSearch" type="text" class="search-box" placeholder="Search snippets..." />`;
    navContainer.appendChild(searchContainer);

    const allSnippets = configData.flatMap((cat: Category) =>
      cat.snippets.map((s: Snippet) => ({ ...s, categoryName: cat.categoryName, categoryIcon: cat.categoryIcon }))
    );

    const renderSearchResults = (query: string) => {
      const trimmed = query.trim().toLowerCase();
      const resultsSection = document.getElementById("search-results");
      if (resultsSection) resultsSection.remove();

      if (!trimmed) {
        configData.forEach((category: Category) => createCategory(category));
        return;
      }

      const matched = allSnippets.filter(
        (s: any) =>
          s.title.toLowerCase().includes(trimmed) ||
          s.description.toLowerCase().includes(trimmed)
      );

      const searchResults = document.createElement("div");
      searchResults.id = "search-results";
      searchResults.className = "nav-section";
      searchResults.innerHTML = `
        <div class="nav-tab active">
          <span class="tab-icon">🔍</span>
          <span class="tab-text">Search Results (${matched.length})</span>
        </div>
      `;

      const menu = document.createElement("div");
      menu.className = "nav-menu expanded";

      matched.forEach((snippet: any) => {
        const item = document.createElement("div");
        item.className = "menu-item";
        item.dataset.item = snippet.id;
        item.innerHTML = `<span class="sub-icon">${snippet.categoryIcon || "📁"}</span>${snippet.title}`;
        item.title = snippet.description;
        item.addEventListener("click", () => D365Speedup.Handlers.openSnippet(snippet, item));
        menu.appendChild(item);
      });

      searchResults.appendChild(menu);
      navContainer.appendChild(searchResults);
    };

    (searchContainer.querySelector("#snippetSearch") as HTMLInputElement).addEventListener("input", (e: Event) => {
      navContainer.querySelectorAll(".nav-section:not(:first-child):not(.search-container)").forEach((el) => el.remove());
      renderSearchResults((e.target as HTMLInputElement).value);
    });

    function createCategory(category: Category) {
      const section = document.createElement("div");
      section.className = "nav-section";

      const tab = document.createElement("div");
      tab.className = "nav-tab";
      tab.innerHTML = `
        <span class="tab-icon">${category.categoryIcon || "📁"}</span>
        <span class="tab-text">${category.categoryName}</span>
        <span class="tab-arrow">▼</span>
      `;

      const menu = document.createElement("div");
      menu.className = "nav-menu";

      category.snippets.forEach((snippet: Snippet) => {
        const item = document.createElement("div");
        item.className = "menu-item";
        item.dataset.item = snippet.id;
        item.innerHTML = `<span class="sub-icon">🪄</span>${snippet.title}`;
        item.title = snippet.description;
        item.addEventListener("click", () => D365Speedup.Handlers.openSnippet(snippet, item));
        menu.appendChild(item);
      });

      tab.addEventListener("click", () => D365Speedup.Handlers.toggleCategory(tab, menu));

      section.appendChild(tab);
      section.appendChild(menu);
      navContainer.appendChild(section);
    }

    configData.forEach((category: Category) => createCategory(category));
  },

  toggleCategory: function (tab: HTMLElement, menu: HTMLElement): void {
    const isExpanded = tab.classList.contains("expanded");
    document.querySelectorAll(".nav-tab").forEach((t) => t.classList.remove("expanded", "active"));
    document.querySelectorAll(".nav-menu").forEach((m) => m.classList.remove("expanded"));

    if (!isExpanded) {
      tab.classList.add("expanded", "active");
      menu.classList.add("expanded");
    }
  },

  openSnippet: async function (snippet: Snippet, menuItem: Element | null): Promise<void> {

    D365Speedup.State.selectedSnippet = snippet;
    D365Speedup.DOM.mainTabs.classList.add("hidden");
    D365Speedup.DOM.contentTitle.classList.remove("hidden");
    document.getElementById("topbarHome")?.classList.add("hidden");

    document.querySelectorAll(".menu-item").forEach((i) => i.classList.remove("active"));
    if (menuItem) menuItem.classList.add("active");

    const { contentTitle, contentArea } = D365Speedup.DOM;

    const isFav: boolean = await D365Speedup.Storage.isFavorite(snippet.id);
    contentTitle.innerHTML = `
      <span class="back-btn" id="backBtn"><svg viewBox="0 0 13 13" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path d="M8.5 2L4 6.5L8.5 11" stroke="#00b4d8" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"></path>
      </svg></span>
      <span style="vertical-align: top; padding-left: 8px;vertical-align: baseline;padding-right: 8px;">${snippet.title}</span>
      <span style="position:relative; display:inline-block; vertical-align: baseline;">
        <span id="favToggle"
              title="${isFav ? "Remove from favorites" : "Add to favorites"}"
              style="margin-left:8px; cursor:pointer;font-weight: bolder;vertical-align: baseline; font-size:17px; color:${isFav ? "#ffcc00" : "#777"};">
          ${isFav ? "★" : "☆"}
        </span>
        <span class="fav-hint hidden" id="favHint">Add to favorites</span>
      </span>
    `;

    if (!isFav) {
      const favHint = document.getElementById("favHint");
      const favToggleEl = document.getElementById("favToggle");
      if (favHint && favToggleEl) {
        const rect = favToggleEl.getBoundingClientRect();
        favHint.style.left = `${rect.right + 8}px`;
        favHint.style.top = `${rect.top + rect.height / 2}px`;
        favHint.style.transform = "translateY(-50%)";
        favHint.classList.remove("hidden");
        setTimeout(() => favHint.classList.add("fade-out"), 2500);
        setTimeout(() => favHint.classList.add("hidden"), 3000);
      }
    }

    document.getElementById("backBtn")!.addEventListener("click", async () => {
      D365Speedup.DOM.contentTitle.classList.add("hidden");
      document.getElementById("topbarHome")?.classList.remove("hidden");
      D365Speedup.DOM.mainTabs.classList.remove("hidden");

      document.querySelectorAll(".main-tab").forEach((t) => t.classList.remove("active"));
      document.querySelector('.main-tab[data-tab="favorites"]')?.classList.add("active");

      D365Speedup.DOM.contentTitle.innerHTML = "";
      await D365Speedup.Handlers.renderFavoritesInMain();
    });

    const hasInputs = Array.isArray(snippet.inputs) && snippet.inputs.length > 0;
    const context: any = await D365Speedup.Helpers.getContextInfo();

    if (!hasInputs) {
      contentArea.innerHTML = `
        <div class="snippet-panel">
          <div class="output-wrapper">
            <div id="outputText" class="output-text"></div>
          </div>
          <div class="loading-state">
            <div class="spinner"></div>
            <p>Loading...</p>
          </div>
        </div>
      `;

      const favIcon = document.getElementById("favToggle");
      if (favIcon) {
        favIcon.addEventListener("click", async (e: Event) => {
          e.stopPropagation();
          const updatedFavs: string[] = await D365Speedup.Storage.toggleFavorite(snippet.id);
          const isNowFav = updatedFavs.includes(snippet.id);
          favIcon.textContent = isNowFav ? "★" : "☆";
          (favIcon as HTMLElement).style.color = isNowFav ? "#ffcc00" : "#777";
          favIcon.title = isNowFav ? "Remove from favorites" : "Add to favorites";
          D365Speedup.Helpers.showToast(isNowFav ? "Added to favorites" : "Removed from favorites");
          await D365Speedup.Handlers.renderSidebar();
        });
      }

      try {
        await D365Speedup.Core.runSnippet();
      } finally {
        const spinner = contentArea.querySelector(".loading-state");
        if (spinner) {
          spinner.classList.add("fade-out");
          setTimeout(() => spinner.remove(), 10);
        }
      }

      D365Speedup.Handlers.toggleSidebar(false);
      return;
    }

    // === CASE 2: Snippet with inputs ===
    let inputsHTML = "";
    snippet.inputs!.forEach((input: SnippetInput) => {
      if (input.type === "hidden") return;

      const requiredMark = input.required
        ? `<span class="required-mark" style="color:#ff4b4b;font-weight:bold;margin-left:4px;">*</span>`
        : "";

      const isRadio = input.type === "radio";
      const controlHtml =
        input.type === "multi-line"
          ? `<textarea id="${input.id}" placeholder="${input.placeholder || ""}" rows="${input.areaRows || 4}" class="input-textarea"></textarea>`
          : input.type === "select"
            ? `<select id="${input.id}" class="input-select" data-type="select">
                  ${(input.options || []).map((opt: SnippetInputOption) => {
                    const val = (opt?.value ?? "").toString();
                    const lbl = (opt?.label ?? val).toString();
                    return `<option value="${D365Speedup.Helpers.escapeHTML(val)}">${D365Speedup.Helpers.escapeHTML(lbl)}</option>`;
                  }).join("")}
               </select>`
            : isRadio
              ? `<div id="${input.id}" class="radio-group">
                  ${(input.options || []).map((opt: SnippetInputOption) => {
                    const val = (opt?.value ?? "").toString();
                    const lbl = (opt?.label ?? val).toString();
                    const checked = (input.defaultValue === val) ? "checked" : "";
                    return `<label class="radio-label"><input type="radio" name="${input.id}" value="${D365Speedup.Helpers.escapeHTML(val)}" ${checked}> ${D365Speedup.Helpers.escapeHTML(lbl)}</label>`;
                  }).join("")}
                </div>`
              : `<input type="text" id="${input.id}" placeholder="${input.placeholder || ""}" class="input-text" data-type="${input.type}"${input.stripSpaces ? ` data-strip-spaces="true"` : ""}${input.appendSuffix ? ` data-append-suffix="${input.appendSuffix}"` : ""} />`;

      const showWhenAttr = input.showWhen
        ? ` data-show-when-input="${input.showWhen.inputId}" data-show-when-value="${input.showWhen.value}"`
        : "";
      const initialHidden = input.showWhen ? ` style="display:none"` : "";

      inputsHTML += `
        <div class="input-group" data-input-id="${input.id}"${showWhenAttr}${initialHidden}>
          <label${isRadio ? "" : ` for="${input.id}"`}>
            ${input.label}${requiredMark}
          </label>
          <div class="input-wrapper-dv">
            ${controlHtml}
            ${isRadio ? "" : `<span class="clear-btn" title="Clear">✖</span>`}
          </div>
        </div>
      `;
    });

    contentArea.innerHTML = `
      <div class="snippet-panel">
        <div class="tab-header">
          <div class="tab-item active" data-tab="input">Input</div>
          <div class="tab-item" data-tab="output">Output</div>
        </div>

        <div class="tab-content active" id="tab-input">
          <div class="snippet-inputs">${inputsHTML}</div>
          ${snippet.inputNote ? `<p class="snippet-note" style="margin:6px 0 0;padding-left:0;">${snippet.inputNote}</p>` : ""}
          <div class="btn-group">
            <button id="runBtn" class="btn-primary">Run</button>
          </div>
        </div>

        <div class="tab-content" id="tab-output">
          <div class="output-wrapper">
            ${snippet.outputType === "code"
              ? `<pre id="outputText" class="output-code"></pre>`
              : `<div id="outputText" class="output-text"></div>`
            }
            ${snippet.copyButtonRequired
              ? `<button id="hoverCopyBtn" class="hover-copy-btn" title="Copy" style="display:none;">📑Copy</button>`
              : ``
            }
          </div>
          ${snippet.note ? `<p class="snippet-note"${snippet.noteShowWhen ? ` data-show-when-input="${snippet.noteShowWhen.inputId}" data-show-when-value="${snippet.noteShowWhen.value}"` : ""}><em>${snippet.note}</em></p>` : ""}
        </div>
      </div>
    `;

    // --- Initialize context + autocomplete + clear buttons ---
    requestAnimationFrame(() => {
      snippet.inputs!.forEach((input: SnippetInput) => {
        const el = document.getElementById(input.id);
        if (!el) return;

        const wrapper = el.closest(".input-wrapper-dv") || el.closest(".input-wrapper");
        const clearBtn = wrapper?.querySelector(".clear-btn") as HTMLElement | null;

        // Attach autocomplete if defined
        if (input.type === "autocomplete" && input.dataSource) {
          D365Speedup.Helpers.attachAutoCompleteInput(el as HTMLInputElement, input.dataSource);
        }

        // Auto-populate context values
        if (input.autopopulate && input.populateFrom && context[input.populateFrom]) {
          const raw = input.stripSpaces
            ? (context[input.populateFrom] as string).replace(/\s+/g, "")
            : (context[input.populateFrom] as string);
          (el as HTMLInputElement).value = input.appendSuffix && !raw.endsWith(input.appendSuffix)
            ? raw + input.appendSuffix
            : raw;
        }

        // Clear button visibility & behavior (supports select)
        const refreshClear = () => {
          if (!clearBtn) return;
          if (input.type === "select") {
            clearBtn.style.display = ((el as HTMLSelectElement).selectedIndex > 0) ? "block" : "none";
          } else {
            clearBtn.style.display = (el as HTMLInputElement).value ? "block" : "none";
          }
        };

        if (clearBtn) {
          // for text/textarea
          el.addEventListener("input", refreshClear);
          // for select
          el.addEventListener("change", refreshClear);

          clearBtn.addEventListener("click", () => {
            if (input.type === "select") {
              (el as HTMLSelectElement).selectedIndex = 0;
            } else {
              (el as HTMLInputElement).value = "";
            }
            refreshClear();
            (el as HTMLElement).focus?.();
          });

          // initial
          refreshClear();
        }

        // Wire up showWhen: radio inputs drive visibility of dependent inputs
        if (input.type === "radio") {
          const updateDependents = () => {
            const checked = el.querySelector("input[type=\"radio\"]:checked") as HTMLInputElement | null;
            const val = checked?.value || "";
            // toggle input groups
            snippet.inputs!.forEach((dep: SnippetInput) => {
              if (dep.showWhen?.inputId === input.id) {
                const depGroup = contentArea.querySelector(`[data-input-id="${dep.id}"]`) as HTMLElement | null;
                if (depGroup) depGroup.style.display = (dep.showWhen.value === val) ? "" : "none";
              }
            });
            // toggle any other elements (e.g. notes) with data-show-when-input
            contentArea.querySelectorAll<HTMLElement>(`[data-show-when-input="${input.id}"]`).forEach(elem => {
              const required = elem.dataset.showWhenValue;
              elem.style.display = (required === val) ? "" : "none";
            });
          };
          el.querySelectorAll("input[type=\"radio\"]").forEach(r => r.addEventListener("change", updateDependents));
          updateDependents();
        }
      });
    });

    // Tab switching
    const tabItems = contentArea.querySelectorAll(".tab-item");
    const tabContents = contentArea.querySelectorAll(".tab-content");
    tabItems.forEach((tab) => {
      tab.addEventListener("click", () => {
        tabItems.forEach((t) => t.classList.remove("active"));
        tabContents.forEach((c) => c.classList.remove("active"));
        tab.classList.add("active");
        const target = contentArea.querySelector(`#tab-${(tab as HTMLElement).dataset.tab}`);
        if (target) target.classList.add("active");
      });
    });

    // Run button
    const runBtn = document.getElementById("runBtn");
    if (runBtn) {
      runBtn.addEventListener("click", async () => {

        const missing = snippet.inputs!.filter((input: SnippetInput) => {
          if (input.required) {
            const el = document.getElementById(input.id);
            if (!el) return true;
            // skip validation for inputs hidden by showWhen
            const group = el.closest("[data-input-id]") as HTMLElement | null;
            if (group && group.style.display === "none") return false;
            if (input.type === "select") return (el as HTMLSelectElement).selectedIndex < 0 || !(el as HTMLSelectElement).value;
            return !String((el as HTMLInputElement).value || "").trim();
          }
          return false;
        });

        if (missing.length > 0) {
          const names = missing.map((m: SnippetInput) => m.label).join(", ");
          D365Speedup.Helpers.showToast(`Required: ${names}`);
          missing.forEach((m: SnippetInput) => {
            const el = document.getElementById(m.id);
            if (el) {
              el.classList.add("input-error");
              setTimeout(() => el.classList.remove("input-error"), 1500);
            }
          });
          return;
        }

        const outputTab = document.querySelector('.tab-item[data-tab="output"]') as HTMLElement | null;
        if (outputTab) outputTab.click();

        const output = document.getElementById("outputText");
        if (output) {
          output.innerHTML = `
            <div class="loading-state">
              <div class="spinner"></div>
              <p>Running... please wait</p>
            </div>
          `;
        }

        const hoverCopyBtn = document.getElementById("hoverCopyBtn");
        if (hoverCopyBtn) {
          hoverCopyBtn.style.display = "none";
          hoverCopyBtn.onclick = () => {
            const text = document.getElementById("outputText")?.innerText || "";
            D365Speedup.Helpers.copyToClipboard(text);
          };
        }

        try {
          await D365Speedup.Core.runSnippet();
        } finally {
          const spinner = output?.querySelector(".loading-state");
          if (spinner) {
            spinner.classList.add("fade-out");
            setTimeout(() => spinner.remove(), 400);
          }
        }
      });
    }

    // Favorite toggle
    const favIcon = document.getElementById("favToggle");
    if (favIcon) {
      favIcon.addEventListener("click", async (e: Event) => {
        e.stopPropagation();
        const updatedFavs: string[] = await D365Speedup.Storage.toggleFavorite(snippet.id);
        const isNowFav = updatedFavs.includes(snippet.id);
        favIcon.textContent = isNowFav ? "★" : "☆";
        (favIcon as HTMLElement).style.color = isNowFav ? "#ffcc00" : "#777";
        favIcon.title = isNowFav ? "Remove from favorites" : "Add to favorites";

        D365Speedup.Helpers.showToast(isNowFav ? "Added to favorites" : "Removed from favorites");
        await D365Speedup.Handlers.renderSidebar();
      });
    }

    D365Speedup.Handlers.toggleSidebar(false);
  },

  renderQuickAccess: async function (): Promise<void> {
    const { contentArea, contentTitle } = D365Speedup.DOM;

    contentTitle.innerHTML = "";

    const context: any = await D365Speedup.Helpers.getContextInfo();
    const baseUrl = (context.clientUrl || "").replace(/\/$/, "");
    const environmentId = context.environmentId || "";

    if (!baseUrl) {
      contentArea.innerHTML = `
        <div class="empty-state">
          <div class="empty-state-icon">🔗</div>
          <p>Open a Dynamics 365 page to use Quick Access.</p>
        </div>
      `;
      return;
    }

    const oldSolutionsUrl =
      `${baseUrl}/tools/solution/home_solution.aspx?etc=7100&utm_source=MakerPortal`;

    const newSolutionsUrl = environmentId
      ? `https://make.powerapps.com/environments/${environmentId}/solutions`
      : `https://make.powerapps.com`;

    const adminPowerAppsUrl = environmentId
      ? `https://admin.powerplatform.microsoft.com/manage/environments/environment/${environmentId}/hub`
      : `https://admin.powerplatform.microsoft.com`;

    const links: { title: string; url: string }[] = [
      { title: "Advanced Find (Classic)", url: `${baseUrl}/main.aspx?pagetype=advancedfind` },
      { title: "Solutions (Classic)", url: oldSolutionsUrl },
      { title: "Solutions (Modern)", url: newSolutionsUrl },
      { title: "Plugin Trace Logs", url: `${baseUrl}/main.aspx?etn=plugintracelog&pagetype=entitylist` },
      { title: "Data Import", url: `${baseUrl}/main.aspx?etc=4412&pagetype=entitylist&forceClassic=1` },
      { title: "Duplicate Detection Rules", url: `${baseUrl}/main.aspx?etn=duplicaterule&pagetype=entitylist` },
      { title: "Bulk Delete Jobs", url: `${baseUrl}/main.aspx?etn=bulkdeleteoperation&pagetype=entitylist` },
      { title: "Power Apps Admin", url: adminPowerAppsUrl },
      { title: "Known Issues", url: "https://admin.powerplatform.microsoft.com/support/knownIssues" }
    ];

    contentArea.innerHTML = `
      <div class="quick-grid">
        ${links.map((l) => `
          <div class="quick-card" data-url="${l.url}">
            <div class="quick-icon">🔗</div>
            <div class="quick-title">${l.title}</div>
          </div>
        `).join("")}
      </div>
    `;

    contentArea.querySelectorAll(".quick-card").forEach((card) => {
      card.addEventListener("click", () => {
        const url = (card as HTMLElement).dataset.url;
        if (url) chrome.tabs.create({ url });
      });
    });
  },

  setupPanelToggle: async function (): Promise<void> {
    const btn = D365Speedup.DOM.panelToggleBtn;
    if (!btn) return;

    const inSidebar = D365Speedup.Helpers.isSidebarContext();

    btn.title = inSidebar ? "Switch to Popup" : "Switch to Side Panel";
    btn.classList.toggle("active-sidebar-mode", inSidebar);

    btn.addEventListener("click", async () => {
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

      if (D365Speedup.Helpers.isSidebarContext()) {
        // sidebar → popup: clear sidebar tab, restore popup, disable panel, open popup
        if (tab?.id) {
          await chrome.storage.session.remove("sidebarTabId");
          await chrome.action.setPopup({ tabId: tab.id, popup: "speedup.html" });
          await (chrome as any).sidePanel.setOptions({ tabId: tab.id, enabled: false });
        }
        try {
          await (chrome as any).action.openPopup({ windowId: tab?.windowId });
        } catch (_e) {}
        window.close();
      } else {
        // popup → sidebar: register this tab, disable panel for ALL other tabs immediately
        // (proactive disable prevents the sidebar content flashing on tab switch)
        if (tab?.id && tab?.windowId) {
          await chrome.storage.session.set({ sidebarTabId: tab.id });
          await (chrome as any).sidePanel.setOptions({ tabId: tab.id, path: "speedup.html?mode=sidebar", enabled: true });
          await chrome.action.setPopup({ tabId: tab.id, popup: "" });
          const allTabs = await chrome.tabs.query({});
          await Promise.all(
            allTabs
              .filter(t => t.id && t.id !== tab.id)
              .map(t => (chrome as any).sidePanel.setOptions({ tabId: t.id, enabled: false }).catch(() => {}))
          );
          await (chrome as any).sidePanel.open({ windowId: tab.windowId });
        }
        window.close();
      }
    });
  },
};

// ============================================================================
// STORAGE HANDLER
// ============================================================================
D365Speedup.Storage = {
  FAVORITE_KEY: "favorites",

  getFavorites: async function (): Promise<string[]> {
    return new Promise((resolve) => {
      chrome.storage.local.get([D365Speedup.Storage.FAVORITE_KEY], (result: Record<string, any>) => {
        resolve(result[D365Speedup.Storage.FAVORITE_KEY] || []);
      });
    });
  },

  saveFavorites: async function (favorites: string[]): Promise<void> {
    return new Promise((resolve) => {
      chrome.storage.local.set({ [D365Speedup.Storage.FAVORITE_KEY]: favorites }, resolve);
    });
  },

  toggleFavorite: async function (id: string): Promise<string[]> {
    const favorites: string[] = await D365Speedup.Storage.getFavorites();
    const idx = favorites.indexOf(id);
    if (idx >= 0) favorites.splice(idx, 1);
    else favorites.push(id);
    await D365Speedup.Storage.saveFavorites(favorites);
    return favorites;
  },

  isFavorite: async function (id: string): Promise<boolean> {
    const favorites: string[] = await D365Speedup.Storage.getFavorites();
    return favorites.includes(id);
  },

};

// ============================================================================
// CORE LOGIC
// ============================================================================
D365Speedup.Core = {

  runSnippet: async function (): Promise<void> {
    const snippet = D365Speedup.State.selectedSnippet;
    if (!snippet) return;

    const values: Record<string, string> = {};
    (snippet.inputs || []).forEach((i: SnippetInput) => {
      if (i.type === "hidden") { values[i.id] = i.defaultValue || ""; return; }
      const el = document.getElementById(i.id);
      if (!el) values[i.id] = "";
      else if (i.type === "select") values[i.id] = (el as HTMLSelectElement).value || "";
      else if (i.type === "radio") {
        const checked = el.querySelector("input[type=\"radio\"]:checked") as HTMLInputElement | null;
        values[i.id] = checked?.value || i.defaultValue || "";
      }
      else values[i.id] = (el as HTMLInputElement).value || "";
    });

    const output = document.getElementById("outputText");
    if (output) output.textContent = "Running...";

    try {
      const moduleUrl = chrome.runtime.getURL(snippet.script);

      if (snippet.runMode === "extension") {
        const module = await import(moduleUrl);
        const result = await module.run(values);
        D365Speedup.Core.renderOutput(result, snippet.outputType, snippet.outputSubType);
      } else if (snippet.runMode === "page") {
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
        const [{ result }] = await chrome.scripting.executeScript({
          target: { tabId: tab.id! },
          world: D365Speedup.Constants.PAGE_WORLD,
          func: async (url: string, values: Record<string, string>) => {
            const module = await import(url);
            return await module.run(values);
          },
          args: [moduleUrl, values],
        });
        D365Speedup.Core.renderOutput(result, snippet.outputType, snippet.outputSubType);
      }
    } catch (err: any) {
      D365Speedup.Core.renderOutput(`Error: ${err.message}`, "text", "");
    }
  },

  renderOutput: function (result: any, type: string | undefined, subType: string | undefined): void {
    const output = document.getElementById("outputText");
    if (!output) return;

    const snippet: Partial<Snippet> = D365Speedup.State.selectedSnippet || {};
    const showCopyBtn = snippet.copyButtonRequired !== false;

    if (type === "code" || type === "text") {
      const lang = subType || "plaintext";
      output.innerHTML = `<pre><code class="language-${lang}">${D365Speedup.Helpers.escapeHTML(result || "")}</code></pre>`;
      if ((window as any).Prism) (window as any).Prism.highlightAllUnder(output);
    } else if (type === "table" || type === "dynamic") {
      if (result && typeof result === "object" && result.__type === "interactiveTables") {
        const wrapId = "multiGrid_" + Math.random().toString(36).slice(2, 9);
        output.innerHTML = `<div id="${wrapId}"></div>`;

        requestAnimationFrame(() => {
          const wrap = document.getElementById(wrapId);
          if (!wrap) return;

          (result.tables || []).forEach((t: any) => {
            const gridId = "grid_" + Math.random().toString(36).slice(2, 9);
            const holder = document.createElement("div");
            holder.id = gridId;

            // small gap between tables
            holder.style.marginBottom = "10px";

            wrap.appendChild(holder);

            D365Speedup.Helpers.bindInteractiveGrid(
              holder,
              t.rows || [],
              t.datasetName || "Result",
              t.gridOptions || {}
            );
          });
        });

        return;
      }

      if (result && typeof result === "object" && result.__type === "collapsibleTables") {
        output.innerHTML = D365Speedup.Helpers.renderCollapsibleTables(result.sections || [], result.options || {});
        return;
      }

      output.innerHTML = result;
    }

    const tables = output.querySelectorAll(".advanced-table");
    tables.forEach((tbl) => {
      const ths = tbl.querySelectorAll("th.sortable");
      let sortCol: number | null = null;
      let sortDir = "asc";

      ths.forEach((th, i) => {
        th.addEventListener("click", () => {
          if (sortCol === i) sortDir = sortDir === "asc" ? "desc" : "asc";
          else { sortCol = i; sortDir = "asc"; }

          ths.forEach((t) => t.classList.remove("sort-asc", "sort-desc"));
          th.classList.add(sortDir === "asc" ? "sort-asc" : "sort-desc");

          const rowsArr = Array.from(tbl.querySelectorAll("tbody tr")) as HTMLTableRowElement[];
          rowsArr.sort((a, b) => {
            const av = (a.children[i] as HTMLElement).innerText.toLowerCase();
            const bv = (b.children[i] as HTMLElement).innerText.toLowerCase();
            if (av < bv) return sortDir === "asc" ? -1 : 1;
            if (av > bv) return sortDir === "asc" ? 1 : -1;
            return 0;
          });

          const tb = tbl.querySelector("tbody") as HTMLElement;
          tb.innerHTML = "";
          rowsArr.forEach((r) => tb.appendChild(r));
        });
      });
    });

    const hoverCopyBtn = document.getElementById("hoverCopyBtn");

    if (showCopyBtn) {
      if (hoverCopyBtn) hoverCopyBtn.style.display = "block";

      output.querySelectorAll(".copy-btn").forEach((btn) => {
        (btn as HTMLElement).textContent = "Copy";
        btn.addEventListener("click", async (e: Event) => {
          const el = e.currentTarget as HTMLElement;
          const text = el.getAttribute("data-copy") || "";
          try {
            await navigator.clipboard.writeText(text);
            el.textContent = "Copied!";
            el.classList.add("copied");
            setTimeout(() => {
              el.textContent = "Copy";
              el.classList.remove("copied");
            }, 1500);
          } catch (err) {
            el.textContent = "Error";
            console.error("Clipboard failed:", err);
          }
        });
      });
    } else {
      if (hoverCopyBtn) hoverCopyBtn.style.display = "none";
    }
  }
};

// ============================================================================
// ONLOAD INITIALIZATION...
// ============================================================================
D365Speedup.Handlers.Onload = async function (): Promise<void> {
  const inSidebar = D365Speedup.Helpers.isSidebarContext();

  if (!inSidebar) {
    document.body.classList.add("popup-mode");
  } else {
    // Restore per-tab popup when sidebar is closed (X button or navigation)
    chrome.tabs.query({ active: true, currentWindow: true }).then(([tab]) => {
      if (tab?.id) {
        const tabId = tab.id;
        window.addEventListener("pagehide", () => {
          chrome.storage.session.remove("sidebarTabId");
          chrome.action.setPopup({ tabId, popup: "speedup.html" });
          (chrome as any).sidePanel.setOptions({ tabId, enabled: false });
        });
      }
    });

    // Sidebar: set content-body height directly so CSS specificity fights don't matter
    const applySidebarHeight = () => {
      const topbar = document.querySelector(".topbar") as HTMLElement | null;
      const tabs = document.querySelector(".main-tabs") as HTMLElement | null;
      const cb = document.getElementById("contentBody");
      const ca = document.getElementById("contentArea");
      if (cb && topbar) {
        const used = topbar.offsetHeight + (tabs ? tabs.offsetHeight : 0);
        cb.style.height = `${window.innerHeight - used - 6}px`;
        if (ca) ca.style.height = "100%";
      }
    };
    requestAnimationFrame(applySidebarHeight);
    window.addEventListener("resize", applySidebarHeight);
  }

  // Open all links in a new tab
  D365Speedup.DOM.contentArea.addEventListener("click", (e: MouseEvent) => {
    const a = (e.target as Element).closest("a");
    if (a && a.href) {
      e.preventDefault();
      chrome.tabs.create({ url: a.href });
    }
  });

  await D365Speedup.Handlers.loadConfig();
  await D365Speedup.Handlers.renderSidebar();
  D365Speedup.Handlers.setupSidebarToggle();
  await D365Speedup.Handlers.renderFavoritesInMain();
  D365Speedup.Handlers.setupMainTabs();
  await D365Speedup.Handlers.setupPanelToggle();


  // Show burger hint briefly if no favorites
  const favorites: string[] = await D365Speedup.Storage.getFavorites();
  if (!favorites.length) {
    const hint = document.getElementById("burgerHint");
    if (hint) {
      hint.classList.remove("hidden");
      setTimeout(() => hint.classList.add("fade-out"), 2500);
      setTimeout(() => hint.classList.add("hidden"), 3000);
    }
  }
};

D365Speedup.Handlers.Onload();
