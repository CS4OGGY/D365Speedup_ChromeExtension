
interface RunValues {
  logicalName?: string;
  timeRange?: string;
}

export async function run({ logicalName, timeRange }: RunValues = {}): Promise<any> {
  try {
    if (!logicalName) throw new Error("Table logical name is required.");
    const range = (timeRange || "24h").toLowerCase();

    const win = window as any;
    const XrmCtx = win.Xrm || win.parent?.Xrm || win.top?.Xrm;
    if (!XrmCtx) throw new Error("Xrm context not found. Please open a Dynamics 365 CE page.");

    // ---- Compute cutoff ----
    const now = new Date();
    const cutoff = new Date(now);

    if (range === "1h") cutoff.setHours(cutoff.getHours() - 1);
    else if (range === "24h") cutoff.setHours(cutoff.getHours() - 24);
    else if (range === "7d") cutoff.setDate(cutoff.getDate() - 7);
    else if (range === "30d" || range === "1m") cutoff.setDate(cutoff.getDate() - 30);
    else throw new Error("Invalid timeRange. Use 1h / 24h / 7d / 30d.");

    const iso: string = cutoff.toISOString();

    const escXml = (s: any): string =>
      String(s ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&apos;");

    // ---- FetchXML paging helper ----
    const fetchAll = async (fxBase: string): Promise<any[]> => {
      const all: any[] = [];
      let page = 1;
      let cookie: string | null = null;

      while (true) {
        const doc = new DOMParser().parseFromString(fxBase, "text/xml");
        const fetchEl = doc.getElementsByTagName("fetch")[0];
        fetchEl.setAttribute("page", String(page));
        fetchEl.setAttribute("count", "5000");
        if (cookie) fetchEl.setAttribute("paging-cookie", cookie);

        const fx = new XMLSerializer().serializeToString(doc);

        const res = await XrmCtx.WebApi.retrieveMultipleRecords(
          "plugintracelog",
          `?fetchXml=${encodeURIComponent(fx)}`
        );

        all.push(...(res.entities || []));

        cookie = res.pagingCookie || null;
        if (!cookie || !(res.entities || []).length) break;

        page++;
        if (page > 200) break; // safety
      }

      return all;
    };

    // ---- FetchXML (primaryentity is STRING in your env) ----
    const fx = `
      <fetch version="1.0" mapping="logical" distinct="false">
        <entity name="plugintracelog">
          <attribute name="plugintracelogid" />
          <attribute name="typename" />
          <attribute name="messagename" />
          <attribute name="operationtype" />
          <attribute name="mode" />
          <attribute name="depth" />
          <attribute name="primaryentity" />
          <attribute name="performanceexecutionstarttime" />
          <attribute name="performanceexecutionduration" />
          <attribute name="createdon" />
          <attribute name="messageblock" />
          <filter type="and">
            <condition attribute="primaryentity" operator="eq" value="${escXml(logicalName)}" />
            <condition attribute="performanceexecutionstarttime" operator="on-or-after" value="${escXml(iso)}" />
          </filter>
          <order attribute="performanceexecutionstarttime" descending="true" />
        </entity>
      </fetch>
    `;

    const t0 = performance.now();
    const logs = await fetchAll(fx);
    const fetchMs = performance.now() - t0;

    const modeLabel = (v: any): string => ({ 0: "Sync", 1: "Async" } as Record<number, string>)[v] ?? String(v);
    const opLabel = (v: any): string =>
      ({
        0: "Unknown",
        1: "Create",
        2: "Update",
        3: "Delete",
        4: "Assign",
        5: "Associate",
        6: "Disassociate",
        7: "SetState",
        8: "Retrieve",
        9: "RetrieveMultiple",
      } as Record<number, string>)[v] ?? String(v);

    // Trace log record link
    const clientUrl: string = XrmCtx.Utility.getGlobalContext().getClientUrl();
    const stripBraces = (g: any): string => String(g || "").replace(/[{}]/g, "");
    const urlTrace = (id: any): string =>
      `${clientUrl}/main.aspx?etn=plugintracelog&pagetype=entityrecord&id=${stripBraces(id)}`;

    const rows = (logs || []).map((l: any) => ({
      Trace: l.plugintracelogid
        ? `<a href="${urlTrace(l.plugintracelogid)}" target="_blank" class="link-cell">Open</a>`
        : "",
      "Type Name": l.typename || "",
      Message: l.messagename || "",
      Operation: opLabel(l.operationtype),
      Mode: modeLabel(l.mode),
      Depth: l.depth ?? "",
      "Execution Start": l.performanceexecutionstarttime || "",
      "Duration (ms)": l.performanceexecutionduration ?? "",
    }));

    // ✅ Return as interactiveTables with a SINGLE grid
    return {
      __type: "interactiveTables",
      meta: { retrievedMs: Math.round(fetchMs), cutoffIso: iso, range },
      tables: [
        {
          datasetName: `🧾 Logs: ${logicalName}`,
          gridOptions: {
            allowHtml: true,
            showRenderTime: true,
            enableSearch: true,
            enableFilters: true,
            enableSorting: true,
            enableResizing: true,
            collapsed: false,
            columnOrder: [
             "Execution Start",
              "Trace",
              "Type Name",
              "Message",
              "Operation",
              "Mode",
              "Depth",
              "Duration (ms)",
            ],
          },
          rows,
        },
      ],
    };
  } catch (err: any) {
    const msg = String(err?.message || err || "")
      .replace(/correlationid\s*:\s*[0-9a-f-]+/ig, "")
      .trim();
    return `<div class="error-box">❌ Error: ${msg}</div>`;
  }
}
