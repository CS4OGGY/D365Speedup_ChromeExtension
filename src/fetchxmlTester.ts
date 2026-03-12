// ============================================================================
// FetchXML Tester (Full Retrieval + Attractive Grid)
// - Adds "Record" link column (opens record in new tab)
// - Hides the GUID primary id column
// - Uses popup.js existing renderer: { __type:"interactiveTables", tables:[...] }
// ============================================================================

interface RunValues {
  fetchXml?: string;
}
 
export async function run({ fetchXml }: RunValues = {}): Promise<any> {
  try {
    const win = window as any;
    const Xrm = win.Xrm || win.parent?.Xrm || win.top?.Xrm;
    if (!Xrm) throw new Error("❌ Xrm context not found. Please open inside a Dynamics 365 CE page.");
    if (!fetchXml?.trim()) throw new Error("⚠️ Please enter a FetchXML query.");

    const clientUrl: string = Xrm.Utility.getGlobalContext().getClientUrl();

    // --- Parse entity name ---
    const getEntityName = (xml: string): string | null => {
      try {
        const doc = new DOMParser().parseFromString(xml, "text/xml");
        return doc.getElementsByTagName("entity")[0]?.getAttribute("name") ?? null;
      } catch {
        const m = xml.match(/<\s*entity\b[^>]*\bname\s*=\s*(['"])([^'"]+)\1/i);
        return m ? m[2] : null;
      }
    };

    const entity = getEntityName(fetchXml);
    if (!entity) throw new Error("⚠️ Could not detect <entity name='...'> in FetchXML.");

    // --- Entity metadata: PrimaryIdAttribute + PrimaryNameAttribute ---
    const metaUrl =
      `${clientUrl}/api/data/v9.2/EntityDefinitions(LogicalName='${encodeURIComponent(entity)}')` +
      `?$select=PrimaryIdAttribute,PrimaryNameAttribute,LogicalName`;

    const metaRes = await fetch(metaUrl, {
      headers: { Accept: "application/json" },
      credentials: "include",
    });
    if (!metaRes.ok) throw new Error("Failed to read entity metadata (PrimaryIdAttribute).");
    const meta = await metaRes.json();

    const primaryIdAttr: string = meta?.PrimaryIdAttribute || "";
    const primaryNameAttr: string = meta?.PrimaryNameAttribute || "";

    const stripBraces = (g: any): string => String(g ?? "").replace(/[{}]/g, "");
    const urlRecord = (id: any): string =>
      `${clientUrl}/main.aspx?etn=${encodeURIComponent(entity)}&pagetype=entityrecord&id=${stripBraces(id)}`;

    // --- Fetch ALL paged results ---
    const all: any[] = [];
    let next: string | null = `?fetchXml=${encodeURIComponent(fetchXml)}`;

    const t0 = performance.now();
    while (next) {
      const res: any = await Xrm.WebApi.retrieveMultipleRecords(entity, next, 5000);
      all.push(...(res.entities || []));
      next = res.nextLink || res["@odata.nextLink"] || null;
    }
    const retrieveSeconds = ((performance.now() - t0) / 1000).toFixed(2);

    // --- Normalize + add Record link + hide GUID id column ---
    const rows = (all || []).map((o: any) => {
      const r: Record<string, any> = {};

      // Build record link (prefer primary name if present)
      const idVal = primaryIdAttr ? o?.[primaryIdAttr] : "";
      const nameVal = primaryNameAttr ? o?.[primaryNameAttr] : "";
      const linkText = nameVal || "Open";

      r["Record"] = idVal
        ? `<a href="${urlRecord(idVal)}" target="_blank" class="link-cell">${linkText}</a>`
        : "";

      for (const [k, v] of Object.entries(o || {})) {
        // hide the raw GUID primary id attribute column
        if (primaryIdAttr && k.toLowerCase() === primaryIdAttr.toLowerCase()) continue;

        r[k] =
          v && typeof v === "object"
            ? (v as any).name || (v as any).FormattedValue || JSON.stringify(v)
            : v;
      }

      return r;
    });

    return {
      __type: "interactiveTables",
      meta: {
        retrievedSeconds: Number(retrieveSeconds),
        total: all.length,
        entity,
      },
      tables: [
        {
          datasetName: "Result",
          gridOptions: {
            allowHtml: true,
            showRenderTime: true,
            enableSearch: true,
            enableFilters: true,
            enableSorting: true,
            enableResizing: true,
            collapsed: false,
          },
          rows,
        },
      ],
    };
  } catch (err: any) {
    return `<div class="error-box">❌ Error: ${err.message}</div>`;
  }
}
