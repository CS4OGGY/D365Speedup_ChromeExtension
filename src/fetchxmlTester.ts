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

    // --- Fetch results (respect count/page if set, else retrieve all pages) ---
    const doc2 = new DOMParser().parseFromString(fetchXml, "text/xml");
    const fetchEl = doc2.getElementsByTagName("fetch")[0];
    const fxCount = fetchEl?.getAttribute("count");
    const hasPaging = !!fxCount || fetchEl?.hasAttribute("page");
    const maxPageSize = fxCount ? parseInt(fxCount, 10) : 5000;

    const all: any[] = [];
    let next: string | null = `?fetchXml=${encodeURIComponent(fetchXml)}`;

    const t0 = performance.now();
    while (next) {
      const res: any = await Xrm.WebApi.retrieveMultipleRecords(entity, next, maxPageSize);
      all.push(...(res.entities || []));
      next = hasPaging ? null : (res.nextLink || res["@odata.nextLink"] || null);
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

// ── Builder metadata helpers ─────────────────────────────────────────────────

export async function fetchEntities(): Promise<{ name: string; display: string }[]> {
  try {
    const win = window as any;
    const Xrm = win.Xrm || win.parent?.Xrm || win.top?.Xrm;
    if (!Xrm) return [];

    win.__d365_fxb_entities ||= null;
    if (win.__d365_fxb_entities) return win.__d365_fxb_entities;

    const clientUrl: string = Xrm.Utility.getGlobalContext().getClientUrl();
    const headers = { Accept: "application/json", "OData-MaxVersion": "4.0", "OData-Version": "4.0" };
    const url = `${clientUrl}/api/data/v9.2/EntityDefinitions?$select=LogicalName,DisplayName&$filter=IsValidForAdvancedFind eq true`;
    const res = await fetch(url, { headers, credentials: "include" });
    if (!res.ok) return [];
    const data = await res.json();

    const entities: { name: string; display: string }[] = (data.value || [])
      .map((e: any) => ({
        name: e.LogicalName as string,
        display: (e.DisplayName?.UserLocalizedLabel?.Label || e.DisplayName?.LocalizedLabels?.[0]?.Label || e.LogicalName) as string,
      }))
      .sort((a: { display: string }, b: { display: string }) => a.display.localeCompare(b.display));

    win.__d365_fxb_entities = entities;
    return entities;
  } catch {
    return [];
  }
}

export async function fetchEntityMeta({ logicalName }: { logicalName: string }): Promise<{
  attrs: { n: string; d: string; t: string; targets?: string[] }[];
  rels: { name: string; display: string; fromAttr: string; toEntity: string; toAttr: string }[];
  views?: { id: string; name: string; type: string; fx: string }[];
  primaryName?: string;
  primaryId?: string;
}> {
  const empty = { attrs: [], rels: [], views: [], primaryName: '', primaryId: '' };
  try {
    const win = window as any;
    const Xrm = win.Xrm || win.parent?.Xrm || win.top?.Xrm;
    if (!Xrm) return empty;

    const clientUrl: string = Xrm.Utility.getGlobalContext().getClientUrl();
    const base = `${clientUrl}/api/data/v9.2`;
    const enc = encodeURIComponent(logicalName);
    const headers = { Accept: "application/json", "OData-MaxVersion": "4.0", "OData-Version": "4.0" };
    const opts: RequestInit = { headers, credentials: "include" };

    // Start lookup-targets + entity primary name fetch independently — failures must not block main data.
    const lkAttrsPromise: Promise<{ value: any[] }> = fetch(
      `${base}/EntityDefinitions(LogicalName='${enc}')/Attributes/Microsoft.Dynamics.CRM.LookupAttributeMetadata?$select=LogicalName,Targets`,
      opts
    ).then(r => r.ok ? r.json() : { value: [] }).catch(() => ({ value: [] }));

    const primaryMetaPromise: Promise<{ primaryName: string; primaryId: string }> = fetch(
      `${base}/EntityDefinitions(LogicalName='${enc}')?$select=PrimaryNameAttribute,PrimaryIdAttribute`,
      opts
    ).then(r => r.ok ? r.json() : {}).then((j: any) => ({
      primaryName: (j?.PrimaryNameAttribute as string) || '',
      primaryId: (j?.PrimaryIdAttribute as string) || '',
    })).catch(() => ({ primaryName: '', primaryId: '' }));

    // Critical fetches: attributes and relationships in parallel.
    const [attrsRes, otnRes, ntoRes] = await Promise.all([
      fetch(`${base}/EntityDefinitions(LogicalName='${enc}')/Attributes?$select=LogicalName,DisplayName,AttributeType,AttributeTypeName,IsValidForAdvancedFind,IsPrimaryName&$filter=IsValidForAdvancedFind/Value eq true`, opts),
      fetch(`${base}/EntityDefinitions(LogicalName='${enc}')/OneToManyRelationships?$select=SchemaName,ReferencingEntity,ReferencingAttribute,ReferencedAttribute`, opts),
      fetch(`${base}/EntityDefinitions(LogicalName='${enc}')/ManyToOneRelationships?$select=SchemaName,ReferencedEntity,ReferencingAttribute,ReferencedAttribute`, opts),
    ]);

    const attrsData = attrsRes.ok ? await attrsRes.json() : { value: [] };
    const otnData   = otnRes.ok   ? await otnRes.json()   : { value: [] };
    const ntoData   = ntoRes.ok   ? await ntoRes.json()   : { value: [] };

    // Fetch system views (savedqueries) and personal views (userqueries) in parallel.
    // returnedtypecode on savedquery/userquery is Edm.String — filter by logical name directly.
    const views: { id: string; name: string; type: string; fx: string }[] = [];
    try {
      const safeEnt = logicalName.replace(/'/g, "''");
      const vFilter = `?$select=savedqueryid,name,fetchxml&$filter=returnedtypecode eq '${safeEnt}' and statecode eq 0 and querytype eq 0&$orderby=name`;
      const uFilter = `?$select=userqueryid,name,fetchxml&$filter=returnedtypecode eq '${safeEnt}' and statecode eq 0 and querytype eq 0&$orderby=name`;
      const [sysResult, userResult] = await Promise.allSettled([
        (Xrm.WebApi.retrieveMultipleRecords("savedquery", vFilter) as Promise<any>),
        (Xrm.WebApi.retrieveMultipleRecords("userquery", uFilter) as Promise<any>),
      ]);
      if (sysResult.status === "fulfilled") {
        for (const v of ((sysResult.value as any).entities || [])) {
          if (v.fetchxml) views.push({ id: v.savedqueryid, name: v.name, type: "S", fx: v.fetchxml });
        }
      }
      if (userResult.status === "fulfilled") {
        for (const v of ((userResult.value as any).entities || [])) {
          if (v.fetchxml) views.push({ id: v.userqueryid, name: `★ ${v.name}`, type: "P", fx: v.fetchxml });
        }
      }
      // Sort: system views first, then personal, both alphabetical.
      views.sort((a, b) => (a.type === b.type ? a.name.localeCompare(b.name) : a.type === "S" ? -1 : 1));
    } catch { /* views are best-effort; attrs/rels continue regardless */ }

    // Await lookup targets — already running since before the main batch.
    const lkAttrsData = await lkAttrsPromise;
    const targetsMap = new Map<string, string[]>();
    for (const a of (lkAttrsData.value || [])) {
      if (a.LogicalName && Array.isArray(a.Targets) && a.Targets.length) {
        targetsMap.set(a.LogicalName as string, a.Targets as string[]);
      }
    }

    const { primaryName, primaryId } = await primaryMetaPromise;
    const attrs: { n: string; d: string; t: string; targets?: string[]; primary?: boolean }[] = (attrsData.value || []).map((a: any) => {
      const entry: { n: string; d: string; t: string; targets?: string[]; primary?: boolean } = {
        n: a.LogicalName as string,
        d: (a.DisplayName?.UserLocalizedLabel?.Label || a.DisplayName?.LocalizedLabels?.[0]?.Label || a.LogicalName) as string,
        t: (a.AttributeTypeName?.Value === "MultiSelectPicklistType" ? "MultiSelectPicklist" : a.AttributeType || "String") as string,
      };
      if (primaryName && a.LogicalName === primaryName) entry.primary = true;
      const tgts = targetsMap.get(a.LogicalName);
      if (tgts) entry.targets = tgts;
      return entry;
    });

    // 1:N — current entity is the "one"; related entity is the "many"
    const oneToMany = (otnData.value || []).map((r: any) => ({
      name: r.SchemaName as string,
      display: `${r.ReferencingEntity} (${r.ReferencingAttribute})`,
      fromAttr: r.ReferencedAttribute as string,
      toEntity: r.ReferencingEntity as string,
      toAttr: r.ReferencingAttribute as string,
    }));

    // N:1 — current entity is the "many"; related entity is the "one"
    const manyToOne = (ntoData.value || []).map((r: any) => ({
      name: r.SchemaName as string,
      display: `${r.ReferencedEntity} (${r.ReferencingAttribute})`,
      fromAttr: r.ReferencingAttribute as string,
      toEntity: r.ReferencedEntity as string,
      toAttr: r.ReferencedAttribute as string,
    }));

    const rels = [...manyToOne, ...oneToMany];
    return { attrs, rels, views, primaryName, primaryId };
  } catch {
    return empty;
  }
}

// ── Option set values ─────────────────────────────────────────────────────────
// Returns {v: integer value, l: display label} pairs for Picklist/State/Status/MultiSelectPicklist.
export async function fetchAttrOptions({ entityName, attrName }: { entityName: string; attrName: string }): Promise<{ v: number; l: string }[]> {
  try {
    const win = window as any;
    const Xrm = win.Xrm || win.parent?.Xrm || win.top?.Xrm;
    if (!Xrm) return [];

    // Virtual name attributes (e.g. statuscodename) must be resolved to their base attribute (statuscode)
    if (attrName.endsWith('name') && attrName.length > 4) attrName = attrName.slice(0, -4);

    const clientUrl: string = Xrm.Utility.getGlobalContext().getClientUrl();
    const base = `${clientUrl}/api/data/v9.2`;
    const enc = encodeURIComponent(entityName);
    const encAttr = encodeURIComponent(attrName);
    const headers = { Accept: "application/json", "OData-MaxVersion": "4.0", "OData-Version": "4.0" };
    const opts: RequestInit = { headers, credentials: "include" };

    // Try each possible metadata type in order — return the first that has options.
    const suffixes = [
      "Microsoft.Dynamics.CRM.PicklistAttributeMetadata",
      "Microsoft.Dynamics.CRM.MultiSelectPicklistAttributeMetadata",
      "Microsoft.Dynamics.CRM.StateAttributeMetadata",
      "Microsoft.Dynamics.CRM.StatusAttributeMetadata",
    ];
    for (const suffix of suffixes) {
      try {
        const res = await fetch(
          `${base}/EntityDefinitions(LogicalName='${enc}')/Attributes(LogicalName='${encAttr}')/${suffix}?$expand=OptionSet($select=Options)`,
          opts
        );
        if (!res.ok) continue;
        const data = await res.json();
        const options: any[] = data?.OptionSet?.Options;
        if (!Array.isArray(options) || !options.length) continue;
        return options
          .filter((o: any) => o.Value !== null && o.Value !== undefined)
          .map((o: any) => ({
            v: o.Value as number,
            l: (o.Label?.UserLocalizedLabel?.Label || o.Label?.LocalizedLabels?.[0]?.Label || String(o.Value)) as string,
          }));
      } catch { /* try next */ }
    }
    return [];
  } catch {
    return [];
  }
}

// ── Lookup record search ──────────────────────────────────────────────────────
// Searches records of targetEntity whose primary name contains searchTerm.
const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export async function fetchLookupRecords({ entityName, searchTerm, searchField }: { entityName: string; searchTerm: string; searchField?: string }): Promise<{ id: string; name: string; sub?: string; url?: string }[]> {
  try {
    const win = window as any;
    const Xrm = win.Xrm || win.parent?.Xrm || win.top?.Xrm;
    if (!Xrm) return [];

    const clientUrl: string = Xrm.Utility.getGlobalContext().getClientUrl();
    const base = `${clientUrl}/api/data/v9.2`;
    const enc = encodeURIComponent(entityName);

    const metaRes = await fetch(
      `${base}/EntityDefinitions(LogicalName='${enc}')?$select=PrimaryIdAttribute,PrimaryNameAttribute`,
      { headers: { Accept: "application/json" }, credentials: "include" }
    );
    if (!metaRes.ok) return [];
    const meta = await metaRes.json();
    const primaryId: string   = meta.PrimaryIdAttribute   || "";
    const primaryName: string = meta.PrimaryNameAttribute || "";
    if (!primaryId || !primaryName) return [];

    const safeQ  = searchTerm.replace(/'/g, "''");
    const isGuid = GUID_RE.test(searchTerm.trim());
    const field  = searchField || primaryName;
    const byOtherField = !!searchField && searchField !== primaryName;

    let filter: string;
    if (!safeQ) {
      filter = `?$top=10&$orderby=${primaryName}`;
    } else if (isGuid) {
      const cleanGuid = searchTerm.trim().toLowerCase().replace(/[{}]/g, "");
      filter = `?$filter=${primaryId} eq ${cleanGuid}&$top=1&$select=${primaryId},${primaryName}`;
    } else {
      const extra = byOtherField ? `,${field}` : "";
      filter = `?$filter=contains(${field},'${safeQ}')&$top=15&$orderby=${primaryName}&$select=${primaryId},${primaryName}${extra}`;
    }

    const res: any = await Xrm.WebApi.retrieveMultipleRecords(entityName, filter);
    return (res.entities || []).map((r: any) => {
      const id = String(r[primaryId] || "").replace(/[{}]/g, "");
      return {
        id,
        name: String(r[primaryName] || r[primaryId] || "?"),
        sub:  byOtherField && r[field] != null ? String(r[field]) : undefined,
        url:  id ? `${clientUrl}/main.aspx?etn=${encodeURIComponent(entityName)}&pagetype=entityrecord&id=${id}` : undefined,
      };
    });
  } catch {
    return [];
  }
}
