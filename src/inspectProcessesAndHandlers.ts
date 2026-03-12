
interface RunValues {
  logicalName?: string;
}

export async function run({ logicalName }: RunValues): Promise<any> {
  try {
    if (!logicalName) throw new Error("Table logical name is required.");

    const win = window as any;
    const XrmContext = win.Xrm || win.parent?.Xrm || win.top?.Xrm;
    if (!XrmContext) throw new Error("Xrm context not found. Please open a Dynamics 365 CE page.");

    const globalCtx = XrmContext.Utility.getGlobalContext();
    const client: string = globalCtx.getClientUrl();
    const envId: string = globalCtx.organizationSettings.bapEnvironmentId;

    // ---------- Helpers ----------
    const stripBraces = (g: any): string => String(g || "").replace(/[{}]/g, "");
    const link = (href: string, text: string): string => `<a href="${href}" target="_blank" class="link-cell">${text || ""}</a>`;

    const pad = (n: number): string => String(n).padStart(2, "0");
    const fmtUtc = (d: any): string => {
      if (!d) return "";
      const dt = new Date(d);
      return `${pad(dt.getUTCDate())}-${pad(dt.getUTCMonth() + 1)}-${dt.getUTCFullYear()} ${pad(
        dt.getUTCHours()
      )}:${pad(dt.getUTCMinutes())}`;
    };

    const wfCat: Record<number, string> = { 0: "Workflow", 1: "Dialog", 2: "Business Rule", 3: "Action", 4: "BPF", 5: "Cloud Flow" };
    const wfState = (v: any): string => (v === 1 ? "Activated" : v === 0 ? "Draft" : String(v));

    const fetchPaged = async (entity: string, query: string): Promise<any[]> => {
      const all: any[] = [];
      let next: string | null = query;
      do {
        const r: any = await XrmContext.WebApi.retrieveMultipleRecords(entity, next);
        all.push(...(r.entities || []));
        next = r.nextLink || null;
      } while (next);
      return all;
    };

    // URLs
    const urlWF = (id: any): string => `${client}/sfa/workflow/edit.aspx?id=${stripBraces(id)}`;
    const urlFlow = (id: any): string =>
      envId ? `https://make.powerautomate.com/environments/${envId}/flows/${stripBraces(id)}/details` : "";
    const urlForm = (formId: any, otc: any): string =>
      `${client}/main.aspx?etc=${otc}&extraqs=formtype%3dmain%26formId=${stripBraces(formId)}&pagetype=formeditor`;

    const urlPluginStep = (id: any): string =>
      `${client}/main.aspx?etn=sdkmessageprocessingstep&pagetype=entityrecord&id=${stripBraces(id)}`;
    const urlPluginType = (id: any): string =>
      `${client}/main.aspx?etn=plugintype&pagetype=entityrecord&id=${stripBraces(id)}`;
    const urlServiceEndpoint = (id: any): string =>
      `${client}/main.aspx?etn=serviceendpoint&pagetype=entityrecord&id=${stripBraces(id)}`;

    // ---------- Get table metadata ----------
    const metaRes = await fetch(
      `${client}/api/data/v9.2/EntityDefinitions(LogicalName='${encodeURIComponent(logicalName)}')?$select=ObjectTypeCode,EntitySetName,LogicalName`,
      { headers: { Accept: "application/json" }, credentials: "include" }
    );
    if (!metaRes.ok) throw new Error("Failed to retrieve table metadata.");
    const meta = await metaRes.json();
    const otc: any = meta.ObjectTypeCode;
    const entitySet: string = (meta.EntitySetName || `${logicalName}s`).toLowerCase();

    // =====================================================================
    // 1️⃣ CLASSIC PROCESSES
    // =====================================================================
    const keyOf = (w: any): string =>
      `${(w.name || "").trim().toLowerCase()}|${w.category}|${(w.primaryentity || "").toLowerCase()}`;
    const pickBest = (a: any, b: any): any =>
      a.statecode === 1 && b.statecode !== 1
        ? a
        : b.statecode === 1 && a.statecode !== 1
        ? b
        : new Date(b.modifiedon || 0) > new Date(a.modifiedon || 0)
        ? b
        : a;

    const allWF = await fetchPaged(
      "workflow",
      "?$select=workflowid,name,category,primaryentity,statecode,modifiedon&$orderby=modifiedon desc&$top=5000"
    );

    const procs = allWF.filter(
      (w: any) => (w.primaryentity || "").toLowerCase() === logicalName.toLowerCase() && w.category !== 5
    );

    const chosen = new Map<string, any>();
    for (const w of procs) {
      const k = keyOf(w);
      chosen.set(k, chosen.has(k) ? pickBest(chosen.get(k), w) : w);
    }

    const procRows = Array.from(chosen.values()).map((w: any) => ({
      Name: link(urlWF(w.workflowid), w.name),
      Category: wfCat[w.category] ?? w.category,
      State: wfState(w.statecode),
      "Modified (UTC)": fmtUtc(w.modifiedon),
    }));

    // =====================================================================
    // 2️⃣ CLOUD FLOWS
    // =====================================================================
    const flowsAll = await fetchPaged(
      "workflow",
      "?$select=workflowid,name,statecode,modifiedon,clientdata,category&$filter=category eq 5&$orderby=modifiedon desc&$top=5000"
    );

    const toJSON = (v: any): any => {
      try {
        return typeof v === "string" ? JSON.parse(v) : v || {};
      } catch {
        return {};
      }
    };

    const collectTables = (obj: any, out: Set<string>): void => {
      if (!obj || typeof obj !== "object") return;
      for (const [k, v] of Object.entries(obj)) {
        const kl = k.toLowerCase();
        if ((kl.includes("table") || kl.includes("entity")) && typeof v === "string" && (v as string).trim()) {
          out.add((v as string).toLowerCase());
        }
        collectTables(v, out);
      }
    };

    const triggerTargets = (def: any): string[] => {
      const hits = new Set<string>();
      const triggers = def?.triggers || def?.properties?.definition?.triggers || {};
      for (const t of Object.values(triggers)) {
        const inp = (t as any)?.inputs || {};
        const cand = [
          inp?.parameters?.entityName,
          inp?.parameters?.tableName,
          inp?.parameters?.entity,
          inp?.path,
        ].filter(Boolean);

        for (const s of cand) {
          const str = String(s).toLowerCase();
          if (str.includes(entitySet) || str.includes(logicalName)) hits.add(entitySet);
        }

        const tmp = new Set<string>();
        collectTables(t, tmp);
        if (tmp.has(logicalName.toLowerCase()) || tmp.has(entitySet)) hits.add(entitySet);
      }
      return [...hits];
    };

    const flowRows: any[] = [];
    const seenFlows = new Set<string>();

    for (const w of flowsAll) {
      const cd = toJSON(w.clientdata);
      const def = toJSON(cd.definition || cd.properties?.definition || cd.Definition);

      const trigHits = triggerTargets(def);
      const hitTrigger = trigHits.includes(entitySet);

      const tabs = new Set<string>();
      collectTables(def, tabs);
      if (hitTrigger) tabs.add(entitySet);

      const hitAny = tabs.has(logicalName.toLowerCase()) || tabs.has(entitySet);
      if (!hitAny) continue;

      const key = (w.name || "").toLowerCase();
      if (seenFlows.has(key)) continue;
      seenFlows.add(key);

      flowRows.push({
        Name: link(urlFlow(w.workflowid), w.name),
        State: wfState(w.statecode),
        Tables: [...tabs].join(", "),
        "Modified (UTC)": fmtUtc(w.modifiedon),
      });
    }

    // =====================================================================
    // 3️⃣ FORM JS HANDLERS
    // =====================================================================
    const xAll = (ctx: any, xpath: string): any[] => {
      const doc = ctx.ownerDocument || ctx;
      const out: any[] = [];
      const it = doc.evaluate(xpath, ctx, null, XPathResult.ORDERED_NODE_ITERATOR_TYPE, null);
      for (let n = it.iterateNext(); n; n = it.iterateNext()) out.push(n);
      return out;
    };
    const xOne = (ctx: any, xpath: string): any => {
      const doc = ctx.ownerDocument || ctx;
      return doc.evaluate(xpath, ctx, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue || null;
    };
    const getAttr = (el: any, names: string[]): string => {
      for (const n of names) {
        const v = el.getAttribute && el.getAttribute(n);
        if (v != null && v !== "") return v;
      }
      return "";
    };
    const cleanLib = (v: any): string => (v || "").replace(/^\s*\$webresource:/i, "").trim();

    const parser = new DOMParser();
    const handlers: any[] = [];

    const fxForms = `
      <fetch>
        <entity name='systemform'>
          <attribute name='name'/><attribute name='type'/><attribute name='formxml'/><attribute name='formid'/>
          <filter><condition attribute='objecttypecode' operator='eq' value='${otc}'/></filter>
        </entity>
      </fetch>`;

    const forms = await fetchPaged("systemform", `?fetchXml=${encodeURIComponent(fxForms)}&$top=5000`);

    const formTypeName = (t: any): string =>
      ({
        2: "Main",
        6: "Quick View",
        7: "Quick Create",
        8: "Mobile (deprecated)",
        11: "Card",
        12: "Main (Interactive Experience)",
      } as Record<number, string>)[t] || String(t);

    for (const f of forms) {
      if (!f.formxml) continue;

      const xml = parser.parseFromString(f.formxml, "text/xml");
      if (xml.getElementsByTagName("parsererror").length) continue;

      const fType = formTypeName(f.type);
      const formName: string = f.name || "";

      const events = xAll(xml, "//*[local-name()='Event'] | //*[local-name()='event']");
      for (const ev of events) {
        const eventName: string = (getAttr(ev, ["name"]) || "").toLowerCase();

        let control = "(form)";
        if (eventName === "onchange") {
          control =
            getAttr(ev, ["attribute", "attribname"]) ||
            (xOne(ev, "ancestor::*[local-name()='control'][1]")?.getAttribute("datafieldname") ||
              xOne(ev, "ancestor::*[local-name()='control'][1]")?.getAttribute("name") ||
              xOne(ev, "ancestor::*[local-name()='control'][1]")?.getAttribute("id") ||
              "(control)");
        } else if (eventName !== "onload" && eventName !== "onsave") {
          const ctrl = xOne(ev, "ancestor::*[local-name()='control'][1]");
          control = ctrl
            ? ctrl.getAttribute("datafieldname") || ctrl.getAttribute("name") || ctrl.getAttribute("id") || "(control)"
            : "(form)";
        }

        const handlerNodes = xAll(ev, ".//*[local-name()='Handler'] | .//*[local-name()='handler']");
        for (const h of handlerNodes) {
          const lib = cleanLib(getAttr(h, ["libraryName", "libraryname", "library"]));
          const fn = getAttr(h, ["functionName", "functionname", "function"]);
          const enabled = getAttr(h, ["enabled"]) || "true";
          const libUrl = lib ? `${client}/WebResources/${lib}` : "";

          handlers.push({
            Form: link(urlForm(f.formid, otc), formName),
            FormType: fType,
            Event: eventName,
            Control: control,
            Library: lib ? link(libUrl, lib) : "",
            Function: fn,
            Enabled: enabled,
          });
        }
      }
    }

    // =====================================================================
    // 4️⃣ PLUGIN STEPS + SERVICE ENDPOINTS
    // =====================================================================
    const stageLabel = (v: any): string =>
      ({
        5: "PreValidation",
        10: "PreValidation",
        15: "PreOperation (Legacy)",
        20: "PreOperation",
        25: "MainOperation",
        30: "PostOperation",
        35: "PostOperation",
        40: "PostOperation",
        50: "Internal",
      } as Record<number, string>)[v] || `Stage ${v}`;

    const modeLabel = (v: any): string => ({ 0: "Sync", 1: "Async" } as Record<number, string>)[v] || String(v);
    const statusLabel = (v: any): string => ({ 1: "Enabled", 2: "Disabled" } as Record<number, string>)[v] || String(v);

    const isServiceBus = (s: any): boolean => /servicebus/i.test(String(s || ""));
    const isMicrosoft = (s: any): boolean => /^(microsoft\.crm\.|microsoft\.dynamics\.|microsoft\.)/i.test(String(s || "").trim());

    const cleanPluginName2 = (name: any, fullName: any): string => {
      const s = (name || fullName || "").trim();
      return s.replace(/^Microsoft\.Crm\./i, "").replace(/^Microsoft\.Dynamics\./i, "").replace(/^Microsoft\./i, "").trim();
    };

    const fxSteps = `
      <fetch distinct="true" top="5000">
        <entity name="sdkmessageprocessingstep">
          <attribute name="sdkmessageprocessingstepid" />
          <attribute name="name" />
          <attribute name="stage" />
          <attribute name="mode" />
          <attribute name="statuscode" />
          <attribute name="rank" />
          <attribute name="filteringattributes" />
          <attribute name="modifiedon" />
          <attribute name="eventhandler" />

          <link-entity name="sdkmessage" from="sdkmessageid" to="sdkmessageid" alias="msg">
            <attribute name="name" alias="messageName" />
          </link-entity>

          <link-entity name="plugintype" from="plugintypeid" to="plugintypeid" alias="pt">
            <attribute name="name" alias="pluginTypeName" />
            <attribute name="typename" alias="pluginTypeFullName" />
            <attribute name="plugintypeid" alias="pluginTypeId" />
          </link-entity>

          <link-entity name="sdkmessagefilter" from="sdkmessagefilterid" to="sdkmessagefilterid" alias="f">
            <filter type="or">
              <condition attribute="primaryobjecttypecode" operator="eq" value="${otc}" />
              <condition attribute="secondaryobjecttypecode" operator="eq" value="${otc}" />
            </filter>
          </link-entity>
        </entity>
      </fetch>`;

    const steps = await fetchPaged("sdkmessageprocessingstep", `?fetchXml=${encodeURIComponent(fxSteps)}&$top=5000`);

    const serviceBusSteps = (steps || []).filter((s: any) => {
      const pName = s.pluginTypeName || s["pt.pluginTypeName"] || s["pt.name"] || "";
      const pFull = s.pluginTypeFullName || s["pt.pluginTypeFullName"] || s["pt.typename"] || "";
      return isServiceBus(pName) || isServiceBus(pFull);
    });

    const endpointIds = Array.from(new Set(serviceBusSteps.map((s: any) => stripBraces(s._eventhandler_value)).filter(Boolean)));

    const endpointMap = new Map<string, any>();
    for (const id of endpointIds) {
      try {
        const url = `${client}/api/data/v9.2/serviceendpoints(${id})?$select=serviceendpointid,name`;
        const r = await fetch(url, { headers: { Accept: "application/json" }, credentials: "include" });
        if (!r.ok) continue;
        endpointMap.set(id as string, await r.json());
      } catch {}
    }

    const serviceEndpointRows = serviceBusSteps.map((s: any) => {
      const stepId = s.sdkmessageprocessingstepid;
      const msg: string = s.messageName || s["msg.messageName"] || s["msg.name"] || "";
      const pName: string = s.pluginTypeName || s["pt.pluginTypeName"] || s["pt.name"] || "";
      const pFull: string = s.pluginTypeFullName || s["pt.pluginTypeFullName"] || s["pt.typename"] || "";
      const pluginLabel: string = cleanPluginName2(pName, pFull) || "ServiceBus";

      const ep = endpointMap.get(stripBraces(s._eventhandler_value));

      return {
        Step: link(urlPluginStep(stepId), s.name || "(no name)"),
        Plugin: pluginLabel,
        Endpoint: ep?.serviceendpointid ? link(urlServiceEndpoint(ep.serviceendpointid), ep.name || "") : "",
        Message: msg,
        Stage: stageLabel(s.stage),
        Mode: modeLabel(s.mode),
        Status: statusLabel(s.statuscode),
        Rank: s.rank ?? "",
        Filtering: s.filteringattributes || "",
      };
    });

    const pluginRows = (steps || [])
      .map((s: any) => {
        const stepId = s.sdkmessageprocessingstepid;
        const msg: string = s.messageName || s["msg.messageName"] || s["msg.name"] || "";

        const pName: string = s.pluginTypeName || s["pt.pluginTypeName"] || s["pt.name"] || "";
        const pFull: string = s.pluginTypeFullName || s["pt.pluginTypeFullName"] || s["pt.typename"] || "";
        const pId = s.pluginTypeId || s["pt.pluginTypeId"] || s["pt.plugintypeid"] || s.plugintypeid;

        if (isServiceBus(pName) || isServiceBus(pFull)) return null;
        if (isMicrosoft(pName) || isMicrosoft(pFull)) return null;

        const pluginLabel = cleanPluginName2(pName, pFull);

        return {
          Step: link(urlPluginStep(stepId), s.name || "(no name)"),
          Plugin: pId ? link(urlPluginType(pId), pluginLabel || "(plugin)") : (pluginLabel || ""),
          Message: msg,
          Stage: stageLabel(s.stage),
          Mode: modeLabel(s.mode),
          Status: statusLabel(s.statuscode),
          Rank: s.rank ?? "",
          Filtering: s.filteringattributes || "",
          "Modified (UTC)": fmtUtc(s.modifiedon),
        };
      })
      .filter(Boolean);

    // Shared grid options
    const baseGridOptions = {
      allowHtml: true,
      showRenderTime: true,
      enableSearch: true,
        collapsed: true,
      enableFilters: true,
      enableSorting: true,
      enableResizing: true,
    };

    // ✅ Return multiple grids (no wrapper sections)
    return {
      __type: "interactiveTables",
      tables: [
        {
          datasetName: `🧠 Classic Processes (${procRows.length})`,
          gridOptions: { ...baseGridOptions, columnOrder: ["Name", "Category", "State", "Modified (UTC)"] },
          rows: procRows,
        },
        {
          datasetName: `⚡ Cloud Flows (${flowRows.length})`,
          gridOptions: { ...baseGridOptions, columnOrder: ["Name", "State", "Tables", "Modified (UTC)"] },
          rows: flowRows,
        },
        {
          datasetName: `🧩 Form JavaScript Handlers (${handlers.length})`,
          gridOptions: { ...baseGridOptions, columnOrder: ["Form", "FormType", "Event", "Control", "Library", "Function", "Enabled"] },
          rows: handlers,
        },
        {
          datasetName: `🔌 Plugin Steps (Custom) (${pluginRows.length})`,
          gridOptions: { ...baseGridOptions, columnOrder: ["Step", "Plugin", "Message", "Stage", "Mode", "Status", "Rank", "Filtering", "Modified (UTC)"] },
          rows: pluginRows,
        },
        {
          datasetName: `🔌 Service Endpoints (${serviceEndpointRows.length})`,
          gridOptions: { ...baseGridOptions, columnOrder: ["Step", "Endpoint", "Message", "Stage", "Mode", "Status", "Rank", "Filtering"] },
          rows: serviceEndpointRows,
        },
      ],
    };
  } catch (err: any) {
    return `<div class="error-box">❌ Error: ${err.message}</div>`;
  }
}
