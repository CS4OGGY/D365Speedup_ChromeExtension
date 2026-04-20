
interface RunValues {
  logicalName?: string;
  flowFilter?: string; // "all" | "trigger" | "reference" (default)
}

export async function run({ logicalName, flowFilter = "reference" }: RunValues): Promise<any> {
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
    const toJSON = (v: any): any => {
      try { return typeof v === "string" ? JSON.parse(v) : v || {}; } catch { return {}; }
    };

    const clean2 = (v: any): string => (v == null ? "" : String(v).trim());

    const normalizeCsv = (v: any): string => {
      if (!v) return "";
      if (Array.isArray(v)) return v.map((x: any) => clean2(x)).filter(Boolean).join(", ");
      return String(v).split(",").map((x: string) => x.trim()).filter(Boolean).join(", ");
    };

    const getByPath = (obj: any, path: string): any => {
      if (!obj) return undefined;
      let cur = obj;
      for (const p of path.split(".")) {
        if (cur == null || typeof cur !== "object" || !(p in cur)) return undefined;
        cur = cur[p];
      }
      return cur;
    };

    const getParam = (params: any, ...names: string[]): any => {
      for (const name of names) {
        const direct = params?.[name];
        if (clean2(direct) !== "") return direct;
        const dotted = getByPath(params, name);
        if (clean2(dotted) !== "") return dotted;
        const slashVal = params?.[name.replace(/\./g, "/")];
        if (clean2(slashVal) !== "") return slashVal;
      }
      return "";
    };

    const pickFirst = (...vals: any[]): any => vals.find((v: any) => clean2(v) !== "") || "";

    const prettifyChangeType = (v: any): string => {
      const raw = clean2(v).toLowerCase();
      if (!raw) return "";
      const map: Record<string, string> = {
        "1": "Added", "2": "Removed", "3": "Modified", "4": "Added or Modified",
        "1,3": "Added or Modified", "3,1": "Added or Modified",
        "1,2,3": "Added, Modified or Removed",
        "create": "Added", "created": "Added", "add": "Added", "added": "Added",
        "update": "Modified", "updated": "Modified", "modify": "Modified", "modified": "Modified",
        "delete": "Removed", "deleted": "Removed",
      };
      const compact = raw.replace(/\s+/g, "");
      if (map[compact]) return map[compact];
      const parts = raw.split(",").map((x: string) => x.trim()).filter(Boolean);
      const hasAdd = parts.some((x: string) => ["1","4","create","created","add","added"].includes(x));
      const hasMod = parts.some((x: string) => ["3","4","update","updated","modify","modified"].includes(x));
      const hasDel = parts.some((x: string) => ["2","delete","deleted","remove","removed"].includes(x));
      if (hasAdd && hasMod && hasDel) return "Added, Modified or Removed";
      if (hasAdd && hasMod) return "Added or Modified";
      if (hasAdd && hasDel) return "Added or Removed";
      if (hasMod && hasDel) return "Modified or Removed";
      if (hasAdd) return "Added";
      if (hasMod) return "Modified";
      if (hasDel) return "Removed";
      return clean2(v);
    };

    const prettifyScope = (v: any): string => {
      const raw = clean2(v).toLowerCase();
      const map: Record<string, string> = {
        "1": "User", "2": "Business Unit", "3": "Parent and Child Business Units", "4": "Organization",
        "user": "User", "businessunit": "Business Unit",
        "parentchildbusinessunit": "Parent and Child Business Units", "organization": "Organization",
      };
      return map[raw] || clean2(v);
    };

    const parseDataverseTrigger = (trigger: any): Record<string, string> | null => {
      const inputs = trigger?.inputs || {};
      const params = inputs?.parameters || {};
      const host = inputs?.host || trigger?.host || {};
      const apiId = clean2(host?.apiId).toLowerCase();
      const connName = clean2(host?.connectionName).toLowerCase();
      const opId = clean2(trigger?.operationId || inputs?.operationId || host?.operationId).toLowerCase();
      const paramsText = JSON.stringify(params).toLowerCase();

      const isDataverse =
        apiId.includes("commondataserviceforapps") ||
        connName.includes("commondataserviceforapps") ||
        opId.includes("subscribewebhooktrigger") ||
        paramsText.includes("subscriptionrequest/entityname") ||
        paramsText.includes('"entityname"');

      if (!isDataverse) return null;

      const tableName = clean2(pickFirst(
        getParam(params, "subscriptionRequest.entityname"),
        getParam(params, "subscriptionRequest.tablename"),
        getParam(params, "entityName"),
        getParam(params, "tableName"),
      )).toLowerCase();

      const rawChangeType = pickFirst(
        getParam(params, "subscriptionRequest.message"),
        getParam(params, "subscriptionRequest.event"),
        getParam(params, "subscriptionRequest.sdkmessage"),
        getParam(params, "changeType"),
        getParam(params, "message"),
      );

      const rawScope = pickFirst(
        getParam(params, "subscriptionRequest.scope"),
        getParam(params, "scope"),
      );

      const selectColumns = normalizeCsv(pickFirst(
        getParam(params, "subscriptionRequest.filteringattributes"),
        getParam(params, "subscriptionRequest.selectcolumns"),
        getParam(params, "filteringattributes"),
        getParam(params, "filteringAttributes"),
        getParam(params, "selectColumns"),
      ));

      const filterRows = clean2(pickFirst(
        getParam(params, "subscriptionRequest.filterexpression"),
        getParam(params, "filterExpression"),
        getParam(params, "filterRows"),
      ));

      const triggerConditions = ([] as any[])
        .concat(trigger?.conditions || [])
        .concat(trigger?.runtimeConfiguration?.conditions || [])
        .concat(trigger?.metadata?.triggerConditions || [])
        .map((x: any) => (typeof x === "string" ? x : JSON.stringify(x)))
        .filter(Boolean)
        .join(" | ");

      return {
        "Table Name": tableName,
        "Change Type": prettifyChangeType(rawChangeType),
        "Scope": prettifyScope(rawScope),
        "Select Columns": selectColumns,
        "Filter Rows": filterRows,
        "Trigger Conditions": triggerConditions,
      };
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

    const principalCache = new Map<string, string>();
    const getPrincipalName = async (id: string): Promise<string> => {
      if (!id) return "";
      const key = id.toLowerCase();
      if (principalCache.has(key)) return principalCache.get(key)!;
      try {
        const u = await XrmContext.WebApi.retrieveRecord("systemuser", id, "?$select=fullname");
        principalCache.set(key, u.fullname || "");
      } catch {
        try {
          const t = await XrmContext.WebApi.retrieveRecord("team", id, "?$select=name");
          principalCache.set(key, t.name || "");
        } catch { principalCache.set(key, ""); }
      }
      return principalCache.get(key)!;
    };

    const flowsAll = await fetchPaged(
      "workflow",
      "?$select=workflowid,name,statecode,modifiedon,clientdata,category,_ownerid_value,_modifiedby_value&$filter=category eq 5&$orderby=modifiedon desc&$top=5000"
    );

    const seenFlows = new Set<string>();
    const rawFlowRows: any[] = [];

    for (const w of flowsAll) {
      const key = (w.name || "").toLowerCase();
      if (seenFlows.has(key)) continue;
      seenFlows.add(key);

      const cd = toJSON(w.clientdata);
      const def = toJSON(cd.definition || cd.properties?.definition || cd.Definition);
      const triggers = def?.triggers || def?.properties?.definition?.triggers || {};

      // Find best matching Dataverse trigger
      const allTrigInfos: any[] = [];
      for (const t of Object.values(triggers)) {
        const info = parseDataverseTrigger(t);
        if (info) allTrigInfos.push(info);
      }
      const trigInfo = allTrigInfos.find((x: any) =>
        x["Table Name"] === logicalName.toLowerCase() || x["Table Name"] === entitySet
      ) || allTrigInfos[0] || null;

      const triggerMatches = trigInfo && (trigInfo["Table Name"] === logicalName.toLowerCase() || trigInfo["Table Name"] === entitySet);

      const tabs = new Set<string>();
      collectTables(def, tabs);
      if (triggerMatches) tabs.add(entitySet);

      const hitAny = tabs.has(logicalName.toLowerCase()) || tabs.has(entitySet) || triggerMatches;
      const wantTrigger   = !flowFilter || flowFilter.includes("trigger");
      const wantReference = !flowFilter || flowFilter.includes("reference");
      const include = (wantTrigger && triggerMatches) || (wantReference && hitAny);
      if (!include) continue;

      rawFlowRows.push({ w, trigInfo, tabs });
    }

    const [ownerNames, modifiedByNames] = await Promise.all([
      Promise.all(rawFlowRows.map((r: any) => getPrincipalName(r.w._ownerid_value))),
      Promise.all(rawFlowRows.map((r: any) => getPrincipalName(r.w._modifiedby_value))),
    ]);

    const flowVizIcon = `<svg width="13" height="13" viewBox="0 0 13 13" fill="none" xmlns="http://www.w3.org/2000/svg" style="vertical-align:-2px"><circle cx="2" cy="6.5" r="1.6" fill="currentColor"/><circle cx="6.5" cy="2" r="1.6" fill="currentColor"/><circle cx="6.5" cy="11" r="1.6" fill="currentColor"/><circle cx="11" cy="6.5" r="1.6" fill="currentColor"/><line x1="3.55" y1="5.8" x2="5.4" y2="3.1" stroke="currentColor" stroke-width="1.1"/><line x1="3.55" y1="7.2" x2="5.4" y2="9.9" stroke="currentColor" stroke-width="1.1"/><line x1="7.6" y1="3.1" x2="9.45" y2="5.8" stroke="currentColor" stroke-width="1.1"/><line x1="7.6" y1="9.9" x2="9.45" y2="7.2" stroke="currentColor" stroke-width="1.1"/></svg>`;
    const flowRows: any[] = rawFlowRows.map((r: any, i: number) => ({
      Name: `<span class="flow-viz-icon" data-fid="${r.w.workflowid}" title="Open flow visualizer"><span class="fvi-badge">${flowVizIcon}</span>${r.w.name || ""}</span>`,
      State: wfState(r.w.statecode),
      "Change Type": r.trigInfo?.["Change Type"] || "",
      "Table Name": r.trigInfo?.["Table Name"] || "",
      "Scope": r.trigInfo?.["Scope"] || "",
      "Select Columns": r.trigInfo?.["Select Columns"] || "",
      "Filter Rows": r.trigInfo?.["Filter Rows"] || "",
      "Trigger Conditions": r.trigInfo?.["Trigger Conditions"] || "",
      "Tables": [...r.tabs].join(", "),
      Owner: ownerNames[i] || "",
      "Modified By": modifiedByNames[i] || "",
      "Modified On": fmtUtc(r.w.modifiedon),
    }));

    const flowDataMap: Record<string, any> = {};
    rawFlowRows.forEach((r: any) => {
      flowDataMap[r.w.workflowid] = { name: r.w.name, __paUrl: urlFlow(r.w.workflowid), ...toJSON(r.w.clientdata) };
    });

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
      __flowDataMap: flowDataMap,
      tables: [
        {
          datasetName: `🧠 Classic Processes (${procRows.length})`,
          gridOptions: { ...baseGridOptions, columnOrder: ["Name", "Category", "State", "Modified (UTC)"] },
          rows: procRows,
        },
        {
          datasetName: `⚡ Cloud Flows (${flowRows.length})`,
          gridOptions: { ...baseGridOptions, columnOrder: ["Name", "State", "Change Type", "Table Name", "Scope", "Select Columns", "Filter Rows", "Trigger Conditions", "Tables", "Owner", "Modified By", "Modified On"] },
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
