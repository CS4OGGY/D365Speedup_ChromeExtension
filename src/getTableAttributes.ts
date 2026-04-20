

interface RunValues {
  logicalName?: string;
}

export async function run({ logicalName }: RunValues = {}): Promise<any> {
  try {
    const logical = (logicalName || "").trim();
    if (!logical) throw new Error("Table logical name is required.");

    const win = window as any;
    const XrmCtx = win.Xrm || win.parent?.Xrm || win.top?.Xrm;
    if (!XrmCtx) throw new Error("Xrm context not found. Please open a Dynamics 365 CE page.");

    const clientUrl: string = XrmCtx.Utility.getGlobalContext().getClientUrl();
    const API_VER = "v9.2";

    // -------------------- Helpers --------------------
    const typeMap: Record<number, string> = {
      0: "Boolean", 1: "Customer", 2: "DateTime", 3: "Decimal", 4: "Double", 5: "Integer",
      6: "Lookup", 7: "Memo", 8: "Money", 9: "Owner", 10: "PartyList", 11: "Picklist",
      12: "State", 13: "Status", 14: "String", 15: "Uniqueidentifier", 16: "CalendarRules",
      17: "Virtual", 18: "BigInt", 19: "ManagedProperty", 20: "EntityName"
    };

    const getAttrType = (a: any): string => {
      const t = a.AttributeType;
      return (typeof t === "number") ? (typeMap[t] || String(t)) : (t || "");
    };

    const getLabel = (labelObj: any): string =>
      (labelObj && labelObj.UserLocalizedLabel && labelObj.UserLocalizedLabel.Label) || "";

    async function fetchJson(url: string): Promise<any> {
      const res = await fetch(url, {
        method: "GET",
        credentials: "include",
        headers: {
          Accept: "application/json",
          "OData-MaxVersion": "4.0",
          "OData-Version": "4.0"
        }
      });
      const text = await res.text().catch(() => "");
      if (!res.ok) throw new Error(text || `HTTP ${res.status} ${res.statusText}`);
      return text ? JSON.parse(text) : {};
    }

    // Limited concurrency mapper
    const mapLimit = async (items: any[], limit: number, fn: (item: any, i: number) => Promise<any>): Promise<any[]> => {
      const results = new Array(items.length);
      let idx = 0;
      const workers = Array.from({ length: Math.min(limit, items.length) }, async () => {
        while (idx < items.length) {
          const i = idx++;
          results[i] = await fn(items[i], i);
        }
      });
      await Promise.all(workers);
      return results;
    };

    // -------------------- Entity Display Name (optional) --------------------
    let entityDisplay: string = logical;
    try {
      const ent = await fetchJson(
        `${clientUrl}/api/data/${API_VER}/EntityDefinitions(LogicalName='${encodeURIComponent(logical)}')?$select=LogicalName,DisplayName`
      );
      entityDisplay =
        ent?.DisplayName?.UserLocalizedLabel?.Label ||
        ent?.LogicalName ||
        logical;
    } catch {
      // ignore
    }

    // -------------------- Base attribute list (paged) --------------------
    const fetchAttributesBase = async (): Promise<any[]> => {
      let url: string | null =
        `${clientUrl}/api/data/${API_VER}/EntityDefinitions(LogicalName='${encodeURIComponent(logical)}')/Attributes` +
        `?$select=LogicalName,SchemaName,DisplayName,AttributeType,IsCustomAttribute,IsPrimaryId,IsPrimaryName,IsValidForRead,IsValidForAdvancedFind`;

      const list: any[] = [];
      while (url) {
        const data = await fetchJson(url);
        (data.value || []).forEach((r: any) => { if (r.LogicalName) list.push(r); });
        url = data["@odata.nextLink"] || null;
      }
      return list;
    };

    // -------------------- OptionSet extraction (robust) --------------------
    const MAX_OPTIONS = 30;

    const formatOptions = (options: any[]): string => {
      if (!Array.isArray(options) || options.length === 0) return "";
      const shown = options.slice(0, MAX_OPTIONS).map((o: any) => `${o.Value}:${getLabel(o.Label)}`);
      const more = options.length > MAX_OPTIONS ? ` ...(+${options.length - MAX_OPTIONS} more)` : "";
      return shown.join(", ") + more;
    };

    // log-once flags (optional; avoids spamming in devtools)
    const logged: Record<string, boolean> = { picklist: false, status: false, state: false, boolean: false };

    const getOptionSetDetails = async (baseAttrUrl: string, kind: string): Promise<any> => {
      // kind: Picklist | Status | State | Boolean
      const typeName = `${kind}AttributeMetadata`;
      const castBase = `${baseAttrUrl}/Microsoft.Dynamics.CRM.${typeName}`;

      // 1) Try expand OptionSet
      try {
        const d = await fetchJson(`${castBase}?$expand=OptionSet`);
        const os = d.OptionSet;
        if (os && Array.isArray(os.Options)) return { options: os.Options };
        if (os && os.TrueOption && os.FalseOption) return { bool: os };
      } catch (e: any) {
        const key = kind.toLowerCase();
        if (!logged[key]) {
          // keep it quiet in extension; but safe to leave (remove if you want)
          // console.warn(`[attributes] OptionSet expand failed for ${kind} (fallbacks will be tried).`, String(e?.message || e).slice(0, 200));
          logged[key] = true;
        }
      }

      // 2) Try without expand (sometimes inline)
      try {
        const d = await fetchJson(castBase);
        const os = d.OptionSet;

        if (os && Array.isArray(os.Options)) return { options: os.Options };
        if (os && os.TrueOption && os.FalseOption) return { bool: os };

        // 3) Global option set fallback
        const name = os && (os.Name || os.name);
        const isGlobal = os && (os.IsGlobal ?? os.isGlobal);

        if (name && isGlobal) {
          const g = await fetchJson(
            `${clientUrl}/api/data/${API_VER}/GlobalOptionSetDefinitions(Name='${encodeURIComponent(name)}')?$expand=Options`
          );
          const opts = g.Options || (g.OptionSet && g.OptionSet.Options);
          if (Array.isArray(opts)) return { options: opts };
        }
      } catch {
        // swallow
      }

      return null;
    };

    // -------------------- Additional Details enrich --------------------
    const enrichAdditionalDetails = async (attr: any): Promise<string> => {
      const ln: string = attr.LogicalName;
      const type: string = getAttrType(attr);

      const parts: string[] = [];
      const add = (k: string, v: any) => {
        if (v === null || v === undefined || v === "") return;
        parts.push(`${k}=${v}`);
      };

      const baseAttrUrl =
        `${clientUrl}/api/data/${API_VER}/EntityDefinitions(LogicalName='${encodeURIComponent(logical)}')` +
        `/Attributes(LogicalName='${encodeURIComponent(ln)}')`;

      try {
        if (type === "String") {
          const d = await fetchJson(`${baseAttrUrl}/Microsoft.Dynamics.CRM.StringAttributeMetadata?$select=MaxLength,Format,FormatName`);
          add("MaxLength", d.MaxLength);
          add("Format", (d.Format && d.Format.Value) || d.Format);
          add("FormatName", (d.FormatName && d.FormatName.Value) || d.FormatName);
        } else if (type === "Memo") {
          const d = await fetchJson(`${baseAttrUrl}/Microsoft.Dynamics.CRM.MemoAttributeMetadata?$select=MaxLength,Format`);
          add("MaxLength", d.MaxLength);
          add("Format", (d.Format && d.Format.Value) || d.Format);
        } else if (type === "Lookup" || type === "Customer" || type === "Owner") {
          const d = await fetchJson(`${baseAttrUrl}/Microsoft.Dynamics.CRM.LookupAttributeMetadata?$select=Targets`);
          add("Targets", Array.isArray(d.Targets) ? d.Targets.join(",") : d.Targets);
        } else if (type === "Money") {
          const d = await fetchJson(`${baseAttrUrl}/Microsoft.Dynamics.CRM.MoneyAttributeMetadata?$select=Precision,MinValue,MaxValue`);
          add("Precision", d.Precision); add("Min", d.MinValue); add("Max", d.MaxValue);
        } else if (type === "Decimal") {
          const d = await fetchJson(`${baseAttrUrl}/Microsoft.Dynamics.CRM.DecimalAttributeMetadata?$select=Precision,MinValue,MaxValue`);
          add("Precision", d.Precision); add("Min", d.MinValue); add("Max", d.MaxValue);
        } else if (type === "Double") {
          const d = await fetchJson(`${baseAttrUrl}/Microsoft.Dynamics.CRM.DoubleAttributeMetadata?$select=Precision,MinValue,MaxValue`);
          add("Precision", d.Precision); add("Min", d.MinValue); add("Max", d.MaxValue);
        } else if (type === "Integer" || type === "BigInt") {
          const d = await fetchJson(`${baseAttrUrl}/Microsoft.Dynamics.CRM.IntegerAttributeMetadata?$select=MinValue,MaxValue`);
          add("Min", d.MinValue); add("Max", d.MaxValue);
        } else if (type === "DateTime") {
          const d = await fetchJson(`${baseAttrUrl}/Microsoft.Dynamics.CRM.DateTimeAttributeMetadata?$select=DateTimeBehavior,Format,CanChangeDateTimeBehavior`);
          const beh = (d.DateTimeBehavior && d.DateTimeBehavior.Value) || d.DateTimeBehavior;
          const fmt = (d.Format && d.Format.Value) || d.Format;
          add("Behavior", beh); add("Format", fmt);
          add("CanChangeBehavior", d.CanChangeDateTimeBehavior);
        } else if (type === "Picklist") {
          const os = await getOptionSetDetails(baseAttrUrl, "Picklist");
          if (os?.options) add("Options", formatOptions(os.options));
        } else if (type === "Status") {
          const os = await getOptionSetDetails(baseAttrUrl, "Status");
          if (os?.options) add("Options", formatOptions(os.options));
        } else if (type === "State") {
          const os = await getOptionSetDetails(baseAttrUrl, "State");
          if (os?.options) add("Options", formatOptions(os.options));
        } else if (type === "Boolean") {
          const os = await getOptionSetDetails(baseAttrUrl, "Boolean");
          if (os?.bool) {
            const t = os.bool.TrueOption;
            const f = os.bool.FalseOption;
            add("Options", `true:${getLabel(t.Label) || "true"}, false:${getLabel(f.Label) || "false"}`);
          }
        }
      } catch {
        // ignore per-attribute failures
      }

      return parts.join("; ");
    };

    // -------------------- Execute --------------------
    const t0 = performance.now();
    const attrs = await fetchAttributesBase();
    const baseMs = performance.now() - t0;

    const t1 = performance.now();
    const details = await mapLimit(attrs, 6, enrichAdditionalDetails);
    const enrichMs = performance.now() - t1;

    const rows = attrs.map((a: any, i: number) => ({
      "Logical Name": a.LogicalName || "",
      "Display Name": a.DisplayName?.UserLocalizedLabel?.Label || "",
      Type: getAttrType(a),
      "Additional Details": details[i] || "",
      Schema: a.SchemaName || "",
      Custom: !!a.IsCustomAttribute,
      "Primary Id": !!a.IsPrimaryId,
      "Primary Name": !!a.IsPrimaryName,
      "Valid For Read": a.IsValidForRead,
      Searchable: a.IsValidForAdvancedFind?.Value ?? a.IsValidForAdvancedFind,
    }));

    rows.sort((x: any, y: any) => String(x["Logical Name"]).localeCompare(String(y["Logical Name"])));

    return {
      __type: "interactiveTables",
      meta: {
        baseRetrievedMs: Math.round(baseMs),
        enrichRetrievedMs: Math.round(enrichMs),
        totalAttributes: rows.length
      },
      tables: [
        {
          datasetName: `🔖 ${entityDisplay} (${logical})`,
          gridOptions: {
            allowHtml: true,
            showRenderTime: true,
            enableSearch: true,
            enableFilters: true,
            enableSorting: true,
            enableResizing: true,
            collapsed: false,
            columnOrder: [
              "Logical Name",
              "Display Name",
              "Type",
              "Additional Details",
              "Schema",
              "Custom",
              "Primary Id",
              "Primary Name",
              "Valid For Read",
              "Searchable",
            ]
          },
          rows
        }
      ]
    };
  } catch (err: any) {
    return `<div class="error-box">❌ Error: ${err?.message || err}</div>`;
  }
}
