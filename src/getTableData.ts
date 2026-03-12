

interface RunValues {
  logicalName?: string;
  recordId?: string;
}

export async function run({ logicalName, recordId }: RunValues = {}): Promise<any> {
  try {
    const win = window as any;
    const XrmCtx = win.Xrm || win.parent?.Xrm || win.top?.Xrm;
    if (!XrmCtx) throw new Error("Xrm context not found. Please run on a Dynamics page.");

    const logical = (logicalName || "").trim();
    if (!logical) throw new Error("Table logical name is required.");

    const rawRecordId: string = (recordId || "").trim().replace(/[{}]/g, "");
    const isGuid = (v: string): boolean =>
      /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(v);

    const hasRecord = !!rawRecordId && isGuid(rawRecordId);

    const clientUrl: string = XrmCtx.Utility.getGlobalContext().getClientUrl();
    const API_VER = "v9.2";
    const CONCURRENCY = 6;
    const MAX_OPTIONS = 30;

    const typeMap: Record<number, string> = {
      0: "Boolean", 1: "Customer", 2: "DateTime", 3: "Decimal", 4: "Double", 5: "Integer",
      6: "Lookup", 7: "Memo", 8: "Money", 9: "Owner", 10: "PartyList", 11: "Picklist",
      12: "State", 13: "Status", 14: "String", 15: "Uniqueidentifier", 16: "CalendarRules",
      17: "Virtual", 18: "BigInt", 19: "ManagedProperty", 20: "EntityName"
    };

    const getAttrType = (a: any): string => {
      const t = a.AttributeType;
      return typeof t === "number" ? (typeMap[t] || String(t)) : (t || "");
    };

    const getLabel = (labelObj: any): string =>
      (labelObj && labelObj.UserLocalizedLabel && labelObj.UserLocalizedLabel.Label) || "";

    async function fetchJson(url: string, headers: Record<string, string> = {}): Promise<any> {
      const res = await fetch(url, {
        method: "GET",
        credentials: "include",
        headers: {
          Accept: "application/json",
          "OData-MaxVersion": "4.0",
          "OData-Version": "4.0",
          ...headers
        }
      });
      const text = await res.text().catch(() => "");
      if (!res.ok) throw new Error(text || `HTTP ${res.status} ${res.statusText}`);
      return text ? JSON.parse(text) : {};
    }

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

    const getEntitySetName = async (): Promise<string> => {
      const data = await fetchJson(
        `${clientUrl}/api/data/${API_VER}/EntityDefinitions(LogicalName='${encodeURIComponent(logical)}')?$select=EntitySetName`
      );
      if (!data.EntitySetName) throw new Error("EntitySetName not found for " + logical);
      return data.EntitySetName;
    };

    const fetchAttributesBase = async (): Promise<any[]> => {
      let url: string | null =
        `${clientUrl}/api/data/${API_VER}/EntityDefinitions(LogicalName='${encodeURIComponent(logical)}')/Attributes` +
        `?$select=LogicalName,SchemaName,AttributeType,AttributeOf,IsCustomAttribute,IsPrimaryId,IsPrimaryName,IsValidForRead`;

      const list: any[] = [];
      while (url) {
        const data = await fetchJson(url);
        (data.value || []).forEach((r: any) => {
          if (r.LogicalName) list.push(r);
        });
        url = data["@odata.nextLink"] || null;
      }
      return list;
    };

    // -------------------- OptionSet helpers (Additional Details) --------------------
    const formatOptions = (options: any[]): string => {
      if (!Array.isArray(options) || options.length === 0) return "";
      const shown = options.slice(0, MAX_OPTIONS).map((o: any) => `${o.Value}:${getLabel(o.Label)}`);
      const more = options.length > MAX_OPTIONS ? ` ...(+${options.length - MAX_OPTIONS} more)` : "";
      return shown.join(", ") + more;
    };

    const getOptionSetDetails = async (baseAttrUrl: string, kind: string): Promise<any> => {
      const castBase = `${baseAttrUrl}/Microsoft.Dynamics.CRM.${kind}AttributeMetadata`;

      try {
        const d = await fetchJson(`${castBase}?$expand=OptionSet`);
        const os = d.OptionSet;
        if (os && Array.isArray(os.Options)) return { options: os.Options };
        if (os && os.TrueOption && os.FalseOption) return { bool: os };
      } catch {}

      try {
        const d = await fetchJson(castBase);
        const os = d.OptionSet;

        if (os && Array.isArray(os.Options)) return { options: os.Options };
        if (os && os.TrueOption && os.FalseOption) return { bool: os };

        const name = os && (os.Name || os.name);
        const isGlobal = os && (os.IsGlobal ?? os.isGlobal);
        if (name && isGlobal) {
          const g = await fetchJson(
            `${clientUrl}/api/data/${API_VER}/GlobalOptionSetDefinitions(Name='${encodeURIComponent(name)}')?$expand=Options`
          );
          const opts = g.Options || (g.OptionSet && g.OptionSet.Options);
          if (Array.isArray(opts)) return { options: opts };
        }
      } catch {}

      return null;
    };

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
          const d = await fetchJson(`${baseAttrUrl}/Microsoft.Dynamics.CRM.StringAttributeMetadata?$select=MaxLength`);
          add("MaxLength", d.MaxLength);
        } else if (type === "Memo") {
          const d = await fetchJson(`${baseAttrUrl}/Microsoft.Dynamics.CRM.MemoAttributeMetadata?$select=MaxLength`);
          add("MaxLength", d.MaxLength);
        } else if (type === "Lookup" || type === "Customer" || type === "Owner") {
          const d = await fetchJson(`${baseAttrUrl}/Microsoft.Dynamics.CRM.LookupAttributeMetadata?$select=Targets`);
          add("Targets", Array.isArray(d.Targets) ? d.Targets.join(",") : d.Targets);
        } else if (type === "Money") {
          const d = await fetchJson(
            `${baseAttrUrl}/Microsoft.Dynamics.CRM.MoneyAttributeMetadata?$select=Precision,MinValue,MaxValue`
          );
          add("Precision", d.Precision);
          add("Min", d.MinValue);
          add("Max", d.MaxValue);
        } else if (type === "Decimal") {
          const d = await fetchJson(
            `${baseAttrUrl}/Microsoft.Dynamics.CRM.DecimalAttributeMetadata?$select=Precision,MinValue,MaxValue`
          );
          add("Precision", d.Precision);
          add("Min", d.MinValue);
          add("Max", d.MaxValue);
        } else if (type === "Double") {
          const d = await fetchJson(
            `${baseAttrUrl}/Microsoft.Dynamics.CRM.DoubleAttributeMetadata?$select=Precision,MinValue,MaxValue`
          );
          add("Precision", d.Precision);
          add("Min", d.MinValue);
          add("Max", d.MaxValue);
        } else if (type === "DateTime") {
          const d = await fetchJson(
            `${baseAttrUrl}/Microsoft.Dynamics.CRM.DateTimeAttributeMetadata?$select=DateTimeBehavior,Format`
          );
          const beh = (d.DateTimeBehavior && d.DateTimeBehavior.Value) || d.DateTimeBehavior;
          const fmt = (d.Format && d.Format.Value) || d.Format;
          add("Behavior", beh);
          add("Format", fmt);
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
      } catch {}

      return parts.join("; ");
    };

    // -------------------- Record fetch (for Value columns) --------------------
    const buildSelectColumns = (attrs: any[]): string[] => {
      const cols: string[] = [];
      for (const a of attrs) {
        const ln: string = a.LogicalName;
        if (!ln) continue;
        if (a.IsValidForRead === false) continue;

        const isDerived = a.AttributeOf !== null && a.AttributeOf !== undefined && a.AttributeOf !== "";
        if (isDerived) continue;

        const type = getAttrType(a);
        if (type === "Lookup" || type === "Customer" || type === "Owner") cols.push(`_${ln}_value`);
        else cols.push(ln);
      }
      return Array.from(new Set(cols));
    };

    const parseMissingProperty = (errText: string): string | null => {
      try {
        const j = JSON.parse(errText);
        const msg = j?.error?.message || "";
        const m = msg.match(/property named '([^']+)'/i);
        return m ? m[1] : null;
      } catch {
        const m = String(errText).match(/property named '([^']+)'/i);
        return m ? m[1] : null;
      }
    };

    const retrieveRecordAll = async (entitySetName: string, cols: string[]): Promise<{ record: any; skipped: string[] }> => {
      if (!hasRecord) return { record: null, skipped: [] };

      const remaining = cols.slice();
      const skipped: string[] = [];
      const merged: Record<string, any> = {};
      const chunkSize = 40;

      for (let i = 0; i < remaining.length; ) {
        const chunk = remaining.slice(i, i + chunkSize);
        let cur = chunk.slice();

        while (cur.length) {
          const select = cur.map(encodeURIComponent).join(",");
          const url = `${clientUrl}/api/data/${API_VER}/${entitySetName}(${rawRecordId})?$select=${select}`;
          try {
            const part = await fetchJson(url, { Prefer: 'odata.include-annotations="*"' });
            Object.assign(merged, part);
            break;
          } catch (e: any) {
            const raw = e?.message || String(e);
            const bad = parseMissingProperty(raw);
            if (!bad) throw e;

            skipped.push(bad);
            cur = cur.filter((c) => c !== bad);

            const idx = remaining.indexOf(bad);
            if (idx >= 0) remaining.splice(idx, 1);
            if (idx >= 0 && idx < i) i--;
          }
        }

        i += chunkSize;
      }

      return { record: merged, skipped };
    };

    const safeJson = (obj: any): string => {
      try {
        return JSON.stringify(obj);
      } catch {
        return "";
      }
    };

    const getValuePair = (attr: any, record: any): { display: any; raw: string } => {
      if (!record) return { display: "", raw: "" };

      const ln: string = attr.LogicalName;
      const type: string = getAttrType(attr);
      const fmt = (k: string) => record[`${k}@OData.Community.Display.V1.FormattedValue`];

      // Lookup-like
      if (type === "Lookup" || type === "Customer" || type === "Owner") {
        const k = `_${ln}_value`;
        const id = record[k];
        if (!id) return { display: "", raw: "" };

        const name = fmt(k) || "";
        const entity = record[`${k}@Microsoft.Dynamics.CRM.lookuplogicalname`] || "";
        return {
          display: name || id,
          raw: safeJson({ id, name: name || null, entity: entity || null })
        };
      }

      // DateTime
      if (type === "DateTime") {
        const iso = record[ln] ?? null;
        const formatted = fmt(ln) ?? null;
        return {
          display: formatted || (iso || ""),
          raw: safeJson({ iso, formatted })
        };
      }

      // Picklist/Status/State/Boolean
      if (type === "Picklist" || type === "Status" || type === "State" || type === "Boolean") {
        const value = record[ln];
        const label = fmt(ln);
        return {
          display: label ? `${label}` : (value ?? ""),
          raw: safeJson({ value: value ?? null, label: label ?? null })
        };
      }

      // MultiSelect heuristic
      {
        const rawVal = record[ln];
        const labels = fmt(ln);
        const looksMulti =
          Array.isArray(rawVal) ||
          (typeof rawVal === "string" &&
            rawVal.includes(",") &&
            typeof labels === "string" &&
            labels.includes(","));

        if (looksMulti) {
          const values = Array.isArray(rawVal)
            ? rawVal
            : (typeof rawVal === "string" ? rawVal.split(",").map((s: string) => s.trim()).filter(Boolean) : []);
          const labelArr = typeof labels === "string"
            ? labels.split(",").map((s: string) => s.trim()).filter(Boolean)
            : [];
          return {
            display: labels || (Array.isArray(rawVal) ? rawVal.join(",") : (rawVal ?? "")),
            raw: safeJson({ values, labels: labelArr.length ? labelArr : null })
          };
        }
      }

      // Default
      {
        const rawVal = record[ln];
        const formatted = fmt(ln);
        return {
          display: formatted !== undefined && formatted !== null && formatted !== "" ? String(formatted) : (rawVal ?? ""),
          raw: safeJson({ raw: rawVal ?? null, formatted: formatted ?? null })
        };
      }
    };

    // -------------------- Execute --------------------
    const attrs = await fetchAttributesBase();
    const details = await mapLimit(attrs, CONCURRENCY, enrichAdditionalDetails);

    let record: any = null;
    let skippedCols: string[] = [];

    if (rawRecordId && !hasRecord) {
      // Non-fatal: show grid but no values
      // (You can change this to throw if you prefer.)
      console.warn("Record id provided but not a valid GUID. Values will be blank.");
    }

    if (hasRecord) {
      const entitySetName = await getEntitySetName();
      const cols = buildSelectColumns(attrs);
      const out = await retrieveRecordAll(entitySetName, cols);
      record = out.record;
      skippedCols = out.skipped || [];
    }

    const rows = attrs
      .map((a: any, i: number) => {
        const pair = getValuePair(a, record);
        return {
          "Logical Name": a.LogicalName || "",
          "Schema": a.SchemaName || "",
          "Type": getAttrType(a),
          "Custom": !!a.IsCustomAttribute,
          "Primary Id": !!a.IsPrimaryId,
          "Primary Name": !!a.IsPrimaryName,
          "Valid For Read": a.IsValidForRead,
          "Additional Details": details[i] || "",
          "Value (Display)": pair.display,
          "Value (Raw)": pair.raw
        };
      })
     .sort((x: any, y: any) => {
  // 1) Primary Name first
  const pn = (y["Primary Name"] === true ? 1 : 0) - (x["Primary Name"] === true ? 1 : 0);
  if (pn !== 0) return pn;

  // 2) Then Primary Id
  const pid = (y["Primary Id"] === true ? 1 : 0) - (x["Primary Id"] === true ? 1 : 0);
  if (pid !== 0) return pid;

  // 3) Then alphabetical by Logical Name
  return String(x["Logical Name"]).localeCompare(String(y["Logical Name"]));
});

    // Optional second small table for skipped cols
    const tables: any[] = [
      {
        datasetName: `📋 Data: `,
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
             "Value (Display)",
            "Value (Raw)",
            "Schema",
            "Type",
            "Custom",
            "Primary Id",
            "Primary Name",
            "Valid For Read",
            "Additional Details"

          ]
        },
        rows
      }
    ];

    if (skippedCols.length) {
      tables.push({
        datasetName: `⚠️ Skipped non-selectable columns (${skippedCols.length})`,
        gridOptions: {
          allowHtml: true,
          enableSearch: true,
          enableSorting: true,
          collapsed: true,
          columnOrder: ["Column"]
        },
        rows: skippedCols.map((c: string) => ({ Column: c }))
      });
    }

    return {
      __type: "interactiveTables",
      tables,
      meta: {
        table: logical,
        recordId: hasRecord ? rawRecordId : null,
        attributes: attrs.length,
        valuesPopulated: hasRecord
      }
    };
  } catch (err: any) {
    return `<div class="error-box">❌ Error: ${err?.message || err}</div>`;
  }
}
