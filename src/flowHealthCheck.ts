
interface FlowRun {
    _workflow_value?: string;
    "_workflow_value@OData.Community.Display.V1.FormattedValue"?: string;
    "_ownerid_value@OData.Community.Display.V1.FormattedValue"?: string;
    status?: string;
    starttime?: string;
    endtime?: string;
    errormessage?: string;
    duration?: string;
    triggertype?: string;
}

interface RetrieveResult {
    entities: FlowRun[];
    "@odata.nextLink"?: string;
}

interface GroupedFlow {
    workflowId: string;
    flowName: string;
    runs: FlowRun[];
}

interface FlowRow {
    "Flow Name": string;
    "Total Runs": number;
    Succeeded: number;
    Failed: number;
    "Last Run Status": string;
    "Last Run Start": string;
    "Last Run End": string;
    "Last Failed Time": string;
    "Last Failed Owner": string;
    "Last Error": string;
    __lastFailedIso?: string;
}

interface GridOptions {
    allowHtml: boolean;
    showRenderTime: boolean;
    enableSearch: boolean;
    enableFilters: boolean;
    enableSorting: boolean;
    enableResizing: boolean;
    collapsed: boolean;
    columnOrder: string[];
}

interface TableResult {
    datasetName: string;
    gridOptions: GridOptions;
    rows: Omit<FlowRow, "__lastFailedIso">[];
}

interface InteractiveTables {
    __type: "interactiveTables";
    meta: { retrievedMs: number };
    tables: TableResult[];
}

const COLUMN_ORDER: string[] = [
    "Flow Name",
    "Total Runs",
    "Succeeded",
    "Failed",
    "Last Run Status",
    "Last Run Start",
    "Last Run End",
    "Last Failed Time",
    "Last Failed Owner",
    "Last Error",
];

const GRID_OPTIONS: GridOptions = {
    allowHtml: true,
    showRenderTime: true,
    enableSearch: true,
    enableFilters: true,
    enableSorting: true,
    enableResizing: true,
    collapsed: false,
    columnOrder: COLUMN_ORDER,
};

export async function run(): Promise<InteractiveTables | string> {
    try {
        const win = window as any;
        const XrmCtx = win.Xrm || win.parent?.Xrm || win.top?.Xrm;
        if (!XrmCtx) throw new Error("Xrm not available. Please run on a Dynamics 365 CE page.");

        const globalCtx = XrmCtx.Utility.getGlobalContext();
        const clientUrl: string = globalCtx.getClientUrl?.();
        if (!clientUrl) throw new Error("clientUrl not found.");

        const envId: string = globalCtx.organizationSettings?.bapEnvironmentId || "";

        const stripBraces = (g: unknown): string => String(g ?? "").replace(/[{}]/g, "");

        const link = (href: string, text: string): string =>
            href ? `<a href="${href}" target="_blank" class="link-cell">${text || ""}</a>` : (text || "");

        const toLocal = (dt: string | undefined): string => (dt ? new Date(dt).toLocaleString() : "—");
        const toIsoOrEmpty = (dt: string | undefined): string => (dt ? new Date(dt).toISOString() : "");

        const urlFlowDetails = (workflowId: string): string => {
            const id = stripBraces(workflowId);
            if (!id) return "";
            if (envId) return `https://make.powerautomate.com/environments/${stripBraces(envId)}/flows/${id}/details`;
            return `${clientUrl}/main.aspx?etn=workflow&pagetype=entityrecord&id=${id}`;
        };

        async function getFlowRuns(limit = 2000): Promise<FlowRun[]> {
            const results: FlowRun[] = [];
            const query =
                `?$select=duration,endtime,errormessage,_ownerid_value,starttime,status,triggertype,_workflow_value` +
                `&$orderby=starttime desc&$top=${Math.min(limit, 5000)}`;

            let res: RetrieveResult = await XrmCtx.WebApi.retrieveMultipleRecords("flowrun", query);
            results.push(...(res.entities || []));

            while (res["@odata.nextLink"] && results.length < limit) {
                const nextQuery = res["@odata.nextLink"].split("?")[1];
                res = await XrmCtx.WebApi.retrieveMultipleRecords("flowrun", nextQuery);
                results.push(...(res.entities || []));
            }

            return results.slice(0, limit);
        }

        const t0 = performance.now();
        const runs = await getFlowRuns(2000);
        const fetchMs = performance.now() - t0;

        if (!runs.length) {
            return {
                __type: "interactiveTables",
                meta: { retrievedMs: Math.round(fetchMs) },
                tables: [{ datasetName: "⚡ Flow Health Check (0)", gridOptions: GRID_OPTIONS, rows: [] }],
            };
        }

        const grouped = new Map<string, GroupedFlow>();
        for (const r of runs) {
            const workflowId = r._workflow_value || "";
            const flowName = r["_workflow_value@OData.Community.Display.V1.FormattedValue"] || "Unknown Flow";
            const key = workflowId || flowName;
            if (!grouped.has(key)) grouped.set(key, { workflowId, flowName, runs: [] });
            grouped.get(key)!.runs.push(r);
        }

        const rows: FlowRow[] = [];
        for (const { workflowId, flowName, runs: list } of grouped.values()) {
            const total = list.length;
            const succeeded = list.filter((x) => String(x.status || "").toLowerCase() === "succeeded").length;
            const failedRuns = list.filter((x) => String(x.status || "").toLowerCase() === "failed");
            const failed = failedRuns.length;
            const lastRun = list[0];
            const lastFailed = failedRuns[0] || null;

            rows.push({
                "Flow Name": link(urlFlowDetails(workflowId), flowName),
                "Total Runs": total,
                Succeeded: succeeded,
                Failed: failed,
                "Last Run Status": lastRun?.status ?? "—",
                "Last Run Start": toLocal(lastRun?.starttime),
                "Last Run End": toLocal(lastRun?.endtime),
                "Last Failed Time": toLocal(lastFailed?.starttime),
                "Last Failed Owner": lastFailed?.["_ownerid_value@OData.Community.Display.V1.FormattedValue"] ?? "",
                "Last Error": (lastFailed?.errormessage || "").slice(0, 200),
                __lastFailedIso: toIsoOrEmpty(lastFailed?.starttime),
            });
        }

        rows.sort((a, b) => {
            const at = Date.parse(a.__lastFailedIso || "") || 0;
            const bt = Date.parse(b.__lastFailedIso || "") || 0;
            if (bt !== at) return bt - at;
            const bf = Number(b.Failed || 0) - Number(a.Failed || 0);
            if (bf !== 0) return bf;
            return Number(b["Total Runs"] || 0) - Number(a["Total Runs"] || 0);
        });

        rows.forEach((r) => delete r.__lastFailedIso);

        return {
            __type: "interactiveTables",
            meta: { retrievedMs: Math.round(fetchMs) },
            tables: [{ datasetName: `⚡ Flow Health Check (${rows.length})`, gridOptions: GRID_OPTIONS, rows }],
        };
    } catch (err: any) {
        return `<div class="error-box">❌ Error: ${err?.message || err}</div>`;
    }
}
