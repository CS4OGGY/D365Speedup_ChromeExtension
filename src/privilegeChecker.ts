

interface RunValues {
  email?: string;
  logicalName?: string;
}

export async function run({ email, logicalName }: RunValues = {}): Promise<any> {
  try {
    const win = window as any;
    const XrmCtx = win.Xrm || win.parent?.Xrm || win.top?.Xrm;
    const clientUrl: string | undefined = XrmCtx?.Utility?.getGlobalContext?.().getClientUrl?.();
    if (!clientUrl) throw new Error("Xrm not found. Run inside a model-driven app.");

    const userEmail = (email || "").trim();
    const table = (logicalName || "").trim();
    if (!userEmail || !table) throw new Error("Both user email and table logical name are required.");

    const API_VER = "v9.2";

    const RIGHTS: Array<{ name: string; keyword: string }> = [
      { name: "Read", keyword: "Read" },
      { name: "Write", keyword: "Write" },
      { name: "Append", keyword: "Append" },
      { name: "AppendTo", keyword: "AppendTo" },
      { name: "Create", keyword: "Create" },
      { name: "Delete", keyword: "Delete" },
      { name: "Share", keyword: "Share" },
      { name: "Assign", keyword: "Assign" }
    ];

    const DEPTH_LABEL: Record<number, string> = {
      0: "None",
      1: "User (Basic)",
      2: "Business Unit (Local)",
      4: "Parent:Child BU (Deep)",
      8: "Organization (Global)"
    };

    const q = (s: any): string => String(s).replace(/'/g, "''");
    const toGuid = (g: any): string => String(g || "").replace(/[{}]/g, "");
    const normDepth = (d: any): number => {
      const n = Number(d) || 0;
      if (n >= 8) return 8;
      if (n >= 4) return 4;
      if (n >= 2) return 2;
      if (n >= 1) return 1;
      return 0;
    };

    async function get(path: string): Promise<any> {
      const url = `${clientUrl}/api/data/${API_VER}/${path.replace(/^\//, "")}`;
      const res = await fetch(url, {
        method: "GET",
        headers: { Accept: "application/json", "OData-MaxVersion": "4.0", "OData-Version": "4.0" },
        credentials: "same-origin"
      });
      const txt = await res.text().catch(() => "");
      let data: any = {};
      try {
        data = txt ? JSON.parse(txt) : {};
      } catch {}
      if (!res.ok) throw new Error(data?.error?.message || `${res.status} ${res.statusText}`);
      return data;
    }

    async function tryGet(paths: string[]): Promise<any> {
      let last: any;
      for (const p of paths) {
        try {
          return await get(p);
        } catch (e: any) {
          last = e;
        }
      }
      throw last || new Error("All attempts failed.");
    }

    async function getInChunks(entitySet: string, select: string, filterExprs: string[], chunkSize: number = 10): Promise<any[]> {
      const out: any[] = [];
      for (let i = 0; i < filterExprs.length; i += chunkSize) {
        const chunk = filterExprs.slice(i, i + chunkSize);
        const data = await get(`${entitySet}?$select=${select}&$filter=${chunk.join(" or ")}`);
        out.push(...(data.value || []));
      }
      return out;
    }

    // -------------------- 1) Find user --------------------
    const userRes = await get(
      `systemusers?$select=systemuserid,fullname,internalemailaddress&$filter=internalemailaddress eq '${q(
        userEmail
      )}'&$top=5`
    );

    const user = userRes.value?.[0];
    if (!user) {
      return `<div class="error-box">❌ User not found for email: <b>${userEmail}</b></div>`;
    }
    const userId: string = toGuid(user.systemuserid);

    // -------------------- 2) Get table privileges (metadata) --------------------
    const privRes = await tryGet([
      `EntityDefinitions(LogicalName='${q(table)}')/Privileges`,
      `EntityDefinitions(LogicalName='${q(table)}')/Microsoft.Dynamics.CRM.Privileges`
    ]);

    const tablePrivs: any[] = privRes.value || [];
    if (!tablePrivs.length) {
      return `<div class="error-box">❌ No privileges returned for table: <b>${table}</b> (metadata restrictions or wrong table name)</div>`;
    }

    // Map right -> privilegeId (GUID)
    const privIdsByRight = new Map<string, string>();

    for (const r of RIGHTS) {
      // 1st choice: PrivilegeType exact match (if present)
      let match = tablePrivs.find((p: any) => (p.PrivilegeType || "").toLowerCase() === r.name.toLowerCase());

      // Fallback: stricter "prvRead" etc
      if (!match) {
        const re = new RegExp(`\\bprv${r.keyword}\\b`, "i");
        match = tablePrivs.find((p: any) => re.test(p.Name || ""));
      }

      // Last fallback: includes
      if (!match) match = tablePrivs.find((p: any) => (p.Name || "").includes(r.keyword));

      if (match?.PrivilegeId) privIdsByRight.set(r.name, toGuid(match.PrivilegeId).toLowerCase());
    }

    // -------------------- 3) Get user roles --------------------
    const userRolesRes = await tryGet([
      `systemuserrolescollection?$select=roleid&$filter=systemuserid eq ${userId}`,
      `systemuserroles?$select=roleid&$filter=systemuserid eq ${userId}`
    ]);

    const roleIds: string[] = (userRolesRes.value || []).map((r: any) => toGuid(r.roleid));
    if (!roleIds.length) {
      return `<div class="error-box">⚠️ User has no roles (or you don't have permission to read user roles).</div>`;
    }

    // -------------------- 4) Role names (chunked) --------------------
    const roleFilters = roleIds.map((id: string) => `roleid eq ${id}`);
    const roles = await getInChunks("roles", "roleid,name", roleFilters, 12);
    const roleNameById = new Map<string, string>((roles || []).map((r: any) => [toGuid(r.roleid), r.name]));

    // -------------------- 5) Role privileges (chunked + fallback entity set) --------------------
    let rolePrivs: any[];
    try {
      rolePrivs = await getInChunks(
        "roleprivilegescollection",
        "roleid,privilegeid,privilegedepthmask",
        roleIds.map((id: string) => `roleid eq ${id}`),
        8
      );
    } catch {
      rolePrivs = await getInChunks(
        "roleprivileges",
        "roleid,privilegeid,privilegedepthmask",
        roleIds.map((id: string) => `roleid eq ${id}`),
        8
      );
    }

    // -------------------- 6) Compute best depth --------------------
    const bestDepth = new Map<string, number>(RIGHTS.map((r) => [r.name, 0]));
    const grantingRoles = new Map<string, Set<string>>(RIGHTS.map((r) => [r.name, new Set<string>()]));

    for (const rp of rolePrivs || []) {
      const pid: string = toGuid(rp.privilegeid).toLowerCase();
      const depth: number = normDepth(rp.privilegedepthmask);
      const roleName: string = roleNameById.get(toGuid(rp.roleid)) || toGuid(rp.roleid);

      for (const r of RIGHTS) {
        if (privIdsByRight.get(r.name) === pid) {
          const current = bestDepth.get(r.name) || 0;
          if (depth > current) bestDepth.set(r.name, depth);
          grantingRoles.get(r.name)!.add(`${roleName} (depth=${depth})`);
        }
      }
    }

    // -------------------- 7) Grid output --------------------
    const rows = RIGHTS.map((r) => {
      const d = bestDepth.get(r.name) || 0;
      return {
        Privilege: r.name,
        Access: d ? "Yes" : "No",
        Level: DEPTH_LABEL[d] || `Mask ${d}`,
        "Roles Granting": [...grantingRoles.get(r.name)!].sort().join("; ")
      };
    });

    const infoRows = [
      { Key: "User", Value: `${user.fullname || ""}`.trim() || "(no name)" },
      { Key: "Email", Value: user.internalemailaddress || userEmail },
      { Key: "UserId", Value: userId },
      { Key: "Table", Value: table }

    ];

    // stash for debugging if you like
    (window as any).__privcheck = {
      user,
      table,
      rows,
      roleIds,
      privIdsByRight: Object.fromEntries(privIdsByRight)
    };

    return {
      __type: "interactiveTables",
      tables: [
        {
          datasetName: "ℹ️ Context",
          gridOptions: {
            allowHtml: true,
            enableSearch: false,
            enableFilters: false,
            enableSorting: false,
            enableResizing: true,
            collapsed: false,
            columnOrder: ["Key", "Value"]
          },
          rows: infoRows
        },
        {
          datasetName: `🔐 Privileges for ${table}`,
          gridOptions: {
            allowHtml: true,
            showRenderTime: true,
            enableSearch: true,
            enableFilters: true,
            enableSorting: true,
            enableResizing: true,
            collapsed: false,
            columnOrder: ["Privilege", "Access", "Level", "Roles Granting"]
          },
          rows
        }
      ],
      meta: {
        email: userEmail,
        table,
        userId,
        roles: roleIds.length
      }
    };
  } catch (e: any) {
    return `<div class="error-box">❌ Privilege check failed: ${e?.message || e}</div>`;
  }
}
