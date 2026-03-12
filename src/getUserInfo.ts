
interface RunValues {
  username?: string;
}

export async function run({ username }: RunValues = {}): Promise<any> {
  try {
    const API_VER = "v9.2";

    const win = window as any;
    const XrmCtx = win.Xrm || win.parent?.Xrm || win.top?.Xrm;
    const globalCtx = XrmCtx?.Utility?.getGlobalContext?.();
    const clientUrl: string | undefined = globalCtx?.getClientUrl?.();
    if (!clientUrl) throw new Error("Xrm/clientUrl not found. Run inside a model-driven app.");

    // ✅ orgId (NOT environmentId) for PPAC redirect
    const orgId: string = globalCtx?.organizationSettings?.organizationId || "";

    const emailRaw = (username || "").trim();
    if (!emailRaw) return `<div class="no-data">No user selected.</div>`;

    const stripBraces = (g: any): string => String(g ?? "").replace(/[{}]/g, "");
    const odataEscape = (s: any): string => String(s ?? "").trim().replace(/'/g, "''");

    // -------------------- Record URLs --------------------
    const urlUser = (id: any): string =>
      `${clientUrl}/main.aspx?etn=systemuser&pagetype=entityrecord&id=${stripBraces(id)}`;

    const urlTeam = (id: any): string =>
      `${clientUrl}/main.aspx?etn=team&pagetype=entityrecord&id=${stripBraces(id)}`;

    const urlFsp = (id: any): string =>
      `${clientUrl}/main.aspx?etn=fieldsecurityprofile&pagetype=entityrecord&id=${stripBraces(id)}`;

    const urlRoleRecord = (id: any): string =>
      `${clientUrl}/main.aspx?etn=role&pagetype=entityrecord&id=${stripBraces(id)}`;

    // ✅ PPAC role editor link
    const urlRoleEditorPPAC = (roleId: any): string => {
      const oid = stripBraces(orgId);
      const rid = stripBraces(roleId);
      if (!oid || !rid) return "";
      return `https://admin.powerplatform.microsoft.com/settingredirect/${oid}/securityroles/${rid}/roleeditor`;
    };

    const link = (href: string, text: string): string =>
      href ? `<a href="${href}" target="_blank" class="link-cell">${text || ""}</a>` : (text || "");

    const linkRole = (roleId: any, name: string): string => {
      const ppac = urlRoleEditorPPAC(roleId);
      return ppac ? link(ppac, name) : link(urlRoleRecord(roleId), name);
    };

    // -------------------- Web API helpers --------------------
    async function apiGet(path: string): Promise<any> {
      const url = encodeURI(`${clientUrl}/api/data/${API_VER}${path}`);
      const res = await fetch(url, {
        method: "GET",
        credentials: "include",
        headers: {
          Accept: "application/json",
          "OData-MaxVersion": "4.0",
          "OData-Version": "4.0",
        },
      });
      if (!res.ok) {
        const text = await res.text().catch(() => "");
        throw new Error(`HTTP ${res.status} ${res.statusText}\n${text}`);
      }
      return res.json();
    }

    async function tryGet(fn: () => Promise<any>): Promise<any> {
      try {
        return await fn();
      } catch (e: any) {
        return { __error: e?.message || String(e) };
      }
    }

    const isSegmentNotFound = (msg: string): boolean => /Resource not found for the segment/i.test(msg || "");
    async function apiGetFirst(paths: string[]): Promise<any> {
      let lastErr: any = null;
      for (const p of paths) {
        try {
          return await apiGet(p);
        } catch (e: any) {
          lastErr = e;
          const msg = e?.message || String(e);
          if (isSegmentNotFound(msg)) continue;
          throw e;
        }
      }
      throw lastErr || new Error("All candidate requests failed.");
    }

    async function resolveEntityMeta(logicalName: string): Promise<{ entitySetName: string | null; primaryIdAttribute: string | null }> {
      try {
        const r = await apiGet(
          `/EntityDefinitions(LogicalName='${logicalName}')?$select=EntitySetName,PrimaryIdAttribute`
        );
        return {
          entitySetName: r?.EntitySetName || null,
          primaryIdAttribute: r?.PrimaryIdAttribute || null,
        };
      } catch {
        return { entitySetName: null, primaryIdAttribute: null };
      }
    }

    function teamTypeLabel(v: any): string {
      if (v === 0) return "Owner";
      if (v === 1) return "Access";
      return String(v);
    }

    // -------------------- User lookup --------------------
    const emailEsc = odataEscape(emailRaw);

    const userSelect = [
      "systemuserid",
      "fullname",
      "internalemailaddress",
      "mobilephone",
      "isdisabled",
      "_businessunitid_value",
      "_parentsystemuserid_value",
    ].join(",");

    const userResp = await tryGet(async () =>
      apiGet(`/systemusers?$select=${userSelect}&$filter=internalemailaddress eq '${emailEsc}'`)
    );
    if (userResp.__error) throw new Error(`User lookup failed:\n${userResp.__error}`);

    const user = userResp?.value?.[0];
    if (!user) return `<div class="no-data">No user found for exact email: <b>${emailRaw}</b></div>`;

    const userId: string = stripBraces(user.systemuserid);

    // -------------------- Business Unit + chain --------------------
    async function getBusinessUnit(buId: any): Promise<any> {
      if (!buId) return null;
      const id = stripBraces(buId);
      return apiGet(`/businessunits(${id})?$select=businessunitid,name,_parentbusinessunitid_value`);
    }

    async function getBuChain(startBu: any): Promise<any[]> {
      const chain: any[] = [];
      let cur = startBu;
      let guard = 0;
      while (cur && guard++ < 20) {
        chain.unshift({ businessunitid: cur.businessunitid, name: cur.name });
        const parentId = cur._parentbusinessunitid_value;
        if (!parentId) break;
        cur = await getBusinessUnit(parentId);
      }
      return chain;
    }

    const businessUnit = await tryGet(async () => getBusinessUnit(user._businessunitid_value));
    const businessUnitChain =
      businessUnit?.__error || !businessUnit ? businessUnit : await tryGet(async () => getBuChain(businessUnit));

    // -------------------- Manager --------------------
    const manager = await tryGet(async () => {
      const mid = user._parentsystemuserid_value;
      if (!mid) return null;
      return apiGet(`/systemusers(${stripBraces(mid)})?$select=systemuserid,fullname,internalemailaddress`);
    });

    // -------------------- Direct Roles --------------------
    const directRoles = await tryGet(async () => {
      const fx = `
        <fetch distinct="true">
          <entity name="role">
            <attribute name="name" />
            <attribute name="roleid" />
            <link-entity name="businessunit" from="businessunitid" to="businessunitid" alias="rbu">
              <attribute name="name" alias="roleBusinessUnit" />
            </link-entity>
            <link-entity name="systemuserroles" from="roleid" to="roleid" intersect="true">
              <filter>
                <condition attribute="systemuserid" operator="eq" value="${userId}" />
              </filter>
            </link-entity>
          </entity>
        </fetch>`;
      const r = await apiGet(`/roles?fetchXml=${encodeURIComponent(fx)}`);
      return (r.value || [])
        .map((x: any) => ({
          Name: linkRole(x.roleid, x.name),
          "Role BU": x.roleBusinessUnit || "",
        }))
        .sort((a: any, b: any) => String(a.Name || "").localeCompare(String(b.Name || "")));
    });

    // -------------------- Teams --------------------
    const teams = await tryGet(async () => {
      const fx = `
        <fetch distinct="true">
          <entity name="team">
            <attribute name="name" />
            <attribute name="teamid" />
            <attribute name="teamtype" />
            <link-entity name="businessunit" from="businessunitid" to="businessunitid" alias="tbu">
              <attribute name="name" alias="teamBusinessUnit" />
            </link-entity>
            <link-entity name="teammembership" from="teamid" to="teamid" intersect="true">
              <filter>
                <condition attribute="systemuserid" operator="eq" value="${userId}" />
              </filter>
            </link-entity>
          </entity>
        </fetch>`;
      const r = await apiGet(`/teams?fetchXml=${encodeURIComponent(fx)}`);
      return (r.value || [])
        .map((x: any) => ({
          Name: link(urlTeam(x.teamid), x.name),
          Type: teamTypeLabel(x.teamtype),
          "Team BU": x.teamBusinessUnit || "",
        }))
        .sort((a: any, b: any) => String(a.Name || "").localeCompare(String(b.Name || "")));
    });

    // -------------------- Roles via Teams --------------------
    const teamRoles = await tryGet(async () => {
      const fx = `
        <fetch distinct="true">
          <entity name="role">
            <attribute name="name" />
            <attribute name="roleid" />
            <link-entity name="businessunit" from="businessunitid" to="businessunitid" alias="rbu">
              <attribute name="name" alias="roleBusinessUnit" />
            </link-entity>
            <link-entity name="teamroles" from="roleid" to="roleid" intersect="true">
              <link-entity name="team" from="teamid" to="teamid" alias="t">
                <attribute name="name" alias="teamName" />
                <attribute name="teamid" alias="teamId" />
                <link-entity name="teammembership" from="teamid" to="teamid" intersect="true">
                  <filter>
                    <condition attribute="systemuserid" operator="eq" value="${userId}" />
                  </filter>
                </link-entity>
              </link-entity>
            </link-entity>
          </entity>
        </fetch>`;
      const r = await apiGet(`/roles?fetchXml=${encodeURIComponent(fx)}`);
      return (r.value || [])
        .map((x: any) => ({
          Team: x.teamId ? link(urlTeam(x.teamId), x.teamName || "") : (x.teamName || ""),
          Role: linkRole(x.roleid, x.name),
          "Role BU": x.roleBusinessUnit || "",
        }))
        .sort((a: any, b: any) => {
          const at = String(a.Team || "").toLowerCase();
          const bt = String(b.Team || "").toLowerCase();
          if (at < bt) return -1;
          if (at > bt) return 1;
          const ar = String(a.Role || "").toLowerCase();
          const br = String(b.Role || "").toLowerCase();
          if (ar < br) return -1;
          if (ar > br) return 1;
          return 0;
        });
    });

    // -------------------- Field Security Profiles --------------------
    const fieldSecurityProfiles = await tryGet(async () => {
      const fx = `
        <fetch distinct="true">
          <entity name="fieldsecurityprofile">
            <attribute name="name" />
            <attribute name="fieldsecurityprofileid" />
            <link-entity name="systemuserprofiles" from="fieldsecurityprofileid" to="fieldsecurityprofileid" intersect="true">
              <filter>
                <condition attribute="systemuserid" operator="eq" value="${userId}" />
              </filter>
            </link-entity>
          </entity>
        </fetch>`;
      const r = await apiGet(`/fieldsecurityprofiles?fetchXml=${encodeURIComponent(fx)}`);
      return (r.value || [])
        .map((x: any) => ({
          Name: link(urlFsp(x.fieldsecurityprofileid), x.name),
        }))
        .sort((a: any, b: any) => String(a.Name || "").localeCompare(String(b.Name || "")));
    });

    // -------------------- User Settings --------------------
    const userSettings = await tryGet(async () => {
      const meta = await resolveEntityMeta("usersettings");
      const setCandidates = Array.from(new Set([meta.entitySetName, "usersettingses", "usersettings"].filter(Boolean))) as string[];

      const wanted = ["timezonecode", "dateformatstring", "timeformatstring", "currencysymbol"];
      let attrs = [...wanted];

      async function fetchWithAttrs(attrList: string[]): Promise<any> {
        const pid = meta.primaryIdAttribute || "usersettingsid";
        const attrXml = [pid, ...attrList].map((a: string) => `<attribute name="${a}" />`).join("\n");

        const fx = `
          <fetch top="1">
            <entity name="usersettings">
              ${attrXml}
              <filter>
                <condition attribute="systemuserid" operator="eq" value="${userId}" />
              </filter>
            </entity>
          </fetch>`;
        return apiGetFirst(setCandidates.map((s: string) => `/${s}?fetchXml=${encodeURIComponent(fx)}`));
      }

      let found: any;
      for (let i = 0; i < 5; i++) {
        try {
          found = await fetchWithAttrs(attrs);
          break;
        } catch (e: any) {
          const msg = e?.message || String(e);
          const m = msg.match(/doesn't contain attribute with Name = '([^']+)'/i);
          if (m?.[1]) {
            const bad = m[1];
            attrs = attrs.filter((a: string) => a.toLowerCase() !== bad.toLowerCase());
            if (!attrs.length) throw e;
            continue;
          }
          throw e;
        }
      }

      const row0 = found?.value?.[0];
      if (!row0) return { note: "No usersettings row returned (permissions or no record)." };

      // timezoneName
      let timezoneName = "";
      if (typeof row0.timezonecode === "number") {
        const tzMeta = await resolveEntityMeta("timezonedefinition");
        const tzSetCandidates = Array.from(new Set([tzMeta.entitySetName, "timezonedefinitions"].filter(Boolean))) as string[];

        const tz = await apiGetFirst(
          tzSetCandidates.map(
            (s: string) =>
              `/${s}?$select=timezonecode,userinterfacename,standardname&$filter=timezonecode eq ${row0.timezonecode}`
          )
        );
        const tzRow = tz?.value?.[0];
        timezoneName = tzRow?.userinterfacename || tzRow?.standardname || "";
      }

      return {
        dateformatstring: row0.dateformatstring || "",
        timeformatstring: row0.timeformatstring || "",
        currencysymbol: row0.currencysymbol || "",
        timezoneName: timezoneName || "",
      };
    });

    // -------------------- Build rows (User table) --------------------
    const buName: string = businessUnit && !businessUnit.__error ? businessUnit.name : "";
    const buChainText: string = Array.isArray(businessUnitChain)
      ? businessUnitChain.map((x: any) => x.name).filter(Boolean).join(" → ")
      : "";

    const userRow = {
      fullname: link(urlUser(user.systemuserid), user.fullname || ""),
      email: user.internalemailaddress || "",
      isdisabled: user.isdisabled,
      mobilephone: user.mobilephone || "",
      businessUnit: buName,
      buChain: buChainText,
      manager: manager && !manager.__error ? (manager.fullname || "") : "",
    };

    // Shared grid options
    const baseGridOptions = {
      allowHtml: true,
      showRenderTime: false,
      enableSearch: false,
      enableFilters: true,
      enableSorting: true,
      enableResizing: true,
    };

    return {
      __type: "interactiveTables",
      tables: [
        {
          datasetName: `👤 User (key fields)`,
          gridOptions: {
            ...baseGridOptions,
            collapsed: true,
            columnOrder: ["fullname", "email", "isdisabled", "mobilephone", "businessUnit", "buChain", "manager"],
          },
          rows: [userRow],
        },
        {
          datasetName: `🕒 User settings`,
          gridOptions: {
            ...baseGridOptions,
            collapsed: true,
            columnOrder: ["dateformatstring", "timeformatstring", "currencysymbol", "timezoneName", "note", "Error"],
          },
          rows: userSettings?.__error ? [{ Error: userSettings.__error }] : [userSettings || {}],
        },
        {
          datasetName: `🎭 Direct roles (${(directRoles?.__error ? 0 : (directRoles || []).length)})`,
          gridOptions: {
            ...baseGridOptions,
            collapsed: true,
            columnOrder: ["Name", "Role BU", "Error"],
          },
          rows: directRoles?.__error ? [{ Error: directRoles.__error }] : (directRoles || []),
        },
        {
          datasetName: `👥 Teams (${(teams?.__error ? 0 : (teams || []).length)})`,
          gridOptions: {
            ...baseGridOptions,
            collapsed: true,
            columnOrder: ["Name", "Type", "Team BU", "Error"],
          },
          rows: teams?.__error ? [{ Error: teams.__error }] : (teams || []),
        },
        {
          datasetName: `🧩 Roles via teams (${(teamRoles?.__error ? 0 : (teamRoles || []).length)})`,
          gridOptions: {
            ...baseGridOptions,
            collapsed: true,
            columnOrder: ["Role", "Team", "Role BU", "Error"],
          },
          rows: teamRoles?.__error ? [{ Error: teamRoles.__error }] : (teamRoles || []),
        },
        {
          datasetName: `🔒 Field Security Profiles (${(fieldSecurityProfiles?.__error ? 0 : (fieldSecurityProfiles || []).length)})`,
          gridOptions: {
            ...baseGridOptions,
            collapsed: true,
            columnOrder: ["Name", "Error"],
          },
          rows: fieldSecurityProfiles?.__error ? [{ Error: fieldSecurityProfiles.__error }] : (fieldSecurityProfiles || []),
        },
      ],
    };
  } catch (err: any) {
    return `<div class="error-box">❌ Error: ${err.message}</div>`;
  }
}
