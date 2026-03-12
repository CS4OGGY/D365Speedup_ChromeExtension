interface RunValues {
  solutionName?: string;
  logicalName?: string;
}

export async function run(values: RunValues): Promise<any> {
  const solutionName = values.solutionName?.trim();
  const tableName = values.logicalName?.trim()?.toLowerCase() || null;

  if (!solutionName) throw new Error("Solution name is required.");

  const win = window as any;
  const XrmContext = win.Xrm || win.parent?.Xrm || win.top?.Xrm;
  if (!XrmContext) throw new Error("Xrm not available. Run inside Dynamics 365 CE page.");

  const baseUrl: string = XrmContext.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2";

  async function fetchJson(url: string): Promise<any> {
    const res = await fetch(url, { headers: { Accept: "application/json" } });
    const json = await res.json();
    if (!res.ok) throw new Error(json.error?.message || res.statusText);
    return json;
  }

  console.log(`🔍 Searching for solution '${solutionName}'...`);
  const solRes = await fetchJson(
    `${baseUrl}/solutions?$select=solutionid,friendlyname,uniquename&$filter=uniquename eq '${solutionName}' or friendlyname eq '${solutionName}'`
  );
  if (!solRes.value?.length) throw new Error(`Solution '${solutionName}' not found.`);
  const sol = solRes.value[0];
  const solId: string = sol.solutionid.replace(/[{}]/g, "");

  const compRes = await fetchJson(
    `${baseUrl}/solutioncomponents?$select=objectid,componenttype&$filter=_solutionid_value eq ${solId} and (componenttype eq 10 or componenttype eq 11 or componenttype eq 12)`
  );
  if (!compRes.value?.length) throw new Error("No relationship components found in this solution.");
  const relIds = new Set(compRes.value.map((c: any) => c.objectid?.replace(/[{}]/g, "").toLowerCase()));
  console.log(`✅ Found ${relIds.size} relationship components in solution '${solutionName}'.`);

  const metaRes = await fetchJson(
    `${baseUrl}/EntityDefinitions?$select=LogicalName&$expand=OneToManyRelationships($select=SchemaName,MetadataId,ReferencedEntity,ReferencingEntity,RelationshipType),ManyToOneRelationships($select=SchemaName,MetadataId,ReferencedEntity,ReferencingEntity,RelationshipType),ManyToManyRelationships($select=SchemaName,MetadataId,Entity1LogicalName,Entity2LogicalName,RelationshipType)`
  );

  const entities: any[] = metaRes.value || [];
  const matched: any[] = [];

  for (const e of entities) {
    for (const r of e.OneToManyRelationships || [])
      if (relIds.has(r.MetadataId?.replace(/[{}]/g, "").toLowerCase()))
        matched.push({ ...r, type: "1:N" });
    for (const r of e.ManyToOneRelationships || [])
      if (relIds.has(r.MetadataId?.replace(/[{}]/g, "").toLowerCase()))
        matched.push({ ...r, type: "N:1" });
    for (const r of e.ManyToManyRelationships || [])
      if (relIds.has(r.MetadataId?.replace(/[{}]/g, "").toLowerCase()))
        matched.push({ ...r, type: "N:N" });
  }

  if (!matched.length) throw new Error("No matching relationships found.");

  const filtered = tableName
    ? matched.filter((r: any) =>
        [r.ReferencedEntity, r.ReferencingEntity, r.Entity1LogicalName, r.Entity2LogicalName]
          .map((x: any) => x?.toLowerCase())
          .includes(tableName)
      )
    : matched;

  if (!filtered.length)
    throw new Error(`No relationships found for table '${tableName}' in solution '${solutionName}'.`);

  const unique: any[] = [];
  const seen = new Set<string>();
  for (const r of filtered) {
    const schema: string = r.SchemaName || "unknown_schema";
    if (seen.has(schema)) continue;
    seen.add(schema);
    unique.push(r);
  }

  const displayNameMap: Record<string, string> = {};
  const entitiesToFetch = new Set<string>();
  unique.forEach((r: any) => {
    entitiesToFetch.add(r.ReferencedEntity || r.Entity1LogicalName);
    entitiesToFetch.add(r.ReferencingEntity || r.Entity2LogicalName);
  });

  for (const name of entitiesToFetch) {
    if (!name) continue;
    try {
      const meta = await fetchJson(
        `${baseUrl}/EntityDefinitions(LogicalName='${name}')?$select=DisplayName,LogicalName`
      );
      const label: string =
        meta?.DisplayName?.UserLocalizedLabel?.Label ||
        name.charAt(0).toUpperCase() + name.slice(1);
      displayNameMap[name] = label.replace(/\s+/g, "");
    } catch {
      displayNameMap[name] = name;
    }
  }

  let diagram = "erDiagram\n";

  for (const r of unique) {
    const from: string = r.ReferencedEntity || r.Entity1LogicalName;
    const to: string = r.ReferencingEntity || r.Entity2LogicalName;
    const fromDisp: string = displayNameMap[from];
    const toDisp: string = displayNameMap[to];
    const schema: string = r.SchemaName || "unknown_relationship";

    if (!from || !to || from === to) continue;

    let connector = "||--o{";
    if (r.type === "N:1") connector = "o{--||";
    if (r.type === "N:N") connector = "}o--o{";

    diagram += `  ${fromDisp} ${connector} ${toDisp} : "${r.type} (${schema})"\n`;
  }

  try {
    await navigator.clipboard.writeText(diagram);
    console.log("📋ER diagram copied to clipboard!");
  } catch {
    console.warn("⚠️ Unable to copy to clipboard automatically.");
  }

  return diagram;
}
