
let _nid = 0;

const esc = (s: any): string => {
    if (s == null) return '';
    return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
};
const clean = (v: any): string => v == null ? '' : String(v).trim();

const getDef = (o: any): any => o?.properties?.definition ?? o?.definition ?? o;
const getTriggers = (d: any): any => d?.triggers ?? {};
const getActions = (d: any): any => d?.actions ?? {};
const getConns = (o: any, d: any): any =>
    o?.properties?.connectionReferences ?? o?.connectionReferences ?? d?.connectionReferences ?? {};

const getParam = (p: any, ...ns: string[]): any => {
    for (const n of ns) {
        const v = p?.[n] ?? p?.[n.replace(/\./g, '/')];
        if (clean(v)) return v;
    }
    return '';
};

const MSG: Record<string, string> = { 1:'Added', 2:'Deleted', 3:'Modified', 4:'Added or Modified', 5:'Ownership Changed' };
const SCOPE: Record<string, string> = { 1:'User', 2:'Business Unit', 3:'Parent & Child BU', 4:'Organisation' };

const parseTrig = (trigs: any): any => {
    const entries = Object.entries(trigs ?? {});
    if (!entries.length) return null;
    const [name, t] = entries[0] as [string, any];
    if (!t) return null;
    const p = t?.inputs?.parameters ?? {};
    return {
        name, type: t.type,
        entity: clean(getParam(p, 'subscriptionRequest/entityname', 'subscriptionRequest.entityname', 'entityName')),
        changeType: MSG[clean(getParam(p, 'subscriptionRequest/message', 'subscriptionRequest.message'))] || '',
        scope: SCOPE[clean(getParam(p, 'subscriptionRequest/scope', 'subscriptionRequest.scope'))] || '',
        filterCols: clean(getParam(p, 'subscriptionRequest/filteringattributes', 'subscriptionRequest.filteringattributes')),
        filterExpr: clean(getParam(p, 'subscriptionRequest/filterexpression', 'subscriptionRequest.filterexpression')),
    };
};

const OP_LABEL: Record<string, string> = {
    GetItem:'Get record', GetItemV2:'Get record', GetItemV3:'Get record',
    ListRecords:'List records', ListRecordsV2:'List records',
    CreateRecord:'Create record', CreateRecordV2:'Create record',
    UpdateRecord:'Update record', UpdateRecordV2:'Update record', UpdateOnlyRecord:'Update record',
    DeleteItem:'Delete record', DeleteItemV2:'Delete record', DeleteItemV3:'Delete record',
    WhoAmI:'Get current user', SubscribeWebhookTrigger:'Webhook',
    SendEmailV2:'Send email', SendEmail:'Send email', ReplyTo:'Reply to email', Forward:'Forward email', GetEmail:'Get email',
    PostMessage:'Post Teams msg', PostMessageToConversation:'Post Teams msg',
    UpsertAdaptiveCard:'Post adaptive card',
    StartAnApproval:'Start approval', WaitForAnApproval:'Wait for approval',
};

const descStep = (_name: string, type: string, raw: any): string => {
    const inp = raw?.inputs ?? {};
    const p = inp?.parameters ?? {};
    const op = clean(inp?.host?.operationId ?? inp?.operationId ?? '');
    const entity = clean(p?.entityName ?? p?.entityname ?? '');
    const method = clean(inp?.method ?? '');
    const varN = clean(inp?.variables?.[0]?.name ?? inp?.name ?? '');
    if (type === 'OpenApiConnection' || type === 'OpenApiConnectionWebhook') {
        const lbl = OP_LABEL[op] || (op ? op.replace(/([A-Z])/g, ' $1').trim() : 'Connector action');
        return entity ? `${lbl} · ${entity}` : lbl;
    }
    if (type === 'Http') {
        let host = '';
        try { const uri = (inp?.uri || '').replace(/@\{[^}]+\}/g, 'x'); if (uri.startsWith('http')) host = new URL(uri).hostname.replace(/^www\./, ''); } catch (_e) {}
        return `HTTP ${method || 'request'}${host ? ' → ' + host : ''}`;
    }
    const D: Record<string, string> = {
        If: 'Condition – branches on a rule', Switch: 'Switch – multi-branch',
        Foreach: 'For Each – loops through list', Until: 'Until – repeats until condition',
        InitializeVariable: varN ? `Init var: ${varN}` : 'Initialize variable',
        SetVariable: varN ? `Set var: ${varN}` : 'Set variable',
        AppendToArrayVariable: varN ? `Append to: ${varN}` : 'Append to array',
        AppendToStringVariable: varN ? `Append to: ${varN}` : 'Append to string',
        IncrementVariable: varN ? `Increment: ${varN}` : 'Increment variable',
        DecrementVariable: varN ? `Decrement: ${varN}` : 'Decrement variable',
        Compose: 'Compose – builds a value', ParseJson: 'Parse JSON payload',
        Terminate: 'Terminate – end flow', Response: 'Send HTTP response',
        Scope: 'Scope – group of steps', Delay: 'Delay / wait',
    };
    return D[type] || type.replace(/([A-Z])/g, ' $1').trim();
};

const VAR_TYPES = ['InitializeVariable','SetVariable','AppendToArrayVariable','AppendToStringVariable','IncrementVariable','DecrementVariable'];

const typeColorOp = (t: string, op: string): string => {
    if (['SendEmailV2','SendEmail','ReplyTo','Forward'].includes(op)) return 'pink';
    if (['StartAnApproval','WaitForAnApproval'].includes(op)) return 'amber';
    if (['CreateRecord','CreateRecordV2'].includes(op)) return 'green';
    if (['UpdateRecord','UpdateRecordV2','UpdateOnlyRecord'].includes(op)) return 'amber';
    if (['DeleteItem','DeleteItemV2','DeleteItemV3'].includes(op)) return 'coral';
    const m: Record<string, string> = {
        OpenApiConnectionWebhook:'teal', OpenApiConnection:'blue', Http:'purple',
        If:'amber', Switch:'amber', Until:'amber', Foreach:'green',
        Terminate:'coral', Response:'teal', Scope:'gray', Compose:'gray', ParseJson:'gray',
    };
    return m[t] || 'gray';
};

const typeLabelOp = (t: string, op: string): string => {
    if (['SendEmailV2','SendEmail'].includes(op)) return 'Email';
    if (['ReplyTo','Forward'].includes(op)) return 'Email';
    if (['StartAnApproval','WaitForAnApproval'].includes(op)) return 'Approval';
    if (['CreateRecord','CreateRecordV2'].includes(op)) return 'Create';
    if (['UpdateRecord','UpdateRecordV2','UpdateOnlyRecord'].includes(op)) return 'Update';
    if (['DeleteItem','DeleteItemV2','DeleteItemV3'].includes(op)) return 'Delete';
    if (['GetItem','GetItemV2','GetItemV3'].includes(op)) return 'Get';
    if (['ListRecords','ListRecordsV2'].includes(op)) return 'List';
    const m: Record<string, string> = {
        OpenApiConnectionWebhook:'Webhook', OpenApiConnection:'Connector', Http:'HTTP',
        If:'Condition', Switch:'Switch', Until:'Until', Foreach:'For Each',
        Terminate:'Terminate', Response:'Response', Scope:'Scope', Compose:'Compose', ParseJson:'Parse JSON',
    };
    return m[t] || t;
};

const DC: Record<string, string> = {
    teal:'#00e5b0', blue:'#5badff', purple:'#a78bfa', amber:'#f5a623',
    coral:'#ff6b6b', green:'#4ade80', pink:'#f472b6', gray:'#344260',
};

const flatten = (m: any, out: any[] = [], d = 0): any[] => {
    for (const [n, a] of Object.entries(m ?? {})) {
        const action = a as any;
        out.push({ name: n, type: action.type ?? 'Unknown', raw: action, depth: d });
        if (action.actions) flatten(action.actions, out, d + 1);
        if (action.else?.actions) flatten(action.else.actions, out, d + 1);
        if (action.cases && typeof action.cases === 'object')
            Object.values(action.cases).forEach((c: any) => flatten(c?.actions, out, d + 1));
    }
    return out;
};

const topoSort = (m: any): any[] => {
    const vis = new Set<string>(), res: any[] = [];
    const visit = (n: string) => {
        if (vis.has(n)) return; vis.add(n);
        const a = m[n]; if (!a) return;
        Object.keys(a.runAfter ?? {}).forEach((d: string) => visit(d));
        res.push({ name: n, ...a });
    };
    Object.keys(m ?? {}).forEach(visit);
    return res;
};

const topoSortRaw = (m: any): any[] => {
    const vis = new Set<string>(), res: any[] = [];
    const visit = (n: string) => {
        if (vis.has(n)) return; vis.add(n);
        const a = m[n]; if (!a) return;
        Object.keys(a.runAfter ?? {}).forEach((d: string) => visit(d));
        res.push({ name: n, ...a });
    };
    Object.keys(m || {}).forEach(visit);
    return res;
};

const countN = (m: any): number => {
    let n = 0;
    for (const a of Object.values(m ?? {})) {
        const action = a as any; n++;
        if (action.actions) n += countN(action.actions);
        if (action.else?.actions) n += countN(action.else.actions);
        if (action.cases && typeof action.cases === 'object')
            Object.values(action.cases).forEach((c: any) => { n += countN(c?.actions); });
    }
    return n;
};

const collectTables = (o: any, s = new Set<string>()): Set<string> => {
    if (!o || typeof o !== 'object') return s;
    for (const [k, v] of Object.entries(o)) {
        const kl = k.toLowerCase();
        if (typeof v === 'string' && (v as string).trim() &&
            (kl === 'entityname' || kl === 'tablename' || kl.endsWith('/entityname')))
            s.add((v as string).trim().toLowerCase());
        collectTables(v, s);
    }
    return s;
};

// Convert OData entity set names (plural) to logical names (singular).
// Dataverse uses simple "name + s" pluralisation for ~95% of tables.
// Also handles "ies→y" (e.g. "categories"→"category") and "es→e" (e.g. "geofences"→"geofence").
const singularizeTables = (tables: Set<string>): Set<string> => {
    const out = new Set<string>();
    for (const t of tables) {
        if (!t) continue;
        let s = t;
        if (s.endsWith('ies') && s.length > 4)      s = s.slice(0, -3) + 'y';
        else if (s.endsWith('ses') && s.length > 4)  s = s.slice(0, -2); // processes→process
        else if (s.endsWith('s')   && s.length > 3)  s = s.slice(0, -1);
        out.add(s);
    }
    return out;
};

type RiskItem = { sev: 'warn' | 'info'; label: string; text: string };

const detectRisks = (flat: any[], trig: any, tables: Set<string>): RiskItem[] => {
    const risks: RiskItem[] = [];
    const warn = (label: string, text: string) => risks.push({ sev: 'warn', label, text });
    const info = (label: string, text: string) => risks.push({ sev: 'info', label, text });

    const scopes    = flat.filter(x => x.type === 'Scope');
    const loops     = flat.filter(x => x.type === 'Foreach');
    const loopN     = loops.length;
    const termN     = flat.filter(x => x.type === 'Terminate').length;

    // ── Error handling pattern ──────────────────────────────────────────────
    const scopeNames = scopes.map(x => (x.name as string).toLowerCase().replace(/_/g, ' '));
    const hasTry   = scopeNames.some(n => n.includes('try'));
    const hasCatch = scopeNames.some(n => n.includes('catch'));
    if (!hasTry && !hasCatch) {
        warn('No Try/Catch Scopes', 'No Try / Catch scope pattern detected. Wrap your actions in a Try scope and add a Catch scope (configured to run after failure) for robust error handling.');
    }

    // ── Terminate with Failed status ────────────────────────────────────────
    const hasFailedTerminate = flat.filter(x => x.type === 'Terminate').some(x => {
        const status = (x.raw?.inputs?.runStatus || x.raw?.inputs?.status || '').toLowerCase();
        return status === 'failed' || status === 'cancelled';
    });
    if (termN === 0) {
        info('No Terminate Action', 'No Terminate action found. Without it, the flow always shows "Succeeded" even when a catch block executes – add a Terminate (Failed) at the end of your Catch scope.');
    } else if (!hasFailedTerminate) {
        info('Terminate Not Set to Failed', 'Terminate action(s) found but none set to "Failed" status – errors may be silently swallowed and the flow will report success.');
    }

    // ── Run After failure paths ─────────────────────────────────────────────
    const hasRunAfterFailure = flat.some(x => {
        const ra = x.raw?.runAfter;
        if (!ra || typeof ra !== 'object') return false;
        return Object.values(ra).some((v: any) => Array.isArray(v) && (v.includes('Failed') || v.includes('TimedOut') || v.includes('Skipped')));
    });
    if (!hasRunAfterFailure && flat.length > 4) {
        info('No Failure Run-After Paths', 'No actions are configured to run after failure or timeout. Use "Run After" settings to handle error paths and avoid silent failures.');
    }

    // ── HTTP calls inside loops ─────────────────────────────────────────────
    let httpInLoop = false, listInLoop = false;
    for (const loop of loops) {
        const inner = flatten(loop.raw?.actions ?? {});
        if (inner.some((x: any) => x.type === 'Http')) httpInLoop = true;
        if (inner.some((x: any) => {
            const op = (x.raw?.inputs?.host?.operationId || '').toLowerCase();
            return x.type === 'OpenApiConnection' && (op.includes('list') || op.includes('getitems') || op.includes('retrievemultiple'));
        })) listInLoop = true;
    }
    if (httpInLoop) warn('HTTP Call in Loop', 'HTTP call(s) inside a loop – large datasets may hit API throttling limits. Consider batching or moving calls outside the loop.');
    if (listInLoop) warn('List Query in Loop', 'List/query operation inside a loop – filter at the data source instead of fetching rows inside each iteration.');

    // ── Concurrency on Apply to Each ────────────────────────────────────────
    const concurrentLoops = loops.filter(x => (x.raw?.runtimeConfiguration?.concurrency?.repetitions ?? 1) > 1);
    if (concurrentLoops.length > 0) {
        warn('Concurrent Loop Iterations', `${concurrentLoops.length} "Apply to Each" loop(s) run with concurrency > 1. Parallel iterations sharing variables or connectors can cause race conditions and throttling.`);
    }

    // ── Pagination on list operations ───────────────────────────────────────
    const listOps = flat.filter(x => {
        const op = (x.raw?.inputs?.host?.operationId || '').toLowerCase();
        return x.type === 'OpenApiConnection' && (op.includes('list') || op.includes('getitems') || op.includes('retrievemultiple'));
    });
    const withoutPagination = listOps.filter(x => !x.raw?.runtimeConfiguration?.paginationPolicy?.minimumItemCount);
    if (withoutPagination.length > 0) {
        info('Pagination Not Enabled', `${withoutPagination.length} list/query operation(s) without pagination enabled – results beyond the default page size (typically 5 000 rows) will be silently dropped.`);
    }

    // ── Self-update recursion ───────────────────────────────────────────────
    const selfUpd = flat.some(x => {
        const b = JSON.stringify(x.raw ?? {}).toLowerCase();
        return trig?.entity && b.includes(trig.entity.toLowerCase()) && b.includes('update');
    });
    if (selfUpd) warn('Possible Recursion', 'May update the same table it is triggered from – ensure the trigger has a filter condition or recursion guard to avoid infinite loops.');

    // ── Large flow without scopes ───────────────────────────────────────────
    if (flat.length > 20 && scopes.length === 0) {
        info('No Scope Containers', `${flat.length} actions with no Scope containers. Group related actions into named scopes for readability, easier debugging, and scoped error handling.`);
    }

    // ── Multiple loops ──────────────────────────────────────────────────────
    if (loopN > 1) {
        info('Multiple Loops', `${loopN} loops detected. Nested or sequential loops on large datasets can cause long run times and throttling – verify that each loop is necessary.`);
    }

    // ── High table count ────────────────────────────────────────────────────
    if (tables.size > 5) {
        info('Many Tables', `Touches ${tables.size} tables – review dependencies carefully and consider whether a child flow could encapsulate some operations.`);
    }

    return risks;
};

function getBranches(a: any): Array<{ label: string; color: string; actions: any }> {
    const branches: Array<{ label: string; color: string; actions: any }> = [];
    if (a.type === 'If') {
        if (a.actions && Object.keys(a.actions).length) branches.push({ label: 'If True', color: '#4ade80', actions: a.actions });
        if (a.else?.actions && Object.keys(a.else.actions).length) branches.push({ label: 'If False', color: '#ff6b6b', actions: a.else.actions });
    } else if (a.type === 'Switch') {
        if (a.cases && typeof a.cases === 'object')
            Object.entries(a.cases).forEach(([k, v]: [string, any]) => {
                if (v?.actions) branches.push({ label: 'Case: ' + k, color: '#f5a623', actions: v.actions });
            });
        if (a.default?.actions) branches.push({ label: 'Default', color: '#6b82a8', actions: a.default.actions });
    } else if (a.type === 'Foreach' || a.type === 'Until') {
        if (a.actions && Object.keys(a.actions).length)
            branches.push({ label: a.type === 'Foreach' ? 'Loop body' : 'Repeat body', color: '#4ade80', actions: a.actions });
    } else if (a.type === 'Scope') {
        if (a.actions && Object.keys(a.actions).length)
            branches.push({ label: 'Scope actions', color: '#6b82a8', actions: a.actions });
    } else if (a.actions && Object.keys(a.actions).length) {
        branches.push({ label: 'Actions', color: '#6b82a8', actions: a.actions });
    }
    return branches;
}

function rNestedStep(a: any, depth: number): string {
    if (!a || depth > 8) return '';
    const op = clean(a.inputs?.host?.operationId || a.inputs?.operationId || '');
    const c = typeColorOp(a.type, op);
    const lbl = typeLabelOp(a.type, op);
    const desc = descStep(a.name, a.type, a);
    const dotCol = DC[c] || DC.gray;
    const tagStyle = `background:var(--${c}-bg);color:var(--${c});border:1px solid var(--${c}-bd)`;
    const branches = getBranches(a);
    const hasChildren = branches.length > 0;
    const nid = 'n' + (++_nid);
    let inner = '';
    if (hasChildren) {
        const brHtml = branches.map(br => {
            const s2 = topoSortRaw(br.actions);
            return `<div class="branch-label"><span class="branch-label-dot" style="background:${br.color}"></span>${esc(br.label)}</div>${s2.map(ch => rNestedStep(ch, depth + 1)).join('')}`;
        }).join('');
        inner = `<div class="sub-tree" id="${nid}">${brHtml}</div>`;
    }
    const expBtn = hasChildren ? `<span class="nl-exp" id="e${nid}">&#9658;</span>` : '';
    const tneAttr = hasChildren ? `data-tne="${nid}" style="cursor:pointer"` : '';
    return `<div><div class="nl-step${hasChildren ? ' has-ch' : ''}" ${tneAttr}><span class="nl-dot" style="background:${dotCol}"></span><span class="nl-name">${esc(a.name.replace(/_/g, ' '))}</span><span class="nl-desc">${esc(desc)}</span><span class="nl-tag" style="${tagStyle}">${esc(lbl)}</span>${expBtn}</div>${inner}</div>`;
}

function rOverviewCard(trig: any, flowName: string, paUrl: string): string {
    const paBtn = paUrl ? `<a href="${esc(paUrl)}" target="_blank" class="pa-full-btn">&#8599; Open in Power Automate</a>` : '';
    return `<div class="card">
    <div class="ch"><div class="ci" style="background:var(--teal-bg);border:1px solid var(--teal-bd)">&#9889;</div>
    <div class="ct"><div class="clabel">Flow Info</div><div class="ctitle">${esc(flowName)}</div></div></div>
    <div class="cb"><div class="mg">
      <div class="mc"><div class="ml">Entity</div><div class="mv mv-sm" style="color:var(--teal)">${esc(trig?.entity || '–')}</div></div>
      <div class="mc"><div class="ml">Change</div><div class="mv mv-sm" style="color:var(--blue)">${esc(trig?.changeType || '–')}</div></div>
    </div>${paBtn}</div></div>`;
}

function rTrigCard(trig: any): string {
    if (!trig) return '';
    const rows = [
        ['Entity / Table', trig.entity], ['Change type', trig.changeType], ['Scope', trig.scope],
        ['Filter columns', trig.filterCols], ['Filter expression', trig.filterExpr],
    ].filter(([, v]) => v);
    return `<div class="card">
    <div class="ch"><div class="ci" style="background:var(--teal-bg);border:1px solid var(--teal-bd)">&#128276;</div>
    <div class="ct"><div class="clabel">Trigger</div><div class="ctitle">${esc(trig.name || 'Unknown')}</div></div></div>
    <div class="cb"><div class="kv">${rows.map(([k, v]) => `<div class="kk">${esc(k)}</div><div class="kvv">${esc(v)}</div>`).join('')}</div></div>
    </div>`;
}

function rFlow(sorted: any[]): string {
    const segs: any[] = [];
    let i = 0;
    while (i < sorted.length) {
        if (VAR_TYPES.includes(sorted[i].type)) {
            const vs: any[] = [];
            while (i < sorted.length && VAR_TYPES.includes(sorted[i].type)) { vs.push(sorted[i]); i++; }
            segs.push({ k: 'vars', items: vs });
        } else { segs.push({ k: 'step', item: sorted[i] }); i++; }
    }
    let n = 0;
    const html = segs.map((seg, si) => {
        const isLast = si === segs.length - 1;
        if (seg.k === 'vars') {
            const items = seg.items;
            if (items.length === 1) {
                n++; const a = items[0];
                return `<div class="fs ct-gray"><div class="fsc"><div class="fsd"></div><div class="fsl"></div></div><div class="fsb"><div class="fsc2"><div class="fsnum">${n}</div><div class="fsi"><div class="fsname">${esc(a.name.replace(/_/g, ' '))}</div><div class="fsdesc">${esc(descStep(a.name, a.type, a))}</div></div><div class="fstag" style="background:var(--gray-bg);color:var(--muted);border:1px solid var(--gray-bd)">Init Var</div></div></div></div>`;
            }
            const gid = 'vg' + si;
            const rows = items.map((v: any) => {
                const vn = v.raw?.inputs?.variables?.[0]?.name || v.raw?.inputs?.name || v.name;
                const vt = v.raw?.inputs?.variables?.[0]?.type || '';
                const tag = v.type === 'InitializeVariable' ? 'init' : v.type.replace('Variable', '').toLowerCase();
                return `<div class="vg-item"><span class="vg-tag">${tag}</span><span class="vg-nm">${esc(vn)}</span><span class="vg-tp">${esc(vt)}</span></div>`;
            }).join('');
            return `<div class="var-group"><div class="vgc"><div class="vg-dot"></div><div class="vg-line"${isLast ? ' style="display:none"' : ''}></div></div><div class="vgb"><div class="vg-hd" data-tvg="${gid}"><div class="vg-ico">&#128230;</div><span class="vg-ttl">Variable Setup</span><span class="vg-cnt">${items.length} vars</span><span class="vg-chv" id="c${gid}">&#9660;</span></div><div id="${gid}" style="display:none"><div class="vg-items">${rows}</div></div></div></div>`;
        }
        n++; const a = seg.item;
        const op = clean(a.inputs?.host?.operationId || a.inputs?.operationId || '');
        const c = typeColorOp(a.type, op);
        const l = typeLabelOp(a.type, op);
        const desc = descStep(a.name, a.type, a);
        const branches = getBranches(a);
        const hasNested = branches.length > 0;
        const tid = 't' + si;
        let nestedHtml = '';
        if (hasNested) {
            const brHtml = branches.map(br => {
                const s2 = topoSortRaw(br.actions);
                return `<div class="branch-label"><span class="branch-label-dot" style="background:${br.color}"></span>${esc(br.label)}</div>${s2.map(ch => rNestedStep(ch, 1)).join('')}`;
            }).join('');
            nestedHtml = `<div class="nested-tree" id="${tid}">${brHtml}</div>`;
        }
        const totalNested = countN(a.actions) + countN(a.else?.actions);
        const expBtn = hasNested ? `<span class="expand-btn" id="eb${tid}">&#9658;</span>` : '';
        return `<div class="fs ct-${c}"><div class="fsc"><div class="fsd"></div><div class="fsl"></div></div><div class="fsb">
      <div class="fsc2${hasNested ? ' clickable' : ''}"${hasNested ? ` data-tnt="${tid}"` : ''}>
        <div class="fsnum">${n}</div>
        <div class="fsi"><div class="fsname">${esc(a.name.replace(/_/g, ' '))}</div>
        <div class="fsdesc">${esc(desc)}${totalNested > 0 ? ` &middot; <span style="color:var(--dim)">${totalNested} nested</span>` : ''}</div></div>
        <div class="fstag">${esc(l)}</div>${expBtn}
      </div>${nestedHtml}</div></div>`;
    }).join('');
    return `<div class="card">
    <div class="ch"><div class="ci" style="background:var(--gray-bg);border:1px solid var(--gray-bd)">&#128203;</div>
    <div class="ct"><div class="clabel">Execution Flow</div><div class="ctitle">Steps in Order – ${sorted.length} top-level</div></div></div>
    <div class="cb"><div class="flowline">${html}</div></div></div>`;
}

function rDataMap(tables: Set<string>, conns: any[]): string {
    const tc = [...singularizeTables(tables)].sort().map(t => `<span class="chip teal">${esc(t)}</span>`).join('');
    const connRows = conns.map(c => `<div class="cr"><div class="ci2">&#128279;</div><div><div class="cname">${esc(c.apiName.replace('shared_', ''))}</div><div class="cref">${esc(c.refName || c.key)}</div></div></div>`).join('');
    return [
        `<div class="card"><div class="ch"><div class="ci" style="background:var(--teal-bg);border:1px solid var(--teal-bd)">&#128451;</div><div class="ct"><div class="clabel">Data Map</div><div class="ctitle">Tables</div></div></div><div class="cb">${tc ? `<div class="chips">${tc}</div>` : '<span style="font-size:12px;color:var(--muted)">No tables detected</span>'}</div></div>`,
        conns.length ? `<div class="card"><div class="ch"><div class="ci" style="background:var(--blue-bg);border:1px solid var(--blue-bd)">&#128268;</div><div class="ct"><div class="clabel">Connections</div><div class="ctitle">Services &amp; APIs</div></div></div><div class="cb"><div class="connrows">${connRows}</div></div></div>` : '',
    ].join('');
}

function buildSummary(trig: any, flat: any[]): string {
    const sentences: string[] = [];

    // Trigger sentence
    if (trig?.entity) {
        const changeMap: Record<string, string> = {
            'Created': 'created', 'Modified': 'modified', 'Deleted': 'deleted',
            'Created or Modified': 'created or modified',
            'Created, Modified or Deleted': 'created, modified, or deleted',
        };
        const change = changeMap[trig.changeType] || (trig.changeType ? trig.changeType.toLowerCase() : 'changed');
        let s = `Triggered when a <em>${esc(trig.entity)}</em> record is <em>${esc(change)}</em>`;
        if (trig.filterCols) s += `, specifically when <em>${esc(trig.filterCols)}</em> field(s) change`;
        if (trig.scope && trig.scope !== 'Organisation') s += ` within the <em>${esc(trig.scope)}</em> scope`;
        sentences.push(s);
    } else if (trig?.type) {
        sentences.push(`Triggered by a <em>${esc(typeLabelOp(trig.type, ''))}</em> event`);
    }

    const dvOps = flat.filter(x => x.type === 'OpenApiConnection' || x.type === 'OpenApiConnectionWebhook');
    const getOp  = (x: any): string => (x.raw?.inputs?.host?.operationId ?? x.raw?.inputs?.operationId ?? '').toLowerCase();
    const getEnt = (x: any): string => {
        const p = x.raw?.inputs?.parameters ?? x.raw?.inputs ?? {};
        return clean(p?.entityName ?? p?.entityname ?? p?.dataset ?? '');
    };
    const uniqEnts = (ops: any[]) => [...new Set(ops.map(getEnt).filter(Boolean))];

    const creates    = dvOps.filter(x => /create|insert/i.test(getOp(x)));
    const updates    = dvOps.filter(x => /update|patch/i.test(getOp(x)));
    const deletes    = dvOps.filter(x => /delet/i.test(getOp(x)));
    const reads      = dvOps.filter(x => /list|getitem|retrieve/i.test(getOp(x)));
    const emails     = dvOps.filter(x => /sendemail|sendmail|send_email/i.test(getOp(x)));
    const approvals  = dvOps.filter(x => /approval|startapproval/i.test(getOp(x)));
    const childFlows = dvOps.filter(x => /childflow|invokeflow/i.test(getOp(x)));
    const httpCalls  = flat.filter(x => x.type === 'Http');
    const loops      = flat.filter(x => x.type === 'Foreach');
    const conditions = flat.filter(x => x.type === 'If' || x.type === 'Switch');
    const scopeNames = flat.filter(x => x.type === 'Scope').map(x => (x.name as string).toLowerCase().replace(/_/g, ' '));
    const hasTryCatch = scopeNames.some(n => n.includes('try')) && scopeNames.some(n => n.includes('catch'));

    const clauses: string[] = [];
    if (reads.length)    { const e = uniqEnts(reads);    clauses.push(e.length ? `retrieves <em>${esc(e.join(', '))}</em> records` : 'queries Dataverse records'); }
    if (creates.length)  { const e = uniqEnts(creates);  clauses.push(e.length ? `creates <em>${esc(e.join(', '))}</em> records` : `creates ${creates.length > 1 ? creates.length + ' ' : ''}record${creates.length > 1 ? 's' : ''}`); }
    if (updates.length)  { const e = uniqEnts(updates);  clauses.push(e.length ? `updates <em>${esc(e.join(', '))}</em> records` : `updates ${updates.length > 1 ? updates.length + ' ' : ''}record${updates.length > 1 ? 's' : ''}`); }
    if (deletes.length)  { const e = uniqEnts(deletes);  clauses.push(e.length ? `deletes <em>${esc(e.join(', '))}</em> records` : 'deletes records'); }
    if (emails.length)   clauses.push(`sends ${emails.length > 1 ? emails.length + ' ' : ''}email notification${emails.length > 1 ? 's' : ''}`);
    if (approvals.length) clauses.push('starts an approval workflow');
    if (httpCalls.length) {
        const hosts = [...new Set(httpCalls.map(x => {
            try { const u = (x.raw?.inputs?.uri || '').replace(/@\{[^}]+\}/g, 'x'); return u.startsWith('http') ? new URL(u).hostname.replace(/^www\./, '') : ''; } catch { return ''; }
        }).filter(Boolean))];
        clauses.push(hosts.length ? `calls external API${hosts.length > 1 ? 's' : ''} (<em>${esc(hosts.join(', '))}</em>)` : 'makes HTTP calls to an external service');
    }
    if (childFlows.length) clauses.push(`invokes ${childFlows.length > 1 ? childFlows.length + ' ' : 'a '}child flow${childFlows.length > 1 ? 's' : ''}`);

    if (clauses.length > 0) {
        const prefix = loops.length > 0 ? 'It loops through records and '
            : conditions.length > 0 ? 'It conditionally '
            : 'It ';
        const body = clauses.length === 1 ? clauses[0]
            : clauses.length === 2 ? clauses.join(' and ')
            : clauses.slice(0, -1).join(', ') + ', and ' + clauses[clauses.length - 1];
        sentences.push(prefix + body);
    }

    if (hasTryCatch) sentences.push('Error handling is in place using Try/Catch scopes');

    if (!sentences.length) return `<span style="color:var(--muted);font-size:12px">Unable to determine flow purpose from the definition.</span>`;
    return sentences.map(s => `<p class="sum-p">${s}.</p>`).join('');
}

function rAnalysis(risks: RiskItem[], allFlat: any[], trig: any): string {
    const typeCounts: Record<string, number> = {};
    allFlat.forEach(a => {
        const op = clean(a.raw?.inputs?.host?.operationId || '');
        const l = typeLabelOp(a.type, op);
        typeCounts[l] = (typeCounts[l] || 0) + 1;
    });
    const topTypes = Object.entries(typeCounts).sort((a, b) => b[1] - a[1]).slice(0, 8);
    const typeHtml = topTypes.map(([l, n]) => `<div style="display:flex;align-items:center;justify-content:space-between;padding:6px 0;border-bottom:1px solid var(--border)"><span style="font-size:12px;color:var(--text)">${esc(l)}</span><span style="font-family:var(--mono);font-size:11px;font-weight:700;color:var(--muted)">${n}</span></div>`).join('');
    const sevIcon = (sev: string) => sev === 'warn' ? '&#9888;' : '&#8505;';
    const riskItems = risks.length
        ? risks.map(r => `<div class="ri ${r.sev}"><span style="font-size:13px;flex-shrink:0;margin-top:1px">${sevIcon(r.sev)}</span><div><div class="ri-label">${esc(r.label)}</div><div class="ri-text">${esc(r.text)}</div></div></div>`).join('')
        : `<div class="ri good"><span style="font-size:13px;flex-shrink:0;margin-top:1px">&#10003;</span><div><div class="ri-label">No obvious risks</div><div class="ri-text">No major structural issues found.</div></div></div>`;
    const summaryHtml = buildSummary(trig, allFlat);
    return [
        `<div class="card"><div class="ch"><div class="ci" style="background:var(--teal-bg);border:1px solid var(--teal-bd)">&#129504;</div><div class="ct"><div class="clabel">Smart Summary</div><div class="ctitle">What This Flow Does</div></div></div><div class="cb"><div class="sum-body">${summaryHtml}</div></div></div>`,
        `<div class="card"><div class="ch"><div class="ci" style="background:var(--amber-bg);border:1px solid var(--amber-bd)">&#9888;</div><div class="ct"><div class="clabel">Analysis</div><div class="ctitle">Watch-outs &amp; Risks</div></div></div><div class="cb"><div class="risklist">${riskItems}</div></div></div>`,
        `<div class="card"><div class="ch"><div class="ci" style="background:var(--blue-bg);border:1px solid var(--blue-bd)">&#128202;</div><div class="ct"><div class="clabel">Breakdown</div><div class="ctitle">Action Types Used</div></div></div><div class="cb"><div style="padding:0 2px">${typeHtml}</div></div></div>`,
    ].join('');
}

function switchTabFn(shadow: ShadowRoot, id: string): void {
    shadow.querySelectorAll('.tab').forEach(t => t.classList.toggle('active', (t as HTMLElement).dataset.tab === id));
    shadow.querySelectorAll('.tabpane').forEach(p => p.classList.toggle('active', p.id === 'pane-' + id));
}

function handleClick(e: Event, shadow: ShadowRoot): void {
    const target = e.target as HTMLElement;
    const tab = target.closest('.tab[data-tab]') as HTMLElement | null;
    if (tab?.dataset.tab) { switchTabFn(shadow, tab.dataset.tab); return; }

    const tntEl = target.closest('[data-tnt]') as HTMLElement | null;
    if (tntEl) {
        const id = tntEl.dataset.tnt!;
        const tree = shadow.querySelector('#' + id) as HTMLElement | null;
        const btn = shadow.querySelector('#eb' + id);
        if (tree) { const open = tree.classList.toggle('open'); btn?.classList.toggle('open', open); }
        return;
    }
    const tvgEl = target.closest('[data-tvg]') as HTMLElement | null;
    if (tvgEl) {
        const id = tvgEl.dataset.tvg!;
        const el = shadow.querySelector('#' + id) as HTMLElement | null;
        const chev = shadow.querySelector('#c' + id);
        if (el) { const open = el.style.display === 'none'; el.style.display = open ? 'block' : 'none'; chev?.classList.toggle('open', open); }
        return;
    }
    const tneEl = target.closest('[data-tne]') as HTMLElement | null;
    if (tneEl) {
        const id = tneEl.dataset.tne!;
        const el = shadow.querySelector('#' + id) as HTMLElement | null;
        const btn = shadow.querySelector('#e' + id);
        if (el) { const open = el.classList.toggle('open'); btn?.classList.toggle('open', open); }
    }
}

const FV_CSS = `
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:host{display:block;width:100%;height:100%;
  --bg:#080d18;--panel:#0f1624;--panel2:#141e30;--panel3:#192236;
  --border:#1e2d45;--border2:#253550;--text:#dde6f5;--muted:#6b82a8;--dim:#2d3f5a;
  --teal:#00e5b0;--teal-bg:rgba(0,229,176,.07);--teal-bd:rgba(0,229,176,.2);
  --blue:#5badff;--blue-bg:rgba(91,173,255,.07);--blue-bd:rgba(91,173,255,.2);
  --purple:#a78bfa;--purple-bg:rgba(167,139,250,.07);--purple-bd:rgba(167,139,250,.2);
  --amber:#f5a623;--amber-bg:rgba(245,166,35,.07);--amber-bd:rgba(245,166,35,.2);
  --coral:#ff6b6b;--coral-bg:rgba(255,107,107,.07);--coral-bd:rgba(255,107,107,.2);
  --green:#4ade80;--green-bg:rgba(74,222,128,.07);--green-bd:rgba(74,222,128,.2);
  --pink:#f472b6;--pink-bg:rgba(244,114,182,.07);--pink-bd:rgba(244,114,182,.2);
  --gray-bg:rgba(255,255,255,.04);--gray-bd:rgba(255,255,255,.09);
  --r:8px;--rl:13px;--rxl:18px;
  --sans:'Plus Jakarta Sans',system-ui,sans-serif;--mono:'IBM Plex Mono',monospace;
  --shadow:0 0 0 1px rgba(255,255,255,.035),0 6px 24px rgba(0,0,0,.4);
}
/* Light mode overrides */
.fvroot.light{
  --bg:#f4f7fc;--panel:#ffffff;--panel2:#edf1f8;--panel3:#e2e9f3;
  --border:#cdd6e8;--border2:#b8c6db;--text:#1a2640;--muted:#536882;--dim:#8fa5bf;
  --teal:#00997a;--teal-bg:rgba(0,153,122,.07);--teal-bd:rgba(0,153,122,.25);
  --blue:#1a73d4;--blue-bg:rgba(26,115,212,.07);--blue-bd:rgba(26,115,212,.25);
  --purple:#7c52e8;--purple-bg:rgba(124,82,232,.07);--purple-bd:rgba(124,82,232,.25);
  --amber:#c47a00;--amber-bg:rgba(196,122,0,.07);--amber-bd:rgba(196,122,0,.25);
  --coral:#d93838;--coral-bg:rgba(217,56,56,.07);--coral-bd:rgba(217,56,56,.25);
  --green:#1a8a3a;--green-bg:rgba(26,138,58,.07);--green-bd:rgba(26,138,58,.25);
  --pink:#c0397a;--pink-bg:rgba(192,57,122,.07);--pink-bd:rgba(192,57,122,.25);
  --gray-bg:rgba(0,0,0,.04);--gray-bd:rgba(0,0,0,.1);
  --shadow:0 0 0 1px rgba(0,0,0,.06),0 6px 24px rgba(0,0,0,.1);
}
.fvroot.light .nl-step:hover{background:#d8e2f0}
.fvroot.light .tabcontent::-webkit-scrollbar-thumb{background:var(--border2)}
.fvroot{display:flex;flex-direction:column;width:100%;height:100%;background:var(--bg);color:var(--text);font-family:var(--sans);font-size:13px;line-height:1.5;overflow:hidden}
.tabbar{display:flex;align-items:stretch;border-bottom:1px solid var(--border);background:var(--panel);flex-shrink:0;padding:0 20px;gap:2px}
.pa-full-btn{display:flex;align-items:center;justify-content:center;gap:6px;margin-top:10px;padding:8px 12px;border-radius:var(--rl);background:var(--blue-bg);border:1px solid var(--blue-bd);color:#1298af;font-size:12px;font-weight:600;text-decoration:none;transition:opacity .15s}.pa-full-btn:hover{opacity:.75}
.tab{display:flex;align-items:center;gap:6px;padding:0 14px;height:42px;font-size:12px;font-weight:600;color:var(--muted);cursor:pointer;border-bottom:2px solid transparent;transition:color .15s,border-color .15s;white-space:nowrap;user-select:none}
.tab:hover{color:var(--text)}.tab.active{color:var(--text);border-bottom-color:var(--teal)}
.tab-ico{font-size:13px}
.tab-badge{font-size:9px;font-weight:700;padding:1px 5px;border-radius:10px;background:var(--amber-bg);color:var(--amber);border:1px solid var(--amber-bd)}
.tabcontent{flex:1;overflow-y:auto;min-height:0}
.tabcontent::-webkit-scrollbar{width:5px}.tabcontent::-webkit-scrollbar-thumb{background:var(--border2);border-radius:3px}
.tabpane{display:none;padding:20px 22px;flex-direction:column;gap:14px}.tabpane.active{display:flex}
.card{background:var(--panel);border:1px solid var(--border);border-radius:var(--rxl);box-shadow:var(--shadow);overflow:hidden}
.ch{display:flex;align-items:center;gap:9px;padding:12px 15px;border-bottom:1px solid var(--border);background:linear-gradient(180deg,rgba(255,255,255,.018) 0%,transparent 100%)}
.ci{width:26px;height:26px;border-radius:7px;display:flex;align-items:center;justify-content:center;font-size:13px;flex-shrink:0}
.ct{flex:1;min-width:0}.clabel{font-size:10px;font-weight:700;letter-spacing:.7px;text-transform:uppercase;color:var(--muted)}
.ctitle{font-size:13px;font-weight:700;letter-spacing:-.1px;margin-top:1px;overflow-wrap:break-word}.cb{padding:14px 16px}
.row2{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.mg{display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:8px}
.mc{background:var(--panel2);border:1px solid var(--border);border-radius:var(--rl);padding:10px 12px}
.ml{font-size:10px;font-weight:700;letter-spacing:.5px;text-transform:uppercase;color:var(--muted);margin-bottom:3px}
.mv{font-size:16px;font-weight:700;font-family:var(--mono);letter-spacing:-.4px;line-height:1.2}.mv-sm{font-size:13px}
.kv{display:grid;grid-template-columns:88px 1fr;gap:5px 12px;font-size:12px}
.kk{color:var(--muted);font-weight:500}.kvv{font-family:var(--mono);font-size:11px;word-break:break-all;color:var(--text)}
.chips{display:flex;flex-wrap:wrap;gap:5px;margin-top:7px}
.chip{display:inline-flex;align-items:center;gap:4px;padding:3px 8px;border-radius:20px;font-size:11px;font-weight:500;background:var(--panel3);border:1px solid var(--border2);color:var(--text)}
.chip.teal{background:var(--teal-bg);border-color:var(--teal-bd);color:var(--teal)}
.flowline{display:flex;flex-direction:column}
.fs{display:flex;align-items:stretch}
.fsc{display:flex;flex-direction:column;align-items:center;width:36px;flex-shrink:0}
.fsd{width:10px;height:10px;border-radius:50%;border:2px solid;margin-top:12px;flex-shrink:0;z-index:1}
.fsl{flex:1;width:2px;margin-top:3px}.fs:last-child .fsl{background:transparent!important}
.fsb{flex:1;padding-bottom:6px}
.fsc2{display:flex;align-items:center;gap:8px;padding:8px 11px;background:var(--panel2);border:1px solid var(--border);border-radius:var(--rl);transition:background .12s,border-color .12s}
.fsc2.clickable{cursor:pointer}.fsc2:hover{background:var(--panel3);border-color:var(--border2)}
.fsnum{width:20px;height:20px;border-radius:5px;display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:800;font-family:var(--mono);flex-shrink:0}
.fsi{flex:1;min-width:0}
.fsname{font-family:var(--mono);font-size:12px;font-weight:500;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.fsdesc{font-size:11px;color:var(--muted);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;margin-top:1px}
.fstag{font-size:9px;font-weight:700;padding:2px 6px;border-radius:5px;flex-shrink:0;letter-spacing:.3px;white-space:nowrap}
.expand-btn{width:16px;height:16px;border-radius:4px;background:var(--gray-bg);border:1px solid var(--gray-bd);display:flex;align-items:center;justify-content:center;font-size:8px;color:var(--muted);flex-shrink:0;transition:background .12s,transform .2s}
.expand-btn.open{background:var(--panel3);transform:rotate(90deg)}
.nested-tree{margin-top:4px;padding-left:16px;border-left:2px solid var(--border);margin-left:5px;display:none}.nested-tree.open{display:block}
.branch-label{font-size:10px;font-weight:700;letter-spacing:.5px;text-transform:uppercase;color:var(--muted);padding:5px 0 4px 2px;display:flex;align-items:center;gap:5px;margin-top:3px}
.branch-label:first-child{margin-top:0}.branch-label-dot{width:5px;height:5px;border-radius:50%;flex-shrink:0}
.nl-step{display:flex;align-items:center;gap:7px;padding:6px 10px;background:var(--panel3);border:1px solid var(--border);border-radius:var(--r);margin-bottom:3px;transition:background .1s}
.nl-step:hover{background:#1f2d42}.nl-step.has-ch{cursor:pointer}
.nl-dot{width:6px;height:6px;border-radius:50%;flex-shrink:0}
.nl-name{font-family:var(--mono);font-size:11px;font-weight:500;flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.nl-desc{font-size:10px;color:var(--muted);flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.nl-tag{font-size:9px;font-weight:700;padding:1px 5px;border-radius:4px;flex-shrink:0;white-space:nowrap}
.nl-exp{width:13px;height:13px;border-radius:3px;background:var(--gray-bg);border:1px solid var(--gray-bd);display:flex;align-items:center;justify-content:center;font-size:8px;color:var(--muted);flex-shrink:0;transition:transform .2s}
.nl-exp.open{transform:rotate(90deg)}
.sub-tree{padding-left:12px;border-left:2px solid var(--border);margin-left:3px;display:none;margin-top:3px}.sub-tree.open{display:block}
.var-group{display:flex;align-items:stretch}
.vgc{display:flex;flex-direction:column;align-items:center;width:36px;flex-shrink:0}
.vg-dot{width:10px;height:10px;border-radius:50%;border:2px solid var(--border2);background:var(--panel3);margin-top:12px;flex-shrink:0;z-index:1}
.vg-line{flex:1;width:2px;background:linear-gradient(var(--border2),var(--border));margin-top:3px}.vgb{flex:1;padding-bottom:6px}
.vg-hd{display:flex;align-items:center;gap:8px;padding:8px 11px;background:var(--panel2);border:1px solid var(--border);border-radius:var(--rl);cursor:pointer;transition:background .12s;user-select:none}
.vg-hd:hover{background:var(--panel3);border-color:var(--border2)}
.vg-ico{width:20px;height:20px;border-radius:5px;background:var(--gray-bg);border:1px solid var(--gray-bd);display:flex;align-items:center;justify-content:center;font-size:10px;flex-shrink:0}
.vg-ttl{font-family:var(--mono);font-size:11px;font-weight:500;flex:1}.vg-cnt{font-size:10px;color:var(--muted);font-family:var(--mono)}
.vg-chv{font-size:9px;color:var(--muted);transition:transform .2s;flex-shrink:0}.vg-chv.open{transform:rotate(180deg)}
.vg-items{padding:3px 0 0}
.vg-item{display:flex;align-items:center;gap:7px;padding:5px 10px;background:var(--panel3);border:1px solid var(--border);border-radius:var(--r);font-size:11px;font-family:var(--mono);margin-bottom:3px}
.vg-tag{font-size:9px;font-weight:700;padding:1px 5px;border-radius:4px;background:var(--purple-bg);color:var(--purple);border:1px solid var(--purple-bd);flex-shrink:0}
.vg-nm{color:var(--text);flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.vg-tp{font-size:10px;color:var(--muted)}
.ct-teal .fsd{border-color:var(--teal);background:var(--teal)}.ct-teal .fsl{background:linear-gradient(var(--teal),var(--border))}.ct-teal .fsnum{background:var(--teal-bg);color:var(--teal)}.ct-teal .fsname{color:var(--teal)}.ct-teal .fstag{background:var(--teal-bg);color:var(--teal);border:1px solid var(--teal-bd)}
.ct-blue .fsd{border-color:var(--blue);background:var(--blue)}.ct-blue .fsl{background:linear-gradient(var(--blue),var(--border))}.ct-blue .fsnum{background:var(--blue-bg);color:var(--blue)}.ct-blue .fstag{background:var(--blue-bg);color:var(--blue);border:1px solid var(--blue-bd)}
.ct-purple .fsd{border-color:var(--purple);background:var(--purple)}.ct-purple .fsl{background:linear-gradient(var(--purple),var(--border))}.ct-purple .fsnum{background:var(--purple-bg);color:var(--purple)}.ct-purple .fstag{background:var(--purple-bg);color:var(--purple);border:1px solid var(--purple-bd)}
.ct-amber .fsd{border-color:var(--amber);background:var(--amber)}.ct-amber .fsl{background:linear-gradient(var(--amber),var(--border))}.ct-amber .fsnum{background:var(--amber-bg);color:var(--amber)}.ct-amber .fstag{background:var(--amber-bg);color:var(--amber);border:1px solid var(--amber-bd)}
.ct-coral .fsd{border-color:var(--coral);background:var(--coral)}.ct-coral .fsl{background:linear-gradient(var(--coral),var(--border))}.ct-coral .fsnum{background:var(--coral-bg);color:var(--coral)}.ct-coral .fstag{background:var(--coral-bg);color:var(--coral);border:1px solid var(--coral-bd)}
.ct-pink .fsd{border-color:var(--pink);background:var(--pink)}.ct-pink .fsl{background:linear-gradient(var(--pink),var(--border))}.ct-pink .fsnum{background:var(--pink-bg);color:var(--pink)}.ct-pink .fstag{background:var(--pink-bg);color:var(--pink);border:1px solid var(--pink-bd)}
.ct-green .fsd{border-color:var(--green);background:var(--green)}.ct-green .fsl{background:linear-gradient(var(--green),var(--border))}.ct-green .fsnum{background:var(--green-bg);color:var(--green)}.ct-green .fstag{background:var(--green-bg);color:var(--green);border:1px solid var(--green-bd)}
.ct-gray .fsd{border-color:var(--border2);background:var(--panel3)}.ct-gray .fsl{background:linear-gradient(var(--border2),var(--border))}.ct-gray .fsnum{background:var(--gray-bg);color:var(--muted)}.ct-gray .fstag{background:var(--gray-bg);color:var(--muted);border:1px solid var(--gray-bd)}
.sum-body{padding:2px 0}.sum-p{margin:0 0 10px;font-size:13px;line-height:1.7;color:var(--text)}.sum-p:last-child{margin-bottom:0}.sum-p em{font-style:normal;color:var(--teal);font-weight:600}
.risklist{display:flex;flex-direction:column;gap:7px}
.ri{display:flex;align-items:flex-start;gap:9px;padding:9px 12px;border-radius:var(--rl);font-size:12px;line-height:1.6}
.ri.warn{background:var(--amber-bg);border:1px solid var(--amber-bd)}.ri.good{background:var(--green-bg);border:1px solid var(--green-bd)}.ri.info{background:var(--blue-bg);border:1px solid var(--blue-bd)}
.ri-label{font-size:9px;font-weight:800;letter-spacing:.6px;text-transform:uppercase;margin-bottom:2px}
.ri.warn .ri-label{color:var(--amber)}.ri.good .ri-label{color:var(--green)}.ri.info .ri-label{color:var(--blue)}.ri-text{color:var(--muted)}
.connrows{display:flex;flex-direction:column;gap:6px}
.cr{display:flex;align-items:center;gap:10px;padding:9px 12px;background:var(--panel2);border:1px solid var(--border);border-radius:var(--rl)}
.ci2{width:26px;height:26px;border-radius:7px;background:var(--blue-bg);border:1px solid var(--blue-bd);display:flex;align-items:center;justify-content:center;font-size:12px;flex-shrink:0}
.cname{font-size:12px;font-weight:600;font-family:var(--mono)}.cref{font-size:10px;color:var(--muted);margin-top:1px}
`;

export function render(container: HTMLElement, json: any): void {
    let shadow = container.shadowRoot;
    if (!shadow) {
        shadow = container.attachShadow({ mode: 'open' });
        const style = document.createElement('style');
        style.textContent = FV_CSS;
        shadow.appendChild(style);

        const root = document.createElement('div');
        root.className = 'fvroot';
        root.innerHTML = `
          <div class="tabbar">
            <div class="tab active" data-tab="overview"><span class="tab-ico">&#9889;</span> Overview</div>
            <div class="tab" data-tab="datamap"><span class="tab-ico">&#128451;</span> Data Map &amp; Connections</div>
            <div class="tab" data-tab="analysis"><span class="tab-ico">&#128202;</span> Analysis <span class="tab-badge" id="riskBadge" style="display:none"></span></div>
          </div>
          <div class="tabcontent">
            <div class="tabpane active" id="pane-overview"></div>
            <div class="tabpane" id="pane-datamap"></div>
            <div class="tabpane" id="pane-analysis"></div>
          </div>`;
        shadow.appendChild(root);
        root.addEventListener('click', (e: Event) => handleClick(e, shadow!));

        // Sync light/dark mode with parent body and watch for live changes
        const syncMode = () => root.classList.toggle('light', document.body.classList.contains('light-mode'));
        syncMode();
        const observer = new MutationObserver(syncMode);
        observer.observe(document.body, { attributes: true, attributeFilter: ['class'] });
        (container as any).__fvObserver = observer;
    } else {
        // Sync mode on subsequent renders (modal reopened)
        const root = shadow.querySelector('.fvroot') as HTMLElement | null;
        root?.classList.toggle('light', document.body.classList.contains('light-mode'));
    }

    // Reset tab to overview
    switchTabFn(shadow, 'overview');

    _nid = 0;
    try {
        const def = getDef(json);
        const trig = parseTrig(getTriggers(def));
        const amap = getActions(def);
        const allFlat = flatten(amap);
        const sorted = topoSort(amap);
        const cobj = getConns(json, def);
        const conns = Object.entries(cobj).map(([k, c]: [string, any]) => ({
            key: k,
            apiName: c?.api?.name || k,
            refName: c?.connection?.connectionReferenceLogicalName || '',
        }));
        const tables = collectTables(def);
        const localRisks = detectRisks(allFlat, trig, tables);
        const flowName = clean(
            json?.name || json?.properties?.displayName ||
            json?.properties?.definitionSummary?.displayName || 'Flow'
        );

        const paUrl = clean(json?.__paUrl);
        shadow.querySelector('#pane-overview')!.innerHTML =
            `<div class="row2">${rOverviewCard(trig, flowName, paUrl)}${rTrigCard(trig)}</div>` +
            rFlow(sorted);
        shadow.querySelector('#pane-datamap')!.innerHTML = rDataMap(tables, conns);
        shadow.querySelector('#pane-analysis')!.innerHTML = rAnalysis(localRisks, allFlat, trig);

        const riskBadge = shadow.querySelector('#riskBadge') as HTMLElement | null;
        if (riskBadge) {
            if (localRisks.length > 0) { riskBadge.textContent = String(localRisks.length); riskBadge.style.display = 'inline-flex'; }
            else { riskBadge.style.display = 'none'; }
        }
    } catch (e: any) {
        shadow.querySelector('#pane-overview')!.innerHTML =
            `<div style="padding:20px;color:#ff6b6b;font-size:13px">Parse error: ${esc(e?.message || 'Unknown error')}</div>`;
    }
}
