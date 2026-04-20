// ======================================================================
// FetchXML Builder — Shadow DOM Component
// Renders a live query builder wired to real Dataverse metadata.
// Adapted from template/FetchXMLTemplate.html — no inline handlers.
// ======================================================================

export interface BuilderCallbacks {
  fetchAllEntities: () => Promise<{ name: string; display: string }[]>;
  fetchEntityMeta: (logicalName: string) => Promise<{
    attrs: { n: string; d: string; t: string; targets?: string[]; primary?: boolean }[];
    rels: { name: string; display: string; fromAttr: string; toEntity: string; toAttr: string }[];
    views?: { id: string; name: string; type: string; fx: string }[];
    primaryName?: string;
    primaryId?: string;
    objectTypeCode?: number;
  }>;
  fetchAttrOptions?: (entityName: string, attrName: string) => Promise<{ v: number; l: string }[]>;
  fetchLookupRecords?: (targetEntity: string, searchTerm: string, searchField?: string) => Promise<{ id: string; name: string; sub?: string; url?: string }[]>;
  onXmlChange: (xml: string) => void;
}

const OPS: Record<string, string[]> = {
  string:          ['eq','ne','contains','not-contains','starts-with','not-starts-with','ends-with','not-ends-with','not-null','null'],
  memo:            ['contains','not-contains','not-null','null'],
  integer:         ['eq','ne','gt','ge','lt','le','not-null','null'],
  bigint:          ['eq','ne','gt','ge','lt','le','not-null','null'],
  decimal:         ['eq','ne','gt','ge','lt','le','not-null','null'],
  double:          ['eq','ne','gt','ge','lt','le','not-null','null'],
  money:           ['eq','ne','gt','ge','lt','le','not-null','null'],
  datetime:        [
    'on','on-or-after','on-or-before',
    'yesterday','today','tomorrow',
    'next-seven-days','last-seven-days','next-week','last-week','this-week',
    'next-month','last-month','this-month',
    'next-year','last-year','this-year',
    'last-x-hours','next-x-hours','last-x-days','next-x-days','last-x-weeks','next-x-weeks','last-x-months','next-x-months','last-x-years','next-x-years',
    'olderthan-x-minutes','olderthan-x-hours','olderthan-x-days','olderthan-x-weeks','olderthan-x-months','olderthan-x-years',
    'not-null','null',
    'in-fiscal-year','in-fiscal-period','in-fiscal-period-and-year','in-or-after-fiscal-period-and-year','in-or-before-fiscal-period-and-year',
    'last-fiscal-year','this-fiscal-year','next-fiscal-year','last-x-fiscal-years','next-x-fiscal-years',
    'last-fiscal-period','this-fiscal-period','next-fiscal-period','last-x-fiscal-periods','next-x-fiscal-periods',
  ],
  boolean:         ['eq','not-null','null'],
  picklist:        ['in','not-in','not-null','null','contains','not-contains','starts-with','not-starts-with','ends-with','not-ends-with'],
  state:           ['in','not-in','not-null','null','contains','not-contains','starts-with','not-starts-with','ends-with','not-ends-with'],
  status:          ['in','not-in','not-null','null','contains','not-contains','starts-with','not-starts-with','ends-with','not-ends-with'],
  multiselect:     ['eq','ne','not-null','null','contain-values','not-contain-values'],
  lookup:          ['eq','ne','eq-userid','ne-userid','eq-businessid','not-null','null','contains','not-contains','starts-with','not-starts-with','ends-with','not-ends-with'],
  customer:        ['eq','ne','not-null','null','contains','not-contains','starts-with','not-starts-with','ends-with','not-ends-with'],
  owner:           [
    'eq-userid','ne-userid',
    'eq-useroruserhierarchy','eq-useroruserhierarchyandteams','eq-userteams','eq-useroruserteams',
    'eq','ne','not-null','null',
    'contains','not-contains','starts-with','not-starts-with','ends-with','not-ends-with',
  ],
  uniqueidentifier:['eq','ne','not-null','null'],
};
// Internal virtual ops → FetchXML like/not-like with % wrapping
const LIKE_WRAP: Record<string,{op:string;pre:string;suf:string}> = {
  'contains':       {op:'like',     pre:'%', suf:'%'},
  'not-contains':   {op:'not-like', pre:'%', suf:'%'},
  'starts-with':    {op:'like',     pre:'',  suf:'%'},
  'not-starts-with':{op:'not-like', pre:'',  suf:'%'},
  'ends-with':      {op:'like',     pre:'%', suf:''},
  'not-ends-with':  {op:'not-like', pre:'%', suf:''},
};
const OL: Record<string,string> = {
  eq:'Equals', ne:'Does Not Equal',
  'contains':'Contains', 'not-contains':'Does Not Contain',
  'starts-with':'Begins With', 'not-starts-with':'Does Not Begin With',
  'ends-with':'Ends With', 'not-ends-with':'Does Not End With',
  'not-null':'Contains Data', null:'Does Not Contain Data',
  in:'Equals', 'not-in':'Does Not Equal',
  gt:'Is Greater Than', ge:'Is Greater Than or Equal To',
  lt:'Is Less Than', le:'Is Less Than or Equal To',
  on:'On', 'on-or-after':'On or After', 'on-or-before':'On or Before',
  yesterday:'Yesterday', today:'Today', tomorrow:'Tomorrow',
  'next-seven-days':'Next 7 Days', 'last-seven-days':'Last 7 Days',
  'next-week':'Next Week', 'last-week':'Last Week', 'this-week':'This Week',
  'next-month':'Next Month', 'last-month':'Last Month', 'this-month':'This Month',
  'next-year':'Next Year', 'last-year':'Last Year', 'this-year':'This Year',
  'last-x-hours':'Last X Hours', 'next-x-hours':'Next X Hours',
  'last-x-days':'Last X Days', 'next-x-days':'Next X Days',
  'last-x-weeks':'Last X Weeks', 'next-x-weeks':'Next X Weeks',
  'last-x-months':'Last X Months', 'next-x-months':'Next X Months',
  'last-x-years':'Last X Years', 'next-x-years':'Next X Years',
  'olderthan-x-minutes':'Older Than X Minutes', 'olderthan-x-hours':'Older Than X Hours',
  'olderthan-x-days':'Older Than X Days', 'olderthan-x-weeks':'Older Than X Weeks',
  'olderthan-x-months':'Older Than X Months', 'olderthan-x-years':'Older Than X Years',
  'in-fiscal-year':'In Fiscal Year', 'in-fiscal-period':'In Fiscal Period',
  'in-fiscal-period-and-year':'In Fiscal Period and Year',
  'in-or-after-fiscal-period-and-year':'In or After Fiscal Period',
  'in-or-before-fiscal-period-and-year':'In or Before Fiscal Period',
  'last-fiscal-year':'Last Fiscal Year', 'this-fiscal-year':'This Fiscal Year', 'next-fiscal-year':'Next Fiscal Year',
  'last-x-fiscal-years':'Last X Fiscal Years', 'next-x-fiscal-years':'Next X Fiscal Years',
  'last-fiscal-period':'Last Fiscal Period', 'this-fiscal-period':'This Fiscal Period', 'next-fiscal-period':'Next Fiscal Period',
  'last-x-fiscal-periods':'Last X Fiscal Periods', 'next-x-fiscal-periods':'Next X Fiscal Periods',
  'eq-userid':'Equals Current User', 'ne-userid':'Does Not Equal Current User',
  'eq-useroruserhierarchy':'Equals Current User Or Their Reporting Hierarchy',
  'eq-useroruserhierarchyandteams':'Equals Current User And Their Teams Or Their Reporting Hierarchy And Their Teams',
  'eq-userteams':"Equals Current User's Teams",
  'eq-useroruserteams':"Equals Current User Or User's Teams",
  'eq-businessid':'Equals Current Business Unit',
  'contain-values':'Contains Values', 'not-contain-values':'Does Not Contain Values',
};
const NO_VAL   = new Set([
  'null','not-null',
  'yesterday','today','tomorrow',
  'next-seven-days','last-seven-days','next-week','last-week','this-week',
  'next-month','last-month','this-month',
  'next-year','last-year','this-year',
  'last-fiscal-year','this-fiscal-year','next-fiscal-year',
  'last-fiscal-period','this-fiscal-period','next-fiscal-period',
  'eq-userid','ne-userid','eq-businessid',
  'eq-useroruserhierarchy','eq-useroruserhierarchyandteams','eq-userteams','eq-useroruserteams',
]);
const X_VAL = new Set([
  'last-x-hours','next-x-hours','last-x-days','next-x-days',
  'last-x-weeks','next-x-weeks','last-x-months','next-x-months','last-x-years','next-x-years',
  'olderthan-x-minutes','olderthan-x-hours','olderthan-x-days','olderthan-x-weeks','olderthan-x-months','olderthan-x-years',
  'in-fiscal-year','in-fiscal-period','in-fiscal-period-and-year',
  'in-or-after-fiscal-period-and-year','in-or-before-fiscal-period-and-year',
  'last-x-fiscal-years','next-x-fiscal-years','last-x-fiscal-periods','next-x-fiscal-periods',
]);
const MULTI_VAL = new Set(['in','not-in','contain-values','not-contain-values']);
const NAME_ATTR_TYPES = new Set(['picklist','state','status','lookup','customer','owner']);
const TYPE_MAP: Record<string,string> = {
  String:'string', Memo:'memo', Integer:'integer', BigInt:'bigint',
  Decimal:'decimal', Double:'double', Money:'money', DateTime:'datetime',
  Boolean:'boolean', Picklist:'picklist', State:'state', Status:'status',
  Lookup:'lookup', Customer:'customer', Owner:'owner',
  Uniqueidentifier:'uniqueidentifier', Virtual:'string', EntityName:'string',
  MultiSelectPicklist:'multiselect', Image:'string', File:'string',
};
const normType = (t: string) => TYPE_MAP[t] || 'string';
const opsFor   = (t: string) => OPS[t] || ['eq','ne','null','not-null'];

const SVG_SEARCH = `<svg width="12" height="12" viewBox="0 0 20 20" fill="currentColor"><path fill-rule="evenodd" d="M9 3.5a5.5 5.5 0 100 11 5.5 5.5 0 000-11zM2 9a7 7 0 1112.452 4.391l3.328 3.329a.75.75 0 11-1.06 1.06l-3.329-3.328A7 7 0 012 9z" clip-rule="evenodd"/></svg>`;
const POPOUT_ICON = String.fromCharCode(0x29C9);

// ── CSS ────────────────────────────────────────────────────────────────
const CSS = `
:host { display:flex; flex-direction:column; height:100%; font-family:"Segoe UI",Tahoma,Geneva,Verdana,sans-serif; font-size:13px; color:#e0e0e0; background:#111; }
*{margin:0;padding:0;box-sizing:border-box}
label{color:#939393;font-size:12px;padding-left:2px;display:block;margin-bottom:2px}
#fxb-root{display:flex;flex-direction:column;height:100%;overflow:hidden}
.tab-header{display:flex;border-bottom:1px solid #222;background:#141414;flex-shrink:0}
.tab-item{flex:1;text-align:center;padding:6px 8px;cursor:pointer;color:#888;font-weight:500;transition:all .2s;border-bottom:2px solid transparent;font-size:11px;letter-spacing:.01em}
.tab-item:hover{color:#cfcfcf;background:rgb(121 136 139/.1)}
.tab-item.active{color:#1abcd7;border-bottom:2px solid rgb(91 134 143/.7)}
:host-context(body.light-mode) .tab-header{background:#e8eaeb;border-bottom-color:#c8cdd0}
:host-context(body.light-mode) .tab-item{color:#666}
:host-context(body.light-mode) .tab-item:hover{color:#2d4a52;background:rgb(91 134 143/.1)}
:host-context(body.light-mode) .tab-item.active{color:#2d4a52;background:rgb(217 222 223/.15);border-bottom-color:rgb(91 134 143/.8)}
#panes{flex:1;overflow:hidden;display:flex;flex-direction:column}
.tab-content{display:none!important;opacity:0}
.tab-content.active{display:flex!important;flex-direction:column;opacity:1;flex:1;overflow:hidden}
.snippet-inputs{flex:1;overflow-y:auto;scrollbar-width:thin;scrollbar-color:#333 #181818}
.snippet-inputs::-webkit-scrollbar{width:5px}
.snippet-inputs::-webkit-scrollbar-thumb{background:#333}
.content-card{background:#1e1e1e;margin-bottom:1px;border-bottom:1px solid #1a1a1a}
.section-header{display:flex;align-items:center;gap:8px;padding:7px 10px;cursor:pointer;background:#181818;border-bottom:1px solid #2a2a2a;transition:background .15s;user-select:none}
.section-header:hover{background:#1e1e1e}
.section-header.locked{cursor:default;pointer-events:none}
.step-badge{width:18px;height:18px;background:#00b4d8;color:#000;font-size:10px;font-weight:700;display:flex;align-items:center;justify-content:center;flex-shrink:0}
.locked .step-badge{background:#2a2a2a;color:#555}
.section-title{font-size:12.5px;font-weight:600;flex:1;color:#e0e0e0}
.locked .section-title{color:#444}
.section-meta{font-size:11px;color:#00b4d8;font-family:Consolas,monospace;font-weight:600}
.locked .section-meta{color:#333}
.section-chev{font-size:10px;color:#444;transition:transform .15s}
.section-header.open .section-chev{transform:rotate(180deg)}
.section-body{display:none;padding:10px}
.section-header.open+.section-body{display:block}
.input-group{margin-bottom:8px}
.input-text,.input-select{width:100%;border:1px solid #353232;background:#181818;color:#e4e6eb;padding:6px 8px;font-size:13px;font-family:inherit;outline:none;transition:border .2s;margin-top:3px}
.input-text:focus,.input-select:focus{border-color:#434343}
.dd-trigger{width:100%;background:#181818;border:1px solid #353232;color:#e4e6eb;padding:6px 8px;font-size:13px;font-family:inherit;display:flex;align-items:center;gap:6px;cursor:pointer;text-align:left;outline:none;transition:border .15s;margin-top:3px}
.dd-trigger:hover{border-color:#434343}
.dd-trigger.open{border-color:#00b4d8}
.dd-trigger-lbl{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.dd-trigger-lbl.ph{color:#555}
.dd-arrow{font-size:9px;color:#555;flex-shrink:0;transition:transform .15s}
.dd-trigger.open .dd-arrow{transform:rotate(180deg)}
.dd-popup{position:absolute;left:0;right:0;background:#252525;border:1px solid #2a2a2a;z-index:9999;display:none;box-shadow:0 6px 24px rgba(0,0,0,.8)}
.dd-popup.open{display:block}
.dd-below{top:calc(100% + 1px)}
.dd-top{bottom:calc(100% + 1px)}
.dd-search{display:flex;align-items:center;gap:6px;padding:6px 8px;background:#1e1e1e;border-bottom:1px solid #2a2a2a}
.dd-search svg{flex-shrink:0;color:#555}
.dd-search input{flex:1;border:none;outline:none;background:transparent;color:#e0e0e0;font-size:12px;font-family:inherit}
.dd-search input::placeholder{color:#555}
.dd-close{flex-shrink:0;width:24px;height:24px;border:1px solid rgba(0,180,216,.22);background:rgba(0,180,216,.08);color:#00b4d8;cursor:pointer;font-size:16px;font-weight:700;line-height:1;display:flex;align-items:center;justify-content:center;padding:0;border-radius:3px;transition:color .15s,background .15s,border-color .15s,transform .15s}
.dd-close:hover{color:#fff;background:#00b4d8;border-color:#00b4d8;transform:scale(1.03)}
.dd-toolbar{display:flex;align-items:center;gap:5px;padding:4px 8px;background:#1a1a1a;border-bottom:1px solid #2a2a2a}
.dd-tbar-btn{font-size:11px;font-weight:600;font-family:inherit;padding:2px 8px;cursor:pointer;border:1px solid #333;background:#202020;color:#888;transition:all .15s}
.dd-tbar-btn:hover{background:#2a2a2a;color:#00b4d8;border-color:#00b4d8}
.dd-tbar-count{font-size:11px;color:#555;margin-left:auto;font-family:Consolas,monospace}
.dd-list{max-height:220px;overflow-y:auto;background:#252525;scrollbar-width:thin;scrollbar-color:#333 #1a1a1a}
.dd-list::-webkit-scrollbar{width:6px}
.dd-list::-webkit-scrollbar-thumb{background:#333}
.dd-item{padding:7px 10px;cursor:pointer;color:#ddd;font-size:13px;border-bottom:1px solid #1e1e1e;display:flex;align-items:center;gap:8px;transition:background .1s,color .1s}
.dd-item:hover{background:#2b2b2b;color:#00b4d8}
.dd-item.sel{background:#0d2226;color:#00b4d8}
.dd-item:last-child{border-bottom:none}
.ent-icon{width:22px;height:22px;display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:700;flex-shrink:0;background:#1e2d3a;color:#00b4d8}
.ent-names{flex:1;min-width:0}
.ent-display{font-size:12.5px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;color:#dfe1e1}
.ent-logical{font-size:10.5px;color:#555;font-family:Consolas,monospace}
.view-badge{font-size:9px;font-weight:700;padding:1px 5px;text-transform:uppercase;letter-spacing:.4px;flex-shrink:0}
.vb-s{background:rgba(0,180,216,.12);color:#00b4d8;border:1px solid rgba(0,180,216,.2)}
.vb-p{background:rgba(255,204,0,.1);color:#ffcc00;border:1px solid rgba(255,204,0,.2)}
.col-item{display:flex;align-items:center;gap:8px;padding:6px 10px;cursor:pointer;border-bottom:1px solid #1e1e1e;color:#ddd;transition:background .1s;user-select:none}
.col-item:hover{background:#2b2b2b;color:#e0e0e0}
.col-item.chk{background:#0d2226}
.col-item:last-child{border-bottom:none}
.col-item input[type=checkbox]{accent-color:#00b4d8;cursor:pointer;flex-shrink:0;width:13px;height:13px;pointer-events:none}
.col-names{flex:1;min-width:0}
.col-disp{font-size:12.5px;font-weight:500;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.col-logical{font-size:10.5px;color:#555;font-family:Consolas,monospace}
.type-badge{font-size:9px;font-weight:700;padding:2px 5px;flex-shrink:0;text-transform:uppercase;letter-spacing:.3px}
.tb-string,.tb-memo{background:rgba(0,160,80,.12);color:#3ecf8e;border:1px solid rgba(0,160,80,.18)}
.tb-integer,.tb-bigint,.tb-decimal,.tb-double,.tb-money{background:rgba(190,120,0,.12);color:#d4893a;border:1px solid rgba(190,120,0,.18)}
.tb-datetime{background:rgba(110,70,210,.12);color:#a78bfa;border:1px solid rgba(110,70,210,.18)}
.tb-boolean{background:rgba(0,160,80,.12);color:#3ecf8e;border:1px solid rgba(0,160,80,.18)}
.tb-lookup,.tb-customer,.tb-owner{background:rgba(0,180,216,.1);color:#00b4d8;border:1px solid rgba(0,180,216,.2)}
.tb-picklist,.tb-state,.tb-status{background:rgba(110,70,210,.12);color:#a78bfa;border:1px solid rgba(110,70,210,.18)}
.tb-uniqueidentifier{background:rgba(200,50,50,.1);color:#e07070;border:1px solid rgba(200,50,50,.18)}
.tb-default{background:#252525;color:#555;border:1px solid #333}
.col-chips{min-height:32px;padding:4px;background:#181818;border:1px solid #353232;margin-top:3px;display:flex;flex-wrap:wrap;gap:3px;cursor:pointer;transition:border .15s;align-items:flex-start}
.col-chips:hover{border-color:#434343}
.col-chips.focused{border-color:#00b4d8}
.chip{display:inline-flex;align-items:center;gap:3px;background:#0d2226;color:#00b4d8;border:1px solid rgba(0,180,216,.22);font-size:10.5px;padding:2px 5px;font-family:Consolas,monospace;max-width:120px}
.chip span{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.chip-x{cursor:pointer;color:#224d5a;font-size:13px;line-height:1;flex-shrink:0;transition:color .15s}
.chip-x:hover{color:#ff5555}
.chip-ph{font-size:12px;color:#444;padding:3px 2px;align-self:center}
.chip-more{font-size:11px;color:#555;padding:3px 4px;align-self:center;cursor:pointer}
.chip-more:hover{color:#00b4d8}
.filter-group{border:1px solid #2a2a2a;margin-bottom:5px}
.fg-bar{display:flex;align-items:center;gap:6px;padding:5px 8px;background:#181818;border-bottom:1px solid #2a2a2a}
.logic-tog{display:flex;border:1px solid #2a2a2a;overflow:hidden}
.lg-b{padding:2px 9px;font-size:10.5px;font-weight:700;background:#181818;color:#555;border:none;cursor:pointer;font-family:inherit;transition:all .15s;letter-spacing:.3px}
.lg-b.on{background:#00b4d8;color:#000}
.fg-lbl{font-size:10.5px;color:#555;flex:1}
.fg-btns{display:flex;gap:3px}
.cond-row{display:flex;align-items:center;gap:4px;padding:5px 8px;border-bottom:1px solid #1a1a1a}
.cond-row:last-child{border-bottom:none}
.cr-s{background:#181818;border:1px solid #2a2a2a;color:#e4e6eb;font-size:11.5px;font-family:inherit;padding:4px 20px 4px 6px;outline:none;min-width:0;cursor:pointer;appearance:none;flex:1;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='8' height='5'%3E%3Cpath d='M0 0l4 5 4-5z' fill='%23555'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 5px center;transition:border .12s;scrollbar-width:thin;scrollbar-color:#333 #181818}
.cr-s:focus{border-color:#00b4d8}
.cr-s::-webkit-scrollbar{width:6px}
.cr-s::-webkit-scrollbar-track{background:#181818}
.cr-s::-webkit-scrollbar-thumb{background:#333;border-radius:3px}
.cr-s::-webkit-scrollbar-thumb:hover{background:#555}
.cr-v{background:#181818;border:1px solid #2a2a2a;color:#e4e6eb;font-size:11.5px;font-family:inherit;padding:4px 6px;outline:none;min-width:0;flex:1;transition:border .12s}
.cr-v:focus{border-color:#00b4d8}
.cr-noval{flex:1;font-size:10px;color:#444;font-style:italic;padding:4px 6px}
.fp-wrap{position:relative;flex:1.2;min-width:0}
.fp-trigger{display:flex;align-items:center;gap:6px;padding:0 6px;height:26px;background:#181818;border:1px solid #2a2a2a;cursor:pointer;font-size:11.5px;color:#c8d6e8;user-select:none;transition:border .12s;overflow:hidden}
.fp-trigger:hover{border-color:#00b4d8}
.fp-trigger-name{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.fp-trigger-badge{flex-shrink:0;font-size:9px;font-weight:700;border:1px solid;padding:0 4px;line-height:14px;font-family:Consolas,monospace}
.fp-trigger-arr{flex-shrink:0;opacity:0.35;font-size:8px}
.fp-drop{position:absolute;left:0;right:0;top:calc(100% + 2px);z-index:9999;background:#1a1a1a;border:1px solid #2a2a2a;box-shadow:0 6px 20px rgba(0,0,0,.7);display:none;flex-direction:column}
.fp-drop.open{display:flex}
.fp-search{background:#141414;border:none;border-bottom:1px solid #2a2a2a;color:#e4e6eb;font-size:11.5px;font-family:inherit;padding:6px 8px;outline:none;width:100%;box-sizing:border-box}
.fp-list{max-height:230px;overflow-y:auto;scrollbar-width:thin;scrollbar-color:#333 #1a1a1a}
.fp-item{display:flex;align-items:center;gap:6px;padding:5px 8px;cursor:pointer;font-size:11.5px;color:#c8d6e8;border-bottom:1px solid #1e1e1e;transition:background .1s}
.fp-item:last-child{border-bottom:none}
.fp-item:hover{background:#0d2d35;color:#e4e6eb}
.fp-item.sel{background:rgba(0,180,216,0.07);color:#c8d6e8}
.fp-item-name{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.fp-item-logical{flex-shrink:0;font-size:10px;color:#3a3a3a;font-family:Consolas,monospace;white-space:nowrap}
.fp-item:hover .fp-item-logical{color:#555}
.fp-item-badge{flex-shrink:0;font-size:9px;font-weight:700;border:1px solid transparent;padding:0 3px;line-height:13px;font-family:Consolas,monospace}
.fp-item-ph{color:#555;font-style:italic}
.fp-empty{padding:10px;font-size:11px;color:#555;font-style:italic;text-align:center}
.lk-card{border:1px solid #2a2a2a;margin-bottom:5px}
.lk-depth-0{border-left:2px solid #00b4d8}
.lk-depth-1{border-left:2px solid #d4893a}
.lk-depth-2{border-left:2px solid #a78bfa}
.lk-depth-3{border-left:2px solid #3ecf8e}
.lk-depth-4{border-left:2px solid #e07070}
.lk-hdr{display:flex;align-items:center;gap:6px;padding:6px 10px;background:#181818;border-bottom:1px solid #2a2a2a;cursor:pointer;transition:background .15s}
.lk-hdr:hover{background:#1e1e1e}
.lk-depth-0 .lk-badge{background:rgba(0,180,216,.12);color:#00b4d8;border:1px solid rgba(0,180,216,.2)}
.lk-depth-1 .lk-badge{background:rgba(190,120,0,.12);color:#d4893a;border:1px solid rgba(190,120,0,.2)}
.lk-depth-2 .lk-badge{background:rgba(110,70,210,.12);color:#a78bfa;border:1px solid rgba(110,70,210,.2)}
.lk-badge{font-size:9px;font-weight:700;padding:2px 6px;text-transform:uppercase;flex-shrink:0}
.lk-name{font-size:12.5px;font-weight:600;flex:1;color:#dfe1e1}
.lk-path{font-size:10px;color:#444;font-family:Consolas,monospace;margin-right:4px}
.lk-alias-disp{font-size:10.5px;color:#555;font-family:Consolas,monospace}
.lk-body{padding:8px;display:none;background:#161616}
.lk-card.open>.lk-body{display:block}
.lk-row{display:flex;align-items:center;gap:8px;margin-bottom:7px}
.lk-row>label{font-size:11px;color:#666;white-space:nowrap;min-width:30px;margin:0;padding:0}
.jp-grp{display:flex}
.jp{padding:3px 9px;font-size:10.5px;font-weight:700;border:1px solid #2a2a2a;cursor:pointer;background:#181818;color:#555;font-family:inherit;transition:all .15s}
.jp.on{background:#00b4d8;color:#000;border-color:#00b4d8}
.alias-inp{background:#181818;border:1px solid #2a2a2a;color:#e4e6eb;font-size:11.5px;font-family:Consolas,monospace;padding:3px 6px;outline:none;width:80px;transition:border .12s}
.alias-inp:focus{border-color:#00b4d8}
.lk-sec-lbl{font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.6px;color:#555;margin-bottom:4px;margin-top:8px}
.lk-sec-lbl:first-child{margin-top:0}
.nested-lk-list{margin-top:6px;padding-left:10px;border-left:1px dashed #2a2a2a}
.sort-row{display:flex;align-items:center;gap:6px;padding:5px 8px;background:#181818;border:1px solid #2a2a2a;margin-bottom:4px}
.sort-attr{font-size:11.5px;font-family:Consolas,monospace;flex:1;color:#00b4d8}
.sort-alias-lbl{font-size:10px;color:#555;font-family:Consolas,monospace}
.sd-grp{display:flex} 
.sd{padding:2px 8px;font-size:10.5px;font-weight:600;border:1px solid #2a2a2a;cursor:pointer;font-family:inherit;background:#181818;color:#555;transition:all .15s}
.sd.on{background:rgba(0,180,216,.12);color:#00b4d8;border-color:#00b4d8}
.sd.on.d{background:rgba(190,120,0,.1);color:#d4893a;border-color:rgba(190,120,0,.4)}
.ib{padding:3px 7px;font-size:11px;cursor:pointer;background:none;border:none;color:#555;transition:all .15s;font-family:inherit}
.ib:hover{background:#2a2a2a;color:#e0e0e0}
.ib.del:hover{background:rgba(200,50,50,.12);color:#ff5555}
.add-btn{width:100%;background:transparent;border:1px dashed #2a2a2a;color:#555;font-size:12px;font-weight:600;font-family:inherit;padding:6px;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:5px;transition:all .15s;margin-top:4px}
.add-btn:hover{border-color:#00b4d8;color:#00b4d8;background:rgba(0,180,216,.04)}
.add-btn.sm{padding:4px;font-size:11px;margin-top:3px}
.opt-row{display:flex;align-items:center;gap:10px;margin-bottom:7px}
.opt-lbl{font-size:12px;color:#939393;flex:1}
.tog{width:32px;height:17px;background:#2a2a2a;position:relative;cursor:pointer;transition:background .15s;flex-shrink:0}
.tog.on{background:#00b4d8}
.tog::after{content:'';position:absolute;top:2.5px;left:2.5px;width:12px;height:12px;background:#fff;transition:transform .15s}
.tog.on::after{transform:translateX(15px)}
.num-in{width:65px;background:#181818;border:1px solid #353232;color:#e4e6eb;padding:4px 6px;font-size:12px;font-family:inherit;outline:none;transition:border .15s}
.num-in:focus{border-color:#434343}
.xml-label{display:flex;align-items:center;padding:5px 10px;background:#181818;border-bottom:1px solid #2a2a2a;font-size:11.5px;color:#555;flex-shrink:0;gap:6px}
.xml-editor-wrap{flex:1;position:relative;background:#181818;overflow:hidden;display:flex;flex-direction:column;min-height:120px}
#xml-highlight{position:absolute;top:0;left:0;right:0;bottom:0;padding:10px;margin:0;font-family:Consolas,'Fira Code',monospace;font-size:12.5px;line-height:1.7;white-space:pre-wrap;overflow:hidden;pointer-events:none;word-break:break-all;color:transparent}
#xml-editor{position:absolute;top:0;left:0;right:0;bottom:0;width:100%;height:100%;padding:10px;margin:0;font-family:Consolas,'Fira Code',monospace;font-size:12.5px;line-height:1.7;white-space:pre-wrap;background:transparent;border:none;color:transparent;caret-color:#e4e6eb;outline:none;resize:none;overflow:auto;z-index:1;word-break:break-all;scrollbar-width:thin;scrollbar-color:#333 #181818}
.xt-tag{color:#7dd3fc}
.xt-attr{color:#93c5fd}
.xt-val{color:#86efac}
.xt-comment{color:#4b5563;font-style:italic}
#val-msg{font-size:11.5px;padding:5px 10px;display:flex;align-items:center;gap:6px;font-weight:500;flex-shrink:0}
#val-msg.ok{background:#091a0e;color:#3ecf8e;border-top:1px solid rgba(62,207,142,.15)}
#val-msg.warn{background:#1c1408;color:#d4893a;border-top:1px solid rgba(190,120,0,.2)}
.hint{text-align:center;padding:14px;color:#444;font-size:12px;font-style:italic}
.load-spinner{text-align:center;padding:14px;color:#555;font-size:11px}
.btn-secondary{border-radius: 3px;background:#2a2a2a;color:#d1d1d1;border:1px solid #333;padding:3px 10px;cursor:pointer;font-family:inherit;font-size:11px;font-weight:600;transition:background .2s}
.btn-secondary:hover{background:#333;color:#e0e0e0}
.lk-srch{position:relative;flex:1;min-width:0;display:flex;align-items:center}
.lk-ctrl-row{display:flex;align-items:center;flex:1;min-width:0}
.lk-trigger{display:flex;align-items:center;gap:6px;padding:0 8px;height:26px;flex:1;background:#181818;border:1px solid #2a2a2a;cursor:pointer;font-size:11.5px;color:#c8d6e8;user-select:none;transition:border .12s;overflow:hidden;min-width:0}
.lk-trigger:hover{border-color:#00b4d8}
.lk-trigger-lbl{overflow:hidden;text-overflow:ellipsis;flex:1;white-space:nowrap}
.lk-trigger-badge{flex-shrink:0;background:#00b4d8;color:#000;font-size:10px;font-weight:700;border-radius:10px;padding:0 5px;line-height:16px}
.lk-search-btn{flex-shrink:0;width:26px;height:26px;background:#181818;border:1px solid #2a2a2a;border-left:0;color:#666;cursor:pointer;display:flex;align-items:self-end;justify-content:center;font-size:20px;padding:0;transition:all .15s;font-family:inherit;overflow:hidden}
.lk-search-btn:hover{border-color:#00b4d8;color:#00b4d8;background:#0d2d35;border-left:1px solid #00b4d8}
.lk-search-btn.active{border-color:#00b4d8;color:#00b4d8;background:rgba(0,180,216,.1);border-left:1px solid #00b4d8}
.lk-sel-drop{position:absolute;left:0;right:0;top:calc(100% + 3px);z-index:9999;background:#1a1a1a;border:1px solid #2a2a2a;box-shadow:0 6px 20px rgba(0,0,0,.7);display:none;overflow-y:auto;max-height:220px;scrollbar-width:thin;scrollbar-color:#333 #1a1a1a}
.lk-sel-drop.open{display:block}
.lk-sel-item{display:flex;align-items:center;gap:8px;padding:6px 10px;border-bottom:1px solid #1e1e1e;font-size:11.5px;color:#c8d6e8}
.lk-sel-item:last-child{border-bottom:none}
.lk-sel-item-name{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.lk-sel-rm{flex-shrink:0;background:none;border:none;color:#555;cursor:pointer;font-size:10px;padding:2px 5px;font-family:inherit;line-height:1;transition:color .1s}
.lk-sel-rm:hover{color:#ff5555;background:rgba(200,50,50,.1)}
.lk-drop{position:absolute;left:0;right:0;top:calc(100% + 3px);z-index:9999;background:#1a1a1a;border:1px solid #2a2a2a;box-shadow:0 6px 20px rgba(0,0,0,.7);display:none;overflow:hidden}
.lk-drop.open{display:flex;flex-direction:column}
.lk-drop-field-row{display:flex;align-items:center;gap:6px;padding:5px 8px;border-bottom:1px solid #242424;background:#141414}
.lk-drop-field-lbl{font-size:10px;color:#555;white-space:nowrap;flex-shrink:0}
.lk-drop-field-sel{flex:1;background:#181818;border:1px solid #2a2a2a;color:#c8d6e8;font-size:10.5px;font-family:inherit;padding:2px 4px;outline:none;cursor:pointer;min-width:0}
.lk-drop-field-sel:focus{border-color:#00b4d8}
.lk-drop-search{display:flex;align-items:center;gap:6px;padding:7px 8px;border-bottom:1px solid #242424}
.lk-srch-inp{flex:1;background:none;border:none;outline:none;color:#e4e6eb;font-size:11.5px;font-family:inherit;padding:0}
.lk-drop-close{flex-shrink:0;width:24px;height:24px;border:1px solid rgba(0,180,216,.22);background:rgba(0,180,216,.08);color:#00b4d8;cursor:pointer;font-size:16px;font-weight:700;line-height:1;display:flex;align-items:center;justify-content:center;padding:0;border-radius:3px;transition:color .15s,background .15s,border-color .15s,transform .15s}
.lk-drop-close:hover{color:#fff;background:#00b4d8;border-color:#00b4d8;transform:scale(1.03)}
.lk-drop-list{max-height:200px;overflow-y:auto;scrollbar-width:thin;scrollbar-color:#333 #1a1a1a}
.lk-srch-item{display:flex;align-items:center;gap:8px;padding:6px 10px;cursor:pointer;font-size:11.5px;color:#c8d6e8;border-bottom:1px solid #222;transition:background .1s}
.lk-srch-item:last-child{border-bottom:none}
.lk-srch-item:hover{background:#0d2d35;color:#00b4d8}
.lk-srch-item.sel{background:rgba(0,180,216,0.08);color:#00b4d8}
.lk-srch-item.sel:hover{background:#0d2d35}
.lk-srch-item-check{width:13px;height:13px;border-radius:2px;border:1px solid #444;flex-shrink:0;display:flex;align-items:center;justify-content:center;font-size:9px;color:transparent;background:#111}
.lk-srch-item.sel .lk-srch-item-check{background:#00b4d8;border-color:#00b4d8;color:#000}
.lk-srch-item-body{display:flex;flex-direction:column;flex:1;min-width:0;gap:1px}
.lk-srch-item-pop{flex-shrink:0;color:#444;font-size:12px;text-decoration:none;padding:0 4px;line-height:1;opacity:1;transition:color .15s}
.lk-srch-item-pop:hover{color:#00b4d8}
.lk-srch-item-name{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.lk-srch-item-sub{font-size:10px;color:#555;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-family:Consolas,monospace}
.lk-srch-item.sel .lk-srch-item-sub{color:#007a8a}
.lk-srch-empty{padding:12px 10px;font-size:11px;color:#555;font-style:italic;text-align:center;line-height:1.6}
.cr-ms{background:#181818;border:1px solid #2a2a2a;color:#e4e6eb;font-size:11.5px;font-family:inherit;padding:2px;outline:none;flex:1;min-width:0;scrollbar-width:thin;scrollbar-color:#333 #181818}
.cr-ms option{padding:3px 6px;background:#181818;color:#e4e6eb}
.cr-ms option:checked{background:#0d2226 !important;color:#00b4d8}
.cr-cbp{position:relative;flex:1;min-width:0}
.cr-cbp-trigger{display:flex;align-items:center;justify-content:space-between;gap:4px;padding:0 8px;height:26px;background:#181818;border:1px solid #2a2a2a;border-radius:4px;cursor:pointer;font-size:11.5px;color:#c8d6e8;user-select:none;white-space:nowrap;overflow:hidden}
.cr-cbp-trigger:hover{border-color:#00b4d8}
.cr-cbp-trigger-lbl{overflow:hidden;text-overflow:ellipsis;flex:1}
.cr-cbp-trigger-arr{flex-shrink:0;opacity:0.5;font-size:9px}
.cr-cbp-drop{position:absolute;top:calc(100% + 2px);left:0;right:0;z-index:999;background:#1e1e1e;border:1px solid #2a2a2a;border-radius:4px;max-height:160px;overflow-y:auto;scrollbar-width:thin;scrollbar-color:#333 #1e1e1e;box-shadow:0 4px 12px rgba(0,0,0,0.5);display:none}
.cr-cbp-drop.open{display:block}
.cr-cbp-close-row{display:flex;justify-content:flex-end;padding:4px 6px;border-bottom:1px solid #2a2a2a;background:#181818;position:sticky;top:0;z-index:1}
.cr-cbp-close{border:1px solid rgba(0,180,216,.22);background:rgba(0,180,216,.08);color:#00b4d8;cursor:pointer;font-size:16px;font-weight:700;line-height:1;display:flex;align-items:center;justify-content:center;width:24px;height:24px;padding:0;border-radius:3px;transition:color .15s,background .15s,border-color .15s,transform .15s}
.cr-cbp-close:hover{color:#fff;background:#00b4d8;border-color:#00b4d8;transform:scale(1.03)}
.cr-cbp-item{display:flex;align-items:center;gap:6px;padding:4px 8px;cursor:pointer;font-size:11.5px;color:#c8d6e8;user-select:none}
.cr-cbp-item:hover{background:rgba(255,255,255,0.07)}
.cr-cbp-item input[type="checkbox"]{accent-color:#00b4d8;cursor:pointer;width:12px;height:12px;flex-shrink:0}

/* ── Aggregate rows ──────────────────────────────────────────────── */
.agg-row{display:flex;align-items:center;gap:4px;padding:4px 6px;background:#181818;border:1px solid #2a2a2a;margin-bottom:4px}
.agg-row.is-groupby{border-left:2px solid #00b4d8}
.agg-row.is-agg{border-left:2px solid #d4893a}
.agg-fn-sel{background-color:#181818;border:1px solid #2a2a2a;color:#e4e6eb;font-size:11.5px;font-family:inherit;padding:4px 20px 4px 6px;outline:none;width:118px;cursor:pointer;appearance:none;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='8' height='5'%3E%3Cpath d='M0 0l4 5 4-5z' fill='%23555'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 5px center;transition:border .12s;flex-shrink:0}
.agg-fn-sel:focus{border-color:#00b4d8}
.agg-alias-inp{background:#181818;border:1px solid #2a2a2a;color:#e4e6eb;font-size:11.5px;font-family:Consolas,monospace;padding:4px 6px;outline:none;width:120px;flex-shrink:0;transition:border .12s}
.agg-alias-inp:focus{border-color:#00b4d8}
.agg-alias-inp.invalid{border-color:rgba(200,50,50,.5)}
.agg-mode-badge{font-size:9px;font-weight:700;padding:2px 5px;background:rgba(190,120,0,.15);color:#d4893a;border:1px solid rgba(190,120,0,.3)}

/* ── Light Mode ─────────────────────────────────────────────────────── */
:host-context(body.light-mode){background:#f5f6f8;color:#1a1a1a}
:host-context(body.light-mode) label{color:#5a6068}
:host-context(body.light-mode) .snippet-inputs{scrollbar-color:#c8cdd0 #f5f6f8}
:host-context(body.light-mode) .snippet-inputs::-webkit-scrollbar-thumb{background:#c8cdd0}
:host-context(body.light-mode) .content-card{background:#fff;border-bottom-color:#e8eaec}
:host-context(body.light-mode) .section-header{background:#eef0f2;border-bottom-color:#dde0e3}
:host-context(body.light-mode) .section-header:hover{background:#e4e7ea}
:host-context(body.light-mode) .locked .step-badge{background:#d0d4d8;color:#999}
:host-context(body.light-mode) .section-title{color:#1a1a1a}
:host-context(body.light-mode) .locked .section-title{color:#bbb}
:host-context(body.light-mode) .locked .section-meta{color:#bbb}
:host-context(body.light-mode) .section-chev{color:#999}
:host-context(body.light-mode) .hint{color:#aaa}
:host-context(body.light-mode) .load-spinner{color:#999}

/* inputs */
:host-context(body.light-mode) .input-text,:host-context(body.light-mode) .input-select{background:#fff;border-color:#c8cdd0;color:#1a1a1a}
:host-context(body.light-mode) .input-text:focus,:host-context(body.light-mode) .input-select:focus{border-color:#00b4d8}
:host-context(body.light-mode) .num-in{background:#fff;border-color:#c8cdd0;color:#1a1a1a}
:host-context(body.light-mode) .num-in:focus{border-color:#00b4d8}

/* dropdown trigger */
:host-context(body.light-mode) .dd-trigger{background:#fff;border-color:#c8cdd0;color:#1a1a1a}
:host-context(body.light-mode) .dd-trigger:hover{border-color:#00b4d8}
:host-context(body.light-mode) .dd-trigger-lbl.ph{color:#aaa}
:host-context(body.light-mode) .dd-arrow{color:#aaa}

/* dropdown popup */
:host-context(body.light-mode) .dd-popup{background:#fff;border-color:#d0d4d8;box-shadow:0 6px 24px rgba(0,0,0,.12)}
:host-context(body.light-mode) .dd-search{background:#f5f6f8;border-bottom-color:#e0e3e6}
:host-context(body.light-mode) .dd-search svg{color:#aaa}
:host-context(body.light-mode) .dd-search input{color:#1a1a1a}
:host-context(body.light-mode) .dd-search input::placeholder{color:#aaa}
:host-context(body.light-mode) .dd-close{color:#007a95;background:#e0f5fa;border-color:rgba(0,180,216,.3)}
:host-context(body.light-mode) .dd-toolbar{background:#eef0f2;border-bottom-color:#dde0e3}
:host-context(body.light-mode) .dd-tbar-btn{background:#fff;border-color:#c8cdd0;color:#555}
:host-context(body.light-mode) .dd-tbar-btn:hover{background:#eef7fb;color:#00b4d8;border-color:#00b4d8}
:host-context(body.light-mode) .dd-tbar-count{color:#aaa}
:host-context(body.light-mode) .dd-list{background:#fff;scrollbar-color:#c8cdd0 #f5f6f8}
:host-context(body.light-mode) .dd-list::-webkit-scrollbar-thumb{background:#c8cdd0}
:host-context(body.light-mode) .dd-item{color:#1a1a1a;border-bottom-color:#eef0f2}
:host-context(body.light-mode) .dd-item:hover{background:#eef7fb;color:#00b4d8}
:host-context(body.light-mode) .dd-item.sel{background:#e0f5fa;color:#007a95}
:host-context(body.light-mode) .ent-icon{background:#e0f5fa;color:#007a95}
:host-context(body.light-mode) .ent-display{color:#1a1a1a}
:host-context(body.light-mode) .ent-logical{color:#8a9099}

/* column chips */
:host-context(body.light-mode) .col-item{color:#1a1a1a;border-bottom-color:#eef0f2}
:host-context(body.light-mode) .col-item:hover{background:#eef7fb;color:#1a1a1a}
:host-context(body.light-mode) .col-item.chk{background:#e0f5fa}
:host-context(body.light-mode) .col-logical{color:#8a9099}
:host-context(body.light-mode) .col-chips{background:#fff;border-color:#c8cdd0}
:host-context(body.light-mode) .col-chips:hover{border-color:#00b4d8}
:host-context(body.light-mode) .col-chips.focused{border-color:#00b4d8}
:host-context(body.light-mode) .chip{background:#e0f5fa;color:#007a95;border-color:rgba(0,180,216,.3)}
:host-context(body.light-mode) .chip-x{color:#99cfe0}
:host-context(body.light-mode) .chip-ph{color:#aaa}
:host-context(body.light-mode) .chip-more{color:#aaa}

/* filters */
:host-context(body.light-mode) .filter-group{border-color:#d0d4d8}
:host-context(body.light-mode) .fg-bar{background:#eef0f2;border-bottom-color:#dde0e3}
:host-context(body.light-mode) .logic-tog{border-color:#c8cdd0}
:host-context(body.light-mode) .lg-b{background:#fff;color:#888;border-color:#c8cdd0}
:host-context(body.light-mode) .lg-b.on{background:#00b4d8;color:#fff}
:host-context(body.light-mode) .fg-lbl{color:#888}
:host-context(body.light-mode) .cond-row{border-bottom-color:#eef0f2}
:host-context(body.light-mode) .cr-s,:host-context(body.light-mode) .cr-v{background:#fff;border-color:#c8cdd0;color:#1a1a1a}
:host-context(body.light-mode) .lk-drop-field-row{background:#f8f9fa;border-bottom-color:#eef0f2}
:host-context(body.light-mode) .lk-drop-field-lbl{color:#aaa}
:host-context(body.light-mode) .lk-drop-field-sel{background:#fff;border-color:#c8cdd0;color:#1a1a1a}
:host-context(body.light-mode) .fp-trigger{background:#fff;border-color:#c8cdd0;color:#1a1a1a}
:host-context(body.light-mode) .fp-trigger:hover{border-color:#00b4d8}
:host-context(body.light-mode) .fp-drop{background:#fff;border-color:#c8cdd0;box-shadow:0 6px 20px rgba(0,0,0,.12)}
:host-context(body.light-mode) .fp-search{background:#f8f9fa;border-bottom-color:#eef0f2;color:#1a1a1a}
:host-context(body.light-mode) .fp-list{scrollbar-color:#c8cdd0 #fff}
:host-context(body.light-mode) .fp-item{color:#1a1a1a;border-bottom-color:#f0f2f5}
:host-context(body.light-mode) .fp-item:hover{background:#eef7fb;color:#1a1a1a}
:host-context(body.light-mode) .fp-item.sel{background:#e0f5fa}
:host-context(body.light-mode) .fp-item-logical{color:#c8cdd0}
:host-context(body.light-mode) .fp-item:hover .fp-item-logical{color:#aaa}
:host-context(body.light-mode) .fp-empty{color:#aaa}
:host-context(body.light-mode) .cr-noval{color:#aaa}
:host-context(body.light-mode) .cr-ms{background:#fff;border-color:#c8cdd0;color:#1a1a1a;scrollbar-color:#c8cdd0 #fff}
:host-context(body.light-mode) .cr-ms option{background:#fff;color:#1a1a1a}
:host-context(body.light-mode) .cr-ms option:checked{background:#e0f5fa !important;color:#007a95}
:host-context(body.light-mode) .cr-cbp-trigger{background:#fff;border-color:#c8cdd0;color:#1a1a1a}
:host-context(body.light-mode) .cr-cbp-trigger:hover{border-color:#00b4d8}
:host-context(body.light-mode) .cr-cbp-drop{background:#fff;border-color:#c8cdd0;scrollbar-color:#c8cdd0 #fff;box-shadow:0 4px 12px rgba(0,0,0,0.12)}
:host-context(body.light-mode) .cr-cbp-close-row{background:#f8f9fa;border-bottom-color:#eef0f2}
:host-context(body.light-mode) .cr-cbp-close{color:#007a95;background:#e0f5fa;border-color:rgba(0,180,216,.3)}
:host-context(body.light-mode) .cr-cbp-item{color:#1a1a1a}
:host-context(body.light-mode) .cr-cbp-item:hover{background:#f0f2f5}

/* related tables / links */
:host-context(body.light-mode) .lk-card{border-color:#d0d4d8}
:host-context(body.light-mode) .lk-hdr{background:#eef0f2;border-bottom-color:#dde0e3}
:host-context(body.light-mode) .lk-hdr:hover{background:#e4e7ea}
:host-context(body.light-mode) .lk-body{background:#f8f9fa}
:host-context(body.light-mode) .lk-name{color:#1a1a1a}
:host-context(body.light-mode) .lk-path{color:#8a9099}
:host-context(body.light-mode) .lk-alias-disp{color:#8a9099}
:host-context(body.light-mode) .lk-row>label{color:#5a6068}
:host-context(body.light-mode) .jp{background:#fff;border-color:#c8cdd0;color:#888}
:host-context(body.light-mode) .alias-inp{background:#fff;border-color:#c8cdd0;color:#1a1a1a}
:host-context(body.light-mode) .lk-sec-lbl{color:#888}
:host-context(body.light-mode) .nested-lk-list{border-left-color:#d0d4d8}
:host-context(body.light-mode) .lk-trigger{background:#fff;border-color:#c8cdd0;color:#1a1a1a}
:host-context(body.light-mode) .lk-trigger:hover{border-color:#00b4d8}
:host-context(body.light-mode) .lk-drop{background:#fff;border-color:#c8cdd0;box-shadow:0 6px 20px rgba(0,0,0,.12)}
:host-context(body.light-mode) .lk-drop-search{border-bottom-color:#eef0f2}
:host-context(body.light-mode) .lk-srch-inp{color:#1a1a1a}
:host-context(body.light-mode) .lk-drop-close{color:#007a95;background:#e0f5fa;border-color:rgba(0,180,216,.3)}
:host-context(body.light-mode) .lk-drop-list{scrollbar-color:#c8cdd0 #fff}
:host-context(body.light-mode) .lk-srch-item{color:#1a1a1a;border-bottom-color:#f0f2f5}
:host-context(body.light-mode) .lk-srch-item:hover{background:#eef7fb;color:#007a95}
:host-context(body.light-mode) .lk-srch-item.sel{background:#e0f5fa;color:#007a95}
:host-context(body.light-mode) .lk-srch-item-check{background:#f5f5f5;border-color:#c8cdd0}
:host-context(body.light-mode) .lk-srch-item.sel .lk-srch-item-check{background:#00b4d8;border-color:#00b4d8;color:#fff}
:host-context(body.light-mode) .lk-srch-item-sep{color:#aaa;background:#f8f9fa}
:host-context(body.light-mode) .lk-srch-empty{color:#aaa}
:host-context(body.light-mode) .lk-srch-item{color:#1a1a1a;border-bottom-color:#eef0f2}
:host-context(body.light-mode) .lk-srch-item:hover,:host-context(body.light-mode) .lk-srch-item.sel{background:#eef7fb;color:#007a95}
:host-context(body.light-mode) .lk-search-btn{background:#fff;border-color:#c8cdd0;color:#666}
:host-context(body.light-mode) .lk-search-btn:hover{border-color:#00b4d8;color:#007a95;background:#eef7fb}
:host-context(body.light-mode) .lk-search-btn.active{border-color:#00b4d8;color:#007a95;background:#e0f5fa}
:host-context(body.light-mode) .lk-sel-drop{background:#fff;border-color:#c8cdd0;box-shadow:0 6px 20px rgba(0,0,0,.12)}
:host-context(body.light-mode) .lk-sel-item{color:#1a1a1a;border-bottom-color:#f0f2f5}
:host-context(body.light-mode) .lk-sel-rm{color:#aaa}
:host-context(body.light-mode) .lk-sel-rm:hover{color:#cc3333;background:rgba(200,50,50,.08)}

/* sort */
:host-context(body.light-mode) .sort-row{background:#eef0f2;border-color:#dde0e3}
:host-context(body.light-mode) .sd{background:#fff;border-color:#c8cdd0;color:#888}
:host-context(body.light-mode) .ib{color:#888}
:host-context(body.light-mode) .ib:hover{background:#eef0f2;color:#1a1a1a}
:host-context(body.light-mode) .add-btn{border-color:#c8cdd0;color:#888}
:host-context(body.light-mode) .add-btn:hover{border-color:#00b4d8;color:#00b4d8;background:rgba(0,180,216,.04)}

/* aggregate rows */
:host-context(body.light-mode) .agg-row{background:#fff;border-color:#d0d4d8}
:host-context(body.light-mode) .agg-fn-sel{background-color:#fff;border-color:#c8cdd0;color:#1a1a1a;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='8' height='5'%3E%3Cpath d='M0 0l4 5 4-5z' fill='%23999'/%3E%3C/svg%3E")}
:host-context(body.light-mode) .agg-alias-inp{background:#fff;border-color:#c8cdd0;color:#1a1a1a}

/* options */
:host-context(body.light-mode) .opt-lbl{color:#5a6068}
:host-context(body.light-mode) .tog{background:#c8cdd0}

/* FetchXML editor */
:host-context(body.light-mode) .xml-label{background:#eef0f2;border-bottom-color:#dde0e3;color:#888}
:host-context(body.light-mode) .xml-editor-wrap{background:#f8f9fa}
:host-context(body.light-mode) #xml-editor{caret-color:#1a1a1a;scrollbar-color:#c8cdd0 #f8f9fa}
:host-context(body.light-mode) .xt-tag{color:#0066b8}
:host-context(body.light-mode) .xt-attr{color:#0050a0}
:host-context(body.light-mode) .xt-val{color:#1a7a40}
:host-context(body.light-mode) #val-msg.ok{background:#e8faf0;color:#1a7a40;border-top-color:rgba(26,122,64,.2)}
:host-context(body.light-mode) #val-msg.warn{background:#fff8e6;color:#a06000;border-top-color:rgba(190,120,0,.2)}

/* buttons */
:host-context(body.light-mode) .btn-secondary{background:#eef0f2;color:#444;border-color:#c8cdd0}
:host-context(body.light-mode) .btn-secondary:hover{background:#e4e7ea;color:#1a1a1a}
`;

// ── Static HTML ─────────────────────────────────────────────────────
const STATIC_HTML = `
<div id="fxb-root">
  <div class="tab-header">
    <div class="tab-item active" data-action="switchTab" data-tab="design">Design</div>
    <div class="tab-item" data-action="switchTab" data-tab="xml">FetchXML</div>
  </div>
  <div id="panes">
    <!-- DESIGN PANE -->
    <div class="tab-content active" id="pane-design">
      <div class="snippet-inputs" id="design-body">

        <!-- 1. TABLE -->
        <div class="content-card" id="sec-table">
          <div class="section-header open" data-action="secTog">
            <div class="step-badge">1</div>
            <div class="section-title">Table</div>
            <div class="section-meta" id="sm-table">—</div>
            <button class="ib" data-action="reset" style="font-size:10px;padding:2px 7px;flex-shrink:0">↺ Reset</button>
            <div class="section-chev">▾</div>
          </div>
          <div class="section-body">
            <div style="display:flex;gap:8px;align-items:flex-start">
              <div style="flex:1;min-width:0;position:relative">
                <label>Select table</label>
                <button class="dd-trigger" id="trig-entity" data-action="tog" data-key="entity">
                  <span class="dd-trigger-lbl ph" id="lbl-entity">Choose a table…</span>
                  <span class="dd-arrow">▾</span>
                </button>
                <div class="dd-popup dd-below" id="dd-entity">
                  <div class="dd-search">${SVG_SEARCH}<input type="text" placeholder="Search tables…" data-input-action="rl_ent"></div>
                  <div class="dd-list" id="list-entity"><div class="load-spinner">Loading tables…</div></div>
                </div>
              </div>
              <div style="flex:1;min-width:0;position:relative">
                <label style="color:#555">Select view <span style="font-size:10px;color:#333">(optional)</span></label>
                <button class="dd-trigger" id="trig-view" data-action="tog" data-key="view" disabled style="pointer-events:none;opacity:.35">
                  <span class="dd-trigger-lbl ph" id="lbl-view">Choose a view…</span>
                  <span class="dd-arrow">▾</span>
                </button>
                <div class="dd-popup dd-below" id="dd-view">
                  <div class="dd-search">${SVG_SEARCH}<input type="text" placeholder="Search views…" data-input-action="rl_view"></div>
                  <div class="dd-list" id="list-view"></div>
                </div>
              </div>
            </div>
            <div style="display:flex;gap:8px;margin-top:4px">
              <div style="flex:1;font-size:10.5px;color:#444;font-family:Consolas,monospace" id="sm-table-sub"></div>
              <div style="flex:1;font-size:10.5px;color:#444;font-family:Consolas,monospace" id="sm-view">—</div>
            </div>
          </div>
        </div>

        <!-- 2. COLUMNS -->
        <div class="content-card" id="sec-cols">
          <div class="section-header locked" data-action="secTog">
            <div class="step-badge">2</div>
            <div class="section-title">Columns</div>
            <div class="section-meta" id="sm-cols">0 selected</div>
            <div class="section-chev">▾</div>
          </div>
          <div class="section-body" style="position:relative">
            <div id="cols-chip-wrap">
              <label>Selected columns - click to open picker</label>
              <div class="col-chips" id="chips-primary" data-action="tog" data-key="cols"><span class="chip-ph">Click to select columns…</span></div>
              <div class="dd-popup dd-below" id="dd-cols">
                <div class="dd-search">${SVG_SEARCH}<input type="text" placeholder="Search columns…" data-input-action="rl_cols" data-lkid=""></div>
                <div class="dd-toolbar">
                  <button class="dd-tbar-btn" data-action="selAll" data-lkid="">Select All</button>
                  <button class="dd-tbar-btn" data-action="clrCols" data-lkid="">Clear</button>
                  <span class="dd-tbar-count" id="cnt-cols">0 selected</span>
                </div>
                <div class="dd-list" id="list-cols"></div>
              </div>
            </div>
            <div id="agg-rows-wrap" style="display:none">
              <div id="agg-rows-list"></div>
              <button class="add-btn sm" data-action="addAggRow">＋ Add aggregate field</button>
            </div>
          </div>
        </div>

        <!-- 3. FILTERS -->
        <div class="content-card" id="sec-filt">
          <div class="section-header locked" data-action="secTog">
            <div class="step-badge">3</div>
            <div class="section-title">Filters</div>
            <div class="section-meta" id="sm-filt">0 conditions</div>
            <div class="section-chev">▾</div>
          </div>
          <div class="section-body">
            <div id="filter-root"></div>
            <div style="display:flex;gap:4px;margin-top:5px">
              <button class="add-btn" style="flex:1" data-action="addCondRoot">＋ Condition</button>
              <button class="add-btn" style="flex:none;width:auto;padding:6px 14px" data-action="addGrpRoot">＋ Group</button>
            </div>
          </div>
        </div>

        <!-- 4. SORT -->
        <div class="content-card" id="sec-sort">
          <div class="section-header locked" data-action="secTog">
            <div class="step-badge">4</div>
            <div class="section-title">Sort Order</div>
            <div class="section-meta" id="sm-sort">0 sorts</div>
            <div class="section-chev">▾</div>
          </div>
          <div class="section-body">
            <div id="sort-list"></div>
            <div style="position:relative">
              <button class="add-btn" data-action="tog" data-key="sf">＋ Add sort field</button>
              <div class="dd-popup dd-top" id="dd-sf">
                <div class="dd-search">${SVG_SEARCH}<input type="text" placeholder="Search fields…" data-input-action="rl_sf"></div>
                <div class="dd-list" id="list-sf" style="max-height:200px"></div>
              </div>
            </div>
          </div>
        </div>

        <!-- 5. RELATED TABLES -->
        <div class="content-card" id="sec-links">
          <div class="section-header locked" data-action="secTog">
            <div class="step-badge">5</div>
            <div class="section-title">Related Tables</div>
            <div class="section-meta" id="sm-links">0 joined</div>
            <div class="section-chev">▾</div>
          </div>
          <div class="section-body">
            <div id="lk-root"></div>
            <div style="position:relative" id="add-root-lk-wrap">
              <button class="add-btn" data-action="togRelDD" data-lkid="">＋ Join related table</button>
              <div class="dd-popup dd-top" id="dd-rel-root">
                <div class="dd-search">${SVG_SEARCH}<input type="text" placeholder="Search relationships…" data-input-action="rl_rel" data-lkid=""></div>
                <div class="dd-list" id="list-rel-root" style="max-height:180px"></div>
              </div>
            </div>
          </div>
        </div>

        <!-- 6. OPTIONS -->
        <div class="content-card" id="sec-opts">
          <div class="section-header locked" data-action="secTog">
            <div class="step-badge">6</div>
            <div class="section-title">Query Options</div>
            <div class="section-meta" id="sm-opts"></div>
            <div class="section-chev">▾</div>
          </div>
          <div class="section-body">
            <div class="opt-row"><div class="opt-lbl">Aggregate Mode (GROUP BY / SUM / AVG…)</div><div class="tog" id="tog-agg" data-action="togOpt" data-key="aggregate"></div></div>
            <div class="opt-row" id="opt-row-distinct"><div class="opt-lbl">Distinct results only</div><div class="tog" id="tog-dist" data-action="togOpt" data-key="distinct"></div></div>
            <div class="opt-row"><div class="opt-lbl">Page size (count)</div><input class="num-in" type="number" id="opt-count" min="1" placeholder="—" data-input-action="optInput"></div>
            <div class="opt-row"><div class="opt-lbl">Page number</div><input class="num-in" type="number" id="opt-page" min="1" placeholder="—" data-input-action="optInput"></div>
          </div>
        </div>

        <div style="height:6px"></div>
      </div>
    </div>

    <!-- FETCHXML PANE -->
    <div class="tab-content" id="pane-xml">
      <div class="xml-label">
        <span style="flex:1"></span>
        <button id="btn-copy-xml" class="btn-secondary" data-action="copyXML" style="display:none">Copy ⎘</button>
        <button id="btn-apply-xml" class="btn-secondary" style="margin-left:4px;display:none" data-action="loadFromXML">← Apply to Design</button>
      </div>
      <div class="xml-editor-wrap">
        <pre id="xml-highlight" aria-hidden="true"></pre>
        <textarea id="xml-editor" spellcheck="false" placeholder="Already have a FetchXML query? Paste it here.&#10;&#10;&lt;fetch&gt;&#10;  &lt;entity name='account'&gt;...&lt;/entity&gt;&#10;&lt;/fetch&gt;" data-input-action="xmlEdit"></textarea>
      </div>
      <div id="val-msg" style="display:none"></div>
    </div>
  </div>
</div>
`;

// ── Main export ─────────────────────────────────────────────────────
export function render(container: HTMLElement, callbacks: BuilderCallbacks): void {
  let shadow = container.shadowRoot;
  if (!shadow) shadow = container.attachShadow({ mode: 'open' });
  shadow.innerHTML = '';

  const styleEl = document.createElement('style');
  styleEl.textContent = CSS;
  shadow.appendChild(styleEl);

  const wrapper = document.createElement('div');
  wrapper.innerHTML = STATIC_HTML;
  shadow.appendChild(wrapper.firstElementChild!);

  const $ = (id: string) => shadow!.getElementById(id);

  const ensureColumnPopupCloseButtons = () => {
    shadow!.querySelectorAll<HTMLElement>('.dd-popup .dd-search').forEach(search => {
      const inp = search.querySelector<HTMLInputElement>('input[data-input-action="rl_cols"]');
      if (!inp || search.querySelector('.dd-close')) return;
      const popup = search.closest<HTMLElement>('.dd-popup');
      if (!popup?.id) return;
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'dd-close';
      btn.title = 'Close';
      btn.textContent = '×';
      btn.dataset.action = 'closeDD';
      btn.dataset.ddid = popup.id;
      search.appendChild(btn);
    });
  };
  ensureColumnPopupCloseButtons();

  // ── Data ──────────────────────────────────────────────────────────
  const DATA: {
    entities: { name: string; display: string }[];
    meta: Record<string, { attrs: { n: string; d: string; t: string; targets?: string[]; primary?: boolean }[]; rels: any[]; views?: any[]; primaryName?: string; primaryId?: string; objectTypeCode?: number }>;
    optionSets: Record<string, { v: number; l: string }[]>;
    loaded: boolean;
    loading: boolean;
  } = { entities: [], meta: {}, optionSets: {}, loaded: false, loading: false };

  // ── State ─────────────────────────────────────────────────────────
  let _ic = 0;
  const nid  = () => 'x' + (++_ic);
  const newG = (l = 'and') => ({ id: nid(), logic: l, conds: [] as any[], kids: [] as any[] });
  const newC = () => ({ id: nid(), field: '', op: 'eq', val: '', valLabel: '', type: '' });
  const newL = (rel: any, alias: string, jt: string) =>
    ({ id: nid(), rel, alias, joinType: jt, fields: [] as string[], filter: newG('and'), links: [] as any[] });

  const S: any = {
    entity: null as string | null,
    fields: [] as { attr: string; alias: string | null; aggr?: string }[],
    links:  [] as any[],
    rootF:  newG('and'),
    sorts:  [] as { attr: string; alias: string | null; desc: boolean; isAggAlias?: boolean }[],
    opts:   { distinct: false, aggregate: false, count: '', page: '' },
    findLink(links: any[], id: string): any {
      for (const lk of links) { if (lk.id === id) return lk; const f = this.findLink(lk.links||[], id); if (f) return f; } return null;
    },
    removeLink(links: any[], id: string): boolean {
      const i = links.findIndex((l: any) => l.id === id); if (i >= 0) { links.splice(i, 1); return true; }
      for (const lk of links) { if (this.removeLink(lk.links||[], id)) return true; } return false;
    },
    countLinks(links: any[]): number { let c = links.length; links.forEach((lk:any)=>c+=this.countLinks(lk.links||[])); return c; },
    collectSortFields(links: any[], acc: any[] = []): any[] {
      links.forEach((lk:any) => { (lk.fields||[]).forEach((a:string)=>acc.push({attr:a,alias:lk.alias,entName:lk.rel.toEntity})); this.collectSortFields(lk.links||[],acc); }); return acc;
    },
    init() { this.entity=null; this.fields=[]; this.links=[]; this.rootF=newG('and'); this.sorts=[]; this.opts={distinct:false,aggregate:false,count:'',page:''}; },
  };

  // ── Generator ─────────────────────────────────────────────────────
  const Gen = {
    run(): string | null {
      if (!S.entity) return null;
      return S.opts.aggregate ? this.runAggregate() : this.runNormal();
    },
    runNormal(): string | null {
      if (!S.entity) return null;
      const fa = ['version="1.0"','output-format="xml-platform"','mapping="logical"'];
      if (S.opts.distinct) fa.push('distinct="true"');
      if (S.opts.count) fa.push(`count="${S.opts.count}"`);
      if (S.opts.count && S.opts.page) fa.push(`page="${S.opts.page}"`);
      const L = [`<fetch ${fa.join(' ')}>`, `  <entity name="${S.entity}">`];
      const pf = S.fields.filter((f:any)=>!f.alias);
      if (!pf.length) L.push('    <all-attributes/>');
      else pf.forEach((f:any)=>L.push(`    <attribute name="${f.attr}"/>`));
      S.sorts.filter((s:any)=>!s.alias).forEach((s:any)=>L.push(`    <order attribute="${s.attr}" descending="${s.desc?'true':'false'}"/>`));
      L.push(...this.filt(S.rootF, 4, S.entity || ''));
      S.links.forEach((lk:any)=>L.push(...this.link(lk, 4)));
      L.push('  </entity>', '</fetch>');
      return L.join('\n');
    },
    runAggregate(): string | null {
      if (!S.entity) return null;
      const fa = ['version="1.0"','output-format="xml-platform"','mapping="logical"','aggregate="true"'];
      if (S.opts.count) fa.push(`count="${S.opts.count}"`);
      if (S.opts.count && S.opts.page) fa.push(`page="${S.opts.page}"`);
      const L = [`<fetch ${fa.join(' ')}>`, `  <entity name="${S.entity}">`];
      S.fields.forEach((f:any) => {
        if (!f.alias) return; // skip rows missing alias — they generate a warning
        if (f.aggr === 'groupby') {
          L.push(`    <attribute name="${f.attr}" groupby="true" alias="${f.alias}"/>`);
        } else if (f.aggr === 'count') {
          const attr = f.attr || S.entity;
          L.push(`    <attribute name="${attr}" aggregate="count" alias="${f.alias}"/>`);
        } else if (f.aggr) {
          L.push(`    <attribute name="${f.attr}" aggregate="${f.aggr}" alias="${f.alias}"/>`);
        }
      });
      S.sorts.forEach((s:any) => {
        if (s.isAggAlias) {
          L.push(`    <order alias="${s.attr}" descending="${s.desc?'true':'false'}"/>`);
        }
      });
      L.push(...this.filt(S.rootF, 4, S.entity || ''));
      L.push('  </entity>', '</fetch>');
      return L.join('\n');
    },
    filt(g: any, i: number, entName: string): string[] {
      const p = ' '.repeat(i);
      const vc = (g.conds||[]).filter((c:any)=>c.field&&c.op&&(NO_VAL.has(c.op)||c.val!==''));
      const vk = (g.kids||[]).filter((k:any)=>this.hc(k));
      if (!vc.length&&!vk.length) return [];
      const L = [`${p}<filter type="${g.logic}">`];
      const esc = (v:string)=>String(v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
      vc.forEach((c:any)=>{
        const lw = LIKE_WRAP[c.op];
        const attrType = (DATA.meta[entName]?.attrs||[]).find((a:any)=>a.n===c.field)?.t || c.type || '';
        // Text-style filters on option sets and lookup-style fields target the virtual *name attribute.
        const fld = (lw && NAME_ATTR_TYPES.has(attrType)) ? c.field+'name' : c.field;
        if (NO_VAL.has(c.op)) {
          L.push(`${p}  <condition attribute="${c.field}" operator="${c.op}"/>`);
        } else if (lw) {
          L.push(`${p}  <condition attribute="${fld}" operator="${lw.op}" value="${esc(lw.pre + c.val + lw.suf)}"/>`);
        } else if (MULTI_VAL.has(c.op)) {
          const vals = c.val.split(',').map((v:string)=>v.trim()).filter(Boolean);
          if (!vals.length) return;
          L.push(`${p}  <condition attribute="${c.field}" operator="${c.op}">`);
          vals.forEach((v:string)=>L.push(`${p}    <value>${esc(v)}</value>`));
          L.push(`${p}  </condition>`);
        } else if ((c.op==='eq'||c.op==='ne') && c.val && c.val.includes(',')) {
          // Multi-selected lookup values → in / not-in with <value> children
          const vals = c.val.split(',').map((v:string)=>v.trim()).filter(Boolean);
          const multiOp = c.op==='eq'?'in':'not-in';
          // Parse entity (uitype) per value from valLabel (format: "name|entity")
          const labels = (c.valLabel||'').split('||');
          const entityPerVal = labels.map((lbl:string)=>{ const sep=lbl.lastIndexOf('|'); return sep>0?lbl.slice(sep+1):''; });
          L.push(`${p}  <condition attribute="${c.field}" operator="${multiOp}">`);
          vals.forEach((v:string,idx:number)=>{
            const ut=entityPerVal[idx];
            L.push(`${p}    <value${ut?` uitype="${ut}"`:''} >${esc(v)}</value>`);
          });
          L.push(`${p}  </condition>`);
        } else {
          // Single lookup value — check for entity (uitype) in valLabel "name|entity"
          const lbl0=(c.valLabel||'').split('||')[0]||'';
          const sep=lbl0.lastIndexOf('|');
          const ut0=sep>0?lbl0.slice(sep+1):'';
          const uitypeAttr=ut0?` uitype="${ut0}"`:'';
          L.push(`${p}  <condition attribute="${c.field}" operator="${c.op}"${uitypeAttr} value="${esc(c.val)}"/>`);
        }
      });
      vk.forEach((k:any)=>L.push(...this.filt(k,i+2,entName)));
      L.push(`${p}</filter>`); return L;
    },
    hc(g:any): boolean { return (g.conds||[]).some((c:any)=>c.field&&c.op&&(NO_VAL.has(c.op)||c.val!==''))||(g.kids||[]).some((k:any)=>this.hc(k)); },
    link(lk: any, i: number): string[] {
      const p = ' '.repeat(i); const r = lk.rel;
      const al = lk.alias ? ` alias="${lk.alias}"` : '';
      const L = [`${p}<link-entity name="${r.toEntity}" from="${r.toAttr}" to="${r.fromAttr}" link-type="${lk.joinType}"${al}>`];
      (lk.fields||[]).forEach((a:string)=>L.push(`${p}  <attribute name="${a}"/>`));
      S.sorts.filter((s:any)=>s.alias===lk.alias).forEach((s:any)=>L.push(`${p}  <order attribute="${s.attr}" descending="${s.desc?'true':'false'}"/>`));
      if (lk.filter&&this.hc(lk.filter)) L.push(...this.filt(lk.filter,i+2,lk.rel.toEntity));
      (lk.links||[]).forEach((ch:any)=>L.push(...this.link(ch,i+2)));
      L.push(`${p}</link-entity>`); return L;
    },
  };

  // ── Parser ────────────────────────────────────────────────────────
  const Parser = {
    parse(xml: string): any {
      const doc = new DOMParser().parseFromString(xml.trim(), 'text/xml');
      if (doc.querySelector('parsererror')) return { err: 'XML parse error' };
      const fe = doc.querySelector('fetch'); const en = fe&&fe.querySelector(':scope>entity');
      if (!en) return { err: 'No <entity> element found' };
      const eName = en.getAttribute('name')!;
      const isAgg = fe!.getAttribute('aggregate') === 'true';
      const opts = { distinct: !isAgg && fe!.getAttribute('distinct')==='true', aggregate: isAgg, count: fe!.getAttribute('count')||'', page: fe!.getAttribute('page')||'' };
      const attrs = [...en.querySelectorAll(':scope>attribute')].map(a=>a.getAttribute('name')!);
      // Aggregate field objects (used when aggregate=true)
      const aggFields: any[] = [...en.querySelectorAll(':scope>attribute')].map(a => {
        const name = a.getAttribute('name') || '';
        const alias = a.getAttribute('alias') || '';
        const groupby = a.getAttribute('groupby') === 'true';
        const aggAttr = a.getAttribute('aggregate') || '';
        let aggr = '';
        if (groupby) aggr = 'groupby';
        else if (aggAttr) aggr = aggAttr;
        else aggr = 'groupby'; // fallback
        return { attr: name, alias: alias || null, aggr };
      });
      const rootF = newG('and');
      [...en.querySelectorAll(':scope>filter')].forEach(f=>this.pF(f,rootF));
      const orders: any[] = [...en.querySelectorAll(':scope>order')].map(o=>{
        const aliasAttr = o.getAttribute('alias');
        const attrAttr = o.getAttribute('attribute');
        if (aliasAttr && !attrAttr) {
          return { attr: aliasAttr, alias: null, desc: o.getAttribute('descending')==='true', isAggAlias: true };
        }
        return { attr: attrAttr, alias: null, desc: o.getAttribute('descending')==='true' };
      });
      const links = [...en.querySelectorAll(':scope>link-entity')].map(le=>this.pL(le,orders));
      return { eName, attrs, aggFields, rootF, orders, links, opts };
    },
    pF(el: Element, parent: any) {
      const g = newG(el.getAttribute('type')||'and');
      [...el.querySelectorAll(':scope>condition')].forEach(ce=>{
        const c=newC();
        c.field=ce.getAttribute('attribute')||'';
        c.op=ce.getAttribute('operator')||'eq';
        const singleVal=ce.getAttribute('value');
        if (singleVal !== null) {
          c.val = singleVal;
        } else {
          // Multi-value condition: <value> children (in, not-in, contain-values, etc.)
          c.val = [...ce.querySelectorAll(':scope>value')].map(v=>v.textContent||'').join(',');
        }
        // If attribute is a virtual *name field (e.g. cs_authcodename), strip suffix → base field
        // and mark as option set type so Gen.filt re-appends it on output
        if ((c.op === 'like' || c.op === 'not-like') && c.field.endsWith('name') && c.field.length > 4) {
          c.field = c.field.slice(0, -4);
          c.type = 'picklist'; // marks it as option set so Gen.filt uses *name attr again
        }
        // Reverse-translate like/not-like + % patterns → virtual ops
        if (c.op === 'like' || c.op === 'not-like') {
          const neg = c.op === 'not-like';
          const v = c.val as string;
          if (v.startsWith('%') && v.endsWith('%')) {
            c.op = neg ? 'not-contains' : 'contains'; c.val = v.slice(1, -1);
          } else if (v.endsWith('%')) {
            c.op = neg ? 'not-starts-with' : 'starts-with'; c.val = v.slice(0, -1);
          } else if (v.startsWith('%')) {
            c.op = neg ? 'not-ends-with' : 'ends-with'; c.val = v.slice(1);
          }
        }
        // Translate native FetchXML begins-with/ends-with → virtual ops
        if (c.op === 'begins-with') c.op = 'starts-with';
        if (c.op === 'not-begins-with') c.op = 'not-starts-with';
        if (c.op === 'ends-with') c.op = 'ends-with'; // already mapped
        if (c.op === 'not-ends-with') c.op = 'not-ends-with'; // already mapped
        g.conds.push(c);
      });
      [...el.querySelectorAll(':scope>filter')].forEach(fe=>this.pF(fe,g));
      (parent.kids||(parent.kids=[])).push(g);
    },
    pL(le: Element, orders: any[]): any {
      const toEnt=le.getAttribute('name')!; const fromAttr=le.getAttribute('to')!; const toAttr=le.getAttribute('from')!;
      const alias=le.getAttribute('alias')||toEnt; const joinType=le.getAttribute('link-type')||'inner';
      const fields=[...le.querySelectorAll(':scope>attribute')].map(a=>a.getAttribute('name')!);
      const rel={name:'',display:toEnt,fromAttr,toEntity:toEnt,toAttr};
      [...le.querySelectorAll(':scope>order')].forEach(o=>orders.push({attr:o.getAttribute('attribute'),alias,desc:o.getAttribute('descending')==='true'}));
      const filter=newG('and');
      [...le.querySelectorAll(':scope>filter')].forEach(f=>Parser.pF(f,filter));
      const links=[...le.querySelectorAll(':scope>link-entity')].map(nle=>Parser.pL(nle,orders));
      return {id:nid(),rel,alias,joinType,fields,filter,links};
    },
  };

  // ── XML button visibility ─────────────────────────────────────────
  function syncXmlButtons() {
    const hasXml = !!($('xml-editor') as HTMLTextAreaElement)?.value.trim();
    const d = hasXml ? '' : 'none';
    const btnCopy  = $('btn-copy-xml')  as HTMLElement | null;
    const btnApply = $('btn-apply-xml') as HTMLElement | null;
    if (btnCopy)  btnCopy.style.display  = d;
    if (btnApply) btnApply.style.display = d;
  }

  // ── XML Highlighter ───────────────────────────────────────────────
  let _hlScrollBound = false;
  function hlXML() {
    const ta=$('xml-editor') as HTMLTextAreaElement; const pre=$('xml-highlight');
    if (!ta||!pre) return;
    if (!_hlScrollBound) {
      _hlScrollBound = true;
      ta.addEventListener('scroll', () => { pre.scrollTop = ta.scrollTop; pre.scrollLeft = ta.scrollLeft; });
    }
    pre.innerHTML = ta.value
      .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;')
      .replace(/&gt;([^&<]+)&lt;/g,'&gt;<span class="xt-val">$1</span>&lt;')
      .replace(/(&lt;\/?)([\w:-]+)/g,'<span class="xt-tag">$1$2</span>')
      .replace(/(\/?&gt;)/g,'<span class="xt-tag">$1</span>')
      .replace(/\s([\w:-]+)=(&quot;[^&]*&quot;)/g,' <span class="xt-attr">$1=</span><span class="xt-val">$2</span>');
    pre.scrollTop = ta.scrollTop;
    pre.scrollLeft = ta.scrollLeft;
  }

  // ── UI helpers ────────────────────────────────────────────────────
  let _openDD: string | null = null;

  function closeAll() {
    shadow!.querySelectorAll('.dd-popup').forEach(d=>d.classList.remove('open'));
    shadow!.querySelectorAll('.dd-trigger').forEach(t=>t.classList.remove('open'));
    shadow!.querySelectorAll('.col-chips').forEach(c=>c.classList.remove('focused'));
    _openDD = null;
  }

  function togDD(key: string) {
    const dd = $('dd-'+key); if (!dd) return;
    const was = dd.classList.contains('open'); closeAll(); if (was) return;
    dd.classList.add('open'); _openDD = key;
    $('trig-'+key)?.classList.add('open');
    $('chips-'+key)?.classList.add('focused');
    const inp = dd.querySelector<HTMLInputElement>('input[type=text]');
    if (inp) { inp.value=''; inp.focus(); }
    if (key==='entity') rlEnt('');
    else if (key==='view') rlView('');
    else if (key==='cols') rlCols('',null);
    else if (key==='sf') rlSf('');
  }

  function togRelDD(parentLkId: string|null) {
    const ddId='dd-rel-'+(parentLkId||'root'); const dd=$(ddId); if(!dd) return;
    const was=dd.classList.contains('open'); closeAll(); if(was) return;
    dd.classList.add('open'); _openDD=ddId;
    const inp=dd.querySelector<HTMLInputElement>('input'); if(inp){inp.value='';inp.focus();}
    rlRel(parentLkId,'');
  }

  function togLkCols(lkId: string) {
    const dd=$('dd-lkc-'+lkId); if(!dd) return;
    const was=dd.classList.contains('open'); closeAll(); if(was) return;
    dd.classList.add('open'); _openDD='lkc'+lkId;
    $('chips-lk-'+lkId)?.classList.add('focused');
    const inp=dd.querySelector<HTMLInputElement>('input'); if(inp){inp.value='';inp.focus();}
    rlCols('',lkId);
  }

  function rlEnt(q: string) {
    const el=$('list-entity'); if(!el) return;
    if (DATA.loading) { el.innerHTML='<div class="load-spinner">Loading tables…</div>'; return; }
    const lq=q.toLowerCase();
    el.innerHTML = DATA.entities.filter(e=>!lq||e.name.includes(lq)||e.display.toLowerCase().includes(lq))
      .map(e=>`<div class="dd-item${S.entity===e.name?' sel':''}" data-action="selEnt" data-name="${e.name}">
        <div class="ent-icon">${e.display[0].toUpperCase()}</div>
        <div class="ent-names"><div class="ent-display">${e.display}</div><div class="ent-logical">${e.name}</div></div>
      </div>`).join('') || '<div class="hint">No tables found</div>';
  }

  function rlView(q: string) {
    if (!S.entity) return; const el=$('list-view'); if(!el) return;
    const views=DATA.meta[S.entity]?.views||[]; const lq=q.toLowerCase();
    el.innerHTML=views.filter(v=>!lq||v.name.toLowerCase().includes(lq))
      .map(v=>`<div class="dd-item" data-action="loadView" data-vid="${v.id}">
        <span class="view-badge ${v.type==='P'?'vb-p':'vb-s'}">${v.type==='P'?'P':'S'}</span>
        <div class="ent-names"><div class="ent-display">${v.name}</div></div>
      </div>`).join('')||'<div class="hint">No views</div>';
  }

  function rlRel(parentLkId: string|null, q: string) {
    const listId='list-rel-'+(parentLkId||'root'); const el=$(listId); if(!el) return;
    let entName=S.entity; if(parentLkId){const lk=S.findLink(S.links,parentLkId);if(lk)entName=lk.rel.toEntity;}
    if(!entName){el.innerHTML='<div class="hint">No table selected</div>';return;}
    const rels=DATA.meta[entName]?.rels||[]; const lq=q.toLowerCase();
    const parentAttr=parentLkId?`data-parent-lkid="${parentLkId}"`:'data-parent-lkid=""';
    el.innerHTML=rels.filter(r=>!lq||r.display.toLowerCase().includes(lq)||r.toEntity.includes(lq))
      .map(r=>`<div class="dd-item" data-action="addLink" data-relname="${r.name}" data-parent-ent="${entName}" ${parentAttr}>
        <div class="ent-names">
          <div class="ent-display">${r.display} <span style="color:#444;font-size:10px">(${r.toEntity})</span></div>
          <div class="ent-logical">${entName}.${r.fromAttr} → ${r.toEntity}.${r.toAttr}</div>
        </div>
      </div>`).join('')||'<div class="hint">No relationships found</div>';
  }

  function rlCols(q: string, lkId: string|null) {
    const entName=lkId?S.findLink(S.links,lkId)?.rel.toEntity:S.entity;
    const meta=entName?DATA.meta[entName]:null; const lq=q.toLowerCase();
    const lid=lkId?'list-lkc-'+lkId:'list-cols'; const cid=lkId?'cnt-lkc-'+lkId:'cnt-cols';
    const el=$(lid); if(!el) return;
    const sel:string[]=lkId?(S.findLink(S.links,lkId)?.fields||[]):S.fields.filter((f:any)=>!f.alias).map((f:any)=>f.attr);
    const isSpecialCol = (s: string) => !/^[a-zA-Z]/i.test(s);
    const attrs=(meta?.attrs||[]).filter(a=>!lq||a.d.toLowerCase().includes(lq)||a.n.includes(lq)).slice().sort((a,b)=>{
      const as=isSpecialCol(a.d), bs=isSpecialCol(b.d);
      if (as!==bs) return as?1:-1;
      return a.d.localeCompare(b.d);
    });
    const lkAttr=lkId?`data-lkid="${lkId}"`:'data-lkid=""';
    el.innerHTML=attrs.map(a=>{const chk=sel.includes(a.n);
      return `<div class="col-item${chk?' chk':''}" data-action="togCol" data-attr="${a.n}" ${lkAttr}>
        <input type="checkbox"${chk?' checked':''}><span class="type-badge tb-${a.t}">${a.t}</span>
        <div class="col-names"><div class="col-disp">${a.d}</div><div class="col-logical">${a.n}</div></div>
      </div>`;}).join('')||'<div class="hint">No attributes found</div>';
    const ce=$(cid); if(ce) ce.textContent=sel.length+' selected';
  }

  function rlSf(q: string) {
    if(!S.entity) return; const el=$('list-sf'); if(!el) return;
    const lq=q.toLowerCase();

    // Aggregate mode: list aliases from agg fields as sort targets
    if (S.opts.aggregate) {
      const items = S.fields.filter((f:any) => f.alias && (!lq || f.alias.includes(lq) || f.attr.includes(lq)));
      el.innerHTML = items.map((f:any) => {
        const fnLabel = f.aggr === 'groupby' ? 'GROUP BY' : (f.aggr||'').toUpperCase();
        return `<div class="col-item" data-action="addAggSort" data-alias="${f.alias}">
          <span class="type-badge tb-default">${fnLabel}</span>
          <div class="col-names"><div class="col-disp" style="color:#d4893a">${f.alias}</div><div class="col-logical">${f.attr||'—'}</div></div>
        </div>`;
      }).join('') || '<div class="hint">No aggregate aliases — add fields first</div>';
      return;
    }

    const items:any[]=[];
    (DATA.meta[S.entity]?.attrs||[]).filter(a=>!lq||a.d.toLowerCase().includes(lq)||a.n.includes(lq))
      .forEach(a=>items.push({attr:a.n,d:a.d,t:a.t,alias:null}));
    S.collectSortFields(S.links).filter((f:any)=>!lq||f.attr.includes(lq)||f.entName.includes(lq))
      .forEach((f:any)=>{const la=(DATA.meta[f.entName]?.attrs||[]).find((a:any)=>a.n===f.attr);items.push({attr:f.attr,d:la?.d||f.attr,t:la?.t||'string',alias:f.alias});});
    items.sort((a,b)=>{const as=!/^[a-zA-Z]/i.test(a.d),bs=!/^[a-zA-Z]/i.test(b.d);if(as!==bs)return as?1:-1;return a.d.localeCompare(b.d);});
    el.innerHTML=items.map(it=>`<div class="col-item" data-action="addSort" data-attr="${it.attr}" data-alias="${it.alias||''}">
      <span class="type-badge tb-${it.t}">${it.t}</span>
      <div class="col-names"><div class="col-disp">${it.alias?`[${it.alias}] `:''}${it.d}</div><div class="col-logical">${it.attr}</div></div>
    </div>`).join('')||'<div class="hint">No fields</div>';
  }

  // ── Aggregate helpers ─────────────────────────────────────────────
  function autoAlias(aggr: string, attr: string): string {
    if (aggr === 'groupby')      return 'grp_' + attr;
    if (aggr === 'count')        return 'count_all';
    if (aggr === 'countcolumn')  return 'cnt_' + attr;
    return aggr + '_' + attr;
  }

  function renderAggRows() {
    const el = $('agg-rows-list'); if (!el) return;
    const attrs = S.entity ? (DATA.meta[S.entity]?.attrs || []) : [];
    if (!S.fields.length) {
      el.innerHTML = '<div class="hint">No aggregate fields — click ＋ to add one</div>';
    } else {
      el.innerHTML = S.fields.map((f: any, i: number) => {
        const isGrp = f.aggr === 'groupby';
        const cls = isGrp ? 'is-groupby' : 'is-agg';
        const attrOpts = attrs.map((a: any) =>
          `<option value="${a.n}"${f.attr === a.n ? ' selected' : ''}>${a.d}</option>`
        ).join('');
        const fnVal = f.aggr || 'groupby';
        return `<div class="agg-row ${cls}">
          <select class="cr-s" style="flex:1" data-action="setAggField" data-idx="${i}">
            <option value="">— Field —</option>${attrOpts}
          </select>
          <select class="agg-fn-sel" data-action="setAggFn" data-idx="${i}">
            <option value="groupby"${fnVal==='groupby'?' selected':''}>Group By</option>
            <option value="sum"${fnVal==='sum'?' selected':''}>SUM</option>
            <option value="avg"${fnVal==='avg'?' selected':''}>AVG</option>
            <option value="count"${fnVal==='count'?' selected':''}>COUNT(*)</option>
            <option value="countcolumn"${fnVal==='countcolumn'?' selected':''}>COUNT(Col)</option>
            <option value="min"${fnVal==='min'?' selected':''}>MIN</option>
            <option value="max"${fnVal==='max'?' selected':''}>MAX</option>
          </select>
          <input class="agg-alias-inp${!f.alias?' invalid':''}" value="${f.alias||''}" placeholder="alias" data-action="setAggAlias" data-idx="${i}">
          <button class="ib del" data-action="rmAggRow" data-idx="${i}">✕</button>
        </div>`;
      }).join('');
    }
    // Update section meta
    const sm = $('sm-cols');
    if (sm) {
      const grpCnt = S.fields.filter((f: any) => f.aggr === 'groupby').length;
      const aggCnt = S.fields.filter((f: any) => f.aggr && f.aggr !== 'groupby').length;
      const parts: string[] = [];
      if (grpCnt) parts.push(grpCnt + ' group');
      if (aggCnt) parts.push(aggCnt + ' agg');
      sm.innerHTML = parts.length
        ? parts.join(', ') + ' <span class="agg-mode-badge">AGG</span>'
        : '<span class="agg-mode-badge">AGG</span>';
    }
  }

  function renderChips(lkId: string|null) {
    const entName=lkId?S.findLink(S.links,lkId)?.rel.toEntity:S.entity;
    const meta=entName?DATA.meta[entName]:null;
    const sel:string[]=lkId?(S.findLink(S.links,lkId)?.fields||[]):S.fields.filter((f:any)=>!f.alias).map((f:any)=>f.attr);
    const el=$(lkId?'chips-lk-'+lkId:'chips-primary'); if(!el) return;
    if(!sel.length){el.innerHTML='<span class="chip-ph">Click to select columns…</span>';return;}
    const MAX=7,show=sel.slice(0,MAX),rest=sel.length-MAX;
    const lkAttr=lkId?`data-lkid="${lkId}"`:'data-lkid=""';
    el.innerHTML=show.map(an=>{const a=meta?.attrs.find(x=>x.n===an);
      return `<span class="chip"><span title="${a?a.d:an}">${an}</span><span class="chip-x" data-action="togCol" data-attr="${an}" ${lkAttr}>×</span></span>`;
    }).join('')+(rest>0?`<span class="chip-more">+${rest}</span>`:'');
    if(!lkId){const sm=$('sm-cols');if(sm)sm.textContent=sel.length+' selected';}
  }

  function renderFilters() {
    const el=$('filter-root'); if(!el) return; el.innerHTML=''; if(!S.entity) return;
    el.appendChild(buildFG(S.rootF,true,S.entity,null));
    const cnt=cntC(S.rootF); const sm=$('sm-filt'); if(sm) sm.textContent=cnt+' condition'+(cnt!==1?'s':'');
  }

  function cntC(g:any):number{let c=(g.conds||[]).filter((c:any)=>c.field&&c.op).length;(g.kids||[]).forEach((k:any)=>c+=cntC(k));return c;}

  function buildFG(g:any, isRoot:boolean, entName:string, lkId:string|null): HTMLElement {
    const lkAttr=lkId?`data-lkid="${lkId}"`:'data-lkid=""';
    const _rawAttrs=(DATA.meta[entName]||DATA.meta[S.entity!]||{attrs:[]}).attrs||[];
    const attrs=[..._rawAttrs].sort((a,b)=>{const as=!/^[a-zA-Z]/i.test(a.d),bs=!/^[a-zA-Z]/i.test(b.d);if(as!==bs)return as?1:-1;return a.d.localeCompare(b.d);});
    const div=document.createElement('div'); div.className='filter-group'; div.id='fg_'+g.id;
    div.innerHTML=`<div class="fg-bar">
      <div class="logic-tog">
        <button class="lg-b${g.logic==='and'?' on':''}" data-action="setL" data-gid="${g.id}" data-logic="and" ${lkAttr}>AND</button>
        <button class="lg-b${g.logic==='or'?' on':''}" data-action="setL" data-gid="${g.id}" data-logic="or" ${lkAttr}>OR</button>
      </div>
      <span class="fg-lbl">${isRoot?'Root filter':''}</span>
      <div class="fg-btns">
        <button class="ib" data-action="addCond" data-gid="${g.id}" ${lkAttr}>＋</button>
        <button class="ib" data-action="addGrp" data-gid="${g.id}" ${lkAttr} style="font-size:9px">＋grp</button>
        ${!isRoot?`<button class="ib del" data-action="rmGrp" data-gid="${g.id}" ${lkAttr}>✕</button>`:''}
      </div>
    </div><div id="conds_${g.id}"></div>`;
    const condsEl=div.querySelector('#conds_'+g.id)!;
    (g.conds||[]).forEach((c:any)=>condsEl.appendChild(buildCondRow(c,g.id,attrs,lkId,entName)));
    if((g.kids||[]).length){const kd=document.createElement('div');kd.style.cssText='padding:0 8px 5px;border-top:1px dashed #2a2a2a';(g.kids||[]).forEach((k:any)=>kd.appendChild(buildFG(k,false,entName,lkId)));div.appendChild(kd);}
    return div;
  }

  const TYPE_COLOR: Record<string,string> = {
    string:'#7dd3fc',memo:'#7dd3fc',
    integer:'#86efac',bigint:'#86efac',decimal:'#86efac',double:'#86efac',money:'#86efac',
    datetime:'#c4b5fd',
    boolean:'#fcd34d',
    picklist:'#fb923c',state:'#fb923c',status:'#fb923c',
    multiselect:'#f472b6',
    lookup:'#00b4d8',customer:'#00b4d8',owner:'#00b4d8',
    uniqueidentifier:'#9ca3af',
  };
  const TYPE_SHORT_FP: Record<string,string> = {
    string:'Text',memo:'Memo',integer:'Int',bigint:'BigInt',decimal:'Dec',
    double:'Float',money:'Money',datetime:'Date',boolean:'Bool',
    picklist:'Option',state:'State',status:'Status',multiselect:'Multi',
    lookup:'Lookup',customer:'Customer',owner:'Owner',uniqueidentifier:'GUID',
  };

  function buildFieldPicker(c:any, attrs:any[], gId:string, lkId:string|null): HTMLElement {
    const wrap=document.createElement('div'); wrap.className='fp-wrap';

    const trigger=document.createElement('div'); trigger.className='fp-trigger';
    const trigName=document.createElement('span'); trigName.className='fp-trigger-name';
    const trigBadge=document.createElement('span'); trigBadge.className='fp-trigger-badge';
    const trigArr=document.createElement('span'); trigArr.className='fp-trigger-arr'; trigArr.textContent='▾';
    trigger.appendChild(trigName); trigger.appendChild(trigBadge); trigger.appendChild(trigArr);
    wrap.appendChild(trigger);

    const drop=document.createElement('div'); drop.className='fp-drop'; wrap.appendChild(drop);
    const searchInp=document.createElement('input') as HTMLInputElement;
    searchInp.className='fp-search'; searchInp.type='text'; searchInp.placeholder='Search fields…';
    drop.appendChild(searchInp);
    const list=document.createElement('div'); list.className='fp-list'; drop.appendChild(list);

    const updateTrigger=()=>{
      const sel=attrs.find((a:any)=>a.n===c.field);
      if(sel){
        trigName.textContent=sel.d; trigName.style.opacity='1';
        trigBadge.textContent=TYPE_SHORT_FP[sel.t]||sel.t;
        trigBadge.style.color=TYPE_COLOR[sel.t]||'#9ca3af';
        trigBadge.style.borderColor=(TYPE_COLOR[sel.t]||'#9ca3af')+'55';
        trigBadge.style.display='';
      } else {
        trigName.textContent=c.field?`[${c.field}]`:'— Field —'; trigName.style.opacity=c.field?'1':'0.45';
        trigBadge.style.display='none';
      }
    };

    const renderList=(q:string)=>{
      list.innerHTML='';
      const lq=q.toLowerCase();
      const filtered=q?attrs.filter((a:any)=>a.d.toLowerCase().includes(lq)||a.n.toLowerCase().includes(lq)):attrs;
      // blank option
      const ph=document.createElement('div'); ph.className='fp-item fp-item-ph';
      ph.textContent='— Field —';
      ph.addEventListener('mousedown',(ev)=>{ ev.preventDefault(); drop.classList.remove('open'); App.setCF(gId,c.id,'',lkId); });
      list.appendChild(ph);
      if(!filtered.length){ list.innerHTML+='<div class="fp-empty">No fields found</div>'; return; }
      filtered.forEach((a:any)=>{
        const item=document.createElement('div'); item.className='fp-item'+(c.field===a.n?' sel':'');
        const nm=document.createElement('span'); nm.className='fp-item-name'; nm.textContent=a.d;
        const lg=document.createElement('span'); lg.className='fp-item-logical'; lg.textContent=a.n;
        const bd=document.createElement('span'); bd.className='fp-item-badge';
        bd.textContent=TYPE_SHORT_FP[a.t]||a.t;
        bd.style.color=TYPE_COLOR[a.t]||'#9ca3af';
        bd.style.borderColor=(TYPE_COLOR[a.t]||'#9ca3af')+'55';
        item.appendChild(nm); item.appendChild(lg); item.appendChild(bd);
        item.addEventListener('mousedown',(ev)=>{ ev.preventDefault(); drop.classList.remove('open'); App.setCF(gId,c.id,a.n,lkId); });
        list.appendChild(item);
      });
    };

    updateTrigger();

    trigger.addEventListener('click',(e)=>{ e.stopPropagation();
      const wasOpen=drop.classList.contains('open');
      drop.classList.toggle('open',!wasOpen);
      if(!wasOpen){ searchInp.value=''; renderList(''); setTimeout(()=>searchInp.focus(),40); }
    });
    document.addEventListener('click',(e)=>{ if(!e.composedPath().includes(wrap)) drop.classList.remove('open'); },{capture:false,passive:true});
    searchInp.addEventListener('input',()=>renderList(searchInp.value.trim()));

    return wrap;
  }

  function buildCondRow(c:any, gId:string, attrs:any[], lkId:string|null, entName:string): HTMLElement {
    const sa=attrs.find((a:any)=>a.n===c.field); const type=sa?sa.t:'string';
    const ops=opsFor(type);
    const row=document.createElement('div'); row.className='cond-row'; row.id='cr_'+c.id;

    row.appendChild(buildFieldPicker(c, attrs, gId, lkId));

    const oSel=document.createElement('select'); oSel.className='cr-s'; oSel.style.flex='0 0 170px';
    Object.assign(oSel.dataset,{action:'setCO',gid:gId,cid:c.id}); if(lkId) oSel.dataset.lkid=lkId;
    oSel.innerHTML=ops.map(op=>`<option value="${op}"${c.op===op?' selected':''}>${OL[op]||op}</option>`).join('');
    row.appendChild(oSel);

    row.appendChild(buildCondVal(c, type, gId, lkId, entName));

    const rm=document.createElement('button'); rm.className='ib del'; rm.textContent='✕';
    Object.assign(rm.dataset,{action:'rmCond',gid:gId,cid:c.id}); if(lkId) rm.dataset.lkid=lkId;
    row.appendChild(rm);
    return row;
  }

  function buildCondVal(c:any, type:string, gId:string, lkId:string|null, entName:string): HTMLElement {
    if (NO_VAL.has(c.op)) {
      const s=document.createElement('span'); s.className='cr-noval'; s.textContent='—'; return s;
    }
    if (X_VAL.has(c.op)) {
      const inp=document.createElement('input') as HTMLInputElement;
      inp.className='cr-v'; inp.type='number'; inp.min='1'; inp.placeholder='N';
      inp.value=c.val||'';
      Object.assign(inp.dataset,{action:'setCV',gid:gId,cid:c.id}); if(lkId) inp.dataset.lkid=lkId;
      return inp;
    }
    if (type==='boolean') {
      const sel=document.createElement('select') as HTMLSelectElement;
      sel.className='cr-s'; sel.style.flex='1';
      Object.assign(sel.dataset,{action:'setCV',gid:gId,cid:c.id}); if(lkId) sel.dataset.lkid=lkId;
      sel.innerHTML=`<option value="">—</option><option value="1"${c.val==='1'?' selected':''}>True</option><option value="0"${c.val==='0'?' selected':''}>False</option>`;
      return sel;
    }
    if (['integer','bigint','decimal','money','double'].includes(type)) {
      const inp=document.createElement('input') as HTMLInputElement;
      inp.className='cr-v'; inp.type='number'; inp.placeholder='Value';
      inp.value=c.val||'';
      Object.assign(inp.dataset,{action:'setCV',gid:gId,cid:c.id}); if(lkId) inp.dataset.lkid=lkId;
      return inp;
    }
    if (type==='datetime' && (c.op==='on'||c.op==='on-or-after'||c.op==='on-or-before')) {
      const inp=document.createElement('input') as HTMLInputElement;
      inp.className='cr-v'; inp.type='date';
      inp.value=c.val||'';
      Object.assign(inp.dataset,{action:'setCV',gid:gId,cid:c.id}); if(lkId) inp.dataset.lkid=lkId;
      return inp;
    }
    if (['picklist','state','status'].includes(type)) {
      if (MULTI_VAL.has(c.op)) return buildCheckboxPicker(c, gId, lkId, entName);
      // LIKE_WRAP ops (contains/not-contains/starts-with/not-starts-with/ends-with/not-ends-with) → text input
      const inp=document.createElement('input') as HTMLInputElement;
      inp.className='cr-v'; inp.type='text'; inp.placeholder='Value';
      inp.value=c.val||'';
      Object.assign(inp.dataset,{action:'setCV',gid:gId,cid:c.id}); if(lkId) inp.dataset.lkid=lkId;
      inp.addEventListener('input',()=>{ c.val=inp.value; syncXML(); });
      return inp;
    }
    if (type==='multiselect') {
      if (c.op==='contain-values'||c.op==='not-contain-values'||c.op==='eq'||c.op==='ne') return buildCheckboxPicker(c, gId, lkId, entName);
      // ends-with / not-ends-with → text input
      const inp=document.createElement('input') as HTMLInputElement;
      inp.className='cr-v'; inp.type='text'; inp.placeholder='Value';
      inp.value=c.val||'';
      Object.assign(inp.dataset,{action:'setCV',gid:gId,cid:c.id}); if(lkId) inp.dataset.lkid=lkId;
      inp.addEventListener('input',()=>{ c.val=inp.value; syncXML(); });
      return inp;
    }
    if ((['lookup','customer','owner'].includes(type) || isPrimaryIdLookupField(entName, c.field, type)) && (c.op==='eq'||c.op==='ne')) {
      return buildLookupWidget(c, gId, lkId, entName);
    }
    // Default: text input (string, memo, uniqueidentifier, in/not-in multi-text, lookup other ops)
    const inp=document.createElement('input') as HTMLInputElement;
    inp.className='cr-v'; inp.type='text';
    inp.placeholder=MULTI_VAL.has(c.op)?'value1, value2, …':'Value';
    inp.value=c.val||'';
    Object.assign(inp.dataset,{action:'setCV',gid:gId,cid:c.id}); if(lkId) inp.dataset.lkid=lkId;
    return inp;
  }


  function buildCheckboxPicker(c:any, _gId:string, _lkId:string|null, entName:string): HTMLElement {
    const cacheKey=`${entName}|${c.field}`;
    const wrap=document.createElement('div'); wrap.className='cr-cbp';
    const trigger=document.createElement('div'); trigger.className='cr-cbp-trigger';
    const lbl=document.createElement('span'); lbl.className='cr-cbp-trigger-lbl'; lbl.textContent='Select values…';
    const arr=document.createElement('span'); arr.className='cr-cbp-trigger-arr'; arr.textContent='▼';
    trigger.appendChild(lbl); trigger.appendChild(arr); wrap.appendChild(trigger);
    const drop=document.createElement('div'); drop.className='cr-cbp-drop'; wrap.appendChild(drop);
    const closeDrop=()=>drop.classList.remove('open');
    const closeRow=document.createElement('div'); closeRow.className='cr-cbp-close-row';
    const closeBtn=document.createElement('button'); closeBtn.type='button'; closeBtn.className='cr-cbp-close'; closeBtn.title='Close'; closeBtn.textContent='×';
    closeRow.appendChild(closeBtn);

    const updateLabel=()=>{
      const checked=[...drop.querySelectorAll<HTMLInputElement>('input:checked')];
      lbl.textContent=checked.length?checked.map(el=>el.dataset.label||el.value).join(', '):'Select values…';
    };
    const attachHandlers=()=>{
      drop.querySelectorAll<HTMLInputElement>('input[type="checkbox"]').forEach(cb=>{
        cb.addEventListener('change',()=>{
          c.val=[...drop.querySelectorAll<HTMLInputElement>('input:checked')].map(el=>el.value).join(',');
          updateLabel(); syncXML();
        });
      });
    };
    const renderOpts=(opts:{v:number;l:string}[])=>{
      const selVals=c.val?c.val.split(',').map((v:string)=>v.trim()).filter(Boolean):[];
      drop.innerHTML='';
      drop.appendChild(closeRow);
      drop.insertAdjacentHTML('beforeend',opts.map(o=>`<label class="cr-cbp-item"><input type="checkbox" value="${o.v}" data-label="${o.l}"${selVals.includes(String(o.v))?' checked':''}><span>${o.l}</span></label>`).join(''));
      attachHandlers(); updateLabel();
    };

    trigger.addEventListener('click',(e)=>{
      e.stopPropagation();
      const isOpen=drop.classList.contains('open');
      document.querySelectorAll('.cr-cbp-drop.open').forEach(d=>d.classList.remove('open'));
      drop.classList.toggle('open',!isOpen);
    });
    document.addEventListener('click',(e)=>{
      if(!e.composedPath().includes(wrap)) closeDrop();
    },{capture:false,passive:true});
    closeBtn.addEventListener('click',(e)=>{ e.stopPropagation(); closeDrop(); });

    if (DATA.optionSets[cacheKey]) {
      renderOpts(DATA.optionSets[cacheKey]);
    } else {
      drop.innerHTML='<div class="load-spinner">Loading…</div>';
      if (callbacks.fetchAttrOptions) {
        callbacks.fetchAttrOptions(entName, c.field).then(opts=>{
          DATA.optionSets[cacheKey]=opts; renderOpts(opts);
        }).catch(()=>{ drop.innerHTML=''; });
      }
    }
    return wrap;
  }

  function buildLookupWidget(c:any, _gId:string, _lkId:string|null, entName:string): HTMLElement {
    const getTargets=():string[]=>{
      const a=DATA.meta[entName]?.attrs.find((x:any)=>x.n===c.field);
      if ((a as any)?.targets?.length) return (a as any).targets;
      if (DATA.meta[entName]?.primaryId===c.field) return entName ? [entName] : [];
      return [];
    };
    const getTargetEnt=():string=>{ const ts=getTargets(); return activeEntity||ts[0]||''; };

    // Active entity for multi-target lookups (e.g. customerid → account or contact)
    let activeEntity='';

    type LkEntry={id:string;name:string;entity?:string;url?:string};
    const selected:LkEntry[]=c.valLabel
      ? c.val.split(',').map((id:string,i:number)=>{
          const lbl=(c.valLabel.split('||')[i]||id).trim();
          // format: "name|entity" for poly lookups, plain name otherwise
          const sep=lbl.lastIndexOf('|');
          if(sep>0){ return {id:id.trim(),name:lbl.slice(0,sep),entity:lbl.slice(sep+1)}; }
          return {id:id.trim(),name:lbl};
        }).filter((e:LkEntry)=>e.id)
      : [];

    const saveState=()=>{
      c.val=selected.map(e=>e.id).join(',');
      c.valLabel=selected.map(e=>e.entity?`${e.name}|${e.entity}`:e.name).join('||');
      syncXML();
    };

    // Root
    const wrap=document.createElement('div'); wrap.className='lk-srch';

    // Top row: trigger (selected view) + search button
    const row=document.createElement('div'); row.className='lk-ctrl-row'; wrap.appendChild(row);

    const trigger=document.createElement('div'); trigger.className='lk-trigger';
    const trigLbl=document.createElement('span'); trigLbl.className='lk-trigger-lbl';
    const trigBadge=document.createElement('span'); trigBadge.className='lk-trigger-badge'; trigBadge.style.display='none';
    trigger.appendChild(trigLbl); trigger.appendChild(trigBadge);
    row.appendChild(trigger);

    const searchBtn=document.createElement('button'); searchBtn.className='lk-search-btn'; searchBtn.type='button'; searchBtn.title='Search & add'; searchBtn.textContent='⌕';
    row.appendChild(searchBtn);

    // Selected items panel
    const selDrop=document.createElement('div'); selDrop.className='lk-sel-drop'; wrap.appendChild(selDrop);

    // Search panel
    const searchDrop=document.createElement('div'); searchDrop.className='lk-drop'; wrap.appendChild(searchDrop);

    // Entity picker row — only shown when lookup has multiple target entities
    const entRow=document.createElement('div'); entRow.className='lk-drop-field-row'; entRow.style.display='none';
    const entLbl=document.createElement('span'); entLbl.className='lk-drop-field-lbl'; entLbl.textContent='Search in';
    const entSel=document.createElement('select') as HTMLSelectElement; entSel.className='lk-drop-field-sel';
    entRow.appendChild(entLbl); entRow.appendChild(entSel); searchDrop.appendChild(entRow);

    // Field-by picker row — populated lazily on open so meta is guaranteed loaded
    const fieldRow=document.createElement('div'); fieldRow.className='lk-drop-field-row';
    const fieldLbl=document.createElement('span'); fieldLbl.className='lk-drop-field-lbl'; fieldLbl.textContent='Search by';
    const fieldSel=document.createElement('select') as HTMLSelectElement; fieldSel.className='lk-drop-field-sel';
    fieldSel.innerHTML='<option value="">Loading…</option>';
    const populateFieldSel=()=>{
      const te=getTargetEnt();
      const mAttrs:(typeof DATA.meta[string]['attrs'][number])[]=(te&&DATA.meta[te]?.attrs)||[];
      const primaryAttr=mAttrs.find((a:any)=>a.primary===true);
      const pnVal=primaryAttr?.n||'';
      const pnDisplay=primaryAttr?primaryAttr.d:'Name';
      const others=[...mAttrs.filter((a:any)=>(a.t==='string'||a.t==='memo')&&!a.primary)]
        .sort((a:any,b:any)=>a.d.localeCompare(b.d));
      fieldSel.innerHTML=`<option value="${pnVal}">${pnDisplay}</option>`
        +others.map((a:any)=>`<option value="${a.n}">${a.d}</option>`).join('');
    };
    const populateEntSel=(targets:string[])=>{
      const disp=(t:string)=>DATA.entities.find(e=>e.name===t)?.display||t;
      if(!activeEntity) activeEntity=targets[0]||'';
      entSel.innerHTML=targets.map(t=>`<option value="${t}"${t===activeEntity?' selected':''}>${disp(t)}</option>`).join('');
      entRow.style.display=targets.length>1?'':'none';
    };
    fieldRow.appendChild(fieldLbl); fieldRow.appendChild(fieldSel); searchDrop.appendChild(fieldRow);
    // Search input
    const searchBar=document.createElement('div'); searchBar.className='lk-drop-search';
    const inp=document.createElement('input') as HTMLInputElement;
    inp.className='lk-srch-inp'; inp.type='text'; inp.placeholder='Type to search or paste a GUID…';
    const closeBtn=document.createElement('button'); closeBtn.type='button'; closeBtn.className='lk-drop-close'; closeBtn.title='Close search'; closeBtn.textContent='×';
    searchBar.appendChild(inp); searchBar.appendChild(closeBtn); searchDrop.appendChild(searchBar);
    const list=document.createElement('div'); list.className='lk-drop-list'; searchDrop.appendChild(list);

    const closeBoth=()=>{ selDrop.classList.remove('open'); searchDrop.classList.remove('open'); searchBtn.classList.remove('active'); };

    const updateTrigger=()=>{
      const te=getTargetEnt();
      if(!selected.length){
        trigLbl.textContent=te?`Select ${te}…`:'Select…';
        trigLbl.style.opacity='0.45';
        trigBadge.style.display='none';
      } else {
        trigLbl.textContent=selected[0].name;
        trigLbl.style.opacity='1';
        if(selected.length>1){ trigBadge.textContent=`+${selected.length-1}`; trigBadge.style.display=''; }
        else { trigBadge.style.display='none'; }
      }
    };

    const renderSelDrop=()=>{
      selDrop.innerHTML='';
      if(!selected.length){
        const emp=document.createElement('div'); emp.className='lk-srch-empty'; emp.innerHTML='Nothing selected yet.<br>Click <b>⌕</b> to search and add records.'; selDrop.appendChild(emp);
        return;
      }
      selected.forEach((e,i)=>{
        const item=document.createElement('div'); item.className='lk-sel-item';
        if(e.url){
          const pop=document.createElement('a'); pop.className='lk-srch-item-pop'; pop.href=e.url; pop.target='_blank'; pop.rel='noopener'; pop.title='Open record'; pop.textContent='â§‰';
          pop.textContent=POPOUT_ICON;
          pop.addEventListener('mousedown',(ev)=>ev.stopPropagation());
          item.appendChild(pop);
        }
        const body=document.createElement('span'); body.style.cssText='flex:1;min-width:0;display:flex;flex-direction:column;gap:1px';
        const nm=document.createElement('span'); nm.className='lk-sel-item-name'; nm.textContent=e.name; nm.title=e.name;
        body.appendChild(nm);
        if(e.entity){ const et=document.createElement('span'); et.className='lk-srch-item-sub'; et.textContent=DATA.entities.find(x=>x.name===e.entity)?.display||e.entity; body.appendChild(et); }
        const rm=document.createElement('button'); rm.type='button'; rm.className='lk-sel-rm'; rm.textContent='✕'; rm.title='Remove';
        rm.addEventListener('mousedown',(ev)=>{ ev.preventDefault(); selected.splice(i,1); saveState(); updateTrigger(); renderSelDrop(); renderSearchList(lastRecs); });
        item.appendChild(body); item.appendChild(rm); selDrop.appendChild(item);
      });
    };

    const renderSearchList=(recs:{id:string;name:string;sub?:string;url?:string}[])=>{
      list.innerHTML='';
      if(!recs.length){ list.innerHTML='<div class="lk-srch-empty">No records found</div>'; return; }
      recs.forEach(r=>{
        const isSel=!!selected.find(s=>s.id===r.id);
        const item=document.createElement('div'); item.className='lk-srch-item'+(isSel?' sel':'');
        const check=document.createElement('span'); check.className='lk-srch-item-check'; check.textContent=isSel?'✓':'';
        const body=document.createElement('span'); body.className='lk-srch-item-body';
        const nm=document.createElement('span'); nm.className='lk-srch-item-name'; nm.textContent=r.name; nm.title=r.name;
        body.appendChild(nm);
        if(r.sub){ const sub=document.createElement('span'); sub.className='lk-srch-item-sub'; sub.textContent=r.sub; body.appendChild(sub); }
        item.appendChild(check); item.appendChild(body);
        if(r.url){
          const pop=document.createElement('a'); pop.className='lk-srch-item-pop'; pop.href=r.url; pop.target='_blank'; pop.rel='noopener'; pop.title='Open record'; pop.textContent='⧉';
          pop.textContent=POPOUT_ICON;
          pop.addEventListener('mousedown',(ev)=>ev.stopPropagation());
          item.appendChild(pop);
        }
        item.addEventListener('mousedown',(ev)=>{
          ev.preventDefault();
          if(isSel){ const idx=selected.findIndex(s=>s.id===r.id); if(idx>=0) selected.splice(idx,1); }
          else {
            const isMulti=getTargets().length>1;
            const ent=isMulti?getTargetEnt():undefined;
            selected.push({id:r.id,name:r.name,entity:ent,url:r.url});
          }
          saveState(); updateTrigger(); renderSearchList(lastRecs);
        });
        list.appendChild(item);
      });
    };

    let lastRecs:{id:string;name:string;sub?:string;url?:string}[]=[];
    let debTimer:any;
    const search=(q:string)=>{
      clearTimeout(debTimer);
      list.innerHTML='<div class="lk-srch-empty">Searching…</div>';
      debTimer=setTimeout(async()=>{
        const te=getTargetEnt();
        if(!callbacks.fetchLookupRecords||!te){list.innerHTML='<div class="lk-srch-empty">No target entity</div>';return;}
        try {
          const sf=fieldSel.value||null;
          lastRecs=await callbacks.fetchLookupRecords(te,q,sf as string|undefined);
          renderSearchList(lastRecs);
        }
        catch { list.innerHTML='<div class="lk-srch-empty">Error searching</div>'; }
      },250);
    };

    updateTrigger();

    // Trigger click → open selected panel
    trigger.addEventListener('click',(e)=>{
      e.stopPropagation();
      const wasOpen=selDrop.classList.contains('open');
      closeBoth();
      if(!wasOpen){ renderSelDrop(); selDrop.classList.add('open'); }
    });

    // Search button click → open search panel
    searchBtn.addEventListener('click',(e)=>{
      e.stopPropagation();
      const wasOpen=searchDrop.classList.contains('open');
      closeBoth();
      if(!wasOpen){
        searchDrop.classList.add('open'); searchBtn.classList.add('active');
        const targets=getTargets();
        if(!activeEntity) activeEntity=targets[0]||'';
        populateEntSel(targets);
        inp.value=''; search(''); setTimeout(()=>inp.focus(),40);
        const te=getTargetEnt();
        if(te && !DATA.meta[te]?.attrs?.length) ensureEntityMeta(te,()=>populateFieldSel());
        else populateFieldSel();
      }
    });

    document.addEventListener('click',(e)=>{ if(!e.composedPath().includes(wrap)) closeBoth(); },{capture:false,passive:true});
    inp.addEventListener('input',()=>search(inp.value.trim()));
    closeBtn.addEventListener('click',(e)=>{ e.stopPropagation(); closeBoth(); });
    entSel.addEventListener('change',()=>{
      activeEntity=entSel.value;
      fieldSel.innerHTML='<option value="">Loading…</option>';
      inp.value=''; search('');
      const te=getTargetEnt();
      if(te && !DATA.meta[te]?.attrs?.length) ensureEntityMeta(te,()=>populateFieldSel());
      else populateFieldSel();
      setTimeout(()=>inp.focus(),40);
    });
    fieldSel.addEventListener('change',()=>{ inp.value=''; search(''); setTimeout(()=>inp.focus(),40); });

    return wrap;
  }

  function isPrimaryIdLookupField(entName:string, field:string, type:string): boolean {
    return type==='uniqueidentifier' && !!field && DATA.meta[entName]?.primaryId===field;
  }

  function renderLkF(lk:any) {
    const el=$('lkf-'+lk.id); if(!el) return; el.innerHTML='';
    const g=lk.filter||(lk.filter=newG('and'));
    el.appendChild(buildFG(g,true,lk.rel.toEntity,lk.id));
  }

  function renderLinks() {
    const el=$('lk-root'); if(!el) return; el.innerHTML='';
    S.links.forEach((lk:any)=>el.appendChild(buildLkCard(lk,0)));
    ensureColumnPopupCloseButtons();
    const sm=$('sm-links'); if(sm) sm.textContent=S.countLinks(S.links)+' joined';
  }

  function buildLkCard(lk:any, depth:number): HTMLElement {
    const display=DATA.entities.find(e=>e.name===lk.rel.toEntity)?.display||lk.rel.toEntity;
    const div=document.createElement('div');
    div.className=`lk-card lk-depth-${Math.min(depth,4)} open`; div.id='lk_'+lk.id;

    const hdr=document.createElement('div'); hdr.className='lk-hdr'; hdr.dataset.action='togLkOpen'; hdr.dataset.id=lk.id;
    hdr.innerHTML=`${depth>0?`<span class="lk-path">L${depth}</span>`:''}<span class="lk-badge">${lk.joinType}</span><span class="lk-name">${display}</span><span class="lk-alias-disp">${lk.alias}</span>`;
    const rmBtn=document.createElement('button'); rmBtn.className='ib del'; rmBtn.textContent='✕';
    rmBtn.dataset.action='rmLink'; rmBtn.dataset.id=lk.id;
    rmBtn.addEventListener('click',e=>{e.stopPropagation();App.rmLink(lk.id);});
    hdr.appendChild(rmBtn); div.appendChild(hdr);

    const body=document.createElement('div'); body.className='lk-body';

    const joinRow=document.createElement('div'); joinRow.className='lk-row';
    joinRow.innerHTML=`<label>Join</label>
      <div class="jp-grp">
        <button class="jp${lk.joinType==='inner'?' on':''}" data-action="setLkJ" data-id="${lk.id}" data-jt="inner">Inner</button>
        <button class="jp${lk.joinType==='outer'?' on':''}" data-action="setLkJ" data-id="${lk.id}" data-jt="outer">Outer</button>
      </div>
      <label style="margin-left:8px">Alias</label>
      <input class="alias-inp" value="${lk.alias}" data-action="setLkA" data-id="${lk.id}">`;
    body.appendChild(joinRow);

    if(lk.rel.fromAttr){const info=document.createElement('div');info.style.cssText='font-size:10px;color:#333;margin-bottom:6px;font-family:Consolas,monospace';info.textContent=`${lk.rel.fromAttr} → ${lk.rel.toEntity}.${lk.rel.toAttr}`;body.appendChild(info);}

    const clbl=document.createElement('div'); clbl.className='lk-sec-lbl'; clbl.textContent='Columns'; body.appendChild(clbl);
    const cw=document.createElement('div'); cw.style.position='relative';
    cw.innerHTML=`<div class="col-chips" id="chips-lk-${lk.id}" data-action="togLkColsBtn" data-lkid="${lk.id}"><span class="chip-ph">Click to select columns…</span></div>
      <div class="dd-popup dd-below" id="dd-lkc-${lk.id}">
        <div class="dd-search">${SVG_SEARCH}<input type="text" placeholder="Search columns…" data-input-action="rl_cols" data-lkid="${lk.id}"></div>
        <div class="dd-toolbar">
          <button class="dd-tbar-btn" data-action="selAll" data-lkid="${lk.id}">All</button>
          <button class="dd-tbar-btn" data-action="clrCols" data-lkid="${lk.id}">Clear</button>
          <span class="dd-tbar-count" id="cnt-lkc-${lk.id}">0</span>
        </div>
        <div class="dd-list" id="list-lkc-${lk.id}" style="max-height:200px"></div>
      </div>`;
    body.appendChild(cw);

    const flbl=document.createElement('div'); flbl.className='lk-sec-lbl'; flbl.textContent='Filters'; body.appendChild(flbl);
    const fel=document.createElement('div'); fel.id='lkf-'+lk.id; body.appendChild(fel);

    const nlbl=document.createElement('div'); nlbl.className='lk-sec-lbl'; nlbl.style.marginTop='10px'; nlbl.textContent='Nested Related Tables'; body.appendChild(nlbl);
    const nl=document.createElement('div'); nl.className='nested-lk-list'; nl.id='nested-lk-'+lk.id;
    (lk.links||[]).forEach((ch:any)=>nl.appendChild(buildLkCard(ch,depth+1)));
    body.appendChild(nl);

    const aw=document.createElement('div'); aw.style.position='relative';
    aw.innerHTML=`<button class="add-btn sm" data-action="togRelDD" data-lkid="${lk.id}">＋ Join from ${display}</button>
      <div class="dd-popup dd-top" id="dd-rel-${lk.id}">
        <div class="dd-search">${SVG_SEARCH}<input type="text" placeholder="Search relationships…" data-input-action="rl_rel" data-lkid="${lk.id}"></div>
        <div class="dd-list" id="list-rel-${lk.id}" style="max-height:160px"></div>
      </div>`;
    body.appendChild(aw);
    div.appendChild(body);

    setTimeout(()=>{renderChips(lk.id);renderLkF(lk);},0);
    return div;
  }

  function renderSorts() {
    const el=$('sort-list'); if(!el) return;
    el.innerHTML=S.sorts.map((s:any,i:number)=>`<div class="sort-row">
      ${s.alias?`<span class="sort-alias-lbl">[${s.alias}]</span>`:''}
      ${s.isAggAlias
        ? `<span class="sort-attr" style="color:#d4893a">${s.attr}</span><span class="sort-alias-lbl" style="margin-left:3px">alias</span>`
        : `<span class="sort-attr">${s.attr}</span>`}
      <div class="sd-grp">
        <button class="sd${!s.desc?' on':''}" data-action="setSortDir" data-idx="${i}" data-desc="false">↑ ASC</button>
        <button class="sd${s.desc?' on d':''}" data-action="setSortDir" data-idx="${i}" data-desc="true">↓ DESC</button>
      </div>
      <button class="ib del" data-action="rmSort" data-idx="${i}">✕</button>
    </div>`).join('')||'<div class="hint">No sort orders defined</div>';
    const sm=$('sm-sort'); if(sm) sm.textContent=S.sorts.length+' sort'+(S.sorts.length!==1?'s':'');
  }

  function syncXML() {
    const xml=Gen.run()||'';
    const ta=$('xml-editor') as HTMLTextAreaElement;
    if(ta&&ta.value!==xml) ta.value=xml;
    hlXML();
    syncXmlButtons();
    const vm=$('val-msg');
    if(!S.entity){if(vm)vm.style.display='none';callbacks.onXmlChange('');return;}
    callbacks.onXmlChange(xml);
    if(!vm) return;
    const warns:string[]=[];
    if (S.opts.aggregate) {
      const missingAlias = S.fields.filter((f:any) => !f.alias).length;
      if (missingAlias) warns.push(`⚠ ${missingAlias} aggregate field${missingAlias>1?'s':''} missing alias`);
      const hasAggFn = S.fields.some((f:any) => f.aggr && f.aggr !== 'groupby');
      const hasGroupBy = S.fields.some((f:any) => f.aggr === 'groupby');
      if (hasAggFn && !hasGroupBy) warns.push('⚠ No Group By field — query may return unexpected results');
      if (S.opts.distinct) warns.push('⚠ Distinct and Aggregate are mutually exclusive');
    } else {
      if(!S.fields.filter((f:any)=>!f.alias).length) warns.push('⚠ No columns selected — using all-attributes');
    }
    const bc=badC(S.rootF); if(bc) warns.push(`✕ ${bc} filter condition${bc>1?'s':''} missing value`);
    vm.style.display='flex';
    if(warns.length){vm.className='warn';vm.textContent=warns.join('  ·  ');}
    else{vm.className='ok';vm.textContent='✓ Valid FetchXML';}
  }

  function badC(g:any):number{let c=(g.conds||[]).filter((c:any)=>c.field&&c.op&&!NO_VAL.has(c.op)&&c.val==='').length;(g.kids||[]).forEach((k:any)=>c+=badC(k));return c;}

  function onAggModeChange() {
    const chipWrap = $('cols-chip-wrap');
    const aggWrap  = $('agg-rows-wrap');
    const distRow  = $('opt-row-distinct');
    if (S.opts.aggregate) {
      if (chipWrap) (chipWrap as HTMLElement).style.display = 'none';
      if (aggWrap)  (aggWrap  as HTMLElement).style.display = '';
      if (distRow)  (distRow  as HTMLElement).style.opacity = '0.4';
      if (distRow)  (distRow  as HTMLElement).style.pointerEvents = 'none';
      // Convert existing plain fields to groupby rows
      if (S.fields.length && !S.fields.some((f:any) => f.aggr)) {
        S.fields = S.fields
          .filter((f:any) => !f.alias) // keep only primary fields (not link-entity alias fields)
          .map((f:any) => ({ attr: f.attr, alias: autoAlias('groupby', f.attr), aggr: 'groupby' }));
      }
      // Discard non-aggregate sorts
      S.sorts = S.sorts.filter((s:any) => s.isAggAlias);
      renderAggRows();
      renderSorts();
    } else {
      if (chipWrap) (chipWrap as HTMLElement).style.display = '';
      if (aggWrap)  (aggWrap  as HTMLElement).style.display = 'none';
      if (distRow)  (distRow  as HTMLElement).style.opacity = '';
      if (distRow)  (distRow  as HTMLElement).style.pointerEvents = '';
      // Keep groupby fields as plain columns; discard agg-function rows
      S.fields = S.fields
        .filter((f:any) => f.aggr === 'groupby' || !f.aggr)
        .map((f:any) => ({ attr: f.attr, alias: null }));
      // Discard alias-based sorts
      S.sorts = S.sorts.filter((s:any) => !s.isAggAlias);
      renderChips(null);
      renderSorts();
    }
  }

  // ── Lazy metadata loader ───────────────────────────────────────────
  // Fetches and caches an entity's metadata via the callback, then
  // calls onLoaded() so callers can refresh affected UI elements.
  function ensureEntityMeta(entityName: string, onLoaded: () => void): void {
    if (DATA.meta[entityName]?.attrs?.length) { onLoaded(); return; }
    callbacks.fetchEntityMeta(entityName).then((raw: any) => {
      const safeAttrs: any[] = Array.isArray(raw?.attrs) ? raw.attrs : [];
      const safeRels:  any[] = Array.isArray(raw?.rels)  ? raw.rels  : [];
      DATA.meta[entityName] = {
        attrs: safeAttrs.map((a: any) => ({...a, t: normType(a.t)})),
        rels:  safeRels,
        views: Array.isArray(raw?.views) ? raw.views : [],
        primaryName: raw?.primaryName || '',
        primaryId: raw?.primaryId || '',
        objectTypeCode: raw?.objectTypeCode || 0,
      };
      onLoaded();
    }).catch(() => {
      if (!DATA.meta[entityName]) DATA.meta[entityName] = { attrs: [], rels: [], views: [] };
      onLoaded();
    });
  }

  function unlockSecs() {
    ['sec-cols','sec-links','sec-filt','sec-sort','sec-opts'].forEach(id=>shadow!.getElementById(id)?.querySelector('.section-header')?.classList.remove('locked'));
    const vt=$('trig-view') as HTMLButtonElement;
    if(vt){vt.removeAttribute('disabled');vt.style.pointerEvents='';vt.style.opacity='';}
    shadow!.getElementById('sec-cols')?.querySelector('.section-header')?.classList.add('open');
  }

  function switchTabInner(t: string) {
    shadow!.querySelectorAll('.tab-item[data-action="switchTab"]').forEach(el=>el.classList.remove('active'));
    shadow!.querySelector(`.tab-item[data-tab="${t}"]`)?.classList.add('active');
    shadow!.querySelectorAll('.tab-content').forEach(p=>p.classList.remove('active'));
    shadow!.getElementById('pane-'+t)?.classList.add('active');
    if(t==='xml') App.sync();
  }

  function secTog(hdr: HTMLElement) {
    if(hdr.classList.contains('locked')) return;
    hdr.classList.toggle('open');
  }

  // Returns a string describing the input widget a given type+op combination needs.
  // Value is only cleared in setCO when the widget kind changes.
  function widgetKind(type: string, op: string): string {
    if (NO_VAL.has(op)) return 'none';
    if (X_VAL.has(op)) return 'xval';
    const OPTSET = new Set(['picklist','state','status','multiselect']);
    if (OPTSET.has(type) && (MULTI_VAL.has(op) || op==='eq' || op==='ne')) return 'checkbox';
    if (type==='boolean') return 'boolean';
    if (type==='datetime') return op==='on'||op==='on-or-after'||op==='on-or-before'?'date':'none';
    if (['integer','bigint','decimal','money','double'].includes(type)) return 'number';
    return 'text';
  }

  // ── App actions ───────────────────────────────────────────────────
  const App = {
    async selEnt(name: string) {
      S.entity=name; S.fields=[]; S.links=[]; S.rootF=newG('and'); S.sorts=[];
      closeAll();

      // Re-lock all dependent sections immediately so the UI doesn't show
      // stale content from a previously selected entity during the async fetch.
      ['sec-cols','sec-links','sec-filt','sec-sort','sec-opts'].forEach(id => {
        const hdr = shadow!.getElementById(id)?.querySelector<HTMLElement>('.section-header');
        if (hdr) { hdr.classList.add('locked'); hdr.classList.remove('open'); }
      });
      const vtLock = $('trig-view') as HTMLButtonElement;
      if (vtLock) { vtLock.setAttribute('disabled',''); vtLock.style.pointerEvents='none'; vtLock.style.opacity='.35'; }
      const cpReset = $('chips-primary'); if (cpReset) cpReset.innerHTML='<span class="chip-ph">Click to select columns…</span>';
      const frReset = $('filter-root'); if (frReset) frReset.innerHTML='';
      const lrReset = $('lk-root'); if (lrReset) lrReset.innerHTML='';
      const slReset = $('sort-list'); if (slReset) slReset.innerHTML='<div class="hint">No sort orders defined</div>';

      const ent=DATA.entities.find(e=>e.name===name);
      const lbl=$('lbl-entity'); if(lbl){lbl.textContent=ent?.display||name;lbl.classList.remove('ph');}
      const sm=$('sm-table'); if(sm) sm.textContent=ent?.display||name;
      const sub=$('sm-table-sub'); if(sub) sub.textContent=name;
      const vlbl=$('lbl-view'); if(vlbl){vlbl.textContent='Choose a view…';vlbl.classList.add('ph');}
      const smv=$('sm-view'); if(smv) smv.textContent='—';
      const smc=$('sm-cols'); if(smc) smc.textContent='0 selected';

      // Always fetch fresh on explicit entity selection — views change when customizations
      // are deployed, and the page-world has no session cache to return stale data from.
      const lc=$('list-cols'); if(lc) lc.innerHTML='<div class="load-spinner">Loading metadata…</div>';
      try {
        const raw = await callbacks.fetchEntityMeta(name);
        const safeAttrs: any[] = Array.isArray(raw?.attrs) ? raw.attrs : [];
        const safeRels:  any[] = Array.isArray(raw?.rels)  ? raw.rels  : [];
        DATA.meta[name] = {
          attrs: safeAttrs.map((a: any) => ({...a, t: normType(a.t)})),
          rels:  safeRels,
          views: Array.isArray(raw?.views) ? raw.views : [],
          primaryName: raw?.primaryName || '',
          primaryId: raw?.primaryId || '',
        };
      } catch(e) {
        // Preserve any previously loaded data; ensure entry exists so rlView/rlCols see an empty array
        if (!DATA.meta[name]) DATA.meta[name] = { attrs: [], rels: [], views: [] };
      }
      unlockSecs();
      renderChips(null); renderLinks();
      const fr=$('filter-root'); if(fr) fr.innerHTML='';
      const sl=$('sort-list'); if(sl) sl.innerHTML='<div class="hint">No sort orders defined</div>';
      renderFilters(); syncXML();
    },

    loadView(viewId: string) {
      const v=DATA.meta[S.entity!]?.views?.find(v=>v.id===viewId); if(!v) return;
      const r:any=Parser.parse(v.fx); if(r.err) return;
      // Tag the view name onto the parsed result so _apply can update the label
      // regardless of whether it's called directly or via selEnt().then().
      r.__viewName = v.name;
      if(S.entity!==r.eName) { this.selEnt(r.eName).then(()=>this._apply(r)); return; }
      this._apply(r);
      closeAll();
    },

    _apply(r:any) {
      S.links=r.links; S.rootF=r.rootF; S.sorts=r.orders; Object.assign(S.opts,r.opts);
      const oc=$('opt-count') as HTMLInputElement; if(oc) oc.value=S.opts.count||'';
      const op=$('opt-page') as HTMLInputElement; if(op) op.value=S.opts.page||'';
      const td=$('tog-dist'); S.opts.distinct?td?.classList.add('on'):td?.classList.remove('on');
      const ta=$('tog-agg'); S.opts.aggregate?ta?.classList.add('on'):ta?.classList.remove('on');
      if (S.opts.aggregate) {
        S.fields = r.aggFields || [];
        onAggModeChange();
      } else {
        S.fields=r.attrs.map((an:string)=>({attr:an,alias:null}));
        onAggModeChange(); // ensures chip-wrap shown, agg-wrap hidden
        renderChips(null);
      }
      renderLinks(); renderFilters(); renderSorts(); syncXML();
      // Update view label if this came from loadView
      if(r.__viewName) {
        const lbl=$('lbl-view'); if(lbl){lbl.textContent=r.__viewName;lbl.classList.remove('ph');}
        const smv=$('sm-view'); if(smv) smv.textContent=r.__viewName;
      }
      // Fetch metadata for every linked entity so their column pickers and
      // filter fields are populated (handles views with link-entity joins).
      const loadLinksMeta = (links: any[]) => {
        links.forEach((lk: any) => {
          if (lk.rel?.toEntity) {
            const lkId = lk.id;
            ensureEntityMeta(lk.rel.toEntity, () => {
              renderChips(lkId);
              const lkObj = S.findLink(S.links, lkId);
              if (lkObj) renderLkF(lkObj);
            });
          }
          if (lk.links?.length) loadLinksMeta(lk.links);
        });
      };
      if (S.links.length) loadLinksMeta(S.links);
    },

    togCol(attr:string, lkId:string|null) {
      if(!lkId){const idx=S.fields.findIndex((f:any)=>f.attr===attr&&!f.alias);if(idx>=0)S.fields.splice(idx,1);else S.fields.push({attr,alias:null});}
      else{const lk=S.findLink(S.links,lkId);if(!lk)return;if(!lk.fields)lk.fields=[];const idx=lk.fields.indexOf(attr);if(idx>=0)lk.fields.splice(idx,1);else lk.fields.push(attr);}
      const q=shadow!.querySelector(lkId?`#dd-lkc-${lkId} input`:'#dd-cols input') as HTMLInputElement;
      rlCols(q?.value||'',lkId); renderChips(lkId); this.sync();
    },

    selAll(lkId:string|null) {
      if (!lkId && S.opts.aggregate) return; // no-op in aggregate mode
      const entName=lkId?S.findLink(S.links,lkId)?.rel.toEntity:S.entity;
      const attrs=entName?DATA.meta[entName]?.attrs||[]:[];
      if(!lkId) attrs.forEach((a:any)=>{if(!S.fields.find((f:any)=>f.attr===a.n&&!f.alias))S.fields.push({attr:a.n,alias:null});});
      else{const lk=S.findLink(S.links,lkId);if(!lk)return;if(!lk.fields)lk.fields=[];attrs.forEach((a:any)=>{if(!lk.fields.includes(a.n))lk.fields.push(a.n);});}
      rlCols('',lkId); renderChips(lkId); this.sync();
    },

    clrCols(lkId:string|null) {
      if (!lkId && S.opts.aggregate) return; // no-op in aggregate mode
      if(!lkId) S.fields=S.fields.filter((f:any)=>f.alias);
      else{const lk=S.findLink(S.links,lkId);if(lk)lk.fields=[];}
      rlCols('',lkId); renderChips(lkId); this.sync();
    },

    addLink(relName:string, parentEntName:string, parentLkId:string|null) {
      const rel=DATA.meta[parentEntName]?.rels.find((r:any)=>r.name===relName); if(!rel) return;
      const alias=rel.toEntity.substring(0,4)+(S.countLinks(S.links)+1);
      const lk=newL(rel,alias,'inner');
      if(!parentLkId) S.links.push(lk);
      else{const p=S.findLink(S.links,parentLkId);if(!p)return;if(!p.links)p.links=[];p.links.push(lk);}
      closeAll(); renderLinks(); this.sync();
      shadow!.getElementById('sec-links')?.querySelector('.section-header')?.classList.add('open');
      // Load the related entity's metadata so its column picker and filter
      // fields are populated.  Re-render the card's chips + filter after load.
      const lkId = lk.id;
      ensureEntityMeta(rel.toEntity, () => {
        renderChips(lkId);
        const lkObj = S.findLink(S.links, lkId);
        if (lkObj) renderLkF(lkObj);
      });
    },

    rmLink(id:string){S.removeLink(S.links,id);renderLinks();this.sync();},

    setLkJ(id:string,jt:string){const lk=S.findLink(S.links,id);if(lk){lk.joinType=jt;renderLinks();this.sync();}},
    setLkA(id:string,a:string){const lk=S.findLink(S.links,id);if(!lk||!a.trim())return;const old=lk.alias;lk.alias=a.trim();S.sorts.forEach((s:any)=>{if(s.alias===old)s.alias=lk.alias;});this.sync();},

    fgRoot(lkId:string|null):any{if(!lkId)return S.rootF;const lk=S.findLink(S.links,lkId);if(!lk)return null;if(!lk.filter)lk.filter=newG('and');return lk.filter;},
    fg(root:any,id:string):any{if(!root)return null;if(root.id===id)return root;for(const k of(root.kids||[])){const f=this.fg(k,id);if(f)return f;}return null;},
    reRenderFilt(lkId:string|null){if(!lkId)renderFilters();else{const lk=S.findLink(S.links,lkId);if(lk)renderLkF(lk);}},

    addCond(gId:string,lkId:string|null){const g=this.fg(this.fgRoot(lkId),gId);if(g){g.conds.push(newC());this.reRenderFilt(lkId);this.sync();}},
    addGrp(gId:string,lkId:string|null){const g=this.fg(this.fgRoot(lkId),gId);if(g){if(!g.kids)g.kids=[];g.kids.push(newG('and'));this.reRenderFilt(lkId);this.sync();}},
    rmGrp(gId:string,lkId:string|null){
      const rmG=(p:any,id:string):boolean=>{if(!p.kids)return false;const i=p.kids.findIndex((k:any)=>k.id===id);if(i>=0){p.kids.splice(i,1);return true;}return p.kids.some((k:any)=>rmG(k,id));};
      rmG(this.fgRoot(lkId),gId);this.reRenderFilt(lkId);this.sync();
    },
    setL(gId:string,l:string,lkId:string|null){const g=this.fg(this.fgRoot(lkId),gId);if(g){g.logic=l;this.reRenderFilt(lkId);this.sync();}},
    setCF(gId:string,cId:string,f:string,lkId:string|null){const g=this.fg(this.fgRoot(lkId),gId);if(!g)return;const c=g.conds.find((c:any)=>c.id===cId);if(c){c.field=f;c.val='';c.valLabel='';const entName=lkId?S.findLink(S.links,lkId)?.rel.toEntity:S.entity;const attr=entName?DATA.meta[entName]?.attrs.find((a:any)=>a.n===f):null;const t=attr?.t||'string';c.type=t;const ops=opsFor(t);c.op=ops[0];this.reRenderFilt(lkId);this.sync();}},
    setCO(gId:string,cId:string,o:string,lkId:string|null){const g=this.fg(this.fgRoot(lkId),gId);if(!g)return;const c=g.conds.find((c:any)=>c.id===cId);if(c){const entName=lkId?S.findLink(S.links,lkId)?.rel.toEntity||'':S.entity||'';const rawType=c.type||c.t||'string';const lookupLike=(['lookup','customer','owner'].includes(rawType)||isPrimaryIdLookupField(entName,c.field,rawType));const prevKind=(lookupLike&&(c.op==='eq'||c.op==='ne'))?'lookup':widgetKind(rawType,c.op);const nextKind=(lookupLike&&(o==='eq'||o==='ne'))?'lookup':widgetKind(rawType,o);if(prevKind!==nextKind){c.val='';c.valLabel='';}c.op=o;this.reRenderFilt(lkId);this.sync();}},
    setCV(gId:string,cId:string,v:string,lkId:string|null){const g=this.fg(this.fgRoot(lkId),gId);if(!g)return;const c=g.conds.find((c:any)=>c.id===cId);if(c){c.val=v;this.sync();}},
    rmCond(gId:string,cId:string,lkId:string|null){const g=this.fg(this.fgRoot(lkId),gId);if(g){g.conds=g.conds.filter((c:any)=>c.id!==cId);this.reRenderFilt(lkId);this.sync();}},

    // ── Aggregate App methods ────────────────────────────────────────
    addAggRow() {
      S.fields.push({ attr: '', alias: '', aggr: 'groupby' });
      renderAggRows(); this.sync();
    },
    rmAggRow(idx: number) {
      S.fields.splice(idx, 1);
      renderAggRows(); this.sync();
    },
    setAggField(idx: number, attr: string) {
      const f = S.fields[idx]; if (!f) return;
      f.attr = attr;
      // Auto-fill alias if blank or still auto-generated
      const expected = f.alias ? autoAlias(f.aggr || 'groupby', f.attr) : '';
      if (!f.alias || f.alias === expected || f.alias === autoAlias(f.aggr || 'groupby', '')) {
        f.alias = attr ? autoAlias(f.aggr || 'groupby', attr) : '';
      }
      renderAggRows(); this.sync();
    },
    setAggFn(idx: number, aggr: string) {
      const f = S.fields[idx]; if (!f) return;
      const wasAuto = f.attr ? f.alias === autoAlias(f.aggr || 'groupby', f.attr) : false;
      f.aggr = aggr;
      if (wasAuto || !f.alias) f.alias = f.attr ? autoAlias(aggr, f.attr) : '';
      renderAggRows(); this.sync();
    },
    setAggAlias(idx: number, alias: string) {
      const f = S.fields[idx]; if (!f) return;
      f.alias = alias;
      this.sync(); // no re-render to avoid losing focus
    },
    addAggSort(alias: string) {
      if (S.sorts.find((s:any) => s.isAggAlias && s.attr === alias)) return;
      S.sorts.push({ attr: alias, alias: null, desc: false, isAggAlias: true });
      closeAll(); renderSorts(); this.sync();
    },

    addSort(attr:string,alias:string|null){const a=alias||null;if(S.sorts.find((s:any)=>s.attr===attr&&s.alias===a))return;S.sorts.push({attr,alias:a,desc:false});closeAll();renderSorts();this.sync();},
    setSortDir(i:number,desc:boolean){S.sorts[i].desc=desc;renderSorts();this.sync();},
    rmSort(i:number){S.sorts.splice(i,1);renderSorts();this.sync();},

    togOpt(key:string,el:HTMLElement){el.classList.toggle('on');(S.opts as any)[key]=el.classList.contains('on');this.sync();if(key==='aggregate')onAggModeChange();},

    sync(){
      const oc=$('opt-count') as HTMLInputElement; S.opts.count=oc?.value||'';
      const op=$('opt-page') as HTMLInputElement; S.opts.page=op?.value||'';
      syncXML();
    },

    copyXML(){const xml=Gen.run()||($('xml-editor') as HTMLTextAreaElement)?.value||'';navigator.clipboard.writeText(xml).catch(()=>{});},

    loadFromXML(){
      const xml=($('xml-editor') as HTMLTextAreaElement)?.value.trim(); if(!xml) return;
      const r:any=Parser.parse(xml); if(r.err) return;
      if(S.entity!==r.eName||!DATA.meta[r.eName]) this.selEnt(r.eName).then(()=>this._apply(r));
      else{this._apply(r);switchTabInner('design');}
    },

    reset(){
      S.init();
      const lbl=$('lbl-entity'); if(lbl){lbl.textContent='Choose a table…';lbl.classList.add('ph');}
      const vlbl=$('lbl-view'); if(vlbl){vlbl.textContent='Choose a view…';vlbl.classList.add('ph');}
      const sub=$('sm-table-sub'); if(sub) sub.textContent='';
      const vt=$('trig-view') as HTMLButtonElement; if(vt){vt.setAttribute('disabled','');vt.style.pointerEvents='none';vt.style.opacity='.35';}
      ['sm-table','sm-view','sm-links','sm-filt','sm-sort'].forEach(id=>{const el=$(id);if(el)el.textContent='—';});
      const smc=$('sm-cols'); if(smc) smc.textContent='0 selected';
      const cp=$('chips-primary'); if(cp) cp.innerHTML='<span class="chip-ph">Click to select columns…</span>';
      const lr=$('lk-root'); if(lr) lr.innerHTML='';
      const fr=$('filter-root'); if(fr) fr.innerHTML='';
      const sl=$('sort-list'); if(sl) sl.innerHTML='<div class="hint">No sort orders defined</div>';
      const xe=$('xml-editor') as HTMLTextAreaElement; if(xe) xe.value='';
      ['opt-count','opt-page'].forEach(id=>{const el=$(id) as HTMLInputElement;if(el) el.value='';});
      $('tog-dist')?.classList.remove('on');
      $('tog-agg')?.classList.remove('on');
      onAggModeChange(); // resets chip-wrap/agg-rows-wrap visibility
      const vm=$('val-msg'); if(vm) vm.style.display='none';
      ['sec-cols','sec-links','sec-filt','sec-sort','sec-opts'].forEach(id=>{const h=shadow!.getElementById(id)?.querySelector('.section-header');if(h){h.classList.add('locked');h.classList.remove('open');}});
      shadow!.getElementById('sec-table')?.querySelector('.section-header')?.classList.add('open');
      closeAll(); callbacks.onXmlChange('');
    },
  };

  // ── Event delegation ──────────────────────────────────────────────
  (shadow as EventTarget).addEventListener('click', (e: Event) => {
    const t = e.target as HTMLElement;
    if (!t.closest('.dd-popup') && !t.closest('.dd-trigger') && !t.closest('.col-chips')) closeAll();

    const el = t.closest('[data-action]') as HTMLElement|null; if (!el) return;
    const action = el.dataset.action!;
    const d = el.dataset;
    const lkId = d.lkid !== undefined ? (d.lkid||null) : null;

    switch(action) {
      case 'switchTab':     switchTabInner(d.tab!); break;
      case 'secTog': { const h=el.classList.contains('section-header')?el:el.closest<HTMLElement>('.section-header'); if(h) secTog(h); break; }
      case 'reset':         App.reset(); break;
      case 'tog':           e.stopPropagation(); togDD(d.key!); break;
      case 'togRelDD':      e.stopPropagation(); togRelDD(lkId); break;
      case 'togLkColsBtn':  e.stopPropagation(); togLkCols(d.lkid!); break;
      case 'closeDD':       e.stopPropagation(); shadow!.getElementById(d.ddid!)?.classList.remove('open'); break;
      case 'togLkOpen':     shadow!.getElementById('lk_'+d.id)?.classList.toggle('open'); break;
      case 'selEnt':        App.selEnt(d.name!); break;
      case 'loadView':      App.loadView(d.vid!); break;
      case 'selAll':        App.selAll(lkId); break;
      case 'clrCols':       App.clrCols(lkId); break;
      case 'togCol':        App.togCol(d.attr!, lkId); break;
      case 'addCondRoot':   App.addCond(S.rootF.id, null); break;
      case 'addGrpRoot':    App.addGrp(S.rootF.id, null); break;
      case 'addCond':       App.addCond(d.gid!, lkId); break;
      case 'addGrp':        App.addGrp(d.gid!, lkId); break;
      case 'rmGrp':         App.rmGrp(d.gid!, lkId); break;
      case 'setL':          App.setL(d.gid!, d.logic!, lkId); break;
      case 'rmCond':        App.rmCond(d.gid!, d.cid!, lkId); break;
      case 'addLink':       App.addLink(d.relname!, d.parentEnt!, d.parentLkid||null); break;
      case 'setLkJ':        App.setLkJ(d.id!, d.jt!); break;
      case 'addSort':       App.addSort(d.attr!, d.alias||null); break;
      case 'setSortDir':    App.setSortDir(Number(d.idx), d.desc==='true'); break;
      case 'rmSort':        App.rmSort(Number(d.idx)); break;
      case 'addAggRow':     App.addAggRow(); break;
      case 'rmAggRow':      App.rmAggRow(Number(d.idx)); break;
      case 'addAggSort':    App.addAggSort(d.alias!); break;
      case 'togOpt':        App.togOpt(d.key!, el); break;
      case 'copyXML':       App.copyXML(); break;
      case 'loadFromXML':   App.loadFromXML(); break;
    }
  });

  (shadow as EventTarget).addEventListener('change', (e: Event) => {
    const t = e.target as HTMLElement; const d = (t as any).dataset; if (!d) return;
    const val = (t as HTMLInputElement|HTMLSelectElement).value;
    const lkId = d.lkid||null;
    switch(d.action) {
      case 'setCF': App.setCF(d.gid!, d.cid!, val, lkId); break;
      case 'setCO': App.setCO(d.gid!, d.cid!, val, lkId); break;
      case 'setCV': App.setCV(d.gid!, d.cid!, val, lkId); break;
      case 'setLkA': App.setLkA(d.id!, val); break;
      case 'setAggField': App.setAggField(Number(d.idx), val); break;
      case 'setAggFn':    App.setAggFn(Number(d.idx), val); break;
    }
  });

  (shadow as EventTarget).addEventListener('input', (e: Event) => {
    const t = e.target as HTMLElement; const d = (t as any).dataset; if (!d) return;
    const val = (t as HTMLInputElement).value;
    const lkId = d.lkid||null;
    const ia = d.inputAction;
    switch(ia) {
      case 'rl_ent':  rlEnt(val); break;
      case 'rl_view': rlView(val); break;
      case 'rl_cols': rlCols(val, lkId); break;
      case 'rl_sf':   rlSf(val); break;
      case 'rl_rel':  rlRel(lkId, val); break;
      case 'optInput': App.sync(); break;
      case 'xmlEdit': hlXML(); syncXmlButtons(); callbacks.onXmlChange(val); break;
    }
    // inline data-action="setCV" on input elements fires via input event
    if (!ia && d.action==='setCV') App.setCV(d.gid!, d.cid!, val, lkId);
    if (!ia && d.action==='setAggAlias') App.setAggAlias(Number(d.idx), val);
  });

  // ── Init: load entities ────────────────────────────────────────────
  DATA.loading = true;
  callbacks.fetchAllEntities().then(ents => {
    const isSpecial = (s: string) => !/^[a-zA-Z]/i.test(s);
    DATA.entities = ents.sort((a, b) => {
      const as = isSpecial(a.display), bs = isSpecial(b.display);
      if (as !== bs) return as ? 1 : -1;
      return a.display.localeCompare(b.display);
    });
    DATA.loaded = true; DATA.loading = false;
    // If entity list dropdown is open, refresh it
    if (_openDD === 'entity') rlEnt('');
  }).catch(() => { DATA.loading = false; });
}
