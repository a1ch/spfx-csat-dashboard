// Dashboard shell. All CSS scoped under .astRoot so nothing leaks onto the
// host SharePoint page. The core fills the filters, KPIs and the four views.
export const ASSET_DASHBOARD_HTML: string = `
<style>
  .astRoot, .astRoot *, .astRoot *::before, .astRoot *::after { box-sizing:border-box; }
  .astRoot * { margin:0; padding:0; }
  .astRoot {
    --navy:#1F3864; --blue:#2E5496; --accent:#2E5496; --accent-lt:#D9E1F2;
    --bg:#f4f6f9; --surface:#fff; --surface2:#f2f2f2; --border:rgba(0,0,0,0.10);
    --border-md:rgba(0,0,0,0.16); --text:#1F3864; --text-2:#5a6478; --text-3:#8a93a6;
    --green:#0F6E56; --red:#A32D2D; --radius:8px; --radius-lg:12px;
    font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif; font-size:14px;
    color:var(--text); background:var(--bg); display:flex; flex-direction:column;
    border:1px solid var(--border); border-radius:var(--radius-lg); overflow:hidden;
  }

  .astRoot header { background:var(--navy); padding:14px 20px; display:flex; align-items:center; justify-content:space-between; gap:12px; flex-wrap:wrap; }
  .astRoot .h-title { font-size:16px; font-weight:600; color:#fff; }
  .astRoot .h-sub { font-size:11px; color:rgba(255,255,255,0.5); margin-top:2px; }
  .astRoot .app-version { font-size:10px; font-weight:600; color:rgba(255,255,255,0.4); margin-left:8px; }
  .astRoot .h-actions { display:flex; gap:8px; align-items:center; flex-wrap:wrap; }
  .astRoot .hbtn { font-size:13px; padding:6px 12px; border-radius:var(--radius); border:1px solid rgba(255,255,255,0.18); background:rgba(255,255,255,0.10); color:#fff; cursor:pointer; text-decoration:none; display:inline-flex; align-items:center; gap:5px; white-space:nowrap; }
  .astRoot .hbtn:hover { background:rgba(255,255,255,0.18); }
  .astRoot .hbtn.accent { background:var(--blue); border-color:var(--blue); }
  .astRoot .hbtn:disabled { opacity:.5; cursor:default; }

  .astRoot .tabs { display:flex; gap:2px; background:var(--navy); padding:0 20px; }
  .astRoot .tab { font-size:13px; padding:9px 16px; color:rgba(255,255,255,0.6); background:none; border:none; cursor:pointer; border-bottom:3px solid transparent; }
  .astRoot .tab:hover { color:#fff; }
  .astRoot .tab.active { color:#fff; border-bottom-color:#7FA8DC; background:rgba(255,255,255,0.06); }

  .astRoot .filters { display:flex; gap:10px; align-items:center; flex-wrap:wrap; padding:12px 20px; background:var(--surface); border-bottom:1px solid var(--border); }
  .astRoot .filters select, .astRoot .filters input { font-size:13px; padding:6px 10px; border-radius:var(--radius); border:1px solid var(--border-md); background:#fff; color:var(--text); outline:none; }
  .astRoot .filters input { min-width:200px; }
  .astRoot .filters select:focus, .astRoot .filters input:focus { border-color:var(--accent); }
  .astRoot .flabel { font-size:11px; color:var(--text-3); text-transform:uppercase; letter-spacing:.04em; }
  .astRoot .lnk { font-size:12px; color:var(--accent); background:none; border:none; cursor:pointer; text-decoration:underline; }

  .astRoot main { padding:18px 20px; display:flex; flex-direction:column; gap:16px; max-height:78vh; overflow-y:auto; }
  .astRoot .view { display:none; flex-direction:column; gap:16px; }
  .astRoot .view.active { display:flex; }

  .astRoot .kpi-grid { display:grid; grid-template-columns:repeat(4,1fr); gap:12px; }
  @media(max-width:900px){ .astRoot .kpi-grid { grid-template-columns:repeat(2,1fr); } }
  .astRoot .kpi { background:var(--surface); border:1px solid var(--border); border-radius:var(--radius-lg); padding:14px 16px; }
  .astRoot .kpi-label { font-size:11px; color:var(--text-3); text-transform:uppercase; letter-spacing:.05em; }
  .astRoot .kpi-value { font-size:24px; font-weight:600; margin-top:5px; }
  .astRoot .kpi-sub { font-size:11px; color:var(--text-3); margin-top:4px; }

  .astRoot .panel { background:var(--surface); border:1px solid var(--border); border-radius:var(--radius-lg); padding:16px 18px; }
  .astRoot .panel-title { font-size:11px; font-weight:600; color:var(--text-3); text-transform:uppercase; letter-spacing:.06em; margin-bottom:12px; }
  .astRoot .grid2 { display:grid; grid-template-columns:1fr 1fr; gap:16px; }
  @media(max-width:900px){ .astRoot .grid2 { grid-template-columns:1fr; } }

  .astRoot .tbl-wrap { overflow-x:auto; }
  .astRoot table { width:100%; border-collapse:collapse; font-size:13px; }
  .astRoot thead th { font-size:11px; font-weight:600; color:var(--text-3); text-transform:uppercase; letter-spacing:.04em; padding:8px 10px; background:var(--surface2); border-bottom:1px solid var(--border); text-align:left; white-space:nowrap; cursor:pointer; }
  .astRoot tbody td { padding:8px 10px; border-bottom:1px solid var(--border); vertical-align:middle; }
  .astRoot tbody tr:hover td { background:#fafbfd; }
  .astRoot .num { text-align:right; white-space:nowrap; }
  .astRoot .fw { font-weight:600; } .astRoot .muted { color:var(--text-2); }
  .astRoot .pill { display:inline-block; font-size:11px; font-weight:600; padding:2px 8px; border-radius:20px; background:var(--accent-lt); color:var(--blue); }
  .astRoot .bar-track { height:6px; background:var(--border); border-radius:3px; overflow:hidden; min-width:70px; }
  .astRoot .bar-fill { height:100%; background:var(--blue); border-radius:3px; }
  .astRoot input.gl { width:120px; font-size:12px; padding:4px 6px; border:1px solid var(--border-md); border-radius:6px; background:#fffdf0; }
  .astRoot input.gl:focus { outline:none; border-color:var(--accent); }

  .astRoot .state { padding:40px 20px; text-align:center; color:var(--text-2); }
  .astRoot .state-title { font-size:16px; font-weight:600; color:var(--text); margin-bottom:6px; }
  .astRoot .state-detail { font-size:12px; color:var(--text-3); margin-top:10px; word-break:break-word; }
  .astRoot .warn { font-size:12px; color:#8a6d3b; background:#fcf8e3; border:1px solid #faebcc; border-radius:var(--radius); padding:8px 12px; }
</style>

<div class="astRoot">
  <header>
    <div>
      <div class="h-title">Stream-Flo IT &mdash; Equipment Assets<span class="app-version" data-el="appVersion"></span></div>
      <div class="h-sub" data-el="headerSub">Built from the Equipment Asset Allocation forms</div>
    </div>
    <div class="h-actions">
      <button type="button" class="hbtn" data-el="refreshBtn" title="Re-read the forms folder">&#8635; Refresh</button>
      <a class="hbtn accent" data-el="uploadBtn" target="_blank" rel="noopener" title="Open the forms folder to upload a new allocation form">&#10514; Upload form</a>
      <button type="button" class="hbtn" data-el="exportBtn" title="Download the styled Excel asset log">&#8681; Excel log</button>
    </div>
  </header>

  <div class="tabs">
    <button type="button" class="tab active" data-tab="dash">Dashboard</button>
    <button type="button" class="tab" data-tab="items">Items</button>
    <button type="button" class="tab" data-tab="emps">Employees</button>
    <button type="button" class="tab" data-tab="forms">Forms</button>
  </div>

  <div class="filters">
    <span class="flabel">Filter</span>
    <select data-el="fEmp"><option value="">All employees</option></select>
    <select data-el="fDept"><option value="">All departments</option></select>
    <select data-el="fLoc"><option value="">All locations</option></select>
    <select data-el="fCls"><option value="">All classes</option></select>
    <select data-el="fYear"><option value="">All years</option></select>
    <input type="search" data-el="fQ" placeholder="Search item, serial, PO, ticket&hellip;">
    <button type="button" class="lnk" data-el="fReset">Reset</button>
    <span class="flabel" data-el="fCount"></span>
  </div>

  <main>
    <div data-el="view-dash"  class="view active"></div>
    <div data-el="view-items" class="view"></div>
    <div data-el="view-emps"  class="view"></div>
    <div data-el="view-forms" class="view"></div>
  </main>
</div>
`;
