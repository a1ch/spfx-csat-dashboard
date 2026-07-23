// The dashboard shell. All CSS is scoped under .opsRoot so nothing leaks onto
// the host SharePoint page, and the original full-viewport app layout is
// adapted to a bounded, self-scrolling panel. The core populates the filters,
// the branch nav, and the four views at runtime.
export const OPS_DASHBOARD_HTML: string = `
<style>
  .opsRoot, .opsRoot *, .opsRoot *::before, .opsRoot *::after { box-sizing:border-box; }
  .opsRoot * { margin:0; padding:0; }
  .opsRoot {
    --navy:#1a2740; --navy-mid:#243351; --accent:#185FA5; --accent-lt:#E6F1FB;
    --accent-bdr:rgba(24,95,165,0.25); --bg:#f4f6f9; --surface:#ffffff;
    --surface2:#f0f2f5; --border:rgba(0,0,0,0.08); --border-md:rgba(0,0,0,0.14);
    --text:#1a2740; --text-2:#5a6478; --text-3:#8a93a6;
    --green:#0F6E56; --green-bg:#E1F5EE; --red:#A32D2D; --red-bg:#FCEBEB;
    --radius:8px; --radius-lg:12px; --sidebar-w:220px; --header-h:60px;
    font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif; font-size:14px;
    color:var(--text); background:var(--bg); height:82vh; min-height:520px; overflow:hidden;
    display:flex; flex-direction:column; border:1px solid var(--border); border-radius:var(--radius-lg);
  }

  .opsRoot header { height:var(--header-h); background:var(--navy); display:flex; align-items:center; justify-content:space-between; padding:0 24px; flex-shrink:0; border-radius:var(--radius-lg) var(--radius-lg) 0 0; }
  .opsRoot .header-left { display:flex; align-items:center; gap:14px; }
  .opsRoot .header-logo { width:32px; height:32px; border-radius:8px; background:var(--accent); display:flex; align-items:center; justify-content:center; font-size:15px; font-weight:700; color:#fff; }
  .opsRoot .header-title { font-size:16px; font-weight:600; color:#fff; letter-spacing:-0.01em; }
  .opsRoot .header-sub { font-size:11px; color:rgba(255,255,255,0.45); margin-top:1px; }
  .opsRoot .app-version { font-size:10px; font-weight:600; color:rgba(255,255,255,0.4); margin-left:8px; }
  .opsRoot .header-controls { display:flex; gap:10px; align-items:center; }
  .opsRoot .header-controls select { font-size:13px; padding:6px 10px; border-radius:var(--radius); border:1px solid rgba(255,255,255,0.15); background:rgba(255,255,255,0.08); color:#fff; cursor:pointer; outline:none; }
  .opsRoot .header-controls select option { background:var(--navy); color:#fff; }
  .opsRoot .header-controls select:hover { border-color:rgba(255,255,255,0.3); }

  .opsRoot .body-wrap { display:flex; flex:1; overflow:hidden; }

  .opsRoot aside { width:var(--sidebar-w); background:var(--navy-mid); flex-shrink:0; display:flex; flex-direction:column; padding:16px 0; overflow-y:auto; }
  .opsRoot .nav-section-label { font-size:10px; font-weight:600; letter-spacing:.08em; color:rgba(255,255,255,0.3); text-transform:uppercase; padding:16px 18px 6px; }
  .opsRoot .nav-section-label:first-child { padding-top:4px; }
  .opsRoot .nav-item { display:flex; align-items:center; gap:10px; padding:9px 18px; cursor:pointer; font-size:13px; color:rgba(255,255,255,0.55); border:none; background:none; width:100%; text-align:left; border-left:3px solid transparent; transition:background 0.15s,color 0.15s; }
  .opsRoot .nav-item:hover { background:rgba(255,255,255,0.06); color:rgba(255,255,255,0.85); }
  .opsRoot .nav-item.active { color:#fff; background:rgba(24,95,165,0.25); border-left-color:var(--accent); }
  .opsRoot .nav-icon { font-size:15px; opacity:0.8; width:18px; text-align:center; }
  .opsRoot .nav-divider { height:1px; background:rgba(255,255,255,0.07); margin:8px 0; }
  .opsRoot .sidebar-footer { margin-top:auto; padding:14px 18px; font-size:11px; color:rgba(255,255,255,0.25); line-height:1.5; border-top:1px solid rgba(255,255,255,0.07); }

  .opsRoot main { flex:1; overflow-y:auto; padding:24px 28px; display:flex; flex-direction:column; gap:20px; min-width:0; }
  .opsRoot .view { display:none; flex-direction:column; gap:20px; }
  .opsRoot .view.active { display:flex; }

  .opsRoot .page-title-row { display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:10px; }
  .opsRoot .page-title { font-size:20px; font-weight:600; color:var(--text); letter-spacing:-0.02em; }
  .opsRoot .page-sub { font-size:12px; color:var(--text-2); margin-top:3px; }
  .opsRoot .page-title-right { display:flex; align-items:center; gap:10px; flex-wrap:wrap; }
  .opsRoot .period-badge { font-size:11px; font-weight:500; padding:4px 10px; background:var(--accent-lt); color:var(--accent); border-radius:20px; border:1px solid var(--accent-bdr); white-space:nowrap; }

  .opsRoot .period-selector { display:flex; align-items:center; gap:8px; }
  .opsRoot .period-selector label { font-size:12px; color:var(--text-2); white-space:nowrap; }
  .opsRoot .period-selector select { font-size:13px; padding:5px 10px; border-radius:var(--radius); border:1px solid var(--border-md); background:var(--surface); color:var(--text); cursor:pointer; outline:none; }
  .opsRoot .period-selector select:focus { border-color:var(--accent); }

  .opsRoot .kpi-grid { display:grid; grid-template-columns:repeat(4,1fr); gap:14px; }
  @media(max-width:900px){ .opsRoot .kpi-grid { grid-template-columns:repeat(2,1fr); } }
  .opsRoot .kpi-card { background:var(--surface); border-radius:var(--radius-lg); border:1px solid var(--border); padding:16px 18px; }
  .opsRoot .kpi-label { font-size:11px; color:var(--text-3); text-transform:uppercase; letter-spacing:.05em; margin-bottom:6px; }
  .opsRoot .kpi-value { font-size:26px; font-weight:600; color:var(--text); line-height:1; }
  .opsRoot .kpi-value.pos { color:var(--green); }
  .opsRoot .kpi-value.neg { color:var(--red); }
  .opsRoot .kpi-sub { font-size:11px; color:var(--text-3); margin-top:5px; }

  .opsRoot .panel { background:var(--surface); border-radius:var(--radius-lg); border:1px solid var(--border); padding:20px 22px; }
  .opsRoot .panel-title { font-size:11px; font-weight:600; color:var(--text-3); text-transform:uppercase; letter-spacing:.06em; margin-bottom:16px; }

  .opsRoot .chart-wrap { position:relative; width:100%; }
  .opsRoot .legend { display:flex; gap:16px; margin-bottom:12px; font-size:12px; color:var(--text-2); align-items:center; flex-wrap:wrap; }
  .opsRoot .leg-sq { width:10px; height:10px; border-radius:2px; display:inline-block; margin-right:5px; vertical-align:middle; }

  .opsRoot .tbl-wrap { overflow-x:auto; }
  .opsRoot table { width:100%; border-collapse:collapse; font-size:13px; }
  .opsRoot thead th { font-size:11px; font-weight:600; color:var(--text-3); text-transform:uppercase; letter-spacing:.04em; padding:8px 12px; background:var(--surface2); border-bottom:1px solid var(--border); text-align:left; white-space:nowrap; }
  .opsRoot tbody td { padding:9px 12px; border-bottom:1px solid var(--border); color:var(--text); vertical-align:middle; }
  .opsRoot tbody tr:last-child td { border-bottom:none; }
  .opsRoot tbody tr:hover td { background:#fafbfd; }
  .opsRoot .pos { color:var(--green); } .opsRoot .neg { color:var(--red); }
  .opsRoot .fw { font-weight:600; } .opsRoot .muted { color:var(--text-2); }

  .opsRoot .mom-pill { display:inline-flex; align-items:center; gap:3px; font-size:12px; font-weight:600; padding:2px 8px; border-radius:20px; }
  .opsRoot .mom-pill.up   { background:var(--green-bg); color:#085041; }
  .opsRoot .mom-pill.down { background:var(--red-bg);   color:#791F1F; }
  .opsRoot .mom-pill.flat { background:var(--surface2); color:var(--text-3); }

  .opsRoot .pbar-row { display:flex; align-items:center; gap:10px; }
  .opsRoot .pbar-track { flex:1; height:6px; background:var(--border); border-radius:3px; overflow:hidden; min-width:80px; }
  .opsRoot .pbar-fill { height:100%; border-radius:3px; }
  .opsRoot .pbar-label { font-size:12px; font-weight:600; min-width:36px; color:var(--text); }

  .opsRoot textarea.note { width:100%; font-size:12px; font-family:inherit; padding:5px 8px; border:1px solid var(--border-md); border-radius:var(--radius); background:#fffdf0; color:var(--text); resize:vertical; min-height:34px; transition:border-color 0.15s; }
  .opsRoot textarea.note:focus { outline:none; border-color:var(--accent); }

  .opsRoot .branch-section { margin-bottom:28px; }
  .opsRoot .branch-heading { font-size:13px; font-weight:700; color:var(--text); padding:10px 12px; background:var(--surface2); border-radius:var(--radius) var(--radius) 0 0; border:1px solid var(--border); border-bottom:none; text-transform:uppercase; letter-spacing:.05em; }
  .opsRoot .branch-table-wrap { border:1px solid var(--border); border-radius:0 0 var(--radius) var(--radius); overflow:hidden; }

  .opsRoot .state-screen { padding:48px 24px; text-align:center; color:var(--text-2); }
  .opsRoot .state-title { font-size:16px; font-weight:600; color:var(--text); margin:10px 0 6px; }
  .opsRoot .state-sub { font-size:13px; max-width:520px; margin:0 auto; }
  .opsRoot .state-detail { font-size:12px; color:var(--text-3); margin-top:12px; word-break:break-word; }
  .opsRoot .btn { font-size:13px; font-weight:600; padding:6px 16px; border-radius:var(--radius); border:none; background:var(--accent); color:#fff; cursor:pointer; white-space:nowrap; }
  .opsRoot .btn:hover { background:#145088; }

  .opsRoot main::-webkit-scrollbar { width:6px; }
  .opsRoot main::-webkit-scrollbar-track { background:transparent; }
  .opsRoot main::-webkit-scrollbar-thumb { background:var(--border-md); border-radius:3px; }
  .opsRoot aside::-webkit-scrollbar { width:4px; }
  .opsRoot aside::-webkit-scrollbar-thumb { background:rgba(255,255,255,0.1); border-radius:2px; }
</style>

<div class="opsRoot">
  <header>
    <div class="header-left">
      <div class="header-logo">FY</div>
      <div>
        <div class="header-title">L48 FY27 Operations Performance<span class="app-version" data-el="appVersion"></span></div>
        <div class="header-sub" data-el="headerSub">Revenue vs. budget &middot; FY start Apr 1 2026</div>
      </div>
    </div>
    <div class="header-controls">
      <select data-el="branchSel"><option value="ALL">All branches</option></select>
      <select data-el="metricSel"><option value="ALL">All metrics</option></select>
    </div>
  </header>

  <div class="body-wrap">
    <aside>
      <div class="nav-section-label">Views</div>
      <button class="nav-item active" data-tab="overview"><span class="nav-icon">&#9672;</span> Overview</button>
      <button class="nav-item" data-tab="monthly"><span class="nav-icon">&#8801;</span> Monthly</button>
      <button class="nav-item" data-tab="quarterly"><span class="nav-icon">&#11041;</span> Quarterly</button>
      <button class="nav-item" data-tab="branch"><span class="nav-icon">&#9678;</span> Branch detail</button>
      <div class="nav-divider"></div>
      <div class="nav-section-label">Branches</div>
      <div data-el="branchNav"></div>
      <div class="sidebar-footer" data-el="footer">FY27 &middot; Apr 2026 &ndash; Mar 2027</div>
    </aside>

    <main data-el="main">
      <div data-el="view-overview"  class="view active"></div>
      <div data-el="view-monthly"   class="view"></div>
      <div data-el="view-quarterly" class="view"></div>
      <div data-el="view-branch"    class="view"></div>
    </main>
  </div>
</div>
`;
