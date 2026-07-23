/**
 * Static markup + styles for the CSAT dashboard. Ported from the standalone
 * HTML dashboard. Inline on* handlers were removed; the web part attaches
 * event listeners by id/data-action after injecting this markup, so nothing
 * depends on global functions (which is not allowed inside an SPFx bundle).
 *
 * Styles are scoped under .sfCsatRoot so they don't leak into the SharePoint
 * page around the web part.
 */
export const DASHBOARD_HTML: string = `
<style>
  .sfCsatRoot, .sfCsatRoot *, .sfCsatRoot *::before, .sfCsatRoot *::after { box-sizing: border-box; }
  .sfCsatRoot {
    --blue-dark: #1a3a5c; --blue-mid: #2a5a8c; --blue-light: #e6f1fb; --blue-link: #185fa5;
    --green-bg: #e1f5ee; --green-text: #0f6e56; --green-border: #5dcaa5;
    --amber-bg: #faeeda; --amber-text: #854f0b;
    --red-bg: #fcebeb; --red-text: #a32d2d;
    --gray-border: rgba(0,0,0,0.12);
    --text-primary: #1a1a1a; --text-secondary: #555; --text-tertiary: #888; --white: #fff;
    --radius-md: 8px; --radius-lg: 12px;
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    color: var(--text-primary);
  }
  .sfCsatRoot .app-header { background: var(--blue-dark); padding: 0.85rem 1.1rem; border-radius: var(--radius-lg) var(--radius-lg) 0 0; display: flex; align-items: center; justify-content: space-between; gap: 0.75rem; flex-wrap: wrap; }
  .sfCsatRoot .app-header-left { display: flex; align-items: center; gap: 0.75rem; }
  .sfCsatRoot .app-header-logo { width: 34px; height: 34px; background: var(--blue-mid); border-radius: 8px; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
  .sfCsatRoot .app-header-logo svg { width: 19px; height: 19px; fill: #4fc3f7; }
  .sfCsatRoot .app-header-title { color: #fff; font-size: 16px; font-weight: 600; line-height: 1.2; }
  .sfCsatRoot .app-header-sub { color: #7ab3d8; font-size: 11px; margin-top: 2px; }
  .sfCsatRoot .app-header-actions { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
  .sfCsatRoot .live-badge { display: flex; align-items: center; gap: 6px; background: rgba(255,255,255,0.12); color: #9fe1cb; font-size: 12px; font-weight: 500; padding: 6px 14px; border-radius: 20px; }
  .sfCsatRoot .live-dot { width: 6px; height: 6px; border-radius: 50%; background: #5dcaa5; }
  .sfCsatRoot .live-dot.error { background: #f09595; }
  .sfCsatRoot .live-dot.loading { background: #fac775; animation: sfpulse 1s infinite; }
  @keyframes sfpulse { 0%,100% { opacity: 1; } 50% { opacity: 0.3; } }
  .sfCsatRoot .hbtn { background: rgba(255,255,255,0.12); border: none; color: #fff; font-size: 12px; padding: 6px 12px; border-radius: 20px; cursor: pointer; display: flex; align-items: center; gap: 6px; }
  .sfCsatRoot .hbtn:hover { background: rgba(255,255,255,0.2); }
  .sfCsatRoot .hbtn:disabled { opacity: 0.5; cursor: default; }
  .sfCsatRoot .hbtn.active { background: #5dcaa5; color: #08331f; }
  .sfCsatRoot .dropdown { position: relative; }
  .sfCsatRoot .dropdown-menu { position: absolute; right: 0; top: calc(100% + 6px); background: #fff; border: 1px solid var(--gray-border); border-radius: var(--radius-md); box-shadow: 0 8px 24px rgba(0,0,0,0.18); min-width: 180px; overflow: hidden; display: none; z-index: 200; }
  .sfCsatRoot .dropdown-menu.open { display: block; }
  .sfCsatRoot .dropdown-menu button { display: block; width: 100%; text-align: left; background: none; border: none; padding: 10px 14px; font-size: 13px; color: var(--text-primary); cursor: pointer; }
  .sfCsatRoot .dropdown-menu button:hover { background: var(--blue-light); }
  .sfCsatRoot .container { background: #f0f2f5; padding: 1.1rem; border-radius: 0 0 var(--radius-lg) var(--radius-lg); }

  .sfCsatRoot .state-screen { text-align: center; padding: 3rem 1.5rem; background: var(--white); border-radius: var(--radius-lg); border: 1px solid var(--gray-border); }
  .sfCsatRoot .state-icon { font-size: 40px; margin-bottom: 1rem; }
  .sfCsatRoot .state-title { font-size: 18px; font-weight: 600; margin-bottom: 8px; }
  .sfCsatRoot .state-sub { font-size: 14px; color: var(--text-secondary); margin-bottom: 1.25rem; line-height: 1.5; }
  .sfCsatRoot .state-detail { font-size: 12px; color: var(--text-tertiary); background: #fafafa; border-radius: var(--radius-md); padding: 10px 14px; text-align: left; font-family: monospace; white-space: pre-wrap; word-break: break-word; margin-top: 1rem; }

  .sfCsatRoot .filter-row { display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; margin-bottom: 1.1rem; }
  .sfCsatRoot .filter-row select { width: 100%; padding: 8px 12px; border: 1px solid var(--gray-border); border-radius: var(--radius-md); font-size: 13px; background: #fff; color: var(--text-primary); font-family: inherit; }
  @media (max-width: 600px) { .sfCsatRoot .filter-row { grid-template-columns: 1fr; } }

  .sfCsatRoot .kpi-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 1.1rem; }
  @media (max-width: 700px) { .sfCsatRoot .kpi-grid { grid-template-columns: repeat(2, 1fr); } }
  .sfCsatRoot .kpi-card { background: var(--white); border: 1px solid var(--gray-border); border-radius: var(--radius-md); padding: 1rem; }
  .sfCsatRoot .kpi-label { font-size: 12px; color: var(--text-tertiary); margin-bottom: 6px; }
  .sfCsatRoot .kpi-value { font-size: 26px; font-weight: 600; color: var(--text-primary); }
  .sfCsatRoot .kpi-value small { font-size: 14px; color: var(--text-tertiary); font-weight: 400; }
  .sfCsatRoot .kpi-sub { font-size: 11px; color: var(--text-tertiary); margin-top: 4px; }
  .sfCsatRoot .delta { font-size: 12px; font-weight: 600; margin-left: 6px; }
  .sfCsatRoot .delta.up { color: var(--green-text); }
  .sfCsatRoot .delta.down { color: var(--red-text); }
  .sfCsatRoot .delta.flat { color: var(--text-tertiary); }

  .sfCsatRoot .section-card { background: var(--white); border: 1px solid var(--gray-border); border-radius: var(--radius-lg); padding: 1.25rem; margin-bottom: 1rem; }
  .sfCsatRoot .section-title { font-size: 13px; font-weight: 600; color: var(--text-secondary); margin-bottom: 1rem; text-transform: uppercase; letter-spacing: 0.04em; }
  .sfCsatRoot .section-title .count-pill { text-transform: none; letter-spacing: normal; background: var(--blue-light); color: var(--blue-link); font-size: 11px; padding: 2px 8px; border-radius: 12px; margin-left: 8px; }
  .sfCsatRoot .empty-note { font-size: 13px; color: var(--text-tertiary); text-align: center; padding: 1.5rem 0; }

  .sfCsatRoot table { width: 100%; font-size: 13px; border-collapse: collapse; }
  .sfCsatRoot th { font-size: 11px; font-weight: 600; color: var(--text-tertiary); text-align: left; padding: 0 8px 8px; border-bottom: 1px solid var(--gray-border); text-transform: uppercase; letter-spacing: 0.03em; }
  .sfCsatRoot td { padding: 9px 8px; border-bottom: 1px solid #f0f0f0; color: var(--text-primary); }
  .sfCsatRoot tr:last-child td { border-bottom: none; }
  .sfCsatRoot .badge { display: inline-block; padding: 2px 10px; border-radius: 20px; font-size: 11px; font-weight: 600; }
  .sfCsatRoot .badge-good { background: var(--green-bg); color: var(--green-text); }
  .sfCsatRoot .badge-warn { background: var(--amber-bg); color: var(--amber-text); }
  .sfCsatRoot .badge-bad { background: var(--red-bg); color: var(--red-text); }

  .sfCsatRoot .comment-card { background: #fafafa; border-radius: var(--radius-md); padding: 10px 14px; margin-bottom: 8px; border-left: 3px solid var(--blue-link); }
  .sfCsatRoot .comment-card.attention { border-left-color: var(--red-text); background: #fff7f7; }
  .sfCsatRoot .comment-text { font-size: 13px; color: var(--text-primary); line-height: 1.5; margin-bottom: 5px; }
  .sfCsatRoot .comment-meta { font-size: 11px; color: var(--text-tertiary); }

  .sfCsatRoot .improv-bar { margin-bottom: 10px; }
  .sfCsatRoot .improv-label { display: flex; justify-content: space-between; font-size: 12px; color: var(--text-secondary); margin-bottom: 4px; }
  .sfCsatRoot .improv-track { background: #f0f0f0; border-radius: 4px; height: 10px; overflow: hidden; }
  .sfCsatRoot .improv-fill { background: var(--blue-dark); border-radius: 4px; height: 10px; }

  .sfCsatRoot .nps-gauge { display: flex; gap: 10px; flex-wrap: wrap; }
  .sfCsatRoot .nps-box { flex: 1; min-width: 90px; border-radius: var(--radius-md); padding: 12px; text-align: center; }
  .sfCsatRoot .nps-box .val { font-size: 20px; font-weight: 600; }
  .sfCsatRoot .nps-box .lbl { font-size: 11px; margin-top: 3px; }
  .sfCsatRoot .nps-det { background: var(--red-bg); color: var(--red-text); }
  .sfCsatRoot .nps-pas { background: var(--amber-bg); color: var(--amber-text); }
  .sfCsatRoot .nps-pro { background: var(--green-bg); color: var(--green-text); }
  .sfCsatRoot .nps-score-box { background: var(--blue-dark); color: #fff; border-radius: var(--radius-md); padding: 12px; text-align: center; min-width: 100px; }
  .sfCsatRoot .nps-score-box .val { font-size: 26px; font-weight: 600; }
  .sfCsatRoot .nps-score-box .lbl { font-size: 10px; color: #7ab3d8; margin-top: 3px; }

  .sfCsatRoot .star-row { display: flex; align-items: center; gap: 8px; font-size: 12px; padding: 6px 0; border-bottom: 1px solid #f0f0f0; }
  .sfCsatRoot .star-row:last-child { border-bottom: none; }
  .sfCsatRoot .star-name { color: var(--text-secondary); flex: 1; min-width: 0; }
  .sfCsatRoot .star-score { font-weight: 600; color: var(--text-primary); min-width: 28px; text-align: right; }
  .sfCsatRoot .stars-mini { color: #f5a623; font-size: 14px; letter-spacing: -1px; }

  .sfCsatRoot .last-sync { font-size: 11px; color: var(--text-tertiary); text-align: right; margin-top: 4px; }
  .sfCsatRoot .app-version { font-size: 10px; font-weight: 600; letter-spacing: .3px; color: var(--text-tertiary); opacity: .75; margin-left: 8px; white-space: nowrap; }
</style>

<div class="sfCsatRoot">
  <div class="app-header">
    <div class="app-header-left">
      <div class="app-header-logo"><svg viewBox="0 0 24 24"><path d="M12 2C6 8 4 12 4 15a8 8 0 0016 0c0-3-2-7-8-13z"/></svg></div>
      <div>
        <div class="app-header-title">Stream-Flo &mdash; Live CSAT Dashboard<span class="app-version" data-el="appVersion"></span></div>
        <div class="app-header-sub" data-el="connSub">Connecting to SharePoint&hellip;</div>
      </div>
    </div>
    <div class="app-header-actions">
      <div class="live-badge"><span class="live-dot loading" data-el="liveDot"></span><span data-el="liveText">Connecting&hellip;</span></div>
      <button class="hbtn" data-action="auto" title="Auto-refresh">&#9711; Auto</button>
      <div class="dropdown">
        <button class="hbtn" data-action="exportToggle">&#10515; Export &#9662;</button>
        <div class="dropdown-menu" data-el="exportMenu">
          <button data-action="exportCsv">Download CSV (raw data)</button>
          <button data-action="exportXlsx">Download Excel &mdash; charts + branch sheets</button>
        </div>
      </div>
      <button class="hbtn" data-action="refresh">&#10226; Refresh</button>
    </div>
  </div>

  <div class="container">
    <div data-el="stateScreen"></div>

    <div data-el="dashboardContent" style="display:none">
      <div class="filter-row">
        <select data-el="branchFilter"><option value="">All branches</option></select>
        <select data-el="rangeFilter">
          <option value="9999">All time</option>
          <option value="90">Last 90 days</option>
          <option value="30">Last 30 days</option>
          <option value="7">Last 7 days</option>
        </select>
        <select data-el="techFilter"><option value="">All technicians</option></select>
      </div>

      <div class="kpi-grid">
        <div class="kpi-card"><div class="kpi-label">Total responses</div><div class="kpi-value" data-el="kTotal">&mdash;</div><div class="kpi-sub" data-el="kTotalSub"></div></div>
        <div class="kpi-card"><div class="kpi-label">Avg CSAT</div><div class="kpi-value"><span data-el="kCsat">&mdash;</span><small>/5</small><span class="delta" data-el="kCsatDelta"></span></div></div>
        <div class="kpi-card"><div class="kpi-label">NPS score</div><div class="kpi-value"><span data-el="kNps">&mdash;</span><span class="delta" data-el="kNpsDelta"></span></div></div>
        <div class="kpi-card"><div class="kpi-label">Safety score</div><div class="kpi-value"><span data-el="kSafety">&mdash;</span><small>/5</small><span class="delta" data-el="kSafetyDelta"></span></div></div>
      </div>

      <div class="section-card" data-el="attentionCard" style="display:none">
        <div class="section-title">Needs attention <span class="count-pill" data-el="attnCount">0</span></div>
        <div data-el="attentionList"></div>
      </div>

      <div class="section-card">
        <div class="section-title">Rating breakdown by category</div>
        <div data-el="starBreakdown"></div>
      </div>

      <div class="section-card">
        <div class="section-title">NPS breakdown</div>
        <div class="nps-gauge">
          <div class="nps-score-box"><div class="val" data-el="npsScoreVal">&mdash;</div><div class="lbl">Net Promoter Score</div></div>
          <div class="nps-box nps-det"><div class="val" data-el="detCount">0</div><div class="lbl">Detractors (0&ndash;6)</div></div>
          <div class="nps-box nps-pas"><div class="val" data-el="pasCount">0</div><div class="lbl">Passives (7&ndash;8)</div></div>
          <div class="nps-box nps-pro"><div class="val" data-el="proCount">0</div><div class="lbl">Promoters (9&ndash;10)</div></div>
        </div>
      </div>

      <div class="section-card">
        <div class="section-title">CSAT trend over time</div>
        <div style="position:relative; height:240px"><canvas data-el="trendChart" role="img" aria-label="Line chart of average CSAT score over time"></canvas></div>
      </div>

      <div class="section-card">
        <div class="section-title">Performance by branch</div>
        <div style="position:relative; height:260px"><canvas data-el="branchChart" role="img" aria-label="Bar chart comparing average CSAT score across branches"></canvas></div>
      </div>

      <div class="section-card">
        <div class="section-title">Improvement opportunities &mdash; most frequently cited</div>
        <div data-el="improvBars"></div>
      </div>

      <div class="section-card">
        <div class="section-title">Technician performance</div>
        <table>
          <thead><tr><th>Technician</th><th>Jobs</th><th>Avg rating</th><th>Status</th></tr></thead>
          <tbody data-el="techTable"></tbody>
        </table>
      </div>

      <div class="section-card">
        <div class="section-title">Exemplary performance &mdash; customer comments</div>
        <div data-el="exemplaryList"></div>
      </div>

      <div class="section-card">
        <div class="section-title">Improvement suggestions &mdash; verbatim</div>
        <div data-el="improveList"></div>
      </div>

      <div class="last-sync" data-el="lastSync"></div>
    </div>
  </div>
  <div data-el="exportChartHolder" aria-hidden="true" style="position:fixed; left:-10000px; top:0; width:900px; height:500px; background:#fff;"></div>
</div>
`;