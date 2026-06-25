<%@ Page Language="C#" %>
<%@ Register TagPrefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>Training Feedback Portal</title>
<!-- SheetJS for Excel parsing and CSV generation -->
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<!-- JSZip for bundling CSV + feedback files into a ZIP -->
<script src="https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js"></script>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Tahoma,sans-serif;font-size:14px;background:#f3f2f1;color:#323130;min-height:100vh}

/* ── SUITE BAR ── */
.sp-bar{background:#0078d4;height:42px;display:flex;align-items:center;padding:0 20px;gap:10px;position:sticky;top:0;z-index:100}
.sp-bar .app{color:#fff;font-size:15px;font-weight:600}
.sp-bar .site{color:rgba(255,255,255,.72);font-size:13px}
.sp-bar .divdr{color:rgba(255,255,255,.35)}

/* ── PAGE HEADER ── */
.ph{background:#fff;border-bottom:1px solid #edebe9;padding:18px 32px 0}
.ph-top{display:flex;align-items:flex-start;justify-content:space-between;gap:16px;flex-wrap:wrap}
.ph h1{font-size:22px;font-weight:600;color:#201f1e;margin-bottom:3px}
.ph .sub{font-size:13px;color:#605e5c}
.ph-actions{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.ph-sp{height:14px}

/* ── BUTTONS ── */
.btn{display:inline-flex;align-items:center;gap:6px;padding:6px 16px;font-size:14px;font-family:inherit;border-radius:2px;cursor:pointer;border:1px solid transparent;transition:background .1s,opacity .1s;white-space:nowrap}
.btn-primary{background:#0078d4;color:#fff;border-color:#0078d4}
.btn-primary:hover{background:#106ebe}
.btn-primary:active{background:#005a9e}
.btn-success{background:#107c41;color:#fff;border-color:#107c41}
.btn-success:hover{background:#0e6b38}
.btn-default{background:#fff;color:#323130;border-color:#8a8886}
.btn-default:hover{background:#f3f2f1}
.btn-sm{padding:4px 10px;font-size:12px}
.btn[disabled]{opacity:.45;cursor:default;pointer-events:none}

/* ── MAIN ── */
.main{max-width:1400px;margin:0 auto;padding:20px 32px}

/* ── LOADING / ERROR ── */
.state-panel{background:#fff;border:1px solid #edebe9;border-radius:2px;padding:60px 24px;text-align:center;margin-bottom:20px}
.spinner{width:36px;height:36px;border:3px solid #edebe9;border-top-color:#0078d4;border-radius:50%;animation:spin .7s linear infinite;margin:0 auto 14px}
.spinner.green{border-top-color:#107c41}
@keyframes spin{to{transform:rotate(360deg)}}
.state-panel h3{font-size:16px;font-weight:600;margin-bottom:6px;color:#323130}
.state-panel p{font-size:13px;color:#605e5c;margin-bottom:4px}
.err-box{background:#fde7e9;border:1px solid #f1707b;border-radius:2px;padding:16px 20px;margin-bottom:20px;font-size:13px;color:#6e0811;display:none}
.err-box strong{display:block;margin-bottom:6px;font-size:15px}
.err-box code{font-family:Consolas,monospace;background:rgba(0,0,0,.07);padding:2px 6px;border-radius:2px;font-size:12px}

/* ── DATA BAR ── */
.data-bar{display:none;align-items:center;gap:12px;background:#e6f4ea;border:1px solid #b7dfbf;border-radius:2px;padding:10px 16px;margin-bottom:16px;flex-wrap:wrap}
.data-bar .dname{font-weight:600;font-size:14px;color:#054d1e}
.data-bar .dmeta{font-size:12px;color:#107c41;margin-top:2px}

/* ── SEARCH CARD ── */
.search-card{background:#fff;border:1px solid #edebe9;border-radius:2px;margin-bottom:16px;display:none}
.sc-head{padding:13px 18px;border-bottom:1px solid #edebe9;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px}
.sc-head h3{font-size:14px;font-weight:600;display:flex;align-items:center;gap:7px;color:#323130}
.sc-hint{font-size:12px;color:#605e5c;background:#f3f2f1;padding:4px 10px;border-radius:12px}
.sc-body{padding:18px;display:flex;gap:16px;flex-wrap:wrap;align-items:flex-end}
.fg{display:flex;flex-direction:column;gap:5px;position:relative}
.fg label{font-size:12px;font-weight:600;color:#323130;letter-spacing:.03em}
.fg label .req{color:#a4262c;margin-left:2px}
.fg input[type=text],.fg input[type=date]{height:34px;padding:0 10px;font-family:inherit;font-size:14px;color:#323130;background:#fff;border:1px solid #8a8886;border-radius:2px;outline:none;min-width:210px;transition:border-color .1s}
.fg input:focus{border-color:#0078d4;box-shadow:0 0 0 1px #0078d4}
.fg input.has-error{border-color:#a4262c!important;box-shadow:0 0 0 1px #a4262c!important}
.sc-div{width:1px;height:58px;background:#edebe9;align-self:flex-end;margin:0 4px}
.sc-actions{display:flex;gap:8px;align-self:flex-end;margin-left:auto}

/* ── AUTOCOMPLETE ── */
.ac-list{position:absolute;top:100%;left:0;right:0;background:#fff;border:1px solid #8a8886;border-top:none;border-radius:0 0 2px 2px;max-height:180px;overflow-y:auto;z-index:200;display:none;box-shadow:0 4px 12px rgba(0,0,0,.1)}
.ac-list.open{display:block}
.ac-item{padding:7px 10px;font-size:13px;cursor:pointer;color:#323130}
.ac-item:hover,.ac-item.hi{background:#eff6fc;color:#0078d4}

/* ── VALIDATION ── */
.val-banner{display:none;align-items:center;gap:8px;background:#fde7e9;border:1px solid #f1707b;border-radius:2px;padding:9px 14px;margin:0 18px 14px;font-size:13px;color:#6e0811}
.val-banner.show{display:flex}
.date-note{font-size:11px;color:#605e5c;padding:6px 18px 10px;border-top:1px solid #f3f2f1}

/* ── SUMMARY CARDS ── */
.sum-grid{display:none;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:12px;margin-bottom:16px}
.sum-card{background:#fff;border:1px solid #edebe9;border-top:3px solid #0078d4;border-radius:2px;padding:12px 16px}
.sum-card.g{border-top-color:#107c41}.sum-card.a{border-top-color:#d98000}
.sum-card.b{border-top-color:#005a9e}.sum-card.r{border-top-color:#a4262c}
.slabel{font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.06em;color:#605e5c;margin-bottom:5px}
.sval{font-size:26px;font-weight:600;color:#323130;line-height:1}

/* ── CMD BAR ── */
.cmd-bar{display:none;align-items:center;justify-content:space-between;margin-bottom:8px;flex-wrap:wrap;gap:8px}
.cmd-left{display:flex;align-items:center;gap:8px}
.cmd-right{display:flex;align-items:center;gap:8px}
.rec-count{font-size:13px;color:#605e5c}
.rec-count strong{color:#323130}

/* ── TABLE ── */
#tbl-section{display:none}
.tbl-wrap{background:#fff;border:1px solid #edebe9;border-radius:2px;overflow-x:auto}
table{width:100%;border-collapse:collapse;min-width:1020px}
thead tr{background:#faf9f8;border-bottom:2px solid #edebe9}
th{padding:10px 12px;font-size:12px;font-weight:600;color:#605e5c;text-align:left;white-space:nowrap;cursor:pointer;user-select:none;border-right:1px solid #f3f2f1}
th:last-child{border-right:none;cursor:default}
th:hover:not(:last-child){background:#f3f2f1;color:#323130}
th.asc .sarr::after{content:' ▲';color:#0078d4}
th.dsc .sarr::after{content:' ▼';color:#0078d4}
tbody tr{border-bottom:1px solid #edebe9}
tbody tr:last-child{border-bottom:none}
tbody tr:hover td{background:#f3f2f1}
td{padding:9px 12px;font-size:14px;color:#323130;vertical-align:middle;border-right:1px solid #f3f2f1}
td:last-child{border-right:none}

/* ── RATING BAR ── */
.rbar{display:flex;align-items:center;gap:7px}
.rfill{height:8px;border-radius:4px;min-width:2px;max-width:70px}
.rfill.good{background:#107c41}.rfill.mid{background:#d98000}.rfill.low{background:#a4262c}
.rnum{font-size:13px;font-weight:600;min-width:28px}

/* ── FEEDBACK FILE CELL ── */
.fb-cell{display:flex;align-items:center;gap:8px;flex-wrap:nowrap}
/* Hyperlink style — the main clickable link */
.fb-link{color:#0078d4;text-decoration:none;font-size:13px;display:inline-flex;align-items:center;gap:4px;background:none;border:none;cursor:pointer;font-family:inherit;padding:0;max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.fb-link:hover{text-decoration:underline;color:#004578}
/* Preview icon button */
.fb-preview-btn{background:none;border:1px solid #c8c6c4;border-radius:2px;padding:2px 7px;font-size:11px;cursor:pointer;font-family:inherit;color:#605e5c;display:inline-flex;align-items:center;gap:3px;white-space:nowrap}
.fb-preview-btn:hover{background:#f3f2f1;border-color:#8a8886}
.fb-na{color:#a19f9d;font-size:13px}

/* ── EXPORT / DOWNLOAD BAR ── */
.dl-bar{display:none;padding:14px 16px;border-top:1px solid #edebe9;background:#f0f6ff;border-bottom:1px solid #c7d9f5;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px}
.dl-bar-info{font-size:13px;color:#323130}
.dl-bar-info strong{color:#0078d4}
.dl-bar-actions{display:flex;gap:8px;flex-wrap:wrap}
/* ZIP progress overlay */
.zip-progress{display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:2000;align-items:center;justify-content:center}
.zip-progress.open{display:flex}
.zip-box{background:#fff;border-radius:4px;padding:32px 40px;text-align:center;min-width:300px;box-shadow:0 12px 40px rgba(0,0,0,.25)}
.zip-box h3{font-size:15px;font-weight:600;margin-bottom:6px;color:#323130}
.zip-box p{font-size:13px;color:#605e5c;margin-bottom:4px}
.zip-pbar-wrap{background:#edebe9;border-radius:4px;height:6px;margin-top:14px;overflow:hidden}
.zip-pbar{height:6px;background:#0078d4;border-radius:4px;transition:width .3s;width:0%}

/* ── EMPTY STATE ── */
.empty{text-align:center;padding:50px 24px;background:#fff;border:1px solid #edebe9;border-top:none;display:none}
.empty .eicon{font-size:40px;margin-bottom:10px}
.empty h3{font-size:16px;font-weight:600;margin-bottom:6px}
.empty p{font-size:13px;color:#605e5c}

/* ── PAGINATION ── */
.pg-bar{display:none;align-items:center;justify-content:space-between;padding:10px 16px;border-top:1px solid #edebe9;background:#faf9f8;flex-wrap:wrap;gap:8px}
.pg-info{font-size:13px;color:#605e5c}
.pg-btns{display:flex;gap:4px}
.pgb{min-width:32px;height:32px;padding:0 8px;font-family:inherit;font-size:13px;border:1px solid #8a8886;background:#fff;color:#323130;border-radius:2px;cursor:pointer;display:inline-flex;align-items:center;justify-content:center}
.pgb:hover:not([disabled]){background:#f3f2f1}
.pgb.on{background:#0078d4;color:#fff;border-color:#0078d4}
.pgb[disabled]{opacity:.4;cursor:default}

/* ── FEEDBACK MODAL ── */
.modal-bg{display:none;position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:1000;align-items:center;justify-content:center;padding:20px}
.modal-bg.open{display:flex}
.modal-box{background:#fff;border-radius:2px;width:100%;max-width:860px;max-height:92vh;display:flex;flex-direction:column;box-shadow:0 16px 48px rgba(0,0,0,.28)}
.mh{display:flex;align-items:flex-start;justify-content:space-between;padding:14px 20px;border-bottom:1px solid #edebe9;flex-shrink:0;gap:12px}
.mh-left{flex:1;min-width:0}
.mh-title{font-size:15px;font-weight:600;color:#201f1e;word-break:break-all}
.mh-sub{font-size:11px;color:#605e5c;margin-top:3px;font-family:Consolas,monospace;word-break:break-all}
.mh-actions{display:flex;align-items:center;gap:8px;flex-shrink:0}
.m-close{background:none;border:none;cursor:pointer;font-size:22px;color:#605e5c;line-height:1;padding:0 4px}
.m-close:hover{color:#323130}
.mb{flex:1;overflow:auto;padding:20px}
.fv-spin-wrap{text-align:center;padding:60px 20px;color:#605e5c}
.fv-pdf iframe{width:100%;height:62vh;border:none;border-radius:2px}
.fv-tbl-wrap{overflow:auto;max-height:62vh;border:1px solid #edebe9;border-radius:2px}
.fv-tbl-wrap table{width:100%;border-collapse:collapse;font-size:13px;min-width:unset}
.fv-tbl-wrap thead tr{background:#faf9f8;position:sticky;top:0;z-index:1}
.fv-tbl-wrap th{padding:8px 10px;font-weight:600;color:#605e5c;border:1px solid #edebe9}
.fv-tbl-wrap td{padding:7px 10px;border:1px solid #edebe9;color:#323130}
.fv-tbl-wrap tbody tr:nth-child(even){background:#faf9f8}
.fv-err{background:#fde7e9;border:1px solid #f1707b;border-radius:2px;padding:14px 16px;font-size:13px;color:#6e0811;line-height:1.6;margin-bottom:12px}
.fv-info{background:#eff6fc;border:1px solid #90c8ff;border-radius:2px;padding:14px 16px;font-size:13px;color:#004578;line-height:1.6;margin-bottom:12px}
.fv-err code,.fv-info code{font-family:Consolas,monospace;background:rgba(0,0,0,.06);padding:1px 5px;border-radius:2px;font-size:12px}

@media(max-width:700px){
  .main{padding:12px 14px}.ph{padding:14px 14px 0}
  .sc-div{display:none}.sc-actions{margin-left:0;width:100%}
}
</style>
</head>
<body>

<!-- SUITE BAR -->
<div class="sp-bar">
  <svg width="20" height="20" viewBox="0 0 20 20" fill="none">
    <rect width="20" height="20" rx="3" fill="rgba(255,255,255,.18)"/>
    <path d="M4 6h12M4 10h12M4 14h8" stroke="#fff" stroke-width="1.5" stroke-linecap="round"/>
  </svg>
  <span class="app">SharePoint</span>
  <span class="divdr">›</span>
  <span class="site">Training Feedback Portal</span>
</div>

<!-- PAGE HEADER -->
<div class="ph">
  <div class="ph-top">
    <div>
      <h1>Training Feedback Records</h1>
      <p class="sub">Search by Program Name or Trainer · click feedback links to preview · download all as ZIP</p>
    </div>
    <div class="ph-actions">
      <button class="btn btn-default btn-sm" id="btn-reload" onclick="loadExcelData()" style="display:none">
        <svg width="13" height="13" viewBox="0 0 16 16" fill="currentColor"><path d="M8 3a5 5 0 104.546 2.914.5.5 0 01.908-.417A6 6 0 118 2v1z"/><path d="M8 4.466V.534a.25.25 0 01.41-.192l2.36 1.966c.12.1.12.284 0 .384L8.41 4.658A.25.25 0 018 4.466z"/></svg>
        Refresh Data
      </button>
    </div>
  </div>
  <div class="ph-sp"></div>
</div>

<!-- MAIN -->
<div class="main">

  <!-- Loading -->
  <div class="state-panel" id="loading-panel">
    <div class="spinner"></div>
    <h3>Loading training data…</h3>
    <p id="loading-path-label" style="margin-top:8px;font-family:Consolas,monospace;font-size:12px;color:#0078d4"></p>
  </div>

  <!-- Error -->
  <div class="err-box" id="err-box">
    <strong id="err-title">Could not load the data file</strong>
    <div id="err-msg" style="margin-bottom:8px"></div>
    <div>Check file exists at: <code id="err-path"></code></div>
    <div style="margin-top:12px"><button class="btn btn-default btn-sm" onclick="loadExcelData()">↺ Retry</button></div>
  </div>

  <!-- Data loaded banner -->
  <div class="data-bar" id="data-bar">
    <svg width="20" height="20" viewBox="0 0 24 24" fill="#107c41"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8l-6-6zm-1 1.5L18.5 9H13V3.5zM6 20V4h5v7h7v9H6z"/></svg>
    <div style="flex:1;min-width:180px">
      <div class="dname" id="data-fname"></div>
      <div class="dmeta" id="data-fmeta"></div>
    </div>
    <span style="font-size:12px;color:#107c41;font-weight:600">✔ Data loaded successfully</span>
  </div>

  <!-- Search card -->
  <div class="search-card" id="search-card">
    <div class="sc-head">
      <h3>
        <svg width="15" height="15" viewBox="0 0 16 16" fill="#0078d4"><path d="M11.742 10.344a6.5 6.5 0 10-1.397 1.398l3.85 3.85a1 1 0 001.415-1.414l-3.868-3.834zm-5.242 1.156a5.5 5.5 0 110-11 5.5 5.5 0 010 11z"/></svg>
        Search Records
      </h3>
      <span class="sc-hint">At least <strong>Program Name</strong> or <strong>Trainer Name</strong> is required</span>
    </div>
    <div class="sc-body">
      <div class="fg">
        <label for="f-prog">Program Name <span class="req">*</span></label>
        <input type="text" id="f-prog" placeholder="Type to search…" autocomplete="off"
               oninput="acInput('prog')" onkeydown="acKey(event,'prog')" onblur="acBlur('prog')"/>
        <div class="ac-list" id="ac-prog"></div>
      </div>
      <div class="fg">
        <label for="f-train">Trainer Name <span class="req">*</span></label>
        <input type="text" id="f-train" placeholder="Type to search…" autocomplete="off"
               oninput="acInput('train')" onkeydown="acKey(event,'train')" onblur="acBlur('train')"/>
        <div class="ac-list" id="ac-train"></div>
      </div>
      <div class="sc-div"></div>
      <div class="fg">
        <label for="f-start">Start Date — From</label>
        <input type="date" id="f-start"/>
      </div>
      <div class="fg">
        <label for="f-end">End Date — To</label>
        <input type="date" id="f-end"/>
      </div>
      <div class="sc-actions">
        <button class="btn btn-primary" onclick="doSearch()">
          <svg width="13" height="13" viewBox="0 0 16 16" fill="currentColor"><path d="M11.742 10.344a6.5 6.5 0 10-1.397 1.398l3.85 3.85a1 1 0 001.415-1.414l-3.868-3.834zm-5.242 1.156a5.5 5.5 0 110-11 5.5 5.5 0 010 11z"/></svg>
          Search
        </button>
        <button class="btn btn-default" onclick="clearSearch()">Clear</button>
      </div>
    </div>
    <div class="val-banner" id="val-banner">
      <svg width="15" height="15" viewBox="0 0 16 16" fill="currentColor"><path d="M8 1a7 7 0 100 14A7 7 0 008 1zm-.75 3.5h1.5v5h-1.5v-5zm0 6.25h1.5V12h-1.5v-1.25z"/></svg>
      Please enter at least a <strong>Program Name</strong> or <strong>Trainer Name</strong> before searching.
    </div>
    <div class="date-note" id="date-note" style="display:none">
      ⚡ Date filter active — start_date ≥ From &amp; end_date ≤ To
    </div>
  </div>

  <!-- Summary cards -->
  <div class="sum-grid" id="sum-grid">
    <div class="sum-card">   <div class="slabel">Classes Found</div>       <div class="sval" id="s-cls">0</div></div>
    <div class="sum-card g"> <div class="slabel">Participants</div>        <div class="sval" id="s-part">—</div></div>
    <div class="sum-card a"> <div class="slabel">Feedback Given</div>      <div class="sval" id="s-fb">—</div></div>
    <div class="sum-card b"> <div class="slabel">Avg Q1 Score</div>        <div class="sval" id="s-q1">—</div></div>
    <div class="sum-card b"> <div class="slabel">Avg Q2 Score</div>        <div class="sval" id="s-q2">—</div></div>
    <div class="sum-card r"> <div class="slabel">With Feedback File</div>  <div class="sval" id="s-files">0</div></div>
  </div>

  <!-- Command bar -->
  <div class="cmd-bar" id="cmd-bar">
    <div class="cmd-left">
      <span class="rec-count" id="rec-count"></span>
    </div>
    <div class="cmd-right">
      <label for="ps-sel" style="font-size:13px;color:#605e5c">Rows:</label>
      <select id="ps-sel" style="height:32px;padding:0 6px;font-size:13px;border:1px solid #8a8886;border-radius:2px;font-family:inherit" onchange="changePS()">
        <option value="10">10</option><option value="25">25</option>
        <option value="50">50</option><option value="100">100</option>
      </select>
    </div>
  </div>

  <!-- Table section -->
  <div id="tbl-section">
    <div class="tbl-wrap">
      <table id="data-table">
        <thead id="t-head"></thead>
        <tbody id="t-body"></tbody>
      </table>
    </div>
    <div class="empty" id="empty-state">
      <div class="eicon">🔍</div>
      <h3>No records match your search</h3>
      <p>Try a broader program name, different trainer, or wider date range.</p>
    </div>

    <!--
      DOWNLOAD BAR
      ─────────────────────────────────────────────────────────────
      Appears below results. Has TWO download options:
        1. CSV only  — just the feedback scores spreadsheet
        2. ZIP       — CSV scores + all linked feedback files bundled
    -->
    <div class="dl-bar" id="dl-bar">
      <div class="dl-bar-info">
        <strong id="dl-bar-count">0</strong> records found &nbsp;·&nbsp;
        <span id="dl-bar-files-note" style="color:#605e5c"></span>
      </div>
      <div class="dl-bar-actions">
        <button class="btn btn-default" onclick="downloadCSV()">
          <svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor"><path d="M14 4.5V14a2 2 0 01-2 2H4a2 2 0 01-2-2V2a2 2 0 012-2h6l4 4.5zM9.5 3V1.5L13 4.5H9.5V3zM4 7h8v1H4V7zm0 2h8v1H4V9zm0 2h5v1H4v-1z"/></svg>
          Download Scores CSV
        </button>
        <button class="btn btn-success" id="btn-zip" onclick="downloadZip()">
          <svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor"><path d="M6.5 7.5a1 1 0 012 0v.5h.5a1.5 1.5 0 010 3H7a1.5 1.5 0 01-1.5-1.5v-1A1.5 1.5 0 017 7h-.5v-.5zM5 1h1v1H5V1zm1 1h1v1H6V2zM5 3h1v1H5V3zm1 1h1v1H6V4zM5 5h1v1H5V5zM9 1a5 5 0 015 5v7a2 2 0 01-2 2H4a2 2 0 01-2-2V6a5 5 0 015-5h2z"/></svg>
          Download ZIP (Scores + Feedback Files)
        </button>
      </div>
    </div>

    <div class="pg-bar" id="pg-bar">
      <span class="pg-info" id="pg-info"></span>
      <div class="pg-btns" id="pg-btns"></div>
    </div>
  </div>

</div><!-- /.main -->

<!-- ZIP PROGRESS OVERLAY -->
<div class="zip-progress" id="zip-progress">
  <div class="zip-box">
    <div class="spinner green" style="margin-bottom:14px"></div>
    <h3>Building ZIP file…</h3>
    <p id="zip-status">Fetching feedback files</p>
    <p id="zip-substatus" style="font-size:12px;color:#a19f9d;margin-top:4px"></p>
    <div class="zip-pbar-wrap"><div class="zip-pbar" id="zip-pbar"></div></div>
  </div>
</div>

<!-- FEEDBACK FILE PREVIEW MODAL -->
<div class="modal-bg" id="modal-bg" onclick="closeMod(event)">
  <div class="modal-box" onclick="event.stopPropagation()">
    <div class="mh">
      <div class="mh-left">
        <div class="mh-title" id="mod-title">Feedback File</div>
        <div class="mh-sub"   id="mod-sub"></div>
      </div>
      <div class="mh-actions">
        <button class="btn btn-default btn-sm" onclick="dlFbFromModal()">
          <svg width="12" height="12" viewBox="0 0 16 16" fill="currentColor"><path d="M8 10.5l-4-4h2.5V2h3v4.5H12L8 10.5zM3 13h10v1.5H3V13z"/></svg>
          Download this file
        </button>
        <button class="btn btn-primary btn-sm" id="mod-newtab" onclick="openNewTab()" style="display:none">Open in new tab ↗</button>
        <button class="m-close" onclick="closeMod(null)">×</button>
      </div>
    </div>
    <div class="mb" id="mod-body">
      <div class="fv-spin-wrap"><div class="spinner"></div><p>Loading…</p></div>
    </div>
  </div>
</div>

<script>
/* ════════════════════════════════════════════════════════════
   ★  CONFIGURATION  — edit ONLY this block
   ════════════════════════════════════════════════════════════
   LOCAL TESTING:
     excelPath      → filename only, e.g. 'FeedbackMaster.xlsx'
     feedbackFolder → relative subfolder, e.g. 'Feedback/'
   ON SHAREPOINT:
     excelPath      → '/sites/HR/Shared Documents/Training/FeedbackMaster.xlsx'
     feedbackFolder → '/sites/HR/Shared Documents/Feedback/'
════════════════════════════════════════════════════════════ */
var CONFIG = {
  excelPath:      'Feedbackdata.xlsx',
  feedbackFolder: 'final/',
  displayName:    'FeedbackMaster.xlsx'
};
/* ═══════════════════════════════════════════════════════════ */

/* ── STATE ── */
var allData      = [];
var filteredData = [];
var rawCols      = [];
var COL          = {};
var currentPage  = 1;
var pageSize     = 10;
var sortKey      = '';
var sortDir      = 1;
var acFocus      = {prog:-1, train:-1};
var currentFbPath= '';
var currentFbBlob= null;
/* _fbMap: random id → resolved file path, populated when table rows render */
window._fbMap    = {};

/* ── UTILITIES ── */
function $e(id){ return document.getElementById(id); }
function val(id){ var e=$e(id); return e?e.value.trim():''; }
function esc(s){
  return String(s==null?'':s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function pad(n){ return String(n).padStart(2,'0'); }
function parseDate(s){
  if(!s&&s!==0) return null;
  if(typeof s==='number'){ var d=new Date(Math.round((s-25569)*86400000)); return isNaN(d)?null:d; }
  var d=new Date(s); return isNaN(d)?null:d;
}
function fmtDate(raw){
  var d=parseDate(raw); if(!d) return raw||'';
  return pad(d.getDate())+'/'+pad(d.getMonth()+1)+'/'+d.getFullYear();
}
function resolvePath(p){
  /* Convert a raw value from feedback_file_location into a fetchable URL.
     If it's just a filename with no slashes, prepend feedbackFolder.      */
  if(!p) return '';
  p=String(p).trim();
  if(/^https?:\/\//i.test(p)) return p;          // absolute URL — use as-is
  if(/^\\\\/.test(p))          return p;          // UNC — flag later
  if(!/[\/\\]/.test(p))        return CONFIG.feedbackFolder+p; // bare filename
  return p;                                        // already has path
}

/* ── COLUMN ALIASES ── */
var ALIASES={
  classid:                ['classid','class_id','class id','id'],
  program_name:           ['program_name','program name','programname','program'],
  trainer_name:           ['trainer_name','trainer name','trainername','trainer'],
  participant_count:      ['participant_count','participant count','participants','no of participants','count'],
  fb_given:               ['fb_given','fb given','feedback given','feedbackgiven','feedback count'],
  avg_q1:                 ['avg_q1','avg q1','avgq1','average q1','q1','avg q1 score'],
  avg_q2:                 ['avg_q2','avg q2','avgq2','average q2','q2','avg q2 score'],
  feedback_file_location: ['feedback_file_location','feedback file location','file location','feedbackfile','feedback file','file path','filepath','feedback path'],
  start_date:             ['start_date','start date','startdate','program start','program_start','from date'],
  end_date:               ['end_date','end date','enddate','program end','program_end','to date'],
};
function buildColMap(keys){
  COL={};
  var lk=keys.map(function(k){return{orig:k,low:k.toLowerCase().replace(/\s+/g,' ').trim()};});
  Object.keys(ALIASES).forEach(function(std){
    var al=ALIASES[std];
    for(var i=0;i<lk.length;i++){
      if(al.indexOf(lk[i].low)!==-1){COL[std]=lk[i].orig;break;}
    }
  });
}
function gc(row,std){
  if(!COL[std]) return '';
  var v=row[COL[std]]; return v===undefined?'':v;
}

/* ── LOAD EXCEL FROM SHAREPOINT / LOCAL SERVER ── */
function loadExcelData(){
  $e('loading-panel').style.display='block';
  $e('err-box').style.display='none';
  $e('data-bar').style.display='none';
  $e('search-card').style.display='none';
  $e('sum-grid').style.display='none';
  $e('tbl-section').style.display='none';
  $e('cmd-bar').style.display='none';
  $e('dl-bar').style.display='none';
  $e('btn-reload').style.display='none';
  $e('loading-path-label').textContent=CONFIG.excelPath;

  fetch(CONFIG.excelPath,{credentials:'same-origin',cache:'no-cache'})
    .then(function(r){
      if(!r.ok) throw new Error('HTTP '+r.status+' — '+r.statusText);
      return r.arrayBuffer();
    })
    .then(function(buf){ parseExcel(buf); })
    .catch(function(err){ showLoadError(err.message); });
}

function parseExcel(buf){
  try{
    var wb=XLSX.read(buf,{type:'array',cellDates:true,dateNF:'yyyy-mm-dd'});
    var ws=wb.Sheets[wb.SheetNames[0]];
    var rows=XLSX.utils.sheet_to_json(ws,{raw:false,dateNF:'yyyy-mm-dd'});
    if(!rows.length){ showLoadError('The first sheet has no data rows.'); return; }

    rawCols=Object.keys(rows[0]);
    buildColMap(rawCols);
    allData=rows; filteredData=[];

    $e('loading-panel').style.display='none';
    $e('data-bar').style.display='flex';
    $e('search-card').style.display='block';
    $e('btn-reload').style.display='inline-flex';
    $e('data-fname').textContent=CONFIG.displayName||CONFIG.excelPath;
    $e('data-fmeta').textContent=rows.length+' rows · '+rawCols.length+' columns';

    var miss=[];
    if(!COL.program_name) miss.push('program_name');
    if(!COL.trainer_name)  miss.push('trainer_name');
    if(miss.length){
      var w=document.createElement('div');
      w.style.cssText='background:#fff4ce;border:1px solid #c7a614;border-radius:2px;padding:9px 14px;margin:12px 18px;font-size:13px;color:#463203';
      w.innerHTML='⚠ Column(s) not found in Excel: <strong>'+miss.join(', ')+'</strong>. Check your column headers.';
      $e('search-card').appendChild(w);
    }
  }catch(err){ showLoadError('Parse error — '+err.message); }
}

function showLoadError(msg){
  $e('loading-panel').style.display='none';
  $e('err-box').style.display='block';
  $e('err-msg').textContent=msg;
  $e('err-path').textContent=CONFIG.excelPath;
}

/* ── AUTOCOMPLETE ── */
var AC={
  prog: {inp:'f-prog', list:'ac-prog', std:'program_name'},
  train:{inp:'f-train',list:'ac-train',std:'trainer_name'}
};
function acInput(k){
  liveValidate();
  var c=AC[k],q=val(c.inp).toLowerCase();
  if(!q){closeAc(k);return;}
  var set={};
  allData.forEach(function(r){var v=String(gc(r,c.std)||'').trim();if(v)set[v]=true;});
  var m=Object.keys(set).filter(function(v){return v.toLowerCase().indexOf(q)!==-1;}).sort().slice(0,14);
  if(!m.length){closeAc(k);return;}
  acFocus[k]=-1;
  $e(c.list).innerHTML=m.map(function(v){
    return '<div class="ac-item" data-v="'+esc(v)+'" onmousedown="acPick(event,\''+k+'\',this.dataset.v)">'+esc(v)+'</div>';
  }).join('');
  $e(c.list).classList.add('open');
}
function acKey(e,k){
  var items=$e(AC[k].list).querySelectorAll('.ac-item');
  if(!items.length) return;
  if(e.key==='ArrowDown'){e.preventDefault();acFocus[k]=Math.min(acFocus[k]+1,items.length-1);acHL(k,items);}
  else if(e.key==='ArrowUp'){e.preventDefault();acFocus[k]=Math.max(acFocus[k]-1,0);acHL(k,items);}
  else if(e.key==='Enter'&&acFocus[k]>=0){e.preventDefault();acPick(null,k,items[acFocus[k]].dataset.v);}
  else if(e.key==='Escape'){closeAc(k);}
}
function acHL(k,items){items.forEach(function(el,i){el.classList.toggle('hi',i===acFocus[k]);});}
function acPick(e,k,v){if(e)e.preventDefault();$e(AC[k].inp).value=v;closeAc(k);liveValidate();}
function acBlur(k){setTimeout(function(){closeAc(k);},160);}
function closeAc(k){$e(AC[k].list).classList.remove('open');acFocus[k]=-1;}

/* ── VALIDATION ── */
function liveValidate(){
  var ok=!!(val('f-prog')||val('f-train'));
  $e('val-banner').classList.toggle('show',!ok);
  $e('f-prog').classList.toggle('has-error',!ok);
  $e('f-train').classList.toggle('has-error',!ok);
  return ok;
}
function validate(){
  if(val('f-prog')||val('f-train')){
    $e('val-banner').classList.remove('show');
    $e('f-prog').classList.remove('has-error');
    $e('f-train').classList.remove('has-error');
    return true;
  }
  $e('val-banner').classList.add('show');
  $e('f-prog').classList.add('has-error');
  $e('f-train').classList.add('has-error');
  $e('f-prog').focus();
  return false;
}

/* ── SEARCH ── */
function doSearch(){
  if(!validate()) return;
  var prog=val('f-prog').toLowerCase();
  var train=val('f-train').toLowerCase();
  var sf=val('f-start'), ef=val('f-end');
  $e('date-note').style.display=(sf||ef)?'block':'none';

  filteredData=allData.filter(function(row){
    var pn=String(gc(row,'program_name')||'').toLowerCase();
    var tn=String(gc(row,'trainer_name')||'').toLowerCase();
    if(prog  && pn.indexOf(prog)===-1)  return false;
    if(train && tn.indexOf(train)===-1) return false;
    if(sf||ef){
      var sd=parseDate(gc(row,'start_date'));
      var ed=parseDate(gc(row,'end_date'));
      if(sf && sd && sd<new Date(sf)) return false;
      if(ef && ed && ed>new Date(ef)) return false;
    }
    return true;
  });

  applySortToFiltered();
  currentPage=1;
  window._fbMap={};   /* reset feedback path map on each new search */

  $e('tbl-section').style.display='block';
  $e('cmd-bar').style.display='flex';

  renderSummary();
  renderTable();
  updateDlBar();
}

function clearSearch(){
  ['f-prog','f-train','f-start','f-end'].forEach(function(id){var e=$e(id);if(e)e.value='';});
  $e('val-banner').classList.remove('show');
  $e('f-prog').classList.remove('has-error');
  $e('f-train').classList.remove('has-error');
  $e('date-note').style.display='none';
  filteredData=[];
  $e('tbl-section').style.display='none';
  $e('cmd-bar').style.display='none';
  $e('sum-grid').style.display='none';
  $e('dl-bar').style.display='none';
  sortKey='';sortDir=1;
  window._fbMap={};
}

/* ── DOWNLOAD BAR ── */
function updateDlBar(){
  var has=filteredData.length>0;
  $e('dl-bar').style.display=has?'flex':'none';
  if(!has) return;
  var filesCount=filteredData.filter(function(r){
    return !!String(gc(r,'feedback_file_location')||'').trim();
  }).length;
  $e('dl-bar-count').textContent=filteredData.length;
  $e('dl-bar-files-note').textContent=filesCount
    ? filesCount+' feedback file'+(filesCount!==1?'s':'')+' attached'
    : 'No feedback files attached';
  /* disable ZIP button if no feedback files */
  $e('btn-zip').disabled=filesCount===0;
}

/* ── SUMMARY ── */
function renderSummary(){
  var parts=0,pn=0,fb=0,fn=0,q1=0,q1n=0,q2=0,q2n=0,files=0;
  filteredData.forEach(function(r){
    var p=parseFloat(gc(r,'participant_count'));if(!isNaN(p)){parts+=p;pn++;}
    var f=parseFloat(gc(r,'fb_given'));         if(!isNaN(f)){fb+=f;fn++;}
    var a=parseFloat(gc(r,'avg_q1'));           if(!isNaN(a)){q1+=a;q1n++;}
    var b=parseFloat(gc(r,'avg_q2'));           if(!isNaN(b)){q2+=b;q2n++;}
    if(String(gc(r,'feedback_file_location')||'').trim()) files++;
  });
  $e('s-cls').textContent =filteredData.length;
  $e('s-part').textContent=pn ?parts.toLocaleString():'—';
  $e('s-fb').textContent  =fn ?fb.toLocaleString()   :'—';
  $e('s-q1').textContent  =q1n?(q1/q1n).toFixed(1)  :'—';
  $e('s-q2').textContent  =q2n?(q2/q2n).toFixed(1)  :'—';
  $e('s-files').textContent=files;
  $e('sum-grid').style.display=filteredData.length?'grid':'none';
}

/* ── TABLE ── */
var DCOLS=[
  {std:'classid',               label:'Class ID',        sort:true},
  {std:'program_name',          label:'Program Name',    sort:true},
  {std:'trainer_name',          label:'Trainer',         sort:true},
  {std:'start_date',            label:'Start Date',      sort:true},
  {std:'end_date',              label:'End Date',        sort:true},
  {std:'participant_count',     label:'Participants',    sort:true},
  {std:'fb_given',              label:'Feedback Given',  sort:true},
  {std:'avg_q1',                label:'Avg Q1',          sort:true},
  {std:'avg_q2',                label:'Avg Q2',          sort:true},
  {std:'feedback_file_location',label:'Feedback File',   sort:false},
];

function renderTable(){
  var shown=DCOLS.filter(function(d){return !!COL[d.std];});
  var mapped=Object.values(COL);
  var extras=rawCols.filter(function(c){return mapped.indexOf(c)===-1;});

  /* HEAD */
  var hh='<tr>';
  shown.forEach(function(d){
    var ok=COL[d.std];
    var sc=(sortKey===ok)?(sortDir===1?'asc':'dsc'):'';
    hh+='<th class="'+sc+'"'+(d.sort?' onclick="sortBy(\''+esc(ok)+'\')"':'')+'>'+
        esc(d.label)+(d.sort?'<span class="sarr"></span>':'')+'</th>';
  });
  extras.forEach(function(c){
    var sc=(sortKey===c)?(sortDir===1?'asc':'dsc'):'';
    hh+='<th class="'+sc+'" onclick="sortBy(\''+esc(c)+'\')">'+esc(c)+'<span class="sarr"></span></th>';
  });
  hh+='</tr>';
  $e('t-head').innerHTML=hh;

  if(!filteredData.length){
    $e('t-body').innerHTML='';
    $e('empty-state').style.display='block';
    $e('pg-bar').style.display='none';
    $e('rec-count').innerHTML='No records found';
    return;
  }
  $e('empty-state').style.display='none';

  var start=(currentPage-1)*pageSize;
  var page=filteredData.slice(start,start+pageSize);
  var bh='';
  page.forEach(function(row){
    bh+='<tr>';
    shown.forEach(function(d){bh+='<td>'+renderCell(d.std,gc(row,d.std))+'</td>';});
    extras.forEach(function(c){bh+='<td>'+esc(row[c]||'')+'</td>';});
    bh+='</tr>';
  });
  $e('t-body').innerHTML=bh;

  var s=start+1,e2=Math.min(start+pageSize,filteredData.length),tot=filteredData.length;
  $e('rec-count').innerHTML='Showing <strong>'+s+'–'+e2+'</strong> of <strong>'+tot+'</strong> record'+(tot!==1?'s':'');
  renderPagination(tot);
}

function renderCell(std,raw){
  var v=(raw==null)?'':raw;

  if(std==='start_date'||std==='end_date')
    return v?esc(fmtDate(v)):'<span style="color:#a19f9d">—</span>';

  if(std==='avg_q1'||std==='avg_q2'){
    var n=parseFloat(v); if(isNaN(n)) return v?esc(v):'<span style="color:#a19f9d">—</span>';
    var pct=Math.round(Math.min(5,Math.max(0,n))/5*100);
    var cls=n>=4?'good':n>=2.5?'mid':'low';
    return '<div class="rbar"><div class="rfill '+cls+'" style="width:'+pct+'%"></div><span class="rnum">'+n.toFixed(1)+'</span></div>';
  }

  if(std==='feedback_file_location'){
    var rawPath=String(v||'').trim();
    if(!rawPath) return '<span class="fb-na">—</span>';

    var resolvedPath=resolvePath(rawPath);
    var fname=rawPath.split(/[\/\\]/).pop()||rawPath;

    /* store resolved path in _fbMap under a random key */
    var eid='fb_'+Math.random().toString(36).slice(2,9);
    window._fbMap[eid]=resolvedPath;

    /*
      HYPERLINK  — clicking the filename opens the preview modal
      PREVIEW    — small button also opens preview modal
      The text is a proper clickable link styled with .fb-link
    */
    return '<div class="fb-cell">'
      /* main hyperlink — the filename is the clickable text */
      +'<button class="fb-link" onclick="openFb(\''+eid+'\')" title="Preview: '+esc(resolvedPath)+'">'
      +'<svg width="13" height="13" viewBox="0 0 16 16" fill="currentColor" style="flex-shrink:0">'
      +'<path d="M14 4.5V14a2 2 0 01-2 2H4a2 2 0 01-2-2V2a2 2 0 012-2h6l4 4.5z"/>'
      +'<path d="M9.5 3V1.5L13 4.5H9.5V3zM4 7h8v1H4V7zm0 2h8v1H4V9zm0 2h5v1H4v-1z" fill="#fff"/>'
      +'</svg>'
      +esc(fname)
      +'</button>'
      /* small preview label */
      +'<button class="fb-preview-btn" onclick="openFb(\''+eid+'\')" title="Preview file">👁 Preview</button>'
      +'</div>';
  }

  if(std==='participant_count'||std==='fb_given'){
    var n=parseFloat(v); return isNaN(n)?esc(v)||'<span style="color:#a19f9d">—</span>':n.toLocaleString();
  }
  return v!==''?esc(v):'<span style="color:#a19f9d">—</span>';
}

/* ── SORT ── */
function sortBy(col){
  if(sortKey===col)sortDir*=-1;else{sortKey=col;sortDir=1;}
  applySortToFiltered();currentPage=1;renderTable();
}
function applySortToFiltered(){
  if(!sortKey) return;
  filteredData.sort(function(a,b){
    var av=a[sortKey]||'',bv=b[sortKey]||'';
    var an=parseFloat(av),bn=parseFloat(bv);
    if(!isNaN(an)&&!isNaN(bn)) return(an-bn)*sortDir;
    var ad=parseDate(av),bd=parseDate(bv);
    if(ad&&bd) return(ad-bd)*sortDir;
    return String(av).localeCompare(String(bv))*sortDir;
  });
}

/* ── PAGINATION ── */
function renderPagination(total){
  var pb=$e('pg-bar'),pages=Math.ceil(total/pageSize);
  if(pages<=1){pb.style.display='none';return;}
  pb.style.display='flex';
  $e('pg-info').textContent='Page '+currentPage+' of '+pages;
  var lo=Math.max(1,currentPage-2),hi=Math.min(pages,currentPage+2);
  var b='<button class="pgb" onclick="goPage('+(currentPage-1)+')" '+(currentPage===1?'disabled':'')+'>◀</button>';
  if(lo>1){b+=pgbtn(1);if(lo>2)b+='<span style="padding:0 4px">…</span>';}
  for(var i=lo;i<=hi;i++)b+=pgbtn(i);
  if(hi<pages){if(hi<pages-1)b+='<span style="padding:0 4px">…</span>';b+=pgbtn(pages);}
  b+='<button class="pgb" onclick="goPage('+(currentPage+1)+')" '+(currentPage===pages?'disabled':'')+'>▶</button>';
  $e('pg-btns').innerHTML=b;
}
function pgbtn(p){return '<button class="pgb'+(p===currentPage?' on':'')+'" onclick="goPage('+p+')">'+p+'</button>';}
function goPage(p){var pages=Math.ceil(filteredData.length/pageSize);if(p<1||p>pages)return;currentPage=p;renderTable();window.scrollTo({top:0,behavior:'smooth'});}
function changePS(){pageSize=parseInt($e('ps-sel').value,10);currentPage=1;renderTable();}

/* ══════════════════════════════════════════════════════════
   DOWNLOAD — CSV SCORES ONLY
══════════════════════════════════════════════════════════ */
function buildCSVBlob(){
  var lines=[rawCols.map(function(c){return '"'+c+'"';}).join(',')];
  filteredData.forEach(function(row){
    lines.push(rawCols.map(function(c){
      return '"'+String(row[c]!=null?row[c]:'').replace(/"/g,'""')+'"';
    }).join(','));
  });
  return new Blob(['\ufeff'+lines.join('\r\n')],{type:'text/csv;charset=utf-8;'});
}

function csvFileName(){
  var p=(val('f-prog')||'all').replace(/\s+/g,'_');
  var t=(val('f-train')||'all').replace(/\s+/g,'_');
  return 'feedback_scores_'+p+'_'+t+'_'+new Date().toISOString().slice(0,10)+'.csv';
}

function downloadCSV(){
  if(!filteredData.length){alert('No records to export.');return;}
  triggerDL(buildCSVBlob(), csvFileName());
}

/* ══════════════════════════════════════════════════════════
   DOWNLOAD — ZIP  (CSV scores + all feedback files)
   ──────────────────────────────────────────────────────────
   1. Build the CSV blob (scores)
   2. Collect unique feedback file paths from filteredData
   3. fetch() each file (same-origin credentials)
   4. Bundle everything into a ZIP using JSZip
   5. Trigger download of the ZIP
══════════════════════════════════════════════════════════ */
function downloadZip(){
  if(!filteredData.length){alert('No records to download.');return;}
  if(typeof JSZip==='undefined'){
    alert('JSZip library failed to load. Check your internet connection and retry.');
    return;
  }

  /* collect unique feedback file paths */
  var pathSet={};
  filteredData.forEach(function(row){
    var raw=String(gc(row,'feedback_file_location')||'').trim();
    if(raw){ var rp=resolvePath(raw); if(rp) pathSet[rp]=true; }
  });
  var paths=Object.keys(pathSet);

  /* show progress overlay */
  $e('zip-progress').classList.add('open');
  $e('zip-status').textContent='Preparing CSV…';
  $e('zip-substatus').textContent='';
  $e('zip-pbar').style.width='0%';

  var zip=new JSZip();

  /* add CSV */
  zip.file(csvFileName(), buildCSVBlob());

  if(!paths.length){
    /* no feedback files — just zip the CSV */
    $e('zip-status').textContent='Generating ZIP…';
    zip.generateAsync({type:'blob'}).then(function(blob){
      $e('zip-progress').classList.remove('open');
      triggerDL(blob, zipFileName());
    });
    return;
  }

  /* fetch each feedback file */
  var total=paths.length, done=0, failed=[];

  $e('zip-status').textContent='Fetching '+total+' feedback file'+(total!==1?'s':'')+'…';

  var folder=zip.folder('FeedbackFiles');

  var promises=paths.map(function(path){
    /* skip UNC paths — can't fetch from browser */
    if(/^\\\\/.test(path)){
      failed.push(path);
      done++;
      updateZipProgress(done,total);
      return Promise.resolve();
    }
    return fetch(path,{credentials:'same-origin'})
      .then(function(r){
        if(!r.ok) throw new Error('HTTP '+r.status);
        return r.arrayBuffer();
      })
      .then(function(buf){
        var fname=path.split(/[\/\\]/).pop()||('file_'+done);
        folder.file(fname, buf);
        done++;
        $e('zip-substatus').textContent='Fetched: '+fname;
        updateZipProgress(done,total);
      })
      .catch(function(err){
        failed.push(path+' ('+err.message+')');
        done++;
        updateZipProgress(done,total);
      });
  });

  Promise.all(promises).then(function(){
    $e('zip-status').textContent='Building ZIP file…';
    $e('zip-substatus').textContent='';
    $e('zip-pbar').style.width='95%';

    /* if some files failed, add a notes file */
    if(failed.length){
      var note='The following feedback files could not be fetched and are not included in this ZIP:\n\n'+
               failed.join('\n')+
               '\n\nMake sure the paths in the feedback_file_location column are valid server-relative URLs accessible from this page.';
      zip.file('_MISSING_FILES.txt', note);
    }

    return zip.generateAsync({type:'blob', compression:'DEFLATE', compressionOptions:{level:6}},
      function(meta){ $e('zip-pbar').style.width=Math.round(95+meta.percent*0.05)+'%'; });
  }).then(function(blob){
    $e('zip-progress').classList.remove('open');
    $e('zip-pbar').style.width='0%';
    triggerDL(blob, zipFileName());
  }).catch(function(err){
    $e('zip-progress').classList.remove('open');
    alert('ZIP generation failed: '+err.message);
  });
}

function updateZipProgress(done,total){
  var pct=Math.round((done/total)*90);
  $e('zip-pbar').style.width=pct+'%';
  $e('zip-status').textContent='Fetching files… ('+done+'/'+total+')';
}

function zipFileName(){
  var p=(val('f-prog')||'all').replace(/\s+/g,'_');
  var t=(val('f-train')||'all').replace(/\s+/g,'_');
  return 'training_feedback_'+p+'_'+t+'_'+new Date().toISOString().slice(0,10)+'.zip';
}

function triggerDL(blob,name){
  var url=URL.createObjectURL(blob),a=document.createElement('a');
  a.href=url;a.download=name;document.body.appendChild(a);a.click();
  document.body.removeChild(a);URL.revokeObjectURL(url);
}

/* ══════════════════════════════════════════════════════════
   FEEDBACK FILE PREVIEW MODAL
   Supported: .pdf → iframe  |  .xlsx/.xls/.csv → table
              .txt/.log/.json → pre  |  .docx/.pptx → info
══════════════════════════════════════════════════════════ */
function openFb(eid){
  var path=(window._fbMap&&window._fbMap[eid])||'';
  if(!path) return;
  currentFbPath=path; currentFbBlob=null;
  var fname=path.split(/[\/\\]/).pop()||path;
  var ext=(fname.split('.').pop()||'').toLowerCase();
  $e('mod-title').textContent=fname;
  $e('mod-sub').textContent=path;
  $e('mod-newtab').style.display='none';
  $e('mod-body').innerHTML='<div class="fv-spin-wrap"><div class="spinner"></div><p>Loading…</p></div>';
  $e('modal-bg').classList.add('open');

  if(/^\\\\/.test(path)){ showUncMsg(path); return; }

  fetch(path,{credentials:'same-origin'})
    .then(function(r){if(!r.ok)throw new Error('HTTP '+r.status+' — '+r.statusText);return r.arrayBuffer();})
    .then(function(buf){currentFbBlob=new Blob([buf]);renderFbContent(buf,ext,fname);})
    .catch(function(err){showFetchErr(path,err.message);});
}

function renderFbContent(buf,ext,fname){
  var body=$e('mod-body');
  if(ext==='pdf'){
    var ou=URL.createObjectURL(new Blob([buf],{type:'application/pdf'}));
    body.innerHTML='<div class="fv-pdf"><iframe src="'+ou+'#toolbar=1" title="'+esc(fname)+'"></iframe></div>';
    return;
  }
  if(ext==='xlsx'||ext==='xls'||ext==='csv'){
    try{
      var wb=XLSX.read(ext==='csv'?new TextDecoder().decode(buf):buf,
                       {type:ext==='csv'?'string':'array',cellDates:true,dateNF:'yyyy-mm-dd'});
      var ws=wb.Sheets[wb.SheetNames[0]];
      var rows=XLSX.utils.sheet_to_json(ws,{raw:false,header:1});
      if(!rows.length){body.innerHTML='<p style="color:#605e5c;padding:16px">File is empty.</p>';return;}
      var hdrs=rows[0],dr=rows.slice(1);
      body.innerHTML='<div class="fv-tbl-wrap"><table><thead><tr>'+
        hdrs.map(function(h){return '<th>'+esc(h==null?'':h)+'</th>';}).join('')+
        '</tr></thead><tbody>'+
        dr.map(function(r){return '<tr>'+hdrs.map(function(_,i){return '<td>'+esc(r[i]==null?'':r[i])+'</td>';}).join('')+'</tr>';}).join('')+
        '</tbody></table></div>'+
        '<p style="font-size:12px;color:#605e5c;margin-top:8px">'+dr.length+' rows · '+hdrs.length+' columns</p>';
    }catch(e){body.innerHTML='<div class="fv-err">Could not parse: '+esc(e.message)+'</div>';}
    return;
  }
  if(ext==='txt'||ext==='log'||ext==='json'||ext==='xml'){
    body.innerHTML='<pre style="font-size:13px;white-space:pre-wrap;word-break:break-all;line-height:1.5;color:#323130">'+
      esc(new TextDecoder().decode(buf))+'</pre>';
    return;
  }
  if(ext==='docx'||ext==='pptx'){
    $e('mod-newtab').style.display='inline-flex';
    body.innerHTML='<div class="fv-info">In-browser preview not available for <strong>.'+esc(ext)+'</strong> files.<br>Use <strong>Download this file</strong> or <strong>Open in new tab</strong>.</div>';
    return;
  }
  body.innerHTML='<div style="text-align:center;padding:40px 20px;color:#605e5c">'+
    '<p style="margin-bottom:14px">Preview not available for <strong>.'+esc(ext)+'</strong> files.</p>'+
    '<button class="btn btn-primary" onclick="dlFbFromModal()">Download file</button></div>';
}

function showUncMsg(path){
  $e('mod-body').innerHTML='<div class="fv-err">UNC paths cannot be opened from a browser:<br><code>'+esc(path)+'</code></div>'+
    '<div class="fv-info">Move this file into a SharePoint document library and update the <em>feedback_file_location</em> column with a relative path like <code>Feedback/filename.pdf</code></div>';
}
function showFetchErr(path,msg){
  $e('mod-newtab').style.display='inline-flex';
  $e('mod-body').innerHTML='<div class="fv-err">Could not load: '+esc(msg)+'<br><code>'+esc(path)+'</code></div>'+
    '<div class="fv-info">Make sure the file exists at that path or click <strong>Open in new tab ↗</strong>.</div>';
}
function dlFbFromModal(){
  if(currentFbBlob){ triggerDL(currentFbBlob,(currentFbPath.split(/[\/\\]/).pop()||'feedback_file')); }
  else if(currentFbPath){ window.open(currentFbPath,'_blank'); }
}
function openNewTab(){if(currentFbPath)window.open(currentFbPath,'_blank');}
function closeMod(e){if(e===null||e.target===$e('modal-bg'))$e('modal-bg').classList.remove('open');}
document.addEventListener('keydown',function(e){if(e.key==='Escape')$e('modal-bg').classList.remove('open');});

/* ── AUTO-LOAD ── */
document.addEventListener('DOMContentLoaded', loadExcelData);
</script>
</body>
</html>
