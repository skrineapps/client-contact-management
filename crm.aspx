<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<!--
  Skrine Client Contact Dashboard
  ================================
  Deployment: upload this file as-is into a SharePoint Document Library and open it
  directly via its file URL (not through a list view). It is a fully static
  HTML/CSS/JS page - no server code runs, so SharePoint Online just streams the
  bytes to the browser.

  Note on text encoding: this file intentionally avoids every non-ASCII character
  (em dashes, curly quotes, etc.) in anything that gets rendered on the page. Some
  SharePoint document libraries serve files with an HTTP Content-Type charset that
  does not match this file's actual UTF-8 bytes, and that HTTP header overrides
  this file's own <meta charset> tag by spec - which is what turned dashes into
  mangled "a-euro" style characters even after the meta tag was correct. Plain
  ASCII text renders identically under any of those charsets, so don't reintroduce
  typographic dashes/quotes/bullets into strings that reach the DOM.

  Before this will work:
  1. In the Azure AD app registration (client ID b679e920-ad66-4f7a-95d7-9bb252c0a1d5),
     under Authentication, add this file's exact URL as a redirect URI under the
     "Single-page application" platform (not "Web").
  2. Confirm the site allows custom script execution for uploaded .aspx files
     (classic/communication sites generally do; modern group-connected team sites
     may block script execution - if the page renders blank/inert, this is why).
  3. Adjust CONFIG / FIELD_ALIASES below if your SharePoint column internal names
     differ from what's assumed here - open a contact's "Raw fields (debug)"
     panel in the UI to see the exact field keys Graph is returning.
-->
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Skrine - Client Contacts</title>
<link rel="icon" type="image/png" href="https://raw.githubusercontent.com/skrineapps/skr-background-assets/main/phone-guide.png">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Plus+Jakarta+Sans:wght@700;800&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css">
<script src="https://cdn.jsdelivr.net/npm/@azure/msal-browser@3/lib/msal-browser.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
  :root{
    --red:#c51f34;
    --red-wash:#fbeaec;
    --page:#f6f5f3;
    --card:#ffffff;
    --ink:#191817;
    --muted:#78756f;
    --faint:#a8a49c;
    --line:#eae8e3;
    --amber:#c9861a;
    --amber-wash:#fbf1e2;
    --green:#2f7a4f;
    --green-wash:#e9f5ee;
    --font-display:'Plus Jakarta Sans',sans-serif;
  }
  @media (prefers-color-scheme: dark) {
    :root{
      --page:#141311;
      --card:#1e1c19;
      --ink:#f3f1ed;
      --muted:#b4afa5;
      --faint:#7d786f;
      --line:#332f2a;
      --red-wash:rgba(197,31,52,0.16);
      --amber-wash:rgba(201,134,26,0.16);
      --green-wash:rgba(47,122,79,0.16);
    }
  }
  *{box-sizing:border-box;}
  body{
    margin:0;
    background-color:var(--page);
    font-family:'Inter',sans-serif;color:var(--ink);-webkit-font-smoothing:antialiased;
  }
  .hidden{display:none !important;}
  button,input,select,textarea{font-family:inherit;}
  button:focus-visible,input:focus-visible,a:focus-visible{outline:2px solid var(--red);outline-offset:2px;}
  button:active{transform:scale(0.97);}

  @keyframes fadeInUp{from{opacity:0;transform:translateY(6px);}to{opacity:1;transform:translateY(0);}}

  .top-accent{height:3px;background:var(--red);position:relative;overflow:hidden;}
  .top-accent.loading::after{content:"";position:absolute;top:0;left:-40%;width:40%;height:100%;background:rgba(255,255,255,0.75);animation:topProgressSweep 1.1s ease-in-out infinite;}
  @keyframes topProgressSweep{0%{left:-40%;}100%{left:140%;}}
  .app-header{background:var(--card);border-bottom:1px solid var(--line);padding:15px 36px;display:flex;justify-content:space-between;align-items:center;position:sticky;top:0;z-index:10;}
  .brand-block{display:flex;align-items:center;gap:12px;}
  .logo{height:26px;width:auto;display:block;}
  .brand-text{display:flex;flex-direction:column;line-height:1.25;}
  .brand-title{font-family:var(--font-display);font-weight:800;font-size:16px;letter-spacing:-0.01em;color:var(--ink);}
  .brand-eyebrow{font-size:10.5px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--red);}
  .navbar-user{position:relative;display:flex;align-items:center;}
  .user-trigger{display:flex;align-items:center;gap:11px;background:none;border:none;padding:5px;border-radius:12px;cursor:pointer;transition:background .15s ease;}
  .user-trigger:hover{background:var(--page);}
  .user-trigger:hover .avatar,.user-trigger[aria-expanded="true"] .avatar{box-shadow:0 0 0 3px var(--red-wash);}
  .navbar-user-text{display:flex;flex-direction:column;align-items:flex-end;line-height:1.35;}
  .navbar-user-name{font-size:13px;font-weight:600;color:var(--ink);}
  .navbar-user-email{font-size:11.5px;color:var(--muted);}
  .avatar{width:36px;height:36px;border-radius:50%;background:var(--line);display:flex;align-items:center;justify-content:center;font-size:12.5px;font-weight:600;color:var(--muted);border:1px solid var(--line);overflow:hidden;flex:0 0 auto;transition:box-shadow .15s ease;}
  .avatar img{width:100%;height:100%;object-fit:cover;display:block;}
  .user-menu{position:absolute;top:calc(100% + 10px);right:0;background:var(--card);border:1px solid var(--line);border-radius:12px;box-shadow:0 14px 32px rgba(20,18,15,0.14);padding:6px;min-width:170px;z-index:20;}
  .user-menu.hidden{display:none;}
  .user-menu-item{display:flex;align-items:center;gap:9px;width:100%;background:none;border:none;padding:9px 11px;border-radius:8px;font-size:13px;color:var(--ink);cursor:pointer;text-align:left;}
  .user-menu-item i{width:14px;text-align:center;color:var(--muted);}
  .user-menu-item:hover{background:var(--red-wash);color:var(--red);}
  .user-menu-item:hover i{color:var(--red);}

  .page{max-width:1700px;margin:0 auto;padding:28px 36px 90px;position:relative;}
  .page-intro{font-size:13.5px;color:var(--muted);margin:0 0 20px;}

  .error-banner{background:var(--red-wash);color:var(--red);border:1px solid var(--red);padding:12px 16px;border-radius:10px;margin-bottom:20px;font-size:13.5px;}

  .loading-overlay{position:fixed;inset:0;background:var(--page);display:flex;align-items:center;justify-content:center;z-index:500;}
  .loading-box{display:flex;flex-direction:column;align-items:center;gap:16px;text-align:center;}
  .spinner-lg{width:40px;height:40px;border-radius:50%;border:3px solid var(--line);border-top-color:var(--red);animation:spin .8s linear infinite;}
  .loading-title{font-family:var(--font-display);font-size:15px;font-weight:800;color:var(--ink);}
  .loading-sub{font-size:12.5px;color:var(--muted);}
  .spinner{width:16px;height:16px;border-radius:50%;border:2px solid var(--line);border-top-color:var(--red);animation:spin .8s linear infinite;flex:0 0 auto;}
  @keyframes spin{to{transform:rotate(360deg);}}

  .kpi-strip{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px;margin-bottom:22px;}
  .kpi-card{background:var(--card);border:1px solid var(--line);border-left:3px solid var(--red);border-radius:10px;padding:14px 16px;text-align:left;font-family:inherit;transition:box-shadow .2s ease,transform .2s ease;animation:fadeInUp .35s ease both;}
  .kpi-card:hover{box-shadow:0 8px 20px rgba(20,18,15,0.08);transform:translateY(-2px);}
  .kpi-card:nth-child(1){animation-delay:.02s;}
  .kpi-card:nth-child(2){animation-delay:.06s;}
  .kpi-card:nth-child(3){animation-delay:.10s;}
  .kpi-card:nth-child(4){animation-delay:.14s;}
  .kpi-icon{display:block;color:var(--red);font-size:13px;margin-bottom:9px;}
  .kpi-value{font-family:var(--font-display);font-size:22px;font-weight:800;color:var(--ink);line-height:1;letter-spacing:-0.01em;display:inline-block;transition:color .3s ease;}
  .kpi-value.kpi-value-pulse{animation:kpiPulse .9s ease;color:var(--red);}
  .kpi-label{font-size:11.5px;color:var(--muted);margin-top:4px;white-space:nowrap;}
  @keyframes kpiPulse{0%{transform:scale(1);}35%{transform:scale(1.16);}100%{transform:scale(1);}}

  .controls{display:flex;gap:8px;margin-bottom:16px;flex-wrap:wrap;align-items:center;}
  .controls button{border:1px solid var(--line);background:var(--card);font-size:12.5px;color:var(--muted);padding:9px 16px;border-radius:8px;cursor:pointer;transition:color .15s ease,border-color .15s ease,background .15s ease;}
  .controls button:hover{color:var(--ink);border-color:#ddd9d1;background:var(--page);}
  .controls button i{transition:transform .3s ease;}
  .controls button:hover i{transform:rotate(25deg);}
  .controls button.primary{background:var(--red);border-color:var(--red);color:#fff;margin-left:auto;}
  .controls button.primary:hover{background:var(--red);color:#fff;}
  .controls button.primary:hover i{transform:none;}
  .controls button:disabled{opacity:.6;cursor:not-allowed;}
  .controls button.spinning i{animation:spin .8s linear infinite;}
  .search-wrap{position:relative;flex:1;min-width:220px;}
  .search-wrap input{width:100%;padding:9px 16px 9px 36px;border:1px solid var(--line);background:var(--card);font-size:13px;border-radius:8px;color:var(--ink);}
  .search-wrap input::placeholder{color:var(--faint);}
  .search-icon{position:absolute;left:14px;top:50%;transform:translateY(-50%);color:var(--faint);font-size:13px;pointer-events:none;}

  .table-wrap{overflow-x:auto;background:var(--card);border:1px solid var(--line);border-radius:12px;box-shadow:0 6px 20px rgba(20,18,15,0.06);}
  .contacts-table{table-layout:fixed;width:100%;min-width:2080px;border-collapse:collapse;font-size:13px;}
  .contacts-table thead th{position:relative;text-align:left;padding:11px 18px 11px 16px;font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:#fff;font-weight:700;background:var(--red);overflow:hidden;border-right:1px solid rgba(255,255,255,0.16);}
  .contacts-table thead th:last-child{border-right:none;}
  .th-inner{display:flex;align-items:center;justify-content:space-between;gap:6px;}
  .th-label{display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
  .th-filter-btn{flex:0 0 auto;background:none;border:none;padding:3px 5px;color:rgba(255,255,255,0.75);cursor:pointer;font-size:10px;border-radius:5px;line-height:1;transition:background .15s ease,color .15s ease;}
  .th-filter-btn:hover{color:#fff;background:rgba(255,255,255,0.18);}
  .th-filter-btn.active{color:#fff;background:rgba(255,255,255,0.32);}
  .th-resizer{position:absolute;top:0;right:0;bottom:0;width:8px;cursor:col-resize;background:transparent;touch-action:none;}
  .th-resizer:hover,.th-resizer.active{background:rgba(255,255,255,0.28);}

  .col-menu{position:fixed;background:var(--card);border:1px solid var(--line);border-radius:10px;box-shadow:0 14px 32px rgba(20,18,15,0.14);padding:6px;width:240px;z-index:2000;font-size:13px;}
  .col-menu.hidden{display:none;}
  .col-menu-item{display:flex;align-items:center;gap:9px;width:100%;background:none;border:none;padding:9px 11px;border-radius:8px;font-size:13px;color:var(--ink);cursor:pointer;text-align:left;}
  .col-menu-item i{width:14px;text-align:center;color:var(--muted);}
  .col-menu-item:hover{background:var(--red-wash);color:var(--red);}
  .col-menu-item:hover i{color:var(--red);}
  .col-menu-divider{height:1px;background:var(--line);margin:6px 4px;}
  .col-menu-label{font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:var(--faint);font-weight:700;padding:8px 11px 6px;}
  .col-menu-search{margin:0 6px 6px;padding:6px 10px;border:1px solid var(--line);border-radius:6px;background:var(--page);color:var(--ink);font-size:12px;width:calc(100% - 12px);}
  .col-menu-values{max-height:220px;overflow-y:auto;padding:2px 2px;}
  .col-menu-range{display:flex;gap:8px;padding:4px 6px 8px;}
  .col-menu-range label{flex:1;display:flex;flex-direction:column;gap:4px;font-size:11px;color:var(--muted);font-weight:600;text-transform:uppercase;letter-spacing:.03em;}
  .col-menu-range-input{font-size:13px;font-weight:400;text-transform:none;letter-spacing:0;padding:7px 10px;border:1px solid var(--line);border-radius:6px;background:var(--page);color:var(--ink);width:100%;box-sizing:border-box;}
  .col-menu-range-input:focus{outline:2px solid var(--red);outline-offset:1px;}
  .col-menu-value{display:flex !important;align-items:center !important;gap:8px;padding:6px 9px;border-radius:6px;font-size:12.5px;cursor:pointer;}
  .col-menu-value:hover{background:var(--page);}
  .col-menu-value input[type="checkbox"]{position:static !important;opacity:1 !important;appearance:auto !important;-webkit-appearance:checkbox !important;float:none !important;display:inline-block !important;width:16px !important;height:16px !important;min-width:16px !important;margin:0 !important;padding:0 !important;flex:0 0 auto !important;accent-color:var(--red);}
  .col-menu-selectall{font-weight:600;border-bottom:1px solid var(--line);margin-bottom:2px;padding-bottom:8px;}
  .col-menu-footer{display:flex;justify-content:space-between;gap:8px;padding:8px 4px 2px;border-top:1px solid var(--line);margin-top:6px;}
  .col-menu-footer button{border:1px solid var(--line);background:var(--card);font-size:12px;font-weight:600;padding:7px 14px;border-radius:8px;cursor:pointer;color:var(--ink);}
  .col-menu-footer button.primary{background:var(--red);border-color:var(--red);color:#fff;}

  @keyframes rowFadeIn{from{opacity:0;transform:translateY(4px);}to{opacity:1;transform:translateY(0);}}
  .contacts-table tbody td{padding:11px 16px;border-bottom:1px solid var(--line);vertical-align:middle;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
  .contacts-table tbody tr{cursor:pointer;transition:background .15s ease;animation:rowFadeIn .25s ease both;}
  .contacts-table tbody tr:nth-child(even){background:rgba(197,31,52,0.025);}
  .contacts-table tbody tr:hover{background:var(--page);}
  .contacts-table tbody tr:hover td:first-child{box-shadow:inset 3px 0 0 var(--red);}
  .contacts-table tbody tr:last-child td{border-bottom:none;}
  .contacts-table td.id-cell{color:var(--faint);font-variant-numeric:tabular-nums;}
  .contacts-table td.name-cell{font-family:var(--font-display);font-weight:800;color:var(--ink);}
  .muted-cell{color:var(--muted);}
  .contacts-table tbody tr:nth-child(1){animation-delay:0ms;}
  .contacts-table tbody tr:nth-child(2){animation-delay:15ms;}
  .contacts-table tbody tr:nth-child(3){animation-delay:30ms;}
  .contacts-table tbody tr:nth-child(4){animation-delay:45ms;}
  .contacts-table tbody tr:nth-child(5){animation-delay:60ms;}
  .contacts-table tbody tr:nth-child(6){animation-delay:75ms;}
  .contacts-table tbody tr:nth-child(7){animation-delay:90ms;}
  .contacts-table tbody tr:nth-child(8){animation-delay:105ms;}
  .contacts-table tbody tr:nth-child(9){animation-delay:120ms;}
  .contacts-table tbody tr:nth-child(10){animation-delay:135ms;}
  .chip{display:inline-flex;align-items:center;gap:6px;font-size:11.5px;font-weight:600;padding:4px 12px;border-radius:20px;white-space:nowrap;}
  .chip.neutral{background:var(--page);color:var(--muted);border:1px solid var(--line);}
  .chip-extra{color:var(--faint);font-size:11.5px;margin-left:4px;}

  .empty{text-align:center;padding:60px 20px;color:var(--muted);}
  .empty-icon{width:44px;height:44px;border-radius:50%;background:var(--green-wash);color:var(--green);display:flex;align-items:center;justify-content:center;margin:0 auto 12px;font-size:18px;}
  .empty-title{font-family:var(--font-display);font-size:14.5px;font-weight:800;color:var(--ink);margin-bottom:2px;}
  .empty-sub{font-size:12.5px;color:var(--muted);}

  .pagination{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-top:16px;flex-wrap:wrap;}
  .pagination-info{font-size:12px;color:var(--muted);}
  .pagination-nav{display:flex;align-items:center;gap:10px;}
  .page-btn{width:32px;height:32px;border-radius:8px;border:1px solid var(--line);background:var(--card);color:var(--ink);display:flex;align-items:center;justify-content:center;cursor:pointer;font-size:12px;transition:border-color .15s ease,color .15s ease,background .15s ease;}
  .page-btn:hover:not(:disabled){border-color:var(--red);color:var(--red);background:var(--red-wash);}
  .page-btn:disabled{opacity:.35;cursor:not-allowed;}
  .pagination-pages{font-size:12.5px;color:var(--muted);font-variant-numeric:tabular-nums;white-space:nowrap;}

  @keyframes modalPop{from{opacity:0;transform:scale(0.96) translateY(6px);}to{opacity:1;transform:scale(1) translateY(0);}}
  .overlay{position:fixed;inset:0;background:rgba(20,18,15,0.45);z-index:1000;backdrop-filter:blur(2px);display:flex;align-items:center;justify-content:center;padding:24px;}
  .overlay.hidden{display:none;}
  .modal{position:relative;width:100%;max-width:720px;max-height:88vh;background:var(--card);border-top:3px solid var(--red);overflow-y:auto;padding:34px 38px;border-radius:16px;box-shadow:0 28px 64px rgba(20,18,15,0.22);animation:modalPop .18s ease;scrollbar-width:thin;scrollbar-color:var(--line) transparent;}
  .modal::-webkit-scrollbar{width:6px;}
  .modal::-webkit-scrollbar-track{background:transparent;}
  .modal::-webkit-scrollbar-thumb{background:var(--line);border-radius:20px;}
  .modal::-webkit-scrollbar-thumb:hover{background:var(--faint);}
  .modal-top{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;}
  .modal h2{font-family:var(--font-display);font-size:21px;font-weight:800;margin:0 0 4px;letter-spacing:-0.01em;}
  .modal .modal-sub{color:var(--muted);font-size:13.5px;margin-bottom:20px;}
  .close{background:var(--page);border:none;width:30px;height:30px;border-radius:50%;font-size:14px;cursor:pointer;color:var(--muted);flex:0 0 auto;display:flex;align-items:center;justify-content:center;transition:background .15s ease,color .15s ease;}
  .close:hover{color:var(--red);background:var(--red-wash);}
  .section-label{font-size:11px;text-transform:uppercase;letter-spacing:0.08em;color:var(--faint);margin:20px 0 10px;font-weight:700;}
  .section-label:first-of-type{margin-top:0;}
  .field-grid{display:grid;grid-template-columns:170px 1fr;gap:12px 16px;font-size:14px;}
  .field-grid dt{color:var(--muted);}
  .field-grid dd{margin:0;color:var(--ink);display:flex;flex-wrap:wrap;gap:6px;align-items:center;}
  .modal .close-row{display:flex;justify-content:flex-end;gap:10px;margin-top:18px;}
  .btn{border:1px solid var(--line);background:var(--card);font-size:13px;font-weight:700;padding:10px 22px;border-radius:8px;cursor:pointer;color:var(--ink);transition:border-color .15s ease,transform .15s ease;}
  .btn.primary{background:var(--red);border-color:var(--red);color:#fff;}
  .btn:hover{border-color:var(--red);transform:translateY(-1px);}
  .btn:active{transform:scale(0.97);}
  .btn:disabled{opacity:.55;cursor:not-allowed;transform:none;}

  .form-grid{display:grid;grid-template-columns:1fr 1fr;gap:16px 18px;align-items:start;}
  .form-field{display:flex;flex-direction:column;gap:6px;}
  .form-field-wide{grid-column:1 / -1;}
  .form-field label{font-size:12px;color:var(--muted);font-weight:600;text-transform:uppercase;letter-spacing:.03em;line-height:1.3;}
  .form-field .req{color:var(--red);text-transform:none;}
  .form-field input,.form-field select{display:block;width:100%;box-sizing:border-box;font-size:13.5px;font-weight:400;padding:9px 12px;border:1px solid var(--line);border-radius:8px;background:var(--page);color:var(--ink);}
  .form-field select{cursor:pointer;}
  .form-field input:focus,.form-field select:focus{outline:2px solid var(--red);outline-offset:1px;}
  .form-field .hint{font-size:11px;color:var(--faint);font-weight:400;}
  .dropdown-check{position:relative;}
  .dropdown-trigger{display:flex;align-items:center;justify-content:space-between;gap:8px;width:100%;box-sizing:border-box;font-size:13.5px;font-weight:400;padding:9px 12px;border:1px solid var(--line);border-radius:8px;background:var(--page);color:var(--ink);cursor:pointer;text-align:left;}
  .dropdown-trigger:hover{border-color:#ddd9d1;}
  .dropdown-trigger-text{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
  .dropdown-trigger i{color:var(--faint);font-size:11px;flex:0 0 auto;transition:transform .15s ease;}
  .dropdown-trigger.open i{transform:rotate(180deg);}
  .checklist{position:absolute;top:calc(100% + 4px);left:0;right:0;z-index:30;border:1px solid var(--line);border-radius:8px;background:var(--card);box-shadow:0 14px 32px rgba(20,18,15,0.16);overflow:hidden;}
  .checklist-search{width:100%;box-sizing:border-box;font-size:13px;font-weight:400;padding:9px 12px;border:none;border-bottom:1px solid var(--line);background:var(--card);color:var(--ink);}
  .checklist-search:focus{outline:none;background:var(--page);}
  .checklist-items{max-height:200px;overflow-y:auto;padding:6px;display:flex !important;flex-direction:column !important;gap:2px;}
  .checklist-item{display:flex !important;flex-direction:row !important;align-items:center !important;gap:8px;padding:6px 8px;border-radius:6px;font-size:13px;font-weight:400;text-transform:none;letter-spacing:0;color:var(--ink);cursor:pointer;width:100%;box-sizing:border-box;}
  .checklist-item:hover{background:var(--page);}
  .checklist-item input[type="checkbox"]{position:static !important;opacity:1 !important;appearance:auto !important;-webkit-appearance:checkbox !important;float:none !important;display:inline-block !important;width:16px !important;height:16px !important;min-width:16px !important;margin:0 !important;padding:0 !important;flex:0 0 auto !important;accent-color:var(--red);}
  .checklist-item span{flex:1 1 auto;text-align:left;}
  .checklist-item.filtered-out{display:none !important;}
  .checklist-empty{font-size:12.5px;color:var(--faint);padding:6px 8px;}
  .form-error{background:var(--red-wash);color:var(--red);padding:9px 12px;border-radius:8px;font-size:12.5px;margin-top:12px;}

  .copilot-fab{position:fixed;bottom:24px;right:24px;width:56px;height:56px;border-radius:50%;background:var(--card);border:1px solid var(--line);padding:0;box-shadow:0 10px 24px rgba(20,18,15,0.18);cursor:pointer;display:flex;align-items:center;justify-content:center;z-index:1500;transition:transform .15s ease,box-shadow .15s ease;}
  .copilot-fab:hover{transform:translateY(-2px) scale(1.04);box-shadow:0 14px 28px rgba(20,18,15,0.24);}
  .copilot-fab.open{box-shadow:0 0 0 3px var(--red-wash),0 10px 24px rgba(20,18,15,0.18);}
  .copilot-fab img{width:30px;height:30px;object-fit:contain;display:block;}
  .copilot-launcher{position:fixed;bottom:92px;right:24px;width:320px;max-width:calc(100vw - 32px);background:var(--card);border:1px solid var(--line);border-radius:16px;box-shadow:0 24px 56px rgba(20,18,15,0.24);padding:10px;z-index:1499;animation:modalPop .18s ease;}
  .copilot-menu-item{display:flex;align-items:center;gap:12px;width:100%;background:none;border:none;padding:11px 12px;border-radius:10px;font-size:13.5px;color:var(--ink);cursor:pointer;text-align:left;}
  .copilot-menu-item i{width:16px;text-align:center;color:var(--muted);font-size:14px;}
  .copilot-menu-item:hover{background:var(--page);}
  .copilot-launcher-divider{height:1px;background:var(--line);margin:6px 4px;}
  .copilot-open-chat-btn{display:flex;align-items:center;gap:10px;width:100%;background:var(--card);border:1.5px solid var(--ink);border-radius:10px;padding:11px 12px;font-size:13.5px;font-weight:600;color:var(--ink);cursor:pointer;text-align:left;margin-bottom:8px;}
  .copilot-open-chat-btn:hover{background:var(--page);}
  .copilot-quick-ask{display:flex;align-items:center;gap:8px;border:1px solid var(--line);border-radius:999px;padding:6px 6px 6px 14px;}
  .copilot-quick-ask input{flex:1;border:none;background:none;font-size:13px;color:var(--ink);outline:none;min-width:0;}
  .copilot-quick-ask input::placeholder{color:var(--faint);}
  .copilot-quick-ask button{flex:0 0 auto;width:30px;height:30px;border-radius:50%;background:var(--red);color:#fff;border:none;cursor:pointer;display:flex;align-items:center;justify-content:center;font-size:12px;}
  .copilot-panel{position:fixed;bottom:92px;right:24px;width:360px;max-width:calc(100vw - 32px);height:500px;max-height:calc(100vh - 140px);background:var(--card);border:1px solid var(--line);border-radius:16px;box-shadow:0 24px 56px rgba(20,18,15,0.24);display:flex;flex-direction:column;overflow:hidden;z-index:1499;animation:modalPop .18s ease;}
  .copilot-header{display:flex;align-items:center;justify-content:space-between;padding:14px 16px;background:var(--red);color:#fff;flex:0 0 auto;}
  .copilot-title{font-family:var(--font-display);font-weight:800;font-size:14px;}
  .copilot-subtitle{font-size:11px;opacity:.85;margin-top:1px;}
  .copilot-header .close{background:rgba(255,255,255,0.18);color:#fff;}
  .copilot-header .close:hover{background:rgba(255,255,255,0.3);color:#fff;}
  .copilot-messages{flex:1 1 auto;overflow-y:auto;padding:14px 16px;display:flex;flex-direction:column;gap:10px;scrollbar-width:thin;scrollbar-color:var(--line) transparent;}
  .copilot-messages::-webkit-scrollbar{width:6px;}
  .copilot-messages::-webkit-scrollbar-thumb{background:var(--line);border-radius:20px;}
  .copilot-msg{max-width:85%;padding:9px 13px;border-radius:14px;font-size:13px;line-height:1.45;white-space:pre-wrap;word-break:break-word;}
  .copilot-msg.user{align-self:flex-end;background:var(--red);color:#fff;border-bottom-right-radius:4px;}
  .copilot-msg.assistant{align-self:flex-start;background:var(--page);color:var(--ink);border:1px solid var(--line);border-bottom-left-radius:4px;}
  .copilot-thinking{display:flex;align-items:center;gap:8px;padding:0 16px 10px;font-size:12px;color:var(--muted);flex:0 0 auto;}
  .copilot-input-row{display:flex;align-items:flex-end;gap:8px;padding:12px;border-top:1px solid var(--line);flex:0 0 auto;}
  .copilot-input-row textarea{flex:1;resize:none;max-height:90px;font-size:13px;padding:9px 12px;border:1px solid var(--line);border-radius:10px;background:var(--page);color:var(--ink);font-family:inherit;}
  .copilot-input-row textarea:focus{outline:2px solid var(--red);outline-offset:1px;}
  .copilot-send{flex:0 0 auto;width:38px;height:38px;border-radius:10px;background:var(--red);color:#fff;border:none;cursor:pointer;display:flex;align-items:center;justify-content:center;font-size:14px;transition:opacity .15s ease;}
  .copilot-send:disabled{opacity:.5;cursor:not-allowed;}

  @media (max-width:640px){
    .kpi-strip{grid-template-columns:repeat(2,1fr);}
    .search-wrap{flex-basis:100%;}
    .navbar-user-text{display:none;}
    .brand-eyebrow{display:none;}
    .form-grid{grid-template-columns:1fr;}
    .controls button.primary{margin-left:0;}
    .copilot-panel{right:16px;bottom:84px;width:calc(100vw - 32px);}
    .copilot-launcher{right:16px;bottom:84px;width:calc(100vw - 32px);}
    .copilot-fab{right:16px;bottom:16px;}
  }
</style>
</head>
<body>

<div class="top-accent" id="top-accent"></div>
<header class="app-header">
  <div class="brand-block">
    <img class="logo" src="https://raw.githubusercontent.com/skrineapps/skr-background-assets/main/logo.png" alt="Skrine" />
    <div class="brand-text">
      <div class="brand-title">Client Contacts Management System</div>
      <!-- <div class="brand-eyebrow">Skrine CRM</div> -->
    </div>
  </div>
  <div class="navbar-user" id="navbar-user">
    <button type="button" class="user-trigger" id="user-trigger" aria-haspopup="true" aria-expanded="false">
      <div class="navbar-user-text">
        <div class="navbar-user-name" id="user-name">&nbsp;</div>
        <div class="navbar-user-email" id="user-email"></div>
      </div>
      <div class="avatar" id="avatar">--</div>
    </button>
    <div class="user-menu hidden" id="user-menu">
      <button type="button" class="user-menu-item" id="sign-out-btn">
        <i class="fa-solid fa-arrow-right-from-bracket" aria-hidden="true"></i> Sign out
      </button>
    </div>
  </div>
</header>

<div class="page">
  <div class="error-banner hidden" id="error-banner"></div>
  <div class="loading-overlay" id="loading-state">
    <div class="loading-box">
      <span class="spinner-lg"></span>
      <div class="loading-title" id="loading-title">Signing you in...</div>
      <div class="loading-sub">This should only take a moment</div>
    </div>
  </div>

  <div id="app-content" class="hidden">
    <p class="page-intro" id="header-subtitle">Contacts linked to you, searchable and filterable.</p>

    <div class="kpi-strip" id="kpi-strip">
      <div class="kpi-card">
        <i class="fa-solid fa-user kpi-icon" aria-hidden="true"></i>
        <div class="kpi-value" id="stat-mine">0</div>
        <div class="kpi-label">My Contacts</div>
      </div>
      <div class="kpi-card">
        <i class="fa-solid fa-building kpi-icon" aria-hidden="true"></i>
        <div class="kpi-value" id="stat-companies">0</div>
        <div class="kpi-label">Companies</div>
      </div>
      <div class="kpi-card">
        <i class="fa-solid fa-earth-asia kpi-icon" aria-hidden="true"></i>
        <div class="kpi-value" id="stat-countries">0</div>
        <div class="kpi-label">Countries</div>
      </div>
      <div class="kpi-card">
        <i class="fa-solid fa-gift kpi-icon" aria-hidden="true"></i>
        <div class="kpi-value" id="stat-cards">0</div>
        <div class="kpi-label">On Greeting List</div>
      </div>
      <div class="kpi-card">
        <i class="fa-solid fa-scale-balanced kpi-icon" aria-hidden="true"></i>
        <div class="kpi-value" id="stat-practice-areas">0</div>
        <div class="kpi-label">Practice Areas</div>
      </div>
      <div class="kpi-card">
        <i class="fa-solid fa-graduation-cap kpi-icon" aria-hidden="true"></i>
        <div class="kpi-value" id="stat-alumni">0</div>
        <div class="kpi-label">Alumni / Foreign Law</div>
      </div>
    </div>

    <div class="controls">
      <div class="search-wrap">
        <i class="fa-solid fa-magnifying-glass search-icon" aria-hidden="true"></i>
        <input id="search" placeholder="Search name, company, position, email..." aria-label="Search contacts">
      </div>
      <button type="button" id="clear-filters">Clear filters</button>
      <button type="button" id="refresh-contacts" title="Reload contacts from SharePoint"><i class="fa-solid fa-arrows-rotate" aria-hidden="true"></i> Refresh</button>
      <button type="button" id="add-contact-btn" class="primary"><i class="fa-solid fa-plus" aria-hidden="true"></i> Add Contact</button>
    </div>

    <div class="table-wrap">
      <table class="contacts-table">
        <colgroup id="col-group"></colgroup>
        <thead id="table-head"></thead>
        <tbody id="contacts-list"></tbody>
      </table>
    </div>
    <div class="empty hidden" id="empty-state">
      <div class="empty-icon"><i class="fa-solid fa-address-card" aria-hidden="true"></i></div>
      <div class="empty-title" id="empty-title">No contacts here</div>
      <div class="empty-sub" id="empty-sub">Try clearing your search or column filters.</div>
    </div>
    <div class="pagination hidden" id="pagination">
      <div class="pagination-info" id="pagination-info"></div>
      <div class="pagination-nav">
        <button type="button" class="page-btn" id="page-prev" aria-label="Previous page"><i class="fa-solid fa-chevron-left" aria-hidden="true"></i></button>
        <span class="pagination-pages" id="pagination-pages"></span>
        <button type="button" class="page-btn" id="page-next" aria-label="Next page"><i class="fa-solid fa-chevron-right" aria-hidden="true"></i></button>
      </div>
    </div>
  </div>
</div>

<div class="col-menu hidden" id="col-menu"></div>

<div class="overlay hidden" id="overlay">
  <div class="modal" id="modal-box"></div>
</div>

<button type="button" class="copilot-fab" id="copilot-fab" aria-label="Open Copilot">
  <img src="https://raw.githubusercontent.com/skrineapps/skr-background-assets/main/Copilot-Logo.png" alt="Copilot" />
</button>

<div class="copilot-launcher hidden" id="copilot-launcher">
  <button type="button" class="copilot-menu-item" data-prompt="Summarize my contacts for me.">
    <i class="fa-solid fa-chart-pie" aria-hidden="true"></i> Summarize my contacts
  </button>
  <button type="button" class="copilot-menu-item" data-prompt="Which practice areas do our contacts fall under, and which are most common?">
    <i class="fa-solid fa-scale-balanced" aria-hidden="true"></i> Practice area breakdown
  </button>
  <button type="button" class="copilot-menu-item" data-prompt="Which countries do our contacts come from?">
    <i class="fa-solid fa-earth-asia" aria-hidden="true"></i> Contacts by country
  </button>
  <div class="copilot-launcher-divider"></div>
  <button type="button" class="copilot-open-chat-btn" id="copilot-open-chat">
    <i class="fa-solid fa-comment-dots" aria-hidden="true"></i> Open chat
  </button>
  <form class="copilot-quick-ask" id="copilot-quick-ask">
    <input type="text" id="copilot-quick-input" placeholder="Ask a question about your contacts" />
    <button type="submit" aria-label="Send"><i class="fa-solid fa-paper-plane" aria-hidden="true"></i></button>
  </form>
</div>

<div class="copilot-panel hidden" id="copilot-panel">
  <div class="copilot-header">
    <div class="copilot-header-text">
      <div class="copilot-title">Skrine Copilot</div>
      <div class="copilot-subtitle">Ask about your contacts</div>
    </div>
    <button type="button" class="close" id="copilot-close" aria-label="Close Copilot"><i class="fa-solid fa-xmark" aria-hidden="true"></i></button>
  </div>
  <div class="copilot-messages" id="copilot-messages"></div>
  <div class="copilot-thinking hidden" id="copilot-thinking">
    <span class="spinner"></span> Thinking...
  </div>
  <div class="copilot-input-row">
    <textarea id="copilot-input" placeholder="Ask a question..." rows="1"></textarea>
    <button type="button" class="copilot-send" id="copilot-send" aria-label="Send"><i class="fa-solid fa-paper-plane" aria-hidden="true"></i></button>
  </div>
</div>

<script>
(function () {
  "use strict";

  const CONFIG = {
    clientId: "b679e920-ad66-4f7a-95d7-9bb252c0a1d5",
    tenantId: "c8ad7181-aeee-4642-822e-51c231e5f9b5",
    siteHostname: "skrineonline.sharepoint.com",
    sitePath: "/sites/SkrineApps",
    listDisplayName: "CRM INFO",
    // Sites.ReadWrite.All (delegated) replaces Sites.Read.All so Add/Edit Contact can write
    // back to the list. GroupMember.Read.All lets the app list members of the Partner In
    // Charge / Lawyers security groups below. Both need to be added to the app registration
    // and admin-consented in Azure AD, or the relevant feature fails outright.
    graphScopes: ["User.Read", "Sites.ReadWrite.All", "GroupMember.Read.All"],
    // Members of this Azure AD security group see every contact under "My Contacts" instead
    // of just the ones matched via Lawyers/CreatedByEmailText. Read from the ID token's
    // "groups" claim (Token configuration > Add groups claim > Security groups in the app
    // registration) - no extra Graph permission needed.
    viewAllContactsGroupId: "56587622-97af-4e39-bbc2-cb6b52d40765",
    // Microsoft Graph's SharePoint list-items API generally can't resolve Person/Group
    // columns (Lawyers, Partner In Charge) at all, with or without an explicit $select -
    // it's a longstanding Graph limitation, not something fixable from the query string.
    // The classic SharePoint REST API resolves (and, for writes, sets) them correctly, so
    // that's used just for those two fields. This needs the "SharePoint" API's AllSites.Write
    // delegated permission added to the app registration and admin-consented (a different
    // API than Microsoft Graph, which is what Sites.ReadWrite.All above belongs to) - without
    // it, reads of these two fields silently stay blank and saves to them fail with an error.
    sharePointResource: "https://skrineonline.sharepoint.com",
    sharePointScopes: ["https://skrineonline.sharepoint.com/AllSites.Write"],
    // Add Contact / Edit Contact populates the Partner In Charge dropdown and the Lawyers
    // checklist from the (transitive/nested) members of these two Azure AD security groups.
    partnerInChargeGroupId: "69f85e06-b41b-4e15-8e15-1bb7c7524f15",
    lawyersGroupId: "07409114-d634-491c-a7a7-c08d7da57734",
    // Copilot chat, via Azure OpenAI - authenticated with the same Azure AD sign-in as
    // everything else (no API key embedded in this file, since a static page has no server
    // to keep a key secret from anyone who views its source). Needs:
    // 1. An Azure OpenAI resource with a model deployed - fill in endpoint/deploymentName
    //    below once you have one (see the comment above CONFIG for setup steps).
    // 2. The signed-in users granted the "Cognitive Services OpenAI User" role on that
    //    resource (Azure Portal > the resource > Access control (IAM) > Add role assignment).
    // 3. This app registration granted the "Azure Cognitive Services" API's delegated
    //    user_impersonation permission, admin-consented (a third API, distinct from both
    //    Microsoft Graph and the SharePoint API used elsewhere in this file).
    azureOpenAI: {
      endpoint: "https://skrine-copilot.services.ai.azure.com/openai/v1",
      deploymentName: "gpt-4.1-mini",
      scopes: ["https://cognitiveservices.azure.com/.default"],
    },
  };

  // If your SharePoint list's internal column names differ from these guesses,
  // add the real internal name to the relevant array (check via a contact's
  // "Raw fields" debug panel in the UI).
  const FIELD_ALIASES = {
    firstName: ["FirstName", "First_x0020_Name"],
    lastName: ["LastName", "Last_x0020_Name"],
    email: ["Email"],
    companyName: ["CompanyName", "Company_x0020_Name"],
    salutation: ["Salutation"],
    position: ["Position"],
    phoneNumber: ["PhoneNumber", "Phone_x0020_Number"],
    country: ["Country"],
    partner: ["Partner"],
    partnerInCharge: ["PartnerInCharge", "Partner_x0020_In_x0020_Charge"],
    practiceArea: ["PracticeArea", "Practice_x0020_Area"],
    greetingCards: ["GreetingCards", "Greeting_x0020_Cards"],
    alumniForeign: ["Alumni_x002f_Foreign", "AlumniForeign"],
    // Confirmed separate from "Contact Owner" (a plain text field, unrelated). No column
    // matched "Lawyers" by internal name as of the last schema scan - fetchListSchema also
    // tries matching by current display name as a fallback and logs the full column list if
    // neither finds it, since the internal name may not resemble "Lawyers" at all.
    lawyers: ["Lawyers"],
    contactOwner: ["ContactOwner"],
  };

  // Fallback match target for fetchListSchema's field-name resolver, used when a column's
  // internal name doesn't match any FIELD_ALIASES candidate (most often because the column
  // was renamed after creation - SharePoint keeps the original internal name forever).
  const DISPLAY_NAME_GUESSES = {
    firstName: ["First Name"],
    lastName: ["Last Name"],
    email: ["Email"],
    companyName: ["Company Name"],
    salutation: ["Salutation"],
    position: ["Position"],
    phoneNumber: ["Phone Number"],
    country: ["Country"],
    partner: ["Partner"],
    partnerInCharge: ["Partner In Charge"],
    practiceArea: ["Practice Area"],
    greetingCards: ["Greeting Cards", "Greeting Card"],
    alumniForeign: ["Alumni / Foreign", "Alumni/Foreign", "Alumni / Foreign Law"],
    lawyers: ["Lawyers", "Lawyer In Charge", "Lawyer"],
    contactOwner: ["Contact Owner"],
  };

  const PAGE_SIZE = 25;

  // Single source of truth for the table: header labels and how to pull a plain-text value
  // out of a contact for both rendering and per-column filtering. "chip" columns render as
  // chips instead of plain text. Default column widths aren't listed here - they're derived
  // from each label's character length (see defaultColumnWidths below), so a column starts
  // just wide enough for its own header title; resizing by hand still overrides per-column.
  const COLUMNS = [
    { key: "id", label: "ID", cellClass: "id-cell", get: (c) => c.id, filterType: "range" },
    { key: "firstName", label: "First Name", cellClass: "name-cell", get: (c) => c.firstName },
    { key: "lastName", label: "Last Name", cellClass: "name-cell", get: (c) => c.lastName },
    { key: "email", label: "Email", cellClass: "muted-cell", get: (c) => c.email },
    { key: "companyName", label: "Company Name", get: (c) => c.companyName },
    { key: "salutation", label: "Salutation", cellClass: "muted-cell", get: (c) => c.salutation },
    { key: "position", label: "Position", cellClass: "muted-cell", get: (c) => c.position },
    { key: "phoneNumber", label: "Phone Number", cellClass: "muted-cell", get: (c) => c.phoneNumber },
    { key: "country", label: "Country", cellClass: "muted-cell", get: (c) => c.country },
    { key: "partner", label: "Partner", cellClass: "muted-cell", get: (c) => c.partner },
    { key: "practiceArea", label: "Practice Area", get: (c) => c.practiceArea.join(", "), chip: (c) => c.practiceArea },
    { key: "greetingCards", label: "Greeting Card", get: (c) => c.greetingCards.join(", "), chip: (c) => c.greetingCards, chipColor: greetingChipColor },
    { key: "alumniForeign", label: "Alumni / Foreign Law", cellClass: "muted-cell", get: (c) => c.alumniForeign },
    { key: "lawyers", label: "Lawyers", get: (c) => c.lawyers.join(", "), chip: (c) => c.lawyers },
    { key: "created", label: "Created Date", cellClass: "muted-cell", get: (c) => formatDateTime(c.created), filterType: "date" },
  ];

  // Each column's share of the table's width, proportional to its header title's length -
  // a few extra "characters" are padded in per column to leave room for the filter icon and
  // padding so very short headers like "ID" aren't crushed. Percentages sum to ~100.
  function defaultColumnWidths() {
    const CHROME_CHARS = 4;
    const weights = COLUMNS.map((col) => col.label.length + CHROME_CHARS);
    const total = weights.reduce((a, b) => a + b, 0);
    return weights.map((w) => (w / total) * 100);
  }

  let msalInstance, activeAccount, currentUser;
  let allContacts = [];
  let searchTerm = "";
  let currentPage = 1;
  let userCanViewAllContacts = false;

  // Populated once at bootstrap: dropdown choices read from the SharePoint list's own Choice
  // columns, and the Partner In Charge / Lawyers people-pickers read from their AD security
  // groups. The Add/Edit form renders from these rather than hardcoding any options.
  let choiceOptions = { salutation: [], country: [], practiceArea: [], greetingCards: [], alumniForeign: [] };
  let partnerInChargeOptions = [];
  let lawyerOptions = [];
  let listEntityTypeName = null;
  const spUserIdCache = new Map();

  // The real internal SharePoint column names for the two Person/Group fields, discovered
  // from the list schema rather than guessed - see fetchListSchema. A guessed name (e.g.
  // FIELD_ALIASES.lawyers[0] === "Lawyers") can be wrong even when it matches the CSV export
  // header, because SharePoint keeps a column's original internal name forever even after
  // the display name (and therefore the export header) changes.
  let personColumnNames = { partnerInCharge: null, lawyers: null };
  // Whether each resolved Person/Group column actually allows multiple people (from Graph's
  // own personOrGroup.allowMultipleSelection facet) - determines whether writes need the
  // {results: [...]} wrapper or a bare single ID. Don't assume from the UI (a single-select
  // dropdown doesn't prove the underlying column is single-value) or from FIELD_ALIASES.
  let personColumnMultiValue = { partnerInCharge: false, lawyers: false };

  // The real internal SharePoint column name for every known field, resolved from the list
  // schema the same way as personColumnNames above - see fetchListSchema. writeKey() reads
  // from this first, falling back to the original FIELD_ALIASES guess only if the schema
  // scan hasn't run yet or genuinely found no match, since a guessed name that's wrong is
  // exactly what causes Graph's unhelpful generic 400 "Invalid request" on save.
  let resolvedFieldNames = {};

  // Column filters work like Excel/SharePoint's column header menu: columnFilterValues[key]
  // is a Set of the values allowed through for that column, or absent entirely when no
  // filter is applied (i.e. everything is shown). sortColumn/sortDirection likewise come
  // from that same per-column dropdown instead of a separate control.
  let columnFilterValues = {};
  let sortColumn = "id";
  let sortDirection = "desc";
  let openMenuKey = null;

  var LOADING_MESSAGES = ["Signing you in...", "Loading your profile...", "Loading your contacts..."];
  var loadingMessageTimer = null;

  function setLoadingText(title) {
    var el = document.getElementById("loading-title");
    if (el) el.textContent = title;
  }

  // ---------- small helpers ----------

  // Short-TTL sessionStorage cache for data that rarely changes (list schema, AD group
  // membership) but was otherwise being re-fetched from scratch on every single browser
  // refresh, even though nothing about it had changed since the last load a moment ago.
  // sessionStorage survives a browser refresh (unlike page-lifetime variables) but not a
  // closed tab, so this only ever saves round trips within the same active session.
  function getSessionCache(key, maxAgeMs) {
    try {
      const raw = sessionStorage.getItem(key);
      if (!raw) return null;
      const { value, savedAt } = JSON.parse(raw);
      if (Date.now() - savedAt > maxAgeMs) return null;
      return value;
    } catch (e) {
      return null;
    }
  }
  function setSessionCache(key, value) {
    try {
      sessionStorage.setItem(key, JSON.stringify({ value, savedAt: Date.now() }));
    } catch (e) { /* storage full or unavailable - just skip caching this time */ }
  }

  function normalizeKey(k) { return k.toLowerCase().replace(/[^a-z0-9]/g, ""); }

  function getField(fields, aliasKey) {
    const candidates = FIELD_ALIASES[aliasKey] || [aliasKey];
    for (const c of candidates) { if (fields[c] !== undefined) return fields[c]; }
    const targets = candidates.map(normalizeKey);
    for (const [k, v] of Object.entries(fields)) {
      if (targets.includes(normalizeKey(k))) return v;
    }
    return undefined;
  }

  function parseMaybeJsonArray(value) {
    // Every path trims each element - inconsistent trailing/leading whitespace on individual
    // choice values (varies per item, likely from how they were originally entered) was
    // otherwise producing "duplicate" entries in filter dropdowns that look identical but
    // aren't the same string, since only the delimiter-split fallback used to trim.
    if (Array.isArray(value)) return value.map((v) => String(v).trim()).filter(Boolean);
    if (value == null || value === "") return [];
    if (typeof value === "string") {
      const trimmed = value.trim();
      if (trimmed.startsWith("[")) {
        try {
          const parsed = JSON.parse(trimmed);
          return Array.isArray(parsed) ? parsed.map((v) => String(v).trim()).filter(Boolean) : [trimmed];
        } catch (e) { /* fall through to delimiter split */ }
      }
      return trimmed.split(/[;,]/).map((s) => s.trim()).filter(Boolean);
    }
    return [String(value).trim()].filter(Boolean);
  }

  // Plain text columns can come back as SharePoint Person/Group objects if a column
  // turns out to be a lookup rather than text (e.g. {LookupValue, Email, Title}) -
  // this always renders something readable instead of "[object Object]".
  function asPlainText(value) {
    if (value == null) return "";
    if (typeof value === "string") return value;
    if (Array.isArray(value)) return value.map(asPlainText).filter(Boolean).join(", ");
    if (typeof value === "object") return value.Email || value.LookupValue || value.Title || value.displayName || "";
    return String(value);
  }

  // Person/Group columns (Created By, Partner In Charge, Lawyers) come back via Graph
  // as an object (single-value) or array of objects (multi-value) shaped roughly like
  // {LookupId, LookupValue, Email} rather than a plain string.
  function parsePersonField(value) {
    if (value == null) return [];
    const arr = Array.isArray(value) ? value : [value];
    return arr
      .map((v) => {
        if (typeof v === "string") return { displayName: v, email: "" };
        if (v && typeof v === "object") {
          return {
            displayName: v.LookupValue || v.Title || v.displayName || v.DisplayName || "",
            email: v.Email || v.EMail || v.email || v.WorkEmail || "",
          };
        }
        return { displayName: "", email: "" };
      })
      .filter((p) => p.displayName || p.email);
  }

  function normalizeName(s) {
    return (s || "").toLowerCase().replace(/["'.,]/g, "").replace(/\s+/g, " ").trim();
  }
  function normalizeEmail(s) { return (s || "").toLowerCase().trim(); }

  function nameMatches(a, b) {
    const na = normalizeName(a), nb = normalizeName(b);
    if (!na || !nb) return false;
    return na === nb || na.includes(nb) || nb.includes(na);
  }

  // Specific colors for the known greeting-card occasions. Matched by keyword rather than
  // exact string so minor variations in how the choice is worded in SharePoint (e.g. "Hari
  // Raya" vs "Hari Raya Aidilfitri") still land on the right color.
  const GREETING_COLOR_RULES = [
    { test: /chinese/i, bg: "rgba(224,86,43,0.16)", fg: "#c9491e" }, // red-orange
    { test: /deepavali|diwali/i, bg: "rgba(74,58,167,0.14)", fg: "#4a3aa7" }, // purple
    { test: /hari\s*raya|aidilfitri/i, bg: "var(--green-wash)", fg: "var(--green)" }, // green
    { test: /season'?s?\s*greetings|christmas/i, bg: "rgba(42,120,214,0.14)", fg: "#2a78d6" }, // blue
  ];

  // Anything not matching a known occasion above (a new choice added later, or "(Blank)")
  // still gets a color - picked deterministically by hashing the text, so the same label
  // always lands on the same fallback color instead of looking uncategorized.
  const CHIP_FALLBACK_PALETTE = [
    { bg: "var(--red-wash)", fg: "var(--red)" },
    { bg: "var(--amber-wash)", fg: "var(--amber)" },
    { bg: "var(--green-wash)", fg: "var(--green)" },
    { bg: "rgba(74,58,167,0.14)", fg: "#4a3aa7" },
    { bg: "rgba(42,120,214,0.14)", fg: "#2a78d6" },
    { bg: "rgba(232,123,164,0.18)", fg: "#c24a7c" },
  ];

  function hashString(s) {
    let hash = 0;
    for (let i = 0; i < s.length; i++) hash = (hash * 31 + s.charCodeAt(i)) | 0;
    return Math.abs(hash);
  }

  function greetingChipColor(label) {
    const rule = GREETING_COLOR_RULES.find((r) => r.test.test(label || ""));
    if (rule) return rule;
    return CHIP_FALLBACK_PALETTE[hashString(label || "") % CHIP_FALLBACK_PALETTE.length];
  }

  function chipCell(items, colorFn) {
    if (!items || !items.length) return `<span class="muted-cell">-</span>`;
    const [first, ...rest] = items;
    const style = colorFn ? ` style="background:${colorFn(first).bg};color:${colorFn(first).fg};"` : "";
    return `<span class="chip"${colorFn ? "" : " neutral"}${style}>${escapeHtml(first)}</span>${rest.length ? `<span class="chip-extra">+${rest.length}</span>` : ""}`;
  }

  function escapeHtml(s) {
    return (s || "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }

  // "Others" is a catch-all salutation choice, not an actual title, so it reads oddly
  // prefixed onto a name (e.g. "Others Kean Lee Swee") - every real salutation (Mr, Datuk,
  // Dr, etc.) still shows normally.
  function contactFullName(c) {
    const salutation = c.salutation && c.salutation.trim().toLowerCase() !== "others" ? c.salutation : "";
    return [salutation, c.firstName, c.lastName].filter(Boolean).join(" ");
  }

  function initialsOf(fullname) {
    if (!fullname) return "--";
    var parts = fullname.trim().split(/\s+/);
    var first = parts[0] ? parts[0][0] : "";
    var last = parts.length > 1 ? parts[parts.length - 1][0] : "";
    return (first + last).toUpperCase() || "--";
  }

  function formatDateTime(iso) {
    if (!iso) return "";
    const d = new Date(iso);
    const datePart = d.toLocaleDateString();
    const timePart = d.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit", hour12: false });
    return `${datePart} ${timePart}`;
  }

  // ---------- MSAL / Graph ----------

  async function ensureSignedIn() {
    msalInstance = new msal.PublicClientApplication({
      auth: {
        clientId: CONFIG.clientId,
        authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
        redirectUri: window.location.href.split("#")[0].split("?")[0],
      },
      cache: { cacheLocation: "sessionStorage", storeAuthStateInCookie: false },
    });
    if (typeof msalInstance.initialize === "function") await msalInstance.initialize();

    const redirectResult = await msalInstance.handleRedirectPromise();
    if (redirectResult && redirectResult.account) {
      activeAccount = redirectResult.account;
      msalInstance.setActiveAccount(activeAccount);
      return;
    }
    const existing = msalInstance.getAllAccounts();
    if (existing.length) {
      activeAccount = existing[0];
      msalInstance.setActiveAccount(activeAccount);
      return;
    }
    try {
      const res = await msalInstance.ssoSilent({ scopes: CONFIG.graphScopes });
      activeAccount = res.account;
    } catch (silentErr) {
      // No popup: if there's truly no ambient session to pick up silently, the whole page
      // navigates to Microsoft sign-in and back, rather than opening a popup window. This is
      // also just plain more reliable than a popup when the page is loaded inside an iframe
      // (as it sometimes is here, embedded in SharePoint's own page chrome) - a popup can be
      // blocked in that context, and login.microsoftonline.com refuses to load inside an
      // iframe at all, so redirecting the top-level page is the option that actually works.
      console.warn("[Contacts] Silent sign-in failed, redirecting to Microsoft sign-in (no popup):", silentErr);
      await msalInstance.loginRedirect({ scopes: CONFIG.graphScopes });
      return; // the page is navigating away - nothing after this point will run
    }
    msalInstance.setActiveAccount(activeAccount);
  }

  async function getGraphToken() {
    const request = { scopes: CONFIG.graphScopes, account: activeAccount };
    try {
      const res = await msalInstance.acquireTokenSilent(request);
      return res.accessToken;
    } catch (e) {
      // Same no-popup reasoning as ensureSignedIn: this navigates the whole page away and
      // back rather than opening a popup. In practice this should be rare - it only happens
      // if a token silently expired in a way acquireTokenSilent can't refresh on its own.
      await msalInstance.acquireTokenRedirect(request);
      return null; // unreachable in practice - the page is navigating away
    }
  }

  // A token for the SharePoint REST API resource, separate from the Graph token above -
  // used only for resolving Person/Group columns (see fetchPersonColumnsViaSharePointRest).
  async function getSharePointToken() {
    const request = { scopes: CONFIG.sharePointScopes, account: activeAccount };
    try {
      const res = await msalInstance.acquireTokenSilent(request);
      return res.accessToken;
    } catch (e) {
      await msalInstance.acquireTokenRedirect(request);
      return null; // unreachable in practice - the page is navigating away
    }
  }

  // A token for the Azure Cognitive Services resource (what Azure OpenAI is built on) -
  // a third, separate token audience from Graph and SharePoint above. Used only by the
  // Copilot chat.
  async function getAzureOpenAIToken() {
    const request = { scopes: CONFIG.azureOpenAI.scopes, account: activeAccount };
    try {
      const res = await msalInstance.acquireTokenSilent(request);
      return res.accessToken;
    } catch (e) {
      await msalInstance.acquireTokenRedirect(request);
      return null; // unreachable in practice - the page is navigating away
    }
  }

  async function graphGet(url) {
    const token = await getGraphToken();
    const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!res.ok) throw new Error(`Graph request failed (${res.status}): ${await res.text()}`);
    return res.json();
  }

  async function graphGetAllPages(url) {
    let items = [];
    let next = url;
    while (next) {
      const data = await graphGet(next);
      items = items.concat(data.value || []);
      next = data["@odata.nextLink"] || null;
    }
    return items;
  }

  async function resolveSiteAndListIds() {
    const cacheKey = "crmSiteListIds:v1";
    const cached = sessionStorage.getItem(cacheKey);
    if (cached) return JSON.parse(cached);

    const site = await graphGet(
      `https://graph.microsoft.com/v1.0/sites/${CONFIG.siteHostname}:${CONFIG.sitePath}`
    );
    const lists = await graphGet(
      `https://graph.microsoft.com/v1.0/sites/${site.id}/lists?$select=id,displayName,name`
    );
    const match = (lists.value || []).find(
      (l) => (l.displayName || "").toLowerCase() === CONFIG.listDisplayName.toLowerCase()
    );
    if (!match) {
      const available = (lists.value || []).map((l) => l.displayName).join(", ");
      throw new Error(`List "${CONFIG.listDisplayName}" not found on site. Available lists: ${available}`);
    }
    const ids = { siteId: site.id, listId: match.id };
    sessionStorage.setItem(cacheKey, JSON.stringify(ids));
    return ids;
  }

  function mapItemToContact(item) {
    const f = item.fields || {};
    const partnerInChargePeople = parsePersonField(getField(f, "partnerInCharge"));
    const lawyerPeople = parsePersonField(getField(f, "lawyers"));
    return {
      id: item.id,
      firstName: asPlainText(getField(f, "firstName")),
      lastName: asPlainText(getField(f, "lastName")),
      email: asPlainText(getField(f, "email")),
      companyName: asPlainText(getField(f, "companyName")),
      salutation: asPlainText(getField(f, "salutation")),
      position: asPlainText(getField(f, "position")),
      phoneNumber: asPlainText(getField(f, "phoneNumber")),
      country: asPlainText(getField(f, "country")),
      partner: asPlainText(getField(f, "partner")),
      partnerInCharge: partnerInChargePeople.map((p) => p.displayName).filter(Boolean).join(", "),
      partnerInChargeEmails: partnerInChargePeople.map((p) => p.email).filter(Boolean),
      practiceArea: parseMaybeJsonArray(getField(f, "practiceArea")),
      greetingCards: parseMaybeJsonArray(getField(f, "greetingCards")),
      alumniForeign: asPlainText(getField(f, "alumniForeign")),
      lawyers: lawyerPeople.map((p) => p.displayName).filter(Boolean),
      lawyerEmails: lawyerPeople.map((p) => p.email).filter(Boolean),
      // item.createdBy is Graph's own resolved item metadata (who actually created the list
      // item), not a custom column - unlike Lawyers/Partner In Charge it's always properly
      // resolved with no Person/Group-field limitations, so it's used instead of the old
      // CreatedByEmailText text column.
      createdBy: (item.createdBy && item.createdBy.user && item.createdBy.user.displayName) || "",
      createdByEmail: (item.createdBy && item.createdBy.user && item.createdBy.user.email) || "",
      contactOwner: asPlainText(getField(f, "contactOwner")),
      created: item.createdDateTime || "",
      modified: item.lastModifiedDateTime || "",
      raw: f,
    };
  }

  // Graph's list-items endpoint often fails to resolve multi-value Person/Group columns
  // (like Lawyers) unless they're explicitly named in the $select - a wildcard "expand=fields"
  // silently comes back empty for them. Naming every known column forces Graph to resolve
  // them properly. If any guessed internal name is wrong, Graph 400s on the whole request,
  // so this falls back to the plain wildcard expand (better degraded than fully broken).
  // SharePoint REST wraps an expanded multi-value (Person/Group) property in a
  // {results: [...]} envelope; a single-value one comes back as a plain object. This
  // normalizes either shape into an array of {displayName, email}.
  function personArrayFromRest(value) {
    if (!value) return [];
    const arr = Array.isArray(value) ? value : Array.isArray(value.results) ? value.results : [value];
    return arr.map((v) => ({ displayName: v.Title || "", email: v.EMail || "" })).filter((p) => p.displayName || p.email);
  }

  // Resolves Lawyers and Partner In Charge via the classic SharePoint REST API, which
  // (unlike Graph) correctly expands Person/Group columns. Uses odata=verbose specifically -
  // the lighter nometadata/minimalmetadata formats are known to be unreliable for $expand on
  // multi-value Person/Lookup columns, which is exactly the scenario this needs to work.
  // Returns { people, error }: a Map from item ID to { lawyers, partnerInCharge }, plus the
  // failure message if the whole request failed (e.g. missing permission, CORS, wrong list/
  // field name) - the caller surfaces that error directly in the UI instead of console-only,
  // since this exact failure has been hard to diagnose blind.
  async function fetchPersonColumnsViaSharePointRest() {
    const people = new Map();
    const lawyersField = personColumnNames.lawyers;
    const partnerField = personColumnNames.partnerInCharge;
    if (!lawyersField && !partnerField) {
      return { people, error: "Could not determine the Lawyers / Partner In Charge column names from the list schema (see the console log from fetchListSchema)." };
    }
    try {
      const token = await getSharePointToken();
      const listTitle = encodeURIComponent(CONFIG.listDisplayName);
      const selectParts = ["Id"];
      const expandParts = [];
      if (lawyersField) { selectParts.push(`${lawyersField}/Title`, `${lawyersField}/EMail`); expandParts.push(lawyersField); }
      if (partnerField) { selectParts.push(`${partnerField}/Title`, `${partnerField}/EMail`); expandParts.push(partnerField); }
      let url =
        `${CONFIG.sharePointResource}${CONFIG.sitePath}/_api/web/lists/getbytitle('${listTitle}')/items` +
        `?$select=${selectParts.join(",")}&$expand=${expandParts.join(",")}&$top=2000`;
      while (url) {
        const res = await fetch(url, {
          headers: { Authorization: `Bearer ${token}`, Accept: "application/json;odata=verbose" },
        });
        if (!res.ok) throw new Error(`SharePoint REST request failed (${res.status}): ${await res.text()}`);
        const data = await res.json();
        const items = (data.d && data.d.results) || [];
        items.forEach((item) => {
          people.set(String(item.Id), {
            lawyers: lawyersField ? personArrayFromRest(item[lawyersField]) : [],
            partnerInCharge: partnerField ? personArrayFromRest(item[partnerField]) : [],
          });
        });
        url = (data.d && data.d.__next) || null;
      }
      return { people, error: null };
    } catch (e) {
      const message = e && e.message ? e.message : String(e);
      console.warn(
        "[Contacts] Could not resolve Lawyers / Partner In Charge via SharePoint REST - check that the SharePoint API's AllSites.Write permission is added and admin-consented in the app registration (a different API than Microsoft Graph), that a fresh sign-in has happened since granting it, and that the page is being loaded from the actual SharePoint site (not a local file or different host, which would hit CORS). Those two fields will show blank until then.",
        e
      );
      return { people, error: message };
    }
  }

  async function fetchContacts() {
    const { siteId, listId } = await resolveSiteAndListIds();
    // Prefer the names discovered by fetchListSchema for the two Person/Group columns over
    // the guessed FIELD_ALIASES candidate - see fetchListSchema for why the guess can be wrong.
    const selectNames = Object.keys(FIELD_ALIASES)
      .map((key) => {
        if (key === "partnerInCharge" && personColumnNames.partnerInCharge) return personColumnNames.partnerInCharge;
        if (key === "lawyers" && personColumnNames.lawyers) return personColumnNames.lawyers;
        return writeKey(key);
      })
      .join(",");
    const preciseUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?expand=fields(select=${selectNames})&$top=200`;
    const fallbackUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?expand=fields&$top=200`;

    // The Graph items fetch and the SharePoint REST person-columns fetch each independently
    // read the whole list - they don't depend on each other's results, only get merged
    // together at the end, so they run concurrently instead of one after the other. This
    // roughly halves the time this function takes on a list with many contacts.
    const itemsPromise = (async () => {
      try {
        return await graphGetAllPages(preciseUrl);
      } catch (e) {
        console.warn(
          "[Contacts] Explicit field selection failed (a guessed internal column name in FIELD_ALIASES is likely wrong) - falling back to the default field expansion. Person/Group columns like Lawyers may come back empty under this fallback.",
          e
        );
        return graphGetAllPages(fallbackUrl);
      }
    })();

    const [items, personResult] = await Promise.all([itemsPromise, fetchPersonColumnsViaSharePointRest()]);
    const contacts = items.map(mapItemToContact);

    // Best-effort supplementary lookup - see fetchPersonColumnsViaSharePointRest. Note this
    // means the "Raw fields (debug)" panel (which only ever shows item.fields, i.e. what
    // Graph itself returned) will never show Lawyers/Partner In Charge even when this
    // succeeds - that data is added here, not to c.raw, which is why it's echoed into a
    // clearly-labeled extra debug key instead (including the exact failure reason, if any),
    // so the panel reflects the full picture without needing the browser console open.
    const { people: personColumns, error: personColumnsError } = personResult;
    contacts.forEach((c) => {
      const extra = personColumns.get(String(c.id));
      c.raw = {
        ...c.raw,
        _resolvedViaSharePointRest: personColumnsError
          ? `ERROR: ${personColumnsError}`
          : extra || "(no SharePoint REST match for this item ID)",
      };
      if (!extra) return;
      if (extra.lawyers.length) {
        c.lawyers = extra.lawyers.map((p) => p.displayName).filter(Boolean);
        c.lawyerEmails = extra.lawyers.map((p) => p.email).filter(Boolean);
      }
      if (extra.partnerInCharge.length) {
        c.partnerInCharge = extra.partnerInCharge.map((p) => p.displayName).filter(Boolean).join(", ");
        c.partnerInChargeEmails = extra.partnerInCharge.map((p) => p.email).filter(Boolean);
      }
    });
    return contacts;
  }

  // Reads the SharePoint list's own column schema via Graph and uses it for two things:
  // (1) each Choice column's configured options, for the Salutation/Country/Practice Area/
  // Greeting Card/Alumni dropdowns; (2) the REAL internal names of the two Person/Group
  // columns (Partner In Charge, Lawyers), via Graph's personOrGroup facet on columns.
  // The internal name is auto-detected rather than guessed because SharePoint keeps a
  // column's original internal name forever even after its display name is changed later -
  // "Lawyers" being the export header does not guarantee "Lawyers" is the internal name, and
  // in this list it turned out not to be.
  const SCHEMA_CACHE_KEY = "crmSchemaCache:v1";
  const SCHEMA_CACHE_MAX_AGE_MS = 10 * 60 * 1000;

  async function fetchListSchema() {
    const cached = getSessionCache(SCHEMA_CACHE_KEY, SCHEMA_CACHE_MAX_AGE_MS);
    if (cached) {
      choiceOptions = cached.choiceOptions;
      personColumnNames = cached.personColumnNames;
      personColumnMultiValue = cached.personColumnMultiValue;
      resolvedFieldNames = cached.resolvedFieldNames;
      console.log("[Contacts] Using cached list schema (under 10 minutes old) - skipped the columns fetch.");
      return;
    }
    try {
      const { siteId, listId } = await resolveSiteAndListIds();
      const data = await graphGet(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/columns?$select=name,displayName,choice,personOrGroup,text,boolean,dateTime,number`);
      const columns = data.value || [];

      // Internal name can drift from what a column is actually labeled in SharePoint (a
      // rename keeps the old internal name forever), so this is logged in full rather than
      // just the matched result - if the alias/name matching below ever gets it wrong again,
      // this table shows exactly what's really on the list to fix it from.
      console.log(
        "[Contacts] Full column schema (internal name -> display name, type):",
        columns.map((col) => ({
          name: col.name,
          displayName: col.displayName,
          type: col.personOrGroup
            ? `personOrGroup${col.personOrGroup.allowMultipleSelection ? " (multi)" : " (single)"}`
            : col.choice ? "choice" : col.text ? "text" : col.number ? "number" : col.boolean ? "boolean" : col.dateTime ? "dateTime" : "other",
        }))
      );

      const byNormalizedName = {};
      columns.forEach((col) => {
        if (col.choice && Array.isArray(col.choice.choices)) {
          byNormalizedName[normalizeKey(col.name)] = col.choice.choices;
        }
      });
      const lookupChoice = (aliasKey) => {
        const candidates = FIELD_ALIASES[aliasKey] || [aliasKey];
        for (const c of candidates) {
          const hit = byNormalizedName[normalizeKey(c)];
          if (hit) return hit;
        }
        return [];
      };
      choiceOptions = {
        salutation: lookupChoice("salutation"),
        country: lookupChoice("country"),
        practiceArea: lookupChoice("practiceArea"),
        greetingCards: lookupChoice("greetingCards"),
        alumniForeign: lookupChoice("alumniForeign"),
      };

      const personColumns = columns.filter((col) => col.personOrGroup);
      // Matches by internal name first (the guesses in FIELD_ALIASES), then falls back to
      // matching by the column's current display name - catches a column that was renamed
      // in SharePoint after creation, where the internal name no longer resembles either.
      // Returns the whole column (not just its name) so the caller can also read its actual
      // allowMultipleSelection flag rather than assuming single vs. multi.
      const matchAliasColumn = (aliasKey, displayNameGuesses) => {
        const candidates = FIELD_ALIASES[aliasKey] || [aliasKey];
        const byInternal = personColumns.find((col) => candidates.some((c) => normalizeKey(c) === normalizeKey(col.name)));
        if (byInternal) return byInternal;
        return personColumns.find((col) => displayNameGuesses.some((d) => normalizeKey(d) === normalizeKey(col.displayName))) || null;
      };
      const partnerCol = matchAliasColumn("partnerInCharge", ["Partner In Charge", "PartnerInCharge"]);
      const lawyerCol = matchAliasColumn("lawyers", ["Lawyers", "Lawyer In Charge", "Lawyer"]);
      personColumnNames.partnerInCharge = partnerCol ? partnerCol.name : null;
      personColumnNames.lawyers = lawyerCol ? lawyerCol.name : null;
      personColumnMultiValue.partnerInCharge = !!(partnerCol && partnerCol.personOrGroup && partnerCol.personOrGroup.allowMultipleSelection);
      personColumnMultiValue.lawyers = !!(lawyerCol && lawyerCol.personOrGroup && lawyerCol.personOrGroup.allowMultipleSelection);
      console.log("[Contacts] Resolved Person/Group field names:", personColumnNames, "multi-value:", personColumnMultiValue);
      if (!personColumnNames.lawyers) {
        console.warn(
          "[Contacts] No column matched \"Lawyers\" by internal name or display name. Full Person/Group column list is logged above - if one of those is really meant to be Lawyers under a different label, tell me its displayName and I'll wire it in directly instead of guessing."
        );
      }

      // Same internal-name-then-display-name resolution, but for every known field, not just
      // the two Person/Group ones - this is what writeKey() actually uses, so every Add/Edit
      // save targets a verified column instead of an unverified FIELD_ALIASES guess.
      resolvedFieldNames = {};
      const unresolved = [];
      Object.keys(FIELD_ALIASES).forEach((aliasKey) => {
        const candidates = FIELD_ALIASES[aliasKey];
        const byInternal = columns.find((col) => candidates.some((c) => normalizeKey(c) === normalizeKey(col.name)));
        if (byInternal) {
          resolvedFieldNames[aliasKey] = byInternal.name;
          return;
        }
        const guesses = DISPLAY_NAME_GUESSES[aliasKey] || [];
        const byDisplay = columns.find((col) => guesses.some((d) => normalizeKey(d) === normalizeKey(col.displayName)));
        if (byDisplay) {
          resolvedFieldNames[aliasKey] = byDisplay.name;
          return;
        }
        resolvedFieldNames[aliasKey] = candidates[0];
        unresolved.push(aliasKey);
      });
      console.log("[Contacts] Resolved write target for every field:", resolvedFieldNames);
      if (unresolved.length) {
        console.warn(
          `[Contacts] Could not verify these fields against the list schema, so they're still using an unverified guess and may fail to save: ${unresolved.join(", ")}. Check the full column schema logged above for their real names.`
        );
      }
      setSessionCache(SCHEMA_CACHE_KEY, { choiceOptions, personColumnNames, personColumnMultiValue, resolvedFieldNames });
    } catch (e) {
      console.warn(
        "[Contacts] Could not read the SharePoint list's column schema - dropdown choices will be empty, and Lawyers/Partner In Charge reads and writes will fail.",
        e
      );
    }
  }

  // Lists every person in a group, flattening nested/sub-groups (transitiveMembers), for the
  // Partner In Charge and Lawyers people-pickers in the Add/Edit form. Requires the
  // GroupMember.Read.All delegated Graph permission, admin-consented.
  const GROUP_CACHE_MAX_AGE_MS = 10 * 60 * 1000;

  async function fetchGroupMembers(groupId) {
    const cacheKey = `crmGroupMembers:${groupId}`;
    const cached = getSessionCache(cacheKey, GROUP_CACHE_MAX_AGE_MS);
    if (cached) return cached;

    const people = [];
    try {
      let url = `https://graph.microsoft.com/v1.0/groups/${groupId}/transitiveMembers?$select=id,displayName,mail,userPrincipalName`;
      while (url) {
        const data = await graphGet(url);
        (data.value || []).forEach((m) => {
          if (m["@odata.type"] === "#microsoft.graph.user") {
            people.push({ id: m.id, displayName: m.displayName || "", mail: m.mail || m.userPrincipalName || "" });
          }
        });
        url = data["@odata.nextLink"] || null;
      }
    } catch (e) {
      console.warn(
        `[Contacts] Could not load members of group ${groupId} - check that GroupMember.Read.All is granted and admin-consented. This dropdown will be empty until then.`,
        e
      );
    }
    people.sort((a, b) => a.displayName.localeCompare(b.displayName));
    if (people.length) setSessionCache(cacheKey, people);
    return people;
  }

  // Manually injected people-picker entries, for testing/specific accounts that aren't (or
  // can't yet be) added to the actual AD security groups. Kept separate per dropdown so an
  // entry can be added to just one of the two without affecting the other.
  const MANUAL_PARTNER_IN_CHARGE_OPTIONS = ["la.test@skrine.com"];
  const MANUAL_LAWYER_OPTIONS = ["la.test@skrine.com", "rayhan.kass@skrine.com"];

  // Same idea, but for "view all contacts" access (see computeViewAllAccess below) rather
  // than a dropdown - accounts listed here get full view-all/export access even before (or
  // without) actually being added to the real "view all contacts" AD security group.
  const MANUAL_VIEW_ALL_CONTACTS_OPTIONS = ["jc@skrine.com"];

  function addManualPersonOptions(list, emails) {
    emails.forEach((email) => {
      if (!list.some((p) => (p.mail || "").toLowerCase() === email.toLowerCase())) {
        list.push({ id: null, displayName: email, mail: email });
      }
    });
    list.sort((a, b) => a.displayName.localeCompare(b.displayName));
  }

  // Reads the "groups" claim off the signed-in account's ID token (populated by the
  // app registration's Token configuration > Add groups claim setting) to decide whether
  // this user is in the "view all contacts" security group. If the account belongs to too
  // many groups, Azure AD omits the claim entirely (an "overage" indicator shows up instead)
  // and we fall back to the restrictive personal view rather than a Graph lookup.
  function computeViewAllAccess() {
    const claims = activeAccount && activeAccount.idTokenClaims;
    const groups = claims && claims.groups;
    const username = ((activeAccount && activeAccount.username) || "").toLowerCase();
    if (MANUAL_VIEW_ALL_CONTACTS_OPTIONS.some((email) => email.toLowerCase() === username)) {
      userCanViewAllContacts = true;
      console.log("[Contacts] userCanViewAllContacts: true (manual override for", username, ")");
      return;
    }
    console.log("[Contacts] ID token groups claim:", groups);
    console.log("[Contacts] Looking for group ID:", CONFIG.viewAllContactsGroupId);
    if (Array.isArray(groups)) {
      userCanViewAllContacts = groups.includes(CONFIG.viewAllContactsGroupId);
      console.log("[Contacts] userCanViewAllContacts:", userCanViewAllContacts);
      return;
    }
    userCanViewAllContacts = false;
    if (claims && claims._claim_names && claims._claim_names.groups) {
      console.warn(
        "[Contacts] This account belongs to too many groups for the ID token to list them directly (overage). Defaulting to personal contacts only - ask the app owner for a Graph-based fallback if this account should see all contacts."
      );
    } else {
      console.warn("[Contacts] No 'groups' claim found on the ID token at all - see console tips for what to check.");
    }
  }

  function contactMatchesUser(contact, user) {
    const byLawyerEmail = contact.lawyerEmails.some((e) => {
      const ne = normalizeEmail(e);
      return ne === normalizeEmail(user.mail) || ne === normalizeEmail(user.userPrincipalName);
    });
    const byLawyerName = contact.lawyers.some((l) => nameMatches(l, user.displayName));
    const createdByEmail = normalizeEmail(contact.createdByEmail);
    const byCreatedBy = !!createdByEmail && (createdByEmail === normalizeEmail(user.mail) || createdByEmail === normalizeEmail(user.userPrincipalName));
    return byLawyerEmail || byLawyerName || byCreatedBy;
  }

  // ---------- UI: overlay / errors ----------

  function hideLoading() {
    stopLoadingMessageLoop();
    document.getElementById("loading-state").classList.add("hidden");
    document.getElementById("app-content").classList.remove("hidden");
  }

  // The full-screen "loading-state" overlay (spinner + rotating status text) is reserved for
  // the one moment it's actually needed: initial sign-in and first data load, when there's
  // nothing on screen yet anyway. Everything after that - the Refresh button, and re-syncing
  // after a save - uses this instead: a thin sweep across the existing top accent bar, so the
  // current view (and the user's place in it) never disappears for a background data refresh.
  function showTopProgress() {
    document.getElementById("top-accent").classList.add("loading");
    const btn = document.getElementById("refresh-contacts");
    if (btn) { btn.disabled = true; btn.classList.add("spinning"); }
  }
  function hideTopProgress() {
    document.getElementById("top-accent").classList.remove("loading");
    const btn = document.getElementById("refresh-contacts");
    if (btn) { btn.disabled = false; btn.classList.remove("spinning"); }
  }

  let backgroundRefreshInFlight = false;

  async function refreshContactsInBackground() {
    if (backgroundRefreshInFlight) return;
    backgroundRefreshInFlight = true;
    showTopProgress();
    try {
      allContacts = await fetchContacts();
      renderAll();
    } catch (e) {
      console.warn("[Contacts] Background refresh failed - the view keeps showing its current data until the next successful refresh.", e);
    } finally {
      backgroundRefreshInFlight = false;
      hideTopProgress();
    }
  }

  function startLoadingMessageLoop() {
    stopLoadingMessageLoop();
    var idx = 0;
    setLoadingText(LOADING_MESSAGES[idx]);
    loadingMessageTimer = setInterval(function () {
      idx = (idx + 1) % LOADING_MESSAGES.length;
      setLoadingText(LOADING_MESSAGES[idx]);
    }, 1800);
  }
  function stopLoadingMessageLoop() {
    if (loadingMessageTimer) { clearInterval(loadingMessageTimer); loadingMessageTimer = null; }
  }

  function showError(msg, err) {
    stopLoadingMessageLoop();
    document.getElementById("loading-state").classList.add("hidden");
    var el = document.getElementById("error-banner");
    el.textContent = err && err.message ? `${msg} (${err.message})` : msg;
    el.classList.remove("hidden");
    console.error(msg, err);
  }

  function showSignInButton(err) {
    stopLoadingMessageLoop();
    var box = document.querySelector("#loading-state .loading-box");
    box.innerHTML = `
      <div class="loading-title">Sign-in needs your input this time</div>
      <button class="btn primary" id="manual-signin">Sign in with Microsoft</button>`;
    document.getElementById("manual-signin").onclick = async () => {
      try {
        // No popup - navigates the whole page to Microsoft sign-in and back. On return,
        // main() runs again from the top and ensureSignedIn() picks up the result via
        // handleRedirectPromise(), so nothing further is needed here.
        await msalInstance.loginRedirect({ scopes: CONFIG.graphScopes });
      } catch (e2) { showError("Sign-in failed.", e2); }
    };
    console.warn("Silent sign-in failed, manual sign-in required:", err);
  }

  // ---------- UI: render ----------

  function currentScopeContacts() {
    if (userCanViewAllContacts) return allContacts;
    return allContacts.filter((c) => contactMatchesUser(c, currentUser));
  }

  // The distinct value(s) a contact contributes to a column's filter checklist. Chip
  // columns (Practice Area, Greeting Card, Lawyers) contribute one entry per tag so the
  // checklist lets you pick individual tags rather than whole combinations.
  function columnValues(col, contact) {
    if (col.chip) {
      const items = col.chip(contact);
      return items.length ? items : ["(Blank)"];
    }
    const v = col.get(contact);
    const trimmed = v == null ? "" : String(v).trim();
    return [trimmed || "(Blank)"];
  }

  function getUniqueValuesForColumn(col) {
    const scope = currentScopeContacts();
    const set = new Set();
    scope.forEach((c) => columnValues(col, c).forEach((v) => set.add(v)));
    return [...set].sort((a, b) => a.localeCompare(b));
  }

  function filteredContacts() {
    let rows = currentScopeContacts();

    if (searchTerm) {
      const term = searchTerm.toLowerCase();
      rows = rows.filter((c) => COLUMNS.some((col) => String(col.get(c) || "").toLowerCase().includes(term)));
    }

    Object.keys(columnFilterValues).forEach((key) => {
      const filter = columnFilterValues[key];
      if (!filter) return;
      const col = COLUMNS.find((c) => c.key === key);
      if (!col) return;
      if (col.filterType === "range") {
        if (filter.min == null && filter.max == null) return;
        rows = rows.filter((c) => {
          const num = parseFloat(col.get(c));
          if (isNaN(num)) return false;
          if (filter.min != null && num < filter.min) return false;
          if (filter.max != null && num > filter.max) return false;
          return true;
        });
        return;
      }
      if (col.filterType === "date") {
        if (filter.year == null && filter.month == null) return;
        rows = rows.filter((c) => {
          if (!c.created) return false;
          const d = new Date(c.created);
          if (isNaN(d.getTime())) return false;
          if (filter.year != null && d.getFullYear() !== filter.year) return false;
          if (filter.month != null && d.getMonth() !== filter.month) return false;
          return true;
        });
        return;
      }
      if (!filter.size) return;
      rows = rows.filter((c) => columnValues(col, c).some((v) => filter.has(v)));
    });

    return rows;
  }

  function sortLabelsFor(col) {
    if (col.key === "id") return ["Smallest to largest", "Largest to smallest"];
    if (col.key === "created") return ["Older to newer", "Newer to older"];
    return ["A to Z", "Z to A"];
  }

  function compareContactsForSort(a, b) {
    if (!sortColumn) {
      const an = `${a.lastName} ${a.firstName}`.toLowerCase();
      const bn = `${b.lastName} ${b.firstName}`.toLowerCase();
      return an.localeCompare(bn);
    }
    const col = COLUMNS.find((c) => c.key === sortColumn);
    let result;
    if (col.key === "id") {
      result = (parseInt(a.id, 10) || 0) - (parseInt(b.id, 10) || 0);
    } else if (col.key === "created") {
      result = new Date(a.created || 0) - new Date(b.created || 0);
    } else {
      result = String(col.get(a) || "").localeCompare(String(col.get(b) || ""));
    }
    return sortDirection === "desc" ? -result : result;
  }

  // Animates a KPI card's number from its current displayed value to `to` (a short
  // count-up tween plus a pulse), instead of just snapping the text - makes it obvious
  // the cards are reacting live to the table's filters rather than being static totals.
  function animateKpiValue(id, to) {
    const el = document.getElementById(id);
    const from = parseInt(el.textContent, 10) || 0;
    if (from === to) return;
    const duration = 900;
    const startTime = performance.now();
    el.classList.remove("kpi-value-pulse");
    void el.offsetWidth; // restart the pulse animation even if it's still mid-flight
    el.classList.add("kpi-value-pulse");
    function tick(now) {
      const progress = Math.min(1, (now - startTime) / duration);
      const eased = 1 - Math.pow(1 - progress, 3);
      el.textContent = Math.round(from + (to - from) * eased);
      if (progress < 1) {
        requestAnimationFrame(tick);
      } else {
        el.textContent = to;
        el.classList.remove("kpi-value-pulse");
      }
    }
    requestAnimationFrame(tick);
  }

  function renderKpis() {
    const scope = filteredContacts();
    animateKpiValue("stat-mine", scope.length);
    animateKpiValue("stat-companies", new Set(scope.map((c) => c.companyName).filter(Boolean)).size);
    animateKpiValue("stat-countries", new Set(scope.map((c) => c.country).filter(Boolean)).size);
    animateKpiValue("stat-cards", scope.filter((c) => c.greetingCards.length > 0).length);
    animateKpiValue("stat-practice-areas", new Set(scope.flatMap((c) => c.practiceArea).filter(Boolean)).size);
    animateKpiValue("stat-alumni", scope.filter((c) => c.alumniForeign).length);
  }

  function renderList() {
    const rows = filteredContacts().slice().sort(compareContactsForSort);
    const list = document.getElementById("contacts-list");
    const empty = document.getElementById("empty-state");
    const pagination = document.getElementById("pagination");

    if (!rows.length) {
      list.innerHTML = "";
      empty.classList.remove("hidden");
      pagination.classList.add("hidden");
      return;
    }
    empty.classList.add("hidden");

    // Only the current page's rows ever hit the DOM, so the list stays fast
    // no matter how many contacts are in the underlying view.
    const totalPages = Math.max(1, Math.ceil(rows.length / PAGE_SIZE));
    if (currentPage > totalPages) currentPage = totalPages;
    if (currentPage < 1) currentPage = 1;
    const start = (currentPage - 1) * PAGE_SIZE;
    const pageRows = rows.slice(start, start + PAGE_SIZE);

    list.innerHTML = pageRows
      .map((c) => `
        <tr data-id="${escapeHtml(c.id)}">
          ${COLUMNS.map((col) => {
            if (col.chip) return `<td>${chipCell(col.chip(c), col.chipColor)}</td>`;
            const cls = col.cellClass ? ` class="${col.cellClass}"` : "";
            return `<td${cls}>${escapeHtml(col.get(c)) || "-"}</td>`;
          }).join("")}
        </tr>`)
      .join("");
    list.querySelectorAll("tr").forEach((row) => {
      row.addEventListener("click", () => {
        const contact = allContacts.find((c) => c.id === row.dataset.id);
        if (contact) openModal(contact);
      });
    });

    if (totalPages <= 1) {
      pagination.classList.add("hidden");
    } else {
      pagination.classList.remove("hidden");
      document.getElementById("pagination-info").textContent =
        `Showing ${start + 1}-${Math.min(start + PAGE_SIZE, rows.length)} of ${rows.length}`;
      document.getElementById("pagination-pages").textContent = `Page ${currentPage} of ${totalPages}`;
      document.getElementById("page-prev").disabled = currentPage <= 1;
      document.getElementById("page-next").disabled = currentPage >= totalPages;
    }
  }

  function renderAll() {
    renderKpis();
    renderList();
  }

  // Builds the header (labels + resize handles) and the per-column filter row from
  // COLUMNS, and gives every column an explicit starting width via <col> elements so
  // table-layout:fixed has something concrete to resize.
  function initTable() {
    const widths = defaultColumnWidths();
    document.getElementById("col-group").innerHTML = COLUMNS
      .map((col, i) => `<col style="width:${widths[i]}%">`)
      .join("");
    document.getElementById("table-head").innerHTML = `
      <tr>
        ${COLUMNS.map((col, i) => `
          <th>
            <div class="th-inner">
              <span class="th-label">${escapeHtml(col.label)}</span>
              <button type="button" class="th-filter-btn" data-col-index="${i}" aria-label="Sort and filter ${escapeHtml(col.label)}">
                <i class="fa-solid fa-chevron-down" aria-hidden="true"></i>
              </button>
            </div>
            <span class="th-resizer" data-col-index="${i}"></span>
          </th>`).join("")}
      </tr>`;

    document.querySelectorAll(".th-filter-btn").forEach((btn) => {
      btn.addEventListener("click", (e) => {
        e.stopPropagation();
        const col = COLUMNS[Number(btn.dataset.colIndex)];
        if (openMenuKey === col.key) { closeColumnMenu(); return; }
        openColumnMenu(col, btn);
      });
    });
    document.addEventListener("click", (e) => {
      const menu = document.getElementById("col-menu");
      if (!menu.classList.contains("hidden") && !menu.contains(e.target) && !e.target.closest(".th-filter-btn")) {
        closeColumnMenu();
      }
    });

    makeColumnsResizable();
  }

  function closeColumnMenu() {
    document.getElementById("col-menu").classList.add("hidden");
    openMenuKey = null;
    updateFilterButtonStates();
  }

  function isColumnFilterActive(col) {
    const filter = columnFilterValues[col.key];
    if (!filter) return false;
    if (col.filterType === "range") return filter.min != null || filter.max != null;
    if (col.filterType === "date") return filter.year != null || filter.month != null;
    return filter.size > 0;
  }

  const MONTH_NAMES = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

  function getDistinctYearsForColumn(col) {
    const scope = currentScopeContacts();
    const years = new Set();
    scope.forEach((c) => {
      if (!c.created) return;
      const d = new Date(c.created);
      if (!isNaN(d.getTime())) years.add(d.getFullYear());
    });
    return [...years].sort((a, b) => b - a);
  }

  function updateFilterButtonStates() {
    document.querySelectorAll(".th-filter-btn").forEach((btn) => {
      const col = COLUMNS[Number(btn.dataset.colIndex)];
      btn.classList.toggle("active", isColumnFilterActive(col));
    });
  }

  function openColumnMenu(col, anchorEl) {
    const menu = document.getElementById("col-menu");
    const [ascLabel, descLabel] = sortLabelsFor(col);
    const isRange = col.filterType === "range";
    const isDate = col.filterType === "date";
    const currentFilter = columnFilterValues[col.key];

    const sortButtons = `
      <button type="button" class="col-menu-item" data-action="sort-asc"><i class="fa-solid fa-arrow-up-short-wide" aria-hidden="true"></i>${escapeHtml(ascLabel)}</button>
      <button type="button" class="col-menu-item" data-action="sort-desc"><i class="fa-solid fa-arrow-down-wide-short" aria-hidden="true"></i>${escapeHtml(descLabel)}</button>
      <div class="col-menu-divider"></div>
      <div class="col-menu-label">Filter by ${escapeHtml(col.label)}</div>`;

    let values = [];
    if (isRange) {
      const min = currentFilter && currentFilter.min != null ? currentFilter.min : "";
      const max = currentFilter && currentFilter.max != null ? currentFilter.max : "";
      menu.innerHTML = `
        ${sortButtons}
        <div class="col-menu-range">
          <label>Min<input type="number" class="col-menu-range-input" data-bound="min" value="${escapeHtml(String(min))}" /></label>
          <label>Max<input type="number" class="col-menu-range-input" data-bound="max" value="${escapeHtml(String(max))}" /></label>
        </div>
        <div class="col-menu-footer">
          <button type="button" class="btn" data-action="clear">Clear filter</button>
          <button type="button" class="btn primary" data-action="apply">OK</button>
        </div>`;
    } else if (isDate) {
      const years = getDistinctYearsForColumn(col);
      const selYear = currentFilter && currentFilter.year != null ? currentFilter.year : "";
      const selMonth = currentFilter && currentFilter.month != null ? currentFilter.month : "";
      menu.innerHTML = `
        ${sortButtons}
        <div class="col-menu-range">
          <label>Year
            <select class="col-menu-range-input" data-bound="year">
              <option value="">Any</option>
              ${years.map((y) => `<option value="${y}" ${String(y) === String(selYear) ? "selected" : ""}>${y}</option>`).join("")}
            </select>
          </label>
          <label>Month
            <select class="col-menu-range-input" data-bound="month">
              <option value="">Any</option>
              ${MONTH_NAMES.map((name, i) => `<option value="${i}" ${String(i) === String(selMonth) ? "selected" : ""}>${escapeHtml(name)}</option>`).join("")}
            </select>
          </label>
        </div>
        <div class="col-menu-footer">
          <button type="button" class="btn" data-action="clear">Clear filter</button>
          <button type="button" class="btn primary" data-action="apply">OK</button>
        </div>`;
    } else {
      values = getUniqueValuesForColumn(col);
      const selected = currentFilter;
      menu.innerHTML = `
        ${sortButtons}
        <input type="text" class="col-menu-search" placeholder="Search values" />
        <label class="col-menu-value col-menu-selectall">
          <input type="checkbox" data-select-all ${selected ? "" : "checked"} />
          <span>Select all</span>
        </label>
        <div class="col-menu-values">
          ${values.map((v) => `
            <label class="col-menu-value">
              <input type="checkbox" data-value="${escapeHtml(v)}" ${!selected || selected.has(v) ? "checked" : ""} />
              <span>${escapeHtml(v)}</span>
            </label>`).join("")}
        </div>
        <div class="col-menu-footer">
          <button type="button" class="btn" data-action="clear">Clear filter</button>
          <button type="button" class="btn primary" data-action="apply">OK</button>
        </div>`;
    }

    const rect = anchorEl.getBoundingClientRect();
    menu.classList.remove("hidden");
    const menuWidth = menu.offsetWidth || 240;
    menu.style.top = `${rect.bottom + 6}px`;
    menu.style.left = `${Math.min(rect.left, window.innerWidth - menuWidth - 12)}px`;
    openMenuKey = col.key;
    updateFilterButtonStates();

    menu.querySelector('[data-action="sort-asc"]').addEventListener("click", () => {
      sortColumn = col.key;
      sortDirection = "asc";
      closeColumnMenu();
      renderList();
    });
    menu.querySelector('[data-action="sort-desc"]').addEventListener("click", () => {
      sortColumn = col.key;
      sortDirection = "desc";
      closeColumnMenu();
      renderList();
    });
    menu.querySelector('[data-action="clear"]').addEventListener("click", () => {
      delete columnFilterValues[col.key];
      closeColumnMenu();
      currentPage = 1;
      renderAll();
    });

    if (isRange) {
      menu.querySelector('[data-action="apply"]').addEventListener("click", () => {
        const minRaw = menu.querySelector('[data-bound="min"]').value;
        const maxRaw = menu.querySelector('[data-bound="max"]').value;
        const min = minRaw === "" ? null : parseFloat(minRaw);
        const max = maxRaw === "" ? null : parseFloat(maxRaw);
        if (min == null && max == null) {
          delete columnFilterValues[col.key];
        } else {
          columnFilterValues[col.key] = { min, max };
        }
        closeColumnMenu();
        currentPage = 1;
        renderAll();
      });
      return;
    }

    if (isDate) {
      menu.querySelector('[data-action="apply"]').addEventListener("click", () => {
        const yearRaw = menu.querySelector('[data-bound="year"]').value;
        const monthRaw = menu.querySelector('[data-bound="month"]').value;
        const year = yearRaw === "" ? null : parseInt(yearRaw, 10);
        const month = monthRaw === "" ? null : parseInt(monthRaw, 10);
        if (year == null && month == null) {
          delete columnFilterValues[col.key];
        } else {
          columnFilterValues[col.key] = { year, month };
        }
        closeColumnMenu();
        currentPage = 1;
        renderAll();
      });
      return;
    }

    menu.querySelector("[data-select-all]").addEventListener("change", (e) => {
      menu.querySelectorAll(".col-menu-values input[type=checkbox]").forEach((cb) => { cb.checked = e.target.checked; });
    });
    menu.querySelector(".col-menu-search").addEventListener("input", (e) => {
      const term = e.target.value.toLowerCase();
      menu.querySelectorAll(".col-menu-values .col-menu-value").forEach((label) => {
        label.style.display = label.textContent.toLowerCase().includes(term) ? "" : "none";
      });
    });
    menu.querySelector('[data-action="apply"]').addEventListener("click", () => {
      const checked = [...menu.querySelectorAll(".col-menu-values input[type=checkbox]:checked")].map((cb) => cb.dataset.value);
      if (checked.length === values.length) {
        delete columnFilterValues[col.key];
      } else {
        columnFilterValues[col.key] = new Set(checked);
      }
      closeColumnMenu();
      currentPage = 1;
      renderAll();
    });
  }

  function makeColumnsResizable() {
    const cols = document.querySelectorAll("#col-group col");
    document.querySelectorAll(".th-resizer").forEach((handle) => {
      handle.addEventListener("mousedown", (e) => {
        e.preventDefault();
        const col = cols[Number(handle.dataset.colIndex)];
        const startX = e.clientX;
        const startWidth = col.getBoundingClientRect().width;
        handle.classList.add("active");

        function onMove(moveEvent) {
          const next = Math.max(60, startWidth + (moveEvent.clientX - startX));
          col.style.width = `${next}px`;
        }
        function onUp() {
          handle.classList.remove("active");
          document.removeEventListener("mousemove", onMove);
          document.removeEventListener("mouseup", onUp);
        }
        document.addEventListener("mousemove", onMove);
        document.addEventListener("mouseup", onUp);
      });
    });
  }

  // ---------- write (Add / Edit) ----------

  // The internal SharePoint column name to write to for a given form field - prefers the
  // name resolved from the live list schema (see fetchListSchema), falling back to the
  // first FIELD_ALIASES guess only if the schema scan hasn't run yet or found no match.
  function writeKey(aliasKey) {
    return resolvedFieldNames[aliasKey] || FIELD_ALIASES[aliasKey][0];
  }

  // Optional fields are omitted entirely when blank rather than sent as "" or [] - Graph can
  // reject an empty string against a Choice column outright (a 400 "Invalid request" with no
  // useful detail) instead of treating it as "no value", so leaving the key out avoids
  // asserting a value for those columns at all.
  function formValuesToFields(values) {
    const fields = {
      [writeKey("firstName")]: values.firstName || "",
      [writeKey("lastName")]: values.lastName || "",
      [writeKey("email")]: values.email || "",
      [writeKey("companyName")]: values.companyName || "",
      [writeKey("country")]: values.country || "",
    };
    if (values.salutation) fields[writeKey("salutation")] = values.salutation;
    if (values.position) fields[writeKey("position")] = values.position;
    if (values.phoneNumber) fields[writeKey("phoneNumber")] = values.phoneNumber;
    if (values.practiceArea && values.practiceArea.length) fields[writeKey("practiceArea")] = values.practiceArea;
    if (values.greetingCards && values.greetingCards.length) fields[writeKey("greetingCards")] = values.greetingCards;
    if (values.alumniForeign) fields[writeKey("alumniForeign")] = values.alumniForeign;
    return fields;
  }

  // Resolves an email/UPN to a SharePoint site user ID via the classic "ensure user" REST
  // call - required before a Person/Group field can reference that person. Cached per
  // session since the same Partner In Charge / Lawyer is looked up repeatedly.
  async function ensureSharePointUserId(email) {
    if (!email) return null;
    const cacheKey = email.toLowerCase();
    if (spUserIdCache.has(cacheKey)) return spUserIdCache.get(cacheKey);
    const token = await getSharePointToken();
    const res = await fetch(`${CONFIG.sharePointResource}${CONFIG.sitePath}/_api/web/ensureuser`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json;odata=verbose",
        Accept: "application/json;odata=verbose",
      },
      body: JSON.stringify({ logonName: email }),
    });
    if (!res.ok) throw new Error(`Could not resolve SharePoint user for ${email} (${res.status}): ${await res.text()}`);
    const data = await res.json();
    const id = data.d && data.d.Id;
    spUserIdCache.set(cacheKey, id);
    return id;
  }

  async function getListEntityTypeName() {
    if (listEntityTypeName) return listEntityTypeName;
    const token = await getSharePointToken();
    const listTitle = encodeURIComponent(CONFIG.listDisplayName);
    const res = await fetch(
      `${CONFIG.sharePointResource}${CONFIG.sitePath}/_api/web/lists/getbytitle('${listTitle}')?$select=ListItemEntityTypeFullName`,
      { headers: { Authorization: `Bearer ${token}`, Accept: "application/json;odata=nometadata" } }
    );
    if (!res.ok) throw new Error(`Could not read the list schema (${res.status}): ${await res.text()}`);
    const data = await res.json();
    listEntityTypeName = data.ListItemEntityTypeFullName;
    return listEntityTypeName;
  }

  // Builds the full SharePoint REST item payload - simple fields plus the Partner In Charge /
  // Lawyers Person fields - shared by create and update. Both go through SharePoint REST
  // entirely rather than Graph: Graph's write path turned out to be just as unreliable for
  // this list as its read path was for Person/Group fields (same generic "invalidRequest"
  // 400 on PATCH as on POST), so REST is used consistently instead of splitting writes
  // between the two APIs.
  async function buildSharePointItemBody(values) {
    const entityType = await getListEntityTypeName();
    const [partnerInChargeId, lawyerIds] = await Promise.all([
      values.partnerInCharge ? ensureSharePointUserId(values.partnerInCharge) : Promise.resolve(null),
      Promise.all((values.lawyers || []).map((email) => ensureSharePointUserId(email))),
    ]);

    const graphFields = formValuesToFields(values);
    const body = { __metadata: { type: entityType } };
    // The classic SharePoint REST API (unlike Graph) wraps multi-value fields - Choice or
    // Person/Group alike - in a {results: [...]} envelope rather than a plain array.
    Object.keys(graphFields).forEach((key) => {
      const val = graphFields[key];
      body[key] = Array.isArray(val) ? { results: val } : val;
    });
    if (partnerInChargeId != null && personColumnNames.partnerInCharge) {
      body[`${personColumnNames.partnerInCharge}Id`] = personColumnMultiValue.partnerInCharge
        ? { results: [partnerInChargeId] }
        : partnerInChargeId;
    }
    const cleanLawyerIds = lawyerIds.filter((id) => id != null);
    if (personColumnNames.lawyers) {
      body[`${personColumnNames.lawyers}Id`] = personColumnMultiValue.lawyers ? { results: cleanLawyerIds } : cleanLawyerIds[0];
    }
    return body;
  }

  async function postSharePointItem(url, body, extraHeaders, actionLabel) {
    const token = await getSharePointToken();
    const res = await fetch(url, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json;odata=verbose",
        Accept: "application/json;odata=verbose",
        ...extraHeaders,
      },
      body: JSON.stringify(body),
    });
    if (!res.ok) {
      const text = await res.text();
      console.error(
        `[Contacts] SharePoint REST ${actionLabel} failed:\n` +
          `URL: POST ${url}\n` +
          `Headers: ${JSON.stringify(extraHeaders)}\n` +
          `Status: ${res.status}\n` +
          `Request body:\n${JSON.stringify(body, null, 2)}\n` +
          `Response:\n${text}`
      );
      if (res.status === 403) {
        throw new Error("Permission denied by SharePoint (403). The SharePoint API's AllSites.Write delegated permission needs to be added and admin-consented in the app registration to save changes.");
      }
      throw new Error(`Could not ${actionLabel} contact (${res.status}): ${text}`);
    }
    if (res.status === 204) return null;
    const data = await res.json().catch(() => null);
    return data && data.d ? data.d : data;
  }

  // Every SharePoint list has a built-in "Title" column - even when it's hidden or
  // repurposed in the list UI, it commonly stays required. Its internal name is always
  // literally "Title" regardless of how it's displayed, so it's not part of FIELD_ALIASES.
  // Both Person fields are required in this list, so create needs them set atomically in
  // the same request that creates the item - a create request missing them gets rejected by
  // SharePoint outright, no matter what the other fields contain.
  async function createContact(values) {
    const body = await buildSharePointItemBody(values);
    body.Title = [values.firstName, values.lastName].filter(Boolean).join(" ") || values.companyName || "Contact";
    const listTitle = encodeURIComponent(CONFIG.listDisplayName);
    const url = `${CONFIG.sharePointResource}${CONFIG.sitePath}/_api/web/lists/getbytitle('${listTitle}')/items`;
    const created = await postSharePointItem(url, body, {}, "create");
    return created && created.Id != null ? String(created.Id) : null;
  }

  async function updateContact(itemId, values) {
    const body = await buildSharePointItemBody(values);
    const listTitle = encodeURIComponent(CONFIG.listDisplayName);
    const url = `${CONFIG.sharePointResource}${CONFIG.sitePath}/_api/web/lists/getbytitle('${listTitle}')/items(${itemId})`;
    await postSharePointItem(url, body, { "X-HTTP-Method": "MERGE", "IF-MATCH": "*" }, "update");
  }

  // ---------- modal: view / add / edit ----------

  function openModal(c) {
    renderViewModal(c);
    document.getElementById("overlay").classList.remove("hidden");
  }

  function openAddContactForm() {
    renderFormModal(null);
    document.getElementById("overlay").classList.remove("hidden");
  }

  function closeModal() {
    document.getElementById("overlay").classList.add("hidden");
  }

  function renderViewModal(c) {
    document.getElementById("modal-box").innerHTML = `
      <div class="modal-top">
        <h2>${escapeHtml(contactFullName(c)) || "(No name)"}</h2>
        <button class="close" id="modal-close" aria-label="Close"><i class="fa-solid fa-xmark" aria-hidden="true"></i></button>
      </div>
      <div class="modal-sub">${escapeHtml([c.position, c.companyName].filter(Boolean).join(" - ")) || "-"}</div>
      <p class="section-label">Contact details</p>
      <dl class="field-grid">
        <dt>Email</dt><dd>${escapeHtml(c.email) || "-"}</dd>
        <dt>Phone</dt><dd>${escapeHtml(c.phoneNumber) || "-"}</dd>
        <dt>Country</dt><dd>${escapeHtml(c.country) || "-"}</dd>
        <dt>Partner</dt><dd>${escapeHtml(c.partner) || "-"}</dd>
        <dt>Partner In Charge</dt><dd>${escapeHtml(c.partnerInCharge) || "-"}</dd>
        <dt>Practice Area</dt><dd>${c.practiceArea.map((a) => `<span class="chip neutral">${escapeHtml(a)}</span>`).join("") || "-"}</dd>
        <dt>Greeting Cards</dt><dd>${c.greetingCards.map((a) => { const col = greetingChipColor(a); return `<span class="chip" style="background:${col.bg};color:${col.fg};">${escapeHtml(a)}</span>`; }).join("") || "-"}</dd>
        <dt>Alumni / Foreign</dt><dd>${escapeHtml(c.alumniForeign) || "-"}</dd>
        <dt>Lawyers</dt><dd>${c.lawyers.map((a) => `<span class="chip neutral">${escapeHtml(a)}</span>`).join("") || "-"}</dd>
        <dt>Contact Owner</dt><dd>${escapeHtml(c.contactOwner) || "-"}</dd>
        <dt>Created By</dt><dd>${escapeHtml(c.createdBy) || "-"}</dd>
        <dt>Created</dt><dd>${c.created ? new Date(c.created).toLocaleString() : "-"}</dd>
        <dt>Modified</dt><dd>${c.modified ? new Date(c.modified).toLocaleString() : "-"}</dd>
      </dl>
      <div class="close-row">
        <button type="button" class="btn" id="modal-edit-btn">Edit</button>
        <button type="button" class="btn primary" id="modal-close-2">Close</button>
      </div>`;
    document.getElementById("modal-close").addEventListener("click", closeModal);
    document.getElementById("modal-close-2").addEventListener("click", closeModal);
    document.getElementById("modal-edit-btn").addEventListener("click", () => renderFormModal(c));
  }

  function ciIncludes(list, value) {
    const lower = (value || "").toLowerCase();
    return (list || []).some((v) => (v || "").toLowerCase() === lower);
  }

  function personOptionDisplayName(options, email) {
    const match = (options || []).find((p) => (p.mail || "").toLowerCase() === (email || "").toLowerCase());
    return match ? match.displayName : email || "";
  }

  // Builds a contact object straight from the form's own values, without waiting on any
  // server round trip - the save request already told us the values are valid, so there's no
  // need to re-fetch just to know what to show. Used to update the table instantly on save;
  // a background refresh (see refreshContactsInBackground) reconciles with the server after.
  // For an edit, unset-in-the-form fields (Partner, Contact Owner, Created/Modified metadata,
  // etc.) are carried over from the existing contact rather than blanked out.
  function valuesToLocalContact(values, existingContact) {
    const base = existingContact || {
      id: null,
      partner: "",
      contactOwner: "",
      createdBy: currentUser.displayName || "",
      createdByEmail: currentUser.mail || currentUser.userPrincipalName || "",
      created: new Date().toISOString(),
      raw: {},
    };
    return {
      ...base,
      firstName: values.firstName,
      lastName: values.lastName,
      email: values.email,
      companyName: values.companyName,
      salutation: values.salutation,
      position: values.position,
      phoneNumber: values.phoneNumber,
      country: values.country,
      practiceArea: values.practiceArea,
      greetingCards: values.greetingCards,
      alumniForeign: values.alumniForeign,
      partnerInCharge: personOptionDisplayName(partnerInChargeOptions, values.partnerInCharge),
      partnerInChargeEmails: values.partnerInCharge ? [values.partnerInCharge] : [],
      lawyers: values.lawyers.map((email) => personOptionDisplayName(lawyerOptions, email)),
      lawyerEmails: values.lawyers,
      modified: new Date().toISOString(),
    };
  }

  // Wires up every dropdown-style checklist (Practice Areas, Greeting Cards, Lawyer In
  // Charge) within a form: click the trigger to open/close its panel, keep the trigger's
  // summary text in sync as checkboxes change, and close whichever panel is open when the
  // user clicks anywhere outside it.
  function wireDropdownChecklists(container) {
    container.querySelectorAll(".dropdown-trigger").forEach((trigger) => {
      const panel = document.getElementById(trigger.dataset.checklist);
      if (!panel) return;
      const search = panel.querySelector(".checklist-search");
      trigger.addEventListener("click", (e) => {
        e.stopPropagation();
        const willOpen = panel.classList.contains("hidden");
        container.querySelectorAll(".checklist").forEach((p) => p.classList.add("hidden"));
        container.querySelectorAll(".dropdown-trigger").forEach((t) => t.classList.remove("open"));
        if (willOpen) {
          panel.classList.remove("hidden");
          trigger.classList.add("open");
          if (search) {
            search.value = "";
            panel.querySelectorAll(".checklist-item").forEach((item) => item.classList.remove("filtered-out"));
            search.focus();
          }
        }
      });
      panel.addEventListener("change", () => {
        const checked = panel.querySelectorAll("input[type=checkbox]:checked").length;
        trigger.querySelector(".dropdown-trigger-text").textContent = checked ? `${checked} selected` : trigger.dataset.placeholder;
      });
      if (search) {
        search.addEventListener("click", (e) => e.stopPropagation());
        search.addEventListener("input", (e) => {
          const term = e.target.value.trim().toLowerCase();
          panel.querySelectorAll(".checklist-item").forEach((item) => {
            item.classList.toggle("filtered-out", !item.textContent.toLowerCase().includes(term));
          });
        });
      }
    });
    container.addEventListener("click", (e) => {
      if (e.target.closest(".dropdown-check")) return;
      container.querySelectorAll(".checklist").forEach((p) => p.classList.add("hidden"));
      container.querySelectorAll(".dropdown-trigger").forEach((t) => t.classList.remove("open"));
    });
  }

  function renderFormModal(contact) {
    const isNew = !contact;
    const v = contact || {
      salutation: "", firstName: "", lastName: "", email: "", companyName: "", position: "",
      phoneNumber: "", country: "", practiceArea: [], greetingCards: [], alumniForeign: "",
      partnerInChargeEmails: [], lawyerEmails: [],
    };

    const textField = (name, label, opts) => {
      opts = opts || {};
      const req = opts.required ? ` <span class="req">*</span>` : "";
      return `<div class="form-field">
          <label for="cf-${name}">${escapeHtml(label)}${req}</label>
          <input id="cf-${name}" name="${name}" type="${opts.type || "text"}" value="${escapeHtml(String(v[name] ?? ""))}" />
        </div>`;
    };

    // options: array of strings, or array of {value,label} (used for people-pickers).
    // Any current value that isn't in the options list (stale AD/SharePoint choice data)
    // is kept as an extra entry instead of silently dropped.
    const selectField = (name, label, options, selectedValue, opts) => {
      opts = opts || {};
      const req = opts.required ? ` <span class="req">*</span>` : "";
      const norm = options.map((o) => (typeof o === "string" ? { value: o, label: o } : o));
      const extra = selectedValue && !ciIncludes(norm.map((o) => o.value), selectedValue)
        ? [{ value: selectedValue, label: selectedValue }]
        : [];
      const optionTags = norm.concat(extra)
        .map((o) => `<option value="${escapeHtml(o.value)}" ${o.value.toLowerCase() === (selectedValue || "").toLowerCase() ? "selected" : ""}>${escapeHtml(o.label)}</option>`)
        .join("");
      return `<div class="form-field">
          <label for="cf-${name}">${escapeHtml(label)}${req}</label>
          <select id="cf-${name}" name="${name}">
            <option value="">-- Select --</option>
            ${optionTags}
          </select>
          ${!options.length ? `<span class="hint">No choices were found for this column on the SharePoint list</span>` : ""}
        </div>`;
    };

    // Renders as a closed-by-default dropdown (like the table's column filter menus) rather
    // than an always-open checkbox grid - the checkboxes themselves are unchanged underneath,
    // so FormData.getAll(name) still works the same regardless of whether the panel is open.
    const checklistField = (name, label, options, selectedValues, opts) => {
      opts = opts || {};
      const req = opts.required ? ` <span class="req">*</span>` : "";
      const norm = options.map((o) => (typeof o === "string" ? { value: o, label: o } : o));
      const normValues = norm.map((o) => o.value);
      const extras = (selectedValues || []).filter((sv) => !ciIncludes(normValues, sv)).map((sv) => ({ value: sv, label: sv }));
      const items = norm.concat(extras)
        .map((o) => `
          <label class="checklist-item">
            <input type="checkbox" name="${name}" value="${escapeHtml(o.value)}" ${ciIncludes(selectedValues, o.value) ? "checked" : ""} />
            <span>${escapeHtml(o.label)}</span>
          </label>`)
        .join("");
      const placeholder = opts.placeholder || "Select...";
      const summary = (selectedValues || []).length ? `${selectedValues.length} selected` : placeholder;
      return `<div class="form-field form-field-wide">
          <label>${escapeHtml(label)}${req}</label>
          <div class="dropdown-check">
            <button type="button" class="dropdown-trigger" data-checklist="cf-${name}" data-placeholder="${escapeHtml(placeholder)}">
              <span class="dropdown-trigger-text">${escapeHtml(summary)}</span>
              <i class="fa-solid fa-chevron-down" aria-hidden="true"></i>
            </button>
            <div class="checklist hidden" id="cf-${name}">
              <input type="text" class="checklist-search" placeholder="Search..." aria-label="Search ${escapeHtml(label)}" />
              <div class="checklist-items">${items || `<div class="checklist-empty">No options available</div>`}</div>
            </div>
          </div>
          ${opts.hint ? `<span class="hint">${escapeHtml(opts.hint)}</span>` : ""}
        </div>`;
    };

    document.getElementById("modal-box").innerHTML = `
      <div class="modal-top">
        <h2>${isNew ? "Add Contact" : "Edit Contact"}</h2>
        <button class="close" id="modal-close" aria-label="Close"><i class="fa-solid fa-xmark" aria-hidden="true"></i></button>
      </div>
      <form id="contact-form" class="contact-form" novalidate>
        <div class="form-grid">
          ${textField("firstName", "First Name", { required: true })}
          ${textField("lastName", "Last Name", { required: true })}
          ${textField("email", "Email", { type: "email", required: true })}
          ${selectField("salutation", "Salutation", choiceOptions.salutation, v.salutation)}
          ${textField("companyName", "Company Name", { required: true })}
          ${textField("position", "Position")}
          ${textField("phoneNumber", "Phone Number")}
          ${selectField("country", "Country", choiceOptions.country, v.country, { required: true })}
          ${checklistField("practiceArea", "Practice Areas", choiceOptions.practiceArea, v.practiceArea, { required: true })}
          ${checklistField("greetingCards", "Greeting Cards", choiceOptions.greetingCards, v.greetingCards)}
          ${selectField("alumniForeign", "Alumni / Foreign Law Firm (if applicable)", choiceOptions.alumniForeign, v.alumniForeign)}
          ${selectField(
            "partnerInCharge",
            "Partner In Charge",
            partnerInChargeOptions.map((p) => ({ value: p.mail, label: p.displayName })),
            (v.partnerInChargeEmails && v.partnerInChargeEmails[0]) || "",
            { required: true }
          )}
          ${checklistField(
            "lawyers",
            "Lawyer In Charge",
            lawyerOptions.map((p) => ({ value: p.mail, label: p.displayName })),
            v.lawyerEmails,
            { required: true, hint: "Select one or more" }
          )}
        </div>
        <div class="form-error hidden" id="form-error"></div>
        <div class="close-row">
          <button type="button" class="btn" id="form-cancel">Cancel</button>
          <button type="submit" class="btn primary" id="form-save">${isNew ? "Add Contact" : "Save Changes"}</button>
        </div>
      </form>`;

    document.getElementById("modal-close").addEventListener("click", closeModal);
    document.getElementById("form-cancel").addEventListener("click", () => {
      if (isNew) closeModal();
      else renderViewModal(contact);
    });
    wireDropdownChecklists(document.getElementById("contact-form"));
    document.getElementById("contact-form").addEventListener("submit", async (e) => {
      e.preventDefault();
      const fd = new FormData(e.target);
      const values = {
        firstName: (fd.get("firstName") || "").trim(),
        lastName: (fd.get("lastName") || "").trim(),
        email: (fd.get("email") || "").trim(),
        salutation: fd.get("salutation") || "",
        companyName: (fd.get("companyName") || "").trim(),
        position: (fd.get("position") || "").trim(),
        phoneNumber: (fd.get("phoneNumber") || "").trim(),
        country: fd.get("country") || "",
        practiceArea: fd.getAll("practiceArea"),
        greetingCards: fd.getAll("greetingCards"),
        alumniForeign: fd.get("alumniForeign") || "",
        partnerInCharge: fd.get("partnerInCharge") || "",
        lawyers: fd.getAll("lawyers"),
      };

      const missing = [];
      if (!values.firstName) missing.push("First Name");
      if (!values.lastName) missing.push("Last Name");
      if (!values.email) missing.push("Email");
      if (!values.companyName) missing.push("Company Name");
      if (!values.country) missing.push("Country");
      if (!values.practiceArea.length) missing.push("Practice Areas");
      if (!values.partnerInCharge) missing.push("Partner In Charge");
      if (!values.lawyers.length) missing.push("Lawyer In Charge");

      const errorBox = document.getElementById("form-error");
      if (missing.length) {
        errorBox.textContent = `Please fill in: ${missing.join(", ")}.`;
        errorBox.classList.remove("hidden");
        return;
      }

      const saveBtn = document.getElementById("form-save");
      errorBox.classList.add("hidden");
      saveBtn.disabled = true;
      saveBtn.textContent = "Saving...";
      try {
        if (isNew) {
          const newId = await createContact(values);
          const localContact = valuesToLocalContact(values, null);
          localContact.id = newId;
          allContacts = [localContact, ...allContacts];
        } else {
          await updateContact(contact.id, values);
          const localContact = valuesToLocalContact(values, contact);
          allContacts = allContacts.map((c) => (c.id === contact.id ? localContact : c));
        }
        // The table reflects the save immediately, from the values just submitted - no need
        // to wait on a re-fetch (which was also masking a separate bug: SharePoint doesn't
        // always make a write visible to an immediate read-back right away). A background
        // refresh (thin top bar, not the full-screen loader) reconciles afterward.
        closeModal();
        currentPage = 1;
        renderAll();
        refreshContactsInBackground();
      } catch (err) {
        errorBox.textContent = err.message;
        errorBox.classList.remove("hidden");
        saveBtn.disabled = false;
        saveBtn.textContent = isNew ? "Add Contact" : "Save Changes";
      }
    });
  }

  function renderUserChip() {
    document.getElementById("avatar").textContent = initialsOf(currentUser.displayName);
    document.getElementById("user-name").textContent = currentUser.displayName || "";
    document.getElementById("user-email").textContent = currentUser.mail || currentUser.userPrincipalName || "";
  }

  let searchDebounceTimer = null;

  function wireControls() {
    document.getElementById("search").addEventListener("input", (e) => {
      searchTerm = e.target.value.trim();
      currentPage = 1;
      // Debounced so a full re-scan + re-render (table and KPI cards both) only happens
      // once typing pauses briefly, instead of on every single keystroke.
      clearTimeout(searchDebounceTimer);
      searchDebounceTimer = setTimeout(renderAll, 200);
    });
    document.getElementById("clear-filters").addEventListener("click", () => {
      clearTimeout(searchDebounceTimer);
      searchTerm = "";
      currentPage = 1;
      columnFilterValues = {};
      sortColumn = "id";
      sortDirection = "desc";
      document.getElementById("search").value = "";
      updateFilterButtonStates();
      renderAll();
    });
    document.getElementById("refresh-contacts").addEventListener("click", () => {
      sessionStorage.removeItem("crmSiteListIds:v1");
      refreshContactsInBackground();
    });
    document.getElementById("page-prev").addEventListener("click", () => { currentPage--; renderList(); });
    document.getElementById("page-next").addEventListener("click", () => { currentPage++; renderList(); });
    document.getElementById("add-contact-btn").addEventListener("click", openAddContactForm);

    document.getElementById("overlay").addEventListener("click", (e) => { if (e.target.id === "overlay") closeModal(); });

    var trigger = document.getElementById("user-trigger");
    var menu = document.getElementById("user-menu");
    trigger.addEventListener("click", () => {
      var open = !menu.classList.contains("hidden");
      menu.classList.toggle("hidden", open);
      trigger.setAttribute("aria-expanded", String(!open));
    });
    document.addEventListener("click", (e) => {
      if (!document.getElementById("navbar-user").contains(e.target)) {
        menu.classList.add("hidden");
        trigger.setAttribute("aria-expanded", "false");
      }
    });
    document.getElementById("sign-out-btn").addEventListener("click", () => {
      msalInstance.logoutRedirect({ postLogoutRedirectUri: window.location.href.split("#")[0].split("?")[0] });
    });
  }

  // ---------- Copilot chat ----------

  let copilotOpen = false;
  let copilotHistory = []; // [{role: "user"|"assistant", content: string}, ...] - this session only

  function copilotConfigured() {
    return !CONFIG.azureOpenAI.endpoint.includes("YOUR-RESOURCE-NAME") && !CONFIG.azureOpenAI.deploymentName.includes("YOUR-DEPLOYMENT-NAME");
  }

  // Grounds the assistant in the currently loaded contacts without sending every individual
  // record (expensive in tokens, and needlessly exposes more data than most questions need):
  // always include aggregate stats (counts by country/practice area/etc., the same numbers
  // the KPI cards and chart already show), and only pull in specific contacts' details when
  // the user's own message plausibly names one (a simple keyword match against the same
  // fields the search box already searches).
  function buildCopilotContext(userMessage) {
    const scope = allContacts;
    const topEntries = (map, n) =>
      [...map.entries()].sort((a, b) => b[1] - a[1]).slice(0, n).map(([k, v]) => `${k}: ${v}`).join(", ") || "(none)";

    const byCountry = new Map();
    const byPracticeArea = new Map();
    scope.forEach((c) => {
      if (c.country) byCountry.set(c.country, (byCountry.get(c.country) || 0) + 1);
      (c.practiceArea.length ? c.practiceArea : ["(Unspecified)"]).forEach((p) => byPracticeArea.set(p, (byPracticeArea.get(p) || 0) + 1));
    });
    const distinctCompanies = new Set(scope.map((c) => c.companyName).filter(Boolean)).size;

    const summary = [
      `Total contacts visible to this user: ${scope.length}.`,
      `Distinct companies: ${distinctCompanies}.`,
      `Distinct countries: ${byCountry.size}. Contacts by country (top 10): ${topEntries(byCountry, 10)}.`,
      `Contacts by practice area (top 10): ${topEntries(byPracticeArea, 10)}.`,
      `Contacts on a greeting card list: ${scope.filter((c) => c.greetingCards.length > 0).length}.`,
    ].join("\n");

    const words = (userMessage || "").toLowerCase().split(/\s+/).filter((w) => w.length >= 3);
    const matches = words.length
      ? scope
          .filter((c) => words.some((w) => [c.firstName, c.lastName, c.companyName, c.position, c.country].join(" ").toLowerCase().includes(w)))
          .slice(0, 5)
      : [];
    const matchesText = matches.length
      ? matches
          .map((c) => `- ${contactFullName(c)}, ${c.position || "n/a"} at ${c.companyName || "n/a"} (${c.country || "n/a"}); email ${c.email || "n/a"}; practice area(s): ${c.practiceArea.join(", ") || "n/a"}`)
          .join("\n")
      : null;

    return matchesText ? `${summary}\n\nContacts that might be relevant to this question:\n${matchesText}` : summary;
  }

  async function callCopilot(userMessage) {
    const token = await getAzureOpenAIToken();
    const url = `${CONFIG.azureOpenAI.endpoint}/responses`;
    const systemPrompt =
      "You are the Copilot assistant embedded in Skrine's client contact management system. " +
      "Answer questions about the firm's contacts using only the data given to you below - if something " +
      "isn't in it, say you don't have that detail rather than guessing. Be concise.\n\n" +
      buildCopilotContext(userMessage);
    const input = [{ role: "system", content: systemPrompt }, ...copilotHistory, { role: "user", content: userMessage }];

    const res = await fetch(url, {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify({ model: CONFIG.azureOpenAI.deploymentName, input }),
    });
    if (!res.ok) {
      const text = await res.text();
      if (res.status === 403 || res.status === 401) {
        throw new Error("Permission denied by Azure AI Foundry. Check that this account has the \"Cognitive Services OpenAI User\" role on the skrine-copilot resource, and that the app registration's Azure Cognitive Services permission is admin-consented.");
      }
      throw new Error(`Copilot request failed (${res.status}): ${text}`);
    }
    const data = await res.json();
    const message = (data.output || []).find((o) => o.type === "message");
    const textPart = message && (message.content || []).find((c) => c.type === "output_text");
    return (textPart && textPart.text) || "(No response.)";
  }

  function renderCopilotMessages() {
    const box = document.getElementById("copilot-messages");
    box.innerHTML = copilotHistory
      .map((m) => `<div class="copilot-msg ${m.role}">${escapeHtml(m.content)}</div>`)
      .join("");
    box.scrollTop = box.scrollHeight;
  }

  function setCopilotThinking(isThinking) {
    document.getElementById("copilot-thinking").classList.toggle("hidden", !isThinking);
    document.getElementById("copilot-send").disabled = isThinking;
    document.getElementById("copilot-input").disabled = isThinking;
  }

  // ---------- chat-triggered contact export (Excel) ----------
  // Restricted to the same security group as "view all contacts" - exports a single named
  // contact (e.g. "export John Tan's contact into excel"), not the whole list.

  let copilotPendingExportCandidates = null;

  function looksLikeContactExportRequest(text) {
    const t = (text || "").toLowerCase();
    return t.includes("export") && (t.includes("excel") || t.includes("xlsx") || t.includes("spreadsheet"));
  }

  function findContactMatchesInText(text, contacts) {
    const t = (text || "").toLowerCase();
    return contacts.filter((c) => {
      const first = (c.firstName || "").trim().toLowerCase();
      const last = (c.lastName || "").trim().toLowerCase();
      if (!first && !last) return false;
      const full = `${first} ${last}`.trim();
      if (full.length > 2 && t.includes(full)) return true;
      if (first && last && t.includes(first) && t.includes(last)) return true;
      return false;
    });
  }

  function contactDisplayLabel(c) {
    const name = contactFullName(c);
    return c.companyName ? `${name} (${c.companyName})` : name;
  }

  function describeExportCandidates(matches) {
    const list = matches.map((c) => `- ${contactDisplayLabel(c)}`).join("\n");
    return `I found a few contacts with that name - which one did you mean?\n${list}\n\nReply with their company name, or say "first" / "second", etc.`;
  }

  function resolveExportDisambiguation(text, candidates) {
    const t = (text || "").toLowerCase();
    const ordinalWords = ["first", "second", "third", "fourth", "fifth"];
    for (let i = 0; i < ordinalWords.length && i < candidates.length; i++) {
      if (t.includes(ordinalWords[i]) || t.trim() === String(i + 1)) return candidates[i];
    }
    const byCompany = candidates.filter((c) => c.companyName && t.includes(c.companyName.toLowerCase()));
    if (byCompany.length === 1) return byCompany[0];
    const byEmail = candidates.filter((c) => c.email && t.includes(c.email.toLowerCase()));
    if (byEmail.length === 1) return byEmail[0];
    return null;
  }

  function exportContactsToExcel(contacts, fileNameHint) {
    const header = COLUMNS.map((col) => col.label);
    const rows = contacts.map((c) => COLUMNS.map((col) => (col.chip ? col.chip(c).join(", ") : (col.get(c) == null ? "" : col.get(c)))));
    const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
    ws["!cols"] = COLUMNS.map((col) => ({ wch: Math.max(10, col.label.length + 4) }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Contacts");
    const safeName = (fileNameHint || "contact").replace(/[^a-z0-9]+/gi, "-").replace(/^-+|-+$/g, "").toLowerCase();
    XLSX.writeFile(wb, `skrine-${safeName || "contact"}-contact.xlsx`);
  }

  function finishContactExport(contact) {
    exportContactsToExcel([contact], `${contact.firstName}-${contact.lastName}`);
    copilotHistory.push({ role: "assistant", content: `Done - I've exported ${contactDisplayLabel(contact)}'s contact to an Excel file. Check your downloads.` });
    renderCopilotMessages();
  }

  async function sendCopilotMessage() {
    const input = document.getElementById("copilot-input");
    const text = input.value.trim();
    if (!text) return;
    input.value = "";

    if (copilotPendingExportCandidates) {
      const candidates = copilotPendingExportCandidates;
      const resolved = resolveExportDisambiguation(text, candidates);
      if (resolved) {
        copilotPendingExportCandidates = null;
        copilotHistory.push({ role: "user", content: text });
        renderCopilotMessages();
        finishContactExport(resolved);
        return;
      }
      copilotPendingExportCandidates = null;
    }

    if (looksLikeContactExportRequest(text)) {
      copilotHistory.push({ role: "user", content: text });
      renderCopilotMessages();
      if (!userCanViewAllContacts) {
        copilotHistory.push({ role: "assistant", content: "Sorry, exporting a contact to Excel isn't available for your account - it's limited to the \"view all contacts\" access group." });
        renderCopilotMessages();
        return;
      }
      const matches = findContactMatchesInText(text, currentScopeContacts());
      if (!matches.length) {
        copilotHistory.push({ role: "assistant", content: "I couldn't find a contact matching that name - could you check the spelling or try their full name?" });
        renderCopilotMessages();
        return;
      }
      if (matches.length === 1) {
        finishContactExport(matches[0]);
        return;
      }
      copilotPendingExportCandidates = matches;
      copilotHistory.push({ role: "assistant", content: describeExportCandidates(matches) });
      renderCopilotMessages();
      return;
    }

    if (!copilotConfigured()) {
      copilotHistory.push({ role: "user", content: text });
      copilotHistory.push({ role: "assistant", content: "Copilot isn't set up yet - the Azure OpenAI endpoint and deployment name in CONFIG.azureOpenAI still need to be filled in." });
      renderCopilotMessages();
      return;
    }

    copilotHistory.push({ role: "user", content: text });
    renderCopilotMessages();
    setCopilotThinking(true);
    try {
      const reply = await callCopilot(text);
      copilotHistory.push({ role: "assistant", content: reply });
    } catch (e) {
      copilotHistory.push({ role: "assistant", content: `Sorry, something went wrong: ${e.message}` });
      console.error("[Copilot] Request failed:", e);
    } finally {
      setCopilotThinking(false);
      renderCopilotMessages();
    }
  }

  let copilotLauncherOpen = false;

  function closeCopilotLauncher() {
    copilotLauncherOpen = false;
    document.getElementById("copilot-launcher").classList.add("hidden");
    document.getElementById("copilot-fab").classList.remove("open");
  }

  function openCopilotPanel() {
    copilotOpen = true;
    document.getElementById("copilot-panel").classList.remove("hidden");
    document.getElementById("copilot-fab").classList.add("open");
    if (!copilotHistory.length) {
      copilotHistory.push({ role: "assistant", content: "Hi, I'm your Skrine CRM Copilot. Ask me about your contacts - e.g. \"how many contacts do we have in Malaysia\" or \"what practice areas are most common\"." });
      renderCopilotMessages();
    }
    document.getElementById("copilot-input").focus();
  }

  function toggleCopilot() {
    if (copilotOpen) {
      copilotOpen = false;
      document.getElementById("copilot-panel").classList.add("hidden");
      document.getElementById("copilot-fab").classList.remove("open");
    } else {
      openCopilotPanel();
    }
  }

  // The FAB opens the quick-action launcher first (matching the Microsoft 365 Copilot
  // pattern) rather than jumping straight into the full chat - "Open chat", a menu item, or
  // the inline quick-ask box all transition from there into the same chat panel used
  // throughout the rest of this file.
  function toggleCopilotLauncher() {
    if (copilotOpen) { toggleCopilot(); return; }
    copilotLauncherOpen = !copilotLauncherOpen;
    document.getElementById("copilot-launcher").classList.toggle("hidden", !copilotLauncherOpen);
    document.getElementById("copilot-fab").classList.toggle("open", copilotLauncherOpen);
  }

  function openCopilotChatWithPrompt(promptText) {
    closeCopilotLauncher();
    if (!copilotOpen) openCopilotPanel();
    if (promptText) {
      document.getElementById("copilot-input").value = promptText;
      sendCopilotMessage();
    }
  }

  function wireCopilot() {
    document.getElementById("copilot-fab").addEventListener("click", toggleCopilotLauncher);
    document.getElementById("copilot-close").addEventListener("click", toggleCopilot);
    document.getElementById("copilot-send").addEventListener("click", sendCopilotMessage);
    document.getElementById("copilot-input").addEventListener("keydown", (e) => {
      if (e.key === "Enter" && !e.shiftKey) {
        e.preventDefault();
        sendCopilotMessage();
      }
    });

    document.querySelectorAll(".copilot-menu-item").forEach((btn) => {
      btn.addEventListener("click", () => openCopilotChatWithPrompt(btn.dataset.prompt));
    });
    document.getElementById("copilot-open-chat").addEventListener("click", () => openCopilotChatWithPrompt(null));
    document.getElementById("copilot-quick-ask").addEventListener("submit", (e) => {
      e.preventDefault();
      const input = document.getElementById("copilot-quick-input");
      const text = input.value.trim();
      input.value = "";
      if (text) openCopilotChatWithPrompt(text);
    });

    document.addEventListener("click", (e) => {
      if (!copilotLauncherOpen) return;
      if (e.target.closest("#copilot-launcher") || e.target.closest("#copilot-fab")) return;
      closeCopilotLauncher();
    });
  }

  // ---------- bootstrap ----------

  async function bootstrapData() {
    computeViewAllAccess();
    setLoadingText("Loading your data...");
    // The profile lookup and both group-member fetches don't depend on the list schema or
    // on each other - they're kicked off immediately and only awaited once everything else
    // is ready, instead of waiting for the schema -> contacts chain to finish first. That
    // chain (schema must resolve before contacts, for its $select) is the only real
    // dependency here; everything else was previously running needlessly one-after-another.
    const meP = graphGet("https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName");
    const partnersP = fetchGroupMembers(CONFIG.partnerInChargeGroupId);
    const lawyersP = fetchGroupMembers(CONFIG.lawyersGroupId);

    await fetchListSchema();
    allContacts = await fetchContacts();

    // Best-effort: each of these already logs and degrades gracefully on its own failure
    // (empty people-picker), so this won't take down the whole app if e.g.
    // GroupMember.Read.All hasn't been consented yet.
    const [me, partners, lawyers] = await Promise.all([meP, partnersP, lawyersP]);
    currentUser = me;
    partnerInChargeOptions = partners;
    lawyerOptions = lawyers;
    addManualPersonOptions(partnerInChargeOptions, MANUAL_PARTNER_IN_CHARGE_OPTIONS);
    addManualPersonOptions(lawyerOptions, MANUAL_LAWYER_OPTIONS);
    hideLoading();
    renderUserChip();
    initTable();
    wireControls();
    wireCopilot();
    renderAll();
  }

  async function main() {
    startLoadingMessageLoop();
    try {
      await ensureSignedIn();
    } catch (e) {
      showSignInButton(e);
      return;
    }
    try {
      await bootstrapData();
    } catch (e) {
      showError("Could not load contacts. Check console / raw fields for details.", e);
    }
  }

  main();
})();
</script>
</body>
</html>
