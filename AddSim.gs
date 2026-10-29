/**
 * Carmel Physics — one form adds a simulation.
 *
 * Setup (use the Google account that will stay at the school):
 * 1. Open the catalog Sheet → Extensions → Apps Script. Paste this file.
 * 2. Edit GITHUB_REPO and PAGES_BASE.
 * 3. Put a GitHub token and a staff code in setSecrets, run setSecrets once, then delete those two strings.
 * 4. Deploy → New deployment → Web app.
 *    Execute as: Me
 *    Who has access: Anyone
 * 5. Open the web app link. That is the add page. Bookmark it for staff.
 *
 * GitHub token: Fine-grained, this repository only, Contents: Read and write.
 * The hub already reads the Sheet. Reload the hub about a minute after saving.
 */
var SHEET_ID = "1MXX-mM_HFfo_gaP4yca4yO_u11LjIeih0lDI-JRiZ9g";
var SHEET_GID = 620668479;
var GITHUB_REPO = "hajime-dot/c-physics";
var PAGES_BASE = "https://hajime-dot.github.io/c-physics/sims/";

function setSecrets() {
  PropertiesService.getScriptProperties().setProperties({
    GITHUB_TOKEN: "paste-token-here",
    STAFF_CODE: "paste-a-code-here",
  });
}

function doGet() {
  return HtmlService.createHtmlOutput(formHtml_())
    .setTitle("Add a simulation")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function addSim(code, title, topic, blurb, html) {
  var props = PropertiesService.getScriptProperties();
  if (!code || code !== props.getProperty("STAFF_CODE")) {
    throw new Error("Staff code is wrong.");
  }
  title = String(title || "").trim();
  topic = String(topic || "").trim();
  blurb = String(blurb || "").trim();
  html = String(html || "");
  if (!title || title.length > 160) throw new Error("Enter a title.");
  if (!topic) throw new Error("Choose a topic.");
  if (html.length < 40 || html.indexOf("<") === -1) throw new Error("Paste the full HTML file.");
  if (html.length > 900000) throw new Error("That HTML is over 900 KB.");

  var token = props.getProperty("GITHUB_TOKEN");
  if (!token || token === "paste-token-here") throw new Error("GitHub token is not set. Run setSecrets.");

  var name = slug_(title) + "-" + Utilities.getUuid().slice(0, 8) + ".html";
  var response = UrlFetchApp.fetch(
    "https://api.github.com/repos/" + GITHUB_REPO + "/contents/sims/" + encodeURIComponent(name),
    {
      method: "put",
      contentType: "application/json",
      headers: {
        Authorization: "Bearer " + token,
        Accept: "application/vnd.github+json",
        "User-Agent": "carmel-physics-hub",
      },
      payload: JSON.stringify({
        message: "Add sim: " + title,
        content: Utilities.base64Encode(html, Utilities.Charset.UTF_8).replace(/\s+/g, ""),
      }),
      muteHttpExceptions: true,
    },
  );
  var status = response.getResponseCode();
  if (status < 200 || status >= 300) {
    throw new Error("GitHub refused the file (" + status + "). " + response.getContentText().slice(0, 180));
  }

  var pageUrl = PAGES_BASE + name;
  var sheet = sheet_();
  var last = sheet.getLastRow();
  var nextSort = 10;
  if (last > 1) {
    var values = sheet.getRange(2, 9, last - 1, 1).getValues();
    for (var i = 0; i < values.length; i++) {
      var n = Number(values[i][0]);
      if (!isNaN(n) && n >= nextSort) nextSort = n + 10;
    }
  }
  sheet.appendRow(["", title, topic, blurb, pageUrl, "", "", new Date(), nextSort]);
  return pageUrl;
}

function sheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === SHEET_GID) return sheets[i];
  }
  return sheets[0];
}

function slug_(value) {
  var out = String(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-|-$/g, "");
  return (out || "sim").slice(0, 40);
}

function formHtml_() {
  return (
    "<!DOCTYPE html><html><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'>" +
    "<style>body{font:16px/1.4 system-ui,sans-serif;background:#f4f5f7;color:#1a1a1a;margin:0}" +
    "main{max-width:720px;margin:0 auto;padding:20px}" +
    "h1{font-size:22px;margin:0 0 8px}p{color:#5c5c5c;margin:0 0 16px}" +
    "label{display:block;font-weight:600;margin:12px 0 6px}" +
    "input,select,textarea{width:100%;box-sizing:border-box;font:inherit;padding:12px;border:1px solid #e2e4e8;border-radius:10px}" +
    "textarea{min-height:220px;font-family:ui-monospace,Menlo,Consolas,monospace;font-size:13px}" +
    "button{margin-top:14px;background:#990011;color:#fff;border:0;border-radius:10px;padding:12px 18px;font-weight:700;cursor:pointer}" +
    "button:disabled{background:#c9c9c9}#msg{min-height:1.4em}</style></head><body><main>" +
    "<h1>Add a simulation</h1><p>Save first, then open the link and check the simulation. It stays off the hub until you type yes in the Sheet column approved.</p>" +
    "<label for='code'>Staff code</label><input id='code' type='password' autocomplete='current-password'>" +
    "<label for='title'>Title</label><input id='title'>" +
    "<label for='topic'>Topic</label><select id='topic'>" +
    "<option>Mechanics</option><option>Waves</option><option>Electricity and Magnetism</option>" +
    "<option>Heat and Gases</option><option>Atomic World</option><option>Energy and Use of Energy</option></select>" +
    "<label for='blurb'>Short note (optional)</label><input id='blurb'>" +
    "<label for='file'>HTML file</label><input id='file' type='file' accept='.html,text/html'>" +
    "<label for='html'>Or paste the HTML</label><textarea id='html' spellcheck='false'></textarea>" +
    "<button id='go' type='button' onclick='onSave()'>Save to hub</button><p id='msg'></p>" +
    "<script>" +
    "document.getElementById('file').addEventListener('change',function(e){" +
    "var f=e.target.files&&e.target.files[0];if(!f)return;var r=new FileReader();" +
    "r.onload=function(){document.getElementById('html').value=String(r.result||'');" +
    "var t=document.getElementById('title');if(!t.value)t.value=f.name.replace(/\\.html?$/i,'').replace(/[-_]+/g,' ');};" +
    "r.readAsText(f);});" +
    "function onSave(){var btn=document.getElementById('go');btn.disabled=true;" +
    "document.getElementById('msg').textContent='Saving…';" +
    "google.script.run.withSuccessHandler(function(url){btn.disabled=false;" +
    "document.getElementById('msg').textContent='Saved. Wait a minute, then open this link and check it. It is not on the hub yet. '+url+' — When it looks right, set approved to yes on that Sheet row.';})" +
    ".withFailureHandler(function(err){btn.disabled=false;document.getElementById('msg').textContent=err.message;})" +
    ".addSim(document.getElementById('code').value,document.getElementById('title').value," +
    "document.getElementById('topic').value,document.getElementById('blurb').value,document.getElementById('html').value);}" +
    "</script></main></body></html>"
  );
}
