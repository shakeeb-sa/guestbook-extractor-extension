let existingDomains = [];
let masterSheetData = {}; // Stores the full content of the uploaded file

// 1. Handle File Upload & Parsing
document.getElementById("excelInput").addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });

    const tempSet = new Set();
    masterSheetData = {}; // Reset

    // Iterate ALL sheets
    workbook.SheetNames.forEach((sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      // Store raw data for Merging later (Skip header row 0 usually, but we keep structure)
      // We filter out empty rows
      const cleanRows = rows.filter((r) => r.length > 0);
      masterSheetData[sheetName] = cleanRows;

      // Build Exclusion List
      cleanRows.forEach((row) => {
        row.forEach((cell) => {
          if (typeof cell === "string" && cell.includes(".")) {
            try {
              const urlToParse = cell.startsWith("http")
                ? cell
                : `http://${cell}`;
              const urlObj = new URL(urlToParse);
              let host = urlObj.hostname.replace(/^www\./, "");
              tempSet.add(host.toLowerCase());
            } catch (err) {}
          }
        });
      });
    });

    existingDomains = Array.from(tempSet);
    document.getElementById(
      "fileStatus"
    ).textContent = `Loaded DB: ${existingDomains.length} domains excluded. Ready to Merge.`;
    document.getElementById("fileStatus").style.color = "green";
  };
  reader.readAsArrayBuffer(file);
});

// 2. Run Button Logic
document.getElementById("run").addEventListener("click", async () => {
  const blacklistEnabled = document.getElementById("blacklistToggle").checked;
  const mergeEnabled = document.getElementById("mergeToggle").checked;

  // FIX 1: Safety Check for Popup Closure
  if (mergeEnabled && Object.keys(masterSheetData).length === 0) {
    alert("❌ Merge Data Missing!\n\nIf you closed this popup window after uploading, the file was lost from memory.\n\nPlease re-upload the Master Sheet before running.");
    return;
  }

  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

  await chrome.scripting.executeScript({
    target: { tabId: tab.id },
    files: ["xlsx.full.min.js"],
  });

  const dataToPass = mergeEnabled ? masterSheetData : {};

  chrome.scripting.executeScript({
    target: { tabId: tab.id },
    func: startExtractionProcess,
    args: [blacklistEnabled, existingDomains, dataToPass],
  });
});

// 3. Stop Button Logic
document.getElementById("stop").addEventListener("click", async () => {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  chrome.scripting.executeScript({
    target: { tabId: tab.id },
    func: () => {
      window.__guestbookStop = true;
    },
  });
});

// --- The Content Script Logic ---
// --- REPLACE THE ENTIRE startExtractionProcess FUNCTION WITH THIS ---

function startExtractionProcess(
  blacklistEnabled,
  exclusionList,
  masterSheetData
) {
  window.__guestbookStop = false;
  window.__clickCounter = 0;

  let totalEntries = 0;
  let batchSize = 20;
  let isAnalyzing = true;

  const EXCLUSION_SET = new Set(exclusionList);

  // Intelligence Buckets
  const CATEGORIES = {
    Profiles: [
      "/profile/", "/user/", "/u/", "/member/", "/members/", "/candidate/", 
      "/author/", "/people/", "/users/", "/profiles/", "/@/",
    ],
    Discussions: [
      "/forum/", "/forums/", "/topic/", "/thread/", "/discuss/", "/discussions/", 
      "/question/", "/issues/", "/group/", "/groups/", "/community/", "/board/", 
      "/viewtopic", "mysite-200-group", "/guestbook.html", "/mn.co/posts/",
    ],
    Business: [
      "/company/", "/companies/", "/employer/", "/employers/", "/listing/", 
      "/classifieds/", "/ads/", "/services/",
    ],
    Content: [
      "/blog/", "/blogs/", "/read-blog/", "/post/", "/posts/", "/p/", "/page/", 
      "/webpage/", "/wiki/", "/media/",
    ],
    Projects: [
      "/snippets/", "/projects", "/idea/", "/ideas/", "/-/projects", 
      "?tab=field_core_pfield",
    ],
    Other: [],
  };

  const BLACKLIST = blacklistEnabled
    ? [
        "carrd.co", "10web.site", "google.com", "wordpress.com", "zendesk.com",
        "livejournal.com", "mystrikingly.com", "bandcamp.com", "about.me",
        "gumroad.com", "webflow.io", "blogspot.com", "linktr.ee", "wix.com",
        "weebly.com", "squarespace.com", "tumblr.com", "facebook.com",
        "twitter.com", "instagram.com", "linkedin.com", "youtube.com",
        "mobirisesite.com",
      ]
    : [];

  // --- UI SETUP (Preserved exactly as requested) ---
  const overlay = document.createElement("div");
  overlay.id = "__gbOverlay";
  Object.assign(overlay.style, {
    position: "fixed", bottom: "20px", right: "20px", width: "300px",
    background: "rgba(0,0,0,0.9)", color: "#fff", padding: "15px",
    zIndex: "10000", borderRadius: "8px", fontFamily: "Arial, sans-serif",
    boxShadow: "0 4px 20px rgba(0,0,0,0.5)", transition: "all 0.3s",
  });

  overlay.innerHTML = `
    <div style="font-weight:bold; font-size:16px; margin-bottom:10px; border-bottom:1px solid #444; padding-bottom:5px;">
      Guestbook Snowball v2.3
    </div>
    <div id="__gbStatus" style="margin-bottom:10px; color:#aaa;">Initializing Probe...</div>
    <div id="__gbStats" style="font-size:13px; display:none;">
      <div>Total Entries: <span id="__gbTotal" style="color:#fff">...</span></div>
      <div>Est. Clicks: <span id="__gbEst" style="color:#fff">...</span></div>
    </div>
    <div id="__gbProgress" style="height:6px; background:#444; margin:10px 0; border-radius:3px; overflow:hidden; display:none;">
      <div id="__gbBar" style="width:0%; height:100%; background:#00e676; transition: width 0.5s;"></div>
    </div>
    <div id="__gbControls" style="display:none; text-align:center;">
      <button id="__gbBtnGo" style="background:#28a745; border:none; color:fff; padding:5px 15px; border-radius:4px; cursor:pointer; color:white; font-weight:bold;">PROCEED</button>
      <button id="__gbBtnStop" style="background:#d9534f; border:none; color:fff; padding:5px 15px; border-radius:4px; cursor:pointer; margin-left:10px; color:white;">CANCEL</button>
    </div>
  `;
  document.body.appendChild(overlay);

  const observer = new PerformanceObserver((list) => {
    list.getEntries().forEach((entry) => {
      if (entry.name.includes("message.asmx") && entry.name.includes("%5D")) {
        const match = entry.name.match(/%2C(\d+)%2C(\d+)%5D/);
        if (match && isAnalyzing) {
          totalEntries = parseInt(match[1]);
          batchSize = parseInt(match[2]);
          presentAnalysis();
        }
      }
    });
  });
  observer.observe({ entryTypes: ["resource"] });

  function updateStatus(text, color = "#aaa") {
    const el = document.getElementById("__gbStatus");
    if (el) {
      el.textContent = text;
      el.style.color = color;
    }
  }

  function presentAnalysis() {
    isAnalyzing = false;
    const remainingEstimate = Math.max(0, Math.ceil(totalEntries / batchSize));

    document.getElementById("__gbTotal").textContent = totalEntries.toLocaleString();
    document.getElementById("__gbEst").textContent = "~" + remainingEstimate;
    document.getElementById("__gbStats").style.display = "block";
    document.getElementById("__gbControls").style.display = "block";
    updateStatus("Analysis Complete. Ready?", "#00e676");

    document.getElementById("__gbBtnGo").onclick = () => {
      document.getElementById("__gbControls").style.display = "none";
      document.getElementById("__gbProgress").style.display = "block";
      updateStatus("Extracting...", "#fff");
      clickButton();
    };

    document.getElementById("__gbBtnStop").onclick = () => {
      document.body.removeChild(overlay);
      window.__guestbookStop = true;
    };
  }

  function updateProgress() {
    const currentEst = window.__clickCounter * batchSize;
    const percent = Math.min(100, (currentEst / totalEntries) * 100);
    const bar = document.getElementById("__gbBar");
    if (bar) bar.style.width = percent + "%";
    updateStatus(`Clicks: ${window.__clickCounter} (Loaded ~${currentEst})`, "#fff");
  }

  function clickButton() {
    const button = document.querySelector("#_aabb-morecount");

    if (window.__guestbookStop) {
      updateStatus("Stopped. Categorizing...", "#d9534f");
      setTimeout(processAndMerge, 100);
      return;
    }

    if (button && typeof button.click === "function") {
      button.click();
      window.__clickCounter++;
      if (!isAnalyzing) {
        updateProgress();
        setTimeout(clickButton, Math.floor(Math.random() * 1000) + 1500);
      } else {
        setTimeout(() => { if (isAnalyzing) {} }, 2000);
      }
    } else {
      updateStatus("Guestbook Finished!", "#00e676");
      if (document.getElementById("__gbBar"))
        document.getElementById("__gbBar").style.width = "100%";
      setTimeout(processAndMerge, 1000);
    }
  }

  // --- PROCESSING LOGIC ---
  function processAndMerge() {
    try {
      const links = Array.from(document.querySelectorAll('a[href^="http"]'));
      const uniqueDomains = new Set();

      // Data Structures
      const oldDataMap = {};
      const newDataMap = {};
      
      // Initialize structures
      Object.keys(CATEGORIES).forEach((key) => {
        oldDataMap[key] = [];
        newDataMap[key] = [];
      });
      // Ensure "Other" exists
      if (!oldDataMap["Other"]) oldDataMap["Other"] = [];
      if (!newDataMap["Other"]) newDataMap["Other"] = [];

      // 1. Load OLD Data (Master File) safely
      if (masterSheetData && typeof masterSheetData === 'object' && Object.keys(masterSheetData).length > 0) {
        for (const [sheetName, rows] of Object.entries(masterSheetData)) {
          if(!Array.isArray(rows)) continue;

          // --- FIX: SKIP OLD REPORT SHEETS ---
          // This prevents the old report from being dumped into "Other"
          if (sheetName.toLowerCase().includes("report")) {
            continue; 
          }

          // Strip Header: Remove first row or row with "Domain"
          const dataRows = rows.filter((r, index) => {
            if (index === 0) return false; 
            if (r[0] === "Domain") return false;
            return true;
          });

          // Match Sheet Name to Category (Case Insensitive)
          let targetKey = "Other";
          const matchedCategory = Object.keys(CATEGORIES).find(
            cat => cat.toLowerCase() === sheetName.toLowerCase()
          );

          if (matchedCategory) {
            targetKey = matchedCategory;
          }

          // If mapping exists, add data. If not, dump in Other.
          if (!oldDataMap[targetKey]) {
             targetKey = "Other"; 
          }
          oldDataMap[targetKey] = oldDataMap[targetKey].concat(dataRows);
        }
      }

      // 2. Process NEW Links
      let totalNewLinks = 0;

      links.forEach((a) => {
        try {
          const url = a.href.trim();
          const fullObj = new URL(url);
          let host = fullObj.hostname.replace(/^www\./, "").toLowerCase();
          let path = (fullObj.pathname + fullObj.search).toLowerCase();

          // Filters
          if (BLACKLIST.some((blocked) => host.endsWith(blocked))) return;
          if (EXCLUSION_SET.has(host)) return;
          if ((fullObj.pathname === "/" || fullObj.pathname === "") && fullObj.search === "") return;
          
          const domainParts = host.split(".");
          if (domainParts.length > 2 && host.includes("blog")) return;
          if (uniqueDomains.has(host)) return;

          uniqueDomains.add(host);
          totalNewLinks++;

          // Categorization
          let targetCat = "Other";
          for (const [catName, patterns] of Object.entries(CATEGORIES)) {
            if (patterns.some((p) => path.includes(p.toLowerCase()))) {
              targetCat = catName;
              break;
            }
          }

          // Add to NEW bucket [Domain, URL, DA, PA, SS, Status]
          newDataMap[targetCat].push([host, url, "", "", "", "New"]);
        } catch (e) {
          // invalid url ignored
        }
      });

      // --- EXCEL GENERATION ---
      if (typeof XLSX === 'undefined') {
        alert("XLSX Library not found. Please reload.");
        return;
      }

      const workbook = XLSX.utils.book_new();

      // -- FEATURE: Updated Report Sheet (Preserved & Fixed) --
      const reportRows = [
        ["Guestbook Extraction Report", new Date().toLocaleString()],
        ["", "", "", ""],
        ["Category", "Previous Links", "New Links", "Total Database"]
      ];
      
      let grandTotalOld = 0;
      let grandTotalNew = 0;

      Object.keys(CATEGORIES).forEach((cat) => {
        const oldCt = oldDataMap[cat] ? oldDataMap[cat].length : 0;
        const newCt = newDataMap[cat] ? newDataMap[cat].length : 0;
        
        if (oldCt > 0 || newCt > 0) {
          reportRows.push([cat, oldCt, newCt, oldCt + newCt]);
          grandTotalOld += oldCt;
          grandTotalNew += newCt;
        }
      });

      // Add Grand Totals
      reportRows.push(["", "", "", ""]);
      reportRows.push(["TOTALS", grandTotalOld, grandTotalNew, grandTotalOld + grandTotalNew]);

      // Create Report Sheet
      const reportSheet = XLSX.utils.aoa_to_sheet(reportRows);
      reportSheet["!cols"] = [{ wch: 20 }, { wch: 15 }, { wch: 15 }, { wch: 15 }];
      XLSX.utils.book_append_sheet(workbook, reportSheet, "Extraction Report");

      // -- MERGING & DATA SHEETS --
      let hasData = false;
      const allSheetNames = Object.keys(CATEGORIES); 

      allSheetNames.forEach((sheetName) => {
        const oldRows = oldDataMap[sheetName] || [];
        const newRows = newDataMap[sheetName] || [];

        if (oldRows.length > 0 || newRows.length > 0) {
          // Define Header
          const finalRows = [["Domain", "Full URL", "DA", "PA", "SS", "Status"]];

          // Add Old Data
          if (oldRows.length > 0) {
            oldRows.forEach((r) => finalRows.push(r));
          }

          // INSERT SEPARATOR IF MERGING
          if (oldRows.length > 0 && newRows.length > 0) {
            finalRows.push(["--- NEW DATA BELOW ---", "", "", "", "", ""]);
          }

          // Add New Data
          if (newRows.length > 0) {
            newRows.forEach((r) => finalRows.push(r));
          }

          const worksheet = XLSX.utils.aoa_to_sheet(finalRows);
          // Set nice column widths
          worksheet["!cols"] = [
            { wch: 30 }, { wch: 40 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 15 }
          ];

          XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
          hasData = true;
        }
      });

      if (!hasData) {
          updateStatus("No new unique links found.", "#d9534f");
          setTimeout(() => { 
            const ov = document.getElementById("__gbOverlay");
            if(ov) ov.remove(); 
          }, 3000);
          return;
      }

      XLSX.writeFile(workbook, "Guestbook_Smart_DB.xlsx");

      updateStatus(`Success! +${totalNewLinks} New Links.`, "#00e676");
      setTimeout(() => {
        const ov = document.getElementById("__gbOverlay");
        if (ov) ov.remove();
      }, 5000);

    } catch(err) {
      console.error(err);
      alert("Error during processing: " + err.message);
      updateStatus("Error Occurred", "red");
    }
  }

  updateStatus("Probing Network...");
  clickButton();
}