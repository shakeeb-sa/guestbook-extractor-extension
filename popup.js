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

  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

  await chrome.scripting.executeScript({
    target: { tabId: tab.id },
    files: ["xlsx.full.min.js"],
  });

  // If Merge is OFF, we pass an empty object for masterData
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
      "/profile/",
      "/user/",
      "/u/",
      "/member/",
      "/members/",
      "/candidate/",
      "/author/",
      "/people/",
      "/users/",
      "/profiles/",
      "/@/",
    ],
    Discussions: [
      "/forum/",
      "/forums/",
      "/topic/",
      "/thread/",
      "/discuss/",
      "/discussions/",
      "/question/",
      "/issues/",
      "/group/",
      "/groups/",
      "/community/",
      "/board/",
      "/viewtopic",
      "mysite-200-group",
      "/guestbook.html",
      "/mn.co/posts/",
    ],
    Business: [
      "/company/",
      "/companies/",
      "/employer/",
      "/employers/",
      "/listing/",
      "/classifieds/",
      "/ads/",
      "/services/",
    ],
    Content: [
      "/blog/",
      "/blogs/",
      "/read-blog/",
      "/post/",
      "/posts/",
      "/p/",
      "/page/",
      "/webpage/",
      "/wiki/",
      "/media/",
    ],
    Projects: [
      "/snippets/",
      "/projects",
      "/idea/",
      "/ideas/",
      "/-/projects",
      "?tab=field_core_pfield",
    ],
    Other: [],
  };

  const BLACKLIST = blacklistEnabled
    ? [
        "carrd.co",
        "10web.site",
        "google.com",
        "wordpress.com",
        "zendesk.com",
        "livejournal.com",
        "mystrikingly.com",
        "bandcamp.com",
        "about.me",
        "gumroad.com",
        "webflow.io",
        "blogspot.com",
        "linktr.ee",
        "wix.com",
        "weebly.com",
        "squarespace.com",
        "tumblr.com",
        "facebook.com",
        "twitter.com",
        "instagram.com",
        "linkedin.com",
        "youtube.com",
        "mobirisesite.com",
      ]
    : [];

  // --- UI SETUP (No changes here) ---
  const overlay = document.createElement("div");
  overlay.id = "__gbOverlay";
  Object.assign(overlay.style, {
    position: "fixed",
    bottom: "20px",
    right: "20px",
    width: "300px",
    background: "rgba(0,0,0,0.9)",
    color: "#fff",
    padding: "15px",
    zIndex: "10000",
    borderRadius: "8px",
    fontFamily: "Arial, sans-serif",
    boxShadow: "0 4px 20px rgba(0,0,0,0.5)",
    transition: "all 0.3s",
  });

  overlay.innerHTML = `
    <div style="font-weight:bold; font-size:16px; margin-bottom:10px; border-bottom:1px solid #444; padding-bottom:5px;">
      Guestbook Snowball v2.2
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

    document.getElementById("__gbTotal").textContent =
      totalEntries.toLocaleString();
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
    updateStatus(
      `Clicks: ${window.__clickCounter} (Loaded ~${currentEst})`,
      "#fff"
    );
  }

  function clickButton() {
    const button = document.querySelector("#_aabb-morecount");

    if (window.__guestbookStop) {
      updateStatus("Stopped. Categorizing...", "#d9534f");
      processAndMerge();
      return;
    }

    if (button && typeof button.click === "function") {
      button.click();
      window.__clickCounter++;
      if (!isAnalyzing) {
        updateProgress();
        setTimeout(clickButton, Math.floor(Math.random() * 1000) + 1500);
      } else {
        setTimeout(() => {
          if (isAnalyzing) {
          }
        }, 2000);
      }
    } else {
      updateStatus("Guestbook Finished!", "#00e676");
      if (document.getElementById("__gbBar"))
        document.getElementById("__gbBar").style.width = "100%";
      setTimeout(processAndMerge, 1000);
    }
  }

  // --- UPGRADED PROCESSING LOGIC ---
  function processAndMerge() {
    const links = Array.from(document.querySelectorAll('a[href^="http"]'));
    const uniqueDomains = new Set();

    const oldDataMap = {};
    const newDataMap = {};
    const statsMap = {};

    // Initialize structures
    Object.keys(CATEGORIES).forEach((key) => {
      oldDataMap[key] = [];
      newDataMap[key] = [];
      statsMap[key] = 0;
    });
    // Add Other bucket
    oldDataMap["Other"] = [];
    newDataMap["Other"] = [];
    statsMap["Other"] = 0;

    // 1. Load OLD Data (Master File) - PRESERVE ALL COLUMNS
    if (masterSheetData && Object.keys(masterSheetData).length > 0) {
      for (const [sheetName, rows] of Object.entries(masterSheetData)) {
        // Filter out header row ("Domain") to get pure data
        // We assume the first row is header if the first cell is "Domain"
        const dataRows = rows.filter((r, index) => {
          if (index === 0 && r[0] === "Domain") return false;
          return true;
        });

        if (oldDataMap[sheetName] !== undefined) {
          oldDataMap[sheetName] = dataRows;
        } else {
          oldDataMap[sheetName] = dataRows;
        }
      }
    }

    // 2. Process NEW Links
    let totalNewLinks = 0;

    links.forEach((a) => {
      try {
        const url = a.href.trim();
        const fullObj = new URL(url);
        let host = fullObj.hostname.replace(/^www\./, "").toLowerCase();
        let path =
          fullObj.pathname.toLowerCase() + fullObj.search.toLowerCase();

        // Filters
        if (BLACKLIST.some((blocked) => host.endsWith(blocked))) return;
        if (EXCLUSION_SET.has(host)) return;
        if (
          (fullObj.pathname === "/" || fullObj.pathname === "") &&
          fullObj.search === ""
        )
          return;
        const domainParts = host.split(".");
        if (domainParts.length > 2 && host.includes("blog")) return;
        if (uniqueDomains.has(host)) return;

        uniqueDomains.add(host);
        totalNewLinks++;

        // Categorization
        let assigned = false;
        let targetCat = "Other";

        for (const [catName, patterns] of Object.entries(CATEGORIES)) {
          if (patterns.some((p) => path.includes(p.toLowerCase()))) {
            targetCat = catName;
            assigned = true;
            break;
          }
        }

        // Add to NEW bucket with "New" flag in Column F (Index 5)
        // Structure: [Domain, URL, Empty(DA), Empty(PA), Empty(SS), "New"]
        newDataMap[targetCat].push([host, url, "", "", "", "New"]);
        statsMap[targetCat]++;
      } catch (e) {}
    });

    // --- EXCEL GENERATION ---
    const workbook = XLSX.utils.book_new();

    // -- FEATURE: Report Sheet --
    const reportData = [
      ["Run Report", new Date().toLocaleString()],
      ["Total New Links", totalNewLinks],
      ["", ""],
      ["Category", "New Links Added"],
    ];
    Object.entries(statsMap).forEach(([cat, count]) => {
      if (count > 0) reportData.push([cat, count]);
    });
    if (totalNewLinks > 0) {
      const reportSheet = XLSX.utils.aoa_to_sheet(reportData);
      // Set Report Column Widths too for niceness
      reportSheet["!cols"] = [{ wch: 20 }, { wch: 20 }];
      XLSX.utils.book_append_sheet(workbook, reportSheet, "Extraction Report");
    }

    // -- MERGING --
    let hasData = false;
    const allSheetNames = new Set([
      ...Object.keys(oldDataMap),
      ...Object.keys(newDataMap),
    ]);

    allSheetNames.forEach((sheetName) => {
      const oldRows = oldDataMap[sheetName] || [];
      const newRows = newDataMap[sheetName] || [];

      if (oldRows.length > 0 || newRows.length > 0) {
        // Define Header with placeholders for DA/PA/SS
        const finalRows = [["Domain", "Full URL", "DA", "PA", "SS", "Status"]];

        // Add Old Data (Preserve whatever columns they had)
        if (oldRows.length > 0) {
          oldRows.forEach((r) => finalRows.push(r));
        }

        // INSERT GAP
        if (oldRows.length > 0 && newRows.length > 0) {
          finalRows.push(["", "", "", "", "", ""]);
        }

        // Add New Data
        if (newRows.length > 0) {
          newRows.forEach((r) => finalRows.push(r));
        }

        const worksheet = XLSX.utils.aoa_to_sheet(finalRows);

        // -- FEATURE: Column Widths --
        // Set Cols A and B to width 35. Others to default.
        worksheet["!cols"] = [
          { wch: 35 }, // A (Domain)
          { wch: 35 }, // B (URL)
          { wch: 10 }, // C (DA)
          { wch: 10 }, // D (PA)
          { wch: 10 }, // E (SS)
          { wch: 10 }, // F (Status/New)
        ];

        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
        hasData = true;
      }
    });

    if (!hasData && totalNewLinks === 0) {
      if (Object.keys(masterSheetData).length > 0) {
        // Just saving master data back
      } else {
        updateStatus("No new unique links found.", "#d9534f");
        setTimeout(() => {
          document.getElementById("__gbOverlay").remove();
        }, 3000);
        return;
      }
    }

    XLSX.writeFile(workbook, "Guestbook_Smart_DB.xlsx");

    updateStatus(`Success! +${totalNewLinks} New Links.`, "#00e676");
    setTimeout(() => {
      const ov = document.getElementById("__gbOverlay");
      if (ov) ov.remove();
    }, 5000);
  }

  updateStatus("Probing Network...");
  clickButton();
}
