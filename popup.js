let existingDomains = [];

// 1. Handle File Upload & Parsing
document.getElementById("excelInput").addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });

    const tempSet = new Set();

    // Iterate ALL sheets
    workbook.SheetNames.forEach((sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      // Convert sheet to JSON array of arrays (rows)
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      rows.forEach((row) => {
        row.forEach((cell) => {
          if (typeof cell === "string" && cell.includes(".")) {
            // Basic normalization to extract domain
            try {
              // If it doesn't have http, add it to parse correctly
              const urlToParse = cell.startsWith("http")
                ? cell
                : `http://${cell}`;
              const urlObj = new URL(urlToParse);
              let host = urlObj.hostname.replace(/^www\./, "");
              tempSet.add(host.toLowerCase());
            } catch (err) {
              // Not a valid URL, skip
            }
          }
        });
      });
    });

    existingDomains = Array.from(tempSet);
    document.getElementById(
      "fileStatus"
    ).textContent = `Loaded ${existingDomains.length} domains to exclude.`;
    document.getElementById("fileStatus").style.color = "green";
  };
  reader.readAsArrayBuffer(file);
});

// 2. Run Button Logic
document.getElementById("run").addEventListener("click", async () => {
  const blacklistEnabled = document.getElementById("blacklistToggle").checked;
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

  // INJECT SheetJS into the page so we can use it for the export later
  await chrome.scripting.executeScript({
    target: { tabId: tab.id },
    files: ["xlsx.full.min.js"],
  });

  // INJECT Main Logic
  chrome.scripting.executeScript({
    target: { tabId: tab.id },
    func: startExtractionProcess,
    args: [blacklistEnabled, existingDomains], // Pass the parsed domains here
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
function startExtractionProcess(blacklistEnabled, exclusionList) {
  window.__guestbookStop = false;
  window.__clickCounter = 0;

  // State for our analysis
  let totalEntries = 0;
  let batchSize = 20;
  let isAnalyzing = true; // Start in analysis mode

  const EXCLUSION_SET = new Set(exclusionList);
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

  // --- UI CREATION ---
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

  // Inner HTML for the dashboard
  overlay.innerHTML = `
    <div style="font-weight:bold; font-size:16px; margin-bottom:10px; border-bottom:1px solid #444; padding-bottom:5px;">
      Guestbook Intelligence
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

  // --- NETWORK SNIFFER (The Magic) ---
  // Observes network requests to find the .asmx URL
  const observer = new PerformanceObserver((list) => {
    list.getEntries().forEach((entry) => {
      // Look for the specific API signature
      if (entry.name.includes("message.asmx") && entry.name.includes("%5D")) {
        // Regex to find the pattern: %2C (comma) NUMBER %2C (comma) NUMBER %5D (closing bracket)
        // Matches: ... , 1663 , 20 ]
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

  // --- FUNCTIONS ---

  function updateStatus(text, color = "#aaa") {
    const el = document.getElementById("__gbStatus");
    if (el) {
      el.textContent = text;
      el.style.color = color;
    }
  }

  function presentAnalysis() {
    isAnalyzing = false; // Stop sniffing
    const remainingEstimate = Math.max(0, Math.ceil(totalEntries / batchSize)); // Rough calc

    document.getElementById("__gbTotal").textContent =
      totalEntries.toLocaleString();
    document.getElementById("__gbEst").textContent = "~" + remainingEstimate;

    document.getElementById("__gbStats").style.display = "block";
    document.getElementById("__gbControls").style.display = "block";

    updateStatus("Analysis Complete. Ready?", "#00e676");

    // Bind Buttons
    document.getElementById("__gbBtnGo").onclick = () => {
      document.getElementById("__gbControls").style.display = "none";
      document.getElementById("__gbProgress").style.display = "block";
      updateStatus("Extracting...", "#fff");
      clickButton(); // RESUME LOOP
    };

    document.getElementById("__gbBtnStop").onclick = () => {
      document.body.removeChild(overlay);
      window.__guestbookStop = true;
    };
  }

  function updateProgress() {
    // Update bar based on clicks vs estimated total
    // This is an approximation since we don't know exactly how many we've loaded without parsing DOM
    // But we can estimate based on clicks
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
      updateStatus("Stopped. Exporting...", "#d9534f");
      processAndExport();
      return;
    }

    if (button && typeof button.click === "function") {
      button.click();
      window.__clickCounter++;

      // If we are past the analysis phase, update UI
      if (!isAnalyzing) {
        updateProgress();
        // Random delay 1.5s - 2.5s
        setTimeout(clickButton, Math.floor(Math.random() * 1000) + 1500);
      } else {
        // If analyzing, wait a bit longer to ensure network request fires
        setTimeout(() => {
          // If observer failed to catch it (network lag), try one more click or default
          if (isAnalyzing) {
            // Fallback if we missed the network packet
            console.log("Network probe timed out, retrying...");
            // We could force analysis end here if needed, but let's just loop
            // For now, let's assume it works or we manually proceed
          }
        }, 2000);
      }
    } else {
      updateStatus("Guestbook Finished!", "#00e676");
      if (document.getElementById("__gbBar"))
        document.getElementById("__gbBar").style.width = "100%";
      setTimeout(processAndExport, 1000);
    }
  }

  function processAndExport() {
    const links = Array.from(document.querySelectorAll('a[href^="http"]'));
    const finalDomains = new Map();

    links.forEach((a) => {
      try {
        const url = a.href.trim();
        const fullObj = new URL(url);
        let host = fullObj.hostname.replace(/^www\./, "").toLowerCase();

        // 1. Check Web 2.0 Blacklist (The basic known ones)
        const isBlacklisted = BLACKLIST.some((blocked) =>
          host.endsWith(blocked)
        );

        // 2. Check Excel Exclusion List
        const isExcluded = EXCLUSION_SET.has(host);

        // 3. Check for Root/Naked Domains (Spam Filter)
        const isRootDomain =
          (fullObj.pathname === "/" || fullObj.pathname === "") &&
          fullObj.search === "";

        // 4. AUTOMATIC SPAM FARM DETECTION (The New Logic)
        // Logic: If the domain has 3+ parts (subdomain.name.com) AND contains "blog"
        // Example: "nikita.blog-gold.com" -> Parts: 3, Contains "blog" -> BANNED.
        const domainParts = host.split(".");
        const hasSubdomain = domainParts.length > 2;
        const containsSpamKeyword = host.includes("blog");

        const isSpamFarm = hasSubdomain && containsSpamKeyword;

        // 5. Final Gatekeeper
        if (
          !isBlacklisted &&
          !isExcluded &&
          !isRootDomain &&
          !isSpamFarm &&
          !finalDomains.has(host)
        ) {
          finalDomains.set(host, url);
        }
      } catch (e) {}
    });

    exportToExcel(finalDomains);
  }

  function exportToExcel(domainMap) {
    const data = [["Domain", "Full URL"]];
    domainMap.forEach((url, domain) => {
      data.push([domain, url]);
    });

    const worksheet = XLSX.utils.aoa_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "New Domains");
    XLSX.writeFile(workbook, "Guestbook_Links.xlsx");

    updateStatus("Download Started!", "#00e676");
    setTimeout(() => {
      const ov = document.getElementById("__gbOverlay");
      if (ov) ov.remove();
    }, 5000);
  }

  // --- KICKOFF ---
  updateStatus("Probing Network...");
  // Trigger the FIRST click to generate the network request
  clickButton();
}
