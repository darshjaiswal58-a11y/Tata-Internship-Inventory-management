const $ = (selector) => document.querySelector(selector);
let currentUser = null;
let searchTimer = null;
let currentCategory = "ml_spare";
let currentMaterialGroup = "";
let latestSummary = null;
let materialGroups = [];
let criticalTimer = null;
let uploadProgressTimer = null;
let inventoryProgressTimer = null;
const RED_ZONE_VALUE = "__red_zone__";
let selectedZone = "red";

const zoneSeries = [
  { key: "red", label: "Red Zone", color: "#c2410c" },
  { key: "yellow", label: "Yellow Zone", color: "#d99a00" },
  { key: "green", label: "Green Zone", color: "#16803c" },
];

const categoryLabels = {
  ml_spare: "Machinery Spare",
  tools: "Tools",
};

const stockFilterCriteriaText = {
  ml_spare: "Machinery stock values show plant 3002 materials from the 180 stock-item Excel.",
  tools: "Tools stock values show plant 3004 materials from the 180 stock-item Excel.",
};

function currentStockFilterCriteriaText() {
  return stockFilterCriteriaText[currentCategory] || stockFilterCriteriaText.ml_spare;
}

async function api(path, options = {}) {
  const response = await fetch(path, { credentials: "same-origin", ...options });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(payload.error || "Request failed");
  }
  return payload;
}

function createJobId() {
  if (window.crypto?.randomUUID) return window.crypto.randomUUID();
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function renderUploadProgress(percent, message) {
  const progress = $("#uploadProgress");
  const fill = $("#uploadProgressBar");
  const percentLabel = $("#uploadProgressPercent");
  const label = $("#uploadProgressLabel");
  const value = Math.max(0, Math.min(100, Math.round(Number(percent) || 0)));
  progress.hidden = false;
  fill.style.width = `${value}%`;
  percentLabel.textContent = `${value}%`;
  label.textContent = message || "Processing Excel...";
}

function setUploadBusy(isBusy) {
  const button = $("#processExcelButton");
  const fileInput = $("#excelFile");
  if (button) {
    button.disabled = isBusy;
    button.textContent = isBusy ? "Processing..." : "Process Excel";
  }
  if (fileInput) fileInput.disabled = isBusy;
}

function stopUploadProgressPolling() {
  if (uploadProgressTimer) {
    clearInterval(uploadProgressTimer);
    uploadProgressTimer = null;
  }
}

function startUploadProgressPolling(jobId, categoryLabel) {
  stopUploadProgressPolling();
  renderUploadProgress(2, `Uploading ${categoryLabel} Excel...`);
  uploadProgressTimer = setInterval(async () => {
    try {
      const progress = await api(`/api/upload-progress?job_id=${encodeURIComponent(jobId)}`);
      const label = progress.category_label || categoryLabel;
      renderUploadProgress(progress.percent || 0, progress.message ? `${label}: ${progress.message}` : `Processing ${label} Excel...`);
      if (progress.state === "done" || progress.state === "error") {
        stopUploadProgressPolling();
      }
    } catch (error) {
      renderUploadProgress(2, `Uploading ${categoryLabel} Excel...`);
    }
  }, 500);
}

function fmt(value) {
  if (value === undefined || value === null || value === "") return "-";
  return value;
}

function escapeAttr(value) {
  return String(value ?? "").replaceAll("&", "&amp;").replaceAll('"', "&quot;").replaceAll("<", "&lt;");
}

function groupLabel(code) {
  if (code === RED_ZONE_VALUE) return "Red Zone Material";
  if (!code) return "-";
  const group = materialGroups.find((item) => item.code === code);
  return group?.label || code;
}

function isRedZoneSelected() {
  return currentMaterialGroup === RED_ZONE_VALUE;
}

function updateCriticalFormState() {
  const redZone = isRedZoneSelected();
  const canEditCriteria = currentUser?.role === "admin" && !redZone;
  $("#criticalStockInput").disabled = !canEditCriteria;
  $("#groupReorderInput").disabled = !canEditCriteria;
  $("#groupCriticalForm").querySelector("button").disabled = !canEditCriteria;
  $("#criticalStockInput").placeholder = redZone ? "From selected area stock-zone data" : "Try 2, 5, or 10";
}

function setSummary(summary) {
  latestSummary = summary;
  renderCategoryCards(summary.categories || {});
  renderCategoryMaterialGroupOptions();
  const selected = summary.categories?.[currentCategory] || {};
  const materialsCount = $("#materialsCount");
  const criteriaCount = $("#criteriaCount");
  if (materialsCount) materialsCount.textContent = selected.materials_count || 0;
  if (criteriaCount) criteriaCount.textContent = selected.criteria_count || 0;
  $("#lowStockCount").textContent = selected.low_stock_count || 0;
  $("#activeStockCount").textContent = selected.active_stocks_count || 0;
  renderMaterialGroupStats();
}

function roleLabel(role) {
  if (role === "admin") return "Admin";
  if (role === "employee") return "Employee";
  if (role === "uploader") return "Uploader";
  if (role === "criteria") return "Criteria Manager";
  return role || "User";
}

function applyAuthState(user) {
  currentUser = user;
  document.body.classList.toggle("logged-in", Boolean(user));
  document.body.classList.toggle("logged-out", !user);
  $("#userBadge").textContent = user ? `${user.name} | ${roleLabel(user.role)}` : "";

  const canEditCriteria = user?.role === "admin";
  document.body.classList.toggle("is-admin", canEditCriteria);
  $("#criteriaForm").querySelectorAll("input, button").forEach((element) => {
    element.disabled = !canEditCriteria;
  });
  $("#criteriaAreaSelect").disabled = !canEditCriteria;
  $("#criteriaHint").textContent = canEditCriteria
    ? "The report includes materials where current stock is at or below the minimum stock."
    : "Only the admin can set or edit stock criteria.";
  $("#groupCriticalForm").querySelectorAll("input, button").forEach((element) => {
    element.disabled = !canEditCriteria;
  });
  updateCriticalFormState();
}

function renderCategoryCards(categories) {
  const mlSpare = categories.ml_spare || {};
  const tools = categories.tools || {};
  $("#mlSpareMaterials").textContent = mlSpare.materials_count || 0;
  $("#mlSpareCriteria").textContent = mlSpare.criteria_count || 0;
  $("#mlSpareActive").textContent = mlSpare.active_stocks_count || 0;
  $("#mlSpareLow").textContent = mlSpare.low_stock_count || 0;
  $("#toolsMaterials").textContent = tools.materials_count || 0;
  $("#toolsCriteria").textContent = tools.criteria_count || 0;
  $("#toolsActive").textContent = tools.active_stocks_count || 0;
  $("#toolsLow").textContent = tools.low_stock_count || 0;
}

function showDashboard() {
  $("#inventoryWorkspace").hidden = true;
  $("#workflow").hidden = true;
  $("#dashboard").hidden = false;
}

function showInventoryWorkspace() {
  $("#dashboard").hidden = true;
  $("#workflow").hidden = true;
  $("#inventoryWorkspace").hidden = false;
  refreshInventory().catch((error) => {
    $("#inventoryStatus").className = "status error";
    $("#inventoryStatus").textContent = error.message;
  });
}

async function selectCategory(category) {
  currentCategory = categoryLabels[category] ? category : "ml_spare";
  currentMaterialGroup = "";
  $("#selectedCategoryTitle").textContent = categoryLabels[currentCategory];
  $("#categorySelect").value = currentCategory;
  $("#uploadCategory").value = currentCategory;
  $("#criteriaCategory").value = currentCategory;
  $("#criteriaAreaSelect").value = currentCategory;
  $("#materialGroupSelect").value = "";
  $("#materialSearch").value = "";
  renderMaterialResults([]);
  $("#dashboard").hidden = true;
  $("#inventoryWorkspace").hidden = true;
  $("#workflow").hidden = false;
  await refresh();
  const rows = await api(`/api/materials?category=${encodeURIComponent(currentCategory)}&material_group=${encodeURIComponent(currentMaterialGroup)}`);
  renderMaterialResults(rows);
}

function renderCriteria(criteria) {
  const rows = Object.values(criteria);
  $("#criteriaRows").innerHTML = rows.length
    ? rows.map((row) => `
      <tr>
        <td>${fmt(row.material)}</td>
        <td>${groupLabel(row.material_group)}</td>
        <td>${fmt(row.description)}</td>
        <td>${fmt(row.minimum_stock)}</td>
        <td>${fmt(row.reorder_quantity)}</td>
        <td>${row.active ? "Active" : "Inactive"}</td>
        <td>${row.keep_stock ? "Kept" : "-"}</td>
        <td><button class="small-danger" type="button" data-criteria-delete="${escapeAttr(row.material)}">Delete</button></td>
      </tr>
    `).join("")
    : `<tr><td class="empty" colspan="8">No criteria added yet.</td></tr>`;
}

function renderUploads(uploads) {
  $("#uploadRows").innerHTML = uploads.length
    ? uploads.map((row) => `
      <tr>
        <td>${fmt(row.uploaded_at)}</td>
        <td>${fmt(row.file_name)}</td>
        <td>${fmt(row.uploaded_by)}</td>
        <td>${fmt(row.imported_rows)}</td>
        <td>${fmt(row.failed_count)}</td>
        <td>${fmt(row.no_criteria_count)}</td>
        <td>
          <div class="download-stack">
            <a class="download" href="/api/download?file=${encodeURIComponent(row.report_file)}">Failed Criteria</a>
            <a class="download" href="/api/no-criteria-export?upload_id=${encodeURIComponent(row.id)}&category=${encodeURIComponent(row.category || currentCategory)}">No Criteria</a>
          </div>
        </td>
      </tr>
    `).join("")
    : `<tr><td class="empty" colspan="7">No uploads processed yet.</td></tr>`;
}

function renderUploadActions(upload) {
  const container = $("#uploadActions");
  if (!upload) {
    container.innerHTML = "";
    return;
  }
  container.innerHTML = `
    <a class="download" href="/api/download?file=${encodeURIComponent(upload.report_file)}">Download failed criteria Excel</a>
    <a class="download" href="/api/no-criteria-export?upload_id=${encodeURIComponent(upload.id)}&category=${encodeURIComponent(upload.category || currentCategory)}">Download no-criteria items</a>
  `;
}

function overrideRows(overrides) {
  return [
    ...Object.values(overrides.keep || {}),
    ...Object.values(overrides.remove || {}),
  ].sort((a, b) => String(a.material).localeCompare(String(b.material)));
}

function renderStockOverrides(overrides) {
  const rows = overrideRows(overrides || {});
  $("#stockOverrideRows").innerHTML = rows.length
    ? rows.map((row) => `
      <tr>
        <td>${fmt(row.material)}</td>
        <td>${fmt(row.description)}</td>
        <td>${categoryLabels[row.category] || row.category}</td>
        <td>${row.action === "remove" ? "Remove from Stock" : "Add / Keep"}</td>
        <td>${fmt(row.updated_at)}</td>
        <td><button class="small-danger" type="button" data-stock-override-delete="${String(row.material).replaceAll('"', "&quot;")}">Delete</button></td>
      </tr>
    `).join("")
    : `<tr><td class="empty" colspan="6">No stock overrides added yet.</td></tr>`;
}

function renderMaterialResults(rows) {
  const container = $("#materialResults");
  if (!rows.length) {
    container.innerHTML = `<div class="result-empty">No matching materials found.</div>`;
    return;
  }
  container.innerHTML = rows.map((row) => `
    <button class="material-result" type="button"
      data-material="${String(row.material).replaceAll('"', "&quot;")}"
      data-description="${String(row.description || "").replaceAll('"', "&quot;")}"
      data-minimum-stock="${row.minimum_stock ?? ""}"
      data-reorder-quantity="${row.reorder_quantity ?? ""}"
      data-keep-stock="${row.keep_stock ? "true" : "false"}">
      <strong>${fmt(row.material)}</strong>
      <span>${fmt(row.description)}</span>
      <small>Group: ${groupLabel(row.material_group)}</small>
      <small>Stock: ${fmt(row.current_stock)}</small>
      <small>Net: ${fmt(row.net_consumption)}</small>
      <small>${row.has_criteria ? "Criteria set" : "No criteria"}</small>
    </button>
  `).join("");
}

async function searchMaterials(query) {
  if (!query || query.trim().length < 2) {
    $("#materialResults").innerHTML = `<div class="result-empty">Type at least 2 characters to search imported materials.</div>`;
    return;
  }
  const rows = await api(`/api/materials?category=${encodeURIComponent(currentCategory)}&material_group=${encodeURIComponent(currentMaterialGroup)}&q=${encodeURIComponent(query.trim())}`);
  renderMaterialResults(rows);
}

function renderFailed(rows) {
  $("#failedRows").innerHTML = rows.length
    ? rows.map((row) => `
      <tr>
        <td>${fmt(row.material)}</td>
        <td>${groupLabel(row.material_group)}</td>
        <td>${fmt(row.description)}</td>
        <td>${fmt(row.current_stock)}</td>
        <td>${fmt(row.minimum_stock)}</td>
        <td>${fmt(row.net_consumption)}</td>
        <td>${fmt(row.reorder_quantity)}</td>
        <td>${fmt(row.reason)}</td>
      </tr>
    `).join("")
    : `<tr><td class="empty" colspan="8">Upload a file to see failed criteria here.</td></tr>`;
}

function renderMaterialGroupOptions(groups) {
  materialGroups = groups || [];
  renderCategoryMaterialGroupOptions();
}

function renderCategoryMaterialGroupOptions() {
  const selected = latestSummary?.categories?.[currentCategory] || {};
  const categoryGroups = Object.values(selected.material_groups || {});
  const options = categoryGroups.length ? categoryGroups : materialGroups;
  const validValues = new Set(options.map((group) => String(group.code)));
  if (currentMaterialGroup && currentMaterialGroup !== RED_ZONE_VALUE && !validValues.has(currentMaterialGroup)) {
    currentMaterialGroup = "";
  }
  $("#materialGroupSelect").innerHTML = [
    `<option value="">All material groups</option>`,
    `<option value="${RED_ZONE_VALUE}">Red Zone Material</option>`,
    ...options.map((group) => (
      `<option value="${String(group.code).replaceAll('"', "&quot;")}">${fmt(group.label)}</option>`
    )),
  ].join("");
  $("#materialGroupSelect").value = currentMaterialGroup;
}

function renderFilterCriteriaSummary() {
  const title = $("#filterCriteriaTitle");
  const detail = $("#filterCriteriaDetail");
  const criteriaText = currentStockFilterCriteriaText();
  if (isRedZoneSelected()) {
    title.textContent = "Zone Status = Red Zone";
    detail.textContent = `${criteriaText} Showing materials where Zone Status is Red.`;
    return;
  }
  if (currentMaterialGroup) {
    title.textContent = `Material Group = ${groupLabel(currentMaterialGroup)}`;
    detail.textContent = `${criteriaText} Showing parts whose material group matches the selected part type.`;
    return;
  }
  title.textContent = "All material groups";
  detail.textContent = criteriaText;
}

function currentZoneAnalysis() {
  const selected = latestSummary?.categories?.[currentCategory] || {};
  if (isRedZoneSelected()) return selected.zone_analysis || {};
  const groupCards = selected.material_groups || {};
  const group = currentMaterialGroup ? groupCards[currentMaterialGroup] || {} : null;
  return group ? group.zone_analysis || {} : selected.zone_analysis || {};
}

function zoneValue(zoneAnalysis, zone) {
  return Number(zoneAnalysis?.[zone] || 0);
}

function renderZoneChart(zoneAnalysis) {
  const svg = $("#zoneLineChart");
  const width = 640;
  const height = 280;
  const left = 58;
  const right = 24;
  const top = 24;
  const bottom = 46;
  const chartWidth = width - left - right;
  const chartHeight = height - top - bottom;
  const values = zoneSeries.map((zone) => zoneValue(zoneAnalysis, zone.key));
  const maxValue = Math.max(10, ...values);
  const yMax = Math.ceil(maxValue / 5) * 5;
  const xFor = (index) => left + (chartWidth / (zoneSeries.length - 1)) * index;
  const yFor = (value) => top + chartHeight - (value / yMax) * chartHeight;
  const baseline = yFor(0);

  const grid = [];
  for (let i = 0; i <= 5; i += 1) {
    const value = Math.round((yMax / 5) * i);
    const y = yFor(value);
    grid.push(`<line class="zone-grid-line" x1="${left}" y1="${y}" x2="${width - right}" y2="${y}"></line>`);
    grid.push(`<text class="zone-axis-label" x="16" y="${y + 4}">${value}</text>`);
  }
  zoneSeries.forEach((zone, index) => {
    const x = xFor(index);
    grid.push(`<line class="zone-grid-line" x1="${x}" y1="${top}" x2="${x}" y2="${baseline}"></line>`);
    grid.push(`<text class="zone-axis-label" x="${x - 24}" y="${height - 14}">${zone.label.replace(" Zone", "")}</text>`);
  });

  const lines = zoneSeries.map((zone, index) => {
    const value = zoneValue(zoneAnalysis, zone.key);
    const x = xFor(index);
    const y = yFor(value);
    const path = `M ${left} ${baseline} Q ${x} ${y} ${x} ${y} T ${width - right} ${baseline}`;
    return `
      <path class="zone-line" data-zone="${zone.key}" data-value="${value}" d="${path}" stroke="${zone.color}"></path>
      <circle class="zone-point" tabindex="0" data-zone="${zone.key}" data-value="${value}" cx="${x}" cy="${y}" r="8" fill="${zone.color}"></circle>
      <text class="zone-axis-label" x="${x - 8}" y="${y - 14}">${value}</text>
    `;
  }).join("");

  svg.innerHTML = `${grid.join("")}${lines}`;
  $("#zoneLegend").innerHTML = zoneSeries.map((zone) => `
    <button type="button" data-zone="${zone.key}">
      <span class="zone-swatch" style="background:${zone.color}"></span>
      ${zone.label}: ${zoneValue(zoneAnalysis, zone.key)}
    </button>
  `).join("");
  const selected = zoneSeries.find((zone) => zone.key === selectedZone) || zoneSeries[0];
  $("#zoneChartTitle").textContent = `${selected.label}: ${zoneValue(zoneAnalysis, selected.key)}`;
}

function showZoneTooltip(target) {
  const zone = zoneSeries.find((item) => item.key === target.dataset.zone);
  if (!zone) return;
  const tooltip = $("#zoneTooltip");
  tooltip.textContent = `${zone.label}: ${target.dataset.value}`;
  tooltip.hidden = false;
}

function hideZoneTooltip() {
  $("#zoneTooltip").hidden = true;
}

async function loadZoneParts(zoneKey = selectedZone) {
  selectedZone = zoneKey;
  const groupParam = currentMaterialGroup && currentMaterialGroup !== RED_ZONE_VALUE ? currentMaterialGroup : "";
  const result = await api(`/api/zone-parts?category=${encodeURIComponent(currentCategory)}&zone=${encodeURIComponent(selectedZone)}&material_group=${encodeURIComponent(groupParam)}`);
  const zone = zoneSeries.find((item) => item.key === selectedZone) || zoneSeries[0];
  $("#zoneChartStatus").textContent = `${zone.label}: ${result.value || 0} part(s).`;
  $("#zoneDownloadLink").href = `/api/zone-export?category=${encodeURIComponent(currentCategory)}&zone=${encodeURIComponent(selectedZone)}&material_group=${encodeURIComponent(groupParam)}`;
  $("#zonePartRows").innerHTML = result.rows?.length
    ? result.rows.map((row) => `
      <tr>
        <td>${fmt(row.material)}</td>
        <td>${groupLabel(row.material_group)}</td>
        <td>${fmt(row.description)}</td>
        <td>${fmt(row.zone)}</td>
        <td>${fmt(row.current_stock)}</td>
        <td>${fmt(row.critical_value)}</td>
      </tr>
    `).join("")
    : `<tr><td class="empty" colspan="6">No parts found in ${zone.label}.</td></tr>`;
}

function renderMaterialGroupStats() {
  renderFilterCriteriaSummary();
  const selected = latestSummary?.categories?.[currentCategory] || {};
  if (isRedZoneSelected()) {
    const zoneAnalysis = selected.zone_analysis || {};
    $("#groupSourceParts").textContent = selected.zone_total || zoneAnalysis.total || 0;
    $("#groupActiveStocks").textContent = zoneAnalysis.red || 0;
    $("#groupCriticalParts").textContent = zoneAnalysis.red || 0;
    renderZoneChart({ red: zoneAnalysis.red || 0, yellow: 0, green: 0 });
    loadZoneParts("red").catch((error) => {
      $("#zoneChartStatus").textContent = error.message;
    });
    return;
  }
  const groupCards = selected.material_groups || {};
  const group = currentMaterialGroup ? groupCards[currentMaterialGroup] || {} : null;
  const zoneAnalysis = group ? group.zone_analysis || {} : selected.zone_analysis || {};
  const sourceParts = group ? group.source_parts_count || 0 : selected.master_parts_count || 0;
  $("#groupSourceParts").textContent = sourceParts;
  $("#groupActiveStocks").textContent = zoneAnalysis.active || (group ? group.active_stocks_count || 0 : selected.active_stocks_count || 0);
  $("#groupCriticalParts").textContent = zoneAnalysis.critical || (group ? group.low_stock_count || 0 : selected.low_stock_count || 0);
  renderZoneChart(zoneAnalysis);
  loadZoneParts(selectedZone).catch((error) => {
    $("#zoneChartStatus").textContent = error.message;
  });
}

function renderGroupCriticalRows(rows) {
  $("#groupCriticalRows").innerHTML = rows.length
    ? rows.map((row) => `
      <tr>
        <td>${fmt(row.material)}</td>
        <td>${groupLabel(row.material_group)}</td>
        <td>${fmt(row.description)}</td>
        <td>${fmt(row.current_stock)}</td>
        <td>${fmt(row.critical_value)}</td>
        <td>${fmt(row.net_consumption)}</td>
        <td>${fmt(row.last_quantity)} ${fmt(row.movement_type)}</td>
        <td>${fmt(row.reason)}</td>
      </tr>
    `).join("")
    : `<tr><td class="empty" colspan="8">No critical parts found for this value.</td></tr>`;
}

function renderExampleButtons(values) {
  $("#criticalExamples").innerHTML = values.length
    ? `Example values from current stock data: ${values.map((value) => `<button type="button" data-critical-value="${value}">${value}</button>`).join("")}`
    : "";
}

async function analyzeGroupCritical() {
  $("#criticalStatus").className = "status";
  const value = $("#criticalStockInput").value;
  if (!currentMaterialGroup) {
    $("#criticalStatus").textContent = "Select a material group to preview critical parts.";
    renderGroupCriticalRows([]);
    return;
  }
  const result = await api(`/api/group-critical?category=${encodeURIComponent(currentCategory)}&material_group=${encodeURIComponent(currentMaterialGroup)}&critical_stock=${encodeURIComponent(value)}`);
  renderExampleButtons(result.examples || []);
  renderGroupCriticalRows(result.rows || []);
  if (result.zone === "red") {
    $("#criticalStatus").textContent =
      `${result.critical_count || 0} red zone material(s) shown from ${result.total_parts || 0} stock-zone material(s).`;
    return;
  }
  if (result.source === "latest_upload_analysis") {
    const fileName = result.upload?.file_name || "latest uploaded Excel";
    $("#criticalStatus").textContent =
      `${result.critical_count || 0} critical part(s) found from ${result.total_parts || 0} analyzed stock item(s) in ${groupLabel(currentMaterialGroup)} using ${fileName}.`;
    return;
  }
  if (result.source === "stock_zone_analysis") {
    $("#criticalStatus").textContent =
      `${result.critical_count || 0} critical part(s) found from ${result.total_parts || 0} Excel-analyzed stock part(s) in ${currentMaterialGroup ? groupLabel(currentMaterialGroup) : "all material groups"}.`;
    return;
  }
  if (result.saved_criteria?.critical_stock !== undefined && !value) {
    $("#criticalStockInput").value = result.saved_criteria.critical_stock;
    $("#groupReorderInput").value = result.saved_criteria.reorder_quantity || "";
  }
  $("#criticalStatus").textContent =
    `${result.critical_count || 0} critical part(s) found from ${result.total_parts || 0} uploaded stock item(s) in ${groupLabel(currentMaterialGroup)}.`;
}

function renderAnalysis(result) {
  const upload = result.upload;
  const rows = result.failed_rows || [];
  if (!upload) {
    $("#analysisSummary").textContent = "Upload an Excel sheet to analyze stock against saved criteria.";
    renderFailed([]);
    return;
  }
  $("#analysisSummary").textContent =
    `${upload.file_name}: ${rows.length} material(s) did not match criteria. ${result.no_criteria_count || 0} row(s) had no criteria set.`;
  renderFailed(rows);
}

function number(value) {
  return Number(value || 0);
}

function inventoryNumber(value) {
  return new Intl.NumberFormat("en-IN", { maximumFractionDigits: 0 }).format(number(value));
}

function chartEmpty(svg, message) {
  svg.innerHTML = `<text class="empty-label" x="340" y="150" text-anchor="middle">${message}</text>`;
}

function renderInventoryPlantChart(plants) {
  const svg = $("#inventoryPlantChart");
  if (!plants.length) return chartEmpty(svg, "Upload an inventory Excel to view plant ageing.");
  const width = 680;
  const height = 300;
  const left = 58;
  const right = 22;
  const top = 26;
  const bottom = 58;
  const plotHeight = height - top - bottom;
  const maxValue = Math.max(1, ...plants.flatMap((row) => [number(row.qty_0_30), number(row.over_six_months), number(row.qty_over_2_years)]));
  const groupWidth = (width - left - right) / plants.length;
  const barWidth = Math.max(8, Math.min(28, groupWidth / 4));
  let markup = `<line class="axis" x1="${left}" y1="${height - bottom}" x2="${width - right}" y2="${height - bottom}"/>`;
  [0, 0.5, 1].forEach((ratio) => {
    const y = top + plotHeight * (1 - ratio);
    markup += `<line class="axis" x1="${left}" y1="${y}" x2="${width - right}" y2="${y}" opacity="${ratio ? 0.5 : 1}"/><text class="axis-label" x="${left - 8}" y="${y + 4}" text-anchor="end">${inventoryNumber(maxValue * ratio)}</text>`;
  });
  plants.forEach((row, index) => {
    const x = left + index * groupWidth + groupWidth / 2;
    const series = [
      { value: number(row.qty_0_30), color: "#3b82f6" },
      { value: number(row.over_six_months), color: "#e3a008" },
      { value: number(row.qty_over_2_years), color: "#c2410c" },
    ];
    series.forEach((item, seriesIndex) => {
      const barHeight = item.value * plotHeight / maxValue;
      const barX = x + (seriesIndex - 1) * (barWidth + 4) - barWidth / 2;
      markup += `<rect x="${barX}" y="${height - bottom - barHeight}" width="${barWidth}" height="${barHeight}" fill="${item.color}" rx="2"><title>${row.plant}: ${inventoryNumber(item.value)}</title></rect>`;
    });
    markup += `<text class="chart-label" x="${x}" y="${height - bottom + 22}" text-anchor="middle">${String(row.plant).replaceAll("&", "&amp;").replaceAll("<", "&lt;")}</text>`;
  });
  markup += `<text class="axis-label" x="${left}" y="16">0-30 days</text><rect x="${left + 58}" y="7" width="12" height="12" fill="#3b82f6"/><text class="axis-label" x="${left + 78}" y="16">Over 6 months</text><rect x="${left + 172}" y="7" width="12" height="12" fill="#e3a008"/><text class="axis-label" x="${left + 192}" y="16">Over 2 years</text><rect x="${left + 282}" y="7" width="12" height="12" fill="#c2410c"/>`;
  svg.innerHTML = markup;
}

function renderInventoryHistoryChart(history) {
  const svg = $("#inventoryHistoryChart");
  if (!history.length) return chartEmpty(svg, "Upload more than one inventory Excel to compare history.");
  const width = 680;
  const height = 300;
  const left = 58;
  const right = 22;
  const top = 30;
  const bottom = 64;
  const plotHeight = height - top - bottom;
  const plotWidth = width - left - right;
  const values = history.map((row) => number(row.total_value));
  const maxValue = Math.max(1, ...values);
  const points = values.map((value, index) => {
    const x = history.length === 1 ? left + plotWidth / 2 : left + index * plotWidth / (history.length - 1);
    const y = top + (1 - value / maxValue) * plotHeight;
    return { x, y, value, label: String(history[index].uploaded_at || "").slice(0, 10) };
  });
  let markup = `<line class="axis" x1="${left}" y1="${height - bottom}" x2="${width - right}" y2="${height - bottom}"/>`;
  [0, 0.5, 1].forEach((ratio) => {
    const y = top + plotHeight * (1 - ratio);
    markup += `<line class="axis" x1="${left}" y1="${y}" x2="${width - right}" y2="${y}" opacity="${ratio ? 0.5 : 1}"/><text class="axis-label" x="${left - 8}" y="${y + 4}" text-anchor="end">${inventoryNumber(maxValue * ratio)}</text>`;
  });
  markup += `<polyline fill="none" stroke="#0f8b8d" stroke-width="3" points="${points.map((point) => `${point.x},${point.y}`).join(" ")}"/>`;
  points.forEach((point) => {
    markup += `<circle cx="${point.x}" cy="${point.y}" r="5" fill="#0f8b8d"><title>${point.label}: ${inventoryNumber(point.value)}</title></circle><text class="axis-label" x="${point.x}" y="${height - bottom + 22}" text-anchor="middle">${point.label}</text>`;
  });
  markup += `<text class="axis-label" x="${left}" y="16">Total inventory value</text>`;
  svg.innerHTML = markup;
}

function renderInventory(data) {
  const plants = data.plants || [];
  const parts = data.parts || [];
  $("#inventoryPlantCount").textContent = plants.length;
  $("#inventoryPartCount").textContent = inventoryNumber(plants.reduce((sum, row) => sum + number(row.parts_count), 0));
  $("#inventoryOverSixMonths").textContent = inventoryNumber(plants.reduce((sum, row) => sum + number(row.over_six_months), 0));
  $("#inventoryOverTwoYears").textContent = inventoryNumber(plants.reduce((sum, row) => sum + number(row.qty_over_2_years), 0));
  $("#inventoryMethod").textContent = data.method || "Upload an agewise inventory Excel to start analysis.";
  $("#inventoryPlantRows").innerHTML = plants.length ? plants.map((row) => `
    <tr><td>${fmt(row.plant)}</td><td>${inventoryNumber(row.parts_count)}</td><td>${inventoryNumber(row.total_stock)}</td><td>${inventoryNumber(row.total_value)}</td><td>${inventoryNumber(row.qty_0_30)}</td><td>${inventoryNumber(row.over_six_months)}</td><td>${inventoryNumber(row.qty_over_2_years)}</td><td>${inventoryNumber(row.attention_parts)}</td></tr>
  `).join("") : `<tr><td class="empty" colspan="8">No inventory Excel has been analyzed yet.</td></tr>`;
  $("#inventoryPartRows").innerHTML = parts.length ? parts.map((row) => `
    <tr><td>${fmt(row.plant)}</td><td>${fmt(row.material)}</td><td>${fmt(row.description)}</td><td>${inventoryNumber(row.total_stock)}</td><td>${inventoryNumber(row.total_value)}</td><td>${inventoryNumber(row.qty_0_30)}</td><td>${inventoryNumber(row.over_six_months)}</td><td>${inventoryNumber(row.qty_over_2_years)}</td></tr>
  `).join("") : `<tr><td class="empty" colspan="8">No inventory parts to show yet.</td></tr>`;
  renderInventoryPlantChart(plants);
  renderInventoryHistoryChart(data.history || []);
}

function renderInventoryProgress(percent, message) {
  const value = Math.max(0, Math.min(100, Math.round(number(percent))));
  $("#inventoryProgress").hidden = false;
  $("#inventoryProgressBar").style.width = `${value}%`;
  $("#inventoryProgressPercent").textContent = `${value}%`;
  $("#inventoryProgressLabel").textContent = message || "Processing inventory Excel...";
}

function setInventoryBusy(isBusy) {
  $("#processInventoryButton").disabled = isBusy;
  $("#inventoryExcelFile").disabled = isBusy;
  $("#processInventoryButton").textContent = isBusy ? "Analyzing..." : "Analyze Inventory";
}

function stopInventoryProgressPolling() {
  if (inventoryProgressTimer) clearInterval(inventoryProgressTimer);
  inventoryProgressTimer = null;
}

function startInventoryProgressPolling(jobId) {
  stopInventoryProgressPolling();
  renderInventoryProgress(2, "Uploading inventory Excel...");
  inventoryProgressTimer = setInterval(async () => {
    try {
      const progress = await api(`/api/upload-progress?job_id=${encodeURIComponent(jobId)}`);
      renderInventoryProgress(progress.percent, progress.message);
      if (progress.state === "done" || progress.state === "error") stopInventoryProgressPolling();
    } catch (_) {
      renderInventoryProgress(2, "Uploading inventory Excel...");
    }
  }, 500);
}

async function refreshInventory() {
  renderInventory(await api("/api/inventory-analysis"));
}

async function refresh() {
  const groupParam = encodeURIComponent(currentMaterialGroup);
  const [summary, criteria, uploads, analysis, overrides] = await Promise.all([
    api("/api/summary"),
    api(`/api/criteria?category=${encodeURIComponent(currentCategory)}&material_group=${groupParam}`),
    api(`/api/uploads?category=${encodeURIComponent(currentCategory)}`),
    api(`/api/analysis?category=${encodeURIComponent(currentCategory)}&material_group=${groupParam}`),
    api("/api/stock-overrides"),
  ]);
  setSummary(summary);
  renderCriteria(criteria);
  renderUploads(uploads);
  renderAnalysis(analysis);
  renderStockOverrides(overrides);
}

async function loadSession() {
  const result = await api("/api/me");
  applyAuthState(result.user);
  if (result.user) {
    const groupResult = await api("/api/material-groups");
    renderMaterialGroupOptions(groupResult.groups || []);
    await refresh();
    showDashboard();
  }
}

let authMode = "login";

function updateAuthForm() {
  const isSignup = authMode === "signup";
  $("#authTitle").textContent = isSignup ? "Create Your Account" : "Login To Continue";
  $("#authSubmit").textContent = isSignup ? "Sign up" : "Login";
  $("#nameInput").required = isSignup;
  $("#roleField").classList.toggle("hidden", !isSignup);
  $("#passwordInput").placeholder = isSignup ? "Create a strong password" : "Enter password";
  $("#toggleAuthBtn").textContent = isSignup ? "Already have an account? Login" : "Create account";
}

$("#toggleAuthBtn").addEventListener("click", () => {
  authMode = authMode === "login" ? "signup" : "login";
  $("#loginStatus").textContent = "";
  updateAuthForm();
});

$("#loginForm").addEventListener("submit", async (event) => {
  event.preventDefault();
  const status = $("#loginStatus");
  status.className = "status";
  status.textContent = authMode === "signup" ? "Creating account..." : "Logging in...";
  try {
    const data = Object.fromEntries(new FormData(event.currentTarget).entries());
    const endpoint = authMode === "signup" ? "/api/signup" : "/api/login";
    const result = await api(endpoint, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(data),
    });
    status.textContent = "";
    applyAuthState(result.user);
    const groupResult = await api("/api/material-groups");
    renderMaterialGroupOptions(groupResult.groups || []);
    await refresh();
    showDashboard();
  } catch (error) {
    status.className = "status error";
    status.textContent = error.message;
  }
});

updateAuthForm();

$("#logoutBtn").addEventListener("click", async () => {
  await api("/api/logout", { method: "POST" });
  applyAuthState(null);
  showDashboard();
});

$("#criteriaForm").addEventListener("submit", async (event) => {
  event.preventDefault();
  const form = event.currentTarget;
  if (currentUser?.role !== "admin") {
    alert("Only the admin can edit stock criteria.");
    return;
  }
  const data = Object.fromEntries(new FormData(form).entries());
  data.active = form.active.checked;
  data.keep_stock = form.keep_stock.checked;
  try {
    const result = await api("/api/criteria", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(data),
    });
    setSummary(result.summary);
    form.reset();
    $("#criteriaCategory").value = currentCategory;
    form.active.checked = true;
    form.keep_stock.checked = false;
    await refresh();
  } catch (error) {
    alert(error.message);
  }
});

$("#criteriaAreaSelect").addEventListener("change", async (event) => {
  await selectCategory(event.target.value);
  location.hash = "criteria";
});

$("#criteriaRows").addEventListener("click", async (event) => {
  const button = event.target.closest("button[data-criteria-delete]");
  if (!button) return;
  if (currentUser?.role !== "admin") {
    alert("Only the admin can delete stock criteria.");
    return;
  }
  try {
    const material = button.dataset.criteriaDelete;
    const result = await api(`/api/criteria?category=${encodeURIComponent(currentCategory)}&material=${encodeURIComponent(material)}`, {
      method: "DELETE",
    });
    setSummary(result.summary);
    await refresh();
  } catch (error) {
    alert(error.message);
  }
});

$("#stockOverrideForm").addEventListener("submit", async (event) => {
  event.preventDefault();
  if (currentUser?.role !== "admin") {
    alert("Only the admin can edit stock overrides.");
    return;
  }
  const form = event.currentTarget;
  const data = Object.fromEntries(new FormData(form).entries());
  try {
    const result = await api("/api/stock-overrides", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(data),
    });
    setSummary(result.summary);
    renderStockOverrides(result.overrides);
    form.reset();
    form.category.value = currentCategory;
    await refresh();
  } catch (error) {
    alert(error.message);
  }
});

$("#stockOverrideRows").addEventListener("click", async (event) => {
  const button = event.target.closest("button[data-stock-override-delete]");
  if (!button) return;
  try {
    const material = button.dataset.stockOverrideDelete;
    const result = await api(`/api/stock-overrides?material=${encodeURIComponent(material)}`, {
      method: "DELETE",
    });
    setSummary(result.summary);
    renderStockOverrides(result.overrides);
    await refresh();
  } catch (error) {
    alert(error.message);
  }
});

$("#materialSearch").addEventListener("input", (event) => {
  window.clearTimeout(searchTimer);
  searchTimer = window.setTimeout(() => {
    searchMaterials(event.target.value).catch((error) => {
      $("#materialResults").innerHTML = `<div class="result-empty">${error.message}</div>`;
    });
  }, 250);
});

$("#materialGroupSelect").addEventListener("change", async (event) => {
  currentMaterialGroup = event.target.value;
  selectedZone = currentMaterialGroup === RED_ZONE_VALUE ? "red" : selectedZone;
  $("#materialSearch").value = "";
  $("#criticalStockInput").value = "";
  $("#groupReorderInput").value = "";
  updateCriticalFormState();
  renderMaterialResults([]);
  renderMaterialGroupStats();
  await refresh();
  await analyzeGroupCritical();
});

$("#zoneLineChart").addEventListener("click", (event) => {
  const target = event.target.closest("[data-zone]");
  if (!target) return;
  selectedZone = target.dataset.zone;
  renderZoneChart(currentZoneAnalysis());
  loadZoneParts(selectedZone).catch((error) => {
    $("#zoneChartStatus").textContent = error.message;
  });
});

$("#zoneLineChart").addEventListener("mouseover", (event) => {
  const target = event.target.closest("[data-zone]");
  if (target) showZoneTooltip(target);
});

$("#zoneLineChart").addEventListener("mouseout", hideZoneTooltip);

$("#zoneLegend").addEventListener("click", (event) => {
  const button = event.target.closest("button[data-zone]");
  if (!button) return;
  selectedZone = button.dataset.zone;
  renderZoneChart(currentZoneAnalysis());
  loadZoneParts(selectedZone).catch((error) => {
    $("#zoneChartStatus").textContent = error.message;
  });
});

$("#categorySelect").addEventListener("change", async (event) => {
  selectCategory(event.target.value).catch((error) => {
    $("#uploadStatus").className = "status error";
    $("#uploadStatus").textContent = error.message;
  });
});

$("#criticalStockInput").addEventListener("input", () => {
  window.clearTimeout(criticalTimer);
  criticalTimer = window.setTimeout(() => {
    analyzeGroupCritical().catch((error) => {
      $("#criticalStatus").textContent = error.message;
      $("#criticalStatus").className = "status error";
    });
  }, 250);
});

$("#criticalExamples").addEventListener("click", (event) => {
  const button = event.target.closest("button[data-critical-value]");
  if (!button) return;
  $("#criticalStockInput").value = button.dataset.criticalValue;
  analyzeGroupCritical().catch((error) => {
    $("#criticalStatus").textContent = error.message;
  });
});

$("#groupCriticalForm").addEventListener("submit", async (event) => {
  event.preventDefault();
  if (isRedZoneSelected()) {
    $("#criticalStatus").textContent = "Red Zone Material uses the 330 material zone data and does not need saved criteria.";
    return;
  }
  if (currentUser?.role !== "admin") {
    alert("Only the admin can save group criteria.");
    return;
  }
  try {
    const result = await api("/api/group-critical", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        category: currentCategory,
        material_group: currentMaterialGroup,
        critical_stock: $("#criticalStockInput").value,
        reorder_quantity: $("#groupReorderInput").value,
      }),
    });
    renderGroupCriticalRows(result.analysis.rows || []);
    renderMaterialGroupStats();
    $("#criticalStatus").className = "status";
    $("#criticalStatus").textContent = "Group critical value saved and analyzed.";
  } catch (error) {
    $("#criticalStatus").className = "status error";
    $("#criticalStatus").textContent = error.message;
  }
});

$("#materialResults").addEventListener("click", (event) => {
  const button = event.target.closest(".material-result");
  if (!button) return;
  const form = $("#criteriaForm");
  form.material.value = button.dataset.material || "";
  form.description.value = button.dataset.description || "";
  form.minimum_stock.value = button.dataset.minimumStock || "";
  form.reorder_quantity.value = button.dataset.reorderQuantity || "";
  form.active.checked = true;
  form.keep_stock.checked = button.dataset.keepStock === "true";
});

$("#uploadForm").addEventListener("submit", async (event) => {
  event.preventDefault();
  $("#uploadCategory").value = currentCategory;
  const status = $("#uploadStatus");
  status.className = "status";
  status.textContent = "Processing Excel...";
  const categoryLabel = categoryLabels[currentCategory] || "selected category";
  const jobId = createJobId();
  const formData = new FormData(event.currentTarget);
  formData.set("job_id", jobId);
  setUploadBusy(true);
  startUploadProgressPolling(jobId, categoryLabel);
  try {
    const result = await api("/api/upload", {
      method: "POST",
      body: formData,
    });
    const categoryMessage = result.upload.material_group_column_found
      ? ` Material groups updated from Excel (${result.upload.learned_material_groups} group(s)).`
      : " No material group column found; using saved master lookup.";
    stopUploadProgressPolling();
    renderUploadProgress(100, `${categoryLabel}: Excel processing completed.`);
    status.textContent = `Imported ${result.upload.imported_rows} rows. Failed criteria: ${result.upload.failed_count}.${categoryMessage}`;
    renderUploadActions(result.upload);
    setSummary(result.summary);
    renderAnalysis({
      upload: result.upload,
      failed_rows: result.failed_rows || [],
      no_criteria_count: result.upload.no_criteria_count || 0,
    });
    await refresh();
  } catch (error) {
    stopUploadProgressPolling();
    renderUploadProgress(100, `${categoryLabel}: Processing stopped.`);
    status.className = "status error";
    status.textContent = error.message;
  } finally {
    setUploadBusy(false);
  }
});

$("#inventoryUploadForm").addEventListener("submit", async (event) => {
  event.preventDefault();
  const status = $("#inventoryStatus");
  status.className = "status";
  status.textContent = "Analyzing inventory Excel...";
  const jobId = createJobId();
  const formData = new FormData(event.currentTarget);
  formData.set("job_id", jobId);
  setInventoryBusy(true);
  startInventoryProgressPolling(jobId);
  try {
    const result = await api("/api/inventory-upload", { method: "POST", body: formData });
    stopInventoryProgressPolling();
    renderInventoryProgress(100, "Inventory analysis completed.");
    status.textContent = `Analyzed ${result.upload.parts_count} parts across ${result.upload.plants_count} plant(s).`;
    renderInventory(result.analysis);
  } catch (error) {
    stopInventoryProgressPolling();
    renderInventoryProgress(100, "Inventory processing stopped.");
    status.className = "status error";
    status.textContent = error.message;
  } finally {
    setInventoryBusy(false);
  }
});

renderFailed([]);
renderMaterialResults([]);
renderGroupCriticalRows([]);
renderUploadActions(null);
renderInventory({});
applyAuthState(null);
loadSession().catch((error) => {
  $("#uploadStatus").className = "status error";
  $("#uploadStatus").textContent = error.message;
});

document.querySelectorAll(".category-card").forEach((button) => {
  button.addEventListener("click", () => {
    selectCategory(button.dataset.category).catch((error) => {
      $("#uploadStatus").className = "status error";
      $("#uploadStatus").textContent = error.message;
    });
  });
});

$("#changeCategoryBtn").addEventListener("click", showDashboard);
$("#inventoryLaunchButton").addEventListener("click", showInventoryWorkspace);
$("#inventoryBackButton").addEventListener("click", showDashboard);
$("#inventoryLink").addEventListener("click", (event) => {
  event.preventDefault();
  showInventoryWorkspace();
});
$("#dashboardLink").addEventListener("click", (event) => {
  event.preventDefault();
  showDashboard();
});
