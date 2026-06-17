const $ = (selector) => document.querySelector(selector);
let currentUser = null;
let searchTimer = null;
let currentCategory = "ml_spare";
let currentMaterialGroup = "";
let latestSummary = null;
let materialGroups = [];
let criticalTimer = null;

const categoryLabels = {
  ml_spare: "M/L Spare",
  tools: "Tools",
};

async function api(path, options = {}) {
  const response = await fetch(path, options);
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(payload.error || "Request failed");
  }
  return payload;
}

function fmt(value) {
  if (value === undefined || value === null || value === "") return "-";
  return value;
}

function setSummary(summary) {
  latestSummary = summary;
  renderCategoryCards(summary.categories || {});
  const selected = summary.categories?.[currentCategory] || {};
  $("#materialsCount").textContent = selected.materials_count || 0;
  $("#criteriaCount").textContent = selected.criteria_count || 0;
  $("#lowStockCount").textContent = selected.low_stock_count || 0;
  $("#activeStockCount").textContent = selected.active_stocks_count || 0;
  renderMaterialGroupStats();
}

function roleLabel(role) {
  if (role === "admin") return "Admin";
  if (role === "employee") return "Employee";
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
  $("#criteriaHint").textContent = canEditCriteria
    ? "The report includes materials where current stock is at or below the minimum stock."
    : "Only the admin can set or edit stock criteria.";
  $("#groupCriticalForm").querySelectorAll("input, button").forEach((element) => {
    element.disabled = !canEditCriteria;
  });
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
  $("#workflow").hidden = true;
  $("#dashboard").hidden = false;
}

async function selectCategory(category) {
  currentCategory = categoryLabels[category] ? category : "ml_spare";
  $("#selectedCategoryTitle").textContent = categoryLabels[currentCategory];
  $("#uploadCategory").value = currentCategory;
  $("#criteriaCategory").value = currentCategory;
  $("#materialSearch").value = "";
  renderMaterialResults([]);
  $("#dashboard").hidden = true;
  $("#workflow").hidden = false;
  await refresh();
}

function renderCriteria(criteria) {
  const rows = Object.values(criteria);
  $("#criteriaRows").innerHTML = rows.length
    ? rows.map((row) => `
      <tr>
        <td>${fmt(row.material)}</td>
        <td>${fmt(row.material_group)}</td>
        <td>${fmt(row.description)}</td>
        <td>${fmt(row.minimum_stock)}</td>
        <td>${fmt(row.reorder_quantity)}</td>
        <td>${row.active ? "Active" : "Inactive"}</td>
      </tr>
    `).join("")
    : `<tr><td class="empty" colspan="6">No criteria added yet.</td></tr>`;
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
        <td><a class="download" href="/api/download?file=${encodeURIComponent(row.report_file)}">Download Excel</a></td>
      </tr>
    `).join("")
    : `<tr><td class="empty" colspan="7">No uploads processed yet.</td></tr>`;
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
      data-reorder-quantity="${row.reorder_quantity ?? ""}">
      <strong>${fmt(row.material)}</strong>
      <span>${fmt(row.description)}</span>
      <small>Group: ${fmt(row.material_group)}</small>
      <small>Stock: ${fmt(row.current_stock)}</small>
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
        <td>${fmt(row.material_group)}</td>
        <td>${fmt(row.description)}</td>
        <td>${fmt(row.current_stock)}</td>
        <td>${fmt(row.minimum_stock)}</td>
        <td>${fmt(row.reorder_quantity)}</td>
        <td>${fmt(row.reason)}</td>
      </tr>
    `).join("")
    : `<tr><td class="empty" colspan="7">Upload a file to see failed criteria here.</td></tr>`;
}

function renderMaterialGroupOptions(groups) {
  materialGroups = groups || [];
  $("#materialGroupSelect").innerHTML = [
    `<option value="">All material groups</option>`,
    ...materialGroups.map((group) => (
      `<option value="${String(group.code).replaceAll('"', "&quot;")}">${fmt(group.label)}</option>`
    )),
  ].join("");
}

function renderMaterialGroupStats() {
  const selected = latestSummary?.categories?.[currentCategory] || {};
  const groupCards = selected.material_groups || {};
  const group = currentMaterialGroup ? groupCards[currentMaterialGroup] || {} : null;
  $("#groupSourceParts").textContent = group ? group.source_parts_count || 0 : materialGroups.reduce((sum, item) => sum + (Number(item.source_parts_count) || 0), 0);
  $("#groupTotalParts").textContent = group ? group.materials_count || 0 : selected.materials_count || 0;
  $("#groupCriteriaSet").textContent = group ? group.criteria_count || 0 : selected.criteria_count || 0;
  $("#groupCriticalParts").textContent = group ? group.low_stock_count || 0 : selected.low_stock_count || 0;
  $("#groupActiveStocks").textContent = group ? group.active_stocks_count || 0 : selected.active_stocks_count || 0;
}

function renderGroupCriticalRows(rows) {
  $("#groupCriticalRows").innerHTML = rows.length
    ? rows.map((row) => `
      <tr>
        <td>${fmt(row.material)}</td>
        <td>${fmt(row.material_group)}</td>
        <td>${fmt(row.description)}</td>
        <td>${fmt(row.current_stock)}</td>
        <td>${fmt(row.last_quantity)} ${fmt(row.movement_type)}</td>
        <td>${fmt(row.reason)}</td>
      </tr>
    `).join("")
    : `<tr><td class="empty" colspan="6">No critical parts found for this value.</td></tr>`;
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
  $("#groupCriticalParts").textContent = result.critical_count || 0;
  if (result.saved_criteria?.critical_stock !== undefined && !value) {
    $("#criticalStockInput").value = result.saved_criteria.critical_stock;
    $("#groupReorderInput").value = result.saved_criteria.reorder_quantity || "";
  }
  $("#criticalStatus").textContent =
    `${result.critical_count || 0} critical part(s) found from ${result.total_parts || 0} uploaded stock item(s) in ${currentMaterialGroup}.`;
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

async function refresh() {
  const groupParam = encodeURIComponent(currentMaterialGroup);
  const [summary, criteria, uploads, analysis] = await Promise.all([
    api("/api/summary"),
    api(`/api/criteria?category=${encodeURIComponent(currentCategory)}&material_group=${groupParam}`),
    api(`/api/uploads?category=${encodeURIComponent(currentCategory)}`),
    api(`/api/analysis?category=${encodeURIComponent(currentCategory)}&material_group=${groupParam}`),
  ]);
  setSummary(summary);
  renderCriteria(criteria);
  renderUploads(uploads);
  renderAnalysis(analysis);
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

$("#loginForm").addEventListener("submit", async (event) => {
  event.preventDefault();
  const status = $("#loginStatus");
  status.className = "status";
  status.textContent = "Logging in...";
  try {
    const data = Object.fromEntries(new FormData(event.currentTarget).entries());
    const result = await api("/api/login", {
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
  $("#materialSearch").value = "";
  $("#criticalStockInput").value = "";
  $("#groupReorderInput").value = "";
  renderMaterialResults([]);
  renderMaterialGroupStats();
  await refresh();
  await analyzeGroupCritical();
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
    $("#groupCriticalParts").textContent = result.analysis.critical_count || 0;
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
});

$("#uploadForm").addEventListener("submit", async (event) => {
  event.preventDefault();
  const status = $("#uploadStatus");
  status.className = "status";
  status.textContent = "Processing Excel...";
  const formData = new FormData(event.currentTarget);
  try {
    const result = await api("/api/upload", {
      method: "POST",
      body: formData,
    });
    status.textContent = `Imported ${result.upload.imported_rows} rows. Failed criteria: ${result.upload.failed_count}.`;
    setSummary(result.summary);
    renderAnalysis({
      upload: result.upload,
      failed_rows: result.failed_rows || [],
      no_criteria_count: result.upload.no_criteria_count || 0,
    });
    await refresh();
  } catch (error) {
    status.className = "status error";
    status.textContent = error.message;
  }
});

renderFailed([]);
renderMaterialResults([]);
renderGroupCriticalRows([]);
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
$("#dashboardLink").addEventListener("click", (event) => {
  event.preventDefault();
  showDashboard();
});
