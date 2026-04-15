const $ = (id) => document.getElementById(id);

let summaryChart = null;

let assetsState = {
  cars: [],
  investments: [],
  leumiBalance: 0,
};
let editingAsset = null;
let reviewState = null;
let progressPollTimer = null;

function stopProgressPolling() {
  if (progressPollTimer) {
    clearInterval(progressPollTimer);
    progressPollTimer = null;
  }
}

function generateRunToken() {
  if (window.crypto && typeof window.crypto.randomUUID === "function") {
    return window.crypto.randomUUID();
  }
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

async function pollProgress(runToken) {
  if (!runToken) return;
  try {
    const resp = await fetch(`/run-month/progress/${encodeURIComponent(runToken)}`);
    if (!resp.ok) return;
    const data = await resp.json();
    const step = String(data?.step || "").trim();
    if (step) {
      setStatus(`Current step: ${step}`, data?.done ? "ok" : "info");
    }
  } catch (_err) {
    // Ignore transient polling failures.
  }
}

function startProgressPolling(runToken) {
  stopProgressPolling();
  void pollProgress(runToken);
  progressPollTimer = setInterval(() => {
    void pollProgress(runToken);
  }, 1200);
}

function setStatus(msg, kind = "info") {
  const el = $("status");
  el.textContent = msg || "";
  el.classList.remove("status-error", "status-ok");
  if (kind === "error") el.classList.add("status-error");
  if (kind === "ok") el.classList.add("status-ok");
}

function getSelectedYearMonth() {
  const year = $("year").value;
  const month = Number($("month").value);
  return { year, month };
}

function monthName(m) {
  const names = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
  ];
  return names[m - 1] || String(m);
}

function safeNumber(x) {
  if (x === null || x === undefined) return 0;
  if (typeof x === "number") return Number.isFinite(x) ? x : 0;
  if (typeof x === "string") {
    const cleaned = x
      .replaceAll("\u200f", "")
      .replaceAll("\u200e", "")
      .trim();

    // Try a direct parse first for normal numeric strings.
    const direct = Number(cleaned.replaceAll(",", ""));
    if (Number.isFinite(direct)) return direct;

    // Fallback: extract first numeric token from messy text (e.g. "₪ 12,345.67").
    const token = cleaned.match(/[-+]?\d[\d,]*(?:\.\d+)?/);
    if (!token) return 0;
    const n = Number(token[0].replaceAll(",", ""));
    return Number.isFinite(n) ? n : 0;
  }
  const n = Number(x);
  return Number.isFinite(n) ? n : 0;
}

function toMoneyString(x) {
  const n = safeNumber(x);
  // Keep it simple; backend stores numbers as strings occasionally.
  return new Intl.NumberFormat(undefined, {
    maximumFractionDigits: 2,
  }).format(n);
}

function randomColors(count) {
  // Deterministic-ish palette: spread hues across 360 degrees.
  const colors = [];
  for (let i = 0; i < count; i++) {
    const hue = Math.round((i * 360) / Math.max(1, count));
    colors.push(`hsl(${hue} 70% 55%)`);
  }
  return colors;
}

function renderAssets() {
  const carsList = $("carsList");
  const investmentsList = $("investmentsList");
  carsList.innerHTML = "";
  investmentsList.innerHTML = "";

  const carsTotal = assetsState.cars.reduce((acc, c) => acc + safeNumber(c.value), 0);
  const investmentsTotal = assetsState.investments.reduce(
    (acc, i) => acc + safeNumber(i.value),
    0
  );
  const totalNetWorth = safeNumber(assetsState.leumiBalance) + carsTotal + investmentsTotal;
  $("carsTotal").textContent = toMoneyString(carsTotal);
  $("investmentsTotal").textContent = toMoneyString(investmentsTotal);
  $("totalNetWorth").textContent = toMoneyString(totalNetWorth);

  for (const [carIndex, car] of assetsState.cars.entries()) {
    const row = document.createElement("div");
    row.className = "list-row";
    const isEditing =
      editingAsset?.kind === "cars" && editingAsset?.index === carIndex;
    row.innerHTML = `
      <div class="list-main">
        <div class="list-title">${escapeHtml(car.name ?? "")}</div>
        ${
          isEditing
            ? `<div class="list-edit-row">
                <input
                  class="asset-edit-input"
                  type="number"
                  inputmode="decimal"
                  step="0.01"
                  value="${escapeAttr(String(safeNumber(car.value)))}"
                />
              </div>`
            : `<div class="list-sub">${toMoneyString(car.value)}</div>`
        }
      </div>
      <div class="list-actions">
        ${
          isEditing
            ? `
              <button class="primary" data-action="edit-save" data-kind="cars" data-index="${carIndex}">Save</button>
              <button class="secondary" data-action="edit-cancel" data-kind="cars" data-index="${carIndex}">Cancel</button>
            `
            : `
              <button class="secondary" data-action="edit" data-kind="cars" data-index="${carIndex}">Edit</button>
            `
        }
        <button
          class="danger"
          data-action="remove"
          data-kind="cars"
          data-index="${carIndex}"
        >
          Remove
        </button>
      </div>
    `;
    carsList.appendChild(row);
  }

  for (const [invIndex, inv] of assetsState.investments.entries()) {
    const row = document.createElement("div");
    row.className = "list-row";
    const isEditing =
      editingAsset?.kind === "investments" && editingAsset?.index === invIndex;
    row.innerHTML = `
      <div class="list-main">
        <div class="list-title">${escapeHtml(inv.name ?? "")}</div>
        ${
          isEditing
            ? `<div class="list-edit-row">
                <input
                  class="asset-edit-input"
                  type="number"
                  inputmode="decimal"
                  step="0.01"
                  value="${escapeAttr(String(safeNumber(inv.value)))}"
                />
              </div>`
            : `<div class="list-sub">${toMoneyString(inv.value)}</div>`
        }
      </div>
      <div class="list-actions">
        ${
          isEditing
            ? `
              <button class="primary" data-action="edit-save" data-kind="investments" data-index="${invIndex}">Save</button>
              <button class="secondary" data-action="edit-cancel" data-kind="investments" data-index="${invIndex}">Cancel</button>
            `
            : `
              <button class="secondary" data-action="edit" data-kind="investments" data-index="${invIndex}">Edit</button>
            `
        }
        <button
          class="danger"
          data-action="remove"
          data-kind="investments"
          data-index="${invIndex}"
        >
          Remove
        </button>
      </div>
    `;
    investmentsList.appendChild(row);
  }
}

function escapeHtml(str) {
  return String(str)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function escapeAttr(str) {
  // For attribute values in our simple HTML template.
  return escapeHtml(str).replaceAll("`", "&#096;");
}

function normalizeExpenseName(name) {
  return String(name || "")
    .replaceAll("\u200f", "")
    .replaceAll("\u200e", "")
    .replace(/\s+/g, " ")
    .trim();
}

function isSplitExpenseName(name) {
  return normalizeExpenseName(name) === "העברה דיגיטל";
}

function removeReviewPanel() {
  const el = $("reviewPanel");
  if (el) el.remove();
  reviewState = null;
}

function createSplitPartRow(partNo, amount, category, allowedCategories) {
  const row = document.createElement("div");
  row.className = "split-row";
  row.innerHTML = `
    <span class="split-label">Part ${partNo}</span>
    <input class="split-amount" type="number" step="0.01" value="${escapeAttr(String(amount))}" />
    <select class="split-category">
      ${allowedCategories
        .map(
          (cat) =>
            `<option value="${escapeAttr(cat)}"${cat === category ? " selected" : ""}>${escapeHtml(cat)}</option>`
        )
        .join("")}
    </select>
  `;
  return row;
}

function collectReviewedItemsFromPanel() {
  const panel = $("reviewPanel");
  if (!panel || !reviewState) {
    throw new Error("Review panel is not available.");
  }

  const reviewedItems = [];
  const rows = panel.querySelectorAll(".review-row");
  for (const row of rows) {
    const include = row.querySelector(".review-include");
    if (!include || !include.checked) continue;

    const name = row.getAttribute("data-name") || "";
    const cost = safeNumber(row.getAttribute("data-cost"));
    const isSplit = row.getAttribute("data-is-split") === "1";

    if (isSplit) {
      const splitRows = row.querySelectorAll(".split-row");
      let splitTotal = 0;
      let added = 0;
      let idx = 1;
      for (const splitRow of splitRows) {
        const amountEl = splitRow.querySelector(".split-amount");
        const categoryEl = splitRow.querySelector(".split-category");
        const amount = Math.abs(safeNumber(amountEl?.value));
        const category = String(categoryEl?.value || "").trim();
        if (amount > 0) {
          reviewedItems.push({
            name: `${name} חלק ${idx}`,
            cost: amount,
            category,
          });
          splitTotal += amount;
          added += 1;
        }
        idx += 1;
      }
      if (added === 0) {
        throw new Error(`Please enter at least one positive split amount for '${name}'.`);
      }
      if (Math.abs(splitTotal - Math.abs(cost)) > 0.01) {
        throw new Error(
          `Split total for '${name}' must equal original amount (${cost}). Current total: ${splitTotal}.`
        );
      }
    } else {
      const categoryEl = row.querySelector(".review-category");
      reviewedItems.push({
        name,
        cost: Math.abs(cost),
        category: String(categoryEl?.value || "").trim(),
      });
    }
  }
  return reviewedItems;
}

function renderReviewPanel(data, year, month) {
  removeReviewPanel();
  reviewState = {
    runToken: data.run_token,
    reviewToken: data.review_token,
    outputExcel: data.output_excel,
    allowedCategories: Array.isArray(data.allowed_categories) ? data.allowed_categories : [],
  };

  const panel = document.createElement("div");
  panel.id = "reviewPanel";
  panel.className = "card";

  const parseErrors = Array.isArray(data.parse_errors) ? data.parse_errors : [];
  const items = Array.isArray(data.items) ? data.items : [];
  const allowedCategories = reviewState.allowedCategories;

  panel.innerHTML = `
    <h2>Review categories before fill (${monthName(month)} ${year})</h2>
    <div class="muted">${items.length} item(s) ready for review.</div>
    ${parseErrors.length > 0 ? `<div class="status status-error">Skipped ${parseErrors.length} malformed line(s).</div>` : ""}
    <div class="review-table-wrap">
      <table class="review-table">
        <thead>
          <tr>
            <th>Expense</th>
            <th>Amount</th>
            <th>Category / Split</th>
            <th>Include</th>
          </tr>
        </thead>
        <tbody id="reviewTbody"></tbody>
      </table>
    </div>
    <div class="review-actions">
      <button id="applyReviewBtn" class="primary">Apply and Fill</button>
      <button id="cancelReviewBtn" class="secondary">Cancel</button>
    </div>
  `;

  const tbody = panel.querySelector("#reviewTbody");
  for (const item of items) {
    const tr = document.createElement("tr");
    tr.className = "review-row";
    tr.setAttribute("data-name", String(item.name || ""));
    tr.setAttribute("data-cost", String(item.cost ?? 0));
    tr.setAttribute("data-is-split", isSplitExpenseName(item.name) ? "1" : "0");

    const duplicateBadge = item.is_possible_duplicate
      ? `<div class="muted review-warning">Possible duplicate</div>`
      : "";
    const includeChecked = item.is_possible_duplicate ? "" : "checked";

    if (isSplitExpenseName(item.name)) {
      const defaultCategory =
        allowedCategories.includes(item.category) && item.category
          ? item.category
          : allowedCategories[0] || "";
      tr.innerHTML = `
        <td>
          ${escapeHtml(String(item.name || ""))}
          ${duplicateBadge}
        </td>
        <td>${toMoneyString(item.cost)}</td>
        <td>
          <div class="split-box">
            <div class="split-list"></div>
            <div class="split-actions">
              <button type="button" class="secondary split-add-btn">Add split row</button>
              <button type="button" class="secondary split-remove-btn">Remove last row</button>
            </div>
          </div>
        </td>
        <td><input class="review-include" type="checkbox" ${includeChecked} /></td>
      `;
      const splitList = tr.querySelector(".split-list");
      splitList.appendChild(createSplitPartRow(1, item.cost, defaultCategory, allowedCategories));
      splitList.appendChild(createSplitPartRow(2, 0, defaultCategory, allowedCategories));
      tr.querySelector(".split-add-btn")?.addEventListener("click", () => {
        const nextNo = splitList.querySelectorAll(".split-row").length + 1;
        splitList.appendChild(createSplitPartRow(nextNo, 0, defaultCategory, allowedCategories));
      });
      tr.querySelector(".split-remove-btn")?.addEventListener("click", () => {
        const rows = splitList.querySelectorAll(".split-row");
        if (rows.length <= 1) return;
        rows[rows.length - 1].remove();
      });
    } else {
      tr.innerHTML = `
        <td>
          ${escapeHtml(String(item.name || ""))}
          ${duplicateBadge}
        </td>
        <td>${toMoneyString(item.cost)}</td>
        <td>
          <select class="review-category">
            ${allowedCategories
              .map((cat) => {
                const selected = cat === item.category ? " selected" : "";
                return `<option value="${escapeAttr(cat)}"${selected}>${escapeHtml(cat)}</option>`;
              })
              .join("")}
          </select>
        </td>
        <td><input class="review-include" type="checkbox" ${includeChecked} /></td>
      `;
    }
    tbody.appendChild(tr);
  }

  const mountPoint = document.querySelector(".controls");
  mountPoint?.insertAdjacentElement("afterend", panel);

  $("cancelReviewBtn")?.addEventListener("click", () => {
    removeReviewPanel();
    setStatus("Review canceled. No workbook changes were applied.", "info");
    $("runBtn").disabled = false;
  });

  $("applyReviewBtn")?.addEventListener("click", async () => {
    try {
      const reviewedItems = collectReviewedItemsFromPanel();
      if (reviewedItems.length === 0) {
        throw new Error("Nothing selected to include.");
      }
      setStatus("Applying reviewed categories and filling workbook...", "info");
      $("applyReviewBtn").disabled = true;
      $("cancelReviewBtn").disabled = true;
      startProgressPolling(reviewState.runToken);

      const finalizeResp = await fetch("/run-month/finalize", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          run_token: reviewState.runToken,
          review_token: reviewState.reviewToken,
          reviewed_items: reviewedItems,
        }),
      });
      const finalizeData = await finalizeResp.json();
      if (!finalizeResp.ok) {
        throw new Error(finalizeData?.detail || "Failed to finalize reviewed expenses.");
      }

      removeReviewPanel();
      setStatus(`Done. Output: ${finalizeData.output_excel}`, "ok");
      await loadStateAndAssets();
      await loadSummaryChart();
    } catch (err) {
      setStatus(String(err?.message || err), "error");
      $("applyReviewBtn").disabled = false;
      $("cancelReviewBtn").disabled = false;
    } finally {
      stopProgressPolling();
      $("runBtn").disabled = false;
    }
  });
}

async function loadStateAndAssets() {
  setStatus("Loading state/assets...", "info");
  const [stateResp, assetsResp] = await Promise.all([
    fetch("/state"),
    fetch("/assets"),
  ]);

  const state = await stateResp.json();
  const assets = await assetsResp.json();

  const lastYear = state?.last_filled_year;
  const lastMonth = state?.last_filled_month;
  if (lastYear && lastMonth) {
    $("lastFilled").textContent = `Last filled: ${monthName(Number(lastMonth))} ${lastYear}`;
  } else {
    $("lastFilled").textContent = "";
  }

  $("leumiBalance").textContent = assets?.leumi_balance
    ? String(assets.leumi_balance)
    : "-";
  assetsState.leumiBalance = safeNumber(assets?.leumi_balance);

  assetsState.cars = Array.isArray(assets?.cars) ? assets.cars : [];
  assetsState.investments = Array.isArray(assets?.investments)
    ? assets.investments
    : [];

  renderAssets();
}

function computeNextMonthFromState(state) {
  const lastYear = Number(state?.last_filled_year);
  const lastMonth = Number(state?.last_filled_month);
  if (!Number.isFinite(lastYear) || !Number.isFinite(lastMonth) || lastMonth < 1 || lastMonth > 12) {
    // Fallback: current year/month.
    const d = new Date();
    return { year: String(d.getFullYear()), month: d.getMonth() + 1 };
  }
  const next = new Date(Number(lastYear), lastMonth, 1); // lastMonth is 1-based
  return { year: String(next.getFullYear()), month: next.getMonth() + 1 };
}

async function loadSummaryChart() {
  const { year, month } = getSelectedYearMonth();
  setStatus(`Loading category summary for ${monthName(month)} ${year}...`, "info");

  const resp = await fetch(`/expenses/summary?year=${encodeURIComponent(year)}&month=${encodeURIComponent(month)}`);
  const summary = await resp.json();
  const entries = Object.entries(summary || {});

  if (entries.length === 0) {
    setStatus("No summary available for the selected month/workbook yet.", "error");
    if (summaryChart) summaryChart.data = { labels: [], datasets: [] };
    if (summaryChart) summaryChart.update();
    return;
  }

  entries.sort((a, b) => (safeNumber(b[1]) - safeNumber(a[1])));
  const top = entries.slice(0, 8);
  const other = entries.slice(8);
  const otherSum = other.reduce((acc, e) => acc + safeNumber(e[1]), 0);

  const labels = top.map((e) => e[0]);
  const values = top.map((e) => safeNumber(e[1]));
  if (otherSum > 0) {
    labels.push("Other");
    values.push(otherSum);
  }

  const colors = randomColors(labels.length);

  const ctx = $("summaryChart").getContext("2d");
  if (!summaryChart) {
    summaryChart = new Chart(ctx, {
      type: "pie",
      data: {
        labels,
        datasets: [
          {
            data: values,
            backgroundColor: colors,
          },
        ],
      },
      options: {
        responsive: true,
        plugins: {
          legend: { position: "bottom" },
        },
      },
    });
  } else {
    summaryChart.data.labels = labels;
    summaryChart.data.datasets[0].data = values;
    summaryChart.data.datasets[0].backgroundColor = colors;
    summaryChart.update();
  }

  setStatus("Ready.", "ok");
}

async function saveAssets() {
  const payload = {
    cars: assetsState.cars,
    investments: assetsState.investments,
  };
  const resp = await fetch("/assets", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  if (!resp.ok) throw new Error(`Failed to save assets: ${resp.status}`);
  await resp.json();
}

async function addCar() {
  const name = $("carName").value.trim();
  const value = safeNumber($("carValue").value);
  if (!name) return setStatus("Car name is required.", "error");
  if (value <= 0) return setStatus("Car value must be > 0.", "error");

  assetsState.cars.push({ name, value });
  $("carName").value = "";
  $("carValue").value = "";
  setStatus("Saving cars...", "info");
  await saveAssets();
  renderAssets();
  setStatus("Cars saved.", "ok");
}

async function addInvestment() {
  const name = $("investmentName").value.trim();
  const value = safeNumber($("investmentValue").value);
  if (!name) return setStatus("Investment name is required.", "error");
  if (value <= 0) return setStatus("Investment value must be > 0.", "error");

  assetsState.investments.push({ name, value });
  $("investmentName").value = "";
  $("investmentValue").value = "";
  setStatus("Saving investments...", "info");
  await saveAssets();
  renderAssets();
  setStatus("Investments saved.", "ok");
}

function wireRemoveButtons() {
  // Event delegation for dynamically rendered list actions.
  document.addEventListener("click", async (e) => {
    const btn = e.target.closest("button[data-kind][data-action][data-index]");
    if (!btn) return;

    const action = btn.getAttribute("data-action");
    const kind = btn.getAttribute("data-kind");
    const index = Number(btn.getAttribute("data-index"));
    if (!action || !kind || !Number.isInteger(index) || index < 0) return;

    if (kind !== "cars" && kind !== "investments") return;
    const list = assetsState[kind];
    if (index >= list.length) return;

    if (action === "edit") {
      editingAsset = { kind, index };
      renderAssets();
      const activeInput = document.querySelector(
        `.list-actions button[data-kind="${kind}"][data-index="${index}"][data-action="edit-save"]`
      )
        ?.closest(".list-row")
        ?.querySelector(".asset-edit-input");
      activeInput?.focus();
      setStatus("Edit the amount and click Save.", "info");
      return;
    }

    if (action === "edit-cancel") {
      editingAsset = null;
      renderAssets();
      setStatus("Edit canceled.", "info");
      return;
    }

    if (action === "remove") {
      if (editingAsset?.kind === kind && editingAsset?.index === index) {
        editingAsset = null;
      }
      assetsState[kind] = list.filter((_, idx) => idx !== index);
    } else if (action === "edit-save") {
      const row = btn.closest(".list-row");
      const input = row?.querySelector(".asset-edit-input");
      const nextValue = safeNumber(input?.value);
      if (nextValue <= 0) {
        setStatus("Amount must be greater than 0.", "error");
        input?.focus();
        return;
      }
      list[index].value = nextValue;
      editingAsset = null;
    } else {
      return;
    }

    setStatus("Saving assets...", "info");
    try {
      await saveAssets();
      renderAssets();
      setStatus(action === "remove" ? "Asset removed." : "Asset updated.", "ok");
    } catch (err) {
      setStatus(String(err?.message || err), "error");
    }
  });
}

async function runMonth() {
  const { year, month } = getSelectedYearMonth();
  const include_leumi = $("includeLeumi").checked;
  const runToken = generateRunToken();

  removeReviewPanel();
  setStatus(`Running for ${monthName(month)} ${year} (this can take a while)...`, "info");
  $("runBtn").disabled = true;
  startProgressPolling(runToken);

  try {
    const resp = await fetch("/run-month/prepare", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ year, month, include_leumi, run_token: runToken }),
    });

    let data = null;
    try {
      data = await resp.json();
    } catch (_jsonErr) {
      // Backend returned non-JSON (for example a plain-text 500 body).
      data = null;
    }

    if (!resp.ok) {
      throw new Error(
        data?.detail ||
          "Failed to run this month right now. Please try again in a moment."
      );
    }
    stopProgressPolling();
    renderReviewPanel(data, year, month);
    setStatus("Run finished. Review categories and click Apply and Fill.", "ok");
  } catch (err) {
    stopProgressPolling();
    setStatus(String(err?.message || err), "error");
    $("runBtn").disabled = false;
  }
}

function populateYearMonthSelectors(nextSuggestion) {
  const yearSel = $("year");
  const monthSel = $("month");

  // Years: last 3 + nextSuggestion year if it is out of range.
  const now = new Date();
  const years = new Set([String(now.getFullYear()), String(now.getFullYear() - 1), String(now.getFullYear() - 2)]);
  if (nextSuggestion?.year) years.add(nextSuggestion.year);
  const sortedYears = Array.from(years).sort((a, b) => Number(a) - Number(b));

  yearSel.innerHTML = "";
  for (const y of sortedYears) {
    const opt = document.createElement("option");
    opt.value = y;
    opt.textContent = y;
    yearSel.appendChild(opt);
  }

  // Months: 1..12
  monthSel.innerHTML = "";
  for (let m = 1; m <= 12; m++) {
    const opt = document.createElement("option");
    opt.value = String(m);
    opt.textContent = String(m);
    monthSel.appendChild(opt);
  }

  if (nextSuggestion?.year) yearSel.value = nextSuggestion.year;
  if (nextSuggestion?.month) monthSel.value = String(nextSuggestion.month);
}

async function init() {
  wireRemoveButtons();

  // Populate selectors from state "next month".
  const stateResp = await fetch("/state");
  const state = await stateResp.json();
  const suggestion = computeNextMonthFromState(state);
  populateYearMonthSelectors(suggestion);

  $("includeLeumi").checked = true;

  await loadStateAndAssets();
  await loadSummaryChart();

  $("runBtn").addEventListener("click", runMonth);
  $("addCarBtn").addEventListener("click", addCar);
  $("addInvestmentBtn").addEventListener("click", addInvestment);
}

init().catch((err) => {
  setStatus(String(err?.message || err), "error");
});

