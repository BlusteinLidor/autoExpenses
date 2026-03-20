const $ = (id) => document.getElementById(id);

let summaryChart = null;

let assetsState = {
  cars: [],
  investments: [],
};

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
  $("carsTotal").textContent = toMoneyString(carsTotal);
  $("investmentsTotal").textContent = toMoneyString(investmentsTotal);

  for (const car of assetsState.cars) {
    const row = document.createElement("div");
    row.className = "list-row";
    row.innerHTML = `
      <div class="list-main">
        <div class="list-title">${escapeHtml(car.name ?? "")}</div>
        <div class="list-sub">${toMoneyString(car.value)}</div>
      </div>
      <button class="danger" data-kind="cars" data-name="${escapeAttr(
        car.name ?? ""
      )}">Remove</button>
    `;
    carsList.appendChild(row);
  }

  for (const inv of assetsState.investments) {
    const row = document.createElement("div");
    row.className = "list-row";
    row.innerHTML = `
      <div class="list-main">
        <div class="list-title">${escapeHtml(inv.name ?? "")}</div>
        <div class="list-sub">${toMoneyString(inv.value)}</div>
      </div>
      <button
        class="danger"
        data-kind="investments"
        data-name="${escapeAttr(inv.name ?? "")}"
      >
        Remove
      </button>
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
  // Event delegation for dynamically rendered lists.
  document.addEventListener("click", async (e) => {
    const btn = e.target.closest("button[data-kind][data-name]");
    if (!btn) return;

    const kind = btn.getAttribute("data-kind");
    const name = btn.getAttribute("data-name");
    if (!kind || !name) return;

    if (kind !== "cars" && kind !== "investments") return;
    const list = assetsState[kind];
    const nextList = list.filter((item) => String(item.name ?? "") !== name);
    assetsState[kind] = nextList;

    setStatus("Saving assets...", "info");
    try {
      await saveAssets();
      renderAssets();
      setStatus("Assets saved.", "ok");
    } catch (err) {
      setStatus(String(err?.message || err), "error");
    }
  });
}

async function runMonth() {
  const { year, month } = getSelectedYearMonth();
  const include_leumi = $("includeLeumi").checked;

  setStatus(`Running for ${monthName(month)} ${year} (this can take a while)...`, "info");
  $("runBtn").disabled = true;

  try {
    const resp = await fetch("/run-month", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ year, month, include_leumi }),
    });

    const data = await resp.json();
    if (!resp.ok) {
      throw new Error(data?.detail || `Run failed with status ${resp.status}`);
    }

    setStatus(`Done. Output: ${data.output_excel}`, "ok");
    // Refresh state/assets and re-draw chart.
    await loadStateAndAssets();
    await loadSummaryChart();
  } catch (err) {
    setStatus(String(err?.message || err), "error");
  } finally {
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

