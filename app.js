const DEFAULT_BACKEND_URL = "https://backend.vermieter.techem.lol";
const DEFAULT_IMPORT_DIRECTORY = "backend/data/sample_csvs";
const CHAT_PROVIDER = "gemini";
const IMPORT_DIRECTORY_STORAGE_KEY = "techem-import-directory";
const THEME_STORAGE_KEY = "techem-theme";
const IMPORT_PANEL_STORAGE_KEY = "techem-import-panel-open";

const state = {
  backendUrl: DEFAULT_BACKEND_URL,
  currentScope: { type: "total", id: null, label: "Gesamtbestand" },
  period: "month",
  offset: 0,
  analysisProvider: "vertex",
  includeWeather: true,
  overview: null,
  navigationCache: {
    city: new Map(),
    building: new Map(),
    apartment: new Map(),
  },
  chartCache: new Map(),
  reportCache: new Map(),
  chart: null,
  report: null,
  reportLoading: false,
  chartLoading: false,
  chatLoading: false,
  theme: "light",
  tips: [],
  challenges: [],
  suggestions: [],
  chartController: null,
  reportController: null,
  chartRequestId: 0,
  reportRequestId: 0,
  chatRequestId: 0,
  chartVisibility: {
    actual_energy: true,
    predicted_energy: true,
    optimized_energy: true,
    actual_co2: true,
    predicted_co2: true,
    optimized_co2: false,
  },
  currentScopeSummary: null,
  chatMessages: [
    {
      role: "assistant",
      text: "Ich kann den aktuellen Scope, die geöffnete Zeitreihe und den Gesamtbestand für Rückfragen zusammenfassen.",
      provider: "system",
      fallbackUsed: false,
    },
  ],
};

const SERIES_CONFIG = {
  actual_energy: { label: "Ist Energie", unit: "kWh", color: "#e30613", axis: "energy", className: "actual" },
  predicted_energy: {
    label: "Erwartet Energie",
    unit: "kWh",
    color: "#f36a76",
    axis: "energy",
    className: "predicted",
  },
  optimized_energy: {
    label: "Optimiert Energie",
    unit: "kWh",
    color: "#0f8a6b",
    axis: "energy",
    className: "optimized",
  },
  actual_co2: { label: "Ist CO2", unit: "kg", color: "#6f5fe1", axis: "co2", className: "actual-co2" },
  predicted_co2: {
    label: "Erwartet CO2",
    unit: "kg",
    color: "#9b8cf0",
    axis: "co2",
    className: "predicted-co2",
  },
  optimized_co2: {
    label: "Optimiert CO2",
    unit: "kg",
    color: "#ff9665",
    axis: "co2",
    className: "optimized-co2",
  },
};

const dom = {};

document.addEventListener("DOMContentLoaded", async () => {
  cacheDom();
  bindEvents();
  renderStaticControls();
  await bootstrap();
});

function cacheDom() {
  dom.importDirectoryInput = document.querySelector("#import-directory-input");
  dom.backendUrlInput = document.querySelector("#backend-url-input");
  dom.importSampleButton = document.querySelector("#import-sample-button");
  dom.importToggleButton = document.querySelector("#import-toggle-button");
  dom.importPanel = document.querySelector("#import-panel");
  dom.themeToggleButton = document.querySelector("#theme-toggle-button");
  dom.scopeTotalButton = document.querySelector("#scope-total-button");
  dom.navigationTree = document.querySelector("#navigation-tree");
  dom.selectionBadge = document.querySelector("#selection-badge");
  dom.heroStatus = document.querySelector("#hero-status");
  dom.metricEnergy = document.querySelector("#metric-energy");
  dom.metricEnergySub = document.querySelector("#metric-energy-sub");
  dom.metricCo2 = document.querySelector("#metric-co2");
  dom.metricCo2Sub = document.querySelector("#metric-co2-sub");
  dom.metricCost = document.querySelector("#metric-cost");
  dom.metricCostSub = document.querySelector("#metric-cost-sub");
  dom.metricAssets = document.querySelector("#metric-assets");
  dom.metricAssetsSub = document.querySelector("#metric-assets-sub");
  dom.chartTitle = document.querySelector("#chart-title");
  dom.chartSubtitle = document.querySelector("#chart-subtitle");
  dom.chartCanvas = document.querySelector("#chart-canvas");
  dom.chartWarningBar = document.querySelector("#chart-warning-bar");
  dom.seriesToggles = document.querySelector("#series-toggles");
  dom.anomaliesList = document.querySelector("#anomalies-list");
  dom.factorsList = document.querySelector("#factors-list");
  dom.dataSourceList = document.querySelector("#data-source-list");
  dom.reportContent = document.querySelector("#report-content");
  dom.reportStatus = document.querySelector("#report-status");
  dom.refreshReportButton = document.querySelector("#refresh-report-button");
  dom.chatMessages = document.querySelector("#chat-messages");
  dom.chatContext = document.querySelector("#chat-context");
  dom.chatForm = document.querySelector("#chat-form");
  dom.chatInput = document.querySelector("#chat-input");
  dom.chatSubmitButton = dom.chatForm.querySelector("button[type='submit']");
  dom.chatScopeToggle = document.querySelector("#chat-scope-toggle");
  dom.detailsContent = document.querySelector("#details-content");
  dom.detailsCaption = document.querySelector("#details-caption");
  dom.searchInput = document.querySelector("#search-input");
  dom.suggestionsList = document.querySelector("#suggestions-list");
  dom.tipsList = document.querySelector("#tips-list");
  dom.tipsStatus = document.querySelector("#tips-status");
  dom.challengesList = document.querySelector("#challenges-list");
  dom.weatherToggle = document.querySelector("#weather-toggle");
  dom.periodSwitcher = document.querySelector("#period-switcher");
  dom.providerSwitcher = document.querySelector("#provider-switcher");
  dom.offsetPrevButton = document.querySelector("#offset-prev-button");
  dom.offsetNextButton = document.querySelector("#offset-next-button");
  dom.offsetResetButton = document.querySelector("#offset-reset-button");
}

function bindEvents() {
  dom.importDirectoryInput.addEventListener("change", () => {
    localStorage.setItem(IMPORT_DIRECTORY_STORAGE_KEY, dom.importDirectoryInput.value.trim());
  });

  dom.backendUrlInput.addEventListener("change", async () => {
    state.backendUrl = normalizeBackendUrl(dom.backendUrlInput.value);
    dom.backendUrlInput.value = state.backendUrl;
    state.chartCache.clear();
    state.reportCache.clear();
    await bootstrap();
  });

  dom.importSampleButton.addEventListener("click", importCsvDirectory);
  dom.importToggleButton.addEventListener("click", toggleImportPanel);
  dom.themeToggleButton.addEventListener("click", toggleTheme);

  dom.scopeTotalButton.addEventListener("click", async () => {
    setScope({ type: "total", id: null, label: "Gesamtbestand" });
    await refreshMainView();
  });

  dom.weatherToggle.addEventListener("change", async (event) => {
    state.includeWeather = event.target.checked;
    await loadChart();
  });

  dom.periodSwitcher.addEventListener("click", async (event) => {
    const button = event.target.closest("button[data-period]");
    if (!button) {
      return;
    }
    state.period = button.dataset.period;
    renderPeriodButtons();
    await refreshMainView();
  });

  dom.providerSwitcher.addEventListener("click", async (event) => {
    const button = event.target.closest("button[data-provider]");
    if (!button) {
      return;
    }
    state.analysisProvider = button.dataset.provider;
    renderProviderButtons();
    await loadChart();
  });

  dom.offsetPrevButton.addEventListener("click", async () => {
    state.offset -= 1;
    renderOffsetControls();
    await loadChart();
  });

  dom.offsetNextButton.addEventListener("click", async () => {
    state.offset += 1;
    renderOffsetControls();
    await loadChart();
  });

  dom.offsetResetButton.addEventListener("click", async () => {
    alignOffsetToLatestData({ summary: state.currentScopeSummary });
    renderOffsetControls();
    await loadChart();
  });

  dom.refreshReportButton.addEventListener("click", async () => {
    await loadReport(true);
  });

  dom.chatForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const message = dom.chatInput.value.trim();
    if (!message || state.chatLoading) {
      return;
    }
    dom.chatInput.value = "";
    state.chatMessages.push({ role: "user", text: message });
    renderChatMessages();
    await sendChat(message);
  });

  dom.searchInput.addEventListener("input", async () => {
    const query = dom.searchInput.value.trim();
    if (!query) {
      state.suggestions = [];
      renderSuggestions();
      return;
    }
    await loadSuggestions(query);
  });

  document.addEventListener("click", (event) => {
    if (!event.target.closest(".search-panel")) {
      state.suggestions = [];
      renderSuggestions();
    }
  });
}

function renderStaticControls() {
  state.theme = loadTheme();
  dom.importDirectoryInput.value = loadImportDirectory();
  dom.backendUrlInput.value = state.backendUrl;
  applyTheme(state.theme);
  applyImportPanelState(loadImportPanelState());
  renderPeriodButtons();
  renderOffsetControls();
  renderProviderButtons();
  renderSeriesToggles();
  renderWarnings(dom.chartWarningBar, []);
  renderChatMessages();
  updateChatComposer();
}

async function bootstrap() {
  try {
    dom.tipsStatus.textContent = "Backend verbunden";
    await Promise.all([loadOverview(), loadTips(), loadChallenges()]);
    await refreshMainView();
  } catch (error) {
    renderConnectionError(error);
  }
}

function normalizeBackendUrl(url) {
  const trimmed = (url || DEFAULT_BACKEND_URL).trim().replace(/\/+$/, "");
  return trimmed || DEFAULT_BACKEND_URL;
}

function loadTheme() {
  return localStorage.getItem(THEME_STORAGE_KEY) || "light";
}

function loadImportPanelState() {
  return localStorage.getItem(IMPORT_PANEL_STORAGE_KEY) === "true";
}

function toggleTheme() {
  state.theme = state.theme === "dark" ? "light" : "dark";
  localStorage.setItem(THEME_STORAGE_KEY, state.theme);
  applyTheme(state.theme);
}

function toggleImportPanel() {
  const nextState = !dom.importPanel.classList.contains("is-open");
  applyImportPanelState(nextState);
  localStorage.setItem(IMPORT_PANEL_STORAGE_KEY, String(nextState));
}

function applyTheme(theme) {
  document.documentElement.dataset.theme = theme;
  document.body.dataset.theme = theme;
  dom.themeToggleButton.textContent = theme === "dark" ? "Light Mode" : "Dark Mode";
}

function applyImportPanelState(isOpen) {
  dom.importPanel.classList.toggle("is-open", isOpen);
  dom.importToggleButton.textContent = isOpen ? "Datenimport schließen" : "Datenimport";
}

function buildScopeCachePart() {
  return `${state.currentScope.type}:${state.currentScope.id || "all"}`;
}

function buildChartCacheKey() {
  return [
    state.backendUrl,
    buildScopeCachePart(),
    state.period,
    state.offset,
    state.includeWeather ? "weather" : "plain",
    state.analysisProvider,
  ].join("|");
}

function buildReportCacheKey() {
  return [state.backendUrl, buildScopeCachePart(), state.period, state.offset, state.analysisProvider].join("|");
}

async function apiFetch(path, options = {}) {
  const response = await fetch(`${state.backendUrl}${path}`, {
    headers: {
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
    ...options,
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`${response.status} ${response.statusText}: ${text}`);
  }

  if (response.status === 204) {
    return null;
  }

  return response.json();
}

async function loadOverview() {
  state.overview = await apiFetch("/api/navigation/overview");
  renderOverviewNavigation();
}

async function loadTips() {
  const response = await apiFetch("/api/tips");
  state.tips = response.tips || [];
  renderTips();
}

async function loadChallenges() {
  const response = await apiFetch("/api/challenges");
  state.challenges = response.challenges || [];
  renderChallenges();
}

async function refreshMainView() {
  const payload = await loadScopeDetails();
  alignOffsetToLatestData(payload);
  renderOffsetControls();
  await loadChart();
}

async function loadScopeDetails() {
  const { type, id } = state.currentScope;
  let payload = state.overview;

  if (type === "city" && id) {
    if (!state.navigationCache.city.has(id)) {
      state.navigationCache.city.set(id, await apiFetch(`/api/navigation/cities/${id}`));
    }
    payload = state.navigationCache.city.get(id);
  } else if (type === "building" && id) {
    if (!state.navigationCache.building.has(id)) {
      state.navigationCache.building.set(id, await apiFetch(`/api/navigation/buildings/${id}`));
    }
    payload = state.navigationCache.building.get(id);
  } else if (type === "apartment" && id) {
    if (!state.navigationCache.apartment.has(id)) {
      state.navigationCache.apartment.set(id, await apiFetch(`/api/navigation/apartments/${id}`));
    }
    payload = state.navigationCache.apartment.get(id);
  }

  renderScopeSummary(payload);
  renderNavigationTree();
  renderScopeDetails(payload);
  state.currentScopeSummary = payload?.summary || null;
  return payload;
}

async function loadChart() {
  const chartCacheKey = buildChartCacheKey();
  const cachedChart = state.chartCache.get(chartCacheKey);
  if (cachedChart) {
    if (state.chartController) {
      state.chartController.abort();
      state.chartController = null;
    }
    state.chart = cachedChart;
    state.report = state.reportCache.get(buildReportCacheKey()) || null;
    state.chartLoading = false;
    dom.chartCanvas.classList.remove("is-loading");
    renderChartSection();
    renderChatContext();
    renderReport();
    dom.refreshReportButton.disabled = false;
    return;
  }

  if (state.chartController) {
    state.chartController.abort();
  }

  const controller = new AbortController();
  const requestId = ++state.chartRequestId;
  state.chartController = controller;
  state.chartLoading = true;
  dom.chartCanvas.classList.add("is-loading");
  dom.chartSubtitle.textContent = "Chart wird aktualisiert";
  dom.refreshReportButton.disabled = true;

  const scopeId = state.currentScope.type === "total" ? "" : `&scope_id=${encodeURIComponent(state.currentScope.id)}`;

  try {
    const chart = await apiFetch(
      `/api/chart?scope_type=${state.currentScope.type}${scopeId}&period=${state.period}&offset=${state.offset}&include_weather=${state.includeWeather}&analysis_provider=${state.analysisProvider}`,
      { signal: controller.signal }
    );

    if (requestId !== state.chartRequestId) {
      return;
    }

    state.chart = chart;
    state.report = null;
    state.chartCache.set(chartCacheKey, chart);
    renderChartSection();
    renderChatContext();
    renderReport();
  } catch (error) {
    if (error.name === "AbortError") {
      return;
    }
    dom.chartCanvas.innerHTML = `<div class="empty-placeholder">Chart-Fehler: ${escapeHtml(error.message)}</div>`;
    dom.chartSubtitle.textContent = "Chart konnte nicht geladen werden";
    renderWarnings(dom.chartWarningBar, []);
  } finally {
    if (requestId === state.chartRequestId) {
      state.chartLoading = false;
      state.chartController = null;
      dom.chartCanvas.classList.remove("is-loading");
      dom.refreshReportButton.disabled = !state.chart;
    }
  }
}

async function loadReport(force = false) {
  const reportCacheKey = buildReportCacheKey();
  if (!force && state.reportCache.has(reportCacheKey)) {
    state.report = state.reportCache.get(reportCacheKey);
    renderReport();
    return;
  }

  if (!force && state.report) {
    return;
  }

  if (state.reportController) {
    state.reportController.abort();
  }

  const controller = new AbortController();
  const requestId = ++state.reportRequestId;
  state.reportController = controller;
  state.reportLoading = true;
  renderReport();

  const scopeId = state.currentScope.type === "total" ? "" : `&scope_id=${encodeURIComponent(state.currentScope.id)}`;

  try {
    const report = await apiFetch(
      `/api/report?scope_type=${state.currentScope.type}${scopeId}&period=${state.period}&offset=${state.offset}&analysis_provider=${state.analysisProvider}`,
      { signal: controller.signal }
    );

    if (requestId !== state.reportRequestId) {
      return;
    }

    state.report = report;
    state.reportCache.set(reportCacheKey, report);
  } catch (error) {
    if (error.name === "AbortError") {
      return;
    }
    state.report = {
      provider: "error",
      title: "Report konnte nicht erzeugt werden",
      overview: error.message,
      main_findings: [],
      influencing_factors: [],
      forecast_notes: [],
      optimization_notes: [],
      risks_and_uncertainties: [],
      plain_text_report: "Der Report-Request ist fehlgeschlagen.",
      used_context: {},
    };
  } finally {
    if (requestId === state.reportRequestId) {
      state.reportLoading = false;
      state.reportController = null;
      renderReport();
    }
  }
}

async function loadSuggestions(query) {
  const response = await apiFetch(`/api/search/suggestions?q=${encodeURIComponent(query)}`);
  state.suggestions = response.suggestions || [];
  renderSuggestions();
}

async function sendChat(message) {
  const pendingId = `pending-${++state.chatRequestId}`;
  const payload = {
    message,
    use_current_scope: dom.chatScopeToggle.checked,
    scope_type: state.currentScope.type,
    scope_id: state.currentScope.type === "total" ? null : state.currentScope.id,
    period: state.period,
    offset: state.offset,
    analysis_provider: CHAT_PROVIDER,
  };

  state.chatLoading = true;
  updateChatComposer();
  state.chatMessages.push({
    id: pendingId,
    role: "assistant",
    text: "",
    provider: CHAT_PROVIDER,
    fallbackUsed: false,
    pending: true,
  });
  renderChatMessages();

  try {
    const response = await apiFetch("/api/chat", {
      method: "POST",
      body: JSON.stringify(payload),
    });

    replaceChatMessage(pendingId, {
      id: pendingId,
      role: "assistant",
      text: response.answer,
      provider: response.provider || "unknown",
      fallbackUsed: Boolean(response.fallback_used),
      caveats: response.caveats || [],
      context: response.used_context,
    });
  } catch (error) {
    replaceChatMessage(pendingId, {
      id: pendingId,
      role: "assistant",
      text: `Chat-Fehler: ${error.message}`,
      provider: "error",
      fallbackUsed: false,
    });
  } finally {
    state.chatLoading = false;
    updateChatComposer();
    renderChatMessages();
  }
}

function loadImportDirectory() {
  return localStorage.getItem(IMPORT_DIRECTORY_STORAGE_KEY) || DEFAULT_IMPORT_DIRECTORY;
}

async function importCsvDirectory() {
  const directoryPath = dom.importDirectoryInput.value.trim();
  if (!directoryPath) {
    showHeroMessage(
      "CSV-Ordner fehlt",
      "Trage einen absoluten oder relativen Ordnerpfad ein, bevor du den Import startest.",
      true
    );
    dom.importDirectoryInput.focus();
    return;
  }

  dom.importSampleButton.disabled = true;
  dom.importSampleButton.textContent = "Import läuft...";

  try {
    const response = await apiFetch("/api/import/csv-directory", {
      method: "POST",
      body: JSON.stringify({ directory_path: directoryPath }),
    });

    localStorage.setItem(IMPORT_DIRECTORY_STORAGE_KEY, directoryPath);
    state.navigationCache.city.clear();
    state.navigationCache.building.clear();
    state.navigationCache.apartment.clear();
    state.chartCache.clear();
    state.reportCache.clear();

    await loadOverview();
    await refreshMainView();

    showHeroMessage(
      "CSV-Daten importiert",
      `${response.imported_files ?? 0} CSV-Dateien aus ${directoryPath} wurden eingelesen und das Dashboard wurde aktualisiert.`
    );
  } catch (error) {
    showHeroMessage("Import fehlgeschlagen", error.message, true);
  } finally {
    dom.importSampleButton.disabled = false;
    dom.importSampleButton.textContent = "CSV-Ordner importieren";
  }
}

function setScope(scope) {
  state.currentScope = { ...scope };
  state.offset = 0;
  dom.selectionBadge.textContent = scope.label;
}

function renderOverviewNavigation() {
  const summary = state.overview?.summary;
  if (!summary) {
    return;
  }

  if (summary.record_count === 0) {
    showHeroMessage(
      "Noch keine CSV-Daten aktiv",
      "Importiere einen CSV-Ordner über die Leiste oben oder lade später einen anderen lokalen Ordner.",
      false,
      true
    );
    return;
  }

  showHeroMessage(
    state.currentScope.label,
    `${summary.record_count} Datensätze im aktiven Bestand, von ${summary.date_from} bis ${summary.date_to}.`
  );
}

function renderNavigationTree() {
  const overview = state.overview;
  if (!overview) {
    dom.navigationTree.innerHTML = "";
    return;
  }

  const sections = [
    `
      <section class="portfolio-summary">
        <div class="portfolio-summary-title">Portfolio</div>
        <div class="portfolio-summary-grid">
          <div><strong>${overview.summary.city_count}</strong><span>Städte</span></div>
          <div><strong>${overview.summary.building_count}</strong><span>Gebäude</span></div>
          <div><strong>${overview.summary.apartment_count}</strong><span>Wohnungen</span></div>
        </div>
      </section>
    `,
  ];

  for (const city of overview.cities) {
    const cityPayload = state.navigationCache.city.get(city.city_id);
    const isCityActive = state.currentScope.type === "city" && state.currentScope.id === city.city_id;
    const cityChildren = [];

    if (cityPayload?.buildings) {
      for (const building of cityPayload.buildings) {
        const buildingPayload = state.navigationCache.building.get(building.building_id);
        const isBuildingActive = state.currentScope.type === "building" && state.currentScope.id === building.building_id;

        cityChildren.push(`
          <button class="nav-item nav-item-building ${isBuildingActive ? "is-active" : ""}" data-scope-type="building" data-scope-id="${building.building_id}" data-scope-label="${escapeHtml(`${building.address}, ${building.city}`)}">
            <span class="nav-level-tag">Gebäude</span>
            <span class="nav-title">${escapeHtml(building.address)}</span>
            <span class="nav-meta">${formatNumber(building.total_energy_kwh, 1)} kWh | ${formatNumber(building.total_co2_kg, 1)} kg CO2 | ${building.apartment_count} Wohnungen</span>
          </button>
        `);

        if (buildingPayload?.apartments) {
          for (const apartment of buildingPayload.apartments) {
            const isApartmentActive =
              state.currentScope.type === "apartment" && state.currentScope.id === apartment.apartment_id;
            cityChildren.push(`
              <button class="nav-item nav-item-apartment ${isApartmentActive ? "is-active" : ""}" data-scope-type="apartment" data-scope-id="${apartment.apartment_id}" data-scope-label="${escapeHtml(`Wohnung ${apartment.apartment_number}`)}">
                <span class="nav-level-tag">Wohnung</span>
                <span class="nav-title">Wohnung ${escapeHtml(apartment.apartment_number)}</span>
                <span class="nav-meta">${formatNumber(apartment.total_energy_kwh, 1)} kWh | ${formatNumber(apartment.livingspace_m2, 1)} m² | ${apartment.roomnumber} Zimmer</span>
              </button>
            `);
          }
        }
      }
    }

    sections.push(`
      <section class="nav-group ${isCityActive ? "is-active-group" : ""}">
        <button class="nav-item nav-item-city ${isCityActive ? "is-active" : ""}" data-scope-type="city" data-scope-id="${city.city_id}" data-scope-label="${escapeHtml(city.city)}">
          <span class="nav-level-tag">Stadt</span>
          <span class="nav-title">${escapeHtml(city.city)}</span>
          <span class="nav-meta">${formatNumber(city.total_energy_kwh, 1)} kWh | ${formatNumber(city.total_co2_kg, 1)} kg CO2 | ${city.building_count} Gebäude</span>
          <span class="nav-chip-row">
            <span class="nav-chip">${city.apartment_count} Wohnungen</span>
            <span class="nav-chip">${city.building_count} Immobilien</span>
          </span>
        </button>
        ${
          cityChildren.length
            ? `<div class="nav-children"><div class="nav-section-label">Gebäudemanagement</div>${cityChildren.join("")}</div>`
            : `<div class="nav-empty-state">Stadt öffnen, um Gebäude und Wohnungen in der Verwaltung zu sehen.</div>`
        }
      </section>
    `);
  }

  dom.navigationTree.innerHTML = sections.join("") || `<div class="empty-placeholder">Noch keine Navigation verfügbar.</div>`;
  dom.navigationTree.querySelectorAll("[data-scope-type]").forEach((button) => {
    button.addEventListener("click", async () => {
      setScope({
        type: button.dataset.scopeType,
        id: button.dataset.scopeId,
        label: button.dataset.scopeLabel,
      });
      await refreshMainView();
    });
  });
}

function renderScopeSummary(payload) {
  const summary = payload?.summary || {
    total_energy_kwh: 0,
    total_co2_kg: 0,
    estimated_cost_eur: 0,
    record_count: 0,
    city_count: 0,
    building_count: 0,
    apartment_count: 0,
    date_from: null,
    date_to: null,
  };
  const scope = payload?.scope || state.currentScope;

  dom.metricEnergy.textContent = `${formatNumber(summary.total_energy_kwh, 1)} kWh`;
  dom.metricEnergySub.textContent = summary.date_from
    ? `${summary.date_from} bis ${summary.date_to}`
    : "Noch keine aktiven Messwerte";

  dom.metricCo2.textContent = `${formatNumber(summary.total_co2_kg, 1)} kg`;
  dom.metricCo2Sub.textContent = "Deterministisch berechnet";

  dom.metricCost.textContent = `${formatCurrency(summary.estimated_cost_eur)}`;
  dom.metricCostSub.textContent = "Schätzung auf Basis fixer Demo-Kosten";

  const assetLabel =
    scope.type === "total"
      ? `${summary.city_count} Städte`
      : scope.type === "city"
        ? `${summary.building_count} Gebäude`
        : scope.type === "building"
          ? `${summary.apartment_count} Wohnungen`
          : `${summary.record_count} Messpunkte`;
  dom.metricAssets.textContent = assetLabel;
  dom.metricAssetsSub.textContent = `${summary.record_count} Datensätze im Scope`;

  dom.selectionBadge.textContent = scope.label || state.currentScope.label;
  showHeroMessage(
    scope.label || state.currentScope.label,
    summary.date_from
      ? `${summary.record_count} Datensätze von ${summary.date_from} bis ${summary.date_to}.`
      : "Noch keine aktiven Messwerte im aktuellen Scope."
  );
}

function renderScopeDetails(payload) {
  let cells = [];
  let caption = "Metadaten";

  if (payload?.metadata) {
    caption = "Wohnungsdetails";
    cells = [
      detailCell("Adresse", `${payload.metadata.street} ${payload.metadata.housenumber}, ${payload.metadata.city}`),
      detailCell("Wohnung", payload.metadata.apartment_number),
      detailCell("Postleitzahl", payload.metadata.zipcode),
      detailCell("Energiequelle", payload.metadata.energysource),
      detailCell("Wohnfläche", `${formatNumber(payload.metadata.livingspace_m2, 1)} m²`),
      detailCell("Zimmer", `${payload.metadata.roomnumber}`),
      detailCell("Schimmelrisiko", formatMoldRisk(payload.metadata.mold_risk)),
      detailCell("Unit", payload.metadata.unitnumber),
      detailCell("Gebäude-ID", payload.metadata.building_id),
    ];
  } else if (payload?.buildings) {
    caption = "Stadtfokus";
    cells = payload.buildings.map((building) =>
      detailCell(
        building.address,
        `${formatNumber(building.total_co2_kg, 1)} kg CO2 | ${building.apartment_count} Wohnungen | Schimmelrisiko ${formatMoldRisk(building.mold_risk)}`
      )
    );
  } else if (payload?.apartments) {
    caption = "Gebäudefokus";
    cells = payload.apartments.map((apartment) =>
      detailCell(
        `Wohnung ${apartment.apartment_number}`,
        `${formatNumber(apartment.total_energy_kwh, 1)} kWh | ${formatNumber(apartment.livingspace_m2, 1)} m²`
      )
    );
  } else if (payload?.scope?.type === "total") {
    caption = "Bestandsfokus";
    cells = [
      detailCell("Städte", `${payload.summary.city_count}`),
      detailCell("Gebäude", `${payload.summary.building_count}`),
      detailCell("Wohnungen", `${payload.summary.apartment_count}`),
      detailCell("Zeitraum", `${payload.summary.date_from || "-"} bis ${payload.summary.date_to || "-"}`),
    ];
  }

  dom.detailsCaption.textContent = caption;
  dom.detailsContent.innerHTML = cells.length
    ? `<div class="details-grid">${cells.join("")}</div>`
    : `<div class="empty-placeholder">Keine Detaildaten für den aktuellen Scope.</div>`;
}

function renderChartSection() {
  if (!state.chart) {
    return;
  }

  dom.chartTitle.textContent = `Verlauf: ${state.chart.scope.label}`;
  dom.chartSubtitle.textContent = `${currentPeriodLabel()} · ${state.chart.period.start} bis ${state.chart.period.end} · Provider ${state.chart.ai_explanations.provider}`;
  renderWarnings(dom.chartWarningBar, state.chart.warnings || []);
  renderSeriesToggles();
  renderChartSvg();
  renderList(dom.anomaliesList, state.chart.anomalies, "Keine Ausreißer im aktuellen Scope markiert.");
  renderList(dom.factorsList, state.chart.ai_explanations.influencing_factors, "Noch keine Einflussfaktoren verfügbar.");

  const dataSourceEntries = Object.entries(state.chart.data_source_info || {}).map(
    ([key, value]) => `${humanizeKey(key)}: ${value}`
  );
  renderList(dom.dataSourceList, dataSourceEntries, "Keine Datenquellenbeschreibung verfügbar.");
}

function renderSeriesToggles() {
  dom.seriesToggles.innerHTML = Object.entries(SERIES_CONFIG)
    .map(([key, config]) => {
      const checked = state.chartVisibility[key];
      return `
        <label class="series-toggle ${checked ? "" : "is-off"}">
          <input type="checkbox" data-series-key="${key}" ${checked ? "checked" : ""} />
          <span class="series-swatch" style="background:${config.color}"></span>
          <span class="series-toggle-label">
            <strong>${config.label}</strong>
            <span>${config.unit}</span>
          </span>
        </label>
      `;
    })
    .join("");

  dom.seriesToggles.querySelectorAll("input[data-series-key]").forEach((checkbox) => {
    checkbox.addEventListener("change", () => {
      state.chartVisibility[checkbox.dataset.seriesKey] = checkbox.checked;
      renderSeriesToggles();
      renderChartSvg();
    });
  });
}

function renderChartSvg() {
  const chart = state.chart;
  if (!chart || !chart.x_axis?.length) {
    dom.chartCanvas.innerHTML = `<div class="empty-placeholder">Noch keine Chartdaten.</div>`;
    return;
  }

  const width = 920;
  const height = 340;
  const padding = { top: 18, right: 56, bottom: 46, left: 56 };
  const xAxis = chart.x_axis;
  const energySeries = ["actual_energy", "predicted_energy", "optimized_energy"].filter((key) => state.chartVisibility[key]);
  const co2Series = ["actual_co2", "predicted_co2", "optimized_co2"].filter((key) => state.chartVisibility[key]);
  const energyValues = energySeries.flatMap((key) => chart.series[key] || []);
  const co2Values = co2Series.flatMap((key) => chart.series[key] || []);
  const energyMax = Math.max(...energyValues, 0);
  const co2Max = Math.max(...co2Values, 0);
  const plotWidth = width - padding.left - padding.right;
  const plotHeight = height - padding.top - padding.bottom;

  const xForIndex = (index) =>
    padding.left + (xAxis.length <= 1 ? plotWidth / 2 : (index / (xAxis.length - 1)) * plotWidth);
  const yForEnergy = (value) =>
    padding.top + plotHeight - ((value || 0) / (energyMax || 1)) * plotHeight;
  const yForCo2 = (value) =>
    padding.top + plotHeight - ((value || 0) / (co2Max || 1)) * plotHeight;

  const gridLines = [];
  for (let step = 0; step <= 4; step += 1) {
    const y = padding.top + (plotHeight / 4) * step;
    gridLines.push(`<line class="grid-line" x1="${padding.left}" y1="${y}" x2="${width - padding.right}" y2="${y}" />`);
  }

  const axisLabels = xAxis
    .map((label, index) => {
      const x = xForIndex(index);
      const shortLabel = label.length > 7 ? label.slice(2) : label;
      return `<text class="axis-text" x="${x}" y="${height - 14}" text-anchor="middle">${escapeHtml(shortLabel)}</text>`;
    })
    .join("");

  const energyAxisLabels = [0, 0.5, 1].map((factor) => {
    const value = energyMax * (1 - factor);
    const y = padding.top + plotHeight * factor;
    return `<text class="axis-text" x="${padding.left - 12}" y="${y + 4}" text-anchor="end">${formatNumber(value, 1)}</text>`;
  });

  const co2AxisLabels = [0, 0.5, 1].map((factor) => {
    const value = co2Max * (1 - factor);
    const y = padding.top + plotHeight * factor;
    return `<text class="axis-text" x="${width - padding.right + 12}" y="${y + 4}" text-anchor="start">${formatNumber(value, 1)}</text>`;
  });

  const lineMarkup = Object.entries(SERIES_CONFIG)
    .filter(([key]) => state.chartVisibility[key])
    .map(([key, config]) => {
      const values = chart.series[key] || [];
      const path = values
        .map((value, index) => `${index === 0 ? "M" : "L"} ${xForIndex(index)} ${config.axis === "energy" ? yForEnergy(value) : yForCo2(value)}`)
        .join(" ");
      return `<path class="series-line ${config.className}" d="${path}" stroke="${config.color}" />`;
    })
    .join("");

  dom.chartCanvas.innerHTML = `
    <svg class="chart-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="Energie- und CO2-Verlauf">
      <defs>
        <linearGradient id="chartGlow" x1="0%" y1="0%" x2="100%" y2="0%">
          <stop offset="0%" stop-color="rgba(53,212,255,0.1)" />
          <stop offset="50%" stop-color="rgba(85,242,177,0.22)" />
          <stop offset="100%" stop-color="rgba(255,179,71,0.1)" />
        </linearGradient>
      </defs>
      <rect x="${padding.left}" y="${padding.top}" width="${plotWidth}" height="${plotHeight}" fill="url(#chartGlow)" opacity="0.28" rx="14" />
      ${gridLines.join("")}
      <line class="axis-line" x1="${padding.left}" y1="${padding.top}" x2="${padding.left}" y2="${padding.top + plotHeight}" />
      <line class="axis-line" x1="${width - padding.right}" y1="${padding.top}" x2="${width - padding.right}" y2="${padding.top + plotHeight}" />
      <line class="axis-line" x1="${padding.left}" y1="${padding.top + plotHeight}" x2="${width - padding.right}" y2="${padding.top + plotHeight}" />
      ${energyAxisLabels.join("")}
      ${co2AxisLabels.join("")}
      ${lineMarkup}
      ${axisLabels}
    </svg>
  `;
}

function renderReport() {
  if (state.reportLoading) {
    dom.reportStatus.textContent = "Report wird erzeugt";
    dom.reportContent.innerHTML = `<div class="empty-placeholder">Report wird geladen...</div>`;
    return;
  }

  if (!state.report) {
    dom.reportStatus.textContent = "Noch nicht geladen";
    dom.reportContent.classList.add("empty-state");
    dom.reportContent.innerHTML = `Wähle einen Scope oder klicke auf "Report erzeugen".`;
    return;
  }

  dom.reportContent.classList.remove("empty-state");
  dom.reportStatus.textContent = `Provider ${state.report.provider}`;
  dom.reportContent.innerHTML = `
    <div class="report-block">
      <h3>${escapeHtml(state.report.title)}</h3>
      <p>${escapeHtml(state.report.overview)}</p>
    </div>
    ${reportMoldRiskBlock(state.report.mold_risk)}
    ${reportListBlock("Kernaussagen", state.report.main_findings)}
    ${reportListBlock("Einflussfaktoren", state.report.influencing_factors)}
    ${reportListBlock("Forecast", state.report.forecast_notes)}
    ${reportListBlock("Optimierung", state.report.optimization_notes)}
    ${reportListBlock("Unsicherheiten", state.report.risks_and_uncertainties)}
    <div class="report-block">
      <h3>Volltext</h3>
      <p>${escapeHtml(state.report.plain_text_report)}</p>
    </div>
  `;
}

function renderChatMessages() {
  dom.chatMessages.innerHTML = state.chatMessages
    .map((message) => {
      if (message.pending) {
        return `
          <article class="chat-message assistant is-pending">
            <strong>Antwort | ${escapeHtml(message.provider || "assistant")}</strong>
            <div class="typing-bubble" aria-label="Antwort wird geladen">
              <span></span>
              <span></span>
              <span></span>
            </div>
          </article>
        `;
      }

      const caveats =
        message.caveats && message.caveats.length
          ? `<ul class="data-list">${message.caveats.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>`
          : "";

      return `
        <article class="chat-message ${message.role}">
          <strong>${message.role === "user" ? "Frage" : `Antwort | ${escapeHtml(message.provider || "unknown")}${message.fallbackUsed ? " | fallback" : ""}`}</strong>
          <div>${escapeHtml(message.text)}</div>
          ${caveats}
        </article>
      `;
    })
    .join("");
  dom.chatMessages.scrollTop = dom.chatMessages.scrollHeight;
}

function renderChatContext() {
  if (!state.chart) {
    dom.chatContext.textContent = "Noch kein aktueller Scope aktiv.";
    dom.chatContext.classList.add("empty-state");
    return;
  }

  dom.chatContext.classList.remove("empty-state");
  dom.chatContext.textContent = `${state.chart.scope.label} · ${currentPeriodLabel()} · ${state.chart.period.start} bis ${state.chart.period.end} · Chat ${CHAT_PROVIDER} · Chart ${state.analysisProvider}`;
}

function renderSuggestions() {
  if (!state.suggestions.length) {
    dom.suggestionsList.classList.remove("is-visible");
    dom.suggestionsList.innerHTML = "";
    return;
  }

  dom.suggestionsList.classList.add("is-visible");
  dom.suggestionsList.innerHTML = state.suggestions
    .map(
      (suggestion) => `
        <button class="suggestion-item" data-type="${suggestion.type}" data-id="${suggestion.id}" data-label="${escapeHtml(suggestion.label)}">
          <span class="suggestion-label">${escapeHtml(suggestion.label)}</span>
          <span class="suggestion-meta">${suggestion.type} | Score ${suggestion.score}</span>
        </button>
      `
    )
    .join("");

  dom.suggestionsList.querySelectorAll(".suggestion-item").forEach((button) => {
    button.addEventListener("click", async () => {
      setScope({
        type: button.dataset.type,
        id: button.dataset.id,
        label: button.dataset.label,
      });
      state.suggestions = [];
      dom.searchInput.value = button.dataset.label;
      renderSuggestions();
      await refreshMainView();
    });
  });
}

function renderTips() {
  dom.tipsList.innerHTML = state.tips
    .map(
      (tip) => `
        <article class="tip-card">
          <h3>${escapeHtml(tip.title)}</h3>
          <p>${escapeHtml(tip.description)}</p>
          <span class="tip-meta">${escapeHtml(tip.category)} | ${escapeHtml(tip.estimated_impact)}</span>
        </article>
      `
    )
    .join("");
}

function renderChallenges() {
  dom.challengesList.innerHTML = state.challenges
    .map((challenge) => {
      const progressRatio = challenge.goal ? Math.min(challenge.progress / challenge.goal, 1) : 0;
      const progressLabel = challenge.progress_label || `${challenge.progress}/${challenge.goal}`;
      const productDropdown =
        challenge.product_url && challenge.cta_label
          ? `
            <details class="challenge-product-dropdown">
              <summary>${escapeHtml(challenge.cta_label)}</summary>
              <div class="challenge-product-panel">
                ${
                  challenge.locked_reason
                    ? `<p class="challenge-product-warning">${escapeHtml(challenge.locked_reason)}</p>`
                    : ""
                }
                ${
                  challenge.product_image_url
                    ? `<img src="${escapeHtml(challenge.product_image_url)}" alt="${escapeHtml(challenge.product_name || "Techem Produkt")}" loading="lazy" referrerpolicy="no-referrer" />`
                    : ""
                }
                <strong>${escapeHtml(challenge.product_name || "Zusatzsensorik")}</strong>
                <a class="secondary-button challenge-action" href="${escapeHtml(challenge.product_url)}" target="_blank" rel="noreferrer">Techem-Seite öffnen</a>
              </div>
            </details>
          `
          : "";

      return `
        <article class="challenge-card ${challenge.enabled ? "" : "disabled"}">
          <h3>${escapeHtml(challenge.title)}</h3>
          <p>${escapeHtml(challenge.description)}</p>
          <div class="challenge-meta">${escapeHtml(challenge.enabled ? `${challenge.reward} | ${challenge.category}` : challenge.reward)}</div>
          ${productDropdown}
          <div class="progress-track"><div class="progress-bar" style="width:${progressRatio * 100}%"></div></div>
          <p class="muted-note">${escapeHtml(progressLabel)}</p>
        </article>
      `;
    })
    .join("");
}
function renderConnectionError(error) {
  dom.tipsStatus.textContent = "Backend nicht erreichbar";
  showHeroMessage("Backend nicht erreichbar", error.message, true, true);
}

function showHeroMessage(title, description, isError = false, showImportHint = false) {
  dom.heroStatus.innerHTML = `
    <div class="status-copy">
      <span class="status-kicker">Datensatz</span>
      <h2>${escapeHtml(title)}</h2>
      <p>${escapeHtml(description)}</p>
    </div>
    <div class="status-actions">
      ${showImportHint ? `<button class="ghost-button" id="hero-import-button">CSV-Ordner importieren</button>` : ""}
      <span class="selection-badge hero-badge ${isError ? "is-error" : ""}">${isError ? "Fehler" : "Live"}</span>
    </div>
  `;

  const importButton = document.querySelector("#hero-import-button");
  if (importButton) {
    importButton.addEventListener("click", importCsvDirectory);
  }
}

function renderList(target, items, emptyText) {
  if (!items || !items.length) {
    target.innerHTML = `<li>${escapeHtml(emptyText)}</li>`;
    return;
  }
  target.innerHTML = items.map((item) => `<li>${escapeHtml(item)}</li>`).join("");
}

function renderWarnings(target, warnings) {
  if (!target) {
    return;
  }

  if (!warnings || !warnings.length) {
    target.innerHTML = "";
    target.classList.remove("is-visible");
    return;
  }

  target.classList.add("is-visible");
  target.innerHTML = warnings
    .map(
      (warning) => `
        <div class="warning-pill">
          <strong>Wetter-Hinweis</strong>
          <span>${escapeHtml(warning)}</span>
        </div>
      `
    )
    .join("");
}

function renderPeriodButtons() {
  dom.periodSwitcher.querySelectorAll("button").forEach((button) => {
    button.classList.toggle("is-active", button.dataset.period === state.period);
  });
  renderOffsetControls();
}

function renderProviderButtons() {
  dom.providerSwitcher.querySelectorAll("button").forEach((button) => {
    button.classList.toggle("is-active", button.dataset.provider === state.analysisProvider);
  });
}

function parseDataDate(value) {
  if (!value) {
    return null;
  }
  const parsed = new Date(`${value}T12:00:00`);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function monthDiff(fromDate, toDate) {
  return (toDate.getFullYear() - fromDate.getFullYear()) * 12 + (toDate.getMonth() - fromDate.getMonth());
}

function startOfWeek(value) {
  const result = new Date(value);
  const day = (result.getDay() + 6) % 7;
  result.setDate(result.getDate() - day);
  result.setHours(12, 0, 0, 0);
  return result;
}

function quarterIndex(value) {
  return value.getFullYear() * 4 + Math.floor(value.getMonth() / 3);
}

function offsetForDate(period, targetDate) {
  const today = new Date();
  today.setHours(12, 0, 0, 0);

  if (period === "day") {
    return Math.round((targetDate - today) / 86400000);
  }
  if (period === "week") {
    return Math.round((startOfWeek(targetDate) - startOfWeek(today)) / (7 * 86400000));
  }
  if (period === "month") {
    return monthDiff(today, targetDate);
  }
  if (period === "quarter") {
    return quarterIndex(targetDate) - quarterIndex(today);
  }
  return targetDate.getFullYear() - today.getFullYear();
}

function alignOffsetToLatestData(payload) {
  const latestDate = parseDataDate(payload?.summary?.date_to);
  if (!latestDate) {
    return;
  }
  state.offset = offsetForDate(state.period, latestDate);
}

function renderOffsetControls() {
  const currentLabel = currentPeriodLabel();
  dom.offsetResetButton.textContent = currentLabel;
  dom.offsetResetButton.title = `${currentLabel} anzeigen`;
  dom.offsetPrevButton.title = previousPeriodLabel();
  dom.offsetNextButton.title = nextPeriodLabel();
}

function reportListBlock(title, items) {
  if (!items || !items.length) {
    return "";
  }
  return `
    <div class="report-block">
      <h3>${escapeHtml(title)}</h3>
      <ul>${items.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>
    </div>
  `;
}

function reportMoldRiskBlock(moldRisk) {
  if (!moldRisk || !moldRisk.label) {
    return "";
  }
  const reasons = moldRisk.reasons?.length
    ? `<ul>${moldRisk.reasons.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>`
    : "";
  return `
    <div class="report-block">
      <h3>Schimmelrisiko</h3>
      <p>${escapeHtml(formatMoldRisk(moldRisk))} (${escapeHtml(String(moldRisk.score ?? 0))}/100)</p>
      <p>${escapeHtml(moldRisk.caveat || "")}</p>
      ${reasons}
    </div>
  `;
}

function detailCell(label, value) {
  return `
    <article class="detail-cell">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
    </article>
  `;
}

function replaceChatMessage(messageId, nextMessage) {
  const index = state.chatMessages.findIndex((message) => message.id === messageId);
  if (index === -1) {
    state.chatMessages.push(nextMessage);
    return;
  }
  state.chatMessages[index] = nextMessage;
}

function updateChatComposer() {
  dom.chatInput.disabled = state.chatLoading;
  dom.chatScopeToggle.disabled = state.chatLoading;
  dom.chatSubmitButton.disabled = state.chatLoading;
  dom.chatSubmitButton.textContent = state.chatLoading ? "Antwort wird erstellt..." : "Frage senden";
}

function formatNumber(value, digits = 1) {
  return Number(value || 0).toLocaleString("de-DE", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
  });
}

function formatCurrency(value) {
  return Number(value || 0).toLocaleString("de-DE", {
    style: "currency",
    currency: "EUR",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function formatMoldRisk(moldRisk) {
  if (!moldRisk || !moldRisk.label) {
    return "Nicht verfügbar";
  }
  return `${moldRisk.label} (${moldRisk.score ?? 0}/100)`;
}

function humanizeKey(value) {
  return value.replace(/_/g, " ").replace(/\b\w/g, (char) => char.toUpperCase());
}

function capitalize(value) {
  return value ? value.charAt(0).toUpperCase() + value.slice(1) : "";
}

function currentPeriodLabel() {
  return periodLabelForOffset(state.period, state.offset);
}

function previousPeriodLabel() {
  return periodLabelForOffset(state.period, state.offset - 1);
}

function nextPeriodLabel() {
  return periodLabelForOffset(state.period, state.offset + 1);
}

function periodLabelForOffset(period, offset) {
  const absOffset = Math.abs(offset);

  if (period === "day") {
    if (offset === 0) return "Heute";
    if (offset === -1) return "Gestern";
    if (offset === 1) return "Morgen";
    return offset < 0 ? `Vor ${absOffset} Tagen` : `In ${absOffset} Tagen`;
  }

  if (period === "week") {
    if (offset === 0) return "Diese Woche";
    if (offset === -1) return "Letzte Woche";
    if (offset === 1) return "Nächste Woche";
    return offset < 0 ? `Vor ${absOffset} Wochen` : `In ${absOffset} Wochen`;
  }

  if (period === "month") {
    if (offset === 0) return "Dieser Monat";
    if (offset === -1) return "Letzter Monat";
    if (offset === 1) return "Nächster Monat";
    return offset < 0 ? `Vor ${absOffset} Monaten` : `In ${absOffset} Monaten`;
  }

  if (period === "quarter") {
    if (offset === 0) return "Dieses Quartal";
    if (offset === -1) return "Letztes Quartal";
    if (offset === 1) return "Nächstes Quartal";
    return offset < 0 ? `Vor ${absOffset} Quartalen` : `In ${absOffset} Quartalen`;
  }

  if (offset === 0) return "Dieses Jahr";
  if (offset === -1) return "Letztes Jahr";
  if (offset === 1) return "Nächstes Jahr";
  return offset < 0 ? `Vor ${absOffset} Jahren` : `In ${absOffset} Jahren`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}
