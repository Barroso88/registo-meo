const API_BASE = "/api";
const AUTH_TOKEN_KEY = "registoMensalAuthToken";

const tabs = document.querySelectorAll(".tab");
const panels = document.querySelectorAll(".panel");

const billingForm = document.getElementById("sales-form");
const mobileForm = document.getElementById("mobile-form");
const energyForm = document.getElementById("energy-form");
const goalsForm = document.getElementById("goals-form");
const tasksForm = document.getElementById("tasks-form");
const stockForm = document.getElementById("stock-form");

const billingTableBody = document.getElementById("sales-table-body");
const mobileTableBody = document.getElementById("mobile-table-body");
const energyTableBody = document.getElementById("energy-table-body");
const reportTableBody = document.getElementById("report-table-body");
const generalTableBody = document.getElementById("general-table-body");
const tasksTableBody = document.getElementById("tasks-table-body");
const stockTableBody = document.getElementById("stock-table-body");
const accessButtons = document.getElementById("access-buttons");
const accessDetail = document.getElementById("access-detail");
const accessTotalCount = document.getElementById("access-total-count");
const accessFilledCount = document.getElementById("access-filled-count");

const monthSelector = document.getElementById("month-selector");
const selectedMonthLabel = document.getElementById("selected-month-label");
const themeSelector = document.getElementById("theme-selector");
const loginScreen = document.getElementById("login-screen");
const loginForm = document.getElementById("login-form");
const loginError = document.getElementById("login-error");
const logoutButton = document.getElementById("logout-button");
const authSettingsForm = document.getElementById("auth-settings-form");
const authUsernameInput = document.getElementById("auth-username");
const authCurrentPasswordInput = document.getElementById("auth-current-password");
const authPasswordInput = document.getElementById("auth-password");
const authPasswordConfirmInput = document.getElementById("auth-password-confirm");
const authSettingsMessage = document.getElementById("auth-settings-message");
const tasksNotificationButton = document.getElementById("tasks-notification-button");
const tasksNotificationText = document.getElementById("tasks-notification-text");
const tasksNotificationBadge = document.getElementById("tasks-notification-badge");
const stockNotificationButton = document.getElementById("stock-notification-button");
const stockNotificationText = document.getElementById("stock-notification-text");
const stockNotificationBadge = document.getElementById("stock-notification-badge");
const todayButtons = document.querySelectorAll(".date-today-button");
const supportTabs = document.querySelectorAll(".support-tab");
const supportPanels = document.querySelectorAll(".support-panel");
const calculatorDisplay = document.getElementById("calculator-display");
const calculatorWidthInput = document.getElementById("calculator-width");
const calculatorWidthValue = document.getElementById("calculator-width-value");
const incentiveTabs = document.querySelectorAll(".incentive-tab");
const menuFontSizeInput = document.getElementById("menu-font-size");
const titleFontSizeInput = document.getElementById("title-font-size");
const fieldFontSizeInput = document.getElementById("field-font-size");
const panelHeaderSizeInput = document.getElementById("panel-header-size");
const densityModeInput = document.getElementById("density-mode");
const primaryColorInput = document.getElementById("primary-color");
const secondaryColorInput = document.getElementById("secondary-color");
const sidebarColorInput = document.getElementById("sidebar-color");
const toolbarColorInput = document.getElementById("toolbar-color");
const formWidthInput = document.getElementById("form-width");
const cardRadiusInput = document.getElementById("card-radius");
const pageColorInputs = {
  relatorios: document.getElementById("page-color-relatorios"),
  faturacao: document.getElementById("page-color-faturacao"),
  movel: document.getElementById("page-color-movel"),
  "meo-energia": document.getElementById("page-color-meo-energia"),
  objetivo: document.getElementById("page-color-objetivo"),
  geral: document.getElementById("page-color-geral"),
  tarefas: document.getElementById("page-color-tarefas"),
  apoio: document.getElementById("page-color-apoio"),
  incentivos: document.getElementById("page-color-incentivos"),
  stock: document.getElementById("page-color-stock"),
  acessos: document.getElementById("page-color-acessos"),
  configuracoes: document.getElementById("page-color-configuracoes"),
};
const panelColorInputs = document.querySelectorAll(".panel-color-input");
const panelBlockInputs = document.querySelectorAll(".panel-block-input");
const panelFieldInputs = document.querySelectorAll(".panel-field-input");
const sidebarTabColorInputs = {
  relatorios: document.getElementById("sidebar-tab-color-relatorios"),
  faturacao: document.getElementById("sidebar-tab-color-faturacao"),
  movel: document.getElementById("sidebar-tab-color-movel"),
  "meo-energia": document.getElementById("sidebar-tab-color-meo-energia"),
  objetivo: document.getElementById("sidebar-tab-color-objetivo"),
  geral: document.getElementById("sidebar-tab-color-geral"),
  tarefas: document.getElementById("sidebar-tab-color-tarefas"),
  apoio: document.getElementById("sidebar-tab-color-apoio"),
  incentivos: document.getElementById("sidebar-tab-color-incentivos"),
  stock: document.getElementById("sidebar-tab-color-stock"),
  acessos: document.getElementById("sidebar-tab-color-acessos"),
  configuracoes: document.getElementById("sidebar-tab-color-configuracoes"),
};

const ACCESS_SERVICES = [
  { id: "sap-comercial", label: "SAP Comercial" },
  { id: "portal-sfa", label: "Portal SFA" },
  { id: "user-rede", label: "User de Rede" },
  { id: "pos-venda", label: "Pós-Venda" },
  { id: "blueticket", label: "Blueticket" },
  { id: "intelcia", label: "Intelcia" },
  { id: "eform", label: "EForm" },
  { id: "bec-deposito", label: "BEC Depósito" },
];

const DEFAULT_SIDEBAR_TAB_COLORS = {
  relatorios: "#b7f43f",
  faturacao: "#8fd7ff",
  movel: "#74c7ff",
  "meo-energia": "#9be85b",
  objetivo: "#6fe0b7",
  geral: "#8ab6ff",
  tarefas: "#7fdcc8",
  apoio: "#9bd0ff",
  incentivos: "#ffd166",
  stock: "#a8ea33",
  acessos: "#c7b2ff",
  configuracoes: "#9dc2ff",
};

const DEFAULT_PAGE_BLOCK_COLORS = {
  relatorios: "#13214d",
  faturacao: "#13214d",
  movel: "#13214d",
  "meo-energia": "#13214d",
  objetivo: "#13214d",
  geral: "#13214d",
  tarefas: "#13214d",
  apoio: "#13214d",
  incentivos: "#13214d",
  stock: "#13214d",
  acessos: "#13214d",
  configuracoes: "#13214d",
};

const DEFAULT_PAGE_FIELD_COLORS = {
  relatorios: "#0f1f56",
  faturacao: "#0f1f56",
  movel: "#0f1f56",
  "meo-energia": "#0f1f56",
  objetivo: "#0f1f56",
  geral: "#0f1f56",
  tarefas: "#0f1f56",
  apoio: "#0f1f56",
  incentivos: "#0f1f56",
  stock: "#0f1f56",
  acessos: "#0f1f56",
  configuracoes: "#0f1f56",
};

function buildDefaultAccessEntries(source = {}) {
  return Object.fromEntries(
    ACCESS_SERVICES.map(({ id }) => {
      const entry = source && typeof source === "object" ? source[id] : null;
      return [
        id,
        {
          user: String(entry?.user || ""),
          password: String(entry?.password || ""),
        },
      ];
    })
  );
}

const totalSales = document.getElementById("total-sales");
const totalAmount = document.getElementById("total-amount");
const grossValueInput = billingForm.elements.namedItem("valor");
const netValueInput = document.getElementById("net-value");
const insuranceValueInput = document.getElementById("insurance-value");
const insuranceAnnualValueInput = document.getElementById("insurance-annual-value");

const reportTotalNet = document.getElementById("report-total-net");
const reportTotalSmartphone = document.getElementById("report-total-smartphone");
const reportTotalDiversidade = document.getElementById("report-total-diversidade");
const reportTotalSeguros = document.getElementById("report-total-seguros");
const reportTotalMovel = document.getElementById("report-total-movel");
const reportTotalEnergia = document.getElementById("report-total-energia");

const goalTotalProgress = document.getElementById("goal-total-progress");
const goalBillingProgress = document.getElementById("goal-billing-progress");
const goalMobileProgress = document.getElementById("goal-mobile-progress");
const goalEnergyProgress = document.getElementById("goal-energy-progress");
const goalInsuranceProgress = document.getElementById("goal-insurance-progress");
const goalDiversityProgress = document.getElementById("goal-diversity-progress");

const goalActualBilling = document.getElementById("goal-actual-billing");
const goalTargetBilling = document.getElementById("goal-target-billing");
const goalPercentBilling = document.getElementById("goal-percent-billing");
const goalWeightedBilling = document.getElementById("goal-weighted-billing");
const goalActualMobile = document.getElementById("goal-actual-mobile");
const goalTargetMobile = document.getElementById("goal-target-mobile");
const goalPercentMobile = document.getElementById("goal-percent-mobile");
const goalWeightedMobile = document.getElementById("goal-weighted-mobile");
const goalActualEnergy = document.getElementById("goal-actual-energy");
const goalTargetEnergy = document.getElementById("goal-target-energy");
const goalPercentEnergy = document.getElementById("goal-percent-energy");
const goalWeightedEnergy = document.getElementById("goal-weighted-energy");
const goalActualInsurance = document.getElementById("goal-actual-insurance");
const goalTargetInsurance = document.getElementById("goal-target-insurance");
const goalPercentInsurance = document.getElementById("goal-percent-insurance");
const goalWeightedInsurance = document.getElementById("goal-weighted-insurance");
const goalActualDiversity = document.getElementById("goal-actual-diversity");
const goalTargetDiversity = document.getElementById("goal-target-diversity");
const goalPercentDiversity = document.getElementById("goal-percent-diversity");
const goalWeightedDiversity = document.getElementById("goal-weighted-diversity");
const tasksTotalCount = document.getElementById("tasks-total-count");
const tasksOpenCount = document.getElementById("tasks-open-count");
const tasksDoneCount = document.getElementById("tasks-done-count");
const stockTotalCount = document.getElementById("stock-total-count");
const stockPendingCount = document.getElementById("stock-pending-count");
const stockCompletedCount = document.getElementById("stock-completed-count");
const incentivesContent = document.getElementById("incentives-content");
const incentivesTotalValue = document.getElementById("incentives-total-value");
const incentivesTotalQuantity = document.getElementById("incentives-total-quantity");
const incentivesActiveModels = document.getElementById("incentives-active-models");

const INCENTIVES_CATALOG = [
  {
    brand: "Samsung",
    sections: [
      {
        title: "Smartphones",
        tone: "samsung",
        items: [
          ["A16 4G 128GB", 1.2],
          ["A17 4G 128GB", 0.8],
          ["A17 4G 256GB", 1.2],
          ["A17 5G 128GB", 1.2],
          ["A17 5G 256GB", 1.6],
          ["A26 5G 128GB", 1.6],
          ["A26 5G 256GB", 2],
          ["A36 5G 128GB", 2],
          ["A36 5G 256GB", 2.4],
          ["A37 5G 128GB", 4],
          ["A37 5G 256GB", 4],
          ["A56 5G 128GB", 2.8],
          ["A56 5G 256GB", 2.4],
          ["A57 5G 128GB", 4],
          ["A57 5G 256GB", 4.8],
          ["S25 FE 5G 128GB", 4.8],
          ["S25 FE 5G 256GB", 5.2],
          ["S25 128GB", 5.6],
          ["S25 256GB", 6],
          ["S25 512GB", 6.4],
          ["S25 Ultra 256GB", 7.6],
          ["S25 Ultra 512GB", 7.6],
          ["S25 Edge 256GB", 6.4],
          ["S25 Edge 512GB", 6.8],
          ["S26 256GB", 7.2],
          ["S26 512GB", 7.6],
          ["S26+ 256GB", 8],
          ["S26+ 512GB", 8.4],
          ["S26 Ultra 256GB", 8.8],
          ["S26 Ultra 512GB", 10.4],
          ["S26 Ultra 1TB", 12],
          ["Z Flip7 FE 256GB", 9.6],
          ["Z Flip7 5G 256GB", 10.4],
          ["Z Flip7 5G 512GB", 11.2],
          ["Z Fold7 5G 256GB", 13.6],
          ["Z Fold7 5G 512GB", 14.4],
          ["Z Fold7 5G 1TB", 16],
        ],
      },
    ],
  },
  {
    brand: "Wearables",
    sections: [
      {
        title: "Wearables",
        tone: "wearables",
        items: [
          ["Watch Ultra", 4],
          ["Watch7 40mm BT", 1.2],
          ["Watch8 40mm BT", 2.4],
          ["Watch8 40mm LTE", 2.8],
          ["Watch8 44mm BT", 2.8],
          ["Watch8 44mm LTE", 3.2],
          ["Watch8 Classic BT", 3.2],
          ["Watch8 Classic LTE", 3.6],
          ["Buds3 FE", 1.2],
          ["Buds4", 2.4],
          ["Buds4 Pro", 4.8],
        ],
      },
    ],
  },
  {
    brand: "tablets",
    displayName: "Tablet's",
    sections: [
      {
        title: "Tablet's",
        tone: "tablets",
        items: [
          ["Tab A9 64GB WiFi", 1.6],
          ["Tab A9+ 128GB WiFi", 2],
          ["Tab A11+ 128GB WiFi", 2.4],
          ["Tab A11+ 256GB WiFi", 2.8],
          ["Tab A11+ 128GB 5G", 2.8],
          ["Tab S10 Lite WiFi 128GB", 3.2],
          ["Tab S10 FE WiFi 128GB", 6],
          ["Tab S10 FE+ WiFi 128GB", 6.8],
          ["Tab S11 WiFi 128GB", 8],
          ["Tab S11 Ultra WiFi 256GB", 9.6],
        ],
      },
    ],
  },
  {
    brand: "Honor",
    sections: [
      {
        title: "Smartphones",
        tone: "honor",
        items: [
          ["HONOR 400 Smart", 6],
          ["HONOR 400 Lite", 8],
          ["HONOR 400", 28],
          ["HONOR 400 Pro", 40],
          ["HONOR Magic 7 Lite", 16],
          ["HONOR Magic 7 Pro", 60],
          ["HONOR Magic 8 Lite", 24],
          ["HONOR Magic 8 Pro", 60],
          ["HONOR Magic V5", 72],
        ],
      },
    ],
  },
  {
    brand: "Xiaomi",
    sections: [
      {
        title: "Smartphones",
        tone: "xiaomi",
        items: [
          ["Redmi 15C 128GB", 0.8],
          ["Redmi 15C 256GB", 1.6],
          ["Redmi 15C 5G 128GB", 1.6],
          ["Redmi 15C 5G 256GB", 2],
          ["Redmi Note 15 128GB", 2.4],
          ["Redmi Note 15 256GB", 2.4],
          ["Redmi Note 15 5G 256GB", 3.2],
          ["Redmi Note 15 Pro 5G 256GB", 4],
          ["Redmi Note 15 Pro+ 5G 512GB", 5.6],
          ["Xiaomi 17 512GB", 24],
          ["Xiaomi 17 Ultra 512GB", 32],
        ],
      },
    ],
  },
  {
    brand: "Motorola",
    sections: [
      {
        title: "Smartphones",
        tone: "motorola",
        items: [
          ["Razr 60 Ultra", 80],
          ["Razr 60", 72],
          ["Edge 70", 32],
          ["Edge 60 Fusion", 16],
          ["G35", 4],
          ["G06 8/256 Power", 3.2],
          ["G06 4/256", 2.4],
          ["G06 4/64", 1.6],
        ],
      },
    ],
  },
  {
    brand: "TCL",
    sections: [
      {
        title: "Smartphones / Tablets",
        tone: "tcl",
        items: [
          ["TCL NXTPAPER 11 PLUS Bundle", 16],
          ["TCL NXTPAPER 14", 24],
          ["TCL 60R 5G", 6.4],
          ["TCL 60 SE NXTPAPER 5G", 6.4],
        ],
      },
    ],
  },
];

const DEFAULT_PAGE_COLORS = {
  relatorios: "#14226d",
  faturacao: "#16286f",
  movel: "#132762",
  "meo-energia": "#17304f",
  objetivo: "#18295f",
  geral: "#11285b",
  tarefas: "#14325a",
  apoio: "#16245b",
  incentivos: "#17265a",
  stock: "#15315b",
  acessos: "#18295f",
  configuracoes: "#1b2557",
};

const state = {
  billingSales: [],
  mobileSales: [],
  energySales: [],
  tasks: [],
  stockEntries: [],
  accessEntries: buildDefaultAccessEntries(),
  monthlyIncentives: {},
  monthlyGoals: {},
  auth: {
    username: "admin",
    passwordHash: "",
  },
  settings: {
    theme: "light",
    menuFontSize: 15,
    titleFontSize: 22,
    fieldFontSize: 13,
    panelHeaderSize: 13,
    densityMode: "compact",
    primaryColor: "#b6f23b",
    secondaryColor: "#1730a1",
    sidebarColor: "#102061",
    toolbarColor: "#b3262d",
    sidebarTabColors: { ...DEFAULT_SIDEBAR_TAB_COLORS },
    pageColors: { ...DEFAULT_PAGE_COLORS },
    pageBlockColors: { ...DEFAULT_PAGE_BLOCK_COLORS },
    pageFieldColors: { ...DEFAULT_PAGE_FIELD_COLORS },
    calculatorWidth: 340,
    formWidth: 720,
    cardRadius: 24,
  },
};

let editingBillingIndex = null;
let editingMobileIndex = null;
let editingEnergyIndex = null;
let editingTaskIndex = null;
let editingStockIndex = null;
let selectedMonth = getCurrentMonth();
let selectedIncentiveBrand = "samsung";
let selectedAccessService = ACCESS_SERVICES[0].id;
let accessPasswordVisible = false;
let accessStatusMessage = "";
let accessStatusError = false;
let calculatorExpression = "0";
let calculatorJustEvaluated = false;

function clampCalculatorWidth(value) {
  return Math.min(460, Math.max(280, Number(value) || 360));
}

tabs.forEach((tab) => {
  tab.addEventListener("click", () => {
    activateTab(tab.dataset.tab);
  });
});

supportTabs.forEach((tab) => {
  tab.addEventListener("click", () => {
    if (tab.classList.contains("incentive-tab")) return;
    const target = tab.dataset.supportTab;
    supportTabs.forEach((item) => item.classList.toggle("active", item === tab));
    supportPanels.forEach((panel) =>
      panel.classList.toggle("active", panel.id === `support-${target}`)
    );
  });
});

incentiveTabs.forEach((tab) => {
  tab.addEventListener("click", () => {
    selectedIncentiveBrand = tab.dataset.incentiveTab || "samsung";
    incentiveTabs.forEach((item) => item.classList.toggle("active", item === tab));
    renderIncentives();
  });
});

accessButtons?.addEventListener("click", (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return;

  const button = target.closest("[data-access-service]");
  if (!(button instanceof HTMLElement)) return;

  const serviceId = button.dataset.accessService;
  if (!serviceId) return;

  selectedAccessService = serviceId;
  accessPasswordVisible = false;
  accessStatusMessage = "";
  accessStatusError = false;
  renderAccesses();
});

accessDetail?.addEventListener("click", async (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return;

  if (target.dataset.toggleAccessPassword !== undefined) {
    accessPasswordVisible = !accessPasswordVisible;
    renderAccesses();
    return;
  }

  if (target.dataset.copyAccessValue !== undefined) {
    const service = ACCESS_SERVICES.find((item) => item.id === selectedAccessService);
    const entry = state.accessEntries?.[selectedAccessService] || { user: "", password: "" };
    const value =
      target.dataset.copyAccessValue === "user"
        ? String(entry.user || "")
        : String(entry.password || "");

    if (!value) {
      accessStatusMessage = `Nao existe ${target.dataset.copyAccessValue === "user" ? "user" : "password"} guardado para ${service?.label || "este serviço"}.`;
      accessStatusError = true;
      renderAccesses();
      return;
    }

    try {
      await copyTextToClipboard(value);
      accessStatusMessage = `${target.dataset.copyAccessValue === "user" ? "User" : "Password"} copiado para ${service?.label || "o serviço selecionado"}.`;
      accessStatusError = false;
      renderAccesses();
    } catch {
      accessStatusMessage = "Nao foi possivel copiar o valor.";
      accessStatusError = true;
      renderAccesses();
    }
  }
});

accessDetail?.addEventListener("submit", async (event) => {
  const form = event.target;
  if (!(form instanceof HTMLFormElement) || form.id !== "access-form") return;
  event.preventDefault();

  const formData = new FormData(form);
  const serviceId = String(formData.get("service") || selectedAccessService);
  state.accessEntries = {
    ...buildDefaultAccessEntries(state.accessEntries),
    [serviceId]: {
      user: valueAsText(formData.get("user")),
      password: valueAsText(formData.get("password")),
    },
  };

  selectedAccessService = serviceId;
  accessPasswordVisible = false;
  accessStatusMessage = "Credenciais guardadas com sucesso.";
  accessStatusError = false;
  await persistAndRender();
});

document.addEventListener("click", (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return;

  const calculatorButton = target.closest("[data-calculator-value], [data-calculator-action]");
  if (!(calculatorButton instanceof HTMLElement)) return;
  if (!calculatorButton.closest("#support-calculadora")) return;

  const value = calculatorButton.dataset.calculatorValue;
  const action = calculatorButton.dataset.calculatorAction;

  if (value !== undefined) {
    appendCalculatorValue(value);
    return;
  }

  if (action === "clear") {
    clearCalculator();
    return;
  }

  if (action === "backspace") {
    backspaceCalculator();
    return;
  }

  if (action === "equals") {
    evaluateCalculator();
  }
});

document.addEventListener("keydown", (event) => {
  if (!isCalculatorKeyboardActive()) return;

  const key = event.key;
  if (event.ctrlKey || event.metaKey || event.altKey) return;

  if (/^\d$/.test(key)) {
    event.preventDefault();
    appendCalculatorValue(key);
    return;
  }

  if (key === "," || key === ".") {
    event.preventDefault();
    appendCalculatorValue(".");
    return;
  }

  if (["+", "-", "*", "/"].includes(key)) {
    event.preventDefault();
    appendCalculatorValue(key);
    return;
  }

  if (key === "%") {
    event.preventDefault();
    appendCalculatorValue("%");
    return;
  }

  if (key === "Enter" || key === "=") {
    event.preventDefault();
    evaluateCalculator();
    return;
  }

  if (key === "Backspace") {
    event.preventDefault();
    backspaceCalculator();
    return;
  }

  if (key === "Escape") {
    event.preventDefault();
    clearCalculator();
  }
});

tasksNotificationButton.addEventListener("click", () => {
  activateTab("tarefas");
});

stockNotificationButton?.addEventListener("click", () => {
  activateTab("stock");
});

loginForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  loginError.textContent = "";

  const formData = new FormData(loginForm);
  const username = String(formData.get("username") || "");
  const password = String(formData.get("password") || "");

  try {
    const response = await fetch(`${API_BASE}/auth/login`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      credentials: "include",
      body: JSON.stringify({ username, password }),
    });

    if (!response.ok) {
      const data = await response.json().catch(() => ({}));
      throw new Error(data.message || "Nao foi possivel entrar.");
    }

    const data = await response.json().catch(() => ({}));
    if (data?.token) {
      setAuthToken(data.token);
    }

    await startAuthenticatedApp();
  } catch (error) {
    loginError.textContent = error.message || "Nao foi possivel entrar.";
  }
});

logoutButton.addEventListener("click", async () => {
  clearAuthToken();
  try {
    await fetch(`${API_BASE}/auth/logout`, {
      method: "POST",
      headers: buildAuthHeaders(),
      credentials: "include",
    });
  } finally {
    window.location.reload();
  }
});

authSettingsForm?.addEventListener("submit", async (event) => {
  event.preventDefault();
  setAuthMessage("");

  const formData = new FormData(authSettingsForm);
  const username = String(formData.get("authUsername") || "").trim();
  const currentPassword = String(formData.get("authCurrentPassword") || "");
  const password = String(formData.get("authPassword") || "");
  const confirmPassword = String(formData.get("authPasswordConfirm") || "");

  if (!username) {
    setAuthMessage("O utilizador não pode estar vazio.", true);
    return;
  }

  if ((password || confirmPassword) && password !== confirmPassword) {
    setAuthMessage("A nova password e a confirmação não coincidem.", true);
    return;
  }

  try {
    const response = await fetch(`${API_BASE}/auth/settings`, {
      method: "PUT",
      headers: buildAuthHeaders({
        "Content-Type": "application/json",
      }),
      credentials: "include",
      body: JSON.stringify({ username, currentPassword, password }),
    });

    if (!response.ok) {
      const data = await response.json().catch(() => ({}));
      const debugSuffix = data.debug
        ? ` [host=${data.debug.host || "-"}, ip=${data.debug.ip || "-"}, hostname=${data.debug.hostname || "-"}]`
        : "";
      throw new Error((data.message || "Nao foi possivel guardar as credenciais.") + debugSuffix);
    }

    const data = await response.json().catch(() => ({}));
    state.auth = {
      username: String(data.username || username),
      passwordHash: String(data.passwordHash || state.auth.passwordHash || ""),
    };
    if (authUsernameInput) authUsernameInput.value = username;
    if (authCurrentPasswordInput) authCurrentPasswordInput.value = "";
    if (authPasswordInput) authPasswordInput.value = "";
    if (authPasswordConfirmInput) authPasswordConfirmInput.value = "";
    const loginUsernameInput = loginForm.elements.namedItem("username");
    if (loginUsernameInput instanceof HTMLInputElement) {
      loginUsernameInput.value = username;
    }
    setAuthMessage("Credenciais atualizadas com sucesso.");
  } catch (error) {
    setAuthMessage(error.message || "Nao foi possivel guardar as credenciais.", true);
  }
});

monthSelector.value = selectedMonth;
themeSelector.value = state.settings.theme;
hydrateVisualEditor();
hydrateAuthSettings();
updateMonthLabel();
syncFormDatesToMonth();
applyTheme();
applyVisualSettings();

monthSelector.addEventListener("input", () => {
  selectedMonth = monthSelector.value || getCurrentMonth();
  updateMonthLabel();
  syncFormDatesToMonth();
  renderAll();
});

themeSelector.addEventListener("input", async () => {
  state.settings.theme = themeSelector.value;
  applyTheme();
  await persistState();
});

menuFontSizeInput.addEventListener("input", handleVisualSettingsChange);
titleFontSizeInput.addEventListener("input", handleVisualSettingsChange);
fieldFontSizeInput.addEventListener("input", handleVisualSettingsChange);
panelHeaderSizeInput.addEventListener("input", handleVisualSettingsChange);
densityModeInput.addEventListener("input", handleVisualSettingsChange);
primaryColorInput.addEventListener("input", handleVisualSettingsChange);
secondaryColorInput.addEventListener("input", handleVisualSettingsChange);
sidebarColorInput.addEventListener("input", handleVisualSettingsChange);
toolbarColorInput.addEventListener("input", handleVisualSettingsChange);
calculatorWidthInput.addEventListener("input", handleCalculatorWidthChange);
formWidthInput.addEventListener("input", handleVisualSettingsChange);
cardRadiusInput.addEventListener("input", handleVisualSettingsChange);
Object.values(pageColorInputs).forEach((input) => {
  input.addEventListener("input", handleVisualSettingsChange);
});
panelBlockInputs.forEach((input) => {
  input.addEventListener("input", handlePanelBlockChange);
});
panelColorInputs.forEach((input) => {
  input.addEventListener("input", handlePanelColorChange);
});
panelFieldInputs.forEach((input) => {
  input.addEventListener("input", handlePanelFieldChange);
});
Object.values(sidebarTabColorInputs).forEach((input) => {
  input?.addEventListener("input", handleSidebarTabColorChange);
});

todayButtons.forEach((button) => {
  button.addEventListener("click", () => {
    const today = getTodayForSelectedMonth();
    const target = button.dataset.dateTarget;

    if (target === "billing") billingForm.elements.namedItem("data").value = today;
    if (target === "mobile") mobileForm.elements.namedItem("data").value = today;
    if (target === "energy") energyForm.elements.namedItem("data").value = today;
    if (target === "task") tasksForm.elements.namedItem("date").value = today;
  });
});

billingForm.addEventListener("submit", async (event) => {
  event.preventDefault();

  const formData = new FormData(billingForm);
  const grossValue = parseMoneyInput(formData.get("valor"));
  const sale = {
    data: valueAsText(formData.get("data")),
    nif: valueAsText(formData.get("nif")),
    equipamento: valueAsText(formData.get("equipamento")),
    valor: grossValue,
    valorSemIva: calculateNetValue(grossValue),
    seguro: valueAsText(formData.get("seguro")),
    valorSeguro: Number(formData.get("valorSeguro") || 0),
    valorSeguroAnual: calculateInsuranceAnnual(Number(formData.get("valorSeguro") || 0)),
    categoria: valueAsText(formData.get("categoria")),
    imei: valueAsText(formData.get("imei")),
    fatura: valueAsText(formData.get("fatura")),
    modalidade: valueAsText(formData.get("modalidade")),
    notas: valueAsText(formData.get("notas")) || "-",
  };

  if (editingBillingIndex === null) {
    state.billingSales.unshift(sale);
  } else {
    state.billingSales[editingBillingIndex] = sale;
  }

  editingBillingIndex = null;
  billingForm.reset();
  netValueInput.value = "";
  insuranceAnnualValueInput.value = "";
  syncFormDatesToMonth();
  await persistAndRender();
});

mobileForm.addEventListener("submit", async (event) => {
  event.preventDefault();

  const formData = new FormData(mobileForm);
  const sale = {
    data: valueAsText(formData.get("data")),
    nome: valueAsText(formData.get("nome")),
    nif: valueAsText(formData.get("nif")),
    numero: valueAsText(formData.get("numero")),
    tarifario: valueAsText(formData.get("tarifario")),
    idVenda: valueAsText(formData.get("idVenda")),
    fidelizacao: valueAsText(formData.get("fidelizacao")),
    portabilidade: valueAsText(formData.get("portabilidade")),
    bestOffer: valueAsText(formData.get("bestOffer")),
    netSegura: valueAsText(formData.get("netSegura")),
    rceAceite: valueAsText(formData.get("rceAceite")),
    formaAceitacao: valueAsText(formData.get("formaAceitacao")),
  };

  if (editingMobileIndex === null) {
    state.mobileSales.unshift(sale);
  } else {
    state.mobileSales[editingMobileIndex] = sale;
  }

  editingMobileIndex = null;
  mobileForm.reset();
  syncFormDatesToMonth();
  await persistAndRender();
});

energyForm.addEventListener("submit", async (event) => {
  event.preventDefault();

  const formData = new FormData(energyForm);
  const sale = {
    data: valueAsText(formData.get("data")),
    nome: valueAsText(formData.get("nome")),
    nif: valueAsText(formData.get("nif")),
    idVenda: valueAsText(formData.get("idVenda")),
    requisicao: valueAsText(formData.get("requisicao")),
    potencia: valueAsText(formData.get("potencia")),
    estado: valueAsText(formData.get("estado")),
    observacoes: valueAsText(formData.get("observacoes")) || "-",
  };

  if (editingEnergyIndex === null) {
    state.energySales.unshift(sale);
  } else {
    state.energySales[editingEnergyIndex] = sale;
  }

  editingEnergyIndex = null;
  energyForm.reset();
  syncFormDatesToMonth();
  await persistAndRender();
});

goalsForm.addEventListener("submit", async (event) => {
  event.preventDefault();

  const formData = new FormData(goalsForm);
  state.monthlyGoals[selectedMonth] = {
    billing: Number(formData.get("billingGoal") || 0),
    mobile: Number(formData.get("mobileGoal") || 0),
    energy: Number(formData.get("energyGoal") || 0),
  };

  await persistAndRender();
});

tasksForm.addEventListener("submit", async (event) => {
  event.preventDefault();

  const formData = new FormData(tasksForm);
  const task = {
    date: valueAsText(formData.get("date")),
    title: valueAsText(formData.get("title")),
    notes: valueAsText(formData.get("notes")) || "-",
    done: editingTaskIndex !== null ? Boolean(state.tasks[editingTaskIndex]?.done) : false,
  };

  if (editingTaskIndex === null) {
    state.tasks.unshift(task);
  } else {
    state.tasks[editingTaskIndex] = task;
  }

  editingTaskIndex = null;
  tasksForm.reset();
  syncFormDatesToMonth();
  await persistAndRender();
});

stockForm.addEventListener("submit", async (event) => {
  event.preventDefault();

  const formData = new FormData(stockForm);
  const entry = {
    dataPedido: valueAsText(formData.get("dataPedido")),
    clienteColaborador: valueAsText(formData.get("clienteColaborador")),
    identificador: valueAsText(formData.get("identificador")),
    contacto: valueAsText(formData.get("contacto")),
    equipamento: valueAsText(formData.get("equipamento")),
    estadoPedido: valueAsText(formData.get("estadoPedido")),
    contactado: valueAsText(formData.get("contactado")),
  };

  if (editingStockIndex === null) {
    state.stockEntries.unshift(entry);
  } else {
    state.stockEntries[editingStockIndex] = entry;
  }

  editingStockIndex = null;
  stockForm.reset();
  syncFormDatesToMonth();
  await persistAndRender();
});

grossValueInput.addEventListener("input", () => {
  const value = parseMoneyInput(grossValueInput.value);
  const hasValue = grossValueInput.value.trim().length > 0;
  netValueInput.value = hasValue
    ? formatPlainCurrency(calculateNetValue(value))
    : "";
});

insuranceValueInput.addEventListener("input", () => {
  const value = Number(insuranceValueInput.value || 0);
  insuranceAnnualValueInput.value = insuranceValueInput.value
    ? formatPlainCurrency(calculateInsuranceAnnual(value))
    : "";
});

billingTableBody.addEventListener("click", async (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return;

  if (target.dataset.editBilling !== undefined) {
    editBillingSale(Number(target.dataset.editBilling));
  }

  if (target.dataset.deleteBilling !== undefined) {
    deleteBillingSale(Number(target.dataset.deleteBilling));
    await persistAndRender();
  }
});

mobileTableBody.addEventListener("click", async (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return;

  if (target.dataset.editMobile !== undefined) {
    editMobileSale(Number(target.dataset.editMobile));
  }

  if (target.dataset.deleteMobile !== undefined) {
    deleteMobileSale(Number(target.dataset.deleteMobile));
    await persistAndRender();
  }
});

energyTableBody.addEventListener("click", async (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return;

  if (target.dataset.editEnergy !== undefined) {
    editEnergySale(Number(target.dataset.editEnergy));
  }

  if (target.dataset.deleteEnergy !== undefined) {
    deleteEnergySale(Number(target.dataset.deleteEnergy));
    await persistAndRender();
  }
});

tasksTableBody.addEventListener("click", async (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return;

  if (target.dataset.editTask !== undefined) {
    editTask(Number(target.dataset.editTask));
  }

  if (target.dataset.deleteTask !== undefined) {
    state.tasks.splice(Number(target.dataset.deleteTask), 1);
    if (editingTaskIndex !== null && Number(target.dataset.deleteTask) <= editingTaskIndex) {
      editingTaskIndex = editingTaskIndex === Number(target.dataset.deleteTask) ? null : editingTaskIndex - 1;
    }
    tasksForm.reset();
    syncFormDatesToMonth();
    await persistAndRender();
  }

  if (target.dataset.toggleTask !== undefined) {
    const index = Number(target.dataset.toggleTask);
    const task = state.tasks[index];
    if (task) {
      task.done = !task.done;
      await persistAndRender();
    }
  }
});

stockTableBody.addEventListener("click", async (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return;

  if (target.dataset.editStock !== undefined) {
    editStockEntry(Number(target.dataset.editStock));
  }

  if (target.dataset.deleteStock !== undefined) {
    deleteStockEntry(Number(target.dataset.deleteStock));
    await persistAndRender();
  }
});

incentivesContent.addEventListener("change", async (event) => {
  const target = event.target;
  if (!(target instanceof HTMLSelectElement)) return;
  if (!target.dataset.incentiveId) return;

  const monthlyValues = getCurrentIncentives();
  const nextValue = Number(target.value || 0);

  if (nextValue > 0) {
    monthlyValues[target.dataset.incentiveId] = nextValue;
  } else {
    delete monthlyValues[target.dataset.incentiveId];
  }

  state.monthlyIncentives[selectedMonth] = monthlyValues;
  renderIncentives();
  await persistState();
});

async function initialize() {
  try {
    const remoteState = await fetchState();
    state.billingSales = Array.isArray(remoteState.billingSales) ? remoteState.billingSales : [];
    state.mobileSales = Array.isArray(remoteState.mobileSales) ? remoteState.mobileSales : [];
    state.energySales = Array.isArray(remoteState.energySales) ? remoteState.energySales : [];
    state.tasks = Array.isArray(remoteState.tasks) ? remoteState.tasks : [];
    state.stockEntries = Array.isArray(remoteState.stockEntries) ? remoteState.stockEntries : [];
    state.accessEntries = buildDefaultAccessEntries(remoteState.accessEntries);
    state.monthlyIncentives =
      remoteState.monthlyIncentives && typeof remoteState.monthlyIncentives === "object"
        ? remoteState.monthlyIncentives
        : {};
    state.monthlyGoals =
      remoteState.monthlyGoals && typeof remoteState.monthlyGoals === "object"
        ? remoteState.monthlyGoals
        : {};
    state.auth = {
      username: String(remoteState.auth?.username || "admin"),
      passwordHash: String(remoteState.auth?.passwordHash || ""),
    };
    state.settings = {
      theme: remoteState.settings?.theme === "dark" ? "dark" : "light",
      menuFontSize: Number(remoteState.settings?.menuFontSize || 15),
      titleFontSize: Number(remoteState.settings?.titleFontSize || 22),
      fieldFontSize: Number(remoteState.settings?.fieldFontSize || 13),
      panelHeaderSize: Number(remoteState.settings?.panelHeaderSize || 13),
      densityMode: remoteState.settings?.densityMode === "normal" ? "normal" : "compact",
      primaryColor: remoteState.settings?.primaryColor || "#b6f23b",
      secondaryColor: remoteState.settings?.secondaryColor || "#1730a1",
      sidebarColor: remoteState.settings?.sidebarColor || "#102061",
      toolbarColor: remoteState.settings?.toolbarColor || "#b3262d",
      sidebarTabColors: {
        ...DEFAULT_SIDEBAR_TAB_COLORS,
        ...(remoteState.settings?.sidebarTabColors && typeof remoteState.settings.sidebarTabColors === "object"
          ? remoteState.settings.sidebarTabColors
          : {}),
      },
      pageColors: {
        ...DEFAULT_PAGE_COLORS,
        ...(remoteState.settings?.pageColors && typeof remoteState.settings.pageColors === "object"
          ? remoteState.settings.pageColors
          : {}),
      },
      pageBlockColors: {
        ...DEFAULT_PAGE_BLOCK_COLORS,
        ...(remoteState.settings?.pageBlockColors &&
        typeof remoteState.settings.pageBlockColors === "object"
          ? remoteState.settings.pageBlockColors
          : {}),
      },
      pageFieldColors: {
        ...DEFAULT_PAGE_FIELD_COLORS,
        ...(remoteState.settings?.pageFieldColors &&
        typeof remoteState.settings.pageFieldColors === "object"
          ? remoteState.settings.pageFieldColors
          : {}),
      },
      calculatorWidth: clampCalculatorWidth(remoteState.settings?.calculatorWidth || 340),
      formWidth: Number(remoteState.settings?.formWidth || 720),
      cardRadius: Number(remoteState.settings?.cardRadius || 24),
    };
  } catch (error) {
    console.error("Erro ao carregar dados da API:", error);
  }

  themeSelector.value = state.settings.theme;
  hydrateVisualEditor();
  hydrateAuthSettings();
  applyTheme();
  applyVisualSettings();
  renderAll();
}

async function fetchState() {
  const response = await fetch(`${API_BASE}/state`, {
    headers: buildAuthHeaders(),
    credentials: "include",
  });
  if (!response.ok) {
    throw new Error(`Falha ao carregar estado: ${response.status}`);
  }
  return response.json();
}

async function persistState() {
  const response = await fetch(`${API_BASE}/state`, {
    method: "PUT",
    headers: buildAuthHeaders({
      "Content-Type": "application/json",
    }),
    credentials: "include",
    body: JSON.stringify(state),
  });

  if (!response.ok) {
    throw new Error(`Falha ao guardar estado: ${response.status}`);
  }
}

async function fetchAuthStatus() {
  const response = await fetch(`${API_BASE}/auth/status`, {
    headers: buildAuthHeaders(),
    credentials: "include",
  });
  if (!response.ok) return { authenticated: false };
  return response.json();
}

function hydrateAuthSettings() {
  if (!authSettingsForm) return;
  if (authUsernameInput) authUsernameInput.value = state.auth.username || "admin";
  if (authCurrentPasswordInput) authCurrentPasswordInput.value = "";
  if (authPasswordInput) authPasswordInput.value = "";
  if (authPasswordConfirmInput) authPasswordConfirmInput.value = "";
  setAuthMessage("");
}

async function startAuthenticatedApp() {
  document.body.classList.add("authenticated");
  await initialize();
}

async function bootstrap() {
  try {
    const status = await fetchAuthStatus();
    if (!status.authenticated) {
      const loginUsernameInput = loginForm.elements.namedItem("username");
      if (loginUsernameInput instanceof HTMLInputElement) {
        loginUsernameInput.value = status.username || "admin";
      }
      clearAuthToken();
      document.body.classList.remove("authenticated");
      return;
    }

    await startAuthenticatedApp();
  } catch (error) {
    console.error("Falha ao verificar autenticacao:", error);
    clearAuthToken();
    document.body.classList.remove("authenticated");
  }
}

async function persistAndRender() {
  renderAll();
  await persistState();
}

function renderAll() {
  hydrateGoalsForm();
  renderBillingSales();
  renderMobileSales();
  renderEnergySales();
  renderTasks();
  renderTasksNotification();
  renderStockNotification();
  renderStockEntries();
  renderAccesses();
  renderReport();
  renderGoals();
  renderGeneral();
  renderIncentives();
  updateCalculatorDisplay();
}

function isCalculatorKeyboardActive() {
  const calculatorPanel = document.getElementById("support-calculadora");
  return Boolean(calculatorPanel?.classList.contains("active"));
}

function renderBillingSales() {
  const visibleSales = state.billingSales.filter((sale) => isInSelectedMonth(sale.data));

  if (visibleSales.length === 0) {
    billingTableBody.innerHTML = emptyRow(14, "Ainda não existem registos de faturação.");
  } else {
    billingTableBody.innerHTML = visibleSales
      .map(
        (sale) => `
          <tr>
            <td>${formatDate(sale.data)}</td>
            <td>${escapeHtml(sale.nif || "")}</td>
            <td>${escapeHtml(sale.equipamento || "")}</td>
            <td>${formatCurrency(Number(sale.valorSemIva || 0))}</td>
            <td>${escapeHtml(sale.seguro || "")}</td>
            <td>${formatCurrency(Number(sale.valorSeguro || 0))}</td>
            <td>${escapeHtml(sale.categoria || "")}</td>
            <td>${escapeHtml(sale.imei || "")}</td>
            <td>${escapeHtml(sale.fatura || "")}</td>
            <td>${escapeHtml(sale.modalidade || "")}</td>
            <td>${escapeHtml(sale.notas || "")}</td>
            <td class="actions-cell">
              <button type="button" class="table-action" data-edit-billing="${state.billingSales.indexOf(sale)}">Editar</button>
              <button type="button" class="table-action delete" data-delete-billing="${state.billingSales.indexOf(sale)}">Eliminar</button>
            </td>
          </tr>
        `
      )
      .join("");
  }

  totalSales.textContent = String(visibleSales.length);
  totalAmount.textContent = formatCurrency(
    visibleSales.reduce((sum, sale) => sum + getBillingEffectiveValue(sale), 0)
  );
}

function renderMobileSales() {
  const visibleSales = state.mobileSales.filter((sale) => isInSelectedMonth(sale.data));

  if (visibleSales.length === 0) {
    mobileTableBody.innerHTML = emptyRow(13, "Ainda não existem registos móveis.");
    return;
  }

  mobileTableBody.innerHTML = visibleSales
    .map(
      (sale) => `
        <tr>
          <td>${formatDate(sale.data)}</td>
          <td>${escapeHtml(sale.nome || "")}</td>
          <td>${escapeHtml(sale.nif || "")}</td>
          <td>${escapeHtml(sale.numero || "")}</td>
          <td>${escapeHtml(sale.tarifario || "")}</td>
          <td>${escapeHtml(sale.idVenda || "")}</td>
          <td>${escapeHtml(sale.fidelizacao || "")}</td>
          <td>${escapeHtml(sale.portabilidade || "")}</td>
          <td>${escapeHtml(sale.bestOffer || "")}</td>
          <td>${escapeHtml(sale.netSegura || "")}</td>
          <td>${escapeHtml(sale.rceAceite || "")}</td>
          <td>${escapeHtml(sale.formaAceitacao || "")}</td>
          <td class="actions-cell">
            <button type="button" class="table-action" data-edit-mobile="${state.mobileSales.indexOf(sale)}">Editar</button>
            <button type="button" class="table-action delete" data-delete-mobile="${state.mobileSales.indexOf(sale)}">Eliminar</button>
          </td>
        </tr>
      `
    )
    .join("");
}

function renderEnergySales() {
  const visibleSales = state.energySales.filter((sale) => isInSelectedMonth(sale.data));

  if (visibleSales.length === 0) {
    energyTableBody.innerHTML = emptyRow(9, "Ainda não existem registos de energia.");
    return;
  }

  energyTableBody.innerHTML = visibleSales
    .map(
      (sale) => `
        <tr>
          <td>${formatDate(sale.data)}</td>
          <td>${escapeHtml(sale.nome || "")}</td>
          <td>${escapeHtml(sale.nif || "")}</td>
          <td>${escapeHtml(sale.idVenda || "")}</td>
          <td>${escapeHtml(sale.requisicao || "")}</td>
          <td>${escapeHtml(sale.potencia || "")}</td>
          <td>${escapeHtml(sale.estado || "")}</td>
          <td>${escapeHtml(sale.observacoes || "")}</td>
          <td class="actions-cell">
            <button type="button" class="table-action" data-edit-energy="${state.energySales.indexOf(sale)}">Editar</button>
            <button type="button" class="table-action delete" data-delete-energy="${state.energySales.indexOf(sale)}">Eliminar</button>
          </td>
        </tr>
      `
    )
    .join("");
}

function renderTasks() {
  const visibleTasks = state.tasks.filter((task) => isInSelectedMonth(task.date));
  const doneCount = visibleTasks.filter((task) => task.done).length;

  tasksTotalCount.textContent = String(visibleTasks.length);
  tasksOpenCount.textContent = String(visibleTasks.length - doneCount);
  tasksDoneCount.textContent = String(doneCount);

  if (visibleTasks.length === 0) {
    tasksTableBody.innerHTML = emptyRow(5, "Ainda não existem tarefas.");
    return;
  }

  tasksTableBody.innerHTML = visibleTasks
    .map(
      (task) => `
        <tr class="${task.done ? "task-done" : ""}">
          <td>${formatDate(task.date)}</td>
          <td>${escapeHtml(task.title || "")}</td>
          <td>${escapeHtml(task.notes || "")}</td>
          <td class="task-status ${task.done ? "done" : ""}">${task.done ? "Concluída" : "Pendente"}</td>
          <td class="actions-cell">
            <button type="button" class="table-action" data-toggle-task="${state.tasks.indexOf(task)}">${task.done ? "Reabrir" : "Concluir"}</button>
            <button type="button" class="table-action" data-edit-task="${state.tasks.indexOf(task)}">Editar</button>
            <button type="button" class="table-action delete" data-delete-task="${state.tasks.indexOf(task)}">Eliminar</button>
          </td>
        </tr>
      `
    )
    .join("");
}

function renderStockEntries() {
  const visibleEntries = state.stockEntries.filter((entry) => isInSelectedMonth(entry.dataPedido));
  const completedCount = visibleEntries.filter(
    (entry) => entry.estadoPedido === "Pedido Abastecido"
  ).length;

  stockTotalCount.textContent = String(visibleEntries.length);
  stockPendingCount.textContent = String(visibleEntries.length - completedCount);
  stockCompletedCount.textContent = String(completedCount);

  if (visibleEntries.length === 0) {
    stockTableBody.innerHTML = emptyRow(8, "Ainda não existem pedidos de stock.");
    return;
  }

  stockTableBody.innerHTML = visibleEntries
    .map(
      (entry) => `
        <tr>
          <td>${formatDate(entry.dataPedido)}</td>
          <td>${escapeHtml(entry.clienteColaborador || "")}</td>
          <td>${escapeHtml(entry.identificador || "")}</td>
          <td>${escapeHtml(entry.contacto || "")}</td>
          <td>${escapeHtml(entry.equipamento || "")}</td>
          <td>${escapeHtml(entry.estadoPedido || "")}</td>
          <td>${escapeHtml(entry.contactado || "")}</td>
          <td class="actions-cell">
            <button type="button" class="table-action" data-edit-stock="${state.stockEntries.indexOf(entry)}">Editar</button>
            <button type="button" class="table-action delete" data-delete-stock="${state.stockEntries.indexOf(entry)}">Eliminar</button>
          </td>
        </tr>
      `
    )
    .join("");
}

function renderAccesses() {
  if (!accessButtons || !accessDetail) return;

  const entries = state.accessEntries || buildDefaultAccessEntries();
  if (!ACCESS_SERVICES.some((service) => service.id === selectedAccessService)) {
    selectedAccessService = ACCESS_SERVICES[0].id;
  }

  const activeService = ACCESS_SERVICES.find((service) => service.id === selectedAccessService) || ACCESS_SERVICES[0];
  const activeEntry = entries[activeService.id] || { user: "", password: "" };
  const filledCount = ACCESS_SERVICES.filter((service) => {
    const entry = entries[service.id];
    return Boolean(entry?.user || entry?.password);
  }).length;

  if (accessTotalCount) accessTotalCount.textContent = String(ACCESS_SERVICES.length);
  if (accessFilledCount) accessFilledCount.textContent = String(filledCount);

  accessButtons.innerHTML = ACCESS_SERVICES.map((service) => {
    const entry = entries[service.id] || { user: "", password: "" };
    const hasData = Boolean(entry.user || entry.password);
    return `
      <button type="button" class="access-service-button ${service.id === activeService.id ? "active" : ""}" data-access-service="${service.id}">
        <span class="access-service-label">${escapeHtml(service.label)}</span>
        <span class="access-service-meta">${hasData ? "Credenciais guardadas" : "Ainda vazio"}</span>
      </button>
    `;
  }).join("");

  accessDetail.innerHTML = `
    <div class="access-detail-card">
      <div class="access-detail-head">
        <div>
          <p class="eyebrow">Acesso rápido</p>
          <h3>${escapeHtml(activeService.label)}</h3>
          <p class="access-detail-copy">Guarda o user e a password deste serviço para consulta rápida.</p>
        </div>
      </div>

      <form class="access-form" id="access-form">
        <input type="hidden" name="service" value="${escapeHtml(activeService.id)}" />

        <label class="settings-field">
          User
          <input
            type="text"
            name="user"
            autocomplete="off"
            placeholder="Escrever user"
            value="${escapeHtml(activeEntry.user || "")}"
          />
        </label>

        <label class="settings-field">
          Password
          <div class="access-password-row">
            <input
              type="${accessPasswordVisible ? "text" : "password"}"
              name="password"
              autocomplete="off"
              placeholder="Escrever password"
              value="${escapeHtml(activeEntry.password || "")}"
            />
            <button type="button" class="secondary-button access-inline-button" data-toggle-access-password>
              ${accessPasswordVisible ? "Ocultar" : "Mostrar"}
            </button>
          </div>
        </label>

        <div class="access-actions">
          <button type="button" class="secondary-button access-inline-button" data-copy-access-value="user">Copiar user</button>
          <button type="button" class="secondary-button access-inline-button" data-copy-access-value="password">Copiar password</button>
          <button type="submit" class="primary-button">Guardar</button>
        </div>

        <p class="access-message ${accessStatusError ? "is-error" : ""}" id="access-message" aria-live="polite">
          ${escapeHtml(accessStatusMessage || "Escolhe um serviço na lista ao lado para editar as credenciais.")}
        </p>
      </form>
    </div>
  `;
}

function renderTasksNotification() {
  const openTasks = state.tasks.filter((task) => isInSelectedMonth(task.date) && !task.done);
  const pendingCount = openTasks.length;

  tasksNotificationBadge.textContent = String(pendingCount);
  tasksNotificationText.textContent =
    pendingCount > 0
      ? `${pendingCount} tarefa${pendingCount === 1 ? "" : "s"} por concluir`
      : "Sem tarefas pendentes";

  tasksNotificationButton.classList.toggle("has-pending", pendingCount > 0);
}

function renderStockNotification() {
  if (!stockNotificationButton || !stockNotificationText || !stockNotificationBadge) return;

  const pendingCount = state.stockEntries.filter(
    (entry) =>
      isInSelectedMonth(entry.dataPedido) &&
      entry.estadoPedido === "Pedido Abastecido" &&
      entry.contactado === "Não"
  ).length;

  stockNotificationBadge.textContent = String(pendingCount);
  stockNotificationText.textContent =
    pendingCount > 0
      ? `${pendingCount} pedido${pendingCount === 1 ? "" : "s"} por contactar`
      : "Sem pedidos por contactar";

  stockNotificationButton.classList.toggle("has-pending", pendingCount > 0);
}

function renderReport() {
  const summaryByDate = new Map();
  const totals = {
    totalSemIva: 0,
    smartphone: 0,
    diversidade: 0,
    seguros: 0,
    movel: 0,
    energia: 0,
  };

  state.billingSales.forEach((sale) => {
    if (!isInSelectedMonth(sale.data)) return;
    const entry = getReportEntry(summaryByDate, sale.data);
    entry.totalSemIva += getBillingEffectiveValue(sale);
    totals.totalSemIva += getBillingEffectiveValue(sale);
    if (sale.categoria === "Smartphone") {
      entry.smartphone += 1;
      totals.smartphone += 1;
    }
    if (sale.categoria === "Diversidade") {
      entry.diversidade += 1;
      totals.diversidade += 1;
    }
    if (sale.seguro === "Sim") {
      entry.seguros += 1;
      totals.seguros += 1;
    }
  });

  state.mobileSales.forEach((sale) => {
    if (!isInSelectedMonth(sale.data)) return;
    const entry = getReportEntry(summaryByDate, sale.data);
    entry.movel += 1;
    totals.movel += 1;
  });

  state.energySales.forEach((sale) => {
    if (!isInSelectedMonth(sale.data)) return;
    const entry = getReportEntry(summaryByDate, sale.data);
    entry.energia += 1;
    totals.energia += 1;
  });

  reportTotalNet.textContent = formatCurrency(totals.totalSemIva);
  reportTotalSmartphone.textContent = String(totals.smartphone);
  reportTotalDiversidade.textContent = String(totals.diversidade);
  reportTotalSeguros.textContent = String(totals.seguros);
  reportTotalMovel.textContent = String(totals.movel);
  reportTotalEnergia.textContent = String(totals.energia);

  const rows = Array.from(summaryByDate.entries()).sort(([a], [b]) =>
    a < b ? 1 : a > b ? -1 : 0
  );

  if (rows.length === 0) {
    reportTableBody.innerHTML = emptyRow(
      7,
      "Ainda não existem dados para apresentar no relatório."
    );
    return;
  }

  reportTableBody.innerHTML = rows
    .map(
      ([date, entry]) => `
        <tr>
          <td>${formatDate(date)}</td>
          <td>${formatCurrency(entry.totalSemIva)}</td>
          <td>${entry.smartphone}</td>
          <td>${entry.diversidade}</td>
          <td>${entry.seguros}</td>
          <td>${entry.movel}</td>
          <td>${entry.energia}</td>
        </tr>
      `
    )
    .join("");
}

function renderGoals() {
  const metrics = getGoalMetrics(selectedMonth);

  goalActualBilling.textContent = formatCurrency(metrics.actualBilling);
  goalTargetBilling.textContent = formatCurrency(metrics.goals.billing);
  goalPercentBilling.textContent = formatPercent(metrics.billingRatio * 100);
  goalWeightedBilling.textContent = formatPercent(metrics.billingWeighted);

  goalActualMobile.textContent = String(metrics.actualMobile);
  goalTargetMobile.textContent = String(metrics.goals.mobile);
  goalPercentMobile.textContent = formatPercent(metrics.mobileRatio * 100);
  goalWeightedMobile.textContent = formatPercent(metrics.mobileWeighted);

  goalActualEnergy.textContent = String(metrics.actualEnergy);
  goalTargetEnergy.textContent = String(metrics.goals.energy);
  goalPercentEnergy.textContent = formatPercent(metrics.energyRatio * 100);
  goalWeightedEnergy.textContent = formatPercent(metrics.energyWeighted);

  goalBillingProgress.textContent = formatPercent(metrics.billingRatio * 100);
  goalMobileProgress.textContent = formatPercent(metrics.mobileRatio * 100);
  goalEnergyProgress.textContent = formatPercent(metrics.energyRatio * 100);
  goalTotalProgress.textContent = formatPercent(metrics.totalWeighted);
  goalInsuranceProgress.textContent = formatPercent(metrics.insuranceProgressRatio * 100);
  goalDiversityProgress.textContent = formatPercent(metrics.diversityProgressRatio * 100);

  goalActualInsurance.textContent = formatPercent(metrics.insuranceActualRatio * 100);
  goalTargetInsurance.textContent = formatPercent(metrics.insuranceGoalRatio * 100);
  goalPercentInsurance.textContent = formatPercent(metrics.insuranceProgressRatio * 100);
  goalWeightedInsurance.textContent = `${metrics.smartphoneWithInsurance} / ${metrics.smartphoneCount}`;

  goalActualDiversity.textContent = formatPercent(metrics.diversityActualRatio * 100);
  goalTargetDiversity.textContent = formatPercent(metrics.diversityGoalRatio * 100);
  goalPercentDiversity.textContent = formatPercent(metrics.diversityProgressRatio * 100);
  goalWeightedDiversity.textContent = `${metrics.diversidadeCount} / ${metrics.smartphoneCount}`;
}

function renderGeneral() {
  const months = getTrackedMonths();

  generalTableBody.innerHTML = months
    .map((month) => {
      const metrics = getGoalMetrics(month);
      return `
        <tr>
          <td>${formatMonth(month)}</td>
          <td>${formatPercent(metrics.billingRatio * 100)}</td>
          <td>${formatPercent(metrics.mobileRatio * 100)}</td>
          <td>${formatPercent(metrics.energyRatio * 100)}</td>
          <td>${formatPercent(metrics.insuranceProgressRatio * 100)}</td>
          <td>${formatPercent(metrics.diversityProgressRatio * 100)}</td>
          <td>${formatPercent(metrics.totalWeighted)}</td>
        </tr>
      `;
    })
    .join("");
}

function renderIncentives() {
  const monthlyValues = getCurrentIncentives();
  const totals = {
    value: 0,
    quantity: 0,
    models: 0,
  };

  incentivesContent.innerHTML = INCENTIVES_CATALOG.map((brand) => {
    const brandId = slugify(brand.brand);
    const brandLabel = brand.displayName || brand.brand;
    const sectionsMarkup = brand.sections
      .map((section) => {
        const rows = section.items
          .map(([name, incentiveValue]) => {
            const id = createIncentiveId(brand.brand, section.title, name);
            const quantity = Number(monthlyValues[id] || 0);
            const subtotal = quantity * incentiveValue;

            totals.value += subtotal;
            totals.quantity += quantity;
            if (quantity > 0) totals.models += 1;

            return `
              <tr>
                <td>${escapeHtml(name)}</td>
                <td>${formatCurrency(incentiveValue)}</td>
                <td>
                  <select class="incentive-quantity" data-incentive-id="${id}" aria-label="Quantidade vendida de ${escapeHtml(name)}">
                    ${buildQuantityOptions(quantity)}
                  </select>
                </td>
                <td>${formatCurrency(subtotal)}</td>
              </tr>
            `;
          })
          .join("");

        return `
          <div class="support-block incentive-block incentive-tone-${section.tone}">
            <div class="incentive-section-header">
              <p class="support-title">${escapeHtml(section.title)}</p>
            </div>
            <div class="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>Modelo</th>
                    <th>Incentivo</th>
                    <th>Qtd.</th>
                    <th>Total</th>
                  </tr>
                </thead>
                <tbody>${rows}</tbody>
              </table>
            </div>
          </div>
        `;
      })
      .join("");

    return `
      <section class="incentive-brand ${selectedIncentiveBrand === brandId ? "active" : ""}">
        <div class="section-heading compact">
          <div>
            <p class="eyebrow">Marca</p>
            <h3>${escapeHtml(brandLabel)}</h3>
          </div>
        </div>
        ${sectionsMarkup}
      </section>
    `;
  }).join("");

  incentivesTotalValue.textContent = formatCurrency(totals.value);
  incentivesTotalQuantity.textContent = String(totals.quantity);
  incentivesActiveModels.textContent = String(totals.models);
}

function editBillingSale(index) {
  const sale = state.billingSales[index];
  if (!sale) return;

  editingBillingIndex = index;
  fillForm(billingForm, sale);
  billingForm.elements.namedItem("notas").value = sale.notas === "-" ? "" : sale.notas;
  netValueInput.value = formatPlainCurrency(Number(sale.valorSemIva || 0));
  billingForm.elements.namedItem("valorSeguro").value = sale.valorSeguro
    ? String(Number(sale.valorSeguro).toFixed(2))
    : "";
  insuranceAnnualValueInput.value = sale.valorSeguro
    ? formatPlainCurrency(Number(sale.valorSeguroAnual || 0))
    : "";
  activateTab("faturacao");
}

function deleteBillingSale(index) {
  state.billingSales.splice(index, 1);
  if (editingBillingIndex === index) {
    editingBillingIndex = null;
    billingForm.reset();
    netValueInput.value = "";
    insuranceAnnualValueInput.value = "";
    syncFormDatesToMonth();
  } else if (editingBillingIndex !== null && index < editingBillingIndex) {
    editingBillingIndex -= 1;
  }
}

function editMobileSale(index) {
  const sale = state.mobileSales[index];
  if (!sale) return;

  editingMobileIndex = index;
  fillForm(mobileForm, sale);
  activateTab("movel");
}

function deleteMobileSale(index) {
  state.mobileSales.splice(index, 1);
  if (editingMobileIndex === index) {
    editingMobileIndex = null;
    mobileForm.reset();
    syncFormDatesToMonth();
  } else if (editingMobileIndex !== null && index < editingMobileIndex) {
    editingMobileIndex -= 1;
  }
}

function editEnergySale(index) {
  const sale = state.energySales[index];
  if (!sale) return;

  editingEnergyIndex = index;
  fillForm(energyForm, sale);
  energyForm.elements.namedItem("observacoes").value =
    sale.observacoes === "-" ? "" : sale.observacoes;
  activateTab("meo-energia");
}

function editTask(index) {
  const task = state.tasks[index];
  if (!task) return;

  editingTaskIndex = index;
  fillForm(tasksForm, task);
  tasksForm.elements.namedItem("notes").value = task.notes === "-" ? "" : task.notes;
  activateTab("tarefas");
}

function editStockEntry(index) {
  const entry = state.stockEntries[index];
  if (!entry) return;

  editingStockIndex = index;
  fillForm(stockForm, entry);
  activateTab("stock");
}

function deleteEnergySale(index) {
  state.energySales.splice(index, 1);
  if (editingEnergyIndex === index) {
    editingEnergyIndex = null;
    energyForm.reset();
    syncFormDatesToMonth();
  } else if (editingEnergyIndex !== null && index < editingEnergyIndex) {
    editingEnergyIndex -= 1;
  }
}

function deleteStockEntry(index) {
  state.stockEntries.splice(index, 1);
  if (editingStockIndex === index) {
    editingStockIndex = null;
    stockForm.reset();
    syncFormDatesToMonth();
  } else if (editingStockIndex !== null && index < editingStockIndex) {
    editingStockIndex -= 1;
  }
}

function fillForm(form, values) {
  Object.entries(values).forEach(([key, value]) => {
    const field = form.elements.namedItem(key);
    if (field instanceof HTMLInputElement || field instanceof HTMLTextAreaElement || field instanceof HTMLSelectElement) {
      field.value = value ?? "";
    }
  });
}

function activateTab(targetId) {
  tabs.forEach((item) => item.classList.toggle("active", item.dataset.tab === targetId));
  panels.forEach((panel) => panel.classList.toggle("active", panel.id === targetId));
  applyVisualSettings();
}

function getReportEntry(summaryByDate, date) {
  const key = date || "sem-data";

  if (!summaryByDate.has(key)) {
    summaryByDate.set(key, {
      totalSemIva: 0,
      smartphone: 0,
      diversidade: 0,
      seguros: 0,
      movel: 0,
      energia: 0,
    });
  }

  return summaryByDate.get(key);
}

function hydrateGoalsForm() {
  const goals = getCurrentGoals();
  goalsForm.elements.namedItem("billingGoal").value = goals.billing || "";
  goalsForm.elements.namedItem("mobileGoal").value = goals.mobile || "";
  goalsForm.elements.namedItem("energyGoal").value = goals.energy || "";
}

function getCurrentGoals() {
  return (
    state.monthlyGoals[selectedMonth] || {
      billing: 0,
      mobile: 0,
      energy: 0,
    }
  );
}

function getCurrentIncentives() {
  const incentives = state.monthlyIncentives[selectedMonth];
  if (incentives && typeof incentives === "object" && !Array.isArray(incentives)) {
    return incentives;
  }
  state.monthlyIncentives[selectedMonth] = {};
  return state.monthlyIncentives[selectedMonth];
}

function getGoalMetrics(month) {
  const goals = state.monthlyGoals[month] || {
    billing: 0,
    mobile: 0,
    energy: 0,
  };
  const monthlyBillingSales = state.billingSales.filter((sale) => belongsToMonth(sale.data, month));
  const actualBilling = monthlyBillingSales.reduce(
    (sum, sale) => sum + getBillingEffectiveValue(sale),
    0
  );
  const actualMobile = state.mobileSales.filter((sale) => belongsToMonth(sale.data, month)).length;
  const actualEnergy = state.energySales.filter((sale) => belongsToMonth(sale.data, month)).length;
  const smartphoneCount = monthlyBillingSales.filter((sale) => sale.categoria === "Smartphone").length;
  const diversidadeCount = monthlyBillingSales.filter((sale) => sale.categoria === "Diversidade").length;
  const smartphoneWithInsurance = monthlyBillingSales.filter(
    (sale) => sale.categoria === "Smartphone" && sale.seguro === "Sim"
  ).length;

  const insuranceActualRatio = smartphoneCount > 0 ? smartphoneWithInsurance / smartphoneCount : 0;
  const insuranceGoalRatio = 0.15;
  const insuranceProgressRatio = insuranceGoalRatio > 0
    ? Math.min(insuranceActualRatio / insuranceGoalRatio, 1)
    : 0;
  const diversityActualRatio = smartphoneCount > 0 ? diversidadeCount / smartphoneCount : 0;
  const diversityGoalRatio = 1.6;
  const diversityProgressRatio = diversityGoalRatio > 0
    ? Math.min(diversityActualRatio / diversityGoalRatio, 1)
    : 0;

  const billingRatio = computeGoalRatio(actualBilling, goals.billing);
  const mobileRatio = computeGoalRatio(actualMobile, goals.mobile);
  const energyRatio = computeGoalRatio(actualEnergy, goals.energy);

  return {
    goals,
    actualBilling,
    actualMobile,
    actualEnergy,
    smartphoneCount,
    diversidadeCount,
    smartphoneWithInsurance,
    insuranceActualRatio,
    insuranceGoalRatio,
    insuranceProgressRatio,
    diversityActualRatio,
    diversityGoalRatio,
    diversityProgressRatio,
    billingRatio,
    mobileRatio,
    energyRatio,
    billingWeighted: billingRatio * 50,
    mobileWeighted: mobileRatio * 30,
    energyWeighted: energyRatio * 20,
    totalWeighted: billingRatio * 50 + mobileRatio * 30 + energyRatio * 20,
  };
}

function getTrackedMonths() {
  const [selectedYear] = selectedMonth.split("-");

  return Array.from({ length: 12 }, (_item, index) => {
    const month = String(index + 1).padStart(2, "0");
    return `${selectedYear}-${month}`;
  });
}

function getCurrentMonth() {
  return new Date().toISOString().slice(0, 7);
}

function isInSelectedMonth(date) {
  return belongsToMonth(date, selectedMonth);
}

function belongsToMonth(date, month) {
  return Boolean(date) && date.startsWith(month);
}

function updateMonthLabel() {
  const [year, month] = selectedMonth.split("-");
  const date = new Date(Number(year), Number(month) - 1, 1);
  selectedMonthLabel.textContent = formatMonth(selectedMonth);
}

function syncFormDatesToMonth() {
  [billingForm, mobileForm, energyForm, stockForm].forEach((form) => {
    const dateInput =
      form === stockForm ? form.elements.namedItem("dataPedido") : form.elements.namedItem("data");
    if (!(dateInput instanceof HTMLInputElement)) return;

    dateInput.min = `${selectedMonth}-01`;
    dateInput.max = `${selectedMonth}-${getLastDayOfMonth(selectedMonth)}`;

    if (!dateInput.value || !dateInput.value.startsWith(selectedMonth)) {
      dateInput.value = `${selectedMonth}-01`;
    }
  });
}

function getLastDayOfMonth(month) {
  const [year, monthNumber] = month.split("-").map(Number);
  return String(new Date(year, monthNumber, 0).getDate()).padStart(2, "0");
}

function getTodayForSelectedMonth() {
  const todayIso = new Date().toISOString().slice(0, 10);
  return todayIso.startsWith(selectedMonth) ? todayIso : `${selectedMonth}-01`;
}

function applyTheme() {
  document.body.classList.toggle("theme-dark", state.settings.theme === "dark");
}

function hydrateVisualEditor() {
  menuFontSizeInput.value = String(state.settings.menuFontSize);
  titleFontSizeInput.value = String(state.settings.titleFontSize);
  fieldFontSizeInput.value = String(state.settings.fieldFontSize);
  panelHeaderSizeInput.value = String(state.settings.panelHeaderSize);
  densityModeInput.value = state.settings.densityMode;
  primaryColorInput.value = state.settings.primaryColor;
  secondaryColorInput.value = state.settings.secondaryColor;
  sidebarColorInput.value = state.settings.sidebarColor;
  toolbarColorInput.value = state.settings.toolbarColor;
  const calculatorWidth = clampCalculatorWidth(state.settings.calculatorWidth);
  state.settings.calculatorWidth = calculatorWidth;
  if (calculatorWidthInput) calculatorWidthInput.value = String(calculatorWidth);
  if (calculatorWidthValue) calculatorWidthValue.textContent = `${calculatorWidth}px`;
  Object.entries(sidebarTabColorInputs).forEach(([tabId, input]) => {
    if (input) {
      input.value = state.settings.sidebarTabColors?.[tabId] || DEFAULT_SIDEBAR_TAB_COLORS[tabId];
    }
  });
  Object.entries(pageColorInputs).forEach(([panelId, input]) => {
    input.value = state.settings.pageColors?.[panelId] || DEFAULT_PAGE_COLORS[panelId];
  });
  hydratePanelColorPickers();
  hydratePanelBlockPickers();
  hydratePanelFieldPickers();
  formWidthInput.value = String(state.settings.formWidth);
  cardRadiusInput.value = String(state.settings.cardRadius);
}

function hydratePanelColorPickers() {
  panelColorInputs.forEach((input) => {
    const panelId = input.dataset.panelColor;
    if (!panelId) return;
    input.value = state.settings.pageColors?.[panelId] || DEFAULT_PAGE_COLORS[panelId];
  });
}

function hydratePanelBlockPickers() {
  panelBlockInputs.forEach((input) => {
    const panelId = input.dataset.panelBlockColor;
    if (!panelId) return;
    input.value = state.settings.pageBlockColors?.[panelId] || DEFAULT_PAGE_BLOCK_COLORS[panelId];
  });
}

function hydratePanelFieldPickers() {
  panelFieldInputs.forEach((input) => {
    const panelId = input.dataset.panelFieldColor;
    if (!panelId) return;
    input.value = state.settings.pageFieldColors?.[panelId] || DEFAULT_PAGE_FIELD_COLORS[panelId];
  });
}

async function handleVisualSettingsChange() {
  state.settings.menuFontSize = Number(menuFontSizeInput.value);
  state.settings.titleFontSize = Number(titleFontSizeInput.value);
  state.settings.fieldFontSize = Number(fieldFontSizeInput.value);
  state.settings.panelHeaderSize = Number(panelHeaderSizeInput.value);
  state.settings.densityMode = densityModeInput.value;
  state.settings.primaryColor = primaryColorInput.value;
  state.settings.secondaryColor = secondaryColorInput.value;
  state.settings.sidebarColor = sidebarColorInput.value;
  state.settings.toolbarColor = toolbarColorInput.value;
  state.settings.sidebarTabColors = Object.fromEntries(
    Object.entries(sidebarTabColorInputs).map(([tabId, input]) => [tabId, input?.value || DEFAULT_SIDEBAR_TAB_COLORS[tabId]])
  );
  state.settings.pageColors = Object.fromEntries(
    Object.entries(pageColorInputs).map(([panelId, input]) => [panelId, input.value])
  );
  state.settings.calculatorWidth = clampCalculatorWidth(calculatorWidthInput?.value || 340);
  state.settings.formWidth = Number(formWidthInput.value);
  state.settings.cardRadius = Number(cardRadiusInput.value);
  applyVisualSettings();
  await persistState();
}

async function handlePanelColorChange(event) {
  const target = event.target;
  if (!(target instanceof HTMLInputElement)) return;
  const panelId = target.dataset.panelColor;
  if (!panelId) return;

  state.settings.pageColors = {
    ...state.settings.pageColors,
    [panelId]: target.value,
  };

  document
    .querySelectorAll(`[data-panel-color="${panelId}"]`)
    .forEach((input) => {
      if (input !== target && input instanceof HTMLInputElement) {
        input.value = target.value;
      }
    });

  applyVisualSettings();
  await persistState();
}

async function handlePanelBlockChange(event) {
  const target = event.target;
  if (!(target instanceof HTMLInputElement)) return;
  const panelId = target.dataset.panelBlockColor;
  if (!panelId) return;

  state.settings.pageBlockColors = {
    ...state.settings.pageBlockColors,
    [panelId]: target.value,
  };

  document
    .querySelectorAll(`[data-panel-block-color="${panelId}"]`)
    .forEach((input) => {
      if (input !== target && input instanceof HTMLInputElement) {
        input.value = target.value;
      }
    });

  applyVisualSettings();
  await persistState();
}

async function handlePanelFieldChange(event) {
  const target = event.target;
  if (!(target instanceof HTMLInputElement)) return;
  const panelId = target.dataset.panelFieldColor;
  if (!panelId) return;

  state.settings.pageFieldColors = {
    ...state.settings.pageFieldColors,
    [panelId]: target.value,
  };

  applyVisualSettings();
  await persistState();
}

async function handleSidebarTabColorChange(event) {
  const target = event.target;
  if (!(target instanceof HTMLInputElement)) return;
  const tabId = target.dataset.sidebarTabColor;
  if (!tabId) return;

  state.settings.sidebarTabColors = {
    ...state.settings.sidebarTabColors,
    [tabId]: target.value,
  };

  document.querySelectorAll(`[data-sidebar-tab-color="${tabId}"]`).forEach((input) => {
    if (input !== target && input instanceof HTMLInputElement) {
      input.value = target.value;
    }
  });

  applyVisualSettings();
  await persistState();
}

async function handleCalculatorWidthChange(event) {
  const target = event.target;
  if (!(target instanceof HTMLInputElement)) return;

  state.settings.calculatorWidth = clampCalculatorWidth(target.value);
  if (calculatorWidthValue) calculatorWidthValue.textContent = `${state.settings.calculatorWidth}px`;

  applyVisualSettings();
  await persistState();
}

function applyVisualSettings() {
  const densityPadding = state.settings.densityMode === "normal" ? 24 : 18;
  document.documentElement.style.setProperty("--menu-font-size", `${state.settings.menuFontSize}px`);
  document.documentElement.style.setProperty("--title-font-size", `${state.settings.titleFontSize}px`);
  document.documentElement.style.setProperty("--field-font-size", `${state.settings.fieldFontSize}px`);
  document.documentElement.style.setProperty("--panel-header-size", `${state.settings.panelHeaderSize}px`);
  document.documentElement.style.setProperty(
    "--panel-value-size",
    `${Math.max(state.settings.panelHeaderSize + 8, 18)}px`
  );
  document.documentElement.style.setProperty("--accent", state.settings.primaryColor);
  document.documentElement.style.setProperty("--accent-dark", shadeColor(state.settings.primaryColor, -18));
  document.documentElement.style.setProperty("--surface-strong", state.settings.secondaryColor);
  document.documentElement.style.setProperty("--accent-soft", state.settings.secondaryColor);
  document.documentElement.style.setProperty("--sidebar-bg", state.settings.sidebarColor);
  document.documentElement.style.setProperty("--month-toolbar-bg", state.settings.toolbarColor);
  document.documentElement.style.setProperty("--month-toolbar-text", getReadableTextColor(state.settings.toolbarColor));
  const calculatorWidth = clampCalculatorWidth(state.settings.calculatorWidth);
  state.settings.calculatorWidth = calculatorWidth;
  document.documentElement.style.setProperty("--calculator-width", `${calculatorWidth}px`);
  tabs.forEach((tab) => {
    const tabId = tab.dataset.tab;
    if (!tabId) return;
    const tabColor = state.settings.sidebarTabColors?.[tabId] || DEFAULT_SIDEBAR_TAB_COLORS[tabId] || state.settings.primaryColor;
    tab.style.color = tab.classList.contains("active") ? "#ffffff" : tabColor;
  });
  Object.entries(DEFAULT_PAGE_COLORS).forEach(([panelId, fallbackColor]) => {
    const pageColor = state.settings.pageColors?.[panelId] || fallbackColor;
    const blockColor = state.settings.pageBlockColors?.[panelId] || DEFAULT_PAGE_BLOCK_COLORS[panelId];
    const fieldColor = state.settings.pageFieldColors?.[panelId] || DEFAULT_PAGE_FIELD_COLORS[panelId];
    const panel = document.getElementById(panelId);
    if (panel) {
      panel.style.background = pageColor;
      panel.style.setProperty("--panel-block-bg", blockColor);
      panel.style.setProperty("--panel-block-text", getReadableTextColor(blockColor));
      panel.style.setProperty("--panel-block-muted", getReadableTextColorWithAlpha(blockColor, 0.7));
      panel.style.setProperty("--panel-field-bg", fieldColor);
      panel.style.setProperty("--panel-field-text", getReadableTextColor(fieldColor));
    }
  });
  document.documentElement.style.setProperty("--form-max-width", `${state.settings.formWidth}px`);
  document.documentElement.style.setProperty("--card-radius", `${state.settings.cardRadius}px`);
  document.documentElement.style.setProperty("--density-padding", `${densityPadding}px`);
}

function shadeColor(hex, percent) {
  const value = hex.replace("#", "");
  const num = Number.parseInt(value, 16);
  const amount = Math.round(2.55 * percent);
  const r = Math.min(255, Math.max(0, (num >> 16) + amount));
  const g = Math.min(255, Math.max(0, ((num >> 8) & 0x00ff) + amount));
  const b = Math.min(255, Math.max(0, (num & 0x0000ff) + amount));
  return `#${((1 << 24) | (r << 16) | (g << 8) | b).toString(16).slice(1)}`;
}

function getReadableTextColor(hex) {
  const value = hex.replace("#", "");
  const num = Number.parseInt(value, 16);
  const r = (num >> 16) & 0xff;
  const g = (num >> 8) & 0xff;
  const b = num & 0xff;
  const luminance = (r * 299 + g * 587 + b * 114) / 1000;
  return luminance > 150 ? "#08112f" : "#f3f8ff";
}

function getReadableTextColorWithAlpha(hex, alpha) {
  const textColor = getReadableTextColor(hex).replace("#", "");
  const num = Number.parseInt(textColor, 16);
  const r = (num >> 16) & 0xff;
  const g = (num >> 8) & 0xff;
  const b = num & 0xff;
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

function computeGoalRatio(actual, target) {
  if (target <= 0) return 0;
  return actual / target;
}

function buildQuantityOptions(selectedValue) {
  const options = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 15, 20, 25, 30];
  return options
    .map(
      (value) =>
        `<option value="${value}" ${value === selectedValue ? "selected" : ""}>${value}</option>`
    )
    .join("");
}

function formatPercent(value) {
  return `${value.toFixed(1)}%`;
}

function formatCurrency(value) {
  return new Intl.NumberFormat("pt-PT", {
    style: "currency",
    currency: "EUR",
  }).format(value);
}

function formatDate(value) {
  if (!value || value === "sem-data") return "-";
  return new Intl.DateTimeFormat("pt-PT").format(new Date(value));
}

function formatMonth(value) {
  if (!value) return "-";
  const [year, month] = value.split("-");
  return new Intl.DateTimeFormat("pt-PT", {
    month: "long",
    year: "numeric",
  })
    .format(new Date(Number(year), Number(month) - 1, 1))
    .toUpperCase();
}

function calculateNetValue(value) {
  return Number((value / 1.23).toFixed(2));
}

function calculateInsuranceAnnual(value) {
  return Number((value * 12).toFixed(2));
}

function formatPlainCurrency(value) {
  return value.toFixed(2);
}

function setAuthMessage(message = "", isError = false) {
  if (!authSettingsMessage) return;
  authSettingsMessage.textContent = message;
  authSettingsMessage.classList.toggle("is-error", isError);
}

function getAuthToken() {
  try {
    return window.localStorage.getItem(AUTH_TOKEN_KEY) || "";
  } catch (error) {
    return "";
  }
}

function setAuthToken(token) {
  try {
    window.localStorage.setItem(AUTH_TOKEN_KEY, token);
  } catch (error) {
    // Keep cookie-based auth as fallback.
  }
}

function clearAuthToken() {
  try {
    window.localStorage.removeItem(AUTH_TOKEN_KEY);
  } catch (error) {
    // Ignore storage failures.
  }
}

function buildAuthHeaders(extraHeaders = {}) {
  const headers = { ...extraHeaders };
  const token = getAuthToken();
  if (token) {
    headers.Authorization = `Bearer ${token}`;
  }
  return headers;
}

function getBillingEffectiveValue(sale) {
  return Number(sale.valorSemIva || 0) + Number(sale.valorSeguroAnual || 0);
}

function valueAsText(value) {
  return value?.toString().trim() || "";
}

function parseMoneyInput(value) {
  const normalized = String(value || "")
    .trim()
    .replace(/\s/g, "")
    .replace(",", ".");

  if (!normalized) return 0;

  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
}

function createIncentiveId(brand, section, name) {
  return `${brand}-${section}-${name}`
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-|-$/g, "");
}

function slugify(value) {
  return String(value)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-|-$/g, "");
}

function emptyRow(colspan, text) {
  return `
    <tr class="empty-row">
      <td colspan="${colspan}">${text}</td>
    </tr>
  `;
}

function updateCalculatorDisplay() {
  if (!calculatorDisplay) return;
  calculatorDisplay.value = formatCalculatorDisplayValue(calculatorExpression);
}

function formatCalculatorDisplayValue(value) {
  return String(value || "0").replaceAll(".", ",");
}

function clearCalculator() {
  calculatorExpression = "0";
  calculatorJustEvaluated = false;
  updateCalculatorDisplay();
}

function backspaceCalculator() {
  if (calculatorExpression === "Erro" || calculatorJustEvaluated) {
    calculatorExpression = "0";
    calculatorJustEvaluated = false;
    updateCalculatorDisplay();
    return;
  }

  calculatorExpression = calculatorExpression.slice(0, -1) || "0";
  if (calculatorExpression === "-" || calculatorExpression === "") {
    calculatorExpression = "0";
  }
  updateCalculatorDisplay();
}

function appendCalculatorValue(value) {
  if (!value) return;

  if (calculatorExpression === "Erro" || calculatorJustEvaluated) {
    if (/[0-9.]/.test(value)) {
      calculatorExpression = value === "." ? "0." : value;
      calculatorJustEvaluated = false;
      updateCalculatorDisplay();
      return;
    }

    calculatorJustEvaluated = false;
  }

  if (calculatorExpression === "0") {
    if (/[0-9]/.test(value)) {
      calculatorExpression = value;
      updateCalculatorDisplay();
      return;
    }

    if (value === ".") {
      calculatorExpression = "0.";
      updateCalculatorDisplay();
      return;
    }

    if (value === "-") {
      calculatorExpression = "-";
      updateCalculatorDisplay();
      return;
    }

    if (value === "%") {
      return;
    }
  }

  const lastChar = calculatorExpression.slice(-1);
  if (/[+\-*/]/.test(lastChar) && /[+\-*/]/.test(value)) {
    calculatorExpression = `${calculatorExpression.slice(0, -1)}${value}`;
    updateCalculatorDisplay();
    return;
  }

  if (value === ".") {
    const segments = calculatorExpression.split(/[+\-*/]/);
    const currentSegment = segments[segments.length - 1] || "";
    if (currentSegment.includes(".")) return;
    if (/[+\-*/]$/.test(calculatorExpression)) {
      calculatorExpression += "0.";
    } else {
      calculatorExpression += ".";
    }
    updateCalculatorDisplay();
    return;
  }

  if (value === "%") {
    if (!/[0-9)%]$/.test(calculatorExpression)) return;
    calculatorExpression += "%";
    updateCalculatorDisplay();
    return;
  }

  calculatorExpression += value;
  updateCalculatorDisplay();
}

function evaluateCalculator() {
  let expression = calculatorExpression.trim();
  if (!expression || expression === "Erro") {
    clearCalculator();
    return;
  }

  while (/[+\-*/.]$/.test(expression)) {
    expression = expression.slice(0, -1);
  }

  if (!expression) {
    clearCalculator();
    return;
  }

  try {
    const result = calculateCalculatorExpression(expression);
    if (!Number.isFinite(result)) throw new Error("Resultado inválido.");

    calculatorExpression = formatCalculatorResult(result);
    calculatorJustEvaluated = true;
    updateCalculatorDisplay();
  } catch {
    calculatorExpression = "Erro";
    calculatorJustEvaluated = false;
    updateCalculatorDisplay();
  }
}

function calculateCalculatorExpression(rawExpression) {
  const tokens = tokenizeCalculatorExpression(rawExpression);
  let index = 0;

  function peek() {
    return tokens[index] || null;
  }

  function consume() {
    return tokens[index++] || null;
  }

  function isUnaryPosition() {
    const previous = tokens[index - 1];
    return (
      !previous ||
      previous.type === "operator" ||
      (previous.type === "paren" && previous.value === "(")
    );
  }

  function parseExpression() {
    let left = parseTerm();
    if (left.percent) {
      left.value /= 100;
      left.percent = false;
    }

    while (peek() && peek().type === "operator" && (peek().value === "+" || peek().value === "-")) {
      const operator = consume().value;
      const right = parseTerm();
      const rightValue = right.percent ? (left.value * right.value) / 100 : right.value;
      left = {
        value: operator === "+" ? left.value + rightValue : left.value - rightValue,
        percent: false,
      };
    }

    return left;
  }

  function parseTerm() {
    let left = parseFactor();

    if (left.percent && peek() && peek().type === "operator" && (peek().value === "*" || peek().value === "/")) {
      left.value /= 100;
      left.percent = false;
    }

    while (peek() && peek().type === "operator" && (peek().value === "*" || peek().value === "/")) {
      const operator = consume().value;
      const right = parseFactor();
      const rightValue = right.percent ? right.value / 100 : right.value;

      left = {
        value: operator === "*" ? left.value * rightValue : left.value / rightValue,
        percent: false,
      };
    }

    return left;
  }

  function parseFactor() {
    let sign = 1;
    while (peek() && peek().type === "operator" && (peek().value === "+" || peek().value === "-") && isUnaryPosition()) {
      if (consume().value === "-") sign *= -1;
    }

    const token = peek();
    if (!token) throw new Error("Expressão incompleta.");

    let result;
    if (token.type === "number") {
      consume();
      result = { value: Number(token.value), percent: false };
    } else if (token.type === "paren" && token.value === "(") {
      consume();
      result = parseExpression();
      const closing = consume();
      if (!closing || closing.type !== "paren" || closing.value !== ")") {
        throw new Error("Parênteses inválidos.");
      }
    } else {
      throw new Error("Expressão inválida.");
    }

    if (peek() && peek().type === "percent") {
      consume();
      result.percent = true;
    }

    result.value *= sign;
    return result;
  }

  const parsed = parseExpression();
  if (index < tokens.length) {
    throw new Error("Expressão inválida.");
  }

  return parsed.value;
}

function tokenizeCalculatorExpression(expression) {
  const tokens = [];
  let index = 0;
  const normalized = String(expression).replaceAll(",", ".");

  while (index < normalized.length) {
    const character = normalized[index];

    if (/\s/.test(character)) {
      index += 1;
      continue;
    }

    if (/[0-9.]/.test(character)) {
      let number = character;
      index += 1;
      while (index < normalized.length && /[0-9.]/.test(normalized[index])) {
        number += normalized[index];
        index += 1;
      }

      if ((number.match(/\./g) || []).length > 1 || number === ".") {
        throw new Error("Número inválido.");
      }

      tokens.push({ type: "number", value: number });
      continue;
    }

    if ("+-*/".includes(character)) {
      tokens.push({ type: "operator", value: character });
      index += 1;
      continue;
    }

    if (character === "(" || character === ")") {
      tokens.push({ type: "paren", value: character });
      index += 1;
      continue;
    }

    if (character === "%") {
      tokens.push({ type: "percent", value: character });
      index += 1;
      continue;
    }

    throw new Error("Caractere inválido.");
  }

  return tokens;
}

function formatCalculatorResult(value) {
  const normalized = Number.parseFloat(Number(value).toFixed(12));
  if (Number.isInteger(normalized)) {
    return String(normalized);
  }

  return String(normalized).replace(/\.0+$/, "").replace(/(\.\d*?)0+$/, "$1");
}

async function copyTextToClipboard(value) {
  if (navigator.clipboard?.writeText) {
    try {
      await navigator.clipboard.writeText(value);
      return;
    } catch {
      // Fallback below.
    }
  }

  const helper = document.createElement("textarea");
  helper.value = value;
  helper.setAttribute("readonly", "true");
  helper.style.position = "fixed";
  helper.style.top = "-9999px";
  helper.style.left = "-9999px";
  helper.style.opacity = "0";
  document.body.appendChild(helper);
  helper.select();
  helper.setSelectionRange(0, helper.value.length);

  const copied = document.execCommand("copy");
  document.body.removeChild(helper);

  if (!copied) {
    throw new Error("Falha ao copiar para a área de transferência.");
  }
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

bootstrap();
