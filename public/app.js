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
const mobileTable = document.querySelector(".mobile-table");
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
const sidebarSearchInput = document.getElementById("sidebar-search");
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
const sidebarButtonColorInput = document.getElementById("sidebar-button-color");
const sidebarButtonStyleInput = document.getElementById("sidebar-button-style");
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
  { id: "sap-comercial", label: "SAP Comercial", color: "#8fd7ff" },
  { id: "portal-sfa", label: "Portal SFA", color: "#7ee081" },
  { id: "user-rede", label: "User de Rede", color: "#ff9f43" },
  { id: "pos-venda", label: "Pós-Venda", color: "#ff6b6b" },
  { id: "blueticket", label: "Blueticket", color: "#d68bff" },
  { id: "intelcia", label: "Intelcia", color: "#7db7ff" },
  { id: "eform", label: "EForm", color: "#f8d84c" },
  { id: "bec-deposito", label: "BEC Depósito", color: "#5ef0b2" },
];

const DEFAULT_SIDEBAR_TAB_COLORS = {
  relatorios: "#f8d84c",
  faturacao: "#4fd3ff",
  movel: "#ff6a9a",
  "meo-energia": "#77f25f",
  objetivo: "#7c7cff",
  geral: "#7fe0ff",
  tarefas: "#5ef0b2",
  apoio: "#ff9f43",
  incentivos: "#d68bff",
  stock: "#a3ff5c",
  acessos: "#ff7d7d",
  configuracoes: "#f77dd6",
};

const DEFAULT_MOBILE_COLUMN_WIDTHS = {
  data: 60,
  nome: 126,
  nif: 66,
  numero: 74,
  tarifario: 118,
  idVenda: 64,
  fidelizacao: 58,
  portabilidade: 58,
  bestOffer: 58,
  netSegura: 62,
  rceAceite: 58,
  actions: 88,
};

const MOBILE_COLUMN_KEYS = [
  "data",
  "nome",
  "nif",
  "numero",
  "tarifario",
  "idVenda",
  "fidelizacao",
  "portabilidade",
  "bestOffer",
  "netSegura",
  "rceAceite",
  "actions",
];

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
const mobileTotalSales = document.getElementById("mobile-total-sales");
const energyTotalSales = document.getElementById("energy-total-sales");
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
const goalMissingBilling = document.getElementById("goal-missing-billing");
const goalWeightedBilling = document.getElementById("goal-weighted-billing");
const goalActualMobile = document.getElementById("goal-actual-mobile");
const goalTargetMobile = document.getElementById("goal-target-mobile");
const goalPercentMobile = document.getElementById("goal-percent-mobile");
const goalMissingMobile = document.getElementById("goal-missing-mobile");
const goalWeightedMobile = document.getElementById("goal-weighted-mobile");
const goalActualEnergy = document.getElementById("goal-actual-energy");
const goalTargetEnergy = document.getElementById("goal-target-energy");
const goalPercentEnergy = document.getElementById("goal-percent-energy");
const goalMissingEnergy = document.getElementById("goal-missing-energy");
const goalWeightedEnergy = document.getElementById("goal-weighted-energy");
const goalActualInsurance = document.getElementById("goal-actual-insurance");
const goalTargetInsurance = document.getElementById("goal-target-insurance");
const goalPercentInsurance = document.getElementById("goal-percent-insurance");
const goalMissingInsurance = document.getElementById("goal-missing-insurance");
const goalWeightedInsurance = document.getElementById("goal-weighted-insurance");
const goalActualDiversity = document.getElementById("goal-actual-diversity");
const goalTargetDiversity = document.getElementById("goal-target-diversity");
const goalPercentDiversity = document.getElementById("goal-percent-diversity");
const goalMissingDiversity = document.getElementById("goal-missing-diversity");
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
  sidebarButtonColor: "#20316f",
  sidebarButtonStyle: "style-1",
  toolbarColor: "#b3262d",
    sidebarTabColors: { ...DEFAULT_SIDEBAR_TAB_COLORS },
    mobileColumnWidths: { ...DEFAULT_MOBILE_COLUMN_WIDTHS },
    pageColors: { ...DEFAULT_PAGE_COLORS },
    pageBlockColors: { ...DEFAULT_PAGE_BLOCK_COLORS },
    pageFieldColors: { ...DEFAULT_PAGE_FIELD_COLORS },
    calculatorWidth: 340,
    formWidth: 720,
    cardRadius: 24,
  },
};

const BILLING_FILTER_FIELDS = {
  data: {
    label: "Data",
  },
  seguro: {
    label: "Seguro",
  },
  categoria: {
    label: "Categoria",
  },
  modalidade: {
    label: "Modalidade",
  },
};

let editingBillingIndex = null;
let billingInlineEditingIndex = null;
let mobileInlineEditingIndex = null;
let editingMobileIndex = null;
let editingEnergyIndex = null;
let energyInlineEditingIndex = null;
let editingTaskIndex = null;
let editingStockIndex = null;
let selectedMonth = getCurrentMonth();
let mobileColumnResizeState = null;
let selectedIncentiveBrand = "samsung";
let selectedAccessService = ACCESS_SERVICES[0].id;
let accessPasswordVisible = false;
let accessStatusMessage = "";
let accessStatusError = false;
let calculatorExpression = "0";
let calculatorJustEvaluated = false;
let searchNavigationActive = false;
let sidebarSearchState = {
  query: "",
  matches: [],
  index: -1,
  activeKey: "",
};
let reportDateFilter = "";
let reportDateFilterOpen = false;
let billingFilters = {
  data: "",
  seguro: "",
  categoria: "",
  modalidade: "",
};
let billingFilterOpenKey = "";
let mobileFilters = {
  data: "",
  tarifario: "",
  fidelizacao: "",
  portabilidade: "",
  bestOffer: "",
  netSegura: "",
  rceAceite: "",
};
let mobileFilterOpenKey = "";

function getSelectedMonthBillingSales() {
  return state.billingSales
    .filter((sale) => isInSelectedMonth(sale.data))
    .sort((a, b) => compareSaleDates(a.data, b.data));
}

function getReportSummaryDates() {
  const dates = new Set();

  state.billingSales.forEach((sale) => {
    if (isInSelectedMonth(sale.data) && sale.data) dates.add(sale.data);
  });
  state.mobileSales.forEach((sale) => {
    if (isInSelectedMonth(sale.data) && sale.data) dates.add(sale.data);
  });
  state.energySales.forEach((sale) => {
    if (isInSelectedMonth(sale.data) && sale.data) dates.add(sale.data);
  });

  return [...dates].sort();
}

function renderReportDateMenu(dates) {
  const menu = document.querySelector('[data-report-filter-menu="data"]');
  const trigger = document.querySelector('[data-report-filter-trigger="data"]');
  if (!(menu instanceof HTMLElement) || !(trigger instanceof HTMLElement)) return;

  trigger.classList.toggle("is-filtered", Boolean(reportDateFilter));
  trigger.setAttribute("aria-expanded", reportDateFilterOpen ? "true" : "false");

  const options = [
    `
      <button type="button" class="report-filter-option ${reportDateFilter === "" ? "active" : ""}" data-report-filter-choice="">
        Todos
      </button>
    `,
    ...(dates.length > 0
      ? dates.map(
          (value) => `
            <button type="button" class="report-filter-option ${reportDateFilter === value ? "active" : ""}" data-report-filter-choice="${escapeHtml(value)}">
              ${escapeHtml(formatDate(value))}
            </button>
          `
        )
      : [`<div class="report-filter-empty">Sem datas</div>`]),
  ];

  menu.innerHTML = options.join("");
  menu.classList.toggle("open", reportDateFilterOpen);
  menu.setAttribute("aria-hidden", reportDateFilterOpen ? "false" : "true");

  if (reportDateFilterOpen) {
    positionBillingFilterMenu(menu, trigger);
  } else {
    menu.style.left = "";
    menu.style.top = "";
    menu.style.transform = "";
    menu.style.maxHeight = "";
    menu.style.position = "";
    menu.style.zIndex = "";
  }
}

function refreshReportDateMenu() {
  renderReportDateMenu(getReportSummaryDates());
}

function toggleReportDateMenu() {
  reportDateFilterOpen = reportDateFilterOpen ? false : true;
  refreshReportDateMenu();
}

function closeReportDateMenu() {
  if (!reportDateFilterOpen) return;
  reportDateFilterOpen = false;
  refreshReportDateMenu();
}

function positionBillingFilterMenu(menu, trigger) {
  const rect = trigger.getBoundingClientRect();
  const viewportWidth = window.innerWidth;
  const viewportHeight = window.innerHeight;
  const menuWidth = Math.min(Math.max(menu.offsetWidth || 176, 176), viewportWidth - 16);
  const anchorX = rect.left + rect.width / 2;
  const left = Math.min(Math.max(anchorX, menuWidth / 2 + 8), viewportWidth - menuWidth / 2 - 8);
  const spaceBelow = viewportHeight - rect.bottom;
  const spaceAbove = rect.top;
  const openUp = spaceBelow < 260 && spaceAbove > spaceBelow;
  const top = openUp ? rect.top - 8 : rect.bottom + 8;
  const maxHeight = Math.max(160, Math.min(280, openUp ? spaceAbove - 16 : spaceBelow - 16));

  if (menu.parentElement !== document.body) {
    document.body.appendChild(menu);
  }
  menu.style.position = "fixed";
  menu.style.left = `${left}px`;
  menu.style.top = `${top}px`;
  menu.style.maxHeight = `${maxHeight}px`;
  menu.style.transform = openUp ? "translate(-50%, -100%)" : "translate(-50%, 0)";
  menu.style.zIndex = "9999";
}

function refreshBillingFilterMenus() {
  renderBillingFilterMenus(getSelectedMonthBillingSales());
}

function clampCalculatorWidth(value) {
  return Math.min(460, Math.max(280, Number(value) || 360));
}

tabs.forEach((tab) => {
  tab.addEventListener("click", () => {
    if (!searchNavigationActive) clearSearchHighlights();
    activateTab(tab.dataset.tab);
  });
});

supportTabs.forEach((tab) => {
  tab.addEventListener("click", () => {
    if (tab.classList.contains("incentive-tab")) return;
    if (!searchNavigationActive) clearSearchHighlights();
    const target = tab.dataset.supportTab;
    supportTabs.forEach((item) => item.classList.toggle("active", item === tab));
    supportPanels.forEach((panel) =>
      panel.classList.toggle("active", panel.id === `support-${target}`)
    );
  });
});

incentiveTabs.forEach((tab) => {
  tab.addEventListener("click", () => {
    if (!searchNavigationActive) clearSearchHighlights();
    selectedIncentiveBrand = tab.dataset.incentiveTab || "samsung";
    incentiveTabs.forEach((item) => item.classList.toggle("active", item === tab));
    renderIncentives();
  });
});

sidebarSearchInput?.addEventListener("keydown", (event) => {
  if (event.key !== "Enter") return;
  event.preventDefault();
  performSidebarSearch(sidebarSearchInput.value);
});

accessButtons?.addEventListener("click", (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return;

  const button = target.closest("[data-access-service]");
  if (!(button instanceof HTMLElement)) return;

  const serviceId = button.dataset.accessService;
  if (!serviceId) return;

  if (!searchNavigationActive) clearSearchHighlights();
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

document.addEventListener("click", (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return;

  const reportFilterTrigger = target.closest("[data-report-filter-trigger]");
  if (reportFilterTrigger instanceof HTMLElement) {
    event.preventDefault();
    toggleReportDateMenu();
    return;
  }

  const reportFilterChoice = target.closest("[data-report-filter-choice]");
  if (reportFilterChoice instanceof HTMLElement) {
    reportDateFilter = reportFilterChoice.dataset.reportFilterChoice || "";
    reportDateFilterOpen = false;
    refreshReportDateMenu();
    renderReport();
    return;
  }

  if (!target.closest(".report-filter-th")) {
    closeReportDateMenu();
  }

  const mobileFilterTrigger = target.closest("[data-mobile-filter-trigger]");
  if (mobileFilterTrigger instanceof HTMLElement) {
    event.preventDefault();
    toggleMobileFilterMenu(mobileFilterTrigger.dataset.mobileFilterTrigger || "");
    return;
  }

  const mobileFilterChoice = target.closest("[data-mobile-filter-choice]");
  if (mobileFilterChoice instanceof HTMLElement) {
    const field = mobileFilterChoice.dataset.mobileFilterChoice || "";
    const value = mobileFilterChoice.dataset.mobileFilterValue || "";
    if (field in mobileFilters) {
      mobileFilters = {
        ...mobileFilters,
        [field]: value,
      };
    }
    closeMobileFilterMenu();
    renderMobileSales();
    return;
  }

  if (!target.closest(".mobile-filter-th")) {
    closeMobileFilterMenu();
  }

  const filterTrigger = target.closest("[data-billing-filter-trigger]");
  if (filterTrigger instanceof HTMLElement) {
    event.preventDefault();
    toggleBillingFilterMenu(filterTrigger.dataset.billingFilterTrigger || "");
    return;
  }

  const filterChoice = target.closest("[data-billing-filter-choice]");
  if (filterChoice instanceof HTMLElement) {
    const field = filterChoice.dataset.billingFilterValue || "";
    const value = filterChoice.dataset.billingFilterChoice || "";
    setBillingFilter(field, value);
    closeBillingFilterMenu();
    renderBillingSales();
    return;
  }

  if (!target.closest(".billing-filter-th")) {
    closeBillingFilterMenu();
  }
});

document.addEventListener("pointermove", updateMobileColumnResize);
document.addEventListener("pointerup", finishMobileColumnResize);
document.addEventListener("pointercancel", finishMobileColumnResize);

window.addEventListener("resize", () => {
  if (billingFilterOpenKey) {
    refreshBillingFilterMenus();
  }
});

window.addEventListener(
  "scroll",
  () => {
    if (billingFilterOpenKey) {
      refreshBillingFilterMenus();
    }
  },
  true
);

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
  closeBillingFilterMenu();
  closeReportDateMenu();
  closeMobileFilterMenu();
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
sidebarButtonColorInput?.addEventListener("input", handleVisualSettingsChange);
sidebarButtonStyleInput?.addEventListener("input", handleVisualSettingsChange);
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

  if (target.dataset.cancelBillingInline !== undefined) {
    cancelBillingInlineEdit(Number(target.dataset.cancelBillingInline));
  }

  if (target.dataset.deleteBilling !== undefined) {
    deleteBillingSale(Number(target.dataset.deleteBilling));
    await persistAndRender();
  }
});

billingTableBody.addEventListener("submit", async (event) => {
  const target = event.target;
  if (!(target instanceof HTMLFormElement)) return;

  if (target.matches("[data-billing-inline-form]")) {
    event.preventDefault();
    await saveBillingInlineEdit(target);
  }
});

mobileTableBody.addEventListener("click", async (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return;

  if (target.dataset.editMobile !== undefined) {
    editMobileSale(Number(target.dataset.editMobile));
  }

  if (target.dataset.cancelMobileInline !== undefined) {
    cancelMobileInlineEdit(Number(target.dataset.cancelMobileInline));
  }

  if (target.dataset.deleteMobile !== undefined) {
    deleteMobileSale(Number(target.dataset.deleteMobile));
    await persistAndRender();
  }
});

mobileTableBody.addEventListener("submit", async (event) => {
  const target = event.target;
  if (!(target instanceof HTMLFormElement)) return;

  if (target.matches("[data-mobile-inline-form]")) {
    event.preventDefault();
    await saveMobileInlineEdit(target);
  }
});

energyTableBody.addEventListener("click", async (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return;

  if (target.dataset.editEnergy !== undefined) {
    editEnergySale(Number(target.dataset.editEnergy));
  }

  if (target.dataset.cancelEnergyInline !== undefined) {
    cancelEnergyInlineEdit(Number(target.dataset.cancelEnergyInline));
  }

  if (target.dataset.deleteEnergy !== undefined) {
    deleteEnergySale(Number(target.dataset.deleteEnergy));
    await persistAndRender();
  }
});

energyTableBody.addEventListener("submit", async (event) => {
  const target = event.target;
  if (!(target instanceof HTMLFormElement)) return;

  if (target.matches("[data-energy-inline-form]")) {
    event.preventDefault();
    await saveEnergyInlineEdit(target);
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
      sidebarButtonColor: remoteState.settings?.sidebarButtonColor || "#20316f",
      sidebarButtonStyle: remoteState.settings?.sidebarButtonStyle || "style-1",
      toolbarColor: remoteState.settings?.toolbarColor || "#b3262d",
      sidebarTabColors: {
        ...DEFAULT_SIDEBAR_TAB_COLORS,
        ...(remoteState.settings?.sidebarTabColors && typeof remoteState.settings.sidebarTabColors === "object"
          ? remoteState.settings.sidebarTabColors
          : {}),
      },
      mobileColumnWidths: {
        ...DEFAULT_MOBILE_COLUMN_WIDTHS,
        ...(remoteState.settings?.mobileColumnWidths && typeof remoteState.settings.mobileColumnWidths === "object"
          ? remoteState.settings.mobileColumnWidths
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
  initializeConfigAccordion();
  renderAll();
  initializeMobileColumnResizing();
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
  const monthSales = getSelectedMonthBillingSales();
  renderBillingFilterMenus(monthSales);

  const visibleSales = monthSales.filter((sale) => {
    if (billingFilters.data && sale.data !== billingFilters.data) return false;
    if (billingFilters.seguro && (sale.seguro || "") !== billingFilters.seguro) return false;
    if (billingFilters.categoria && (sale.categoria || "") !== billingFilters.categoria) return false;
    if (billingFilters.modalidade && (sale.modalidade || "") !== billingFilters.modalidade) return false;
    return true;
  });

  if (visibleSales.length === 0) {
    billingTableBody.innerHTML = emptyRow(11, "Ainda não existem registos de faturação.");
  } else {
    billingTableBody.innerHTML = visibleSales
      .map((sale) => {
        const billingIndex = state.billingSales.indexOf(sale);
        const isInlineEditing = billingInlineEditingIndex === billingIndex;
        const rowAccent = getBillingDateAccent(sale.data);

        return `
        <tr data-search-row-index="${billingIndex}" class="${isInlineEditing ? "billing-inline-active" : ""} billing-date-row ledger-date-row" style="--billing-row-text: ${rowAccent.accent}; --row-text-accent: ${rowAccent.accent}; --ledger-row-bg: ${rowAccent.soft}; --ledger-row-bg-hover: ${rowAccent.softer}; --ledger-row-border: ${hexToRgba(rowAccent.accent, 0.38)};">
            <td>${formatDate(sale.data)}</td>
            <td>${escapeHtml(sale.nif || "")}</td>
            <td>${escapeHtml(sale.equipamento || "")}</td>
            <td>${formatCurrency(Number(sale.valorSemIva || 0))}</td>
            <td>${escapeHtml(sale.seguro || "")}</td>
            <td>${escapeHtml(sale.categoria || "")}</td>
            <td>${escapeHtml(sale.imei || "")}</td>
            <td>${escapeHtml(sale.fatura || "")}</td>
            <td>${escapeHtml(sale.modalidade || "")}</td>
            <td>${escapeHtml(sale.notas || "")}</td>
            <td class="actions-cell">
              <button type="button" class="table-action icon edit" data-edit-billing="${billingIndex}" aria-label="Editar" title="Editar">
                <span class="billing-action-icon billing-action-icon-edit" aria-hidden="true"></span>
              </button>
              <button type="button" class="table-action icon delete" data-delete-billing="${billingIndex}" aria-label="Eliminar" title="Eliminar">
                <span class="billing-action-icon billing-action-icon-delete" aria-hidden="true"></span>
              </button>
            </td>
          </tr>
          ${
            isInlineEditing
              ? `
                <tr class="billing-inline-editor-row">
                  <td colspan="11">
                    <form class="billing-inline-form" data-billing-inline-form="${billingIndex}">
                      <div class="billing-inline-grid">
                        <label>
                          <span>Data</span>
                          <input type="date" name="data" value="${escapeHtml(sale.data || "")}" />
                        </label>
                        <label>
                          <span>NIF</span>
                          <input type="text" name="nif" value="${escapeHtml(sale.nif || "")}" />
                        </label>
                        <label>
                          <span>Equipamento</span>
                          <input type="text" name="equipamento" value="${escapeHtml(sale.equipamento || "")}" />
                        </label>
                        <label>
                          <span>Valor em €</span>
                          <input type="text" name="valor" inputmode="decimal" value="${escapeHtml(formatMoneyInputValue(Number(sale.valor || 0)))}" />
                        </label>
                        <label>
                          <span>Seguro</span>
                          <select name="seguro">
                            ${buildSelectOptions(["Sim", "Não"], sale.seguro || "", true)}
                          </select>
                        </label>
                        <label>
                          <span>Valor mensal seguro</span>
                          <input type="text" name="valorSeguro" inputmode="decimal" value="${escapeHtml(sale.valorSeguro ? formatMoneyInputValue(Number(sale.valorSeguro || 0)) : "")}" />
                        </label>
                        <label>
                          <span>Categoria</span>
                          <select name="categoria">
                            ${buildSelectOptions(["Smartphone", "Diversidade", "Reserva", "Outro"], sale.categoria || "", true)}
                          </select>
                        </label>
                        <label>
                          <span>IMEI/SN</span>
                          <input type="text" name="imei" value="${escapeHtml(sale.imei || "")}" />
                        </label>
                        <label>
                          <span>Nº Fatura</span>
                          <input type="text" name="fatura" value="${escapeHtml(sale.fatura || "")}" />
                        </label>
                        <label>
                          <span>Modalidade</span>
                          <select name="modalidade">
                            ${buildSelectOptions(["Pronto Pagamento", "MEOS", "Vencimento", "Prestações"], sale.modalidade || "", true)}
                          </select>
                        </label>
                        <label class="billing-inline-notes">
                          <span>Notas</span>
                          <textarea name="notas" rows="2">${escapeHtml(sale.notas === "-" ? "" : sale.notas || "")}</textarea>
                        </label>
                      </div>
                      <div class="billing-inline-actions">
                        <button type="submit" class="primary-button compact">Guardar</button>
                        <button type="button" class="ghost-button compact" data-cancel-billing-inline="${billingIndex}">Cancelar</button>
                      </div>
                    </form>
                  </td>
                </tr>
              `
              : ""
          }
        `;
      })
      .join("");
  }

  totalSales.textContent = String(monthSales.length);
  totalAmount.textContent = formatCurrency(
    monthSales.reduce((sum, sale) => sum + getBillingEffectiveValue(sale), 0)
  );
}

function renderBillingFilterMenus(monthSales) {
  const menuDefinitions = [
    {
      field: "data",
      values: [...new Set(monthSales.map((sale) => sale.data).filter(Boolean))].sort(),
      formatter: (value) => formatDate(value),
      emptyLabel: "Sem datas",
    },
    {
      field: "seguro",
      values: ["Sim", "Não"],
      formatter: (value) => value,
    },
    {
      field: "categoria",
      values: [...new Set(monthSales.map((sale) => sale.categoria).filter(Boolean))].sort((a, b) =>
        a.localeCompare(b)
      ),
      formatter: (value) => value,
      emptyLabel: "Sem categorias",
    },
    {
      field: "modalidade",
      values: [...new Set(monthSales.map((sale) => sale.modalidade).filter(Boolean))].sort((a, b) =>
        a.localeCompare(b)
      ),
      formatter: (value) => value,
      emptyLabel: "Sem modalidades",
    },
  ];

  menuDefinitions.forEach(({ field, values, formatter, emptyLabel }) => {
    const menu = document.querySelector(`[data-billing-filter-menu="${field}"]`);
    const trigger = document.querySelector(`[data-billing-filter-trigger="${field}"]`);
    if (!(menu instanceof HTMLElement) || !(trigger instanceof HTMLElement)) return;

    const currentValue = billingFilters[field] || "";
    trigger.classList.toggle("is-filtered", Boolean(currentValue));
    trigger.setAttribute("aria-expanded", billingFilterOpenKey === field ? "true" : "false");

    const items = [
      `
        <button type="button" class="billing-filter-option ${currentValue === "" ? "active" : ""}" data-billing-filter-value="${field}" data-billing-filter-choice="">
          Todos
        </button>
      `,
      ...(values.length > 0
        ? values.map(
            (value) => `
              <button type="button" class="billing-filter-option ${currentValue === value ? "active" : ""}" data-billing-filter-value="${field}" data-billing-filter-choice="${escapeHtml(value)}">
                ${escapeHtml(formatter(value))}
              </button>
            `
          )
        : [
            `<div class="billing-filter-empty">${emptyLabel || "Sem opções"}</div>`,
          ]),
    ];

    menu.innerHTML = items.join("");
    menu.classList.toggle("open", billingFilterOpenKey === field);
    menu.setAttribute("aria-hidden", billingFilterOpenKey === field ? "false" : "true");
    if (billingFilterOpenKey === field) {
      positionBillingFilterMenu(menu, trigger);
    } else {
      menu.style.left = "";
      menu.style.top = "";
      menu.style.transform = "";
      menu.style.maxHeight = "";
      menu.style.position = "";
    }
  });
}

function renderMobileFilterMenus(visibleSales) {
  const menuDefinitions = [
    {
      field: "data",
      values: [...new Set(visibleSales.map((sale) => sale.data).filter(Boolean))].sort(),
      formatter: (value) => formatDate(value),
      emptyLabel: "Sem datas",
    },
    {
      field: "tarifario",
      values: [...new Set(visibleSales.map((sale) => sale.tarifario).filter(Boolean))].sort((a, b) =>
        a.localeCompare(b)
      ),
      formatter: (value) => value,
      emptyLabel: "Sem tarifários",
    },
    {
      field: "fidelizacao",
      values: ["Sim", "Não"],
      formatter: (value) => value,
    },
    {
      field: "portabilidade",
      values: ["Sim", "Não"],
      formatter: (value) => value,
    },
    {
      field: "bestOffer",
      values: ["Sim", "Não"],
      formatter: (value) => value,
    },
    {
      field: "netSegura",
      values: ["Sim", "Não"],
      formatter: (value) => value,
    },
    {
      field: "rceAceite",
      values: ["Sim", "Não"],
      formatter: (value) => value,
    },
  ];

  menuDefinitions.forEach(({ field, values, formatter, emptyLabel }) => {
    const menu = document.querySelector(`[data-mobile-filter-menu="${field}"]`);
    const trigger = document.querySelector(`[data-mobile-filter-trigger="${field}"]`);
    if (!(menu instanceof HTMLElement) || !(trigger instanceof HTMLElement)) return;

    const currentValue = mobileFilters[field] || "";
    trigger.classList.toggle("is-filtered", Boolean(currentValue));
    trigger.setAttribute("aria-expanded", mobileFilterOpenKey === field ? "true" : "false");

    const items = [
      `
        <button type="button" class="mobile-filter-option ${currentValue === "" ? "active" : ""}" data-mobile-filter-choice="${field}" data-mobile-filter-value="">
          Todos
        </button>
      `,
      ...(values.length > 0
        ? values.map(
            (value) => `
              <button type="button" class="mobile-filter-option ${currentValue === value ? "active" : ""}" data-mobile-filter-choice="${field}" data-mobile-filter-value="${escapeHtml(value)}">
                ${escapeHtml(formatter(value))}
              </button>
            `
          )
        : [`<div class="mobile-filter-empty">${emptyLabel || "Sem opções"}</div>`]),
    ];

    menu.innerHTML = items.join("");
    menu.classList.toggle("open", mobileFilterOpenKey === field);
    menu.setAttribute("aria-hidden", mobileFilterOpenKey === field ? "false" : "true");
    if (mobileFilterOpenKey === field) {
      positionBillingFilterMenu(menu, trigger);
    } else {
      menu.style.left = "";
      menu.style.top = "";
      menu.style.transform = "";
      menu.style.maxHeight = "";
      menu.style.position = "";
      menu.style.zIndex = "";
    }
  });
}

function refreshMobileFilterMenus() {
  renderMobileFilterMenus(state.mobileSales.filter((sale) => isInSelectedMonth(sale.data)));
}

function toggleMobileFilterMenu(field) {
  mobileFilterOpenKey = mobileFilterOpenKey === field ? "" : field;
  refreshMobileFilterMenus();
}

function closeMobileFilterMenu() {
  if (!mobileFilterOpenKey) return;
  mobileFilterOpenKey = "";
  refreshMobileFilterMenus();
}

function buildSelectOptions(values, selectedValue, includeEmpty = false) {
  return [
    includeEmpty ? `<option value="">Selecionar</option>` : "",
    ...values.map(
      (value) => `<option value="${escapeHtml(value)}" ${selectedValue === value ? "selected" : ""}>${escapeHtml(value)}</option>`
    ),
  ].join("");
}

function formatMoneyInputValue(value) {
  return Number.isFinite(value) ? Number(value).toFixed(2) : "";
}

function setBillingFilter(field, value) {
  if (!(field in billingFilters)) return;
  billingFilters = {
    ...billingFilters,
    [field]: value,
  };
}

function toggleBillingFilterMenu(field) {
  billingFilterOpenKey = billingFilterOpenKey === field ? "" : field;
  refreshBillingFilterMenus();
}

function closeBillingFilterMenu() {
  if (!billingFilterOpenKey) return;
  billingFilterOpenKey = "";
  refreshBillingFilterMenus();
}

function renderMobileSales() {
  const visibleSales = state.mobileSales
    .filter((sale) => isInSelectedMonth(sale.data))
    .sort((a, b) => compareSaleDates(a.data, b.data));
  renderMobileFilterMenus(visibleSales);
  if (mobileTotalSales instanceof HTMLElement) {
    mobileTotalSales.textContent = String(visibleSales.length);
  }

  const filteredSales = visibleSales.filter((sale) => {
    if (mobileFilters.data && sale.data !== mobileFilters.data) return false;
    if (mobileFilters.tarifario && (sale.tarifario || "") !== mobileFilters.tarifario) return false;
    if (mobileFilters.fidelizacao && (sale.fidelizacao || "") !== mobileFilters.fidelizacao) return false;
    if (mobileFilters.portabilidade && (sale.portabilidade || "") !== mobileFilters.portabilidade) return false;
    if (mobileFilters.bestOffer && (sale.bestOffer || "") !== mobileFilters.bestOffer) return false;
    if (mobileFilters.netSegura && (sale.netSegura || "") !== mobileFilters.netSegura) return false;
    if (mobileFilters.rceAceite && (sale.rceAceite || "") !== mobileFilters.rceAceite) return false;
    return true;
  });

  if (filteredSales.length === 0) {
    mobileTableBody.innerHTML = emptyRow(12, "Ainda não existem registos móveis.");
    return;
  }

  mobileTableBody.innerHTML = filteredSales
    .map(
      (sale) => {
        const mobileIndex = state.mobileSales.indexOf(sale);
        const isInlineEditing = mobileInlineEditingIndex === mobileIndex;
        const rowAccent = getBillingDateAccent(sale.data);

        return `
        <tr data-search-row-index="${mobileIndex}" class="${isInlineEditing ? "mobile-inline-active" : ""} mobile-date-row ledger-date-row" style="--row-text-accent: ${rowAccent.accent}; --ledger-row-bg: ${rowAccent.soft}; --ledger-row-bg-hover: ${rowAccent.softer}; --ledger-row-border: ${hexToRgba(rowAccent.accent, 0.38)};">
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
          <td class="actions-cell">
            <button type="button" class="table-action icon edit" data-edit-mobile="${mobileIndex}" aria-label="Editar" title="Editar">
              <span class="billing-action-icon billing-action-icon-edit" aria-hidden="true"></span>
            </button>
            <button type="button" class="table-action icon delete" data-delete-mobile="${mobileIndex}" aria-label="Eliminar" title="Eliminar">
              <span class="billing-action-icon billing-action-icon-delete" aria-hidden="true"></span>
            </button>
          </td>
        </tr>
        ${
          isInlineEditing
            ? `
              <tr class="mobile-inline-editor-row">
                <td colspan="12">
                  <form class="billing-inline-form mobile-inline-form" data-mobile-inline-form="${mobileIndex}">
                    <div class="billing-inline-grid">
                      <label>
                        <span>Data</span>
                        <input type="date" name="data" value="${escapeHtml(sale.data || "")}" />
                      </label>
                      <label>
                        <span>Nome</span>
                        <input type="text" name="nome" value="${escapeHtml(sale.nome || "")}" />
                      </label>
                      <label>
                        <span>NIF</span>
                        <input type="text" name="nif" value="${escapeHtml(sale.nif || "")}" />
                      </label>
                      <label>
                        <span>Número</span>
                        <input type="text" name="numero" value="${escapeHtml(sale.numero || "")}" />
                      </label>
                      <label>
                        <span>Tarifário</span>
                        <input type="text" name="tarifario" value="${escapeHtml(sale.tarifario || "")}" />
                      </label>
                      <label>
                        <span>ID</span>
                        <input type="text" name="idVenda" value="${escapeHtml(sale.idVenda || "")}" />
                      </label>
                      <label>
                        <span>Fidelização</span>
                        <select name="fidelizacao">
                          ${buildSelectOptions(["Sim", "Não"], sale.fidelizacao || "", true)}
                        </select>
                      </label>
                      <label>
                        <span>Portabilidade</span>
                        <select name="portabilidade">
                          ${buildSelectOptions(["Sim", "Não"], sale.portabilidade || "", true)}
                        </select>
                      </label>
                      <label>
                        <span>NBO</span>
                        <select name="bestOffer">
                          ${buildSelectOptions(["Sim", "Não"], sale.bestOffer || "", true)}
                        </select>
                      </label>
                      <label>
                        <span>Segura</span>
                        <select name="netSegura">
                          ${buildSelectOptions(["Sim", "Não"], sale.netSegura || "", true)}
                        </select>
                      </label>
                      <label>
                        <span>RCE Aceite</span>
                        <select name="rceAceite">
                          ${buildSelectOptions(["Sim", "Não"], sale.rceAceite || "", true)}
                        </select>
                      </label>
                    </div>
                    <div class="billing-inline-actions">
                      <button type="submit" class="primary-button compact">Guardar</button>
                      <button type="button" class="ghost-button compact" data-cancel-mobile-inline="${mobileIndex}">Cancelar</button>
                    </div>
                  </form>
                </td>
              </tr>
            `
            : ""
        }
      `;
      }
    )
    .join("");
}

function renderEnergySales() {
  const visibleSales = state.energySales.filter((sale) => isInSelectedMonth(sale.data));
  if (energyTotalSales instanceof HTMLElement) {
    energyTotalSales.textContent = String(visibleSales.length);
  }

  if (visibleSales.length === 0) {
    energyTableBody.innerHTML = emptyRow(9, "Ainda não existem registos de energia.");
    return;
  }

  energyTableBody.innerHTML = visibleSales
    .map(
      (sale) => {
        const energyIndex = state.energySales.indexOf(sale);
        const isInlineEditing = energyInlineEditingIndex === energyIndex;
        const rowAccent = getBillingDateAccent(sale.data);
        return `
        <tr data-search-row-index="${energyIndex}" class="${isInlineEditing ? "energy-inline-active" : ""} energy-date-row ledger-date-row" style="--row-text-accent: ${rowAccent.accent}; --ledger-row-bg: ${rowAccent.soft}; --ledger-row-bg-hover: ${rowAccent.softer}; --ledger-row-border: ${hexToRgba(rowAccent.accent, 0.38)};">
          <td>${formatDate(sale.data)}</td>
          <td>${escapeHtml(sale.nome || "")}</td>
          <td>${escapeHtml(sale.nif || "")}</td>
          <td>${escapeHtml(sale.idVenda || "")}</td>
          <td>${escapeHtml(sale.requisicao || "")}</td>
          <td>${escapeHtml(sale.potencia || "")}</td>
          <td>${escapeHtml(sale.estado || "")}</td>
          <td>${escapeHtml(sale.observacoes || "")}</td>
          <td class="actions-cell">
            <button type="button" class="table-action icon edit" data-edit-energy="${energyIndex}" aria-label="Editar" title="Editar">
              <span class="billing-action-icon billing-action-icon-edit" aria-hidden="true"></span>
            </button>
            <button type="button" class="table-action icon delete" data-delete-energy="${energyIndex}" aria-label="Eliminar" title="Eliminar">
              <span class="billing-action-icon billing-action-icon-delete" aria-hidden="true"></span>
            </button>
          </td>
        </tr>
        ${
          isInlineEditing
            ? `
              <tr class="energy-inline-editor-row">
                <td colspan="9">
                  <form class="billing-inline-form energy-inline-form" data-energy-inline-form="${energyIndex}">
                    <div class="billing-inline-grid">
                      <label>
                        <span>Data</span>
                        <input type="date" name="data" value="${escapeHtml(sale.data || "")}" />
                      </label>
                      <label>
                        <span>Nome</span>
                        <input type="text" name="nome" value="${escapeHtml(sale.nome || "")}" />
                      </label>
                      <label>
                        <span>NIF</span>
                        <input type="text" name="nif" value="${escapeHtml(sale.nif || "")}" />
                      </label>
                      <label>
                        <span>ID Venda</span>
                        <input type="text" name="idVenda" value="${escapeHtml(sale.idVenda || "")}" />
                      </label>
                      <label>
                        <span>Nº Requisição</span>
                        <input type="text" name="requisicao" value="${escapeHtml(sale.requisicao || "")}" />
                      </label>
                      <label>
                        <span>Potência</span>
                        <select name="potencia">
                          ${buildSelectOptions(["1.15", "2.30", "3.45", "4.60", "5.75", "6.90", "10.35", "13.80", "17.25", "20.70", "27.60", "34.50", "41.40"], sale.potencia || "", true)}
                        </select>
                      </label>
                      <label>
                        <span>Estado</span>
                        <select name="estado">
                          ${buildSelectOptions(["Instalado", "Pendente", "Rejeitado"], sale.estado || "", true)}
                        </select>
                      </label>
                      <label class="billing-inline-notes">
                        <span>Observações</span>
                        <textarea name="observacoes" rows="2">${escapeHtml(sale.observacoes === "-" ? "" : sale.observacoes || "")}</textarea>
                      </label>
                    </div>
                    <div class="billing-inline-actions">
                      <button type="submit" class="primary-button compact">Guardar</button>
                      <button type="button" class="ghost-button compact" data-cancel-energy-inline="${energyIndex}">Cancelar</button>
                    </div>
                  </form>
                </td>
              </tr>
            `
            : ""
        }
      `;
      }
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
      <button type="button" class="access-service-button ${service.id === activeService.id ? "active" : ""}" data-access-service="${service.id}" style="--access-service-tone: ${service.color || "#8fe56b"};">
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

  const reportDates = [...summaryByDate.keys()].sort();
  if (reportDateFilter && !reportDates.includes(reportDateFilter)) {
    reportDateFilter = "";
    reportDateFilterOpen = false;
  }

  reportTotalNet.textContent = formatCurrency(totals.totalSemIva);
  reportTotalSmartphone.textContent = String(totals.smartphone);
  reportTotalDiversidade.textContent = String(totals.diversidade);
  reportTotalSeguros.textContent = String(totals.seguros);
  reportTotalMovel.textContent = String(totals.movel);
  reportTotalEnergia.textContent = String(totals.energia);

  refreshReportDateMenu();

  const rows = reportDates
    .filter((date) => !reportDateFilter || date === reportDateFilter)
    .map((date) => ({ date, entry: summaryByDate.get(date) }))
    .filter((row) => Boolean(row.entry));

  if (rows.length === 0) {
    reportTableBody.innerHTML = emptyRow(
      7,
      "Ainda não existem dados para apresentar no relatório."
    );
    return;
  }

  reportTableBody.innerHTML = rows
    .map(
      ({ date, entry }) => `
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
  goalActualBilling.classList.toggle("goal-value-good", metrics.actualBilling >= metrics.goals.billing);
  goalActualBilling.classList.toggle("goal-value-bad", metrics.actualBilling < metrics.goals.billing);
  goalPercentBilling.textContent = formatPercent(metrics.billingRatio * 100);
  goalMissingBilling.textContent = formatCurrency(Math.max(metrics.goals.billing - metrics.actualBilling, 0));
  goalWeightedBilling.textContent = formatPercent(metrics.billingWeighted);

  goalActualMobile.textContent = String(metrics.actualMobile);
  goalTargetMobile.textContent = String(metrics.goals.mobile);
  goalActualMobile.classList.toggle("goal-value-good", metrics.actualMobile >= metrics.goals.mobile);
  goalActualMobile.classList.toggle("goal-value-bad", metrics.actualMobile < metrics.goals.mobile);
  goalPercentMobile.textContent = formatPercent(metrics.mobileRatio * 100);
  goalMissingMobile.textContent = String(Math.max(metrics.goals.mobile - metrics.actualMobile, 0));
  goalWeightedMobile.textContent = formatPercent(metrics.mobileWeighted);

  goalActualEnergy.textContent = String(metrics.actualEnergy);
  goalTargetEnergy.textContent = String(metrics.goals.energy);
  goalActualEnergy.classList.toggle("goal-value-good", metrics.actualEnergy >= metrics.goals.energy);
  goalActualEnergy.classList.toggle("goal-value-bad", metrics.actualEnergy < metrics.goals.energy);
  goalPercentEnergy.textContent = formatPercent(metrics.energyRatio * 100);
  goalMissingEnergy.textContent = String(Math.max(metrics.goals.energy - metrics.actualEnergy, 0));
  goalWeightedEnergy.textContent = formatPercent(metrics.energyWeighted);

  goalBillingProgress.textContent = formatPercent(metrics.billingRatio * 100);
  goalMobileProgress.textContent = formatPercent(metrics.mobileRatio * 100);
  goalEnergyProgress.textContent = formatPercent(metrics.energyRatio * 100);
  goalTotalProgress.textContent = formatPercent(metrics.totalWeighted);
  goalInsuranceProgress.textContent = formatPercent(metrics.insuranceActualRatio * 100);
  goalDiversityProgress.textContent = formatPercent(metrics.diversityActualRatio * 100);

  goalActualInsurance.textContent = formatPercent(metrics.insuranceActualRatio * 100);
  goalTargetInsurance.textContent = formatPercent(metrics.insuranceGoalRatio * 100);
  goalActualInsurance.classList.toggle(
    "goal-value-good",
    metrics.insuranceActualRatio >= metrics.insuranceGoalRatio
  );
  goalActualInsurance.classList.toggle(
    "goal-value-bad",
    metrics.insuranceActualRatio < metrics.insuranceGoalRatio
  );
  goalPercentInsurance.textContent = "";
  const insuranceTargetCount = metrics.smartphoneCount > 0
    ? Math.ceil(metrics.smartphoneCount * metrics.insuranceGoalRatio)
    : 0;
  goalMissingInsurance.textContent = String(
    Math.max(insuranceTargetCount - metrics.smartphoneWithInsurance, 0)
  );
  goalWeightedInsurance.textContent = `${metrics.smartphoneWithInsurance} / ${metrics.smartphoneCount}`;

  goalActualDiversity.textContent = formatPercent(metrics.diversityActualRatio * 100);
  goalTargetDiversity.textContent = formatPercent(metrics.diversityGoalRatio * 100);
  goalActualDiversity.classList.toggle(
    "goal-value-good",
    metrics.diversityActualRatio >= metrics.diversityGoalRatio
  );
  goalActualDiversity.classList.toggle(
    "goal-value-bad",
    metrics.diversityActualRatio < metrics.diversityGoalRatio
  );
  goalPercentDiversity.textContent = "";
  const diversityTargetCount = metrics.smartphoneCount > 0
    ? Math.ceil(metrics.smartphoneCount * metrics.diversityGoalRatio)
    : 0;
  goalMissingDiversity.textContent = String(
    Math.max(diversityTargetCount - metrics.diversidadeCount, 0)
  );
  goalWeightedDiversity.textContent = `${metrics.diversidadeCount} / ${metrics.smartphoneCount}`;
}

function renderGeneral() {
  const months = getTrackedMonths();
  const monthColors = [
    "#ff8a80",
    "#ffb74d",
    "#fff176",
    "#81c784",
    "#4dd0e1",
    "#64b5f6",
    "#7986cb",
    "#ba68c8",
    "#f06292",
    "#a1887f",
    "#90a4ae",
    "#ffd54f",
  ];

  generalTableBody.innerHTML = months
    .map((month, index) => {
      const metrics = getGoalMetrics(month);
      const color = monthColors[index % monthColors.length];
      return `
        <tr class="${month === selectedMonth ? "general-month-selected" : ""}" style="--month-accent: ${color};">
          <td>${formatMonth(month)}</td>
          <td>${formatPercent(metrics.billingRatio * 100)}</td>
          <td>${formatPercent(metrics.mobileRatio * 100)}</td>
          <td>${formatPercent(metrics.energyRatio * 100)}</td>
          <td>${formatPercent(metrics.insuranceActualRatio * 100)}</td>
          <td>${formatPercent(metrics.diversityActualRatio * 100)}</td>
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
      <section class="incentive-brand ${selectedIncentiveBrand === brandId ? "active" : ""}" data-incentive-brand="${brandId}">
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

  billingInlineEditingIndex = index;
  editingBillingIndex = null;
  activateTab("faturacao");
  renderBillingSales();
}

function cancelBillingInlineEdit(index) {
  if (billingInlineEditingIndex !== index) return;
  billingInlineEditingIndex = null;
  renderBillingSales();
}

async function saveBillingInlineEdit(form) {
  const formData = new FormData(form);
  const billingIndex = Number(form.dataset.billingInlineForm || "");
  const existingSale = state.billingSales[billingIndex];
  if (!existingSale) return;

  const grossValue = parseMoneyInput(formData.get("valor"));
  const valueOnSale = {
    ...existingSale,
    data: valueAsText(formData.get("data")),
    nif: valueAsText(formData.get("nif")),
    equipamento: valueAsText(formData.get("equipamento")),
    valor: grossValue,
    valorSemIva: calculateNetValue(grossValue),
    seguro: valueAsText(formData.get("seguro")),
    valorSeguro: parseMoneyInput(formData.get("valorSeguro")),
    valorSeguroAnual: calculateInsuranceAnnual(parseMoneyInput(formData.get("valorSeguro"))),
    categoria: valueAsText(formData.get("categoria")),
    imei: valueAsText(formData.get("imei")),
    fatura: valueAsText(formData.get("fatura")),
    modalidade: valueAsText(formData.get("modalidade")),
    notas: valueAsText(formData.get("notas")) || "-",
  };

  state.billingSales[billingIndex] = valueOnSale;
  billingInlineEditingIndex = null;
  await persistAndRender();
}

function deleteBillingSale(index) {
  state.billingSales.splice(index, 1);
  if (billingInlineEditingIndex === index) {
    billingInlineEditingIndex = null;
  } else if (billingInlineEditingIndex !== null && index < billingInlineEditingIndex) {
    billingInlineEditingIndex -= 1;
  }
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

  mobileInlineEditingIndex = index;
  editingMobileIndex = null;
  activateTab("movel");
  renderMobileSales();
}

function cancelMobileInlineEdit(index) {
  if (mobileInlineEditingIndex !== index) return;
  mobileInlineEditingIndex = null;
  renderMobileSales();
}

async function saveMobileInlineEdit(form) {
  const formData = new FormData(form);
  const mobileIndex = Number(form.dataset.mobileInlineForm || "");
  const existingSale = state.mobileSales[mobileIndex];
  if (!existingSale) return;

  const sale = {
    ...existingSale,
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
    formaAceitacao: existingSale.formaAceitacao || "",
  };

  state.mobileSales[mobileIndex] = sale;
  mobileInlineEditingIndex = null;
  await persistAndRender();
}

function deleteMobileSale(index) {
  state.mobileSales.splice(index, 1);
  if (mobileInlineEditingIndex === index) {
    mobileInlineEditingIndex = null;
  } else if (mobileInlineEditingIndex !== null && index < mobileInlineEditingIndex) {
    mobileInlineEditingIndex -= 1;
  }
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

  energyInlineEditingIndex = index;
  editingEnergyIndex = null;
  activateTab("meo-energia");
  renderEnergySales();
}

function cancelEnergyInlineEdit(index) {
  if (energyInlineEditingIndex !== index) return;
  energyInlineEditingIndex = null;
  renderEnergySales();
}

async function saveEnergyInlineEdit(form) {
  const formData = new FormData(form);
  const energyIndex = Number(form.dataset.energyInlineForm || "");
  const existingSale = state.energySales[energyIndex];
  if (!existingSale) return;

  const sale = {
    ...existingSale,
    data: valueAsText(formData.get("data")),
    nome: valueAsText(formData.get("nome")),
    nif: valueAsText(formData.get("nif")),
    idVenda: valueAsText(formData.get("idVenda")),
    requisicao: valueAsText(formData.get("requisicao")),
    potencia: valueAsText(formData.get("potencia")),
    estado: valueAsText(formData.get("estado")),
    observacoes: valueAsText(formData.get("observacoes")) || "-",
  };

  state.energySales[energyIndex] = sale;
  energyInlineEditingIndex = null;
  await persistAndRender();
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
  if (energyInlineEditingIndex === index) {
    energyInlineEditingIndex = null;
  } else if (energyInlineEditingIndex !== null && index < energyInlineEditingIndex) {
    energyInlineEditingIndex -= 1;
  }
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

function normalizeSearchText(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

function rankSearchElement(element, query) {
  const normalizedText = normalizeSearchText(element.textContent);
  let score = normalizedText.length;

  if (normalizedText === query) score -= 1000;
  if (element.matches(".tab, .support-tab, .incentive-tab, .access-service-button")) score -= 350;
  if (element.matches("h1, h2, h3, .support-title, .access-service-label, .contact-item span")) {
    score -= 120;
  }
  if (element.matches(".panel, .support-panel, .incentive-brand")) score += 200;

  return score;
}

function getSearchTargetKey(target, element) {
  const fallbackKey = `raw:${element?.tagName || "unknown"}:${normalizeSearchText(element?.textContent || "").slice(0, 60)}`;

  if (!(element instanceof HTMLElement) || !target) {
    return fallbackKey;
  }

  if (target.kind === "main-tab") return `main:${target.tabId}`;
  if (target.kind === "support-tab") return `support:${target.supportTabId}`;
  if (target.kind === "support-panel") return `support-panel:${target.supportTabId}`;
  if (target.kind === "incentive-brand" || target.kind === "incentive-panel") return `brand:${target.brandId}`;
  if (target.kind === "access-service") return `access:${target.serviceId}`;
  if (target.kind === "panel") return `panel:${target.tabId}`;
  if (target.kind === "billing-row") return `billing-row:${target.rowIndex}`;
  if (target.kind === "mobile-row") return `mobile-row:${target.rowIndex}`;
  if (target.kind === "energy-row") return `energy-row:${target.rowIndex}`;

  return fallbackKey;
}

function resolveSearchTarget(element) {
  const accessButton = element.closest(".access-service-button");
  if (accessButton) {
    return {
      kind: "access-service",
      serviceId: accessButton.dataset.accessService || "",
      element: accessButton,
    };
  }

  const tableRow = element.closest("tbody tr");
  if (tableRow && tableRow.dataset.searchRowIndex !== undefined) {
    const billingTable = tableRow.closest(".billing-table");
    if (billingTable) {
      return {
        kind: "billing-row",
        rowIndex: Number(tableRow.dataset.searchRowIndex || "0"),
        element: tableRow,
      };
    }

    const mobileTable = tableRow.closest(".mobile-table");
    if (mobileTable) {
      return {
        kind: "mobile-row",
        rowIndex: Number(tableRow.dataset.searchRowIndex || "0"),
        element: tableRow,
      };
    }

    const energyTable = tableRow.closest(".meo-energy-table");
    if (energyTable) {
      return {
        kind: "energy-row",
        rowIndex: Number(tableRow.dataset.searchRowIndex || "0"),
        element: tableRow,
      };
    }
  }

  const incentiveTab = element.closest(".incentive-tab");
  if (incentiveTab) {
    return {
      kind: "incentive-brand",
      brandId: incentiveTab.dataset.incentiveTab || "",
      element: incentiveTab,
    };
  }

  const supportTab = element.closest(".support-tab");
  if (supportTab && !supportTab.classList.contains("incentive-tab")) {
    return {
      kind: "support-tab",
      supportTabId: supportTab.dataset.supportTab || "",
      element: supportTab,
    };
  }

  const mainTab = element.closest(".tab");
  if (mainTab) {
    return {
      kind: "main-tab",
      tabId: mainTab.dataset.tab || "",
      element: mainTab,
    };
  }

  const supportPanel = element.closest(".support-panel");
  if (supportPanel) {
    return {
      kind: "support-panel",
      supportTabId: supportPanel.id.replace(/^support-/, ""),
      element: supportPanel,
    };
  }

  const incentiveBrand = element.closest(".incentive-brand");
  if (incentiveBrand) {
    return {
      kind: "incentive-panel",
      brandId: incentiveBrand.dataset.incentiveBrand || "",
      element: incentiveBrand,
    };
  }

  const panel = element.closest(".panel");
  if (panel) {
    return {
      kind: "panel",
      tabId: panel.id,
      element: panel,
    };
  }

  return null;
}

function findSearchTarget(query) {
  const normalizedQuery = normalizeSearchText(query);
  if (!normalizedQuery) return null;

  const matches = findSearchMatches(normalizedQuery);
  if (!matches.length) return null;

  const bestMatch = matches[0];
  return resolveSearchTarget(bestMatch) || { kind: "raw", element: bestMatch };
}

function findSearchMatches(query) {
  const normalizedQuery = normalizeSearchText(query);
  if (!normalizedQuery) return [];

  const candidates = Array.from(document.querySelectorAll("body *")).filter((element) => {
    if (!(element instanceof HTMLElement)) return false;
    if (["SCRIPT", "STYLE", "NOSCRIPT"].includes(element.tagName)) return false;
    const text = normalizeSearchText(element.textContent);
    return text.includes(normalizedQuery);
  });

  if (!candidates.length) return [];

  candidates.sort((a, b) => rankSearchElement(a, normalizedQuery) - rankSearchElement(b, normalizedQuery));

  const seen = new Set();
  return candidates
    .map((element) => {
      const target = resolveSearchTarget(element) || { kind: "raw", element };
      const key = getSearchTargetKey(target, element);

      if (seen.has(key)) return null;
      seen.add(key);
      return { ...target, element, searchKey: key };
    })
    .filter(Boolean);
}

function highlightSearchResult(element) {
  if (!(element instanceof HTMLElement)) return;
}

function focusSearchTarget(target) {
  if (!target) return;

  if (target.kind === "main-tab" && target.tabId) {
    searchNavigationActive = true;
    activateTab(target.tabId);
    window.setTimeout(() => {
      const panel = document.getElementById(target.tabId);
      panel?.scrollIntoView({ behavior: "smooth", block: "start" });
      highlightSearchResult(panel || target.element);
      highlightSearchText(panel || target.element, sidebarSearchInput?.value || "");
      searchNavigationActive = false;
    }, 60);
    return;
  }

  if (target.kind === "support-tab" && target.supportTabId) {
    searchNavigationActive = true;
    activateTab("apoio");
    const button = Array.from(supportTabs).find(
      (item) => !item.classList.contains("incentive-tab") && item.dataset.supportTab === target.supportTabId
    );
    button?.click();
    window.setTimeout(() => {
      const panel = document.getElementById(`support-${target.supportTabId}`);
      panel?.scrollIntoView({ behavior: "smooth", block: "start" });
      highlightSearchResult(panel || target.element);
      highlightSearchText(panel || target.element, sidebarSearchInput?.value || "");
      searchNavigationActive = false;
    }, 80);
    return;
  }

  if (target.kind === "support-panel" && target.supportTabId) {
    searchNavigationActive = true;
    activateTab("apoio");
    const button = Array.from(supportTabs).find(
      (item) => !item.classList.contains("incentive-tab") && item.dataset.supportTab === target.supportTabId
    );
    button?.click();
    window.setTimeout(() => {
      const panel = document.getElementById(`support-${target.supportTabId}`);
      panel?.scrollIntoView({ behavior: "smooth", block: "start" });
      highlightSearchResult(panel || target.element);
      highlightSearchText(panel || target.element, sidebarSearchInput?.value || "");
      searchNavigationActive = false;
    }, 80);
    return;
  }

  if ((target.kind === "incentive-brand" || target.kind === "incentive-panel") && target.brandId) {
    searchNavigationActive = true;
    activateTab("incentivos");
    selectedIncentiveBrand = target.brandId;
    incentiveTabs.forEach((item) => item.classList.toggle("active", item.dataset.incentiveTab === target.brandId));
    renderIncentives();
    window.setTimeout(() => {
      const brand = document.querySelector(`.incentive-brand[data-incentive-brand="${target.brandId}"]`);
      brand?.scrollIntoView({ behavior: "smooth", block: "start" });
      highlightSearchResult(brand || target.element);
      highlightSearchText(brand || target.element, sidebarSearchInput?.value || "");
      searchNavigationActive = false;
    }, 80);
    return;
  }

  if (target.kind === "access-service" && target.serviceId) {
    searchNavigationActive = true;
    activateTab("acessos");
    selectedAccessService = target.serviceId;
    accessPasswordVisible = false;
    accessStatusMessage = "";
    accessStatusError = false;
    renderAccesses();
    window.setTimeout(() => {
      const detail = document.getElementById("access-detail");
      detail?.scrollIntoView({ behavior: "smooth", block: "start" });
      highlightSearchResult(detail || target.element);
      highlightSearchText(detail || target.element, sidebarSearchInput?.value || "");
      searchNavigationActive = false;
    }, 80);
    return;
  }

  if (target.kind === "billing-row") {
    searchNavigationActive = true;
    activateTab("faturacao");
    window.setTimeout(() => {
      const row = document.querySelector(`.billing-table tbody tr[data-search-row-index="${target.rowIndex}"]`);
      row?.scrollIntoView({ behavior: "smooth", block: "center" });
      highlightSearchResult(row || target.element);
      highlightSearchText(row || target.element, sidebarSearchInput?.value || "");
      searchNavigationActive = false;
    }, 60);
    return;
  }

  if (target.kind === "mobile-row") {
    searchNavigationActive = true;
    activateTab("movel");
    window.setTimeout(() => {
      const row = document.querySelector(`.mobile-table tbody tr[data-search-row-index="${target.rowIndex}"]`);
      row?.scrollIntoView({ behavior: "smooth", block: "center" });
      highlightSearchResult(row || target.element);
      highlightSearchText(row || target.element, sidebarSearchInput?.value || "");
      searchNavigationActive = false;
    }, 60);
    return;
  }

  if (target.kind === "energy-row") {
    searchNavigationActive = true;
    activateTab("meo-energia");
    window.setTimeout(() => {
      const row = document.querySelector(`.meo-energy-table tbody tr[data-search-row-index="${target.rowIndex}"]`);
      row?.scrollIntoView({ behavior: "smooth", block: "center" });
      highlightSearchResult(row || target.element);
      highlightSearchText(row || target.element, sidebarSearchInput?.value || "");
      searchNavigationActive = false;
    }, 60);
    return;
  }

  if (target.kind === "panel" && target.tabId) {
    searchNavigationActive = true;
    activateTab(target.tabId);
    window.setTimeout(() => {
      const panel = document.getElementById(target.tabId);
      panel?.scrollIntoView({ behavior: "smooth", block: "start" });
      highlightSearchResult(panel || target.element);
      highlightSearchText(panel || target.element, sidebarSearchInput?.value || "");
      searchNavigationActive = false;
    }, 60);
    return;
  }

  const fallbackElement = target.element instanceof HTMLElement ? target.element : null;
  fallbackElement?.scrollIntoView({ behavior: "smooth", block: "center" });
  highlightSearchResult(fallbackElement || undefined);
  highlightSearchText(fallbackElement || undefined, sidebarSearchInput?.value || "");
  searchNavigationActive = false;
}

function performSidebarSearch(query) {
  const normalizedQuery = normalizeSearchText(query);
  if (!normalizedQuery) return;

  const matches = findSearchMatches(normalizedQuery);
  if (!matches.length) {
    clearSearchHighlights();
    sidebarSearchState = {
      query: normalizedQuery,
      matches: [],
      index: -1,
      activeKey: "",
    };
    return;
  }

  const sameQuery = sidebarSearchState.query === normalizedQuery;
  sidebarSearchState.matches = matches;
  sidebarSearchState.query = normalizedQuery;

  let nextIndex = 0;
  if (sameQuery && sidebarSearchState.activeKey) {
    const currentIndex = matches.findIndex((match) => match.searchKey === sidebarSearchState.activeKey);
    nextIndex =
      currentIndex === -1 ? (sidebarSearchState.index + 1) % matches.length : (currentIndex + 1) % matches.length;
  }

  sidebarSearchState.index = nextIndex;
  sidebarSearchState.activeKey = matches[nextIndex]?.searchKey || "";

  clearSearchHighlights();
  focusSearchTarget(matches[nextIndex]);
}

function clearSearchHighlights() {
  document.querySelectorAll(".search-hit-term").forEach((node) => {
    const textNode = document.createTextNode(node.textContent || "");
    node.replaceWith(textNode);
  });
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function normalizeForSearch(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();
}

function findSearchRange(text, query) {
  const normalizedText = normalizeForSearch(text);
  const normalizedQuery = normalizeForSearch(query);
  if (!normalizedQuery) return null;

  const index = normalizedText.indexOf(normalizedQuery);
  if (index === -1) return null;

  let normalizedPosition = 0;
  let start = -1;
  let end = -1;

  for (let i = 0; i < text.length; i += 1) {
    const normalizedChar = text[i]
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase();

    for (let j = 0; j < normalizedChar.length; j += 1) {
      if (normalizedPosition === index && start === -1) start = i;
      if (normalizedPosition === index + normalizedQuery.length - 1) {
        end = i + 1;
      }
      normalizedPosition += 1;
    }
  }

  if (start === -1 || end === -1) return null;
  return { start, end };
}

function highlightSearchText(element, query) {
  if (!(element instanceof HTMLElement)) return;
  const normalizedQuery = normalizeForSearch(query);
  if (!normalizedQuery) return;

  const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT, {
    acceptNode(node) {
      const parent = node.parentElement;
      if (!parent) return NodeFilter.FILTER_REJECT;
      if (parent.closest(".search-hit-term")) return NodeFilter.FILTER_REJECT;
      if (["SCRIPT", "STYLE", "NOSCRIPT"].includes(parent.tagName)) return NodeFilter.FILTER_REJECT;
      if (!normalizeForSearch(node.textContent).includes(normalizedQuery)) return NodeFilter.FILTER_REJECT;
      return NodeFilter.FILTER_ACCEPT;
    },
  });

  const textNodes = [];
  while (walker.nextNode()) {
    textNodes.push(walker.currentNode);
  }

  textNodes.forEach((node) => {
    const text = node.textContent || "";
    const fragment = document.createDocumentFragment();
    let cursor = 0;
    let matched = false;

    while (cursor < text.length) {
      const range = findSearchRange(text.slice(cursor), normalizedQuery);
      if (!range) break;

      matched = true;
      const start = cursor + range.start;
      const end = cursor + range.end;

      if (start > cursor) {
        fragment.append(document.createTextNode(text.slice(cursor, start)));
      }

      const mark = document.createElement("mark");
      mark.className = "search-hit-term";
      mark.textContent = text.slice(start, end);
      fragment.append(mark);
      cursor = end;
    }

    if (matched) {
      if (cursor < text.length) {
        fragment.append(document.createTextNode(text.slice(cursor)));
      }
      node.replaceWith(fragment);
    }
  });
}

function resetSidebarSearchState() {
  sidebarSearchState = {
    query: "",
    matches: [],
    index: -1,
    activeKey: "",
  };
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

function compareSaleDates(dateA, dateB) {
  const safeA = dateA || "9999-12-31";
  const safeB = dateB || "9999-12-31";
  return safeA.localeCompare(safeB);
}

function getBillingDateAccent(date) {
  const safeDate = String(date || "0000-00-00");
  const parsed = safeDate.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  let accent;

  if (parsed) {
    const year = Number(parsed[1]);
    const month = Number(parsed[2]);
    const day = Number(parsed[3]);
    const utcDayIndex = Math.floor(Date.UTC(year, month - 1, day) / 86400000);
    const hue = (utcDayIndex * 137.508) % 360;
    accent = hslToHex(hue, 82, 56);
  } else {
    let hash = 0;
    for (let index = 0; index < safeDate.length; index += 1) {
      hash = (hash * 31 + safeDate.charCodeAt(index)) >>> 0;
    }
    accent = hslToHex(hash % 360, 82, 56);
  }

  const soft = hexToRgba(accent, 0.18);
  const softer = hexToRgba(accent, 0.28);
  return {
    accent,
    soft,
    softer,
    text: getReadableTextColor(accent),
  };
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
  if (sidebarButtonColorInput) sidebarButtonColorInput.value = state.settings.sidebarButtonColor;
  if (sidebarButtonStyleInput) sidebarButtonStyleInput.value = state.settings.sidebarButtonStyle || "style-1";
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

let configAccordionInitialized = false;

function initializeConfigAccordion() {
  if (configAccordionInitialized) return;

  const sections = Array.from(document.querySelectorAll("#configuracoes .config-section"));
  if (!sections.length) return;

  sections.forEach((section) => {
    section.addEventListener("toggle", () => {
      if (!section.open) return;

      sections.forEach((otherSection) => {
        if (otherSection !== section) otherSection.open = false;
      });
    });
  });

  configAccordionInitialized = true;
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
  state.settings.sidebarButtonColor = sidebarButtonColorInput?.value || "#20316f";
  state.settings.sidebarButtonStyle = sidebarButtonStyleInput?.value || "style-1";
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

function hexToRgb(hex) {
  const normalized = String(hex || "").replace("#", "").trim();
  if (normalized.length !== 6) return null;

  const numeric = Number.parseInt(normalized, 16);
  if (Number.isNaN(numeric)) return null;

  return {
    r: (numeric >> 16) & 255,
    g: (numeric >> 8) & 255,
    b: numeric & 255,
  };
}

function hexToRgba(hex, alpha) {
  const rgb = hexToRgb(hex);
  if (!rgb) return hex;
  return `rgba(${rgb.r}, ${rgb.g}, ${rgb.b}, ${alpha})`;
}

function hslToHex(h, s, l) {
  const hue = ((h % 360) + 360) % 360;
  const saturation = Math.max(0, Math.min(100, s)) / 100;
  const lightness = Math.max(0, Math.min(100, l)) / 100;
  const chroma = (1 - Math.abs(2 * lightness - 1)) * saturation;
  const sector = hue / 60;
  const secondary = chroma * (1 - Math.abs((sector % 2) - 1));
  let red = 0;
  let green = 0;
  let blue = 0;

  if (sector >= 0 && sector < 1) {
    red = chroma;
    green = secondary;
  } else if (sector >= 1 && sector < 2) {
    red = secondary;
    green = chroma;
  } else if (sector >= 2 && sector < 3) {
    green = chroma;
    blue = secondary;
  } else if (sector >= 3 && sector < 4) {
    green = secondary;
    blue = chroma;
  } else if (sector >= 4 && sector < 5) {
    red = secondary;
    blue = chroma;
  } else {
    red = chroma;
    blue = secondary;
  }

  const match = lightness - chroma / 2;
  const toHex = (value) => {
    const channel = Math.round((value + match) * 255);
    return channel.toString(16).padStart(2, "0");
  };

  return `#${toHex(red)}${toHex(green)}${toHex(blue)}`;
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
  const sidebarButtonScale = Math.max(0.88, Math.min(1.38, state.settings.menuFontSize / 15));
  const sidebarButtonBase = state.settings.sidebarButtonColor || "#20316f";
  const sidebarButtonHover = shadeColor(sidebarButtonBase, -10);
  const sidebarButtonActive = shadeColor(sidebarButtonBase, -16);
  const sidebarButtonBorder = shadeColor(sidebarButtonBase, -24);
  const sidebarButtonGlow = hexToRgba(shadeColor(sidebarButtonBase, 18), 0.28);
  const sidebarButtonPadY = Math.round(11 + (sidebarButtonScale - 1) * 10);
  const sidebarButtonPadX = Math.round(15 + (sidebarButtonScale - 1) * 8);
  const sidebarButtonPadLeft = Math.round(17 + (sidebarButtonScale - 1) * 8);
  const sidebarButtonRadius = Math.round(12 + (sidebarButtonScale - 1) * 10);
  const sidebarButtonShadow = `0 ${Math.round(9 + (sidebarButtonScale - 1) * 4)}px ${Math.round(20 + (sidebarButtonScale - 1) * 8)}px rgba(0, 0, 0, 0.18)`;
  const sidebarButtonHoverShadow = `0 ${Math.round(11 + (sidebarButtonScale - 1) * 4)}px ${Math.round(22 + (sidebarButtonScale - 1) * 8)}px rgba(0, 0, 0, 0.22)`;
  const sidebarButtonActiveShadow = `0 ${Math.round(12 + (sidebarButtonScale - 1) * 4)}px ${Math.round(24 + (sidebarButtonScale - 1) * 8)}px rgba(0, 0, 0, 0.25)`;
  const sidebarButtonLift = `${Math.round((sidebarButtonScale - 1) * 2)}px`;
  const sidebarButtonHoverLift = `${Math.round(-2 - (sidebarButtonScale - 1) * 2)}px`;
  const sidebarButtonActiveLift = `${Math.round(-1 - (sidebarButtonScale - 1) * 1)}px`;
  const sidebarButtonWidth = `${Math.max(84, Math.min(100, 92 + (sidebarButtonScale - 1) * 16)).toFixed(2)}%`;
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
  document.documentElement.style.setProperty("--accent-soft", shadeColor(state.settings.secondaryColor, 14));
  document.documentElement.style.setProperty("--sidebar-bg", state.settings.sidebarColor);
  document.body.style.setProperty("--sidebar-tab-bg", sidebarButtonBase);
  document.body.style.setProperty("--sidebar-tab-bg-hover", sidebarButtonHover);
  document.body.style.setProperty("--sidebar-tab-bg-active", sidebarButtonActive);
  document.body.style.setProperty("--sidebar-tab-border", sidebarButtonBorder);
  document.body.style.setProperty("--sidebar-tab-glow", sidebarButtonGlow);
  document.body.style.setProperty("--sidebar-tab-width", sidebarButtonWidth);
  document.body.dataset.sidebarButtonStyle = state.settings.sidebarButtonStyle || "style-1";
  document.documentElement.style.setProperty("--month-toolbar-bg", state.settings.toolbarColor);
  document.documentElement.style.setProperty("--month-toolbar-text", getReadableTextColor(state.settings.toolbarColor));
  const calculatorWidth = clampCalculatorWidth(state.settings.calculatorWidth);
  state.settings.calculatorWidth = calculatorWidth;
  document.documentElement.style.setProperty("--calculator-width", `${calculatorWidth}px`);
  tabs.forEach((tab) => {
    const tabId = tab.dataset.tab;
    if (!tabId) return;
    const tabColor = state.settings.sidebarTabColors?.[tabId] || DEFAULT_SIDEBAR_TAB_COLORS[tabId] || state.settings.primaryColor;
    tab.style.setProperty("color", tabColor, "important");
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
  applyMobileColumnWidths();
}

function applyMobileColumnWidths() {
  if (!(mobileTable instanceof HTMLElement)) return;
  const widths = state.settings.mobileColumnWidths || DEFAULT_MOBILE_COLUMN_WIDTHS;
  MOBILE_COLUMN_KEYS.forEach((key) => {
    const widthValue = Number(widths[key] || DEFAULT_MOBILE_COLUMN_WIDTHS[key] || 60);
    mobileTable.style.setProperty(`--mobile-col-${key}-width`, `${widthValue}px`);
  });
}

function initializeMobileColumnResizing() {
  const handles = document.querySelectorAll("[data-mobile-column-resize]");
  if (!handles.length) return;

  handles.forEach((handle) => {
    handle.addEventListener("pointerdown", (event) => {
      if (!(event.currentTarget instanceof HTMLElement)) return;
      const key = event.currentTarget.dataset.mobileColumnResize;
      if (!key) return;

      event.preventDefault();
      mobileColumnResizeState = {
        key,
        startX: event.clientX,
        startWidth: Number(state.settings.mobileColumnWidths?.[key] || DEFAULT_MOBILE_COLUMN_WIDTHS[key] || 60),
      };
      document.body.classList.add("is-resizing-mobile-columns");
      event.currentTarget.setPointerCapture?.(event.pointerId);
    });
  });
}

function updateMobileColumnResize(event) {
  if (!mobileColumnResizeState) return;
  if (!Number.isFinite(event.clientX)) return;

  const { key, startX, startWidth } = mobileColumnResizeState;
  const delta = event.clientX - startX;
  const nextWidth = clampMobileColumnWidth(key, startWidth + delta);
  state.settings.mobileColumnWidths = {
    ...DEFAULT_MOBILE_COLUMN_WIDTHS,
    ...(state.settings.mobileColumnWidths || {}),
    [key]: nextWidth,
  };
  applyMobileColumnWidths();
}

async function finishMobileColumnResize() {
  if (!mobileColumnResizeState) return;
  mobileColumnResizeState = null;
  document.body.classList.remove("is-resizing-mobile-columns");
  try {
    await persistState();
  } catch (error) {
    console.error("Falha ao guardar larguras das colunas móveis:", error);
  }
}

function clampMobileColumnWidth(key, value) {
  const minWidths = {
    data: 56,
    nome: 96,
    nif: 56,
    numero: 56,
    tarifario: 96,
    idVenda: 52,
    fidelizacao: 48,
    portabilidade: 48,
    bestOffer: 48,
    netSegura: 52,
    rceAceite: 48,
    actions: 78,
  };

  const minWidth = minWidths[key] || 48;
  return Math.max(minWidth, Math.min(260, Math.round(Number(value) || minWidth)));
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
