const express = require("express");
const fs = require("fs/promises");
const crypto = require("crypto");
const path = require("path");

const app = express();
const port = Number(process.env.PORT || 3000);
const dataFile = process.env.DATA_FILE || path.join(__dirname, "..", "data", "store.json");
const publicDir = path.join(__dirname, "..", "public");
const authUsername = process.env.AUTH_USERNAME || "admin";
const authPassword = process.env.AUTH_PASSWORD || "admin123";
const authSecret = process.env.AUTH_SECRET || "registo-mensal-secret";
const authCookieName = "registo_mensal_auth";
const authTtlMs = 1000 * 60 * 60 * 24 * 14;
const authPasswordHash = hashPassword(authPassword);
const accessServiceIds = [
  "sap-comercial",
  "portal-sfa",
  "user-rede",
  "pos-venda",
  "blueticket",
  "intelcia",
  "eform",
  "bec-deposito",
];

function clampCalculatorWidth(value) {
  return Math.min(460, Math.max(280, Number(value) || 360));
}

function buildDefaultAccessEntries(source = {}) {
  return Object.fromEntries(
    accessServiceIds.map((id) => {
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

const defaultState = {
  billingSales: [],
  mobileSales: [],
  energySales: [],
  tasks: [],
  stockEntries: [],
  accessEntries: buildDefaultAccessEntries(),
  monthlyIncentives: {},
  monthlyGoals: {},
  auth: {
    username: authUsername,
    passwordHash: authPasswordHash,
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
    toolbarColor: "#b3262d",
    sidebarTabColors: {
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
    },
    pageColors: {
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
    },
    pageBlockColors: {
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
    },
    pageFieldColors: {
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
    },
    calculatorWidth: 340,
    formWidth: 720,
    cardRadius: 24,
  },
};

app.use(express.json({ limit: "2mb" }));
app.use(express.static(publicDir));

app.get("/api/health", (_request, response) => {
  response.json({ ok: true });
});

app.get("/api/auth/status", async (request, response) => {
  const auth = await readAuthConfig();
  const session = getSessionFromRequest(request);
  response.json({
    authenticated: Boolean(session),
    username: auth.username,
  });
});

app.post("/api/auth/login", async (request, response) => {
  const username = String(request.body?.username || "");
  const password = String(request.body?.password || "");
  const auth = await readAuthConfig();

  if (username !== auth.username || hashPassword(password) !== auth.passwordHash) {
    return response.status(401).json({ ok: false, message: "Credenciais invalidas." });
  }

  const token = createSessionToken();
  response.setHeader("Set-Cookie", buildSessionCookie(token));
  response.json({ ok: true, username, token });
});

app.post("/api/auth/logout", (_request, response) => {
  response.setHeader("Set-Cookie", clearSessionCookie());
  response.json({ ok: true });
});

app.put("/api/auth/settings", async (request, response) => {
  const username = String(request.body?.username || "").trim();
  const currentPassword = String(request.body?.currentPassword || "");
  const password = String(request.body?.password || "");

  if (!username) {
    return response.status(400).json({ ok: false, message: "O utilizador não pode estar vazio." });
  }

  const state = await readState();
  const currentAuth = state.auth || defaultState.auth;
  const session = getSessionFromRequest(request);
  const currentPasswordMatches =
    currentPassword.trim().length > 0 && hashPassword(currentPassword) === currentAuth.passwordHash;
  const hostHeader = String(request.headers.host || "");
  const isLocalRequest =
    ["localhost", "127.0.0.1", "::1"].includes(request.hostname) ||
    ["localhost", "127.0.0.1", "::1"].some((value) => hostHeader.includes(value)) ||
    request.ip === "::1" ||
    request.ip === "127.0.0.1" ||
    request.ip === "::ffff:127.0.0.1";

  if (!session && !currentPasswordMatches && !isLocalRequest) {
    return response.status(401).json({
      ok: false,
      message: "Credenciais atuais inválidas.",
      debug: {
        hostname: request.hostname,
        host: hostHeader,
        ip: request.ip,
      },
    });
  }

  state.auth = {
    username,
    passwordHash:
      password.trim().length > 0 ? hashPassword(password) : currentAuth.passwordHash || authPasswordHash,
  };

  await fs.mkdir(path.dirname(dataFile), { recursive: true });
  await fs.writeFile(dataFile, JSON.stringify(state, null, 2), "utf8");
  response.json({ ok: true, username, passwordHash: state.auth.passwordHash });
});

app.get("/api/state", requireAuth, async (_request, response) => {
  const state = await readState();
  response.json(state);
});

app.put("/api/state", requireAuth, async (request, response) => {
  const nextState = sanitizeState(request.body);
  await fs.mkdir(path.dirname(dataFile), { recursive: true });
  await fs.writeFile(dataFile, JSON.stringify(nextState, null, 2), "utf8");
  response.json({ ok: true });
});

app.get("*", (_request, response) => {
  response.sendFile(path.join(publicDir, "index.html"));
});

app.listen(port, () => {
  console.log(`Servidor ativo em http://0.0.0.0:${port}`);
});

async function readState() {
  try {
    const raw = await fs.readFile(dataFile, "utf8");
    return sanitizeState(JSON.parse(raw));
  } catch (error) {
    if (error.code === "ENOENT") {
      return structuredClone(defaultState);
    }
    throw error;
  }
}

function requireAuth(request, response, next) {
  const session = getSessionFromRequest(request);
  if (session) {
    next();
    return;
  }

  response.status(401).json({ ok: false, message: "Nao autenticado." });
}

function getSessionFromRequest(request) {
  const tokens = getAuthTokensFromRequest(request);
  for (const token of tokens) {
    const session = validateSessionToken(token);
    if (session) return session;
  }

  return null;
}

function getAuthTokensFromRequest(request) {
  const tokens = [];
  const authorization = request.headers.authorization || request.headers.Authorization;
  if (typeof authorization === "string" && authorization.startsWith("Bearer ")) {
    const token = authorization.slice(7).trim();
    if (token) tokens.push(token);
  }

  const headerToken = request.headers["x-auth-token"];
  if (typeof headerToken === "string" && headerToken.trim()) {
    tokens.push(headerToken.trim());
  }

  const cookies = parseCookies(request.headers.cookie || "");
  const cookieToken = cookies[authCookieName];
  if (cookieToken) tokens.push(cookieToken);

  return tokens;
}

function validateSessionToken(token) {
  if (!token) return null;

  const parts = token.split(".");
  if (parts.length !== 3) return null;

  const [firstPart, secondPart, signature] = parts;

  if (/^\d+$/.test(firstPart)) {
    const expiresAt = Number(firstPart);
    if (expiresAt < Date.now()) return null;
    const expectedSignature = signSessionPayload(`${firstPart}.${secondPart}`);
    if (!safeCompare(signature, expectedSignature)) return null;
    return { authenticated: true };
  }

  const username = firstPart;
  const expiresAt = Number(secondPart);
  if (expiresAt < Date.now()) return null;

  const expectedSignature = signSessionPayload(`${username}.${secondPart}`);
  if (!safeCompare(signature, expectedSignature)) return null;

  return { authenticated: true };
}

function createSessionToken() {
  const expiresAt = String(Date.now() + authTtlMs);
  const nonce = crypto.randomBytes(16).toString("hex");
  const payload = `${expiresAt}.${nonce}`;
  const signature = signSessionPayload(payload);
  return `${payload}.${signature}`;
}

function signSessionPayload(payload) {
  return crypto.createHmac("sha256", authSecret).update(payload).digest("hex");
}

function hashPassword(password) {
  return crypto.pbkdf2Sync(password, "registo-mensal-salt", 120000, 32, "sha256").toString("hex");
}

async function readAuthConfig() {
  const state = await readState();
  return state.auth || defaultState.auth;
}

function safeCompare(left, right) {
  if (left.length !== right.length) return false;
  return crypto.timingSafeEqual(Buffer.from(left), Buffer.from(right));
}

function parseCookies(cookieHeader) {
  return cookieHeader.split(";").reduce((cookies, item) => {
    const [rawName, ...rest] = item.trim().split("=");
    if (!rawName || rest.length === 0) return cookies;
    cookies[rawName] = decodeURIComponent(rest.join("="));
    return cookies;
  }, {});
}

function buildSessionCookie(token) {
  return `${authCookieName}=${encodeURIComponent(token)}; HttpOnly; Path=/; SameSite=Lax; Max-Age=${Math.floor(authTtlMs / 1000)}`;
}

function clearSessionCookie() {
  return `${authCookieName}=; HttpOnly; Path=/; SameSite=Lax; Max-Age=0`;
}

function sanitizeState(input) {
  const state = input && typeof input === "object" ? input : {};
  const authInput = state.auth && typeof state.auth === "object" ? state.auth : {};

  return {
    billingSales: Array.isArray(state.billingSales) ? state.billingSales : [],
    mobileSales: Array.isArray(state.mobileSales) ? state.mobileSales : [],
    energySales: Array.isArray(state.energySales) ? state.energySales : [],
    tasks: Array.isArray(state.tasks) ? state.tasks : [],
    stockEntries: Array.isArray(state.stockEntries) ? state.stockEntries : [],
    accessEntries: buildDefaultAccessEntries(
      state.accessEntries && typeof state.accessEntries === "object" ? state.accessEntries : {}
    ),
    monthlyIncentives:
      state.monthlyIncentives && typeof state.monthlyIncentives === "object" && !Array.isArray(state.monthlyIncentives)
        ? state.monthlyIncentives
        : {},
    monthlyGoals:
      state.monthlyGoals && typeof state.monthlyGoals === "object" && !Array.isArray(state.monthlyGoals)
        ? state.monthlyGoals
        : {},
    auth: {
      username: String(authInput.username || authUsername).trim() || authUsername,
      passwordHash:
        typeof authInput.passwordHash === "string" && authInput.passwordHash
          ? authInput.passwordHash
          : authPasswordHash,
    },
    settings: {
      theme: state.settings?.theme === "dark" ? "dark" : "light",
      menuFontSize: Number(state.settings?.menuFontSize || 15),
      titleFontSize: Number(state.settings?.titleFontSize || 22),
      fieldFontSize: Number(state.settings?.fieldFontSize || 13),
      panelHeaderSize: Number(state.settings?.panelHeaderSize || 13),
      densityMode: state.settings?.densityMode === "normal" ? "normal" : "compact",
      primaryColor: state.settings?.primaryColor || "#b6f23b",
      secondaryColor: state.settings?.secondaryColor || "#1730a1",
      sidebarColor: state.settings?.sidebarColor || "#102061",
      sidebarButtonColor: state.settings?.sidebarButtonColor || "#20316f",
      toolbarColor: state.settings?.toolbarColor || "#b3262d",
      sidebarTabColors:
        state.settings?.sidebarTabColors && typeof state.settings.sidebarTabColors === "object"
          ? state.settings.sidebarTabColors
          : {
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
            },
      pageColors:
        state.settings?.pageColors && typeof state.settings.pageColors === "object"
          ? state.settings.pageColors
          : {
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
            },
      pageBlockColors:
        state.settings?.pageBlockColors && typeof state.settings.pageBlockColors === "object"
          ? state.settings.pageBlockColors
          : {
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
            },
      pageFieldColors:
        state.settings?.pageFieldColors && typeof state.settings.pageFieldColors === "object"
          ? state.settings.pageFieldColors
          : {
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
            },
      calculatorWidth: clampCalculatorWidth(state.settings?.calculatorWidth || 340),
      formWidth: Number(state.settings?.formWidth || 720),
      cardRadius: Number(state.settings?.cardRadius || 24),
    },
  };
}
