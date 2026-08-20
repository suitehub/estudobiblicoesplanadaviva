import {
  collection,
  addDoc,
  getDocs,
  updateDoc,
  deleteDoc,
  doc,
  setDoc,
  getDoc,
  query,
  where,
  limit,
  serverTimestamp
} from "https://www.gstatic.com/firebasejs/12.11.0/firebase-firestore.js";

import {
  getAuth,
  createUserWithEmailAndPassword,
  signOut
} from "https://www.gstatic.com/firebasejs/12.11.0/firebase-auth.js";

import {
  initializeApp,
  deleteApp
} from "https://www.gstatic.com/firebasejs/12.11.0/firebase-app.js";

import {
  exportExcelReport,
  exportPdfReport,
  filterInteressadosForReport,
  buildInstructorsData,
  resolveInteressadoFullData
} from "./reports.js";

import { app, db } from "./firebase-config.js";

const DISTRITO_FIXO = "Esplanada";

/* =========================
   ESTADO
========================= */
const state = {
  user: null,
  interessados: [],
  series: [],
  usuarios: [],
  locais: [],
  activeSection: "dashboardSection"
};

const statusOptions = [
  "Todos",
  "Ativo",
  "Pausado",
  "Desinteressado",
  "Concluído",
  "Pronto para apelo",
  "Pronto para batismo",
  "Batismo Realizado"
];

const interestOptions = ["Todos", "Alto", "Médio", "Baixo"];

/* =========================
   ELEMENTOS
========================= */
const pageTitle = document.getElementById("pageTitle");
const globalMessage = document.getElementById("globalMessage");
const loadingOverlay = document.getElementById("loadingOverlay");

const metricsGrid = document.getElementById("metricsGrid");
const recentList = document.getElementById("recentList");
const attentionList = document.getElementById("attentionList");

const searchInput = document.getElementById("searchInput");
const statusFilter = document.getElementById("statusFilter");
const interestFilter = document.getElementById("interestFilter");
const serieFilter = document.getElementById("serieFilter");

const interessadosTable = document.getElementById("interessadosTable");
const seriesList = document.getElementById("seriesList");
const usuariosList = document.getElementById("usuariosList");
const locaisList = document.getElementById("locaisList");

const interessadoModalBackdrop = document.getElementById("interessadoModalBackdrop");
const closeInteressadoModalBtn = document.getElementById("closeInteressadoModalBtn");
const interessadoModalTitle = document.getElementById("interessadoModalTitle");

const interessadoForm = document.getElementById("interessadoForm");
const interessadoId = document.getElementById("interessadoId");
const nome = document.getElementById("nome");
const telefone = document.getElementById("telefone");
const endereco = document.getElementById("endereco");
const igreja = document.getElementById("igreja");
const distrito = document.getElementById("distrito");
const instrutoresContainer = document.getElementById("instrutoresContainer");
const addInstrutorBtn = document.getElementById("addInstrutorBtn");
const serieId = document.getElementById("serieId");
const estudoAtual = document.getElementById("estudoAtual");
const progressoPreview = document.getElementById("progressoPreview");
const status = document.getElementById("status");
const interesse = document.getElementById("interesse");
const ultimoContato = document.getElementById("ultimoContato");
const obsRapidas = document.getElementById("obsRapidas");
const observacoes = document.getElementById("observacoes");
const resetFormBtn = document.getElementById("resetFormBtn");

const openInteressadoModalBtn = document.getElementById("openInteressadoModalBtn");
const openInteressadoModalBtn2 = document.getElementById("openInteressadoModalBtn2");

const serieForm = document.getElementById("serieForm");
const novaSerieNome = document.getElementById("novaSerieNome");
const novaSerieTotal = document.getElementById("novaSerieTotal");

const localForm = document.getElementById("localForm");
const novoLocalNome = document.getElementById("novoLocalNome");
const novoLocalTipo = document.getElementById("novoLocalTipo");

const userForm = document.getElementById("userForm");
const novoUsuarioNome = document.getElementById("novoUsuarioNome");
const novoUsuarioPerfil = document.getElementById("novoUsuarioPerfil");
const novoUsuarioIgreja = document.getElementById("novoUsuarioIgreja");
const novoUsuarioDistrito = document.getElementById("novoUsuarioDistrito");
const novoUsuarioUsername = document.getElementById("novoUsuarioUsername");
const novoUsuarioSenhaLocal = document.getElementById("novoUsuarioSenhaLocal");
const novoUsuarioEmail = document.getElementById("novoUsuarioEmail");
const novoUsuarioSenha = document.getElementById("novoUsuarioSenha");

const novoUsuarioUsernameGroup = document.getElementById("novoUsuarioUsernameGroup");
const novoUsuarioSenhaLocalGroup = document.getElementById("novoUsuarioSenhaLocalGroup");
const novoUsuarioEmailGroup = document.getElementById("novoUsuarioEmailGroup");
const novoUsuarioSenhaGroup = document.getElementById("novoUsuarioSenhaGroup");

/* =========================
   HELPERS
========================= */
function showLoading() {
  loadingOverlay?.classList.remove("hidden");
}

function hideLoading() {
  loadingOverlay?.classList.add("hidden");
}

function showMessage(message, type = "info") {
  if (!globalMessage) return;
  globalMessage.textContent = message;
  globalMessage.classList.remove("hidden", "success", "error", "info");
  globalMessage.classList.add(type);
}

function clearMessage() {
  if (!globalMessage) return;
  globalMessage.textContent = "";
  globalMessage.classList.add("hidden");
  globalMessage.classList.remove("success", "error", "info");
}

function escapeHtml(value = "") {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function normalizeText(text = "") {
  return String(text)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();
}

function slug(text = "") {
  return normalizeText(text).replace(/\s+/g, "-");
}

function isAdmin() {
  return state.user?.perfil === "admin";
}

function isDistrital() {
  return state.user?.perfil === "distrital";
}

function isLocal() {
  return state.user?.perfil === "local";
}

function isMembro() {
  return state.user?.perfil === "membro";
}

function canManageSeries() {
  return isAdmin();
}

function canManageLocais() {
  return isAdmin();
}

function canSeeUsersSection() {
  return isAdmin() || isDistrital() || isLocal();
}

function canCreateUsers() {
  return isAdmin() || isDistrital() || isLocal();
}

function getAllowedProfilesToCreate() {
  if (isAdmin()) return ["admin", "distrital", "local", "membro"];
  if (isDistrital()) return ["local", "membro"];
  if (isLocal()) return ["membro"];
  return [];
}

function canEditUser(targetUser) {
  const currentUser = state.user;

  if (!currentUser || !targetUser) return false;

  if (isAdmin()) return true;

  if (isDistrital()) {
    return ["local", "membro"].includes(targetUser.perfil);
  }

  if (isLocal()) {
    return targetUser.perfil === "membro";
  }

  return false;
}

function canDeleteUser(targetUser) {
  const currentUser = state.user;

  if (!currentUser || !targetUser) return false;

  // não permitir auto-exclusão
  if (targetUser.id === currentUser.uid) return false;

  if (isAdmin()) return true;

  if (isDistrital()) {
    return ["local", "membro"].includes(targetUser.perfil);
  }

  if (isLocal()) {
    return targetUser.perfil === "membro";
  }

  return false;
}

function formatRole(role = "") {
  switch (role) {
    case "admin":
      return "Administrador";
    case "distrital":
      return "Líder Distrital";
    case "local":
      return "Líder Local";
    case "membro":
      return "Membro";
    default:
      return role || "-";
  }
}

function getInstrutorNames(item) {
  if (!item) return [];
  const resolved = resolveInteressadoFullData(item, state);
  return resolved.instrutorNomes || [];
}

function getInstrutorNamesText(item) {
  if (!item) return "-";
  const resolved = resolveInteressadoFullData(item, state);
  return resolved.instrutoresString || "-";
}

function userCanManageInteressado(item) {
  if (isAdmin() || isDistrital()) return true;
  if (isLocal()) return item.igrejaId === getCurrentUserLocalId();
  if (isMembro()) return item.criadoPorId === state.user.uid;
  return false;
}

function canAddStudyToInteressado(item) {
  return userCanManageInteressado(item) && Number(item.estudoAtual || 0) < Number(item.totalEstudos || 0);
}

function calcProgress(estudoAtualValue, totalEstudosValue) {
  const atual = Math.max(0, Number(estudoAtualValue || 0));
  const total = Math.max(1, Number(totalEstudosValue || 1));
  const capped = Math.min(atual, total);
  const porcentagem = Math.round((capped / total) * 100);
  const faltantes = Math.max(0, total - capped);
  return { capped, total, porcentagem, faltantes };
}

function daysSince(dateStr) {
  if (!dateStr) return 999;
  const today = new Date();
  const date = new Date(`${dateStr}T00:00:00`);
  const diff = today - date;
  return Math.floor(diff / (1000 * 60 * 60 * 24));
}

function isAtRisk(item) {
  const activeStudyStatuses = [
    "Ativo",
    "Pronto para apelo",
    "Pronto para batismo"
  ];

  return (
    activeStudyStatuses.includes(item.status) &&
    daysSince(item.ultimoContato) > 14
  );
}

function isDecision(item) {
  return (
    item.status === "Pronto para apelo" ||
    item.status === "Pronto para batismo"
  );
}

/* =========================
   DASHBOARD: PROGRESSO ATUAL
========================= */
function getProgressBuckets(data = []) {
  const buckets = {
    "0–25%": 0,
    "26–50%": 0,
    "51–75%": 0,
    "76–99%": 0,
    "100%": 0
  };

  data.forEach((item) => {
    const value = Number(item.porcentagem || 0);

    if (value <= 25) {
      buckets["0–25%"] += 1;
    } else if (value <= 50) {
      buckets["26–50%"] += 1;
    } else if (value <= 75) {
      buckets["51–75%"] += 1;
    } else if (value < 100) {
      buckets["76–99%"] += 1;
    } else {
      buckets["100%"] += 1;
    }
  });

  return buckets;
}

function getAverageProgressValue(data = []) {
  if (!data.length) return 0;

  const total = data.reduce((sum, item) => {
    return sum + Number(item.porcentagem || 0);
  }, 0);

  return Math.round(total / data.length);
}

function getSpiritualReadinessValue(statusText = "") {
  const map = {
  "Desinteressado": 10,
  "Pausado": 30,
  "Ativo": 55,
  "Concluído": 75,
  "Pronto para apelo": 90,
  "Pronto para batismo": 100,
  "Batismo Realizado": 100
};

  return map[statusText] ?? 0;
}

function getAverageSpiritualReadinessValue(data = []) {
  if (!data.length) return 0;

  const total = data.reduce((sum, item) => {
    return sum + getSpiritualReadinessValue(item.status);
  }, 0);

  return Math.round(total / data.length);
}

function getProgressInsightLabel(avg) {
  if (avg >= 85) return "Grupo muito avançado";
  if (avg >= 65) return "Bom avanço geral";
  if (avg >= 45) return "Avanço moderado";
  if (avg >= 25) return "Base ainda inicial";
  return "Progresso muito inicial";
}

function getReadinessInsightLabel(avg) {
  if (avg >= 85) return "Prontidão espiritual muito alta";
  if (avg >= 65) return "Prontidão espiritual consistente";
  if (avg >= 45) return "Prontidão em desenvolvimento";
  if (avg >= 25) return "Prontidão ainda frágil";
  return "Baixa prontidão no momento";
}

function getMainStatusCounts(data = []) {
  const counts = {
    ativos: 0,
    apelo: 0,
    batismo: 0,
    concluidos: 0,
    batismoRealizado: 0
  };

  data.forEach((item) => {
    if (item.status === "Ativo") counts.ativos += 1;
    if (item.status === "Pronto para apelo") counts.apelo += 1;
    if (item.status === "Pronto para batismo") counts.batismo += 1;
    if (item.status === "Concluído") counts.concluidos += 1;
    if (item.status === "Batismo Realizado") counts.batismoRealizado += 1;
  });

  return counts;
}

function getMainStatusPercentages(data = []) {
  const total = data.length || 1;
  const counts = getMainStatusCounts(data);

  return {
    total: data.length,
    ativos: Math.round((counts.ativos / total) * 100),
    apelo: Math.round((counts.apelo / total) * 100),
    batismo: Math.round((counts.batismo / total) * 100),
    concluidos: Math.round((counts.concluidos / total) * 100),
    batismoRealizado: Math.round((counts.batismoRealizado / total) * 100)
  };
}

function getChurchGroupedData(data = []) {
  const churchMap = new Map();

  data.forEach((item) => {
    const churchId = item.igrejaId || "sem-igreja";
    const churchName = item.igrejaNome || "Sem igreja";

    if (!churchMap.has(churchId)) {
      churchMap.set(churchId, {
        igrejaId: churchId,
        igrejaNome: churchName,
        items: []
      });
    }

    churchMap.get(churchId).items.push(item);
  });

  return Array.from(churchMap.values()).sort((a, b) =>
    a.igrejaNome.localeCompare(b.igrejaNome, "pt-BR")
  );
}

function renderStatusSummaryChart(percentages) {
  return `
    <div class="church-status-chart">
      <div class="church-status-row">
        <div class="church-status-head">
          <span class="church-status-label">Ativos</span>
          <strong class="church-status-value">${percentages.ativos}%</strong>
        </div>
        <div class="church-status-bar">
          <span class="status-ativo-fill" style="width:${percentages.ativos}%"></span>
        </div>
      </div>

      <div class="church-status-row">
        <div class="church-status-head">
          <span class="church-status-label">Apelo</span>
          <strong class="church-status-value">${percentages.apelo}%</strong>
        </div>
        <div class="church-status-bar">
          <span class="status-apelo-fill" style="width:${percentages.apelo}%"></span>
        </div>
      </div>

      <div class="church-status-row">
        <div class="church-status-head">
          <span class="church-status-label">Pronto p/ batismo</span>
          <strong class="church-status-value">${percentages.batismo}%</strong>
        </div>
        <div class="church-status-bar">
          <span class="status-batismo-fill" style="width:${percentages.batismo}%"></span>
        </div>
      </div>

      <div class="church-status-row">
        <div class="church-status-head">
          <span class="church-status-label">Concluídos</span>
          <strong class="church-status-value">${percentages.concluidos}%</strong>
        </div>
        <div class="church-status-bar">
          <span class="status-concluido-fill" style="width:${percentages.concluidos}%"></span>
        </div>
      </div>

      <div class="church-status-row">
        <div class="church-status-head">
          <span class="church-status-label">Batismo realizado</span>
          <strong class="church-status-value">${percentages.batismoRealizado}%</strong>
        </div>
        <div class="church-status-bar">
          <span class="status-batismo-realizado-fill" style="width:${percentages.batismoRealizado}%"></span>
        </div>
      </div>
    </div>
  `;
}

function renderDashboardScales() {
  if (!dashboardScales) return;

  const data = getVisibleInteressados();
  const progressoMedio = getAverageProgressValue(data);
  const prontidaoMedia = getAverageSpiritualReadinessValue(data);
  const buckets = getProgressBuckets(data);

  if (!data.length) {
    dashboardScales.innerHTML = `
      <div class="card">
        <div class="card-header">
          <div>
            <h3>Painel do progresso e status</h3>
            <p>Distribuição dos interessados e panorama espiritual atual.</p>
          </div>
        </div>
        <div class="empty-state">Ainda não há interessados suficientes para exibir o painel.</div>
      </div>
    `;
    return;
  }

  const maxValue = Math.max(...Object.values(buckets), 1);

  const barsHtml = Object.entries(buckets)
    .map(([label, value]) => {
      const height = Math.max((value / maxValue) * 220, value > 0 ? 22 : 10);

      return `
        <div class="progress-chart-col">
          <div class="progress-chart-value">${value}</div>
          <div class="progress-chart-bar-wrap">
            <div class="progress-chart-bar" style="height:${height}px;"></div>
          </div>
          <div class="progress-chart-label">${label}</div>
        </div>
      `;
    })
    .join("");

  const geral = getMainStatusPercentages(data);

  const districtOrAdminView = isAdmin() || isDistrital();
  const churches = getChurchGroupedData(data);

  const churchCardsHtml = districtOrAdminView
    ? `
      <div class="card" style="margin-top:18px;">
        <div class="card-header">
          <div>
            <h3>Panorama por igreja</h3>
            <p>Total de estudos bíblicos e percentuais dos principais níveis por igreja.</p>
          </div>
        </div>

        <div class="stack-list">
          ${churches
            .map((church) => {
              const percentages = getMainStatusPercentages(church.items);

              return `
                <article class="stack-item">
                  <div class="stack-item-top">
                    <h4
  class="dashboard-church-link"
  data-igreja-id="${church.igrejaId}"
>
  ${escapeHtml(church.igrejaNome)}
</h4>
                    <span class="tiny-muted">${percentages.total} estudos bíblicos</span>
                  </div>

                  <p class="stack-item-sub">
                    Visão percentual atual dos principais status nesta igreja.
                  </p>

                  ${renderStatusSummaryChart(percentages)}
                </article>
              `;
            })
            .join("")}
        </div>
      </div>
    `
    : `
      <div class="card" style="margin-top:18px;">
        <div class="card-header">
          <div>
            <h3>Panorama da igreja local</h3>
            <p>Total de estudos bíblicos e percentuais dos principais níveis da sua igreja.</p>
          </div>
        </div>

        <article class="stack-item">
          <div class="stack-item-top">
            <h4>${escapeHtml(getCurrentUserLocalName() || "Igreja local")}</h4>
            <span class="tiny-muted">${geral.total} estudos bíblicos</span>
          </div>

          <p class="stack-item-sub">
            Visão percentual atual dos principais status da igreja local.
          </p>

          ${renderStatusSummaryChart(geral)}
        </article>
      </div>
    `;

  dashboardScales.innerHTML = `
    <div class="card">
      <div class="card-header card-header-wrap">
        <div>
          <h3>Panorama atual do progresso nos estudos</h3>
          <p>Distribuição dos interessados por faixa de avanço no momento atual.</p>
        </div>
      </div>

      <div class="progress-summary-grid">
        <div class="progress-summary-card">
          <span class="progress-summary-kicker">Progresso médio</span>
          <strong class="progress-summary-value">${progressoMedio}%</strong>
          <p class="progress-summary-text">${getProgressInsightLabel(progressoMedio)}</p>
        </div>

        <div class="progress-summary-card">
          <span class="progress-summary-kicker">Prontidão espiritual</span>
          <strong class="progress-summary-value">${prontidaoMedia}%</strong>
          <p class="progress-summary-text">${getReadinessInsightLabel(prontidaoMedia)}</p>
        </div>
      </div>

      <div class="progress-chart-shell">
        <div class="progress-chart-y-axis">
          <span>${maxValue}</span>
          <span>${Math.round(maxValue * 0.75)}</span>
          <span>${Math.round(maxValue * 0.5)}</span>
          <span>${Math.round(maxValue * 0.25)}</span>
          <span>0</span>
        </div>

        <div class="progress-chart-area">
          <div class="progress-chart-grid">
            <span></span>
            <span></span>
            <span></span>
            <span></span>
            <span></span>
          </div>

          <div class="progress-chart-cols">
            ${barsHtml}
          </div>
        </div>
      </div>
    </div>

    <div class="card" style="margin-top:18px;">
      <div class="card-header">
        <div>
          <h3>${districtOrAdminView ? "Panorama geral do distrito" : "Panorama geral da igreja"}</h3>
          <p>Percentuais atuais dos principais níveis espirituais.</p>
        </div>
      </div>

      <article class="stack-item">
        <div class="stack-item-top">
          <h4>${districtOrAdminView ? "Distrito Esplanada" : escapeHtml(getCurrentUserLocalName() || "Igreja local")}</h4>
          <span class="tiny-muted">${geral.total} estudos bíblicos</span>
        </div>

        <p class="stack-item-sub">
          Percentuais principais calculados com base nos estudos atualmente registrados.
        </p>

        ${renderStatusSummaryChart(geral)}
      </article>
    </div>

    ${churchCardsHtml}
  `;
bindDashboardChurchLinks();
}

function bindDashboardChurchLinks() {
  document
    .querySelectorAll(".dashboard-church-link")
    .forEach((element) => {

      element.addEventListener("click", () => {

        const igrejaId =
          element.dataset.igrejaId;

        if (!igrejaId) return;

        serieFilter.value = "Todas";
        statusFilter.value = "Todos";
        interestFilter.value = "Todos";

        setSection("interessadosSection");

        const filtered =
          state.interessados.filter(
            item => item.igrejaId === igrejaId
          );

        renderInteressadosTable(filtered);
      });
    });
}

function sortByUpdatedDesc(arr) {
  return [...arr].sort((a, b) => {
    const aTime = new Date(a.atualizadoEm || a.updatedAt || 0).getTime();
    const bTime = new Date(b.atualizadoEm || b.updatedAt || 0).getTime();
    return bTime - aTime;
  });
}

function sortByName(arr) {
  return [...arr].sort((a, b) =>
    (a.nome || "").localeCompare(b.nome || "", "pt-BR")
  );
}

function getSerieById(id) {
  return state.series.find((s) => s.id === id);
}

function getLocalById(id) {
  return state.locais.find((l) => l.id === id);
}

function getUserById(id) {
  return state.usuarios.find((u) => u.id === id);
}

function getCurrentUserLocalId() {
  return state.user?.igrejaId || "";
}

function getCurrentUserLocalName() {
  return state.user?.igrejaNome || "";
}

function getSectionTitle(sectionId) {
  const map = {
    dashboardSection: "Dashboard",
    interessadosSection: "Interessados",
    instrutoresSection: "Instrutores",
    usuariosSection: "Usuários",
    locaisSection: "Locais",
    catalogoSection: "Catálogo de Estudos",
    relatoriosSection: "Relatórios & Exportações"
  };
  return map[sectionId] || "Esplanada Viva";
}

function getAllowedSections() {
  const base = ["dashboardSection", "interessadosSection", "instrutoresSection", "relatoriosSection"];
  if (canSeeUsersSection()) base.push("usuariosSection");
  if (canManageLocais()) base.push("locaisSection");
  if (canManageSeries()) base.push("catalogoSection");
  return base;
}

function setSection(sectionId) {
  const allowed = getAllowedSections();
  const safeSection = allowed.includes(sectionId) ? sectionId : "dashboardSection";

  state.activeSection = safeSection;

  document.querySelectorAll(".page-section").forEach((section) => {
    section.classList.toggle("active", section.id === safeSection);
  });

  document.querySelectorAll(".nav-btn").forEach((btn) => {
    btn.classList.toggle("active", btn.dataset.section === safeSection);
  });

  const topbar = document.querySelector(".topbar");
  if (topbar) {
    topbar.style.display = safeSection === "instrutoresSection" ? "none" : "";
  }

  if (pageTitle) {
    pageTitle.textContent = getSectionTitle(safeSection);
  }

  if (safeSection === "instrutoresSection") {
    renderInstrutoresSection();
  }

  if (safeSection === "relatoriosSection") {
    renderRelatoriosSection();
  }
}

function applyRoleUI() {
  document.querySelectorAll(".nav-btn").forEach((btn) => {
    const section = btn.dataset.section;
    const allowed = getAllowedSections().includes(section);
    btn.style.display = allowed ? "" : "none";
  });

  document.querySelectorAll(".admin-only").forEach((el) => {
    el.style.display = isAdmin() ? "" : "none";
  });

  if (!getAllowedSections().includes(state.activeSection)) {
    state.activeSection = "dashboardSection";
  }

  setSection(state.activeSection);
}

function buildEmailAuth(username, igrejaId) {
  const safeUsername = normalizeText(username).replace(/[^a-z0-9._-]/g, "");
  const safeChurch = String(igrejaId || "").replace(/[^a-zA-Z0-9]/g, "");
  return `${safeUsername}__${safeChurch}@app.esplanadaviva.local`;
}

function mapUserCreationError(error) {
  const code = error?.code || "";

  switch (code) {
    case "auth/email-already-in-use":
      return "Esse acesso já existe.";
    case "auth/invalid-email":
      return "E-mail inválido.";
    case "auth/weak-password":
      return "A senha precisa ter pelo menos 6 caracteres.";
    default:
      return error?.message || "Não foi possível criar o usuário.";
  }
}

/* =========================
   FIRESTORE LOAD
========================= */
async function loadSeries() {
  const snap = await getDocs(collection(db, "series"));
  state.series = snap.docs.map((item) => ({
    id: item.id,
    ...item.data()
  }));
}

async function loadLocais() {
  const snap = await getDocs(collection(db, "locais"));
  state.locais = snap.docs
    .map((item) => ({
      id: item.id,
      ...item.data()
    }))
    .filter((item) => item.ativo !== false);
}

async function loadUsuarios() {
  let snap;

  if (isAdmin() || isDistrital()) {
    snap = await getDocs(collection(db, "usuarios"));
    state.usuarios = snap.docs
      .map((item) => ({
        id: item.id,
        ...item.data()
      }))
      .filter((item) => item.ativo !== false);
    return;
  }

  if (isLocal()) {
    const localId = getCurrentUserLocalId();
    snap = await getDocs(
      query(collection(db, "usuarios"), where("igrejaId", "==", localId))
    );
    state.usuarios = snap.docs
      .map((item) => ({
        id: item.id,
        ...item.data()
      }))
      .filter((item) => item.ativo !== false);
    return;
  }

  if (isMembro()) {
    state.usuarios = state.user ? [state.user] : [];
    return;
  }

  state.usuarios = [];
}

async function loadInteressados() {
  let snap;

  if (isAdmin()) {
    snap = await getDocs(collection(db, "interessados"));
    state.interessados = snap.docs.map((item) => ({
      id: item.id,
      ...item.data()
    }));
    return;
  }

  if (isDistrital()) {
    const q = query(
      collection(db, "interessados"),
      where("distrito", "==", DISTRITO_FIXO)
    );

    snap = await getDocs(q);
    state.interessados = snap.docs.map((item) => ({
      id: item.id,
      ...item.data()
    }));
    return;
  }

  if (isLocal()) {
    const q = query(
      collection(db, "interessados"),
      where("igrejaId", "==", getCurrentUserLocalId())
    );

    snap = await getDocs(q);
    state.interessados = snap.docs.map((item) => ({
      id: item.id,
      ...item.data()
    }));
    return;
  }

  if (isMembro()) {
    const q = query(
      collection(db, "interessados"),
      where("criadoPorId", "==", state.user.uid)
    );

    snap = await getDocs(q);
    state.interessados = snap.docs.map((item) => ({
      id: item.id,
      ...item.data()
    }));
    return;
  }

  state.interessados = [];
}

async function refreshData() {
  if (!state.user) return;

  showLoading();
  clearMessage();

  try {
    await Promise.all([
      loadLocais(),
      loadSeries(),
      loadUsuarios(),
      loadInteressados()
    ]);

    renderAll();
  } catch (error) {
    console.error("Erro ao carregar dados:", error);
    showMessage("Não foi possível carregar os dados do sistema.", "error");
  } finally {
    hideLoading();
  }
}

/* =========================
   FILTROS
========================= */
function getVisibleInteressados() {
  return sortByUpdatedDesc(state.interessados);
}

function getFilteredInteressados() {
  let data = getVisibleInteressados();

  const search = normalizeText(searchInput?.value || "");
  const selectedStatus = statusFilter?.value || "Todos";
  const selectedInterest = interestFilter?.value || "Todos";
  const selectedSerie = serieFilter?.value || "Todas";

  if (search) {
    data = data.filter((item) => {
      const instrutoresText = getInstrutorNamesText(item);
      const allInstNames = getInstrutorNames(item)
        .map((n) => normalizeText(n))
        .join(" ");

      return (
        normalizeText(item.nome || "").includes(search) ||
        normalizeText(item.telefone || "").includes(search) ||
        normalizeText(item.endereco || "").includes(search) ||
        normalizeText(item.igrejaNome || "").includes(search) ||
        normalizeText(item.distrito || "").includes(search) ||
        normalizeText(instrutoresText).includes(search) ||
        allInstNames.includes(search) ||
        normalizeText(item.serieNome || "").includes(search) ||
        normalizeText(item.status || "").includes(search) ||
        normalizeText(item.interesse || "").includes(search) ||
        normalizeText(item.observacoes || "").includes(search) ||
        normalizeText(item.criadoPorNome || "").includes(search)
      );
    });
  }

  if (selectedStatus && selectedStatus !== "Todos") {
    data = data.filter((item) => item.status === selectedStatus);
  }

  if (selectedInterest && selectedInterest !== "Todos") {
    data = data.filter((item) => item.interesse === selectedInterest);
  }

  if (selectedSerie && selectedSerie !== "Todas") {
    data = data.filter((item) => item.serieId === selectedSerie);
  }

  return data;
}

/* =========================
   RENDERS BÁSICOS
========================= */
function renderStatusAndInterestOptions() {
  status.innerHTML = statusOptions
    .filter((item) => item !== "Todos")
    .map((item) => `<option value="${escapeHtml(item)}">${escapeHtml(item)}</option>`)
    .join("");

  interesse.innerHTML = interestOptions
    .filter((item) => item !== "Todos")
    .map((item) => `<option value="${escapeHtml(item)}">${escapeHtml(item)}</option>`)
    .join("");
}

function renderFilters() {
  if (statusFilter) {
    const current = statusFilter.value || "Todos";
    statusFilter.innerHTML = statusOptions
      .map((item) => `<option value="${escapeHtml(item)}">${escapeHtml(item)}</option>`)
      .join("");
    statusFilter.value = current;
  }

  if (interestFilter) {
    const current = interestFilter.value || "Todos";
    interestFilter.innerHTML = interestOptions
      .map((item) => `<option value="${escapeHtml(item)}">${escapeHtml(item)}</option>`)
      .join("");
    interestFilter.value = current;
  }

  if (serieFilter) {
    const current = serieFilter.value || "Todas";
    serieFilter.innerHTML =
      `<option value="Todas">Todas as séries</option>` +
      sortByName(state.series)
        .map((serie) => `<option value="${serie.id}">${escapeHtml(serie.nome)}</option>`)
        .join("");
    serieFilter.value = current;
  }
}

function renderSerieOptions() {
  if (!serieId) return;

  const current = serieId.value || "";
  serieId.innerHTML =
    `<option value="">Selecionar série</option>` +
    sortByName(state.series)
      .map((serie) => `<option value="${serie.id}">${escapeHtml(serie.nome)} (${serie.totalEstudos})</option>`)
      .join("");

  if (current) {
    serieId.value = current;
  }
}

function renderLocalOptions(targetSelect, { onlyCurrent = false, includeEmpty = true } = {}) {
  if (!targetSelect) return;

  let locais = [...state.locais];

  if (onlyCurrent) {
    const localId = getCurrentUserLocalId();
    locais = locais.filter((item) => item.id === localId);
  }

  locais = sortByName(locais);

  targetSelect.innerHTML =
    `${includeEmpty ? '<option value="">Selecionar igreja</option>' : ""}` +
    locais.map((item) => `<option value="${item.id}">${escapeHtml(item.nome)}</option>`).join("");
}

function renderInteressadoLocalOptions(selectedId = "") {
  if (!igreja) return;

  if (isLocal() || isMembro()) {
    renderLocalOptions(igreja, { onlyCurrent: true });
    igreja.value = getCurrentUserLocalId();
    igreja.disabled = true;
    return;
  }

  renderLocalOptions(igreja);
  igreja.disabled = false;

  if (selectedId) {
    igreja.value = selectedId;
  }
}

function getResponsibleUsersForCurrentContext() {
  if (isAdmin() || isDistrital()) {
    return sortByName(state.usuarios.filter((u) => u.ativo !== false));
  }

  if (isLocal()) {
    const localId = getCurrentUserLocalId();
    return sortByName(
      state.usuarios.filter((u) => u.ativo !== false && u.igrejaId === localId)
    );
  }

  if (isMembro()) {
    return state.user ? [state.user] : [];
  }

  return [];
}

function buildInstrutorOptionsHtml(selectedId = "") {
  const users = getResponsibleUsersForCurrentContext();

  return (
    `<option value="">Selecionar instrutor</option>` +
    users
      .map((u) => {
        const selected = u.id === selectedId ? "selected" : "";
        return `<option value="${u.id}" ${selected}>${escapeHtml(u.nome)}</option>`;
      })
      .join("")
  );
}

function updateInstrutorRowsUI() {
  if (!instrutoresContainer) return;

  const rows = Array.from(instrutoresContainer.querySelectorAll(".instrutor-row"));
  const isSingleLocked = isMembro();

  rows.forEach((row, index) => {
    const removeBtn = row.querySelector(".remove-instrutor-btn");
    const select = row.querySelector(".instrutor-select");

    if (removeBtn) {
      removeBtn.style.display =
        rows.length > 1 && !isSingleLocked ? "" : "none";
    }

    if (select) {
      select.disabled = isSingleLocked;
    }

    if (isSingleLocked && index > 0) {
      row.remove();
    }
  });
}

function createInstrutorRow(selectedId = "") {
  const row = document.createElement("div");
  row.className = "instrutor-row";

  const select = document.createElement("select");
  select.className = "instrutor-select";
  select.required = true;
  select.innerHTML = buildInstrutorOptionsHtml(selectedId);

  const removeBtn = document.createElement("button");
  removeBtn.type = "button";
  removeBtn.className = "btn btn-secondary btn-sm remove-instrutor-btn";
  removeBtn.textContent = "Remover";

  removeBtn.addEventListener("click", () => {
    row.remove();
    updateInstrutorRowsUI();
  });

  row.appendChild(select);
  row.appendChild(removeBtn);

  return row;
}

function renderResponsavelOptions(selectedIds = []) {
  if (!instrutoresContainer) return;

  instrutoresContainer.innerHTML = "";

  const safeIds = Array.isArray(selectedIds) && selectedIds.length
    ? selectedIds
    : [""];

  safeIds.forEach((id) => {
    instrutoresContainer.appendChild(createInstrutorRow(id));
  });

  if (isMembro()) {
    const firstSelect = instrutoresContainer.querySelector(".instrutor-select");
    if (firstSelect) {
      firstSelect.value = state.user?.uid || "";
    }
  }

  updateInstrutorRowsUI();
}

function getInstrutoresFromForm() {
  if (!instrutoresContainer) return [];

  const selects = Array.from(
    instrutoresContainer.querySelectorAll(".instrutor-select")
  );

  const ids = selects
    .map((select) => select.value)
    .filter(Boolean);

  const uniqueIds = [...new Set(ids)];

  const users = uniqueIds
    .map((id) => getUserById(id))
    .filter(Boolean)
    .map((user) => ({
      id: user.id,
      nome: user.nome,
      perfil: user.perfil
    }));

  return users;
}

function updateProgressPreview() {
  const serie = getSerieById(serieId.value);
  const total = serie ? Number(serie.totalEstudos) : 0;
  const atual = Number(estudoAtual.value || 0);
  const progress = calcProgress(atual, total || 1);

  if (!total) {
    progressoPreview.textContent = "0 de 0 • 0% • faltam 0";
    return;
  }

  progressoPreview.textContent =
    `${Math.min(atual, total)} de ${total} • ${progress.porcentagem}% • faltam ${Math.max(0, total - Math.min(atual, total))}`;
}

/* =========================
   MODAL INTERESSADO
========================= */
function openInteressadoModal() {
  interessadoModalBackdrop?.classList.remove("hidden");
}

function closeInteressadoModal() {
  interessadoModalBackdrop?.classList.add("hidden");
}

function resetInteressadoForm() {
  interessadoForm?.reset();
  interessadoId.value = "";
  interessadoModalTitle.textContent = "Novo interessado";

  renderInteressadoLocalOptions();
  renderResponsavelOptions();
  renderSerieOptions();
  renderStatusAndInterestOptions();

  if (distrito) {
    distrito.value = DISTRITO_FIXO;
    distrito.readOnly = true;
  }

  status.value = "Ativo";
  interesse.value = "Médio";
  ultimoContato.value = new Date().toISOString().slice(0, 10);
  estudoAtual.value = 0;

  updateProgressPreview();
}

function fillInteressadoForm(item) {
  interessadoId.value = item.id;
  nome.value = item.nome || "";
  telefone.value = item.telefone || "";
  endereco.value = item.endereco || "";

  renderInteressadoLocalOptions(item.igrejaId || "");
  if (!igreja.disabled) {
    igreja.value = item.igrejaId || "";
  }

  renderResponsavelOptions(item.instrutorIds || []);
  renderSerieOptions();
  serieId.value = item.serieId || "";

  estudoAtual.value = item.estudoAtual ?? 0;
  renderStatusAndInterestOptions();
  status.value = item.status || "Ativo";
  interesse.value = item.interesse || "Médio";
  ultimoContato.value = item.ultimoContato || "";
  observacoes.value = item.observacoes || "";

  if (distrito) {
    distrito.value = DISTRITO_FIXO;
    distrito.readOnly = true;
  }

  interessadoModalTitle.textContent = "Editar interessado";
  updateProgressPreview();
}

/* =========================
   DASHBOARD / TABELAS
========================= */
function renderMetrics() {
  const data = getVisibleInteressados();

  const metrics = [
  {
    label: "Total de interessados",
    metric: "todos",
    value: data.length,
    className: "metric-card soft-blue"
  },
  {
    label: "Ativos",
    metric: "ativos",
    value: data.filter((item) => item.status === "Ativo").length,
    className: "metric-card soft-green"
  },
  {
    label: "Em risco",
    metric: "risco",
    value: data.filter(isAtRisk).length,
    className: "metric-card soft-yellow"
  },
  {
    label: "Concluídos",
    metric: "concluidos",
    value: data.filter((item) => item.status === "Concluído").length,
    className: "metric-card soft-green"
  },
  {
    label: "Prontos para decisão",
    metric: "decisao",
    value: data.filter(isDecision).length,
    className: "metric-card soft-blue"
  },
  {
    label: "Batismo realizado",
    metric: "batismo",
    value: data.filter((item) => item.status === "Batismo Realizado").length,
    className: "metric-card soft-blue"
  }
];

  metricsGrid.innerHTML = metrics
    .map(
  (item) => `
    <div
      class="${item.className} dashboard-metric-link"
      data-metric="${item.metric}"
      style="cursor:pointer"
    >
      <span class="metric-label">${escapeHtml(item.label)}</span>
      <strong class="metric-value">${item.value}</strong>
    </div>
  `
)
    .join("");

bindMetricLinks();
}

function bindMetricLinks() {

  document
    .querySelectorAll(".dashboard-metric-link")
    .forEach((card) => {

      card.addEventListener("click", () => {

        const metric =
          card.dataset.metric;

        let filtered =
          [...state.interessados];

        switch(metric) {

          case "ativos":
            filtered =
              filtered.filter(
                item => item.status === "Ativo"
              );
            break;

          case "risco":
            filtered =
              filtered.filter(isAtRisk);
            break;

          case "concluidos":
            filtered =
              filtered.filter(
                item => item.status === "Concluído"
              );
            break;

          case "decisao":
            filtered =
              filtered.filter(isDecision);
            break;

          case "batismo":
            filtered =
              filtered.filter(
                item =>
                  item.status ===
                  "Batismo Realizado"
              );
            break;
        }

        setSection("interessadosSection");

        renderInteressadosTable(filtered);
      });
    });
}

function renderRecentList() {
  const rawData = getVisibleInteressados().slice(0, 6);
  const data = rawData.map((item) => resolveInteressadoFullData(item, state));

  if (!data.length) {
    recentList.innerHTML = `<div class="empty-state">Nenhum interessado encontrado.</div>`;
    return;
  }

  recentList.innerHTML = data
  .map((item) => {
    const p = calcProgress(item.estudoAtual, item.totalEstudos);

    return `
      <article class="stack-item">
        <div class="stack-item-top">
          <h4>${escapeHtml(item.nome)}</h4>
          <span class="tiny-muted">${escapeHtml(item.instrutoresString)}</span>
        </div>
        <p class="stack-item-sub">
          ${escapeHtml(item.igrejaNome)} • <strong>${escapeHtml(item.serieNome)}</strong>
        </p>
        <div class="progress-meta">
          ${p.capped} de ${p.total} • ${p.porcentagem}% • faltam ${p.faltantes}
        </div>
        <div class="progress-bar">
          <span style="width:${p.porcentagem}%"></span>
        </div>
      </article>
    `;
  })
  .join("");
}

function renderAttentionList() {
  const rawData = getVisibleInteressados()
    .filter((item) => isAtRisk(item) || isDecision(item))
    .slice(0, 8);
  const data = rawData.map((item) => resolveInteressadoFullData(item, state));

  if (!data.length) {
    attentionList.innerHTML = `<div class="empty-state">Nenhum caso crítico no momento.</div>`;
    return;
  }

  attentionList.innerHTML = data
    .map((item) => {
      const note = isAtRisk(item)
        ? `Último contato há ${daysSince(item.ultimoContato)} dias`
        : "Pronto para decisão espiritual";

      return `
        <article class="stack-item">
          <div class="stack-item-top">
            <h4>${escapeHtml(item.nome)}</h4>
          </div>
          <p class="stack-item-sub">${escapeHtml(note)}</p>
          <div class="pill-row">
            <span class="pill status-${slug(item.status)}">${escapeHtml(item.status)}</span>
            <span class="pill interest-${slug(item.interesse)}">${escapeHtml(item.interesse)}</span>
          </div>
        </article>
      `;
    })
    .join("");
}

function renderInteressadosTable(customData = null) {
  const rawData = Array.isArray(customData) ? customData : getFilteredInteressados();
  const data = rawData.map((item) => resolveInteressadoFullData(item, state));

  if (!data.length) {
    interessadosTable.innerHTML = `
      <tr>
        <td colspan="7">
          <div class="empty-state">Nenhum registro encontrado.</div>
        </td>
      </tr>
    `;
    return;
  }

  interessadosTable.innerHTML = data
    .map((item) => {
      const p = calcProgress(item.estudoAtual, item.totalEstudos);
      const canAdvance = canAddStudyToInteressado(item);

      return `
        <tr>
          <td>
            <strong>${escapeHtml(item.nome)}</strong><br>
            <span class="tiny-muted">${escapeHtml(item.telefone)}</span><br>
            <span class="tiny-muted">${escapeHtml(item.igrejaNome)} • ${escapeHtml(item.distrito || DISTRITO_FIXO)}</span>
          </td>
          <td>
            <div style="font-weight:600;color:var(--text);">${escapeHtml(item.instrutoresString)}</div>
            ${item.criadoPorNome && item.criadoPorNome !== "-" && item.criadoPorNome !== item.instrutoresString ? `<span class="tiny-muted">Cadastrado por: ${escapeHtml(item.criadoPorNome)}</span>` : ""}
          </td>
          <td>
            <strong>${escapeHtml(item.serieNome)}</strong>
          </td>
          <td>
            <div class="progress-meta">${p.capped} de ${p.total} • ${p.porcentagem}%</div>
            <div class="progress-bar compact">
              <span style="width:${p.porcentagem}%"></span>
            </div>
            <span class="tiny-muted" style="font-size:11px;">Faltam ${p.faltantes} lições</span>
          </td>
          <td><span class="pill status-${slug(item.status)}">${escapeHtml(item.status)}</span></td>
          <td><span class="pill interest-${slug(item.interesse)}">${escapeHtml(item.interesse)}</span></td>
          <td>
            <div class="action-row">
              <button class="btn btn-secondary btn-sm" data-edit-interessado="${item.id}">Editar</button>
              <button class="btn btn-secondary btn-sm" data-add-study="${item.id}" ${canAdvance ? "" : "disabled"}>+1 estudo</button>
              <button class="btn btn-danger btn-sm" data-delete-interessado="${item.id}">Excluir</button>
            </div>
          </td>
        </tr>
      `;
    })
    .join("");
}

function renderSeriesList() {
  if (!seriesList) return;

  const search = normalizeText(document.getElementById("catalogoSearchInput")?.value || "");
  let data = sortByName(state.series);

  if (search) {
    data = data.filter((s) => normalizeText(s.nome).includes(search));
  }

  if (!data.length) {
    seriesList.innerHTML = `<div class="empty-state">Nenhuma série cadastrada.</div>`;
    return;
  }

  seriesList.innerHTML = data
    .map((serie) => {
      const controls = canManageSeries()
        ? `
          <div class="action-row">
            <button class="btn btn-secondary btn-sm" data-edit-serie="${serie.id}">Editar</button>
            <button class="btn btn-danger btn-sm" data-delete-serie="${serie.id}">Excluir</button>
          </div>
        `
        : "";

      return `
        <article class="stack-item">
          <div class="stack-item-top">
            <h4>${escapeHtml(serie.nome)}</h4>
          </div>
          <p class="stack-item-sub">Total de estudos: <strong>${serie.totalEstudos}</strong></p>
          ${controls}
        </article>
      `;
    })
    .join("");
}

function getVisibleUsersList() {
  if (isAdmin() || isDistrital()) {
    return sortByName(state.usuarios);
  }

  if (isLocal()) {
    const localId = getCurrentUserLocalId();
    return sortByName(state.usuarios.filter((u) => u.igrejaId === localId));
  }

  return [];
}

function renderUsuariosList() {
  if (!usuariosList) return;

  if (!canSeeUsersSection()) {
    usuariosList.innerHTML = `<div class="empty-state">Sem permissão para visualizar usuários.</div>`;
    return;
  }

  const search = normalizeText(document.getElementById("usuariosSearchInput")?.value || "");
  let visible = getVisibleUsersList();

  if (search) {
    visible = visible.filter(
      (u) =>
        normalizeText(u.nome || "").includes(search) ||
        normalizeText(u.email || u.emailAuth || "").includes(search) ||
        normalizeText(u.username || "").includes(search) ||
        normalizeText(formatRole(u.perfil)).includes(search) ||
        normalizeText(u.igrejaNome || "").includes(search)
    );
  }

  if (!visible.length) {
    usuariosList.innerHTML = `<div class="empty-state">Nenhum usuário cadastrado.</div>`;
    return;
  }

  usuariosList.innerHTML = visible
    .map((usuario) => {
      const loginInfo = ["admin", "distrital"].includes(usuario.perfil)
        ? (usuario.email || usuario.emailAuth || "-")
        : (usuario.username || "-");

      const podeEditar = canEditUser(usuario);
      const podeExcluir = canDeleteUser(usuario);

      const controls = (podeEditar || podeExcluir)
        ? `
          <div class="action-row">
            ${podeEditar ? `<button class="btn btn-secondary btn-sm" data-edit-user="${usuario.id}">Editar</button>` : ""}
            ${podeExcluir ? `<button class="btn btn-danger btn-sm" data-delete-user="${usuario.id}">Excluir</button>` : ""}
          </div>
        `
        : "";

      return `
        <article class="stack-item">
          <div class="stack-item-top">
            <h4>${escapeHtml(usuario.nome || "Sem nome")}</h4>
            <span class="tiny-muted">${escapeHtml(formatRole(usuario.perfil))}</span>
          </div>
          <p class="stack-item-sub">
            ${escapeHtml(loginInfo)} • ${escapeHtml(usuario.igrejaNome || "Sem igreja")} • ${escapeHtml(usuario.distrito || DISTRITO_FIXO)}
          </p>
          <div class="pill-row">
            <span class="pill ${usuario.ativo ? "status-ativo" : "status-desinteressado"}">
              ${usuario.ativo ? "Ativo" : "Inativo"}
            </span>
          </div>
          ${controls}
        </article>
      `;
    })
    .join("");
}

function renderLocaisList() {
  if (!locaisList) return;

  if (!canManageLocais()) {
    locaisList.innerHTML = `<div class="empty-state">Apenas administradores podem gerenciar locais.</div>`;
    return;
  }

  const search = normalizeText(document.getElementById("locaisSearchInput")?.value || "");
  let list = sortByName(state.locais);

  if (search) {
    list = list.filter(
      (l) =>
        normalizeText(l.nome || "").includes(search) ||
        normalizeText(l.tipo || "").includes(search)
    );
  }

  if (!list.length) {
    locaisList.innerHTML = `<div class="empty-state">Nenhum local cadastrado.</div>`;
    return;
  }

  locaisList.innerHTML = list
    .map((local) => {
      return `
        <article class="stack-item">
          <div class="stack-item-top">
            <h4>${escapeHtml(local.nome)}</h4>
            <span class="tiny-muted">${escapeHtml(local.tipo || "igreja")}</span>
          </div>
          <p class="stack-item-sub">${escapeHtml(DISTRITO_FIXO)}</p>
          <div class="action-row">
            <button class="btn btn-secondary btn-sm" data-edit-local="${local.id}">Editar</button>
            <button class="btn btn-danger btn-sm" data-delete-local="${local.id}">Excluir</button>
          </div>
        </article>
      `;
    })
    .join("");
}

/* =========================
   INSTRUTORES TAB
========================= */
function formatDate(dateStr) {
  if (!dateStr) return "";
  try {
    const parts = String(dateStr).split("-");
    if (parts.length === 3) {
      return `${parts[2]}/${parts[1]}/${parts[0]}`;
    }
    const d = new Date(dateStr);
    if (!isNaN(d.getTime())) {
      return d.toLocaleDateString("pt-BR");
    }
  } catch (e) {}
  return dateStr;
}

function getRoleLabel(perfil) {
  switch (perfil) {
    case "admin":
      return "PASTOR / ADMINISTRADOR";
    case "distrital":
      return "LÍDER DISTRITAL";
    case "local":
      return "LÍDER LOCAL";
    case "membro":
      return "MEMBRO";
    default:
      return (perfil || "INSTRUTOR").toUpperCase();
  }
}

function getInstructorInitials(name = "") {
  if (!name) return "IN";
  const clean = name.trim().split(/\s+/).filter(Boolean);
  if (!clean.length) return "IN";
  if (clean.length === 1) {
    return clean[0].substring(0, 2).toUpperCase();
  }
  const first = clean[0];
  const second = clean[1];
  if (first.toLowerCase().startsWith("pr") || first.toLowerCase().startsWith("pastor")) {
    return "PR";
  }
  return (first[0] + (second ? second[0] : first[1] || "")).toUpperCase();
}

function getStudentsForInstructor(instructor) {
  const instId = String(instructor.id || "");
  const instNomeNorm = normalizeText(instructor.nome || "");

  return (state.interessados || []).filter((item) => {
    // 1. Array de IDs
    if (Array.isArray(item.instrutorIds) && item.instrutorIds.some((id) => String(id) === instId)) {
      return true;
    }
    // 2. Array de objetos { id, nome }
    if (
      Array.isArray(item.instrutores) &&
      item.instrutores.some(
        (i) => String(i.id) === instId || (i.nome && normalizeText(i.nome) === instNomeNorm)
      )
    ) {
      return true;
    }
    // 3. Array de nomes
    if (
      Array.isArray(item.instrutorNomes) &&
      item.instrutorNomes.some((n) => normalizeText(n) === instNomeNorm)
    ) {
      return true;
    }
    // 4. Criado por
    if (item.criadoPorId && String(item.criadoPorId) === instId) {
      return true;
    }
    // 5. Responsavel
    if (item.responsavelNome && normalizeText(item.responsavelNome) === instNomeNorm) {
      return true;
    }
    return false;
  });
}

function getAllInstructorsData() {
  const map = new Map();

  (state.usuarios || []).forEach((u) => {
    if (u.ativo !== false) {
      map.set(u.id, {
        id: u.id,
        nome: u.nome,
        perfil: u.perfil || "membro",
        igrejaId: u.igrejaId || "",
        igrejaNome: u.igrejaNome || "",
        email: u.email || ""
      });
    }
  });

  if (state.user && state.user.uid && !map.has(state.user.uid)) {
    map.set(state.user.uid, {
      id: state.user.uid,
      nome: state.user.nome,
      perfil: state.user.perfil || "membro",
      igrejaId: state.user.igrejaId || "",
      igrejaNome: state.user.igrejaNome || "",
      email: state.user.email || ""
    });
  }

  (state.interessados || []).forEach((item) => {
    if (Array.isArray(item.instrutores)) {
      item.instrutores.forEach((inst) => {
        if (inst && inst.id && !map.has(inst.id)) {
          map.set(inst.id, {
            id: inst.id,
            nome: inst.nome || "Instrutor",
            perfil: "membro",
            igrejaId: item.igrejaId || "",
            igrejaNome: item.igrejaNome || "",
            email: ""
          });
        }
      });
    }
  });

  const roleRank = { admin: 1, distrital: 2, local: 3, membro: 4 };

  const list = Array.from(map.values()).map((inst) => {
    const students = getStudentsForInstructor(inst);
    const activeStudents = students.filter((s) => s.status === "Ativo");
    return {
      ...inst,
      students,
      activeStudents,
      totalAlunos: students.length,
      estudosAtivos: activeStudents.length
    };
  });

  list.sort((a, b) => {
    const rankA = roleRank[a.perfil] || 99;
    const rankB = roleRank[b.perfil] || 99;
    if (rankA !== rankB) return rankA - rankB;
    return a.nome.localeCompare(b.nome, "pt-BR");
  });

  return list;
}

let instrutorExpandedSet = new Set();
let instrutorSearchQuery = "";
let instrutorIgrejaFilterValue = "Todas";

function renderInstrutoresSection() {
  const container = document.getElementById("instrutoresSection");
  if (!container) return;

  const allInstructors = getAllInstructorsData();

  const totalInstructors = allInstructors.length;
  const activeInstructors = allInstructors.filter((i) => i.totalAlunos > 0).length;
  const totalActiveStudies = state.interessados.filter((i) => i.status === "Ativo").length;
  const totalStudies = state.interessados.length;
  const mediaEstudos = totalInstructors > 0 ? (totalActiveStudies / totalInstructors).toFixed(1) : "0.0";
  const capacidade = totalInstructors > 0 ? Math.round((activeInstructors / totalInstructors) * 100) + "%" : "0%";

  const metricsGrid = document.getElementById("instrutoresMetricsGrid");
  if (metricsGrid) {
    metricsGrid.innerHTML = `
      <div class="inst-metric-card">
        <div class="inst-metric-icon green">
          <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path>
            <circle cx="9" cy="7" r="4"></circle>
            <path d="M23 21v-2a4 4 0 0 0-3-3.87"></path>
            <path d="M16 3.13a4 4 0 0 1 0 7.75"></path>
          </svg>
        </div>
        <div class="inst-metric-info">
          <span class="inst-metric-label">INSTRUTORES</span>
          <span class="inst-metric-value">${totalInstructors}</span>
          <span class="inst-metric-subtext">${activeInstructors} ativos com alunos</span>
        </div>
      </div>

      <div class="inst-metric-card">
        <div class="inst-metric-icon blue">
          <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"></path>
            <path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"></path>
          </svg>
        </div>
        <div class="inst-metric-info">
          <span class="inst-metric-label">ESTUDOS ATIVOS</span>
          <span class="inst-metric-value">${totalActiveStudies}</span>
          <span class="inst-metric-subtext">De ${totalStudies} estudos cadastrados</span>
        </div>
      </div>

      <div class="inst-metric-card">
        <div class="inst-metric-icon purple">
          <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <circle cx="12" cy="12" r="3"></circle>
            <path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"></path>
          </svg>
        </div>
        <div class="inst-metric-info">
          <span class="inst-metric-label">MÉDIA DE ESTUDOS</span>
          <span class="inst-metric-value">${mediaEstudos}</span>
          <span class="inst-metric-subtext">Estudos por instrutor cadastrado</span>
        </div>
      </div>

      <div class="inst-metric-card">
        <div class="inst-metric-icon orange">
          <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <circle cx="12" cy="12" r="10"></circle>
            <polyline points="12 6 12 12 16 14"></polyline>
          </svg>
        </div>
        <div class="inst-metric-info">
          <span class="inst-metric-label">CAPACIDADE</span>
          <span class="inst-metric-value">${capacidade}</span>
          <span class="inst-metric-subtext">De engajamento de obreiros</span>
        </div>
      </div>
    `;
  }

  const churchSelect = document.getElementById("instrutoresIgrejaFilter");
  if (churchSelect) {
    const churches = Array.from(
      new Set(
        [
          ...state.locais.map((l) => l.nome),
          ...allInstructors.map((i) => i.igrejaNome).filter(Boolean),
          ...state.interessados.map((i) => i.igrejaNome).filter(Boolean)
        ].filter(Boolean)
      )
    ).sort((a, b) => a.localeCompare(b, "pt-BR"));

    const currentVal = instrutorIgrejaFilterValue;
    churchSelect.innerHTML =
      `<option value="Todas">Todas as Igrejas</option>` +
      churches
        .map(
          (c) =>
            `<option value="${escapeHtml(c)}" ${c === currentVal ? "selected" : ""}>${escapeHtml(c)}</option>`
        )
        .join("");
  }

  const query = normalizeText(instrutorSearchQuery);
  const selectedChurch = instrutorIgrejaFilterValue;

  const filteredInstructors = allInstructors.filter((inst) => {
    if (selectedChurch && selectedChurch !== "Todas") {
      const matchInstChurch = inst.igrejaNome === selectedChurch;
      const matchStudentsChurch = inst.students.some((s) => s.igrejaNome === selectedChurch);
      if (!matchInstChurch && !matchStudentsChurch) return false;
    }

    if (query) {
      const matchName = normalizeText(inst.nome).includes(query);
      const matchEmail = normalizeText(inst.email).includes(query);
      const matchChurch = normalizeText(inst.igrejaNome).includes(query);
      const matchStudents = inst.students.some((s) => normalizeText(s.nome).includes(query));
      if (!matchName && !matchEmail && !matchChurch && !matchStudents) return false;
    }

    return true;
  });

  const listContainer = document.getElementById("instrutoresListContainer");
  if (listContainer) {
    if (!filteredInstructors.length) {
      listContainer.innerHTML = `<div class="empty-state">Nenhum instrutor encontrado com os filtros selecionados.</div>`;
      return;
    }

    listContainer.innerHTML = filteredInstructors
      .map((inst) => {
        const isExpanded = instrutorExpandedSet.has(inst.id);
        const initials = getInstructorInitials(inst.nome);
        const isSelf = state.user && state.user.uid === inst.id;
        const roleLabel = getRoleLabel(inst.perfil);
        const firstName = inst.nome.trim().split(/\s+/)[0] || "INSTRUTOR";

        const studentsHtml =
          inst.students.length === 0
            ? `<div class="inst-empty-custody">Nenhum estudo bíblico sob custódia deste instrutor no momento.</div>`
            : `
            <div class="students-custody-grid">
              ${inst.students
                .map((student) => {
                  const p = calcProgress(student.estudoAtual, student.totalEstudos);
                  const statusClass =
                    student.status === "Pausado"
                      ? "paused"
                      : student.status === "Concluído"
                      ? "done"
                      : "";

                  return `
                    <div class="inst-student-card">
                      <div class="inst-student-top">
                        <div>
                          <h4 class="inst-student-name">${escapeHtml(student.nome)}</h4>
                          <p class="inst-student-church">Igreja: ${escapeHtml(student.igrejaNome || "Não informada")}</p>
                        </div>
                        <span class="inst-student-badge ${statusClass}">${escapeHtml(student.status || "ATIVO")}</span>
                      </div>

                      <div class="inst-course-box">
                        <div class="inst-course-head">
                          <div>
                            <div class="inst-course-title">${escapeHtml(student.serieNome || "Estudo Bíblico")}</div>
                            <div class="inst-course-lesson">Lição ${p.capped} de ${p.total}</div>
                          </div>
                          <div>
                            <div class="inst-course-pct">${p.porcentagem}%</div>
                            <div class="inst-course-left">Restam ${p.faltantes} ${p.faltantes === 1 ? "lição" : "lições"}</div>
                          </div>
                        </div>
                        <div class="inst-progress-bar-track">
                          <div class="inst-progress-bar-fill" style="width: ${p.porcentagem}%;"></div>
                        </div>
                      </div>

                      <div class="inst-student-bottom">
                        <a href="tel:${escapeHtml(student.telefone || "")}" class="inst-phone-pill">
                          <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                            <path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72 12.84 12.84 0 0 0 .7 2.81 2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45 12.84 12.84 0 0 0 2.81.7A2 2 0 0 1 22 16.92z"></path>
                          </svg>
                          ${escapeHtml(student.telefone || "Sem telefone")}
                        </a>

                        <span class="inst-contact-date">
                          <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                            <rect x="3" y="4" width="18" height="18" rx="2" ry="2"></rect>
                            <line x1="16" y1="2" x2="16" y2="6"></line>
                            <line x1="8" y1="2" x2="8" y2="6"></line>
                            <line x1="3" y1="10" x2="21" y2="10"></line>
                          </svg>
                          Último contato: ${escapeHtml(formatDate(student.ultimoContato) || "Não informado")}
                        </span>
                      </div>
                    </div>
                  `;
                })
                .join("")}
            </div>
          `;

        return `
          <div class="instructor-accordion-card ${isExpanded ? "expanded" : ""}" data-instructor-id="${inst.id}">
            <div class="instructor-card-head" data-toggle-instructor="${inst.id}">
              <div class="instructor-head-left">
                <div class="instructor-avatar-badge">${escapeHtml(initials)}</div>
                <div class="instructor-meta-block">
                  <div class="instructor-meta-name-row">
                    <span class="instructor-meta-name">${escapeHtml(inst.nome)}</span>
                    ${isSelf ? `<span class="badge-you">VOCÊ</span>` : ""}
                  </div>
                  <span class="badge-role-tag">${escapeHtml(roleLabel)}</span>
                </div>
              </div>

              <div class="instructor-head-right">
                <div class="instructor-stat-col">
                  <div class="instructor-stat-label">TOTAL ALUNOS</div>
                  <div class="instructor-stat-val">${inst.totalAlunos} interessado${inst.totalAlunos === 1 ? "" : "s"}</div>
                </div>

                <div class="instructor-stat-col">
                  <div class="instructor-stat-label green">ESTUDOS ATIVOS</div>
                  <div class="instructor-stat-val green">${inst.estudosAtivos} em andamento</div>
                </div>

                <div class="instructor-toggle-btn">
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
                    <polyline points="6 9 12 15 18 9"></polyline>
                  </svg>
                </div>
              </div>
            </div>

            ${
              isExpanded
                ? `
                  <div class="instructor-card-drawer">
                    <div class="custody-header">
                      📖 ESTUDOS BÍBLICOS SOB CUSTÓDIA DE ${escapeHtml(firstName.toUpperCase())}
                    </div>
                    ${studentsHtml}
                  </div>
                `
                : ""
            }
          </div>
        `;
      })
      .join("");
  }
}

/* =========================
   FORM USUÁRIOS
========================= */
function renderUserProfileOptions() {
  if (!novoUsuarioPerfil) return;

  const allowed = getAllowedProfilesToCreate();
  const labels = {
    admin: "Administrador",
    distrital: "Líder Distrital",
    local: "Líder Local",
    membro: "Membro"
  };

  const current = novoUsuarioPerfil.value;
  novoUsuarioPerfil.innerHTML = allowed
    .map((profile) => `<option value="${profile}">${labels[profile]}</option>`)
    .join("");

  if (allowed.includes(current)) {
    novoUsuarioPerfil.value = current;
  } else if (allowed.length) {
    novoUsuarioPerfil.value = allowed[0];
  }

  updateUserFormByProfile();
}

function renderUserLocalOptions() {
  if (!novoUsuarioIgreja) return;

  let locais = [...state.locais];

  if (isLocal()) {
    const localId = getCurrentUserLocalId();
    locais = locais.filter((item) => item.id === localId);
  }

  locais = sortByName(locais);

  novoUsuarioIgreja.innerHTML =
    `<option value="">Selecionar igreja</option>` +
    locais.map((item) => `<option value="${item.id}">${escapeHtml(item.nome)}</option>`).join("");

  if (isLocal() && locais[0]) {
    novoUsuarioIgreja.value = locais[0].id;
    novoUsuarioIgreja.disabled = true;
  } else {
    novoUsuarioIgreja.disabled = false;
  }
}

function updateUserFormByProfile() {
  const profile = novoUsuarioPerfil?.value || "membro";
  const globalAccount = profile === "admin" || profile === "distrital";
  const localAccount = profile === "local" || profile === "membro";

  if (novoUsuarioUsernameGroup) {
    novoUsuarioUsernameGroup.style.display = localAccount ? "" : "none";
  }

  if (novoUsuarioSenhaLocalGroup) {
    novoUsuarioSenhaLocalGroup.style.display = localAccount ? "" : "none";
  }

  if (novoUsuarioEmailGroup) {
    novoUsuarioEmailGroup.style.display = globalAccount ? "" : "none";
  }

  if (novoUsuarioSenhaGroup) {
    novoUsuarioSenhaGroup.style.display = globalAccount ? "" : "none";
  }

  if (novoUsuarioUsername) {
    novoUsuarioUsername.required = localAccount;
    if (!localAccount) novoUsuarioUsername.value = "";
  }

  if (novoUsuarioSenhaLocal) {
    novoUsuarioSenhaLocal.required = localAccount;
    if (!localAccount) novoUsuarioSenhaLocal.value = "";
  }

  if (novoUsuarioEmail) {
    novoUsuarioEmail.required = globalAccount;
    if (!globalAccount) novoUsuarioEmail.value = "";
  }

  if (novoUsuarioSenha) {
    novoUsuarioSenha.required = globalAccount;
    if (!globalAccount) novoUsuarioSenha.value = "";
  }

  if (novoUsuarioIgreja) {
    novoUsuarioIgreja.required = localAccount;
  }
}

async function editUser(userId) {
  const targetUser = getUserById(userId);

  if (!targetUser) {
    showMessage("Usuário não encontrado.", "error");
    return;
  }

  if (!canEditUser(targetUser)) {
    showMessage("Você não tem permissão para editar este usuário.", "error");
    return;
  }

  novoUsuarioNome.value = targetUser.nome || "";
  novoUsuarioPerfil.value = targetUser.perfil || "membro";
  novoUsuarioDistrito.value = targetUser.distrito || DISTRITO_FIXO;
  novoUsuarioIgreja.value = targetUser.igrejaId || "";
  novoUsuarioUsername.value = targetUser.username || "";
  novoUsuarioEmail.value = targetUser.email || "";

  updateUserFormByProfile();

  userForm.dataset.editingUserId = targetUser.id;

  showMessage(`Editando usuário: ${targetUser.nome}`, "info");
}

async function deleteUser(userId) {
  const targetUser = getUserById(userId);

  if (!targetUser) {
    showMessage("Usuário não encontrado.", "error");
    return;
  }

  if (!canDeleteUser(targetUser)) {
    showMessage("Você não tem permissão para excluir este usuário.", "error");
    return;
  }

  const confirmed = window.confirm(`Tem certeza que deseja excluir o usuário "${targetUser.nome}"?`);
  if (!confirmed) return;

  showLoading();
  clearMessage();

  try {
    await deleteDoc(doc(db, "usuarios", userId));

    if (userForm?.dataset?.editingUserId === userId) {
      delete userForm.dataset.editingUserId;
      userForm.reset();
      renderUserProfileOptions();
      renderUserLocalOptions();
      updateUserFormByProfile();
    }

    showMessage("Usuário excluído com sucesso.", "success");
    await refreshData();
  } catch (error) {
    console.error("Erro ao excluir usuário:", error);
    showMessage("Não foi possível excluir o usuário.", "error");
  } finally {
    hideLoading();
  }
}

/* =========================
   CRUD INTERESSADOS
========================= */
function getSelectedResponsavel() {
  const responsavelId = instrutor.value;
  if (!responsavelId) {
    throw new Error("Selecione o responsável.");
  }

  const user = getUserById(responsavelId);

  if (!user) {
    throw new Error("Responsável inválido.");
  }

  return user;
}

function getSelectedInteressadoLocal() {
  const localId = igreja.value;
  if (!localId) {
    throw new Error("Selecione a igreja.");
  }

  const local = getLocalById(localId);

  if (!local) {
    throw new Error("Igreja inválida.");
  }

  return local;
}

async function handleInteressadoSubmit(event) {
  event.preventDefault();
  clearMessage();

  try {
    const recordId = interessadoId.value.trim();
    const selectedSerie = getSerieById(serieId.value);

    if (!selectedSerie) {
      throw new Error("Selecione uma série válida.");
    }

    if (!nome.value.trim()) {
      throw new Error("Informe o nome do interessado.");
    }

    const selectedLocal = getSelectedInteressadoLocal();
    const instrutoresSelecionados = getInstrutoresFromForm();

    if (!instrutoresSelecionados.length) {
      throw new Error("Selecione pelo menos um instrutor.");
    }

    const progress = calcProgress(estudoAtual.value, selectedSerie.totalEstudos);

    if ((isLocal() || isMembro()) && selectedLocal.id !== getCurrentUserLocalId()) {
      throw new Error("Você só pode cadastrar interessados na sua igreja.");
    }

    if (isMembro()) {
      const membroFoiSelecionado = instrutoresSelecionados.some((item) => item.id === state.user.uid);
      if (!membroFoiSelecionado) {
        throw new Error("Membro deve estar incluído entre os instrutores.");
      }
    }

    const payload = {
      nome: nome.value.trim(),
      telefone: telefone.value.trim(),
      endereco: endereco.value.trim(),
      igrejaId: selectedLocal.id,
      igrejaNome: selectedLocal.nome,
      igrejaTipo: selectedLocal.tipo || "igreja",
      distrito: DISTRITO_FIXO,

      instrutorIds: instrutoresSelecionados.map((item) => item.id),
      instrutorNomes: instrutoresSelecionados.map((item) => item.nome),
      instrutores: instrutoresSelecionados,

      criadoPorId: state.user.uid,
      criadoPorNome: state.user.nome,
      criadoPorPerfil: state.user.perfil,

      serieId: selectedSerie.id,
      serieNome: selectedSerie.nome,
      estudoAtual: progress.capped,
      totalEstudos: selectedSerie.totalEstudos,
      porcentagem: progress.porcentagem,
      faltantes: progress.faltantes,
      status: status.value,
      interesse: interesse.value,
      observacoes: observacoes.value.trim(),
      ultimoContato: ultimoContato.value || "",
      atualizadoEm: new Date().toISOString(),
      updatedAt: serverTimestamp()
    };

    showLoading();

    if (recordId) {
      const existing = state.interessados.find((item) => item.id === recordId);

      if (!existing) {
        throw new Error("Registro não encontrado.");
      }

      if (!userCanManageInteressado(existing)) {
        throw new Error("Você não pode editar esse registro.");
      }

      await updateDoc(doc(db, "interessados", recordId), payload);
      showMessage("Interessado atualizado com sucesso.", "success");
    } else {
      payload.criadoEm = new Date().toISOString();
      payload.createdAt = serverTimestamp();

      await addDoc(collection(db, "interessados"), payload);
      showMessage("Interessado cadastrado com sucesso.", "success");
    }

    await refreshData();
    closeInteressadoModal();
    resetInteressadoForm();
  } catch (error) {
    console.error("Erro ao salvar interessado:", error);
    showMessage(error.message || "Não foi possível salvar o interessado.", "error");
  } finally {
    hideLoading();
  }
}

async function advanceInteressadoStudy(id) {
  const item = state.interessados.find((row) => row.id === id);
  if (!item) return;

  if (!canAddStudyToInteressado(item)) {
    showMessage("Você não pode avançar o estudo desse interessado.", "error");
    return;
  }

  try {
    const nextValue = Number(item.estudoAtual || 0) + 1;
    const progress = calcProgress(nextValue, item.totalEstudos);

    showLoading();

    await updateDoc(doc(db, "interessados", id), {
      estudoAtual: progress.capped,
      porcentagem: progress.porcentagem,
      faltantes: progress.faltantes,
      atualizadoEm: new Date().toISOString(),
      updatedAt: serverTimestamp()
    });

    showMessage("Estudo avançado com sucesso.", "success");
    await refreshData();
  } catch (error) {
    console.error("Erro ao avançar estudo:", error);
    showMessage("Não foi possível avançar o estudo.", "error");
  } finally {
    hideLoading();
  }
}

async function editInteressado(id) {
  const item = state.interessados.find((row) => row.id === id);
  if (!item) return;

  if (!userCanManageInteressado(item)) {
    showMessage("Você não pode editar esse registro.", "error");
    return;
  }

  fillInteressadoForm(item);
  openInteressadoModal();
}

async function deleteInteressado(id) {
  const item = state.interessados.find((row) => row.id === id);
  if (!item) return;

  if (!userCanManageInteressado(item)) {
    showMessage("Você não pode excluir esse registro.", "error");
    return;
  }

  const confirmed = window.confirm(`Excluir o registro de "${item.nome}"?`);
  if (!confirmed) return;

  try {
    showLoading();
    await deleteDoc(doc(db, "interessados", id));
    showMessage("Interessado excluído com sucesso.", "success");
    await refreshData();
  } catch (error) {
    console.error("Erro ao excluir interessado:", error);
    showMessage("Não foi possível excluir o interessado.", "error");
  } finally {
    hideLoading();
  }
}

/* =========================
   CRUD SÉRIES
========================= */
async function handleSerieSubmit(event) {
  event.preventDefault();
  clearMessage();

  if (!canManageSeries()) {
    showMessage("Apenas administradores podem gerenciar séries.", "error");
    return;
  }

  try {
    const nomeSerie = novaSerieNome.value.trim();
    const totalEstudos = Number(novaSerieTotal.value);

    if (!nomeSerie) {
      throw new Error("Informe o nome da série.");
    }

    if (!Number.isFinite(totalEstudos) || totalEstudos < 1) {
      throw new Error("Informe um total de estudos válido.");
    }

    const duplicate = state.series.find(
      (serie) => normalizeText(serie.nome) === normalizeText(nomeSerie)
    );

    if (duplicate) {
      throw new Error("Já existe uma série com esse nome.");
    }

    showLoading();

    await addDoc(collection(db, "series"), {
      nome: nomeSerie,
      totalEstudos,
      ativo: true,
      criadoPor: state.user.uid,
      criadoEm: new Date().toISOString(),
      atualizadoEm: new Date().toISOString(),
      createdAt: serverTimestamp(),
      updatedAt: serverTimestamp()
    });

    serieForm.reset();
    showMessage("Série criada com sucesso.", "success");
    await refreshData();
  } catch (error) {
    console.error("Erro ao criar série:", error);
    showMessage(error.message || "Não foi possível criar a série.", "error");
  } finally {
    hideLoading();
  }
}

async function editSerie(id) {
  if (!canManageSeries()) {
    showMessage("Apenas administradores podem editar séries.", "error");
    return;
  }

  const serie = getSerieById(id);
  if (!serie) return;

  const novoNome = window.prompt("Editar nome da série:", serie.nome);
  if (novoNome === null) return;

  const nomeFinal = novoNome.trim();
  if (!nomeFinal) {
    showMessage("O nome da série não pode ficar vazio.", "error");
    return;
  }

  const novoTotal = window.prompt("Editar total de estudos:", String(serie.totalEstudos));
  if (novoTotal === null) return;

  const totalFinal = Number(novoTotal);
  if (!Number.isFinite(totalFinal) || totalFinal < 1) {
    showMessage("O total de estudos precisa ser válido.", "error");
    return;
  }

  const duplicate = state.series.find(
    (item) => item.id !== id && normalizeText(item.nome) === normalizeText(nomeFinal)
  );

  if (duplicate) {
    showMessage("Já existe outra série com esse nome.", "error");
    return;
  }

  try {
    showLoading();

    await updateDoc(doc(db, "series", id), {
      nome: nomeFinal,
      totalEstudos: totalFinal,
      atualizadoEm: new Date().toISOString(),
      updatedAt: serverTimestamp()
    });

    const impacted = state.interessados.filter((item) => item.serieId === id);

    await Promise.all(
      impacted.map(async (item) => {
        const progress = calcProgress(item.estudoAtual, totalFinal);
        await updateDoc(doc(db, "interessados", item.id), {
          serieNome: nomeFinal,
          totalEstudos: totalFinal,
          estudoAtual: progress.capped,
          porcentagem: progress.porcentagem,
          faltantes: progress.faltantes,
          atualizadoEm: new Date().toISOString(),
          updatedAt: serverTimestamp()
        });
      })
    );

    showMessage("Série atualizada com sucesso.", "success");
    await refreshData();
  } catch (error) {
    console.error("Erro ao editar série:", error);
    showMessage("Não foi possível editar a série.", "error");
  } finally {
    hideLoading();
  }
}

async function deleteSerie(id) {
  if (!canManageSeries()) {
    showMessage("Apenas administradores podem excluir séries.", "error");
    return;
  }

  const serie = getSerieById(id);
  if (!serie) return;

  const inUse = state.interessados.some((item) => item.serieId === id);
  if (inUse) {
    showMessage("Essa série está vinculada a interessados e não pode ser excluída agora.", "error");
    return;
  }

  const confirmed = window.confirm(`Excluir a série "${serie.nome}"?`);
  if (!confirmed) return;

  try {
    showLoading();
    await deleteDoc(doc(db, "series", id));
    showMessage("Série excluída com sucesso.", "success");
    await refreshData();
  } catch (error) {
    console.error("Erro ao excluir série:", error);
    showMessage("Não foi possível excluir a série.", "error");
  } finally {
    hideLoading();
  }
}

/* =========================
   CRUD LOCAIS
========================= */
async function handleLocalFormSubmit(event) {
  event.preventDefault();
  clearMessage();

  if (!canManageLocais()) {
    showMessage("Apenas administradores podem gerenciar locais.", "error");
    return;
  }

  try {
    const nomeValue = novoLocalNome.value.trim();
    const tipoValue = novoLocalTipo.value;

    if (!nomeValue) {
      throw new Error("Informe o nome do local.");
    }

    const duplicate = state.locais.find(
      (item) => normalizeText(item.nome) === normalizeText(nomeValue)
    );

    if (duplicate) {
      throw new Error("Já existe um local com esse nome.");
    }

    showLoading();

    await addDoc(collection(db, "locais"), {
      nome: nomeValue,
      tipo: tipoValue || "igreja",
      distrito: DISTRITO_FIXO,
      ativo: true,
      criadoPor: state.user.uid,
      criadoEm: new Date().toISOString(),
      atualizadoEm: new Date().toISOString(),
      createdAt: serverTimestamp(),
      updatedAt: serverTimestamp()
    });

    localForm.reset();
    showMessage("Local criado com sucesso.", "success");
    await refreshData();
  } catch (error) {
    console.error("Erro ao criar local:", error);
    showMessage(error.message || "Não foi possível criar o local.", "error");
  } finally {
    hideLoading();
  }
}

async function editLocal(id) {
  if (!canManageLocais()) {
    showMessage("Apenas administradores podem editar locais.", "error");
    return;
  }

  const local = getLocalById(id);
  if (!local) return;

  const novoNome = window.prompt("Editar nome do local:", local.nome);
  if (novoNome === null) return;

  const nomeFinal = novoNome.trim();
  if (!nomeFinal) {
    showMessage("O nome do local não pode ficar vazio.", "error");
    return;
  }

  const novoTipo = window.prompt("Editar tipo do local (igreja/grupo):", local.tipo || "igreja");
  if (novoTipo === null) return;

  const tipoFinal = novoTipo.trim() || "igreja";

  const duplicate = state.locais.find(
    (item) => item.id !== id && normalizeText(item.nome) === normalizeText(nomeFinal)
  );

  if (duplicate) {
    showMessage("Já existe outro local com esse nome.", "error");
    return;
  }

  try {
    showLoading();

    await updateDoc(doc(db, "locais", id), {
      nome: nomeFinal,
      tipo: tipoFinal,
      distrito: DISTRITO_FIXO,
      atualizadoEm: new Date().toISOString(),
      updatedAt: serverTimestamp()
    });

    await Promise.all([
      ...state.usuarios
        .filter((u) => u.igrejaId === id)
        .map((u) =>
          updateDoc(doc(db, "usuarios", u.id), {
            igrejaNome: nomeFinal,
            distrito: DISTRITO_FIXO,
            atualizadoEm: new Date().toISOString(),
            updatedAt: serverTimestamp()
          })
        ),
      ...state.interessados
        .filter((i) => i.igrejaId === id)
        .map((i) =>
          updateDoc(doc(db, "interessados", i.id), {
            igrejaNome: nomeFinal,
            igrejaTipo: tipoFinal,
            distrito: DISTRITO_FIXO,
            atualizadoEm: new Date().toISOString(),
            updatedAt: serverTimestamp()
          })
        ),
      ...state.usuarios
        .filter((u) => u.igrejaId === id && (u.perfil === "local" || u.perfil === "membro"))
        .map((u) =>
          updateDoc(doc(db, "login_index", u.id), {
            igrejaNome: nomeFinal,
            atualizadoEm: new Date().toISOString(),
            updatedAt: serverTimestamp()
          }).catch(() => {})
        )
    ]);

    showMessage("Local atualizado com sucesso.", "success");
    await refreshData();
  } catch (error) {
    console.error("Erro ao editar local:", error);
    showMessage("Não foi possível editar o local.", "error");
  } finally {
    hideLoading();
  }
}

async function deleteLocal(id) {
  if (!canManageLocais()) {
    showMessage("Apenas administradores podem excluir locais.", "error");
    return;
  }

  const local = getLocalById(id);
  if (!local) return;

  const inUse =
    state.usuarios.some((u) => u.igrejaId === id) ||
    state.interessados.some((i) => i.igrejaId === id);

  if (inUse) {
    showMessage("Esse local está vinculado a usuários ou interessados.", "error");
    return;
  }

  const confirmed = window.confirm(`Excluir o local "${local.nome}"?`);
  if (!confirmed) return;

  try {
    showLoading();
    await deleteDoc(doc(db, "locais", id));
    showMessage("Local excluído com sucesso.", "success");
    await refreshData();
  } catch (error) {
    console.error("Erro ao excluir local:", error);
    showMessage("Não foi possível excluir o local.", "error");
  } finally {
    hideLoading();
  }
}

/* =========================
   CRUD USUÁRIOS
========================= */
function validateUserCreationByRole(profile) {
  const allowed = getAllowedProfilesToCreate();

  if (!allowed.includes(profile)) {
    throw new Error("Você não pode criar esse tipo de usuário.");
  }
}

async function ensureUniqueUsernameInChurch(username, igrejaId) {
  const q = query(
    collection(db, "login_index"),
    where("username", "==", username),
    where("igrejaId", "==", igrejaId),
    limit(1)
  );

  const snap = await getDocs(q);
  return snap.empty;
}

async function createGlobalUserWithSecondaryApp({ nome, email, senha, perfil }) {
  const secondaryName = `secondary-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
  const secondaryApp = initializeApp(app.options, secondaryName);
  const secondaryAuth = getAuth(secondaryApp);

  try {
    const credential = await createUserWithEmailAndPassword(
      secondaryAuth,
      email,
      senha
    );

    const uid = credential.user.uid;

    await setDoc(doc(db, "usuarios", uid), {
      nome,
      email,
      emailAuth: email,
      perfil,
      username: "",
      senhaLocal: "",
      igrejaId: "",
      igrejaNome: "",
      distrito: DISTRITO_FIXO,
      ativo: true,
      criadoPor: state.user.uid,
      criadoEm: new Date().toISOString(),
      atualizadoEm: new Date().toISOString(),
      createdAt: serverTimestamp(),
      updatedAt: serverTimestamp()
    });

    await signOut(secondaryAuth);
    await deleteApp(secondaryApp);

    return uid;
  } catch (error) {
    try {
      await signOut(secondaryAuth);
    } catch (_) {}
    try {
      await deleteApp(secondaryApp);
    } catch (_) {}
    throw error;
  }
}

async function createLocalOrMembroUserWithSecondaryApp({
  nome,
  perfil,
  igrejaId,
  igrejaNome,
  username,
  senha
}) {
  const secondaryName = `secondary-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
  const secondaryApp = initializeApp(app.options, secondaryName);
  const secondaryAuth = getAuth(secondaryApp);

  const emailAuth = buildEmailAuth(username, igrejaId);

  try {
    const credential = await createUserWithEmailAndPassword(
      secondaryAuth,
      emailAuth,
      senha
    );

    const uid = credential.user.uid;

    await setDoc(doc(db, "usuarios", uid), {
      nome,
      email: "",
      emailAuth,
      perfil,
      username,
      senhaLocal: "",
      igrejaId,
      igrejaNome,
      distrito: DISTRITO_FIXO,
      ativo: true,
      criadoPor: state.user.uid,
      criadoEm: new Date().toISOString(),
      atualizadoEm: new Date().toISOString(),
      createdAt: serverTimestamp(),
      updatedAt: serverTimestamp()
    });

    await setDoc(doc(db, "login_index", uid), {
      username,
      igrejaId,
      igrejaNome,
      perfil,
      emailAuth,
      ativo: true,
      criadoEm: new Date().toISOString(),
      atualizadoEm: new Date().toISOString(),
      createdAt: serverTimestamp(),
      updatedAt: serverTimestamp()
    });

    await signOut(secondaryAuth);
    await deleteApp(secondaryApp);

    return uid;
  } catch (error) {
    try {
      await signOut(secondaryAuth);
    } catch (_) {}
    try {
      await deleteApp(secondaryApp);
    } catch (_) {}
    throw error;
  }
}

async function handleUserFormSubmit(event) {
  event.preventDefault();
  clearMessage();

  const editingUserId = userForm?.dataset?.editingUserId || "";
  const editingUser = editingUserId ? getUserById(editingUserId) : null;

  if (!editingUserId && !canCreateUsers()) {
    showMessage("Você não pode criar usuários.", "error");
    return;
  }

  const nomeValue = novoUsuarioNome.value.trim();
  const perfilValue = novoUsuarioPerfil.value;

  if (!nomeValue) {
    showMessage("Informe o nome do usuário.", "error");
    return;
  }

  try {
    showLoading();

    if (editingUserId) {
      if (!editingUser) {
        throw new Error("Usuário não encontrado.");
      }

      if (!canEditUser(editingUser)) {
        throw new Error("Você não tem permissão para editar este usuário.");
      }

      const payload = {
        nome: nomeValue,
        atualizadoEm: new Date().toISOString(),
        updatedAt: serverTimestamp()
      };

      if (isAdmin()) {
        payload.perfil = perfilValue;
      }

      if (["local", "membro"].includes(editingUser.perfil) || ["local", "membro"].includes(payload.perfil || editingUser.perfil)) {
        const igrejaIdValue = novoUsuarioIgreja.value;

        if (!igrejaIdValue) {
          throw new Error("Selecione a igreja.");
        }

        const local = getLocalById(igrejaIdValue);
        if (!local) {
          throw new Error("Igreja inválida.");
        }

        if (isLocal() && igrejaIdValue !== getCurrentUserLocalId()) {
          throw new Error("Você só pode editar usuários da sua igreja.");
        }

        payload.igrejaId = local.id;
        payload.igrejaNome = local.nome;
        payload.distrito = DISTRITO_FIXO;
      }

      await updateDoc(doc(db, "usuarios", editingUserId), payload);

      delete userForm.dataset.editingUserId;
      userForm.reset();
      renderUserProfileOptions();
      renderUserLocalOptions();
      updateUserFormByProfile();

      showMessage("Usuário atualizado com sucesso.", "success");
      await refreshData();
      return;
    }

    validateUserCreationByRole(perfilValue);

    if (perfilValue === "admin" || perfilValue === "distrital") {
      const emailValue = novoUsuarioEmail.value.trim();
      const senhaValue = novoUsuarioSenha.value;

      if (!emailValue || !senhaValue) {
        throw new Error("Preencha e-mail e senha.");
      }

      await createGlobalUserWithSecondaryApp({
        nome: nomeValue,
        email: emailValue,
        senha: senhaValue,
        perfil: perfilValue
      });

      userForm.reset();
      renderUserProfileOptions();
      renderUserLocalOptions();
      updateUserFormByProfile();

      showMessage("Usuário global criado com sucesso.", "success");
      await refreshData();
      return;
    }

    const igrejaIdValue = novoUsuarioIgreja.value;
    const usernameValue = normalizeText(novoUsuarioUsername.value);
    const senhaLocalValue = novoUsuarioSenhaLocal.value;

    if (!igrejaIdValue) {
      throw new Error("Selecione a igreja.");
    }

    if (!usernameValue) {
      throw new Error("Informe o usuário.");
    }

    if (!senhaLocalValue) {
      throw new Error("Informe a senha.");
    }

    const local = getLocalById(igrejaIdValue);
    if (!local) {
      throw new Error("Igreja inválida.");
    }

    if (isLocal() && igrejaIdValue !== getCurrentUserLocalId()) {
      throw new Error("Você só pode criar usuários da sua igreja.");
    }

    await createLocalOrMembroUserWithSecondaryApp({
      nome: nomeValue,
      perfil: perfilValue,
      igrejaId: local.id,
      igrejaNome: local.nome,
      username: usernameValue,
      senha: senhaLocalValue
    });

    userForm.reset();
    renderUserProfileOptions();
    renderUserLocalOptions();
    updateUserFormByProfile();

    showMessage("Usuário criado com sucesso.", "success");
    await refreshData();
  } catch (error) {
    console.error("Erro ao salvar usuário:", error);
    showMessage(mapUserCreationError(error), "error");
  } finally {
    hideLoading();
  }
}

/* =========================
   EVENTOS
========================= */
function bindStaticEvents() {
  document.querySelectorAll(".nav-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      const targetSection = btn.dataset.section;
      setSection(targetSection);
    });
  });

  addInstrutorBtn?.addEventListener("click", () => {
    if (!instrutoresContainer) return;

    instrutoresContainer.appendChild(createInstrutorRow());
    updateInstrutorRowsUI();
  });

  openInteressadoModalBtn?.addEventListener("click", () => {
    resetInteressadoForm();
    openInteressadoModal();
  });

  openInteressadoModalBtn2?.addEventListener("click", () => {
    resetInteressadoForm();
    openInteressadoModal();
  });

  const openInteressadoModalBtn3 = document.getElementById("openInteressadoModalBtn3");
  openInteressadoModalBtn3?.addEventListener("click", () => {
    resetInteressadoForm();
    openInteressadoModal();
  });

  const instrutoresSearchInput = document.getElementById("instrutoresSearchInput");
  instrutoresSearchInput?.addEventListener("input", (e) => {
    instrutorSearchQuery = e.target.value;
    renderInstrutoresSection();
  });

  const instrutoresIgrejaFilter = document.getElementById("instrutoresIgrejaFilter");
  instrutoresIgrejaFilter?.addEventListener("change", (e) => {
    instrutorIgrejaFilterValue = e.target.value;
    renderInstrutoresSection();
  });

  const instrutoresListContainer = document.getElementById("instrutoresListContainer");
  instrutoresListContainer?.addEventListener("click", (e) => {
    const toggleHead = e.target.closest("[data-toggle-instructor]");
    if (toggleHead) {
      const instId = toggleHead.dataset.toggleInstructor;
      if (instrutorExpandedSet.has(instId)) {
        instrutorExpandedSet.delete(instId);
      } else {
        instrutorExpandedSet.add(instId);
      }
      renderInstrutoresSection();
    }
  });

  closeInteressadoModalBtn?.addEventListener("click", closeInteressadoModal);

  interessadoModalBackdrop?.addEventListener("click", (event) => {
    if (event.target === interessadoModalBackdrop) {
      closeInteressadoModal();
    }
  });

  interessadoForm?.addEventListener("submit", handleInteressadoSubmit);
  resetFormBtn?.addEventListener("click", resetInteressadoForm);

  serieForm?.addEventListener("submit", handleSerieSubmit);
  localForm?.addEventListener("submit", handleLocalFormSubmit);
  userForm?.addEventListener("submit", handleUserFormSubmit);

  serieId?.addEventListener("change", updateProgressPreview);
  estudoAtual?.addEventListener("input", updateProgressPreview);

  obsRapidas?.addEventListener("change", (event) => {
    const value = event.target.value;
    if (!value) return;

    observacoes.value = observacoes.value.trim()
      ? `${observacoes.value.trim()} ${value}`
      : value;

    event.target.value = "";
  });

  searchInput?.addEventListener("input", () => renderInteressadosTable());
  statusFilter?.addEventListener("change", () => renderInteressadosTable());
  interestFilter?.addEventListener("change", () => renderInteressadosTable());
  serieFilter?.addEventListener("change", () => renderInteressadosTable());

  const usuariosSearchInput = document.getElementById("usuariosSearchInput");
  usuariosSearchInput?.addEventListener("input", () => renderUsuariosList());

  const locaisSearchInput = document.getElementById("locaisSearchInput");
  locaisSearchInput?.addEventListener("input", () => renderLocaisList());

  const catalogoSearchInput = document.getElementById("catalogoSearchInput");
  catalogoSearchInput?.addEventListener("input", () => renderSeriesList());

  interessadosTable?.addEventListener("click", async (event) => {
    const editBtn = event.target.closest("[data-edit-interessado]");
    const deleteBtn = event.target.closest("[data-delete-interessado]");
    const addStudyBtn = event.target.closest("[data-add-study]");

    if (editBtn) {
      await editInteressado(editBtn.dataset.editInteressado);
    }

    if (addStudyBtn) {
      await advanceInteressadoStudy(addStudyBtn.dataset.addStudy);
    }

    if (deleteBtn) {
      await deleteInteressado(deleteBtn.dataset.deleteInteressado);
    }
  });

  seriesList?.addEventListener("click", async (event) => {
  const editBtn = event.target.closest("[data-edit-serie]");
  const deleteBtn = event.target.closest("[data-delete-serie]");

  if (editBtn) {
    await editSerie(editBtn.dataset.editSerie);
  }

  if (deleteBtn) {
    await deleteSerie(deleteBtn.dataset.deleteSerie);
  }
});

usuariosList?.addEventListener("click", async (event) => {
  const editBtn = event.target.closest("[data-edit-user]");
  const deleteBtn = event.target.closest("[data-delete-user]");

  if (editBtn) {
    await editUser(editBtn.dataset.editUser);
  }

  if (deleteBtn) {
    await deleteUser(deleteBtn.dataset.deleteUser);
  }
});

locaisList?.addEventListener("click", async (event) => {
  const editBtn = event.target.closest("[data-edit-local]");
  const deleteBtn = event.target.closest("[data-delete-local]");

  if (editBtn) {
    await editLocal(editBtn.dataset.editLocal);
  }

  if (deleteBtn) {
    await deleteLocal(deleteBtn.dataset.deleteLocal);
  }
});

  novoUsuarioPerfil?.addEventListener("change", updateUserFormByProfile);

  // RELATÓRIOS LISTENERS
  const reportIgrejaEl = document.getElementById("reportFilterIgreja");
  const reportStatusEl = document.getElementById("reportFilterStatus");
  const reportSerieEl = document.getElementById("reportFilterSerie");
  const reportInteresseEl = document.getElementById("reportFilterInteresse");

  const updateReportsFromFilters = () => {
    reportFilters.igreja = reportIgrejaEl?.value || "Todas";
    reportFilters.status = reportStatusEl?.value || "Todos";
    reportFilters.serie = reportSerieEl?.value || "Todas";
    reportFilters.interesse = reportInteresseEl?.value || "Todos";
    renderRelatoriosSection();
  };

  reportIgrejaEl?.addEventListener("change", updateReportsFromFilters);
  reportStatusEl?.addEventListener("change", updateReportsFromFilters);
  reportSerieEl?.addEventListener("change", updateReportsFromFilters);
  reportInteresseEl?.addEventListener("change", updateReportsFromFilters);

  // Relatório Detalhado (Dossiê Completo com 100% dos dados)
  document.getElementById("btnExportExcelDetalhado")?.addEventListener("click", async () => {
    try {
      showMessage("Gerando arquivo Excel Detalhado...", "info");
      await exportExcelReport({ reportType: "detalhado", filters: reportFilters, state });
      showMessage("Relatório Detalhado em Excel (.xlsx) baixado com sucesso!", "success");
    } catch (err) {
      console.error(err);
      showMessage("Erro ao gerar arquivo Excel Detalhado.", "error");
    }
  });

  document.getElementById("btnExportPdfDetalhado")?.addEventListener("click", async () => {
    try {
      showMessage("Gerando arquivo PDF Detalhado...", "info");
      await exportPdfReport({ reportType: "detalhado", filters: reportFilters, state });
      showMessage("Relatório Detalhado em PDF (.pdf) baixado com sucesso!", "success");
    } catch (err) {
      console.error(err);
      showMessage("Erro ao gerar arquivo PDF Detalhado.", "error");
    }
  });

  // Relatório Geral Completo (Multi-aba)
  document.getElementById("btnExportExcelCompleto")?.addEventListener("click", async () => {
    try {
      showMessage("Gerando arquivo Excel Geral...", "info");
      await exportExcelReport({ reportType: "completo", filters: reportFilters, state });
      showMessage("Relatório Geral em Excel (.xlsx) baixado com sucesso!", "success");
    } catch (err) {
      console.error(err);
      showMessage("Erro ao gerar arquivo Excel.", "error");
    }
  });

  document.getElementById("btnExportPdfCompleto")?.addEventListener("click", async () => {
    try {
      showMessage("Gerando arquivo PDF Geral...", "info");
      await exportPdfReport({ reportType: "completo", filters: reportFilters, state });
      showMessage("Relatório Geral em PDF (.pdf) baixado com sucesso!", "success");
    } catch (err) {
      console.error(err);
      showMessage("Erro ao gerar arquivo PDF.", "error");
    }
  });

  // Delegated buttons for individual module downloads
  document.getElementById("relatoriosSection")?.addEventListener("click", async (e) => {
    const btn = e.target.closest("[data-export-type]");
    if (!btn) return;

    const exportType = btn.dataset.exportType;
    const reportModule = btn.dataset.report;

    try {
      if (exportType === "excel") {
        showMessage(`Gerando relatório de ${reportModule} em Excel...`, "info");
        await exportExcelReport({ reportType: reportModule, filters: reportFilters, state });
        showMessage(`Relatório de ${reportModule} em Excel gerado com sucesso!`, "success");
      } else if (exportType === "pdf") {
        showMessage(`Gerando relatório de ${reportModule} em PDF...`, "info");
        await exportPdfReport({ reportType: reportModule, filters: reportFilters, state });
        showMessage(`Relatório de ${reportModule} em PDF gerado com sucesso!`, "success");
      }
    } catch (err) {
      console.error(err);
      showMessage("Erro ao exportar relatório.", "error");
    }
  });
}

/* =========================================================
   RELATÓRIOS SECTION RENDERING
========================================================= */
const reportFilters = {
  igreja: "Todas",
  status: "Todos",
  serie: "Todas",
  interesse: "Todos"
};

function getUniqueChurchesForReports() {
  const set = new Set();
  if (state.user?.igrejaNome) {
    set.add(state.user.igrejaNome);
  }
  (state.locais || []).forEach((l) => {
    if (l.nome) set.add(l.nome);
  });
  (state.interessados || []).forEach((i) => {
    if (i.igrejaNome) set.add(i.igrejaNome);
  });
  return Array.from(set).sort((a, b) => a.localeCompare(b, "pt-BR"));
}

function renderRelatoriosSection() {
  const container = document.getElementById("relatoriosSection");
  if (!container) return;

  // 1. Populate Igreja filter dropdown
  const igrejaSelect = document.getElementById("reportFilterIgreja");
  if (igrejaSelect) {
    const churches = getUniqueChurchesForReports();
    const currentIgreja = reportFilters.igreja || "Todas";
    igrejaSelect.innerHTML =
      `<option value="Todas" ${currentIgreja === "Todas" ? "selected" : ""}>Todas as Igrejas (${churches.length} cadastradas)</option>` +
      churches
        .map((c) => `<option value="${escapeHtml(c)}" ${c === currentIgreja ? "selected" : ""}>${escapeHtml(c)}</option>`)
        .join("");
  }

  // 2. Populate Serie filter dropdown
  const serieSelect = document.getElementById("reportFilterSerie");
  if (serieSelect) {
    const currentSerie = reportFilters.serie || "Todas";
    serieSelect.innerHTML =
      `<option value="Todas" ${currentSerie === "Todas" ? "selected" : ""}>Todas as Séries (${state.series.length} disponíveis)</option>` +
      state.series
        .map((s) => `<option value="${escapeHtml(s.nome)}" ${s.nome === currentSerie ? "selected" : ""}>${escapeHtml(s.nome)}</option>`)
        .join("");
  }

  // 3. Filter data
  const filtered = filterInteressadosForReport(state.interessados, reportFilters, state);
  const totalChurches = getUniqueChurchesForReports().length;
  const instructorsList = buildInstructorsData(state.interessados, state.usuarios, state);

  // 4. Update preview summary & count
  const summaryEl = document.getElementById("reportPreviewSummary");
  if (summaryEl) {
    summaryEl.innerHTML = `
      <span>📋 <strong>${filtered.length}</strong> Estudantes</span>
      <span>⛪ <strong>${totalChurches}</strong> Igrejas</span>
      <span>👨‍🏫 <strong>${instructorsList.length}</strong> Instrutores</span>
      <span>📖 <strong>${state.series.length}</strong> Séries</span>
    `;
  }

  const countEl = document.getElementById("reportPreviewCount");
  if (countEl) {
    countEl.textContent = `Exibindo ${filtered.length} estudante(s) com base nos filtros selecionados (Igreja: ${reportFilters.igreja}, Status: ${reportFilters.status}, Série: ${reportFilters.serie}).`;
  }

  // 5. Update Preview Table
  const tableBody = document.getElementById("reportPreviewTableBody");
  if (tableBody) {
    if (filtered.length === 0) {
      tableBody.innerHTML = `<tr><td colspan="6" class="empty-state">Nenhum registro encontrado para os filtros selecionados.</td></tr>`;
    } else {
      const topItems = filtered.slice(0, 15);
      tableBody.innerHTML =
        topItems
          .map((item) => {
            const p = calcProgress(item.estudoAtual, item.totalEstudos);

            return `
              <tr>
                <td>
                  <strong>${escapeHtml(item.nome || "-")}</strong><br>
                  <span class="tiny-muted">${escapeHtml(item.telefone)}</span>
                </td>
                <td>${escapeHtml(item.igrejaNome || "-")}</td>
                <td><strong>${escapeHtml(item.serieNome || "-")}</strong></td>
                <td>
                  <div class="progress-meta" style="font-size:11px;margin-bottom:2px;">${p.capped}/${p.total} • ${p.porcentagem}%</div>
                  <div class="progress-bar compact">
                    <span style="width:${p.porcentagem}%"></span>
                  </div>
                  <span class="tiny-muted" style="font-size:11px;">Faltam ${p.faltantes} lições</span>
                </td>
                <td><span class="pill status-${slug(item.status)}">${escapeHtml(item.status || "Ativo")}</span></td>
                <td>
                  <div style="font-weight:600;color:var(--text);font-size:13px;">${escapeHtml(item.instrutoresString || "-")}</div>
                </td>
              </tr>
            `;
          })
          .join("") +
        (filtered.length > 15
          ? `<tr><td colspan="6" style="text-align:center;font-size:12px;color:var(--muted);padding:10px;">+ ${filtered.length - 15} outros estudantes serão incluídos na exportação completa.</td></tr>`
          : "");
    }
  }
}

/* =========================
   AUTH READY
========================= */
window.addEventListener("auth-ready", async (event) => {
  const detail = event.detail;

  if (!detail?.ready) return;

  if (!detail?.profile) {
    state.user = null;
    state.interessados = [];
    state.series = [];
    state.usuarios = [];
    state.locais = [];
    return;
  }

  state.user = {
    uid: detail.profile.uid || detail.profile.id || detail.firebaseUser?.uid || "",
    ...detail.profile
  };

  const allowedPerfis = ["admin", "distrital", "local", "membro"];
  if (!allowedPerfis.includes(state.user.perfil)) {
    showMessage("Perfil sem acesso ao sistema.", "error");
    return;
  }

  applyRoleUI();
  await refreshData();
});

/* =========================
   RENDER FINAL
========================= */
function renderAll() {
  renderFilters();
  renderSerieOptions();
  renderInteressadoLocalOptions();
  renderResponsavelOptions();
  renderStatusAndInterestOptions();
  renderUserProfileOptions();
  renderUserLocalOptions();
  updateProgressPreview();

  if (novoUsuarioDistrito) {
    novoUsuarioDistrito.value = DISTRITO_FIXO;
    novoUsuarioDistrito.readOnly = true;
  }

  if (distrito) {
    distrito.value = DISTRITO_FIXO;
    distrito.readOnly = true;
  }

  renderMetrics();
  renderDashboardScales();
  renderRecentList();
  renderAttentionList();
  renderInteressadosTable();
  renderSeriesList();
  renderUsuariosList();
  renderLocaisList();
  renderInstrutoresSection();
  renderRelatoriosSection();

  applyRoleUI();
}

/* =========================
   START
========================= */
bindStaticEvents();
renderFilters();
renderStatusAndInterestOptions();
renderSerieOptions();
updateProgressPreview();

/* =========================
   DEBUG
========================= */
window.esplanadaApp = {
  state,
  refreshData,
  editInteressado,
  deleteInteressado,
  editSerie,
  deleteSerie,
  editLocal,
  deleteLocal
};
