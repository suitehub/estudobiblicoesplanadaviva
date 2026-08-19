import {
  signInWithEmailAndPassword,
  onAuthStateChanged,
  signOut,
  setPersistence,
  browserLocalPersistence
} from "https://www.gstatic.com/firebasejs/12.11.0/firebase-auth.js";

import {
  doc,
  getDoc,
  getDocs,
  collection,
  query,
  where,
  limit
} from "https://www.gstatic.com/firebasejs/12.11.0/firebase-firestore.js";

import { auth, db } from "./firebase-config.js";

/* =========================
   ELEMENTOS
========================= */
const loginForm = document.getElementById("loginForm");

const loginAccessType = document.getElementById("loginAccessType");
const loginLocal = document.getElementById("loginLocal");
const loginUsername = document.getElementById("loginUsername");
const loginEmail = document.getElementById("loginEmail");
const loginPassword = document.getElementById("loginPassword");

const loginLocalGroup = document.getElementById("loginLocalGroup");
const loginUsernameGroup = document.getElementById("loginUsernameGroup");
const loginEmailGroup = document.getElementById("loginEmailGroup");

const authMessage = document.getElementById("authMessage");
const globalMessage = document.getElementById("globalMessage");
const logoutBtn = document.getElementById("logoutBtn");

const authScreen = document.getElementById("authScreen");
const appScreen = document.getElementById("appScreen");
const loadingOverlay = document.getElementById("loadingOverlay");

const loggedUserName = document.getElementById("loggedUserName");
const loggedUserRole = document.getElementById("loggedUserRole");
const userAvatar = document.getElementById("userAvatar");

/* =========================
   STORE
========================= */
const authStore = {
  firebaseUser: null,
  profile: null,
  ready: false
};

window.authStore = authStore;

/* =========================
   HELPERS
========================= */
function showLoading() {
  loadingOverlay?.classList.remove("hidden");
}

function hideLoading() {
  loadingOverlay?.classList.add("hidden");
}

function showMessage(target, message, type = "error") {
  if (!target) return;
  target.textContent = message;
  target.classList.remove("hidden", "success", "error", "info");
  target.classList.add(type);
}

function hideMessage(target) {
  if (!target) return;
  target.textContent = "";
  target.classList.add("hidden");
  target.classList.remove("success", "error", "info");
}

function normalizeText(text = "") {
  return String(text)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();
}

function getInitials(name = "") {
  const parts = name.trim().split(/\s+/).filter(Boolean);
  if (!parts.length) return "EV";
  if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase();
  return (parts[0][0] + parts[1][0]).toUpperCase();
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
      return "Sem perfil";
  }
}

function isGlobalAccessType(type) {
  return type === "admin" || type === "distrital";
}

function isLocalAccessType(type) {
  return type === "local" || type === "membro";
}

function clearAuthStore() {
  authStore.firebaseUser = null;
  authStore.profile = null;
  authStore.ready = true;
}

function setAuthStore(firebaseUser, profile) {
  authStore.firebaseUser = firebaseUser;
  authStore.profile = profile;
  authStore.ready = true;
}

function dispatchAuthReady() {
  window.dispatchEvent(
    new CustomEvent("auth-ready", {
      detail: {
        firebaseUser: authStore.firebaseUser,
        profile: authStore.profile,
        ready: authStore.ready
      }
    })
  );
}

function showAuthScreen() {
  authScreen?.classList.remove("hidden");
  appScreen?.classList.add("hidden");
}

function showAppScreen() {
  authScreen?.classList.add("hidden");
  appScreen?.classList.remove("hidden");
}

function applyUserUI(profile) {
  if (!profile) return;

  if (loggedUserName) loggedUserName.textContent = profile.nome || "Usuário";
  if (loggedUserRole) loggedUserRole.textContent = formatRole(profile.perfil);
  if (userAvatar) userAvatar.textContent = getInitials(profile.nome || "EV");
}

function updateLoginModeUI() {
  const accessType = loginAccessType?.value || "admin";
  const globalMode = isGlobalAccessType(accessType);
  const localMode = isLocalAccessType(accessType);

  if (loginEmailGroup) loginEmailGroup.style.display = globalMode ? "" : "none";
  if (loginLocalGroup) loginLocalGroup.style.display = localMode ? "" : "none";
  if (loginUsernameGroup) loginUsernameGroup.style.display = localMode ? "" : "none";

  if (loginEmail) {
    loginEmail.required = globalMode;
    if (!globalMode) loginEmail.value = "";
  }

  if (loginLocal) {
    loginLocal.required = localMode;
    if (!localMode) loginLocal.value = "";
  }

  if (loginUsername) {
    loginUsername.required = localMode;
    if (!localMode) loginUsername.value = "";
  }

  if (loginPassword) {
    loginPassword.required = true;
  }
}

function mapFirebaseAuthError(error) {
  const code = error?.code || "";

  switch (code) {
    case "auth/invalid-email":
      return "E-mail inválido.";
    case "auth/missing-password":
      return "Digite sua senha.";
    case "auth/invalid-credential":
      return "Credenciais inválidas.";
    case "auth/user-disabled":
      return "Este usuário foi desativado.";
    case "auth/too-many-requests":
      return "Muitas tentativas. Tente novamente mais tarde.";
    case "auth/network-request-failed":
      return "Falha de conexão. Verifique sua internet.";
    default:
      return error?.message || "Não foi possível entrar.";
  }
}

/* =========================
   LOAD LOCAIS PARA LOGIN
========================= */
async function loadLoginLocais() {
  if (!loginLocal) return;

  try {
    const snap = await getDocs(collection(db, "locais"));
    const locais = snap.docs
      .map((item) => ({ id: item.id, ...item.data() }))
      .filter((item) => item.ativo !== false)
      .sort((a, b) => (a.nome || "").localeCompare(b.nome || "", "pt-BR"));

    loginLocal.innerHTML =
      `<option value="">Selecionar igreja</option>` +
      locais
        .map((item) => {
          const label = `${item.nome}${item.tipo ? ` (${item.tipo})` : ""}`;
          return `<option value="${item.id}">${label}</option>`;
        })
        .join("");
  } catch (error) {
    console.error("Erro ao carregar locais no login:", error);
  }
}

/* =========================
   PERFIL DO USUÁRIO
========================= */
async function fetchUserProfile(uid) {
  const ref = doc(db, "usuarios", uid);
  const snap = await getDoc(ref);

  if (!snap.exists()) {
    throw new Error("Perfil do usuário não encontrado no banco.");
  }

  const data = snap.data();

  if (!data.ativo) {
    throw new Error("Seu acesso está inativo. Fale com o administrador.");
  }

  return {
    uid,
    id: snap.id,
    ...data
  };
}

/* =========================
   LOGIN INDEX
========================= */
async function fetchLoginIndexRecord({ perfil, username, igrejaId }) {
  const q = query(
    collection(db, "login_index"),
    where("perfil", "==", perfil),
    where("username", "==", username),
    where("igrejaId", "==", igrejaId),
    where("ativo", "==", true),
    limit(1)
  );

  const snap = await getDocs(q);

  if (snap.empty) {
    throw new Error("Usuário não encontrado para essa igreja.");
  }

  const docSnap = snap.docs[0];
  const data = docSnap.data();

  if (!data.emailAuth) {
    throw new Error("Esse acesso não possui credencial válida.");
  }

  return {
    id: docSnap.id,
    ...data
  };
}

/* =========================
   LOGIN GLOBAL
========================= */
async function handleGlobalLogin({ accessType, email, password }) {
  if (!email || !password) {
    throw new Error("Preencha e-mail e senha.");
  }

  await setPersistence(auth, browserLocalPersistence);

  const credential = await signInWithEmailAndPassword(auth, email, password);
  const profile = await fetchUserProfile(credential.user.uid);

  if (profile.perfil !== accessType) {
    await signOut(auth).catch(() => {});
    throw new Error("Esse acesso não corresponde ao perfil selecionado.");
  }

  setAuthStore(credential.user, profile);
  applyUserUI(profile);
  showAppScreen();
  dispatchAuthReady();
}

/* =========================
   LOGIN LOCAL / MEMBRO
========================= */
async function handleLocalLogin({ accessType, igrejaId, username, password }) {
  if (!igrejaId) {
    throw new Error("Selecione a igreja.");
  }

  if (!username) {
    throw new Error("Digite o usuário.");
  }

  if (!password) {
    throw new Error("Digite a senha.");
  }

  await setPersistence(auth, browserLocalPersistence);

  const loginRecord = await fetchLoginIndexRecord({
    perfil: accessType,
    username: normalizeText(username),
    igrejaId
  });

  const credential = await signInWithEmailAndPassword(
    auth,
    loginRecord.emailAuth,
    password
  );

  const profile = await fetchUserProfile(credential.user.uid);

  if (profile.perfil !== accessType) {
    await signOut(auth).catch(() => {});
    throw new Error("Esse acesso não corresponde ao perfil selecionado.");
  }

  if (profile.igrejaId !== igrejaId) {
    await signOut(auth).catch(() => {});
    throw new Error("Esse usuário não pertence à igreja selecionada.");
  }

  if (normalizeText(profile.username || "") !== normalizeText(username)) {
    await signOut(auth).catch(() => {});
    throw new Error("Usuário inválido.");
  }

  setAuthStore(credential.user, profile);
  applyUserUI(profile);
  showAppScreen();
  dispatchAuthReady();
}

/* =========================
   SUBMIT LOGIN
========================= */
async function handleLogin(event) {
  event.preventDefault();
  hideMessage(authMessage);

  const accessType = loginAccessType?.value || "admin";

  try {
    showLoading();

    if (isGlobalAccessType(accessType)) {
      await handleGlobalLogin({
        accessType,
        email: loginEmail?.value.trim() || "",
        password: loginPassword?.value || ""
      });
    } else {
      await handleLocalLogin({
        accessType,
        igrejaId: loginLocal?.value || "",
        username: loginUsername?.value || "",
        password: loginPassword?.value || ""
      });
    }

    showMessage(
      globalMessage,
      `Bem-vindo, ${authStore.profile?.nome || "usuário"}.`,
      "success"
    );
  } catch (error) {
    console.error("Erro no login:", error);

    clearAuthStore();
    showAuthScreen();

    const accessTypeNow = loginAccessType?.value || "admin";
    const message = isGlobalAccessType(accessTypeNow)
      ? mapFirebaseAuthError(error)
      : (error?.message || "Não foi possível entrar.");

    showMessage(authMessage, message, "error");
    dispatchAuthReady();
  } finally {
    hideLoading();
  }
}

/* =========================
   LOGOUT
========================= */
async function handleLogout() {
  try {
    showLoading();

    await signOut(auth);

    clearAuthStore();
    loginForm?.reset();
    updateLoginModeUI();

    showAuthScreen();
    hideMessage(globalMessage);
    dispatchAuthReady();
  } catch (error) {
    console.error("Erro ao sair:", error);
    showMessage(globalMessage, "Não foi possível encerrar a sessão.", "error");
  } finally {
    hideLoading();
  }
}

/* =========================
   BOOTSTRAP AUTH
========================= */
function bootstrapAuth() {
  showLoading();

  onAuthStateChanged(auth, async (user) => {
    try {
      hideMessage(authMessage);

      if (!user) {
        clearAuthStore();
        showAuthScreen();
        dispatchAuthReady();
        return;
      }

      const profile = await fetchUserProfile(user.uid);

      setAuthStore(user, profile);
      applyUserUI(profile);
      showAppScreen();
      dispatchAuthReady();
    } catch (error) {
      console.error("Erro ao validar sessão:", error);

      clearAuthStore();

      await signOut(auth).catch(() => {});

      showAuthScreen();
      showMessage(
        authMessage,
        error?.message || "Sua sessão não pôde ser validada.",
        "error"
      );
      dispatchAuthReady();
    } finally {
      hideLoading();
    }
  });
}

/* =========================
   EVENTOS
========================= */
if (loginAccessType) {
  loginAccessType.addEventListener("change", updateLoginModeUI);
}

if (loginForm) {
  loginForm.addEventListener("submit", handleLogin);
}

if (logoutBtn) {
  logoutBtn.addEventListener("click", handleLogout);
}

/* =========================
   START
========================= */
await loadLoginLocais();
updateLoginModeUI();
bootstrapAuth();

export {
  authStore,
  handleLogout
};
