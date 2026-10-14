import { initializeApp, deleteApp } from "https://www.gstatic.com/firebasejs/10.14.1/firebase-app.js";
import {
  getAuth, onAuthStateChanged, signInWithEmailAndPassword, signOut, createUserWithEmailAndPassword,
  sendPasswordResetEmail, updatePassword, reauthenticateWithCredential, EmailAuthProvider,
} from "https://www.gstatic.com/firebasejs/10.14.1/firebase-auth.js";
import {
  getFirestore, doc, getDoc, getDocs, setDoc, updateDoc, deleteDoc, collection,
} from "https://www.gstatic.com/firebasejs/10.14.1/firebase-firestore.js";
import { firebaseConfig } from "./firebase-config.js";
import { recopilar } from "./contenido.js";
import { construirPptx } from "./pptx.js";

const $ = (s) => document.querySelector(s);
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const setMsg = (sel, texto = "", tipo = "") => { const el = $(sel); el.textContent = texto; el.classList.remove("ok", "error"); if (tipo) el.classList.add(tipo); };

// ---------------------------------------------------------------- fechas / estados
const pad = (n) => String(n).padStart(2, "0");
const aIso = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const hoy = () => aIso(new Date());
const sumarDias = (n) => { const d = new Date(); d.setDate(d.getDate() + n); return aIso(d); };
function diasHasta(iso) {
  const [y, m, d] = iso.split("-").map(Number);
  const h = new Date(); h.setHours(0, 0, 0, 0);
  return Math.round((new Date(y, m - 1, d) - h) / 86400000);
}
function estadoUsuario(u) {
  if (u.activo === false) return "Desactivado";
  if (u.vence) {
    const d = diasHasta(u.vence);
    if (d < 0) return "Vencido";
    if (d <= 7) return "Por vencer";
  }
  return "Activo";
}
function validarAcceso(p) {
  if (!p) return "Tu cuenta no tiene permisos. Contacta al administrador.";
  if (p.activo === false) return "Tu cuenta está desactivada. Contacta al administrador.";
  if (p.vence && diasHasta(p.vence) < 0) return "Tu acceso venció. Contacta al administrador para renovarlo.";
  return null;
}
function errorFirebase(err) {
  const c = (err && err.code) || "";
  if (["auth/invalid-credential", "auth/wrong-password", "auth/user-not-found", "auth/invalid-login-credentials"].includes(c)) return "Correo o contraseña incorrectos.";
  if (c === "auth/invalid-email") return "El correo no es válido.";
  if (c === "auth/too-many-requests") return "Demasiados intentos. Espera unos minutos e intenta de nuevo.";
  if (c === "auth/email-already-in-use") return "Ese correo ya tiene una cuenta. Si fue eliminado antes, bórralo también en Firebase > Authentication.";
  if (c === "auth/weak-password") return "La contraseña es muy débil (mínimo 6 caracteres).";
  if (c === "auth/network-request-failed") return "Sin conexión. Revisa tu internet.";
  if (c === "permission-denied" || c === "auth/requires-recent-login") return "No tienes permiso para esta acción.";
  return "Ocurrió un error: " + ((err && err.message) || c || "desconocido");
}

// ---------------------------------------------------------------- Firebase
const configurado = firebaseConfig.apiKey && !String(firebaseConfig.apiKey).startsWith("PEGA");
let app, auth, db;
let perfil = null;       // datos del usuario conectado
let usuarios = [];       // lista (solo administrador)
let avisoLogin = "";     // mensaje a mostrar al volver al login

function mostrarLogin() {
  $("#vista-app").hidden = true;
  $("#vista-login").hidden = false;
  setMsg("#login-msg", avisoLogin, "error");
  avisoLogin = "";
}

function mostrarApp() {
  $("#vista-login").hidden = true;
  $("#vista-app").hidden = false;
  $("#u-nombre").textContent = perfil.nombre || perfil.email;
  $("#u-rol").textContent = perfil.rol === "admin" ? "Administrador" : "Usuario";
  $("#u-avatar").textContent = (perfil.nombre || perfil.email || "?").charAt(0).toUpperCase();
  $("#tab-admin").hidden = perfil.rol !== "admin";
  activarTab("gen");
  if (perfil.rol === "admin") cargarUsuarios();
}

if (!configurado) {
  $("#vista-login").hidden = false;
  $("#aviso-config").hidden = false;
  $("#form-login button").disabled = true;
} else {
  app = initializeApp(firebaseConfig);
  auth = getAuth(app);
  db = getFirestore(app);

  onAuthStateChanged(auth, async (user) => {
    if (!user) { perfil = null; mostrarLogin(); return; }
    try {
      const snap = await getDoc(doc(db, "usuarios", user.uid));
      const datos = snap.exists() ? snap.data() : null;
      const motivo = validarAcceso(datos);
      if (motivo) { avisoLogin = motivo; await signOut(auth); return; }
      perfil = { uid: user.uid, email: user.email, ...datos };
      mostrarApp();
    } catch (err) {
      avisoLogin = "No se pudo verificar tu cuenta. Intenta de nuevo.";
      await signOut(auth);
    }
  });
}

// ---------------------------------------------------------------- login / sesión
$("#form-login").addEventListener("submit", async (e) => {
  e.preventDefault();
  if (!configurado) return;
  const btn = e.submitter || $("#form-login button");
  btn.disabled = true;
  setMsg("#login-msg", "");
  try {
    await signInWithEmailAndPassword(auth, $("#login-email").value.trim(), $("#login-pass").value);
    $("#login-pass").value = "";
  } catch (err) {
    setMsg("#login-msg", errorFirebase(err), "error");
  } finally {
    btn.disabled = false;
  }
});

$("#btn-salir").addEventListener("click", () => signOut(auth));

// Cambiar mi contraseña
$("#btn-clave").addEventListener("click", () => { $("#form-clave").reset(); setMsg("#c-msg"); $("#dlg-clave").showModal(); });
$("#c-cancelar").addEventListener("click", () => $("#dlg-clave").close());
$("#form-clave").addEventListener("submit", async (e) => {
  e.preventDefault();
  setMsg("#c-msg");
  try {
    const user = auth.currentUser;
    await reauthenticateWithCredential(user, EmailAuthProvider.credential(user.email, $("#c-actual").value));
    await updatePassword(user, $("#c-nueva").value);
    setMsg("#c-msg", "Contraseña actualizada.", "ok");
    setTimeout(() => $("#dlg-clave").close(), 1200);
  } catch (err) {
    setMsg("#c-msg", errorFirebase(err), "error");
  }
});

// ---------------------------------------------------------------- pestañas
function activarTab(nombre) {
  document.querySelectorAll(".tab").forEach((t) => t.classList.toggle("activa", t.dataset.tab === nombre));
  $("#panel-gen").hidden = nombre !== "gen";
  $("#panel-admin").hidden = nombre !== "admin";
}
$("#tabs").addEventListener("click", (e) => {
  const t = e.target.closest(".tab");
  if (t && !t.hidden) activarTab(t.dataset.tab);
});

// ---------------------------------------------------------------- administración
const adminsActivos = () => usuarios.filter((u) => u.rol === "admin" && u.activo !== false);
const esUltimoAdmin = (uid) => { const a = adminsActivos(); return a.length === 1 && a[0].uid === uid; };

async function cargarUsuarios() {
  try {
    const snap = await getDocs(collection(db, "usuarios"));
    usuarios = snap.docs.map((d) => ({ uid: d.id, ...d.data() }))
      .sort((a, b) => (a.nombre || "").localeCompare(b.nombre || ""));
    renderAdmin();
  } catch (err) {
    setMsg("#msg-admin", errorFirebase(err), "error");
  }
}

function renderAdmin() {
  const estados = usuarios.map(estadoUsuario);
  const n = (x) => estados.filter((e) => e === x).length;
  $("#stats").innerHTML =
    `<div class="stat"><b>${estados.length}</b><span>Usuarios totales</span></div>` +
    `<div class="stat ok"><b>${n("Activo") + n("Por vencer")}</b><span>Con acceso</span></div>` +
    `<div class="stat warn"><b>${n("Por vencer")}</b><span>Vencen en 7 días</span></div>` +
    `<div class="stat bad"><b>${n("Vencido") + n("Desactivado")}</b><span>Vencidos o desactivados</span></div>`;

  $("#tabla-usuarios").innerHTML = usuarios.map((u) => {
    const est = estadoUsuario(u);
    const cls = est.toLowerCase().replace(" ", "-");
    return `<tr data-uid="${esc(u.uid)}">
      <td><b>${esc(u.nombre || "")}</b><div class="muted">${esc(u.email || "")}</div></td>
      <td>${u.rol === "admin" ? "Administrador" : "Usuario"}</td>
      <td><span class="badge ${cls}">${est}</span></td>
      <td>${esc(u.vence || "Sin vencimiento")}</td>
      <td>
        <button class="btn-mini" data-acc="editar" type="button">Editar</button>
        <button class="btn-mini" data-acc="reset" type="button" title="Envía un correo para que elija una nueva contraseña">Enviar nueva clave</button>
        <button class="btn-mini peligro" data-acc="borrar" type="button">Eliminar</button>
      </td></tr>`;
  }).join("");
}

// Crear usuario (usa una conexión secundaria para no cerrar la sesión del administrador)
$("#a-vence").value = sumarDias(30);
$("#a-sinvence").addEventListener("change", (e) => { $("#a-vence").disabled = e.target.checked; });
$("#form-alta").addEventListener("submit", async (e) => {
  e.preventDefault();
  const btn = e.submitter;
  const datos = {
    nombre: $("#a-nombre").value.trim(),
    email: $("#a-email").value.trim().toLowerCase(),
    clave: $("#a-clave").value,
    rol: $("#a-rol").value,
    vence: $("#a-sinvence").checked ? null : ($("#a-vence").value || null),
  };
  if (usuarios.some((u) => (u.email || "").toLowerCase() === datos.email)) { setMsg("#msg-alta", "Ese correo ya está registrado.", "error"); return; }
  btn.disabled = true;
  setMsg("#msg-alta", "Creando usuario…");
  const sec = initializeApp(firebaseConfig, "alta-" + Date.now());
  try {
    const secAuth = getAuth(sec);
    const cred = await createUserWithEmailAndPassword(secAuth, datos.email, datos.clave);
    await setDoc(doc(db, "usuarios", cred.user.uid), {
      nombre: datos.nombre, email: datos.email, rol: datos.rol, activo: true, vence: datos.vence, creado: hoy(),
    });
    await signOut(secAuth);
    $("#form-alta").reset();
    $("#a-vence").value = sumarDias(30); $("#a-vence").disabled = false;
    setMsg("#msg-alta", `Usuario «${datos.email}» creado. Entrégale su correo y la contraseña inicial.`, "ok");
    await cargarUsuarios();
  } catch (err) {
    setMsg("#msg-alta", errorFirebase(err), "error");
  } finally {
    await deleteApp(sec);
    btn.disabled = false;
  }
});

// Acciones de la tabla
let editando = null;
$("#tabla-usuarios").addEventListener("click", async (e) => {
  const b = e.target.closest("button[data-acc]");
  if (!b) return;
  const uid = b.closest("tr").dataset.uid;
  const u = usuarios.find((x) => x.uid === uid);
  if (!u) return;
  setMsg("#msg-admin");

  if (b.dataset.acc === "editar") {
    editando = u;
    $("#e-email").textContent = u.email || "";
    $("#e-nombre").value = u.nombre || "";
    $("#e-rol").value = u.rol || "usuario";
    $("#e-activo").checked = u.activo !== false;
    $("#e-sinvence").checked = !u.vence;
    $("#e-vence").value = u.vence || sumarDias(30);
    $("#e-vence").disabled = !u.vence;
    setMsg("#e-msg");
    $("#dlg-edit").showModal();
  }

  if (b.dataset.acc === "reset") {
    try {
      await sendPasswordResetEmail(auth, u.email);
      setMsg("#msg-admin", `Se envió a ${u.email} un correo para elegir una nueva contraseña.`, "ok");
    } catch (err) { setMsg("#msg-admin", errorFirebase(err), "error"); }
  }

  if (b.dataset.acc === "borrar") {
    if (uid === perfil.uid) { setMsg("#msg-admin", "No puedes eliminar tu propia cuenta.", "error"); return; }
    if (esUltimoAdmin(uid)) { setMsg("#msg-admin", "Debe quedar al menos un administrador activo.", "error"); return; }
    if (!confirm(`¿Eliminar a ${u.nombre || u.email}? Perderá el acceso de inmediato.`)) return;
    try {
      await deleteDoc(doc(db, "usuarios", uid));
      setMsg("#msg-admin", `Usuario «${u.email}» eliminado.`, "ok");
      await cargarUsuarios();
    } catch (err) { setMsg("#msg-admin", errorFirebase(err), "error"); }
  }
});

$("#e-sinvence").addEventListener("change", (e) => { $("#e-vence").disabled = e.target.checked; });
$("#e-cancelar").addEventListener("click", () => $("#dlg-edit").close());
$("#form-edit").addEventListener("submit", async (e) => {
  e.preventDefault();
  const u = editando;
  const cambios = {
    nombre: $("#e-nombre").value.trim(),
    rol: $("#e-rol").value,
    activo: $("#e-activo").checked,
    vence: $("#e-sinvence").checked ? null : ($("#e-vence").value || null),
  };
  const dejaDeSerAdminActivo = (cambios.rol !== "admin" || !cambios.activo);
  if (u.uid === perfil.uid && (!cambios.activo || cambios.rol !== "admin")) { setMsg("#e-msg", "No puedes quitarte tu propio acceso de administrador.", "error"); return; }
  if (dejaDeSerAdminActivo && esUltimoAdmin(u.uid)) { setMsg("#e-msg", "Debe quedar al menos un administrador activo.", "error"); return; }
  try {
    await updateDoc(doc(db, "usuarios", u.uid), cambios);
    $("#dlg-edit").close();
    setMsg("#msg-admin", `Cambios guardados para «${u.email}».`, "ok");
    await cargarUsuarios();
  } catch (err) { setMsg("#e-msg", errorFirebase(err), "error"); }
});

// ---------------------------------------------------------------- generador
$("#num").addEventListener("input", (e) => { $("#num-val").textContent = e.target.value; });

async function cargarImagen(url) {
  try {
    const r = await fetch(url);
    if (!r.ok) return null;
    const blob = await r.blob();
    if (!/^image\/(jpeg|png)/.test(blob.type)) return null;
    const data = await new Promise((res, rej) => { const fr = new FileReader(); fr.onload = () => res(fr.result); fr.onerror = rej; fr.readAsDataURL(blob); });
    const dim = await new Promise((res) => { const im = new Image(); im.onload = () => res({ w: im.naturalWidth, h: im.naturalHeight }); im.onerror = () => res(null); im.src = data; });
    return dim && dim.w && dim.h ? { data, ...dim } : null;
  } catch { return null; }
}

let ultima = null;   // { pptx, nombre }
const nombreArchivo = (tema) => (tema.replace(/[^\p{L}\p{N}]+/gu, "_").replace(/^_+|_+$/g, "").slice(0, 60) || "presentacion") + ".pptx";

$("#form-gen").addEventListener("submit", async (e) => {
  e.preventDefault();
  const tema = $("#tema").value.trim();
  if (!tema) { setMsg("#estado", "Escribe un tema válido.", "error"); return; }
  const btn = $("#btn-gen");
  btn.disabled = true;
  $("#resultado").hidden = true;
  try {
    setMsg("#estado", "Reuniendo información…");
    const datos = await recopilar(tema, Number($("#num").value));
    if (!datos) { setMsg("#estado", "No encontré información suficiente sobre ese tema. Prueba con otra redacción.", "error"); return; }

    setMsg("#estado", "Preparando imágenes…");
    const [portadaImg] = await Promise.all([
      datos.portada ? cargarImagen(datos.portada.url) : null,
      ...datos.slides.map(async (s) => { s.img = s.imagen ? await cargarImagen(s.imagen.url) : null; }),
    ]);

    setMsg("#estado", "Armando el PowerPoint…");
    const pptx = construirPptx(window.PptxGenJS, {
      tema: datos.tema, slides: datos.slides, fuentes: datos.fuentes,
      portada: datos.portada, portadaImg, incluirYT: $("#yt").checked,
    });
    const nombre = nombreArchivo(tema);
    await pptx.writeFile({ fileName: nombre });
    ultima = { pptx, nombre };

    $("#lista-slides").innerHTML = ["Portada", ...datos.slides.map((s) => s.titulo), "Fuentes y referencias"].map((t) => `<li>${esc(t)}</li>`).join("");
    $("#resultado").hidden = false;
    setMsg("#estado", `Presentación lista: ${datos.slides.length} diapositivas de contenido.`, "ok");
  } catch (err) {
    setMsg("#estado", "Ocurrió un error al generar la presentación. Intenta de nuevo.", "error");
    console.error(err);
  } finally {
    btn.disabled = false;
  }
});

$("#btn-re").addEventListener("click", () => { if (ultima) ultima.pptx.writeFile({ fileName: ultima.nombre }); });
