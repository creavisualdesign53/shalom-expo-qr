
import './style.css'
import * as XLSX from 'xlsx'
import QRCode from 'qrcode'
import { jsPDF } from 'jspdf'
import { Html5Qrcode } from 'html5-qrcode'

import { firebaseConfig, auth, db } from './firebase'
import { initializeApp } from 'firebase/app'

import {
  signInWithEmailAndPassword,
  onAuthStateChanged,
  signOut,
  signInAnonymously,
  getAuth,
  createUserWithEmailAndPassword
} from 'firebase/auth'

import {
  getFirestore,
  doc,
  getDoc,
  setDoc,
  collection,
  getDocs,
  addDoc,
  query,
  where,
  limit,
  updateDoc,
  deleteDoc,
  runTransaction,
  serverTimestamp,
  writeBatch
} from 'firebase/firestore'

const app = document.querySelector('#app')
const rutaActual = window.location.pathname.replace(/\/+$/, '') || '/'

const modoJurado =
  new URLSearchParams(window.location.search).get('modo') ||
  (
    rutaActual === '/staff' ? 'staff' :
    rutaActual === '/jurado' ? 'evaluacion-jurado' :
    null
  )
let administrador = null
let juradoActual = null
let temporizadorCorreccionJurado = null;
function escapar(texto) {
  return String(texto ?? '').replace(/[&<>"']/g, caracter => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  })[caracter])
}
async function cerrarSesion() {
  try {
    await signOut(auth)
  } catch (error) {
    console.error('Error al cerrar sesión:', error)
    alert('No se pudo cerrar la sesión.')
  }
}

const appUsuarios = initializeApp(firebaseConfig, 'shalom-usuarios')
const authUsuarios = getAuth(appUsuarios)

const appJurados = initializeApp(firebaseConfig, 'shalom-jurados')
const authJurados = getAuth(appJurados)
const dbJurados = getFirestore(appJurados)
// =====================================================
// CONEXIÓN INDEPENDIENTE PARA PERSONAL STAFF
// =====================================================

const appStaff = initializeApp(firebaseConfig, 'shalom-staff')
const authStaff = getAuth(appStaff)
const dbStaff = getFirestore(appStaff)

let staffActual = null
async function crearUsuarioSistema(correo, clave, nombre, rol) {
  const usuarioActual = auth.currentUser

  if (!usuarioActual) {
    throw new Error('No hay un administrador conectado.')
  }

  const credencial = await createUserWithEmailAndPassword(
    authUsuarios,
    correo,
    clave
  )

  const uid = credencial.user.uid

  await setDoc(doc(db, 'usuarios', uid), {
    nombre: nombre,
    correo: correo,
    rol: rol,
    activo: true,
    creadoEn: new Date()
  })

  await signOut(authUsuarios)

  return uid
}

function escaparTextoUsuarios(valor) {
  return String(valor ?? '').replace(/[&<>"']/g, caracter => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;'
  })[caracter])
}

async function mostrarUsuarios() {
  const contenido = document.querySelector('#contenido');

  if (administrador?.rol !== 'administrador') {
    contenido.innerHTML = '<h2>No tienes permiso para administrar usuarios.</h2>';
    return;
  }

  contenido.innerHTML = `
    <section class="usuarios-panel">

      <header class="usuarios-encabezado">
  <div>
    <span class="usuarios-etiqueta">GESTIÓN INSTITUCIONAL</span>
    <h1>Administrar usuarios</h1>
    <p>Controla las cuentas, los permisos y el acceso al sistema.</p>
  </div>

  <div style="display:flex; align-items:center; gap:15px;">
    <button type="button" id="volver-inicio-usuarios"
      class="asistencia-btn asistencia-btn-principal">
      ← INICIO
    </button>

    <div class="usuarios-icono-principal">♙</div>
  </div>
</header>

      <div class="usuarios-estadisticas">
        <article class="usuario-stat stat-total">
          <div class="usuario-stat-icono">♙</div>
          <div>
            <span>Total de usuarios</span>
            <strong id="stat-total">0</strong>
          </div>
        </article>

        <article class="usuario-stat stat-activos">
          <div class="usuario-stat-icono">✓</div>
          <div>
            <span>Usuarios activos</span>
            <strong id="stat-activos">0</strong>
          </div>
        </article>

        <article class="usuario-stat stat-inactivos">
          <div class="usuario-stat-icono">◷</div>
          <div>
            <span>Usuarios inactivos</span>
            <strong id="stat-inactivos">0</strong>
          </div>
        </article>

        <article class="usuario-stat stat-admins">
          <div class="usuario-stat-icono">★</div>
          <div>
            <span>Administradores</span>
            <strong id="stat-admins">0</strong>
          </div>
        </article>
      </div>

      <div class="usuarios-tarjeta usuarios-crear">
        <div class="usuarios-titulo-seccion">
          <div>
            <span class="usuarios-etiqueta">NUEVA CUENTA</span>
            <h2>Registrar usuario</h2>
            <p>Completa los datos para crear un nuevo acceso.</p>
          </div>
          <span class="usuarios-titulo-icono">＋</span>
        </div>

        <form id="form-crear-usuario" class="usuarios-formulario">

          <div class="usuarios-campo">
            <label for="usuario-nombre">Nombre completo</label>
            <input id="usuario-nombre" type="text"
              placeholder="Ej. Juan Pérez" required>
          </div>

          <div class="usuarios-campo">
            <label for="usuario-correo">Correo electrónico</label>
            <input id="usuario-correo" type="email"
              placeholder="correo@ejemplo.com" required>
          </div>

          <div class="usuarios-campo">
            <label for="usuario-clave">Contraseña inicial</label>
            <input id="usuario-clave" type="password"
              placeholder="Mínimo 6 caracteres" minlength="6" required>
          </div>

          <div class="usuarios-campo">
            <label for="usuario-rol">Rol de acceso</label>
            <select id="usuario-rol" required>
              <option value="">Seleccionar rol</option>
              <option value="administrador">Administrador</option>
              <option value="director">Director</option>
              <option value="coordinador">Coordinador</option>
              <option value="secretaria">Secretaría</option>
            </select>
          </div>

          <div class="usuarios-form-acciones">
            <button type="submit" class="usuarios-btn-crear">
              ＋ Crear usuario
            </button>
          </div>
        </form>

        <p id="mensaje-usuarios" class="usuarios-mensaje" role="status"></p>
      </div>

      <div class="usuarios-tarjeta usuarios-listado">
        <div class="usuarios-titulo-seccion">
          <div>
            <span class="usuarios-etiqueta">DIRECTORIO INSTITUCIONAL</span>
            <h2>Usuarios registrados</h2>
            <p>Consulta y administra los accesos del sistema.</p>
          </div>
          <span class="usuarios-contador" id="usuarios-contador">0 usuarios</span>
        </div>

        <div class="usuarios-filtros">
          <div class="usuarios-buscador">
            <span>⌕</span>
            <input type="search" id="buscar-usuario"
              placeholder="Buscar por nombre o correo...">
          </div>

          <select id="filtro-rol" aria-label="Filtrar por rol">
            <option value="">Todos los roles</option>
            <option value="administrador">Administrador</option>
            <option value="director">Director</option>
            <option value="coordinador">Coordinador</option>
            <option value="secretaria">Secretaría</option>
          </select>

          <select id="filtro-estado" aria-label="Filtrar por estado">
            <option value="">Todos los estados</option>
            <option value="activo">Activos</option>
            <option value="inactivo">Inactivos</option>
          </select>
        </div>

        <div id="lista-usuarios" class="usuarios-tabla-contenedor">
          <div class="usuarios-cargando">Cargando usuarios...</div>
        </div>
      </div>

      <p class="usuarios-pie">SHALOM EXPO 2026 · Administración institucional</p>
    </section>
  `;
document.querySelector('#volver-inicio-usuarios')
  .addEventListener('click', () => {
    mostrarPanel(administrador?.nombre || 'Administrador');
  });
  const formulario = document.querySelector('#form-crear-usuario');
  const mensaje = document.querySelector('#mensaje-usuarios');
  const lista = document.querySelector('#lista-usuarios');
  const buscador = document.querySelector('#buscar-usuario');
  const filtroRol = document.querySelector('#filtro-rol');
  const filtroEstado = document.querySelector('#filtro-estado');

  let usuariosActuales = [];

  const nombreRol = rol => ({
    administrador: 'Administrador',
    director: 'Director',
    coordinador: 'Coordinador',
    secretaria: 'Secretaría'
  }[rol] || rol || 'Sin rol');

  function actualizarEstadisticas() {
    const total = usuariosActuales.length;
    const activos = usuariosActuales.filter(u => u.activo === true).length;
    const admins = usuariosActuales.filter(u => u.rol === 'administrador').length;

    document.querySelector('#stat-total').textContent = total;
    document.querySelector('#stat-activos').textContent = activos;
    document.querySelector('#stat-inactivos').textContent = total - activos;
    document.querySelector('#stat-admins').textContent = admins;
  }

  function dibujarUsuarios() {
    const texto = buscador.value.trim().toLowerCase();
    const rol = filtroRol.value;
    const estado = filtroEstado.value;

    const filtrados = usuariosActuales.filter(usuario => {
      const coincideTexto =
        (usuario.nombre || '').toLowerCase().includes(texto) ||
        (usuario.correo || '').toLowerCase().includes(texto);

      const coincideRol = !rol || usuario.rol === rol;

      const coincideEstado = !estado ||
        (estado === 'activo' && usuario.activo === true) ||
        (estado === 'inactivo' && usuario.activo !== true);

      return coincideTexto && coincideRol && coincideEstado;
    });

    document.querySelector('#usuarios-contador').textContent =
      `${filtrados.length} ${filtrados.length === 1 ? 'usuario' : 'usuarios'}`;

    if (filtrados.length === 0) {
      lista.innerHTML = `
        <div class="usuarios-vacio">
          <span>⌕</span>
          <h3>No se encontraron usuarios</h3>
          <p>Prueba con otro nombre o cambia los filtros.</p>
        </div>
      `;
      return;
    }

    lista.innerHTML = `
      <table class="usuarios-tabla">
        <thead>
          <tr>
            <th>USUARIO</th>
            <th>ROL</th>
            <th>ESTADO</th>
            <th>ACCIONES</th>
          </tr>
        </thead>
        <tbody>
          ${filtrados.map(usuario => {
            const activo = usuario.activo === true;
            const uid = escaparTextoUsuarios(usuario.uid);
            const nombre = escaparTextoUsuarios(usuario.nombre || 'Sin nombre');
            const correo = escaparTextoUsuarios(usuario.correo || 'Sin correo');
            const rolTexto = escaparTextoUsuarios(nombreRol(usuario.rol));
            const inicial = escaparTextoUsuarios(
              (usuario.nombre || 'U').trim().charAt(0).toUpperCase()
            );

            return `
              <tr>
                <td>
                  <div class="usuario-identidad">
                    <div class="usuario-avatar">${inicial}</div>
                    <div class="usuario-datos">
                      <strong>${nombre}</strong>
                      <span>${correo}</span>
                    </div>
                  </div>
                </td>
                <td><span class="usuario-rol">${rolTexto}</span></td>
                <td>
                  <span class="usuario-estado ${activo ? 'estado-activo' : 'estado-inactivo'}">
                    <i></i>${activo ? 'Activo' : 'Inactivo'}
                  </span>
                </td>
                <td>
                  <div class="usuario-acciones">
                    <button type="button" class="usuario-btn-editar"
                      data-accion="editar" data-uid="${uid}">
                      Editar
                    </button>
                    <button type="button"
                      class="usuario-btn-estado ${activo ? 'es-desactivar' : 'es-activar'}"
                      data-accion="estado" data-uid="${uid}"
                      data-activo="${activo}">
                      ${activo ? 'Desactivar' : 'Activar'}
                    </button>
                    <button type="button" class="usuario-btn-eliminar"
                      data-accion="eliminar" data-uid="${uid}">
                      Eliminar
                    </button>
                  </div>
                </td>
              </tr>
            `;
          }).join('')}
        </tbody>
      </table>
    `;
  }

  async function cargarUsuarios() {
    lista.innerHTML = '<div class="usuarios-cargando">Cargando usuarios...</div>';

    try {
      const resultado = await getDocs(collection(db, 'usuarios'));

      usuariosActuales = resultado.docs.map(documento => ({
        ...documento.data(),
        uid: documento.id
      }));

      actualizarEstadisticas();
      dibujarUsuarios();

    } catch (error) {
      console.error(error);
      lista.innerHTML = `
        <div class="usuarios-error">
          No se pudieron cargar los usuarios. Verifica tu conexión y permisos.
        </div>
      `;
    }
  }

  buscador.addEventListener('input', dibujarUsuarios);
  filtroRol.addEventListener('change', dibujarUsuarios);
  filtroEstado.addEventListener('change', dibujarUsuarios);

  lista.addEventListener('click', async evento => {
    const boton = evento.target.closest('button[data-accion]');
    if (!boton) return;

    const uid = boton.dataset.uid;
    const accion = boton.dataset.accion;
    const usuario = usuariosActuales.find(u => u.uid === uid);

    if (!usuario) return;

    if (uid === auth.currentUser?.uid &&
        (accion === 'estado' || accion === 'eliminar')) {
      alert('No puedes desactivar ni eliminar tu propia cuenta.');
      return;
    }

    if (accion === 'estado') {
      const nuevoEstado = usuario.activo !== true;

      if (!confirm(
        `¿Deseas ${nuevoEstado ? 'activar' : 'desactivar'} la cuenta de ${usuario.nombre || usuario.correo}?`
      )) return;

      try {
        await updateDoc(doc(db, 'usuarios', uid), {
          activo: nuevoEstado
        });

        await cargarUsuarios();
      } catch (error) {
        console.error(error);
        alert('No se pudo actualizar el estado del usuario.');
      }
    }

    if (accion === 'editar') {
      const nuevoNombre = prompt('Nombre completo:', usuario.nombre || '');
      if (nuevoNombre === null) return;

      if (!nuevoNombre.trim()) {
        alert('El nombre no puede estar vacío.');
        return;
      }

      const roles = ['administrador', 'director', 'coordinador', 'secretaria'];
      const nuevoRol = prompt(
        'Escribe el nuevo rol: administrador, director, coordinador o secretaria.',
        usuario.rol || 'coordinador'
      );

      if (nuevoRol === null) return;

      const rolLimpio = nuevoRol.trim().toLowerCase();

      if (!roles.includes(rolLimpio)) {
        alert('El rol indicado no es válido.');
        return;
      }

      if (usuario.rol === 'administrador' && rolLimpio !== 'administrador') {
        const otrosAdmins = usuariosActuales.filter(
          u => u.rol === 'administrador' && u.uid !== uid && u.activo === true
        );

        if (otrosAdmins.length === 0) {
          alert('No puedes quitar el rol al único administrador activo.');
          return;
        }
      }

      try {
        await updateDoc(doc(db, 'usuarios', uid), {
          nombre: nuevoNombre.trim(),
          rol: rolLimpio
        });

        alert('Usuario actualizado correctamente.');
        await cargarUsuarios();
      } catch (error) {
        console.error(error);
        alert('No se pudo editar el usuario.');
      }
    }

    if (accion === 'eliminar') {
      if (!confirm(
        `¿Estás seguro de eliminar el registro de ${usuario.nombre || usuario.correo}?\n\nEsta acción borrará su documento de Firestore, pero NO eliminará su cuenta de Firebase Authentication.`
      )) return;

      if (usuario.rol === 'administrador' && usuario.activo === true) {
        const otrosAdmins = usuariosActuales.filter(
          u => u.rol === 'administrador' && u.activo === true && u.uid !== uid
        );

        if (otrosAdmins.length === 0) {
          alert('No puedes eliminar al único administrador activo.');
          return;
        }
      }

      try {
        await deleteDoc(doc(db, 'usuarios', uid));
        alert('Registro eliminado de Firestore. La cuenta de acceso de Firebase Authentication permanece.');
        await cargarUsuarios();
      } catch (error) {
        console.error(error);
        alert('No se pudo eliminar el registro del usuario.');
      }
    }
  });

  formulario.addEventListener('submit', async evento => {
    evento.preventDefault();

    const nombre = document.querySelector('#usuario-nombre').value.trim();
    const correo = document.querySelector('#usuario-correo').value.trim();
    const clave = document.querySelector('#usuario-clave').value;
    const rol = document.querySelector('#usuario-rol').value;

    mensaje.textContent = 'Creando cuenta...';
    mensaje.className = 'usuarios-mensaje';

    try {
      await crearUsuarioSistema(correo, clave, nombre, rol);

      mensaje.textContent = 'Usuario creado correctamente.';
      mensaje.classList.add('mensaje-exito');

      formulario.reset();
      await cargarUsuarios();

    } catch (error) {
      console.error(error);
      mensaje.textContent = 'No se pudo crear el usuario: ' + error.message;
      mensaje.classList.add('mensaje-error');
    }
  });

  await cargarUsuarios();
}


function mostrarPanel(nombre) {
  detenerEscanerQR();

  const rolUsuario = administrador?.rol || 'usuario';

  const nombresRoles = {
    administrador: 'Administrador',
    director: 'Director',
    coordinador: 'Coordinador',
    secretaria: 'Secretaría'
  };

  const nombreRol = nombresRoles[rolUsuario] || rolUsuario;

  app.innerHTML = `
    <div class="panel">

      <!-- MENÚ LATERAL -->
      <aside class="sidebar">

        <div class="marca-institucional">
          <div class="marca-icono">S</div>
          <div class="marca-texto">
            <h2>SHALOM EXPO</h2>
            <p>GESTIÓN 2026</p>
          </div>
        </div>

        <div class="usuario-sidebar">
          <div class="usuario-avatar">
            ${(escapar(nombre || 'U').charAt(0)).toUpperCase()}
          </div>
          <div class="usuario-info">
            <strong>${escapar(nombre)}</strong>
            <span>${escapar(nombreRol)}</span>
          </div>
        </div>

        <div class="menu-titulo">MENÚ PRINCIPAL</div>

        <nav class="menu">

          <button class="activo" data-vista="inicio">
            <span class="menu-icono">⌂</span>
            <span>Inicio</span>
          </button>

          <button data-vista="estudiantes">
            <span class="menu-icono">♙</span>
            <span>Estudiantes</span>
          </button>

          <button data-vista="docentes">
            <span class="menu-icono">♧</span>
            <span>Docentes</span>
          </button>

          <button data-vista="asistencia">
            <span class="menu-icono">▤</span>
            <span>Asistencia</span>
          </button>

          <button data-vista="qr">
            <span class="menu-icono">▣</span>
            <span>Códigos QR</span>
          </button>

          <div class="menu-titulo">EXPO 2026</div>

          <button data-vista="sta">
            <span class="menu-icono">▦</span>
            <span>Proyectos y stands</span>
          </button>

          <button data-vista="visitas">
            <span class="menu-icono">♧</span>
            <span>Visitas</span>
          </button>
<button data-vista="evaluacion">
  <span class="menu-icono">☆</span>
  <span>Evaluación</span>
</button>

<button data-vista="jurados">
  <span class="menu-icono">♙</span>
  <span>Jurados registrados</span>
</button>

<div class="menu-titulo">INFORMACIÓN</div>

          <button data-vista="cie">
            <span class="menu-icono">▥</span>
            <span>CIE – Estadísticas</span>
          </button>

          ${rolUsuario === 'administrador' ? `
            <div class="menu-titulo">CONFIGURACIÓN</div>

            <button data-vista="usuarios">
              <span class="menu-icono">⚙</span>
              <span>Administrar usuarios</span>
            </button>
          ` : ''}

        </nav>

        <div class="sidebar-pie">
          <div class="sidebar-version">
            <span class="estado-punto"></span>
            Sistema operativo
          </div>

          <button id="cerrar-sesion" class="btn-cerrar">
            <span>↪</span>
            Cerrar sesión
          </button>

          <p class="sidebar-creditos">
            COLEGIO ADVENTISTA SHALOM
          </p>
        </div>

      </aside>

      <div class="fondo-menu"></div>

      <!-- CONTENIDO PRINCIPAL -->
      <main id="contenido" class="contenido">

        <header class="encabezado">

          <div class="encabezado-izquierda">

            <button
              id="btn-menu-movil"
              class="btn-menu-movil"
              aria-label="Abrir o cerrar menú"
              aria-expanded="false">
              ☰
            </button>

            <div class="encabezado-titulo">
              <span>COLEGIO ADVENTISTA SHALOM</span>
              <h1>Panel de administración</h1>
            </div>

          </div>

          <div class="encabezado-derecha">
            <div class="indicador-sistema">
              <span></span>
              Sistema activo
            </div>
          </div>

        </header>

        <!-- BIENVENIDA -->
        <section class="dashboard-bienvenida">

          <div class="dashboard-bienvenida-texto">
            <span class="dashboard-etiqueta">
              CENTRO DE CONTROL · SHALOM EXPO 2026
            </span>

            <h2>¡Bienvenido, ${escapar(nombre)}!</h2>

            <p>
              Administra estudiantes, controla la asistencia,
              organiza proyectos y supervisa las evaluaciones
              desde una sola plataforma.
            </p>

            <div class="dashboard-bienvenida-pie">
              <span>GESTIÓN ACADÉMICA 2026</span>
              <span>COLEGIO ADVENTISTA SHALOM</span>
            </div>
          </div>

          <div class="dashboard-bienvenida-grafico">
            <div class="dashboard-circulo">
              <span>EXPO</span>
              <strong>2026</strong>
            </div>
          </div>

        </section>

        <!-- INDICADORES -->
        <section class="dashboard-indicadores">

          <article class="dashboard-indicador">
            <div class="dashboard-indicador-icono morado">♙</div>
            <div>
              <span>ESTUDIANTES</span>
              <strong>—</strong>
              <small>Registro institucional</small>
            </div>
          </article>

          <article class="dashboard-indicador">
            <div class="dashboard-indicador-icono naranja">▦</div>
            <div>
              <span>PROYECTOS</span>
              <strong>—</strong>
              <small>Proyectos registrados</small>
            </div>
          </article>

          <article class="dashboard-indicador">
            <div class="dashboard-indicador-icono verde">♧</div>
            <div>
              <span>JURADOS</span>
              <strong>—</strong>
              <small>Jurados registrados</small>
            </div>
          </article>

          <article class="dashboard-indicador">
            <div class="dashboard-indicador-icono azul">☆</div>
            <div>
              <span>EVALUACIONES</span>
              <strong>—</strong>
              <small>Registros de evaluación</small>
            </div>
          </article>

        </section>

        <!-- ACCESOS RÁPIDOS -->
        <section class="dashboard-seccion">

          <div class="dashboard-seccion-titulo">
            <div>
              <span>GESTIÓN INSTITUCIONAL</span>
              <h2>Accesos rápidos</h2>
              <p>
                Ingresa directamente a las principales herramientas.
              </p>
            </div>
          </div>

          <div class="dashboard-accesos">

            <button class="dashboard-acceso" id="inicio-ir-estudiantes">
              <span class="dashboard-acceso-icono morado">♙</span>
              <span>
                <strong>Estudiantes</strong>
                <small>Registrar y administrar estudiantes</small>
              </span>
              <b>→</b>
            </button>

            <button class="dashboard-acceso" id="inicio-ir-asistencia">
              <span class="dashboard-acceso-icono verde">▤</span>
              <span>
                <strong>Asistencia</strong>
                <small>Consultar registros de asistencia</small>
              </span>
              <b>→</b>
            </button>

            <button class="dashboard-acceso" id="inicio-ir-qr">
              <span class="dashboard-acceso-icono naranja">▣</span>
              <span>
                <strong>Escáner QR</strong>
                <small>Registrar asistencia con credenciales</small>
              </span>
              <b>→</b>
            </button>

            <button class="dashboard-acceso" id="inicio-ir-proyectos">
              <span class="dashboard-acceso-icono azul">▦</span>
              <span>
                <strong>Proyectos y stands</strong>
                <small>Gestionar proyectos de la Expo</small>
              </span>
              <b>→</b>
            </button>

            <button class="dashboard-acceso" id="inicio-ir-visitas">
              <span class="dashboard-acceso-icono naranja">♧</span>
              <span>
                <strong>Visitas</strong>
                <small>Controlar visitantes de la feria</small>
              </span>
              <b>→</b>
            </button>
<button class="dashboard-acceso" id="inicio-ir-jurados">
  <span class="dashboard-acceso-icono verde">♙</span>
  <span>
    <strong>Jurados registrados</strong>
    <small>Consultar todos los jurados inscritos</small>
  </span>
  <b>→</b>
</button>
            <button class="dashboard-acceso" id="inicio-ir-evaluacion">
              <span class="dashboard-acceso-icono morado">☆</span>
              <span>
                <strong>Evaluaciones</strong>
                <small>Gestionar calificaciones de jurados</small>
              </span>
              <b>→</b>
            </button>

            <button class="dashboard-acceso" id="inicio-ir-cie">
              <span class="dashboard-acceso-icono verde">▥</span>
              <span>
                <strong>Estadísticas CIE</strong>
                <small>Consultar información y estadísticas</small>
              </span>
              <b>→</b>
            </button>

            ${rolUsuario === 'administrador' ? `
              <button class="dashboard-acceso" id="inicio-ir-usuarios">
                <span class="dashboard-acceso-icono azul">⚙</span>
                <span>
                  <strong>Usuarios del sistema</strong>
                  <small>Administrar cuentas y permisos</small>
                </span>
                <b>→</b>
              </button>
            ` : ''}

          </div>

        </section>

        <!-- ACTIVIDAD RECIENTE -->
        <section class="dashboard-actividad">

          <div>
            <span>SEGUIMIENTO DEL SISTEMA</span>
            <h2>Actividad reciente</h2>
            <p>
              Consulta los movimientos y registros de la plataforma.
            </p>
          </div>

          <div class="dashboard-actividad-vacia">
            <span>◷</span>
            <strong>Panel de actividad</strong>
            <p>
              Aquí incorporaremos los últimos registros de estudiantes,
              asistencias y evaluaciones.
            </p>
          </div>

        </section>

      </main>
    </div>
  `;

  // ==========================================
  // MENÚ RESPONSIVE
  // ==========================================

  const panel = document.querySelector('.panel');
  const botonMenu = document.querySelector('#btn-menu-movil');
  const fondoMenu = document.querySelector('.fondo-menu');

  function alternarMenu() {
    if (window.innerWidth <= 768) {
      panel.classList.toggle('menu-abierto');
    } else {
      panel.classList.toggle('menu-colapsado');
    }

    botonMenu.setAttribute(
      'aria-expanded',
      panel.classList.contains('menu-abierto') ? 'true' : 'false'
    );
  }

  function cerrarMenu() {
    panel.classList.remove('menu-abierto');
    botonMenu.setAttribute('aria-expanded', 'false');
  }

  botonMenu.addEventListener('click', alternarMenu);
  fondoMenu.addEventListener('click', cerrarMenu);

  // ==========================================
  // CERRAR SESIÓN
  // ==========================================

  document.querySelector('#cerrar-sesion')
    .addEventListener('click', cerrarSesion);

  // ==========================================
  // NAVEGACIÓN
  // ==========================================

  function abrirVista(selector, funcion) {
    const boton = document.querySelector(
      `[data-vista="${selector}"]`
    );

    if (!boton) return;

    boton.addEventListener('click', () => {
      document.querySelectorAll('.menu [data-vista]')
        .forEach(item => item.classList.remove('activo'));

      boton.classList.add('activo');

      if (window.innerWidth <= 768) {
        cerrarMenu();
      }

      funcion();
    });
  }

  abrirVista('inicio', () => {
    mostrarPanel(nombre || 'Usuario');
  });

  abrirVista('estudiantes', mostrarEstudiantes);
  abrirVista('docentes', mostrarDocentes);
  abrirVista('asistencia', mostrarAsistencias);
  abrirVista('qr', mostrarEscanerQR);
  abrirVista('sta', mostrarProyectos);
  abrirVista('visitas', mostrarVisitas);
  abrirVista('jurados', mostrarJuradosRegistrados);
  abrirVista('cie', mostrarCIE);
  abrirVista('evaluacion', mostrarEvaluacion);

  if (rolUsuario === 'administrador') {
    abrirVista('usuarios', mostrarUsuarios);
  }

  // ==========================================
  // ACCESOS RÁPIDOS DEL DASHBOARD
  // ==========================================
  const accesosInicio = {
  'inicio-ir-estudiantes': 'estudiantes',
  'inicio-ir-asistencia': 'asistencia',
  'inicio-ir-qr': 'qr',
  'inicio-ir-proyectos': 'sta',
  'inicio-ir-visitas': 'visitas',
  'inicio-ir-jurados': 'jurados',
  'inicio-ir-evaluacion': 'evaluacion',
  'inicio-ir-cie': 'cie',
  'inicio-ir-usuarios': 'usuarios'
};

  Object.entries(accesosInicio).forEach(([id, vista]) => {
    const acceso = document.getElementById(id);
    const botonMenu = document.querySelector(
      `.menu [data-vista="${vista}"]`
    );

    if (acceso && botonMenu) {
      acceso.addEventListener('click', () => {
        botonMenu.click();
      });
    }
  });
}

// ==========================================
// JURADOS REGISTRADOS
// ==========================================


async function mostrarJuradosRegistrados() {

  const contenido = document.querySelector('#contenido');

  if (!contenido) return;

  // SOLO EL ADMINISTRADOR PUEDE GESTIONAR JURADOS
  if (administrador?.rol !== 'administrador') {
    contenido.innerHTML = `
      <section class="dashboard-seccion jurados-admin">
        <h2>Acceso restringido</h2>
        <p>No tienes permiso para gestionar jurados registrados.</p>
      </section>
    `;
    return;
  }

  contenido.innerHTML = `
    <section class="dashboard-seccion jurados-admin">

      <div class="dashboard-seccion-titulo">
        <div>
          <span>SHALOM EXPO 2026</span>
          <h2>Jurados registrados</h2>
          <p>
            Consulta y administra los jurados inscritos
            y sus evaluaciones realizadas.
          </p>
        </div>
      </div>

      <div class="dashboard-indicadores">

        <article class="dashboard-indicador">
          <div class="dashboard-indicador-icono morado">♙</div>
          <div>
            <span>TOTAL REGISTRADOS</span>
            <strong id="jurados-total">—</strong>
            <small>Jurados inscritos</small>
          </div>
        </article>

        <article class="dashboard-indicador">
          <div class="dashboard-indicador-icono verde">✓</div>
          <div>
            <span>JURADOS ACTIVOS</span>
            <strong id="jurados-activos">—</strong>
            <small>Registros activos</small>
          </div>
        </article>

        <article class="dashboard-indicador">
          <div class="dashboard-indicador-icono morado">☆</div>
          <div>
            <span>TOTAL EVALUACIONES</span>
            <strong id="jurados-evaluaciones">—</strong>
            <small>Evaluaciones registradas</small>
          </div>
        </article>

      </div>

      <div class="dashboard-seccion" style="margin-top:25px;">

        <div style="
          display:flex;
          justify-content:space-between;
          align-items:center;
          flex-wrap:wrap;
          gap:15px;
          margin-bottom:20px;
        ">

          <input
            type="search"
            id="buscar-jurado"
            placeholder="Buscar por nombre o código..."
            style="
              width:100%;
              max-width:420px;
              padding:12px;
              border:1px solid #ddd;
              border-radius:10px;
            "
          >

          <button
            type="button"
            id="jurados-eliminar-todos"
            style="
              background:#b42332;
              color:white;
              border:none;
              border-radius:9px;
              padding:12px 18px;
              font-weight:600;
              cursor:pointer;
            "
          >
            ELIMINAR TODOS LOS JURADOS
          </button>

        </div>

        <div style="overflow-x:auto;">

          <table class="tabla">
            <thead>
              <tr>
                <th>N.º</th>
                <th>Código</th>
                <th>Nombre completo</th>
                <th>Fecha de registro</th>
                <th>Estado</th>
                <th>Evaluaciones</th>
                <th>Acciones</th>
              </tr>
            </thead>

            <tbody id="tabla-jurados-registrados">
              <tr>
                <td colspan="7">Cargando jurados...</td>
              </tr>
            </tbody>
          </table>

        </div>

        <p id="jurados-mensaje" style="margin-top:15px;"></p>

      </div>

      <button
        id="jurados-volver"
        class="btn-secundario"
        style="margin-top:20px;"
      >
        ← Volver al inicio
      </button>

    </section>
  `;

  const tabla = document.querySelector('#tabla-jurados-registrados');
  const mensaje = document.querySelector('#jurados-mensaje');
  const buscar = document.querySelector('#buscar-jurado');
  const botonEliminarTodos = document.querySelector(
    '#jurados-eliminar-todos'
  );

  document.querySelector('#jurados-volver')
    ?.addEventListener('click', () => {
      document.querySelector('[data-vista="inicio"]')?.click();
    });

  let jurados = [];
  let evaluaciones = [];

  // FORMATEAR FECHA
  function formatearFecha(valor) {

    if (!valor) return 'No disponible';

    let fechaObjeto;

    if (typeof valor.toDate === 'function') {
      fechaObjeto = valor.toDate();
    } else {
      fechaObjeto = new Date(valor);
    }

    if (Number.isNaN(fechaObjeto.getTime())) {
      return 'No disponible';
    }

    return fechaObjeto.toLocaleDateString('es-BO');
  }

  // CARGAR JURADOS Y EVALUACIONES
  async function cargarDatosJurados() {

    const [respuestaJurados, respuestaEvaluaciones] =
      await Promise.all([
        getDocs(collection(db, 'jurados')),
        getDocs(collection(db, 'evaluaciones_feria'))
      ]);

    jurados = respuestaJurados.docs.map(documento => ({
      id: documento.id,
      ...documento.data()
    }));

    evaluaciones = respuestaEvaluaciones.docs.map(documento => ({
      id: documento.id,
      ...documento.data()
    }));

    // ORDENAR POR FECHA DE REGISTRO
    jurados.sort((a, b) => {
      const fechaA = new Date(a.creadoEn || 0).getTime() || 0;
      const fechaB = new Date(b.creadoEn || 0).getTime() || 0;

      return fechaB - fechaA;
    });

    // CONTADORES
    const activos = jurados.filter(
      jurado => jurado.activo === true
    ).length;

    document.querySelector('#jurados-total').textContent =
      jurados.length;

    document.querySelector('#jurados-activos').textContent =
      activos;

    document.querySelector('#jurados-evaluaciones').textContent =
      evaluaciones.length;

    dibujarJurados();

    mensaje.textContent =
      `Se encontraron ${jurados.length} jurados y ${evaluaciones.length} evaluaciones registradas.`;
  }

  // DIBUJAR TABLA
  function dibujarJurados(filtro = '') {

    const termino = filtro.trim().toLowerCase();

    const filtrados = jurados.filter(jurado => {

      const nombre = jurado.nombreCompleto ||
        `${jurado.nombres || ''} ${jurado.apellidos || ''}`;

      const codigo = jurado.codigo || '';

      return (
        nombre.toLowerCase().includes(termino) ||
        codigo.toLowerCase().includes(termino)
      );
    });

    if (filtrados.length === 0) {

      tabla.innerHTML = `
        <tr>
          <td colspan="7" style="text-align:center; padding:25px;">
            No se encontraron jurados registrados.
          </td>
        </tr>
      `;

      return;
    }

    tabla.innerHTML = filtrados.map((jurado, indice) => {

      const nombre = jurado.nombreCompleto ||
        `${jurado.nombres || ''} ${jurado.apellidos || ''}`;

      const cantidadEvaluaciones = evaluaciones.filter(
        evaluacion => evaluacion.jurado_uid === jurado.id
      ).length;

      const estado = jurado.activo === true
        ? '<span style="color:#16834a;font-weight:600;">Activo</span>'
        : '<span style="color:#b42332;font-weight:600;">Inactivo</span>';

      return `
        <tr>
          <td>${indice + 1}</td>

          <td>
            ${escapar(jurado.codigo || '—')}
          </td>

          <td>
            ${escapar(nombre.trim() || 'Sin nombre')}
          </td>

          <td>
            ${escapar(formatearFecha(jurado.creadoEn))}
          </td>

          <td>${estado}</td>

          <td>
            <strong>${cantidadEvaluaciones}</strong>
          </td>

          <td>
            <button
              type="button"
              data-jurado-eliminar="${escapar(jurado.id)}"
              style="
                background:#b42332;
                color:white;
                border:none;
                border-radius:7px;
                padding:8px 12px;
                cursor:pointer;
              "
            >
              Eliminar
            </button>
          </td>
        </tr>
      `;

    }).join('');
  }

  // BUSCADOR
  buscar.addEventListener('input', () => {
    dibujarJurados(buscar.value);
  });

  // ELIMINAR UN JURADO
  tabla.addEventListener('click', async evento => {

    const boton = evento.target.closest('[data-jurado-eliminar]');

    if (!boton) return;

    if (administrador?.rol !== 'administrador') {
      alert('No tienes permiso para eliminar jurados.');
      return;
    }

    const id = boton.dataset.juradoEliminar;

    const jurado = jurados.find(item => item.id === id);

    if (!jurado) return;

    const nombre = jurado.nombreCompleto ||
      `${jurado.nombres || ''} ${jurado.apellidos || ''}`;

    const confirmar = confirm(
      `¿Deseas eliminar el registro del jurado ${nombre}?\n\n` +
      'También se eliminará su código reservado.\n\n' +
      'IMPORTANTE: Sus evaluaciones históricas se conservarán.'
    );

    if (!confirmar) return;

    boton.disabled = true;
    boton.textContent = 'Eliminando...';

    try {

      const lote = writeBatch(db);

      lote.delete(doc(db, 'jurados', id));

      if (jurado.codigo) {
        lote.delete(
          doc(db, 'codigos_jurados', String(jurado.codigo))
        );
      }

      await lote.commit();

      alert(
        'Jurado eliminado correctamente. Sus evaluaciones se conservaron.'
      );

      await cargarDatosJurados();

    } catch (error) {

      console.error('Error al eliminar jurado:', error);

      alert(
        'No se pudo eliminar el jurado. Verifica los permisos de Firebase.'
      );

      boton.disabled = false;
      boton.textContent = 'Eliminar';
    }
  });

  // ELIMINAR TODOS LOS JURADOS
  botonEliminarTodos.addEventListener('click', async () => {

    if (administrador?.rol !== 'administrador') {
      alert('No tienes permiso para eliminar jurados.');
      return;
    }

    if (jurados.length === 0) {
      alert('No hay jurados registrados para eliminar.');
      return;
    }

    const primeraConfirmacion = confirm(
      `ATENCIÓN\n\n` +
      `Se eliminarán ${jurados.length} perfiles de jurados ` +
      'y sus códigos reservados.\n\n' +
      'Las evaluaciones históricas NO serán eliminadas.\n\n' +
      '¿Deseas continuar?'
    );

    if (!primeraConfirmacion) return;

    const segundaConfirmacion = prompt(
      'Para confirmar la eliminación de todos los jurados, escribe exactamente:\n\nELIMINAR TODOS'
    );

    if (segundaConfirmacion !== 'ELIMINAR TODOS') {
      alert('Operación cancelada. No se eliminó ningún jurado.');
      return;
    }

    botonEliminarTodos.disabled = true;
    botonEliminarTodos.textContent = 'ELIMINANDO...';

    try {

      // CREAR LISTA DE DOCUMENTOS A ELIMINAR
      const referencias = [];

      jurados.forEach(jurado => {

        referencias.push(
          doc(db, 'jurados', jurado.id)
        );

        if (jurado.codigo) {
          referencias.push(
            doc(db, 'codigos_jurados', String(jurado.codigo))
          );
        }

      });

      // FIRESTORE ADMITE HASTA 500 OPERACIONES POR LOTE.
      // USAMOS 400 PARA MANTENER UN MARGEN DE SEGURIDAD.
      for (let inicio = 0; inicio < referencias.length; inicio += 400) {

        const lote = writeBatch(db);

        const grupo = referencias.slice(inicio, inicio + 400);

        grupo.forEach(referencia => {
          lote.delete(referencia);
        });

        await lote.commit();
      }

      alert(
        'Se eliminaron todos los perfiles de jurados y sus códigos. ' +
        'Las evaluaciones históricas permanecen intactas.'
      );

      await cargarDatosJurados();

    } catch (error) {

      console.error('Error al eliminar todos los jurados:', error);

      alert(
        'No se pudo completar la eliminación. Revisa los permisos de Firebase.'
      );

    } finally {

      botonEliminarTodos.disabled = false;
      botonEliminarTodos.textContent = 'ELIMINAR TODOS LOS JURADOS';

    }
  });

  // INICIAR CARGA
  try {

    await cargarDatosJurados();

  } catch (error) {

    console.error('Error al cargar jurados:', error);

    tabla.innerHTML = `
      <tr>
        <td colspan="7" style="text-align:center; padding:25px;">
          No se pudieron cargar los registros.
        </td>
      </tr>
    `;

    mensaje.textContent =
      'Error: ' + (error.message || 'No se pudieron consultar los jurados.');

  }

}

let cieCache = null;

async function mostrarCIE() {
  await detenerEscanerQR();

  const contenido = document.querySelector('#contenido');

  contenido.innerHTML = `
    <header class="encabezado">
    <button id="volver-inicio" class="asistencia-btn asistencia-btn-principal">
  ← INICIO
</button>
      <div>
        <p>COLEGIO ADVENTISTA SHALOM</p>
        <h1>Centro de Inteligencia y Estadística</h1>
      </div>
      <span class="estado">● SISTEMA ACTIVO</span>
    </header>

    <section class="cie-portada">
      <div>
        <span>EXPO SHALOM 2026</span>
        <h2>Panel de control institucional</h2>
        <p>Estadísticas, indicadores y seguimiento en tiempo real.</p>
      </div>
      <div class="cie-portada-icono">📊</div>
    </section>

    <section class="cie-filtros">
      <div>
        <label for="cie-periodo">PERÍODO ESTADÍSTICO</label>
        <select id="cie-periodo">
          <option value="hoy">Hoy</option>
          <option value="mes" selected>Este mes</option>
          <option value="todo">Todo el historial</option>
          <option value="rango">Personalizado</option>
        </select>
      </div>

      <div id="cie-rango-fechas" style="display:none;">
        <label for="cie-desde">Desde</label>
        <input type="date" id="cie-desde">
        <label for="cie-hasta">Hasta</label>
        <input type="date" id="cie-hasta">
      </div>

      <button id="cie-aplicar">APLICAR FILTRO</button>
      <button id="cie-actualizar">↻ ACTUALIZAR</button>
    </section>

    <div id="cie-mensaje" class="cie-mensaje">
      Consultando información de Firebase...
    </div>

    <section class="cie-subtitulo">
      <div>
        <h2>Indicadores institucionales</h2>
        <p>Resumen general del sistema</p>
      </div>
    </section>

    <section class="cie-indicadores">
      <article class="cie-indicador violeta">
        <span>Estudiantes registrados</span>
        <strong id="cie-total">0</strong>
        <small>Base institucional</small>
      </article>

      <article class="cie-indicador azul">
        <span>Estudiantes activos</span>
        <strong id="cie-activos">0</strong>
        <small>Estudiantes habilitados</small>
      </article>

      <article class="cie-indicador naranja">
        <span>QR generados</span>
        <strong id="cie-qr">0</strong>
        <small>Credenciales con código</small>
      </article>

      <article class="cie-indicador turquesa">
        <span>Escaneos / asistencias</span>
        <strong id="cie-escaneos">0</strong>
        <small>Registros del período</small>
      </article>

      <article class="cie-indicador verde">
        <span>Visitas a stands</span>
        <strong id="cie-visitas-stands">0</strong>
        <small>Registros del período</small>
      </article>

      <article class="cie-indicador rosa">
        <span>Visitas a estudiantes</span>
        <strong id="cie-visitas-estudiantes">0</strong>
        <small>Seguimiento estudiantil</small>
      </article>

      <article class="cie-indicador amarillo">
        <span>Evaluaciones</span>
        <strong id="cie-evaluaciones">0</strong>
        <small>Evaluaciones registradas</small>
      </article>

      <article class="cie-indicador oscuro">
        <span>Proyectos</span>
        <strong id="cie-proyectos">0</strong>
        <small>Proyectos registrados</small>
      </article>

      <article class="cie-indicador azul">
        <span>Usuarios registrados</span>
        <strong id="cie-usuarios">0</strong>
        <small>Cuentas del sistema</small>
      </article>

      <article class="cie-indicador verde">
        <span>Usuarios activos</span>
        <strong id="cie-usuarios-activos">0</strong>
        <small>Cuentas habilitadas</small>
      </article>
    </section>

    <section class="cie-subtitulo">
      <div>
        <h2>Estadísticas y gráficos</h2>
        <p id="cie-periodo-texto">Resultados del mes seleccionado</p>
      </div>
    </section>

    <section class="cie-graficos">
      <article class="cie-grafico">
        <h3>Actividad general</h3>
        <p>Comparación de registros por módulo</p>
        <div id="cie-grafico-actividad"></div>
      </article>

      <article class="cie-grafico">
        <h3>Estudiantes por nivel</h3>
        <p>Distribución de estudiantes registrados</p>
        <div id="cie-grafico-niveles"></div>
      </article>

      <article class="cie-grafico">
        <h3>Asistencia por curso</h3>
        <p>Registros de asistencia del período</p>
        <div id="cie-grafico-cursos"></div>
      </article>

      <article class="cie-grafico">
        <h3>Resultados de evaluación</h3>
        <p>Promedio de las evaluaciones registradas</p>
        <div id="cie-grafico-evaluaciones"></div>
      </article>
    </section>

    <section class="cie-pie">
      <span id="cie-fecha">Esperando actualización...</span>
      <span>SHALOM EXPO 2026 · CIE</span>
    </section>
  `;
document.querySelector('#volver-inicio')
  .addEventListener('click', () => {
    mostrarPanel(administrador?.nombre || 'Usuario');
  });
  const hoy = new Date();
  const fechaHoy = [
    hoy.getFullYear(),
    String(hoy.getMonth() + 1).padStart(2, '0'),
    String(hoy.getDate()).padStart(2, '0')
  ].join('-');

  document.querySelector('#cie-desde').value =
    fechaHoy.slice(0, 8) + '01';

  document.querySelector('#cie-hasta').value = fechaHoy;

  document.querySelector('#cie-periodo').addEventListener('change', e => {
    document.querySelector('#cie-rango-fechas').style.display =
      e.target.value === 'rango' ? 'flex' : 'none';
  });

  document.querySelector('#cie-aplicar')
    .addEventListener('click', aplicarFiltroCIE);

  document.querySelector('#cie-actualizar')
    .addEventListener('click', async () => {
      cieCache = null;
      await cargarDatosCIE();
      aplicarFiltroCIE();
    });

  await cargarDatosCIE();
  aplicarFiltroCIE();
}

async function cargarDatosCIE() {
  const mensaje = document.querySelector('#cie-mensaje');
  mensaje.textContent = 'Consultando las bases de datos...';

  const consultar = async nombre => {
    try {
      const resultado = await getDocs(collection(db, nombre));
      return resultado.docs.map(d => ({
        id: d.id,
        ...d.data()
      }));
    } catch (error) {
      console.warn('No se pudo consultar ' + nombre, error);
      return null;
    }
  };

  try {
    const resultados = await Promise.all([
      consultar('estudiantes'),
      consultar('asistencias'),
      consultar('visitas_stands'),
      consultar('visitas_estudiantes'),
      consultar('evaluaciones_feria'),
      consultar('proyectos'),
      consultar('usuarios')
    ]);

    cieCache = {
      estudiantes: resultados[0] || [],
      asistencias: resultados[1] || [],
      visitasStands: resultados[2] || [],
      visitasEstudiantes: resultados[3] || [],
      evaluaciones: resultados[4] || [],
      proyectos: resultados[5] || [],
      usuarios: resultados[6] || []
    };

    mensaje.textContent =
      'Información consultada. Puedes seleccionar el período que deseas analizar.';

  } catch (error) {
    console.error('Error CIE:', error);
    mensaje.textContent = 'No se pudieron consultar los datos.';
  }
}

function fechaRegistroCIE(registro) {
  const valor =
    registro.fecha ||
    registro.fechaEvaluacion ||
    registro.fechaRegistro ||
    registro.creadoEn ||
    registro.createdAt ||
    registro.timestamp;

  if (!valor) return '';

  if (typeof valor === 'string' && /^\d{4}-\d{2}-\d{2}/.test(valor)) {
    return valor.slice(0, 10);
  }

  let fecha;

  if (valor && typeof valor.toDate === 'function') {
    fecha = valor.toDate();
  } else {
    fecha = new Date(valor);
  }

  if (isNaN(fecha.getTime())) return '';

  return [
    fecha.getFullYear(),
    String(fecha.getMonth() + 1).padStart(2, '0'),
    String(fecha.getDate()).padStart(2, '0')
  ].join('-');
}

function aplicarFiltroCIE() {
  if (!cieCache) return;

  const periodo = document.querySelector('#cie-periodo').value;
  const hoy = new Date();

  const hoyTexto = [
    hoy.getFullYear(),
    String(hoy.getMonth() + 1).padStart(2, '0'),
    String(hoy.getDate()).padStart(2, '0')
  ].join('-');

  let desde = '';
  let hasta = '';
  let descripcion = '';

  if (periodo === 'hoy') {
    desde = hoyTexto;
    hasta = hoyTexto;
    descripcion = 'Estadísticas de hoy';

  } else if (periodo === 'mes') {
    desde = hoyTexto.slice(0, 8) + '01';
    hasta = hoyTexto;
    descripcion = 'Estadísticas de este mes';

  } else if (periodo === 'rango') {
    desde = document.querySelector('#cie-desde').value;
    hasta = document.querySelector('#cie-hasta').value;

    if (!desde || !hasta || desde > hasta) {
      alert('Selecciona un rango de fechas válido.');
      return;
    }

    descripcion = `Del ${desde} al ${hasta}`;

  } else {
    descripcion = 'Historial completo';
  }

  const filtrar = lista => lista.filter(registro => {
    if (!desde && !hasta) return true;

    const fecha = fechaRegistroCIE(registro);
    return fecha && fecha >= desde && fecha <= hasta;
  });

  const estudiantes = cieCache.estudiantes;
  const activos = estudiantes.filter(e => e.activo !== false);
  const asistencias = filtrar(cieCache.asistencias);
  const visitasStands = filtrar(cieCache.visitasStands);
  const visitasEstudiantes = filtrar(cieCache.visitasEstudiantes);
  const evaluaciones = filtrar(cieCache.evaluaciones);
  const proyectos = filtrar(cieCache.proyectos);

  const usuarios = cieCache.usuarios;
  const qr = estudiantes.filter(e => e.qr_id && String(e.qr_id).trim());

  const presentesHoy = asistencias.filter(a => {
    const estado = String(a.estado || '').toLowerCase();
    return ['presente', 'a tiempo', 'atrasado', 'tarde'].includes(estado);
  }).length;

  const ausentesHoy = periodo === 'hoy'
    ? Math.max(0, activos.length - presentesHoy)
    : 0;

  const porcentaje = activos.length
    ? Math.round((presentesHoy / activos.length) * 100)
    : 0;

  const poner = (id, valor) => {
    const elemento = document.querySelector('#' + id);
    if (elemento) elemento.textContent = valor;
  };

  poner('cie-total', estudiantes.length);
  poner('cie-activos', activos.length);
  poner('cie-qr', qr.length);
  poner('cie-escaneos', asistencias.length);
  poner('cie-visitas-stands', visitasStands.length);
  poner('cie-visitas-estudiantes', visitasEstudiantes.length);
  poner('cie-evaluaciones', evaluaciones.length);
  poner('cie-proyectos', proyectos.length);
  poner('cie-usuarios', usuarios.length);
  poner('cie-usuarios-activos', usuarios.filter(u => u.activo === true).length);

  poner('cie-periodo-texto', descripcion);

  if (periodo === 'hoy') {
    document.querySelector('#cie-periodo-texto').textContent =
      `${descripcion} · ${presentesHoy} presentes · ${ausentesHoy} ausentes · ${porcentaje}% de asistencia`;
  }

  const contar = lista => {
    const grupos = {};

    lista.forEach(item => {
      const nombre = item.curso || item.nivel || 'Sin clasificar';
      grupos[nombre] = (grupos[nombre] || 0) + 1;
    });

    return Object.entries(grupos)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10);
  };

  const escapar = texto => String(texto).replace(/[&<>"']/g, letra => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  })[letra]);

  const dibujar = (id, datos, color) => {
    const contenedor = document.querySelector('#' + id);
    if (!contenedor) return;

    const maximo = Math.max(1, ...datos.map(d => d[1]));

    if (!datos.length) {
      contenedor.innerHTML = '<p class="cie-sin-datos">No hay datos para mostrar.</p>';
      return;
    }

    contenedor.innerHTML = datos.map(d => `
      <div class="cie-barra-item">
        <div class="cie-barra-texto">
          <span>${escapar(d[0])}</span>
          <strong>${d[1]}</strong>
        </div>
        <div class="cie-barra-fondo">
          <div class="cie-barra-progreso"
            style="width:${Math.max(2, d[1] / maximo * 100)}%;background:${color}">
          </div>
        </div>
      </div>
    `).join('');
  };

  dibujar('cie-grafico-actividad', [
    ['Asistencias', asistencias.length],
    ['Visitas a stands', visitasStands.length],
    ['Visitas a estudiantes', visitasEstudiantes.length],
    ['Evaluaciones', evaluaciones.length],
    ['Proyectos', proyectos.length]
  ], '#7853a6');

  const niveles = {};
  estudiantes.forEach(e => {
    const nivel = e.nivel || 'Sin nivel';
    niveles[nivel] = (niveles[nivel] || 0) + 1;
  });

  dibujar(
    'cie-grafico-niveles',
    Object.entries(niveles).sort((a, b) => b[1] - a[1]),
    '#e99a39'
  );

  dibujar(
    'cie-grafico-cursos',
    contar(asistencias),
    '#299b91'
  );

  const promedio = evaluaciones.length
    ? evaluaciones.reduce((suma, e) => suma + Number(e.total || 0), 0) / evaluaciones.length
    : 0;

  dibujar('cie-grafico-evaluaciones', [
    ['Promedio obtenido', Math.round(promedio * 10) / 10],
    ['Máximo posible', 45]
  ], '#d56a91');

  document.querySelector('#cie-fecha').textContent =
    'Actualizado: ' + new Date().toLocaleString('es-BO');
}

// ==========================================
// MÓDULO VISITAS – SHALOM EXPO 2026
// ==========================================


async function mostrarVisitas() {
  await detenerEscanerQR();

  const contenido = document.querySelector('#contenido');

  contenido.innerHTML = `
    <header class="encabezado">
    <header class="encabezado">
  <div>
    <p>COLEGIO ADVENTISTA SHALOM</p>
    <h1>Gestión de visitas</h1>
  </div>

  <div style="display:flex; align-items:center; gap:12px;">
    <span class="estado">● SISTEMA ACTIVO</span>
    <button type="button" id="volver-inicio-visitas" class="asistencia-btn asistencia-btn-principal">
      ← INICIO
    </button>
  </div>
</header>
      <div>
        <p>COLEGIO ADVENTISTA SHALOM</p>
        <h1>Gestión de visitas</h1>
      </div>
      <span class="estado">● SISTEMA ACTIVO</span>
    </header>

    <section class="cie-portada">
      <div>
        <span>SHALOM EXPO 2026</span>
        <h2>VISITAS – SHALOM</h2>
        <p>Registro de visitantes y seguimiento de visitas a estudiantes.</p>
      </div>
      <div class="cie-portada-icono">🤝</div>
    </section>

    <section class="cie-indicadores">
      <article class="cie-indicador violeta">
        <span>Visitantes del estand</span>
        <strong id="total-visitantes">0</strong>
        <small>Personas registradas</small>
      </article>

      <article class="cie-indicador azul">
        <span>Visitas programadas</span>
        <strong id="total-programadas">0</strong>
        <small>Agenda estudiantil</small>
      </article>

      <article class="cie-indicador naranja">
        <span>Visitas pendientes</span>
        <strong id="total-pendientes">0</strong>
        <small>Seguimientos por realizar</small>
      </article>

      <article class="cie-indicador verde">
        <span>Visitas realizadas</span>
        <strong id="total-realizadas">0</strong>
        <small>Seguimientos completados</small>
      </article>
    </section>

    <section class="cie-graficos" style="margin-top:22px;">

      <article class="cie-grafico">
        <h3>REGISTRO DE VISITANTES DEL ESTAND</h3>
        <p>Personas que visitan nuestro estand en la Feria Multidisciplinaria.</p>

        <form id="form-visitante" class="formulario-visitas">

          <label>Nombre completo</label>
          <input id="visitante-nombre" required
            placeholder="Nombre y apellidos">

          <label>Institución o procedencia</label>
          <input id="visitante-institucion"
            placeholder="Colegio, institución o procedencia">

          <label>Teléfono</label>
          <input id="visitante-telefono"
            placeholder="Número de contacto">

          <label>Área de interés / Materia de exposición</label>
          <select id="visitante-interes" required>
            <option value="">Seleccionar área de exposición</option>
            <option>Artes Plásticas</option>
            <option>Biología</option>
            <option>Ciencias Sociales</option>
            <option>Educación Física</option>
            <option>Filosofía</option>
            <option>Física</option>
            <option>Formación Cristiana</option>
            <option>Inglés</option>
            <option>Lenguaje</option>
            <option>Matemática</option>
            <option>Orientación Vocacional</option>
            <option>Psicología</option>
            <option>Química</option>
            <option>Robótica</option>
            <option value="Otro">Otro – Especificar</option>
          </select>

          <div id="otro-area-contenedor" style="display:none;">
            <label>Especifique otra área</label>
            <input id="visitante-otro-interes"
              placeholder="Nombre de la exposición">
          </div>

          <label>Estudiante que atendió (opcional)</label>
          <input id="visitante-atendido"
            placeholder="Nombre del estudiante">

          <button type="submit">REGISTRAR VISITANTE</button>
        </form>
      </article>

      <article class="cie-grafico">
        <h3>AGENDA DE VISITAS A ESTUDIANTES</h3>
        <p>Programa y controla las visitas a los estudiantes.</p>

        <form id="form-visita-estudiante" class="formulario-visitas">

          <label>Estudiante</label>
          <select id="visita-estudiante" required>
            <option value="">Cargando estudiantes...</option>
          </select>

          <label>Fecha de visita</label>
          <input type="date" id="visita-fecha" required>

          <label>Hora</label>
          <input type="time" id="visita-hora" required>

          <label>Persona que realizará la visita</label>
          <input id="visita-responsable" required
            placeholder="Nombre del responsable">

          <label>Motivo de la visita</label>
          <textarea id="visita-motivo" required
            placeholder="Describe el motivo de la visita"></textarea>

          <label>Recordatorio</label>
          <select id="visita-recordatorio">
            <option value="Sin recordatorio">Sin recordatorio</option>
            <option value="El mismo día">El mismo día</option>
            <option value="Un día antes">Un día antes</option>
          </select>

          <button type="submit">PROGRAMAR VISITA</button>
        </form>
      </article>
    </section>

    <section class="cie-grafico" style="margin-top:22px;">
      <h3>VISITANTES REGISTRADOS</h3>
      <button type="button" id="limpiar-visitantes"
  class="eval-btn-secundario" style="margin:10px 0;">
  Limpiar todo
</button>
      <p>Historial de personas que visitaron nuestro estand.</p>

      <div class="tabla-responsive">
        <table>
          <thead>
            <tr>
              <th>N.º</th>
              <th>VISITANTE</th>
              <th>PROCEDENCIA</th>
              <th>INTERÉS</th>
              <th>FECHA</th>
            </tr>
          </thead>
          <tbody id="tabla-visitantes">
            <tr><td colspan="5">Consultando registros...</td></tr>
          </tbody>
        </table>
      </div>
    </section>

    <section class="cie-grafico" style="margin-top:22px;">
      <h3>VISITAS PROGRAMADAS</h3>
      <button type="button" id="limpiar-visitas"
  class="eval-btn-secundario" style="margin:10px 0;">
  Limpiar todo
</button>
      <p>Control y seguimiento de las visitas a estudiantes.</p>

      <div class="tabla-responsive">
        <table>
          <thead>
            <tr>
              <th>ESTUDIANTE</th>
              <th>CURSO</th>
              <th>FECHA Y HORA</th>
              <th>RESPONSABLE</th>
              <th>ESTADO</th>
              <th>ACCIÓN</th>
            </tr>
          </thead>
          <tbody id="tabla-visitas">
            <tr><td colspan="6">Consultando agenda...</td></tr>
          </tbody>
        </table>
      </div>
    </section>
    <section class="cie-grafico" style="margin-top:22px;">
      <h3>VISITAS A PROYECTOS Y STANDS</h3>
      <button type="button" id="limpiar-visitas-proyectos"
  class="eval-btn-secundario" style="margin:10px 0;">
  Limpiar todo
</button>
      <p>Registro de las personas que STAFF ha atendido al visitar los proyectos.</p>

      <div class="tabla-responsive">
        <table>
          <thead>
            <tr>
              <th>N.º</th>
              <th>PROYECTO</th>
              <th>ÁREA</th>
              <th>STAND</th>
              <th>FECHA</th>
              <th>HORA</th>
              <th>STAFF RESPONSABLE</th>
              <th>ACCIÓN</th>
            </tr>
          </thead>
          <tbody id="tabla-visitas-proyectos">
            <tr>
              <td colspan="7">Consultando visitas a proyectos...</td>
            </tr>
          </tbody>
        </table>
      </div>
    </section>
    <div id="visitas-mensaje"
      class="cie-mensaje"
      style="margin-top:20px;">
      Preparando módulo de visitas...
    </div>
  `;
document.querySelector('#volver-inicio-visitas')
  .addEventListener('click', async () => {
    await detenerEscanerQR();
    mostrarPanel(administrador?.nombre || 'Administrador');
  });
  document.querySelector('#visita-fecha').value =
    new Date().toLocaleDateString('en-CA');

  const selectorInteres =
    document.querySelector('#visitante-interes');

  const otroContenedor =
    document.querySelector('#otro-area-contenedor');

  const otroInteres =
    document.querySelector('#visitante-otro-interes');

  selectorInteres.addEventListener('change', () => {
    const mostrar = selectorInteres.value === 'Otro';
    otroContenedor.style.display = mostrar ? 'block' : 'none';
    otroInteres.required = mostrar;

    if (!mostrar) otroInteres.value = '';
  });

  document.querySelector('#form-visitante')
    .addEventListener('submit', guardarVisitanteShalom);

  document.querySelector('#form-visita-estudiante')
    .addEventListener('submit', guardarVisitaEstudianteShalom);

  document.querySelector('#tabla-visitas')
    .addEventListener('click', gestionarAccionVisitaShalom);
    document.querySelector('#tabla-visitas-proyectos')
  .addEventListener('click', async evento => {
    const boton = evento.target.closest(
      '[data-eliminar-visita-proyecto]'
    );

    if (!boton) return;

    if (administrador?.rol !== 'administrador') return;

    const id = boton.dataset.eliminarVisitaProyecto;

    const confirmar = confirm(
      '¿Deseas eliminar permanentemente esta visita al proyecto?'
    );

    if (!confirmar) return;

    boton.disabled = true;
    boton.textContent = 'Eliminando...';

    try {
      await deleteDoc(doc(db, 'visitas_proyectos', id));

      await cargarDatosVisitasShalom();

      alert('La visita se eliminó correctamente.');

    } catch (error) {
      console.error('Error eliminando visita:', error);

      alert('No se pudo eliminar la visita. Revisa los permisos de Firebase.');

      boton.disabled = false;
      boton.textContent = 'Eliminar';
    }
  });
    if (String(administrador?.rol || '').toLowerCase() === 'administrador') {

document.querySelector('#limpiar-visitantes')
  .addEventListener('click', () =>
    limpiarColeccionVisitasShalom(
      'visitantes_feria',
      'limpiar-visitantes',
      'Visitantes registrados'
    )
  );

document.querySelector('#limpiar-visitas')
  .addEventListener('click', () =>
    limpiarColeccionVisitasShalom(
      'visitas_estudiantes',
      'limpiar-visitas',
      'Visitas programadas'
    )
  );

document.querySelector('#limpiar-visitas-proyectos')
  .addEventListener('click', () =>
    limpiarColeccionVisitasShalom(
      'visitas_proyectos',
      'limpiar-visitas-proyectos',
      'Visitas a proyectos y stands'
    )
  );
}

  await cargarEstudiantesVisitasShalom();
  await cargarDatosVisitasShalom();
}


// ==========================================
// FUNCIONES AUXILIARES
// ==========================================

function escaparVisitasShalom(valor) {
  return String(valor ?? '').replace(/[&<>"']/g, letra => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  })[letra]);
}

function nombreEstudianteVisitasShalom(e) {
  return (
    e.nombreCompleto ||
    [e.nombres, e.apellidos].filter(Boolean).join(' ') ||
    [e.nombre, e.apellido].filter(Boolean).join(' ') ||
    'Estudiante sin nombre'
  );
}

function cursoEstudianteVisitasShalom(e) {
  return [e.nivel, e.curso, e.paralelo]
    .filter(Boolean)
    .join(' ');
}


// ==========================================
// CARGAR ESTUDIANTES
// ==========================================

async function cargarEstudiantesVisitasShalom() {
  const selector = document.querySelector('#visita-estudiante');

  try {
    const respuesta = await getDocs(collection(db, 'estudiantes'));

    const estudiantes = respuesta.docs
      .map(d => ({
        id: d.id,
        ...d.data()
      }))
      .filter(e => e.activo !== false)
      .sort((a, b) =>
        nombreEstudianteVisitasShalom(a).localeCompare(
          nombreEstudianteVisitasShalom(b)
        )
      );

    selector.innerHTML =
      '<option value="">Seleccionar estudiante</option>';

    estudiantes.forEach(e => {
      const opcion = document.createElement('option');
      opcion.value = e.id;
      opcion.textContent =
        `${nombreEstudianteVisitasShalom(e)} — ${cursoEstudianteVisitasShalom(e)}`;
      selector.appendChild(opcion);
    });

    if (!estudiantes.length) {
      selector.innerHTML =
        '<option value="">No hay estudiantes registrados</option>';
    }

  } catch (error) {
    console.error('Error al cargar estudiantes de visitas:', error);
    selector.innerHTML =
      '<option value="">Error al cargar estudiantes</option>';
  }
}


// ==========================================
// REGISTRAR VISITANTE DEL ESTAND
// ==========================================

async function guardarVisitanteShalom(evento) {
  evento.preventDefault();

  const boton = evento.submitter;
  if (boton) boton.disabled = true;

  const mensaje = document.querySelector('#visitas-mensaje');

  try {
    const interesSeleccionado =
      document.querySelector('#visitante-interes').value;

    const otroInteres =
      document.querySelector('#visitante-otro-interes').value.trim();

    const interes = interesSeleccionado === 'Otro'
      ? otroInteres
      : interesSeleccionado;

    if (!interes) {
      throw new Error('Debes especificar el área de interés.');
    }

    const datos = {
      nombre: document.querySelector('#visitante-nombre').value.trim(),
      institucion: document.querySelector('#visitante-institucion').value.trim(),
      telefono: document.querySelector('#visitante-telefono').value.trim(),
      interes: interes,
      estudianteAtendio: document.querySelector('#visitante-atendido').value.trim(),
      fecha: new Date().toLocaleDateString('en-CA'),
      creadoEn: new Date().toISOString()
    };

    await addDoc(collection(db, 'visitantes_feria'), datos);

    evento.target.reset();
    document.querySelector('#otro-area-contenedor').style.display = 'none';

    mensaje.textContent = 'Visitante registrado correctamente.';

    await cargarDatosVisitasShalom();

  } catch (error) {
    console.error('Error al registrar visitante:', error);
    mensaje.textContent = 'ERROR: ' + error.code + ' - ' + error.message;
  } finally {
    if (boton) boton.disabled = false;
  }
}


// ==========================================
// PROGRAMAR VISITA A ESTUDIANTE
// ==========================================

async function guardarVisitaEstudianteShalom(evento) {
  evento.preventDefault();

  const boton = evento.submitter;
  if (boton) boton.disabled = true;

  const mensaje = document.querySelector('#visitas-mensaje');

  try {
    const estudianteId =
      document.querySelector('#visita-estudiante').value;

    const respuesta = await getDocs(
      query(
        collection(db, 'estudiantes'),
        where('__name__', '==', estudianteId),
        limit(1)
      )
    );

 if (respuesta.empty) {
  throw new Error('No se encontró al estudiante seleccionado.');
}

const estudiante = respuesta.docs[0].data();

    const datos = {
      estudianteId: estudianteId,
      estudianteNombre: nombreEstudianteVisitasShalom(estudiante),
      curso: cursoEstudianteVisitasShalom(estudiante),
      fecha: document.querySelector('#visita-fecha').value,
      hora: document.querySelector('#visita-hora').value,
      responsable: document.querySelector('#visita-responsable').value.trim(),
      motivo: document.querySelector('#visita-motivo').value.trim(),
      recordatorio: document.querySelector('#visita-recordatorio').value,
      estado: 'Pendiente',
      creadoEn: new Date().toISOString()
    };

    await addDoc(collection(db, 'visitas_estudiantes'), datos);

    evento.target.reset();
    document.querySelector('#visita-fecha').value =
      new Date().toLocaleDateString('en-CA');

    mensaje.textContent = 'Visita programada correctamente.';

    await cargarDatosVisitasShalom();

  } catch (error) {
    console.error('Error al programar visita:', error);
    mensaje.textContent = 'No se pudo programar la visita: ' + error.message;
  } finally {
    if (boton) boton.disabled = false;
  }
}

// ==========================================
// CARGAR REGISTROS Y ESTADÍSTICAS
// ==========================================


async function cargarDatosVisitasShalom() {
  const mensaje = document.querySelector('#visitas-mensaje');

  const tablaVisitantes = document.querySelector('#tabla-visitantes');
  const tablaVisitas = document.querySelector('#tabla-visitas');
  const tablaVisitasProyectos = document.querySelector(
    '#tabla-visitas-proyectos'
  );

  try {
    mensaje.textContent = 'Consultando registros de Firebase...';

    const [
      respuestaVisitantes,
      respuestaVisitas,
      respuestaVisitasProyectos
    ] = await Promise.all([
      getDocs(collection(db, 'visitantes_feria')),
      getDocs(collection(db, 'visitas_estudiantes')),
      getDocs(collection(db, 'visitas_proyectos'))
    ]);

    const visitantes = respuestaVisitantes.docs.map(d => ({
      id: d.id,
      ...d.data()
    }));

    const visitas = respuestaVisitas.docs.map(d => ({
      id: d.id,
      ...d.data()
    }));

    const visitasProyectos = respuestaVisitasProyectos.docs.map(d => ({
      id: d.id,
      ...d.data()
    }));

    // INDICADORES
    document.querySelector('#total-visitantes').textContent =
      visitantes.length;

    document.querySelector('#total-programadas').textContent =
      visitas.length;

    document.querySelector('#total-pendientes').textContent =
      visitas.filter(v => v.estado === 'Pendiente').length;

    document.querySelector('#total-realizadas').textContent =
      visitas.filter(v => v.estado === 'Realizada').length;

    // ORDENAR REGISTROS
    visitantes.sort((a, b) =>
      String(b.fecha || '').localeCompare(String(a.fecha || ''))
    );

    visitas.sort((a, b) =>
      String(a.fecha || '').localeCompare(String(b.fecha || ''))
    );

    visitasProyectos.sort((a, b) => {
      const fechaA = String(a.fecha || '') + ' ' + String(a.hora || '');
      const fechaB = String(b.fecha || '') + ' ' + String(b.hora || '');

      return fechaB.localeCompare(fechaA);
    });

    // TABLA DE VISITANTES
    if (!visitantes.length) {
      tablaVisitantes.innerHTML =
        '<tr><td colspan="5">Todavía no hay visitantes registrados.</td></tr>';
    } else {
      tablaVisitantes.innerHTML = visitantes.map((v, i) => `
        <tr>
          <td>${i + 1}</td>
          <td>${escaparVisitasShalom(v.nombre)}</td>
          <td>${escaparVisitasShalom(v.institucion || 'No especificada')}</td>
          <td>${escaparVisitasShalom(v.interes || 'No especificado')}</td>
          <td>${escaparVisitasShalom(v.fecha || 'Sin fecha')}</td>
        </tr>
      `).join('');
    }

    // TABLA DE VISITAS PROGRAMADAS
    if (!visitas.length) {
      tablaVisitas.innerHTML =
        '<tr><td colspan="6">Todavía no hay visitas programadas.</td></tr>';
    } else {
      tablaVisitas.innerHTML = visitas.map(v => `
        <tr>
          <td>${escaparVisitasShalom(v.estudianteNombre)}</td>
          <td>${escaparVisitasShalom(v.curso || 'Sin curso')}</td>
          <td>
            ${escaparVisitasShalom(v.fecha)}
            <br>
            ${escaparVisitasShalom(v.hora)}
          </td>
          <td>${escaparVisitasShalom(v.responsable)}</td>
          <td>
            <span class="badge">
              ${escaparVisitasShalom(v.estado || 'Pendiente')}
            </span>
          </td>
          <td>
            <button type="button"
              data-visita-id="${escaparVisitasShalom(v.id)}"
              data-visita-estado="Realizada">
              REALIZADA
            </button>

            <button type="button"
              data-visita-id="${escaparVisitasShalom(v.id)}"
              data-visita-estado="Pendiente">
              PENDIENTE
            </button>
          </td>
        </tr>
      `).join('');
    }

    // TABLA DE VISITAS A PROYECTOS Y STANDS
    if (!tablaVisitasProyectos) {
      console.error('No se encontró la tabla #tabla-visitas-proyectos.');
    } else if (!visitasProyectos.length) {
      tablaVisitasProyectos.innerHTML =
        '<tr><td colspan="8">Todavía no hay visitas registradas a proyectos.</td></tr>';
    } else {
      tablaVisitasProyectos.innerHTML = visitasProyectos.map((v, i) => `
        <tr>
          <td>${i + 1}</td>
          <td>${escaparVisitasShalom(v.proyectoNombre || 'Sin nombre')}</td>
          <td>${escaparVisitasShalom(v.area || 'Sin área')}</td>
          <td>${escaparVisitasShalom(v.stand || 'Sin stand')}</td>
          <td>${escaparVisitasShalom(v.fecha || 'Sin fecha')}</td>
          <td>${escaparVisitasShalom(v.hora || 'Sin hora')}</td>
          <td>${escaparVisitasShalom(v.staff_nombre || 'Sin responsable')}</td>
<td>
  ${administrador?.rol === 'administrador' ? `
    <button type="button"
      data-eliminar-visita-proyecto="${escaparVisitasShalom(v.id)}">
      Eliminar
    </button>
  ` : ''}
</td>
        </tr>
      `).join('');
    }

    mensaje.textContent =
      'Información actualizada correctamente desde Firebase.';

  } catch (error) {
    console.error('Error real al consultar visitas:', error);

    tablaVisitantes.innerHTML =
      '<tr><td colspan="5">No se pudieron consultar los visitantes.</td></tr>';

    tablaVisitas.innerHTML =
      '<tr><td colspan="6">No se pudieron consultar las visitas.</td></tr>';

    if (tablaVisitasProyectos) {
      tablaVisitasProyectos.innerHTML =
        '<tr><td colspan="8">No se pudieron consultar las visitas a proyectos.</td></tr>';
    }

    mensaje.textContent =
      'Error de Firebase: ' + (error.code || error.message);
  }
}
if (String(administrador?.rol || '').toLowerCase() === 'administrador') {

  document.querySelector('#limpiar-visitantes')
    .addEventListener('click', () =>
      limpiarColeccionVisitasShalom(
        'visitantes_feria',
        'limpiar-visitantes',
        'Visitantes registrados'
      )
    );

  document.querySelector('#limpiar-visitas')
    .addEventListener('click', () =>
      limpiarColeccionVisitasShalom(
        'visitas_estudiantes',
        'limpiar-visitas',
        'Visitas programadas'
      )
    );

  document.querySelector('#limpiar-visitas-proyectos')
    .addEventListener('click', () =>
      limpiarColeccionVisitasShalom(
        'visitas_proyectos',
        'limpiar-visitas-proyectos',
        'Visitas a proyectos y stands'
      )
    );
}
async function limpiarColeccionVisitasShalom(nombreColeccion, botonId, nombreSeccion) {
  const boton = document.querySelector(`#${botonId}`);

  if (
    String(administrador?.rol || '').toLowerCase() !== 'administrador' ||
    !boton
  ) return;

  boton.disabled = true;

  try {
    const respuesta = await getDocs(collection(db, nombreColeccion));
    const registros = respuesta.docs;

    if (!registros.length) {
      alert(`No hay registros en ${nombreSeccion}.`);
      return;
    }

    if (!confirm(`Se eliminarán PERMANENTEMENTE ${registros.length} registros de ${nombreSeccion}. ¿Deseas continuar?`)) return;

    if (!confirm('CONFIRMACIÓN FINAL: ¿Autorizas esta eliminación definitiva?')) return;

    boton.textContent = 'Eliminando...';

    for (let i = 0; i < registros.length; i += 100) {
      await Promise.all(
        registros.slice(i, i + 100).map(registro =>
          deleteDoc(doc(db, nombreColeccion, registro.id))
        )
      );
    }

    await cargarDatosVisitasShalom();
    alert('Registros eliminados correctamente.');

  } catch (error) {
    console.error('Error al limpiar registros:', error);
    alert('Error al eliminar. Revisa los permisos de Firebase.');
    await cargarDatosVisitasShalom();

  } finally {
    boton.disabled = false;
    boton.textContent = 'Limpiar todo';
  }
}
// ==========================================
// CAMBIAR ESTADO DE UNA VISITA
// ==========================================

async function gestionarAccionVisitaShalom(evento) {
  const boton = evento.target.closest('[data-visita-id]');

  if (!boton) return;

  const id = boton.dataset.visitaId;
  const estado = boton.dataset.visitaEstado;

  try {
    boton.disabled = true;

    await updateDoc(
      doc(db, 'visitas_estudiantes', id),
      { estado: estado }
    );

    document.querySelector('#visitas-mensaje').textContent =
      'Estado de la visita actualizado correctamente.';

    await cargarDatosVisitasShalom();

  } catch (error) {
    console.error('Error al actualizar visita:', error);

    document.querySelector('#visitas-mensaje').textContent =
      'No se pudo actualizar la visita: ' + error.message;

    boton.disabled = false;
  }
}

async function guardarVisitante(evento) {
  evento.preventDefault();

  const boton = evento.submitter;
  boton.disabled = true;

  try {
    await addDoc(collection(db, 'visitantes_feria'), {
      nombre: document.querySelector('#visitante-nombre').value.trim(),
      institucion: document.querySelector('#visitante-institucion').value.trim(),
      telefono: document.querySelector('#visitante-telefono').value.trim(),
      interes: document.querySelector('#visitante-interes').value === 'Otro'
  ? document.querySelector('#visitante-otro-interes').value.trim()
  : document.querySelector('#visitante-interes').value,
      atendido: document.querySelector('#visitante-atendido').value.trim(),
      fecha: new Date().toISOString(),
      creadoEn: new Date().toISOString()
    });

    alert('Visitante registrado correctamente.');
    evento.target.reset();
    await cargarDatosVisitas();

  } catch (error) {
    console.error('Error al registrar visitante:', error);
    alert('No se pudo registrar al visitante.');
  } finally {
    boton.disabled = false;
  }
}

async function cargarEstudiantesVisitas() {
  const selector = document.querySelector('#visita-estudiante');

  try {
    const resultado = await getDocs(collection(db, 'estudiantes'));

    const estudiantes = resultado.docs
      .map(documento => ({
        id: documento.id,
        ...documento.data()
      }))
      .filter(estudiante => estudiante.activo !== false)
      .sort((a, b) =>
        `${a.apellidos || ''} ${a.nombres || ''}`.localeCompare(
          `${b.apellidos || ''} ${b.nombres || ''}`, 'es'
        )
      );

    selector.innerHTML = '<option value="">Seleccionar estudiante</option>' +
      estudiantes.map(estudiante => {
        const nombre = `${estudiante.apellidos || ''} ${estudiante.nombres || ''}`.trim();
        const curso = estudiante.curso || '';
        const paralelo = estudiante.paralelo || '';
        const nivel = estudiante.nivel || '';
        return `<option value="${escapar(estudiante.id)}"
          data-nombre="${escapar(nombre)}"
          data-curso="${escapar(curso)}"
          data-paralelo="${escapar(paralelo)}"
          data-nivel="${escapar(nivel)}">
          ${escapar(nombre)} – ${escapar(curso)} ${escapar(paralelo)}
        </option>`;
      }).join('');

  } catch (error) {
    console.error('Error al cargar estudiantes:', error);
    selector.innerHTML = '<option value="">Error al cargar estudiantes</option>';
  }
}

async function guardarVisitaEstudiante(evento) {
  evento.preventDefault();

  const boton = evento.submitter;
  boton.disabled = true;

  try {
    const selector = document.querySelector('#visita-estudiante');
    const opcion = selector.options[selector.selectedIndex];

    if (!selector.value || !opcion) {
      alert('Selecciona un estudiante.');
      return;
    }

    await addDoc(collection(db, 'visitas_estudiantes'), {
      estudianteId: selector.value,
      estudiante: opcion.dataset.nombre,
      curso: opcion.dataset.curso,
      paralelo: opcion.dataset.paralelo,
      nivel: opcion.dataset.nivel,
      fecha: document.querySelector('#visita-fecha').value,
      hora: document.querySelector('#visita-hora').value,
      responsable: document.querySelector('#visita-responsable').value.trim(),
      motivo: document.querySelector('#visita-motivo').value.trim(),
      recordatorio: document.querySelector('#visita-recordatorio').value,
      estado: 'Pendiente',
      creadoEn: new Date().toISOString()
    });

    alert('Visita programada correctamente.');
    evento.target.reset();
    await cargarDatosVisitas();

  } catch (error) {
    console.error('Error al programar visita:', error);
    alert('No se pudo programar la visita.');
  } finally {
    boton.disabled = false;
  }
}

async function cargarDatosVisitas() {
  try {
    const [visitantesDB, visitasDB] = await Promise.all([
      getDocs(collection(db, 'visitantes_feria')),
      getDocs(collection(db, 'visitas_estudiantes'))
    ]);

    const visitantes = visitantesDB.docs.map(doc => doc.data());

    const visitas = visitasDB.docs
      .map(documento => ({
        id: documento.id,
        ...documento.data()
      }))
      .sort((a, b) =>
        `${a.fecha || ''} ${a.hora || ''}`.localeCompare(
          `${b.fecha || ''} ${b.hora || ''}`
        )
      );

    document.querySelector('#total-visitantes').textContent = visitantes.length;
    document.querySelector('#total-programadas').textContent = visitas.length;
    document.querySelector('#total-pendientes').textContent =
      visitas.filter(visita => visita.estado === 'Pendiente').length;

    document.querySelector('#tabla-visitantes').innerHTML =
      visitantes.length
        ? visitantes.map((visitante, indice) => `
            <tr>
              <td>${indice + 1}</td>
              <td>${escapar(visitante.nombre || '')}</td>
              <td>${escapar(visitante.institucion || '—')}</td>
              <td>${escapar(visitante.interes || '')}</td>
              <td>${escapar(new Date(visitante.fecha).toLocaleString('es-BO'))}</td>
            </tr>
          `).join('')
        : '<tr><td colspan="5">Aún no hay visitantes registrados.</td></tr>';

    document.querySelector('#tabla-visitas').innerHTML =
      visitas.length
        ? visitas.map(visita => `
            <tr>
              <td>${escapar(visita.estudiante || '')}</td>
              <td>${escapar(`${visita.nivel || ''} ${visita.curso || ''} ${visita.paralelo || ''}`)}</td>
              <td>${escapar(`${visita.fecha || ''} ${visita.hora || ''}`)}</td>
              <td>${escapar(visita.responsable || '')}</td>
              <td>
                <select onchange="cambiarEstadoVisita('${visita.id}', this.value)">
                  <option ${visita.estado === 'Pendiente' ? 'selected' : ''}>Pendiente</option>
                  <option ${visita.estado === 'Realizada' ? 'selected' : ''}>Realizada</option>
                  <option ${visita.estado === 'Reprogramada' ? 'selected' : ''}>Reprogramada</option>
                  <option ${visita.estado === 'Cancelada' ? 'selected' : ''}>Cancelada</option>
                </select>
              </td>
              <td>
                <button onclick="eliminarVisita('${visita.id}')">Eliminar</button>
              </td>
            </tr>
          `).join('')
        : '<tr><td colspan="6">No hay visitas programadas.</td></tr>';

  } catch (error) {
    console.error('Error al cargar visitas:', error);
    alert('No se pudieron consultar los registros de visitas.');
  }
}

async function cambiarEstadoVisita(id, estado) {
  try {
    await updateDoc(doc(db, 'visitas_estudiantes', id), { estado });
    await cargarDatosVisitas();
  } catch (error) {
    console.error(error);
    alert('No se pudo actualizar el estado.');
  }
}

async function eliminarVisita(id) {
  if (!confirm('¿Deseas eliminar esta visita programada?')) return;

  try {
    await deleteDoc(doc(db, 'visitas_estudiantes', id));
    await cargarDatosVisitas();
  } catch (error) {
    console.error(error);
    alert('No se pudo eliminar la visita.');
  }
}

// Permite que los botones de las tablas llamen a estas funciones
window.cambiarEstadoVisita = cambiarEstadoVisita;
window.eliminarVisita = eliminarVisita;

// ==========================================
// MÓDULO DE EVALUACIÓN - SHALOM EXPO 2026
// ==========================================

let evaluacionEditandoId = null;
let estudiantesEvaluacion = [];
let seleccionEvaluacion = [];
let evaluacionesCache = [];

const criteriosEvaluacion = [
  {
    nombre: 'Pertinencia curricular',
    maximo: 8,
    detalles: [
      ['Relación con el área o materia', 3],
      ['Correspondencia curricular', 3],
      ['Aplicación de aprendizajes', 2]
    ]
  },
  {
    nombre: 'Texto bíblico',
    maximo: 5,
    detalles: [
      ['Selección del texto bíblico', 2],
      ['Relación entre Biblia y tema', 2],
      ['Aplicación del principio bíblico', 1]
    ]
  },
  {
    nombre: 'Dominio del tema',
    maximo: 8,
    detalles: [
      ['Conocimiento del tema', 3],
      ['Seguridad en las respuestas', 2],
      ['Investigación realizada', 2],
      ['Uso correcto de conceptos', 1]
    ]
  },
  {
    nombre: 'Producto y creatividad',
    maximo: 8,
    detalles: [
      ['Calidad del producto', 3],
      ['Originalidad', 2],
      ['Diseño y presentación', 2],
      ['Uso de materiales', 1]
    ]
  },
  {
    nombre: 'Exposición y participación',
    maximo: 7,
    detalles: [
      ['Claridad de exposición', 2],
      ['Organización', 2],
      ['Participación', 2],
      ['Manejo del tiempo', 1]
    ]
  },
  {
    nombre: 'Responsabilidad y valores',
    maximo: 5,
    detalles: [
      ['Puntualidad y cumplimiento', 2],
      ['Orden y cuidado', 1],
      ['Respeto y cooperación', 2]
    ]
  },
  {
    nombre: 'Integración y aplicación',
    maximo: 4,
    detalles: [
      ['Aplicación a la vida cotidiana', 2],
      ['Aporte a la comunidad', 2]
    ]
  }
];

function escaparEvaluacion(valor) {
  return String(valor ?? '').replace(/[&<>"']/g, caracter => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  })[caracter]);
}

function fechaLocalEvaluacion() {
  const fecha = new Date();
  const zona = fecha.getTimezoneOffset();
  return new Date(fecha.getTime() - zona * 60000)
    .toISOString().slice(0, 10);
}

function nombreCompletoEvaluacion(estudiante) {
  return `${estudiante.apellidos || ''} ${estudiante.nombres || ''}`.trim();
}

function cursoCompletoEvaluacion(estudiante) {
  return [
    estudiante.nivel,
    estudiante.curso,
    estudiante.paralelo
  ].filter(Boolean).join(' ');
}

function calcularTotalEvaluacion() {
  let total = 0;

  criteriosEvaluacion.forEach((criterio, indice) => {
    const campo = document.querySelector(`#eval-criterio-${indice}`);
    total += Number(campo?.value || 0);
  });

  return total;
}

async function cargarEstudiantesEvaluacion() {
  const resultado = await getDocs(collection(db, 'estudiantes'));

  estudiantesEvaluacion = resultado.docs.map(documento => {
    const datos = documento.data();

    return {
      id: documento.id,
      ...datos
    };
  })
  .filter(estudiante => estudiante.activo !== false)
  .sort((a, b) =>
    nombreCompletoEvaluacion(a).localeCompare(
      nombreCompletoEvaluacion(b), 'es'
    )
  );
}

function renderizarListaEstudiantesEvaluacion() {
  const lista = document.querySelector('#eval-lista-estudiantes');
  const busqueda = document.querySelector('#eval-buscar-estudiante');
  const tipo = document.querySelector('#eval-tipo');

  if (!lista) return;

  const texto = (busqueda?.value || '').toLowerCase();

  const filtrados = estudiantesEvaluacion.filter(estudiante => {
    const nombre = nombreCompletoEvaluacion(estudiante).toLowerCase();
    const curso = cursoCompletoEvaluacion(estudiante).toLowerCase();

    return nombre.includes(texto) || curso.includes(texto);
  });

  if (!filtrados.length) {
    lista.innerHTML = '<p class="eval-vacio">No se encontraron estudiantes.</p>';
    return;
  }

  lista.innerHTML = filtrados.map(estudiante => {
    const seleccionado = seleccionEvaluacion.some(
      item => item.id === estudiante.id
    );

    return `
      <label class="eval-estudiante">
        <input
          type="checkbox"
          class="eval-check-estudiante"
          value="${escaparEvaluacion(estudiante.id)}"
          ${seleccionado ? 'checked' : ''}
        >
        <span>
          <strong>${escaparEvaluacion(nombreCompletoEvaluacion(estudiante))}</strong>
          <small>${escaparEvaluacion(cursoCompletoEvaluacion(estudiante))}</small>
        </span>
      </label>
    `;
  }).join('');
}

function actualizarResumenSeleccionEvaluacion() {
  const resumen = document.querySelector('#eval-resumen-seleccion');
  const tipo = document.querySelector('#eval-tipo');

  if (resumen) {
    resumen.textContent =
      `${seleccionEvaluacion.length} estudiante(s) seleccionado(s)`;
  }

  const selector = document.querySelector('#eval-tipo');
  if (selector) {
    selector.value = tipo?.value || 'Grupal';
  }
}

function renderizarCriteriosEvaluacion() {
  return criteriosEvaluacion.map((criterio, indice) => `
    <div class="evaluacion-criterio">
      <div class="eval-criterio-cabecera">
        <strong>${indice + 1}. ${criterio.nombre}</strong>
        <span>Máximo ${criterio.maximo} puntos</span>
      </div>

      ${criterio.detalles.map(detalle => `
        <div class="eval-detalle">
          <span>${detalle[0]} <small>(${detalle[1]} pts.)</small></span>
        </div>
      `).join('')}

      <label class="eval-puntaje-label">
        Puntaje obtenido
        <input
          type="number"
          id="eval-criterio-${indice}"
          class="eval-puntaje"
          min="0"
          max="${criterio.maximo}"
          step="1"
          value="0"
          required
        >
      </label>
    </div>
  `).join('');
}

async function mostrarEvaluacion(proyecto = null) {
    const contenido = document.querySelector('#contenido') || app;

  contenido.innerHTML = `
    <div class="evaluacion-panel ${modoJurado === 'evaluacion-jurado' ? 'evaluacion-panel-jurado' : 'evaluacion-panel-admin'}">

      ${modoJurado === 'evaluacion-jurado' && juradoActual ? `
        <section class="eval-tarjeta">
          <div class="eval-titulo-seccion">
            <div>
              <span class="eval-etiqueta">JURADO IDENTIFICADO</span>
              <h3>${escaparEvaluacion(juradoActual.nombreCompleto)}</h3>
              <p>Código personal: ${escaparEvaluacion(juradoActual.codigo)}</p>
            </div>
            <span class="badge-jurado">IDENTIFICADO</span>
          </div>
        </section>
      ` : ''}

      <header class="eval-encabezado">
        <div>
          <span class="eval-etiqueta">SHALOM EXPO 2026</span>
          <h2>Evaluación de proyectos</h2>
          <p>Registro y calificación de las exposiciones estudiantiles.</p>
          <button type="button" id="volver-inicio-evaluacion" class="asistencia-btn asistencia-btn-principal">
            ← INICIO
          </button>
        </div>

        <div class="eval-total-general">
          <strong>45</strong>
          <span>PUNTOS</span>
        </div>
      </header>

      <section class="eval-tarjeta" style="display:${proyecto ? 'none' : 'block'};">
        <h3>1. Datos de la evaluación</h3>

        <div class="eval-grid">
          <div>
            <label>Modalidad de evaluación</label>
            <select id="eval-tipo">
              <option value="Grupal">Evaluación grupal</option>
              <option value="Personal">Evaluación personal</option>
            </select>
          </div>

          <div>
            <label>Área o materia</label>
            <select id="eval-area" required>
              <option value="">Seleccionar área</option>
              <option>Artes Plásticas</option>
              <option>Biología</option>
              <option>Ciencias Sociales</option>
              <option>Educación Física</option>
              <option>Filosofía</option>
              <option>Física</option>
              <option>Formación Cristiana</option>
              <option>Inglés</option>
              <option>Lenguaje</option>
              <option>Matemática</option>
              <option>Orientación Vocacional</option>
              <option>Psicología</option>
              <option>Química</option>
              <option>Robótica</option>
              <option>Otra</option>
            </select>
          </div>

          <div class="eval-campo-completo">
            <label>Título del tema o proyecto</label>
            <input
              type="text"
              id="eval-tema"
              placeholder="Ej. La importancia del cuidado del medio ambiente"
              required
            >
          </div>

          ${modoJurado === 'evaluacion-jurado' ? '' : `
            <div>
              <label>Nombre del jurado o evaluador</label>
              <input
                type="text"
                id="eval-jurado"
                placeholder="Nombre del evaluador"
                required
              >
            </div>
          `}

          <div>
            <label>Fecha de evaluación</label>
            <input type="date" id="eval-fecha" value="${fechaLocalEvaluacion()}" required>
          </div>
        </div>
      </section>

      <section class="eval-tarjeta" style="display:${proyecto ? 'none' : 'block'};">
        <div class="eval-titulo-seccion">
          <div>
            <h3>2. Selección de estudiantes</h3>
            <p>Selecciona uno o varios estudiantes registrados.</p>
          </div>
          <span id="eval-resumen-seleccion">0 estudiantes seleccionados</span>
        </div>

        <input
          type="search"
          id="eval-buscar-estudiante"
          placeholder="Buscar por nombre, curso o paralelo..."
        >

        <div id="eval-lista-estudiantes" class="eval-lista-estudiantes">
          <p class="eval-vacio">Cargando estudiantes...</p>
        </div>

        <div class="eval-qr">
          <h3>REGISTRAR ESTUDIANTES CON QR</h3>

          <div class="eval-qr-opciones">
            <button type="button" id="eval-qr-individual">
              MODO INDIVIDUAL
            </button>

            <button type="button" id="eval-qr-grupal">
              MODO GRUPAL
            </button>
          </div>

          <div id="eval-lector-qr" style="display:none;">
            <div id="lector-qr-evaluacion"></div>
            <p id="eval-qr-resultado">
              Selecciona una modalidad para comenzar.
            </p>

            <button type="button" id="eval-qr-detener">
              DETENER CÁMARA
            </button>
          </div>
        </div>

        <div class="eval-confirmacion">
          <button type="button" id="eval-confirmar-grupo">
            CONFIRMAR GRUPO
          </button>
          <span id="eval-estado-grupo">
            Selecciona tus estudiantes y confirma.
          </span>
        </div>
      </section>

      ${proyecto ? `
        <section class="eval-tarjeta">
          <h3>2. Datos del proyecto e integrantes</h3>

          <div class="eval-grid">
            <div>
              <label>Nombre del proyecto</label>
              <strong>${escaparEvaluacion(proyecto.nombre || '')}</strong>
            </div>

            <div>
              <label>Área</label>
              <strong>${escaparEvaluacion(proyecto.area || '')}</strong>
            </div>

            <div>
              <label>Stand</label>
              <strong>${escaparEvaluacion(proyecto.stand || 'Sin asignar')}</strong>
            </div>

            <div class="eval-campo-completo">
              <label>Descripción</label>
              <p>${escaparEvaluacion(proyecto.descripcion || 'Sin descripción')}</p>
            </div>
          </div>

          <h4>Estudiantes integrantes</h4>

          <div class="eval-lista-estudiantes">
            ${(proyecto.estudiantes || []).map((estudiante, indice) => `
              <div class="eval-estudiante">
                <span>
                  <strong>${indice + 1}. ${escaparEvaluacion(estudiante.nombre || 'Estudiante')}</strong>
                </span>
              </div>
            `).join('') || '<p>No hay integrantes registrados.</p>'}
          </div>
        </section>
      ` : ''}

      <section class="eval-tarjeta">
        <h3>3. Rúbrica de evaluación</h3>
        <p class="eval-ayuda">Ingresa el puntaje obtenido en cada criterio.</p>

        <div class="eval-rubrica">
          ${renderizarCriteriosEvaluacion()}
        </div>

        <div class="eval-total-final">
          <span>CALIFICACIÓN TOTAL</span>
          <strong id="eval-total">0 / 45</strong>
        </div>
      </section>

      <section class="eval-tarjeta">
        <h3>4. Observaciones</h3>
        <textarea
          id="eval-observaciones"
          rows="3"
          placeholder="Observaciones y recomendaciones..."
        ></textarea>
      </section>

      <div class="eval-acciones">
        <button id="eval-guardar" class="eval-btn-principal">
          Guardar evaluación
        </button>
        <button id="eval-limpiar" class="eval-btn-secundario">
          Limpiar formulario
        </button>
      </div>

      <section class="eval-tarjeta">
        <div class="eval-titulo-seccion">
          <div>
            <h3>5. Evaluaciones registradas</h3>
            <p>Consulta, modifica o exporta los resultados.</p>
          </div>

          ${modoJurado !== 'evaluacion-jurado' ? `
  <div class="eval-acciones-historial">
    <button id="eval-exportar" class="eval-btn-exportar">
      EXPORTAR EXCEL
    </button>

    <button id="eval-limpiar-todo" class="eval-btn-secundario">
      LIMPIAR TODO
    </button>
  </div>
` : ''}
        </div>

        <input
          type="search"
          id="eval-buscar-registro"
          placeholder="Buscar por tema, estudiante, área o jurado..."
        >

        <div class="eval-tabla-contenedor">
          <table class="tabla-evaluaciones">
            <thead>
              <tr>
                <th>TEMA</th>
                <th>TIPO</th>
                <th>ESTUDIANTES</th>
                <th>ÁREA</th>
                <th>JURADO</th>
                <th>FECHA</th>
                <th>PUNTAJE</th>
                <th>ACCIÓN</th>
              </tr>
            </thead>
            <tbody id="eval-tabla">
              <tr><td colspan="8">Consultando registros...</td></tr>
            </tbody>
          </table>
        </div>
      </section>

    </div>
  `;

try {
  await cargarEstudiantesEvaluacion();
// CARGAR DATOS DEL PROYECTO CUANDO SE INGRESA POR QR
if (proyecto) {

  document.querySelector('#eval-area').value =
    proyecto.area || '';

  document.querySelector('#eval-tema').value =
    proyecto.nombre || '';

  const idsEstudiantesProyecto = (proyecto.estudiantes || [])
    .map(estudiante => estudiante.id);

  seleccionEvaluacion = estudiantesEvaluacion.filter(
    estudiante => idsEstudiantesProyecto.includes(estudiante.id)
  );

  document.querySelector('#eval-tipo').value =
    seleccionEvaluacion.length > 1 ? 'Grupal' : 'Personal';

  const botonConfirmarProyecto =
    document.querySelector('#eval-confirmar-grupo');

  botonConfirmarProyecto.disabled = true;
  botonConfirmarProyecto.textContent =
    'GRUPO DEL PROYECTO CONFIRMADO';
}
  // MODO INDIVIDUAL
  document.querySelector('#eval-qr-individual')
    .addEventListener('click', () => {
      tipo.value = 'Personal';

      const botonConfirmar =
        document.querySelector('#eval-confirmar-grupo');

      botonConfirmar.disabled = true;
      botonConfirmar.style.display = 'none';

      document.querySelector('#eval-estado-grupo').textContent =
        'Modo individual: escanea un estudiante para evaluarlo.';

      iniciarQREvaluacion('individual');
    });

  // MODO GRUPAL
  document.querySelector('#eval-qr-grupal')
    .addEventListener('click', () => {
      tipo.value = 'Grupal';

      const botonConfirmar =
        document.querySelector('#eval-confirmar-grupo');

      botonConfirmar.disabled = false;
      botonConfirmar.style.display = '';
      botonConfirmar.textContent = 'CONFIRMAR GRUPO';

      document.querySelector('#eval-estado-grupo').textContent =
        'Escanea los estudiantes y confirma el grupo para continuar.';

      iniciarQREvaluacion('grupal');
    });

  // DETENER CÁMARA
  document.querySelector('#eval-qr-detener')
    .addEventListener('click', () => {
      detenerQREvaluacion();
      document.querySelector('#eval-qr-resultado').textContent =
        'Cámara detenida.';
    });

  renderizarListaEstudiantesEvaluacion();
  await cargarEvaluaciones();

} catch (error) {
  console.error('Error al cargar evaluación:', error);

  document.querySelector('#eval-lista-estudiantes').innerHTML =
    '<p class="eval-error">No se pudieron cargar los estudiantes. Verifica Firebase.</p>';
}

  const tipo = document.querySelector('#eval-tipo');

  tipo.addEventListener('change', () => {
  if (tipo.value === 'Personal' && seleccionEvaluacion.length > 1) {
    seleccionEvaluacion = seleccionEvaluacion.slice(0, 1);
    renderizarListaEstudiantesEvaluacion();
    alert('La evaluación personal permite seleccionar solamente un estudiante.');
  }

  const botonConfirmarGrupo =
    document.querySelector('#eval-confirmar-grupo');

  if (tipo.value === 'Personal') {
    botonConfirmarGrupo.disabled = true;
    botonConfirmarGrupo.style.display = 'none';
  } else {
    botonConfirmarGrupo.disabled = false;
    botonConfirmarGrupo.style.display = '';
    botonConfirmarGrupo.textContent = 'CONFIRMAR GRUPO';
  }

  actualizarResumenSeleccionEvaluacion();
});

  let grupoConfirmado = false;

  const botonConfirmar = document.querySelector('#eval-confirmar-grupo');

botonConfirmar.addEventListener('click', async () => {

  if (seleccionEvaluacion.length === 0) {
    alert('Primero escanea o selecciona al menos un estudiante.');
    return;
  }

  grupoConfirmado = true;

  await detenerQREvaluacion();

  document.querySelectorAll('.eval-check-estudiante').forEach(check => {
    check.disabled = true;
  });

  document.querySelector('#eval-qr-individual').disabled = true;
  document.querySelector('#eval-qr-grupal').disabled = true;

  botonConfirmar.textContent = 'GRUPO CONFIRMADO';
  botonConfirmar.disabled = true;

  document.querySelector('#eval-estado-grupo').textContent =
    `${seleccionEvaluacion.length} estudiante(s) confirmados para recibir una calificación grupal.`;

});

  document.querySelector('#eval-buscar-estudiante')
    .addEventListener('input', renderizarListaEstudiantesEvaluacion);

  document.querySelector('#eval-lista-estudiantes')
    .addEventListener('change', event => {
      const check = event.target.closest('.eval-check-estudiante');
      if (!check) return;

      const estudiante = estudiantesEvaluacion.find(
        item => item.id === check.value
      );

      if (!estudiante) return;

      if (check.checked) {
        if (
          tipo.value === 'Personal' &&
          seleccionEvaluacion.length >= 1
        ) {
          check.checked = false;
          alert('En evaluación personal solo puedes seleccionar un estudiante.');
          return;
        }

        if (!seleccionEvaluacion.some(item => item.id === estudiante.id)) {
          seleccionEvaluacion.push(estudiante);
        }
      } else {
        seleccionEvaluacion = seleccionEvaluacion.filter(
          item => item.id !== estudiante.id
        );
      }

      actualizarResumenSeleccionEvaluacion();
    });

  document.querySelectorAll('.eval-puntaje').forEach(campo => {
    campo.addEventListener('input', () => {
      const maximo = Number(campo.max);
      let valor = Number(campo.value || 0);

      if (valor < 0) valor = 0;
      if (valor > maximo) valor = maximo;

      campo.value = valor;
      document.querySelector('#eval-total').textContent =
        `${calcularTotalEvaluacion()} / 45`;
    });
  });
// REGRESAR AL PANEL CORRECTO
document
  .querySelector('#volver-inicio-evaluacion')
  .addEventListener('click', async () => {

    if (modoJurado === 'evaluacion-jurado') {

      mostrarPanelJurado();

    } else {

      await detenerEscanerQR();
      mostrarPanel(administrador?.nombre || 'Administrador');

    }

  });
  document.querySelector('#eval-guardar')
    .addEventListener('click', guardarEvaluacion);

  document.querySelector('#eval-limpiar')
    .addEventListener('click', limpiarFormularioEvaluacion);

  document.querySelector('#eval-buscar-registro')
    .addEventListener('input', mostrarTablaEvaluaciones);

  if (modoJurado !== 'evaluacion-jurado') {
    document.querySelector('#eval-exportar')
      ?.addEventListener('click', exportarEvaluaciones);
  }
if (modoJurado !== 'evaluacion-jurado') {
  document.querySelector('#eval-limpiar-todo')
    ?.addEventListener('click', limpiarTodasLasEvaluaciones);
}
  document.querySelector('#eval-tabla')
    .addEventListener('click', async event => {
      const editar = event.target.closest('[data-eval-editar]');
      const eliminar = event.target.closest('[data-eval-eliminar]');

      if (editar) await editarEvaluacion(editar.dataset.evalEditar);
      if (eliminar) await eliminarEvaluacion(eliminar.dataset.evalEliminar);
    });
}

let lectorQREvaluacion = null;
let modoQREvaluacion = 'individual';
let procesandoQREvaluacion = false;

async function iniciarQREvaluacion(modo) {
  modoQREvaluacion = modo;

  const contenedor = document.querySelector('#eval-lector-qr');
  const resultado = document.querySelector('#eval-qr-resultado');

  contenedor.style.display = 'block';
  resultado.textContent = 'Preparando cámara...';

  await detenerQREvaluacion();

  lectorQREvaluacion = new Html5Qrcode('lector-qr-evaluacion');

  try {
    await lectorQREvaluacion.start(
      { facingMode: 'environment' },
      { fps: 10, qrbox: { width: 250, height: 250 } },
      procesarQREvaluacion,
      () => {}
    );

    resultado.textContent = 'Cámara activa. Escanea el QR del estudiante.';
  } catch (error) {
    console.error(error);
    resultado.textContent = 'No se pudo activar la cámara. Verifica los permisos.';
  }
}

async function procesarQREvaluacion(textoQR) {
  if (procesandoQREvaluacion) return;

  procesandoQREvaluacion = true;

  const resultado = document.querySelector('#eval-qr-resultado');

  try {
    const consulta = query(
      collection(db, 'estudiantes'),
      where('qr_id', '==', textoQR),
      limit(1)
    );

    const respuesta = await getDocs(consulta);

    if (respuesta.empty) {
      resultado.textContent = 'Estudiante no encontrado.';
      return;
    }

    const documento = respuesta.docs[0];
    const estudiante = documento.data();

    if (estudiante.activo === false) {
      resultado.textContent = 'Este estudiante está inactivo.';
      return;
    }

    const id = documento.id;

    if (seleccionEvaluacion.some(e => e.id === id)) {
      resultado.textContent = 'Este estudiante ya está seleccionado.';
      return;
    }

    if (modoQREvaluacion === 'individual' && seleccionEvaluacion.length > 0) {
      resultado.textContent = 'Ya existe un estudiante seleccionado. Limpia la selección para evaluar otro.';
      return;
    }

    seleccionEvaluacion.push({
      id,
      nombre: nombreCompletoEvaluacion(estudiante),
      curso: estudiante.curso || '',
      paralelo: estudiante.paralelo || '',
      nivel: estudiante.nivel || ''
    });

    renderizarListaEstudiantesEvaluacion();
    actualizarResumenSeleccionEvaluacion();

    resultado.textContent =
      `Estudiante agregado: ${nombreCompletoEvaluacion(estudiante)}`;

    if (modoQREvaluacion === 'individual') {
      await detenerQREvaluacion();
    }

  } catch (error) {
    console.error(error);
    resultado.textContent = 'Error al consultar el estudiante.';
  } finally {
    setTimeout(() => {
      procesandoQREvaluacion = false;
    }, 1200);
  }
}

async function detenerQREvaluacion() {
  if (lectorQREvaluacion) {
    try {
      if (lectorQREvaluacion.isScanning) {
        await lectorQREvaluacion.stop();
      }

      await lectorQREvaluacion.clear();
    } catch (error) {
      console.error(error);
    }

    lectorQREvaluacion = null;
  }
}


async function guardarEvaluacion() {
  const boton = document.querySelector('#eval-guardar');
  const tipo = document.querySelector('#eval-tipo').value;
  const tema = document.querySelector('#eval-tema').value.trim();
  const area = document.querySelector('#eval-area').value;
  const fecha = document.querySelector('#eval-fecha').value;

  const esModoJurado = modoJurado === 'evaluacion-jurado';

  const campoJurado = document.querySelector('#eval-jurado');
  let jurado = campoJurado?.value.trim() || '';
  let identidadJurado = null;

  // Identificar automáticamente al jurado
  if (esModoJurado) {
    const usuarioJurado = authJurados.currentUser;

    if (!usuarioJurado) {
      alert('No se encontró la sesión del jurado. Vuelve a registrarte.');
      return;
    }

    try {
      const referenciaJurado = doc(
        db,
        'jurados',
        usuarioJurado.uid
      );

      const resultadoJurado = await getDoc(referenciaJurado);

      if (!resultadoJurado.exists() || resultadoJurado.data().activo !== true) {
        alert('El registro del jurado no existe o está inactivo.');
        return;
      }

      identidadJurado = resultadoJurado.data();
      jurado = identidadJurado.nombreCompleto;

    } catch (error) {
      console.error('Error al verificar jurado:', error);
      alert('No se pudo verificar tu identidad como jurado.');
      return;
    }

  } else if (!jurado) {
    alert('Ingresa el nombre del jurado.');
    return;
  }

  if (!tema || !area || !fecha) {
    alert('Completa todos los datos de la evaluación.');
    return;
  }

  const botonConfirmarGrupo =
  document.querySelector('#eval-confirmar-grupo');

if (
  tipo === 'Grupal' &&
  !botonConfirmarGrupo.disabled
) {
  alert('Primero debes confirmar el grupo.');
  return;
}

  if (seleccionEvaluacion.length === 0) {
    alert('Selecciona al menos un estudiante.');
    return;
  }

  if (tipo === 'Personal' && seleccionEvaluacion.length !== 1) {
    alert('La evaluación personal requiere exactamente un estudiante.');
    return;
  }

  const puntajes = criteriosEvaluacion.map((criterio, indice) => {
    const obtenido = Number(
      document.querySelector(`#eval-criterio-${indice}`).value || 0
    );

    return {
      criterio: criterio.nombre,
      maximo: criterio.maximo,
      obtenido
    };
  });

  const total = puntajes.reduce(
    (suma, criterio) => suma + criterio.obtenido, 0
  );

  const estudiantes = seleccionEvaluacion.map(estudiante => ({
    id: estudiante.id,
    nombre: nombreCompletoEvaluacion(estudiante),
    nivel: estudiante.nivel || '',
    curso: estudiante.curso || '',
    paralelo: estudiante.paralelo || ''
  }));

  const datos = {
    tipoEvaluacion: tipo,
    tema,
    area,
    curso: [...new Set(estudiantes.map(e => e.curso).filter(Boolean))].join(', '),
    paralelo: [...new Set(estudiantes.map(e => e.paralelo).filter(Boolean))].join(', '),
    estudiantes,
    integrantes: estudiantes.map(e => e.nombre).join(', '),
    jurado,
    fecha,
    puntajes,
    total,
    observaciones: document.querySelector('#eval-observaciones').value.trim(),
    actualizadoEn: new Date().toISOString()
  };

  // Guardar identidad automática del jurado
  if (identidadJurado) {
    datos.jurado_uid = identidadJurado.uid;
    datos.jurado_codigo = identidadJurado.codigo;
    datos.jurado_nombre = identidadJurado.nombreCompleto;
  }

  const modoEdicion = evaluacionEditandoId;

  boton.disabled = true;
  boton.textContent = 'Guardando...';

  try {
    if (modoEdicion) {
      await updateDoc(
        doc(db, 'evaluaciones_feria', modoEdicion),
        datos
      );

      alert('Evaluación actualizada correctamente.');

        } else {
      datos.creadoEn = new Date().toISOString();
      datos.fechaCreacion = serverTimestamp();

      await addDoc(
        collection(db, 'evaluaciones_feria'),
        datos
      );

      alert('Evaluación guardada correctamente.');
    }

    limpiarFormularioEvaluacion();
    await cargarEvaluaciones();

  } catch (error) {
    console.error('Error al guardar evaluación:', error);
    alert('No se pudo guardar. Verifica tu conexión y los permisos de Firebase.');

  } finally {
    boton.disabled = false;
    boton.textContent = evaluacionEditandoId
      ? 'Actualizar evaluación'
      : 'Guardar evaluación';
  }
}

function limpiarFormularioEvaluacion() {
  evaluacionEditandoId = null;
  seleccionEvaluacion = [];

  const formulario = document.querySelector('.evaluacion-panel');
  if (!formulario) return;

  document.querySelector('#eval-tipo').value = 'Grupal';
  document.querySelector('#eval-tema').value = '';
  document.querySelector('#eval-area').value = '';
  if (modoJurado !== 'evaluacion-jurado') {
  document.querySelector('#eval-jurado').value = ''
}
  document.querySelector('#eval-fecha').value = fechaLocalEvaluacion();
  document.querySelector('#eval-observaciones').value = '';
  document.querySelector('#eval-buscar-estudiante').value = '';
  document.querySelector('#eval-buscar-registro').value = '';

  document.querySelectorAll('.eval-puntaje').forEach(campo => {
    campo.value = 0;
  });

  document.querySelector('#eval-total').textContent = '0 / 45';
  document.querySelector('#eval-guardar').textContent = 'Guardar evaluación';

  actualizarResumenSeleccionEvaluacion();
  renderizarListaEstudiantesEvaluacion();
}


async function cargarEvaluaciones() {
  try {
    let consultaEvaluaciones;

    if (modoJurado === 'evaluacion-jurado') {
      const usuarioJurado = authJurados.currentUser;

      if (!usuarioJurado) {
        throw new Error('No hay un jurado identificado.');
      }

      consultaEvaluaciones = query(
        collection(db, 'evaluaciones_feria'),
        where('jurado_uid', '==', usuarioJurado.uid)
      );

    } else {
      consultaEvaluaciones = collection(
        db,
        'evaluaciones_feria'
      );
    }

    const resultado = await getDocs(consultaEvaluaciones);

    evaluacionesCache = resultado.docs.map(documento => ({
      id: documento.id,
      ...documento.data()
    }));

    evaluacionesCache.sort((a, b) =>
      String(b.fecha || '').localeCompare(String(a.fecha || ''))
    );

    mostrarTablaEvaluaciones();

  } catch (error) {
    console.error('Error al cargar evaluaciones:', error);
    alert('No se pudieron cargar las evaluaciones. Verifica los permisos.');
  }
}

function mostrarTablaEvaluaciones() {
  const tabla = document.querySelector('#eval-tabla');
  if (!tabla) return;
if (temporizadorCorreccionJurado) {
  clearTimeout(temporizadorCorreccionJurado);
  temporizadorCorreccionJurado = null;
}
  const esModoJurado = modoJurado === 'evaluacion-jurado';

  const busqueda = (
    document.querySelector('#eval-buscar-registro')?.value || ''
  ).toLowerCase();

  const filtrados = evaluacionesCache.filter(evaluacion => {
    const nombres = (evaluacion.estudiantes || [])
      .map(e => e.nombre)
      .join(' ');

    const texto = [
      evaluacion.tema,
      evaluacion.tipoEvaluacion,
      nombres,
      evaluacion.area,
      evaluacion.jurado,
      evaluacion.curso,
      evaluacion.paralelo
    ].join(' ').toLowerCase();

    return texto.includes(busqueda);
  });

  if (!filtrados.length) {
    tabla.innerHTML = `
      <tr>
        <td colspan="8">No hay evaluaciones registradas.</td>
      </tr>
    `;
    return;
  }

  tabla.innerHTML = filtrados.map(evaluacion => {
    const fechaCreacion = evaluacion.fechaCreacion?.toMillis?.() || 0;

    const puedeCorregir = fechaCreacion > 0 &&
      Date.now() < fechaCreacion + 10 * 60 * 1000;

    return `
      <tr>
        <td>${escaparEvaluacion(evaluacion.tema)}</td>
        <td>${escaparEvaluacion(evaluacion.tipoEvaluacion)}</td>
        <td>${escaparEvaluacion(
          (evaluacion.estudiantes || []).map(e => e.nombre).join(', ') ||
          evaluacion.integrantes || ''
        )}</td>
        <td>${escaparEvaluacion(evaluacion.area)}</td>
        <td>${escaparEvaluacion(evaluacion.jurado)}</td>
        <td>${escaparEvaluacion(evaluacion.fecha)}</td>
        <td><strong>${Number(evaluacion.total || 0)} / 45</strong></td>
        <td>
          ${esModoJurado ? `
            ${puedeCorregir ? `
              <button class="eval-editar"
                data-eval-editar="${escaparEvaluacion(evaluacion.id)}">
                Corregir
              </button>
            ` : `
              <span>Plazo de corrección finalizado</span>
            `}
          ` : `
            <div class="eval-acciones-tabla">
              <button class="eval-editar"
                data-eval-editar="${escaparEvaluacion(evaluacion.id)}">
                Editar
              </button>
              <button class="eval-eliminar"
                data-eval-eliminar="${escaparEvaluacion(evaluacion.id)}">
                Eliminar
              </button>
            </div>
          `}
        </td>
      </tr>
    `;
  }).join('');

  if (esModoJurado) {
    const expiraciones = filtrados
      .map(evaluacion => evaluacion.fechaCreacion?.toMillis?.())
      .filter(fecha => fecha && fecha + 10 * 60 * 1000 > Date.now())
      .map(fecha => fecha + 10 * 60 * 1000);

    if (expiraciones.length > 0) {
      const proximaExpiracion = Math.min(...expiraciones);

      temporizadorCorreccionJurado = setTimeout(() => {
        temporizadorCorreccionJurado = null;
        mostrarTablaEvaluaciones();
      }, proximaExpiracion - Date.now() + 50);
    }
  }
}
async function editarEvaluacion(id) {
  const evaluacion = evaluacionesCache.find(item => item.id === id);
  if (!evaluacion) return;

  evaluacionEditandoId = id;
  seleccionEvaluacion = (evaluacion.estudiantes || []).map(estudiante => {
    return estudiantesEvaluacion.find(item => item.id === estudiante.id) || {
      ...estudiante,
      apellidos: estudiante.nombre || '',
      nombres: ''
    };
  });

  document.querySelector('#eval-tipo').value =
    evaluacion.tipoEvaluacion || 'Grupal';

  document.querySelector('#eval-tema').value = evaluacion.tema || '';
  document.querySelector('#eval-area').value = evaluacion.area || '';
  if (modoJurado !== 'evaluacion-jurado') {
  document.querySelector('#eval-jurado').value = evaluacion.jurado || '';
}
  document.querySelector('#eval-fecha').value = evaluacion.fecha || fechaLocalEvaluacion();
  document.querySelector('#eval-observaciones').value = evaluacion.observaciones || '';

  criteriosEvaluacion.forEach((criterio, indice) => {
    const puntaje = (evaluacion.puntajes || []).find(
      item => item.criterio === criterio.nombre
    );

    document.querySelector(`#eval-criterio-${indice}`).value =
      puntaje?.obtenido || 0;
  });

  document.querySelector('#eval-total').textContent =
    `${calcularTotalEvaluacion()} / 45`;

  document.querySelector('#eval-guardar').textContent =
    'Actualizar evaluación';

  renderizarListaEstudiantesEvaluacion();
  actualizarResumenSeleccionEvaluacion();

  document.querySelector('.eval-tarjeta').scrollIntoView({
    behavior: 'smooth',
    block: 'start'
  });
}

async function eliminarEvaluacion(id) {
  const confirmar = confirm('¿Estás seguro de eliminar esta evaluación?');
  if (!confirmar) return;

  try {
    await deleteDoc(doc(db, 'evaluaciones_feria', id));

    if (evaluacionEditandoId === id) {
      limpiarFormularioEvaluacion();
    }

    await cargarEvaluaciones();
    alert('Evaluación eliminada correctamente.');

  } catch (error) {
    console.error('Error al eliminar evaluación:', error);
    alert('No se pudo eliminar la evaluación.');
  }
}

async function limpiarTodasLasEvaluaciones() {

  // SOLO ADMINISTRADOR
  if (administrador?.rol !== 'administrador') {
    alert('No tienes permiso para eliminar todas las evaluaciones.');
    return;
  }

  const confirmar = confirm(
    '¿ESTÁS SEGURO?\n\n' +
    'Se eliminarán PERMANENTEMENTE todas las evaluaciones registradas.\n\n' +
    'Esta acción no eliminará estudiantes, proyectos ni códigos QR.'
  );

  if (!confirmar) return;

  const segundaConfirmacion = confirm(
    'CONFIRMACIÓN FINAL\n\n' +
    '¿Deseas eliminar todo el historial de evaluaciones?'
  );

  if (!segundaConfirmacion) return;

  const boton = document.querySelector('#eval-limpiar-todo');

  try {
    boton.disabled = true;
    boton.textContent = 'ELIMINANDO...';

    const referencia = collection(db, 'evaluaciones_feria');
    const resultado = await getDocs(referencia);

    await Promise.all(
      resultado.docs.map(documento =>
        deleteDoc(doc(db, 'evaluaciones_feria', documento.id))
      )
    );

    evaluacionesCache = [];

    if (evaluacionEditandoId) {
      limpiarFormularioEvaluacion();
    }

    await cargarEvaluaciones();

    alert('Se eliminaron todas las evaluaciones correctamente.');

  } catch (error) {
    console.error('Error al limpiar evaluaciones:', error);
    alert('No se pudieron eliminar todas las evaluaciones.');

  } finally {
    if (boton) {
      boton.disabled = false;
      boton.textContent = 'LIMPIAR TODO';
    }
  }
}

function exportarEvaluaciones() {
  if (!evaluacionesCache.length) {
    alert('No existen evaluaciones para exportar.');
    return;
  }

  const filas = evaluacionesCache.map(evaluacion => {
    const fila = {
      TIPO: evaluacion.tipoEvaluacion || '',
      TEMA: evaluacion.tema || '',
      ESTUDIANTES: (evaluacion.estudiantes || [])
        .map(e => e.nombre).join(', '),
      AREA: evaluacion.area || '',
      CURSO: evaluacion.curso || '',
      PARALELO: evaluacion.paralelo || '',
      JURADO: evaluacion.jurado || '',
      FECHA: evaluacion.fecha || ''
    };

    (evaluacion.puntajes || []).forEach((puntaje, indice) => {
      fila[puntaje.criterio || `Criterio ${indice + 1}`] =
        puntaje.obtenido;
    });

    fila.TOTAL = evaluacion.total || 0;
    fila.OBSERVACIONES = evaluacion.observaciones || '';

    return fila;
  });

  const hoja = XLSX.utils.json_to_sheet(filas);
  const libro = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(libro, hoja, 'Evaluaciones');
  XLSX.writeFile(libro, 'Evaluaciones_Shalom_Expo_2026.xlsx');
}

async function actualizarEstadisticasCIE() {
  try {
    const [estudiantesDB, asistenciasDB] = await Promise.all([
      getDocs(collection(db, 'estudiantes')),
      getDocs(collection(db, 'asistencias'))
    ]);

    const estudiantes = estudiantesDB.docs
      .map(doc => doc.data())
      .filter(estudiante => estudiante.activo !== false);

    const ahora = new Date();

    const fecha = [
      ahora.getFullYear(),
      String(ahora.getMonth() + 1).padStart(2, '0'),
      String(ahora.getDate()).padStart(2, '0')
    ].join('-');

    const asistenciasHoy = asistenciasDB.docs
      .map(doc => doc.data())
      .filter(registro => registro.fecha === fecha);

    const codigosPresentes = new Set(
      asistenciasHoy.map(registro => String(registro.codigo))
    );

    const total = estudiantes.length;

    const presentes = estudiantes.filter(estudiante =>
      codigosPresentes.has(String(estudiante.codigo))
    ).length;

    const ausentes = Math.max(0, total - presentes);

    const porcentaje = total > 0
      ? ((presentes / total) * 100).toFixed(1)
      : '0.0';

    document.querySelector('#cie-total').textContent = total;
    document.querySelector('#cie-presentes').textContent = presentes;
    document.querySelector('#cie-ausentes').textContent = ausentes;
    document.querySelector('#cie-porcentaje').textContent = porcentaje + '%';
   generarGraficosCIE(estudiantes, codigosPresentes);

    document.querySelector('#cie-fecha').textContent =
      'Fecha de consulta: ' + fecha;

    document.querySelector('#cie-mensaje').textContent =
      'Se encontraron ' + asistenciasHoy.length +
      ' registros de asistencia para hoy.';

  } catch (error) {
    console.error('Error en estadísticas CIE:', error);
    document.querySelector('#cie-mensaje').textContent =
      'No se pudieron cargar las estadísticas.';
  }
}

function dibujarGraficoCircularCIE(presentes, ausentes) {
  const canvas = document.querySelector('#grafico-asistencia');
  if (!canvas) return;

  canvas.width = 500;
  canvas.height = 260;

  const ctx = canvas.getContext('2d');
  const total = presentes + ausentes;

  const cx = 145;
  const cy = 125;
  const radio = 85;

  ctx.clearRect(0, 0, canvas.width, canvas.height);

  if (total === 0) {
    ctx.fillStyle = '#888';
    ctx.font = '16px Arial';
    ctx.fillText('Sin datos disponibles', 120, 130);
    return;
  }

  const angulo = (presentes / total) * Math.PI * 2;

  ctx.beginPath();
  ctx.arc(cx, cy, radio, 0, Math.PI * 2);
  ctx.fillStyle = '#E5E7EB';
  ctx.fill();

  ctx.beginPath();
  ctx.moveTo(cx, cy);
  ctx.arc(cx, cy, radio, -Math.PI / 2, -Math.PI / 2 + angulo);
  ctx.closePath();
  ctx.fillStyle = '#7652A8';
  ctx.fill();

  ctx.fillStyle = '#302052';
  ctx.font = 'bold 25px Arial';
  ctx.textAlign = 'center';
  ctx.fillText(Math.round(presentes / total * 100) + '%', cx, cy + 8);

  ctx.textAlign = 'left';
  ctx.font = '15px Arial';

  ctx.fillStyle = '#7652A8';
  ctx.fillRect(285, 85, 14, 14);
  ctx.fillStyle = '#333';
  ctx.fillText('Presentes: ' + presentes, 310, 98);

  ctx.fillStyle = '#E5E7EB';
  ctx.fillRect(285, 125, 14, 14);
  ctx.fillStyle = '#333';
  ctx.fillText('Ausentes: ' + ausentes, 310, 138);
}

function dibujarBarrasCIE(id, grupos) {
  const canvas = document.querySelector('#' + id);
  if (!canvas) return;

 const datos = grupos;

  canvas.width = 640;
 canvas.height = Math.max(220, datos.length * 55 + 65);

  const ctx = canvas.getContext('2d');
  ctx.clearRect(0, 0, canvas.width, canvas.height);

  if (!datos.length) {
    ctx.fillStyle = '#888';
    ctx.font = '16px Arial';
    ctx.fillText('Sin datos disponibles', 20, 40);
    return;
  }

  const maximo = Math.max(
    1,
    ...datos.map(d => d.presentes + d.ausentes)
  );

  datos.forEach((grupo, i) => {
    const y = 35 + i * 55;
    const anchoMaximo = 390;

    ctx.fillStyle = '#333';
    ctx.font = '13px Arial';
    ctx.textAlign = 'left';
    ctx.fillText(grupo.nombre.substring(0, 25), 10, y);

    const anchoPresentes =
      (grupo.presentes / maximo) * anchoMaximo;

    const anchoAusentes =
      (grupo.ausentes / maximo) * anchoMaximo;

    ctx.fillStyle = '#7652A8';
    ctx.fillRect(190, y - 15, anchoPresentes, 17);

    ctx.fillStyle = '#E5E7EB';
    ctx.fillRect(
      190 + anchoPresentes,
      y - 15,
      anchoAusentes,
      17
    );

    ctx.fillStyle = '#333';
    ctx.font = '12px Arial';
    ctx.fillText(
      grupo.presentes + ' / ' + grupo.ausentes,
      200 + anchoPresentes + anchoAusentes,
      y
    );
  });
}

function generarGraficosCIE(estudiantes, codigosPresentes) {
  const presentes = estudiantes.filter(estudiante =>
    codigosPresentes.has(String(estudiante.codigo))
  );

  const ausentes = estudiantes.filter(estudiante =>
    !codigosPresentes.has(String(estudiante.codigo))
  );

  dibujarGraficoCircularCIE(presentes.length, ausentes.length);

  function agrupar(campo, usarParalelo = false) {
    const grupos = {};

    estudiantes.forEach(estudiante => {
      const nombre = usarParalelo
        ? `${estudiante[campo] || 'Sin curso'} ${estudiante.paralelo || ''}`.trim()
        : estudiante[campo] || 'Sin especificar';

      if (!grupos[nombre]) {
        grupos[nombre] = { nombre, presentes: 0, ausentes: 0 };
      }

      if (codigosPresentes.has(String(estudiante.codigo))) {
        grupos[nombre].presentes++;
      } else {
        grupos[nombre].ausentes++;
      }
    });

    const ordenCursos = [
  'Primera Sección',
  'Segunda Sección',
  'Primero de Primaria',
  'Segundo de Primaria',
  'Tercero de Primaria',
  'Cuarto de Primaria',
  'Quinto de Primaria',
  'Sexto de Primaria',
  'Primero de Secundaria',
  'Segundo de Secundaria',
  'Tercero de Secundaria',
  'Cuarto de Secundaria',
  'Quinto de Secundaria',
  'Sexto de Secundaria'
];

return Object.values(grupos).sort((a, b) => {
  const posicionA = ordenCursos.findIndex(nombre =>
    a.nombre.toLowerCase().includes(nombre.toLowerCase())
  );

  const posicionB = ordenCursos.findIndex(nombre =>
    b.nombre.toLowerCase().includes(nombre.toLowerCase())
  );

  return (posicionA === -1 ? 999 : posicionA) -
         (posicionB === -1 ? 999 : posicionB);
});
  }

  dibujarBarrasCIE('grafico-niveles', agrupar('nivel'));
  dibujarBarrasCIE('grafico-cursos', agrupar('curso', true));
}


async function mostrarAsistencias() {
  const contenido = document.querySelector('#contenido');

  contenido.innerHTML = `
    <div class="asistencia-vista">

      <header class="asistencia-portada">
      <button id="volver-inicio-asistencia" class="asistencia-btn asistencia-btn-principal">
  ← INICIO
</button>
        <div class="asistencia-portada-texto">
          <span class="asistencia-etiqueta">SHALOM EXPO 2026 · GESTIÓN ACADÉMICA</span>
          <h1>Control de asistencia</h1>
          <p>Supervisa los ingresos y consulta los registros de asistencia estudiantil.</p>
        </div>
        <div class="asistencia-portada-estado">
          <span class="asistencia-punto"></span>
          SISTEMA ACTIVO
        </div>
      </header>

      <section class="asistencia-indicadores">
        <article class="asistencia-indicador indicador-total">
          <div class="indicador-icono">01</div>
          <div>
            <span>TOTAL ESTUDIANTES</span>
            <strong id="total-asistencia">...</strong>
            <small>Estudiantes registrados</small>
          </div>
        </article>

        <article class="asistencia-indicador indicador-presentes">
          <div class="indicador-icono">02</div>
          <div>
            <span>PRESENTES</span>
            <strong id="total-presentes">...</strong>
            <small>Registros del día</small>
          </div>
        </article>

        <article class="asistencia-indicador indicador-ausentes">
          <div class="indicador-icono">03</div>
          <div>
            <span>AUSENTES</span>
            <strong id="total-ausentes">...</strong>
            <small>Sin registro de ingreso</small>
          </div>
        </article>
      </section>

      <section class="asistencia-panel">
        <div class="asistencia-panel-cabecera">
          <div>
            <span class="asistencia-subtitulo">REGISTRO DIARIO</span>
            <h2>Estudiantes presentes</h2>
            <p>Consulta los ingresos registrados durante el día.</p>
          </div>

          <div class="asistencia-acciones">
            <button id="actualizar-asistencias" class="asistencia-btn asistencia-btn-principal">
              <span>↻</span> Actualizar
            </button>
            <button id="limpiar-asistencias" class="asistencia-btn asistencia-btn-peligro">
              Limpiar pruebas
            </button>
          </div>
        </div>

        <div class="asistencia-filtros">
          <div class="asistencia-busqueda">
            <label for="buscar-asistencia">Buscar estudiante</label>
            <input type="text" id="buscar-asistencia"
              placeholder="Nombre o código del estudiante...">
          </div>

          <div>
            <label for="filtro-nivel">Nivel</label>
            <select id="filtro-nivel">
              <option value="">Todos los niveles</option>
            </select>
          </div>

          <div>
            <label for="filtro-curso">Curso</label>
            <select id="filtro-curso">
              <option value="">Todos los cursos</option>
            </select>
          </div>

          <div>
            <label for="filtro-paralelo">Paralelo</label>
            <select id="filtro-paralelo">
              <option value="">Todos los paralelos</option>
            </select>
          </div>
        </div>

        <div class="asistencia-tabla-contenedor">
          <table class="asistencia-tabla">
            <thead>
              <tr>
                <th>N.º</th>
                <th>ESTUDIANTE</th>
                <th>CÓDIGO</th>
                <th>CURSO</th>
                <th>HORA</th>
                <th>ESTADO</th>
              </tr>
            </thead>
            <tbody id="lista-asistencias">
              <tr><td colspan="6">Cargando registros...</td></tr>
            </tbody>
          </table>
        </div>
      </section>

      <section class="asistencia-panel asistencia-reportes">
        <div class="asistencia-panel-cabecera">
          <div>
            <span class="asistencia-subtitulo">DOCUMENTACIÓN</span>
            <h2>Reportes de asistencia</h2>
            <p>Genera informes en PDF o Excel según el periodo que necesites.</p>
          </div>
          <div class="reporte-icono">PDF / XLS</div>
        </div>

        <div class="asistencia-reporte-controles">
          <div class="asistencia-campo-reporte">
            <label for="tipo-reporte">Periodo del reporte</label>
            <select id="tipo-reporte">
              <option value="dia">Asistencia del día</option>
              <option value="mes">Asistencia del mes</option>
              <option value="fechas">Rango de fechas</option>
            </select>
          </div>

          <div class="asistencia-campo-reporte">
            <label for="mes-reporte">Mes</label>
            <input type="month" id="mes-reporte" style="display:none;">
          </div>

          <div class="asistencia-campo-reporte">
            <label for="anio-reporte">Año</label>
            <input type="number" id="anio-reporte"
              value="${new Date().getFullYear()}"
              min="2020" max="2100" style="display:none;">
          </div>

          <div class="asistencia-campo-reporte">
            <label for="fecha-inicio-reporte">Fecha inicial</label>
            <input type="date" id="fecha-inicio-reporte" style="display:none;">
          </div>

          <div class="asistencia-campo-reporte">
            <label for="fecha-fin-reporte">Fecha final</label>
            <input type="date" id="fecha-fin-reporte" style="display:none;">
          </div>
        </div>

        <div class="asistencia-descargas">
          <button id="descargar-pdf" class="asistencia-btn asistencia-btn-pdf">
            Descargar reporte PDF
          </button>
          <button id="descargar-excel" class="asistencia-btn asistencia-btn-excel">
            Descargar reporte Excel
          </button>
        </div>
      </section>

    </div>
  `;
document.querySelector('#volver-inicio-asistencia')
  .addEventListener('click', () => {
    mostrarPanel(administrador?.nombre || 'Usuario');
  });
  document.querySelector('#actualizar-asistencias')
    .addEventListener('click', cargarPanelAsistencias);

  document.querySelector('#limpiar-asistencias')
    .addEventListener('click', limpiarAsistenciasPrueba);

  const tipoReporte = document.querySelector('#tipo-reporte');
  const mesReporte = document.querySelector('#mes-reporte');
  const anioReporte = document.querySelector('#anio-reporte');
  const fechaInicio = document.querySelector('#fecha-inicio-reporte');
  const fechaFin = document.querySelector('#fecha-fin-reporte');

  const hoy = new Date();
  mesReporte.value = `${hoy.getFullYear()}-${String(hoy.getMonth() + 1).padStart(2, '0')}`;

  function actualizarCamposReporte() {
    const tipo = tipoReporte.value;

    mesReporte.style.display = tipo === 'mes' ? 'block' : 'none';
    anioReporte.style.display = tipo === 'mes' ? 'block' : 'none';
    fechaInicio.style.display = tipo === 'fechas' ? 'block' : 'none';
    fechaFin.style.display = tipo === 'fechas' ? 'block' : 'none';
  }

  tipoReporte.addEventListener('change', actualizarCamposReporte);
  actualizarCamposReporte();

  document.querySelector('#descargar-pdf')
    .addEventListener('click', descargarAsistenciaPDF);

  document.querySelector('#descargar-excel')
    .addEventListener('click', descargarAsistenciaExcel);

  await cargarPanelAsistencias();
}

let datosEstudiantesAsistencia = []
let datosPresentesAsistencia = []
let datosTodasAsistencias = []

async function limpiarAsistenciasPrueba() {
  const confirmar = confirm(
    '¿ESTÁS SEGURO?\n\nSe eliminarán TODOS los registros de asistencia de prueba.\n\nLos estudiantes y sus códigos QR NO serán eliminados.'
  )

  if (!confirmar) return

  try {
    const registros = await getDocs(collection(db, 'asistencias'))

    if (registros.empty) {
      alert('No hay registros de asistencia para eliminar.')
      return
    }

    const documentos = registros.docs

    // Firestore permite hasta 500 operaciones por lote.
    for (let i = 0; i < documentos.length; i += 450) {
      const lote = writeBatch(db)

      documentos.slice(i, i + 450).forEach(documento => {
        lote.delete(documento.ref)
      })

      await lote.commit()
    }

    alert(`Limpieza completada. Se eliminaron ${documentos.length} registros.`)

    await cargarPanelAsistencias()

  } catch (error) {
    console.error('Error al limpiar asistencias:', error)
    alert('No se pudieron eliminar los registros. Revisa la consola.')
  }
}

async function cargarPanelAsistencias() {
  const lista = document.querySelector('#lista-asistencias')

  try {
    const ahora = new Date()
    const fecha = [
      ahora.getFullYear(),
      String(ahora.getMonth() + 1).padStart(2, '0'),
      String(ahora.getDate()).padStart(2, '0')
    ].join('-')

    const [estudiantesDB, asistenciasDB] = await Promise.all([
      getDocs(collection(db, 'estudiantes')),
      getDocs(collection(db, 'asistencias'))
    ])

  datosEstudiantesAsistencia = estudiantesDB.docs
  .map(documento => documento.data())
  .filter(estudiante => estudiante.activo !== false)

// Guardar todos los registros para los reportes
datosTodasAsistencias = asistenciasDB.docs
  .map(documento => documento.data())

// Mostrar únicamente la asistencia de hoy
datosPresentesAsistencia = datosTodasAsistencias
  .filter(registro => registro.fecha === fecha)

    const llenarFiltro = (id, campo, texto) => {
      const selector = document.querySelector(id)
      const valorAnterior = selector.value

      const opciones = [...new Set(
        datosEstudiantesAsistencia
          .map(estudiante => estudiante[campo])
          .filter(Boolean)
      )].sort((a, b) => a.localeCompare(b, 'es'))

      selector.innerHTML = `<option value="">${texto}</option>` +
        opciones.map(opcion =>
          `<option value="${escapar(opcion)}">${escapar(opcion)}</option>`
        ).join('')

      selector.value = opciones.includes(valorAnterior) ? valorAnterior : ''
    }

    llenarFiltro('#filtro-nivel', 'nivel', 'Todos los niveles')
    llenarFiltro('#filtro-curso', 'curso', 'Todos los cursos')
    llenarFiltro('#filtro-paralelo', 'paralelo', 'Todos los paralelos')

    document.querySelector('#filtro-nivel').onchange = aplicarFiltrosAsistencia
    document.querySelector('#filtro-curso').onchange = aplicarFiltrosAsistencia
    document.querySelector('#filtro-paralelo').onchange = aplicarFiltrosAsistencia
    document.querySelector('#buscar-asistencia').oninput = aplicarFiltrosAsistencia

    aplicarFiltrosAsistencia()

  } catch (error) {
    console.error('Error al cargar asistencias:', error)
    lista.innerHTML = '<tr><td colspan="6">Error al consultar las asistencias.</td></tr>'
  }
}

function aplicarFiltrosAsistencia() {
  const nivel = document.querySelector('#filtro-nivel').value
  const curso = document.querySelector('#filtro-curso').value
  const paralelo = document.querySelector('#filtro-paralelo').value
  const busqueda = document.querySelector('#buscar-asistencia').value
    .trim().toLowerCase()

  const coincide = estudiante => {
    const nombre = `${estudiante.nombres || ''} ${estudiante.apellidos || ''}`.toLowerCase()
    const codigo = String(estudiante.codigo || '').toLowerCase()

    return (!nivel || estudiante.nivel === nivel) &&
      (!curso || estudiante.curso === curso) &&
      (!paralelo || estudiante.paralelo === paralelo) &&
      (!busqueda || nombre.includes(busqueda) || codigo.includes(busqueda))
  }

  const estudiantesFiltrados = datosEstudiantesAsistencia.filter(coincide)
  const presentes = datosPresentesAsistencia
    .filter(registro =>
      registro.estado === 'PRESENTE' &&
      coincide(registro)
    )

  document.querySelector('#total-asistencia').textContent = estudiantesFiltrados.length
  document.querySelector('#total-presentes').textContent = presentes.length
  document.querySelector('#total-ausentes').textContent =
    Math.max(0, estudiantesFiltrados.length - presentes.length)

  presentes.sort((a, b) => (a.hora || '').localeCompare(b.hora || ''))

  const lista = document.querySelector('#lista-asistencias')

  if (presentes.length === 0) {
    lista.innerHTML = '<tr><td colspan="6">No hay estudiantes presentes con estos filtros.</td></tr>'
    return
  }

  lista.innerHTML = presentes.map((estudiante, indice) => `
    <tr>
      <td class="numero-fila">${String(indice + 1).padStart(2, '0')}</td>
      <td>
        <div class="nombre-estudiante">
          <strong>${escapar(estudiante.nombres || '')} ${escapar(estudiante.apellidos || '')}</strong>
          <small>Estudiante registrado</small>
        </div>
      </td>
      <td><span class="codigo-estudiante">${escapar(estudiante.codigo || '')}</span></td>
      <td><strong>${escapar(estudiante.curso || '')} ${escapar(estudiante.paralelo || '')}</strong></td>
      <td>${escapar(estudiante.hora || '')}</td>
      <td><span class="estado-asistencia"><span>●</span> PRESENTE</span></td>
    </tr>
  `).join('')
}

// Obtener registros según el periodo y los filtros seleccionados
function obtenerRegistrosReporte() {
  const tipo = document.querySelector('#tipo-reporte').value
  const ahora = new Date()

  const fechaActual = [
    ahora.getFullYear(),
    String(ahora.getMonth() + 1).padStart(2, '0'),
    String(ahora.getDate()).padStart(2, '0')
  ].join('-')

  let registros = [...datosTodasAsistencias]

  if (tipo === 'hoy') {
    registros = registros.filter(r => r.fecha === fechaActual)
  }

  if (tipo === 'mes') {
    const mes = document.querySelector('#mes-reporte').value
    const anio = document.querySelector('#anio-reporte').value

    if (!mes || !anio || Number(anio) < 2020) {
      alert('Selecciona un mes y un año válidos.')
      return null
    }

    const periodo = `${anio}-${mes}`
    registros = registros.filter(r =>
      String(r.fecha || '').startsWith(periodo)
    )
  }

  if (tipo === 'fechas') {
    const inicio = document.querySelector('#fecha-inicio-reporte').value
    const fin = document.querySelector('#fecha-fin-reporte').value

    if (!inicio || !fin) {
      alert('Selecciona la fecha inicial y final.')
      return null
    }

    if (inicio > fin) {
      alert('La fecha inicial no puede ser posterior a la fecha final.')
      return null
    }

    registros = registros.filter(r =>
      r.fecha >= inicio && r.fecha <= fin
    )
  }

  // Aplicar los filtros de nivel, curso, paralelo y búsqueda
  const nivel = document.querySelector('#filtro-nivel').value
  const curso = document.querySelector('#filtro-curso').value
  const paralelo = document.querySelector('#filtro-paralelo').value
  const busqueda = document.querySelector('#buscar-asistencia').value
    .trim().toLowerCase()

  registros = registros.filter(r => {
    const nombre = `${r.nombres || ''} ${r.apellidos || ''}`.toLowerCase()
    const codigo = String(r.codigo || '').toLowerCase()

    return r.estado === 'PRESENTE' &&
      (!nivel || r.nivel === nivel) &&
      (!curso || r.curso === curso) &&
      (!paralelo || r.paralelo === paralelo) &&
      (!busqueda || nombre.includes(busqueda) || codigo.includes(busqueda))
  })

  return registros.sort((a, b) =>
    `${a.fecha || ''} ${a.hora || ''}`.localeCompare(
      `${b.fecha || ''} ${b.hora || ''}`
    )
  )
}


// DESCARGAR REPORTE PDF
function descargarAsistenciaPDF() {
  const registros = obtenerRegistrosReporte()

  if (registros === null) return

  if (registros.length === 0) {
    alert('No existen registros para descargar con los filtros seleccionados.')
    return
  }

  const pdf = new jsPDF({ orientation: 'landscape' })

  pdf.setFontSize(17)
  pdf.setFont('helvetica', 'bold')
  pdf.text('COLEGIO ADVENTISTA SHALOM', 14, 16)

  pdf.setFontSize(13)
  pdf.text('REPORTE DE CONTROL DE ASISTENCIA', 14, 24)

  pdf.setFontSize(9)
  pdf.setFont('helvetica', 'normal')
  pdf.text(`Fecha de generación: ${new Date().toLocaleDateString('es-BO')}`, 14, 31)
  pdf.text(`Total de registros: ${registros.length}`, 240, 31)

  const columnas = [
    ['N.º', 14],
    ['FECHA', 25],
    ['HORA', 53],
    ['CÓDIGO', 78],
    ['ESTUDIANTE', 108],
    ['CURSO / PARALELO', 200],
    ['ESTADO', 260]
  ]

  let y = 40

  const dibujarEncabezado = () => {
    pdf.setFillColor(35, 48, 75)
    pdf.rect(14, y - 6, 269, 9, 'F')
    pdf.setTextColor(255, 255, 255)
    pdf.setFont('helvetica', 'bold')
    pdf.setFontSize(8)

    columnas.forEach(([titulo, x]) => pdf.text(titulo, x, y))
    pdf.setTextColor(0, 0, 0)
    pdf.setFont('helvetica', 'normal')
  }

  dibujarEncabezado()
  y += 9

  registros.forEach((r, i) => {
    if (y > 190) {
      pdf.addPage()
      y = 18
      dibujarEncabezado()
      y += 9
    }

    const nombre = `${r.nombres || ''} ${r.apellidos || ''}`.trim()
    const curso = `${r.curso || ''} ${r.paralelo || ''}`.trim()

    pdf.setFontSize(8)
    pdf.text(String(i + 1), 14, y)
    pdf.text(String(r.fecha || ''), 25, y)
    pdf.text(String(r.hora || ''), 53, y)
    pdf.text(String(r.codigo || ''), 78, y)
    pdf.text(nombre.substring(0, 45), 108, y)
    pdf.text(curso.substring(0, 27), 200, y)
    pdf.text('PRESENTE', 260, y)

    pdf.setDrawColor(220, 220, 220)
    pdf.line(14, y + 2, 283, y + 2)

    y += 8
  })

  pdf.save('Reporte_Asistencia_Shalom.pdf')
}

// DESCARGAR REPORTE EXCEL
function descargarAsistenciaExcel() {
  const registros = obtenerRegistrosReporte()

  if (registros === null) return

  if (registros.length === 0) {
    alert('No existen registros para descargar con los filtros seleccionados.')
    return
  }

  const datos = registros.map((r, i) => ({
    'N.º': i + 1,
    'FECHA': r.fecha || '',
    'HORA': r.hora || '',
    'CÓDIGO': r.codigo || '',
    'APELLIDOS': r.apellidos || '',
    'NOMBRES': r.nombres || '',
    'NIVEL': r.nivel || '',
    'CURSO': r.curso || '',
    'PARALELO': r.paralelo || '',
    'ESTADO': r.estado || ''
  }))

  const hoja = XLSX.utils.json_to_sheet(datos)
  hoja['!cols'] = [
    { wch: 7 },
    { wch: 14 },
    { wch: 12 },
    { wch: 14 },
    { wch: 25 },
    { wch: 25 },
    { wch: 18 },
    { wch: 18 },
    { wch: 12 },
    { wch: 15 }
  ]

  const libro = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(libro, hoja, 'Asistencia')

  XLSX.writeFile(libro, 'Reporte_Asistencia_Shalom.xlsx')
}

async function actualizarTotalEstudiantesPanel() {
  const elemento = document.querySelector('#total-estudiantes-panel')

  if (!elemento) return

  try {
    const resultado = await getDocs(collection(db, 'estudiantes'))

const correcciones = resultado.docs.filter(documento =>
  String(documento.data().curso || '').trim().toUpperCase() === 'SEXO'
)

if (correcciones.length > 0) {
  const lote = writeBatch(db)

  correcciones.forEach(documento => {
    lote.update(documento.ref, { curso: 'SEXTO' })
  })

  await lote.commit()

  console.log(`Se corrigieron ${correcciones.length} cursos de SEXO a SEXTO.`)
}

    elemento.textContent = resultado.size
  } catch (error) {
    console.error('Error al contar estudiantes:', error)
    elemento.textContent = '—'
  }
}


async function cargarOpcionesCredenciales() {
  try {
    const resultado = await getDocs(collection(db, 'estudiantes'))

    const cursos = new Map()

    resultado.docs.forEach(documento => {
      const estudiante = documento.data()

      if (!estudiante.nivel || !estudiante.curso) return

      const nivel = estudiante.nivel.trim()
      const curso = estudiante.curso.trim()
      const clave = JSON.stringify({ nivel, curso })

      cursos.set(clave, { nivel, curso })
    })

    const selector = document.querySelector('#filtro-curso')

    selector.innerHTML = '<option value="">Seleccionar curso</option>'

    Array.from(cursos.entries())
      .sort((a, b) =>
        `${a[1].nivel} ${a[1].curso}`.localeCompare(
          `${b[1].nivel} ${b[1].curso}`,
          'es',
          { numeric: true }
        )
      )
      .forEach(([clave, datos]) => {
        const opcion = document.createElement('option')
        opcion.value = clave
        opcion.textContent = `${datos.nivel} - ${datos.curso}`
        selector.appendChild(opcion)
      })

  } catch (error) {
    console.error('Error al cargar cursos:', error)
    alert('No se pudieron cargar los cursos.')
  }
}

async function cargarParalelosCredenciales() {
  try {
    const selectorCurso = document.querySelector('#filtro-curso')
    const selectorParalelo = document.querySelector('#filtro-paralelo')

    selectorParalelo.innerHTML =
      '<option value="">Seleccionar paralelo</option>'

    if (!selectorCurso.value) return

    const seleccionado = JSON.parse(selectorCurso.value)

    const resultado = await getDocs(collection(db, 'estudiantes'))

    const paralelos = new Set()

    resultado.docs.forEach(documento => {
      const estudiante = documento.data()

      if (
        estudiante.nivel?.trim() === seleccionado.nivel &&
        estudiante.curso?.trim() === seleccionado.curso &&
        estudiante.paralelo
      ) {
        paralelos.add(estudiante.paralelo.trim())
      }
    })

    Array.from(paralelos)
      .sort((a, b) => a.localeCompare(b, 'es', { numeric: true }))
      .forEach(paralelo => {
        const opcion = document.createElement('option')
        opcion.value = paralelo
        opcion.textContent = `Paralelo ${paralelo}`
        selectorParalelo.appendChild(opcion)
      })

  } catch (error) {
    console.error('Error al cargar paralelos:', error)
    alert('No se pudieron cargar los paralelos.')
  }
}

async function generarCredencialesGrupo() {
  const cursoSeleccionado = document.querySelector('#filtro-curso').value
  const paraleloSeleccionado = document.querySelector('#filtro-paralelo').value
  const mensaje = document.querySelector('#mensaje-credenciales')

  if (!cursoSeleccionado) {
    alert('Primero selecciona un curso.')
    return
  }

  if (!paraleloSeleccionado) {
    alert('Primero selecciona un paralelo.')
    return
  }

  const datos = JSON.parse(cursoSeleccionado)

  await generarCredenciales({
    nivel: datos.nivel,
    curso: datos.curso,
    paralelo: paraleloSeleccionado
  })
}

async function generarCredenciales(filtros = null) {
  if (!filtros) {
    alert('Selecciona un curso y paralelo para generar las credenciales.')
    return
  }
  try {
    const resultado = await getDocs(collection(db, 'estudiantes'))

    const estudiantes = resultado.docs
  .map(documento => ({
    id: documento.id,
    ...documento.data()
  }))
  .filter(estudiante =>
  estudiante.qr_id &&
  estudiante.activo !== false &&
  (
    !filtros ||
    (
      estudiante.nivel?.trim() === filtros.nivel &&
      estudiante.curso?.trim() === filtros.curso &&
      estudiante.paralelo?.trim() === filtros.paralelo
    )
  )
)
    if (estudiantes.length === 0) {
      alert('No hay estudiantes activos con identificador QR.')
      return
    }

    const confirmar = confirm(
      `Se generarán ${estudiantes.length} credenciales en PDF. ¿Deseas continuar?`
    )

    if (!confirmar) return

    const cargarImagen = async ruta => {
      const respuesta = await fetch(ruta)

      if (!respuesta.ok) {
        throw new Error('No se encontró la imagen: ' + ruta)
      }

      const blob = await respuesta.blob()

      return await new Promise((resolve, reject) => {
        const lector = new FileReader()
        lector.onload = () => resolve(lector.result)
        lector.onerror = reject
        lector.readAsDataURL(blob)
      })
    }
const frente = await cargarImagen('/anverso.png')
const reverso = await cargarImagen('/reverso.png')

const pdf = new jsPDF({
  orientation: 'landscape',
  unit: 'mm',
  format: 'letter'
})

// MEDIDAS EXACTAS: 10,40 x 6,70 CM
const ancho = 104
const alto = 67

// SEIS CREDENCIALES POR HOJA
const posiciones = [
  [34.7, 5.45],
  [140.7, 5.45],
  [34.7, 74.45],
  [140.7, 74.45],
  [34.7, 143.45],
  [140.7, 143.45]
]

// GENERAR CREDENCIALES EN GRUPOS DE 6
for (let inicio = 0; inicio < estudiantes.length; inicio += 6) {

  const grupo = estudiantes.slice(inicio, inicio + 6)

  // ==========================================
  // PÁGINA DEL ANVERSO
  // ==========================================

  if (inicio > 0) {
    pdf.addPage('letter', 'landscape')
  }

  for (let i = 0; i < grupo.length; i++) {

    const estudiante = grupo[i]
    const [x, y] = posiciones[i]

    // FONDO ANVERSO
    pdf.addImage(
      frente,
      'PNG',
      x,
      y,
      ancho,
      alto
    )


 // NOMBRES
const nombres = (estudiante.nombres || '').toUpperCase()
const apellidos = (estudiante.apellidos || '').toUpperCase()

pdf.setFont('helvetica', 'bold')
pdf.setFontSize(14)
pdf.setTextColor(255, 255, 255)

pdf.text(nombres, x + 72, y + 35, {
  align: 'center',
  maxWidth: 51
})

// APELLIDOS
pdf.setFont('helvetica', 'normal')
pdf.setFontSize(10)
pdf.setTextColor(55, 55, 55)

pdf.text(apellidos, x + 72, y + 42, {
  align: 'center',
  maxWidth: 51
})

// CURSO Y PARALELO
const curso = (
  `${estudiante.curso || ''} ${estudiante.paralelo || ''}`
).trim().toUpperCase()

pdf.setFont('helvetica', 'bold')
pdf.setFontSize(9)
pdf.setTextColor(255, 193, 7)

pdf.text(curso, x + 72, y + 49, {
  align: 'center',
  maxWidth: 51
})

// CÓDIGO
pdf.setFont('helvetica', 'normal')
pdf.setFontSize(8)
pdf.setTextColor(255, 255, 255)

pdf.text(
  `CÓDIGO: ${estudiante.codigo || ''}`,
  x + 72,
  y + 55,
  { align: 'center' }
)
    // QR DEL ANVERSO
    const qr = await QRCode.toDataURL(
      estudiante.qr_id,
      {
        width: 500,
        margin: 2,
        errorCorrectionLevel: 'H'
      }
    )

    pdf.addImage(
      qr,
      'PNG',
      x + 81,
      y + 7,
      20,
      20
    )
  }

  // ==========================================
  // PÁGINA DEL REVERSO
  // ==========================================

  pdf.addPage('letter', 'landscape')

  for (let i = 0; i < grupo.length; i++) {

    const estudiante = grupo[i]
    const [x, y] = posiciones[i]

    // FONDO COMPLETO DEL REVERSO
    pdf.addImage(
      reverso,
      'PNG',
      x,
      y,
      ancho,
      alto
    )

    // MISMO QR DEL ESTUDIANTE
    const qr = await QRCode.toDataURL(
      estudiante.qr_id,
      {
        width: 500,
        margin: 2,
        errorCorrectionLevel: 'H'
      }
    )
  
// FONDO BLANCO
pdf.roundedRect(
  x + 44,
  y + 44,
  15,
  15,
  1,
  1,
  'F'
)

// QR TRASERO
pdf.addImage(
  qr,
  'PNG',
  x + 44,
  y + 44,
  15,
  15
)
  }
}

// DESCARGAR PDF
pdf.save('credenciales-completas-shalom.pdf')

alert(
  `¡Credenciales generadas!\n` +
  `Total: ${estudiantes.length}\n` +
  `Hojas de impresión: ${Math.ceil(estudiantes.length / 6)} por cada cara.`
)
  } catch (error) {
    console.error('Error al generar credenciales:', error)
    alert('No se pudo generar el PDF. Revisa la consola.')
  }
}

async function generarQRParaEstudiantesExistentes() {
  const confirmar = confirm(
    'Se asignará un identificador QR único a cada estudiante que todavía no tenga uno. ¿Deseas continuar?'
  )
  if (!confirmar) return

  try {
    const resultado = await getDocs(collection(db, 'estudiantes'))

    let lote = writeBatch(db)
    let cantidadLote = 0
    let generados = 0

    for (const documento of resultado.docs) {
      const datos = documento.data()

      if (!datos.qr_id) {
        lote.update(documento.ref, {
          qr_id: crypto.randomUUID(),
          qr_generado: true
        })

        cantidadLote++
        generados++

        if (cantidadLote === 400) {
          await lote.commit()
          lote = writeBatch(db)
          cantidadLote = 0
        }
      }
    }

    if (cantidadLote > 0) {
      await lote.commit()
    }

    alert(`Proceso terminado. Se asignaron ${generados} identificadores QR.`)

  } catch (error) {
    console.error('Error al generar identificadores:', error)
    alert('No se pudieron asignar los identificadores. Revisa la consola.')
  }
}

let lectorQR = null


async function mostrarEscanerQR() {
  const contenido = document.querySelector('#contenido')

  contenido.innerHTML = `
    <main class="asistencia-qr">

      <header class="asistencia-qr-header">
        <div class="asistencia-qr-identidad">
          <span>SHALOM EXPO 2026</span>
          <h1>Control de asistencia</h1>
          <p>Registro de estudiantes mediante códigos QR</p>
        </div>

        <div class="asistencia-qr-estado">
          <span class="asistencia-punto"></span>
          SISTEMA DE ESCANEO
        </div>
      </header>

      <section class="asistencia-qr-resumen">
        <article class="asistencia-qr-indicador">
          <span>ESCANEADOS</span>
          <strong id="qr-total-escaneados">0</strong>
        </article>

        <article class="asistencia-qr-indicador">
          <span>PRESENTES</span>
          <strong id="qr-total-presentes">0</strong>
        </article>

        <article class="asistencia-qr-indicador">
          <span>DUPLICADOS</span>
          <strong id="qr-total-duplicados">0</strong>
        </article>
      </section>

      <section class="asistencia-qr-contenido">

        <div class="asistencia-qr-camara">
          <div class="asistencia-qr-titulo">
            <div>
              <span>LECTURA AUTOMÁTICA</span>
              <h2>Escáner de códigos QR</h2>
              <p>Ubica el código QR de la credencial frente a la cámara.</p>
            </div>
            <span class="asistencia-qr-icono">⌗</span>
          </div>

          <div class="asistencia-qr-visor">
            <div id="lector-qr"></div>
          </div>

          <div class="asistencia-qr-instruccion">
            <span>●</span>
            Mantén el código QR dentro del área de lectura.
          </div>

          <div id="resultado-qr" class="asistencia-qr-resultado">
            <span class="asistencia-resultado-icono">⌁</span>
            <div>
              <strong>Esperando lectura</strong>
              <p>El resultado aparecerá aquí cuando se escanee una credencial.</p>
            </div>
          </div>

          <button id="detener-lector" class="asistencia-qr-detener">
            DETENER CÁMARA
          </button>
        </div>

        <aside class="asistencia-qr-lateral">
          <div class="asistencia-qr-ayuda">
            <span>INSTRUCCIONES</span>
            <h3>¿Cómo registrar asistencia?</h3>
            <ol>
              <li>Activa la cámara del dispositivo.</li>
              <li>Solicita al estudiante mostrar su credencial.</li>
              <li>Centra el código QR frente a la cámara.</li>
              <li>Espera la confirmación del registro.</li>
            </ol>
          </div>

          <div class="asistencia-qr-seguridad">
            <strong>Control de registros</strong>
            <p>El sistema verifica si el estudiante ya registró su asistencia para evitar duplicaciones.</p>
          </div>

          <button id="volver-inicio" class="asistencia-qr-volver">
            ← VOLVER AL INICIO
          </button>
        </aside>

      </section>
    </main>
  `

  document.querySelector('#volver-inicio').addEventListener('click', async () => {
    await detenerEscanerQR()
    mostrarPanel(administrador.nombre || 'Administrador')
  })

  document.querySelector('#detener-lector')
    .addEventListener('click', detenerEscanerQR)

  lectorQR = new Html5Qrcode('lector-qr')

  try {
    await lectorQR.start(
      { facingMode: 'environment' },
      { fps: 10, qrbox: { width: 250, height: 250 } },
      buscarEstudianteQR,
      () => {}
    )
  } catch (error) {
    console.error(error)
    document.querySelector('#resultado-qr').innerHTML =
      '<p>No se pudo activar la cámara. Verifica los permisos del navegador.</p>'
  }
}

async function detenerEscanerQR() {
  if (lectorQR && lectorQR.isScanning) {
    await lectorQR.stop()
    await lectorQR.clear()
  }

  lectorQR = null
}

let procesandoQR = false

async function buscarEstudianteQR(textoQR) {
  if (procesandoQR) return

  procesandoQR = true

  const resultadoQR = document.querySelector('#resultado-qr')

  try {
    const consulta = query(
      collection(db, 'estudiantes'),
      where('qr_id', '==', textoQR),
      limit(1)
    )

    const resultado = await getDocs(consulta)

    if (resultado.empty) {
      resultadoQR.innerHTML = '<h3>ESTUDIANTE NO ENCONTRADO</h3>'
      return
    }

    const documento = resultado.docs[0]
    const estudiante = documento.data()

    if (estudiante.activo === false) {
      resultadoQR.innerHTML = '<h3>ESTUDIANTE INACTIVO</h3>'
      return
    }

    const ahora = new Date()
    const fecha = [
      ahora.getFullYear(),
      String(ahora.getMonth() + 1).padStart(2, '0'),
      String(ahora.getDate()).padStart(2, '0')
    ].join('-')

    const hora = ahora.toLocaleTimeString('es-BO')

    const asistenciaRef = doc(
      db,
      'asistencias',
      `${documento.id}_${fecha}`
    )

    const asistenciaExistente = await getDoc(asistenciaRef)

    if (asistenciaExistente.exists()) {
      resultadoQR.innerHTML = `
        <h2 style="color:#d97706">ASISTENCIA YA REGISTRADA</h2>
        <h3>${escapar(estudiante.nombres || '')} ${escapar(estudiante.apellidos || '')}</h3>
        <p>Primera entrada: ${escapar(asistenciaExistente.data().hora || '')}</p>
      `
      return
    }

    await setDoc(asistenciaRef, {
      estudiante_id: documento.id,
      codigo: estudiante.codigo || '',
      nombres: estudiante.nombres || '',
      apellidos: estudiante.apellidos || '',
      nivel: estudiante.nivel || '',
      curso: estudiante.curso || '',
      paralelo: estudiante.paralelo || '',
      fecha: fecha,
      hora: hora,
      fecha_hora: ahora,
      estado: 'PRESENTE'
    })

    resultadoQR.innerHTML = `
      <div class="form-estudiante">
        <h2 style="color:#15803d">ASISTENCIA REGISTRADA</h2>
        <h3>${escapar(estudiante.nombres || '')} ${escapar(estudiante.apellidos || '')}</h3>
        <p><strong>Código:</strong> ${escapar(estudiante.codigo || '')}</p>
        <p><strong>Curso:</strong> ${escapar(estudiante.curso || '')} ${escapar(estudiante.paralelo || '')}</p>
        <p><strong>Fecha:</strong> ${fecha}</p>
        <p><strong>Hora:</strong> ${hora}</p>
        <span class="estado">● PRESENTE</span>
      </div>
    `

  } catch (error) {
    console.error('Error al registrar asistencia:', error)
    resultadoQR.textContent = 'Error al consultar o guardar la asistencia en Firebase.'
  } finally {
    setTimeout(() => {
      procesandoQR = false
    }, 2000)
  }
}

async function mostrarEstudiantes() {
  const contenido = document.querySelector('#contenido');

  contenido.innerHTML = `
    <div class="estudiantes-vista">

      <!-- ENCABEZADO -->
      <header class="est-encabezado">
        <div class="est-encabezado-texto">
          <span class="est-etiqueta">SHALOM EXPO 2026 · GESTIÓN ACADÉMICA</span>
          <h1>Administración de estudiantes</h1>
          <p>Registro, organización y control de los participantes de la exposición.</p>
        </div>

        <div class="est-encabezado-acciones">
          <span class="est-conexion">
            <span></span> SISTEMA ACTIVO
          </span>
          <button id="volver-inicio" class="est-btn est-btn-claro">
            ← Volver al inicio
          </button>
        </div>
      </header>

      <!-- RESUMEN -->
      <section class="est-resumen">
        <article class="est-resumen-card est-resumen-principal">
          <div class="est-resumen-icono">01</div>
          <div>
            <span>GESTIÓN INSTITUCIONAL</span>
            <h2>Directorio estudiantil</h2>
            <p>Administra los registros y las credenciales de los estudiantes.</p>
          </div>
        </article>

        <article class="est-resumen-card">
          <div class="est-resumen-icono est-icono-lila">02</div>
          <div>
            <span>BASE DE DATOS</span>
            <h3 id="total-estudiantes">Cargando registros...</h3>
            <p>Estudiantes registrados</p>
          </div>
        </article>
      </section>

      <!-- REGISTRO MANUAL -->
      <section class="est-card">
        <div class="est-card-encabezado">
          <div class="est-numero">01</div>
          <div>
            <h2>Registrar estudiante</h2>
            <p>Completa los datos académicos del nuevo participante.</p>
          </div>
        </div>

        <form id="form-estudiante" class="est-formulario">
          <div class="est-campos">
            <div class="est-campo">
              <label for="codigo">Código del estudiante</label>
              <input id="codigo" placeholder="Ej. SH-0002" required>
            </div>

            <div class="est-campo">
              <label for="nombres">Nombres</label>
              <input id="nombres" placeholder="Nombres del estudiante" required>
            </div>

            <div class="est-campo">
              <label for="apellidos">Apellidos</label>
              <input id="apellidos" placeholder="Apellidos del estudiante" required>
            </div>

            <div class="est-campo">
              <label for="nivel">Nivel</label>
              <select id="nivel" required>
                <option value="">Seleccionar nivel</option>
                <option value="Inicial">Inicial</option>
                <option value="Primaria">Primaria</option>
                <option value="Secundaria">Secundaria</option>
              </select>
            </div>

            <div class="est-campo">
              <label for="curso">Curso</label>
              <input id="curso" placeholder="Ej. 1ro" required>
            </div>

            <div class="est-campo">
              <label for="paralelo">Paralelo</label>
              <select id="paralelo" required>
                <option value="">Seleccionar paralelo</option>
                <option value="A">A</option>
                <option value="B">B</option>
                <option value="C">C</option>
                <option value="D">D</option>
              </select>
            </div>
          </div>

          <div class="est-form-pie">
            <p id="mensaje-estudiante" role="status"></p>
            <button type="submit" id="guardar-estudiante" class="est-btn est-btn-primario">
              GUARDAR ESTUDIANTE
            </button>
          </div>
        </form>
      </section>

      <!-- HERRAMIENTAS -->
      <section class="est-herramientas">

        <article class="est-card est-herramienta">
          <div class="est-card-encabezado">
            <div class="est-numero est-numero-naranja">02</div>
            <div>
              <h2>Importación desde Excel</h2>
              <p>Prepara la carga masiva de estudiantes desde una hoja de cálculo.</p>
            </div>
          </div>

          <div class="est-archivo">
            <label for="archivo-excel">Seleccionar archivo Excel</label>
            <input type="file" id="archivo-excel" accept=".xlsx,.xls">
            <span>Formatos permitidos: .xlsx y .xls</span>
          </div>

          <button type="button" id="previsualizar-excel" class="est-btn est-btn-secundario">
            REVISAR ARCHIVO
          </button>

          <p id="mensaje-excel" role="status"></p>
          <div id="vista-excel" class="tabla-scroll est-vista-excel"></div>
        </article>

        <article class="est-card est-herramienta">
          <div class="est-card-encabezado">
            <div class="est-numero est-numero-azul">03</div>
            <div>
              <h2>Credenciales por curso</h2>
              <p>Genera archivos PDF independientes para cada grupo.</p>
            </div>
          </div>

          <div class="est-campos est-campos-dos">
            <div class="est-campo">
              <label for="filtro-curso">Curso</label>
              <select id="filtro-curso">
                <option value="">Cargando cursos...</option>
              </select>
            </div>

            <div class="est-campo">
              <label for="filtro-paralelo">Paralelo</label>
              <select id="filtro-paralelo">
                <option value="">Seleccionar paralelo</option>
              </select>
            </div>
          </div>

          <button type="button" id="generar-credenciales-grupo" class="est-btn est-btn-primario">
            GENERAR PDF DEL GRUPO
          </button>

          <p id="mensaje-credenciales" role="status"></p>
        </article>
      </section>

      <!-- DIRECTORIO -->
      <section class="est-card est-directorio">
        <div class="est-card-encabezado est-directorio-encabezado">
          <div class="est-numero est-numero-verde">04</div>
          <div>
            <h2>Directorio de estudiantes</h2>
            <p>Consulta, filtra y administra los registros institucionales.</p>
          </div>
        </div>

        <div class="est-filtros">
          <div class="est-buscador">
            <label for="buscar-estudiante">BUSCAR ESTUDIANTE</label>
            <input
              type="text"
              id="buscar-estudiante"
              placeholder="Nombre, apellido o código..."
            >
          </div>

          <div class="est-campo">
            <label for="filtro-nivel-estudiantes">NIVEL</label>
            <select id="filtro-nivel-estudiantes">
              <option value="">Todos los niveles</option>
            </select>
          </div>

          <div class="est-campo">
            <label for="filtro-curso-estudiantes">CURSO</label>
            <select id="filtro-curso-estudiantes">
              <option value="">Todos los cursos</option>
            </select>
          </div>

          <div class="est-campo">
            <label for="filtro-paralelo-estudiantes">PARALELO</label>
            <select id="filtro-paralelo-estudiantes">
              <option value="">Todos los paralelos</option>
            </select>
          </div>
        </div>

        <!-- ACCIONES -->
        <div class="est-acciones">
          <div class="est-acciones-grupo">
            <button id="actualizar-estudiantes" class="est-btn est-btn-secundario">
              ↻ Actualizar
            </button>
            <button type="button" id="generar-qr-estudiantes" class="est-btn est-btn-secundario">
              GENERAR QR PARA TODOS
            </button>
            <button type="button" id="generar-credenciales" class="est-btn est-btn-secundario">
              GENERAR CREDENCIALES
            </button>
            <button type="button" id="respaldo-estudiantes" class="est-btn est-btn-secundario">
              DESCARGAR RESPALDO
            </button>
          </div>

          <button type="button" id="limpiar-estudiantes" class="est-btn est-btn-peligro">
            ELIMINAR TODOS
          </button>
        </div>

        <!-- TABLA -->
        <div class="est-tabla-contenedor">
          <table class="est-tabla">
            <thead>
              <tr>
                <th>CÓDIGO</th>
                <th>ESTUDIANTE</th>
                <th>NIVEL</th>
                <th>CURSO</th>
                <th>PARALELO</th>
                <th>ESTADO</th>
              </tr>
            </thead>
            <tbody id="lista-estudiantes">
              <tr><td colspan="6">Cargando estudiantes...</td></tr>
            </tbody>
          </table>
        </div>
      </section>

      <footer class="est-pie">
        SHALOM EXPO 2026 · SISTEMA DE GESTIÓN INSTITUCIONAL
      </footer>

    </div>
  `;

  document.querySelector('#volver-inicio').addEventListener('click', () => {
    mostrarPanel(administrador.nombre || 'Administrador');
  });

  document.querySelector('#actualizar-estudiantes')
    .addEventListener('click', cargarEstudiantes);

  document.querySelector('#respaldo-estudiantes')
    .addEventListener('click', descargarCopiaEstudiantes);

  document.querySelector('#generar-qr-estudiantes')
    .addEventListener('click', generarQRParaEstudiantesExistentes);

  document.querySelector('#generar-credenciales')
    .addEventListener('click', () => generarCredenciales());

  document.querySelector('#limpiar-estudiantes')
    .addEventListener('click', eliminarTodosEstudiantes);

  document.querySelector('#previsualizar-excel')
    .addEventListener('click', revisarExcel);

  document.querySelector('#form-estudiante')
    .addEventListener('submit', guardarEstudiante);

  document.querySelector('#filtro-curso')
    .addEventListener('change', cargarParalelosCredenciales);

  document.querySelector('#generar-credenciales-grupo')
    .addEventListener('click', generarCredencialesGrupo);

  await cargarOpcionesCredenciales();
  await cargarEstudiantes();
}

async function eliminarEstudiante(id, codigo) {
  const confirmar = confirm(
    `¿Estás seguro de eliminar al estudiante ${codigo}? Esta acción no se puede deshacer.`
  )

  if (!confirmar) return

  try {
    await deleteDoc(doc(db, 'estudiantes', id))

    alert('Estudiante eliminado correctamente.')

    await cargarEstudiantes()

  } catch (error) {
    console.error('Error al eliminar:', error)
    alert('No se pudo eliminar el estudiante.')
  }
}
async function descargarCopiaEstudiantes() {
  try {
    const resultado = await getDocs(collection(db, 'estudiantes'))

    const respaldo = resultado.docs.map(documento => ({
      id: documento.id,
      ...documento.data()
    }))

    const archivo = new Blob(
      [JSON.stringify(respaldo, null, 2)],
      { type: 'application/json' }
    )

    const enlace = document.createElement('a')
    const url = URL.createObjectURL(archivo)

    enlace.href = url
    enlace.download = 'respaldo_estudiantes_shalom.json'
    enlace.click()

    URL.revokeObjectURL(url)

    alert(`Copia de seguridad descargada: ${respaldo.length} estudiantes.`)

  } catch (error) {
    console.error(error)
    alert('No se pudo crear la copia de seguridad.')
  }
}
async function eliminarTodosEstudiantes() {
  const confirmar = confirm(
    '¿ESTÁS SEGURO? Se eliminarán TODOS los estudiantes de Firebase. Esta acción no se puede deshacer.'
  )

  if (!confirmar) return

  const segundaConfirmacion = confirm(
    'CONFIRMACIÓN FINAL: ¿Ya descargaste la copia de seguridad y deseas continuar?'
  )

  if (!segundaConfirmacion) return

  const boton = document.querySelector('#limpiar-estudiantes')

  try {
    boton.disabled = true
    boton.textContent = 'ELIMINANDO...'

    const referencia = collection(db, 'estudiantes')
    let eliminados = 0

    while (true) {
      const resultado = await getDocs(query(referencia, limit(450)))

      if (resultado.empty) break

      const lote = writeBatch(db)

      resultado.docs.forEach(documento => {
        lote.delete(documento.ref)
      })

      await lote.commit()

      eliminados += resultado.size
    }

    alert(`Se eliminaron ${eliminados} estudiantes correctamente.`)

    await cargarEstudiantes()

  } catch (error) {
    console.error('Error al eliminar estudiantes:', error)
    alert('Ocurrió un error. Revisa la consola y Firebase.')

  } finally {
    boton.disabled = false
    boton.textContent = 'ELIMINAR TODOS'
  }
}

async function cargarEstudiantes() {
  const lista = document.querySelector('#lista-estudiantes')
  const total = document.querySelector('#total-estudiantes')

  if (!lista || !total) return

  lista.innerHTML = '<tr><td colspan="8">Consultando Firebase...</td></tr>'

  try {
    const resultado = await getDocs(collection(db, 'estudiantes'))

    const estudiantes = resultado.docs.map(documento => {
      const datos = documento.data()

      return {
        id: documento.id,
        ...datos,
        codigo: String(datos.codigo || '').trim(),
        nombres: String(datos.nombres || '').trim(),
        apellidos: String(datos.apellidos || '').trim(),
        nivel: String(datos.nivel || '').trim(),
        curso: String(datos.curso || '').trim().toUpperCase() === 'SEXO'
          ? 'SEXTO'
          : String(datos.curso || '').trim(),
        paralelo: String(datos.paralelo || '').trim()
      }
    })

    console.log('TOTAL ESTUDIANTES:', estudiantes.length)

    total.textContent = `${estudiantes.length} estudiante(s) registrado(s)`

    const normalizar = texto =>
      String(texto || '')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .trim()
        .toUpperCase()

    const ordenNivel = {
      INICIAL: 1,
      PRIMARIA: 2,
      SECUNDARIA: 3
    }

    function ordenCurso(nivel, curso) {
      const c = normalizar(curso)

      if (normalizar(nivel) === 'INICIAL') {
        if (c.includes('PRIMERA') || c.includes('PRIMER')) return 1
        if (c.includes('SEGUNDA') || c.includes('SEGUNDO')) return 2
        return 99
      }

      const numero = c.match(/PRIMERO|PRIMER|SEGUNDO|TERCERO|CUARTO|QUINTO|SEXTO|[1-6]/)

      if (!numero) return 99

      const equivalencias = {
        PRIMERO: 1,
        PRIMER: 1,
        SEGUNDO: 2,
        TERCERO: 3,
        CUARTO: 4,
        QUINTO: 5,
        SEXTO: 6
      }

      return equivalencias[numero[0]] || Number(numero[0]) || 99
    }

    estudiantes.sort((a, b) => {
      const diferenciaNivel =
        (ordenNivel[normalizar(a.nivel)] || 99) -
        (ordenNivel[normalizar(b.nivel)] || 99)

      if (diferenciaNivel !== 0) return diferenciaNivel

      const diferenciaCurso =
        ordenCurso(a.nivel, a.curso) - ordenCurso(b.nivel, b.curso)

      if (diferenciaCurso !== 0) return diferenciaCurso

      const diferenciaParalelo =
        normalizar(a.paralelo).localeCompare(normalizar(b.paralelo))

      if (diferenciaParalelo !== 0) return diferenciaParalelo

      return normalizar(a.apellidos).localeCompare(normalizar(b.apellidos))
    })

    function llenarFiltro(id, campo, texto) {
      const selector = document.querySelector(id)
      if (!selector) return

      const valorAnterior = selector.value

      const opciones = [...new Set(
        estudiantes
          .map(e => e[campo])
          .filter(Boolean)
      )].sort((a, b) =>
        a.localeCompare(b, 'es', { numeric: true })
      )

      selector.innerHTML =
        `<option value="">${texto}</option>` +
        opciones.map(opcion =>
          `<option value="${escapar(opcion)}">${escapar(opcion)}</option>`
        ).join('')

      selector.value = opciones.includes(valorAnterior) ? valorAnterior : ''
    }

    llenarFiltro(
      '#filtro-nivel-estudiantes',
      'nivel',
      'Todos los niveles'
    )

    llenarFiltro(
      '#filtro-curso-estudiantes',
      'curso',
      'Todos los cursos'
    )

    llenarFiltro(
      '#filtro-paralelo-estudiantes',
      'paralelo',
      'Todos los paralelos'
    )

    function aplicarFiltrosEstudiantes() {
      const nivel = document.querySelector('#filtro-nivel-estudiantes')?.value || ''
      const curso = document.querySelector('#filtro-curso-estudiantes')?.value || ''
      const paralelo = document.querySelector('#filtro-paralelo-estudiantes')?.value || ''
      const busqueda = (
        document.querySelector('#buscar-estudiante')?.value || ''
      ).trim().toLowerCase()

      const filtrados = estudiantes.filter(e => {
        const nombre = `${e.nombres} ${e.apellidos}`.toLowerCase()
        const codigo = e.codigo.toLowerCase()

        return (!nivel || e.nivel === nivel) &&
          (!curso || e.curso === curso) &&
          (!paralelo || e.paralelo === paralelo) &&
          (!busqueda || nombre.includes(busqueda) || codigo.includes(busqueda))
      })

      total.textContent =
        `Mostrando ${filtrados.length} de ${estudiantes.length} estudiantes`

      if (filtrados.length === 0) {
        lista.innerHTML =
          '<tr><td colspan="8">No se encontraron estudiantes.</td></tr>'
        return
      }

      lista.innerHTML = filtrados.map(estudiante => `
        <tr>
          <td>${escapar(estudiante.codigo).toUpperCase()}</td>
          <td>${escapar(estudiante.apellidos).toUpperCase()}</td>
          <td>${escapar(estudiante.nombres).toUpperCase()}</td>
          <td>${escapar(estudiante.nivel).toUpperCase()}</td>
          <td>${escapar(estudiante.curso).toUpperCase()}</td>
          <td>${escapar(estudiante.paralelo).toUpperCase()}</td>
          <td>${estudiante.activo === true ? 'ACTIVO' : 'INACTIVO'}</td>
          <td>
            <button type="button" class="btn-eliminar"
              data-id="${escapar(estudiante.id)}"
              data-codigo="${escapar(estudiante.codigo)}">
              ELIMINAR
            </button>
          </td>
        </tr>
      `).join('')

      lista.querySelectorAll('.btn-eliminar').forEach(boton => {
        boton.addEventListener('click', () => {
          eliminarEstudiante(boton.dataset.id, boton.dataset.codigo)
        })
      })
    }

    document.querySelector('#filtro-nivel-estudiantes')
      ?.addEventListener('change', aplicarFiltrosEstudiantes)

    document.querySelector('#filtro-curso-estudiantes')
      ?.addEventListener('change', aplicarFiltrosEstudiantes)

    document.querySelector('#filtro-paralelo-estudiantes')
      ?.addEventListener('change', aplicarFiltrosEstudiantes)

    document.querySelector('#buscar-estudiante')
      ?.addEventListener('input', aplicarFiltrosEstudiantes)

    aplicarFiltrosEstudiantes()

  } catch (error) {
    console.error('Error al cargar estudiantes:', error)
    lista.innerHTML =
      '<tr><td colspan="8">Error al consultar estudiantes. Revisa la consola.</td></tr>'
    total.textContent = 'No se pudieron cargar los registros.'
  }
}

async function revisarExcel() {
  const archivo = document.querySelector('#archivo-excel').files[0]
  const mensaje = document.querySelector('#mensaje-excel')
  const vista = document.querySelector('#vista-excel')

  mensaje.textContent = ''
  vista.innerHTML = ''

  if (!archivo) {
    mensaje.textContent = 'Primero selecciona un archivo Excel.'
    return
  }

  try {
    const datos = await archivo.arrayBuffer()
    const libro = XLSX.read(datos, { type: 'array' })
    const hoja = libro.Sheets[libro.SheetNames[0]]
    const filas = XLSX.utils.sheet_to_json(hoja, { defval: '' })

    if (filas.length === 0) {
      mensaje.textContent = 'El archivo no contiene registros.'
      return
    }

    const columnas = ['codigo', 'nombres', 'apellidos', 'nivel', 'curso', 'paralelo']
    const encabezados = Object.keys(filas[0])

    const faltantes = columnas.filter(columna => !encabezados.includes(columna))

    if (faltantes.length > 0) {
      mensaje.textContent = `Faltan estas columnas: ${faltantes.join(', ')}`
      return
    }

    const validos = filas.filter(fila =>
      columnas.every(columna => String(fila[columna]).trim() !== '')
    )

    if (validos.length !== filas.length) {
      mensaje.textContent = 'Hay filas incompletas. Corrige el Excel antes de continuar.'
      return
    }

    const codigos = validos.map(fila =>
      String(fila.codigo).trim().toUpperCase()
    )

    if (new Set(codigos).size !== codigos.length) {
      mensaje.textContent = 'El Excel contiene códigos repetidos. Corrígelos antes de continuar.'
      return
    }

    window.estudiantesExcel = validos.map(fila => ({
      codigo: String(fila.codigo).trim().toUpperCase(),
      nombres: String(fila.nombres).trim(),
      apellidos: String(fila.apellidos).trim(),
      nivel: String(fila.nivel).trim(),
      curso: String(fila.curso).trim(),
      paralelo: String(fila.paralelo).trim().toUpperCase(),
      activo: true,
      qr_generado: false
    }))

    vista.innerHTML = `
      <p><strong>Vista previa: ${validos.length} estudiantes</strong></p>
      <table>
        <thead>
          <tr>
            <th>CÓDIGO</th>
            <th>NOMBRES</th>
            <th>APELLIDOS</th>
            <th>NIVEL</th>
            <th>CURSO</th>
            <th>PARALELO</th>
          </tr>
        </thead>
        <tbody>
          ${window.estudiantesExcel.slice(0, 10).map(e => `
            <tr>
              <td>${escapar(e.codigo)}</td>
              <td>${escapar(e.nombres)}</td>
              <td>${escapar(e.apellidos)}</td>
              <td>${escapar(e.nivel)}</td>
              <td>${escapar(e.curso)}</td>
              <td>${escapar(e.paralelo)}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
     <p>Se muestran los primeros 10 registros.</p>

<button type="button" id="importar-estudiantes">
  IMPORTAR ESTUDIANTES A FIREBASE
</button>

<button type="button" id="respaldo-estudiantes">
  DESCARGAR COPIA DE SEGURIDAD
</button>
`
    mensaje.textContent = 'Archivo leído correctamente. Todavía no se ha guardado ningún estudiante.'
document.querySelector('#importar-estudiantes')
  .addEventListener('click', importarEstudiantes)

document.querySelector('#respaldo-estudiantes')
  .addEventListener('click', descargarCopiaEstudiantes)
  } catch (error) {
    console.error(error)
    mensaje.textContent = 'No se pudo leer el archivo Excel.'
  }
}

async function importarEstudiantes() {
  const boton = document.querySelector('#importar-estudiantes')
  const mensaje = document.querySelector('#mensaje-excel')

  const estudiantes = window.estudiantesExcel || []

  if (estudiantes.length === 0) {
    mensaje.textContent = 'Primero selecciona un archivo Excel válido.'
    return
  }

  const confirmar = confirm(
    `¿Deseas importar ${estudiantes.length} estudiantes a Firebase?`
  )

  if (!confirmar) return

  boton.disabled = true
  boton.textContent = 'IMPORTANDO...'
  mensaje.textContent = ''

  let guardados = 0
  let duplicados = 0

  try {
    const existentes = await getDocs(collection(db, 'estudiantes'))

    const codigosExistentes = new Set(
      existentes.docs.map(documento =>
        String(documento.data().codigo || '').trim().toUpperCase()
      )
    )

    for (const estudiante of estudiantes) {
      const codigo = String(estudiante.codigo || '').trim().toUpperCase()

      if (!codigo) {
        duplicados++
        continue
      }

      if (codigosExistentes.has(codigo)) {
        duplicados++
        continue
      }

      await addDoc(collection(db, 'estudiantes'), {
        ...estudiante,
        codigo
      })

      codigosExistentes.add(codigo)
      guardados++
    }

    mensaje.textContent =
      `Importación finalizada. Guardados: ${guardados}. Duplicados omitidos: ${duplicados}.`

    window.estudiantesExcel = []

    await cargarEstudiantes()

  } catch (error) {
    console.error('Error al importar:', error)
    mensaje.textContent = `Error al importar: ${error.code || error.message}`
  } finally {
    boton.disabled = false
    boton.textContent = 'IMPORTAR ESTUDIANTES A FIREBASE'
  }
}

async function guardarEstudiante(evento) {
  evento.preventDefault()

  const boton = document.querySelector('#guardar-estudiante')
  const mensaje = document.querySelector('#mensaje-estudiante')

  const estudiante = {
  codigo: document.querySelector('#codigo').value.trim().toUpperCase(),
  nombres: document.querySelector('#nombres').value.trim(),
  apellidos: document.querySelector('#apellidos').value.trim(),
  nivel: document.querySelector('#nivel').value,
  curso: document.querySelector('#curso').value.trim(),
  paralelo: document.querySelector('#paralelo').value,
  activo: true,
  qr_id: crypto.randomUUID(),
  qr_generado: true
}

  boton.disabled = true
  boton.textContent = 'GUARDANDO...'
  mensaje.textContent = ''

  try {
    const duplicado = query(
      collection(db, 'estudiantes'),
      where('codigo', '==', estudiante.codigo),
      limit(1)
    )

    const existente = await getDocs(duplicado)

    if (!existente.empty) {
      mensaje.textContent = 'Ese código ya está registrado. Utiliza otro.'
      return
    }

    await addDoc(collection(db, 'estudiantes'), estudiante)

    mensaje.textContent = '¡ESTUDIANTE REGISTRADO CORRECTAMENTE!'
    document.querySelector('#form-estudiante').reset()

    await cargarEstudiantes()

  } catch (error) {
    console.error('Error al guardar:', error)
    mensaje.textContent = `No se pudo guardar: ${error.code || error.message}`
  } finally {
    boton.disabled = false
    boton.textContent = 'GUARDAR ESTUDIANTE'
  }
}

onAuthStateChanged(auth, async (usuario) => {
    if (modoJurado) return
  if (!usuario) {
    app.innerHTML = `
      <main class="inicio">
        <section class="bienvenida login-modern">

<div class="login-logo">
  <p class="login-etiqueta">COLEGIO ADVENTISTA SHALOM</p>

  <h2>SHALOM EXPO <span>2026</span></h2>

  <div class="login-linea"></div>

  <p class="login-descripcion">
    Sistema de gestión institucional y evaluación de proyectos
  </p>
</div>


<div class="login-titulo">
  <div class="login-icono-bienvenida">✦</div>
  <span class="login-subtitulo">ACCESO INSTITUCIONAL</span>
  <h3>¡Bienvenido!</h3>
  <p>Ingresa tus credenciales para acceder a la plataforma.</p>
</div>

<form id="login-form" class="login-form">

  <div class="login-campo">
    <label for="correo">CORREO ELECTRÓNICO</label>
    <div class="login-input-contenedor">
      <span>✉</span>
      <input
        id="correo"
        type="email"
        placeholder="nombre@institucion.edu"
        autocomplete="username"
        required
      >
    </div>
  </div>

  <div class="login-campo">
    <label for="clave">CONTRASEÑA</label>
    <div class="login-input-contenedor">
      <span>🔒</span>
      <input
        id="clave"
        type="password"
        placeholder="Ingresa tu contraseña"
        autocomplete="current-password"
        required
      >
      <button
        type="button"
        id="mostrar-clave"
        class="btn-mostrar-clave"
        aria-label="Mostrar contraseña"
      >◉</button>
    </div>
  </div>

  <button type="submit" class="btn-ingresar">
    <span>INGRESAR AL SISTEMA</span>
    <span class="login-flecha">→</span>
  </button>

  <p id="mensaje" role="status"></p>

</form>

<div class="login-pie">
  <div class="login-pie-linea"></div>
  <span>SHALOM EXPO 2026</span>
  <p>Innovación <span>·</span> Educación <span>·</span> Servicio</p>
</div>
</section>
</main>
`

    document.querySelector('#login-form')
      .addEventListener('submit', async (evento) => {
        evento.preventDefault()

        const mensaje = document.querySelector('#mensaje')
        const boton = document.querySelector('.btn-ingresar')

        mensaje.textContent = ''
        boton.disabled = true
        boton.textContent = 'VERIFICANDO ACCESO...'

        try {
          await signInWithEmailAndPassword(
            auth,
            document.querySelector('#correo').value.trim(),
            document.querySelector('#clave').value
          )
        } catch (error) {
          mensaje.textContent = `Error de acceso: ${error.code || error.message}`
          boton.disabled = false
          boton.textContent = 'INGRESAR AL SISTEMA →'
        }
      })

    document.querySelector('#mostrar-clave')
      .addEventListener('click', () => {
        const clave = document.querySelector('#clave')
        const boton = document.querySelector('#mostrar-clave')

        clave.type = clave.type === 'password' ? 'text' : 'password'

        boton.textContent = clave.type === 'password' ? '◉' : '◎'
        boton.setAttribute(
          'aria-label',
          clave.type === 'password' ? 'Mostrar contraseña' : 'Ocultar contraseña'
        )
      })

    return
  }

  try {
    const referencia = doc(db, 'usuarios', usuario.uid)
    const resultado = await getDoc(referencia)

    if (!resultado.exists()) {
      await signOut(auth)
      return
    }

    const datosUsuario = resultado.data()

    const rolesPermitidos = [
      'administrador',
      'director',
      'coordinador',
      'secretaria'
    ]

    if (
      datosUsuario.activo !== true ||
      !rolesPermitidos.includes(datosUsuario.rol)
    ) {
      await signOut(auth)
      return
    }

    administrador = datosUsuario
    mostrarPanel(administrador.nombre || 'Administrador')

  } catch (error) {
    console.error('Error:', error)
    app.innerHTML = '<h2>Error al verificar los permisos del administrador.</h2>'
  }
})

// =====================================================
// REGISTRO AUTOMÁTICO DE JURADOS
// SHALOM EXPO 2026
// =====================================================

async function iniciarRegistroJurados() {

  app.innerHTML = `
  <main class="jurado-page">
    <section class="jurado-card">

      <header class="jurado-header">
        <div class="jurado-marca">
          <span class="jurado-marca-linea"></span>
          <span>COLEGIO ADVENTISTA SHALOM</span>
        </div>

        <p class="jurado-evento">PLATAFORMA INSTITUCIONAL</p>

        <h1>SHALOM EXPO <span>2026</span></h1>

        <p class="jurado-subtitulo">
          Registro oficial de jurados evaluadores
        </p>
      </header>

      <div class="jurado-separador"></div>

      <section class="jurado-formulario">
        <div class="jurado-introduccion">
          <span class="jurado-paso">IDENTIFICACIÓN PERSONAL</span>
          <h2>Registro de jurado</h2>
          <p>
            Completa tus datos personales para generar
            tu identificación digital de evaluador.
          </p>
        </div>

        <form id="form-registro-jurado" class="jurado-form">

          <div class="jurado-input-group">
            <label for="jurado-nombres">Nombres</label>
            <input
              id="jurado-nombres"
              type="text"
              placeholder="Escribe tus nombres"
              autocomplete="given-name"
              required
            >
          </div>

          <div class="jurado-input-group">
            <label for="jurado-apellidos">Apellidos</label>
            <input
              id="jurado-apellidos"
              type="text"
              placeholder="Escribe tus apellidos"
              autocomplete="family-name"
              required
            >
          </div>

          <button type="submit" class="jurado-btn-principal">
            Crear mi identificación
            <span>→</span>
          </button>

          <p id="mensaje-jurado" class="jurado-mensaje" role="status"></p>
        </form>
      </section>


      <div id="identificacion-jurado"></div>

      <div class="jurado-versiculo">
        <div class="jurado-versiculo-linea"></div>
        <p>
          “Y todo lo que hagáis, hacedlo de corazón,
          como para el Señor y no para los hombres”.
        </p>
        <span>COLOSENSES 3:23</span>
      </div>

      <footer class="jurado-footer">

        SHALOM EXPO 2026 · SISTEMA DE EVALUACIÓN
      </footer>

    </section>
  </main>
`

  try {
    const credencial = await signInAnonymously(authJurados)
    const usuario = credencial.user

   const referencia = doc(dbJurados, 'jurados', usuario.uid)
const existente = await getDoc(referencia)

   if (existente.exists()) {

  staffActual = {
    uid: usuario.uid,
    ...existente.data()
  }

  mostrarPanelStaff()
  return
}
    document.querySelector('#form-registro-jurado')
      .addEventListener('submit', async (evento) => {
        evento.preventDefault()

        const boton = evento.currentTarget.querySelector('button')
        const mensaje = document.querySelector('#mensaje-jurado')

        const nombres = document.querySelector('#jurado-nombres')
          .value.trim()

        const apellidos = document.querySelector('#jurado-apellidos')
          .value.trim()

        if (!nombres || !apellidos) {
          mensaje.textContent = 'Completa tus nombres y apellidos.'
          return
        }

        boton.disabled = true
        boton.textContent = 'REGISTRANDO...'

        try {
          
let datosJurado
let codigoReservado = false

for (let intento = 0; intento < 5; intento++) {
  const aleatorios = new Uint8Array(8)
  crypto.getRandomValues(aleatorios)

  const codigoAleatorio = Array.from(aleatorios)
    .map(numero => numero.toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase()

  const codigo = `JUR-${codigoAleatorio}`
  const codigoRef = doc(dbJurados, 'codigos_jurados', codigo)

  try {
    await runTransaction(dbJurados, async (transaccion) => {
      const codigoExistente = await transaccion.get(codigoRef)
      const juradoExistente = await transaccion.get(referencia)

      if (juradoExistente.exists()) {
        throw new Error('Este celular ya tiene un registro.')
      }

      if (codigoExistente.exists()) {
        throw new Error('CODIGO_DUPLICADO')
      }

      datosJurado = {
        uid: usuario.uid,
        codigo,
        nombres,
        apellidos,
        nombreCompleto: `${nombres} ${apellidos}`,
        activo: true,
        creadoEn: new Date().toISOString()
      }

      transaccion.set(codigoRef, {
        uid: usuario.uid,
        creadoEn: new Date().toISOString()
      })

      transaccion.set(referencia, datosJurado)
    })

    codigoReservado = true
    break

  } catch (error) {
    if (error.message !== 'CODIGO_DUPLICADO') {
      throw error
    }
  }
}

if (!codigoReservado) {
  throw new Error('No se pudo generar un código único. Intenta nuevamente.')
}

          await mostrarIdentificacionJurado(datosJurado)

        } catch (error) {
          console.error('Error al registrar jurado:', error)
          mensaje.textContent =
            'No se pudo completar el registro. Verifica los permisos de Firebase.'
          boton.disabled = false
          boton.textContent = 'REGISTRARME →'
        }
      })

  } catch (error) {
    console.error('Error de autenticación de jurado:', error)
    app.innerHTML = '<h2>No se pudo iniciar el registro. Intenta nuevamente.</h2>'
  }
}


// MOSTRAR IDENTIFICACIÓN Y QR PERSONAL
async function mostrarIdentificacionJurado(datos) {
  const formulario = document.querySelector('#form-registro-jurado')
  if (formulario) formulario.style.display = 'none'

  const contenedor = document.querySelector('#identificacion-jurado')

  contenedor.innerHTML = `
    <div style="text-align:center; padding:20px;">
      <h3>¡Registro completado!</h3>
      <p>Tu identificación como jurado es:</p>
      <h2 id="codigo-jurado"></h2>
      <p id="nombre-jurado"></p>
      <img id="qr-jurado" alt="QR personal del jurado"
        style="width:210px; max-width:100%; margin:15px auto; display:block;">
      <p>Conserva esta identificación para el evento.</p>
      <button id="ir-evaluacion-jurado" class="btn-ingresar">
        INGRESAR A EVALUAR →
      </button>
    </div>
  `

  document.querySelector('#codigo-jurado').textContent = datos.codigo
  document.querySelector('#nombre-jurado').textContent =
    datos.nombreCompleto

  try {
    const imagenQR = await QRCode.toDataURL(datos.codigo, {
      width: 300,
      margin: 2
    })

    document.querySelector('#qr-jurado').src = imagenQR
  } catch (error) {
    console.error('Error al generar QR:', error)
  }

  document.querySelector('#ir-evaluacion-jurado')
    .addEventListener('click', () => {
      window.location.href = '?modo=evaluacion-jurado'
    })
}


// ACTIVAR LA RUTA DE REGISTRO
if (modoJurado === 'registro-jurado') {
  iniciarRegistroJurados()
}
// =====================================================
// ACCESO INDEPENDIENTE PARA PERSONAL STAFF
// =====================================================

if (modoJurado === 'staff') {
  iniciarAccesoStaff()
}

if (modoJurado === 'evaluacion-jurado') {
  iniciarEvaluacionJurados()
}

// =====================================================
// ACCESO PÚBLICO DE EVALUACIÓN PARA JURADOS
// =====================================================

async function iniciarEvaluacionJurados() {

  // PORTADA DE VERIFICACIÓN
  app.innerHTML = `
    <main class="inicio">
      <section class="bienvenida login-modern">

        <div class="login-logo">
          <div class="login-icono">S</div>
          <p class="login-etiqueta">COLEGIO ADVENTISTA SHALOM</p>
          <h2>SHALOM EXPO <span>2026</span></h2>
          <div class="login-linea"></div>
        </div>

        <div class="login-titulo">
          <h3>Acceso de Jurados</h3>
          <p>Estamos verificando tu identificación.</p>
        </div>

        <div class="login-verificando">
          <div class="login-spinner"></div>
          <p>Verificando credenciales...</p>
        </div>

      </section>
    </main>
  `;

  try {

    // 1. ESPERAR LA RESTAURACIÓN DE SESIÓN
    await authJurados.authStateReady();

    let usuario = authJurados.currentUser;

    // 2. INICIAR SESIÓN ANÓNIMA SI ES NECESARIO
    if (!usuario) {
      const credencial = await signInAnonymously(authJurados);
      usuario = credencial.user;
    }

    // 3. VERIFICAR IDENTIDAD DEL JURADO
    const referencia = doc(
      dbJurados,
      'jurados',
      usuario.uid
    );

    const resultado = await getDoc(referencia);

    // 4. COMPROBAR QUE EL JURADO ESTÉ REGISTRADO Y ACTIVO
    if (
      !resultado.exists() ||
      resultado.data().activo !== true
    ) {

      app.innerHTML = `
        <main class="inicio">
          <section class="bienvenida login-modern">

            <div class="login-logo">
              <div class="login-icono">S</div>
              <p class="login-etiqueta">COLEGIO ADVENTISTA SHALOM</p>
              <h2>SHALOM EXPO <span>2026</span></h2>
              <div class="login-linea"></div>
            </div>

            <div class="login-titulo">
              <h3>Identificación no encontrada</h3>
              <p>
                No encontramos un registro de jurado activo
                asociado a este dispositivo.
              </p>
            </div>

            <button
              id="volver-registro"
              class="btn-ingresar"
              type="button"
            >
              REGISTRARME COMO JURADO →
            </button>

            <div class="login-pie">
              <span>SHALOM EXPO 2026</span>
              <p>Innovación · Educación · Servicio</p>
            </div>

          </section>
        </main>
      `;

      document
        .querySelector('#volver-registro')
        .addEventListener('click', () => {
          window.location.href = '?modo=registro-jurado';
        });

      return;
    }

    // 5. GUARDAR LOS DATOS DEL JURADO IDENTIFICADO
    juradoActual = {
      uid: usuario.uid,
      ...resultado.data()
    };

    // 6. ABRIR EL PANEL DE EVALUACIÓN
    app.innerHTML = `
      <main id="contenido" class="contenido">
        <section class="evaluacion-panel">
          <p>Preparando tu panel de evaluación...</p>
        </section>
      </main>
    `;

    mostrarPanelJurado();

  } catch (error) {

    console.error(
      'Error al ingresar a evaluación:',
      error
    );

    app.innerHTML = `
      <main class="inicio">
        <section class="bienvenida login-modern">

          <div class="login-logo">
            <div class="login-icono">S</div>
            <p class="login-etiqueta">COLEGIO ADVENTISTA SHALOM</p>
            <h2>SHALOM EXPO <span>2026</span></h2>
            <div class="login-linea"></div>
          </div>

          <div class="login-titulo">
            <h3>No se pudo ingresar</h3>
            <p>
              Ocurrió un problema al verificar tu identificación.
              Comprueba tu conexión a Internet e inténtalo nuevamente.
            </p>
          </div>

          <button
            id="reintentar-jurado"
            class="btn-ingresar"
            type="button"
          >
            VOLVER A INTENTAR →
          </button>

          <div class="login-pie">
            <span>SHALOM EXPO 2026</span>
            <p>Innovación · Educación · Servicio</p>
          </div>

        </section>
      </main>
    `;

    document
      .querySelector('#reintentar-jurado')
      ?.addEventListener('click', () => {
        iniciarEvaluacionJurados();
      });

  }
}

// ACTIVAR EL ACCESO DE EVALUACIÓN
if (modoJurado === 'evaluacion-jurado') {
  iniciarEvaluacionJurados();
}
// =====================================================
// PANEL PRINCIPAL DEL JURADO
// =====================================================

function mostrarPanelJurado() {

  if (!juradoActual) {
    iniciarEvaluacionJurados();
    return;
  }

  app.innerHTML = `
    <main class="jurado-page">
      <section class="jurado-card">

        <header class="jurado-header">
          <div class="jurado-marca">
            <span class="jurado-marca-linea"></span>
            <span>COLEGIO ADVENTISTA SHALOM</span>
          </div>

          <p class="jurado-evento">PLATAFORMA DE EVALUACIÓN</p>
          <h1>SHALOM EXPO <span>2026</span></h1>
          <p class="jurado-subtitulo">Panel de evaluación del jurado</p>
        </header>

        <div class="jurado-separador"></div>

        <section class="jurado-formulario">

          <div class="jurado-introduccion">
            <span class="jurado-paso">SESIÓN ACTIVA</span>
            <h2>¡Bienvenido!</h2>
            <h3>${escapar(juradoActual.nombreCompleto || '')}</h3>
            <p>
              Selecciona cómo deseas comenzar la evaluación
              de los proyectos de SHALOM EXPO 2026.
            </p>
          </div>

          <div class="staff-panel-acciones">

            <button
              type="button"
              id="jurado-escanear-proyecto"
              class="jurado-btn-principal">
              <span>▦</span>
              ESCANEAR QR DEL PROYECTO
            </button>

            <button
              type="button"
              id="jurado-evaluacion-manual"
              class="jurado-btn-principal">
              <span>✎</span>
              EVALUACIÓN MANUAL
            </button>

            <p class="jurado-mensaje">
              Puedes escanear el código QR del stand
              o ingresar directamente a la evaluación manual.
            </p>

          </div>

        </section>

        <div class="jurado-versiculo">
          <div class="jurado-versiculo-linea"></div>
          <p>
            “Y todo lo que hagáis, hacedlo de corazón,
            como para el Señor”.
          </p>
          <span>COLOSENSES 3:23</span>
        </div>

        <footer class="jurado-footer">
          SHALOM EXPO 2026 · JURADOS
        </footer>

      </section>
    </main>
  `;

  document
    .querySelector('#jurado-escanear-proyecto')
    .addEventListener('click', () => {
      mostrarEscanerProyectoJurado();
    });

  document
    .querySelector('#jurado-evaluacion-manual')
    .addEventListener('click', () => {
      mostrarEvaluacion();
    });
}


// =====================================================
// ESCÁNER QR DE PROYECTOS PARA JURADOS
// =====================================================

async function mostrarEscanerProyectoJurado() {

  if (!juradoActual) {
    iniciarEvaluacionJurados();
    return;
  }

  let lectorQRProyectoJurado = null;
  let procesandoQR = false;

  app.innerHTML = `
    <main class="jurado-page">
      <section class="jurado-card">

        <header class="jurado-header">
          <div class="jurado-marca">
            <span class="jurado-marca-linea"></span>
            <span>COLEGIO ADVENTISTA SHALOM</span>
          </div>

          <p class="jurado-evento">EVALUACIÓN DE PROYECTOS</p>
          <h1>SHALOM EXPO <span>2026</span></h1>
          <p class="jurado-subtitulo">Escáner QR del proyecto</p>
        </header>

        <div class="jurado-separador"></div>

        <section class="jurado-formulario">

          <div class="jurado-introduccion">
            <span class="jurado-paso">IDENTIFICACIÓN DEL PROYECTO</span>
            <h2>Escanear código QR</h2>
            <p>
              Enfoca el código QR del stand que deseas evaluar.
            </p>
            <h3>${escapar(juradoActual.nombreCompleto || '')}</h3>
          </div>

          <div id="lector-qr-proyecto-jurado"></div>

          <div
            id="resultado-proyecto-jurado"
            class="jurado-mensaje"
            role="status">
            Esperando lectura del código QR...
          </div>

          <button
            type="button"
            id="jurado-volver-panel"
            class="jurado-btn-principal">
            ← VOLVER AL PANEL DEL JURADO
          </button>

        </section>

        <footer class="jurado-footer">
          SHALOM EXPO 2026 · JURADOS
        </footer>

      </section>
    </main>
  `;

  const resultado = document.querySelector(
    '#resultado-proyecto-jurado'
  );

  async function detenerEscanerProyectoJurado() {

    if (!lectorQRProyectoJurado) return;

    try {

      if (lectorQRProyectoJurado.isScanning) {
        await lectorQRProyectoJurado.stop();
      }

      await lectorQRProyectoJurado.clear();

    } catch (error) {
      console.warn('No se pudo cerrar el escáner del jurado:', error);
    }

    lectorQRProyectoJurado = null;
  }

  document
    .querySelector('#jurado-volver-panel')
    .addEventListener('click', async () => {
      await detenerEscanerProyectoJurado();
      mostrarPanelJurado();
    });

  try {

    lectorQRProyectoJurado = new Html5Qrcode(
      'lector-qr-proyecto-jurado'
    );

    await lectorQRProyectoJurado.start(
      { facingMode: 'environment' },
      {
        fps: 10,
        qrbox: {
          width: 250,
          height: 250
        }
      },

      async (textoQR) => {

        if (procesandoQR) return;

        const prefijo = 'SHALOM-PROYECTO:';

        if (!textoQR.startsWith(prefijo)) {
          resultado.textContent =
            'Este código QR no corresponde a un proyecto de SHALOM EXPO.';
          return;
        }

        procesandoQR = true;

        try {

          resultado.textContent = 'Identificando proyecto...';

          const proyectoId = textoQR
            .substring(prefijo.length)
            .trim();

          if (!proyectoId) {
            throw new Error('El código QR no contiene un identificador válido.');
          }

          const referenciaProyecto = doc(
            db,
            'proyectos',
            proyectoId
          );

          const respuestaProyecto = await getDoc(
            referenciaProyecto
          );

          if (!respuestaProyecto.exists()) {
            throw new Error('No encontramos este proyecto en el sistema.');
          }

          const proyecto = {
            id: respuestaProyecto.id,
            ...respuestaProyecto.data()
          };

          if (proyecto.activo === false) {
            throw new Error('Este proyecto se encuentra desactivado.');
          }

          await detenerEscanerProyectoJurado();

          app.innerHTML = `
            <main id="contenido" class="contenido">
              <section class="evaluacion-panel">
                <p>Proyecto identificado. Preparando evaluación...</p>
              </section>
            </main>
          `;

          await mostrarEvaluacion(proyecto);

        } catch (error) {

          console.error('Error al identificar proyecto para jurado:', error);

          resultado.textContent =
            error.message || 'No se pudo identificar el proyecto.';

          procesandoQR = false;
        }
      },

      () => {}
    );

  } catch (error) {

    console.error('Error al iniciar escáner del jurado:', error);

    resultado.textContent =
      'No se pudo iniciar la cámara. Verifica los permisos del navegador.';
  }
}

 // =====================================================
// GESTIÓN DE DOCENTES - SHALOM EXPO 2026
// REGISTRO, ASIGNACIONES E IMPORTACIÓN EXCEL
// =====================================================

async function mostrarDocentes() {
  const contenido = document.querySelector('#contenido');

  const puedeEditar = ['administrador', 'coordinador'].includes(
    administrador?.rol
  );

  let docentes = [];
let cursosDisponibles = [];
let docenteEditandoId = null;
let catalogoAcademico = [];
let catalogoEditandoId = null;

  const escapar = valor => String(valor ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');

  function normalizarNivel(valor) {
    const nivel = String(valor || '').trim().toLowerCase();

    if (nivel.includes('inicial')) return 'Inicial';
    if (nivel.includes('primaria')) return 'Primaria';
    if (nivel.includes('secundaria')) return 'Secundaria';

    return '';
  }

  function obtenerNombreCurso(estudiante) {
    const curso = String(
      estudiante.curso || estudiante.grado || ''
    ).trim();

    const paralelo = String(
      estudiante.paralelo || estudiante.seccion || ''
    ).trim();

    if (!curso) return '';

    // Si el curso ya contiene el paralelo, no lo repetimos.
    if (
      paralelo &&
      !curso.toLowerCase().includes(paralelo.toLowerCase())
    ) {
      return `${curso} ${paralelo}`.trim();
    }

    return curso;
  }

  function ordenarCursos(a, b) {
    const orden = [
      'pre-kinder',
      'prekinder',
      'kinder',
      'primero',
      'segundo',
      'tercero',
      'cuarto',
      'quinto',
      'sexto'
    ];

    const textoA = a.toLowerCase();
    const textoB = b.toLowerCase();

    const indiceA = orden.findIndex(x => textoA.includes(x));
    const indiceB = orden.findIndex(x => textoB.includes(x));

    if (indiceA !== indiceB) {
      if (indiceA === -1) return 1;
      if (indiceB === -1) return -1;
      return indiceA - indiceB;
    }

    return a.localeCompare(b, 'es', {
      numeric: true,
      sensitivity: 'base'
    });
  }

  async function cargarCursosDesdeEstudiantes() {
    const resultado = await getDocs(collection(db, 'estudiantes'));

    const cursosUnicos = new Map();

    resultado.docs.forEach(documento => {
      const estudiante = documento.data();

      if (estudiante.activo === false) return;

      const nivel = normalizarNivel(estudiante.nivel);
      const curso = obtenerNombreCurso(estudiante);

      if (!nivel || !curso) return;

      const clave = `${nivel}|${curso}`.toLowerCase();

      cursosUnicos.set(clave, {
        nivel,
        curso
      });
    });

    cursosDisponibles = [...cursosUnicos.values()].sort((a, b) => {
      const niveles = {
        Inicial: 1,
        Primaria: 2,
        Secundaria: 3
      };

      if (niveles[a.nivel] !== niveles[b.nivel]) {
        return niveles[a.nivel] - niveles[b.nivel];
      }

      return ordenarCursos(a.curso, b.curso);
    });
  }

  contenido.innerHTML = `
    <style>
      .docentes-admin {
        --doc-morado: #7045a5;
        --doc-oscuro: #382052;
        --doc-lila: #f2eafa;
        --doc-naranja: #f39a45;
        --doc-texto: #352c40;
        --doc-borde: #e9e2f0;
        color: var(--doc-texto);
      }

      .docentes-admin .docentes-hero {
        background: linear-gradient(120deg, #382052, #7045a5 70%, #9870c8);
        color: white;
        padding: 30px;
        border-radius: 20px;
        margin-bottom: 22px;
        box-shadow: 0 10px 30px rgba(56,32,82,.15);
      }

      .docentes-admin .docentes-hero span {
        font-size: 11px;
        letter-spacing: 2px;
        font-weight: 800;
        opacity: .8;
      }

      .docentes-admin .docentes-hero h2 {
        margin: 9px 0;
        font-size: clamp(23px, 4vw, 32px);
        color: white;
      }

      .docentes-admin .docentes-hero p {
        margin: 0;
        color: #f0e8fa;
        font-size: 14px;
      }

      .docentes-admin .docentes-indicadores {
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 15px;
        margin-bottom: 22px;
      }

      .docentes-admin .docentes-indicador {
        background: white;
        border: 1px solid var(--doc-borde);
        border-radius: 16px;
        padding: 19px;
        display: flex;
        align-items: center;
        gap: 14px;
        box-shadow: 0 4px 15px rgba(56,32,82,.04);
      }

      .docentes-admin .docentes-indicador-icono {
        width: 47px;
        height: 47px;
        flex-shrink: 0;
        border-radius: 14px;
        display: grid;
        place-items: center;
        font-size: 23px;
        background: var(--doc-lila);
        color: var(--doc-morado);
      }

      .docentes-admin .docentes-indicador-icono.verde {
        background: #e5f6ec;
        color: #21844b;
      }

      .docentes-admin .docentes-indicador span {
        display: block;
        font-size: 10px;
        letter-spacing: 1px;
        font-weight: 800;
        color: #80758b;
      }

      .docentes-admin .docentes-indicador strong {
        display: block;
        font-size: 26px;
        margin: 3px 0;
        color: var(--doc-oscuro);
      }

      .docentes-admin .docentes-indicador small {
        color: #8c8495;
        font-size: 11px;
      }

      .docentes-admin .docentes-tarjeta {
        background: white;
        border: 1px solid var(--doc-borde);
        border-radius: 18px;
        padding: 23px;
        margin-bottom: 22px;
        box-shadow: 0 5px 20px rgba(56,32,82,.045);
      }

      .docentes-admin .docentes-seccion-titulo {
        margin-bottom: 20px;
      }

      .docentes-admin .docentes-seccion-titulo h3 {
        margin: 0 0 6px;
        color: var(--doc-oscuro);
        font-size: 19px;
      }

      .docentes-admin .docentes-seccion-titulo p {
        margin: 0;
        color: #89808f;
        font-size: 13px;
      }

      .docentes-admin .docentes-campos {
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 16px;
      }

      .docentes-admin label {
        display: block;
        font-size: 12px;
        font-weight: 700;
        color: #51465d;
      }

      .docentes-admin input,
      .docentes-admin select {
        width: 100%;
        min-width: 0;
        margin-top: 7px;
        padding: 12px 13px;
        border: 1px solid #e4dce9;
        border-radius: 10px;
        background: #fff;
        color: #352c40;
        font: inherit;
        font-size: 13px;
        box-sizing: border-box;
        outline: none;
      }

      .docentes-admin input:focus,
      .docentes-admin select:focus {
        border-color: var(--doc-morado);
        box-shadow: 0 0 0 3px rgba(112,69,165,.10);
      }

      .docentes-admin .docentes-asignacion {
        margin-top: 20px;
        border: 1px solid var(--doc-borde);
        border-radius: 15px;
        padding: 18px;
        background: #fdfbff;
      }

      .docentes-admin .docentes-asignacion-cabecera {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 12px;
        margin-bottom: 17px;
      }

      .docentes-admin .docentes-asignacion-cabecera h4 {
        margin: 0 0 5px;
        color: var(--doc-oscuro);
        font-size: 15px;
      }

      .docentes-admin .docentes-asignacion-cabecera p {
        margin: 0;
        color: #8a8092;
        font-size: 12px;
      }

      .docentes-admin .docentes-asignacion-fila {
        background: white;
        border: 1px solid #ece5f1;
        border-radius: 13px;
        padding: 16px;
        margin-bottom: 14px;
      }

      .docentes-admin .docentes-asignacion-grid {
        display: grid;
        grid-template-columns: 1fr 1.3fr auto;
        gap: 12px;
        align-items: end;
      }

      .docentes-admin .docentes-cursos-cabecera {
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 10px;
        margin: 17px 0 10px;
      }

      .docentes-admin .docentes-cursos-cabecera strong {
        color: var(--doc-oscuro);
        font-size: 12px;
      }

      .docentes-admin .docentes-cursos-cabecera small {
        color: #8a8092;
        font-size: 11px;
      }

      .docentes-admin .docentes-cursos-grid {
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: 8px;
      }

      .docentes-admin .docente-curso-opcion {
        display: flex;
        align-items: center;
        gap: 8px;
        padding: 10px;
        border: 1px solid #e9e1ef;
        border-radius: 9px;
        background: #fff;
        cursor: pointer;
        font-size: 12px;
        font-weight: 600;
        min-width: 0;
      }

      .docentes-admin .docente-curso-opcion:has(input:checked) {
        background: #f2eafa;
        border-color: #9870c8;
        color: #54327e;
      }

      .docentes-admin .docente-curso-opcion input {
        width: 16px;
        height: 16px;
        min-width: 16px;
        margin: 0;
        accent-color: #7045a5;
      }

      .docentes-admin .docentes-sin-cursos {
        padding: 15px;
        background: #fff8ed;
        border: 1px dashed #edc58f;
        border-radius: 10px;
        color: #89602d;
        font-size: 12px;
      }

      .docentes-admin .docentes-boton {
        border: 0;
        border-radius: 10px;
        padding: 11px 16px;
        background: var(--doc-morado);
        color: white;
        font-size: 12px;
        font-weight: 800;
        cursor: pointer;
      }

      .docentes-admin .docentes-boton:hover {
        filter: brightness(1.08);
      }

      .docentes-admin .docentes-boton-naranja {
        background: var(--doc-naranja);
        color: #382052;
      }

      .docentes-admin .docentes-boton-claro {
        background: var(--doc-lila);
        color: var(--doc-morado);
      }

      .docentes-admin .docentes-boton-peligro {
        background: #fff0f0;
        color: #b33d3d;
      }

      .docentes-admin .docentes-boton:disabled {
        opacity: .55;
        cursor: wait;
      }

      .docentes-admin .docentes-acciones {
        display: flex;
        flex-wrap: wrap;
        gap: 10px;
        margin-top: 18px;
      }

      .docentes-admin .docentes-mensaje {
        margin: 13px 0 0;
        font-size: 13px;
        font-weight: 700;
      }

      .docentes-admin .docentes-listado-cabecera {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 15px;
        margin-bottom: 16px;
      }

      .docentes-admin .docentes-listado-cabecera input {
        max-width: 310px;
      }

      .docentes-admin .docentes-tabla-contenedor {
        overflow-x: auto;
        width: 100%;
      }

      .docentes-admin .docentes-tabla-contenedor table {
        min-width: 850px;
        width: 100%;
        border-collapse: collapse;
      }

      .docentes-admin .docentes-tabla-contenedor th {
        background: #f5effa;
        color: #59416e;
        font-size: 10px;
        letter-spacing: .6px;
        padding: 13px 10px;
        text-align: left;
      }

      .docentes-admin .docentes-tabla-contenedor td {
        padding: 13px 10px;
        border-bottom: 1px solid #eee8f2;
        font-size: 12px;
        vertical-align: top;
      }

      .docentes-admin .docente-asignacion-resumen {
        margin-bottom: 7px;
      }

      .docentes-admin .docente-asignacion-resumen strong {
        display: block;
        color: #513477;
      }

      .docentes-admin .docente-asignacion-resumen small {
        display: block;
        color: #8a8092;
        margin-top: 3px;
      }

      .docentes-admin .docentes-estado {
        display: inline-block;
        padding: 5px 9px;
        border-radius: 20px;
        font-size: 10px;
        font-weight: 800;
      }

      .docentes-admin .docentes-estado.activo {
        background: #e5f6ec;
        color: #21844b;
      }

      .docentes-admin .docentes-estado.inactivo {
        background: #fff0f0;
        color: #b33d3d;
      }

      .docentes-admin .docentes-acciones-tabla {
        display: flex;
        flex-wrap: wrap;
        gap: 5px;
      }

      .docentes-admin .docentes-acciones-tabla button {
        border: 0;
        border-radius: 7px;
        padding: 7px 9px;
        cursor: pointer;
        font-size: 10px;
        font-weight: 700;
        background: #f2eafa;
        color: #7045a5;
      }

      .docentes-admin .docentes-acciones-tabla button:last-child {
        background: #fff0f0;
        color: #b33d3d;
      }

      .docentes-admin .docentes-volver {
        margin: 0 0 20px;
        background: #eee7f4;
        color: #50336e;
      }

      @media(max-width: 700px) {
        .docentes-admin .docentes-hero {
          padding: 22px;
        }

        .docentes-admin .docentes-indicadores {
          grid-template-columns: 1fr;
        }

        .docentes-admin .docentes-tarjeta {
          padding: 16px;
        }

        .docentes-admin .docentes-campos,
        .docentes-admin .docentes-asignacion-grid {
          grid-template-columns: 1fr;
        }

        .docentes-admin .docentes-cursos-grid {
          grid-template-columns: repeat(2, minmax(0, 1fr));
        }

        .docentes-admin .docentes-asignacion-cabecera,
        .docentes-admin .docentes-listado-cabecera {
          align-items: stretch;
          flex-direction: column;
        }

        .docentes-admin .docentes-listado-cabecera input {
          max-width: none;
        }
      }
    </style>

    <section class="dashboard-seccion docentes-admin">

      <header class="docentes-hero">
        <span>COLEGIO ADVENTISTA SHALOM · SHALOM EXPO 2026</span>
        <h2>GESTIÓN DE DOCENTES</h2>
        <p>Administración académica, niveles, materias y asignaciones por curso.</p>
      </header>
${puedeEditar ? `
  <section class="docentes-tarjeta">
    <div class="docentes-asignacion" id="catalogo-academico-panel">

      <div class="docentes-asignacion-cabecera">
        <div>
          <h4>Catálogo académico institucional</h4>
          <p>Administra las materias y funciones disponibles para los tres niveles.</p>
        </div>
      </div>

      <div class="docentes-campos">
        <label>
          Nivel educativo
          <select id="catalogo-nivel">
            <option value="Inicial">Inicial</option>
            <option value="Primaria">Primaria</option>
            <option value="Secundaria">Secundaria</option>
          </select>
        </label>

        <label>
          Tipo de registro
          <select id="catalogo-tipo">
            <option value="materia">Materia académica</option>
            <option value="funcion">Función institucional</option>
          </select>
        </label>

        <label style="grid-column:1/-1">
          Nombre de la materia o función
          <input id="catalogo-nombre" type="text"
            placeholder="Ej. Matemática, Maestro/a de Grado, Asesor/a">
        </label>
      </div>

      <div class="docentes-acciones">
        <button type="button" class="docentes-boton"
          id="catalogo-guardar">
          + Agregar al catálogo
        </button>

        <button type="button"
          class="docentes-boton docentes-boton-claro"
          id="catalogo-cancelar" hidden>
          Cancelar edición
        </button>
      </div>

      <p id="catalogo-mensaje" class="docentes-mensaje"></p>

      <div style="margin-top:20px">
        <h4 style="color:#382052">Registros disponibles</h4>

        <div class="docentes-tabla-contenedor">
          <table style="min-width:650px">
            <thead>
              <tr>
                <th>NIVEL</th>
                <th>NOMBRE</th>
                <th>TIPO</th>
                <th>ESTADO</th>
                <th>ACCIONES</th>
              </tr>
            </thead>
            <tbody id="catalogo-tabla">
              <tr>
                <td colspan="5">Cargando catálogo...</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>

    </div>
  </section>
` : ''}
      <div class="docentes-indicadores">
        <div class="docentes-indicador">
          <div class="docentes-indicador-icono">♙</div>
          <div>
            <span>TOTAL DE DOCENTES</span>
            <strong id="docentes-total">0</strong>
            <small>Registros institucionales</small>
          </div>
        </div>

        <div class="docentes-indicador">
          <div class="docentes-indicador-icono verde">✓</div>
          <div>
            <span>DOCENTES ACTIVOS</span>
            <strong id="docentes-activos">0</strong>
            <small>Disponibles para proyectos</small>
          </div>
        </div>
      </div>

      ${puedeEditar ? `
        <section class="docentes-tarjeta docentes-formulario">

          <div class="docentes-seccion-titulo">
            <h3 id="docentes-form-titulo">Registrar nuevo docente</h3>
            <p>Registra una sola vez a cada docente y agrega todas sus asignaciones.</p>
          </div>

          <div class="docentes-campos">
            <label>
              Nombres
              <input id="docente-nombres" type="text"
                placeholder="Ej. María">
            </label>

            <label>
              Apellidos
              <input id="docente-apellidos" type="text"
                placeholder="Ej. Pérez Mamani">
            </label>

            <label>
              Tipo de docente
              <select id="docente-tipo">
                <option value="Docente de curso">Docente de curso / grado</option>
                <option value="Especialista">Docente especialista</option>
                <option value="Ambos">Docente de curso y especialista</option>
              </select>
            </label>
          </div>

                <h4>Asignaciones académicas</h4>
                <p>Selecciona nivel, materia y los cursos correspondientes.</p>
              </div>

              <button type="button" class="docentes-boton docentes-boton-claro"
                id="docente-agregar-asignacion">
                + Agregar nivel o materia
              </button>
            </div>

            <div id="docente-lista-asignaciones"></div>
          </div>

          <div class="docentes-acciones">
            <button type="button" class="docentes-boton"
              id="docente-guardar">
              Guardar docente
            </button>

            <button type="button"
              class="docentes-boton docentes-boton-claro"
              id="docente-cancelar" hidden>
              Cancelar edición
            </button>
          </div>

          <p id="docente-mensaje" class="docentes-mensaje"></p>
        </section>
      ` : `
        <section class="docentes-tarjeta">
          Consulta de docentes. No tienes permisos para modificar registros.
        </section>
      `}

      <section class="docentes-tarjeta docentes-listado">
        <div class="docentes-listado-cabecera">
          <div class="docentes-seccion-titulo">
            <h3>Directorio institucional de docentes</h3>
            <p>Consulta sus niveles, materias y cursos asignados.</p>
          </div>

          <input id="docentes-buscar" type="search"
            placeholder="Buscar docente, nivel o materia...">
            ${puedeEditar ? `
  <button type="button"
    class="docentes-boton docentes-boton-peligro"
    id="docentes-limpiar-todo">
    Limpiar todo
  </button>
` : ''}
        </div>

        <div class="docentes-tabla-contenedor">
          <table>
            <thead>
              <tr>
                <th>N.º</th>
                <th>DOCENTE</th>
                <th>TIPO</th>
                <th>NIVELES</th>
                <th>MATERIAS Y CURSOS</th>
                <th>ESTADO</th>
                ${puedeEditar ? '<th>ACCIONES</th>' : ''}
              </tr>
            </thead>

            <tbody id="docentes-tabla">
              <tr>
                <td colspan="${puedeEditar ? 7 : 6}">Cargando docentes...</td>
              </tr>
            </tbody>
          </table>
        </div>
      </section>

      <button id="docentes-volver"
        class="docentes-boton docentes-volver">
        ← Volver al inicio
      </button>

    </section>
  `;

  const listaAsignaciones = document.querySelector(
    '#docente-lista-asignaciones'
  );

  // =====================================================
  // CATÁLOGO ACADÉMICO COMPARTIDO
  // =====================================================

  async function cargarCatalogoAcademico() {
    const resultado = await getDocs(
      collection(db, 'catalogo_academico')
    );

    catalogoAcademico = resultado.docs.map(documento => ({
      id: documento.id,
      ...documento.data()
    }));

    catalogoAcademico.sort((a, b) =>
      (a.nombre || '').localeCompare(b.nombre || '', 'es')
    );

    mostrarCatalogoAcademico();
  }

  function mostrarCatalogoAcademico() {
    const tabla = document.querySelector('#catalogo-tabla');

    if (!tabla) return;

    if (!catalogoAcademico.length) {
      tabla.innerHTML = `
        <tr>
          <td colspan="5">
            Todavía no hay materias ni funciones registradas.
          </td>
        </tr>
      `;
      return;
    }

    tabla.innerHTML = catalogoAcademico.map(item => `
      <tr>
        <td>${escapar(item.nivel)}</td>
        <td><strong>${escapar(item.nombre)}</strong></td>
        <td>${item.tipo === 'funcion'
          ? 'Función institucional'
          : 'Materia académica'}</td>
        <td>
          <span class="docentes-estado ${
            item.activo === false ? 'inactivo' : 'activo'
          }">
            ${item.activo === false ? 'Inactivo' : 'Activo'}
          </span>
        </td>
        <td>
          <div class="docentes-acciones-tabla">
            <button type="button"
              data-catalogo-editar="${escapar(item.id)}">
              Editar
            </button>

            <button type="button"
              data-catalogo-estado="${escapar(item.id)}">
              ${item.activo === false ? 'Activar' : 'Desactivar'}
            </button>

            <button type="button"
              data-catalogo-eliminar="${escapar(item.id)}">
              Eliminar
            </button>
          </div>
        </td>
      </tr>
    `).join('');
  }

  function limpiarFormularioCatalogo() {
    catalogoEditandoId = null;

    document.querySelector('#catalogo-nivel').value = 'Inicial';
    document.querySelector('#catalogo-tipo').value = 'materia';
    document.querySelector('#catalogo-nombre').value = '';

    document.querySelector('#catalogo-guardar').textContent =
      '+ Agregar al catálogo';

    document.querySelector('#catalogo-cancelar').hidden = true;
  }

  async function guardarElementoCatalogo() {
    const boton = document.querySelector('#catalogo-guardar');
    const mensaje = document.querySelector('#catalogo-mensaje');

    const nivel = document.querySelector('#catalogo-nivel').value;
    const tipo = document.querySelector('#catalogo-tipo').value;
    const nombre = document.querySelector('#catalogo-nombre')
      .value.trim().replace(/\s+/g, ' ');

    if (!nombre) {
      mensaje.textContent = 'Escribe el nombre de la materia o función.';
      mensaje.style.color = '#b33d3d';
      return;
    }

    const normalizar = texto => String(texto || '')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toLowerCase()
      .replace(/\s+/g, ' ')
      .trim();

    const duplicado = catalogoAcademico.find(item =>
      item.id !== catalogoEditandoId &&
      item.nivel === nivel &&
      item.tipo === tipo &&
      normalizar(item.nombre) === normalizar(nombre)
    );

    if (duplicado) {
      mensaje.textContent =
        'Esta materia o función ya está registrada en ese nivel.';
      mensaje.style.color = '#b33d3d';
      return;
    }

    boton.disabled = true;

    try {
      const datos = {
        nivel,
        tipo,
        nombre,
        activo: catalogoEditandoId
          ? catalogoAcademico.find(x => x.id === catalogoEditandoId)?.activo !== false
          : true,
        actualizadoEn: serverTimestamp()
      };

      if (catalogoEditandoId) {
        await updateDoc(
          doc(db, 'catalogo_academico', catalogoEditandoId),
          datos
        );
      } else {
        datos.creadoEn = serverTimestamp();

        await addDoc(
          collection(db, 'catalogo_academico'),
          datos
        );
      }

      await cargarCatalogoAcademico();
      limpiarFormularioCatalogo();

      mensaje.textContent = 'Catálogo actualizado correctamente.';
      mensaje.style.color = '#21844b';

    } catch (error) {
      console.error('Error guardando catálogo:', error);

      mensaje.textContent =
        'No se pudo guardar. Revisa los permisos de Firebase.';
      mensaje.style.color = '#b33d3d';

    } finally {
      boton.disabled = false;
    }
  }


async function comprobarUsoCatalogo(item) {
  const [resultadoDocentes, resultadoProyectos] = await Promise.all([
    getDocs(collection(db, 'docentes')),
    getDocs(collection(db, 'proyectos'))
  ]);

  const normalizar = texto => String(texto || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();

  const nombreNormalizado = normalizar(item.nombre);

  // Comprobar asignaciones de docentes nuevas y antiguas.
  const usadoEnDocentes = resultadoDocentes.docs.some(documento => {
    const asignaciones = documento.data().asignaciones || [];

    return asignaciones.some(asignacion => {
      const mismoCatalogo = asignacion.catalogoId === item.id;

      const mismaMateriaAnterior =
        normalizar(asignacion.nivel) === normalizar(item.nivel) &&
        normalizar(asignacion.materia) === nombreNormalizado;

      return mismoCatalogo || mismaMateriaAnterior;
    });
  });

  // Comprobar proyectos nuevos y anteriores.
  const usadoEnProyectos = resultadoProyectos.docs.some(documento => {
    const proyecto = documento.data();

    const mismoCatalogo = proyecto.catalogoMateriaId === item.id;

    const mismaMateriaAnterior =
      normalizar(proyecto.nivel) === normalizar(item.nivel) &&
      normalizar(proyecto.area) === nombreNormalizado;

    return mismoCatalogo || mismaMateriaAnterior;
  });

  return usadoEnDocentes || usadoEnProyectos;
}

  async function gestionarAccionesCatalogo(evento) {
    const editar = evento.target.closest('[data-catalogo-editar]');
    const estado = evento.target.closest('[data-catalogo-estado]');
    const eliminar = evento.target.closest('[data-catalogo-eliminar]');

    if (editar) {
      const item = catalogoAcademico.find(x =>
        x.id === editar.dataset.catalogoEditar
      );

      if (!item) return;

      catalogoEditandoId = item.id;

      document.querySelector('#catalogo-nivel').value = item.nivel;
      document.querySelector('#catalogo-tipo').value = item.tipo;
      document.querySelector('#catalogo-nombre').value = item.nombre;

      document.querySelector('#catalogo-guardar').textContent =
        'Actualizar registro';

      document.querySelector('#catalogo-cancelar').hidden = false;

      document.querySelector('#catalogo-academico-panel')
        .scrollIntoView({ behavior: 'smooth', block: 'start' });
    }

    if (estado) {
      const item = catalogoAcademico.find(x =>
        x.id === estado.dataset.catalogoEstado
      );

      if (!item) return;

      const nuevoEstado = item.activo === false;

      if (!confirm(
        `¿Deseas ${nuevoEstado ? 'activar' : 'desactivar'} "${item.nombre}"?`
      )) return;

      try {
        await updateDoc(
          doc(db, 'catalogo_academico', item.id),
          {
            activo: nuevoEstado,
            actualizadoEn: serverTimestamp()
          }
        );

        await cargarCatalogoAcademico();

      } catch (error) {
        console.error(error);
        alert('No se pudo cambiar el estado del registro.');
      }
    }

    if (eliminar) {
      const item = catalogoAcademico.find(x =>
        x.id === eliminar.dataset.catalogoEliminar
      );

      if (!item) return;

      try {
        const estaEnUso = await comprobarUsoCatalogo(item);

        if (estaEnUso) {
          alert(
            'Esta materia o función ya está utilizada en docentes o proyectos. No se eliminará para proteger el historial. Puedes desactivarla.'
          );
          return;
        }

        if (!confirm(
          `¿Eliminar definitivamente "${item.nombre}" del nivel ${item.nivel}?`
        )) return;

        await deleteDoc(
          doc(db, 'catalogo_academico', item.id)
        );

        await cargarCatalogoAcademico();

        alert('Registro eliminado correctamente.');

      } catch (error) {
        console.error(error);
        alert('No se pudo comprobar o eliminar el registro.');
      }
    }
  }

  if (puedeEditar) {
    document.querySelector('#catalogo-guardar')
      .addEventListener('click', guardarElementoCatalogo);

    document.querySelector('#catalogo-cancelar')
      .addEventListener('click', limpiarFormularioCatalogo);

    document.querySelector('#catalogo-tabla')
      .addEventListener('click', gestionarAccionesCatalogo);
  }

  function obtenerCursosDeNivel(nivel) {
    return cursosDisponibles
      .filter(c => c.nivel === nivel)
      .map(c => c.curso);
  }

function agregarAsignacion(datos = {}) {
  const fila = document.createElement('div');
  fila.className = 'docentes-asignacion-fila';

  fila.innerHTML = `
    <div class="docentes-asignacion-grid">

      <label>
        Nivel educativo
        <select class="asignacion-nivel">
          <option value="">Seleccionar nivel</option>
          <option value="Inicial">Inicial</option>
          <option value="Primaria">Primaria</option>
          <option value="Secundaria">Secundaria</option>
        </select>
      </label>

      <label>
        Materia o función
        <select class="asignacion-materia" disabled>
          <option value="">Primero selecciona un nivel</option>
        </select>
      </label>

      <button type="button"
        class="docentes-boton docentes-boton-peligro asignacion-eliminar">
        Quitar
      </button>

    </div>

    <div class="docentes-cursos-cabecera">
      <strong>Cursos y paralelos disponibles</strong>
      <button type="button"
        class="docentes-boton docentes-boton-claro asignacion-todos">
        Seleccionar todos
      </button>
    </div>

    <div class="docentes-cursos-grid asignacion-lista-cursos">
      <div class="docentes-sin-cursos">
        Primero selecciona un nivel educativo.
      </div>
    </div>

    <small class="asignacion-contador"
      style="display:block;margin-top:10px;color:#806b91">
      0 cursos seleccionados
    </small>
  `;

  const selectorNivel = fila.querySelector('.asignacion-nivel');
  const campoMateria = fila.querySelector('.asignacion-materia');
  const listaCursos = fila.querySelector('.asignacion-lista-cursos');
  const contador = fila.querySelector('.asignacion-contador');

  selectorNivel.value = datos.nivel || '';

  function actualizarMaterias(nivel, materiaInicial = '', catalogoIdInicial = '') {
    const disponibles = catalogoAcademico.filter(item =>
      item.nivel === nivel && item.activo !== false
    );

    campoMateria.innerHTML = `
      <option value="">Seleccionar materia o función</option>
      ${disponibles.map(item => `
        <option
          value="${escapar(item.nombre)}"
          data-catalogo-id="${escapar(item.id)}"
          data-tipo="${escapar(item.tipo || 'materia')}"
          ${(
            catalogoIdInicial
              ? item.id === catalogoIdInicial
              : item.nombre === materiaInicial
          ) ? 'selected' : ''}>
          ${escapar(item.nombre)}
          ${item.tipo === 'funcion' ? '(Función)' : '(Materia)'}
        </option>
      `).join('')}
    `;

    // Conserva asignaciones antiguas que todavía no estén en el catálogo.
    if (
      materiaInicial &&
      !disponibles.some(item => item.nombre === materiaInicial)
    ) {
      const opcionAntigua = document.createElement('option');
      opcionAntigua.value = materiaInicial;
      opcionAntigua.textContent = `${materiaInicial} (Registro anterior)`;
      opcionAntigua.selected = true;
      opcionAntigua.dataset.catalogoId = '';
      opcionAntigua.dataset.tipo = datos.tipoAsignacion || 'materia';
      campoMateria.appendChild(opcionAntigua);
    }

    campoMateria.disabled = !nivel;
  }

  function seleccionados() {
    return [
      ...listaCursos.querySelectorAll('.asignacion-curso:checked')
    ].map(input => input.value);
  }

  function actualizarContador() {
    contador.textContent =
      `${seleccionados().length} cursos seleccionados`;
  }

  function dibujarCursos(nivel, seleccionInicial = []) {
    const cursos = obtenerCursosDeNivel(nivel);

    if (!nivel) {
      listaCursos.innerHTML = `
        <div class="docentes-sin-cursos">
          Primero selecciona un nivel educativo.
        </div>
      `;
      actualizarContador();
      return;
    }

    if (!cursos.length) {
      listaCursos.innerHTML = `
        <div class="docentes-sin-cursos">
          No se encontraron estudiantes activos para este nivel.
          Verifica los registros de estudiantes.
        </div>
      `;
      actualizarContador();
      return;
    }

    listaCursos.innerHTML = cursos.map(curso => `
      <label class="docente-curso-opcion">
        <input type="checkbox"
          class="asignacion-curso"
          value="${escapar(curso)}"
          ${seleccionInicial.includes(curso) ? 'checked' : ''}>
        <span>${escapar(curso)}</span>
      </label>
    `).join('');

    listaCursos.querySelectorAll('.asignacion-curso')
      .forEach(input => {
        input.addEventListener('change', actualizarContador);
      });

    actualizarContador();
  }

  selectorNivel.addEventListener('change', () => {
    actualizarMaterias(selectorNivel.value);
    dibujarCursos(selectorNivel.value, []);
  });

  campoMateria.addEventListener('change', () => {
    const opcion = campoMateria.selectedOptions[0];

    if (opcion) {
      opcion.dataset.catalogoId = opcion.dataset.catalogoId || '';
    }
  });

  fila.querySelector('.asignacion-todos')
    .addEventListener('click', () => {
      const checks = [
        ...listaCursos.querySelectorAll('.asignacion-curso')
      ];

      if (!checks.length) return;

      const todosMarcados = checks.every(input => input.checked);

      checks.forEach(input => {
        input.checked = !todosMarcados;
      });

      actualizarContador();

      fila.querySelector('.asignacion-todos').textContent =
        todosMarcados ? 'Seleccionar todos' : 'Quitar selección';
    });

  fila.querySelector('.asignacion-eliminar')
    .addEventListener('click', () => fila.remove());

  const cursosIniciales = Array.isArray(datos.cursos)
    ? datos.cursos
    : typeof datos.cursos === 'string'
      ? datos.cursos.split(',').map(x => x.trim()).filter(Boolean)
      : [];

  if (selectorNivel.value) {
    actualizarMaterias(
      selectorNivel.value,
      datos.materia || '',
      datos.catalogoId || ''
    );

    dibujarCursos(selectorNivel.value, cursosIniciales);
  }

  listaAsignaciones.appendChild(fila);
}

 
function obtenerAsignaciones() {
  return [
    ...listaAsignaciones.querySelectorAll('.docentes-asignacion-fila')
  ].map(fila => {
    const nivel = fila.querySelector('.asignacion-nivel').value;

    const selectorMateria = fila.querySelector('.asignacion-materia');
    const materia = selectorMateria.value.trim();

    const opcionSeleccionada = selectorMateria.selectedOptions[0];

    const catalogoId =
      opcionSeleccionada?.dataset.catalogoId || '';

    const tipoAsignacion =
      opcionSeleccionada?.dataset.tipo || 'materia';

    const cursos = [
      ...fila.querySelectorAll('.asignacion-curso:checked')
    ].map(input => input.value);

    return {
      nivel,
      materia,
      catalogoId,
      tipoAsignacion,
      cursos,
      cursosTexto: cursos.join(', ')
    };
  }).filter(a => a.nivel && a.materia && a.cursos.length);
}

  function limpiarFormulario() {
    docenteEditandoId = null;

    document.querySelector('#docente-nombres').value = '';
    document.querySelector('#docente-apellidos').value = '';
    document.querySelector('#docente-tipo').value = 'Docente de curso';

    listaAsignaciones.innerHTML = '';
    agregarAsignacion();

    document.querySelector('#docentes-form-titulo').textContent =
      'Registrar nuevo docente';

    document.querySelector('#docente-guardar').textContent =
      'Guardar docente';

    document.querySelector('#docente-cancelar').hidden = true;
    document.querySelector('#docente-mensaje').textContent = '';
  }

  function mostrarTabla(filtro = '') {
    const tabla = document.querySelector('#docentes-tabla');
    const busqueda = filtro.toLowerCase().trim();

    const filtrados = docentes.filter(docente => {
      const asignacionesTexto = (docente.asignaciones || [])
        .map(a => {
          const cursos = Array.isArray(a.cursos)
            ? a.cursos.join(' ')
            : a.cursosTexto || a.cursos || '';

          return `${a.nivel} ${a.materia} ${cursos}`;
        })
        .join(' ');

      return `${docente.nombreCompleto} ${docente.tipo} ${asignacionesTexto}`
        .toLowerCase()
        .includes(busqueda);
    });

    if (!filtrados.length) {
      tabla.innerHTML = `
        <tr>
          <td colspan="${puedeEditar ? 7 : 6}">
            No se encontraron docentes.
          </td>
        </tr>
      `;
      return;
    }

    tabla.innerHTML = filtrados.map((docente, indice) => {
      const asignaciones = docente.asignaciones || [];
      const niveles = [...new Set(asignaciones.map(a => a.nivel))];

      return `
        <tr>
          <td>${indice + 1}</td>

          <td>
            <strong>${escapar(docente.nombreCompleto)}</strong>
          </td>

          <td>${escapar(docente.tipo || 'No especificado')}</td>

          <td>${niveles.map(escapar).join(', ') || '—'}</td>

          <td>
            ${asignaciones.map(a => {
              const cursos = Array.isArray(a.cursos)
                ? a.cursos.join(', ')
                : a.cursosTexto || a.cursos || 'Asignación anterior';

              return `
                <div class="docente-asignacion-resumen">
                  <strong>${escapar(a.materia || 'Docente de grado')}</strong>
                  <small>${escapar(a.nivel)} · ${escapar(cursos)}</small>
                </div>
              `;
            }).join('') || '—'}
          </td>

          <td>
            <span class="docentes-estado ${
              docente.activo === false ? 'inactivo' : 'activo'
            }">
              ${docente.activo === false ? 'Inactivo' : 'Activo'}
            </span>
          </td>

          ${puedeEditar ? `
            <td>
              <div class="docentes-acciones-tabla">
                <button type="button"
                  data-docente-editar="${escapar(docente.id)}">
                  Editar
                </button>

                <button type="button"
                  data-docente-estado="${escapar(docente.id)}">
                  ${docente.activo === false ? 'Activar' : 'Desactivar'}
                </button>

                <button type="button"
                  data-docente-eliminar="${escapar(docente.id)}">
                  Eliminar
                </button>
              </div>
            </td>
          ` : ''}
        </tr>
      `;
    }).join('');
  }
async function limpiarTodosLosDocentes() {
  const boton = document.querySelector('#docentes-limpiar-todo');

  if (!puedeEditar || !boton) return;

  if (!docentes.length) {
    alert('No hay docentes registrados para eliminar.');
    return;
  }

  const cantidad = docentes.length;

  const primeraConfirmacion = confirm(
    `ATENCIÓN: Se eliminarán PERMANENTEMENTE ${cantidad} docentes de Firebase.\n\nEsta acción no se puede deshacer.\n\n¿Deseas continuar?`
  );

  if (!primeraConfirmacion) return;

  const segundaConfirmacion = confirm(
    `CONFIRMACIÓN FINAL:\n\n¿Autorizas eliminar definitivamente los ${cantidad} registros del directorio institucional de docentes?`
  );

  if (!segundaConfirmacion) return;

  boton.disabled = true;
  boton.textContent = 'Eliminando...';

  try {
    await Promise.all(
      docentes.map(docente =>
        deleteDoc(doc(db, 'docentes', docente.id))
      )
    );

    docenteEditandoId = null;

    await cargarDocentes();

    alert(`Se eliminaron ${cantidad} docentes correctamente.`);

  } catch (error) {
    console.error('Error eliminando docentes:', error);

    alert(
      'No se pudieron eliminar todos los docentes. Revisa los permisos de Firebase.'
    );

    await cargarDocentes();

  } finally {
    boton.disabled = false;
    boton.textContent = 'Limpiar todo';
  }
}
  async function cargarDocentes() {
    const resultado = await getDocs(collection(db, 'docentes'));

    docentes = resultado.docs.map(documento => ({
      id: documento.id,
      ...documento.data()
    }));

    docentes.sort((a, b) =>
      (a.nombreCompleto || '').localeCompare(
        b.nombreCompleto || '',
        'es'
      )
    );

    document.querySelector('#docentes-total').textContent =
      docentes.length;

    document.querySelector('#docentes-activos').textContent =
      docentes.filter(d => d.activo !== false).length;

    mostrarTabla(document.querySelector('#docentes-buscar').value);
  }

  if (puedeEditar) {
    agregarAsignacion();

    document.querySelector('#docente-agregar-asignacion')
      .addEventListener('click', () => agregarAsignacion());

    document.querySelector('#docente-cancelar')
      .addEventListener('click', limpiarFormulario);

    document.querySelector('#docente-guardar')
      .addEventListener('click', async () => {
        const boton = document.querySelector('#docente-guardar');
        const mensaje = document.querySelector('#docente-mensaje');

        const nombres = document.querySelector('#docente-nombres').value.trim();
        const apellidos = document.querySelector('#docente-apellidos').value.trim();
        const tipo = document.querySelector('#docente-tipo').value;

        const asignaciones = obtenerAsignaciones();

        if (!nombres || !apellidos) {
          mensaje.textContent = 'Completa los nombres y apellidos.';
          return;
        }

        if (!asignaciones.length) {
          mensaje.textContent =
            'Selecciona al menos un nivel y uno o varios cursos.';
          return;
        }

        const nombreCompleto = `${nombres} ${apellidos}`
          .replace(/\s+/g, ' ')
          .trim();

        const normalizar = texto => texto.normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .toLowerCase()
          .replace(/\s+/g, ' ')
          .trim();

        const duplicado = docentes.find(d =>
          normalizar(d.nombreCompleto || '') === normalizar(nombreCompleto)
          && d.id !== docenteEditandoId
        );

        if (duplicado) {
          mensaje.textContent =
            'Este docente ya está registrado. Edita su registro para agregar otros niveles o materias.';
          return;
        }

        const datos = {
          nombres,
          apellidos,
          nombreCompleto,
          tipo,
          asignaciones,
          activo: docenteEditandoId
            ? docentes.find(d => d.id === docenteEditandoId)?.activo !== false
            : true,
          actualizadoEn: serverTimestamp()
        };

        boton.disabled = true;
        boton.textContent = 'Guardando...';

        try {
          if (docenteEditandoId) {
            await updateDoc(
              doc(db, 'docentes', docenteEditandoId),
              datos
            );
          } else {
            datos.creadoEn = serverTimestamp();

            await addDoc(
              collection(db, 'docentes'),
              datos
            );
          }

          await cargarDocentes();
          limpiarFormulario();

          mensaje.textContent = 'Docente guardado correctamente.';
          mensaje.style.color = '#21844b';

        } catch (error) {
          console.error('Error guardando docente:', error);

          mensaje.textContent =
            'No se pudo guardar. Verifica los permisos de Firebase.';
          mensaje.style.color = '#b33d3d';

        } finally {
          boton.disabled = false;
          boton.textContent = docenteEditandoId
            ? 'Actualizar docente'
            : 'Guardar docente';
        }
      });

    document.querySelector('#docentes-tabla')
      .addEventListener('click', async evento => {
        const editar = evento.target.closest('[data-docente-editar]');
        const cambiarEstado = evento.target.closest('[data-docente-estado]');
        const eliminar = evento.target.closest('[data-docente-eliminar]');

        if (editar) {
          const docente = docentes.find(d =>
            d.id === editar.dataset.docenteEditar
          );

          if (!docente) return;

          docenteEditandoId = docente.id;

          document.querySelector('#docente-nombres').value =
            docente.nombres || '';

          document.querySelector('#docente-apellidos').value =
            docente.apellidos || '';

          document.querySelector('#docente-tipo').value =
            docente.tipo || 'Docente de curso';

          listaAsignaciones.innerHTML = '';

          (docente.asignaciones || []).forEach(a => {
            agregarAsignacion(a);
          });

          if (!docente.asignaciones?.length) {
            agregarAsignacion();
          }

          document.querySelector('#docentes-form-titulo').textContent =
            'Editar docente';

          document.querySelector('#docente-guardar').textContent =
            'Actualizar docente';

          document.querySelector('#docente-cancelar').hidden = false;

          document.querySelector('.docentes-formulario')
            .scrollIntoView({
              behavior: 'smooth',
              block: 'start'
            });
        }

        if (cambiarEstado) {
          const docente = docentes.find(d =>
            d.id === cambiarEstado.dataset.docenteEstado
          );

          if (!docente) return;

          const nuevoEstado = docente.activo === false;

          if (!confirm(
            `¿Deseas ${nuevoEstado ? 'activar' : 'desactivar'} a ${docente.nombreCompleto}?`
          )) return;

          try {
            await updateDoc(
              doc(db, 'docentes', docente.id),
              {
                activo: nuevoEstado,
                actualizadoEn: serverTimestamp()
              }
            );

            await cargarDocentes();

          } catch (error) {
            console.error(error);
            alert('No se pudo cambiar el estado del docente.');
          }
        }

        if (eliminar) {
          const docente = docentes.find(d =>
            d.id === eliminar.dataset.docenteEliminar
          );

          if (!docente) return;

          if (!confirm(
            `¿Eliminar definitivamente el registro de ${docente.nombreCompleto}?`
          )) return;

          try {
            await deleteDoc(doc(db, 'docentes', docente.id));
            await cargarDocentes();

            alert('Docente eliminado.');

          } catch (error) {
            console.error(error);
            alert('No se pudo eliminar el docente.');
          }
        }
      });
  }

  document.querySelector('#docentes-buscar')
    .addEventListener('input', evento => {
      mostrarTabla(evento.target.value);
    });
if (puedeEditar) {
  document.querySelector('#docentes-limpiar-todo')
    .addEventListener('click', limpiarTodosLosDocentes);
}
  document.querySelector('#docentes-volver')
    .addEventListener('click', () => {
      mostrarPanel(administrador?.nombre || 'Administrador');
    });

  try {
    await cargarCursosDesdeEstudiantes();
await cargarCatalogoAcademico();
await cargarDocentes();

  } catch (error) {
    console.error('Error cargando docentes o cursos:', error);

    document.querySelector('#docentes-tabla').innerHTML = `
      <tr>
        <td colspan="${puedeEditar ? 7 : 6}">
          Error al cargar datos. Revisa los permisos de Firebase.
        </td>
      </tr>
    `;

    const mensaje = document.querySelector('#docente-mensaje');

    if (mensaje) {
      mensaje.textContent =
        'No se pudieron cargar los cursos de estudiantes. Revisa Firebase.';
      mensaje.style.color = '#b33d3d';
    }
  }
}

async function mostrarProyectos() {
  await detenerEscanerQR();

  const contenido = document.querySelector('#contenido');

  const rolesPermitidos = [
    'administrador',
    'director',
    'coordinador',
    'secretaria'
  ];

  if (!rolesPermitidos.includes(administrador?.rol)) {
    contenido.innerHTML =
      '<h2>No tienes permiso para administrar proyectos.</h2>';
    return;
  }

  contenido.innerHTML = `
    <section class="proyectos-panel">

      <header class="eval-encabezado">
        <button id="volver-inicio"
          class="asistencia-btn asistencia-btn-principal">
          ← INICIO
        </button>

        <div>
          <span class="eval-etiqueta">SHALOM EXPO 2026</span>
          <h2>Administración de proyectos</h2>
          <p>Gestión de grupos, stands, docentes, estudiantes y códigos QR.</p>
        </div>

        <div class="eval-total-general">
          <strong id="proyectos-total">0</strong>
          <span>PROYECTOS</span>
        </div>
      </header>

      <section class="eval-tarjeta">
        <h3 id="proyecto-form-titulo">Registrar nuevo proyecto</h3>

        <form id="form-proyecto">
          <input type="hidden" id="proyecto-id">

          <div class="eval-grid">

            <div class="eval-campo-completo">
              <label>Nombre del proyecto</label>
              <input id="proyecto-nombre" type="text"
                placeholder="Título del proyecto" required>
            </div>

            <div>
              <label>Nivel educativo</label>
              <select id="proyecto-nivel" required>
                <option value="">Seleccionar nivel</option>
                <option value="Inicial">Inicial</option>
                <option value="Primaria">Primaria</option>
                <option value="Secundaria">Secundaria</option>
              </select>
            </div>

            <div>
              <label>Curso</label>
              <select id="proyecto-curso" required disabled>
                <option value="">Primero selecciona un nivel</option>
              </select>
            </div>

            <div>
              <label>Grupo automático</label>
              <input id="proyecto-grupo" type="text"
                placeholder="Se genera al guardar" readonly>
            </div>

            <div>
              <label>Stand automático</label>
              <input id="proyecto-stand" type="text"
                placeholder="Se asigna al guardar" readonly>
            </div>

            <div class="eval-campo-completo">
              <label>Área o materia</label>
              <select id="proyecto-area" required>
                <option value="">Seleccionar materia</option>
              </select>
              <small id="proyecto-ayuda-materia">
                Selecciona nivel y curso para ver las materias disponibles.
              </small>
            </div>

          </div>

          <div class="eval-tarjeta">
            <h3>Docentes participantes</h3>
            <p>Se sugieren docentes según el nivel, curso y materia. También puedes seleccionar otros manualmente.</p>

            <input type="search" id="proyecto-buscar-docente"
              placeholder="Buscar docente...">

            <div id="proyecto-docentes">
              <p>Selecciona nivel y curso.</p>
            </div>
          </div>

          <h3>Estudiantes participantes</h3>

          <input type="search" id="proyecto-buscar-estudiante"
            placeholder="Buscar estudiante...">

          <div id="proyecto-estudiantes"
            class="eval-lista-estudiantes">
            <p>Selecciona nivel y curso para mostrar estudiantes.</p>
          </div>

          <div class="eval-campo-completo">
            <label>Descripción del proyecto</label>
            <textarea id="proyecto-descripcion" rows="3"
              placeholder="Breve descripción del proyecto..."></textarea>
          </div>

          <p id="proyecto-mensaje"></p>

          <div class="eval-acciones">
            <button type="submit" class="eval-btn-principal">
              Guardar proyecto
            </button>

            <button type="button" id="proyecto-cancelar"
              class="eval-btn-secundario">
              Limpiar
            </button>
          </div>

        </form>
      </section>

      <section class="eval-tarjeta">
        <h3>Proyectos registrados</h3>

        <input type="search" id="proyecto-buscar"
          placeholder="Buscar por proyecto, materia, curso o stand...">
${administrador?.rol === 'administrador' ? `
  <button type="button" id="proyectos-limpiar-todo"
    class="eval-btn-secundario" style="margin:12px 0;">
    Limpiar todo
  </button>
` : ''}
        <div id="proyectos-listado">
          <p>Cargando proyectos...</p>
        </div>
      </section>

    </section>
  `;

  document.querySelector('#volver-inicio')
    .addEventListener('click', () => {
      mostrarPanel(administrador?.nombre || 'Usuario');
    });

  let estudiantesProyecto = [];
  let docentesProyecto = [];
  let proyectosRegistrados = [];
  let catalogoAcademicoProyecto = [];

  let seleccionEstudiantes = new Set();
  let seleccionDocentes = new Set();

  const formulario = document.querySelector('#form-proyecto');
  const mensaje = document.querySelector('#proyecto-mensaje');

  const nivelCampo = document.querySelector('#proyecto-nivel');
  const cursoCampo = document.querySelector('#proyecto-curso');
  const materiaCampo = document.querySelector('#proyecto-area');

  const materiasBase = [
    'Artes Plásticas',
    'Biología',
    'Ciencias Sociales',
    'Educación Física',
    'Filosofía',
    'Física',
    'Formación Cristiana',
    'Inglés',
    'Lenguaje',
    'Matemática',
    'Orientación Vocacional',
    'Psicología',
    'Química',
    'Robótica',
    'Otra'
  ];

  function normalizarTextoProyecto(valor) {
    return String(valor || '')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toLowerCase()
      .replace(/\s*\/\s*/g, '/')
      .trim();
  }

  function cursoDeEstudiante(estudiante) {
    const curso = String(
      estudiante.curso || estudiante.grado || ''
    ).trim();

    const paralelo = String(
      estudiante.paralelo || estudiante.seccion || ''
    ).trim();

    const turno = String(estudiante.turno || '').trim();

    let resultado = curso;

    if (
      paralelo &&
      !normalizarTextoProyecto(curso).includes(
        normalizarTextoProyecto(paralelo)
      )
    ) {
      resultado += ` ${paralelo}`;
    }

    if (
      turno &&
      !normalizarTextoProyecto(resultado).includes(
        normalizarTextoProyecto(turno)
      )
    ) {
      resultado += ` / ${turno}`;
    }

    return resultado.trim();
  }

  function cursosDeAsignacion(asignacion) {
    if (Array.isArray(asignacion.cursos)) {
      return asignacion.cursos.map(String);
    }

    if (typeof asignacion.cursos === 'string') {
      return asignacion.cursos
        .split(',')
        .map(curso => curso.trim())
        .filter(Boolean);
    }

    if (typeof asignacion.cursosTexto === 'string') {
      return asignacion.cursosTexto
        .split(',')
        .map(curso => curso.trim())
        .filter(Boolean);
    }

    return [];
  }

  function asignacionesDocente(docente) {
    return Array.isArray(docente.asignaciones)
      ? docente.asignaciones
      : [];
  }

  function nombreDocente(docente) {
    return docente.nombreCompleto ||
      `${docente.nombres || ''} ${docente.apellidos || ''}`.trim();
  }

  
async function cargarDatosProyecto() {
  const [
    respuestaEstudiantes,
    respuestaProyectos,
    respuestaDocentes,
    respuestaCatalogo
  ] = await Promise.all([
    getDocs(collection(db, 'estudiantes')),
    getDocs(collection(db, 'proyectos')),
    getDocs(collection(db, 'docentes')),
    getDocs(collection(db, 'catalogo_academico'))
  ]);

  estudiantesProyecto = respuestaEstudiantes.docs
    .map(documento => ({
      id: documento.id,
      ...documento.data()
    }))
    .filter(estudiante => estudiante.activo !== false);

  proyectosRegistrados = respuestaProyectos.docs.map(documento => ({
    id: documento.id,
    ...documento.data()
  }));

  docentesProyecto = respuestaDocentes.docs
    .map(documento => ({
      id: documento.id,
      ...documento.data()
    }))
    .filter(docente => docente.activo !== false);

  catalogoAcademicoProyecto = respuestaCatalogo.docs.map(documento => ({
    id: documento.id,
    ...documento.data()
  }));

  actualizarCursosProyecto();
  actualizarStandDisponible();
  renderizarProyectosRegistrados();
}

function actualizarCursosProyecto() {
  const nivel = nivelCampo.value;

  const cursos = [...new Set(
    estudiantesProyecto
      .filter(estudiante =>
        normalizarTextoProyecto(estudiante.nivel) ===
        normalizarTextoProyecto(nivel)
      )
      .map(cursoDeEstudiante)
      .filter(Boolean)
    )].sort((a, b) => a.localeCompare(b, 'es'));

  cursoCampo.innerHTML = `
    <option value="">Seleccionar curso</option>
    ${cursos.map(curso => `
      <option value="${escapar(curso)}">${escapar(curso)}</option>
    `).join('')}
  `;

  cursoCampo.disabled = !nivel;

  actualizarMateriasProyecto();
  renderizarDocentesProyecto();
  renderizarEstudiantesProyecto();
}
  function docentesCompatibles() {
    const nivel = nivelCampo.value;
    const curso = cursoCampo.value;
    const materia = materiaCampo.value;

    if (!nivel || !curso) return [];

    return docentesProyecto.map(docente => {
      const asignaciones = asignacionesDocente(docente);

      const compatible = asignaciones.some(asignacion => {
        const mismoNivel =
          normalizarTextoProyecto(asignacion.nivel) ===
          normalizarTextoProyecto(nivel);

        const cursos = cursosDeAsignacion(asignacion);

        const mismoCurso = cursos.some(item =>
          normalizarTextoProyecto(item) ===
          normalizarTextoProyecto(curso)
        );

        const mismaMateria = !materia ||
          normalizarTextoProyecto(asignacion.materia) ===
          normalizarTextoProyecto(materia);

        return mismoNivel && mismoCurso && mismaMateria;
      });

      return {
        ...docente,
        compatible
      };
    }).sort((a, b) => {
      if (a.compatible !== b.compatible) {
        return a.compatible ? -1 : 1;
      }

      return nombreDocente(a).localeCompare(nombreDocente(b), 'es');
    });
  }

 
function actualizarMateriasProyecto(
  materiaInicial = materiaCampo.value,
  catalogoIdInicial = ''
) {
  const nivel = nivelCampo.value;

  const materiasDisponibles = catalogoAcademicoProyecto
    .filter(item =>
      normalizarTextoProyecto(item.nivel) ===
        normalizarTextoProyecto(nivel) &&
      item.tipo === 'materia' &&
      item.activo !== false
    )
    .sort((a, b) =>
      a.nombre.localeCompare(b.nombre, 'es')
    );

  materiaCampo.innerHTML = `
    <option value="">Seleccionar materia</option>

    ${materiasDisponibles.map(item => `
      <option
        value="${escapar(item.nombre)}"
        data-catalogo-id="${escapar(item.id)}">
        ${escapar(item.nombre)}
      </option>
    `).join('')}
  `;

  // Restaurar la materia seleccionada al editar.
  if (materiaInicial) {
    const encontrada = materiasDisponibles.find(item =>
      catalogoIdInicial
        ? item.id === catalogoIdInicial
        : normalizarTextoProyecto(item.nombre) ===
          normalizarTextoProyecto(materiaInicial)
    );

    if (encontrada) {
      materiaCampo.value = encontrada.nombre;
    } else {
      // Conserva materias de proyectos anteriores
      // que ya no aparecen en el catálogo activo.
      const opcionAnterior = document.createElement('option');

      opcionAnterior.value = materiaInicial;
      opcionAnterior.textContent = `${materiaInicial} (Registro anterior)`;
      opcionAnterior.dataset.catalogoId = '';
      opcionAnterior.selected = true;

      materiaCampo.appendChild(opcionAnterior);
    }
  }

  materiaCampo.disabled = !nivel;

  const ayuda = document.querySelector('#proyecto-ayuda-materia');

  if (!nivel) {
    ayuda.textContent = 'Primero selecciona un nivel educativo.';
  } else if (!materiasDisponibles.length) {
    ayuda.textContent =
      'No hay materias activas para este nivel. Agrégalas desde DOCENTES > Catálogo académico.';
  } else {
    ayuda.textContent =
      'Materias activas del catálogo académico para el nivel seleccionado.';
  }

  renderizarDocentesProyecto();
}

  function renderizarDocentesProyecto() {
    const lista = document.querySelector('#proyecto-docentes');

    if (!nivelCampo.value || !cursoCampo.value) {
      lista.innerHTML = '<p>Selecciona nivel y curso para ver docentes.</p>';
      return;
    }

    const busqueda = normalizarTextoProyecto(
      document.querySelector('#proyecto-buscar-docente').value
    );

    const docentes = docentesCompatibles().filter(docente =>
      normalizarTextoProyecto(nombreDocente(docente)).includes(busqueda)
    );

    if (!docentes.length) {
      lista.innerHTML = '<p>No se encontraron docentes.</p>';
      return;
    }

    lista.innerHTML = docentes.map(docente => `
      <label class="eval-estudiante-item">
        <input type="checkbox"
          data-docente-proyecto="${escapar(docente.id)}"
          ${seleccionDocentes.has(docente.id) ? 'checked' : ''}>

        <span>
          ${escapar(nombreDocente(docente))}
          <small>
            ${docente.compatible
              ? 'Docente sugerido para este curso'
              : 'Selección manual'}
          </small>
        </span>
      </label>
    `).join('');
  }

  function renderizarEstudiantesProyecto() {
    const lista = document.querySelector('#proyecto-estudiantes');

    if (!nivelCampo.value || !cursoCampo.value) {
      lista.innerHTML =
        '<p>Selecciona nivel y curso para mostrar estudiantes.</p>';
      return;
    }

    const busqueda = normalizarTextoProyecto(
      document.querySelector('#proyecto-buscar-estudiante').value
    );

    const filtrados = estudiantesProyecto.filter(estudiante => {
      const mismoNivel =
        normalizarTextoProyecto(estudiante.nivel) ===
        normalizarTextoProyecto(nivelCampo.value);

      const mismoCurso =
        normalizarTextoProyecto(cursoDeEstudiante(estudiante)) ===
        normalizarTextoProyecto(cursoCampo.value);

      const nombre = nombreCompletoEvaluacion(estudiante);

      return mismoNivel && mismoCurso &&
        normalizarTextoProyecto(nombre).includes(busqueda);
    });

    if (!filtrados.length) {
      lista.innerHTML = '<p>No hay estudiantes en este curso.</p>';
      return;
    }

    lista.innerHTML = filtrados.map(estudiante => `
      <label class="eval-estudiante-item">
        <input type="checkbox"
          data-estudiante-proyecto="${escapar(estudiante.id)}"
          ${seleccionEstudiantes.has(estudiante.id) ? 'checked' : ''}>

        <span>
          ${escapar(nombreCompletoEvaluacion(estudiante))}
          <small>
            ${escapar(cursoDeEstudiante(estudiante))}
          </small>
        </span>
      </label>
    `).join('');
  }

  function obtenerSiguienteStand() {
    const ocupados = new Set(
      proyectosRegistrados
        .filter(proyecto =>
          proyecto.id !== document.querySelector('#proyecto-id').value
        )
        .map(proyecto => String(proyecto.stand || '').trim())
    );

    let numero = 1;

    while (ocupados.has(String(numero).padStart(2, '0'))) {
      numero++;
    }

    return String(numero).padStart(2, '0');
  }

  function actualizarStandDisponible() {
    const proyectoId = document.querySelector('#proyecto-id').value;
    const proyectoActual = proyectosRegistrados.find(
      proyecto => proyecto.id === proyectoId
    );

    document.querySelector('#proyecto-stand').value =
      proyectoActual?.stand || obtenerSiguienteStand();
  }

  function codigoGrupo(proyectoId) {
    return `GRUPO-${proyectoId.slice(0, 8).toUpperCase()}`;
  }
async function limpiarTodosLosProyectos() {
  const boton = document.querySelector('#proyectos-limpiar-todo');

  if (administrador?.rol !== 'administrador' || !boton) return;

  if (!proyectosRegistrados.length) {
    alert('No hay proyectos registrados para eliminar.');
    return;
  }

  const cantidad = proyectosRegistrados.length;

  const primeraConfirmacion = confirm(
    `ATENCIÓN: Se eliminarán PERMANENTEMENTE ${cantidad} proyectos de Firebase.\n\nEsta acción no se puede deshacer.\n\n¿Deseas continuar?`
  );

  if (!primeraConfirmacion) return;

  const segundaConfirmacion = confirm(
    `CONFIRMACIÓN FINAL:\n\n¿Autorizas eliminar definitivamente los ${cantidad} proyectos registrados?`
  );

  if (!segundaConfirmacion) return;

  boton.disabled = true;
  boton.textContent = 'Eliminando...';

  try {
    await Promise.all(
      proyectosRegistrados.map(proyecto =>
        deleteDoc(doc(db, 'proyectos', proyecto.id))
      )
    );

    await cargarDatosProyecto();

    alert(`Se eliminaron ${cantidad} proyectos correctamente.`);

  } catch (error) {
    console.error('Error eliminando proyectos:', error);

    alert(
      'No se pudieron eliminar todos los proyectos. Revisa los permisos de Firebase.'
    );

    await cargarDatosProyecto();

  } finally {
    boton.disabled = false;
    boton.textContent = 'Limpiar todo';
  }
}
  function renderizarProyectosRegistrados() {
    const lista = document.querySelector('#proyectos-listado');

    const busqueda = normalizarTextoProyecto(
      document.querySelector('#proyecto-buscar').value
    );

    const filtrados = proyectosRegistrados.filter(proyecto =>
      normalizarTextoProyecto(`
        ${proyecto.nombre || ''}
        ${proyecto.area || ''}
        ${proyecto.nivel || ''}
        ${proyecto.curso || ''}
        ${proyecto.stand || ''}
        ${proyecto.grupo_codigo || ''}
      `).includes(busqueda)
    );

    document.querySelector('#proyectos-total').textContent =
      proyectosRegistrados.length;

    if (!filtrados.length) {
      lista.innerHTML = '<p>No hay proyectos registrados.</p>';
      return;
    }

    lista.innerHTML = filtrados.map(proyecto => `
      <article class="eval-tarjeta">
        <h3>${escapar(proyecto.nombre || '')}</h3>

        <p><strong>Nivel:</strong> ${escapar(proyecto.nivel || 'No registrado')}</p>
        <p><strong>Curso:</strong> ${escapar(proyecto.curso || 'No registrado')}</p>
        <p><strong>Grupo:</strong> ${escapar(proyecto.grupo_codigo || 'Proyecto anterior')}</p>
        <p><strong>Materia:</strong> ${escapar(proyecto.area || '')}</p>
        <p><strong>Stand:</strong> ${escapar(proyecto.stand || '')}</p>

        <p><strong>Docentes:</strong>
          ${(proyecto.docentes || [])
            .map(docente => escapar(docente.nombre || ''))
            .join(', ') || 'Sin docentes registrados'}
        </p>

        <p><strong>Estudiantes:</strong>
          ${(proyecto.estudiantes || [])
            .map(estudiante => escapar(estudiante.nombre || ''))
            .join(', ')}
        </p>

        <div class="eval-acciones">
          <button type="button"
            data-proyecto-qr="${escapar(proyecto.id)}">
            Ver código QR
          </button>

          <button type="button"
            data-proyecto-editar="${escapar(proyecto.id)}">
            Editar
          </button>

          <button type="button"
            data-proyecto-eliminar="${escapar(proyecto.id)}">
            Eliminar
          </button>
        </div>
      </article>
    `).join('');
  }

  function limpiarFormularioProyecto() {
    formulario.reset();

    seleccionEstudiantes.clear();
    seleccionDocentes.clear();

    document.querySelector('#proyecto-id').value = '';
    document.querySelector('#proyecto-grupo').value = '';

    document.querySelector('#proyecto-form-titulo').textContent =
      'Registrar nuevo proyecto';

    actualizarCursosProyecto();
    actualizarStandDisponible();

    document.querySelector('#proyecto-buscar-docente').value = '';
    document.querySelector('#proyecto-buscar-estudiante').value = '';

    renderizarDocentesProyecto();
    renderizarEstudiantesProyecto();

    mensaje.textContent = '';
  }

  nivelCampo.addEventListener('change', () => {
    seleccionEstudiantes.clear();
    seleccionDocentes.clear();
    actualizarCursosProyecto();
  });

  cursoCampo.addEventListener('change', () => {
    seleccionEstudiantes.clear();
    seleccionDocentes.clear();
    actualizarMateriasProyecto();
    renderizarEstudiantesProyecto();
  });

  materiaCampo.addEventListener('change', () => {
    renderizarDocentesProyecto();
  });

  document.querySelector('#proyecto-buscar-docente')
    .addEventListener('input', renderizarDocentesProyecto);

  document.querySelector('#proyecto-buscar-estudiante')
    .addEventListener('input', renderizarEstudiantesProyecto);

  document.querySelector('#proyecto-buscar')
    .addEventListener('input', renderizarProyectosRegistrados);
if (administrador?.rol === 'administrador') {
  document.querySelector('#proyectos-limpiar-todo')
    .addEventListener('click', limpiarTodosLosProyectos);
}
  document.querySelector('#proyecto-docentes')
    .addEventListener('change', evento => {
      const campo = evento.target.closest('[data-docente-proyecto]');
      if (!campo) return;

      if (campo.checked) {
        seleccionDocentes.add(campo.dataset.docenteProyecto);
      } else {
        seleccionDocentes.delete(campo.dataset.docenteProyecto);
      }
    });

  document.querySelector('#proyecto-estudiantes')
    .addEventListener('change', evento => {
      const campo = evento.target.closest('[data-estudiante-proyecto]');
      if (!campo) return;

      if (campo.checked) {
        seleccionEstudiantes.add(campo.dataset.estudianteProyecto);
      } else {
        seleccionEstudiantes.delete(campo.dataset.estudianteProyecto);
      }
    });

  document.querySelector('#proyecto-cancelar')
    .addEventListener('click', limpiarFormularioProyecto);

  formulario.addEventListener('submit', async evento => {
    evento.preventDefault();

    const boton = formulario.querySelector('[type="submit"]');
    boton.disabled = true;
    mensaje.textContent = 'Guardando proyecto...';

    try {
      const proyectoId = document.querySelector('#proyecto-id').value;

      const nombre = document.querySelector('#proyecto-nombre')
        .value.trim();

      const nivel = nivelCampo.value;
      const curso = cursoCampo.value;
      const area = materiaCampo.value;
      const catalogoMateriaId =
  materiaCampo.selectedOptions[0]?.dataset.catalogoId || '';

      const descripcion = document.querySelector('#proyecto-descripcion')
        .value.trim();

      if (!nivel || !curso || !area) {
        throw new Error('Completa el nivel, curso y materia.');
      }

      const estudiantesSeleccionados = estudiantesProyecto
        .filter(estudiante => seleccionEstudiantes.has(estudiante.id))
        .map(estudiante => ({
          id: estudiante.id,
          nombre: nombreCompletoEvaluacion(estudiante)
        }));

      const docentesSeleccionados = docentesProyecto
        .filter(docente => seleccionDocentes.has(docente.id))
        .map(docente => ({
          id: docente.id,
          nombre: nombreDocente(docente)
        }));

      if (!estudiantesSeleccionados.length) {
        throw new Error('Selecciona al menos un estudiante.');
      }

      if (!docentesSeleccionados.length) {
        throw new Error('Selecciona al menos un docente participante.');
      }

      const proyectoAnterior = proyectosRegistrados.find(
        proyecto => proyecto.id === proyectoId
      );

      const stand = proyectoAnterior?.stand || obtenerSiguienteStand();

      const datos = {
        nombre,
        nivel,
        curso,
        grupo_codigo: proyectoAnterior?.grupo_codigo || '',
        area,
        catalogoMateriaId,  
        stand,
        descripcion,
        docentes: docentesSeleccionados,
        estudiantes: estudiantesSeleccionados,
        activo: proyectoAnterior?.activo !== false
      };

      if (proyectoId) {
        datos.grupo_codigo =
          proyectoAnterior?.grupo_codigo || codigoGrupo(proyectoId);

        await updateDoc(
          doc(db, 'proyectos', proyectoId),
          datos
        );

      } else {
        const referencia = await addDoc(
          collection(db, 'proyectos'),
          datos
        );

        const codigoQR = `SHALOM-PROYECTO:${referencia.id}`;

        await updateDoc(referencia, {
          grupo_codigo: codigoGrupo(referencia.id),
          qr_codigo: codigoQR
        });
      }

      mensaje.textContent = 'Proyecto guardado correctamente.';

      limpiarFormularioProyecto();
      await cargarDatosProyecto();

      mensaje.textContent = 'Proyecto guardado correctamente.';

    } catch (error) {
      console.error('Error al guardar proyecto:', error);
      mensaje.textContent =
        error.message || 'No se pudo guardar el proyecto.';
    } finally {
      boton.disabled = false;
    }
  });

  document.querySelector('#proyectos-listado')
    .addEventListener('click', async evento => {

      const botonQR = evento.target.closest('[data-proyecto-qr]');
      const botonEditar = evento.target.closest('[data-proyecto-editar]');
      const botonEliminar = evento.target.closest('[data-proyecto-eliminar]');

      if (botonQR) {
        const proyecto = proyectosRegistrados.find(
          item => item.id === botonQR.dataset.proyectoQr
        );

        if (proyecto) await mostrarQRProyecto(proyecto);
      }

      if (botonEditar) {
        const proyecto = proyectosRegistrados.find(
          item => item.id === botonEditar.dataset.proyectoEditar
        );

        if (!proyecto) return;

        seleccionEstudiantes = new Set(
          (proyecto.estudiantes || []).map(estudiante => estudiante.id)
        );

        seleccionDocentes = new Set(
          (proyecto.docentes || []).map(docente => docente.id)
        );

        document.querySelector('#proyecto-id').value = proyecto.id;
        document.querySelector('#proyecto-nombre').value =
          proyecto.nombre || '';

        document.querySelector('#proyecto-descripcion').value =
          proyecto.descripcion || '';

        nivelCampo.value = proyecto.nivel || '';

        if (!proyecto.nivel && proyecto.estudiantes?.length) {
          const primerEstudiante = estudiantesProyecto.find(
            estudiante => estudiante.id === proyecto.estudiantes[0].id
          );

          if (primerEstudiante) {
            nivelCampo.value = primerEstudiante.nivel || '';
          }
        }

        actualizarCursosProyecto();

        cursoCampo.value = proyecto.curso || '';

        if (!proyecto.curso && proyecto.estudiantes?.length) {
          const primerEstudiante = estudiantesProyecto.find(
            estudiante => estudiante.id === proyecto.estudiantes[0].id
          );

          if (primerEstudiante) {
            cursoCampo.value = cursoDeEstudiante(primerEstudiante);
          }
        }

        actualizarMateriasProyecto(
  proyecto.area || '',
  proyecto.catalogoMateriaId || ''
);

document.querySelector('#proyecto-stand').value =
  proyecto.stand || '';

        document.querySelector('#proyecto-grupo').value =
          proyecto.grupo_codigo || 'Se conservará el grupo anterior';

        renderizarDocentesProyecto();
        renderizarEstudiantesProyecto();

        document.querySelector('#proyecto-form-titulo').textContent =
          'Editar proyecto';

        document.querySelector('#proyecto-form-titulo')
          .scrollIntoView({ behavior: 'smooth' });
      }

      if (botonEliminar) {
        const id = botonEliminar.dataset.proyectoEliminar;

        if (!confirm('¿Seguro que deseas eliminar este proyecto?')) return;

        try {
          await deleteDoc(doc(db, 'proyectos', id));
          await cargarDatosProyecto();
        } catch (error) {
          console.error(error);
          alert('No se pudo eliminar el proyecto.');
        }
      }
    });

  async function mostrarQRProyecto(proyecto) {
    try {
      const codigo = proyecto.qr_codigo ||
        `SHALOM-PROYECTO:${proyecto.id}`;

      const imagen = await QRCode.toDataURL(codigo, {
        width: 320,
        margin: 2
      });

      const anterior = document.querySelector('#modal-qr-proyecto');
      if (anterior) anterior.remove();

      const modal = document.createElement('div');
      modal.id = 'modal-qr-proyecto';

      modal.innerHTML = `
        <div style="
          position:fixed; inset:0; z-index:9999;
          background:rgba(0,0,0,.75);
          display:flex; align-items:center; justify-content:center;
          padding:20px;">

          <div style="
            background:white; color:#222; border-radius:18px;
            padding:25px; text-align:center; max-width:380px; width:100%;">

            <h2>${escapar(proyecto.nombre)}</h2>
            <p>STAND ${escapar(proyecto.stand || '')}</p>

            <img src="${imagen}" alt="QR del proyecto"
              style="width:260px; max-width:100%;">

            <p>SHALOM EXPO 2026</p>

            <button type="button" id="cerrar-qr-proyecto">
              Cerrar
            </button>
          </div>
        </div>
      `;

      document.body.appendChild(modal);

      modal.querySelector('#cerrar-qr-proyecto')
        .addEventListener('click', () => modal.remove());

    } catch (error) {
      console.error(error);
      alert('No se pudo generar el código QR.');
    }
  }

  try {
    await cargarDatosProyecto();
  } catch (error) {
    console.error('Error al cargar proyectos:', error);

    document.querySelector('#proyectos-listado').innerHTML =
      '<p>No se pudieron cargar los datos. Verifica Firebase.</p>';
  }
}

// =====================================================
// ACCESO Y REGISTRO DEL PERSONAL STAFF
// =====================================================

async function iniciarAccesoStaff() {

  app.innerHTML = `
    <main class="jurado-page">
      <section class="jurado-card">

        <header class="jurado-header">
          <div class="jurado-marca">
            <span class="jurado-marca-linea"></span>
            <span>COLEGIO ADVENTISTA SHALOM</span>
          </div>

          <p class="jurado-evento">PLATAFORMA INSTITUCIONAL</p>
          <h1>SHALOM EXPO <span>2026</span></h1>
          <p class="jurado-subtitulo">Acceso del personal STAFF</p>
        </header>

        <div class="jurado-separador"></div>

        <section class="jurado-formulario">
          <div class="jurado-introduccion">
            <span class="jurado-paso">IDENTIFICACIÓN PERSONAL</span>
            <h2>Bienvenido al equipo STAFF</h2>
            <p>
              Escribe tus nombres y apellidos para ingresar
              al sistema de control de la feria.
            </p>
          </div>

          <form id="form-registro-staff" class="jurado-form">

            <div class="jurado-input-group">
              <label for="staff-nombre">Nombre completo</label>
              <input
                id="staff-nombre"
                type="text"
                placeholder="Ej. Juan Pérez López"
                autocomplete="name"
                required
              >
            </div>

            <button type="submit" class="jurado-btn-principal">
              INGRESAR AL SISTEMA →
            </button>

            <p id="mensaje-staff" class="jurado-mensaje" role="status"></p>
          </form>
        </section>

        <div class="jurado-versiculo">
          <div class="jurado-versiculo-linea"></div>
          <p>
            “Y todo lo que hagáis, hacedlo de corazón,
            como para el Señor”.
          </p>
          <span>COLOSENSES 3:23</span>
        </div>

        <footer class="jurado-footer">
          SHALOM EXPO 2026 · PERSONAL STAFF
        </footer>

      </section>
    </main>
  `
  try {

    await authStaff.authStateReady()

    let usuario = authStaff.currentUser

    if (!usuario) {
      const credencial = await signInAnonymously(authStaff)
      usuario = credencial.user
    }

    const referencia = doc(dbStaff, 'staff', usuario.uid)
    const existente = await getDoc(referencia)

    // SESIÓN STAFF YA REGISTRADA
    if (existente.exists()) {

      staffActual = {
        uid: usuario.uid,
        ...existente.data()
      }

      if (staffActual.activo !== true) {
        throw new Error('Esta sesión STAFF está desactivada.')
      }

      mostrarPanelStaff()
      return
    }

    // REGISTRO DE NUEVO PERSONAL
    const formulario = document.querySelector('#form-registro-staff')

    formulario.addEventListener('submit', async (evento) => {

      evento.preventDefault()

      const boton = evento.currentTarget.querySelector('button')
      const mensaje = document.querySelector('#mensaje-staff')

      const nombreCompleto = document
        .querySelector('#staff-nombre')
        .value
        .trim()
        .replace(/\s+/g, ' ')

      if (nombreCompleto.split(' ').length < 2) {
        mensaje.textContent = 'Escribe tus nombres y apellidos completos.'
        return
      }

      boton.disabled = true
      boton.textContent = 'REGISTRANDO...'

      try {

        const datosStaff = {
          uid: usuario.uid,
          nombreCompleto: nombreCompleto,
          activo: true,
          creadoEn: new Date().toISOString()
        }

        await setDoc(referencia, datosStaff)

        staffActual = datosStaff

        mostrarPanelStaff()

      } catch (error) {

        console.error('Error al registrar STAFF:', error)

        mensaje.textContent =
          'No se pudo completar el registro. Verifica los permisos de Firebase.'

        boton.disabled = false
        boton.textContent = 'INTENTAR NUEVAMENTE →'
      }
    })

  } catch (error) {

    console.error('Error de acceso STAFF:', error)

    app.innerHTML = `
      <main class="inicio">
        <section class="bienvenida">
          <h2>No se pudo iniciar el acceso STAFF</h2>
          <p>${escapar(error.message || 'Verifica tu conexión e intenta nuevamente.')}</p>
        </section>
      </main>
    `
  }
}


// =====================================================
// PANEL PRINCIPAL DEL PERSONAL STAFF
// =====================================================


function mostrarPanelStaff() {

  if (!staffActual) {
    iniciarAccesoStaff()
    return
  }

  app.innerHTML = `
    <main class="jurado-page">
      <section class="jurado-card">

        <header class="jurado-header">
          <div class="jurado-marca">
            <span class="jurado-marca-linea"></span>
            <span>COLEGIO ADVENTISTA SHALOM</span>
          </div>

          <p class="jurado-evento">PLATAFORMA INSTITUCIONAL</p>
          <h1>SHALOM EXPO <span>2026</span></h1>
          <p class="jurado-subtitulo">Panel de control STAFF</p>
        </header>

        <div class="jurado-separador"></div>

        <section class="jurado-formulario">

          <div class="jurado-introduccion">
            <span class="jurado-paso">SESIÓN ACTIVA</span>
            <h2>¡Bienvenido al equipo STAFF!</h2>
            <p>Personal encargado del control de la feria.</p>
            <h3>${escapar(staffActual.nombreCompleto)}</h3>
          </div>

          <div class="staff-panel-acciones">

            <button
              type="button"
              id="staff-escanear-proyecto"
              class="jurado-btn-principal">
              <span>▦</span>
              REGISTRAR VISITA A STANDS
            </button>

            <button
              type="button"
              id="staff-registrar-asistencia"
              class="jurado-btn-principal">
              <span>◉</span>
              REGISTRAR ASISTENCIA
            </button>

            <p class="jurado-mensaje">
              Selecciona la operación que deseas realizar.
            </p>

          </div>

        </section>

        <div class="jurado-versiculo">
          <div class="jurado-versiculo-linea"></div>
          <p>
            “Y todo lo que hagáis, hacedlo de corazón,
            como para el Señor”.
          </p>
          <span>COLOSENSES 3:23</span>
        </div>

        <footer class="jurado-footer">
          SHALOM EXPO 2026 · PERSONAL STAFF
        </footer>

      </section>
    </main>
  `

// =====================================================
// ESCÁNER DE VISITAS A STANDS PARA PERSONAL STAFF
// =====================================================

async function mostrarEscanerProyectoStaff() {

  if (!staffActual) {
    iniciarAccesoStaff()
    return
  }

  let lectorQRProyectoStaff = null
  let procesandoQR = false

  app.innerHTML = `
    <main class="jurado-page">
      <section class="jurado-card">

        <header class="jurado-header">
          <div class="jurado-marca">
            <span class="jurado-marca-linea"></span>
            <span>COLEGIO ADVENTISTA SHALOM</span>
          </div>

          <p class="jurado-evento">CONTROL DE VISITAS A STANDS</p>
          <h1>SHALOM EXPO <span>2026</span></h1>
          <p class="jurado-subtitulo">Registro de visitantes</p>
        </header>

        <div class="jurado-separador"></div>

        <section class="jurado-formulario">

          <div class="jurado-introduccion">
            <span class="jurado-paso">PERSONAL STAFF</span>
            <h2>Escanear QR del proyecto</h2>
            <p>
              Enfoca el código QR del stand que está visitando
              el estudiante o visitante.
            </p>
            <h3>${escapar(staffActual.nombreCompleto)}</h3>
          </div>

          <div id="lector-qr-proyecto-staff"></div>

          <div id="resultado-proyecto-staff"
            class="jurado-mensaje"
            role="status">
            Esperando lectura del código QR...
          </div>

          <button
            type="button"
            id="staff-volver-panel-proyectos"
            class="jurado-btn-principal">
            ← VOLVER AL PANEL STAFF
          </button>

        </section>

        <footer class="jurado-footer">
          SHALOM EXPO 2026 · PERSONAL STAFF
        </footer>

      </section>
    </main>
  `

  async function detenerEscanerProyecto() {

    if (lectorQRProyectoStaff) {

      try {
        if (lectorQRProyectoStaff.isScanning) {
          await lectorQRProyectoStaff.stop()
        }

        await lectorQRProyectoStaff.clear()

      } catch (error) {
        console.warn('No se pudo cerrar el escáner:', error)
      }

      lectorQRProyectoStaff = null
    }
  }

  document
    .querySelector('#staff-volver-panel-proyectos')
    .addEventListener('click', async () => {
      await detenerEscanerProyecto()
      mostrarPanelStaff()
    })

// REGISTRAR VISITA AL PROYECTO AUTOMÁTICAMENTE
async function mostrarSeleccionVisitante(proyecto) {

  const resultado = document.querySelector(
    '#resultado-proyecto-staff'
  );

  if (!resultado) return;

  if (!staffActual || !staffActual.uid) {
    resultado.textContent = 'No se pudo identificar al personal STAFF.';
    return;
  }

  resultado.innerHTML = `
    <h2>REGISTRANDO VISITA...</h2>
    <p>${escapar(proyecto.nombre || 'Proyecto sin nombre')}</p>
  `;

  try {

    const ahora = new Date();

    const fecha = [
      ahora.getFullYear(),
      String(ahora.getMonth() + 1).padStart(2, '0'),
      String(ahora.getDate()).padStart(2, '0')
    ].join('-');

    const hora = ahora.toLocaleTimeString('es-BO');

    const datosVisita = {
      proyectoId: proyecto.id,
      proyectoNombre: proyecto.nombre || '',
      area: proyecto.area || '',
      stand: String(proyecto.stand || ''),

      fecha: fecha,
      hora: hora,

      staff_uid: staffActual.uid,
      staff_nombre: staffActual.nombreCompleto,

      creadoEn: ahora
    };

    await addDoc(
      collection(dbStaff, 'visitas_proyectos'),
      datosVisita
    );

    resultado.innerHTML = `
      <div class="form-estudiante">

        <h2 style="color:#15803d">
          ¡VISITA REGISTRADA!
        </h2>

        <h3>${escapar(proyecto.nombre || 'Proyecto')}</h3>

        <p>
          <strong>Área:</strong>
          ${escapar(proyecto.area || '')}
        </p>

        <p>
          <strong>Stand:</strong>
          ${escapar(String(proyecto.stand || ''))}
        </p>

        <p><strong>Fecha:</strong> ${fecha}</p>

        <p><strong>Hora:</strong> ${hora}</p>

        <p>
          <strong>Registrado por:</strong>
          ${escapar(staffActual.nombreCompleto)}
        </p>

        <span class="estado">● VISITA REGISTRADA</span>

        <button
          type="button"
          id="staff-otra-visita-proyecto"
          class="jurado-btn-principal">
          ESCANEAR OTRO STAND
        </button>

        <button
          type="button"
          id="staff-finalizar-visitas"
          class="jurado-btn-principal">
          VOLVER AL PANEL STAFF
        </button>

      </div>
    `;

    document
      .querySelector('#staff-otra-visita-proyecto')
      .addEventListener('click', () => {
        mostrarEscanerProyectoStaff();
      });

    document
      .querySelector('#staff-finalizar-visitas')
      .addEventListener('click', () => {
        mostrarPanelStaff();
      });

  } catch (error) {

    console.error('Error al registrar visita STAFF:', error);

    resultado.innerHTML = `
      <h2>NO SE PUDO REGISTRAR LA VISITA</h2>
      <p>${escapar(error.code || '')}</p>
      <p>${escapar(error.message || String(error))}</p>

      <button
        type="button"
        id="staff-volver-panel-error"
        class="jurado-btn-principal">
        VOLVER AL PANEL STAFF
      </button>
    `;

    document
      .querySelector('#staff-volver-panel-error')
      ?.addEventListener('click', () => {
        mostrarPanelStaff();
      });
  }
}

  // INICIAR CÁMARA
  try {

    lectorQRProyectoStaff = new Html5Qrcode(
      'lector-qr-proyecto-staff'
    )

    await lectorQRProyectoStaff.start(
      { facingMode: 'environment' },
      {
        fps: 10,
        qrbox: { width: 250, height: 250 }
      },

      async (textoQR) => {

        if (procesandoQR) return

        procesandoQR = true

        try {

          const prefijo = 'SHALOM-PROYECTO:'

          if (!textoQR.startsWith(prefijo)) {
            const resultado = document.querySelector(
              '#resultado-proyecto-staff'
            )

            if (resultado) {
              resultado.textContent =
                'Este QR no corresponde a un proyecto de SHALOM EXPO.'
            }

            procesandoQR = false
            return
          }

          const proyectoId = textoQR
            .substring(prefijo.length)
            .trim()

          if (!proyectoId) {
            throw new Error('El código QR no contiene un proyecto válido.')
          }

          await detenerEscanerProyecto()

          const proyectoRef = doc(db, 'proyectos', proyectoId)
          const respuestaProyecto = await getDoc(proyectoRef)

          if (!respuestaProyecto.exists()) {
            throw new Error('No se encontró el proyecto.')
          }

          const proyecto = {
            id: respuestaProyecto.id,
            ...respuestaProyecto.data()
          }

          if (proyecto.activo === false) {
            throw new Error('Este proyecto se encuentra inactivo.')
          }

          await mostrarSeleccionVisitante(proyecto)

        } catch (error) {

          console.error('Error al leer QR del proyecto:', error)

          const resultado = document.querySelector(
            '#resultado-proyecto-staff'
          )

          if (resultado) {
            resultado.textContent = error.message
          }

        } finally {
          procesandoQR = false
        }
      },

      () => {}
    )

  } catch (error) {

    console.error('Error al iniciar cámara de proyectos:', error)

    const resultado = document.querySelector(
      '#resultado-proyecto-staff'
    )

    if (resultado) {
      resultado.textContent =
        'No se pudo activar la cámara. Verifica los permisos del navegador.'
    }
  }
}

  // REGISTRAR VISITA A STANDS
  document
    .querySelector('#staff-escanear-proyecto')
    .addEventListener('click', () => {
      mostrarEscanerProyectoStaff()
    })

  // REGISTRAR ASISTENCIA
  document
    .querySelector('#staff-registrar-asistencia')
    .addEventListener('click', () => {
      mostrarEscanerAsistenciaStaff()
    })
}
// =====================================================
// ESCÁNER DE ASISTENCIA PARA PERSONAL STAFF
// =====================================================

async function mostrarEscanerAsistenciaStaff() {

  if (!staffActual) {
    iniciarAccesoStaff()
    return
  }

  app.innerHTML = `
    <main class="jurado-page">
      <section class="jurado-card">

        <header class="jurado-header">
          <div class="jurado-marca">
            <span class="jurado-marca-linea"></span>
            <span>COLEGIO ADVENTISTA SHALOM</span>
          </div>

          <p class="jurado-evento">CONTROL DE ASISTENCIA</p>
          <h1>SHALOM EXPO <span>2026</span></h1>
          <p class="jurado-subtitulo">Registro de estudiantes</p>
        </header>

        <div class="jurado-separador"></div>

        <section class="jurado-formulario">

          <div class="jurado-introduccion">
            <span class="jurado-paso">PERSONAL STAFF</span>
            <h2>Escanear credencial</h2>
            <p>
              Enfoca el código QR de la credencial
              del estudiante para registrar su asistencia.
            </p>
            <h3>${escapar(staffActual.nombreCompleto)}</h3>
          </div>

          <div id="lector-qr-staff"></div>

          <div id="resultado-qr-staff"
            class="jurado-mensaje"
            role="status">
            Esperando lectura del código QR...
          </div>

          <button
            type="button"
            id="staff-volver-panel"
            class="jurado-btn-principal">
            ← VOLVER AL PANEL STAFF
          </button>

        </section>

        <footer class="jurado-footer">
          SHALOM EXPO 2026 · PERSONAL STAFF
        </footer>

      </section>
    </main>
  `

  let lectorQRStaff = null
  let procesandoQRStaff = false

  async function detenerLectorStaff() {
    if (lectorQRStaff && lectorQRStaff.isScanning) {
      await lectorQRStaff.stop()
      await lectorQRStaff.clear()
    }

    lectorQRStaff = null
  }

  document
    .querySelector('#staff-volver-panel')
    .addEventListener('click', async () => {
      await detenerLectorStaff()
      mostrarPanelStaff()
    })

  try {

    lectorQRStaff = new Html5Qrcode('lector-qr-staff')

    await lectorQRStaff.start(
      { facingMode: 'environment' },
      {
        fps: 10,
        qrbox: { width: 250, height: 250 }
      },

      async (textoQR) => {

        if (procesandoQRStaff) return

        procesandoQRStaff = true

        const resultadoQR = document.querySelector(
          '#resultado-qr-staff'
        )

        try {

          const consulta = query(
            collection(db, 'estudiantes'),
            where('qr_id', '==', textoQR),
            limit(1)
          )

          const resultado = await getDocs(consulta)

          if (resultado.empty) {
            resultadoQR.innerHTML =
              '<h3>ESTUDIANTE NO ENCONTRADO</h3>'

            procesandoQRStaff = false
            return
          }

          const documento = resultado.docs[0]
          const estudiante = documento.data()

          if (estudiante.activo === false) {
            resultadoQR.innerHTML =
              '<h3>ESTUDIANTE INACTIVO</h3>'

            procesandoQRStaff = false
            return
          }

          const ahora = new Date()

          const fecha = [
            ahora.getFullYear(),
            String(ahora.getMonth() + 1).padStart(2, '0'),
            String(ahora.getDate()).padStart(2, '0')
          ].join('-')

          const hora = ahora.toLocaleTimeString('es-BO')

          const asistenciaRef = doc(
            db,
            'asistencias',
            `${documento.id}_${fecha}`
          )

          const asistenciaExistente = await getDoc(asistenciaRef)

          await detenerLectorStaff()

          if (asistenciaExistente.exists()) {

            resultadoQR.innerHTML = `
              <h2 style="color:#d97706">
                ASISTENCIA YA REGISTRADA
              </h2>

              <h3>
                ${escapar(estudiante.nombres || '')}
                ${escapar(estudiante.apellidos || '')}
              </h3>

              <p>
                Primera entrada:
                ${escapar(asistenciaExistente.data().hora || '')}
              </p>
            `

            return
          }

          await setDoc(asistenciaRef, {

            estudiante_id: documento.id,
            codigo: estudiante.codigo || '',
            nombres: estudiante.nombres || '',
            apellidos: estudiante.apellidos || '',
            nivel: estudiante.nivel || '',
            curso: estudiante.curso || '',
            paralelo: estudiante.paralelo || '',

            fecha: fecha,
            hora: hora,
            fecha_hora: ahora,
            estado: 'PRESENTE',

            staff_uid: staffActual.uid,
            staff_nombre: staffActual.nombreCompleto

          })

          resultadoQR.innerHTML = `
            <div class="form-estudiante">

              <h2 style="color:#15803d">
                ASISTENCIA REGISTRADA
              </h2>

              <h3>
                ${escapar(estudiante.nombres || '')}
                ${escapar(estudiante.apellidos || '')}
              </h3>

              <p>
                <strong>Código:</strong>
                ${escapar(estudiante.codigo || '')}
              </p>

              <p>
                <strong>Curso:</strong>
                ${escapar(estudiante.curso || '')}
                ${escapar(estudiante.paralelo || '')}
              </p>

              <p><strong>Fecha:</strong> ${fecha}</p>
              <p><strong>Hora:</strong> ${hora}</p>

              <p>
                <strong>Registrado por:</strong>
                ${escapar(staffActual.nombreCompleto)}
              </p>

              <span class="estado">● PRESENTE</span>

            </div>
          `

        } catch (error) {

          console.error(
            'Error al registrar asistencia STAFF:',
            error
          )

          resultadoQR.textContent =
            'No se pudo registrar la asistencia. Verifica los permisos de Firebase.'

          if (lectorQRStaff && lectorQRStaff.isScanning) {
            procesandoQRStaff = false
          }

        } finally {
          procesandoQRStaff = false
        }
      },

      () => {}
    )

  } catch (error) {

    console.error('Error al iniciar escáner STAFF:', error)

    const resultadoQR = document.querySelector(
      '#resultado-qr-staff'
    )

    if (resultadoQR) {
      resultadoQR.textContent =
        'No se pudo activar la cámara. Verifica los permisos del navegador.'
    }
  }
}