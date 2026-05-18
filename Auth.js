const CABECERAS_USUARIOS = ['USUARIO', 'CLAVE', 'NOMBRE', 'ROL', 'ESTADO'];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('SCM Piura')
    .addItem('Preparar hoja Usuarios', 'asegurarHojaUsuarios')
    .addToUi();
}

function asegurarHojaUsuarios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.NOMBRE_HOJA_USUARIOS);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.NOMBRE_HOJA_USUARIOS);
  }

  if (sheet.getLastRow() === 0) {
    inicializarHojaUsuarios(sheet);
  }

  return sheet;
}

function inicializarHojaUsuarios(sheet) {
  sheet.clear();
  sheet.getRange(1, 1, 1, CABECERAS_USUARIOS.length).setValues([CABECERAS_USUARIOS]);
  sheet.getRange(2, 1, 1, CABECERAS_USUARIOS.length).setValues([
    ['admin', 'admin123', 'Administrador', 'ADMIN', CONFIG.ESTADO_USUARIO_ACTIVO]
  ]);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, CABECERAS_USUARIOS.length);
}

function validarLogin(usuario, password) {
  const sheet = asegurarHojaUsuarios();
  const data = sheet.getDataRange().getDisplayValues();

  if (data.length <= 1) {
    return {
      success: false,
      mensaje: "No hay usuarios registrados en la hoja 'Usuarios'."
    };
  }

  const col = obtenerColumnasUsuarios(data[0]);
  if (col.usuario === -1 || col.clave === -1) {
    return {
      success: false,
      mensaje: "La hoja 'Usuarios' debe tener las columnas USUARIO y CLAVE."
    };
  }

  const usuarioBuscado = normalizarTexto(usuario).toLowerCase();
  const claveBuscada = normalizarTexto(password);

  if (!usuarioBuscado || !claveBuscada) {
    return { success: false, mensaje: 'Complete usuario y clave.' };
  }

  for (let i = 1; i < data.length; i++) {
    const filaUsuario = normalizarTexto(data[i][col.usuario]).toLowerCase();
    const filaClave = normalizarTexto(data[i][col.clave]);

    if (filaUsuario !== usuarioBuscado || filaClave !== claveBuscada) {
      continue;
    }

    const estado = col.estado === -1
      ? CONFIG.ESTADO_USUARIO_ACTIVO
      : normalizarTexto(data[i][col.estado]).toUpperCase();

    if (estado && estado !== CONFIG.ESTADO_USUARIO_ACTIVO) {
      return { success: false, mensaje: 'Usuario inactivo. Revise la hoja Usuarios.' };
    }

    return {
      success: true,
      nombre: col.nombre === -1 ? data[i][col.usuario] : data[i][col.nombre],
      rol: col.rol === -1 ? '' : data[i][col.rol],
      usuario: data[i][col.usuario]
    };
  }

  return { success: false, mensaje: 'Usuario o clave incorrectos.' };
}

function validarClaveUsuarioActual(usuario, password) {
  const resultado = validarLogin(usuario, password);
  if (!resultado.success) {
    throw new Error(resultado.mensaje || 'No fue posible validar la clave del usuario.');
  }

  return resultado;
}

function obtenerColumnasUsuarios(cabeceras) {
  const normalizadas = cabeceras.map(function(celda) {
    return normalizarTexto(celda).toUpperCase();
  });

  const buscar = function(alias) {
    for (let i = 0; i < alias.length; i++) {
      const indice = normalizadas.indexOf(alias[i]);
      if (indice !== -1) {
        return indice;
      }
    }
    return -1;
  };

  return {
    usuario: buscar(['USUARIO', 'USER', 'LOGIN']),
    clave: buscar(['CLAVE', 'PASSWORD', 'CONTRASENA']),
    nombre: buscar(['NOMBRE', 'NOMBRES']),
    rol: buscar(['ROL', 'PERFIL']),
    estado: buscar(['ESTADO', 'ACTIVO'])
  };
}
