const CONFIG = {
  ID_CARPETA_CUADRILLAS: '14jg7JnmBg-P6yg50zATzbb6knLXKmk8s',
  ID_CARPETA_PROCESADOS: '1-ONXD_6VYbHDSJnGzmdlB2XcTCTznC6k',
  ID_HOJA_MAESTRO: SpreadsheetApp.getActiveSpreadsheet().getId(),
  NOMBRE_HOJA_PADRON: 'Padron_Maestro',
  NOMBRE_HOJA_USUARIOS: 'Usuarios',
  NOMBRE_CARPETA_PADRONES: '00_PADRONES_ENOSA',
  COLUMNAS_PADRON_REQUERIDAS: [
    'SUMINISTRO', 'CARTERA', 'NOMBRE_CLIENTE', 'DIRECCION', 'MARCA', 'MODELO',
    'SERIE_FAB', 'FACTOR', 'RUTA', 'NOMB_RUTA', 'TARIFA', 'PERIODO',
    'LEC_ANT_EAT', 'LEC_ANT_EAHP', 'LEC_ANT_EAFP', 'LEC_ANT_ER', 'LEC_ANT_PHP',
    'LEC_ANT_PFP', 'CUADRILLA', 'ORIGEN_LECTURA', 'ESTADO_LECTURA', 'ESTADO', 'FUENTE_ARCHIVO', 'NOMBRE_ARCHIVO',
    'ARCHIVO_LECTURA', 'TIPO_ARCHIVO_LECTURA', 'EAT', 'EAHP', 'EAFP', 'ER',
    'PHP', 'PFP', 'FECHA REGISTRO', 'OBS'
  ],
  CUADRILLAS: [
    { key: 'cuadrilla_1', label: 'Cuadrilla 1', folderName: '01', folderId: '' },
    { key: 'cuadrilla_2', label: 'Cuadrilla 2', folderName: '02', folderId: '' },
    { key: 'cuadrilla_3', label: 'Cuadrilla 3', folderName: '03', folderId: '' },
    { key: 'cuadrilla_4', label: 'Cuadrilla 4', folderName: '04', folderId: '' },
    { key: 'cuadrilla_5', label: 'Cuadrilla 5', folderName: '05', folderId: '' }
  ],

  OBIS: {
    EAT: { code: '1.8.0', label: 'Energía Activa Total' },
    EAHP: { code: '1.8.1', label: 'Energía Activa Hora Punta' },
    EAFP: { code: '1.8.2', label: 'Energía Activa Fuera de Punta' },
    ER: { code: '3.8.0', label: 'Energía Reactiva' },
    PHP: { code: '1.6.1', label: 'Potencia Hora Punta' },
    PFP: { code: '1.6.2', label: 'Potencia Fuera de Punta' }
  },

  ESTADO_COMPLETO: 'REALIZADO',
  ESTADO_PENDIENTE: 'PENDIENTE',
  ESTADO_LEIDO_SIN_CARGA: 'LEIDO SIN PARAMETROS',
  ESTADO_CORTADO: 'CORTADO',
  ESTADO_RETIRADO: 'RETIRADO',
  ORIGEN_LECTURA_TELEMETRIA: 'TELEMETRIA',
  ORIGEN_LECTURA_CAMPO: 'CAMPO',
  ESTADO_USUARIO_ACTIVO: 'ACTIVO'
};
