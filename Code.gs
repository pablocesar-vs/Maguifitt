/**
 * MAGUI FITT PRO - Backend Google Apps Script
 * Google Sheets + Google Drive + Web App
 */

const MF = {
  appName: 'Magui FITT PRO',
  timezone: 'America/Argentina/Buenos_Aires',
  folderName: 'Magui FITT PRO - Fotos',
  sheets: {
    perfil: 'Perfil',
    mediciones: 'Mediciones',
    fotos: 'Fotos',
    config: 'Configuracion',
    frases: 'Frases'
  }
};

const MF_HEADERS = {
  Perfil: ['ID','Nombre','Apellido','Fecha nacimiento','Edad','Altura','Peso inicial','Peso actual','Peso objetivo','Fecha inicio','Fecha objetivo','Foto perfil URL','Objetivo general','Actualizado'],
  Mediciones: ['ID','Fecha','Hora','Peso','Cuello','Pecho','Cintura','Abdomen','Cadera','Gluteos','Muslo izquierdo','Muslo derecho','Brazo izquierdo','Brazo derecho','IMC','Estado animo','Observaciones','Creado'],
  Fotos: ['ID','Medicion ID','Fecha','Tipo','Nombre archivo','Drive File ID','URL','Creado'],
  Configuracion: ['Clave','Valor'],
  Frases: ['ID','Frase','Activa']
};

function doGet() {
  inicializarProyecto();
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle(MF.appName)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function inicializarProyecto() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  crearHoja_(ss, MF.sheets.perfil, MF_HEADERS.Perfil);
  crearHoja_(ss, MF.sheets.mediciones, MF_HEADERS.Mediciones);
  crearHoja_(ss, MF.sheets.fotos, MF_HEADERS.Fotos);
  crearHoja_(ss, MF.sheets.config, MF_HEADERS.Configuracion);
  crearHoja_(ss, MF.sheets.frases, MF_HEADERS.Frases);
  iniciarConfig_();
  iniciarFrases_();
  iniciarPerfil_();
  return { ok: true };
}

function obtenerDatosApp() {
  inicializarProyecto();
  return {
    ok: true,
    perfil: obtenerPerfil(),
    mediciones: obtenerMediciones(),
    fotos: obtenerFotos(),
    dashboard: obtenerDashboard(),
    frase: obtenerFraseAleatoria(),
    graficos: obtenerDatosGraficos()
  };
}

function obtenerPerfil() {
  const sh = hoja_(MF.sheets.perfil);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return perfilDefault_();
  const obj = filaObjeto_(data[0], data[1]);
  obj.Edad = calcularEdad_(obj['Fecha nacimiento']);
  obj['Peso actual'] = ultimoPeso_() || num_(obj['Peso actual']) || num_(obj['Peso inicial']) || '';
  return obj;
}

function guardarPerfil(data) {
  if (!data) throw new Error('Datos inválidos.');
  const sh = hoja_(MF.sheets.perfil);
  const actual = obtenerPerfil();
  const nacimiento = data.fechaNacimiento || '';
  const row = [
    actual.ID || id_('PER'),
    txt_(data.nombre || 'Magui'),
    txt_(data.apellido),
    nacimiento,
    calcularEdad_(nacimiento),
    numero_(data.altura),
    numero_(data.pesoInicial),
    numero_(data.pesoActual),
    numero_(data.pesoObjetivo),
    data.fechaInicio || actual['Fecha inicio'] || fecha_(new Date()),
    data.fechaObjetivo || '',
    data.fotoPerfilUrl || actual['Foto perfil URL'] || '',
    txt_(data.objetivoGeneral || 'Mejorar mi evolución física de forma saludable.'),
    new Date()
  ];
  if (sh.getLastRow() < 2) sh.appendRow(row);
  else sh.getRange(2, 1, 1, row.length).setValues([row]);
  return { ok: true, perfil: obtenerPerfil(), dashboard: obtenerDashboard() };
}

function guardarMedicion(data) {
  if (!data) throw new Error('Datos inválidos.');
  const perfil = obtenerPerfil();
  const peso = numero_(data.peso);
  if (!peso) throw new Error('El peso es obligatorio.');
  const now = new Date();
  const row = [
    id_('MED'),
    fecha_(now),
    hora_(now),
    peso,
    numero_(data.cuello),
    numero_(data.pecho),
    numero_(data.cintura),
    numero_(data.abdomen),
    numero_(data.cadera),
    numero_(data.gluteos),
    numero_(data.musloIzquierdo),
    numero_(data.musloDerecho),
    numero_(data.brazoIzquierdo),
    numero_(data.brazoDerecho),
    calcularIMC_(peso, perfil.Altura),
    txt_(data.estadoAnimo),
    txt_(data.observaciones),
    now
  ];
  hoja_(MF.sheets.mediciones).appendRow(row);
  actualizarPesoPerfil_(peso);
  return { ok: true, medicionId: row[0], mediciones: obtenerMediciones(), dashboard: obtenerDashboard(), graficos: obtenerDatosGraficos() };
}

function eliminarMedicion(id) {
  if (!id) throw new Error('ID requerido.');
  const sh = hoja_(MF.sheets.mediciones);
  const row = buscarFila_(sh, id);
  if (row < 2) throw new Error('Medición no encontrada.');
  sh.deleteRow(row);
  const ult = obtenerMediciones()[0];
  if (ult) actualizarPesoPerfil_(ult.Peso);
  return { ok: true, mediciones: obtenerMediciones(), dashboard: obtenerDashboard(), graficos: obtenerDatosGraficos() };
}

function duplicarMedicion(id) {
  const m = obtenerMediciones().find(x => x.ID === id);
  if (!m) throw new Error('Medición no encontrada.');
  return guardarMedicion({
    peso: m.Peso,
    cuello: m.Cuello,
    pecho: m.Pecho,
    cintura: m.Cintura,
    abdomen: m.Abdomen,
    cadera: m.Cadera,
    gluteos: m.Gluteos,
    musloIzquierdo: m['Muslo izquierdo'],
    musloDerecho: m['Muslo derecho'],
    brazoIzquierdo: m['Brazo izquierdo'],
    brazoDerecho: m['Brazo derecho'],
    estadoAnimo: m['Estado animo'],
    observaciones: m.Observaciones
  });
}

function guardarFotoPerfil(base64, filename, mimeType) {
  const saved = guardarArchivoImagen_(base64, filename, mimeType);
  const p = obtenerPerfil();
  guardarPerfil({
    nombre: p.Nombre,
    apellido: p.Apellido,
    fechaNacimiento: p['Fecha nacimiento'],
    altura: p.Altura,
    pesoInicial: p['Peso inicial'],
    pesoActual: p['Peso actual'],
    pesoObjetivo: p['Peso objetivo'],
    fechaInicio: p['Fecha inicio'],
    fechaObjetivo: p['Fecha objetivo'],
    fotoPerfilUrl: saved.url,
    objetivoGeneral: p['Objetivo general']
  });
  return { ok: true, url: saved.url };
}

function guardarFotoProgreso(payload) {
  if (!payload || !payload.base64 || !payload.tipo) throw new Error('Datos de foto inválidos.');
  const saved = guardarArchivoImagen_(payload.base64, payload.filename, payload.mimeType);
  const now = new Date();
  const row = [id_('FOTO'), payload.medicionId || '', fecha_(now), txt_(payload.tipo), saved.name, saved.fileId, saved.url, now];
  hoja_(MF.sheets.fotos).appendRow(row);
  return { ok: true, foto: { id: row[0], tipo: row[3], url: row[6] } };
}

function obtenerMediciones() {
  const sh = hoja_(MF.sheets.mediciones);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const head = data[0];
  return data.slice(1).filter(r => r[0]).map(r => filaObjeto_(head, r)).sort((a,b) => new Date(b.Creado) - new Date(a.Creado));
}

function obtenerFotos() {
  const sh = hoja_(MF.sheets.fotos);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const head = data[0];
  return data.slice(1).filter(r => r[0]).map(r => filaObjeto_(head, r)).sort((a,b) => new Date(b.Creado) - new Date(a.Creado));
}

function obtenerDashboard() {
  const p = obtenerPerfil();
  const meds = obtenerMediciones();
  const inicial = num_(p['Peso inicial']) || (meds.length ? num_(meds[meds.length - 1].Peso) : 0);
  const actual = num_(p['Peso actual']) || ultimoPeso_() || inicial;
  const objetivo = num_(p['Peso objetivo']);
  return {
    pesoInicial: inicial || '',
    pesoActual: actual || '',
    pesoObjetivo: objetivo || '',
    diferencia: inicial && actual ? redondear_(actual - inicial) : '',
    imc: calcularIMC_(actual, p.Altura),
    progreso: progreso_(inicial, actual, objetivo),
    ultimaMedicion: meds[0] || null,
    totalMediciones: meds.length,
    diasTranscurridos: diasDesde_(p['Fecha inicio']),
    proximaMedicion: proxima_(meds[0]),
    mensaje: mensaje_(meds)
  };
}

function obtenerDatosGraficos() {
  const meds = obtenerMediciones().reverse();
  const peso = [['Fecha','Peso']];
  const cintura = [['Fecha','Cintura']];
  const gluteos = [['Fecha','Glúteos']];
  const imc = [['Fecha','IMC']];
  meds.forEach(m => {
    peso.push([String(m.Fecha), num_(m.Peso)]);
    cintura.push([String(m.Fecha), num_(m.Cintura)]);
    gluteos.push([String(m.Fecha), num_(m.Gluteos)]);
    imc.push([String(m.Fecha), num_(m.IMC)]);
  });
  return { peso, cintura, gluteos, imc };
}

function obtenerFraseAleatoria() {
  const sh = hoja_(MF.sheets.frases);
  const data = sh.getDataRange().getValues().slice(1).filter(r => r[1] && String(r[2]).toLowerCase() !== 'no');
  if (!data.length) return 'Cada pequeño avance cuenta.';
  return data[Math.floor(Math.random() * data.length)][1];
}

/* Internas */

function crearHoja_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#F8C8DC').setFontColor('#333333');
    sh.setFrozenRows(1);
  }
  sh.autoResizeColumns(1, headers.length);
}

function iniciarConfig_() {
  const sh = hoja_(MF.sheets.config);
  if (sh.getLastRow() > 1) return;
  sh.getRange(2,1,4,2).setValues([
    ['Periodicidad','Semanal'],
    ['Recordatorios','Si'],
    ['Unidad peso','kg'],
    ['Unidad medidas','cm']
  ]);
}

function iniciarFrases_() {
  const sh = hoja_(MF.sheets.frases);
  if (sh.getLastRow() > 1) return;
  sh.getRange(2,1,6,3).setValues([
    ['FRA-001','Cada pequeño avance cuenta.','Si'],
    ['FRA-002','Tu constancia vale más que la perfección.','Si'],
    ['FRA-003','Hoy también suma. Un paso más cerca.','Si'],
    ['FRA-004','El progreso real se construye con hábitos simples.','Si'],
    ['FRA-005','Magui, vas mejorando. Seguí así.','Si'],
    ['FRA-006','No se trata de correr, se trata de no abandonar.','Si']
  ]);
}

function iniciarPerfil_() {
  const sh = hoja_(MF.sheets.perfil);
  if (sh.getLastRow() > 1) return;
  const now = new Date();
  sh.appendRow([id_('PER'),'Magui','','','','','','','',fecha_(now),'','','Mejorar mi evolución física de forma saludable.',now]);
}

function guardarArchivoImagen_(base64, filename, mimeType) {
  if (!base64 || !filename || !mimeType || !String(mimeType).startsWith('image/')) throw new Error('Imagen inválida.');
  const folder = carpeta_();
  const clean = String(base64).includes(',') ? String(base64).split(',')[1] : String(base64);
  const blob = Utilities.newBlob(Utilities.base64Decode(clean), mimeType, Date.now() + '_' + filename.replace(/[^\w.\-]/g,'_'));
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { fileId: file.getId(), url: file.getUrl(), name: file.getName() };
}

function carpeta_() {
  const it = DriveApp.getFoldersByName(MF.folderName);
  return it.hasNext() ? it.next() : DriveApp.createFolder(MF.folderName);
}

function hoja_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('No existe la hoja ' + name);
  return sh;
}

function filaObjeto_(headers, row) {
  const o = {};
  headers.forEach((h,i) => o[h] = row[i] instanceof Date ? Utilities.formatDate(row[i], MF.timezone, 'yyyy-MM-dd') : (row[i] ?? ''));
  return o;
}

function perfilDefault_() {
  return { Nombre:'Magui', Apellido:'', Edad:'', Altura:'', 'Peso inicial':'', 'Peso actual':'', 'Peso objetivo':'', 'Fecha inicio':fecha_(new Date()), 'Foto perfil URL':'', 'Objetivo general':'Mejorar mi evolución física de forma saludable.' };
}

function actualizarPesoPerfil_(peso) {
  const sh = hoja_(MF.sheets.perfil);
  if (sh.getLastRow() < 2) iniciarPerfil_();
  const p = obtenerPerfil();
  sh.getRange(2,7).setValue(num_(p['Peso inicial']) || peso);
  sh.getRange(2,8).setValue(peso);
  sh.getRange(2,14).setValue(new Date());
}

function buscarFila_(sh, id) {
  const data = sh.getDataRange().getValues();
  for (let i=1;i<data.length;i++) if (data[i][0] === id) return i+1;
  return -1;
}

function ultimoPeso_() {
  const meds = obtenerMediciones();
  return meds.length ? num_(meds[0].Peso) : '';
}

function calcularEdad_(fecha) {
  if (!fecha) return '';
  const b = new Date(fecha);
  if (isNaN(b)) return '';
  const t = new Date();
  let e = t.getFullYear() - b.getFullYear();
  if (t.getMonth() < b.getMonth() || (t.getMonth() === b.getMonth() && t.getDate() < b.getDate())) e--;
  return e >= 0 ? e : '';
}

function calcularIMC_(peso, alturaCm) {
  const p = num_(peso);
  const h = num_(alturaCm) / 100;
  return p && h ? redondear_(p / (h*h)) : '';
}

function progreso_(inicial, actual, objetivo) {
  inicial = num_(inicial); actual = num_(actual); objetivo = num_(objetivo);
  if (!inicial || !actual || !objetivo || inicial === objetivo) return 0;
  return Math.max(0, Math.min(100, Math.round((Math.abs(inicial - actual) / Math.abs(inicial - objetivo)) * 100)));
}

function proxima_(ultima) {
  if (!ultima || !ultima.Fecha) return 'Pendiente';
  const d = new Date(ultima.Fecha);
  d.setDate(d.getDate() + 7);
  return fecha_(d);
}

function diasDesde_(fecha) {
  if (!fecha) return 0;
  const d = new Date(fecha);
  if (isNaN(d)) return 0;
  return Math.max(0, Math.floor((new Date() - d) / 86400000));
}

function mensaje_(meds) {
  if (!meds.length) return 'Cargá tu primera medición para empezar a ver tu evolución.';
  if (meds.length === 1) return 'Primera medición registrada. Ya tenés tu punto de partida.';
  const a = num_(meds[0].Peso), b = num_(meds[1].Peso);
  if (!a || !b) return 'Seguimos registrando datos para medir mejor el progreso.';
  const dif = redondear_(a - b);
  if (dif < 0) return 'Bajaste ' + Math.abs(dif) + ' kg desde la medición anterior. Excelente constancia.';
  if (dif > 0) return 'Subiste ' + dif + ' kg desde la medición anterior. Seguimos con foco.';
  return 'Tu peso se mantuvo estable. La constancia también es progreso.';
}

function txt_(v) { return v == null ? '' : String(v).trim(); }
function num_(v) { const n = Number(String(v ?? '').replace(',','.')); return isNaN(n) ? 0 : n; }
function numero_(v) { return v === '' || v == null ? '' : redondear_(num_(v)); }
function redondear_(n) { return Math.round(Number(n) * 100) / 100; }
function id_(p) { return p + '-' + Utilities.getUuid().slice(0,8).toUpperCase(); }
function fecha_(d) { return Utilities.formatDate(new Date(d), MF.timezone, 'yyyy-MM-dd'); }
function hora_(d) { return Utilities.formatDate(new Date(d), MF.timezone, 'HH:mm'); }
