const SHEET_NAME = 'Respuestas';

function getOrCreateSheetId() {
  const props = PropertiesService.getScriptProperties();
  const existing = props.getProperty('SHEET_ID');
  if (existing) return existing;

  const ss = SpreadsheetApp.create('Ejercicios de griego — Respuestas');
  const sheet = ss.getActiveSheet().setName(SHEET_NAME);
  sheet.appendRow(['Fecha', 'Actividad', 'Nombre', 'Email', 'Ejercicio', 'Frase', 'Item', 'Respuesta', 'Correcta']);
  sheet.setFrozenRows(1);
  props.setProperty('SHEET_ID', ss.getId());
  return ss.getId();
}

function configurarPlanilla() {
  const id = getOrCreateSheetId();
  Logger.log('Listo. Planilla: ' + SpreadsheetApp.openById(id).getUrl());
}

function doPost(e) {
  const ss = SpreadsheetApp.openById(getOrCreateSheetId());
  const sheet = ss.getSheetByName(SHEET_NAME);
  const feedbackSheet = ss.getSheetByName('Feedback') || ss.insertSheet('Feedback');
  ensureFeedbackHeader_(feedbackSheet);
  const asistenciaSheet = ss.getSheetByName('Asistencia') || ss.insertSheet('Asistencia');
  ensureAsistenciaHeader_(asistenciaSheet);
  const consultasSheet = ss.getSheetByName('Preguntas y sugerencias espontáneas') || ss.insertSheet('Preguntas y sugerencias espontáneas');
  ensureConsultasHeader_(consultasSheet);

  const data = JSON.parse(e.postData.contents);
  const timestamp = new Date();
  const actividad = data.actividad || '';
  const nombre = data.nombre || '';
  const email = data.email || '';
  const respuestas = data.respuestas || [];

  const rows = respuestas.map(function (r) {
    return [
      timestamp,
      actividad,
      nombre,
      email,
      r.ejercicio || '',
      r.frase != null ? r.frase : '',
      r.item || '',
      r.respuesta != null ? r.respuesta : '',
      r.correcta === true ? 'Sí' : (r.correcta === false ? 'No' : '')
    ];
  });

  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }

  // Pivotear las 3 filas de "Evaluación" de este envío en una sola fila de "Feedback".
  var dificultad = '', comentario = '', sugerencia = '', huboFeedback = false;
  respuestas.forEach(function (r) {
    if (r.ejercicio !== 'Evaluación') return;
    huboFeedback = true;
    if (r.item === 'Dificultad percibida') dificultad = r.respuesta || '';
    if (r.item === 'Comentario sobre la dificultad') comentario = r.respuesta || '';
    if (r.item === 'Sugerencia') sugerencia = r.respuesta || '';
  });
  if (huboFeedback) {
    feedbackSheet.appendRow([timestamp, actividad, nombre, email, dificultad, comentario, sugerencia]);
  }

  // Registrar asistencia: un resumen por envío con el % de respuestas correctas.
  // Solo cuentan los ítems evaluables (correcta true/false) — quedan afuera "Evaluación"
  // y cualquier ítem con correcta:null (p. ej. una palabra no marcada, que ya se cuenta
  // como incorrecta aparte, o metadatos que no son ni acierto ni error).
  var correctas = 0, evaluables = 0;
  respuestas.forEach(function (r) {
    if (r.correcta === true) { correctas++; evaluables++; }
    else if (r.correcta === false) { evaluables++; }
  });
  if (evaluables > 0) {
    var pct = Math.round((correctas / evaluables) * 100);
    asistenciaSheet.appendRow([timestamp, actividad, nombre, email, correctas, evaluables, pct + '%']);
  }

  // Consulta espontánea desde el botón "Preguntas/Propuestas" de la página de entrada.
  if (data.consulta) {
    consultasSheet.appendRow([timestamp, email, data.consulta]);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, filas: rows.length }))
    .setMimeType(ContentService.MimeType.JSON);
}

function ensureConsultasHeader_(sheet) {
  if (sheet.getLastRow() > 0) return;
  sheet.getRange(1, 1, 1, 3).setValues([['Fecha', 'Email', 'Consulta']]);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
}

function ensureFeedbackHeader_(sheet) {
  if (sheet.getLastRow() > 0) return;
  sheet.getRange(1, 1, 1, 7)
    .setValues([['Fecha', 'Actividad', 'Nombre', 'Email', 'Dificultad percibida', 'Comentario', 'Sugerencia']]);
  sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
}

function ensureAsistenciaHeader_(sheet) {
  if (sheet.getLastRow() > 0) return;
  sheet.getRange(1, 1, 1, 7)
    .setValues([['Fecha', 'Actividad', 'Nombre', 'Email', 'Correctas', 'Evaluables', '% Correctas']]);
  sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
}

// ---------------------------------------------------------------------------
// Correr UNA sola vez, a mano (menú "Ejecutar" en el editor de Apps Script), para
// insertar la columna "Actividad" en la hoja principal sin romper las filas que ya
// estaban cargadas desde antes de este cambio (esas quedan con "Actividad" vacío —
// son de antes de que el HTML mandara ese dato, no hay forma de reconstruirlo).
// Segura para correr más de una vez: si la columna B ya dice "Actividad", no hace nada.
// ---------------------------------------------------------------------------
function migrarColumnaActividad() {
  const ss = SpreadsheetApp.openById(getOrCreateSheetId());
  const sheet = ss.getSheetByName(SHEET_NAME);
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (header[1] === 'Actividad') {
    Logger.log('Ya estaba migrada, no se tocó nada.');
    return;
  }
  sheet.insertColumnBefore(2);
  sheet.getRange(1, 2).setValue('Actividad');
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold');
  Logger.log('Columna "Actividad" insertada. Las filas viejas quedan con ese campo vacío.');
}

// ---------------------------------------------------------------------------
// Correr UNA sola vez, a mano, para traer a "Feedback" el feedback que ya estaba
// cargado en la hoja principal desde antes de este cambio. Segura para correr más de
// una vez: no duplica filas que ya estén en "Feedback" (compara por contenido).
// Correr DESPUÉS de migrarColumnaActividad(), si hace falta.
// ---------------------------------------------------------------------------
function backfillFeedback() {
  const ss = SpreadsheetApp.openById(getOrCreateSheetId());
  const sheet = ss.getSheetByName(SHEET_NAME);
  const feedbackSheet = ss.getSheetByName('Feedback') || ss.insertSheet('Feedback');
  ensureFeedbackHeader_(feedbackSheet);

  const data = sheet.getDataRange().getValues();
  const groups = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var fecha = row[0], actividad = row[1], nombre = row[2], email = row[3], ejercicio = row[4], item = row[6], respuesta = row[7];
    if (ejercicio !== 'Evaluación') continue;
    var key = fecha + '|' + actividad + '|' + nombre + '|' + email;
    if (!groups[key]) groups[key] = { fecha: fecha, actividad: actividad, nombre: nombre, email: email, dificultad: '', comentario: '', sugerencia: '' };
    if (item === 'Dificultad percibida') groups[key].dificultad = respuesta;
    if (item === 'Comentario sobre la dificultad') groups[key].comentario = respuesta;
    if (item === 'Sugerencia') groups[key].sugerencia = respuesta;
  }

  const existing = feedbackSheet.getDataRange().getValues();
  const existingKeys = {};
  for (var j = 1; j < existing.length; j++) {
    existingKeys[feedbackContentKey_(existing[j][1], existing[j][2], existing[j][3], existing[j][4], existing[j][5], existing[j][6])] = true;
  }

  const newRows = [];
  Object.keys(groups).forEach(function (key) {
    var g = groups[key];
    var ck = feedbackContentKey_(g.actividad, g.nombre, g.email, g.dificultad, g.comentario, g.sugerencia);
    if (existingKeys[ck]) return;
    existingKeys[ck] = true;
    newRows.push([g.fecha, g.actividad, g.nombre, g.email, g.dificultad, g.comentario, g.sugerencia]);
  });

  if (newRows.length > 0) {
    feedbackSheet.getRange(feedbackSheet.getLastRow() + 1, 1, newRows.length, 7).setValues(newRows);
  }
  Logger.log('Backfill: ' + newRows.length + ' filas nuevas agregadas a Feedback.');
}

function feedbackContentKey_(actividad, nombre, email, dificultad, comentario, sugerencia) {
  return [actividad, nombre, email, dificultad, comentario, sugerencia].join('|');
}

// ---------------------------------------------------------------------------
// Correr UNA sola vez, a mano, para traer a "Asistencia" los envíos que ya estaban
// cargados en la hoja principal desde antes de este cambio. Segura para correr más de
// una vez: no duplica filas que ya estén en "Asistencia" (compara por contenido).
// Correr DESPUÉS de migrarColumnaActividad(), si hace falta.
// ---------------------------------------------------------------------------
function backfillAsistencia() {
  const ss = SpreadsheetApp.openById(getOrCreateSheetId());
  const sheet = ss.getSheetByName(SHEET_NAME);
  const asistenciaSheet = ss.getSheetByName('Asistencia') || ss.insertSheet('Asistencia');
  ensureAsistenciaHeader_(asistenciaSheet);

  const data = sheet.getDataRange().getValues();
  const groups = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var fecha = row[0], actividad = row[1], nombre = row[2], email = row[3], correccion = row[8];
    if (correccion !== 'Sí' && correccion !== 'No') continue;
    var key = fecha + '|' + actividad + '|' + nombre + '|' + email;
    if (!groups[key]) groups[key] = { fecha: fecha, actividad: actividad, nombre: nombre, email: email, correctas: 0, evaluables: 0 };
    groups[key].evaluables++;
    if (correccion === 'Sí') groups[key].correctas++;
  }

  const existing = asistenciaSheet.getDataRange().getValues();
  const existingKeys = {};
  for (var j = 1; j < existing.length; j++) {
    existingKeys[asistenciaContentKey_(existing[j][1], existing[j][2], existing[j][3], existing[j][4], existing[j][5])] = true;
  }

  const newRows = [];
  Object.keys(groups).forEach(function (key) {
    var g = groups[key];
    if (g.evaluables === 0) return;
    var ck = asistenciaContentKey_(g.actividad, g.nombre, g.email, g.correctas, g.evaluables);
    if (existingKeys[ck]) return;
    existingKeys[ck] = true;
    var pct = Math.round((g.correctas / g.evaluables) * 100);
    newRows.push([g.fecha, g.actividad, g.nombre, g.email, g.correctas, g.evaluables, pct + '%']);
  });

  if (newRows.length > 0) {
    asistenciaSheet.getRange(asistenciaSheet.getLastRow() + 1, 1, newRows.length, 7).setValues(newRows);
  }
  Logger.log('Backfill: ' + newRows.length + ' filas nuevas agregadas a Asistencia.');
}

function asistenciaContentKey_(actividad, nombre, email, correctas, evaluables) {
  return [actividad, nombre, email, correctas, evaluables].join('|');
}

// ---------------------------------------------------------------------------
// GET público: sirve el banco de vocabulario de "Repaso nonstop" como JSON, leyendo
// directamente de la hoja "Vocabulario" — así Santiago corrige texto (traducciones,
// citas, frases de contexto) editando la planilla, sin tocar código ni volver a
// publicar el HTML.
// Columnas esperadas (fila 1 = encabezado, se ignora):
//   Greek | Meaning | Kind | Number | CorrectCases | Person | Voice | Citation | ContextGreek | ContextWord | ContextSpanish
// CorrectCases: casos separados por coma (ej. "nom,voc"). Kind: nombre / adjetivo / participio / verbo / infinitivo.
// ---------------------------------------------------------------------------
const VOCAB_SHEET_NAME = 'Vocabulario';

function ensureVocabHeader_(sheet) {
  if (sheet.getLastRow() > 0) return;
  var header = ['Greek', 'Meaning', 'Kind', 'Number', 'CorrectCases', 'Person', 'Voice', 'Citation', 'ContextGreek', 'ContextWord', 'ContextSpanish', 'Semana'];
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.getRange(1, 1, 1, header.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

// doGet rutea por el parámetro ?tipo= (o por ?email=, para "Mi progreso") — así un solo
// Apps Script sirve el banco de vocabulario (Repaso nonstop), el contenido de la página
// de entrada (index.html) y la consulta de progreso por email, sin duplicar backend. Sin
// parámetro, por compatibilidad con lo ya desplegado, sirve vocabulario (repaso_nonstop.html
// no manda parámetro).
function doGet(e) {
  const email = (e && e.parameter && e.parameter.email) || '';
  if (email) return doGetProgreso_(email);
  const tipo = (e && e.parameter && e.parameter.tipo) || 'vocabulario';
  if (tipo === 'indice') return doGetIndice_();
  return doGetVocabulario_();
}

// ---------------------------------------------------------------------------
// GET ?email=... — para el botón "Mi progreso" de la página de entrada. Devuelve
// qué actividades (valores de "Actividad", i.e. document.title) tienen al menos un
// envío de ese email en la hoja "Asistencia". No expone nada del resto de la clase:
// solo lo que coincide con el email pedido.
// ---------------------------------------------------------------------------
function doGetProgreso_(email) {
  const emailNorm = String(email).trim().toLowerCase();
  const ss = SpreadsheetApp.openById(getOrCreateSheetId());
  const asistenciaSheet = ss.getSheetByName('Asistencia');
  const actividades = {};
  if (asistenciaSheet) {
    const data = asistenciaSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const rowEmail = String(data[i][3] || '').trim().toLowerCase();
      if (rowEmail === emailNorm) actividades[data[i][1]] = true; // columna "Actividad"
    }
  }
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, actividades: Object.keys(actividades) }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGetVocabulario_() {
  const ss = SpreadsheetApp.openById(getOrCreateSheetId());
  const sheet = ss.getSheetByName(VOCAB_SHEET_NAME) || ss.insertSheet(VOCAB_SHEET_NAME);
  ensureVocabHeader_(sheet);

  const data = sheet.getDataRange().getValues();
  const bank = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue; // fila vacía
    var entry = {
      greek: row[0],
      meaning: row[1],
      kind: row[2],
      citation: row[7] || '',
      contextGreek: row[8] || '',
      contextWord: row[9] || '',
      contextSpanish: row[10] || ''
    };
    if (row[3]) entry.number = row[3];
    if (row[4]) entry.correctCases = String(row[4]).split(',').map(function (c) { return c.trim(); }).filter(Boolean);
    if (row[5]) entry.person = String(row[5]);
    if (row[6]) entry.voice = row[6];
    if (row[11]) entry.semana = String(row[11]);
    bank.push(entry);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ bank: bank }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------------
// Contenido de la página de entrada (index.html), leído de la hoja "Índice".
// Columnas: Semana | SemanaActual (SI/vacío) | Tipo | Título | Link | Estado
// Tipo: "Lectura fuente" / "Lectura bibliográfica" / "Práctico" / "Ficha" / "Audio" / "Actividad"
// Estado: solo aplica a "Actividad" — "Obligatoria" / "Optativa"
// ---------------------------------------------------------------------------
const INDICE_SHEET_NAME = 'Índice';

function ensureIndiceHeader_(sheet) {
  if (sheet.getLastRow() > 0) return;
  var header = ['Semana', 'SemanaActual', 'Tipo', 'Título', 'Link', 'Estado'];
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.getRange(1, 1, 1, header.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function doGetIndice_() {
  const ss = SpreadsheetApp.openById(getOrCreateSheetId());
  const sheet = ss.getSheetByName(INDICE_SHEET_NAME) || ss.insertSheet(INDICE_SHEET_NAME);
  ensureIndiceHeader_(sheet);

  const data = sheet.getDataRange().getValues();
  const items = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue; // fila vacía
    items.push({
      semana: String(row[0]),
      semanaActual: String(row[1] || '').trim().toUpperCase() === 'SI',
      tipo: row[2] || '',
      titulo: row[3] || '',
      link: row[4] || '',
      estado: row[5] || ''
    });
  }

  return ContentService
    .createTextOutput(JSON.stringify({ items: items }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------------
// Correr UNA sola vez, a mano, para poblar la hoja "Vocabulario" con el banco inicial
// (41 palabras de Frases 1 y 2, todas con una frase de origen real). Si la hoja ya
// tiene filas, no hace nada — para no pisar correcciones manuales de Santiago.
// Para agregar palabras de un ejercicio nuevo más adelante, agregar filas a mano
// directamente en la hoja "Vocabulario" — no hace falta volver a correr esto.
// ---------------------------------------------------------------------------
function poblarVocabularioInicial() {
  const ss = SpreadsheetApp.openById(getOrCreateSheetId());
  const sheet = ss.getSheetByName(VOCAB_SHEET_NAME) || ss.insertSheet(VOCAB_SHEET_NAME);
  ensureVocabHeader_(sheet);

  if (sheet.getLastRow() > 1) {
    Logger.log('La hoja "Vocabulario" ya tiene datos — no se tocó nada.');
    return;
  }

  var F1_1 = ['λόγος ἐστὶ φωνὴ σημαντική', 'El discurso es voz con significado'];
  var F1_2 = ['ἐκ καρποῦ δένδρον γιγνώσκετε', 'Por el fruto conocen el árbol'];
  var F1_3 = ['ὁ ὕπνος τοῦ θανάτου ἀδελφός ἐστιν', 'El sueño es hermano de la muerte'];
  var F1_4 = ['φασὶ γὰρ τὴν τοῦ ἀνθρώπου ψυχὴν εἶναι ἀθάνατον', 'Pues dicen que el alma del hombre es inmortal'];
  var F1_5 = ['Πίνδαρος ὁ ποιητὴς τὸν οἶνον καλέει δρόσον τῆς ἀμπέλου', 'El poeta Píndaro llama rocío de la vid al vino'];
  var F1_6 = ['λόγος κάτοπτρον τοῦ θυμοῦ', 'El discurso es espejo del ánimo'];
  var F2_1 = ['ἡ ψυχὴ καὶ αὐτή, Ἔρως σχέτλιε, ἔχει πτέρυγας', 'También el alma misma, cruel Eros, tiene alas.'];
  var F2_2 = ['τῷ ἀνθρώπῳ σῶμα μὲν θνητόν, ψυχὴ δὲ ἀθάνατος', 'El ser humano tiene el cuerpo mortal, pero el alma inmortal.'];
  var F2_3 = ['τὰ χαλεπὰ τῶν πραγμάτων ὁ χρόνος διαλύει', 'El tiempo disuelve las dificultades (de los asuntos).'];
  var F2_4 = ['κακὴ ἡ δελφῖνος ἐν χέρσῳ βία', 'La fuerza del delfín en tierra firme es mala.'];
  var F2_5 = ['ἐξ ὀνύχων τὸν λέοντα γιγνώσκειν ἔξεστι', 'Es posible conocer al león por sus garras.'];
  var F2_6 = ['δὶς παῖδες οἱ γέροντες', 'Los viejos son niños dos veces.'];
  var F2_7 = ['ἡ Δίκη γὰρ πάντα βλέπει', 'Pues la Justicia todo lo ve.'];
  var F2_8 = ['Ἔνιοι κακῶς φρονοῦσι πράττοντες καλῶς', 'Algunos piensan mal, aunque actúen bien.'];

  // [greek, meaning, kind, number, correctCases, person, voice, citation, contextPair, contextWord]
  var rows = [
    ['λόγος', 'discurso, palabra', 'nombre', 'sg', 'nom', '', '', 'λόγος, -ου', F1_1, 'λόγος'],
    ['φωνὴ', 'voz', 'nombre', 'sg', 'nom,voc', '', '', 'φωνή, -ῆς', F1_1, 'φωνὴ'],
    ['καρποῦ', 'fruto', 'nombre', 'sg', 'gen', '', '', 'καρπός, -οῦ', F1_2, 'καρποῦ'],
    ['δένδρον', 'árbol', 'nombre', 'sg', 'nom,ac,voc', '', '', 'δένδρον, -ου', F1_2, 'δένδρον'],
    ['ὕπνος', 'sueño', 'nombre', 'sg', 'nom', '', '', 'ὕπνος, -ου', F1_3, 'ὕπνος'],
    ['θανάτου', 'muerte', 'nombre', 'sg', 'gen', '', '', 'θάνατος, -ου', F1_3, 'θανάτου'],
    ['ἀδελφός', 'hermano', 'nombre', 'sg', 'nom', '', '', 'ἀδελφός, -οῦ', F1_3, 'ἀδελφός'],
    ['ἀνθρώπου', 'hombre, ser humano', 'nombre', 'sg', 'gen', '', '', 'ἄνθρωπος, -ου', F1_4, 'ἀνθρώπου'],
    ['ψυχὴν', 'alma', 'nombre', 'sg', 'ac', '', '', 'ψυχή, -ῆς', F1_4, 'ψυχὴν'],
    ['Πίνδαρος', 'Píndaro (poeta)', 'nombre', 'sg', 'nom', '', '', 'Πίνδαρος, -ου', F1_5, 'Πίνδαρος'],
    ['ποιητὴς', 'poeta', 'nombre', 'sg', 'nom', '', '', 'ποιητής, -οῦ', F1_5, 'ποιητὴς'],
    ['οἶνον', 'vino', 'nombre', 'sg', 'ac', '', '', 'οἶνος, -ου', F1_5, 'οἶνον'],
    ['δρόσον', 'rocío', 'nombre', 'sg', 'ac', '', '', 'δρόσος, -ου', F1_5, 'δρόσον'],
    ['ἀμπέλου', 'vid', 'nombre', 'sg', 'gen', '', '', 'ἄμπελος, -ου', F1_5, 'ἀμπέλου'],
    ['κάτοπτρον', 'espejo', 'nombre', 'sg', 'nom,ac,voc', '', '', 'κάτοπτρον, -ου', F1_6, 'κάτοπτρον'],
    ['θυμοῦ', 'ánimo, espíritu', 'nombre', 'sg', 'gen', '', '', 'θυμός, -οῦ', F1_6, 'θυμοῦ'],
    ['Ἔρως', 'Eros (dios del amor)', 'nombre', 'sg', 'nom,voc', '', '', 'Ἔρως, -ωτος', F2_1, 'Ἔρως'],
    ['πτέρυγας', 'ala', 'nombre', 'pl', 'ac', '', '', 'πτέρυξ, -υγος', F2_1, 'πτέρυγας'],
    ['σῶμα', 'cuerpo', 'nombre', 'sg', 'nom,ac,voc', '', '', 'σῶμα, -ατος', F2_2, 'σῶμα'],
    ['πραγμάτων', 'asunto, cosa', 'nombre', 'pl', 'gen', '', '', 'πρᾶγμα, -ατος', F2_3, 'πραγμάτων'],
    ['χρόνος', 'tiempo', 'nombre', 'sg', 'nom', '', '', 'χρόνος, -ου', F2_3, 'χρόνος'],
    ['δελφῖνος', 'delfín', 'nombre', 'sg', 'gen', '', '', 'δελφίς, -ῖνος', F2_4, 'δελφῖνος'],
    ['χέρσῳ', 'tierra firme', 'nombre', 'sg', 'dat', '', '', 'χέρσος, -ου', F2_4, 'χέρσῳ'],
    ['βία', 'fuerza', 'nombre', 'sg', 'nom,voc', '', '', 'βία, -ας', F2_4, 'βία'],
    ['ὀνύχων', 'uña, garra', 'nombre', 'pl', 'gen', '', '', 'ὄνυξ, -υχος', F2_5, 'ὀνύχων'],
    ['λέοντα', 'león', 'nombre', 'sg', 'ac', '', '', 'λέων, -οντος', F2_5, 'λέοντα'],
    ['παῖδες', 'niño, niña', 'nombre', 'pl', 'nom,voc', '', '', 'παῖς, παιδός', F2_6, 'παῖδες'],
    ['γέροντες', 'viejo, anciano', 'nombre', 'pl', 'nom,voc', '', '', 'γέρων, -οντος', F2_6, 'γέροντες'],
    ['Δίκη', 'Justicia (personificada)', 'nombre', 'sg', 'nom,voc', '', '', 'Δίκη, -ης', F2_7, 'Δίκη'],

    ['ἐστὶ', 'ser, estar', 'verbo', 'sg', '', '3', '', 'εἰμί', F1_1, 'ἐστὶ'],
    ['γιγνώσκετε', 'conocer', 'verbo', 'pl', '', '2', '', 'γιγνώσκω', F1_2, 'γιγνώσκετε'],
    ['φασὶ', 'decir', 'verbo', 'pl', '', '3', '', 'φημί', F1_4, 'φασὶ'],
    ['καλέει', 'llamar', 'verbo', 'sg', '', '3', '', 'καλέω', F1_5, 'καλέει'],
    ['ἔχει', 'tener', 'verbo', 'sg', '', '3', '', 'ἔχω', F2_1, 'ἔχει'],
    ['διαλύει', 'disolver', 'verbo', 'sg', '', '3', '', 'διαλύω', F2_3, 'διαλύει'],
    ['ἔξεστι', 'ser posible', 'verbo', 'sg', '', '3', '', 'ἔξεστι(ν)', F2_5, 'ἔξεστι'],
    ['βλέπει', 'ver', 'verbo', 'sg', '', '3', '', 'βλέπω', F2_7, 'βλέπει'],
    ['φρονοῦσι', 'pensar', 'verbo', 'pl', '', '3', '', 'φρονέω', F2_8, 'φρονοῦσι'],

    ['εἶναι', 'ser, estar', 'infinitivo', '', '', '', 'act', 'εἰμί', F1_4, 'εἶναι'],
    ['γιγνώσκειν', 'conocer', 'infinitivo', '', '', '', 'act', 'γιγνώσκω', F2_5, 'γιγνώσκειν'],
    ['πράττοντες', 'hacer, actuar', 'participio', 'pl', 'nom,voc', '', '', 'πράττω', F2_8, 'πράττοντες']
  ];

  var out = rows.map(function (r) {
    var contextPair = r[8];
    return [r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], contextPair[0], r[9], contextPair[1]];
  });

  sheet.getRange(2, 1, out.length, out[0].length).setValues(out);
  Logger.log('Listo: ' + out.length + ' palabras cargadas en "Vocabulario". Planilla: ' + ss.getUrl());
}


// ---------------------------------------------------------------------------
// Correr UNA sola vez, a mano, para sumar a "Vocabulario" las formas rastreadas
// desde "Base de frases y formas griegas.xlsx" (pestaña "Frases", filas con
// Traducción propia ya cargada) — 137 formas nuevas de 36 frases: 7 del Práctico
// 2 (20.08) que nunca se habían usado en ningún ejercicio, y 29 de material
// marcado "Griego I" / "cursada anterior" (Nivel: Griego I, o Año 2025) que no
// pertenecen a esta cursada de Griego II. Pensadas para el nivel "Todo (incluye
// Griego I)" del selector de dificultad de Repaso nonstop (ver conversación del
// 2026-08-31) — no se les asignó columna de semana todavía, eso se resuelve en
// una pasada aparte. Segura para correr una sola vez: no repite nada de
// poblarVocabularioInicial() (verificado, cero superposición) — si se corre dos
// veces, sí duplica, porque a diferencia de poblarVocabularioInicial() esta no
// chequea si la hoja ya tiene datos (para no bloquear la carga incremental
// normal). Antes de correrla, confirmar a ojo que "λύσω" todavía no aparece en
// la columna A de "Vocabulario".
// ---------------------------------------------------------------------------
function poblarVocabularioAmpliado() {
  const ss = SpreadsheetApp.openById(getOrCreateSheetId());
  const sheet = ss.getSheetByName(VOCAB_SHEET_NAME) || ss.insertSheet(VOCAB_SHEET_NAME);
  ensureVocabHeader_(sheet);

  const existing = sheet.getDataRange().getValues();
  const yaEsta = existing.some(function (r) { return r[0] === 'λύσω'; });
  if (yaEsta) {
    Logger.log('Ya parece estar cargado (se encontró "λύσω") — no se tocó nada.');
    return;
  }

  var rows = [
    ['ἀνήρ', 'hombre, varón', 'nombre', 'sg', 'nom', '', '', 'ἀνήρ, ἀνδρός', 'Ἀνὴρ ὁ φεύγων καὶ πάλιν μαχήσεται', 'Ἀνὴρ', 'El hombre que huye también luchará de nuevo.'],
    ['φεύγων', 'huir, evitar', 'participio', 'sg', 'nom,voc', '', '', 'φεύγω', 'Ἀνὴρ ὁ φεύγων καὶ πάλιν μαχήσεται', 'φεύγων', 'El hombre que huye también luchará de nuevo.'],
    ['μαχήσεται', 'luchar, combatir', 'verbo', 'sg', '', '3', '', 'μάχομαι', 'Ἀνὴρ ὁ φεύγων καὶ πάλιν μαχήσεται', 'μαχήσεται', 'El hombre que huye también luchará de nuevo.'],
    ['ἄνθρωπος', 'hombre, ser humano', 'nombre', 'sg', 'nom', '', '', 'ἄνθρωπος, -ου', 'Ἄνθρωπος ὢν γίνωσκε τῆς ὀργῆς κρατεῖν', 'Ἄνθρωπος', 'Siendo humano, aprendé a dominar tu ira.'],
    ['ὤν', 'ser, estar', 'participio', 'sg', 'nom,voc', '', '', 'εἰμί', 'Ἄνθρωπος ὢν γίνωσκε τῆς ὀργῆς κρατεῖν', 'ὢν', 'Siendo humano, aprendé a dominar tu ira.'],
    ['γίνωσκε', 'conocer', 'verbo', 'sg', '', '2', '', 'γιγνώσκω', 'Ἄνθρωπος ὢν γίνωσκε τῆς ὀργῆς κρατεῖν', 'γίνωσκε', 'Siendo humano, aprendé a dominar tu ira.'],
    ['ὀργῆς', 'ira, cólera', 'nombre', 'sg', 'gen', '', '', 'ὀργή, -ῆς', 'Ἄνθρωπος ὢν γίνωσκε τῆς ὀργῆς κρατεῖν', 'ὀργῆς', 'Siendo humano, aprendé a dominar tu ira.'],
    ['κρατεῖν', 'dominar, tener poder sobre', 'infinitivo', '', '', '', 'act', 'κρατέω', 'Ἄνθρωπος ὢν γίνωσκε τῆς ὀργῆς κρατεῖν', 'κρατεῖν', 'Siendo humano, aprendé a dominar tu ira.'],
    ['πολλοί', 'mucho, numeroso', 'adjetivo', 'pl', 'nom,voc', '', '', 'πολύς, πολλή, πολύ', 'Πολλοὶ κακῶς πράττουσιν οὐκ ὄντες κακοί', 'Πολλοὶ', 'A muchos les va mal sin ser malos.'],
    ['πράττουσι', 'hacer, actuar', 'verbo', 'pl', '', '3', '', 'πράττω', 'Πολλοὶ κακῶς πράττουσιν οὐκ ὄντες κακοί', 'πράττουσιν', 'A muchos les va mal sin ser malos.'],
    ['ὄντες', 'ser, estar', 'participio', 'pl', 'nom,voc', '', '', 'εἰμί', 'Πολλοὶ κακῶς πράττουσιν οὐκ ὄντες κακοί', 'ὄντες', 'A muchos les va mal sin ser malos.'],
    ['κακοί', 'malo, funesto', 'adjetivo', 'pl', 'nom,voc', '', '', 'κακός, -ή, -όν', 'Πολλοὶ κακῶς πράττουσιν οὐκ ὄντες κακοί', 'κακοί', 'A muchos les va mal sin ser malos.'],
    ['ὄνος', 'burro, asno', 'nombre', 'sg', 'nom', '', '', 'ὄνος, -ου', 'ὄνος ξύλων γόμον φέρων λίμνην διέβαινεν', 'ὄνος', 'Un burro que llevaba una carga de leña, cruzaba un lago.'],
    ['ξύλων', 'madera, leña', 'nombre', 'pl', 'gen', '', '', 'ξύλον, -ου', 'ὄνος ξύλων γόμον φέρων λίμνην διέβαινεν', 'ξύλων', 'Un burro que llevaba una carga de leña, cruzaba un lago.'],
    ['γόμον', 'carga', 'nombre', 'sg', 'ac', '', '', 'γόμος, -ου', 'ὄνος ξύλων γόμον φέρων λίμνην διέβαινεν', 'γόμον', 'Un burro que llevaba una carga de leña, cruzaba un lago.'],
    ['φέρων', 'llevar, cargar', 'participio', 'sg', 'nom,voc', '', '', 'φέρω', 'ὄνος ξύλων γόμον φέρων λίμνην διέβαινεν', 'φέρων', 'Un burro que llevaba una carga de leña, cruzaba un lago.'],
    ['λίμνην', 'lago, laguna, pantano', 'nombre', 'sg', 'ac', '', '', 'λίμνη, -ης', 'ὄνος ξύλων γόμον φέρων λίμνην διέβαινεν', 'λίμνην', 'Un burro que llevaba una carga de leña, cruzaba un lago.'],
    ['διέβαινεν', 'atravesar, cruzar', 'verbo', 'sg', '', '3', '', 'διαβαίνω', 'ὄνος ξύλων γόμον φέρων λίμνην διέβαινεν', 'διέβαινεν', 'Un burro que llevaba una carga de leña, cruzaba un lago.'],
    ['δελφῖνα', 'delfín', 'nombre', 'sg', 'ac', '', '', 'δελφίς, -ῖνος', 'δελφῖνα νήχεσθαι διδάσκεις', 'δελφῖνα', 'Le enseñas al delfín a nadar.'],
    ['νήχεσθαι', 'nadar', 'infinitivo', '', '', '', 'med', 'νήχομαι', 'δελφῖνα νήχεσθαι διδάσκεις', 'νήχεσθαι', 'Le enseñas al delfín a nadar.'],
    ['διδάσκεις', 'enseñar', 'verbo', 'sg', '', '2', '', 'διδάσκω', 'δελφῖνα νήχεσθαι διδάσκεις', 'διδάσκεις', 'Le enseñas al delfín a nadar.'],
    ['φησί', 'decir, afirmar', 'verbo', 'sg', '', '3', '', 'φημί', 'φησὶ γάρ που Πρωταγόρας πάντων χρημάτων μέτρον ἄνθρωπον εἶναι', 'φησὶ', 'Pues dice, creo, Protágoras, que el hombre es la medida de todas las cosas.'],
    ['Πρωταγόρας', 'Protágoras (sofista)', 'nombre', 'sg', 'nom', '', '', 'Πρωταγόρας, -ου', 'φησὶ γάρ που Πρωταγόρας πάντων χρημάτων μέτρον ἄνθρωπον εἶναι', 'Πρωταγόρας', 'Pues dice, creo, Protágoras, que el hombre es la medida de todas las cosas.'],
    ['πάντων', 'todo, cada', 'adjetivo', 'pl', 'gen', '', '', 'πᾶς, πᾶσα, πᾶν', 'φησὶ γάρ που Πρωταγόρας πάντων χρημάτων μέτρον ἄνθρωπον εἶναι', 'πάντων', 'Pues dice, creo, Protágoras, que el hombre es la medida de todas las cosas.'],
    ['χρημάτων', 'cosa, asunto; (pl.) dinero, bienes', 'nombre', 'pl', 'gen', '', '', 'χρῆμα, -ατος', 'φησὶ γάρ που Πρωταγόρας πάντων χρημάτων μέτρον ἄνθρωπον εἶναι', 'χρημάτων', 'Pues dice, creo, Protágoras, que el hombre es la medida de todas las cosas.'],
    ['μέτρον', 'medida', 'nombre', 'sg', 'nom,ac,voc', '', '', 'μέτρον, -ου', 'φησὶ γάρ που Πρωταγόρας πάντων χρημάτων μέτρον ἄνθρωπον εἶναι', 'μέτρον', 'Pues dice, creo, Protágoras, que el hombre es la medida de todas las cosas.'],
    ['ἄνθρωπον', 'hombre, ser humano', 'nombre', 'sg', 'ac', '', '', 'ἄνθρωπος, -ου', 'φησὶ γάρ που Πρωταγόρας πάντων χρημάτων μέτρον ἄνθρωπον εἶναι', 'ἄνθρωπον', 'Pues dice, creo, Protágoras, que el hombre es la medida de todas las cosas.'],
    ['Πέρσαι', 'persa', 'nombre', 'pl', 'nom,voc', '', '', 'Πέρσης, -ου', 'οἱ Πέρσαι διδάσκουσι τοὺς παῖδας πείθεσθαι τοῖς ἄρχουσιν', 'Πέρσαι', 'Los persas enseñan a los niños a obedecer a los gobernantes.'],
    ['διδάσκουσι', 'enseñar', 'verbo', 'pl', '', '3', '', 'διδάσκω', 'οἱ Πέρσαι διδάσκουσι τοὺς παῖδας πείθεσθαι τοῖς ἄρχουσιν', 'διδάσκουσι', 'Los persas enseñan a los niños a obedecer a los gobernantes.'],
    ['παῖδας', 'niño, niña', 'nombre', 'pl', 'ac', '', '', 'παῖς, παιδός', 'οἱ Πέρσαι διδάσκουσι τοὺς παῖδας πείθεσθαι τοῖς ἄρχουσιν', 'παῖδας', 'Los persas enseñan a los niños a obedecer a los gobernantes.'],
    ['πείθεσθαι', 'persuadir; (mediopasivo) obedecer', 'infinitivo', '', '', '', 'med', 'πείθω', 'οἱ Πέρσαι διδάσκουσι τοὺς παῖδας πείθεσθαι τοῖς ἄρχουσιν', 'πείθεσθαι', 'Los persas enseñan a los niños a obedecer a los gobernantes.'],
    ['ἄρχουσιν', 'gobernante, magistrado', 'nombre', 'pl', 'dat', '', '', 'ἄρχων, -οντος', 'οἱ Πέρσαι διδάσκουσι τοὺς παῖδας πείθεσθαι τοῖς ἄρχουσιν', 'ἄρχουσιν', 'Los persas enseñan a los niños a obedecer a los gobernantes.'],
    ['ποταμοί', 'río', 'nombre', 'pl', 'nom,voc', '', '', 'ποταμός, -οῦ', 'οἱ ποταμοὶ ἐκ τῶν πηγῶν κατὰ τῶν ἄκρων διὰ τοῦ πεδίου εἰς τὴν θάλατταν ῥέουσιν', 'ποταμοὶ', 'Los ríos fluyen desde las fuentes por las alturas a través de la llanura hacia el mar.'],
    ['πηγῶν', 'fuente, manantial', 'nombre', 'pl', 'gen', '', '', 'πηγή, -ῆς', 'οἱ ποταμοὶ ἐκ τῶν πηγῶν κατὰ τῶν ἄκρων διὰ τοῦ πεδίου εἰς τὴν θάλατταν ῥέουσιν', 'πηγῶν', 'Los ríos fluyen desde las fuentes por las alturas a través de la llanura hacia el mar.'],
    ['ἄκρων', 'extremo, más alto; (subst.) cumbre, altura', 'adjetivo', 'pl', 'gen', '', '', 'ἄκρος, -α, -ον', 'οἱ ποταμοὶ ἐκ τῶν πηγῶν κατὰ τῶν ἄκρων διὰ τοῦ πεδίου εἰς τὴν θάλατταν ῥέουσιν', 'ἄκρων', 'Los ríos fluyen desde las fuentes por las alturas a través de la llanura hacia el mar.'],
    ['πεδίου', 'llanura', 'nombre', 'sg', 'gen', '', '', 'πεδίον, -ου', 'οἱ ποταμοὶ ἐκ τῶν πηγῶν κατὰ τῶν ἄκρων διὰ τοῦ πεδίου εἰς τὴν θάλατταν ῥέουσιν', 'πεδίου', 'Los ríos fluyen desde las fuentes por las alturas a través de la llanura hacia el mar.'],
    ['θάλατταν', 'mar', 'nombre', 'sg', 'ac', '', '', 'θάλαττα, -ης', 'οἱ ποταμοὶ ἐκ τῶν πηγῶν κατὰ τῶν ἄκρων διὰ τοῦ πεδίου εἰς τὴν θάλατταν ῥέουσιν', 'θάλατταν', 'Los ríos fluyen desde las fuentes por las alturas a través de la llanura hacia el mar.'],
    ['ῥέουσι', 'fluir, correr (de un líquido)', 'verbo', 'pl', '', '3', '', 'ῥέω', 'οἱ ποταμοὶ ἐκ τῶν πηγῶν κατὰ τῶν ἄκρων διὰ τοῦ πεδίου εἰς τὴν θάλατταν ῥέουσιν', 'ῥέουσιν', 'Los ríos fluyen desde las fuentes por las alturas a través de la llanura hacia el mar.'],
    ['δεῖ', 'ser necesario, hacer falta', 'verbo', 'sg', '', '3', '', 'δεῖ', 'δεῖ ἐν τοῖς μὲν ὅπλοις φοβερούς, ἐν δὲ τοῖς δικαστηρίοις φιλανθρώπους εἶναι', 'δεῖ', 'Es necesario ser temibles en las armas, pero benévolos en los tribunales.'],
    ['ὅπλοις', 'arma', 'nombre', 'pl', 'dat', '', '', 'ὅπλον, -ου', 'δεῖ ἐν τοῖς μὲν ὅπλοις φοβερούς, ἐν δὲ τοῖς δικαστηρίοις φιλανθρώπους εἶναι', 'ὅπλοις', 'Es necesario ser temibles en las armas, pero benévolos en los tribunales.'],
    ['φοβερούς', 'temible, terrible', 'adjetivo', 'pl', 'ac', '', '', 'φοβερός, -ά, -όν', 'δεῖ ἐν τοῖς μὲν ὅπλοις φοβερούς, ἐν δὲ τοῖς δικαστηρίοις φιλανθρώπους εἶναι', 'φοβερούς', 'Es necesario ser temibles en las armas, pero benévolos en los tribunales.'],
    ['δικαστηρίοις', 'tribunal', 'nombre', 'pl', 'dat', '', '', 'δικαστήριον, -ου', 'δεῖ ἐν τοῖς μὲν ὅπλοις φοβερούς, ἐν δὲ τοῖς δικαστηρίοις φιλανθρώπους εἶναι', 'δικαστηρίοις', 'Es necesario ser temibles en las armas, pero benévolos en los tribunales.'],
    ['φιλανθρώπους', 'benévolo, humano', 'adjetivo', 'pl', 'ac', '', '', 'φιλάνθρωπος, -ον', 'δεῖ ἐν τοῖς μὲν ὅπλοις φοβερούς, ἐν δὲ τοῖς δικαστηρίοις φιλανθρώπους εἶναι', 'φιλανθρώπους', 'Es necesario ser temibles en las armas, pero benévolos en los tribunales.'],
    ['Αἰγύπτιοι', 'egipcio', 'nombre', 'pl', 'nom,voc', '', '', 'Αἰγύπτιος, -α, -ον', 'οἱ Αἰγύπτιοι τὸν ἥλιον καὶ τὴν σελήνην θεοὺς εἶναι νομίζουσιν', 'Αἰγύπτιοι', 'Los egipcios consideran que el sol y la luna son dioses.'],
    ['ἥλιον', 'sol', 'nombre', 'sg', 'ac', '', '', 'ἥλιος, -ου', 'οἱ Αἰγύπτιοι τὸν ἥλιον καὶ τὴν σελήνην θεοὺς εἶναι νομίζουσιν', 'ἥλιον', 'Los egipcios consideran que el sol y la luna son dioses.'],
    ['σελήνην', 'luna', 'nombre', 'sg', 'ac', '', '', 'σελήνη, -ης', 'οἱ Αἰγύπτιοι τὸν ἥλιον καὶ τὴν σελήνην θεοὺς εἶναι νομίζουσιν', 'σελήνην', 'Los egipcios consideran que el sol y la luna son dioses.'],
    ['θεούς', 'dios, diosa', 'nombre', 'pl', 'ac', '', '', 'θεός, -οῦ', 'οἱ Αἰγύπτιοι τὸν ἥλιον καὶ τὴν σελήνην θεοὺς εἶναι νομίζουσιν', 'θεοὺς', 'Los egipcios consideran que el sol y la luna son dioses.'],
    ['νομίζουσι', 'considerar, creer', 'verbo', 'pl', '', '3', '', 'νομίζω', 'οἱ Αἰγύπτιοι τὸν ἥλιον καὶ τὴν σελήνην θεοὺς εἶναι νομίζουσιν', 'νομίζουσιν', 'Los egipcios consideran que el sol y la luna son dioses.'],
    ['Εὐριπίδης', 'Eurípides (poeta trágico)', 'nombre', 'sg', 'nom', '', '', 'Εὐριπίδης, -ου', 'ὁ Εὐριπίδης τραγικώτατος γε τῶν ποιητῶν φαίνεται', 'Εὐριπίδης', 'Eurípides parece ser, de los poetas, ciertamente el más trágico.'],
    ['τραγικώτατος', 'trágico; (superl.) el más trágico', 'adjetivo', 'sg', 'nom', '', '', 'τραγικός, -ή, -όν (superl. τραγικώτατος)', 'ὁ Εὐριπίδης τραγικώτατος γε τῶν ποιητῶν φαίνεται', 'τραγικώτατος', 'Eurípides parece ser, de los poetas, ciertamente el más trágico.'],
    ['ποιητῶν', 'poeta', 'nombre', 'pl', 'gen', '', '', 'ποιητής, -οῦ', 'ὁ Εὐριπίδης τραγικώτατος γε τῶν ποιητῶν φαίνεται', 'ποιητῶν', 'Eurípides parece ser, de los poetas, ciertamente el más trágico.'],
    ['φαίνεται', 'parecer, aparecer', 'verbo', 'sg', '', '3', '', 'φαίνομαι', 'ὁ Εὐριπίδης τραγικώτατος γε τῶν ποιητῶν φαίνεται', 'φαίνεται', 'Eurípides parece ser, de los poetas, ciertamente el más trágico.'],
    ['πάτριοι', 'ancestral, patrio, tradicional', 'adjetivo', 'pl', 'nom,voc', '', '', 'πάτριος, -α, -ον', 'οἱ πάτριοι νόμοι κελεύουσι τοὺς φόνους θανάτῳ κολάζεσθαι', 'πάτριοι', 'Las leyes patrias ordenan que los asesinatos sean castigados con la muerte.'],
    ['νόμοι', 'ley, costumbre', 'nombre', 'pl', 'nom,voc', '', '', 'νόμος, -ου', 'οἱ πάτριοι νόμοι κελεύουσι τοὺς φόνους θανάτῳ κολάζεσθαι', 'νόμοι', 'Las leyes patrias ordenan que los asesinatos sean castigados con la muerte.'],
    ['κελεύουσι', 'ordenar, mandar', 'verbo', 'pl', '', '3', '', 'κελεύω', 'οἱ πάτριοι νόμοι κελεύουσι τοὺς φόνους θανάτῳ κολάζεσθαι', 'κελεύουσι', 'Las leyes patrias ordenan que los asesinatos sean castigados con la muerte.'],
    ['φόνους', 'asesinato, homicidio', 'nombre', 'pl', 'ac', '', '', 'φόνος, -ου', 'οἱ πάτριοι νόμοι κελεύουσι τοὺς φόνους θανάτῳ κολάζεσθαι', 'φόνους', 'Las leyes patrias ordenan que los asesinatos sean castigados con la muerte.'],
    ['θανάτῳ', 'muerte', 'nombre', 'sg', 'dat', '', '', 'θάνατος, -ου', 'οἱ πάτριοι νόμοι κελεύουσι τοὺς φόνους θανάτῳ κολάζεσθαι', 'θανάτῳ', 'Las leyes patrias ordenan que los asesinatos sean castigados con la muerte.'],
    ['κολάζεσθαι', 'castigar; (mediopasivo) ser castigado', 'infinitivo', '', '', '', 'med', 'κολάζω', 'οἱ πάτριοι νόμοι κελεύουσι τοὺς φόνους θανάτῳ κολάζεσθαι', 'κολάζεσθαι', 'Las leyes patrias ordenan que los asesinatos sean castigados con la muerte.'],
    ['ἄγετε', 'llevar, conducir', 'verbo', 'pl', '', '2', '', 'ἄγω', 'ἄγετε τοὺς ταύρους εἰς τὸ πεδίον καὶ θύετε τοῖς θεοῖς', 'ἄγετε', 'Lleven los toros a la llanura y sacrifíquenlos a los dioses.'],
    ['ταύρους', 'toro', 'nombre', 'pl', 'ac', '', '', 'ταῦρος, -ου', 'ἄγετε τοὺς ταύρους εἰς τὸ πεδίον καὶ θύετε τοῖς θεοῖς', 'ταύρους', 'Lleven los toros a la llanura y sacrifíquenlos a los dioses.'],
    ['πεδίον', 'llanura', 'nombre', 'sg', 'nom,ac,voc', '', '', 'πεδίον, -ου', 'ἄγετε τοὺς ταύρους εἰς τὸ πεδίον καὶ θύετε τοῖς θεοῖς', 'πεδίον', 'Lleven los toros a la llanura y sacrifíquenlos a los dioses.'],
    ['θύετε', 'sacrificar', 'verbo', 'pl', '', '2', '', 'θύω', 'ἄγετε τοὺς ταύρους εἰς τὸ πεδίον καὶ θύετε τοῖς θεοῖς', 'θύετε', 'Lleven los toros a la llanura y sacrifíquenlos a los dioses.'],
    ['θεοῖς', 'dios', 'nombre', 'pl', 'dat', '', '', 'θεός, -οῦ', 'ἄγετε τοὺς ταύρους εἰς τὸ πεδίον καὶ θύετε τοῖς θεοῖς', 'θεοῖς', 'Lleven los toros a la llanura y sacrifíquenlos a los dioses.'],
    ['ἀγρῶν', 'campo', 'nombre', 'pl', 'gen', '', '', 'ἀγρός, -οῦ', 'ἐκ τῶν ἀγρῶν οἱ πολέμιοι φέρονται τοὺς καλοὺς ἵππους', 'ἀγρῶν', 'Los enemigos se llevan de los campos los bellos caballos.'],
    ['πολέμιοι', 'enemigo, hostil', 'adjetivo', 'pl', 'nom,voc', '', '', 'πολέμιος, -α, -ον', 'ἐκ τῶν ἀγρῶν οἱ πολέμιοι φέρονται τοὺς καλοὺς ἵππους', 'πολέμιοι', 'Los enemigos se llevan de los campos los bellos caballos.'],
    ['φέρονται', 'llevar; (mediopasivo) llevarse', 'verbo', 'pl', '', '3', '', 'φέρω', 'ἐκ τῶν ἀγρῶν οἱ πολέμιοι φέρονται τοὺς καλοὺς ἵππους', 'φέρονται', 'Los enemigos se llevan de los campos los bellos caballos.'],
    ['καλούς', 'bello, bueno', 'adjetivo', 'pl', 'ac', '', '', 'καλός, -ή, -όν', 'ἐκ τῶν ἀγρῶν οἱ πολέμιοι φέρονται τοὺς καλοὺς ἵππους', 'καλοὺς', 'Los enemigos se llevan de los campos los bellos caballos.'],
    ['ἵππους', 'caballo', 'nombre', 'pl', 'ac', '', '', 'ἵππος, -ου', 'ἐκ τῶν ἀγρῶν οἱ πολέμιοι φέρονται τοὺς καλοὺς ἵππους', 'ἵππους', 'Los enemigos se llevan de los campos los bellos caballos.'],
    ['χρή', 'ser necesario, convenir', 'verbo', 'sg', '', '3', '', 'χρή', 'χρὴ φιλόσοφον εἶναι φίλον ἀληθείας καὶ δικαιοσύνης', 'χρὴ', 'Es necesario que el filósofo sea amigo de la verdad y la justicia.'],
    ['φιλόσοφον', 'filósofo', 'nombre', 'sg', 'ac', '', '', 'φιλόσοφος, -ου', 'χρὴ φιλόσοφον εἶναι φίλον ἀληθείας καὶ δικαιοσύνης', 'φιλόσοφον', 'Es necesario que el filósofo sea amigo de la verdad y la justicia.'],
    ['φίλον', 'amigo', 'nombre', 'sg', 'ac', '', '', 'φίλος, -ου', 'χρὴ φιλόσοφον εἶναι φίλον ἀληθείας καὶ δικαιοσύνης', 'φίλον', 'Es necesario que el filósofo sea amigo de la verdad y la justicia.'],
    ['ἀληθείας', 'verdad', 'nombre', 'sg', 'gen', '', '', 'ἀλήθεια, -ας', 'χρὴ φιλόσοφον εἶναι φίλον ἀληθείας καὶ δικαιοσύνης', 'ἀληθείας', 'Es necesario que el filósofo sea amigo de la verdad y la justicia.'],
    ['δικαιοσύνης', 'justicia', 'nombre', 'sg', 'gen', '', '', 'δικαιοσύνη, -ης', 'χρὴ φιλόσοφον εἶναι φίλον ἀληθείας καὶ δικαιοσύνης', 'δικαιοσύνης', 'Es necesario que el filósofo sea amigo de la verdad y la justicia.'],
    ['πάτταλοι', 'clavo, estaca', 'nombre', 'pl', 'nom,voc', '', '', 'πάτταλος, -ου', 'οἱ πάτταλοι παττάλοις ἐκκρούονται', 'πάτταλοι', 'Los clavos se sacan con clavos.'],
    ['παττάλοις', 'clavo, estaca', 'nombre', 'pl', 'dat', '', '', 'πάτταλος, -ου', 'οἱ πάτταλοι παττάλοις ἐκκρούονται', 'παττάλοις', 'Los clavos se sacan con clavos.'],
    ['ἐκκρούονται', 'expulsar, sacar (a golpes)', 'verbo', 'pl', '', '3', '', 'ἐκκρούω', 'οἱ πάτταλοι παττάλοις ἐκκρούονται', 'ἐκκρούονται', 'Los clavos se sacan con clavos.'],
    ['κακῶν', 'malo', 'adjetivo', 'pl', 'gen', '', '', 'κακός, -ή, -όν', 'τὰς δὲ τῶν κακῶν συνουσίας φεῦγε ἀμεταστρεπτί', 'κακῶν', 'Huí sin vueltas de la compañía de los malos.'],
    ['συνουσίας', 'compañía, trato, relación', 'nombre', 'pl', 'ac', '', '', 'συνουσία, -ας', 'τὰς δὲ τῶν κακῶν συνουσίας φεῦγε ἀμεταστρεπτί', 'συνουσίας', 'Huí sin vueltas de la compañía de los malos.'],
    ['φεῦγε', 'huir, evitar', 'verbo', 'sg', '', '2', '', 'φεύγω', 'τὰς δὲ τῶν κακῶν συνουσίας φεῦγε ἀμεταστρεπτί', 'φεῦγε', 'Huí sin vueltas de la compañía de los malos.'],
    ['λέγετε', 'decir, hablar', 'verbo', 'pl', '', '2', '', 'λέγω', 'οὐκ ὀρθῶς λέγετε, ὦ ἄνθρωποι, ἀλλὰ ψεύδεσθε', 'λέγετε', 'No hablan correctamente, oh, personas, sino que mienten.'],
    ['ἄνθρωποι', 'hombre, ser humano, persona', 'nombre', 'pl', 'nom,voc', '', '', 'ἄνθρωπος, -ου', 'οὐκ ὀρθῶς λέγετε, ὦ ἄνθρωποι, ἀλλὰ ψεύδεσθε', 'ἄνθρωποι', 'No hablan correctamente, oh, personas, sino que mienten.'],
    ['ψεύδεσθε', 'mentir', 'verbo', 'pl', '', '2', '', 'ψεύδομαι', 'οὐκ ὀρθῶς λέγετε, ὦ ἄνθρωποι, ἀλλὰ ψεύδεσθε', 'ψεύδεσθε', 'No hablan correctamente, oh, personas, sino que mienten.'],
    ['κακοῖς', 'malo; (subst. neutro) desgracia, mal', 'adjetivo', 'pl', 'dat', '', '', 'κακός, -ή, -όν', 'ἐν τοῖς κακοῖς δεῖ τοὺς φίλους εὐεργετεῖν', 'κακοῖς', 'En las desgracias hay que hacer bien a los amigos.'],
    ['φίλους', 'amigo', 'nombre', 'pl', 'ac', '', '', 'φίλος, -ου', 'ἐν τοῖς κακοῖς δεῖ τοὺς φίλους εὐεργετεῖν', 'φίλους', 'En las desgracias hay que hacer bien a los amigos.'],
    ['εὐεργετεῖν', 'hacer bien a, beneficiar', 'infinitivo', '', '', '', 'act', 'εὐεργετέω', 'ἐν τοῖς κακοῖς δεῖ τοὺς φίλους εὐεργετεῖν', 'εὐεργετεῖν', 'En las desgracias hay que hacer bien a los amigos.'],
    ['ἀρετή', 'virtud, excelencia', 'nombre', 'sg', 'nom,voc', '', '', 'ἀρετή, -ῆς', 'ἡ δ᾽ ἀρετὴ πάσης τέχνης ἀκριβεστέρα ἐστίν', 'ἀρετὴ', 'Pero la virtud es más exacta que toda técnica.'],
    ['πάσης', 'todo, cada', 'adjetivo', 'sg', 'gen', '', '', 'πᾶς, πᾶσα, πᾶν', 'ἡ δ᾽ ἀρετὴ πάσης τέχνης ἀκριβεστέρα ἐστίν', 'πάσης', 'Pero la virtud es más exacta que toda técnica.'],
    ['τέχνης', 'arte, técnica, oficio', 'nombre', 'sg', 'gen', '', '', 'τέχνη, -ης', 'ἡ δ᾽ ἀρετὴ πάσης τέχνης ἀκριβεστέρα ἐστίν', 'τέχνης', 'Pero la virtud es más exacta que toda técnica.'],
    ['ἀκριβεστέρα', 'preciso, exacto; (comp.) más exacto', 'adjetivo', 'sg', 'nom,voc', '', '', 'ἀκριβής, -ές (comp. ἀκριβέστερος)', 'ἡ δ᾽ ἀρετὴ πάσης τέχνης ἀκριβεστέρα ἐστίν', 'ἀκριβεστέρα', 'Pero la virtud es más exacta que toda técnica.'],
    ['Πυθαγόρᾳ', 'Pitágoras', 'nombre', 'sg', 'dat', '', '', 'Πυθαγόρας, -ου', 'Πυθαγόρᾳ μὲν μάλιστα ἤρεσκεν ἡ ἐχεμυθία', 'Πυθαγόρᾳ', 'A Pitágoras le agradaba especialmente el silencio.'],
    ['ἤρεσκεν', 'agradar, complacer', 'verbo', 'sg', '', '3', '', 'ἀρέσκω', 'Πυθαγόρᾳ μὲν μάλιστα ἤρεσκεν ἡ ἐχεμυθία', 'ἤρεσκεν', 'A Pitágoras le agradaba especialmente el silencio.'],
    ['ἐχεμυθία', 'discreción, silencio (guardar un secreto)', 'nombre', 'sg', 'nom,voc', '', '', 'ἐχεμυθία, -ας', 'Πυθαγόρᾳ μὲν μάλιστα ἤρεσκεν ἡ ἐχεμυθία', 'ἐχεμυθία', 'A Pitágoras le agradaba especialmente el silencio.'],
    ['θεοί', 'dios', 'nombre', 'pl', 'nom,voc', '', '', 'θεός, -οῦ', 'θεοί τ᾽ εἰσὶν καὶ ἀνθρώπων ἐπιμελοῦνται', 'θεοί', 'Los dioses existen y se preocupan por los seres humanos.'],
    ['εἰσί', 'ser, estar, existir', 'verbo', 'pl', '', '3', '', 'εἰμί', 'θεοί τ᾽ εἰσὶν καὶ ἀνθρώπων ἐπιμελοῦνται', 'εἰσὶν', 'Los dioses existen y se preocupan por los seres humanos.'],
    ['ἀνθρώπων', 'hombre, ser humano', 'nombre', 'pl', 'gen', '', '', 'ἄνθρωπος, -ου', 'θεοί τ᾽ εἰσὶν καὶ ἀνθρώπων ἐπιμελοῦνται', 'ἀνθρώπων', 'Los dioses existen y se preocupan por los seres humanos.'],
    ['ἐπιμελοῦνται', 'preocuparse por, cuidar de', 'verbo', 'pl', '', '3', '', 'ἐπιμελέομαι', 'θεοί τ᾽ εἰσὶν καὶ ἀνθρώπων ἐπιμελοῦνται', 'ἐπιμελοῦνται', 'Los dioses existen y se preocupan por los seres humanos.'],
    ['λοιπόν', 'restante, que queda; (ac. adverbial) en lo sucesivo', 'adjetivo', 'sg', 'nom,ac,voc', '', '', 'λοιπός, -ή, -όν', 'ἡμεῖς δὲ τὸ λοιπὸν Βακχίῳ δουλεύσομεν', 'λοιπὸν', 'Pero nosotros, de aquí en más, serviremos a Baco.'],
    ['Βακχίῳ', 'Baco (=Dioniso)', 'nombre', 'sg', 'dat', '', '', 'Βάκχιος, -ου', 'ἡμεῖς δὲ τὸ λοιπὸν Βακχίῳ δουλεύσομεν', 'Βακχίῳ', 'Pero nosotros, de aquí en más, serviremos a Baco.'],
    ['δουλεύσομεν', 'servir, ser esclavo de', 'verbo', 'pl', '', '1', '', 'δουλεύω', 'ἡμεῖς δὲ τὸ λοιπὸν Βακχίῳ δουλεύσομεν', 'δουλεύσομεν', 'Pero nosotros, de aquí en más, serviremos a Baco.'],
    ['Σώστρατος', 'Sóstrato (nombre propio)', 'nombre', 'sg', 'nom', '', '', 'Σώστρατος, -ου', 'Σώστρατος ἦν μοι ἐπιτήδειος καὶ φίλος', 'Σώστρατος', 'Sóstrato era para mí cercano y amigo.'],
    ['ἦν', 'ser, estar', 'verbo', 'sg', '', '3', '', 'εἰμί', 'Σώστρατος ἦν μοι ἐπιτήδειος καὶ φίλος', 'ἦν', 'Sóstrato era para mí cercano y amigo.'],
    ['ἐπιτήδειος', 'apto, conveniente; cercano, íntimo', 'adjetivo', 'sg', 'nom', '', '', 'ἐπιτήδειος, -α, -ον', 'Σώστρατος ἦν μοι ἐπιτήδειος καὶ φίλος', 'ἐπιτήδειος', 'Sóstrato era para mí cercano y amigo.'],
    ['φίλος', 'amigo', 'nombre', 'sg', 'nom', '', '', 'φίλος, -ου', 'Σώστρατος ἦν μοι ἐπιτήδειος καὶ φίλος', 'φίλος', 'Sóstrato era para mí cercano y amigo.'],
    ['νόμιζε', 'considerar, creer', 'verbo', 'sg', '', '2', '', 'νομίζω', 'νόμιζ᾽ ἀδελφοὺς τοὺς ἀληθινούς φίλους', 'νόμιζ᾽', 'Considera hermanos a los verdaderos amigos.'],
    ['ἀδελφούς', 'hermano', 'nombre', 'pl', 'ac', '', '', 'ἀδελφός, -οῦ', 'νόμιζ᾽ ἀδελφοὺς τοὺς ἀληθινούς φίλους', 'ἀδελφοὺς', 'Considera hermanos a los verdaderos amigos.'],
    ['ἀληθινούς', 'verdadero, genuino', 'adjetivo', 'pl', 'ac', '', '', 'ἀληθινός, -ή, -όν', 'νόμιζ᾽ ἀδελφοὺς τοὺς ἀληθινούς φίλους', 'ἀληθινούς', 'Considera hermanos a los verdaderos amigos.'],
    ['βίος', 'vida', 'nombre', 'sg', 'nom', '', '', 'βίος, -ου', 'ὁ μὲν οὖν βίος τοῦ σοφιστοῦ τοιοῦτος', 'βίος', 'Por un lado tal es, entonces, la vida del sofista.'],
    ['σοφιστοῦ', 'sofista', 'nombre', 'sg', 'gen', '', '', 'σοφιστής, -οῦ', 'ὁ μὲν οὖν βίος τοῦ σοφιστοῦ τοιοῦτος', 'σοφιστοῦ', 'Por un lado tal es, entonces, la vida del sofista.'],
    ['θεῷ', 'dios', 'nombre', 'sg', 'dat', '', '', 'θεός, -οῦ', 'θεῷ μάχεσθαι δεινόν ἐστι καὶ τύχῃ', 'θεῷ', 'Luchar contra un dios y contra la fortuna es terrible.'],
    ['μάχεσθαι', 'luchar, combatir', 'infinitivo', '', '', '', 'med', 'μάχομαι', 'θεῷ μάχεσθαι δεινόν ἐστι καὶ τύχῃ', 'μάχεσθαι', 'Luchar contra un dios y contra la fortuna es terrible.'],
    ['δεινόν', 'terrible, temible; hábil', 'adjetivo', 'sg', 'nom,ac,voc', '', '', 'δεινός, -ή, -όν', 'θεῷ μάχεσθαι δεινόν ἐστι καὶ τύχῃ', 'δεινόν', 'Luchar contra un dios y contra la fortuna es terrible.'],
    ['τύχῃ', 'fortuna, suerte, azar', 'nombre', 'sg', 'dat', '', '', 'τύχη, -ης', 'θεῷ μάχεσθαι δεινόν ἐστι καὶ τύχῃ', 'τύχῃ', 'Luchar contra un dios y contra la fortuna es terrible.'],
    ['προνοίας', 'previsión, providencia', 'nombre', 'sg', 'gen', '', '', 'πρόνοια, -ας', 'πάσης προνοίας ἡ τύχη δυνατωτέρα', 'προνοίας', 'La fortuna es más poderosa que toda previsión.'],
    ['τύχη', 'fortuna, suerte', 'nombre', 'sg', 'nom,voc', '', '', 'τύχη, -ης', 'πάσης προνοίας ἡ τύχη δυνατωτέρα', 'τύχη', 'La fortuna es más poderosa que toda previsión.'],
    ['δυνατωτέρα', 'poderoso, capaz; (comp.) más poderoso', 'adjetivo', 'sg', 'nom,voc', '', '', 'δυνατός, -ή, -όν (comp. δυνατώτερος)', 'πάσης προνοίας ἡ τύχη δυνατωτέρα', 'δυνατωτέρα', 'La fortuna es más poderosa que toda previsión.'],
    ['πολλῶν', 'mucho, numeroso', 'adjetivo', 'pl', 'gen', '', '', 'πολύς, πολλή, πολύ', 'πολλῶν ὁ χρόνος γίγνεται ἰατρός', 'πολλῶν', 'El tiempo se vuelve médico de muchas cosas.'],
    ['γίγνεται', 'llegar a ser, convertirse en; suceder', 'verbo', 'sg', '', '3', '', 'γίγνομαι', 'πολλῶν ὁ χρόνος γίγνεται ἰατρός', 'γίγνεται', 'El tiempo se vuelve médico de muchas cosas.'],
    ['ἰατρός', 'médico', 'nombre', 'sg', 'nom', '', '', 'ἰατρός, -οῦ', 'πολλῶν ὁ χρόνος γίγνεται ἰατρός', 'ἰατρός', 'El tiempo se vuelve médico de muchas cosas.'],
    ['ἐρωτικά', 'erótico, relativo al amor; (subst.) las cosas del amor', 'adjetivo', 'pl', 'nom,ac,voc', '', '', 'ἐρωτικός, -ή, -όν', 'ἡ δὲ ἐμὲ τὰ ἐρωτικὰ ἐδίδασκεν', 'ἐρωτικὰ', 'Y ella me enseñaba las cosas del amor.'],
    ['ἐδίδασκεν', 'enseñar', 'verbo', 'sg', '', '3', '', 'διδάσκω', 'ἡ δὲ ἐμὲ τὰ ἐρωτικὰ ἐδίδασκεν', 'ἐδίδασκεν', 'Y ella me enseñaba las cosas del amor.'],
    ['ζῷα', 'animal, ser viviente', 'nombre', 'pl', 'nom,ac,voc', '', '', 'ζῷον, -ου', 'τὰ ζῷα πείθεται τῷ ἀνθρώπῳ', 'ζῷα', 'Los animales obedecen al ser humano.'],
    ['πείθεται', 'persuadir; (mediopasivo) obedecer', 'verbo', 'sg', '', '3', '', 'πείθω', 'τὰ ζῷα πείθεται τῷ ἀνθρώπῳ', 'πείθεται', 'Los animales obedecen al ser humano.'],
    ['ἀνθρώπῳ', 'hombre, ser humano', 'nombre', 'sg', 'dat', '', '', 'ἄνθρωπος, -ου', 'τὰ ζῷα πείθεται τῷ ἀνθρώπῳ', 'ἀνθρώπῳ', 'Los animales obedecen al ser humano.'],
    ['μαθητῇ', 'estudiante, discípulo', 'nombre', 'sg', 'dat', '', '', 'μαθητής, -οῦ', 'τῷ μαθητῇ πρέπει σπουδή', 'μαθητῇ', 'Al estudiante le corresponde el esfuerzo.'],
    ['πρέπει', 'convenir, corresponder, ser apropiado', 'verbo', 'sg', '', '3', '', 'πρέπει', 'τῷ μαθητῇ πρέπει σπουδή', 'πρέπει', 'Al estudiante le corresponde el esfuerzo.'],
    ['σπουδή', 'esfuerzo, celo, diligencia', 'nombre', 'sg', 'nom,voc', '', '', 'σπουδή, -ῆς', 'τῷ μαθητῇ πρέπει σπουδή', 'σπουδή', 'Al estudiante le corresponde el esfuerzo.'],
    ['ἔλαφοι', 'ciervo', 'nombre', 'pl', 'nom,voc', '', '', 'ἔλαφος, -ου', 'οἱ ἔλαφοι ἔθνησκον δίψῃ', 'ἔλαφοι', 'Los ciervos morían de sed.'],
    ['ἔθνησκον', 'morir', 'verbo', 'pl', '', '3', '', 'θνῄσκω', 'οἱ ἔλαφοι ἔθνησκον δίψῃ', 'ἔθνησκον', 'Los ciervos morían de sed.'],
    ['δίψῃ', 'sed', 'nombre', 'sg', 'dat', '', '', 'δίψα, -ης', 'οἱ ἔλαφοι ἔθνησκον δίψῃ', 'δίψῃ', 'Los ciervos morían de sed.'],
    ['ὁδόν', 'camino, vía', 'nombre', 'sg', 'ac', '', '', 'ὁδός, -οῦ', 'παρὰ τὴν ὁδὸν κρήνη ἦν', 'ὁδὸν', 'Había una fuente junto al camino.'],
    ['κρήνη', 'fuente, manantial', 'nombre', 'sg', 'nom,voc', '', '', 'κρήνη, -ης', 'παρὰ τὴν ὁδὸν κρήνη ἦν', 'κρήνη', 'Había una fuente junto al camino.'],
    ['λύσω', 'liberar, desatar, soltar', 'verbo', 'sg', '', '1', '', 'λύω', 'ἐγὼ δὲ λύσω τόνδε', 'λύσω', 'Yo, por mi parte, lo liberaré.'],
    ['μῦθος', 'relato, fábula, mito', 'nombre', 'sg', 'nom', '', '', 'μῦθος, -ου', 'Ὁ μῦθος δηλοῖ ὅτι τῶν πονηρῶν αἱ χάριτες φοβεραί εἰσιν', 'μῦθος', 'El relato muestra que los favores de los malvados son temibles.'],
    ['δηλοῖ', 'mostrar, revelar, indicar', 'verbo', 'sg', '', '3', '', 'δηλόω', 'Ὁ μῦθος δηλοῖ ὅτι τῶν πονηρῶν αἱ χάριτες φοβεραί εἰσιν', 'δηλοῖ', 'El relato muestra que los favores de los malvados son temibles.'],
    ['πονηρῶν', 'malvado, vil, malo', 'adjetivo', 'pl', 'gen', '', '', 'πονηρός, -ά, -όν', 'Ὁ μῦθος δηλοῖ ὅτι τῶν πονηρῶν αἱ χάριτες φοβεραί εἰσιν', 'πονηρῶν', 'El relato muestra que los favores de los malvados son temibles.'],
    ['χάριτες', 'favor, gracia', 'nombre', 'pl', 'nom,voc', '', '', 'χάρις, -ιτος', 'Ὁ μῦθος δηλοῖ ὅτι τῶν πονηρῶν αἱ χάριτες φοβεραί εἰσιν', 'χάριτες', 'El relato muestra que los favores de los malvados son temibles.'],
    ['φοβεραί', 'temible, terrible', 'adjetivo', 'pl', 'nom,voc', '', '', 'φοβερός, -ά, -όν', 'Ὁ μῦθος δηλοῖ ὅτι τῶν πονηρῶν αἱ χάριτες φοβεραί εἰσιν', 'φοβεραί', 'El relato muestra que los favores de los malvados son temibles.'],
  ];

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  Logger.log('Listo: ' + rows.length + ' formas nuevas agregadas a "Vocabulario". Planilla: ' + ss.getUrl());
}

// ---------------------------------------------------------------------------
// Correr UNA sola vez, a mano, DESPUÉS de poblarVocabularioAmpliado(). Agrega la
// columna "Semana" (si no existe) y le pone valor a cada fila de "Vocabulario"
// según de qué frase salió (comparando por ContextGreek, columna I) — así el
// selector de nivel de Repaso nonstop puede filtrar por "Semana 1/2/3" o "Todo".
// Las frases que no están en el mapa (material de Griego I / cursada 2025) quedan
// con Semana vacío a propósito: solo aparecen en el nivel "Todo". Segura para
// correr más de una vez: vuelve a calcular todo desde cero cada vez, no acumula.
// ---------------------------------------------------------------------------
function migrarColumnaSemana() {
  const ss = SpreadsheetApp.openById(getOrCreateSheetId());
  const sheet = ss.getSheetByName(VOCAB_SHEET_NAME);
  if (!sheet) { Logger.log('No existe la hoja "Vocabulario".'); return; }

  const header = sheet.getRange(1, 1, 1, 12).getValues()[0];
  if (header[11] !== 'Semana') {
    sheet.getRange(1, 12).setValue('Semana');
    sheet.getRange(1, 12).setFontWeight('bold');
  }

  // Semana 1 = Frases 1 completo + Frases 2 (ítems 1-7). Semana 2 = Frases 2 (ítem 8)
  // + las 7 frases del Práctico 2 (20.08) sumadas el 2026-08-31. Todo lo demás (Griego I
  // general, cursada 2025) queda sin semana. Confirmado con Santiago el 2026-08-31.
  const SEMANA_POR_FRASE = {
    'λόγος ἐστὶ φωνὴ σημαντική': '1',
    'ἐκ καρποῦ δένδρον γιγνώσκετε': '1',
    'ὁ ὕπνος τοῦ θανάτου ἀδελφός ἐστιν': '1',
    'φασὶ γὰρ τὴν τοῦ ἀνθρώπου ψυχὴν εἶναι ἀθάνατον': '1',
    'Πίνδαρος ὁ ποιητὴς τὸν οἶνον καλέει δρόσον τῆς ἀμπέλου': '1',
    'λόγος κάτοπτρον τοῦ θυμοῦ': '1',
    'ἡ ψυχὴ καὶ αὐτή, Ἔρως σχέτλιε, ἔχει πτέρυγας': '1',
    'τῷ ἀνθρώπῳ σῶμα μὲν θνητόν, ψυχὴ δὲ ἀθάνατος': '1',
    'τὰ χαλεπὰ τῶν πραγμάτων ὁ χρόνος διαλύει': '1',
    'κακὴ ἡ δελφῖνος ἐν χέρσῳ βία': '1',
    'ἐξ ὀνύχων τὸν λέοντα γιγνώσκειν ἔξεστι': '1',
    'δὶς παῖδες οἱ γέροντες': '1',
    'ἡ Δίκη γὰρ πάντα βλέπει': '1',
    'Ἔνιοι κακῶς φρονοῦσι πράττοντες καλῶς': '2',
    'Ἀνὴρ ὁ φεύγων καὶ πάλιν μαχήσεται': '2',
    'Ἄνθρωπος ὢν γίνωσκε τῆς ὀργῆς κρατεῖν': '2',
    'Πολλοὶ κακῶς πράττουσιν οὐκ ὄντες κακοί': '2',
    'ὄνος ξύλων γόμον φέρων λίμνην διέβαινεν': '2',
    'δελφῖνα νήχεσθαι διδάσκεις': '2',
    'φησὶ γάρ που Πρωταγόρας πάντων χρημάτων μέτρον ἄνθρωπον εἶναι': '2',
    'οἱ Πέρσαι διδάσκουσι τοὺς παῖδας πείθεσθαι τοῖς ἄρχουσιν': '2'
  };

  const data = sheet.getDataRange().getValues();
  let updated = 0;
  for (let i = 1; i < data.length; i++) {
    const contextGreek = data[i][8]; // columna I
    const semana = SEMANA_POR_FRASE.hasOwnProperty(contextGreek) ? SEMANA_POR_FRASE[contextGreek] : '';
    if (data[i][11] !== semana) {
      sheet.getRange(i + 1, 12).setValue(semana);
      updated++;
    }
  }
  Logger.log('Semana actualizada en ' + updated + ' filas.');
}
