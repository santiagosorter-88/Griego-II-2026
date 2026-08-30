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

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, filas: rows.length }))
    .setMimeType(ContentService.MimeType.JSON);
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
// CorrectCases: casos separados por coma (ej. "nom,voc"). Kind: nombre / participio / verbo / infinitivo.
// ---------------------------------------------------------------------------
const VOCAB_SHEET_NAME = 'Vocabulario';

function ensureVocabHeader_(sheet) {
  if (sheet.getLastRow() > 0) return;
  var header = ['Greek', 'Meaning', 'Kind', 'Number', 'CorrectCases', 'Person', 'Voice', 'Citation', 'ContextGreek', 'ContextWord', 'ContextSpanish'];
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.getRange(1, 1, 1, header.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

// doGet rutea por el parámetro ?tipo= — así un solo Apps Script sirve el banco de
// vocabulario (Repaso nonstop) Y el contenido de la página de entrada (index.html),
// sin duplicar backend. Sin parámetro, por compatibilidad con lo ya desplegado, sirve
// vocabulario (repaso_nonstop.html no manda parámetro).
function doGet(e) {
  const tipo = (e && e.parameter && e.parameter.tipo) || 'vocabulario';
  if (tipo === 'indice') return doGetIndice_();
  return doGetVocabulario_();
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
