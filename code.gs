/**
 * TRENING — дневник силовых тренировок
 * Серверная часть (Google Apps Script Web App)
 *
 * Скрипт должен быть ПРИВЯЗАН к таблице «Trening»
 * (открыть таблицу → Расширения → Apps Script).
 *
 * Листы находятся автоматически: сначала по имени (без учёта регистра,
 * пробелов и опечаток вида Workauts), затем по набору столбцов.
 */

var TZ = Session.getScriptTimeZone();
var SHEET_CACHE = {};

/* ============================== WEB APP ============================== */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Trening')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ============================ УТИЛИТЫ ================================ */

function ss_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('Скрипт не привязан к таблице. Откройте таблицу «Trening» → Расширения → Apps Script.');
  return ss;
}

/** Нормализация: нижний регистр, без пробелов/подчёркиваний/дефисов. */
function norm_(s) {
  return String(s).toLowerCase().replace(/[\s_\-–—]/g, '');
}

/**
 * Поиск листа:
 *  1) по имени (нечётко: 'Workouts' найдёт и 'workouts', и ' Workauts ', и 'Workout');
 *  2) по обязательным заголовкам столбцов (mustHave), исключая mustNotHave.
 */
function findSheet_(key, nameHints, mustHave, mustNotHave) {
  if (SHEET_CACHE[key]) return SHEET_CACHE[key];
  var sheets = ss_().getSheets();
  var i, j;

  // 1. точное нечёткое совпадение имени
  for (i = 0; i < sheets.length; i++) {
    var n = norm_(sheets[i].getName());
    for (j = 0; j < nameHints.length; j++) {
      if (n === norm_(nameHints[j])) { SHEET_CACHE[key] = sheets[i]; return sheets[i]; }
    }
  }
  // 2. имя начинается с подсказки (Workouts2, Workauts (копия) и т.п.)
  for (i = 0; i < sheets.length; i++) {
    var n2 = norm_(sheets[i].getName());
    for (j = 0; j < nameHints.length; j++) {
      if (n2.indexOf(norm_(nameHints[j])) === 0 || norm_(nameHints[j]).indexOf(n2) === 0) {
        if (headersOk_(sheets[i], mustHave, mustNotHave)) { SHEET_CACHE[key] = sheets[i]; return sheets[i]; }
      }
    }
  }
  // 3. по заголовкам столбцов
  for (i = 0; i < sheets.length; i++) {
    if (headersOk_(sheets[i], mustHave, mustNotHave)) { SHEET_CACHE[key] = sheets[i]; return sheets[i]; }
  }

  var names = sheets.map(function (s) { return '«' + s.getName() + '»'; }).join(', ');
  throw new Error('Не найден лист ' + key + '. В таблице есть листы: ' + names +
    '. Нужны столбцы: ' + mustHave.join(', ') + '.');
}

function headersOk_(sh, mustHave, mustNotHave) {
  if (sh.getLastRow() < 1 || sh.getLastColumn() < 1) return false;
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm_);
  for (var i = 0; i < mustHave.length; i++) {
    if (head.indexOf(norm_(mustHave[i])) < 0) return false;
  }
  if (mustNotHave) {
    for (var k = 0; k < mustNotHave.length; k++) {
      if (head.indexOf(norm_(mustNotHave[k])) >= 0) return false;
    }
  }
  return true;
}

function shWorkouts_() {
  return findSheet_('Workouts',
    ['Workouts', 'Workout', 'Workauts', 'Тренировки'],
    ['Workout_iD', 'FocusArea'], ['SetID']);
}
function shSets_() {
  return findSheet_('WorkoutSets',
    ['WorkoutSets', 'WorkoutSet', 'Sets', 'Подходы'],
    ['SetID', 'Workout_iD', 'Exercise'], null);
}
function shExercises_() {
  return findSheet_('Exercises',
    ['Exercises', 'Exercise', 'Упражнения'],
    [], ['SetID', 'Workout_iD']);
}

/** Карта "нормализованное имя столбца → индекс (0-based)". */
function headerMap_(sh) {
  var row = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
  var map = {};
  row.forEach(function (v, i) {
    var k = norm_(v);
    if (k && !(k in map)) map[k] = i;
  });
  return map;
}

function col_(map, name, sheetName) {
  var k = norm_(name);
  if (!(k in map)) throw new Error('В листе «' + sheetName + '» не найден столбец «' + name + '».');
  return map[k];
}

/** Любое значение даты → строка 'yyyy-MM-dd'. */
function isoDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  var s = String(v).trim();
  var m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})/);              // dd.MM.yyyy
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);                     // yyyy-MM-dd
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  var d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
  return '';
}

/** 'yyyy-MM-dd' → Date (локальная полночь), чтобы в ячейке была именно дата. */
function fromIso_(iso) {
  var p = String(iso).split('-');
  return new Date(+p[0], +p[1] - 1, +p[2]);
}

/* ======================== ЧТЕНИЕ ДАННЫХ ============================== */

/** Определение столбцов листа Exercises по заголовкам. */
function exCols_(headRow) {
  var head = headRow.map(function (h) { return String(h).toLowerCase(); });
  var exCol = -1, grCol = -1, phCol = -1, tpCol = -1;
  head.forEach(function (h, i) {
    if (exCol < 0 && /(exerc|упраж|назван|name)/.test(h)) exCol = i;
    if (grCol < 0 && /(muscle|group|мышц|групп)/.test(h)) grCol = i;
    if (phCol < 0 && /(photo|image|img|picture|url|link|video|gif|фото|картин|изображ|ссыл|видео)/.test(h)) phCol = i;
    if (tpCol < 0 && /(^тип|type|вид)/.test(h)) tpCol = i;
  });
  if (exCol < 0) exCol = 0;
  if (grCol < 0) grCol = (exCol === 0 ? 1 : 0);
  return { ex: exCol, gr: grCol, ph: phCol, tp: tpCol };
}

/** Справочник упражнений. Столбцы определяются по заголовкам. */
function readExercises_() {
  var sh = shExercises_();
  var vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];

  var cols = exCols_(vals[0]);

  var out = [];
  for (var r = 1; r < vals.length; r++) {
    var name = String(vals[r][cols.ex]).trim();
    var group = String(vals[r][cols.gr]).trim();
    var photo = cols.ph >= 0 ? String(vals[r][cols.ph]).trim() : '';
    var type = cols.tp >= 0 ? String(vals[r][cols.tp]).trim() : '';
    if (name) out.push({ name: name, group: group, photo: photo, type: type });
  }
  return out;
}

/**
 * Фото упражнения из вашей базы: сервер сам скачивает файл
 * (Google Диск или прямая ссылка) и отдаёт приложению как data-URL.
 * Так картинка показывается внутри приложения без сторонних сайтов.
 */
function getPhotoData(exName) {
  var ex = null;
  readExercises_().forEach(function (e) { if (e.name === String(exName)) ex = e; });
  if (!ex || !ex.photo) return { error: 'Для упражнения не указана ссылка на фото.' };

  var url = ex.photo;
  var blob = null;
  try {
    if (/drive\.google\.com|docs\.google\.com/i.test(url)) {
      var m = url.match(/[-\w]{25,}/);                 // ID файла на Диске
      if (!m) return { error: 'Не удалось распознать ссылку Google Диска.' };
      blob = DriveApp.getFileById(m[0]).getBlob();
    } else if (/^https?:\/\//i.test(url)) {
      var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
      if (resp.getResponseCode() >= 400) return { error: 'Файл по ссылке недоступен (' + resp.getResponseCode() + ').' };
      blob = resp.getBlob();
    } else {
      return { error: 'Ссылка должна начинаться с https:// или вести на Google Диск.' };
    }
  } catch (e2) {
    return { error: 'Не удалось загрузить фото: ' + e2.message };
  }

  var bytes = blob.getBytes();
  if (bytes.length > 4.5 * 1024 * 1024) return { error: 'Файл больше 4,5 МБ — уменьшите картинку.' };
  var ct = blob.getContentType() || 'image/jpeg';
  if (ct.indexOf('image/') !== 0) ct = 'image/jpeg';
  return { data: 'data:' + ct + ';base64,' + Utilities.base64Encode(bytes) };
}

/**
 * Добавление нового упражнения в справочник Exercises из приложения.
 * Возвращает свежий getAppData(). Дубликаты (без учёта регистра) не создаются.
 */
function addExercise(name, group) {
  name = String(name || '').trim();
  group = String(group || '').trim();
  if (!name) throw new Error('Укажите название упражнения.');
  if (!group) throw new Error('Укажите группу мышц.');

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = shExercises_();
    var vals = sh.getDataRange().getValues();
    var cols = exCols_(vals[0]);

    for (var r = 1; r < vals.length; r++) {
      if (norm_(vals[r][cols.ex]) === norm_(name)) return getAppData(); // уже есть
    }

    var rowArr = new Array(Math.max(sh.getLastColumn(), cols.gr + 1)).fill('');
    rowArr[cols.ex] = name;
    rowArr[cols.gr] = group;
    sh.appendRow(rowArr);
    return getAppData();
  } finally {
    lock.releaseLock();
  }
}

function readWorkouts_() {
  var sh = shWorkouts_();
  var map = headerMap_(sh);
  var cDate = col_(map, 'Date', 'Workouts');
  var cFocus = col_(map, 'FocusArea', 'Workouts');
  var cNotes = col_(map, 'Notes', 'Workouts');
  var cId = col_(map, 'Workout_iD', 'Workouts');

  var vals = sh.getDataRange().getValues();
  var out = [];
  for (var r = 1; r < vals.length; r++) {
    if (String(vals[r][cId]).trim() === '') continue;
    out.push({
      row: r + 1,
      id: Number(vals[r][cId]),
      date: isoDate_(vals[r][cDate]),
      focus: String(vals[r][cFocus] || ''),
      notes: String(vals[r][cNotes] || '')
    });
  }
  return out;
}

function readSets_() {
  var sh = shSets_();
  var map = headerMap_(sh);
  var c = {
    setId: col_(map, 'SetID', 'WorkoutSets'),
    wId: col_(map, 'Workout_iD', 'WorkoutSets'),
    date: col_(map, 'Date', 'WorkoutSets'),
    group: col_(map, 'MuscleGroup', 'WorkoutSets'),
    ex: col_(map, 'Exercise', 'WorkoutSets'),
    weight: col_(map, 'Weight', 'WorkoutSets'),
    reps: col_(map, 'Reps', 'WorkoutSets')
  };
  var rng = sh.getDataRange();
  var vals = rng.getValues();
  var disp = rng.getDisplayValues();
  var out = [];
  for (var r = 1; r < vals.length; r++) {
    if (String(vals[r][c.setId]).trim() === '') continue;
    out.push({
      row: r + 1,
      setId: Number(vals[r][c.setId]),
      workoutId: Number(vals[r][c.wId]),
      date: isoDate_(vals[r][c.date]),
      group: String(vals[r][c.group] || ''),
      exercise: String(vals[r][c.ex] || ''),
      weight: numCell_(vals[r][c.weight], disp[r][c.weight]),
      reps: numCell_(vals[r][c.reps], disp[r][c.reps])
    });
  }
  return out;
}

/**
 * Число из ячейки. Если Google Таблица распознала «7.5» как ДАТУ
 * (частая проблема в русской локали), берём отображаемое значение
 * ячейки («7.5») и превращаем его в число.
 */
function numCell_(v, d) {
  if (v === '' || v === null || v === undefined) return '';
  if (typeof v === 'number') return v;
  var s = String(d).trim().replace(/\s/g, '').replace(',', '.');
  var m = s.match(/-?\d+(\.\d+)?/);
  if (m) {
    var n = parseFloat(m[0]);
    if (isFinite(n)) return n;
  }
  return '';
}

/** Все данные для клиента одним вызовом. */
function getAppData() {
  var sets = readSets_();
  var workouts = readWorkouts_();

  var counts = {};
  sets.forEach(function (s) { counts[s.workoutId] = (counts[s.workoutId] || 0) + 1; });

  var w = workouts.map(function (x) {
    return { id: x.id, date: x.date, focus: x.focus, notes: x.notes, setCount: counts[x.id] || 0 };
  });
  w.sort(function (a, b) { return a.date < b.date ? 1 : (a.date > b.date ? -1 : 0); });

  var s = sets.map(function (x) {
    return { setId: x.setId, workoutId: x.workoutId, date: x.date, group: x.group,
             exercise: x.exercise, weight: x.weight, reps: x.reps };
  });
  s.sort(function (a, b) { return a.setId - b.setId; });

  return { exercises: readExercises_(), workouts: w, sets: s,
           today: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd') };
}

/* ===================== БИЗНЕС-ЛОГИКА (раздел 5 ТЗ) =================== */

/** Workout_iD по дате: существующий или новый (MAX+1). При необходимости создаёт строку в Workouts. */
function ensureWorkout_(dateIso) {
  var workouts = readWorkouts_();
  for (var i = 0; i < workouts.length; i++) {
    if (workouts[i].date === dateIso) return workouts[i].id;   // одна дата = одна тренировка
  }
  var maxId = 0;
  workouts.forEach(function (w) { if (w.id > maxId) maxId = w.id; });
  var newId = maxId + 1;

  var sh = shWorkouts_();
  var map = headerMap_(sh);
  var rowArr = new Array(sh.getLastColumn()).fill('');
  rowArr[col_(map, 'Date', 'Workouts')] = fromIso_(dateIso);
  rowArr[col_(map, 'Workout_iD', 'Workouts')] = newId;
  sh.appendRow(rowArr);

  // Идентификатор — именно число (5.4)
  var r = sh.getLastRow();
  sh.getRange(r, col_(map, 'Workout_iD', 'Workouts') + 1).setNumberFormat('0');
  sh.getRange(r, col_(map, 'Date', 'Workouts') + 1).setNumberFormat('dd.mm.yyyy');
  return newId;
}

/** Пересчитывает FocusArea; удаляет тренировки, оставшиеся без подходов (если нет заметок). */
function refreshWorkouts_(workoutIds) {
  if (!workoutIds || !workoutIds.length) return;
  var uniq = {};
  workoutIds.forEach(function (id) { if (id || id === 0) uniq[id] = true; });

  var sets = readSets_();
  var groupsByWorkout = {};
  sets.forEach(function (s) {
    if (!groupsByWorkout[s.workoutId]) groupsByWorkout[s.workoutId] = [];
    if (s.group && groupsByWorkout[s.workoutId].indexOf(s.group) < 0) groupsByWorkout[s.workoutId].push(s.group);
  });

  var sh = shWorkouts_();
  var map = headerMap_(sh);
  var cFocus = col_(map, 'FocusArea', 'Workouts') + 1;
  var workouts = readWorkouts_();

  var toDelete = [];
  workouts.forEach(function (w) {
    if (!uniq[w.id]) return;
    var hasSets = sets.some(function (s) { return s.workoutId === w.id; });
    if (hasSets) {
      sh.getRange(w.row, cFocus).setValue((groupsByWorkout[w.id] || []).join(', '));
    } else if (!String(w.notes).trim()) {
      toDelete.push(w.row);                    // пустая тренировка без заметок — убрать
    } else {
      sh.getRange(w.row, cFocus).setValue('');
    }
  });
  toDelete.sort(function (a, b) { return b - a; }).forEach(function (r) { sh.deleteRow(r); });
}

/* ========================= ЗАПИСЬ ДАННЫХ ============================= */

/**
 * Сохранение подхода (добавление или редактирование).
 * payload: { setId|null, date:'yyyy-MM-dd', group, exercise, weight, reps }
 * Возвращает свежий getAppData().
 */
function saveSet(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var dateIso = isoDate_(payload.date);
    if (!dateIso) throw new Error('Некорректная дата.');
    var weight = Number(payload.weight);
    var reps = Number(payload.reps);
    if (isNaN(weight) || isNaN(reps)) throw new Error('Вес и повторения должны быть числами.');
    if (!payload.exercise) throw new Error('Не выбрано упражнение.');

    var workoutId = ensureWorkout_(dateIso);   // 5.2 / 5.3

    var sh = shSets_();
    var map = headerMap_(sh);
    var c = {
      setId: col_(map, 'SetID', 'WorkoutSets'),
      wId: col_(map, 'Workout_iD', 'WorkoutSets'),
      date: col_(map, 'Date', 'WorkoutSets'),
      group: col_(map, 'MuscleGroup', 'WorkoutSets'),
      ex: col_(map, 'Exercise', 'WorkoutSets'),
      weight: col_(map, 'Weight', 'WorkoutSets'),
      reps: col_(map, 'Reps', 'WorkoutSets')
    };

    var touched = [workoutId];

    if (payload.setId === null || payload.setId === undefined || payload.setId === '') {
      /* --- новый подход: SetID = MAX + 1 (5.1) --- */
      var maxId = 0;
      readSets_().forEach(function (s) { if (s.setId > maxId) maxId = s.setId; });
      var newSetId = maxId + 1;

      var rowArr = new Array(sh.getLastColumn()).fill('');
      rowArr[c.setId] = newSetId;
      rowArr[c.wId] = workoutId;
      rowArr[c.date] = fromIso_(dateIso);
      rowArr[c.group] = String(payload.group || '');
      rowArr[c.ex] = String(payload.exercise);
      rowArr[c.weight] = weight;
      rowArr[c.reps] = reps;
      sh.appendRow(rowArr);

      var r = sh.getLastRow();                 // числовые форматы (5.4)
      sh.getRange(r, c.setId + 1).setNumberFormat('0');
      sh.getRange(r, c.wId + 1).setNumberFormat('0');
      sh.getRange(r, c.weight + 1).setNumberFormat('0.###');
      sh.getRange(r, c.reps + 1).setNumberFormat('0');
      sh.getRange(r, c.date + 1).setNumberFormat('dd.mm.yyyy');
    } else {
      /* --- редактирование существующего --- */
      var target = null;
      readSets_().forEach(function (s) { if (s.setId === Number(payload.setId)) target = s; });
      if (!target) throw new Error('Подход SetID=' + payload.setId + ' не найден.');
      touched.push(target.workoutId);

      var rng = sh.getRange(target.row, 1, 1, sh.getLastColumn());
      var vals = rng.getValues()[0];
      vals[c.wId] = workoutId;
      vals[c.date] = fromIso_(dateIso);
      vals[c.group] = String(payload.group || '');
      vals[c.ex] = String(payload.exercise);
      vals[c.weight] = weight;
      vals[c.reps] = reps;
      rng.setValues([vals]);
      sh.getRange(target.row, c.wId + 1).setNumberFormat('0');
      sh.getRange(target.row, c.setId + 1).setNumberFormat('0');
      sh.getRange(target.row, c.weight + 1).setNumberFormat('0.###');
      sh.getRange(target.row, c.reps + 1).setNumberFormat('0');
      sh.getRange(target.row, c.date + 1).setNumberFormat('dd.mm.yyyy');
    }

    refreshWorkouts_(touched);                 // FocusArea + целостность (5.3, 5.4)
    return getAppData();
  } finally {
    lock.releaseLock();
  }
}

/** Удаление подхода по SetID. Возвращает свежий getAppData(). */
function deleteSet(setId) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var target = null;
    readSets_().forEach(function (s) { if (s.setId === Number(setId)) target = s; });
    if (!target) throw new Error('Подход SetID=' + setId + ' не найден.');

    shSets_().deleteRow(target.row);
    refreshWorkouts_([target.workoutId]);
    return getAppData();
  } finally {
    lock.releaseLock();
  }
}

/**
 * Удаление ВСЕЙ тренировки: все её подходы в WorkoutSets
 * и сама строка в Workouts. Возвращает свежий getAppData().
 */
function deleteWorkout(workoutId) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    workoutId = Number(workoutId);

    var shS = shSets_();
    readSets_()
      .filter(function (s) { return s.workoutId === workoutId; })
      .map(function (s) { return s.row; })
      .sort(function (a, b) { return b - a; })
      .forEach(function (r) { shS.deleteRow(r); });

    var shW = shWorkouts_();
    readWorkouts_()
      .filter(function (w) { return w.id === workoutId; })
      .map(function (w) { return w.row; })
      .sort(function (a, b) { return b - a; })
      .forEach(function (r) { shW.deleteRow(r); });

    return getAppData();
  } finally {
    lock.releaseLock();
  }
}

/** Диагностика: запустить в редакторе (Выполнить), посмотреть журнал. */
function diagnose() {
  var sheets = ss_().getSheets();
  sheets.forEach(function (sh) {
    var head = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
    Logger.log('Лист «%s»: %s', sh.getName(), head.join(' | '));
  });
  Logger.log('Workouts → «%s»', shWorkouts_().getName());
  Logger.log('WorkoutSets → «%s»', shSets_().getName());
  Logger.log('Exercises → «%s»', shExercises_().getName());
}

/**
 * ПОЧИНКА ДАННЫХ (запустить ОДИН РАЗ вручную из редактора: выбрать
 * repairNumbers в списке функций → «Выполнить»).
 * Находит в столбцах Weight и Reps ячейки, которые Google Таблица
 * распознала как даты (например «7.5» → 7 мая), и заменяет их
 * настоящими числами. Также ставит числовой формат на все столбцы.
 */
function repairNumbers() {
  var sh = shSets_();
  var map = headerMap_(sh);
  var cols = [
    col_(map, 'Weight', 'WorkoutSets') + 1,
    col_(map, 'Reps', 'WorkoutSets') + 1
  ];
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  var fixed = 0;
  cols.forEach(function (col) {
    var rng = sh.getRange(2, col, lastRow - 1, 1);
    var vals = rng.getValues();
    var disp = rng.getDisplayValues();
    for (var i = 0; i < vals.length; i++) {
      var v = vals[i][0];
      if (v instanceof Date || (typeof v === 'string' && v !== '')) {
        var n = numCell_(v, disp[i][0]);
        if (n !== '') {
          sh.getRange(i + 2, col).setValue(n);
          fixed++;
        }
      }
    }
    rng.setNumberFormat(col === cols[0] ? '0.###' : '0');
  });

  // числовой формат для идентификаторов
  sh.getRange(2, col_(map, 'SetID', 'WorkoutSets') + 1, lastRow - 1, 1).setNumberFormat('0');
  sh.getRange(2, col_(map, 'Workout_iD', 'WorkoutSets') + 1, lastRow - 1, 1).setNumberFormat('0');

  Logger.log('Исправлено ячеек: %s', fixed);
}