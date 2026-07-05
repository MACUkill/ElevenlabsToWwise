/**
 * ============================================================
 *  VO System — Apps Script v2 (Thug Life)  |  Faza 0a
 * ============================================================
 *  Instalacja: Google Sheets → Rozszerzenia → Apps Script →
 *  wklej CAŁOŚĆ zamiast starego kodu → zapisz → odśwież arkusz.
 *
 *  Spec: raport "Dialogue & VO System.md" (§6.1, §14.0, Aneks B)
 *
 *  Zasady twarde:
 *   - Line_ID nadajemy TYLKO pustym wierszom (ID nieśmiertelne).
 *   - Line_ID = {Speaker}_{Type}_{Context}_{NN}
 *       Bark/Effort : NN = auto-increment w grupie Speaker×Type×Context (01–99)
 *       Story/Conv  : NN = Sequence_Order sformatowany do 3 cyfr (010, 020…)
 *   - Walidacja przed nadaniem: Type z listy, Speaker z arkusza Casting,
 *     Context wg klucza Aneks B + słownika Vocabulary. Błąd = brak ID + notatka
 *     na komórce Line_ID (nigdy nadpisanie).
 *
 *  Arkusze meta (wymagane):
 *   - "Casting":    nagłówki min. Speaker | Voice_ID | Gender | EffortSet
 *   - "Vocabulary": nagłówki Intent | Quest | Topic (listy w kolumnach)
 * ============================================================
 */

const TYPES = ['Bark', 'Effort', 'Story', 'ConvB', 'ConvH'];
const SEQUENCE_TYPES = ['Story', 'ConvB', 'ConvH'];
const EFFORT_POOL_SPEAKERS = ['Male', 'Female']; // etap 1 EffortSet (§7.4); Player jest w Casting
const META_SHEETS = ['Casting', 'Vocabulary'];

// Aliasy nagłówków — akceptujemy stare nazwy kolumn na czas migracji
const HEADER_ALIASES = {
  'Character_Name': 'Speaker',
  'Category': 'Type',
  'Intent_Tag': 'Context',
};

// ------------------------------------------------------------
// MENU
// ------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎙️ VO System')
    .addItem('✅ Wygeneruj Line_ID w TYM arkuszu', 'generateIdsActiveSheet')
    .addItem('🔢 Autonumeruj zaznaczony blok (co 10)', 'autoNumberSelection')
    .addItem('🔍 Waliduj wszystkie arkusze (bez zmian)', 'validateAllSheets')
    .addToUi();
}

// ------------------------------------------------------------
// NAGŁÓWKI / DANE POMOCNICZE
// ------------------------------------------------------------
function getHeaderMap_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return {};
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    let name = String(h || '').trim();
    if (!name) return;
    if (HEADER_ALIASES[name]) name = HEADER_ALIASES[name];
    if (!(name in map)) map[name] = i; // pierwsza wygrywa
  });
  return map;
}

function loadCastingSpeakers_(ss) {
  const sh = ss.getSheetByName('Casting');
  if (!sh) return null;
  const map = getHeaderMap_(sh);
  if (!('Speaker' in map)) return null;
  const vals = sh.getRange(2, map['Speaker'] + 1, Math.max(sh.getLastRow() - 1, 0), 1).getValues();
  const set = new Set();
  vals.forEach(r => { const v = String(r[0] || '').trim(); if (v) set.add(v); });
  return set;
}

function loadVocabulary_(ss) {
  const out = { intents: new Set(), quests: new Set(), topics: new Set() };
  const sh = ss.getSheetByName('Vocabulary');
  if (!sh) return null;
  const map = getHeaderMap_(sh);
  const readCol = (header, target) => {
    if (!(header in map)) return;
    const vals = sh.getRange(2, map[header] + 1, Math.max(sh.getLastRow() - 1, 0), 1).getValues();
    vals.forEach(r => { const v = String(r[0] || '').trim(); if (v) target.add(v); });
  };
  readCol('Intent', out.intents);
  readCol('Quest', out.quests);
  readCol('Topic', out.topics);
  return out;
}

/** Zbiera WSZYSTKIE istniejące Line_ID w całym pliku (globalna unikalność)
 *  + max NN per grupa barków/effortów. */
function collectExisting_(ss) {
  const ids = new Set();
  const groupMaxNN = {}; // "Speaker_Type_Context" -> max int NN
  ss.getSheets().forEach(sheet => {
    if (META_SHEETS.includes(sheet.getName())) return;
    const map = getHeaderMap_(sheet);
    if (!('Line_ID' in map)) return;
    const n = sheet.getLastRow() - 1;
    if (n < 1) return;
    const vals = sheet.getRange(2, map['Line_ID'] + 1, n, 1).getValues();
    vals.forEach(r => {
      const id = String(r[0] || '').trim();
      if (!id) return;
      ids.add(id);
      const m = id.match(/^(.+)_(\d{2,3})$/);
      if (m) {
        const nn = parseInt(m[2], 10);
        if (!(m[1] in groupMaxNN) || nn > groupMaxNN[m[1]]) groupMaxNN[m[1]] = nn;
      }
    });
  });
  return { ids: ids, groupMaxNN: groupMaxNN };
}

// ------------------------------------------------------------
// WALIDACJA WIERSZA  → { errors:[], warnings:[] }
// ------------------------------------------------------------
function validateRow_(row, casting, vocab) {
  const errors = [];
  const warnings = [];
  const speaker = row.speaker, type = row.type, context = row.context;

  // Type
  if (!type) errors.push('Brak Type.');
  else if (!TYPES.includes(type)) errors.push('Type "' + type + '" spoza listy: ' + TYPES.join('/'));

  // Speaker
  if (!speaker) errors.push('Brak Speaker.');
  else {
    const inCasting = casting && casting.has(speaker);
    const isPool = EFFORT_POOL_SPEAKERS.includes(speaker);
    if (type === 'Effort') {
      if (!inCasting && !isPool) errors.push('Effort-Speaker "' + speaker + '" nie jest pulą (' + EFFORT_POOL_SPEAKERS.join('/') + ') ani wpisem w Casting.');
    } else if (casting && !inCasting) {
      errors.push('Speaker "' + speaker + '" nie istnieje w arkuszu Casting.');
    }
  }

  // Context — klucz Aneks B
  if (!context) errors.push('Brak Context.');
  else {
    if (!/^[A-Za-z0-9]+$/.test(context)) errors.push('Context "' + context + '" — tylko litery/cyfry, bez podkreślników (Aneks B).');
    const startsQ = /^Q\d{2}/.test(context);

    if (type === 'Bark' || type === 'Effort') {
      if (vocab && vocab.intents.size > 0 && !vocab.intents.has(context))
        errors.push('Intent "' + context + '" spoza Vocabulary→Intent.');
    }
    if (type === 'Story' && !startsQ) errors.push('Story-Context musi zaczynać się od Q## (np. Q02Intro).');
    if (type === 'ConvH' && !startsQ) warnings.push('ConvH bez kotwicy Q## — świadomy wyjątek world-building? (preload strefą, nie questem)');
    if (type === 'ConvB') {
      if (startsQ) errors.push('ConvB nie może mieć kotwicy Q## — treść questowa to ConvH (Aneks B).');
      const base = context.replace(/\d+$/, '');
      if (vocab && vocab.topics.size > 0 && !vocab.topics.has(base))
        errors.push('Topic "' + base + '" spoza Vocabulary→Topic.');
    }
    if ((type === 'Story' || type === 'ConvH') && startsQ && vocab && vocab.quests.size > 0) {
      const q = context.substring(0, 3);
      if (!vocab.quests.has(q)) warnings.push('Quest ' + q + ' nie figuruje w Vocabulary→Quest.');
    }
  }

  // Sequence_Order
  if (SEQUENCE_TYPES.includes(type)) {
    if (!Number.isInteger(row.seq) || row.seq < 1) errors.push('Sekwencja wymaga Sequence_Order ≥ 1 (kroki co 10).');
    else if (row.seq > 999) errors.push('Sequence_Order > 999 nie mieści się w NN.');
  }

  // Treść
  if (!row.hasText && !row.hasS2S) errors.push('Wiersz bez treści (Spoken_Text/S2S_Reference).');

  return { errors: errors, warnings: warnings };
}

// ------------------------------------------------------------
// CZYTANIE WIERSZY ARKUSZA
// ------------------------------------------------------------
function readRows_(sheet) {
  const map = getHeaderMap_(sheet);
  const required = ['Line_ID', 'Speaker', 'Type', 'Context'];
  const missing = required.filter(c => !(c in map));
  if (missing.length) return { error: 'Brak kolumn: ' + missing.join(', '), rows: [], map: map };

  const n = sheet.getLastRow() - 1;
  if (n < 1) return { error: null, rows: [], map: map };
  const data = sheet.getRange(2, 1, n, sheet.getLastColumn()).getValues();

  const get = (r, key) => (key in map) ? String(r[map[key]] === null ? '' : r[map[key]]).trim() : '';
  const rows = data.map((r, i) => {
    const seqRaw = get(r, 'Sequence_Order');
    return {
      rowIndex: i + 2,
      id: get(r, 'Line_ID'),
      speaker: get(r, 'Speaker'),
      type: get(r, 'Type'),
      context: get(r, 'Context'),
      seq: seqRaw === '' ? null : Math.trunc(Number(seqRaw)),
      hasText: get(r, 'Spoken_Text') !== '',
      hasS2S: get(r, 'S2S_Reference') !== '',
      isEmpty: !get(r, 'Speaker') && !get(r, 'Type') && !get(r, 'Context') &&
               get(r, 'Spoken_Text') === '' && get(r, 'S2S_Reference') === '' && !get(r, 'Line_ID'),
    };
  });
  return { error: null, rows: rows, map: map };
}

// ------------------------------------------------------------
// GŁÓWNA: GENEROWANIE ID (aktywny arkusz)
// ------------------------------------------------------------
function generateIdsActiveSheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (META_SHEETS.includes(sheet.getName())) { ui.alert('To arkusz meta (' + sheet.getName() + ') — tu nie ma linii.'); return; }

  const casting = loadCastingSpeakers_(ss);
  const vocab = loadVocabulary_(ss);
  if (!casting) { ui.alert('BŁĄD: brak arkusza "Casting" z kolumną Speaker. Utwórz go najpierw (Faza 0a).'); return; }
  if (!vocab)   { ui.alert('BŁĄD: brak arkusza "Vocabulary" (nagłówki: Intent | Quest | Topic).'); return; }

  const parsed = readRows_(sheet);
  if (parsed.error) { ui.alert('BŁĄD: ' + parsed.error); return; }

  const existing = collectExisting_(ss);
  const idCol = parsed.map['Line_ID'] + 1;

  // dedup Sequence_Order w obrębie rozmowy (Type_Context) — po całym arkuszu
  const seqSeen = {}; // "Type_Context" -> Set(seq)
  parsed.rows.forEach(row => {
    if (SEQUENCE_TYPES.includes(row.type) && Number.isInteger(row.seq)) {
      const key = row.type + '_' + row.context;
      if (!seqSeen[key]) seqSeen[key] = {};
      seqSeen[key][row.seq] = (seqSeen[key][row.seq] || 0) + 1;
    }
  });

  let created = 0, skippedHasId = 0, failed = 0;
  const problems = [];

  parsed.rows.forEach(row => {
    const cell = sheet.getRange(row.rowIndex, idCol);
    if (row.isEmpty) return;
    if (row.id) { skippedHasId++; cell.setNote(null); return; } // ID nieśmiertelne — nie ruszamy

    const v = validateRow_(row, casting, vocab);

    // dubel Sequence_Order w rozmowie?
    if (SEQUENCE_TYPES.includes(row.type) && Number.isInteger(row.seq)) {
      const key = row.type + '_' + row.context;
      if (seqSeen[key] && seqSeen[key][row.seq] > 1)
        v.errors.push('Duplikat Sequence_Order=' + row.seq + ' w rozmowie ' + key + '.');
    }

    if (v.errors.length) {
      failed++;
      cell.setNote('⛔ ' + v.errors.concat(v.warnings.map(w => '⚠️ ' + w)).join('\n'));
      problems.push('w.' + row.rowIndex + ': ' + v.errors[0]);
      return;
    }

    // NN
    let nn;
    if (SEQUENCE_TYPES.includes(row.type)) {
      nn = ('000' + row.seq).slice(-3);
    } else {
      const group = row.speaker + '_' + row.type + '_' + row.context;
      const next = (existing.groupMaxNN[group] || 0) + 1;
      if (next > 99) {
        failed++; cell.setNote('⛔ Grupa ' + group + ' przekroczyła 99 wariantów.');
        problems.push('w.' + row.rowIndex + ': limit wariantów'); return;
      }
      existing.groupMaxNN[group] = next;
      nn = ('00' + next).slice(-2);
    }

    const newId = row.speaker + '_' + row.type + '_' + row.context + '_' + nn;
    if (existing.ids.has(newId)) {
      failed++; cell.setNote('⛔ Duplikat ID: ' + newId + ' już istnieje w pliku.');
      problems.push('w.' + row.rowIndex + ': duplikat ' + newId); return;
    }

    existing.ids.add(newId);
    cell.setValue(newId);
    cell.setNote(v.warnings.length ? '⚠️ ' + v.warnings.join('\n') : null);
    created++;
  });

  let msg = 'Arkusz "' + sheet.getName() + '":\n✅ nadano ' + created + ' ID\n⏭️ pominięto (ID już było): ' + skippedHasId + '\n⛔ odrzucono: ' + failed;
  if (problems.length) msg += '\n\nProblemy (notatki na komórkach Line_ID):\n' + problems.slice(0, 12).join('\n') + (problems.length > 12 ? '\n…' : '');
  ui.alert(msg);
}

// ------------------------------------------------------------
// AUTONUMERACJA zaznaczonego bloku (Sequence_Order co 10)
// ------------------------------------------------------------
function autoNumberSelection() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  if (META_SHEETS.includes(sheet.getName())) { ui.alert('Arkusz meta — nie numerujemy.'); return; }

  const map = getHeaderMap_(sheet);
  if (!('Sequence_Order' in map)) { ui.alert('BŁĄD: brak kolumny Sequence_Order.'); return; }
  const sel = sheet.getActiveRange();
  if (!sel) { ui.alert('Zaznacz wiersze segmentu do ponumerowania.'); return; }

  const startRow = Math.max(sel.getRow(), 2);
  const endRow = sel.getLastRow();
  if (endRow < startRow) { ui.alert('Zaznaczenie nie obejmuje wierszy danych.'); return; }

  const seqCol = map['Sequence_Order'] + 1;
  const range = sheet.getRange(startRow, seqCol, endRow - startRow + 1, 1);
  const current = range.getValues();
  const hasValues = current.some(r => String(r[0]).trim() !== '' && Number(r[0]) !== 0);
  if (hasValues) {
    const resp = ui.alert('Nadpisać istniejące Sequence_Order w zaznaczeniu?\n(ID już nadane NIE zmienią się — kolejność gra kolumna, §6.1)', ui.ButtonSet.YES_NO);
    if (resp !== ui.Button.YES) return;
  }

  let v = 10;
  const out = current.map(() => { const r = [v]; v += 10; return r; });
  range.setValues(out);
  ui.alert('Ponumerowano ' + out.length + ' wierszy: 10…' + (v - 10) + '.');
}

// ------------------------------------------------------------
// WALIDACJA WSZYSTKICH ARKUSZY (dry-run, nic nie zapisuje)
// ------------------------------------------------------------
function validateAllSheets() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const casting = loadCastingSpeakers_(ss);
  const vocab = loadVocabulary_(ss);
  const report = [];
  if (!casting) report.push('⛔ Brak arkusza Casting (kolumna Speaker).');
  if (!vocab) report.push('⛔ Brak arkusza Vocabulary (Intent|Quest|Topic).');

  const allIds = {};
  let totalErr = 0, totalWarn = 0;

  ss.getSheets().forEach(sheet => {
    if (META_SHEETS.includes(sheet.getName())) return;
    const parsed = readRows_(sheet);
    if (parsed.error) { report.push('⛔ [' + sheet.getName() + '] ' + parsed.error); return; }

    parsed.rows.forEach(row => {
      if (row.isEmpty) return;
      const v = validateRow_(row, casting, vocab);
      if (row.id) {
        if (allIds[row.id]) v.errors.push('Duplikat Line_ID (też w ' + allIds[row.id] + ').');
        else allIds[row.id] = sheet.getName() + ' w.' + row.rowIndex;
      }
      v.errors.forEach(e => { totalErr++; if (report.length < 40) report.push('⛔ [' + sheet.getName() + ' w.' + row.rowIndex + '] ' + e); });
      v.warnings.forEach(w => { totalWarn++; if (report.length < 40) report.push('⚠️ [' + sheet.getName() + ' w.' + row.rowIndex + '] ' + w); });
    });
  });

  const head = 'WALIDACJA: ' + (totalErr === 0 ? '✅ zero błędów' : '⛔ błędów: ' + totalErr) + ', ⚠️ ostrzeżeń: ' + totalWarn + '\n\n';
  ui.alert(head + (report.length ? report.join('\n') : 'Wszystko czyste. Można generować.'));
}
