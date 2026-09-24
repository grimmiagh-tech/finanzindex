/**
 * Finanzindex – Google Apps Script Backend
 * Passend zu den Blättern: Konten | Kategorien | Fixkosten_Plan | Transaktionen
 *
 * EINRICHTUNG:
 * 1. Im Google Sheet: Erweiterungen -> Apps Script. Diesen kompletten Code einfügen (alten ersetzen).
 * 2. Unten in KONFIGURATION deine E-Mail-Adresse eintragen (für Fehlerbenachrichtigungen).
 * 3. Projekteinstellungen (Zahnrad links) -> Haken bei "'appsscript.json'-Manifestdatei im Editor anzeigen"
 *    setzen -> appsscript.json öffnen -> "timeZone" auf "Europe/Berlin" stellen -> Speichern.
 * 4. Im Funktions-Dropdown oben "installTrigger" auswählen -> Ausführen (einmalig!). Beim ersten Mal
 *    erscheint ein Berechtigungsfenster -> Konto wählen -> "Erweitert" -> "Zu Finanzindex (unsicher)
 *    wechseln" -> Zulassen. Das richtet die automatische monatliche Ausführung ein. Da diese Version
 *    zusätzlich automatische Backups über Google Drive anlegt, wird bei der Berechtigung jetzt auch
 *    Zugriff auf Google Drive abgefragt - das ist normal und nötig, bitte zulassen.
 * 5. Bereitstellen -> Neue Bereitstellung -> Web-App -> Ausführen als: Ich, Zugriff: Jeder.
 * 6. Die erzeugte URL in der Finanzindex-App unter "Einstellungen" eintragen.
 * 7. Den unten stehenden API_SECRET auch in der App unter "Einstellungen" beim Feld
 *    "API-Schlüssel" eintragen - ohne diesen zweiten, geheimen Wert antwortet das Skript
 *    ab jetzt auf keine Anfrage mehr, selbst wenn die URL bekannt sein sollte.
 *
 * Im Google Sheet im Blatt "Konten" bitte einmalig eine neue Zeile ergänzen:
 *   Konto: Ausgaberest | Aktueller Saldo: 0
 */

// ============================== KONFIGURATION ==============================
const NOTIFY_EMAIL = 'DEINE-EMAIL@beispiel.de'; // <- HIER deine Adresse eintragen
// Zusätzliches, geheimes Kennwort (zufällig erzeugt, 43 Zeichen) - muss exakt so
// auch in der App unter "Einstellungen" -> "API-Schlüssel" eingetragen werden.
// Ohne passenden Schlüssel lehnt das Skript jede Anfrage ab, selbst bei korrekter URL.
const API_SECRET = 'C9qoSKYirav7DZni8DM7yNQKqUe3SwJBVvkhsK7MJMI';
// =============================================================================

const SHEETS = {
  TRANSAKTIONEN: 'Transaktionen',
  KONTEN: 'Konten',
  KATEGORIEN: 'Kategorien',
  FIXKOSTEN: 'Fixkosten_Plan',
  EINSTELLUNGEN: 'Einstellungen',
};

const REAL_ACCOUNTS = ['Trade Republic', 'ING'];
// Diese Töpfe werden beim Monatsabschluss NICHT in den Ausgaberest gefegt.
const POTS_OHNE_AUSGABEREST = ['Model 3', 'Rücklage', 'Jahres Rückstellung', 'Transit', 'Ausgaberest'];

const MONATSNAMEN = ['Januar','Februar','März','April','Mai','Juni','Juli','August','September','Oktober','November','Dezember'];

/* ============================== Web-App-Endpunkte ============================== */

function doGet(e) {
  try {
    if (!e || !e.parameter || e.parameter.key !== API_SECRET) {
      return jsonResponse({ status: 'error', error: 'Nicht autorisiert - falscher oder fehlender API-Schlüssel.' });
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const geraeteStatus = getGeraeteStatus(ss, e.parameter.geraeteId);

    const payload = {
      transaktionen: sheetToObjects(ss.getSheetByName(SHEETS.TRANSAKTIONEN)),
      konten: sheetToObjects(ss.getSheetByName(SHEETS.KONTEN)),
      kategorien: sheetToObjects(ss.getSheetByName(SHEETS.KATEGORIEN)),
      fixkosten: sheetToObjects(ss.getSheetByName(SHEETS.FIXKOSTEN)),
      einstellungen: sheetToObjects(ss.getSheetByName(SHEETS.EINSTELLUNGEN)),
      kontoauszuege: listeKontoauszuege(ss),
      geraete: geraeteStatus.liste,
      gesperrt: geraeteStatus.gesperrt,
    };
    return jsonResponse({ status: 'ok', data: payload });
  } catch (err) {
    return jsonResponse({ status: 'error', error: String(err.message || err) });
  }
}

/* ============================== Geräte-Online-Kontrolle ============================== */
// Sichtbares Tabellenblatt "Geräte" mit Geräte-ID + Zeitpunkt des letzten
// Lebenszeichens. Ein Gerät gilt als "online", wenn sein letztes
// Lebenszeichen höchstens GERAETE_TIMEOUT_MS alt ist - so bleibt die Zahl
// auch dann korrekt, wenn eine App mal nicht sauber beendet wurde (Akku
// leer, App aus der Android-Übersicht weggewischt o. Ä.), statt für immer
// als "online" hängen zu bleiben.
const GERAETE_SHEET = 'Geräte';
const GERAETE_TIMEOUT_MS = 55 * 1000;
const GERAETE_SPALTEN = ['Geräte-ID', 'Name', 'Zuletzt gesehen', 'Gesperrt'];

// Dauerhafte, ausschließlich manuell gepflegte Liste vertrauenswürdiger
// Geräte-IDs (eigene Geräte) - bewusst ein KOMPLETT SEPARATES Blatt,
// getrennt vom "Geräte"-Blatt. Der erste Anlauf hatte "Vertraut" als
// Spalte direkt in der Geräte-Zeile gespeichert - das Problem dabei: genau
// diese Zeile wird ja routinemäßig gelöscht, sobald das Gerät länger
// inaktiv war, und damit ging die Vertraut-Markierung immer wieder
// verloren. Diese Liste hier wird vom Skript NIE automatisch verändert,
// ergänzt oder geleert - nur direkt im Sheet von Hand pflegen.
const VERTRAUTE_GERAETE_SHEET = 'Vertraute Geräte';
const VERTRAUTE_GERAETE_SPALTEN = ['Geräte-ID', 'Notiz (z. B. Gerätename)'];

function getOrCreateVertrauteGeraeteSheet(ss) {
  let sheet = ss.getSheetByName(VERTRAUTE_GERAETE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(VERTRAUTE_GERAETE_SHEET);
    sheet.getRange(1, 1, 1, VERTRAUTE_GERAETE_SPALTEN.length).setValues([VERTRAUTE_GERAETE_SPALTEN]);
  }
  return sheet;
}

function holeVertrauteIds(ss) {
  const sheet = getOrCreateVertrauteGeraeteSheet(ss);
  const values = sheet.getDataRange().getValues();
  const ids = new Set();
  for (let r = 1; r < values.length; r++) {
    if (values[r][0]) ids.add(String(values[r][0]).trim());
  }
  return ids;
}

// Legt das Blatt bei Bedarf an UND repariert eine evtl. fehlende oder
// unvollständige Kopfzeile - ohne dabei bestehende Gerätezeilen zu
// überschreiben. Das war nötig, weil ältere/manuell angelegte Blätter
// teils ganz ohne Kopfzeile begannen.
function getOrCreateGeraeteSheet(ss) {
  let sheet = ss.getSheetByName(GERAETE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(GERAETE_SHEET);
    sheet.getRange(1, 1, 1, GERAETE_SPALTEN.length).setValues([GERAETE_SPALTEN]);
    return sheet;
  }
  const ersteZeile = sheet.getRange(1, 1, 1, GERAETE_SPALTEN.length).getValues()[0];
  if (ersteZeile[0] !== GERAETE_SPALTEN[0]) {
    // Kopfzeile fehlt komplett (Zeile 1 enthält bereits ein echtes Gerät) -
    // neue Zeile oben einfügen, bestehende Daten rutschen nur eine Zeile
    // nach unten, nichts wird überschrieben.
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, GERAETE_SPALTEN.length).setValues([GERAETE_SPALTEN]);
  } else {
    // Kopfzeile (Spalte A) vorhanden, aber evtl. noch mit falschen/alten
    // Beschriftungen aus einer früheren, kleineren Sheet-Version in den
    // übrigen Spalten. Betrifft NUR die Beschriftungen in Zeile 1, nie die
    // eigentlichen Gerätezeilen darunter.
    let passtNicht = false;
    for (let i = 1; i < GERAETE_SPALTEN.length; i++) {
      if (ersteZeile[i] !== GERAETE_SPALTEN[i]) { passtNicht = true; break; }
    }
    if (passtNicht) sheet.getRange(1, 1, 1, GERAETE_SPALTEN.length).setValues([GERAETE_SPALTEN]);
  }
  // Alte "Vertraut"-Spalte (Spalte E) aus dem ersten, verworfenen Ansatz
  // entfernen, falls noch vorhanden - Vertrauen läuft jetzt ausschließlich
  // über das separate Blatt "Vertraute Geräte" oben.
  if (sheet.getLastColumn() > GERAETE_SPALTEN.length) {
    sheet.getRange(1, GERAETE_SPALTEN.length + 1, sheet.getMaxRows(), sheet.getLastColumn() - GERAETE_SPALTEN.length).clearContent();
  }
  return sheet;
}

// Akzeptiert sowohl einen echten Boolean TRUE als auch den Text "TRUE"/
// "true" (falls eine Zelle z. B. wegen des Formats "Nur Text" nicht
// automatisch in einen Wahrheitswert umgewandelt wurde) - damit ein von
// Hand eingetragenes TRUE in "Gesperrt"/"Vertraut" in jedem Fall erkannt
// wird, unabhängig vom genauen Zellformat.
function istWahr(wert) {
  if (wert === true) return true;
  if (typeof wert === 'string' && wert.trim().toLowerCase() === 'true') return true;
  return false;
}

// WICHTIG: Läuft unter einer LockService-Sperre (siehe unten). Ohne diese
// Sperre können zwei fast gleichzeitige Heartbeats (z. B. mehrere Geräte,
// oder ein doppelt ausgelöster Heartbeat desselben Geräts) die Tabelle
// BEIDE lesen, bevor die erste Schreibaktion angekommen ist - beide finden
// dann "keinen Treffer" und hängen je eine eigene Zeile an. Genau das war
// die Ursache für doppelt gelistete Geräte. Die Sperre betrifft nur diesen
// kleinen Lese-und-Schreib-Vorgang auf dem "Geräte"-Blatt, nicht den
// großen Datensync in doGet - dort ändert sich nichts an der Performance.
//
// Räumt bei jedem Aufruf nebenbei auf: Ein Gerät, das in Spalte "Vertraut"
// mit TRUE markiert ist (eigene, bekannte Geräte) und dessen letztes
// Lebenszeichen den Timeout überschritten hat, wird komplett aus dem Blatt
// entfernt - genau wie beim PC, der sich beim Schließen sofort selbst
// abmeldet, nur eben zeitversetzt um den Timeout, weil uns bei
// Android-Geräten (App aus der Übersicht weggewischt) kein zuverlässiges
// "sofort"-Signal zur Verfügung steht. NICHT als vertraut markierte
// (unbekannte/fremde) Geräte werden NIE automatisch entfernt, damit die
// Sicherheitsfunktion (fremde Geräte erkennen) erhalten bleibt. Gesperrte
// Geräte werden ebenfalls nie automatisch entfernt, sonst würde die Sperre
// nach Ablauf des Timeouts wirkungslos.
function upsertGeraet(ss, geraeteId, geraeteName) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getOrCreateGeraeteSheet(ss);
    const vertrauteIds = holeVertrauteIds(ss);
    const values = sheet.getDataRange().getValues();
    const now = new Date();
    const nowMs = now.getTime();

    let gesperrt = false;
    let ownName = geraeteName || '';
    const kept = [];

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      if (!row[0]) continue;

      if (row[0] === geraeteId) {
        // Eigene Zeile(n) zusammenführen (auch falls durch eine frühere
        // Race Condition mehrfach vorhanden) - Name/Sperre übernehmen, die
        // Zeile selbst wird unten einmalig neu geschrieben.
        if (!geraeteName && row[1]) ownName = row[1];
        if (istWahr(row[3])) gesperrt = true;
        continue;
      }

      const istGesperrt = istWahr(row[3]);
      const istVertraut = vertrauteIds.has(String(row[0]).trim());
      const ts = row[2];
      const zuletzt = ts instanceof Date ? ts.getTime() : null;
      if (istVertraut && !istGesperrt && zuletzt !== null && (nowMs - zuletzt) > GERAETE_TIMEOUT_MS) {
        continue; // veraltet und vertraut (laut separater Liste) -> löschen
      }
      // Beim Übernehmen gleich auf einen echten Boolean normalisieren -
      // falls Gesperrt bisher als Text "TRUE"/"FALSE" statt als
      // Wahrheitswert gespeichert war, heilt sich das dauerhaft aus.
      kept.push([row[0], row[1], row[2], istGesperrt]);
    }

    kept.push([geraeteId, ownName, now, gesperrt]);

    sheet.getRange(2, 1, kept.length, GERAETE_SPALTEN.length).setValues(kept);
    const alteDatenzeilen = values.length - 1;
    if (alteDatenzeilen > kept.length) {
      sheet.deleteRows(2 + kept.length, alteDatenzeilen - kept.length);
    }

    return gesperrt;
  } finally {
    lock.releaseLock();
  }
}

// Sperrt ein Gerät dauerhaft (Häkchen in Spalte "Gesperrt"), statt die Zeile
// zu löschen - das gesperrte Gerät bekommt das bei seiner nächsten Anfrage
// mitgeteilt und zeigt dann einen Sperrbildschirm, unabhängig davon, ob auf
// diesem Gerät eine PIN eingerichtet ist. WICHTIG: Das Aufheben der Sperre
// geht absichtlich NICHT aus der App heraus - nur direkt im Sheet das
// Häkchen in Spalte D entfernen. Ein fremdes Gerät kennt zwar ggf. den
// API-Schlüssel, hat aber keinen Zugriff auf das Google-Konto selbst.
function sperreGeraet(ss, geraeteId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getOrCreateGeraeteSheet(ss);
    const values = sheet.getDataRange().getValues();
    let gefunden = false;
    for (let r = 1; r < values.length; r++) {
      if (values[r][0] === geraeteId) {
        sheet.getRange(r + 1, 4).setValue(true);
        gefunden = true;
      }
    }
    if (!gefunden) sheet.appendRow([geraeteId, '', new Date(), true]);
  } finally {
    lock.releaseLock();
  }
}

function entferneGeraet(ss, geraeteId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = ss.getSheetByName(GERAETE_SHEET);
    if (!sheet) return;
    const values = sheet.getDataRange().getValues();
    for (let r = values.length - 1; r >= 1; r--) {
      if (values[r][0] === geraeteId) sheet.deleteRow(r + 1);
    }
  } finally {
    lock.releaseLock();
  }
}

// Räumt veraltete, vertraute und NICHT gesperrte Geräte auf - unabhängig
// davon, ob gerade überhaupt ein Gerät aktiv ist. upsertGeraet räumt zwar
// bei jedem Heartbeat schon "nebenbei" die JEWEILS ANDEREN Geräte auf,
// aber wenn z. B. das Handy das letzte offene Gerät war, schickt danach
// niemand mehr ein Heartbeat - die eigene, jetzt veraltete Zeile des Handys
// würde sonst für immer stehen bleiben. Dieser Trigger (siehe
// installGeraeteAufraeumTrigger, einmalig einrichten) läuft deshalb jede
// Minute unabhängig für sich und schließt genau diese Lücke.
function raeumeVeralteteGeraeteAuf() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateGeraeteSheet(ss);
    const vertrauteIds = holeVertrauteIds(ss);
    const values = sheet.getDataRange().getValues();
    const nowMs = Date.now();
    const kept = [];
    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      if (!row[0]) continue;
      const istGesperrt = istWahr(row[3]);
      const istVertraut = vertrauteIds.has(String(row[0]).trim());
      const ts = row[2];
      const zuletzt = ts instanceof Date ? ts.getTime() : null;
      if (istVertraut && !istGesperrt && zuletzt !== null && (nowMs - zuletzt) > GERAETE_TIMEOUT_MS) {
        continue; // veraltet und vertraut (laut separater Liste) -> löschen
      }
      kept.push([row[0], row[1], row[2], istGesperrt]);
    }
    const alteDatenzeilen = values.length - 1;
    if (kept.length !== alteDatenzeilen) {
      if (kept.length > 0) {
        sheet.getRange(2, 1, kept.length, GERAETE_SPALTEN.length).setValues(kept);
      }
      sheet.deleteRows(2 + kept.length, alteDatenzeilen - kept.length);
    }
  } finally {
    lock.releaseLock();
  }
}

// EINMALIG im Funktions-Dropdown oben auswählen und ausführen (wie bei
// installTrigger für die monatliche Automatik). Richtet einen
// zeitgesteuerten Trigger ein, der raeumeVeralteteGeraeteAuf jede Minute
// aufruft. Entfernt zuerst einen evtl. bereits vorhandenen gleichen
// Trigger, damit bei mehrfachem Ausführen kein Duplikat entsteht.
function installGeraeteAufraeumTrigger() {
  ScriptApp.getProjectTriggers().forEach((t) => {
    if (t.getHandlerFunction() === 'raeumeVeralteteGeraeteAuf') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('raeumeVeralteteGeraeteAuf').timeBased().everyMinutes(1).create();
}

// Liefert Geräteliste UND Sperr-Status des anfragenden Geräts in EINEM
// einzigen Lesevorgang der Tabelle (statt zwei getrennten) - wichtig, da
// Google die Tabelle bei jedem Lesevorgang kurz sperrt und zu viele separate
// Zugriffe bei mehreren gleichzeitig aktiven Geräten sonst zu Wartezeiten
// bis hin zu Zeitüberschreitungen führen können.
function getGeraeteStatus(ss, anfragendesGeraet) {
  const sheet = ss.getSheetByName(GERAETE_SHEET);
  if (!sheet) return { liste: [], gesperrt: false };
  const values = sheet.getDataRange().getValues();
  const now = Date.now();
  // Nach Geräte-ID zusammenführen (defensiv gegen evtl. noch vorhandene
  // Duplikat-Zeilen aus der Zeit vor der LockService-Absicherung): neuester
  // Zeitstempel gewinnt, eine Sperre in irgendeiner Zeile gilt fürs Gerät.
  const byId = new Map();
  let gesperrt = false;
  for (let r = 1; r < values.length; r++) {
    const id = values[r][0];
    if (!id) continue;
    const ts = values[r][2];
    const zuletzt = ts instanceof Date ? ts.getTime() : null;
    const istGesperrt = istWahr(values[r][3]);
    const bestehend = byId.get(id);
    if (!bestehend || (zuletzt || 0) > (bestehend.zuletztGesehen || 0)) {
      byId.set(id, {
        id,
        name: values[r][1] || (bestehend ? bestehend.name : ''),
        zuletztGesehen: zuletzt,
        online: zuletzt !== null && (now - zuletzt) < GERAETE_TIMEOUT_MS,
        gesperrt: istGesperrt || (bestehend ? bestehend.gesperrt : false),
      });
    } else if (istGesperrt) {
      bestehend.gesperrt = true;
    }
    if (anfragendesGeraet && id === anfragendesGeraet && istGesperrt) gesperrt = true;
  }
  const liste = Array.from(byId.values());
  liste.sort((a, b) => (b.zuletztGesehen || 0) - (a.zuletztGesehen || 0));
  return { liste, gesperrt };
}

const KONTOAUSZUEGE_CACHE_KEY = 'kontoauszuege_liste_v1';
const KONTOAUSZUEGE_CACHE_SEKUNDEN = 900; // 15 Minuten

// Listet alle bisher erstellten Kontoauszug-PDFs aus Drive auf (Name + Link),
// damit die App sie direkt anzeigen kann, ohne selbst etwas zu speichern -
// die Datei liegt ja bereits einmal in Drive.
//
// Performance: Drive-Zugriffe sind spürbar langsamer als reine Sheet-
// Zugriffe, diese Liste ändert sich aber in der Praxis nur einmal im Monat
// (automatischer Monatsabschluss). Deshalb 15 Minuten aus dem Cache
// bedienen, statt bei JEDEM Sync erneut den Drive-Ordner zu durchsuchen.
// kontoauszugErstellen leert den Cache zusätzlich sofort nach dem
// Erstellen eines neuen Auszugs, damit ein neuer Auszug trotzdem sofort
// sichtbar ist, statt bis zu 15 Minuten zu warten.
function listeKontoauszuege(ss) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(KONTOAUSZUEGE_CACHE_KEY);
  if (cached !== null) {
    try { return JSON.parse(cached); } catch (parseErr) { /* Cache-Inhalt ignorieren, neu laden */ }
  }
  try {
    const sheetFile = DriveApp.getFileById(ss.getId());
    const parentFolder = sheetFile.getParents().hasNext() ? sheetFile.getParents().next() : DriveApp.getRootFolder();
    const ordnerIter = parentFolder.getFoldersByName('Finanzindex Kontoauszüge');
    if (!ordnerIter.hasNext()) {
      cache.put(KONTOAUSZUEGE_CACHE_KEY, JSON.stringify([]), KONTOAUSZUEGE_CACHE_SEKUNDEN);
      return [];
    }
    const ordner = ordnerIter.next();
    const files = ordner.getFiles();
    const liste = [];
    while (files.hasNext()) {
      const f = files.next();
      liste.push({ name: f.getName(), url: f.getUrl(), erstellt: f.getDateCreated().getTime() });
    }
    liste.sort((a, b) => b.erstellt - a.erstellt);
    cache.put(KONTOAUSZUEGE_CACHE_KEY, JSON.stringify(liste), KONTOAUSZUEGE_CACHE_SEKUNDEN);
    return liste;
  } catch (err) {
    return [];
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.key !== API_SECRET) {
      return jsonResponse({ status: 'error', error: 'Nicht autorisiert - falscher oder fehlender API-Schlüssel.' });
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (body.action === 'create' || body.action === 'update' || body.action === 'delete') {
      const sheet = ss.getSheetByName(SHEETS.TRANSAKTIONEN);
      if (body.action === 'create') {
        insertTransaktionRow(sheet, body);
      } else if (body.action === 'update') {
        const row = findRowById(sheet, body.id);
        if (row === -1) throw new Error('Buchung nicht gefunden: ' + body.id);
        sheet.getRange(row, 1, 1, 9).setValues([[
          body.datum, body.titel, body.kategorie, body.typ,
          body.betrag, body.quelle, body.ziel || '', body.notiz || '', body.id,
        ]]);
      } else {
        const row = findRowById(sheet, body.id);
        if (row === -1) throw new Error('Buchung nicht gefunden: ' + body.id);
        sheet.deleteRow(row);
      }
    } else if (body.action === 'fixkosten-create' || body.action === 'fixkosten-update' || body.action === 'fixkosten-delete') {
      const sheet = ss.getSheetByName(SHEETS.FIXKOSTEN);
      if (body.action === 'fixkosten-create') {
        insertFixkostenRow(sheet, body);
      } else if (body.action === 'fixkosten-update') {
        const row = findFixkostenRowById(sheet, body.id);
        if (row === -1) throw new Error('Fixkosten-Position nicht gefunden: ' + body.id);
        sheet.getRange(row, 1, 1, 9).setValues([[
          body.bezeichnung, body.betrag, body.kategorie, body.typ,
          body.quelle, body.ziel || '', body.gueltigAb || '', body.gueltigBis || '', body.id,
        ]]);
      } else {
        const row = findFixkostenRowById(sheet, body.id);
        if (row === -1) throw new Error('Fixkosten-Position nicht gefunden: ' + body.id);
        sheet.deleteRow(row);
      }
    } else if (body.action === 'heartbeat') {
      const gesperrt = upsertGeraet(ss, body.geraeteId, body.geraeteName);
      // Aktuelle Geräteliste direkt in der Heartbeat-Antwort mitschicken
      // (Performance-Optimierung): dadurch braucht das Frontend keine
      // zusätzliche, separate Abfrage mehr nur für die "Geräte online"-
      // Anzeige - spart bei mehreren Geräten spürbar Anfragen gegenüber
      // Apps Script/Sheets.
      const geraeteStatus = getGeraeteStatus(ss, body.geraeteId);
      return jsonResponse({ status: 'ok', gesperrt: gesperrt, geraete: geraeteStatus.liste });
    } else if (body.action === 'geraet-abmelden') {
      entferneGeraet(ss, body.geraeteId);
    } else if (body.action === 'geraet-sperren') {
      sperreGeraet(ss, body.geraeteId);
    } else if (body.action === 'ai-chat') {
      const ergebnis = handleAiChat(ss, body);
      return jsonResponse({ status: 'ok', antwort: ergebnis.text, buchungsvorschlag: ergebnis.buchungsvorschlag, navigation: ergebnis.navigation });
    } else {
      throw new Error('Unbekannte Aktion: ' + body.action);
    }

    return jsonResponse({ status: 'ok' });
  } catch (err) {
    return jsonResponse({ status: 'error', error: String(err.message || err) });
  }
}

function insertTransaktionRow(sheet, tx) {
  sheet.insertRowBefore(2);
  sheet.getRange(2, 1, 1, 9).setValues([[
    tx.datum, tx.titel, tx.kategorie, tx.typ,
    tx.betrag, tx.quelle, tx.ziel || '', tx.notiz || '', tx.id,
  ]]);
}

// Reihenfolge muss zur Spaltenreihenfolge in Fixkosten_Plan passen:
// Bezeichnung | Betrag | Kategorie | Typ | Konto | Zielkonto | Gültig Ab | Gültig bis | ID
function insertFixkostenRow(sheet, fk) {
  sheet.insertRowBefore(2);
  sheet.getRange(2, 1, 1, 9).setValues([[
    fk.bezeichnung, fk.betrag, fk.kategorie, fk.typ,
    fk.quelle, fk.ziel || '', fk.gueltigAb || '', fk.gueltigBis || '', fk.id,
  ]]);
}

function findFixkostenRowById(sheet, id) {
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf('ID');
  if (idCol === -1) return -1;
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idCol]) === String(id)) return r + 1;
  }
  return -1;
}

function sheetToObjects(sheet) {
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (values.length === 0) return [];
  const tz = Session.getScriptTimeZone() || 'Europe/Berlin';
  const headers = values.shift();
  return values
    .filter((row) => row.some((cell) => cell !== '' && cell !== null))
    .map((row) => {
      const obj = {};
      headers.forEach((h, i) => {
        const cell = row[i];
        // WICHTIG: Datums-Zellen als eindeutigen "JJJJ-MM-TT"-Text statt als
        // rohes Date-Objekt zurückgeben. Ein Date-Objekt würde beim
        // Verpacken in JSON automatisch in Weltzeit (UTC) umgerechnet - das
        // verschiebt ein deutsches Datum (UTC+1/+2) um einen Tag nach vorn.
        obj[h] = cell instanceof Date ? Utilities.formatDate(cell, tz, 'yyyy-MM-dd') : cell;
      });
      return obj;
    });
}

function findRowById(sheet, id) {
  const values = sheet.getDataRange().getValues();
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][8]) === String(id)) return r + 1;
  }
  return -1;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ============================== Hilfsfunktionen: Zahlen & Daten ============================== */

function parseNumberSafe(v) {
  if (typeof v === 'number') return v;
  if (!v) return 0;
  let s = String(v).replace(/[^0-9,.\-]/g, '').trim();
  if (s.indexOf(',') !== -1) s = s.replace(/\./g, '').replace(',', '.');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function toDateSafe(v) {
  if (v instanceof Date) return v;
  if (!v) return null;
  const s = String(v);
  let m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})/);
  if (m) return new Date(parseInt(m[3], 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10));
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
  return null;
}

function monthKeyOf(dateVal) {
  const d = toDateSafe(dateVal);
  if (!d) return '';
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
}

function easterSunday(year) {
  const a = year % 19, b = Math.floor(year / 100), c = year % 100;
  const d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const month = Math.floor((h + l - 7 * m + 114) / 31);
  const day = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(year, month - 1, day);
}

function germanHolidays(year) {
  const easter = easterSunday(year);
  const add = (d, n) => { const r = new Date(d); r.setDate(r.getDate() + n); return r; };
  return [new Date(year,0,1), add(easter,-2), add(easter,1), new Date(year,4,1), add(easter,39), add(easter,50), new Date(year,9,3), new Date(year,11,25), new Date(year,11,26)];
}

function sameDate(a, b) { return a.getFullYear()===b.getFullYear() && a.getMonth()===b.getMonth() && a.getDate()===b.getDate(); }
function isBankBusinessDay(date) { const day = date.getDay(); if (day===0||day===6) return false; return !germanHolidays(date.getFullYear()).some((h) => sameDate(h, date)); }
function lastBankBusinessDayOfMonth(year, monthIndex) { const d = new Date(year, monthIndex+1, 0); while (!isBankBusinessDay(d)) d.setDate(d.getDate()-1); return d; }

/* ============================== Geschäftslogik ============================== */

function getFixkostenAktiv(rows, year, monthIdx) {
  const refEnde = new Date(year, monthIdx, 28);
  const refAnfang = new Date(year, monthIdx, 1);
  return rows.filter((f) => {
    const ab = toDateSafe(f['Gültig Ab'] || f['Gültig Ab ']);
    const bis = toDateSafe(f['Gültig bis']);
    const abOk = !ab || ab <= refEnde;
    const bisOk = !bis || bis >= refAnfang;
    return abOk && bisOk;
  });
}

function getKontenMap(ss) {
  const rows = sheetToObjects(ss.getSheetByName(SHEETS.KONTEN));
  const map = {};
  rows.forEach((r) => {
    if (!r.Konto || r.Konto === 'Gesamt') return;
    map[r.Konto] = parseNumberSafe(r['Aktueller Saldo']);
  });
  return map;
}

function computeBalancesGAS(kontenMap, transaktionen) {
  const balances = {};
  Object.keys(kontenMap).forEach((name) => (balances[name] = kontenMap[name]));
  const external = { extern: true, arbeitgeber: true };
  transaktionen.forEach((t) => {
    const betrag = parseNumberSafe(t.Betrag);
    const quelle = t.Quelle, ziel = t.Ziel;
    if (t.Typ === 'Ausgabe') {
      if (balances[quelle] !== undefined) balances[quelle] -= betrag;
    } else if (t.Typ === 'Einnahme') {
      const qKey = String(quelle || '').toLowerCase();
      if (!external[qKey] && balances[quelle] !== undefined) balances[quelle] += betrag;
      if (balances[ziel] !== undefined) balances[ziel] += betrag;
    } else if (t.Typ === 'Umbuchung') {
      if (balances[quelle] !== undefined) balances[quelle] -= betrag;
      if (balances[ziel] !== undefined) balances[ziel] += betrag;
    }
  });
  return balances;
}

/* ============================== Monatsabschluss: Ausgaberest ============================== */

function ausgaberestSweep(ss, log, heute) {
  const sheet = ss.getSheetByName(SHEETS.TRANSAKTIONEN);
  const kontenMap = getKontenMap(ss);
  const transaktionen = sheetToObjects(sheet);
  const balances = computeBalancesGAS(kontenMap, transaktionen);

  const now = heute || new Date();
  const prev = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const prevMonthKey = prev.getFullYear() + '-' + String(prev.getMonth() + 1).padStart(2, '0');
  const prevMonthLabel = MONATSNAMEN[prev.getMonth()] + ' ' + prev.getFullYear();
  const letzterTagVormonat = new Date(now.getFullYear(), now.getMonth(), 0);
  const titel = 'Ausgaberest ' + prevMonthLabel;

  const potNames = Object.keys(kontenMap).filter((name) => REAL_ACCOUNTS.indexOf(name) === -1 && POTS_OHNE_AUSGABEREST.indexOf(name) === -1);

  // Jeder betroffene Topf wird exakt auf 0 zurückgesetzt - egal ob er im
  // Plus oder im Minus steht. Positive Reste fließen zu Ausgaberest, negative
  // Reste werden AUS Ausgaberest ausgeglichen. In Summe zeigt Ausgaberest
  // dadurch automatisch den echten Netto-Wert über alle Töpfe hinweg
  // (Überzüge werden direkt mit Guthaben aus anderen Töpfen verrechnet).
  let anzahl = 0;
  potNames.forEach((pot) => {
    const rest = Math.round((balances[pot] || 0) * 100) / 100;
    if (Math.abs(rest) <= 0.005) return;
    const schonGebucht = transaktionen.some((t) =>
      t.Titel === titel && monthKeyOf(t.Datum) === prevMonthKey &&
      ((t.Quelle === pot && t.Ziel === 'Ausgaberest') || (t.Quelle === 'Ausgaberest' && t.Ziel === pot))
    );
    if (schonGebucht) return;

    const istPositiv = rest > 0;
    insertTransaktionRow(sheet, {
      datum: letzterTagVormonat,
      titel: titel,
      kategorie: 'Sparen',
      typ: 'Umbuchung',
      betrag: Math.abs(rest),
      quelle: istPositiv ? pot : 'Ausgaberest',
      ziel: istPositiv ? 'Ausgaberest' : pot,
      notiz: 'Automatischer Monatsabschluss' + (istPositiv ? '' : ' (Überzug ausgeglichen)'),
      id: Utilities.getUuid(),
    });
    anzahl++;
    log.push('Ausgaberest: ' + (istPositiv ? pot + ' -> Ausgaberest' : 'Ausgaberest -> ' + pot) + ' (' + Math.abs(rest) + ' EUR)');
  });
  return anzahl;
}

/* ============================== Monatslauf: Fixkosten & Besoldung ============================== */

function fixkostenMonatslauf(ss, log, heute) {
  const sheet = ss.getSheetByName(SHEETS.TRANSAKTIONEN);
  const fixkosten = sheetToObjects(ss.getSheetByName(SHEETS.FIXKOSTEN));
  const transaktionen = sheetToObjects(sheet);

  const now = heute || new Date();
  const year = now.getFullYear(), monthIdx = now.getMonth();
  const targetMonthKey = year + '-' + String(monthIdx + 1).padStart(2, '0');
  const salaryDate = lastBankBusinessDayOfMonth(year, monthIdx - 1);
  const fixedDate = new Date(year, monthIdx, 1);

  const active = getFixkostenAktiv(fixkosten, year, monthIdx);
  let created = 0, skipped = 0;

  active.forEach((f) => {
    const titel = f.Bezeichnung;
    const schonGebucht = transaktionen.some((t) => t.Titel === titel && monthKeyOf(t.Datum) === targetMonthKey);
    if (schonGebucht) { skipped++; return; }

    const isEinnahme = f.Typ === 'Einnahme';
    const zielIstIntern = f.Zielkonto && f.Zielkonto !== 'Extern';
    const typ = isEinnahme ? 'Einnahme' : (zielIstIntern ? 'Umbuchung' : 'Ausgabe');

    insertTransaktionRow(sheet, {
      datum: isEinnahme ? salaryDate : fixedDate,
      titel: titel,
      kategorie: f.Kategorie,
      typ: typ,
      betrag: parseNumberSafe(f.Betrag),
      quelle: f.Konto,
      ziel: (zielIstIntern || isEinnahme) ? f.Zielkonto : '',
      notiz: 'Automatischer Monatslauf' + (typ !== f.Typ ? ' (korrigiert von ' + f.Typ + ' zu ' + typ + ')' : ''),
      id: Utilities.getUuid(),
    });
    created++;
    transaktionen.push({ Titel: titel, Datum: isEinnahme ? salaryDate : fixedDate });
    log.push('Fixkosten: ' + titel + ' (' + typ + ', ' + parseNumberSafe(f.Betrag) + ' EUR)');
  });
  return { created: created, skipped: skipped };
}

/* ============================== Kombinierter Trigger (läuft am 1. jeden Monats) ============================== */

function monatswechselAutomatik() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = [];
  try {
    const ausgaberestAnzahl = ausgaberestSweep(ss, log);
    kontoauszugErstellen(ss, log);
    const fixkostenErgebnis = fixkostenMonatslauf(ss, log);
    backupErstellen(ss, log);
    Logger.log(log.join('\n'));
    // Bei Bedarf: Erfolgsmeldung ebenfalls per Mail verschicken, indem die
    // nächste Zeile einkommentiert wird.
    // notifyErfolg(ausgaberestAnzahl, fixkostenErgebnis, log);
  } catch (err) {
    notifyFehler(err, log);
    throw err;
  }
}

// Erstellt einen PDF-Kontoauszug für den abgelaufenen Monat: Anfangs- und
// Endsaldo je Konto sowie alle einzelnen Buchungen des Monats. Läuft NACH
// dem Ausgaberest-Ausgleich, damit auch diese automatischen
// Abschlussbuchungen (die noch im alten Monat datiert sind) im Auszug
// erscheinen - und VOR dem neuen Fixkosten-Lauf, damit der neue Monat noch
// nicht mit hineinzählt. Landet als PDF im Drive-Ordner
// "Finanzindex Kontoauszüge" neben dem Sheet.
function kontoauszugErstellen(ss, log, heute) {
  try {
    const now = heute || new Date();
    const prev = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    const jahr = prev.getFullYear(), monatIdx = prev.getMonth();
    const monatLabel = MONATSNAMEN[monatIdx] + ' ' + jahr;
    const monatStart = new Date(jahr, monatIdx, 1);
    const monatEnde = new Date(jahr, monatIdx + 1, 0, 23, 59, 59);
    const tz = Session.getScriptTimeZone() || 'Europe/Berlin';

    const kontenMap = getKontenMap(ss);
    const alleTx = sheetToObjects(ss.getSheetByName(SHEETS.TRANSAKTIONEN));

    const vorMonat = alleTx.filter((t) => { const d = toDateSafe(t.Datum); return d && d < monatStart; });
    const bisEndeMonat = alleTx.filter((t) => { const d = toDateSafe(t.Datum); return d && d <= monatEnde; });
    const anfangsSalden = computeBalancesGAS(kontenMap, vorMonat);
    const endSalden = computeBalancesGAS(kontenMap, bisEndeMonat);

    const monatsBuchungen = alleTx
      .filter((t) => { const d = toDateSafe(t.Datum); return d && d >= monatStart && d <= monatEnde; })
      .sort((a, b) => toDateSafe(a.Datum) - toDateSafe(b.Datum));

    const gesamtEinnahmen = monatsBuchungen.filter((t) => t.Typ === 'Einnahme').reduce((s, t) => s + parseNumberSafe(t.Betrag), 0);
    const gesamtAusgaben = monatsBuchungen.filter((t) => t.Typ === 'Ausgabe').reduce((s, t) => s + parseNumberSafe(t.Betrag), 0);

    const doc = DocumentApp.create('Kontoauszug ' + monatLabel);
    const body = doc.getBody();
    body.appendParagraph('Kontoauszug ' + monatLabel).setHeading(DocumentApp.ParagraphHeading.TITLE);
    body.appendParagraph('Einnahmen gesamt: ' + gesamtEinnahmen.toFixed(2) + ' EUR    Ausgaben gesamt: ' + gesamtAusgaben.toFixed(2) + ' EUR    Netto: ' + (gesamtEinnahmen - gesamtAusgaben).toFixed(2) + ' EUR');
    body.appendParagraph('');

    body.appendParagraph('Konten').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    const kontoTabelle = [['Konto', 'Anfangssaldo', 'Endsaldo', 'Veränderung']];
    Object.keys(kontenMap).forEach((name) => {
      const a = anfangsSalden[name] || 0, e = endSalden[name] || 0;
      kontoTabelle.push([name, a.toFixed(2) + ' EUR', e.toFixed(2) + ' EUR', (e - a).toFixed(2) + ' EUR']);
    });
    body.appendTable(kontoTabelle);
    body.appendParagraph('');

    body.appendParagraph('Buchungen').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    const buchungsTabelle = [['Datum', 'Titel', 'Kategorie', 'Typ', 'Betrag', 'Quelle', 'Ziel', 'Notiz']];
    monatsBuchungen.forEach((t) => {
      buchungsTabelle.push([
        Utilities.formatDate(toDateSafe(t.Datum), tz, 'dd.MM.yyyy'),
        t.Titel || '', t.Kategorie || '', t.Typ || '',
        parseNumberSafe(t.Betrag).toFixed(2) + ' EUR', t.Quelle || '', t.Ziel || '', t.Notiz || '',
      ]);
    });
    body.appendTable(buchungsTabelle);
    doc.saveAndClose();

    const docFile = DriveApp.getFileById(doc.getId());
    const pdfBlob = docFile.getAs('application/pdf');
    const sheetFile = DriveApp.getFileById(ss.getId());
    const parentFolder = sheetFile.getParents().hasNext() ? sheetFile.getParents().next() : DriveApp.getRootFolder();
    const ordnerName = 'Finanzindex Kontoauszüge';
    const ordnerIter = parentFolder.getFoldersByName(ordnerName);
    const ordner = ordnerIter.hasNext() ? ordnerIter.next() : parentFolder.createFolder(ordnerName);
    ordner.createFile(pdfBlob).setName('Kontoauszug ' + monatLabel + '.pdf');
    docFile.setTrashed(true); // temporäres Google Doc wird nicht benötigt, nur das PDF bleibt
    CacheService.getScriptCache().remove(KONTOAUSZUEGE_CACHE_KEY); // neuer Auszug soll sofort sichtbar sein, nicht erst nach bis zu 15 Minuten

    log.push('Kontoauszug erstellt: Kontoauszug ' + monatLabel + '.pdf (' + monatsBuchungen.length + ' Buchungen)');
  } catch (err) {
    log.push('Kontoauszug fehlgeschlagen: ' + err.message);
  }
}

// Manuelles Sofort-Erstellen eines Kontoauszugs für den Vormonat, unabhängig
// vom automatischen Monatswechsel - zum Testen oder bei Bedarf zwischendurch.
function kontoauszugJetzt() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = [];
  kontoauszugErstellen(ss, log);
  Logger.log(log.join('\n'));
}

// Legt monatlich eine datierte Kopie des kompletten Sheets im selben
// Drive-Ordner ab - ein unabhängiges Backup, das nichts mit der App im
// Browser zu tun hat und auch dann läuft, wenn die App nie geöffnet wird.
function backupErstellen(ss, log) {
  try {
    const now = new Date();
    const name = 'Finanzindex Backup ' + Utilities.formatDate(now, Session.getScriptTimeZone() || 'Europe/Berlin', 'yyyy-MM-dd');
    const file = DriveApp.getFileById(ss.getId());
    const ordner = file.getParents().hasNext() ? file.getParents().next() : DriveApp.getRootFolder();
    file.makeCopy(name, ordner);
    log.push('Backup erstellt: ' + name);
  } catch (err) {
    log.push('Backup fehlgeschlagen: ' + err.message);
  }
}

// Manuelles Sofort-Backup, unabhängig vom automatischen Monatswechsel.
// Im Apps-Script-Editor im Funktions-Dropdown auswählen und "Ausführen"
// klicken, um jederzeit auf Wunsch eine aktuelle, datierte Kopie des
// gesamten Sheets in Google Drive abzulegen.
function backupJetzt() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = [];
  backupErstellen(ss, log);
  Logger.log(log.join('\n'));
}

function notifyFehler(err, log) {
  try {
    MailApp.sendEmail(
      NOTIFY_EMAIL,
      'Finanzindex: Automatischer Buchungslauf fehlgeschlagen',
      'Der automatische Monatswechsel ist fehlgeschlagen.\n\nFehler: ' + (err && err.message ? err.message : err) +
      '\n\nBisher protokollierte Schritte:\n' + log.join('\n') +
      '\n\nBitte prüfe dein Google Sheet und führe die Buchungen bei Bedarf manuell über die App nach.'
    );
  } catch (mailErr) {
    Logger.log('Konnte Fehler-E-Mail nicht senden: ' + mailErr);
  }
}

function notifyErfolg(ausgaberestAnzahl, fixkostenErgebnis, log) {
  MailApp.sendEmail(
    NOTIFY_EMAIL,
    'Finanzindex: Monatswechsel erfolgreich',
    'Ausgaberest-Buchungen: ' + ausgaberestAnzahl +
    '\nFixkosten gebucht: ' + fixkostenErgebnis.created + ', übersprungen: ' + fixkostenErgebnis.skipped +
    '\n\nDetails:\n' + log.join('\n')
  );
}

/* ============================== Simulation zum Testen (gefahrlos) ============================== */

// Erstellt automatisch eine komplette Kopie des Sheets, spielt darin den
// Monatswechsel mit einem SIMULIERTEN Datum durch und lässt das echte,
// originale Sheet dabei vollständig unangetastet. Ergebnis: eine
// Google-Drive-URL zur Testkopie, die du dir ansehen und danach löschen
// kannst. Datum im Format 'JJJJ-MM-TT', z. B. '2026-10-01'.
function testMonatswechsel(datumString) {
  const simDate = datumString ? new Date(datumString) : new Date();
  const original = SpreadsheetApp.getActiveSpreadsheet();
  const originalFile = DriveApp.getFileById(original.getId());
  const zeitstempel = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Europe/Berlin', 'yyyy-MM-dd HH-mm-ss');
  const kopie = originalFile.makeCopy('TEST-SIMULATION ' + zeitstempel);
  const testSs = SpreadsheetApp.openById(kopie.getId());

  const log = [];
  log.push('Simuliertes Datum: ' + simDate.toDateString());
  const ausgaberestAnzahl = ausgaberestSweep(testSs, log, simDate);
  kontoauszugErstellen(testSs, log, simDate);
  const fixkostenErgebnis = fixkostenMonatslauf(testSs, log, simDate);

  log.push('---');
  log.push('Ausgaberest-Buchungen: ' + ausgaberestAnzahl);
  log.push('Fixkosten gebucht: ' + fixkostenErgebnis.created + ', übersprungen: ' + fixkostenErgebnis.skipped);
  log.push('Testkopie zum Ansehen: ' + kopie.getUrl());
  Logger.log(log.join('\n'));
  return kopie.getUrl();
}

// Bequeme, parameterlose Wrapper-Funktionen, damit du sie direkt über das
// Funktions-Dropdown auswählen und mit "Ausführen" starten kannst, ohne
// selbst Parameter eingeben zu müssen. Datum bei Bedarf anpassen.
function testAlsMonatsanfangOktober() {
  testMonatswechsel('2026-10-01');
}
function testAlsMonatsanfangNovember() {
  testMonatswechsel('2026-11-01');
}

/* ============================== Einmalige Einrichtung: ID-Spalte in Fixkosten_Plan ============================== */

// Einmal manuell ausführen (Funktions-Dropdown -> fixkostenIdsErgaenzen ->
// Ausführen), BEVOR Fixkosten-Bearbeitung aus der App genutzt wird. Legt bei
// Bedarf die Spalte "ID" an und füllt sie für alle bestehenden Zeilen, die
// noch keine haben, mit einer eindeutigen Kennung. Kann gefahrlos mehrfach
// ausgeführt werden - bereits vorhandene IDs werden nie verändert.
function fixkostenIdsErgaenzen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.FIXKOSTEN);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  let idCol = headers.indexOf('ID');
  if (idCol === -1) {
    idCol = headers.length;
    sheet.getRange(1, idCol + 1).setValue('ID');
  }
  let anzahl = 0;
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (row.every((c) => c === '' || c === null)) continue;
    if (!row[idCol]) {
      sheet.getRange(r + 1, idCol + 1).setValue(Utilities.getUuid());
      anzahl++;
    }
  }
  Logger.log('Fixkosten-IDs ergänzt: ' + anzahl);
}

/* ============================== Einmalige Einrichtung des Zeit-Triggers ============================== */

function installTrigger() {
  // Vorhandene Trigger für diese Funktion zuerst entfernen, damit sie nicht doppelt läuft
  ScriptApp.getProjectTriggers().forEach((t) => {
    if (t.getHandlerFunction() === 'monatswechselAutomatik') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('monatswechselAutomatik')
    .timeBased()
    .onMonthDay(1)
    .atHour(1)
    .create();
  Logger.log('Trigger eingerichtet: monatswechselAutomatik läuft künftig am 1. jeden Monats zwischen 1 und 2 Uhr.');
}

/* ============================== KI-Assistent (Gemini) ==============================
 * EINRICHTUNG (einmalig):
 * 1. Kostenlosen API-Schlüssel unter aistudio.google.com erzeugen.
 * 2. Hier im Skript: Projekteinstellungen (Zahnrad links) -> Script-Eigenschaften ->
 *    "Property hinzufügen" -> Name: GEMINI_API_KEY, Wert: der erzeugte Schlüssel.
 *    NICHT hier im Code eintragen - Script-Eigenschaften sind der sichere Ort dafür,
 *    im Gegensatz zum Code landen sie nie im (potenziell einsehbaren) Quelltext.
 * 3. Nach dem Speichern dieses Codes ggf. einmal neu bereitstellen (Bereitstellen ->
 *    Bereitstellungen verwalten -> Bearbeiten -> Neue Version), damit die App die
 *    neue Aktion "ai-chat" kennt.
 */

// Bei Bedarf hier austauschen, falls Google den Modellnamen irgendwann wieder
// ändert - Google meldet einen veralteten Namen im Fehlertext meist mit dem
// aktuell empfohlenen Ersatz gleich mit.
const GEMINI_MODEL = 'gemini-3.6-flash';

function getGeminiApiKey() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
}

// Statischer FAQ-Inhalt (1:1 aus faq.html extrahiert) - Grundlage für
// Bedienungsfragen. Bewusst als fester Text statt live aus faq.html
// nachgeladen: kein zusätzlicher externer Aufruf nötig, dadurch schneller
// und ein Ausfall von GitHub Pages hätte keine Auswirkung. Bei inhaltlichen
// Änderungen an faq.html diesen Text von Hand nachziehen.
const FAQ_TEXT = `## Buchungen & Verteilung
F: Wie wird die monatliche Zinsgutschrift von Trade Republic gebucht?
A: Einnahme buchen: Quelle "extern", Ziel "Transit", Betrag = gesamte Zinsgutschrift. Danach die gewünschte Verteilung als einzelne Umbuchungen von Transit zu den Zielkonten vornehmen (zum Beispiel Transit → Model 3, Transit → Rücklage).
F: Wie werden ungeplante Sondereinnahmen gebucht?
A: Direkt als Einnahme auf das Konto buchen, für das das Geld gedacht ist (Quelle "extern" oder "arbeitgeber").
F: Was ist "Ausgaberest" und wie wird er weiterverwendet?
A: Am Monatsende werden Laden, E-Liquids, Lebensmittel, Freizeit, Kosmetik, Abo & ETF und Puffer automatisch auf 0 zurückgesetzt. Guthaben wandert zu Ausgaberest, Überzüge werden von dort ausgeglichen – so entsteht dort automatisch der echte Netto-Wert. Das gesammelte Geld kann jederzeit per Umbuchung verteilt werden (zum Beispiel Ausgaberest → Model 3 oder Ausgaberest → Jahres Rückstellung).
F: Was ist "Budgetrest" und woraus setzt er sich zusammen?
A: Die Summe der aktuellen Kontostände von Laden, E-Liquids, Lebensmittel, Freizeit, Kosmetik, Abo & ETF und Puffer – also derselben sieben Töpfe, die auch den Ausgaberest speisen. Er zeigt live, wie viel von deinem laufenden Monatsbudget in diesen Töpfen zusammengenommen noch übrig ist, während der Monat noch läuft.
F: Worin unterscheiden sich Budgetrest und Ausgaberest?
A: Budgetrest ist eine reine Anzeige, kein echtes Konto: Er wird bei jedem Öffnen der App live aus den sieben Töpfen neu zusammengerechnet und zeigt "was ist gerade, während der Monat noch läuft, in diesen Töpfen übrig". Ausgaberest dagegen ist ein echtes Konto mit eigenem, dauerhaftem Kontostand, das erst am Monatsende durch den automatischen Ausgleich befüllt wird und "was am Ende des Monats übrig geblieben ist" dauerhaft festhält. Solange der Monat läuft, ist Ausgaberest 0 (oder der Rest vom Vormonat, falls noch nicht verteilt) - Budgetrest ändert sich dagegen laufend mit jeder Buchung.
F: Wie viel muss manuell von ING auf Trade Republic überwiesen werden?
A: Tab "Fixkosten" → Karte "Gehaltsverteilung" zeigt den Betrag, der auf ING bleiben soll, und den Betrag, der auf Trade Republic überwiesen werden muss – berechnet aus den aktuell gültigen Fixkosten-Positionen.
F: Warum sollte "Trade Republic gesamt" im Dashboard immer mit der echten App übereinstimmen?
A: Dieser Wert entsteht rein rechnerisch aus den erfassten Buchungen. Eine Abweichung zeigt zuverlässig: eine Buchung fehlt oder ist falsch erfasst – nicht, dass die Berechnung selbst falsch wäre.

## Automatisierung
F: Was passiert automatisch am 1. jeden Monats?
A: Zuerst wird der Ausgaberest-Ausgleich für den Vormonat gebucht, danach die neuen Fixkosten und die Besoldung für den aktuellen Monat, danach ein automatisches Backup erstellt. Läuft serverseitig über Google Apps Script, unabhängig davon, ob die App-Seite geöffnet ist.
F: Wie lässt sich der Monatswechsel vorab gefahrlos testen?
A: Im Apps-Script-Editor die Funktion testAlsMonatsanfangOktober (oder testAlsMonatsanfangNovember) ausführen. Das erstellt automatisch eine komplette Testkopie des Sheets und lässt das Original unangetastet.
F: Was passiert, wenn die Automatisierung fehlschlägt?
A: Es wird automatisch eine E-Mail an die im Skript hinterlegte Adresse verschickt, inklusive Fehlermeldung und Protokoll der bereits erledigten Schritte.

## Sicherung (Backup)
F: Muss selbst etwas für Backups getan werden?
A: Nein. Sobald installTrigger einmal ausgeführt wurde, legt das Skript automatisch jeden Monat eine datierte Kopie des kompletten Sheets in Google Drive ab (Name: "Finanzindex Backup JJJJ-MM-TT").
F: Wie lässt sich sofort eine Sicherung erstellen, ohne zu warten?
A: Zwei Wege: (1) In Google Sheets oben "Datei" → "Kopie erstellen" – sofort, ganz ohne Skript. (2) Im Apps-Script-Editor die Funktion backupJetzt auswählen und ausführen – erstellt dieselbe Art von Backup wie die automatische, nur sofort auf Wunsch.
F: Wie wird im Notfall aus einer Sicherung wiederhergestellt?
A: Bei kleineren, kürzlichen Fehlern reicht oft "Datei" → "Versionsverlauf" direkt in Google Sheets, ganz ohne Backup-Datei. Bei größeren Problemen: die Backup-Datei in Drive öffnen und die benötigten Werte manuell zurück ins Original-Sheet übertragen.

## Kontoauszug
F: Was ist der Kontoauszug und wann entsteht er?
A: Ein fester, unveränderlicher PDF-Schnappschuss des abgelaufenen Monats – anders als das Sheet, das sich laufend weiter verändert. Er entsteht automatisch als Teil des monatlichen Ablaufs: Ausgaberest-Ausgleich → Kontoauszug → neue Fixkosten → Backup. Läuft also ebenfalls am 1. jeden Monats zwischen 1 und 2 Uhr, sobald installTrigger eingerichtet ist.
F: Was steht genau im Kontoauszug?
A: Oben die Gesamtsummen (Einnahmen, Ausgaben, Netto), danach je Konto der Anfangs- und Endsaldo sowie die Veränderung, danach alle Einzelbuchungen des Monats chronologisch – inklusive der automatischen Ausgaberest-Ausgleichsbuchungen, die noch im alten Monat datiert sind.
F: Wo wird der Kontoauszug gespeichert und wie sehe ich ihn mir an?
A: Als PDF im Drive-Ordner "Finanzindex Kontoauszüge" (wird beim ersten Mal automatisch angelegt), neben dem Sheet. In der App im Tab "Auswertung" erscheint unten eine Liste aller bisher erstellten Kontoauszüge zum direkten Antippen und Ansehen – ohne selbst etwas herunterzuladen, da die Datei ja bereits in Drive liegt.
F: Wie lässt sich sofort ein Kontoauszug erstellen, ohne zu warten?
A: Im Apps-Script-Editor die Funktion kontoauszugJetzt auswählen und ausführen – erstellt den Kontoauszug für den Vormonat sofort, unabhängig vom automatischen Monatswechsel.

## Sicherheit
F: Was bedeutet "X Geräte online" oben im Header?
A: Eine reine Kontrollanzeige. Jedes geöffnete Gerät meldet sich alle 30 Sekunden im Tabellenblatt "Geräte" (Geräte-ID, Name, Zeitpunkt, Gesperrt). Als "online" zählt ein Gerät, dessen letztes Lebenszeichen höchstens 60 Sekunden alt ist - danach fällt es automatisch wieder raus, auch wenn die App mal nicht sauber beendet wurde. Antippen öffnet eine Liste aller bekannten Geräte mit Zeitpunkt und einem "Sperren"-Knopf je fremdem Gerät. Taucht eine höhere Zahl auf, als eigene Geräte gerade wirklich geöffnet sind, ist das ein Hinweis auf einen fremden Zugriff.
F: Was passiert, wenn ich ein Gerät in der Liste "sperre"?
A: Das Gerät bekommt spätestens bei seinem nächsten Lebenszeichen (innerhalb von rund 30 Sekunden) einen vollflächigen Sperrbildschirm angezeigt - unabhängig davon, ob auf diesem Gerät eine PIN eingerichtet ist. Es kann die App danach nicht mehr benutzen. Das eigene, gerade aktiv genutzte Gerät lässt sich absichtlich nicht sperren.
F: Wie hebe ich eine Sperre wieder auf, falls ich versehentlich das falsche Gerät gesperrt habe?
A: Absichtlich nicht aus der App heraus möglich. Im Tabellenblatt "Geräte" in der Spalte "Gesperrt" das Häkchen bei der entsprechenden Zeile per Hand entfernen. Das ist bewusst so gebaut: Ein fremdes Gerät kennt zwar eventuell den API-Schlüssel, hat aber keinen Zugriff auf das Google-Konto selbst - die Sperre lässt sich also nur von dir selbst aufheben, nie vom gesperrten Gerät.
F: Wo lege ich fest, wie ein Gerät in der Liste heißt?
A: In den Einstellungen unter "Gerätename" - einmal pro Gerät ausfüllen (z. B. "Mi 13", "Mi Pad 7 Pro", "PC"). Ohne eingetragenen Namen erscheint das Gerät als "Unbenanntes Gerät".
F: GitHub ist öffentlich - kann jemand Fremdes meine echten Daten sehen?
A: Nein, nicht allein dadurch. Öffentlich einsehbar ist nur der Programmcode (index.html, faq.html) - eine leere Hülle ganz ohne echte Zahlen. Die echten Daten kommen erst über die im Browser hinterlegte Apps-Script-URL, die nirgendwo im Code steht, sondern ausschließlich lokal auf dem jeweiligen Gerät gespeichert ist.
F: Kann jemand über GitHub meine Datei verändern, weil sie öffentlich lesbar ist?
A: Nein. Lesen und Verändern sind bei GitHub getrennte Rechte. Ändern kann ausschließlich der Besitzer des GitHub-Kontos (oder wer ausdrücklich mit Schreibrecht eingeladen wurde) - unabhängig davon, ob die Datei öffentlich sichtbar ist.
F: Was ist der "API-Schlüssel" in den Einstellungen und wofür ist er gut?
A: Ein zusätzliches, zufällig erzeugtes Kennwort, das im Skript hinterlegt ist und bei jeder Anfrage zusätzlich zur Apps-Script-URL mitgeschickt werden muss. Ohne diesen Schlüssel beantwortet das Skript keine Anfrage - selbst wenn die URL allein irgendwo sichtbar werden sollte, reicht das allein nicht mehr für einen Zugriff.
F: Warum sind URL und API-Schlüssel in den Einstellungen wie ein Passwort maskiert?
A: Damit die beiden Werte nicht schon durch einen kurzen Blick über die Schulter sichtbar sind. Ist auf diesem Gerät eine Sicherheitssperre eingerichtet, verlangt "Anzeigen" zusätzlich eine erneute Bestätigung per PIN oder Fingerabdruck, bevor die Werte im Klartext erscheinen - auch innerhalb einer bereits entsperrten App-Sitzung.
F: Ist die FAQ-Seite (diese Seite hier) auch öffentlich?
A: Ja - anders als die eigentliche App hat diese FAQ-Seite keine eigene Sperre und ist für jeden mit der passenden Adresse lesbar. Da hier aber keinerlei echte Finanzdaten stehen, nur allgemeine Erklärungen zur Funktionsweise, ist das unbedenklich.

## Bedienung
F: Was bedeutet der Betrag in Klammern hinter einem Topfnamen?
A: Der aktuell gültige monatliche Fixkosten-Betrag für diesen Topf. Transit und Puffer zeigen keinen, da sie anders funktionieren (mehrere Quellen beziehungsweise reiner Reservetopf).
F: Kann ich Fixkosten-Positionen direkt in der App bearbeiten?
A: Ja. Anlegen, Bearbeiten und Löschen im Tab "Fixkosten" schreibt direkt ins Sheet, genau wie bei normalen Buchungen – inklusive automatischem Nachholen bei einer Verbindungsstörung. Voraussetzung dafür ist eine eindeutige "ID"-Spalte in Fixkosten_Plan (einmalig über die Funktion fixkostenIdsErgaenzen im Apps-Script-Editor eingerichtet).
F: Warum wird ein Betrag rot angezeigt?
A: Rot zeigt einen negativen, überzogenen Kontostand an.
F: Wo befinden sich Fixkosten, Prognose, Auswertung und Einstellungen auf dem Smartphone?
A: Über den Button "Mehr" in der unteren Leiste.
F: Wie wird zwischen hellem und dunklem Design gewechselt?
A: Über das Sonne-/Mond-Symbol oben rechts im Header der App.
F: Was bedeutet die Zeitangabe neben dem Sync-Symbol (z. B. "vor 2 Min")?
A: Das ist der Zeitpunkt der letzten erfolgreich bestätigten Synchronisierung mit dem Google Sheet, keine Live-Uhr. Während gerade synchronisiert wird, steht dort "läuft…", bei einem Fehlschlag "Fehler". Ein Antippen des Symbols stößt jederzeit eine erneute Synchronisierung an.
F: Warum dauert die Synchronisierung auf dem Smartphone manchmal sehr lange?
A: Die Oberfläche zeigt sofort den zuletzt gespeicherten Stand aus einem lokalen Zwischenspeicher an – die Zahlen erscheinen also unmittelbar. Im Hintergrund läuft parallel die eigentliche Verbindung zu Google Apps Script, die bei schwachem Mobilfunkempfang deutlich länger dauern kann, da die Antwort dabei technisch über einen zusätzlichen Zwischenschritt (eine Weiterleitung) ausgeliefert wird. Das ist eine bekannte Einschränkung von Google Apps Script bei mobilen Verbindungen, keine Fehlfunktion der App. Die App bleibt währenddessen normal bedienbar.

## KI-Assistent
F: Was kann der KI-Assistent (Gemini-Symbol in der unteren Leiste)?
A: Vier Dinge: (1) Neue Buchungen per Text oder Sprache anlegen und bestehende Buchungen ändern – immer erst als Vorschlag zur Bestätigung, nie automatisch. (2) Fragen zur Bedienung und Funktionsweise der App beantworten, wie hier in der FAQ. (3) Per Sprache oder Text zu einem beliebigen Menüpunkt springen (z. B. "Zeig mir die Auswertung"). (4) Ganz allgemeine Fragen beantworten, die nichts mit der App zu tun haben. Bewusst nicht dabei: Fragen zu genauen Kontoständen oder Ausgabensummen – Sprachmodelle rechnen bei mehreren Zahlen unzuverlässig, dafür bitte immer direkt ins Dashboard bzw. die Auswertung schauen.
F: Kann der Assistent auch Buchungen löschen?
A: Nein, das ist bewusst nicht vorgesehen. Aus Sicherheitsgründen kann der Assistent ausschließlich neue Buchungen anlegen oder bestehende ändern, niemals löschen. Löschen bleibt ausschließlich manuell im Journal möglich.

## Buchungslogos
F: Warum stehen bei manchen Buchungen echte Firmenlogos (z. B. Netflix, Tesla, Edeka)?
A: Rein zur besseren Übersicht in der eigenen, privaten Buchungsliste – kein Sponsoring, keine Zusammenarbeit und keine Aussage über die dargestellten Marken. Ein Logo wird rein technisch anhand des selbst eingegebenen Buchungstexts erkannt (enthält er z. B. das Wort "Netflix", erscheint das Netflix-Logo) – vergleichbar mit einer Banking-App, die neben Buchungen ebenfalls Händler-Logos zeigt. Die Nutzung erfolgt ausschließlich privat, nicht-kommerziell und rein zur Identifikation der jeweiligen Buchung. Alle Rechte an den gezeigten Logos liegen bei den jeweiligen Markeninhabern. Ist kein passendes Logo hinterlegt, erscheint stattdessen ein grauer Kreis mit dem Anfangsbuchstaben des Buchungstexts.

## Technischer Überblick
F: Wo liegen die Daten wirklich?
A: Ausschließlich im verknüpften Google Sheet (Blätter: Konten, Kategorien, Fixkosten_Plan, Transaktionen, Einstellungen). Das Sheet ist die einzige echte Datenquelle.
F: Was ist die Webseite (index.html) eigentlich?
A: Eine reine Oberfläche ohne eigene Datenspeicherung. Sie liegt bei GitHub Pages, lädt beim Öffnen React, Chart.js und weitere Bausteine live aus dem Internet und ruft anschließend die aktuellen Daten aus dem Google Sheet ab.
F: Was ist das Google Apps Script und wo liegt es?
A: Ein Programm, direkt an das Sheet gebunden (Erweiterungen → Apps Script). Es übernimmt zwei Aufgaben: Als Web-App beantwortet es die Anfragen der Seite (Daten lesen/schreiben), und unabhängig davon läuft es einmal im Monat automatisch (Ausgaberest, Fixkosten, Backup) – auch wenn die Seite nie geöffnet wird.
F: Wie hängen GitHub, die App und das Sheet zusammen?
A: GitHub Pages (Ort für den Programmcode) → Web-App-URL → Google Apps Script → liest und schreibt im Google Sheet. Bei einer Weiterentwicklung der App wird eine neue index.html und/oder .gs-Datei bereitgestellt, die die bisherige Version ersetzt – das Sheet mit den eigenen Daten bleibt davon immer unberührt.
F: Was passiert ohne Internetverbindung?
A: Es wird der zuletzt geladene Stand aus einem lokalen Zwischenspeicher im jeweiligen Browser angezeigt. Neue Buchungen können ohne Internetverbindung nicht gespeichert werden, da nichts ohne Verbindung ins Sheet geschrieben werden kann.
F: Warum wurden Buchungsdaten früher manchmal einen Tag zu früh angezeigt?
A: Google Apps Script rechnete Datumswerte beim Verpacken in die Antwort automatisch in Weltzeit (UTC) um, wodurch ein deutsches Datum um einen Tag zurückrutschen konnte. Behoben, indem Apps Script Datumswerte seitdem als eindeutigen Text (z. B. "2026-09-16") statt als Zeitwert verschickt. Die eigentlichen Buchungen im Sheet waren davon nie betroffen, nur die Anzeige in der App nach dem Laden.`;

// Leichtgewichtiger Kontext nur noch für Buchungsvorschläge: bekannte
// Konten/Töpfe (ohne Salden - die KI soll dazu nichts mehr berechnen/nennen
// müssen) und Kategorien, plus die letzten Buchungen inkl. ID (nötig, damit
// "buchung_bearbeiten" eine konkrete Zeile referenzieren kann). Bewusst OHNE
// jede Saldo-Berechnung mehr - das war zugleich der teuerste Teil pro
// Anfrage UND die Quelle falscher Zahlen, seit die KI keine Kontostände mehr
// nennen soll, wird das schlicht nicht mehr gebraucht.
function buildKIKontext(ss) {
  const now = new Date();
  const kontenMap = getKontenMap(ss);
  const transaktionen = sheetToObjects(ss.getSheetByName(SHEETS.TRANSAKTIONEN));
  const fixkosten = sheetToObjects(ss.getSheetByName(SHEETS.FIXKOSTEN));
  const aktiveFixkosten = getFixkostenAktiv(fixkosten, now.getFullYear(), now.getMonth());

  const kontenListe = Object.keys(kontenMap).join(', ');

  // Auf 15 statt vorher 40 reduziert (reicht für "die Buchung von gestern/
  // letzte Woche") - hält den Prompt an Gemini kleiner, seit zusätzlich der
  // FAQ-Text und zwei weitere Funktionsdeklarationen mitgeschickt werden.
  const letzteBuchungen = transaktionen
    .slice()
    .sort((a, b) => (toDateSafe(b.Datum) || 0) - (toDateSafe(a.Datum) || 0))
    .slice(0, 15)
    .map((t) => 'ID ' + t.ID + ' | ' + t.Datum + ' | ' + t.Titel + ' | ' + t.Typ + ' | ' + parseNumberSafe(t.Betrag).toFixed(2) + ' EUR | ' + t.Kategorie + ' | ' + t.Quelle + (t.Ziel ? ' -> ' + t.Ziel : ''))
    .join('\n');

  const kategorienSet = {};
  aktiveFixkosten.forEach((f) => { if (f.Kategorie) kategorienSet[f.Kategorie] = true; });
  transaktionen.forEach((t) => { if (t.Kategorie) kategorienSet[t.Kategorie] = true; });
  const kategorienListe = Object.keys(kategorienSet).join(', ');

  return [
    'BEKANNTE KONTEN/TÖPFE (für Quelle/Ziel bei Buchungen - keine Salden, dazu nichts berechnen oder nennen):', kontenListe || '(keine)',
    '', 'BEKANNTE KATEGORIEN:', kategorienListe || '(keine)',
    '', 'LETZTE BUCHUNGEN (neueste zuerst, max. 40, inkl. ID zum Referenzieren bei buchung_bearbeiten):', letzteBuchungen || '(keine)',
  ].join('\n');
}

// Ruft die Gemini-API auf. "contents" ist der bisherige Gesprächsverlauf
// (role "user"/"model" abwechselnd), "systemText" die Systemanweisung samt
// FAQ- und Buchungskontext. Deklariert drei Funktionen, die Gemini aufrufen
// kann statt nur Text zu schreiben:
// - buchung_vorschlagen: neue Buchung, immer nur als Vorschlag zur
//   Bestätigung angezeigt, nie automatisch gebucht.
// - buchung_bearbeiten: bestehende Buchung ändern (per ID), ebenfalls nur
//   als Vorschlag. Absichtlich KEINE Lösch-Funktion vorhanden.
// - navigiere_zu: wechselt den aktiven Tab in der App.
function callGemini(systemText, contents) {
  const apiKey = getGeminiApiKey();
  if (!apiKey) {
    throw new Error('Kein Gemini-API-Schlüssel hinterlegt. Siehe Kommentar über diesem Abschnitt für die Einrichtung.');
  }

  const buchungsFelder = {
    datum: { type: 'STRING', description: 'Datum im Format JJJJ-MM-TT, ohne Angabe: heute' },
    titel: { type: 'STRING', description: 'Kurzer Titel/Empfänger, z.B. "Aral Tanken"' },
    typ: { type: 'STRING', enum: ['Ausgabe', 'Einnahme', 'Umbuchung'] },
    betrag: { type: 'NUMBER', description: 'Betrag als positive Zahl' },
    kategorie: { type: 'STRING', description: 'Passende Kategorie, möglichst eine bereits bekannte' },
    quelle: { type: 'STRING', description: 'Herkunfts-Konto/Topf, ohne Angabe: "ING"' },
    ziel: { type: 'STRING', description: 'Nur bei Umbuchung/Einnahme relevant: Ziel-Konto/Topf' },
    notiz: { type: 'STRING', description: 'Optionale zusätzliche Notiz' },
  };

  const buchungsFunktion = {
    name: 'buchung_vorschlagen',
    description: 'Schlägt eine neue Buchung (Ausgabe, Einnahme oder Umbuchung) vor, wenn der Nutzer erkennbar etwas gebucht haben möchte. Wird dem Nutzer nur als Vorschlag zur Bestätigung angezeigt, nie automatisch gebucht.',
    parameters: { type: 'OBJECT', properties: buchungsFelder, required: ['titel', 'typ', 'betrag', 'kategorie', 'quelle'] },
  };

  const bearbeitenFunktion = {
    name: 'buchung_bearbeiten',
    description: 'Ändert eine bestehende Buchung aus der Liste "LETZTE BUCHUNGEN". Nur Felder angeben, die sich wirklich ändern - nicht genannte Felder bleiben unverändert. Wird dem Nutzer nur als Vorschlag zur Bestätigung angezeigt, nie automatisch gespeichert. Zum Löschen NICHT verwenden - Löschen ist nicht möglich.',
    parameters: {
      type: 'OBJECT',
      properties: Object.assign({ id: { type: 'STRING', description: 'ID der zu ändernden Buchung aus der Liste "LETZTE BUCHUNGEN"' } }, buchungsFelder),
      required: ['id'],
    },
  };

  const navigationsFunktion = {
    name: 'navigiere_zu',
    description: 'Wechselt in der App zu einem der Hauptbereiche. Nutzen, wenn der Nutzer erkennbar irgendwohin wechseln oder etwas ansehen möchte, z.B. "zeig mir das Journal" oder "ich will die Auswertung sehen".',
    parameters: {
      type: 'OBJECT',
      properties: { ziel: { type: 'STRING', enum: ['dashboard', 'buchen', 'journal', 'fixkosten', 'prognose', 'auswertung', 'einstellungen'] } },
      required: ['ziel'],
    },
  };

  const payload = {
    system_instruction: { parts: [{ text: systemText }] },
    contents: contents,
    tools: [{ function_declarations: [buchungsFunktion, bearbeitenFunktion, navigationsFunktion] }],
  };

  const res = UrlFetchApp.fetch(
    'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent?key=' + encodeURIComponent(apiKey),
    {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    }
  );

  const status = res.getResponseCode();
  let json;
  try {
    json = JSON.parse(res.getContentText());
  } catch (e) {
    throw new Error('Antwort von Gemini konnte nicht gelesen werden (Status ' + status + ').');
  }
  if (status !== 200) {
    throw new Error('Gemini-Fehler: ' + (json.error && json.error.message ? json.error.message : res.getContentText()));
  }

  const candidate = json.candidates && json.candidates[0];
  const parts = (candidate && candidate.content && candidate.content.parts) || [];
  let text = '';
  let buchungsvorschlag = null;
  let navigation = null;
  parts.forEach((p) => {
    if (p.text) text += p.text;
    if (p.functionCall && (p.functionCall.name === 'buchung_vorschlagen' || p.functionCall.name === 'buchung_bearbeiten')) {
      buchungsvorschlag = p.functionCall.args;
    }
    if (p.functionCall && p.functionCall.name === 'navigiere_zu') {
      navigation = p.functionCall.args && p.functionCall.args.ziel;
    }
  });
  return { text: bereinigeMarkdown(text.trim()), buchungsvorschlag: buchungsvorschlag, navigation: navigation };
}

// Sicherheitsnetz, falls Gemini die Anweisung "kein Markdown" doch mal
// ignoriert - entfernt die gängigsten Formatierungszeichen, damit weder die
// Chat-Anzeige noch die Sprachausgabe sie wörtlich mitliest/-zeigt.
function bereinigeMarkdown(text) {
  return text
    .replace(/\*\*(.*?)\*\*/g, '$1')
    .replace(/(^|\s)\*([^\s*][^*]*?)\*(?=\s|$)/g, '$1$2')
    .replace(/^#{1,6}\s+/gm, '')
    .replace(/^[*-]\s+/gm, '- ')
    .replace(/\s\*\s/g, '. ')
    .replace(/`([^`]+)`/g, '$1')
    .trim();
}

function handleAiChat(ss, body) {
  const nachricht = String(body.nachricht || '').trim();
  if (!nachricht) throw new Error('Keine Nachricht übermittelt.');
  const verlauf = Array.isArray(body.verlauf) ? body.verlauf : [];

  const kontext = buildKIKontext(ss);
  const heute = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Europe/Berlin', 'yyyy-MM-dd');

  const systemText = [
    'Du bist der Assistent in der App "Finanzindex" - antworte kurz, konkret und auf Deutsch.',
    'Heutiges Datum: ' + heute + '.',
    'Du kannst genau vier Dinge tun:',
    '1. Neue Buchungen anlegen: Wenn der Nutzer erkennbar etwas gebucht haben möchte (z.B. "Buche 12 Euro Aral von ING"), rufe buchung_vorschlagen auf.',
    '2. Bestehende Buchungen ändern: Wenn der Nutzer eine vorhandene Buchung ändern möchte (z.B. "ändere die Tanken-Buchung von gestern auf 15 Euro"), suche sie in der Liste "LETZTE BUCHUNGEN" unten anhand von Titel/Datum/Betrag und rufe buchung_bearbeiten mit deren ID auf, nur mit den sich wirklich ändernden Feldern. Findest du keine eindeutig passende Buchung, frag nach statt zu raten.',
    '3. Buchungen LÖSCHEN kannst du NICHT und darfst das auch nicht anbieten oder andeuten - sag stattdessen, dass Löschen nur manuell im Journal geht.',
    '4. Zwischen den Bereichen der App wechseln: Bei Wünschen wie "zeig mir X" oder "öffne Y" rufe navigiere_zu auf.',
    'Sowohl Buchungsvorschläge als auch Änderungsvorschläge werden dem Nutzer nur zur Bestätigung angezeigt, nie automatisch übernommen.',
    'Fragen zur Bedienung/Funktionsweise der App beantwortest du anhand des FAQ-Textes unten, in eigenen Worten statt wortwörtlich vorgelesen.',
    'Allgemeine Fragen ohne Bezug zur App beantwortest du ganz normal wie jede andere Frage auch.',
    'WICHTIG: Du kennst weder genaue Kontostände noch Ausgabensummen und darfst dazu auch nichts berechnen oder schätzen - bei solchen Fragen antworte, dass dafür ein Blick ins Dashboard bzw. die Auswertung in der App zuverlässiger ist.',
    'Antworte in reinem Klartext ohne Markdown-Formatierung - keine Sternchen für fett/kursiv, keine Rauten für Überschriften, keine Sternchen/Bindestriche als Aufzählungszeichen. Bei mehreren Punkten stattdessen einfach neue Zeilen oder normale Sätze verwenden. Die Antwort wird sowohl als Text angezeigt als auch vorgelesen, Formatierungszeichen würden dabei wörtlich mitgelesen.',
    '',
    '=== FAQ (für Bedienungsfragen) ===',
    FAQ_TEXT,
    '',
    '=== KONTEXT FÜR BUCHUNGEN ===',
    kontext,
  ].join('\n');

  const contents = verlauf
    .filter((m) => m && m.text)
    .slice(-10)
    .map((m) => ({ role: m.role === 'assistant' ? 'model' : 'user', parts: [{ text: String(m.text) }] }));
  contents.push({ role: 'user', parts: [{ text: nachricht }] });

  return callGemini(systemText, contents);
}
