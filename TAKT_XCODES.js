// ═══════════════════════════════════════════════════════════════
// X-CODE REFERENZ — 54235 FFV12024 V20 — 76 Codes komplett
// Quelle: Excel Sheet "Ventile" + "Taktung" (V20, Stand aktuell)
// Drop-in für vsp.js — ersetzt das bestehende takt-Objekt
// ═══════════════════════════════════════════════════════════════

const TAKT = {
  // ── Grundstatus ──
  "X":    { t:"Ventil geöffnet", cat:"basis" },

  // ── Getaktete Ventile (Zeitsteuerung) ──
  "X1":   { t:"Getaktet: 30s offen / 60s zu", cat:"takt", on:30, off:60 },
  "X2":   { t:"Getaktet: 20s offen / 20s zu", cat:"takt", on:20, off:20 },
  "X2N":  { t:"Getaktet: 20s zu / 20s offen (invertiert)", cat:"takt", on:20, off:20 },
  "X3":   { t:"Getaktet: 3s offen / 12s zu", cat:"takt", on:3, off:12 },
  "X3N":  { t:"Getaktet: 12s offen / 3s zu (invertiert)", cat:"takt", on:12, off:3 },
  "X4":   { t:"Getaktet: offen bis Signal Füllstandsmessstelle Produktfilter", cat:"takt" },
  "X6":   { t:"Getaktet: 5s offen / 5s zu", cat:"takt", on:5, off:5 },
  "X9":   { t:"Getaktet: 1s offen / 9s zu", cat:"takt", on:1, off:9 },
  "X12":  { t:"Getaktet: 4s offen / 20s zu", cat:"takt", on:4, off:20 },
  "X13":  { t:"Getaktet: 1s offen / 9s zu", cat:"takt", on:1, off:9 },
  "X18":  { t:"Getaktet: 1s offen / 3s zu", cat:"takt", on:1, off:3 },
  "X19":  { t:"Getaktet: 20s zu / 4s offen", cat:"takt", on:4, off:20 },
  "X20":  { t:"Getaktet: 9s zu / 1s offen", cat:"takt", on:1, off:9 },
  "X21":  { t:"Getaktet: 10s offen / 20s zu", cat:"takt", on:10, off:20 },
  "X24":  { t:"Getaktet: 10s offen / 5s zu", cat:"takt", on:10, off:5 },
  "X24N": { t:"Getaktet: 5s offen / 10s zu (invertiert)", cat:"takt", on:5, off:10 },
  "X25":  { t:"Getaktet: 8s offen / 2s zu", cat:"takt", on:8, off:2 },
  "X25N": { t:"Getaktet: 8s zu / 2s offen (invertiert)", cat:"takt", on:2, off:8 },
  "X33":  { t:"Getaktet: 1s offen / 15s zu", cat:"takt", on:1, off:15 },
  "X34":  { t:"Getaktet: 15s offen / 15s zu", cat:"takt", on:15, off:15 },
  "X36":  { t:"Getaktet: 3s offen / 3s zu", cat:"takt", on:3, off:3 },
  "X36N": { t:"Getaktet: 3s zu / 3s offen (invertiert)", cat:"takt", on:3, off:3 },
  "X37":  { t:"Getaktet: 30s offen / 30s zu", cat:"takt", on:30, off:30 },
  "X38":  { t:"Getaktet: 30s zu / 30s offen (invertiert)", cat:"takt", on:30, off:30 },
  "X39":  { t:"Getaktet: 15s zu / 15s offen (invertiert)", cat:"takt", on:15, off:15 },
  "X45":  { t:"Druckgetaktet: offen bei PT5 < 0,5 bar / zu bei PT5 > 1,5 bar", cat:"druck" },
  "X48":  { t:"20s zu, danach restliche Schrittzeit offen", cat:"takt", off:20 },
  "X50":  { t:"Getaktet: 50s offen / 5s zu", cat:"takt", on:50, off:5 },
  "X51":  { t:"60s offen, danach restliche Schrittzeit zu", cat:"takt", on:60 },

  // ── Druckabhängige Ventile ──
  "X5":   { t:"Druckgesteuert: zu bei PT6 < 0,3 bar / offen bei PT6 > 0,5 bar", cat:"druck", sensor:"PT6", lo:0.3, hi:0.5 },
  "X11":  { t:"EPOS-Rampe: 0→36% öffnen; zu bei P < 0,2 bar; offen bei PT5 > 0,5 bar", cat:"druck", sensor:"PT5" },
  "X14":  { t:"Druckgesteuert: zu bei PT5 < 0,6 bar / offen bei PT5 > 1,2 bar", cat:"druck", sensor:"PT5", lo:0.6, hi:1.2 },
  "X15":  { t:"Druckgesteuert: zu bei PT5 < 0,4 bar / offen bei PT5 > 0,8 bar", cat:"druck", sensor:"PT5", lo:0.4, hi:0.8 },
  "X16":  { t:"Druckgesteuert: zu bei PT5 < 0,6 bar / offen bei PT5 > 2,0 bar", cat:"druck", sensor:"PT5", lo:0.6, hi:2.0 },
  "X17":  { t:"Druckgesteuert: zu bei PT5 < 0,4 bar / offen bei PT5 > 1,0 bar", cat:"druck", sensor:"PT5", lo:0.4, hi:1.0 },
  "X22":  { t:"Druckgesteuert: zu bei PT1 < 0,4 bar / offen bei PT1 > 0,8 bar", cat:"druck", sensor:"PT1", lo:0.4, hi:0.8 },
  "X23":  { t:"Druckgesteuert: zu bei PT1 < 0,6 bar / offen bei PT1 > 1,0 bar", cat:"druck", sensor:"PT1", lo:0.6, hi:1.0 },
  "X35":  { t:"Druckgesteuert: zu bei PT5 < 1,0 bar / offen bei PT5 > 3,3 bar", cat:"druck", sensor:"PT5", lo:1.0, hi:3.3 },
  "X40":  { t:"Druckgesteuert: zu bei PT5 < 0,2 bar / offen bei PT5 > 0,4 bar", cat:"druck", sensor:"PT5", lo:0.2, hi:0.4 },

  // ── Bedingte Ventile / Schrittsteuerung ──
  "X7":   { t:"Nur offen bei Prog. 6, 7, 20, 21, 23 (Begasungsstation)", cat:"bed" },
  "X8":   { t:"Nur offen bei 12-stelligem Betrieb", cat:"bed" },
  "X8A":  { t:"Nur im 12-stelligen Betrieb geöffnet", cat:"bed" },
  "X8B":  { t:"Nur im 12-stelligen Betrieb geöffnet", cat:"bed" },
  "X10":  { t:"Offen wenn Taste 'Produktfilter entlüften' betätigt", cat:"bed" },
  "X26":  { t:"Schritt nur aktiv wenn Formatwechsel angewählt", cat:"schritt" },
  "X27":  { t:"Zu wenn keine Objekte im Begasungsbereich", cat:"bed" },
  "X28":  { t:"Ablauf über WAS — produktabhängig lt. Parametersatz", cat:"bed" },
  "X28N": { t:"Ablauf über WAI — produktabhängig lt. Parametersatz", cat:"bed" },
  "X29":  { t:"Nur offen bei Prog. 7, 20, 21, 23 (inkl. SF-Rack)", cat:"bed" },
  "X30":  { t:"Offen wenn Stickstoff im Parametersatz (Standard!)", cat:"bed" },
  "X30N": { t:"Nur offen wenn Prozessluft im Parametersatz", cat:"bed" },
  "X31":  { t:"Aktiviert mit min./max. Kontrolle Dosierbehälter", cat:"bed" },
  "X32":  { t:"Reserve", cat:"bed" },
  "X49":  { t:"Zu wenn max. Level Dosierbehälter erreicht (2s verzögert)", cat:"bed" },
  "X52":  { t:"Schritt aktiv wenn Merker 'Deko erfolgreich' = 0", cat:"schritt" },
  "X53":  { t:"Zu wenn min./max. Level Dosierbehälter erreicht (2s verzögert)", cat:"bed" },
  "X54":  { t:"Schritt aktiv wenn Filterintegritätstest im Rezept angewählt", cat:"schritt" },
  "X55":  { t:"Schritt nur aktiv wenn Single Use angewählt", cat:"schritt" },
  "X55N": { t:"Schritt nur aktiv wenn Single Use NICHT angewählt", cat:"schritt" },
  "X56":  { t:"Schritt aktiv wenn Filter Flush im Rezept angewählt", cat:"schritt" },

  // ── Taktfolgen (0,5s-Sequenzen für V13.5 / V14.7 / V14.8 / V14.14 / V14.15) ──
  "XA":   { t:"Taktfolge: Folge 5;21", cat:"folge", seq:[5,21] },
  "XB":   { t:"Taktfolge: Folge 11", cat:"folge", seq:[11] },
  "XC":   { t:"Taktfolge: Folge 6", cat:"folge", seq:[6] },
  "XD":   { t:"Taktfolge: Folge 1;3;7;9;13;15;17;19;22;24", cat:"folge", seq:[1,3,7,9,13,15,17,19,22,24] },
  "XE":   { t:"Taktfolge: Folge 12;18", cat:"folge", seq:[12,18] },
  "XF":   { t:"Taktfolge: Folge 2;4;8;10;14;16;20;23", cat:"folge", seq:[2,4,8,10,14,16,20,23] },
  "XG":   { t:"Taktfolge V13.5: Folge 1;2;3;4;14;15;16;17", cat:"folge", valve:"V13.5", seq:[1,2,3,4,14,15,16,17] },
  "XH":   { t:"Taktfolge V14.15: Folge 5;6;7;8;18;19;20;21", cat:"folge", valve:"V14.15", seq:[5,6,7,8,18,19,20,21] },
  "XI":   { t:"Taktfolge V14.8: Folge 9;10;11;12;22;23;24;25", cat:"folge", valve:"V14.8", seq:[9,10,11,12,22,23,24,25] },
  "XJ":   { t:"Taktfolge V14.15: Folge 1;2;3;4;14;15;16;17", cat:"folge", valve:"V14.15", seq:[1,2,3,4,14,15,16,17] },
  "XK":   { t:"Taktfolge V14.8: Folge 7;8;9;10;20;21;22;23", cat:"folge", valve:"V14.8", seq:[7,8,9,10,20,21,22,23] },
  "XL":   { t:"Taktfolge V14.7: Folge 13", cat:"folge", valve:"V14.7", seq:[13] },
  "XM":   { t:"Taktfolge V14.14: Folge 26", cat:"folge", valve:"V14.14", seq:[26] },
  "XN":   { t:"Taktfolge V14.7: Folge 13", cat:"folge", valve:"V14.7", seq:[13] },
  "XP":   { t:"Taktfolge V14.14: Folge 26", cat:"folge", valve:"V14.14", seq:[26] }
};

// ═══════════════════════════════════════════════════════════════
// HELPER: Klartext für beliebigen X-Code abrufen
// ═══════════════════════════════════════════════════════════════
function getTaktText(code) {
  if (!code) return '';
  const c = code.trim();
  if (TAKT[c]) return TAKT[c].t;
  // Fallback: "XA / X8A" Kombi-Notation
  if (c.includes('/')) {
    return c.split('/').map(p => {
      const k = p.trim();
      return TAKT[k] ? `${k}: ${TAKT[k].t}` : k;
    }).join(' + ');
  }
  return c; // unbekannter Code → als String zurückgeben
}

// ═══════════════════════════════════════════════════════════════
// HELPER: Kategorie-Label für UI-Badge
// ═══════════════════════════════════════════════════════════════
function getTaktCat(code) {
  const c = code?.trim();
  if (!c || !TAKT[c]) return 'unbekannt';
  const cats = {
    basis:   '🟢 Offen',
    takt:    '⏱️ Getaktet',
    druck:   '📊 Druckgesteuert',
    bed:     '⚙️ Bedingt',
    schritt: '📋 Schrittsteuerung',
    folge:   '🔄 Taktfolge'
  };
  return cats[TAKT[c].cat] || TAKT[c].cat;
}
