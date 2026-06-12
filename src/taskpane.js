/* PWMedia Word Add-in — taskpane.js — v1.0.0 */

var TEMPLATE_BASE = 'https://scubarious.github.io/word-addin/templates/';

// ── Initialisatie ────────────────────────────────────────────────
Office.onReady(function () {
  laadInstellingen();
});

// ── Tab wisselen ─────────────────────────────────────────────────
function switchTab(naam) {
  var namen = ['sjablonen', 'invoegen', 'beelden', 'instellingen'];
  document.querySelectorAll('.tab').forEach(function (t) { t.classList.remove('active'); });
  document.querySelectorAll('.panel').forEach(function (p) { p.classList.remove('active'); });
  var idx = namen.indexOf(naam);
  if (idx >= 0) {
    document.querySelectorAll('.tab')[idx].classList.add('active');
    document.getElementById('panel-' + naam).classList.add('active');
  }
}

// ── Template openen ──────────────────────────────────────────────
function openTemplate(bestandsnaam) {
  var url = TEMPLATE_BASE + bestandsnaam;
  var a = document.createElement('a');
  a.href = url;
  a.download = bestandsnaam;
  a.target = '_blank';
  a.style.display = 'none';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  showToast('Sjabloon wordt geopend...');
}

// ── Bouwblokken invoegen ─────────────────────────────────────────
async function insertBlock(type) {
  try {
    await Word.run(async function (ctx) {
      var sel = ctx.document.getSelection();
      var datum = new Date().toLocaleDateString('nl-NL');
      if (type === 'versiebeheer') {
        var tbl = sel.insertTable(3, 4, 'After', [
          ['Versie', 'Datum', 'Auteur', 'Wijziging'],
          ['1.0', datum, '[Naam]', 'Initiele versie'],
          ['', '', '', '']
        ]);
        tbl.styleBuiltIn = Word.Style.tableGrid;
      } else if (type === 'tabel2') {
        sel.insertTable(4, 2, 'After', [
          ['Kolom 1', 'Kolom 2'],
          ['', ''], ['', ''], ['', '']
        ]).styleBuiltIn = Word.Style.tableGrid;
      } else if (type === 'tabel3') {
        sel.insertTable(4, 3, 'After', [
          ['Kolom 1', 'Kolom 2', 'Kolom 3'],
          ['', '', ''], ['', '', ''], ['', '', '']
        ]).styleBuiltIn = Word.Style.tableGrid;
      } else if (type === 'kader') {
        sel.insertText('[Tekst in accentkader]', 'Replace');
      }
      await ctx.sync();
    });
    showToast('Ingevoegd');
  } catch (e) { showToast('Fout: ' + e.message, true); }
}

// ── Documentvelden invoegen ──────────────────────────────────────
async function insertField(type) {
  var s = getSettings();
  var waarden = {
    datum:   new Date().toLocaleDateString('nl-NL', { day: 'numeric', month: 'long', year: 'numeric' }),
    auteur:  s.auteur  || '[Auteur]',
    bedrijf: s.bedrijf || '[Bedrijfsnaam]',
    versie:  '1.0',
  };
  try {
    await Word.run(async function (ctx) {
      ctx.document.getSelection().insertText(waarden[type] || type, 'Replace');
      await ctx.sync();
    });
    showToast('Veld ingevoegd');
  } catch (e) { showToast('Fout: ' + e.message, true); }
}

// ── Beeldbibliotheek ─────────────────────────────────────────────
var geselecteerdBeeld = null;

function selectImage(id, el) {
  geselecteerdBeeld = id;
  document.querySelectorAll('.image-thumb').forEach(function (t) { t.classList.remove('selected'); });
  el.classList.add('selected');
  document.getElementById('btn-beeld').disabled = false;
}

function insertSelectedImage() {
  showToast('Koppel "' + geselecteerdBeeld + '" aan echte URL in productie');
}

// ── Instellingen ─────────────────────────────────────────────────
function getSettings() {
  return {
    bedrijf: localStorage.getItem('pwm_bedrijf') || '',
    auteur:  localStorage.getItem('pwm_auteur')  || '',
    functie: localStorage.getItem('pwm_functie') || '',
    kleur:   localStorage.getItem('pwm_kleur')   || '#f18700',
  };
}

function laadInstellingen() {
  var s = getSettings();
  document.getElementById('s-bedrijf').value = s.bedrijf;
  document.getElementById('s-auteur').value  = s.auteur;
  document.getElementById('s-functie').value = s.functie;
  document.getElementById('s-kleur').value   = s.kleur;
  updateKleur(s.kleur);
}

function saveSettings() {
  localStorage.setItem('pwm_bedrijf', document.getElementById('s-bedrijf').value);
  localStorage.setItem('pwm_auteur',  document.getElementById('s-auteur').value);
  localStorage.setItem('pwm_functie', document.getElementById('s-functie').value);
  localStorage.setItem('pwm_kleur',   document.getElementById('s-kleur').value);
  showToast('Instellingen opgeslagen');
}

function resetSettings() {
  ['pwm_bedrijf', 'pwm_auteur', 'pwm_functie', 'pwm_kleur'].forEach(function (k) {
    localStorage.removeItem(k);
  });
  laadInstellingen();
  showToast('Instellingen gereset');
}

function updateKleur(val) {
  var el = document.getElementById('kleur-preview');
  if (el && /^#[0-9A-Fa-f]{6}$/.test(val)) el.style.background = val;
}

// ── Toast ────────────────────────────────────────────────────────
var toastTimer;
function showToast(msg, isError) {
  var t = document.getElementById('toast');
  t.textContent = msg;
  t.className = 'toast show' + (isError ? ' error' : '');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(function () { t.className = 'toast'; }, 2800);
}
