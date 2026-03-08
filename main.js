/**
 * Certification Formation Grist — Logique du widget
 * --------------------------------------------------
 * • Lit un fichier .grist (SQLite via sql.js/WASM)
 * • Vérifie les 23 exercices détectables par le schéma
 * • Détecte l'utilisateur principal depuis l'historique des actions
 * • Calcule le niveau de certification (Débutant / Intermédiaire / Avancé)
 * • Rend les résultats dans l'interface index.html
 */

'use strict';

/* ═══════════════════════════════════════════════════════════════
   1. INITIALISATION GRIST
   ═══════════════════════════════════════════════════════════════ */
if (typeof grist !== 'undefined') {
  grist.ready({ requiredAccess: 'none' });
}

/* ═══════════════════════════════════════════════════════════════
   2. HELPERS BASE DE DONNÉES
   ═══════════════════════════════════════════════════════════════ */

/**
 * Retourne toutes les colonnes d'une table Grist sous forme de tableau
 * d'objets { colId, type, formula, isFormula, recalcWhen, reverseCol, widgetOptions, rules }
 */
function getColumns(db, tableId) {
  try {
    const res = db.exec(`
      SELECT c.colId, c.type, c.formula, c.isFormula,
             c.recalcWhen, c.reverseCol, c.widgetOptions, c.rules
      FROM   _grist_Tables_column c
      JOIN   _grist_Tables t ON c.parentId = t.id
      WHERE  t.tableId = '${tableId}'
    `);
    if (!res.length) return [];
    const keys = res[0].columns;
    return res[0].values.map(row => {
      const obj = {};
      keys.forEach((k, i) => { obj[k] = row[i]; });
      return obj;
    });
  } catch (e) {
    return [];
  }
}

/** Trouve une colonne par son colId exact (insensible à la casse) */
function col(cols, colId) {
  const id = colId.toLowerCase();
  return cols.find(c => c.colId.toLowerCase() === id) || null;
}

/** Vérifie si une table Grist existe dans le document */
function tableExists(db, tableId) {
  try {
    const res = db.exec(
      `SELECT 1 FROM _grist_Tables WHERE tableId = '${tableId}' LIMIT 1`
    );
    return res.length > 0 && res[0].values.length > 0;
  } catch (e) {
    return false;
  }
}

/** Parse les widgetOptions JSON d'une colonne sans planter */
function parseOpts(column) {
  if (!column || !column.widgetOptions) return {};
  try { return JSON.parse(column.widgetOptions); } catch (e) { return {}; }
}

/* ═══════════════════════════════════════════════════════════════
   3. DÉFINITION DES 23 EXERCICES VÉRIFIABLES
   ═══════════════════════════════════════════════════════════════ */

const SEQUENCES = {
  1: { label: 'Séquence 1', subtitle: 'Bibliothèque — bases de Grist' },
  2: { label: 'Séquence 2', subtitle: 'Typage de DONNEES_EXCEL_REUNION' },
  3: { label: 'Séquence 3', subtitle: 'Vues avancées' },
  4: { label: 'Séquence 4', subtitle: 'Références avancées' },
  6: { label: 'Séquence 6', subtitle: 'Formules avancées & qualité des données' },
};

const EXERCISES = [

  /* ── SEQ 1 ─────────────────────────────────────────────── */
  {
    id: '1-2', seq: 1, niveau: 'D',
    nom: 'Comprendre les colonnes',
    detail: 'Annee_de_publication→Date, id_categorie→Int + colonne Bool + colonne Choice dans Donnees_Livres',
    check(db) {
      const cols = getColumns(db, 'Donnees_Livres');
      const annee  = col(cols, 'Annee_de_publication');
      const idCat  = col(cols, 'id_categorie');
      const hasBool   = cols.some(c => c.type === 'Bool');
      const hasChoice = cols.some(c => c.type === 'Choice');
      return annee?.type === 'Date'
          && idCat?.type  === 'Int'
          && hasBool
          && hasChoice;
    }
  },

  {
    id: '1-3', seq: 1, niveau: 'D',
    nom: 'Colonne de référence 1/2',
    detail: 'Colonne Auteur dans Donnees_Livres → type Ref:Donnees_Auteurs',
    check(db) {
      const cols   = getColumns(db, 'Donnees_Livres');
      const auteur = col(cols, 'Auteur');
      return !!auteur?.type?.startsWith('Ref:');
    }
  },

  {
    id: '1-4', seq: 1, niveau: 'D',
    nom: 'Colonne de référence 2/2',
    detail: 'Champs rapportés $Auteur.Date_de_naissance et $Auteur.Nationalite dans Donnees_Livres',
    check(db) {
      const cols = getColumns(db, 'Donnees_Livres');
      const lookups = cols.filter(c =>
        c.isFormula === 1 &&
        typeof c.formula === 'string' &&
        c.formula.includes('$Auteur.')
      );
      return lookups.length >= 2;
    }
  },

  {
    id: '1-9', seq: 1, niveau: 'I',
    nom: 'Formules simples',
    detail: 'Colonne formule supplémentaire (LEFT / concat) dans Donnees_Auteurs, au-delà de Nom_complet',
    check(db) {
      const cols = getColumns(db, 'Donnees_Auteurs');
      return cols.some(c =>
        c.isFormula === 1 &&
        c.colId !== 'Nom_complet' &&
        typeof c.formula === 'string' &&
        c.formula.length > 0
      );
    }
  },

  /* ── SEQ 2 ─────────────────────────────────────────────── */
  {
    id: '2-1', seq: 2, niveau: 'D',
    nom: 'Colonne Texte',
    detail: 'Nom_Organisateur et Contact_Telephone → type Text dans DONNEES_EXCEL_REUNION',
    check(db) {
      const cols = getColumns(db, 'Donnees_Excel_Reunion');
      const nom = col(cols, 'Nom_Organisateur');
      const tel = col(cols, 'Contact_Telephone');
      return nom?.type === 'Text' && tel?.type === 'Text';
    }
  },

  {
    id: '2-2', seq: 2, niveau: 'D',
    nom: 'Colonne Entier / Numérique',
    detail: 'Nombre_Participants→Int, Budget_Alloue→Numeric avec format monétaire (numMode:currency)',
    check(db) {
      const cols   = getColumns(db, 'Donnees_Excel_Reunion');
      const nbPart = col(cols, 'Nombre_Participants');
      const budget = col(cols, 'Budget_Alloue');
      if (nbPart?.type !== 'Int')     return false;
      if (budget?.type !== 'Numeric') return false;
      const opts = parseOpts(budget);
      return opts.numMode === 'currency';
    }
  },

  {
    id: '2-3', seq: 2, niveau: 'D',
    nom: 'Colonne Date / Heure',
    detail: 'Date_Reunion→Date, Horaire_Debut et Horaire_Fin→DateTime dans DONNEES_EXCEL_REUNION',
    check(db) {
      const cols  = getColumns(db, 'Donnees_Excel_Reunion');
      const date  = col(cols, 'Date_Reunion');
      const debut = col(cols, 'Horaire_Debut');
      const fin   = col(cols, 'Horaire_Fin');
      return date?.type  === 'Date'
          && debut?.type?.startsWith('DateTime')
          && fin?.type?.startsWith('DateTime');
    }
  },

  {
    id: '2-4', seq: 2, niveau: 'D',
    nom: 'Colonne Choix',
    detail: 'Lieu_Reunion→Choice (≥4 valeurs), Themes_Abordes→ChoiceList dans DONNEES_EXCEL_REUNION',
    check(db) {
      const cols   = getColumns(db, 'Donnees_Excel_Reunion');
      const lieu   = col(cols, 'Lieu_Reunion');
      const themes = col(cols, 'Themes_Abordes');
      if (lieu?.type   !== 'Choice')     return false;
      if (themes?.type !== 'ChoiceList') return false;
      const opts = parseOpts(lieu);
      return Array.isArray(opts.choices) && opts.choices.length >= 4;
    }
  },

  {
    id: '2-5', seq: 2, niveau: 'D',
    nom: 'Colonne Booléenne',
    detail: 'Repas→Bool avec widget Switch dans DONNEES_EXCEL_REUNION',
    check(db) {
      const cols  = getColumns(db, 'Donnees_Excel_Reunion');
      const repas = col(cols, 'Repas');
      if (repas?.type !== 'Bool') return false;
      const opts = parseOpts(repas);
      return opts.widget === 'Switch';
    }
  },

  {
    id: '2-6', seq: 2, niveau: 'D',
    nom: 'Colonne Pièce Jointe',
    detail: 'Compte_Rendu→type Attachments dans DONNEES_EXCEL_REUNION',
    check(db) {
      const cols = getColumns(db, 'Donnees_Excel_Reunion');
      const cr   = col(cols, 'Compte_Rendu');
      return cr?.type === 'Attachments';
    }
  },

  /* ── SEQ 3 ─────────────────────────────────────────────── */
  {
    id: '3-4', seq: 3, niveau: 'I',
    nom: 'Vue Géocodeur',
    detail: 'Colonnes Latitude et Longitude créées dans Donnees_Adherents',
    check(db) {
      const cols = getColumns(db, 'Donnees_Adherents');
      const lat = cols.find(c =>
        c.colId.toLowerCase().includes('latitude') ||
        c.colId.toLowerCase() === 'lat'
      );
      const lon = cols.find(c =>
        c.colId.toLowerCase().includes('longitude') ||
        c.colId.toLowerCase() === 'lon' ||
        c.colId.toLowerCase() === 'lng'
      );
      return !!(lat && lon);
    }
  },

  /* ── SEQ 4 ─────────────────────────────────────────────── */
  {
    id: '4-1', seq: 4, niveau: 'D',
    nom: 'Comprendre les références',
    detail: 'Colonne Emprunteur dans Donnees_Emprunts → type Ref:Donnees_Adherents',
    check(db) {
      const cols       = getColumns(db, 'Donnees_Emprunts');
      const emprunteur = col(cols, 'Emprunteur');
      return !!emprunteur?.type?.startsWith('Ref:');
    }
  },

  {
    id: '4-2', seq: 4, niveau: 'D',
    nom: 'Champs Rapportés',
    detail: 'Colonne formule $Emprunteur.<champ> dans Donnees_Emprunts (champ rapporté)',
    check(db) {
      const cols = getColumns(db, 'Donnees_Emprunts');
      return cols.some(c =>
        c.isFormula === 1 &&
        typeof c.formula === 'string' &&
        c.formula.includes('$Emprunteur.')
      );
    }
  },

  {
    id: '4-5', seq: 4, niveau: 'I',
    nom: 'Références Bi-Directionnelles',
    detail: 'reverseCol != 0 sur Emprunteur dans Donnees_Emprunts',
    check(db) {
      const cols       = getColumns(db, 'Donnees_Emprunts');
      const emprunteur = col(cols, 'Emprunteur');
      return !!(emprunteur?.reverseCol && emprunteur.reverseCol !== 0);
    }
  },

  {
    id: '4-6', seq: 4, niveau: 'I',
    nom: 'Références Multiples',
    detail: 'Colonne Livres_Preferes dans Donnees_Adherents → type RefList:Donnees_Livres',
    check(db) {
      const cols = getColumns(db, 'Donnees_Adherents');
      return cols.some(c => c.type?.startsWith('RefList:'));
    }
  },

  {
    id: '4-7', seq: 4, niveau: 'A',
    nom: 'LookupOne',
    detail: 'Table Calculs existante + colonne formule contenant lookupOne()',
    check(db) {
      if (!tableExists(db, 'Calculs')) return false;
      const cols = getColumns(db, 'Calculs');
      return cols.some(c =>
        c.isFormula === 1 &&
        typeof c.formula === 'string' &&
        c.formula.includes('lookupOne(')
      );
    }
  },

  {
    id: '4-8', seq: 4, niveau: 'A',
    nom: 'LookupRecords',
    detail: 'Table Calculs existante + colonne formule contenant lookupRecords()',
    check(db) {
      if (!tableExists(db, 'Calculs')) return false;
      const cols = getColumns(db, 'Calculs');
      return cols.some(c =>
        c.isFormula === 1 &&
        typeof c.formula === 'string' &&
        c.formula.includes('lookupRecords(')
      );
    }
  },

  /* ── SEQ 6 ─────────────────────────────────────────────── */
  {
    id: '6-2', seq: 6, niveau: 'A',
    nom: 'ID & UUID',
    detail: 'Numero_De_Reservation formula=$id  ET/OU  Numero_De_Dossier formula=UUID() recalcWhen=1',
    check(db) {
      const cols   = getColumns(db, 'Donnees_Reservations');
      const numRes = col(cols, 'Numero_De_Reservation');
      const numDos = col(cols, 'Numero_De_Dossier');
      const idOk   = numRes?.isFormula === 1 && numRes?.formula === '$id';
      const uuidOk = typeof numDos?.formula === 'string' &&
                     numDos.formula.toUpperCase().includes('UUID') &&
                     numDos.recalcWhen === 1;
      return !!(idOk || uuidOk);
    }
  },

  {
    id: '6-3', seq: 6, niveau: 'A',
    nom: 'Label Utilisateur',
    detail: 'Agent_Responsable formula=user.Name (ou user.Email) avec recalcWhen != 0',
    check(db) {
      const cols  = getColumns(db, 'Donnees_Reservations');
      const agent = col(cols, 'Agent_Responsable');
      return !!(
        typeof agent?.formula === 'string' &&
        agent.formula.includes('user.') &&
        agent.recalcWhen !== 0
      );
    }
  },

  {
    id: '6-4', seq: 6, niveau: 'A',
    nom: 'Label Temporel',
    detail: 'Date_Reservation formula=NOW() avec recalcWhen != 0 dans Donnees_Reservations',
    check(db) {
      const cols    = getColumns(db, 'Donnees_Reservations');
      const dateRes = col(cols, 'Date_Reservation');
      return !!(
        typeof dateRes?.formula === 'string' &&
        dateRes.formula.includes('NOW()') &&
        dateRes.recalcWhen !== 0
      );
    }
  },

  {
    id: '6-5', seq: 6, niveau: 'A',
    nom: 'Doublons',
    detail: 'Colonne formule avec lookupRecords() + len() dans Donnees_Reservations',
    check(db) {
      const cols = getColumns(db, 'Donnees_Reservations');
      return cols.some(c =>
        c.isFormula === 1 &&
        typeof c.formula === 'string' &&
        c.formula.includes('lookupRecords(') &&
        c.formula.includes('len(')
      );
    }
  },

  {
    id: '6-6', seq: 6, niveau: 'A',
    nom: 'Valeurs Vides',
    detail: 'Colonne Bool avec formule bool(...) dans Donnees_Reservations',
    check(db) {
      const cols = getColumns(db, 'Donnees_Reservations');
      return cols.some(c =>
        c.type === 'Bool' &&
        c.isFormula === 1 &&
        typeof c.formula === 'string' &&
        c.formula.includes('bool(')
      );
    }
  },

  {
    id: '6-7', seq: 6, niveau: 'A',
    nom: 'Valeurs Aberrantes',
    detail: 'Colonne formule Python IQR (statistics / median / outlier) dans Donnees_Reservations',
    check(db) {
      const cols = getColumns(db, 'Donnees_Reservations');
      return cols.some(c =>
        c.isFormula === 1 &&
        typeof c.formula === 'string' &&
        (c.formula.includes('statistics') ||
         c.formula.includes('median')     ||
         c.formula.includes('outlier')    ||
         c.formula.includes('iqr'))
      );
    }
  },

];

/* ═══════════════════════════════════════════════════════════════
   4. DÉTECTION DE L'UTILISATEUR PRINCIPAL (anti-fraude)
   ═══════════════════════════════════════════════════════════════ */

/**
 * Analyse _gristsys_ActionHistory pour identifier l'utilisateur ayant
 * le plus d'actions dans le document (hors comptes système).
 * Retourne { name, email, actionCount } ou null si indétectable.
 */
function extractTopUser(db) {
  try {
    // Lit les 500 premières entrées de l'historique (body = blob)
    const res = db.exec(
      'SELECT body FROM _gristsys_ActionHistory ORDER BY id DESC LIMIT 500'
    );
    if (!res.length || !res[0].values.length) return null;

    const emailCounts = {};
    const emailRegex  = /[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g;

    // Motifs d'emails système à ignorer
    const SYSTEM_PATTERNS = [
      /^grist(docs|app|bot)?@/i,
      /noreply/i,
      /no-reply/i,
      /support@/i,
      /system@/i,
      /admin@getgrist\.com/i,
    ];

    const decoder = new TextDecoder('utf-8', { fatal: false });

    res[0].values.forEach(([body]) => {
      if (!body) return;

      // Le body peut être un Uint8Array (blob SQLite) ou une chaîne
      let text;
      if (body instanceof Uint8Array || ArrayBuffer.isView(body)) {
        text = decoder.decode(body);
      } else if (typeof body === 'string') {
        text = body;
      } else {
        return;
      }

      // Extrait tous les emails uniques de cet enregistrement
      const matches = text.match(emailRegex) || [];
      const seenInRow = new Set();

      matches.forEach(raw => {
        const email = raw.toLowerCase();
        if (seenInRow.has(email)) return; // un seul incrément par ligne d'historique
        seenInRow.add(email);

        // Filtre les emails système
        if (SYSTEM_PATTERNS.some(p => p.test(email))) return;

        emailCounts[email] = (emailCounts[email] || 0) + 1;
      });
    });

    const entries = Object.entries(emailCounts);
    if (!entries.length) return null;

    // Email le plus fréquent = utilisateur principal
    const [email, actionCount] = entries.sort((a, b) => b[1] - a[1])[0];

    // Tente de récupérer le nom affiché depuis _grist_ACLPrincipals
    let name = null;
    try {
      const safe  = email.replace(/'/g, "''");
      const pRes  = db.exec(
        `SELECT name FROM _grist_ACLPrincipals WHERE email = '${safe}' LIMIT 1`
      );
      if (pRes.length && pRes[0].values.length) {
        name = pRes[0].values[0][0] || null;
      }
    } catch (_) { /* table absente ou vide */ }

    // Fallback : partie locale de l'email
    if (!name) name = email.split('@')[0].replace(/[._\-]/g, ' ');

    return { name, email, actionCount };

  } catch (e) {
    return null;
  }
}

/* ═══════════════════════════════════════════════════════════════
   5. CALCUL DU NIVEAU DE CERTIFICATION
   ═══════════════════════════════════════════════════════════════ */

/** Heures de formation équivalentes par niveau */
const LEVEL_HOURS = { D: 7, I: 14, A: 21 };

/**
 * Seuils :
 *   Débutant      → ≥ 5 exercices D validés
 *   Intermédiaire → ≥ 5 D + ≥ 2 I
 *   Avancé        → ≥ 5 D + ≥ 2 I + ≥ 2 A
 * Retourne 'none' | 'D' | 'I' | 'A'
 */
function getCertificationLevel(results) {
  const doneD = results.filter(r => r.niveau === 'D' && r.done).length;
  const doneI = results.filter(r => r.niveau === 'I' && r.done).length;
  const doneA = results.filter(r => r.niveau === 'A' && r.done).length;

  if (doneD >= 5 && doneI >= 2 && doneA >= 2) return 'A';
  if (doneD >= 5 && doneI >= 2)               return 'I';
  if (doneD >= 5)                             return 'D';
  return 'none';
}

/* ═══════════════════════════════════════════════════════════════
   6. ANALYSE DU FICHIER
   ═══════════════════════════════════════════════════════════════ */

async function analyzeGristFile(arrayBuffer) {
  // Charge sql.js (WASM depuis CDN)
  const SQL = await initSqlJs({
    locateFile: f => `https://cdn.jsdelivr.net/npm/sql.js@1.10.2/dist/${f}`
  });

  const db = new SQL.Database(new Uint8Array(arrayBuffer));

  // Vérifie que c'est bien un fichier Grist
  try {
    db.exec('SELECT 1 FROM _grist_Tables LIMIT 1');
  } catch (e) {
    db.close();
    throw new Error(
      'Ce fichier ne semble pas être un document Grist valide.\n' +
      'Assurez-vous de déposer un fichier .grist issu de votre formation.'
    );
  }

  // Lance les 23 vérifications d'exercices
  const results = EXERCISES.map(ex => {
    let done = false;
    try { done = !!ex.check(db); } catch (_) { done = false; }
    return { ...ex, done };
  });

  // Détecte l'utilisateur principal (peut échouer sans bloquer)
  const userInfo = extractTopUser(db);

  db.close();
  return { results, userInfo };
}

/* ═══════════════════════════════════════════════════════════════
   7. RENDU UI
   ═══════════════════════════════════════════════════════════════ */

/* Libellés lisibles */
const LEVEL_META = {
  none: {
    medal:  '📋',
    label:  'Aucune certification',
    title:  'Formation en cours…',
    next:   'Validez au moins 5 exercices Débutant pour obtenir votre première certification.',
    color:  'var(--grey)',
  },
  D: {
    medal:  '🥉',
    label:  'Débutant',
    title:  'Certification Débutant',
    next:   'Validez 2 exercices Intermédiaire pour passer au niveau supérieur.',
    color:  'var(--blue)',
  },
  I: {
    medal:  '🥈',
    label:  'Intermédiaire',
    title:  'Certification Intermédiaire',
    next:   'Validez 2 exercices Avancé pour atteindre le niveau Avancé.',
    color:  'var(--orange)',
  },
  A: {
    medal:  '🥇',
    label:  'Avancé',
    title:  'Certification Avancé',
    next:   '🎉 Félicitations ! Vous avez complété toutes les certifications.',
    color:  'var(--purple)',
  },
};

const LEVEL_TOTALS = { D: 11, I: 4, A: 8 }; // exercices disponibles par niveau

function niveauLabel(code) {
  return { D: 'Débutant', I: 'Intermédiaire', A: 'Avancé' }[code] || code;
}

function renderResults(results, userInfo) {
  const level = getCertificationLevel(results);
  const meta  = LEVEL_META[level];

  /* ── Carte de certification ───────────────────────────── */
  const card = document.getElementById('cert-card');
  card.className = `level-${level}`;

  document.getElementById('cert-medal').textContent       = meta.medal;
  document.getElementById('cert-level-label').textContent = meta.label;
  document.getElementById('cert-title').textContent       = meta.title;
  document.getElementById('cert-next').textContent        = meta.next;

  // Bloc utilisateur
  const userBlock = document.getElementById('cert-user-block');
  if (userInfo) {
    userBlock.innerHTML = `
      <p class="cert-user">
        Réalisé par <span>${escHtml(userInfo.name)}</span>
        <span class="email-badge">${escHtml(userInfo.email)}</span>
      </p>`;
  } else {
    userBlock.innerHTML =
      '<p class="cert-no-user">Identité non détectée (historique absent ou vide)</p>';
  }

  /* ── Barres de progression par niveau ─────────────────── */
  ['D', 'I', 'A'].forEach(lvl => {
    const done  = results.filter(r => r.niveau === lvl && r.done).length;
    const total = LEVEL_TOTALS[lvl];
    const pct   = total > 0 ? Math.round((done / total) * 100) : 0;

    document.getElementById(`fill-${lvl}`).style.width   = pct + '%';
    document.getElementById(`count-${lvl}`).textContent  = `${done} / ${total}`;
  });

  /* ── Bouton téléchargement certificat ─────────────────── */
  const dlBtn = $('btn-download-cert');
  dlBtn.disabled = (level === 'none');
  // Retire les anciens listeners puis en rattache un nouveau
  const dlBtnClone = dlBtn.cloneNode(true);
  dlBtn.parentNode.replaceChild(dlBtnClone, dlBtn);
  dlBtnClone.disabled = (level === 'none');
  dlBtnClone.addEventListener('click', () => generateAndDownloadCert(results, userInfo));

  /* ── Liste des exercices par séquence ─────────────────── */
  const container = document.getElementById('sequences');
  container.innerHTML = '';

  // Regroupe par séquence
  const bySeq = {};
  results.forEach(r => {
    if (!bySeq[r.seq]) bySeq[r.seq] = [];
    bySeq[r.seq].push(r);
  });

  Object.keys(bySeq).sort((a, b) => Number(a) - Number(b)).forEach(seqKey => {
    const seqResults = bySeq[seqKey];
    const seqDone    = seqResults.filter(r => r.done).length;
    const seqTotal   = seqResults.length;
    const seqInfo    = SEQUENCES[seqKey] || { label: `Séquence ${seqKey}` };

    const group = document.createElement('div');
    group.className = 'seq-group';

    group.innerHTML = `
      <div class="seq-header">
        <div class="seq-left">
          <span class="seq-badge">SEQ ${seqKey}</span>
          <span class="seq-title">${escHtml(seqInfo.label)}</span>
        </div>
        <div class="seq-right">
          <span class="seq-count">${seqDone}&thinsp;/&thinsp;${seqTotal}</span>
          <span class="seq-chevron">▼</span>
        </div>
      </div>
      <div class="seq-body">
        ${seqResults.map(r => `
          <div class="ex-row">
            <div class="ex-icon">${r.done ? '✅' : '❌'}</div>
            <div>
              <div class="ex-name ${r.done ? 'done' : 'miss'}">${escHtml(r.nom)}</div>
              <div class="ex-detail">${escHtml(r.detail)}</div>
            </div>
            <span class="ex-id">${r.id}</span>
            <span class="niv-badge ${r.niveau}">${niveauLabel(r.niveau)}</span>
          </div>
        `).join('')}
      </div>
    `;

    // Collapse / expand au clic sur l'en-tête
    group.querySelector('.seq-header').addEventListener('click', () => {
      group.classList.toggle('collapsed');
    });

    container.appendChild(group);
  });
}

/** Échappe les caractères HTML (XSS-safe) */
function escHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g,  '&amp;')
    .replace(/</g,  '&lt;')
    .replace(/>/g,  '&gt;')
    .replace(/"/g,  '&quot;');
}

/* ═══════════════════════════════════════════════════════════════
   8. GÉNÉRATION PDF DU CERTIFICAT
   ═══════════════════════════════════════════════════════════════ */

/**
 * Coordonnées de placement du texte dans Template_Certificat_Grist_MOOC - 3.pdf
 * A4 = 595 × 842 pts  |  Origine (0,0) en bas-à-gauche dans pdf-lib
 * Ajuster Y si le texte ne tombe pas dans la bonne zone.
 */
const PDF_ZONES = {
  name:  { y: 526, size: 18 },      // Zone blanche entre les deux filets rouges
  level: { y: 414, size: 13 },      // Sous "Niveau de formation accordé :"
  hours: {                           // Cellule droite du tableau Équivalent
    rectX: 222, rectY: 394, rectW: 330, rectH: 22,
    textY: 402, size: 11,
  },
  date:  { y: 335, size: 11 },      // Dans le cadre "Délivré le"
};

/**
 * Génère le certificat PDF rempli et le télécharge.
 * Utilise pdf-lib (chargé depuis CDN dans index.html).
 */
async function generateAndDownloadCert(results, userInfo) {
  const level = getCertificationLevel(results);
  if (level === 'none') return;

  const btn = $('btn-download-cert');
  btn.disabled = true;
  btn.innerHTML = `
    <svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round">
      <circle cx="8" cy="8" r="6" stroke-dasharray="30" stroke-dashoffset="10">
        <animateTransform attributeName="transform" type="rotate" from="0 8 8" to="360 8 8" dur=".7s" repeatCount="indefinite"/>
      </circle>
    </svg>
    Génération en cours…`;

  try {
    const { PDFDocument, StandardFonts, rgb } = PDFLib;

    // Charge le template PDF depuis le même répertoire que la page
    const templateUrl = './Template_Certificat_Grist_MOOC - 3.pdf';
    const templateBytes = await fetch(templateUrl).then(r => {
      if (!r.ok) throw new Error(`Template introuvable (${r.status})`);
      return r.arrayBuffer();
    });

    const pdfDoc = await PDFDocument.load(templateBytes);
    const page   = pdfDoc.getPages()[0];
    const { width } = page.getSize();

    const fontReg  = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const fontBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

    const RED   = rgb(0.78, 0.06, 0.06);
    const DARK  = rgb(0.12, 0.12, 0.12);
    const WHITE = rgb(1,    1,    1);

    /** Centre le texte horizontalement sur la page */
    function drawCentered(text, font, size, y, color = DARK) {
      const tw = font.widthOfTextAtSize(text, size);
      page.drawText(text, { x: (width - tw) / 2, y, size, font, color });
    }

    // 1. Nom du participant
    const name = userInfo ? userInfo.name : 'Participant(e)';
    drawCentered(name, fontBold, PDF_ZONES.name.size, PDF_ZONES.name.y);

    // 2. Niveau de certification (rouge gras, centré)
    const levelLabels = { D: 'Débutant', I: 'Intermédiaire', A: 'Avancé' };
    drawCentered(levelLabels[level], fontBold, PDF_ZONES.level.size, PDF_ZONES.level.y, RED);

    // 3. Heures de formation — couvre le texte existant avec un rect blanc puis réécrit
    const z = PDF_ZONES.hours;
    page.drawRectangle({ x: z.rectX, y: z.rectY, width: z.rectW, height: z.rectH, color: WHITE });
    const hoursText = `${LEVEL_HOURS[level]} heures de formation`;
    const tw = fontBold.widthOfTextAtSize(hoursText, z.size);
    page.drawText(hoursText, {
      x: z.rectX + (z.rectW - tw) / 2,
      y: z.textY,
      size: z.size,
      font: fontBold,
      color: RED,
    });

    // 4. Date de délivrance
    const dateStr = new Date().toLocaleDateString('fr-FR', {
      day: 'numeric', month: 'long', year: 'numeric',
    });
    drawCentered(dateStr, fontReg, PDF_ZONES.date.size, PDF_ZONES.date.y);

    // Sauvegarde + téléchargement
    const pdfBytes = await pdfDoc.save();
    const blob     = new Blob([pdfBytes], { type: 'application/pdf' });
    const url      = URL.createObjectURL(blob);
    const safeName = name.replace(/[^a-zA-Z0-9\u00C0-\u017E _\-]/g, '_');
    const a        = Object.assign(document.createElement('a'), {
      href:     url,
      download: `Certificat_GristMOOC_${safeName}_${levelLabels[level]}.pdf`,
    });
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 8000);

  } catch (err) {
    alert('Erreur lors de la génération du certificat :\n' + err.message);
  } finally {
    btn.disabled = false;
    btn.innerHTML = `
      <svg width="15" height="15" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M8 2v9M5 8l3 3 3-3M2 13h12"/>
      </svg>
      Télécharger mon certificat PDF`;
  }
}

/* ═══════════════════════════════════════════════════════════════
   9. GESTION DES ÉTATS UI
   ═══════════════════════════════════════════════════════════════ */

const $ = id => document.getElementById(id);

function showState(state) {
  ['drop-zone', 'loading', 'error', 'results'].forEach(id => {
    $(id).classList.add('hidden');
  });
  $(state).classList.remove('hidden');
}

function showError(msg) {
  $('error-msg').textContent = msg;
  showState('error');
}

/* ═══════════════════════════════════════════════════════════════
   10. GESTION DU FICHIER (drag & drop + clic)
   ═══════════════════════════════════════════════════════════════ */

async function handleFile(file) {
  if (!file) return;

  if (!file.name.toLowerCase().endsWith('.grist')) {
    showError('Format invalide : veuillez déposer un fichier .grist');
    return;
  }

  showState('loading');

  try {
    const buffer             = await file.arrayBuffer();
    const { results, userInfo } = await analyzeGristFile(buffer);
    renderResults(results, userInfo);
    showState('results');
  } catch (err) {
    showError(err.message || "Une erreur est survenue lors de l'analyse.");
  }
}

/* ── Drag & drop ──────────────────────────────────────────────── */
const dropZone  = $('drop-zone');
const fileInput = $('file-input');

['dragenter', 'dragover', 'dragleave', 'drop'].forEach(ev => {
  dropZone.addEventListener(ev, e => { e.preventDefault(); e.stopPropagation(); });
});

dropZone.addEventListener('dragenter', () => dropZone.classList.add('drag-over'));
dropZone.addEventListener('dragover',  () => dropZone.classList.add('drag-over'));
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));

dropZone.addEventListener('drop', e => {
  dropZone.classList.remove('drag-over');
  handleFile(e.dataTransfer.files[0]);
});

/* ── Clic pour parcourir ──────────────────────────────────────── */
dropZone.addEventListener('click', () => fileInput.click());

fileInput.addEventListener('change', () => {
  if (fileInput.files.length > 0) {
    handleFile(fileInput.files[0]);
    fileInput.value = ''; // reset pour permettre de recharger le même fichier
  }
});

/* ── Bouton reset ─────────────────────────────────────────────── */
$('btn-reset').addEventListener('click', () => {
  showState('drop-zone');
});
