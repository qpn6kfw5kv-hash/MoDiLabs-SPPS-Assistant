// ============================================================
// SPPS Lab Assistant — Amino Acid Database
// Complete Fmoc-amino acid data for SPPS calculations
// ============================================================

const AMINO_ACIDS = {
  G: {
    code1: 'G', code3: 'Gly', name: 'Glycine',
    fmocName: 'Fmoc-Gly-OH', fmocMW: 297.31, residueMW: 57.05,
    protection: null, scavengerNote: null
  },
  A: {
    code1: 'A', code3: 'Ala', name: 'Alanine',
    fmocName: 'Fmoc-Ala-OH', fmocMW: 311.34, residueMW: 71.08,
    protection: null, scavengerNote: null
  },
  V: {
    code1: 'V', code3: 'Val', name: 'Valine',
    fmocName: 'Fmoc-Val-OH', fmocMW: 339.39, residueMW: 99.13,
    protection: null, scavengerNote: null
  },
  L: {
    code1: 'L', code3: 'Leu', name: 'Leucine',
    fmocName: 'Fmoc-Leu-OH', fmocMW: 353.41, residueMW: 113.16,
    protection: null, scavengerNote: null
  },
  I: {
    code1: 'I', code3: 'Ile', name: 'Isoleucine',
    fmocName: 'Fmoc-Ile-OH', fmocMW: 353.41, residueMW: 113.16,
    protection: null, scavengerNote: null
  },
  P: {
    code1: 'P', code3: 'Pro', name: 'Proline',
    fmocName: 'Fmoc-Pro-OH', fmocMW: 337.37, residueMW: 97.12,
    protection: null, scavengerNote: null
  },
  F: {
    code1: 'F', code3: 'Phe', name: 'Phenylalanine',
    fmocName: 'Fmoc-Phe-OH', fmocMW: 387.43, residueMW: 147.18,
    protection: null, scavengerNote: null
  },
  W: {
    code1: 'W', code3: 'Trp', name: 'Tryptophan',
    fmocName: 'Fmoc-Trp(Boc)-OH', fmocMW: 526.58, residueMW: 186.21,
    protection: 'Boc', scavengerNote: 'Aggiungere TIS come scavenger per prevenire alchilazione del Trp'
  },
  M: {
    code1: 'M', code3: 'Met', name: 'Methionine',
    fmocName: 'Fmoc-Met-OH', fmocMW: 371.45, residueMW: 131.20,
    protection: null, scavengerNote: 'Aggiungere EDT o tioanisolo per prevenire ossidazione della Met'
  },
  S: {
    code1: 'S', code3: 'Ser', name: 'Serine',
    fmocName: 'Fmoc-Ser(tBu)-OH', fmocMW: 383.44, residueMW: 87.08,
    protection: 'tBu', scavengerNote: null
  },
  T: {
    code1: 'T', code3: 'Thr', name: 'Threonine',
    fmocName: 'Fmoc-Thr(tBu)-OH', fmocMW: 397.47, residueMW: 101.10,
    protection: 'tBu', scavengerNote: null
  },
  C: {
    code1: 'C', code3: 'Cys', name: 'Cysteine',
    fmocName: 'Fmoc-Cys(Trt)-OH', fmocMW: 585.71, residueMW: 103.14,
    protection: 'Trt', scavengerNote: 'Aggiungere EDT nel cocktail di cleavage per prevenire rialchilazione'
  },
  Y: {
    code1: 'Y', code3: 'Tyr', name: 'Tyrosine',
    fmocName: 'Fmoc-Tyr(tBu)-OH', fmocMW: 459.54, residueMW: 163.18,
    protection: 'tBu', scavengerNote: null
  },
  H: {
    code1: 'H', code3: 'His', name: 'Histidine',
    fmocName: 'Fmoc-His(Trt)-OH', fmocMW: 619.74, residueMW: 137.14,
    protection: 'Trt', scavengerNote: null
  },
  D: {
    code1: 'D', code3: 'Asp', name: 'Aspartic Acid',
    fmocName: 'Fmoc-Asp(OtBu)-OH', fmocMW: 411.45, residueMW: 115.09,
    protection: 'OtBu', scavengerNote: null
  },
  E: {
    code1: 'E', code3: 'Glu', name: 'Glutamic Acid',
    fmocName: 'Fmoc-Glu(OtBu)-OH', fmocMW: 425.48, residueMW: 129.12,
    protection: 'OtBu', scavengerNote: null
  },
  N: {
    code1: 'N', code3: 'Asn', name: 'Asparagine',
    fmocName: 'Fmoc-Asn(Trt)-OH', fmocMW: 596.68, residueMW: 114.10,
    protection: 'Trt', scavengerNote: null
  },
  Q: {
    code1: 'Q', code3: 'Gln', name: 'Glutamine',
    fmocName: 'Fmoc-Gln(Trt)-OH', fmocMW: 610.71, residueMW: 128.13,
    protection: 'Trt', scavengerNote: null
  },
  K: {
    code1: 'K', code3: 'Lys', name: 'Lysine',
    fmocName: 'Fmoc-Lys(Boc)-OH', fmocMW: 468.55, residueMW: 128.17,
    protection: 'Boc', scavengerNote: null
  },
  R: {
    code1: 'R', code3: 'Arg', name: 'Arginine',
    fmocName: 'Fmoc-Arg(Pbf)-OH', fmocMW: 648.77, residueMW: 156.19,
    protection: 'Pbf', scavengerNote: 'Richiede tempo di cleavage più lungo (3-4h) per rimozione completa del Pbf'
  }
};

// ============================================================
// Resins
// ============================================================
const RESINS = [
  { name: 'Fmoc-Rink Amide', type: 'amide', defaultLoading: 0.60, range: '0.4–0.8 mmol/g' },
  { name: 'Rink Amide', type: 'amide', defaultLoading: 0.65, range: '0.3–0.8 mmol/g' },
  { name: 'Rink Amide MBHA', type: 'amide', defaultLoading: 0.52, range: '0.3–0.7 mmol/g' },
  { name: 'Rink Amide ChemMatrix', type: 'amide', defaultLoading: 0.50, range: '0.4–0.6 mmol/g' },
  { name: 'Wang', type: 'acid', defaultLoading: 0.90, range: '0.5–1.2 mmol/g' },
  { name: '2-Chlorotrityl (2-CTC)', type: 'acid', defaultLoading: 1.20, range: '0.8–1.6 mmol/g' },
  { name: 'Sieber Amide', type: 'amide', defaultLoading: 0.70, range: '0.5–0.8 mmol/g' },
];

// ============================================================
// Activators
// ============================================================
const ACTIVATORS = [
  { name: 'HBTU', mw: 379.24, defaultEq: 2.9, base: 'DIPEA', baseMW: 129.24, baseEq: 6 },
  { name: 'HATU', mw: 380.23, defaultEq: 2.9, base: 'DIPEA', baseMW: 129.24, baseEq: 6 },
  { name: 'PyBOP', mw: 520.43, defaultEq: 2.9, base: 'DIPEA', baseMW: 129.24, baseEq: 6 },
  { name: 'HCTU', mw: 413.69, defaultEq: 2.9, base: 'DIPEA', baseMW: 129.24, baseEq: 6 },
  { name: 'DIC/Oxyma', mw: 126.20, defaultEq: 3.0, base: null, baseMW: null, baseEq: 0, coReagent: 'Oxyma', coReagentMW: 142.11, coReagentEq: 3.0 },
  { name: 'DIC/HOBt', mw: 126.20, defaultEq: 3.0, base: null, baseMW: null, baseEq: 0, coReagent: 'HOBt', coReagentMW: 135.13, coReagentEq: 3.0 },
  { name: 'HOBt', mw: 135.13, defaultEq: 3.0, base: null, baseMW: null, baseEq: 0, coReagent: null, coReagentMW: null, coReagentEq: 0 }
];

// ============================================================
// Cleavage Cocktails
// ============================================================
const CLEAVAGE_COCKTAILS = {
  standard: {
    name: 'Standard (Reagent K semplificato)',
    composition: 'TFA/TIS/H₂O (95:2.5:2.5)',
    time: '2–3 ore',
    notes: 'Per peptidi senza residui sensibili'
  },
  withCys: {
    name: 'Con Cys/Met (Reagent K)',
    composition: 'TFA/EDT/TIS/H₂O (94:2.5:1:2.5)',
    time: '2–3 ore',
    notes: 'EDT necessario per Cys(Trt) e Met'
  },
  withArg: {
    name: 'Con Arg(Pbf)',
    composition: 'TFA/TIS/H₂O (95:2.5:2.5)',
    time: '3–4 ore',
    notes: 'Tempo più lungo per rimozione completa del Pbf'
  },
  withCysArg: {
    name: 'Con Cys/Met + Arg(Pbf)',
    composition: 'TFA/EDT/TIS/H₂O (94:2.5:1:2.5)',
    time: '3–4 ore',
    notes: 'EDT + tempo lungo per Pbf'
  },
  shortCleavage: {
    name: 'Short Cleavage (test)',
    composition: 'TFA/TIS/H₂O (95:2.5:2.5)',
    time: '30–60 min',
    notes: 'Mini-cleavage su piccola aliquota di resina per test MS'
  }
};

// ============================================================
// Utility Functions
// ============================================================

/**
 * Tokenize a sequence allowing for (UnconventionalNames)
 */
function findMatchingParen(text, startIndex) {
  let depth = 0;
  for (let i = startIndex; i < text.length; i++) {
    if (text[i] === '(') depth++;
    else if (text[i] === ')') {
      depth--;
      if (depth === 0) return i;
    }
  }
  return -1;
}

function readFmocToken(seq, startIndex) {
  if (seq.slice(startIndex, startIndex + 5).toLowerCase() !== 'fmoc-') return null;

  let depth = 0;
  for (let i = startIndex; i < seq.length; i++) {
    if (seq[i] === '(') depth++;
    else if (seq[i] === ')') {
      if (depth === 0) break;
      depth--;
    }

    if (depth === 0 && seq.slice(i, i + 3).toUpperCase() === '-OH') {
      return {
        token: seq.slice(startIndex, i + 3).trim(),
        nextIndex: i + 3
      };
    }
  }

  return null;
}

function tokenizeSequence(seq) {
  if (!seq) return [];
  const tokens = [];
  let i = 0;
  while(i < seq.length) {
    const fmocToken = readFmocToken(seq, i);
    if (fmocToken) {
      tokens.push(fmocToken.token);
      i = fmocToken.nextIndex;
      continue;
    }

    if(seq[i] === '(') {
      let end = findMatchingParen(seq, i);
      if(end === -1) end = seq.length;
      tokens.push(seq.slice(i+1, end).trim());
      i = end + 1;
    } else if (seq[i].match(/[A-Za-z]/)) {
      tokens.push(seq[i].toUpperCase());
      i++;
    } else {
      i++;
    }
  }
  return tokens;
}

/**
 * Calculate peptide MW from sequence tokens
 * @param {string[]} tokens - Array of amino acid tokens
 * @param {string} cTerminus - 'acid' or 'amide'
 * @returns {number} Molecular weight in Da
 */
function calculatePeptideMW(tokens, cTerminus = 'amide', customMWs = {}) {
  let mw = 0;
  if (!Array.isArray(tokens)) return 0;
  for (const token of tokens) {
    if (AMINO_ACIDS[token]) {
      mw += AMINO_ACIDS[token].residueMW;
    } else if (customMWs && customMWs[token] && !isNaN(parseFloat(customMWs[token]))) {
      // Estimate residue MW from Fmoc-AA-OH MW: Fmoc-AA-OH - 240.26
      mw += parseFloat(customMWs[token]) - 240.26;
    }
  }
  // Add terminal groups
  if (cTerminus === 'amide') {
    mw += 17.031; // NH2 at C-term + H at N-term
  } else {
    mw += 18.015; // OH at C-term + H at N-term
  }
  return Math.round(mw * 100) / 100;
}

/**
 * Calculate mass for a single coupling step
 */
function calculateAAMass(token, equivalents, scale, customFmocMW = 0) {
  let fmocMW = customFmocMW;
  if (AMINO_ACIDS[token]) {
    fmocMW = AMINO_ACIDS[token].fmocMW;
  }
  if (!fmocMW) return 0;
  return Math.round(fmocMW * equivalents * scale * 100) / 100;
}

/**
 * Calculate resin mass
 */
function calculateResinMass(scale, loading) {
  return Math.round((scale / loading) * 1000 * 100) / 100;
}

/**
 * Calculate activator mass
 */
function calculateActivatorMass(activator, scale) {
  const result = {
    activatorMass: Math.round(activator.mw * activator.defaultEq * scale * 100) / 100,
    baseMass: 0,
    baseVolume: 0,
    coReagentMass: 0
  };

  if (activator.base) {
    const baseMassMg = activator.baseMW * activator.baseEq * scale;
    result.baseMass = Math.round(baseMassMg * 100) / 100;
    result.baseVolume = Math.round((baseMassMg / 0.742) * 100) / 100; // uL
  }

  if (activator.coReagent) {
    result.coReagentMass = Math.round(activator.coReagentMW * activator.coReagentEq * scale * 100) / 100;
  }

  return result;
}

/**
 * Suggest cleavage cocktail based on sequence
 */
function suggestCleavageCocktail(tokens) {
  // If tokens is a string, convert to array for includes check? 
  // Standard AAs are 1-letter, so it works.
  const seq = Array.isArray(tokens) ? tokens.join('') : String(tokens);
  const s = seq.toUpperCase();
  const hasCys = s.includes('C');
  const hasMet = s.includes('M');
  const hasArg = s.includes('R');
  const hasTrp = s.includes('W');

  const warnings = [];
  if (hasCys) warnings.push('Cys(Trt) presente -> EDT raccomandato');
  if (hasMet) warnings.push('Met presente -> EDT o tioanisolo raccomandato');
  if (hasArg) warnings.push('Arg(Pbf) presente -> tempo cleavage esteso (3-4h)');
  if (hasTrp) warnings.push('Trp(Boc) presente -> TIS raccomandato come scavenger');

  let cocktail;
  if ((hasCys || hasMet) && hasArg) {
    cocktail = CLEAVAGE_COCKTAILS.withCysArg;
  } else if (hasCys || hasMet) {
    cocktail = CLEAVAGE_COCKTAILS.withCys;
  } else if (hasArg) {
    cocktail = CLEAVAGE_COCKTAILS.withArg;
  } else {
    cocktail = CLEAVAGE_COCKTAILS.standard;
  }

  return { cocktail, warnings };
}

/**
 * Calculate common MS adducts
 */
function calculateMSAdducts(mw) {
  return [
    { name: '[M+H]+', mz: Math.round((mw + 1.008) * 100) / 100 },
    { name: '[M+2H]2+', mz: Math.round((mw + 2.016) / 2 * 100) / 100 },
    { name: '[M+3H]3+', mz: Math.round((mw + 3.024) / 3 * 100) / 100 },
    { name: '[M+Na]+', mz: Math.round((mw + 22.989) * 100) / 100 },
    { name: '[M+K]+', mz: Math.round((mw + 38.963) * 100) / 100 },
    { name: '[M+TFA+H]+', mz: Math.round((mw + 114.02 + 1.008) * 100) / 100 },
  ];
}
